from numpy import add
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import os
import traceback
import json
from datetime import datetime
from db_connector import MongoDBConnector
from io import BytesIO
from openpyxl.utils import get_column_letter
import re
from openpyxl.packaging import manifest



# 记得切环境


class ProcessingError(Exception):
    """处理错误的自定义异常类"""
    pass


def template_handler(keyword):
    """模板处理器装饰器"""
    def decorator(func):
        func._template_keyword = keyword
        return func
    return decorator


class InvoiceGenerator:
    def __init__(self, upload_folder, output_folder, db_connector=None, image_folder=None):
        """
        初始化发票生成器
        :param upload_folder: 上传文件夹路径
        :param output_folder: 输出文件夹路径
        :param db_connector: 数据库连接器（可选）
        :param image_folder: 图片文件夹路径（可选）
        """
        self.db_connector = db_connector if db_connector else MongoDBConnector()
        self.upload_folder = upload_folder
        self.output_folder = output_folder
        # 获取当前文件所在的目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.image_folder = os.path.join(current_dir, '产品图片(1)')  # 图片文件夹路径
        
        # 调试开关和统计
        self.debug_mode = False  # 可以通过环境变量控制
        self.missing_products = set()  # 记录缺失的产品信息
        
        # 产品信息缓存机制
        self.product_cache = {}  # 产品信息缓存字典 {msku: product_info}
        self.cache_enabled = True  # 缓存开关
        
        # 图片信息缓存机制
        self.image_cache = {}  # 图片路径缓存字典 {msku: image_path or None}
        self.image_cache_enabled = True  # 图片缓存开关
        
        # 定义模板配置，只列出不需要编码的模板
        self.template_config = {
            "依诺达": {"requires_code": False},  # 不需要编码的模板
            "罗马尼亚鹏城": {"requires_code": False}  # 不需要编码的模板
        }
        
        # 初始化模板处理器字典
        self._template_handlers = {}
        # 注册所有带有_template_keyword属性的方法
        print("开始注册模板处理器...")
        for name in dir(self):
            method = getattr(self, name)
            if hasattr(method, '_template_keyword'):
                keyword = method._template_keyword
                self._template_handlers[keyword] = method
                print(f"注册模板处理器: {name} -> {keyword}")
        print(f"已注册的模板处理器: {list(self._template_handlers.keys())}")

    def _log_missing_product(self, msku, template_name=""):
        """统一处理产品信息缺失的日志"""
        self.missing_products.add(msku)
        if self.debug_mode:
            print(f"警告: 未找到产品 {msku} 的信息 (模板: {template_name})")
    
    def _log_debug(self, message):
        """调试信息输出"""
        if self.debug_mode:
            print(f"[DEBUG] {message}")
    
    def _log_info(self, message):
        """重要信息输出"""
        print(f"[INFO] {message}")
    
    def should_enable_merge(self, address_info=None):
        """
        判断是否应该启用产品合并功能
        只有美国地址才启用合并
        
        :param address_info: 地址信息
        :return: 是否启用合并
        """
        if not address_info:
            print("[合并判断] 没有地址信息，跳过合并")
            return False
            
        try:
            # address_info的结构是：{'seller_info': {..., 'country_code': '...'}, 'address_info': {..., 'countryCode': '...'}}
            if isinstance(address_info, dict):
                # 优先从seller_info中获取country_code
                if 'seller_info' in address_info and 'country_code' in address_info['seller_info']:
                    country_code = address_info['seller_info']['country_code'].upper()
                # 备选：从address_info中获取countryCode
                elif 'address_info' in address_info and isinstance(address_info['address_info'], dict):
                    country_code = address_info['address_info'].get('countryCode', '').upper()
                # 兼容旧格式：直接从根级别获取countryCode
                else:
                    country_code = address_info.get('countryCode', '').upper()
            else:
                print("[合并判断] 地址信息格式不正确，跳过合并")
                return False
            
            # 只有美国才启用合并
            if country_code == 'US':
                print(f"[合并判断] 检测到美国地址 (countryCode: {country_code})，启用产品合并")
                return True
            else:
                print(f"[合并判断] 检测到非美国地址 (countryCode: {country_code})，跳过产品合并")
                return False
                
        except Exception as e:
            print(f"[合并判断] 判断过程中出现错误: {str(e)}，跳过合并")
            return False

    def merge_items_by_product_name(self, box_data, debug=False):
        """
        根据品名合并每个箱子内相同品名的产品（不跨箱合并）
        
        :param box_data: 箱子数据字典 {box_number: PackingListBox}
        :param debug: 是否输出调试信息
        :return: 合并后的箱子数据
        """
        from get_ticket_data import PackingListBox, PackingListItem
        
        merged_box_data = {}
        
        # 使用单一数据库连接处理所有箱子
        with self.db_connector as db:
            for box_number, box in box_data.items():
                if debug:
                    print(f"处理箱子 {box_number}，原始商品数: {len(box.items)}")
                
                # 创建新的箱子对象
                merged_box = PackingListBox(box_number)
                merged_box.length = box.length
                merged_box.width = box.width  
                merged_box.height = box.height
                merged_box.weight = box.weight
                
                # 按品名分组商品
                product_groups = {}
                
                for item in box.items:
                    # 直接使用装箱单中的product_name作为分组键
                    group_key = item.product_name.strip().lower() if item.product_name else item.msku
                    
                    if debug:
                        print(f"[调试] MSKU: {item.msku}, 品名: {item.product_name}, 分组键: {group_key}")
                        print(f"[调试] 该商品的箱子数量分布: {item.box_quantities}")
                    
                    if group_key not in product_groups:
                        product_groups[group_key] = []
                    product_groups[group_key].append(item)
            
                # 合并每个分组
                for group_key, items in product_groups.items():
                    if debug:
                        print(f"\n[调试] ===== 处理分组: '{group_key}' =====")
                        print(f"[调试] 该分组包含 {len(items)} 个商品")
                        
                        if len(items) > 1:
                            print(f"[调试] 需要合并的商品列表:")
                            for i, item in enumerate(items):
                                print(f"[调试]   商品{i+1}: MSKU={item.msku}, 品名={item.product_name}, 箱子数量={item.box_quantities}")
                        else:
                            print(f"[调试] 单个商品，无需合并: MSKU={items[0].msku}, 品名={items[0].product_name}")
                    
                    # 使用第一个商品作为基础（其他属性取第一个）
                    base_item = items[0]
                    
                    # 创建合并后的商品
                    merged_item = PackingListItem(
                        sequence_no=base_item.sequence_no,
                        msku=base_item.msku,  # 保留第一个MSKU
                        fnsku=base_item.fnsku,
                        product_name=base_item.product_name,  # 保留品名
                        sku=base_item.sku,
                        quantity=0,  # 将重新计算
                        box_quantities={}
                    )
                    
                    # 合并数量、重量、价格等数值字段
                    total_weight_sum = 0  # 总重量和（重量×数量）
                    total_price_sum = 0   # 总价格和（价格×数量）
                    total_quantity = 0    # 总数量
                    
                    if debug:
                        print(f"[调试] 开始合并数量和价格...")
                    
                    for i, item in enumerate(items):
                        if debug:
                            print(f"\n[调试] --- 处理第{i+1}个商品: {item.msku} ---")
                        
                        # 获取该商品的产品信息用于计算重量和价格
                        item_product_info = self._get_product_info(item.msku, db)
                        
                        # 获取该商品的总数量
                        current_item_total_qty = sum(item.box_quantities.values())
                        total_quantity += current_item_total_qty
                        
                        if debug:
                            print(f"[调试] 该商品总数量: {current_item_total_qty}")
                            print(f"[调试] 该商品箱子分布: {item.box_quantities}")
                        
                        # 累加重量（产品重量 × 数量）
                        if item_product_info and item_product_info.get('weight'):
                            try:
                                item_weight = float(item_product_info['weight']) * current_item_total_qty
                                total_weight_sum += item_weight
                                if debug:
                                    print(f"[调试] 该商品重量: {item_product_info.get('weight')} × {current_item_total_qty} = {item_weight}")
                            except (ValueError, TypeError):
                                if debug:
                                    print(f"[调试] 该商品重量数据无效: {item_product_info.get('weight')}")
                        
                        # 累加价格（单价 × 数量）
                        if item_product_info and item_product_info.get('price'):
                            try:
                                unit_price = float(item_product_info['price'])
                                item_price = unit_price * current_item_total_qty
                                total_price_sum += item_price
                                if debug:
                                    print(f"[调试] 该商品价格: {unit_price} × {current_item_total_qty} = {item_price}")
                            except (ValueError, TypeError):
                                if debug:
                                    print(f"[调试] 该商品价格数据无效: {item_product_info.get('price')}")
                        
                        # 合并箱子数量
                        for box_num, qty in item.box_quantities.items():
                            if box_num not in merged_item.box_quantities:
                                merged_item.box_quantities[box_num] = 0
                            merged_item.box_quantities[box_num] += qty
                    
                    # 更新总数量
                    merged_item.quantity = total_quantity
                    
                    # 计算加权平均重量和价格
                    weighted_avg_weight = total_weight_sum / total_quantity if total_quantity > 0 else 0
                    weighted_avg_price = total_price_sum / total_quantity if total_quantity > 0 else 0
                    
                    if debug:
                        print(f"\n[调试] === 合并结果汇总 ===")
                        print(f"[调试] 最终箱子数量分布: {merged_item.box_quantities}")
                        print(f"[调试] 最终总数量: {total_quantity}")
                        print(f"[调试] 总重量和: {total_weight_sum:.2f}")
                        print(f"[调试] 总价格和: {total_price_sum:.2f}")
                        print(f"[调试] 加权平均重量: {weighted_avg_weight:.2f}")
                        print(f"[调试] 加权平均价格: {weighted_avg_price:.2f}")
                    
                    # 为合并后的商品添加合并信息标记
                    if len(items) > 1:
                        merged_item._is_merged = True
                        merged_item._merged_count = len(items)
                        merged_item._merged_mskus = [item.msku for item in items]
                        merged_item._weighted_avg_weight = weighted_avg_weight
                        merged_item._weighted_avg_price = weighted_avg_price
                        merged_item._total_weight_sum = total_weight_sum
                        merged_item._total_price_sum = total_price_sum
                        
                        if debug:
                            print(f"[调试] 合并的MSKU列表: {[item.msku for item in items]}")
                            print(f"[调试] 标记为合并商品")
                    else:
                        merged_item._is_merged = False
                        # 单个商品也保存重量和价格信息
                        merged_item._weighted_avg_weight = weighted_avg_weight
                        merged_item._weighted_avg_price = weighted_avg_price
                        if debug:
                            print(f"[调试] 单个商品，无合并标记")
                    
                    if debug:
                        print(f"[调试] ===== 分组处理完成 =====\n")
                    
                    merged_box.add_item(merged_item)
                
                merged_box_data[box_number] = merged_box
                
                print(f"\n[调试] ##### 箱子 {box_number} 处理完成 #####")
                print(f"[调试] 原始商品数: {len(box.items)}")
                print(f"[调试] 合并后商品数: {len(merged_box.items)}")
                print(f"[调试] 分组数: {len(product_groups)}")
                
                # 显示最终结果
                print(f"[调试] 最终商品列表:")
                for i, final_item in enumerate(merged_box.items):
                    print(f"[调试]   商品{i+1}: MSKU={final_item.msku}, 总数量={final_item.quantity}, 箱子分布={final_item.box_quantities}")
                    if hasattr(final_item, '_is_merged') and final_item._is_merged:
                        print(f"[调试]     -> 这是合并商品，包含{final_item._merged_count}个原商品")
                print(f"[调试] ##### 箱子 {box_number} 处理完成 #####\n")
        
        print(f"\n[调试] ========== 所有箱子合并完成 ==========\n")
        return merged_box_data
    
    def _get_display_price(self, item, product_info):
        """获取显示用的单价"""
        if hasattr(item, '_weighted_avg_price'):
            # 合并商品或单个商品都显示加权平均单价
            avg_price = getattr(item, '_weighted_avg_price', 0)
            return f"{avg_price:.2f}" if avg_price > 0 else "0.00"
        else:
            # 普通商品显示原价格
            return product_info.get('price', '') if product_info else ''
    
    def _get_total_price(self, item, box_number, product_info):
        """获取该箱子中该商品的总价格"""
        if hasattr(item, '_weighted_avg_price'):
            # 合并商品或单个商品：使用加权平均单价 × 该箱子数量
            box_qty = item.box_quantities.get(box_number, 0)
            avg_price = getattr(item, '_weighted_avg_price', 0)
            total_price = avg_price * box_qty
            return f"{total_price:.2f}" if total_price > 0 else "0.00"
        else:
            # 普通商品：数量 × 单价
            if product_info:
                box_qty = item.box_quantities.get(box_number, 0)
                price = product_info.get('price', 0)
                total_price = box_qty * float(price) if price else 0
                return f"{total_price:.2f}" if total_price > 0 else "0.00"
            else:
                return "0.00"
    
    def _extract_asin_from_link(self, link):
        """从Amazon链接中提取ASIN码
        
        Args:
            link (str): Amazon产品链接，例如 'https://www.amazon.com.au/dp/B0FM42H4BX/'
            
        Returns:
            str: 提取的ASIN码，例如 'B0FM42H4BX'，如果提取失败返回空字符串
        """
        import re
        
        if not link or not isinstance(link, str):
            return ''
        
        # Amazon ASIN格式：10个字符，由大写字母和数字组成
        # 常见的Amazon链接格式：
        # https://www.amazon.com/dp/B0FM42H4BX/
        # https://www.amazon.com.au/dp/B0FM42H4BX/
        # https://www.amazon.co.uk/dp/B0FM42H4BX/
        # https://www.amazon.com/gp/product/B0FM42H4BX/
        
        # 匹配 /dp/ 或 /gp/product/ 后面的ASIN
        patterns = [
            r'/dp/([A-Z0-9]{10})(?:/|$|\?)',
            r'/gp/product/([A-Z0-9]{10})(?:/|$|\?)',
            r'[?&]asin=([A-Z0-9]{10})(?:&|$)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, link, re.IGNORECASE)
            if match:
                asin = match.group(1).upper()
                # 验证ASIN格式：10个字符，字母和数字组成
                if len(asin) == 10 and re.match(r'^[A-Z0-9]{10}$', asin):
                    return asin
        
        return ''
    
    def _get_merged_product_display_name(self, item, product_info):
        """获取合并商品的显示名称"""
        if hasattr(item, '_is_merged') and item._is_merged:
            merged_count = getattr(item, '_merged_count', 1)
            # 直接使用装箱单中的品名作为基础名称
            base_name = item.product_name or item.msku
            
            # 如果合并了多个商品，在名称后面添加标识
            if merged_count > 1:
                return f"{base_name} (合并{merged_count}个商品)"
            else:
                return base_name
        else:
            # 普通商品返回装箱单中的品名
            return item.product_name or item.msku
    
    def _print_missing_summary(self):
        """打印缺失产品信息的汇总"""
        if self.missing_products:
            missing_list = sorted(list(self.missing_products))
            print(f"\n=== 产品信息缺失汇总 ===")
            print(f"共有 {len(missing_list)} 个产品缺失信息:")
            for msku in missing_list:
                print(f"  - {msku}")
            print("=" * 30)
        else:
            print("\n=== 产品信息完整 ===")
            print("所有产品信息都已找到")
            print("=" * 20)
    
    def _collect_all_mskus(self, box_data):
        """从箱子数据中收集所有需要的MSKU"""
        mskus = set()
        for box in box_data.values():
            for item in box.items:
                mskus.add(item.msku)
        return list(mskus)
    
    def _preload_product_info(self, mskus):
        """批量预加载产品信息到缓存"""
        if not self.cache_enabled or not mskus:
            return
        
        self._log_info(f"开始预加载 {len(mskus)} 个产品信息...")
        
        try:
            with self.db_connector as db:
                # 批量查询所有产品信息
                products = db['msku_info'].find({'msku': {'$in': mskus}})
                
                loaded_count = 0
                for product in products:
                    msku = product.get('msku')
                    if msku:
                        self.product_cache[msku] = {
                            'cn_name': product.get('productNameZh', ''),
                            'en_name': product.get('productNameEn', ''),
                            'en_usage': product.get('useEn', ''),
                            'ch_usage': product.get('useZh', ''),
                            'material_en': product.get('materialEn', ''),
                            'material_cn': product.get('materialZh', ''),
                            'hs_code': product.get('HS', ''),
                            'usage_en': product.get('useEn', ''),
                            'usage_cn': product.get('useZh', ''),
                            'brand': product.get('brand', ''),
                            'model': product.get('model', ''),
                            'link': product.get('productLink', ''),
                            'price': product.get('askprice', ''),
                            'electrified': product.get('electrified', ''),
                            'magnetic': product.get('magnetic', ''),
                            'weight': product.get('weight', ''),
                        }
                        loaded_count += 1
                
                self._log_info(f"预加载完成: {loaded_count}/{len(mskus)} 个产品信息已缓存")
                
                # 记录未找到的产品
                cached_mskus = set(self.product_cache.keys())
                missing_mskus = set(mskus) - cached_mskus
                if missing_mskus:
                    self._log_debug(f"未找到产品信息的MSKU: {missing_mskus}")
                
        except Exception as e:
            print(f"预加载产品信息时发生错误: {str(e)}")
            # 如果预加载失败，禁用缓存，回退到原有方式
            self.cache_enabled = False
    
    def _preload_image_info(self, mskus):
        """批量预加载图片存在性信息到缓存"""
        if not self.image_cache_enabled or not mskus:
            return
        
        self._log_info(f"开始预加载 {len(mskus)} 个产品的图片信息...")
        
        found_count = 0
        missing_count = 0
        
        for msku in mskus:
            # 检查jpg和png两种格式
            jpg_path = os.path.join(self.image_folder, f"{msku}.jpg")
            png_path = os.path.join(self.image_folder, f"{msku}.png")
            
            if os.path.exists(jpg_path):
                self.image_cache[msku] = jpg_path
                found_count += 1
                self._log_debug(f"找到图片: {msku}.jpg")
            elif os.path.exists(png_path):
                self.image_cache[msku] = png_path
                found_count += 1
                self._log_debug(f"找到图片: {msku}.png")
            else:
                self.image_cache[msku] = None
                missing_count += 1
                self._log_debug(f"图片不存在: {msku}")
        
        self._log_info(f"图片预加载完成: 找到 {found_count} 个，缺失 {missing_count} 个")
    
    def _print_cache_statistics(self):
        """打印缓存使用统计"""
        print(f"\n=== 缓存统计 ===")
        
        # 产品信息缓存统计
        if self.cache_enabled:
            cache_size = len(self.product_cache)
            print(f"产品信息缓存: 已启用")
            print(f"缓存产品数量: {cache_size}")
            if cache_size > 0:
                print(f"缓存产品列表: {', '.join(sorted(self.product_cache.keys())[:5])}{'...' if cache_size > 5 else ''}")
        else:
            print(f"产品信息缓存: 已禁用")
        
        # 图片信息缓存统计
        if self.image_cache_enabled:
            image_cache_size = len(self.image_cache)
            found_images = sum(1 for path in self.image_cache.values() if path is not None)
            missing_images = image_cache_size - found_images
            print(f"图片信息缓存: 已启用")
            print(f"图片缓存数量: {image_cache_size} (找到: {found_images}, 缺失: {missing_images})")
        else:
            print(f"图片信息缓存: 已禁用")
            
        print("=" * 20)
    
    def clear_cache(self):
        """清空产品信息缓存"""
        self.product_cache.clear()
        self._log_info("产品信息缓存已清空")
    
    def disable_cache(self):
        """禁用缓存功能"""
        self.cache_enabled = False
        self.product_cache.clear()
        self._log_info("产品信息缓存已禁用")
    
    def enable_cache(self):
        """启用缓存功能"""
        self.cache_enabled = True
        self._log_info("产品信息缓存已启用")
    
    @template_handler("叮铛卡航限时达")
    def _fill_dingdang_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充叮铛卡航限时达模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:

                # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['模板']  # 获取模板工作表

                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                # 在第一行B列填充编码
                if code:
                    cell = sheet.cell(row=1, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=9)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    try:
                        # 填充收件人信息
                       
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=4, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=3, column=2)  # B2单元格
                            cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        # 填充地址信息
                        if 'addressLine1' in address_info_detail:
                            cell = sheet.cell(row=6, column=2)  # B3单元格
                            cell.value = address_info_detail['addressLine1']

                        # 城市
                        if 'city' in address_info_detail:
                            cell = sheet.cell(row=8, column=2)  # B4单元格
                            cell.value = address_info_detail['city']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=10, column=2)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=11, column=2)  # B7单元格
                            cell.value = address_info_detail['countryCode']

                        if address_parts:
                            # 检查warehouseId并添加到最前面
                            final_address_parts = address_parts.copy()
                            if 'warehouseId' in address_info_detail:
                                warehouse_id = str(address_info_detail['warehouseId'])
                                # 设置第4行的warehouse_id
                                cell_warehouse_4 = sheet.cell(row=4, column=2)  # 第4行B列
                                cell_warehouse_4.value = warehouse_id 

                                # 设置第3行的warehouse_id
                                cell_warehouse_3 = sheet.cell(row=3, column=2)  # 第3行B列
                                cell_warehouse_3.value = warehouse_id
   
                                # 设置第5行的公司名称
                                cell_company = sheet.cell(row=5, column=2)  # 第5行B列
                                if warehouse_id not in address_info_detail['name']:
                                    cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                    cell_company.value = cell_result_value
                                else:
                                    cell_company.value = address_info_detail['name']

                                # 如果地址部分中没有warehouse_id，则添加
                                if warehouse_id not in address_parts:
                                    final_address_parts.insert(0, warehouse_id)
                            
                            # 设置第6行的完整地址
                            cell_address = sheet.cell(row=6, column=2)  # 第6行B列
                            cell_address.value = ', '.join(final_address_parts)
                            
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=16, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 检查所有产品的电磁属性
                has_electric = False
                has_magnetic = False
                for box in box_data.values():
                    for item in box.items:
                        product_info = self._get_product_info(item.msku, db)
                        if product_info:
                            if product_info.get('electrified', '') == '是':
                                has_electric = True
                            if product_info.get('magnetic', '') == '是':
                                has_magnetic = True
                            if has_electric and has_magnetic:
                                break
                    if has_electric and has_magnetic:
                        break

                # 在表格顶部添加电磁属性标记
                if has_electric:
                    cell = sheet.cell(row=1, column=6)  # F列第1行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)
                
                if has_magnetic:
                    cell = sheet.cell(row=2, column=6)  # F列第2行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)

                # 填充数据
                row_num = 18  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始

                # 遍历每个箱子

                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))

                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))

                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 设置单元格值和样式ç
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (3,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (4, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            # (5, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (5, self._get_display_price(item, product_info)),   # 单价

                            (6, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (7, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (8, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (9, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (10, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (11, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (12, product_info.get('link', '') if product_info else ''),
                            (14, ''),  # 图片列 (N列)
                            (15, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                            (17, box.length if box.length is not None else ""),  # 长度 (Q列)
                            (18, box.width if box.width is not None else ""),    # 宽度 (R列)
                            (19, box.height if box.height is not None else "")   # 高度 (S列)
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"N{row_num}"  # 图片列（第14列）
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("顺丰")
    def _fill_sf_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """填充顺丰模板"""
        """
        填充顺丰模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:

                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data


                sheet = wb['Sheet1']  # 获取模板工作表
                # box_Reference_id = ''  # 在方法开始时就初始化
                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                self.unmerge_cells_in_range(sheet, 2, 4, 2, 9)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    # if 'seller_info' in address_info:
                    #     box_Reference_id = address_info['seller_info']['amazonReferenceId']
                    try:
                        # 填充收件人信息
                       
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=2, column=2)  
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=3, column=2)  
                            cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        if address_parts:
                            # 检查warehouseId并添加到最前面
                            final_address_parts = address_parts.copy()
                            
                            if 'warehouseId' in address_info_detail:
                                warehouse_id = str(address_info_detail['warehouseId'])
                                # 如果地址部分中没有warehouse_id，则添加
                                if warehouse_id not in address_parts:
                                    final_address_parts.insert(0, warehouse_id)
                                
                                # 设置第3行的warehouse_id
                                cell_warehouse = sheet.cell(row=3, column=2)  # 第3行B列
                                cell_warehouse.value = warehouse_id 
    
                                # 设置第2行的公司名称
                                cell_company = sheet.cell(row=2, column=2)  # 第2行B列
                                if warehouse_id not in address_info_detail['name']:
                                    cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                    cell_company.value = cell_result_value
                                else:
                                    cell_company.value = address_info_detail['name']
                            
                            # 设置第4行的完整地址
                            cell_address = sheet.cell(row=4, column=2)  # 第4行B列
                            cell_address.value = ', '.join(final_address_parts)

                                
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                # 填充数据
                row_num = 12  
                index = 1    # 添加序号计数器，从1开始

                # 将box_data按箱号排序
                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))

                 # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        price = 0
                        total_price = 0

                        if product_info is not None:
                            item.product_name = product_info.get('cn_name', item.product_name)
                        else:
                            # 处理未找到产品信息的情况
                            self._log_missing_product(item.msku)
                            item.product_name = "需要补数据"
                        
                        # box_number_str = code+f"{box_number:05d}" 
                        box_number_str = code+'U00000'+str(box_number)
                        # Reference_id = ''  # 初始化为None
                        # if box_Reference_id:
                        #     Reference_id = box_Reference_id
                        
                        # 设置单元格值和样式
                        cell_data = [
                            # 基本信息
                            (1, box_number_str),                                   # 货箱编号
                            # (2, Reference_id),                                    # 参考编号
                            (3, item.msku),                                       # 商品编码
                            # 产品名称信息
                            (4, product_info.get('en_name', '') if product_info else ''),                 # 英文名称
                            (5, product_info.get('cn_name', '') if product_info else ''),                 # 中文名称
                            (6, product_info.get('brand', '') if product_info else ''),                   # 品牌
                            (7, product_info.get('model', '') if product_info else ''),                   # 型号
                            
                            # 产品材料和用途
                            (8, product_info.get('material_cn', '') if product_info else ''),             # 中文材料
                            (9, product_info.get('material_en', '') if product_info else ''),             # 英文材料
                            (10, str(product_info.get('usage_en', '') + 
                                   product_info.get('usage_cn', '')) if product_info else ''),            # 用途
                            (11, '纸箱'),                                         # 包装类型
                            
                            # 产品规格信息
                            (12, product_info.get('hs_code', '') if product_info else ''),                # HS编码
                            (13, item.box_quantities.get(box_number, 0)),         # 数量
                            
                            # 箱体信息
                            (16, box.length if box.length is not None else ""),   # 长度
                            (17, box.width if box.width is not None else ""),     # 宽度
                            (18, box.height if box.height is not None else ""),   # 高度
                            (19, box.weight if box.weight is not None else ""),   # 重量
                            
                            # 其他信息
                            (20, product_info.get('link', '') if product_info else ''),                   # 链接
                            (21, '')                                              # 图片占位
                        ]
                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = sheet.row_dimensions[12].height
                        
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"U{row_num}"  # 图片列（第14列）
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1
                
                self.merge_cells_in_range(sheet, 2, 2, 2, 9)
                self.merge_cells_in_range(sheet, 3, 3, 2, 9)
                self.merge_cells_in_range(sheet, 4, 4, 2, 9)

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("依诺达")
    def _fill_ynd_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充叮铛卡航限时达模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:
                sheet = wb['Sheet1']  # 获取模板工作表
                self.unmerge_cells_in_range(sheet, 15, 15, 2, 4)
                self.unmerge_cells_in_range(sheet, 1, 1, 6, 8)
                self.unmerge_cells_in_range(sheet, 2, 2, 6, 8)
                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=15, column=2)  #填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 检查所有产品的电磁属性
                has_electric = False
                has_magnetic = False
                for box in box_data.values():
                    for item in box.items:
                        product_info = self._get_product_info(item.msku, db)
                        if product_info:
                            if product_info.get('electrified', '') == '是':
                                has_electric = True
                            if product_info.get('magnetic', '') == '是':
                                has_magnetic = True
                            if has_electric and has_magnetic:
                                break
                    if has_electric and has_magnetic:
                        break

                # 在表格顶部添加电磁属性标记
                if has_electric:
                    cell = sheet.cell(row=1, column=6)  # F列第1行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)
                
                if has_magnetic:
                    cell = sheet.cell(row=2, column=6)  # F列第2行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)

                # 填充数据
                row_num = 17  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[17].height
                sheet.column_dimensions['Q'].width = row_height/4
                
                # 遍历每个箱子

                sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 设置单元格值和样式ç
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, box.weight if box.weight is not None else ""),  # 重量 
                            (3, box.length if box.length is not None else ""),  # 长度 
                            (4, box.width if box.width is not None else ""),    # 宽度 
                            (5, box.height if box.height is not None else "") ,  # 高度 
                            (6, item.msku),                    
                            (7,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (8, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            # (5, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (10, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (11, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (13, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (12, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (14, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (15, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (16, product_info.get('link', '') if product_info else ''),
                            (17, ''),  # 图片列 (N列)
                            # (15, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                          
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)
                        sheet.row_dimensions[row_num].height = row_height
                       
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"Q{row_num}"
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

                self.merge_cells_in_range(sheet, 15, 15, 2, 4)
                self.merge_cells_in_range(sheet, 1, 1, 6, 8)
                self.merge_cells_in_range(sheet, 2, 2, 6, 8)

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("叮铛(美洲)")
    def _fill_ddmz_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充叮铛卡航限时达模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:

                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['清关发票']  # 获取模板工作表

                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                self.unmerge_cells_in_range(sheet, 3, 3, 8, 15)
                self.unmerge_cells_in_range(sheet, 4, 4, 8, 15)
                self.unmerge_cells_in_range(sheet, 5, 5, 8, 11)
                self.unmerge_cells_in_range(sheet, 5, 5, 13, 15)
                self.unmerge_cells_in_range(sheet, 6, 6, 8, 15)
                self.unmerge_cells_in_range(sheet, 7, 7, 8, 11)
                self.unmerge_cells_in_range(sheet, 7, 7,13, 15)
                self.unmerge_cells_in_range(sheet, 8, 8,13, 15)

                # 在第一行B列填充编码
                if code:
                    cell = sheet.cell(row=3, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=9)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    try:
                        # 填充收件人信息，这里收件人和
                        
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=3, column=8)  # B2单元格
                            # cell.value = address_info_detail['name']

                            # cell = sheet.cell(row=4, column=8)  # B2单元格
                            # cell.value = address_info_detail['name']

                            # cell = sheet.cell(row=8, column=13)  # B2单元格
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        # 填充地址信息
                        # if 'addressLine1' in address_info_detail:
                        #     cell = sheet.cell(row=6, column=2)  # B3单元格
                        #     cell.value = address_info_detail['addressLine1']

                        # 城市
                        if 'city' in address_info_detail:
                            cell = sheet.cell(row=7, column=8)  # B4单元格
                            cell.value = address_info_detail['city']

                        #省份
                        if 'stateOrProvinceCode' in address_info_detail:
                            cell = sheet.cell(row=7, column=13)  # B5单元格
                            cell.value = address_info_detail['stateOrProvinceCode']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=5, column=13)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=5, column=8)  # B7单元格
                            cell.value = address_info_detail['countryCode']

                        if address_parts:
                            # 检查warehouseId并添加到最前面
                            final_address_parts = address_parts.copy()
                            
                            if 'warehouseId' in address_info_detail:
                                warehouse_id = str(address_info_detail['warehouseId'])
                                # 如果地址部分中没有warehouse_id，则添加
                                if warehouse_id not in address_parts:
                                    final_address_parts.insert(0, warehouse_id)
                                
                                # 设置第3行H列的warehouse_id
                                cell_warehouse_3 = sheet.cell(row=3, column=8)  # 第3行H列
                                cell_warehouse_3.value = warehouse_id 
                                
                                # 设置第8行M列的warehouse_id
                                cell_warehouse_8 = sheet.cell(row=8, column=13)  # 第8行M列
                                cell_warehouse_8.value = warehouse_id 

                                # 设置第4行H列的公司名称
                                cell_company = sheet.cell(row=4, column=8)  # 第4行H列
                                if warehouse_id not in address_info_detail['name']:
                                    cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                    cell_company.value = cell_result_value
                                else:
                                    cell_company.value = address_info_detail['name']
                            
                            # 设置第6行H列的完整地址
                            cell_address = sheet.cell(row=6, column=8)  # 第6行H列
                            cell_address.value = ', '.join(final_address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=3, column=6)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")


                # 填充数据
                row_num = 12  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[12].height
                Reference_id = ''
                
                # 遍历每个箱子
                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))

                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        Reference_id = address_info['address_info'].get('amazonReferenceId','') if address_info and address_info.get('address_info') else ''

                        # 设置单元格值和样式
                        cell_data = [
                            (1, code+f"{box_number:05d}"),                    
                            # (2, box.weight if box.weight is not None else ""),  
                            (2,Reference_id),  
                            (3,f"{box.length}*{box.width}*{box.height}"),  
                            (4,box_number), 
                            (5,box.weight), 
                            (6, product_info.get('hs_code', '') if product_info else ''),  
                            (8,product_info.get('en_name', '') if product_info else ''),  
                            (7, product_info.get('cn_name', '') if product_info else ''), 
                            (9, item.box_quantities.get(box_number, 0)),
                            (10,''),
                            (11, product_info.get('brand', '') if product_info else ''),    
                            (12, product_info.get('model', '') if product_info else ''),  
                            (13, str(product_info.get('material_cn', '')) if product_info else ''),  
                            (14, str(product_info.get('usage_cn', '')+product_info.get('usage_en', '' ))if product_info else ''),    # 用途 (H列)
                            (15, ''),  
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                        
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"O{row_num}"  # 图片列（第14列）
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1
                
                self.merge_cells_in_range(sheet, 3, 3, 8, 15)
                self.merge_cells_in_range(sheet, 4, 4, 8, 15)
                self.merge_cells_in_range(sheet, 5, 5, 8, 11)
                self.merge_cells_in_range(sheet, 5, 5, 13, 15)
                self.merge_cells_in_range(sheet, 6, 6, 8, 15)
                self.merge_cells_in_range(sheet, 7, 7, 8, 11)
                self.merge_cells_in_range(sheet, 7, 7,13, 15)
                self.merge_cells_in_range(sheet, 8, 8,13, 15)
                
            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("UPS(美洲)")
    def _fill_ups_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充UPS美洲模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:
                
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['发票']  # 获取发票工作表
                
                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }


                # 解除合并单元格
                print("正在解除合并单元格...")
                # self.unmerge_cells_in_range(sheet, 13, 16, 1, 15)
                self.unmerge_cells_in_range(sheet, 4, 4, 1, 3)
                self.unmerge_cells_in_range(sheet, 7, 11, 1, 3)
                self.unmerge_cells_in_range(sheet, 7, 11, 4, 15)
                
                row_num = 13  # 从第13行开始填充数据
                # 增加一行，用于显示Reference ID
                row_num += 1
           
                if code:
                    cell = sheet.cell(row=4, column=1)  # B列是第2列
                    cell.value = f"运单号码:{code}"
                    cell.font = Font(name='Arial', size=12)

                    cell_another = sheet.cell(row=5, column=1)  # B列是第2列
                    
                    Reference_id = address_info['address_info'].get('amazonReferenceId','') if address_info and address_info.get('address_info') else ''
                    cell_another.value = f"Reference ID:{Reference_id}"
                    cell_another = sheet.cell(row=7, column=1)  # B列是第2列
                    cell_another.value = f"FBA编号:{code}"
                    cell_another.font = Font(name='Arial', size=12)
                
                  # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    try:

                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        if address_parts:
                            cell = sheet.cell(row=7, column=4)  # B3单元格
                            
                            final_address_parts = address_parts.copy()
                            if 'warehouseId' in address_info_detail:
                                warehouse_id = str(address_info_detail['warehouseId'])
                                if warehouse_id not in address_parts:
                                    final_address_parts.insert(0, warehouse_id)
                            cell.value = ', '.join(final_address_parts)
                            cell.font = Font(name='Arial', size=12)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")
            
                # 遍历每个箱子
                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                row_height = sheet.row_dimensions[13].height
                
                # 记录数据起始行，用于后续公式计算
                data_start_row = row_num

                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    if not box.items:
                        continue
                    
                    start_row = row_num  # 记录当前箱子的起始行
                    box_items_count = len(box.items)
                    
                    for item in box.items:
                        # 获取产品信息
                        db_product_info = self._get_product_info(item.msku, db)
                        if not db_product_info:
                            continue

                        # 构建产品名称和获取数量、价格
                        name = f"{db_product_info.get('en_name', '')}({db_product_info.get('cn_name', '')})"
                        quantity = item.box_quantities.get(box_number, 0)
                        price = db_product_info.get('price', 0)
                        if price == '':
                            price = 0

                        # 移除累计变量计算，改为使用Excel公式
                        # 检查重量是否为None（仅用于警告）
                        if box.weight is None:
                            print(f"警告：UPS模板中箱子 {box_number} 的重量数据为None")

                        # 设置单元格值
                        cell_data = [
                            (1, f"{code}U00000{box_number}" if item == box.items[0] else ""),  # FBA号,只在第一行显示
                            (2, box_number if item == box.items[0] else ""),  # 箱号,只在第一行显示
                            (3, db_product_info.get('cn_name', '') if db_product_info else ''),  # 中文品名
                            (4, db_product_info.get('en_name', '') if db_product_info else ''),  # 英文品名
                            
                            (6, item.box_quantities.get(box_number, 0)),  # 数量
                            (5, self._get_display_price(item, db_product_info)),   # 单价
                            (7, self._get_total_price(item, box_number, db_product_info)),   # 总价
                            # (7, float(price) * quantity),  # 总价
                            (8, f"{db_product_info.get('material_cn', '')}/{db_product_info.get('material_en', '')}" if db_product_info else ''),  # 材质
                            (9, f"{db_product_info.get('usage_cn', '')}/{db_product_info.get('usage_en', '')}" if db_product_info else ''),  # 用途
                            (10, box.weight if item == box.items[0] else ""),  # 毛重,只在第一行显示
                            (11, box.length if item == box.items[0] else ""),  # 长,只在第一行显示
                            (12, box.width if item == box.items[0] else ""),  # 宽,只在第一行显示
                            (13, box.height if item == box.items[0] else ""),  # 高,只在第一行显示
                            (14, db_product_info.get('brand', '') if db_product_info else ''),  # 品牌
                            (15, db_product_info.get('hs_code', '') if db_product_info else '')  # HS编码
                        ]
                        
                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                        sheet.column_dimensions['O'].width = row_height
                        row_num += 1
                    
                    # 如果这个箱子有多个产品,需要合并单元格
                    if box_items_count > 1:
                        merge_columns = [1, 2, 10, 11, 12, 13]  # 需要合并的列
                        for col in merge_columns:
                            self.merge_cells_in_range(sheet, start_row, row_num-1, col, col)

                # 添加总计行，使用Excel公式计算总数
                total_row = row_num  # 直接使用当前行号，不再加1
                data_end_row = row_num - 1  # 数据结束行
                
                self._set_cell_value(sheet, total_row, 1, "总件数", style_info)
                self._set_cell_value(sheet, total_row, 2, len(box_data), style_info)
                # 使用SUM公式计算总数量（第6列）
                self._set_cell_value(sheet, total_row, 6, f"=SUM(F{data_start_row}:F{data_end_row})", style_info)
                # 使用SUM公式计算总金额（第7列）
                self._set_cell_value(sheet, total_row, 7, f"=SUM(G{data_start_row}:G{data_end_row})", style_info)
                self._set_cell_value(sheet, total_row, 9, "总重", style_info)
                # 使用SUM公式计算总重量（第10列），只计算非空单元格
                self._set_cell_value(sheet, total_row, 10, f"=SUM(J{data_start_row}:J{data_end_row})", style_info)
                
                def set_cell_value(sheet, row, column, value, font_size=12):
                    cell = sheet.cell(row=row, column=column)
                    cell.value = value
                    cell.font = Font(name='Arial', size=font_size,bold=True)
                    no_border = Border(left=Side(border_style=None),
                                right=Side(border_style=None),
                                top=Side(border_style=None),
                                bottom=Side(border_style=None))
                    cell.border = no_border
                    
                def add_border_to_range(sheet, start_row, end_row, start_col, end_col):
                    border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    for row in range(start_row, end_row + 1):
                        for col in range(start_col, end_col + 1):
                            cell = sheet.cell(row=row, column=col)
                            cell.border = border

                    # 为数据区域添加边框（从表头到合计行）
                add_border_to_range(sheet,total_row, total_row, 1, 15)

                # 使用辅助函数设置单元格
                made_in_row = total_row + 2
                current_date = datetime.now().strftime("%Y.%m.%d")
                set_cell_value(sheet, made_in_row, 11, "DATE")
                set_cell_value(sheet, made_in_row+1, 11, f"签字日期:{current_date}")
                set_cell_value(sheet, made_in_row, 1, "Made in China")

                # 设置行高
                # for row in range(13, row_num):
                #     sheet.row_dimensions[row].height = 
                self.merge_cells_in_range(sheet, 4, 4, 1, 3)
                self.merge_cells_in_range(sheet, 7, 11, 1, 3)
                self.merge_cells_in_range(sheet, 7, 11, 4, 15)
                

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("林道")
    def _fill_lindao_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充林道模板
        :param wb: Excel工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data
                    
                sheet = wb['模板']  # 获取模板工作表

                self.unmerge_cells_in_range(sheet, 3, 3, 2, 4)
                self.unmerge_cells_in_range(sheet, 4, 4, 2, 4)
                self.unmerge_cells_in_range(sheet, 5, 5, 2, 4)
                self.unmerge_cells_in_range(sheet, 6, 6, 2, 4)
                self.unmerge_cells_in_range(sheet, 9, 9, 2, 4)
                self.unmerge_cells_in_range(sheet, 11, 11, 2, 4)
                self.unmerge_cells_in_range(sheet, 12, 12, 2, 4)
                self.unmerge_cells_in_range(sheet, 15, 15, 2, 4)

                self.unmerge_cells_in_range(sheet, 13, 13, 6, 8)
                self.unmerge_cells_in_range(sheet, 14, 14, 6, 8)
                self.unmerge_cells_in_range(sheet, 15, 15, 6, 8)
                self._log_info("开始处理林道模板")
                current_date = datetime.now().strftime("%Y.%m.%d")
                cell = sheet.cell(row=1, column=4)
                cell.value = current_date
                cell.font = Font(name='Arial', size=12)

                box_Reference_id = '' 

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                # 在第一行B列填充编码
                if code:
                    cell = sheet.cell(row=1, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=9)

                    cell = sheet.cell(row=14, column=6)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=9)
                    cell = sheet.cell(row=13, column=6)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=9)


                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    
                    if address_info_detail['amazonReferenceId']:
                        box_Reference_id =address_info_detail['amazonReferenceId']

                        cell = sheet.cell(row=15, column=6)  # B列是第2列
                        cell.value = box_Reference_id
                        cell.font = Font(name='Arial', size=9)

                    try:
                        # 填充收件人信息，这里收件人和
                        
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=4, column=2)  # B2单元格
                            # cell.value = address_info_detail['name']

                            # cell = sheet.cell(row=3, column=2)  # B2单元格
                            # cell.value = address_info_detail['name']

                            # cell = sheet.cell(row=5, column=2)  # B2单元格
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        # 填充地址信息
                        if 'addressLine1' in address_info_detail:
                            cell = sheet.cell(row=6, column=2)  # B3单元格
                            cell.value = address_info_detail['addressLine1']

                        # 城市
                        if 'city' in address_info_detail:
                            cell = sheet.cell(row=9, column=2)  # B4单元格
                            cell.value = address_info_detail['city']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=11, column=2)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=12, column=2)  # B7单元格
                            cell.value = address_info_detail['countryCode']

                        if address_info_detail['type'] == 'amz':
                            if address_parts:
                                # 准备地址部分
                                final_address_parts = address_parts.copy()
                                
                                # 设置第7行B列的地址单元格
                                cell_address = sheet.cell(row=7, column=2)  # 第7行B列
                                
                                if 'warehouseId' in address_info_detail:
                                    warehouse_id = str(address_info_detail['warehouseId'])
                                    
                                    # 如果地址部分中没有warehouse_id，则添加
                                    if warehouse_id not in address_parts:
                                        final_address_parts.insert(0, warehouse_id)
                                    
                                    # 设置第4行B列的warehouse_id
                                    cell_warehouse_4 = sheet.cell(row=4, column=2)  # 第4行B列
                                    cell_warehouse_4.value = warehouse_id 
                                    
                                    # 设置第3行B列的warehouse_id
                                    cell_warehouse_3 = sheet.cell(row=3, column=2)  # 第3行B列
                                    cell_warehouse_3.value = warehouse_id 

                                    # 设置第5行B列的公司名称
                                    cell_company = sheet.cell(row=5, column=2)  # 第5行B列
                                    if warehouse_id not in address_info_detail['name']:
                                        cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                        cell_company.value = cell_result_value
                                    else:
                                        cell_company.value = address_info_detail['name']
                                
                                # 设置完整地址
                                cell_address.value = ', '.join(final_address_parts)
                        else:
                            # 非AMZ类型的地址处理
                            if 'addressLine2' in address_info_detail: 
                                cell_address2 = sheet.cell(row=7, column=2)  # 第7行B列
                                cell_address2.value = address_info_detail['addressLine2']    
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                    
                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=15, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 检查所有产品的电磁属性
                has_electric = False
                has_magnetic = False
                for box in box_data.values():
                    for item in box.items:
                        product_info = self._get_product_info(item.msku, db)
                        if product_info:
                            if product_info.get('electrified', '') == '是':
                                has_electric = True
                            if product_info.get('magnetic', '') == '是':
                                has_magnetic = True
                            if has_electric and has_magnetic:
                                break
                    if has_electric and has_magnetic:
                        break

                # 在表格顶部添加电磁属性标记
                if has_electric:
                    cell = sheet.cell(row=1, column=6)  # F列第1行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)
                
                if has_magnetic:
                    cell = sheet.cell(row=2, column=6)  # F列第2行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)

                # 填充数据
                row_num = 17  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[17].height

                # 遍历每个箱子

                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))

                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)

                        Reference_id = None  # 初始化为None
                        if box_Reference_id:
                            Reference_id = box_Reference_id
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)

                        
                        # 设置单元格值和样式ç
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, code),
                            (3,Reference_id if Reference_id is not None else ""), 
                            (4, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (5, box.length if box.length is not None else ""),  # 长度 (Q列)
                            (6, box.width if box.width is not None else ""),    # 宽度 (R列)
                            (7, box.height if box.height is not None else ""),   # 高度 (S列)
                            (8,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (9, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            # (10, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (10, self._get_display_price(item, product_info)),   # 单价
                            
                            (11,"美元"),  
                            (12, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (13, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (14, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (15, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (16, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (17, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            # (12, product_info.get('link', '') if product_info else ''),
                            (18, ''),  # 图片列 (N列)
                            # (15, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                            # (24,item.SKU)
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                        sheet.column_dimensions['R'].width = row_height
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"R{row_num}" 
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

                self.merge_cells_in_range(sheet, 3, 3, 2, 4)
                self.merge_cells_in_range(sheet, 4, 4, 2, 4)
                self.merge_cells_in_range(sheet, 5, 5, 2, 4)
                self.merge_cells_in_range(sheet, 6, 6, 2, 4)
                self.merge_cells_in_range(sheet, 9, 9, 2, 4)
                self.merge_cells_in_range(sheet, 11, 11, 2, 4)
                self.merge_cells_in_range(sheet, 12, 12, 2, 4)
                self.merge_cells_in_range(sheet, 15, 15, 2, 4)
                self.merge_cells_in_range(sheet, 13, 13, 6, 8)
                self.merge_cells_in_range(sheet, 14, 14, 6, 8)
                self.merge_cells_in_range(sheet, 15, 15, 6, 8)
            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("林道UPS")
    def _fill_lindaoUPS_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充林道UPS模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db: 
            try:
                # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data
                    
                sheet = wb['发票']
                
                # 解除单元格合并
                ranges_to_unmerge = ['A22:D22', 'A23:D23', 'A24:D24', 'A25:D25',
                                'A26:D26', 'A27:D27']
                for range_str in ranges_to_unmerge:
                    try:
                        start_row, end_row, start_col, end_col = self._parse_range(range_str)
                        self.unmerge_cells_in_range(sheet, start_row, end_row, start_col, end_col)
                    except Exception as e:
                        print(f"解除单元格合并时出错 {range_str}: {str(e)}")

                # 保存特定行的高度
                row_23_height = sheet.row_dimensions[23].height if 23 in sheet.row_dimensions else 15
                row_28_height = sheet.row_dimensions[28].height if 28 in sheet.row_dimensions else 15

                # 设置行高
                default_height = sheet.row_dimensions[60].height if 60 in sheet.row_dimensions else 15
                for r in range(19, 22):
                    sheet.row_dimensions[r].height = default_height
                
                # 记录E24单元格的格式
                cell_e24 = sheet.cell(row=24, column=5)
                cell_font_e24 = Font(
                    name=cell_e24.font.name if cell_e24.font.name else 'Arial',
                    size=cell_e24.font.size if cell_e24.font.size else 11,
                    bold=cell_e24.font.bold,
                    italic=cell_e24.font.italic,
                    vertAlign=cell_e24.font.vertAlign,
                    color=cell_e24.font.color
                )
                cell_alignment_e24 = Alignment(
                    horizontal=cell_e24.alignment.horizontal if cell_e24.alignment.horizontal else 'center',
                    vertical=cell_e24.alignment.vertical if cell_e24.alignment.vertical else 'center',
                    text_rotation=cell_e24.alignment.text_rotation,
                    wrap_text=cell_e24.alignment.wrap_text,
                    shrink_to_fit=cell_e24.alignment.shrink_to_fit,
                    indent=cell_e24.alignment.indent
                )

                # 删除指定行
                sheet.delete_rows(22, 29)

                # 记录格式信息
                row_height = sheet.row_dimensions[19].height if 19 in sheet.row_dimensions else 15
                cell_page = sheet.cell(row=19, column=1)
                cell_border = Border(
                    left=cell_page.border.left if cell_page.border.left else Side(style='thin'),
                    right=cell_page.border.right if cell_page.border.right else Side(style='thin'),
                    top=cell_page.border.top if cell_page.border.top else Side(style='thin'),
                    bottom=cell_page.border.bottom if cell_page.border.bottom else Side(style='thin')
                )
                cell_font = Font(
                    name=cell_page.font.name if cell_page.font.name else 'Arial',
                    size=cell_page.font.size if cell_page.font.size else 11,
                    bold=cell_page.font.bold,
                    italic=cell_page.font.italic,
                    vertAlign=cell_page.font.vertAlign
                )
                cell_alignment = Alignment(horizontal='center', vertical='center')
                # 填充数据
                num_row = 19

                
                #由box_data.items()改为
                for box_number, box in processed_data.items():
                    if not box.items:
                        continue
                    
                    for product_info in box.items:
                        # 获取产品信息
                        db_product_info = self._get_product_info(product_info.msku, db)
                        if not db_product_info:
                            continue

                        # 构建产品名称和获取数量、价格
                        name = f"{db_product_info.get('en_name', '')}({db_product_info.get('cn_name', '')})"
                        quantity = product_info.box_quantities.get(box_number, 0)
                        price = db_product_info.get('price', 0)

                        # 累计总数和总金额
                        total_quantity = quantity
                        total_amount = float(price) * quantity

                        # 设置单元格值
                        cell_values = [
                            (1, name), 
                            (2, quantity), 
                            (3, price), 
                            (4, quantity * price),
                            ('CN', 5)
                        ]
                        
                        for value, col in cell_values:
                            cell = sheet.cell(row=num_row, column=col, value=value)
                            cell.font = cell_font
                            cell.alignment = cell_alignment
                            cell.border = cell_border
                        num_row += 1

                # 设置行高和边框
                for row in range(19, num_row + 8):
                    sheet.row_dimensions[row].height = row_height
                    for col in range(1, 6):
                        cell = sheet.cell(row=row, column=col)
                        if not cell.border:
                            cell.border = cell_border

                # 添加底部文本
                declarations = [
                    ('THESE COMMODITIES ARE LICENSED FOR THE UNTIMATE DESTINATION SHOWN.', cell_font),
                    ('以上商品已有到最终目的地的许可。', cell_font),
                    ('', None),
                    ('I DECLARE ALL THE INFORMATION CONTAINED IN THIS INVOICE LIST TO BE TRUE AND CORRECT.',
                    Font(name='Arial', size=9, color='000080')),
                    ('以上申报均属实。', Font(name='宋体', size=11, color='FF0000', bold=True)),
                    ('', None),
                    ('SIGNATURE OF SHIPPER/EXPORTER(TYPE NAME TITLE AND SIGN):    ',
                    Font(name='Arial', size=9, color='000080', bold=True)),
                    ('寄件人/出口商签名(正楷和职位)', Font(name='宋体', size=9, color='000080', bold=True))
                ]

                for i, (text, font) in enumerate(declarations):
                    if text:
                        cell = sheet.cell(row=num_row + i, column=1, value=text)
                        if font:
                            cell.font = font
                        cell.alignment = cell_alignment
                        if i == 1:
                            sheet.row_dimensions[num_row + i].height = row_23_height
                        elif i == 6:
                            sheet.row_dimensions[num_row + i].height = row_28_height

                # 合并单元格
                for row in range(num_row, num_row + 8):
                    try:
                        self.merge_cells_in_range(sheet, row, row, 1, 4)
                    except Exception as e:
                        print(f"合并单元格时出错 row {row}: {str(e)}")

                # 设置右侧文本
                right_text = [
                    (num_row, 5, 'CHECK ONE', cell_font),
                    (num_row + 1, 5, '□ F.O.B', cell_font),
                    (num_row + 2, 5, '', cell_font_e24),
                    (num_row + 6, 4, 'DATE:', cell_font),
                    (num_row + 7, 4, '日期', cell_font)
                ]

                for row, col, text, font in right_text:
                    cell = sheet.cell(row=row, column=col, value=text)
                    cell.font = font
                    cell.alignment = (cell_alignment_e24 if col == 5 and row == num_row + 2 
                                    else cell_alignment)

            except Exception as e:
                print(f"填充林道UPS模板时发生错误: {str(e)}")
                raise

    @template_handler("递信普通")
    def _fill_dixing_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充递信模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:

                # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data
                
                sheet = wb['FBA对应贴标资料']  # 获取模板工作表
                self._log_info("开始处理递信模板")
                current_date = datetime.now().strftime("%Y.%m.%d")

                cell = sheet.cell(row=1, column=4)
                cell.value = current_date
                cell.font = Font(name='Arial', size=12)
                
                # 正确访问shipmentName
                shipment_Name = ''
                if address_info and 'address_info' in address_info and 'shipmentName' in address_info['address_info']:
                    shipment_parts = address_info['address_info']['shipmentName'].split('-')
                    if len(shipment_parts) > 1:
                        shipment_Name = shipment_parts[-2]


                # 记录第3行的格式信息
                row_height = sheet.row_dimensions[3].height

                # 查找并取消合并单元格
                merged_cells = sheet.merged_cells
                cells_to_unmerge = []
                for merged_cell in merged_cells:
                    if merged_cell.min_row >= 3:
                        cells_to_unmerge.append(merged_cell)

                for cell_range in cells_to_unmerge:
                    sheet.unmerge_cells(str(cell_range))

                # 保存最后一行的高度
                last_height = sheet.row_dimensions[21].height

                # 删除原来的内容
                sheet.delete_rows(3, 20)

                address_parts = []
                adress = ''
                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    try:
                        # 填充收件人信息，这里收件人和
                        
                        # 填充地址信息
                       
                        if 'name' in address_info_detail:
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        if address_parts:
                            # cell = sheet.cell(row=3, column=14)  # B3单元格
                            # adress = ', '.join(address_parts)
                            adress = ', \n'.join(address_parts)

                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                # 设置行高
                for r in range(3, 21):
                    sheet.row_dimensions[r].height = sheet.row_dimensions[25].height

                # 初始化统计数据
                total_weight = 0
                total_quantity = 0
                row_num = 3
                ticket = str(code)+"U00000"

                # 设置单元格对齐方式
                center_alignment = Alignment(horizontal='center', vertical='center')

                # 遍历每个箱子

                #由box_data.items()改为processed_data.items()
                for box_number, box in sorted(processed_data.items(), key=lambda x: int(x[0])):
                    # 计算箱子中的产品数量
                    box_products = box.items
                    print(box_products)
                    is_mixed = len(box_products) >= 2
                    # identifier = f"{len(box_data)}-{box_number}{'(混装)' if is_mixed else ''}"
                    # 如果箱号是个位数则补零，否则不补零
                    identifier = f"{shipment_Name}0{box_number}" if int(box_number) < 10 else f"{shipment_Name}{box_number}"

                    merge_start_row = row_num
                    
                    # 计算该箱子的总数量
                    box_total_quantity = sum(
                        getattr(product_info, 'box_quantities', {}).get(box_number, 0)
                        for product_info in box_products
                    )
                    print(f"Box {box_number} total quantity: {box_total_quantity}")

                    # 遍历箱子中的每个产品
                    for index, product_info in enumerate(box_products):
                        if hasattr(product_info, 'msku'):
                            db_product_info = self._get_product_info(product_info.msku, db)
                            if db_product_info:
                                # 更新产品信息
                                for key, value in db_product_info.items():
                                    setattr(product_info, key, value)

                        # 设置单元格值
                        cell_data = [
                            (1, identifier),  # 标识符
                            # (2, f"{getattr(product_info, 'cn_name', '')}\n{getattr(product_info, 'en_name', '')}"),  # 品名
                            # (2, f"{getattr(product_info, 'cn_name', '')} ({getattr(product_info, 'box_quantities', {}).get(box_number, 0)}双) {getattr(product_info, 'en_name', '')}" if getattr(product_info, 'cn_name', '') == '袜子' else f"{getattr(product_info, 'cn_name', '')}\n{getattr(product_info, 'en_name', '')}"),
                            (2, f"{getattr(product_info, 'cn_name', '')} ({self._get_quantity_from_sku(getattr(product_info, 'sku', ''))}双) {getattr(product_info, 'en_name', '')}" if getattr(product_info, 'cn_name', '') == '袜子' else f"{getattr(product_info, 'cn_name', '')}\n{getattr(product_info, 'en_name', '')}"),
                            (3, f"{getattr(product_info, 'material_cn', '')}\n{getattr(product_info, 'material_en', '')}"),  # 材质
                            (4, f"{getattr(product_info, 'usage_en', '')}, {getattr(product_info, 'usage_cn', '')}"),  # 用途
                            (5, 1),  # 箱号
                            (8, box_total_quantity), 
                            (6, getattr(box, 'weight', '')),  # 重量
                            (7, f"{getattr(box, 'length', '')}*{getattr(box, 'width', '')}*{getattr(box, 'height', '')}"),  # 尺寸
                        ]

                        # 处理数量信息
                        quantity = getattr(product_info, 'box_quantities', {}).get(box_number, 0)
                        original_value = getattr(product_info, 'box_original_values', {}).get(box_number, str(quantity))
                        

                        total_quantity = total_quantity + quantity

                        # 只在第一行显示箱子的总数量
                        if quantity:
                            if ' ' in original_value:
                                prefix, number = original_value.split(' ', 1)
                                cell_data.extend([
                                    # (8, number),  # 数量
                                    (9, number),  # 数量（重复）
                                    (10, f"{prefix} {str(quantity)}"),  # 前缀（如 A1）
                                ])
                            else:
                                cell_data.extend([
                                    (9, str(box_total_quantity)),  # 数量（重复）
                                    (10, str(box_total_quantity)),  # 没有前缀
                                ])
                        else:
                            cell_data.extend([
                                (9, ''),  # 数量（重复）
                                (10, ''),  # 没有前缀
                            ])

                        # 设置运单号
                        cell_data.append((13, f"{ticket}{box_number}"))  # 运单号
                        cell_data.append((14, adress))  # 地址信息

                        # 设置单元格值
                        for col, value in cell_data:
                            cell = sheet.cell(row=row_num, column=col, value=value)
                            cell.alignment = center_alignment

                        # 插入产品图片
                        if hasattr(product_info, 'msku') and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"L{row_num}"
                                sheet.row_dimensions[row_num].height = 95
                                # self.insert_product_image(sheet, image_cell, product_info.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, product_info.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1
                        
                    # 合并单元格
                    if row_num > merge_start_row:
                        merge_ranges = [
                            (merge_start_row, 1, row_num - 1, 1),  # 标识符列
                            (merge_start_row, 5, row_num - 1, 5),  # 箱号列
                            (merge_start_row, 6, row_num - 1, 6),  # 重量列
                            (merge_start_row, 7, row_num - 1, 7),  # 尺寸列
                            (merge_start_row, 8, row_num - 1, 8),  # 数量列
                            (merge_start_row, 13, row_num - 1, 13),  # 运单号列
                           
                        ]
                        for start_row, start_col, end_row, end_col in merge_ranges:
                            sheet.merge_cells(
                                start_row=start_row,
                                start_column=start_col,
                                end_row=end_row,
                                end_column=end_col
                            )
                    total_weight += box.weight

                # 合并最后一列
                sheet.merge_cells(start_row=3, start_column=14, end_row=row_num, end_column=14)

                # 删除多余的行
                if row_num < sheet.max_row:
                    sheet.delete_rows(row_num + 1, sheet.max_row - row_num)

                # 设置边框
                cell_border = Border(
                    left=Side(border_style='thin'),
                    right=Side(border_style='thin'),
                    top=Side(border_style='thin'),
                    bottom=Side(border_style='thin')
                )

                # 设置所有单元格的边框和行高
                for row in range(3, row_num + 1):
                    sheet.row_dimensions[row].height = row_height
                    sheet.column_dimensions['L'].width = row_height
                    for col in range(1, 15):
                        cell = sheet.cell(row=row, column=col)
                        cell.border = cell_border

                # 设置最后一行的汇总信息
                sheet.row_dimensions[row_num].height = last_height

                # 设置汇总行的字体和样式
                summary_data = [
                    (1, '汇总', Font(bold=True, name='宋体', size=12)),
                    (5, len(box_data), Font(bold=True, name='微软雅黑', size=11)),
                    (6, total_weight, Font(bold=True, name='微软雅黑', size=9)),
                    (8, total_quantity, Font(bold=True, name='微软雅黑', size=12)),
                    (9, total_quantity, Font(bold=True, name='微软雅黑', size=11))
                ]

                for col, value, font in summary_data:
                    cell = sheet.cell(row=row_num, column=col, value=value)
                    cell.font = font
                    cell.alignment = center_alignment

                print("递信模板填充完成")

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                traceback.print_exc()
                raise
    
    @template_handler("德邦美森限时达")
    def _fill_dbmsxsd_template(self, wb, box_data,code=None, address_info=None, shipment_id=None):
        """
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:
                 # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['FBA专线出货资料模板']  # 获取模板工作表
                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                #先拆分合并的单元格，用于写入
                self.unmerge_cells_in_range(sheet, 4, 4, 3, 7)
                self.unmerge_cells_in_range(sheet, 5, 5, 3, 7)
                self.unmerge_cells_in_range(sheet, 3, 3, 3, 7)


                # 初始化 address_info_detail
                address_info_detail = {}
                if address_info and 'address_info' in address_info:
                    address_info_detail = address_info['address_info'] or {}
                
                if code:
                    cell = sheet.cell(row=4, column=11)  # B列是第2列
                    # cell.value = "FBA 号：" + str(code)
                    # cell.font = Font(name='Arial', size=12,bold=True)
                    cell.value = str(code)

                # 这里名字用物流中心编码
                if 'warehouseId' in address_info_detail:
                    warehouse_id = str(address_info_detail['warehouseId'])
                    cell = sheet.cell(row=4, column=3)  
                    cell.value = str(warehouse_id)

                    cell_company = sheet.cell(row=3, column=3)  # B2单元格
                    if 'warehouseId' not in address_info_detail['name']:
                        cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                        cell_company.value = cell_result_value
                    else:
                        cell_company.value = address_info_detail['name']


                Reference_id = address_info_detail.get('amazonReferenceId', '') if address_info_detail else ''
                if Reference_id:
                    cell = sheet.cell(row=5, column=11)  # B列是第2列
                    cell.value = Reference_id
                    
                # 如果有地址信息，填充到相应的单元格
                if address_info_detail:
                    # 已经初始化过 address_info_detail，这里不需要重复初始化
                    pass
                    try:
                       
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=3, column=3)  
                            # cell.value = address_info_detail['name']
                            # cell = sheet.cell(row=4, column=3)  
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                            
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        if address_info_detail['type'] == 'amz':
                            cell = sheet.cell(row=3, column=3)  
                            cell.value = 'Amazon'
                            if address_parts:
                                # 检查warehouseId并添加到最前面
                                final_address_parts = address_parts.copy()
                                if 'warehouseId' in address_info_detail:
                                    warehouse_id = str(address_info_detail['warehouseId'])
                                    if warehouse_id not in address_parts:
                                        final_address_parts.insert(0, warehouse_id)
                                
                                # cell = sheet.cell(row=2, column=3)  
                                # cell.value = ', '.join(final_address_parts)

                                cell = sheet.cell(row=5, column=3)  
                                cell.value = ', '.join(final_address_parts)

                                
                        else:
                            cell = sheet.cell(row=3, column=3)  
                            cell.value = 'WalMart'
                           
                            if 'logistics_code' in address_info_detail:
                                cell = sheet.cell(row=4,column= 3)
                                cell.value = address_info_detail['logistics_code']
                            if 'addressLine2' in address_info_detail: 
                                cell = sheet.cell(row=2, column=3)  
                                cell.value = address_info_detail['addressLine2']    
                                
                                cell = sheet.cell(row=5, column=3)  
                                cell.value = address_info_detail['addressLine2']    
  
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                # 填充数据
                row_num = 9  
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[9].height

                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")
                    first_row_of_box = row_num  # 记录这个箱子的第一行
                    box_number_str = code + 'U00000' + str(box_number)

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 检查长宽高是否为None，如果是则使用默认值0或跳过计算
                        if box.length is None or box.width is None or box.height is None:
                            print(f"警告：箱子 {box_number} 的尺寸数据不完整：length={box.length}, width={box.width}, height={box.height}")
                            volume = 0  # 设置默认值为0
                        else:
                            volume = box.length * box.width * box.height * 0.000001
                        price = 0
                        total_price = 0

                        # 检查重量是否为None（仅用于警告）
                        if box.weight is None:
                            print(f"警告：箱子 {box_number} 的重量数据为None")
    
                        if product_info is not None:
                            item.product_name = product_info.get('cn_name', item.product_name)
                            print(f"产品信息：{product_info}")
                        else:
                        # 处理未找到产品信息的情况
                            print(f"未找到产品信息，MSKU: {item.msku}")
                            item.product_name = "需要补数据"  # 可以设置一个默认值
                    
                        # 设置单元格值和样式
                        cell_data = [
                            # 基本信息
                            # 产品名称信息
                            (2, box_number_str),
                            (4,f"{product_info.get('en_name', '')}" if product_info else ''),
                            (3,f"{product_info.get('cn_name', '')}" if product_info else ''),
                            (6, product_info.get('hs_code', '') if product_info else ''),                # HS编码
                            # (3, f"{product_info.get('en_name', '')} ({product_info.get('cn_name', '')})" if product_info else ''), 
                            (7, item.box_quantities.get(box_number, 0)),         # 数量
                            # 价格处理：区分合并和非合并商品
                            (8, self._get_display_price(item, product_info)),   # 单价
                            (9, self._get_total_price(item, box_number, product_info)),   # 总价
                            # 产品材料和用途
                            (10, f"{product_info.get('material_en', '')} /{product_info.get('material_cn', '')}" if product_info else ''),            # 中文材料
                            (11, str(product_info.get('usage_en', '') + '/' +
                                   product_info.get('usage_cn', '')) if product_info else ''),            # 用途

                            (12, product_info.get('brand', '') if product_info else ''),
                            (13, product_info.get('model', '') if product_info else ''),
                            (14,1),

                            (15, box.weight if box.weight is not None else ""),  # 重量 
                            (16, box.length if box.length is not None else ""),  # 长度 
                            (17, box.width if box.width is not None else ""),    # 宽度 
                            (18, box.height if box.height is not None else "") ,  # 高度 
                            (19, f"=P{row_num}*Q{row_num}*R{row_num}/1000000*N{row_num}"),  # 体积重公式
                            (20, f"=P{row_num}*Q{row_num}*R{row_num}/6000*N{row_num}"),   # 体积重公式（6000系数）

                            (21,  product_info.get('electrified', '')if product_info else ''),            
                            (5, ''),                                            # 图片占位
                        ]
                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                        
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"E{row_num}"  # 图片列（第14列）
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1
                        
                    # box_info_data = [
                    #     (12, box_number_str),  
                    #     (13, box.weight if box.weight is not None else ""),     # 重量
                    #     (14, box.weight if box.weight is not None else ""),     # 重量
                    #     (15, volume if volume is not None else "")          # 体积
                    # ]

                    # # 设置箱子信息
                    # for column, value in box_info_data:
                    #     cell = sheet.cell(row=first_row_of_box, column=column, value=value)
                    #     cell.font = style_info['font']
                    #     cell.border = style_info['border']
                    #     cell.alignment = style_info['alignment']

                    # # 使用箱子中的产品数量来确定合并范围
                    # if len(box.items) > 1:  # 只有当箱子中有多个产品时才合并
                    #     for column, _ in box_info_data:
                    #         sheet.merge_cells(
                    #             start_row=first_row_of_box,
                    #             start_column=column,
                    #             end_row=first_row_of_box + len(box.items) - 1,
                    #             end_column=column
                    #         )
                    #            # 添加总计行

                total_row = row_num  # 直接使用当前行号，不再加1
                data_start_row = 9  # 数据起始行
                data_end_row = total_row - 1  # 数据结束行
                
                self._set_cell_value(sheet, total_row, 2, "TOTAL:", style_info)
                self._set_cell_value(sheet, total_row, 14, len(box_data), style_info)
                self._set_cell_value(sheet, total_row, 7, f"=SUM(G{data_start_row}:G{data_end_row})", style_info)  # 数量总和
                self._set_cell_value(sheet, total_row, 9, f"=SUM(I{data_start_row}:I{data_end_row})", style_info)  # 总价总和
                self._set_cell_value(sheet, total_row, 15, f"=SUM(O{data_start_row}:O{data_end_row})", style_info)  # 重量总和
                self._set_cell_value(sheet, total_row, 19, f"=SUM(S{data_start_row}:S{data_end_row})", style_info)  # 体积重总和（19列）
                self._set_cell_value(sheet, total_row, 20, f"=SUM(T{data_start_row}:T{data_end_row})", style_info)  # 体积重总和（20列）
                # self._set_cell_value(sheet, total_row, 20, f"=SUM(N{data_start_row}:N{data_end_row})", style_info) 
    
                
                self.merge_cells_in_range(sheet, 4, 4, 3, 7)
                self.merge_cells_in_range(sheet, 5, 5, 3, 5)
                self.merge_cells_in_range(sheet, 3, 3, 3, 5)

                
                thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(border_style='thin'))

                for row_index in range(total_row-1, total_row+1):  # 行索引从1到3（对应A1:C3中的1到3行）
                    for col_index in range(2, 16):  # 列索引从1到3（对应A、B、C三列）
                        cell = sheet.cell(row=row_index, column=col_index)
                        cell.border = thin_border

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("德邦空派")
    def _fill_dbkp_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充德邦空派模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:

                # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['箱单发票']  # 获取模板工作表
                
                # 获取基础单元格样式
                base_cell = sheet.cell(row=9, column=3)
                base_border = Border(
                    left=base_cell.border.left,
                    right=base_cell.border.right,
                    top=base_cell.border.top,
                    bottom=base_cell.border.bottom
                )
                base_font = Font(
                    name=base_cell.font.name,
                    size=base_cell.font.size,
                    bold=base_cell.font.bold,
                    italic=base_cell.font.italic,
                    vertAlign=base_cell.font.vertAlign
                )
                base_alignment = Alignment(horizontal='center', vertical='center')
                style_info = {'font': base_font, 'border': base_border, 'alignment': base_alignment}
                
                # 保存行高信息
                # row_height = sheet.row_dimensions[9].height
                row_height = 60
                row_height_low = 30
                
                # 计算总数据
                total_length = 0
                total_weight = 0
                
                for box_number, box in box_data.items():
                    # 每个箱子中的产品种类数
                    box_product_count = len(box.items)
                    total_length += box_product_count

                    total_weight += box.weight if (hasattr(box, 'weight') and box.weight is not None) else 0
                
                print("需要添加行数为total_length:", total_length)
                # 插入所需行数
                if total_length > 1:
                    sheet.insert_rows(9, total_length)
                
                self.unmerge_cells_in_range(sheet, 4, 4, 3, 7)
                self.unmerge_cells_in_range(sheet, 5, 5, 3, 7)
                self.unmerge_cells_in_range(sheet, 5, 5, 12, 18)
                self.unmerge_cells_in_range(sheet, 4, 4, 12, 13)

                # 解除数据区域内的合并单元格
                print("解除数据区域内的合并单元格...")
                try:
                    data_ranges = []
                    for merged_range in sheet.merged_cells.ranges:
                        range_str = str(merged_range)
                        start_row, end_row, _, _ = self._parse_range(range_str)
                        if start_row >= 9:  # 只解除数据区域的合并单元格
                            data_ranges.append(range_str)
                    
                    for range_str in data_ranges:
                        sheet.unmerge_cells(range_str)
                        print(f"解除合并单元格: {range_str}")
                except Exception as e:
                    print(f"解除合并单元格时发生错误: {str(e)}")
                    traceback.print_exc()

                # 填充日期
                self._set_cell_value(sheet, 4, 12, datetime.now().strftime('%Y-%m-%d'), style_info)
                self.merge_cells_in_range(sheet, 4, 4, 12, 13)        
                
                # if code:
                #     cell = sheet.cell(row=4, column=3)  # B列是第2列
                #     cell.value = code
                #     cell.font = Font(name='Arial', size=9)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    try:
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=5, column=3)  # B2单元格
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        if address_info_detail['type'] == 'amz':
                            if address_parts:
                                # 准备地址部分
                                final_address_parts = address_parts.copy()
                                
                                # 设置第5行L列的地址单元格
                                cell_address = sheet.cell(row=5, column=12)  # 第5行L列
                                
                                if 'warehouseId' in address_info_detail:
                                    warehouse_id = str(address_info_detail['warehouseId'])
                                    
                                    # 如果地址部分中没有warehouse_id，则添加
                                    if warehouse_id not in address_parts:
                                        final_address_parts.insert(0, warehouse_id)

                                    # 设置第6行C列的公司名称
                                    cell_company = sheet.cell(row=6, column=3)  # 第6行C列
                                    if warehouse_id not in address_info_detail['name']:
                                        cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                        cell_company.value = cell_result_value
                                    else:
                                        cell_company.value = address_info_detail['name']

                                # 设置完整地址
                                cell_address.value = ', '.join(final_address_parts)
                        else:
                            # 非AMZ类型的地址处理
                            if 'addressLine2' in address_info_detail: 
                                cell_address2 = sheet.cell(row=5, column=12)  # 第5行L列
                                cell_address2.value = address_info_detail['addressLine2']    
                                  
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                # 填充数据
                row_num = 9
                data_rows = []  # 存储所有产品数据行
                sheet.delete_rows(9)
                
                # 第一步：收集所有产品数据，
                #由box_data.items()改为processed_data.items()
                for box_number, box in sorted(processed_data.items(), key=lambda x: int(x[0])):
                    box_start_row = row_num
                    total_quantity = 0
                    total_price = 0
                    
                    for item in box.items:
                        product_info = self._get_product_info(item.msku, db)
                        # 即使查询不到产品信息，也要生成一行
                        if not product_info:
                            self._log_missing_product(item.msku)
                            # 创建默认的产品信息
                            product_info = {
                                'en_name': f'Product-{item.msku}',
                                'cn_name': '待补充数据',
                                'hs_code': '',
                                'price': 0,
                                'magnetic': '',
                                'brand': '',
                                'model': '',
                                'link': ''
                            }

                        # 定义price变量
                        price_value = product_info.get('price', 0)
                        # 确保price是浮点数
                        price = float(price_value) if price_value else 0

                        # # 序号
                        # self._set_cell_value(sheet, row_num, 1, row_num - 8, style_info)
                        
                        # FBA号
                        fba_number = code+'U00000'+str(box_number)
                        self._set_cell_value(sheet, row_num, 2, fba_number, style_info)
                        # 箱号
                        self._set_cell_value(sheet, row_num, 3, 1, style_info)
                        # 产品名称
                        name = f"{product_info.get('en_name', '')}({product_info.get('cn_name', '')})"
                        self._set_cell_value(sheet, row_num, 4, name, style_info)
                        # HS编码
                        self._set_cell_value(sheet, row_num, 5, product_info.get('hs_code', ''), style_info)
                        # 数量
                        quantity = item.quantity if (hasattr(item, 'quantity') and item.quantity is not None) else 0
                        self._set_cell_value(sheet, row_num, 6, quantity, style_info)

                        self._set_cell_value(sheet,row_num,16,'',style_info)  
                        self._set_cell_value(sheet, row_num, 7, f"${price}", style_info)
                        
                        # 总价
                        total = round(float(quantity) * price, 2)
                        self._set_cell_value(sheet, row_num, 8, f"${total}", style_info)
                        
                        # 重量相关
                        # 修改后
                        box_weight = box.weight if (hasattr(box, 'weight') and box.weight is not None) else 0
                        for col in range(9, 12):
                            self._set_cell_value(sheet, row_num, col, box_weight, style_info)
                        
                        # 箱子尺寸
                        if (hasattr(box, 'length') and hasattr(box, 'width') and hasattr(box, 'height') and 
                            box.length is not None and box.width is not None and box.height is not None):
                            self._set_cell_value(sheet, row_num, 12, box.length, style_info)
                            self._set_cell_value(sheet, row_num, 13, box.width, style_info)
                            self._set_cell_value(sheet, row_num, 14, box.height, style_info)
                            volume = box.length * box.width * box.height * 0.000001
                            self._set_cell_value(sheet, row_num, 15, volume, style_info)
                        else:
                            self._set_cell_value(sheet, row_num, 12, 0, style_info)
                            self._set_cell_value(sheet, row_num, 13, 0, style_info)
                            self._set_cell_value(sheet, row_num, 14, 0, style_info)
                            self._set_cell_value(sheet, row_num, 15, 0, style_info)
                        
                        # 磁性
                        self._set_cell_value(sheet, row_num, 17, product_info.get('magnetic', ''), style_info)

                        # self._set_cell_value(sheet, row_num, 18, product_info.get('link', ''), style_info)

                        # 插入产品图片
                        # self.insert_product_image(sheet, f'P{row_num}', item.msku, self.image_folder)
                        self.insert_original_product_image(sheet, f'P{row_num}', item.msku, self.image_folder)
                        
                        total_quantity += quantity
                        total_price += total
                        
                        # 保存产品信息用于后续创建申报要素表格
                        data_rows.append({
                            'product_info': product_info,
                            'row': row_num
                        })
                        
                        row_num += 1
                    
                    # 合并相同箱号的单元格
                    if row_num - box_start_row > 1:
                        for col in [2, 3, 12, 13, 14, 15]:
                            try:
                                merge_range = f"{get_column_letter(col)}{box_start_row}:{get_column_letter(col)}{row_num-1}"
                                sheet.merge_cells(merge_range)
                            except Exception as e:
                                print(f"合并单元格失败 {merge_range}: {str(e)}")

                # 设置原始数据区域的行高
                for row in range(9, row_num):
                    sheet.row_dimensions[row].height = row_height
                    sheet.column_dimensions['P'].width = row_height
                
                # 空一行开始添加申报要素表格
                row_num += 2
                
                # 第二步：为每个产品创建申报要素表格
                for data_row in data_rows:
                    product_info = data_row['product_info']
                    
                    # 创建申报要素表格
                    table_height = self._create_declaration_table(sheet, row_num+1, product_info)
                    
                    # 设置表格区域的行高，暂时先不用
                    for i in range(row_num, row_num + table_height):
                        sheet.row_dimensions[i].height = row_height_low
                    
                    # 更新行号（表格高度 + 1行间距）
                    row_num += table_height + 1
                
                # 合并底部单元格
                try:
                    sheet.merge_cells(f"F{row_num+2}:M{row_num+8}")
                    sheet.merge_cells(f"B{row_num+2}:D{row_num+2}")
                    sheet.merge_cells(f"B{row_num+10}:D{row_num+10}")
                    sheet.merge_cells(f"B{row_num+18}:D{row_num+18}")
                    
                    for i in range(3, 9):
                        sheet.merge_cells(f"C{row_num+i}:D{row_num+i}")
                    
                    for i in range(11, 17):
                        sheet.merge_cells(f"C{row_num+i}:D{row_num+i}")

                    for i in range(19, 25):
                        sheet.merge_cells(f"C{row_num+i}:D{row_num+i}")

                except Exception as e:
                    print(f"合并底部单元格时发生错误: {str(e)}")

            except Exception as e:
                print(f"填充德邦空派模板时发生错误: {str(e)}")
                traceback.print_exc()
                raise ProcessingError(f"填充德邦空派模板失败: {str(e)}")

    @template_handler("罗马尼亚鹏城")
    def _fill_ropc_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """填充罗马尼亚鹏程模板"""
        with self.db_connector as db:
            try:
                sheet = wb['发票箱单']  # 获取模板工作表
                # 拆开合并单元格
                # self.unmerge_cells_in_range(sheet, 15, 15, 2, 4)
                # self.unmerge_cells_in_range(sheet, 1, 1, 6, 8)
                # self.unmerge_cells_in_range(sheet, 2, 2, 6, 8)
                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=1, column=6)  #填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 检查所有产品的电磁属性
                has_electric = False
                has_magnetic = False
                for box in box_data.values():
                    for item in box.items:
                        product_info = self._get_product_info(item.msku, db)
                        if product_info:
                            if product_info.get('electrified', '') == '是':
                                has_electric = True
                            if product_info.get('magnetic', '') == '是':
                                has_magnetic = True
                            if has_electric and has_magnetic:
                                break
                    if has_electric and has_magnetic:
                        break

                # 在表格顶部添加电磁属性标记
                if has_electric:
                    cell = sheet.cell(row=2, column=6)  # F列第1行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)
                
                if has_magnetic:
                    cell = sheet.cell(row=3, column=6)  # F列第2行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)

                # 填充数据
                row_num = 20  # 从第20行开始填充
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[20].height
                # sheet.column_dimensions['Q'].width = row_height/4
                
                # 遍历每个箱子

                sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 设置单元格值和样式ç
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (10, box.weight if box.weight is not None else ""),  # 重量 
                            (11, box.weight if box.weight is not None else ""),  # 重量 
                            (12, box.length if box.length is not None else ""),  # 长度 
                            (13, box.width if box.width is not None else ""),    # 宽度 
                            (14, box.height if box.height is not None else "") ,  # 高度                  
                            (3,product_info.get('en_name', '') if product_info else ''),  
                            (2, product_info.get('cn_name', '') if product_info else ''),  
                            (6, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (5, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (8, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (4, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (9, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (17, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (18, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (16, product_info.get('link', '') if product_info else ''),
                            (15, ''),  
                            (22,item.sku),
                            # (15, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                          
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)
                        sheet.row_dimensions[row_num].height = row_height
                       
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"O{row_num}"
                                # self.insert_product_image(sheet, image_cellO item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

                # self.merge_cells_in_range(sheet, 15, 15, 2, 4)
                # self.merge_cells_in_range(sheet, 1, 1, 6, 8)
                # self.merge_cells_in_range(sheet, 2, 2, 6, 8)

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise
    
    @template_handler("运达通")
    def _fill_ydt_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充运达通模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:

                # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['模板']  # 获取模板工作表
                self.unmerge_cells_in_range(sheet, 4, 4, 2, 4)
                self.unmerge_cells_in_range(sheet, 6, 6, 2, 4)
                self.unmerge_cells_in_range(sheet, 8, 8, 2, 4)
                self.unmerge_cells_in_range(sheet, 7, 7, 2, 4)
                self.unmerge_cells_in_range(sheet, 12, 12, 2, 4)
                self.unmerge_cells_in_range(sheet, 13, 13, 2, 4)
                self.unmerge_cells_in_range(sheet, 14, 14, 2, 4)
                self.unmerge_cells_in_range(sheet, 10, 10, 2, 4)
                self.unmerge_cells_in_range(sheet, 11, 11, 2, 4)
                print("开始写入模版信息")

                cell = sheet.cell(row=9, column=2)  # B15单元格
                cell.value = ''


                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                # 在第一行B列填充编码
                if code:
                    cell = sheet.cell(row=4, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=9)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    print("地址信息:", address_info_detail)
                    try:
                        # 填充收件人信息
                       
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=7, column=2)  # B2单元格
                            # cell.value = address_info_detail['name'] 

                            # cell = sheet.cell(row=3, column=2)  # B2单元格
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])

                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        # 填充名字信息
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=7, column=2)  # B2单元格
                            cell.value = address_info_detail['name']
                            
                        # 填充地址信息
                        if 'addressLine1' in address_info_detail:
                            cell = sheet.cell(row=8, column=2)  # B3单元格
                            cell.value = address_info_detail['addressLine1']

                        # 城市
                        if 'city' in address_info_detail:
                            cell = sheet.cell(row=10, column=2)  # B4单元格
                            cell.value = address_info_detail['city']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=12, column=2)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=13, column=2)  # B7单元格
                            cell.value = address_info_detail['countryCode']

                        if 'stateOrProvinceCode' in address_info_detail:
                            cell = sheet.cell(row=11, column=2)  # B8单元格
                            cell.value = address_info_detail['stateOrProvinceCode']

                        if address_parts:
                            # 准备地址部分
                            final_address_parts = address_parts.copy()
                            
                            # 设置第6行B列的地址单元格
                            cell_address = sheet.cell(row=6, column=2)  # 第6行B列
                            
                            if 'warehouseId' in address_info_detail:
                                warehouse_id = str(address_info_detail['warehouseId'])
                                
                                # 设置第7行B列的公司名称
                                cell_company = sheet.cell(row=7, column=2)  # 第7行B列
                                if warehouse_id not in address_info_detail['name']:
                                    cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                    cell_company.value = cell_result_value
                                else:
                                    cell_company.value = address_info_detail['name']

                                # 如果地址部分中没有warehouse_id，则添加
                                if warehouse_id not in address_parts:
                                    final_address_parts.insert(0, warehouse_id)
                            
                            # 设置完整地址
                            cell_address.value = ', '.join(final_address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=14, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 填充数据
                row_num = 16  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始

                # 遍历每个箱子

                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))

                 # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                row_height = sheet.row_dimensions[16].height
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        box_number_str = code+'U00000'+str(box_number)
                        # 设置单元格值和样式
                        cell_data = [
                            (1, box_number_str),                    # 货箱编号 (A列)
                            (15, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (3,product_info.get('en_name', '') if product_info else ''),  
                            (2, product_info.get('cn_name', '') if product_info else ''),  
                            (8, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (7, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (10, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (4, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (5, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (9, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            # (12, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (12, ''),
                            (13, product_info.get('link', '') if product_info else ''),
                            (14, ''),  
                            # (15, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                            (16, box.length if box.length is not None else ""),  # 长度 (Q列)
                            (17, box.width if box.width is not None else ""),    # 宽度 (R列)
                            (18, box.height if box.height is not None else ""),  # 高度 (S列)
                            (6,'')
                            # (6,"")
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)
                        sheet.row_dimensions[row_num].height = row_height

                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"K{row_num}"  # 图片列
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1
                
                self.merge_cells_in_range(sheet, 10, 10, 2, 4)
                self.merge_cells_in_range(sheet, 11, 11, 2, 4)
                self.merge_cells_in_range(sheet, 12, 12, 2, 4)
                self.merge_cells_in_range(sheet, 13, 13, 2, 4)
                self.merge_cells_in_range(sheet, 14, 14, 2, 4)
                self.merge_cells_in_range(sheet, 4, 4, 2, 4)
                self.merge_cells_in_range(sheet, 6, 6, 2, 4)
                self.merge_cells_in_range(sheet, 8, 8, 2, 4)
                self.merge_cells_in_range(sheet, 7, 7, 2, 4)

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise
    
    @template_handler("一八供应链")
    def _fill_yiba_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充运达通模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:

                 # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['专线箱单 ']  # 获取模板工作表
               
                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=12),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                # 在第一行B列填充编码
                if code:
                    cell = sheet.cell(row=1, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=12,color="FF0000")

                    cell = sheet.cell(row=2, column=4)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=12)

                    cell = sheet.cell(row=7, column=4)  # B列是第2列
                    cell.value = "亚马逊地址"
                    cell.font = Font(name='Arial', size=12)


                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    print("地址信息:", address_info_detail)
                    try:
                       
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=7, column=2)  # B2单元格
                            # cell.value = address_info_detail['name'] 

                            # cell = sheet.cell(row=3, column=2)  # B2单元格
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])

                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        # # 填充名字信息
                        # if 'name' in address_info_detail:
                        #     cell = sheet.cell(row=8, column=2)  # B2单元格
                        #     cell.value = address_info_detail['name']
                            
                        # 填充地址信息
                        if 'addressLine1' in address_info_detail:
                            cell = sheet.cell(row=15, column=2)  # B3单元格
                            cell.value = address_info_detail['addressLine1']

                        # 城市
                        if 'city' in address_info_detail:
                            cell = sheet.cell(row=14, column=2)  # B4单元格
                            cell.value = address_info_detail['city']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=12, column=2)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=13, column=2)  # B7单元格
                            cell.value = address_info_detail['countryCode']

                        #省份
                        if 'stateOrProvinceCode' in address_info_detail:
                            cell = sheet.cell(row=13, column=2)  # B8单元格
                            cell.value = address_info_detail['stateOrProvinceCode']

                        if address_parts:
                            # 准备地址部分
                            final_address_parts = address_parts.copy()
                            
                            if 'warehouseId' in address_info_detail:
                                warehouse_id = str(address_info_detail['warehouseId'])
                                
                                # 设置第8行B列的warehouse_id
                                cell_warehouse_8 = sheet.cell(row=8, column=2)  # 第8行B列
                                cell_warehouse_8.value = warehouse_id

                                # 设置第10行B列的warehouse_id
                                cell_warehouse_10 = sheet.cell(row=10, column=2)  # 第10行B列
                                cell_warehouse_10.value = warehouse_id

                                # 设置公司名称
                                if 'cell_company' in locals() and warehouse_id not in address_info_detail['name']:
                                    cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                    cell_company.value = cell_result_value
                                elif 'cell_company' in locals():
                                    cell_company.value = address_info_detail['name']

                                # 如果地址部分中没有warehouse_id，则添加
                                if warehouse_id not in address_parts:
                                    final_address_parts.insert(0, warehouse_id)
                            
                            # 设置完整地址
                            # 注意：原代码中没有明确指定地址单元格，使用第15行B列作为默认
                            cell_address = sheet.cell(row=15, column=2)  # 第15行B列
                            cell_address.value = ', '.join(final_address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=17, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)

                    cell.font = Font(name='Arial', size=20, bold=True, color="FF0000") 
                except Exception as e:
                    
                    print(f"填充箱数时发生错误: {str(e)}")

                # 填充数据
                row_num = 19  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始

                # 遍历每个箱子

                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                 # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))

                row_height = sheet.row_dimensions[19].height
                Reference_id = ''
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 根据箱子编号进行进位，格式为U000001, U000002等
                        box_number_str = code + "U" + str(int(box_number)).zfill(6)
                        Reference_id = address_info['address_info'].get('amazonReferenceId','') if address_info and address_info.get('address_info') else ''
                        # 设置单元格值和样式
                        cell_data = [
                            (1, box_number_str),                    # 货箱编号 (A列)
                            (2,Reference_id),
                            (9,item.box_quantities.get(box_number, 0)),
                            (10,'1'),
                            (13, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (4,product_info.get('en_name', '') if product_info else ''),  
                            (3, product_info.get('cn_name', '') if product_info else ''),  
                            (12, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (19, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (20,''),
                            (5, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (7, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (8,''),
                            (14,''),
                            (6, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (15, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (17, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (18, product_info.get('link', '') if product_info else ''),
                            (11,'美金'),
                            
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)
                        sheet.row_dimensions[row_num].height = row_height

                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"N{row_num}"  # 图片列
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise
    
    @template_handler("德邦澳大利亚")
    def _fill_debang_australia_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        填充叮铛卡航限时达模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:

                # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                self.unmerge_cells_in_range(sheet, 3, 3, 2, 4)
                self.unmerge_cells_in_range(sheet, 3, 3, 8, 11)
                self.unmerge_cells_in_range(sheet, 4, 4, 8, 15)
                self.unmerge_cells_in_range(sheet, 5, 5, 8, 11)
                self.unmerge_cells_in_range(sheet, 5, 5, 13, 15)
                self.unmerge_cells_in_range(sheet, 6, 6, 8, 15)
                self.unmerge_cells_in_range(sheet, 7, 7, 8, 11)
                self.unmerge_cells_in_range(sheet, 7, 7, 13, 15)
                
                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=11),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                # 填充编码
                if code:
                    cell = sheet.cell(row=3, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=12)

                total_boxes = len(box_data.keys())
                if total_boxes:
                    cell = sheet.cell(row=3, column=6)  #填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=11)


                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    try:                     
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=4, column=2)  # B2单元格
                            # cell.value = address_info_detail['name']

                            # cell = sheet.cell(row=3, column=2)  # B2单元格
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        # # 填充地址信息
                        # if 'addressLine1' in address_info_detail:
                        #     cell = sheet.cell(row=6, column=8)  # B3单元格
                        #     cell.value = address_info_detail['addressLine1']

                        # 城市
                        # if 'city' in address_info_detail:
                        #     cell = sheet.cell(row=7, column=8)  # B4单元格
                        #     cell.value = address_info_detail['city']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=5, column=13)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=5, column=8)  # B7单元格
                            cell.value = address_info_detail['countryCode']
                        
                        #收件人
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=3, column=8)  # B7单元格
                            cell.value = address_info_detail['name']  
                            
                            cell = sheet.cell(row=4, column=8)  # B7单元格
                            cell.value = address_info_detail['name']  

                        if address_parts:
                            # 准备地址部分
                            final_address_parts = address_parts.copy()
                            
                            if 'warehouseId' in address_info_detail:
                                warehouse_id = str(address_info_detail['warehouseId'])
                                
                                # 设置第3行H列的warehouse_id
                                cell_warehouse = sheet.cell(row=3, column=8)  # 第3行H列
                                cell_warehouse.value = warehouse_id

                                # 设置第4行H列的公司名称
                                cell_company = sheet.cell(row=4, column=8)  # 第4行H列
                                if warehouse_id not in address_info_detail['name']:
                                    cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                    cell_company.value = cell_result_value
                                else:
                                    cell_company.value = address_info_detail['name']
                                
                                # 如果地址部分中没有warehouse_id，则添加
                                if warehouse_id not in address_parts:
                                    final_address_parts.insert(0, warehouse_id)
                            
                            # 设置第6行H列的完整地址
                            cell_address = sheet.cell(row=6, column=8)  # 第6行H列
                            cell_address.value = ', '.join(final_address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=16, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")


                # 填充数据
                row_num = 12  # 从第12行开始填充
                index = 1    # 添加序号计数器，从1开始
                Reference_id = ''
                
                # 记录第12行的行高，用于后续行的统一设置
                template_row_height = sheet.row_dimensions[12].height
             

                # 遍历每个箱子

                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))

                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))

                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        

                        Reference_id = address_info['address_info'].get('amazonReferenceId','') if address_info and address_info.get('address_info') else ''
                        box_number_str = code+'U00000'+str(box_number)
                        # 设置单元格值和样式ç
                        cell_data = [
                            (1, box_number_str),                    # 货箱编号 (A列)
                            (2, Reference_id if Reference_id is not None else ""),  
                            (3,f"{box.length}*{box.width}*{box.height}"),  #箱子的尺寸
                            (4, box_number), 
                            (5,box.weight),
                            (6, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (7, product_info.get('cn_name', '') if product_info else ''),  
                            (8,product_info.get('en_name', '') if product_info else ''),  #
                            (9, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (10, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (11, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (13, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (14, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (12, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (15, ''),  # 图片列 (N列)
                            (16, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                            # (17, box.length if box.length is not None else ""),  # 长度 (Q列)
                            # (18, box.width if box.width is not None else ""),    # 宽度 (R列)
                            # (19, box.height if box.height is not None else "")   # 高度 (S列)
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)
                            sheet.row_dimensions[row_num].height = template_row_height

                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"O{row_num}"  # 图片列（第14列）
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

                self.merge_cells_in_range(sheet, 3, 3, 2, 4)
                self.merge_cells_in_range(sheet, 3, 3, 8, 11)
                self.merge_cells_in_range(sheet, 4, 4, 8, 15)
                self.merge_cells_in_range(sheet, 5, 5, 8, 11)
                self.merge_cells_in_range(sheet, 5, 5, 13, 15)
                self.merge_cells_in_range(sheet, 6, 6, 8, 15)
                self.merge_cells_in_range(sheet, 7, 7, 8, 11)
                self.merge_cells_in_range(sheet, 7, 7, 13, 15)
            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise
    
    @template_handler("德邦英欧")
    def _fill_debang_EU_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:
                sheet = wb['2025.3.12 最新发票模板']  # 获取模板工作表

                  # 可选：应用产品合并
                merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                
                # 使用合并后的数据替代原始数据
                processed_data = merged_box_data  # 或者直接使用 box_data 跳过合并
            
                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                # 填充编码
                if code:
                    cell = sheet.cell(row=13, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=12)

                total_boxes = len(box_data.keys())
                if total_boxes:
                    cell = sheet.cell(row=14, column=2)  #填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=11)

                Reference_id = address_info['address_info'].get('amazonReferenceId','') if address_info and address_info.get('address_info') else ''
                if Reference_id:
                    cell = sheet.cell(row=12, column=6)  #填充箱数
                    cell.value = Reference_id
                    cell.font = Font(name='Arial', size=11)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info'] if address_info and address_info.get('address_info') else {}
                    try:                     
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=4, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                            # cell = sheet.cell(row=3, column=2)  # B2单元格
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        # # 填充地址信息
                        # if 'addressLine1' in address_info_detail:
                        #     cell = sheet.cell(row=6, column=8)  # B3单元格
                        #     cell.value = address_info_detail['addressLine1']

                        # 城市
                        if 'city' in address_info_detail:
                            cell = sheet.cell(row=7, column=2)  # B4单元格
                            cell.value = address_info_detail['city']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=9, column=2)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=10, column=2)  # B7单元格
                            cell.value = address_info_detail['countryCode']
                        
                        #收件人
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=3, column=2)  # B7单元格
                            cell.value = address_info_detail['name']  
                            
        
                        if address_parts:
                            # 准备地址部分
                            final_address_parts = address_parts.copy()
                            
                            if 'warehouseId' in address_info_detail:
                                warehouse_id = str(address_info_detail['warehouseId'])
                                
                                # 设置第3行B列的warehouse_id
                                cell_warehouse = sheet.cell(row=3, column=2)  # 第3行B列
                                cell_warehouse.value = warehouse_id

                                # 设置第4行B列的公司名称
                                cell_company = sheet.cell(row=4, column=2)  # 第4行B列
                                if warehouse_id not in address_info_detail['name']:
                                    cell_result_value = str(address_info_detail['name'])+','+str(warehouse_id)
                                    cell_company.value = cell_result_value
                                else:
                                    cell_company.value = address_info_detail['name']
                                
                                # 如果地址部分中没有warehouse_id，则添加
                                if warehouse_id not in address_parts:
                                    final_address_parts.insert(0, warehouse_id)
                            
                            # 设置第5行B列的完整地址
                            cell_address = sheet.cell(row=5, column=2)  # 第5行B列
                            cell_address.value = ', '.join(final_address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                # # 检查所有产品的电磁属性
                # has_electric = False
                # has_magnetic = False
                # for box in box_data.values():
                #     for item in box.items:
                #         product_info = self._get_product_info(item.msku, db)
                #         if product_info:
                #             if product_info.get('electrified', '') == '是':
                #                 has_electric = True
                #             if product_info.get('magnetic', '') == '是':
                #                 has_magnetic = True
                #             if has_electric and has_magnetic:
                #                 break
                #     if has_electric and has_magnetic:
                #         break

                # # 在表格顶部添加电磁属性标记
                # if has_electric:
                #     cell = sheet.cell(row=1, column=6)  # F列第1行
                #     cell.value = "是"
                #     cell.font = Font(name='Arial', size=9)
                
                # if has_magnetic:
                #     cell = sheet.cell(row=2, column=6)  # F列第2行
                #     cell.value = "是"
                #     cell.font = Font(name='Arial', size=9)

                # 填充数据
                row_num = 18  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[18].height
                sheet.column_dimensions['Q'].width = row_height/4
                
                # 遍历每个箱子

                # sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            self._log_missing_product(item.msku)
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)

                        box_number_str = code+'U00000'+str(box_number)
                        # 设置单元格值和样式ç
                        cell_data = [
                            (1, box_number_str),                    # 货箱编号 (A列)
                            (7, box.weight if box.weight is not None else ""),  # 重量 
                            (8, box.length if box.length is not None else ""),  # 长度 
                            (9, box.width if box.width is not None else ""),    # 宽度 
                            (10, box.height if box.height is not None else "") ,  # 高度 
                            # (19, item.sku),

                            (2,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (3, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)

                            (5,"个"),
                            

                            # (5, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (4, item.box_quantities.get(box_number, 0)),  # 数量 (F列)

                            (13, product_info.get('material_cn', '') if product_info else ''),  # 材料 (D列) 
                            (11, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (15, product_info.get('usage_en', '')if product_info else ''),    # 用途 (H列)
                            (12, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (14, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (16, product_info.get('link', '') if product_info else ''),
                            (17, ''),  # 图片列 (N列)

                            (6, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                          
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)
                        sheet.row_dimensions[row_num].height = row_height
                       
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"Q{row_num}"
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1


            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise
    
    @template_handler("递信日本空派")
    def _fill_dx_japenDP_template(self, wb, box_data,code=None, address_info=None, shipment_id=None):
        """
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:
                 # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['发票']  # 获取模板工作表
                print("开始写入模版信息")
                warehouse_id = ''

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                #先拆分合并的单元格，用于写入
                self.unmerge_cells_in_range(sheet, 6, 6, 7, 10)
                self.unmerge_cells_in_range(sheet, 8, 8, 7, 10)
                self.unmerge_cells_in_range(sheet, 9, 9, 7, 10)


                # 初始化 address_info_detail
                address_info_detail = {}
                if address_info and 'address_info' in address_info:
                    address_info_detail = address_info['address_info'] or {}
                
    
                # 如果有地址信息，填充到相应的单元格
                if address_info_detail:
                    # 已经初始化过 address_info_detail，这里不需要重复初始化
                    pass
                    try:
                       
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=3, column=3)  
                            # cell.value = address_info_detail['name']
                            # cell = sheet.cell(row=4, column=3)  
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                            
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        if address_info_detail['type'] == 'amz':
                            # cell = sheet.cell(row=3, column=3)  
                            # cell.value = 'Amazon'
                            if address_parts:
                                # 检查warehouseId并添加到最前面
                                final_address_parts = address_parts.copy()
                                if 'warehouseId' in address_info_detail:
                                    warehouse_id = str(address_info_detail['warehouseId'])
                                    if warehouse_id not in address_parts:
                                        final_address_parts.insert(0, warehouse_id)
                                    cell = sheet.cell(row=8, column=7)  
                                    cell.value = ', '.join(final_address_parts)
                                

                                    cell_company = sheet.cell(row=9, column=7)  # B2单元格
                                    if warehouse_id not in address_info_detail['name']:
                                        cell_result_value = str("ADDRESS*:") + str(address_info_detail['name'])+','+str(warehouse_id)
                                        cell_company.value = cell_result_value
                                    else:
                                        cell_company.value = str("ADDRESS*:") + address_info_detail['name']
                                    
                                # cell.value = ', '.join(final_address_parts)

                        # else:
                        #     cell = sheet.cell(row=3, column=3)  
                        #     cell.value = 'WalMart'
                           
                        #     if 'logistics_code' in address_info_detail:
                        #         cell = sheet.cell(row=4,column= 3)
                        #         cell.value = address_info_detail['logistics_code']
                        #     if 'addressLine2' in address_info_detail: 
                        #         cell = sheet.cell(row=2, column=3)  
                        #         cell.value = address_info_detail['addressLine2']    
                                
                        #         cell = sheet.cell(row=8, column=10)  
                        #         cell.value = address_info_detail['addressLine2']    
  
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                # 填充数据
                row_num = 13  
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[13].height

                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                total_box_number = len(sorted_boxes)
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")
                    first_row_of_box = row_num  # 记录这个箱子的第一行
                    # box_number_str = code + 'U00000' + str(box_number)
                    box_number_str = code + 'U' + f"{box_number:06d}"  

                    box_items_count = len(box.items)
                    start_row = row_num

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 检查长宽高是否为None，如果是则使用默认值0或跳过计算
                        if box.length is None or box.width is None or box.height is None:
                            print(f"警告：箱子 {box_number} 的尺寸数据不完整：length={box.length}, width={box.width}, height={box.height}")
                            volume = 0  # 设置默认值为0
                        else:
                            volume = box.length * box.width * box.height * 0.000001
                        price = 0
                        total_price = 0

                        # 检查重量是否为None（仅用于警告）
                        if box.weight is None:
                            print(f"警告：箱子 {box_number} 的重量数据为None")
    
                        if product_info is not None:
                            item.product_name = product_info.get('cn_name', item.product_name)
                            print(f"产品信息：{product_info}")
                        else:
                        # 处理未找到产品信息的情况
                            print(f"未找到产品信息，MSKU: {item.msku}")
                            item.product_name = "需要补数据"  # 可以设置一个默认值
                    
                        # 设置单元格值和样式
                        cell_data = [
                            # 基本信息
                            # 产品名称信息
                            # (2, box_number_str),
                            (1, f"F00001\n{warehouse_id}\n{total_box_number}/{box_number}"),
                            (2, '1'),
                            (3, f"{product_info.get('cn_name', '')} ({self._get_quantity_from_sku(item.sku)}双) {product_info.get('en_name', '')}" if product_info and product_info.get('cn_name', '') == '袜子' else f"{product_info.get('cn_name', '') if product_info else ''}\n{product_info.get('en_name', '') if product_info else ''}"),
                            (4,f"{product_info.get('en_name', '')}" if product_info else ''),
                            (8, product_info.get('hs_code', '') if product_info else ''),                # HS编码
                            # (3, f"{product_info.get('en_name', '')} ({product_info.get('cn_name', '')})" if product_info else ''), 
                            (9, item.box_quantities.get(box_number, 0)),         # 数量
                            # 价格处理：区分合并和非合并商品
                            (10, self._get_display_price(item, product_info)),   # 单价
                            (11, self._get_total_price(item, box_number, product_info)),   # 总价
                            # 产品材料和用途
                            (7, f"{product_info.get('material_en', '')}\n{product_info.get('material_cn', '')}" if product_info else ''),            # 中文材料
                            (6, str(product_info.get('usage_en', '') + ',' +
                                   product_info.get('usage_cn', '')) if product_info else ''),            # 用途

                            (4, product_info.get('brand', '') if product_info else ''),
                            (5, product_info.get('model', '') if product_info else ''),
                            (12,''),
                            (13,product_info.get('link', '') if product_info else ''),
                            (14,  product_info.get('electrified', '')if product_info else ''), 
                            (15,'否'),
                            (16,box_number_str),
                            (17, item.fnsku if item.fnsku else ''),
                            
                            
                        ]
                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                        
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"L{row_num}"  # 图片列（第14列）
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        # 每处理完一个产品，行号递增
                        row_num += 1
                    
                    # 处理完这个箱子的所有产品后，如果有多个产品需要合并单元格
                    if box_items_count > 1:
                        merge_columns = [1, 2, 16]  # 需要合并的列：序号列、数量列、箱号列
                        end_row = row_num - 1  # 结束行是当前行的前一行
                        
                        for col in merge_columns:
                            try:
                                print(f"合并箱子 {box_number} 的第 {col} 列，从第 {start_row} 行到第 {end_row} 行")
                                self.merge_cells_in_range(sheet, start_row, end_row, col, col)
                            except Exception as e:
                                print(f"合并第 {col} 列时发生错误: {str(e)}")
                        
                    # box_info_data = [
                    #     (12, box_number_str),  
                    #     (13, box.weight if box.weight is not None else ""),     # 重量
                    #     (14, box.weight if box.weight is not None else ""),     # 重量
                    #     (15, volume if volume is not None else "")          # 体积
                    # ]

                    # # 设置箱子信息
                    # for column, value in box_info_data:
                    #     cell = sheet.cell(row=first_row_of_box, column=column, value=value)
                    #     cell.font = style_info['font']
                    #     cell.border = style_info['border']
                    #     cell.alignment = style_info['alignment']

                    # # 使用箱子中的产品数量来确定合并范围
                    # if len(box.items) > 1:  # 只有当箱子中有多个产品时才合并
                    #     for column, _ in box_info_data:
                    #         sheet.merge_cells(
                    #             start_row=first_row_of_box,
                    #             start_column=column,
                    #             end_row=first_row_of_box + len(box.items) - 1,
                    #             end_column=column
                    #         )
                    #            # 添加总计行

                total_row = row_num  # 直接使用当前行号，不再加1
                data_start_row = 13  # 数据起始行
                data_end_row = total_row - 1  # 数据结束行
                
                self._set_cell_value(sheet, total_row, 1, "总申报合计*", style_info)
                self._set_cell_value(sheet, total_row, 9, f"=SUM(I{data_start_row}:I{data_end_row})", style_info)  # 数量总和
                self._set_cell_value(sheet, total_row, 11, f"=SUM(K{data_start_row}:K{data_end_row})", style_info)  # 总价总和
                
                # self._set_cell_value(sheet, total_row, 20, f"=SUM(N{data_start_row}:N{data_end_row})", style_info) 
    
                
                self.merge_cells_in_range(sheet, 6, 6, 7, 10)
                self.merge_cells_in_range(sheet, 8, 5, 7, 10)
                self.merge_cells_in_range(sheet, 9, 3, 7, 10)

                
                thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(border_style='thin'))

                for row_index in range(total_row-1, total_row+1):  # 行索引从1到3（对应A1:C3中的1到3行）
                    for col_index in range(1, 17):  # 列索引从1到3（对应A、B、C三列）
                        cell = sheet.cell(row=row_index, column=col_index)
                        cell.border = thin_border

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("日本宏川贴标")
    def _fill_dx_japenHC_template(self, wb, box_data,code=None, address_info=None, shipment_id=None):
        """
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        """
        with self.db_connector as db:
            try:
                 # 根据国家判断是否启用产品合并（只有美国才启用）
                if self.should_enable_merge(address_info):
                    merged_box_data = self.merge_items_by_product_name(box_data, debug=True)
                    processed_data = merged_box_data
                else:
                    processed_data = box_data

                sheet = wb['日本空运模板']  # 获取模板工作表
                print("开始写入模版信息")
                warehouse_id = ''

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                #先拆分合并的单元格，用于写入
                # self.unmerge_cells_in_range(sheet, 6, 6, 7, 10)

                self.unmerge_cells_in_range(sheet, 8, 8, 10, 18)
                self.unmerge_cells_in_range(sheet, 9, 11, 10, 18)


                # 初始化 address_info_detail
                address_info_detail = {}
                if address_info and 'address_info' in address_info:
                    address_info_detail = address_info['address_info'] or {}
                
    
                # 如果有地址信息，填充到相应的单元格
                if address_info_detail:
                    # 已经初始化过 address_info_detail，这里不需要重复初始化
                    pass
                    try:
                       
                        # 填充地址信息
                        address_parts = []
                        if 'name' in address_info_detail:
                            # cell = sheet.cell(row=3, column=3)  
                            # cell.value = address_info_detail['name']
                            # cell = sheet.cell(row=4, column=3)  
                            # cell.value = address_info_detail['name']
                            address_parts.append(address_info_detail['name'])
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        
                        if ('addressLine2' in address_info_detail and 
                            address_info_detail['addressLine2'] is not None and 
                            str(address_info_detail['addressLine2']).strip() != '' and 
                            str(address_info_detail['addressLine2']).strip().lower() != 'null'): 
                            address_parts.append(address_info_detail['addressLine2'])
                            
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        if address_info_detail['type'] == 'amz':
                            # cell = sheet.cell(row=3, column=3)  
                            # cell.value = 'Amazon'
                            if address_parts:
                                # 检查warehouseId并添加到最前面
                                final_address_parts = address_parts.copy()
                                if 'warehouseId' in address_info_detail:
                                    warehouse_id = str(address_info_detail['warehouseId'])
                                    if warehouse_id not in address_parts:
                                        final_address_parts.insert(0, warehouse_id)
                                    cell = sheet.cell(row=9, column=10)  
                                    cell.value = ', '.join(final_address_parts)
                                
                                    # cell_company = sheet.cell(row=9, column=7)  # B2单元格
                                    # if warehouse_id not in address_info_detail['name']:
                                    #     cell_result_value = str("ADDRESS*:") + str(address_info_detail['name'])+','+str(warehouse_id)
                                    #     cell_company.value = cell_result_value
                                    # else:
                                    #     cell_company.value = str("ADDRESS*:") + address_info_detail['name']
                                      
  
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                # 填充数据
                row_num = 23  
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[23].height

                # 将processed_data按箱号排序（使用合并后的数据）
                sorted_boxes = sorted(processed_data.items(), key=lambda x: int(x[0]))
                total_box_number = len(sorted_boxes)
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    self._log_debug(f"处理箱子 {box_number}")
                    first_row_of_box = row_num  # 记录这个箱子的第一行
                    # box_number_str = code + 'U00000' + str(box_number)
                    box_number_str = code + 'U' + f"{box_number:06d}"  

                    box_items_count = len(box.items)
                    start_row = row_num

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 检查长宽高是否为None，如果是则使用默认值0或跳过计算
                        if box.length is None or box.width is None or box.height is None:
                            print(f"警告：箱子 {box_number} 的尺寸数据不完整：length={box.length}, width={box.width}, height={box.height}")
                            volume = 0  # 设置默认值为0
                        else:
                            volume = box.length * box.width * box.height * 0.000001
                        price = 0
                        total_price = 0

                        # 检查重量是否为None（仅用于警告）
                        if box.weight is None:
                            print(f"警告：箱子 {box_number} 的重量数据为None")
    
                        if product_info is not None:
                            item.product_name = product_info.get('cn_name', item.product_name)
                            print(f"产品信息：{product_info}")
                        else:
                        # 处理未找到产品信息的情况
                            print(f"未找到产品信息，MSKU: {item.msku}")
                            item.product_name = "需要补数据"  # 可以设置一个默认值
                    
                        # 设置单元格值和样式
                        cell_data = [
                            # 基本信息
                            # 产品名称信息
                            # (2, box_number_str),
                            (2, box_number_str),
                            (3, box.weight if box.weight is not None else ""),
                            (4, max(0, float(box.weight) - 1.1) if box.weight is not None else ""),
                            (5,''),
                            (6, self._extract_asin_from_link(product_info.get('link', '')) if product_info else ''),
                            (7, item.fnsku if item.fnsku else ''),
                            (8,f"{product_info.get('en_name', '')}" if product_info else ''),
                            (9, f"{product_info.get('cn_name', '')}" if product_info else ''),
                            (10,f"{product_info.get('material_cn', '')}" if product_info else ''),
                            (11, f"{product_info.get('material_en', '')}" if product_info else ''),
                            (12,f"{product_info.get('usage_cn', '')}" if product_info else ''),
                            (13, product_info.get('brand', '') if product_info else ''),
                            (14, product_info.get('model', '') if product_info else ''),
                            (15, item.box_quantities.get(box_number, 0)),         # 数量
                            (16, product_info.get('hs_code', '') if product_info else ''),                # HS编码
                            # (3, f"{product_info.get('en_name', '')} ({product_info.get('cn_name', '')})" if product_info else ''), 
                            # 价格处理：区分合并和非合并商品
                            (17, self._get_display_price(item, product_info)),   # 单价
                            # (18, self._get_total_price(item, box_number, product_info)),   # 总价
                            (18, ''), 
                            (19,product_info.get('link', '') if product_info else ''),
                
                        ]
                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                           # 设置总价公式 = O列*Q列
                        self._set_cell_value(sheet, row_num, 18, f"=O{row_num}*Q{row_num}", style_info)
                        
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"E{row_num}"  # 图片列（第14列）
                                # self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                                self.insert_original_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        # 每处理完一个产品，行号递增
                        row_num += 1
                    
                    # 处理完这个箱子的所有产品后，如果有多个产品需要合并单元格
                    if box_items_count > 1:
                        merge_columns = [2,3,4]  # 需要合并的列：序号列、数量列、箱号列
                        end_row = row_num - 1  # 结束行是当前行的前一行
                        
                        for col in merge_columns:
                            try:
                                print(f"合并箱子 {box_number} 的第 {col} 列，从第 {start_row} 行到第 {end_row} 行")
                                self.merge_cells_in_range(sheet, start_row, end_row, col, col)
                            except Exception as e:
                                print(f"合并第 {col} 列时发生错误: {str(e)}")
                        
                    # box_info_data = [
                    #     (12, box_number_str),  
                    #     (13, box.weight if box.weight is not None else ""),     # 重量
                    #     (14, box.weight if box.weight is not None else ""),     # 重量
                    #     (15, volume if volume is not None else "")          # 体积
                    # ]

                    # # 设置箱子信息
                    # for column, value in box_info_data:
                    #     cell = sheet.cell(row=first_row_of_box, column=column, value=value)
                    #     cell.font = style_info['font']
                    #     cell.border = style_info['border']
                    #     cell.alignment = style_info['alignment']

                    # # 使用箱子中的产品数量来确定合并范围
                    # if len(box.items) > 1:  # 只有当箱子中有多个产品时才合并
                    #     for column, _ in box_info_data:
                    #         sheet.merge_cells(
                    #             start_row=first_row_of_box,
                    #             start_column=column,
                    #             end_row=first_row_of_box + len(box.items) - 1,
                    #             end_column=column
                    #         )
                    #            # 添加总计行

                total_row = row_num  # 直接使用当前行号，不再加1
                data_start_row = 23  # 数据起始行
                data_end_row = total_row - 1  # 数据结束行
                
                
                # self._set_cell_value(sheet, total_row, 1, "总申报合计*", style_info)
                # self._set_cell_value(sheet, total_row, 9, f"=SUM(I{data_start_row}:I{data_end_row})", style_info)  # 数量总和
                # self._set_cell_value(sheet, total_row, 11, f"=SUM(K{data_start_row}:K{data_end_row})", style_info)  # 总价总和
                
                
                cell = sheet.cell(row=total_row+2, column=2)
                cell.value = "I DECLARE ALL THE INFORMATION CONTAINED INT THE INVOICE TO BE TRUE "
                cell = sheet.cell(row=total_row+3, column=2)
                cell.value = "AND CORRECT."
                cell = sheet.cell(row=total_row+4, column=2)
                cell.value = "我声明发票上所填信息均真实准确"
                cell =  sheet.cell(row=total_row+4, column=16)
                cell.value = 'Payment Method'

                cell =  sheet.cell(row=total_row+4, column=18)
                cell.value = 'Check one'
                cell =  sheet.cell(row=total_row+6, column=17)
                cell.value = 'Others'
                cell =  sheet.cell(row=total_row+6, column=18)
                cell.value = 'FOB'

                self._set_cell_value(sheet, total_row+3, 16, f"=SUM(R{data_start_row}:R{data_end_row})", style_info)  # 总价总和

                cell = sheet.cell(row=total_row+2, column=18)
                small_font = Font(size=8)
                cell.font = small_font
                cell.value = f'发票登记\n总价值'


                cell = sheet.cell(row=total_row+2, column=16)
                cell.value = 'JPY'
                # 设置字体为红色加粗
                
                red_bold_font = Font(bold=True, color='FF0000', size=12)
                cell.font = red_bold_font

                cell = sheet.cell(row=total_row+5, column=2)
                cell.value = "__________________________________________________________________________________________________________________________"

                cell = sheet.cell(row=total_row+6, column=2)
                cell.value = "SIGNATURE OF SHIPPER/EXPORTER"
                cell = sheet.cell(row=total_row+7, column=2)
                cell.value = "发货人签名"
                cell = sheet.cell(row=total_row+8, column=2)
                cell.value = "NAME(PLEASE PRINT)"
                cell = sheet.cell(row=total_row+9, column=2)
                cell.value = "姓名(请用标准字体填写)"
                # self._set_cell_value(sheet, total_row, 20, f"=SUM(N{data_start_row}:N{data_end_row})", style_info) 

                cell = sheet.cell(row=total_row+9, column=16)
                cell.value = '日期'

                cell = sheet.cell(row=total_row+8, column=16)
                cell.value = 'DATE'

                self.merge_cells_in_range(sheet, 8, 8, 10, 18)
                self.merge_cells_in_range(sheet, 9, 11, 10, 18)
                self.merge_cells_in_range(sheet, total_row+9, total_row+9, 16, 18)
                self.merge_cells_in_range(sheet, total_row+3, total_row+3, 16, 18)

                thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(border_style='thin'))

                for row_index in range(total_row-1, total_row+1):  # 
                    for col_index in range(2, 17):  # 列索引从1到3（对应A、B、C三列）
                        cell = sheet.cell(row=row_index, column=col_index)
                        cell.border = thin_border
                
                # 为total_row到total_row+6行的左侧单元格添加细线边框
                left_border = Border(left=Side(style='thin'))
                for row_index in range(total_row, total_row+7):  # total_row到total_row+6
                    cell = sheet.cell(row=row_index, column=16)  # 左侧单元格（第2列）
                    current_border = cell.border
                    if current_border:
                        # 如果已有边框，保留其他边框样式，只修改左边框
                        new_border = Border(
                            left=Side(style='thin'),
                            right=current_border.right,
                            top=current_border.top,
                            bottom=current_border.bottom
                        )
                        cell.border = new_border
                    else:
                        # 如果没有边框，只添加左边框
                        cell.border = left_border

                for row_index in range(total_row+4, total_row+7):  # total_row到total_row+6
                    cell = sheet.cell(row=row_index, column=17)  # 左侧单元格（第2列）
                    current_border = cell.border
                    if current_border:
                        # 如果已有边框，保留其他边框样式，只修改左边框
                        new_border = Border(
                            right=Side(style='thin'),
                            left=current_border.left,
                            top=current_border.top,
                            bottom=current_border.bottom
                        )
                        cell.border = new_border
                    else:
                        # 如果没有边框，只添加左边框
                        cell.border = left_border


                right_border = Border(right=Side(style='thin'))
                for row_index in range(total_row, total_row+7):  # total_row到total_row+6
                    cell = sheet.cell(row=row_index, column=18)  # 左侧单元格（第2列）
                    current_border = cell.border
                    if current_border:
                        # 如果已有边框，保留其他边框样式，只修改左边框
                        new_border = Border(
                            right=Side(style='thin'),
                            left=current_border.left,
                            top=current_border.top,
                            bottom=current_border.bottom
                        )
                        cell.border = new_border
                    else:
                        # 如果没有边框，只添加左边框
                        cell.border = left_border


                cell = sheet.cell(row=total_row+2, column=18)
                # 设置四边都有边框
                all_borders = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                cell.border = all_borders
                

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    def _fill_default_template(self, wb, box_data, code=None, address_info=None, shipment_id=None):
        """默认的模板处理方法"""
        raise ProcessingError("未找到匹配的模板处理方法，请确保模板文件名包含正确的关键字")

    def generate_invoice(self, template_path, box_data, code=None, address_info=None, shipment_id=None):
        """
        生成发票
        :param template_path: 模板文件路径
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :param shipment_id: Shipment ID（可选）
        :return: 生成的发票文件路径
        """
        try:
            # 清空之前的缓存和记录
            self.missing_products.clear()
            self.product_cache.clear()  # 清空之前的产品缓存
            self.image_cache.clear()  # 清空之前的图片缓存
            
            print(f"开始处理模板文件: {template_path}")
            if not os.path.exists(template_path):
                raise ProcessingError(f"模板文件不存在: {template_path}")
            
            # 预加载产品信息和图片信息
            all_mskus = self._collect_all_mskus(box_data)
            if all_mskus:
                self._preload_product_info(all_mskus)
                self._preload_image_info(all_mskus)

            # 获取当前时间戳
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            
            # 检查产品的电磁属性
            has_electric = False
            has_magnetic = False
            for box in box_data.values():
                for item in box.items:
                    product_info = self._get_product_info(item.msku, self.db_connector)
                    if product_info:
                        if product_info.get('electrified', '') == '是':
                            has_electric = True
                        if product_info.get('magnetic', '') == '是':
                            has_magnetic = True
                        if has_electric and has_magnetic:
                            break
                if has_electric and has_magnetic:
                    break
            
            # 构建文件名后缀
            suffix = ""
            if has_electric:
                suffix += "_带电"
            if has_magnetic:
                suffix += "_带磁"

            # 确保输出目录存在
            if not os.path.exists(self.output_folder):
                os.makedirs(self.output_folder)

            # 构建输出文件路径
            if address_info:
                try:
                    shipment_name = address_info['address_info']['shipmentName'] if address_info and address_info.get('address_info') else ''
                    # 使用正则表达式提取数据
                    time, logistics, number = self.extract_data(shipment_name)
                    
                    if logistics is not None:
                        logistics = logistics.replace('-', '')  # 去掉破折号
                    else:
                        raise ValueError('Logistics cannot be None')

                    # 获取code
                    code_suffix = f"{code}" if code else ""

                    country_name = address_info["seller_info"]["country_name"] if address_info and address_info.get("seller_info") else ''
                    output_filename = f'{code_suffix}-{time}-{logistics}票-{number}件-{country_name}-发票装箱单.xlsx'
                    # 替换任何可能导致路径问题的字符
                    output_filename = "".join(c for c in output_filename if c not in r'<>:"/\|?*')
                    output_path = os.path.join(self.output_folder, output_filename)

                except (TypeError, ValueError) as e:
                    # 处理异常的情况
                    print(f"Error occurred: {e}. Using default filename.")
                    output_filename = f"{timestamp}{suffix}.xlsx"
                    output_path = os.path.join(self.output_folder, output_filename)
            else:
                output_filename = f"{timestamp}{suffix}.xlsx"
                output_path = os.path.join(self.output_folder, output_filename)

            # 使用openpyxl加载模板
            print(f"正在加载模板文件...")
            wb = load_workbook(template_path)
            print(f"成功加载模板文件，工作表: {wb.sheetnames}")

            # 获取对应的模板处理方法
            template_handler = self._get_template_handler(template_path)
            if template_handler is None:
                raise ProcessingError(f"未找到对应的模板处理方法: {template_path}")

            # 处理模板
            template_handler(wb, box_data, code, address_info, shipment_id)

            # 注册webp MIME类型，防止出现KeyError: '.webp'错误
            import mimetypes
            mimetypes.add_type('image/webp', '.webp')
            
            # 保存文件
            wb.save(output_path)
            print(f"发票已生成: {output_path}")
            
            # 打印缓存统计和缺失产品信息汇总
            self._print_cache_statistics()
            self._print_missing_summary()

            return output_path

        except Exception as e:
            error_msg = f"生成发票时发生错误: {str(e)}"
            print(error_msg)
            traceback.print_exc()
            raise ProcessingError(error_msg)

    def _get_template_handler(self, template_path):
        """根据模板文件名选择对应的处理方法"""
        try:
            template_name = os.path.basename(template_path).lower()

            
            for keyword, handler in self._template_handlers.items():

                keyword_lower = keyword.lower()
                if keyword_lower in template_name:
                    print(f"找到匹配的处理器: {handler.__name__}")
                    return handler.__get__(self, type(self))
            print(f"未找到匹配的处理器，可用的关键字: {list(self._template_handlers.keys())}")
            return self._fill_default_template
        except Exception as e:
            print(f"模板处理器匹配过程中出错: {str(e)}")
            return self._fill_default_template

    def _get_product_info(self, msku, db=None):
        """
        从缓存或MongoDB获取产品信息
        :param msku: 产品的MSKU
        :param db: 数据库连接（可选，缓存优先）
        :return: 包含产品信息的字典
        """
        try:
            # 优先从缓存获取
            if self.cache_enabled and msku in self.product_cache:
                self._log_debug(f"从缓存获取产品信息: {msku}")
                return self.product_cache[msku]
            
            # 缓存中没有，从数据库获取
            self._log_debug(f"从数据库获取产品信息: {msku}")
            
            if db is None:
                # 如果没有传入db连接，创建新的连接
                with self.db_connector as database:
                    return self._get_product_info_from_db(msku, database)
            else:
                return self._get_product_info_from_db(msku, db)
                
        except Exception as e:
            print(f"Error fetching product info for MSKU {msku}: {str(e)}")
            return None
    
    def _get_product_info_from_db(self, msku, db):
        """
        从数据库获取产品信息的具体实现
        :param msku: 产品的MSKU
        :param db: 数据库连接
        :return: 包含产品信息的字典
        """
        try:
            # 检查db是否为None
            if db is None:
                print(f"警告: 数据库连接对象为None，无法获取产品信息: {msku}")
                return None

            # 使用传入的db连接
            collection = db['msku_info']
            product = collection.find_one({'msku': msku})
            
            if product:
                product_info = {
                    'cn_name': product.get('productNameZh', ''),
                    'en_name': product.get('productNameEn', ''),
                    'en_usage': product.get('useEn', ''),
                    'ch_usage': product.get('useZh', ''),
                    'material_en': product.get('materialEn', ''),
                    'material_cn': product.get('materialZh', ''),
                    'hs_code': product.get('HS', ''),
                    'usage_en': product.get('useEn', ''),
                    'usage_cn': product.get('useZh', ''),
                    'brand': product.get('brand', ''),
                    'model': product.get('model', ''),
                    'link': product.get('productLink', ''),
                    'price': product.get('askprice', ''),
                    'electrified': product.get('electrified', ''),
                    'magnetic': product.get('magnetic', ''),
                    'weight': product.get('weight', ''),
                }
                
                # 如果缓存开启，将结果添加到缓存
                if self.cache_enabled:
                    self.product_cache[msku] = product_info
                    self._log_debug(f"产品信息已添加到缓存: {msku}")
                
                return product_info
            return None
        except Exception as e:
            print(f"Error fetching product info from DB for MSKU {msku}: {str(e)}")
            return None

    def _set_cell_value(self, sheet, row, column, value, style_info):
        """
        设置单元格的值和样式，支持公式
        :param sheet: 工作表对象
        :param row: 行号
        :param column: 列号
        :param value: 单元格值（可以是普通值或公式）
        :param style_info: 样式信息
        """
        cell = sheet.cell(row=row, column=column)
        
        # 如果值是字符串且以=开头，则作为公式处理
        if isinstance(value, str) and value.startswith('='):
            cell.value = value  # openpyxl会自动识别为公式
        else:
            cell.value = value
            
        # 应用样式
        if style_info:
            cell.font = style_info['font']
            cell.border = style_info['border']
            cell.alignment = style_info['alignment']

    def insert_centered_image(self, worksheet, cell_address, image_path, fixed_width=None, fixed_height=None):
        """
        在指定的单元格中插入居中的图片
        :param worksheet: 工作表对象
        :param cell_address: 单元格地址（例如'A1'）
        :param image_path: 图片文件路径
        :param fixed_width: 固定宽度（可选）
        :param fixed_height: 固定高度（可选）
        :return: 是否成功插入图片
        """
        try:
            # 读取图片
            img = PILImage.open(image_path)
            
            # 获取单元格的宽度和高度（以像素为单位）
            column_width = worksheet.column_dimensions[cell_address[0]].width
            row_height = worksheet.row_dimensions[int(cell_address[1:])].height
            
            # 如果没有指定宽度和高度，使用单元格的大小
            if fixed_width is None:
                fixed_width = column_width * 7  # 转换为像素
            if fixed_height is None:
                fixed_height = row_height * 1.5  # 转换为像素
                
            # 获取原始图片尺寸
            original_width, original_height = img.size
            
            # 计算缩放比例
            width_ratio = fixed_width / original_width
            height_ratio = fixed_height / original_height
            scale = min(width_ratio, height_ratio)
            
            # 计算新的尺寸
            new_width = int(original_width * scale)
            new_height = int(original_height * scale)
            
            # 调整图片大小
            img = img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
            
            # 将图片保存到BytesIO对象
            img_byte_arr = BytesIO()
            img.save(img_byte_arr, format=img.format if img.format else 'PNG')
            img_byte_arr.seek(0)
            
            # 创建Excel图片对象
            xl_img = XLImage(img_byte_arr)
            xl_img.anchor = cell_address
            worksheet.add_image(xl_img)
            
            return True
        except Exception as e:
            print(f"插入图片时发生错误: {str(e)}")
            return False

    def insert_product_image(self, worksheet, cell_address, msku, image_folder, fixed_width=None, fixed_height=None):
        """
        在Excel工作表中插入产品图片
        :param worksheet: openpyxl工作表对象
        :param cell_address: 单元格地址
        :param msku: 产品MSKU
        :param image_folder: 图片文件夹路径
        :param fixed_width: 固定宽度（可选）
        :param fixed_height: 固定高度（可选）
        """
        try:
            # 构建图片文件路径
            image_path_jpg = os.path.join(image_folder, f"{msku}.jpg")
            image_path_png = os.path.join(image_folder, f"{msku}.png")
            print(f"尝试加载图片: {image_path_jpg}") 
            
            # 检查JPEG图片文件是否存在
            if os.path.exists(image_path_jpg):
                return self.insert_centered_image(worksheet, cell_address, image_path_jpg, fixed_width, fixed_height)
            elif os.path.exists(image_path_png):
                print(f"尝试加载PNG图片: {image_path_png}")
                return self.insert_centered_image(worksheet, cell_address, image_path_png, fixed_width, fixed_height)
            else:
                print(f"图片文件不存在: {image_path_jpg} 和 {image_path_png}")
                return False
        except Exception as e:
            print(f"处理产品图片时发生错误: {str(e)}")
            return False

    def insert_original_image(self, worksheet, cell_address, image_path):
        """
        在指定的单元格中插入原始图片，不进行压缩处理
        :param worksheet: 工作表对象
        :param cell_address: 单元格地址（例如'A1'）
        :param image_path: 图片文件路径
        :return: 是否成功插入图片
        """
        try:
            # 检查是否为webp格式
            if image_path.lower().endswith('.webp'):
                print(f"检测到webp格式图片，将转换为PNG: {image_path}")
                try:
                    # 尝试使用webp库处理
                    from webp import WebPHandler
                    # 将WebP转换为PNG并保存到BytesIO
                    img_byte_arr = BytesIO()
                    with open(image_path, 'rb') as webp_file:
                        webp_data = webp_file.read()
                        webp_handler = WebPHandler(webp_data)
                        rgb_data = webp_handler.decode_rgb()
                        # 创建PIL图像
                        width, height = webp_handler.get_info()['width'], webp_handler.get_info()['height']
                        img = PILImage.frombytes('RGB', (width, height), rgb_data)
                        img.save(img_byte_arr, format='PNG')
                        img_byte_arr.seek(0)
                except (ImportError, Exception) as e:
                    print(f"使用webp库处理失败: {str(e)}，尝试使用PIL直接处理")
                    # 退回到使用PIL直接处理
                    img = PILImage.open(image_path)
                    # 将图片转换为RGB模式
                    if img.mode in ('RGBA', 'LA'):
                        img_format = 'PNG'  # 保留透明度
                    else:
                        img = img.convert('RGB')
                        img_format = 'PNG'
                    # 将图片保存到BytesIO对象
                    img_byte_arr = BytesIO()
                    img.save(img_byte_arr, format=img_format)
                    img_byte_arr.seek(0)
            else:
                # 非webp格式处理
                img = PILImage.open(image_path)
                img_format = img.format if img.format else 'PNG'
                img_byte_arr = BytesIO()
                img.save(img_byte_arr, format=img_format)
                img_byte_arr.seek(0)
        
            # 创建Excel图片对象
            xl_img = XLImage(img_byte_arr)
            xl_img.anchor = cell_address
            worksheet.add_image(xl_img)
        
            print(f"成功插入图片: {image_path}")
            return True
        except Exception as e:
            print(f"插入原始图片时发生错误: {str(e)}")
            traceback.print_exc()
            return False

    def insert_original_product_image(self, worksheet, cell_address, msku, image_folder):
        """
        在Excel工作表中插入原始产品图片，支持缓存优化
        :param worksheet: openpyxl工作表对象
        :param cell_address: 单元格地址
        :param msku: 产品MSKU
        :param image_folder: 图片文件夹路径
        """
        try:
            # 优先从缓存获取图片路径
            if self.image_cache_enabled and msku in self.image_cache:
                cached_path = self.image_cache[msku]
                if cached_path:
                    self._log_debug(f"从缓存获取图片路径: {msku} -> {cached_path}")
                    return self.insert_original_image(worksheet, cell_address, cached_path)
                else:
                    self._log_debug(f"缓存显示图片不存在: {msku}")
                    return False
            
            # 缓存未命中，执行原有逻辑
            self._log_debug(f"缓存未命中，检查图片文件: {msku}")
            
            # 构建图片文件路径
            image_path_jpg = os.path.join(image_folder, f"{msku}.jpg")
            image_path_png = os.path.join(image_folder, f"{msku}.png")
            print(f"尝试加载原始图片: {image_path_jpg}") 
            
            # 检查JPEG图片文件是否存在
            if os.path.exists(image_path_jpg):
                # 更新缓存
                if self.image_cache_enabled:
                    self.image_cache[msku] = image_path_jpg
                return self.insert_original_image(worksheet, cell_address, image_path_jpg)
            elif os.path.exists(image_path_png):
                print(f"尝试加载PNG原始图片: {image_path_png}")
                # 更新缓存
                if self.image_cache_enabled:
                    self.image_cache[msku] = image_path_png
                return self.insert_original_image(worksheet, cell_address, image_path_png)
            else:
                print(f"图片文件不存在: {image_path_jpg} 和 {image_path_png}")
                # 更新缓存为None
                if self.image_cache_enabled:
                    self.image_cache[msku] = None
                return False
        except Exception as e:
            print(f"处理原始产品图片时发生错误: {str(e)}")
            return False


    def extract_data(self, ticket_str):
        # 支持两种格式的shipmentName解析
        # 格式1: XX-XX-YYYY.MM.DD-物流公司-数字/数字 (新格式，带前缀)
        # 格式2: YYYY.MM.DD-物流公司-数字/数字 (原格式)
        
        # 先尝试新格式：XX-XX-YYYY.MM.DD-物流公司-数字/数字
        # 匹配两个任意前缀段，然后是日期，物流公司，最后是数字/数字
        new_pattern = r'[^-]+-[^-]+-(?P<time>\d{4}\.\d{1,2}\.\d{1,2})-(?P<logistics>[^-]+(?:-[^-]+)*)-(?P<number>\d+)/(\d+)'
        match = re.search(new_pattern, ticket_str)
        if match:
            print(f"使用新格式解析shipmentName: {ticket_str}")
            return match.group('time'), match.group('logistics').replace('-', ''), match.group(4)
        
        # 再尝试原格式：YYYY.MM.DD-物流公司-数字/数字
        old_pattern = r'(?P<time>\d{4}\.\d{2}\.\d{2})-(?P<logistics>[^-]+(?:-[^-]+)*)-(?P<number>\d+)/(\d+)'
        match = re.search(old_pattern, ticket_str)
        if match:
            print(f"使用原格式解析shipmentName: {ticket_str}")
            return match.group('time'), match.group('logistics').replace('-', ''), match.group(4)
        
        # 都不匹配
        print(f"无法解析shipmentName格式: {ticket_str}")
        return None, None, None

    def unmerge_cells_in_range(self, sheet, start_row, end_row, start_col, end_col):
        """
        解除指定范围内的所有合并单元格。

        :param sheet: 要操作的工作表对象
        :param start_row: 起始行
        :param end_row: 结束行
        :param start_col: 起始列
        :param end_col: 结束列
        """
        print("正在解除合并单元格...")
        merged_cells = list(sheet.merged_cells.ranges)
        cells_to_unmerge = []
        for merged_cell in merged_cells:
            min_row, min_col, max_row, max_col = merged_cell.bounds
            # 检查合并单元格是否在指定范围内
            if (min_row >= start_row and max_row <= end_row and
                min_col >= start_col and max_col <= end_col):
                try:
                    sheet.unmerge_cells(str(merged_cell))
                except Exception as e:
                    print(f"解除合并单元格时发生错误: {str(e)}")
        print("合并单元格解除完成")

    def merge_cells_in_range(self, sheet, start_row, end_row, start_col, end_col):
        """
        合并指定区域内的单元格。

        :param sheet: 要操作的工作表对象
        :param start_row: 起始行
        :param end_row: 结束行
        :param start_col: 起始列
        :param end_col: 结束列
        """
        try:
            # 获取合并区域的范围字符串
            merge_range = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
            print(f"正在合并单元格区域: {merge_range}")
            
            # 合并单元格
            sheet.merge_cells(merge_range)
            
            # 设置合并后的单元格样式（可选）
            merged_cell = sheet.cell(row=start_row, column=start_col)
            merged_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            print(f"单元格合并完成: {merge_range}")
        except Exception as e:
            print(f"合并单元格时发生错误: {str(e)}")

    def _parse_range(self, range_str):
        """
        解析Excel单元格范围字符串

        :param range_str: Excel单元格范围字符串（例如'A1:B2'）
        :return: 解析后的范围元组（start_row, end_row, start_col, end_col）
        """
        start_cell, end_cell = range_str.split(':')
        start_row = int(start_cell[1:])
        start_col = ord(start_cell[0]) - 64
        end_row = int(end_cell[1:])
        end_col = ord(end_cell[0]) - 64
        return start_row, end_row, start_col, end_col

    def _create_declaration_table(self, sheet, row, product_info):
        """
        在指定行创建申报要素表格
        
        :param sheet: 工作表对象
        :param row: 起始行
        :param product_info: 产品信息
        :return: 表格占用的行数
        """
        # 创建表头
        cell = sheet.cell(row=row, column=2)  # 从第二列开始
        cell.value = "申报要素（必填）"
        cell.font = Font(name='SimSun', bold=True, size=11)
        cell.fill = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)
        sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)  # 合并第二列和第三列
        
        # 创建表格内容
        rows = [
            ("HS", product_info.get('hs_code', '')),
            ("品名", f"{product_info.get('en_name', '')}({product_info.get('cn_name', '')})"),
            ("材质", product_info.get('material_cn', '')),
            ("用途", product_info.get('usage_cn', '')),
            ("品牌", product_info.get('brand', '')),
            ("型号", product_info.get('model', '')),
        ]
        
        print(product_info.get('usage_cn', ''))
        for row_index, (label, value) in enumerate(rows, 1):
            # 标签列
            cell = sheet.cell(row=row + row_index, column=2)  # 第二列
            cell.value = label
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                              top=Side(style='thin'), bottom=Side(style='thin'))
            
            # 值列
            cell = sheet.cell(row=row + row_index, column=3)  # 第三列
            cell.value = value
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                              top=Side(style='thin'), bottom=Side(style='thin'))
        
        return 7  # 返回表格占用的行数（1行表头 + 6行内容）

    def _get_quantity_from_sku(self, sku):
        """从SKU中提取数量信息
        如果SKU用-分割后的最后一个部分是数字，就使用这个数字；否则使用默认值5
        """
        default_quantity = 5
        if not sku:
            return default_quantity
        
        parts = sku.split('-')
        if not parts:
            return default_quantity
        
        last_part = parts[-1]
        try:
            # 尝试将最后一部分转换为整数
            quantity = int(last_part)
            return quantity
        except (ValueError, TypeError):
            # 如果转换失败，返回默认值
            return default_quantity