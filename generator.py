from numpy import add
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import os
import traceback
import json
from datetime import datetime
from db_connector import MongoDBConnector
from io import BytesIO


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
        
        # 初始化模板处理器字典
        self._template_handlers = {}
        # 注册所有带有_template_keyword属性的方法
        for name in dir(self):
            method = getattr(self, name)
            if hasattr(method, '_template_keyword'):
                self._template_handlers[method._template_keyword] = method

    @template_handler("叮铛卡航限时达")
    def _fill_dingdang_template(self, wb, box_data, code=None, address_info=None):
        """
        填充叮铛卡航限时达模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        """
        with self.db_connector as db:
            try:
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
                    address_info_detail = address_info['address_info']
                    try:
                        # 填充收件人信息，这里收件人和
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=4, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=3, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                        # 填充地址信息
                        address_parts = []
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
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
                            cell = sheet.cell(row=6, column=2)  # B3单元格
                            cell.value = ', '.join(address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=16, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 先解除所有合并的单元格
                print(f"正在解除合并单元格...")
                merged_ranges = list(sheet.merged_cells.ranges)
                for merged_range in merged_ranges:
                    try:
                        sheet.unmerge_cells(str(merged_range))
                    except:
                        pass
                print(f"合并单元格解除完成")

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
                for box_number, box in box_data.items():
                    print(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # print(f"产品信息：{product_info}")
                        price = product_info.get('price', 0)
                        total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                        if product_info:
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 设置单元格值和样式
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (3,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (4, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            (5, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
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
                                self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("顺丰")
    def _fill_ldmsxsd_template(self, wb, box_data, code=None, address_info=None):
        """填充顺丰模板"""
        """
        填充顺丰模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        """
        with self.db_connector as db:
            try:
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


                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info']
                    try:
                        # 填充收件人信息
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=4, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                        # 填充地址信息
                        address_parts = []
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])


                        if address_parts:
                            cell = sheet.cell(row=4, column=2)  # B3单元格
                            cell.value = ', '.join(address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=16, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 先解除所有合并的单元格
                print(f"正在解除合并单元格...")
                merged_ranges = list(sheet.merged_cells.ranges)
                for merged_range in merged_ranges:
                    try:
                        sheet.unmerge_cells(str(merged_range))
                    except:
                        pass
                print(f"合并单元格解除完成")

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


                # 填充数据
                row_num = 12  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始

                # 遍历每个箱子
                for box_number, box in box_data.items():
                    print(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # print(f"产品信息：{product_info}")
                        price = product_info.get('price', 0)
                        total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                        if product_info:
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 设置单元格值和样式
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (4,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (5, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            (5, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (6, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (9, product_info.get('material_en', '') if product_info else ''),  # 材料 (D列) 
                            (8, product_info.get('material_cn', '') if product_info else ''),  # HS编码 (G列)
                            (12, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (10, str(product_info.get('usage_en', '')+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (11, '纸箱'),    # 品牌 (I列)
                            (6, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (7, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (13, item.box_quantities.get(box_number, 0)),

                            (14, ''),  # 图片列 (N列)
                            (16, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
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
                                self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("法国")
    def _fill_france_template(self, wb, box_data, code=None, address_info=None):
        # TODO: 实现法国模板处理逻辑
        pass

    @template_handler("英国")
    def _fill_uk_template(self, wb, box_data, code=None, address_info=None):
        # TODO: 实现英国模板处理逻辑
        pass

    def generate_invoice(self, template_path, box_data,code=None, address_info=None):
        """
        生成发票
        :param template_path: 模板文件路径
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :return: 生成的发票文件路径
        """
        try:
            # 检查模板文件是否存在
            if not os.path.exists(template_path):
                raise ProcessingError(f"模板文件不存在: {template_path}")

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
                    shipment_name = address_info['address_info']['shipmentName']
                    # 使用正则表达式提取数据
                    time, logistics, number = self.extract_data(shipment_name)
                    
                    if logistics is not None:
                        logistics = logistics.replace('-', '')  # 去掉破折号
                    else:
                        raise ValueError('Logistics cannot be None')

                    # 获取code
                    code_suffix = f"-{code}" if code else ""

                    output_filename = f'百泰{code_suffix}-{time}-{logistics}票-{number}件-{address_info["seller_info"]["country_name"]}-发票装箱单.xlsx'
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
            template_handler(wb, box_data, code, address_info)

            # 保存文件
            wb.save(output_path)
            print(f"发票已生成: {output_path}")

            return output_path

        except Exception as e:
            error_msg = f"生成发票时发生错误: {str(e)}"
            print(error_msg)
            traceback.print_exc()
            raise ProcessingError(error_msg)

    def _get_template_handler(self, template_path):
        """根据模板文件名选择对应的处理方法"""
        template_name = os.path.basename(template_path).lower()
        for keyword, handler in self._template_handlers.items():
            if keyword in template_name:
                return handler.__get__(self, type(self))
        return self._fill_default_template

    def _get_product_info(self, msku, db=None):
        """
        从MongoDB获取产品信息
        :param msku: 产品的MSKU
        :param db: 数据库连接（可选）
        :return: 包含产品信息的字典
        """
        try:
            if db is None:
                # 如果没有传入db连接，创建新的连接
                with self.db_connector as db:
                    return self._get_product_info(msku, db)
            
            # 使用传入的db连接
            collection = db['msku_info']
            product = collection.find_one({'msku': msku})
            
            if product:
                return {
                    'cn_name': product.get('productNameZh', ''),
                    'en_name': product.get('productNameEn', ''),
                    'en_usage': product.get('useEn', ''),
                    'ch_usage':product.get('useZh', ''),
                    'material_en': product.get('materialEn', ''),
                    'material_cn': product.get('materialZh', ''),
                    'hs_code': product.get('HS', ''),
                    'usage_en': product.get('useEn', ''),
                    'usage_cn': product.get('useZh', ''),
                    'brand': product.get('brand', ''),
                    'model': product.get('model', ''),
                    'link': product.get('productLink', ''),
                    'price':product.get('askprice', ''),
                    'electrified':product.get('electrified', ''),
                    'magnetic':product.get('magnetic', ''),
                    'weight':product.get('weight', ''),
                }
            return None
        except Exception as e:
            print(f"Error fetching product info for MSKU {msku}: {str(e)}")
            return None

    def _set_cell_value(self, sheet, row, column, value, style_info):
        """
        设置单元格的值和样式
        :param sheet: 工作表对象
        :param row: 行号
        :param column: 列号
        :param value: 单元格值
        :param style_info: 样式信息
        """
        cell = sheet.cell(row=row, column=column)
        cell.value = value
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

    def extract_data(self, ticket_str):
        import re
        # 定义正则表达式，修改为提取斜杠后的数字
        pattern = r'(?P<time>\d{4}\.\d{2}\.\d{2})-(?P<logistics>[^-]+(?:-[^-]+)*)-(?P<number>\d+)/(\d+)'
        match = re.search(pattern, ticket_str)
        if match:
            # 使用第四个捕获组（斜杠后的数字）作为件数
            return match.group('time'), match.group('logistics').replace('-', ''), match.group(4)
        else:
            return None, None, None
