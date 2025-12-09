#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ERP产品数据同步到MongoDB - 正式版
使用现有的数据库配置和连接
"""

import requests
import login
from db_utils import MongoDBClient  # 使用web_ticket相同的连接器
import time
from datetime import datetime
import json
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys

# 数据库操作锁
db_operation_lock = threading.Lock()

# 进度条锁
progress_lock = threading.Lock()


class ERPProductSync:
    """ERP产品同步类"""
    
    def __init__(self, token, max_retries=3):
        self.token = token
        self.max_retries = max_retries  # 最大重试次数
        self.db_client = MongoDBClient()
        self.db_client.connect()
        self.collection = self.db_client.db['msku_info']
        print(f"✓ 已连接到数据库: {self.db_client.config['database']}")
        print(f"✓ 使用集合: msku_info")
        print(f"✓ 网络重试次数: {self.max_retries}")
    
    def print_progress_bar(self, current, total, prefix='', suffix='', length=50):
        """打印进度条"""
        with progress_lock:
            percent = 100 * (current / float(total))
            filled_length = int(length * current // total)
            bar = '█' * filled_length + '-' * (length - filled_length)
            sys.stdout.write(f'\r{prefix} |{bar}| {percent:.1f}% {suffix}')
            sys.stdout.flush()
            if current == total:
                print()  # 完成后换行
    
    def request_with_retry(self, request_func, *args, **kwargs):
        """带重试机制的请求"""
        for attempt in range(self.max_retries):
            try:
                result = request_func(*args, **kwargs)
                return result
            except requests.exceptions.RequestException as e:
                if attempt < self.max_retries - 1:
                    wait_time = (attempt + 1) * 2  # 递增等待时间：2秒、4秒、6秒
                    print(f"\n  ⚠ 网络错误，{wait_time}秒后重试 ({attempt + 1}/{self.max_retries}): {str(e)}")
                    time.sleep(wait_time)
                else:
                    print(f"\n  ✗ 重试{self.max_retries}次后仍失败: {str(e)}")
                    raise
            except Exception as e:
                print(f"\n  ✗ 请求失败: {str(e)}")
                raise
        return None
        
    def __del__(self):
        """关闭MongoDB连接"""
        if hasattr(self, 'db_client'):
            self.db_client.close()
    
    def get_product_list(self, limit=None):
        """获取产品列表（筛选已配对的产品）"""
        headers = {
            'AK-Client-Type': 'web',
            'AK-Origin': 'https://erp.lingxing.com',
            'Accept': 'application/json, text/plain, */*',
            'Content-Type': 'application/json;charset=UTF-8',
            'X-AK-Company-Id': '901217529031491584',
            'X-AK-ENV-KEY': 'SAAS-101',
            'X-AK-PLATFORM': '1',
            'X-AK-Zid': '10330128',
            'auth-token': self.token,
        }
        
        json_data = {
            'search_field_time': 'create_time',
            'sort_field': 'create_time',
            'sort_type': 'desc',
            'search_field': 'sku',
            'status': [1],              # 在售状态
            'is_related': 1,            # 已配对 (1=已配对, 0=未配对)
            'offset': 0,
            'length': limit if limit else 500,
            'product_type': [1, 2],
            'req_time_sequence': '/api/product/lists$$17',
        }
        
        all_data = []
        offset = 0
        
        while True:
            json_data['offset'] = offset
            print(f"  正在获取产品列表，offset={offset}...")
            
            try:
                response = self.request_with_retry(
                    requests.post,
                    'https://erp.lingxing.com/api/product/lists',
                    headers=headers,
                    json=json_data,
                    timeout=30
                )
                
                if response.status_code == 200:
                    data = response.json()
                    if data.get('code') == 1 and 'list' in data:
                        fetched = data['list']
                        all_data.extend(fetched)
                        print(f"    获取到 {len(fetched)} 条产品")
                        
                        if limit or len(fetched) < 500:
                            break
                        offset += 500
                    else:
                        print(f"  业务错误: {data.get('msg')}")
                        break
                else:
                    print(f"  HTTP请求失败: {response.status_code}")
                    break
            except Exception as e:
                print(f"  获取产品列表失败: {str(e)}")
                break
            
            time.sleep(0.5)
        
        print(f"  共获取 {len(all_data)} 条产品")
        
        # 验证：确保所有产品都是已配对的 (is_related=1)
        matched_count = sum(1 for p in all_data if p.get('is_related') == 1)
        unmatched_count = len(all_data) - matched_count
        print(f"  ✓ 已配对产品: {matched_count} 个")
        if unmatched_count > 0:
            print(f"  ⚠ 未配对产品: {unmatched_count} 个 (将被过滤)")
            # 过滤掉未配对的产品
            all_data = [p for p in all_data if p.get('is_related') == 1]
        
        return all_data
    
    def get_product_detail(self, product_id):
        """获取产品详情（带重试）"""
        headers = {
            'AK-Client-Type': 'web',
            'Accept': 'application/json, text/plain, */*',
            'X-AK-Company-Id': '901217529031491584',
            'X-AK-ENV-KEY': 'SAAS-101',
            'X-AK-PLATFORM': '1',
            'X-AK-Zid': '10330128',
            'auth-token': self.token,
        }
        
        try:
            response = self.request_with_retry(
                requests.get,
                f'https://erp.lingxing.com/api/product/info?id={product_id}',
                headers=headers,
                timeout=30
            )
            
            if response.status_code == 200:
                return response.json()
        except Exception:
            pass
        return None
    
    def get_product_links(self, product_id):
        """获取产品链接（MSKU列表，带重试）"""
        headers = {
            'AK-Client-Type': 'web',
            'Accept': 'application/json, text/plain, */*',
            'X-AK-Company-Id': '901217529031491584',
            'X-AK-ENV-KEY': 'SAAS-101',
            'X-AK-PLATFORM': '1',
            'X-AK-Zid': '10330128',
            'auth-token': self.token,
        }
        
        try:
            response = self.request_with_retry(
                requests.get,
                f'https://erp.lingxing.com/api/module/product/product.view/getProductListing?product_id={product_id}',
                headers=headers,
                timeout=30
            )
            
            if response.status_code == 200:
                return response.json()
        except Exception:
            pass
        return None
    
    def split_field(self, field_value, en_field_value=None):
        """
        拆分中英文字段，支持三种情况：
        1. 分开存储：customs_clearance_material + customs_clearance_en_material
        2. 复合存储：customs_clearance_material = "中文/English" (斜杠分割)
        3. 真实缺失：字段为空或不存在
        
        Args:
            field_value: 中文字段值（可能包含英文）
            en_field_value: 英文字段值（如果分开存储）
        
        Returns:
            tuple: (中文, 英文)
        """
        # 情况1: 分开存储 - 如果有独立的英文字段，优先使用
        if en_field_value and en_field_value.strip():
            zh = field_value.strip() if field_value else ""
            en = en_field_value.strip()
            return zh, en
        
        # 情况3: 真实缺失 - 字段为空
        if not field_value or field_value.strip() == "":
            return "", ""
        
        # 情况2: 复合存储 - 包含斜杠分隔符 (注意: JSON中的\/会被自动解析为/)
        if '/' in field_value:
            parts = field_value.split('/', 1)  # 只分割第一个斜杠
            zh = parts[0].strip() if len(parts) > 0 else ""
            en = parts[1].strip() if len(parts) > 1 else ""
            return zh, en
        
        # 其他情况: 只有中文，没有英文
        return field_value.strip(), ""
    
    def parse_special_attr(self, special_attr):
        """解析special_attr字段"""
        if not special_attr or len(special_attr) == 0:
            return "否", "否"
        
        attr_list = [str(x) for x in special_attr]
        has_electric = "是" if ("1" in attr_list or "2" in attr_list) else "否"
        has_magnetic = "是" if "6" in attr_list else "否"
        
        return has_electric, has_magnetic
    
    def transform_product_data(self, product_info, link_info):
        """
        将ERP产品数据转换为MongoDB格式
        
        处理材质和用途的三种情况：
        1. 分开存储：customs_clearance_material + customs_clearance_en_material
        2. 复合存储：customs_clearance_material = "中文/English"
        3. 真实缺失：字段为空
        """
        info = product_info.get('info', {})
        declaration = info.get('product_declaration_list', {})
        clearance = info.get('product_clearance_list', {})
        special_attr = info.get('special_attr', [])
        spec_info = info.get('spec_info', {})
        
        # 解析电磁属性
        electrified, magnetic = self.parse_special_attr(special_attr)
        
        # 解析材质（支持三种情况）
        material_zh, material_en = self.split_field(
            clearance.get('customs_clearance_material', ''),
            clearance.get('customs_clearance_en_material', '')
        )
        
        # 解析用途（支持三种情况）
        use_zh, use_en = self.split_field(
            clearance.get('customs_clearance_usage', ''),
            clearance.get('customs_clearance_en_usage', '')
        )
        
        # 处理品牌和型号（空值显示"无"）
        brand = info.get('brand_name', '')
        brand = brand.strip() if brand else ''
        brand = brand if brand else "无"
        
        model = info.get('model', '')
        model = model.strip() if model else ''
        model = model if model else "无"
        
        # 获取重量（ERP中单位是克，转换为千克）
        weight_str = spec_info.get('cg_product_net_weight', '0')
        try:
            weight_g = float(weight_str)  # 原始重量（克）
            weight = weight_g * 0.001  # 转换为千克
        except (ValueError, TypeError):
            weight = 0.0
        
        document = {
            "msku": link_info.get('msku', ''),
            "productNameZh": declaration.get('customs_export_name', ''),
            "productNameEn": declaration.get('customs_import_name', ''),
            "price": declaration.get('customs_import_price', ''),
            "brand": brand,
            "model": model,
            "HS": declaration.get('customs_declaration_hs_code', ''),
            "image_url": "",
            "asin": link_info.get('asin', ''),
            "askPrice": "",
            "electrified": electrified,
            "magnetic": magnetic,
            "materialEn": material_en,
            "materialZh": material_zh,
            "outboundFee": "",
            "productLink": link_info.get('asin_url', ''),
            "putAwayFee": "",
            "useEn": use_en,
            "useZh": use_zh,
            "weight": weight,
            "X_ROW_K": ""
            # created_at 和 updated_at 在插入/更新时自动设置
        }
        
        return document
    
    def check_existing_mskus(self):
        """获取现有的MSKU列表（参考web_ticket实现）"""
        with db_operation_lock:
            existing_mskus = set(doc['msku'] for doc in self.collection.find({}, {'msku': 1}))
        return existing_mskus
    
    def get_existing_data(self):
        """获取现有的MSKU完整数据（用于比较更新）"""
        with db_operation_lock:
            existing_data = {}
            for doc in self.collection.find({}):
                msku = doc.get('msku')
                if msku:
                    existing_data[msku] = doc
        return existing_data
    
    # 需要从ERP同步的字段（这些字段以ERP为准）
    ERP_SYNC_FIELDS = [
        'productNameZh', 'productNameEn', 'price', 'brand', 'model', 'HS',
        'asin', 'electrified', 'magnetic', 'materialEn', 'materialZh',
        'productLink', 'useEn', 'useZh', 'weight'
    ]
    
    # 本地维护的字段（不会被ERP覆盖）
    LOCAL_FIELDS = ['image_url', 'askPrice', 'outboundFee', 'putAwayFee', 'X_ROW_K', 'created_at']
    
    def compare_and_get_updates(self, erp_doc, existing_doc):
        """
        比较ERP数据与现有数据，返回需要更新的字段
        
        Args:
            erp_doc: 从ERP获取的数据
            existing_doc: MongoDB中现有的数据
        
        Returns:
            dict: 需要更新的字段，如果没有差异返回空字典
        """
        updates = {}
        
        for field in self.ERP_SYNC_FIELDS:
            erp_value = erp_doc.get(field, '')
            existing_value = existing_doc.get(field, '')
            
            # 统一处理空值比较
            if erp_value is None:
                erp_value = ''
            if existing_value is None:
                existing_value = ''
            
            # 数值类型特殊处理
            if field == 'weight':
                try:
                    erp_float = float(erp_value) if erp_value != '' else 0.0
                    existing_float = float(existing_value) if existing_value != '' else 0.0
                    if abs(erp_float - existing_float) > 0.0001:
                        updates[field] = erp_float
                except (ValueError, TypeError):
                    if str(erp_value) != str(existing_value):
                        updates[field] = erp_value
            elif field == 'price':
                try:
                    erp_price = float(erp_value) if erp_value != '' else 0.0
                    existing_price = float(existing_value) if existing_value != '' else 0.0
                    if abs(erp_price - existing_price) > 0.001:
                        updates[field] = erp_value
                except (ValueError, TypeError):
                    if str(erp_value) != str(existing_value):
                        updates[field] = erp_value
            else:
                # 字符串比较
                if str(erp_value).strip() != str(existing_value).strip():
                    updates[field] = erp_value
        
        return updates
    
    def batch_insert_documents(self, documents):
        """批量插入文档（参考web_ticket实现）"""
        if not documents:
            return 0
        
        with db_operation_lock:
            # 再次检查MSKU是否存在（防止并发插入）
            current_mskus = set(doc['msku'] for doc in self.collection.find({}, {'msku': 1}))
            documents_to_insert = [doc for doc in documents if doc.get('msku') not in current_mskus]
            
            if documents_to_insert:
                # 为新插入的文档添加created_at和updated_at
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                for doc in documents_to_insert:
                    doc['created_at'] = now
                    doc['updated_at'] = now
                self.collection.insert_many(documents_to_insert)
                return len(documents_to_insert)
        
        return 0
    
    def batch_update_documents(self, updates_list):
        """
        批量更新文档
        
        Args:
            updates_list: [(msku, updates_dict), ...] 格式的更新列表
        
        Returns:
            int: 成功更新的数量
        """
        if not updates_list:
            return 0
        
        updated_count = 0
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        with db_operation_lock:
            for msku, updates in updates_list:
                if updates:
                    # 添加updated_at时间戳
                    updates['updated_at'] = now
                    result = self.collection.update_one(
                        {'msku': msku},
                        {'$set': updates}
                    )
                    if result.modified_count > 0:
                        updated_count += 1
        
        return updated_count
    
    def process_single_product(self, product, existing_mskus):
        """处理单个产品（用于并发）"""
        product_id = product.get('id')
        sku = product.get('sku', 'Unknown')
        documents = []
        
        try:
            # 获取产品详情
            detail = self.get_product_detail(product_id)
            if not detail or detail.get('code') != 1:
                return {'success': False, 'error': '获取产品详情失败', 'sku': sku}
            
            # 获取产品链接
            links = self.get_product_links(product_id)
            if not links or links.get('code') != 1:
                return {'success': False, 'error': '获取产品链接失败', 'sku': sku}
            
            link_data = links.get('data', [])
            if not link_data:
                return {'success': False, 'error': '没有MSKU链接', 'sku': sku}
            
            # 处理每个MSKU
            new_mskus = []
            skipped_mskus = []
            for link in link_data:
                msku = link.get('msku', '')
                if msku in existing_mskus:
                    skipped_mskus.append(msku)
                else:
                    document = self.transform_product_data(detail, link)
                    documents.append(document)
                    new_mskus.append(msku)
            
            return {
                'success': True,
                'sku': sku,
                'documents': documents,
                'new_mskus': new_mskus,
                'skipped_mskus': skipped_mskus
            }
            
        except Exception as e:
            return {'success': False, 'error': str(e), 'sku': sku}
    
    def process_single_product_cached(self, product, existing_mskus):
        """处理单个产品（使用缓存的链接数据）- 旧版本，仅插入"""
        product_id = product.get('id')
        sku = product.get('sku', 'Unknown')
        documents = []
        
        try:
            # 添加小延时，避免并发请求过快
            time.sleep(0.1)
            
            # 获取产品详情
            detail = self.get_product_detail(product_id)
            if not detail or detail.get('code') != 1:
                return {'success': False, 'error': '获取产品详情失败', 'sku': sku}
            
            # 使用缓存的链接数据
            link_data = product.get('_links', [])
            if not link_data:
                return {'success': False, 'error': '没有MSKU链接', 'sku': sku}
            
            # 处理每个MSKU
            new_mskus = []
            skipped_mskus = []
            for link in link_data:
                msku = link.get('msku', '')
                if msku in existing_mskus:
                    skipped_mskus.append(msku)
                else:
                    document = self.transform_product_data(detail, link)
                    documents.append(document)
                    new_mskus.append(msku)
            
            return {
                'success': True,
                'sku': sku,
                'documents': documents,
                'new_mskus': new_mskus,
                'skipped_mskus': skipped_mskus
            }
            
        except Exception as e:
            return {'success': False, 'error': str(e), 'sku': sku}
    
    def process_single_product_sync(self, product, existing_data):
        """
        处理单个产品（支持插入和更新）
        
        Args:
            product: 产品信息（包含缓存的_links）
            existing_data: 现有MSKU完整数据字典 {msku: doc}
        
        Returns:
            dict: 包含新增文档、更新列表等信息
        """
        product_id = product.get('id')
        sku = product.get('sku', 'Unknown')
        
        try:
            # 添加小延时，避免并发请求过快
            time.sleep(0.1)
            
            # 获取产品详情
            detail = self.get_product_detail(product_id)
            if not detail or detail.get('code') != 1:
                return {'success': False, 'error': '获取产品详情失败', 'sku': sku}
            
            # 使用缓存的链接数据
            link_data = product.get('_links', [])
            if not link_data:
                return {'success': False, 'error': '没有MSKU链接', 'sku': sku}
            
            # 处理每个MSKU
            new_documents = []      # 新增的文档
            updates_list = []       # 需要更新的 [(msku, updates), ...]
            unchanged_mskus = []    # 无变化的MSKU
            new_mskus = []          # 新增的MSKU
            updated_mskus = []      # 更新的MSKU
            
            for link in link_data:
                msku = link.get('msku', '')
                if not msku:
                    continue
                
                # 转换ERP数据
                erp_doc = self.transform_product_data(detail, link)
                
                if msku in existing_data:
                    # MSKU已存在，比较差异
                    existing_doc = existing_data[msku]
                    updates = self.compare_and_get_updates(erp_doc, existing_doc)
                    
                    if updates:
                        # 有差异，需要更新
                        updates_list.append((msku, updates))
                        updated_mskus.append(msku)
                    else:
                        # 无差异
                        unchanged_mskus.append(msku)
                else:
                    # 新MSKU，需要插入
                    new_documents.append(erp_doc)
                    new_mskus.append(msku)
            
            return {
                'success': True,
                'sku': sku,
                'new_documents': new_documents,
                'updates_list': updates_list,
                'new_mskus': new_mskus,
                'updated_mskus': updated_mskus,
                'unchanged_mskus': unchanged_mskus
            }
            
        except Exception as e:
            return {'success': False, 'error': str(e), 'sku': sku}
    
    def sync_products(self, limit=None, use_concurrent=True, max_workers=5):
        """
        同步产品数据（全量同步版 - 支持插入和更新）
        
        Args:
            limit: 限制处理的产品数量
            use_concurrent: 是否使用并发处理
            max_workers: 并发线程数（默认5，避免触发限流）
        
        同步策略：
            - 新MSKU：插入新记录，设置created_at和updated_at
            - 已存在MSKU：对比ERP_SYNC_FIELDS字段，有差异则更新，同时更新updated_at
            - 本地字段（image_url, askPrice等）：不会被ERP覆盖
        """
        print("\n" + "=" * 70)
        print("开始同步ERP产品数据到MongoDB（全量同步模式）")
        print("  策略：新增插入 + 差异更新（以ERP为准）")
        if use_concurrent:
            print(f"  并发：最大{max_workers}个线程")
        print("=" * 70)
        
        try:
            # 步骤1: 获取产品列表
            print("\n[步骤1] 获取ERP产品列表...")
            products = self.get_product_list(limit=limit)
            
            # 步骤2: 获取现有MSKU完整数据（用于比较更新）
            print("\n[步骤2] 获取数据库现有数据...")
            existing_data = self.get_existing_data()
            print(f"  数据库中已有 {len(existing_data)} 个MSKU")
            
            # 步骤2.5: 获取所有产品的MSKU链接并缓存
            print("\n[步骤2.5] 获取产品MSKU链接...")
            products_to_process = []
            total_products = len(products)
            
            for idx, product in enumerate(products, 1):
                product_id = product.get('id')
                
                # 显示进度条
                self.print_progress_bar(
                    idx, 
                    total_products, 
                    prefix='获取链接',
                    suffix=f'({idx}/{total_products})'
                )
                
                # 获取产品链接
                links = self.get_product_links(product_id)
                if links and links.get('code') == 1:
                    link_data = links.get('data', [])
                    if link_data:
                        product['_links'] = link_data  # 缓存链接数据
                        products_to_process.append(product)
                time.sleep(0.1)  # 避免请求过快
            
            print(f"\n  ✓ 需要处理: {len(products_to_process)} 个产品")
            
            stats = {
                'total': len(products),
                'processed': 0,
                'inserted': 0,
                'updated': 0,
                'unchanged': 0,
                'skipped': 0,  # 兼容旧版调用（web_ticket.py需要此字段）
                'errors': 0,
                'msku_count': 0
            }
            
            if not products_to_process:
                print("\n✓ 没有产品需要处理")
                return stats
            
            # 步骤3: 处理产品并收集待插入/更新的数据
            print(f"\n[步骤3] 开始处理 {len(products_to_process)} 个产品...")
            documents_to_insert = []
            updates_to_apply = []  # [(msku, updates), ...]
            
            if use_concurrent and len(products_to_process) > 1:
                # 并发处理
                print(f"  使用 {max_workers} 个线程并发处理...")
                total_to_process = len(products_to_process)
                
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # 提交所有任务
                    future_to_product = {
                        executor.submit(self.process_single_product_sync, product, existing_data): product
                        for product in products_to_process
                    }
                    
                    # 处理完成的任务
                    completed = 0
                    for future in as_completed(future_to_product):
                        completed += 1
                        product = future_to_product[future]
                        sku = product.get('sku', 'Unknown')
                        
                        try:
                            result = future.result()
                            if result['success']:
                                documents_to_insert.extend(result['new_documents'])
                                updates_to_apply.extend(result['updates_list'])
                                stats['processed'] += 1
                                stats['msku_count'] += len(result['new_mskus']) + len(result['updated_mskus']) + len(result['unchanged_mskus'])
                                
                                # 显示进度条
                                new_count = len(result['new_mskus'])
                                update_count = len(result['updated_mskus'])
                                unchanged_count = len(result['unchanged_mskus'])
                                self.print_progress_bar(
                                    completed,
                                    total_to_process,
                                    prefix='处理进度',
                                    suffix=f'{sku} ✓ 新增{new_count}/更新{update_count}/无变化{unchanged_count}'
                                )
                            else:
                                stats['errors'] += 1
                                self.print_progress_bar(
                                    completed,
                                    total_to_process,
                                    prefix='处理进度',
                                    suffix=f'{sku} ✗ {result["error"]}'
                                )
                        except Exception as e:
                            stats['errors'] += 1
                            self.print_progress_bar(
                                completed,
                                total_to_process,
                                prefix='处理进度',
                                suffix=f'{sku} ✗ {str(e)}'
                            )
            else:
                # 串行处理
                for idx, product in enumerate(products_to_process, 1):
                    sku = product.get('sku', 'Unknown')
                    print(f"\n--- [{idx}/{len(products_to_process)}] 处理产品: {sku} ---")
                    
                    try:
                        result = self.process_single_product_sync(product, existing_data)
                        if result['success']:
                            documents_to_insert.extend(result['new_documents'])
                            updates_to_apply.extend(result['updates_list'])
                            stats['processed'] += 1
                            stats['msku_count'] += len(result['new_mskus']) + len(result['updated_mskus']) + len(result['unchanged_mskus'])
                            print(f"  ✓ 新增{len(result['new_mskus'])}个, 更新{len(result['updated_mskus'])}个, 无变化{len(result['unchanged_mskus'])}个")
                        else:
                            stats['errors'] += 1
                            print(f"  ✗ {result['error']}")
                    except Exception as e:
                        print(f"  ❌ 处理出错: {str(e)}")
                        stats['errors'] += 1
                    
                    time.sleep(0.3)
            
            # 步骤4: 批量插入新数据
            print(f"\n[步骤4] 批量插入新数据...")
            if documents_to_insert:
                print(f"  准备插入 {len(documents_to_insert)} 条新记录...")
                inserted_count = self.batch_insert_documents(documents_to_insert)
                stats['inserted'] = inserted_count
                print(f"  ✓ 成功插入 {inserted_count} 条记录")
            else:
                print(f"  没有新数据需要插入")
            
            # 步骤5: 批量更新现有数据
            print(f"\n[步骤5] 批量更新现有数据...")
            if updates_to_apply:
                print(f"  准备更新 {len(updates_to_apply)} 条记录...")
                updated_count = self.batch_update_documents(updates_to_apply)
                stats['updated'] = updated_count
                print(f"  ✓ 成功更新 {updated_count} 条记录")
            else:
                print(f"  没有数据需要更新")
            
            # 计算无变化的数量
            stats['unchanged'] = stats['msku_count'] - stats['inserted'] - stats['updated']
            stats['skipped'] = stats['unchanged']  # 兼容旧版调用
            
            # 步骤6: 输出统计信息
            print("\n" + "=" * 70)
            print("同步完成！统计信息：")
            print("=" * 70)
            print(f"总产品数:       {stats['total']}")
            print(f"处理产品数:     {stats['processed']}")
            print(f"总MSKU数:       {stats['msku_count']}")
            print(f"  - 新增:       {stats['inserted']}")
            print(f"  - 更新:       {stats['updated']}")
            print(f"  - 无变化:     {stats['unchanged']}")
            print(f"错误数量:       {stats['errors']}")
            print("=" * 70)
            
            return stats
            
        finally:
            # 确保关闭数据库连接
            if hasattr(self, 'db_client'):
                self.db_client.close()
                print("\n✓ 数据库连接已关闭")


def main():
    """主函数"""
    print("=" * 70)
    print("ERP产品数据同步到MongoDB - 正式版")
    print("=" * 70)
    
    # 询问用户
    print("\n请选择运行模式：")
    print("1. 测试模式（只处理前10个产品）")
    print("2. 正式模式（处理所有已配对产品）")
    
    choice = input("\n请输入选择 (1/2): ").strip()
    
    if choice == '1':
        limit = 10
        print(f"\n✓ 选择测试模式，将处理前 {limit} 个产品")
    elif choice == '2':
        limit = None
        confirm = input("\n⚠️  正式模式将处理所有产品，确认继续？(yes/no): ").strip().lower()
        if confirm != 'yes':
            print("已取消")
            return
        print("\n✓ 选择正式模式，将处理所有已配对产品")
    else:
        print("无效选择，退出")
        return
    
    print("\n[1] 登录ERP系统...")
    token = login.run()
    print(f"✓ 登录成功")
    
    print("\n[2] 初始化同步对象...")
    sync = ERPProductSync(token)
    
    print("\n[3] 开始同步数据...")
    stats = sync.sync_products(limit=limit)
    
    print("\n[4] 查看插入的数据示例...")
    if stats['inserted'] > 0:
        sample = sync.collection.find_one()
        if sample:
            print("\n示例记录：")
            print(json.dumps(sample, indent=2, ensure_ascii=False, default=str))
    
    print("\n✓ 同步完成！")


if __name__ == "__main__":
    main()
