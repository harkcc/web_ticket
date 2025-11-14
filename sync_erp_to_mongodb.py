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

# 数据库操作锁（与web_ticket保持一致）
db_operation_lock = threading.Lock()


class ERPProductSync:
    """ERP产品同步类"""
    
    def __init__(self, token):
        self.token = token
        self.db_client = MongoDBClient()
        self.db_client.connect()
        self.collection = self.db_client.db['msku_info']
        print(f"✓ 已连接到数据库: {self.db_client.config['database']}")
        print(f"✓ 使用集合: msku_info")
        
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
            'status': [1],
            'is_matched_alibaba': '1',
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
            
            response = requests.post(
                'https://erp.lingxing.com/api/product/lists',
                headers=headers,
                json=json_data
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
            
            time.sleep(0.5)
        
        print(f"  共获取 {len(all_data)} 条产品")
        return all_data
    
    def get_product_detail(self, product_id):
        """获取产品详情"""
        headers = {
            'AK-Client-Type': 'web',
            'Accept': 'application/json, text/plain, */*',
            'X-AK-Company-Id': '901217529031491584',
            'X-AK-ENV-KEY': 'SAAS-101',
            'X-AK-PLATFORM': '1',
            'X-AK-Zid': '10330128',
            'auth-token': self.token,
        }
        
        response = requests.get(
            f'https://erp.lingxing.com/api/product/info?id={product_id}',
            headers=headers
        )
        
        if response.status_code == 200:
            return response.json()
        return None
    
    def get_product_links(self, product_id):
        """获取产品链接（MSKU列表）"""
        headers = {
            'AK-Client-Type': 'web',
            'Accept': 'application/json, text/plain, */*',
            'X-AK-Company-Id': '901217529031491584',
            'X-AK-ENV-KEY': 'SAAS-101',
            'X-AK-PLATFORM': '1',
            'X-AK-Zid': '10330128',
            'auth-token': self.token,
        }
        
        response = requests.get(
            f'https://erp.lingxing.com/api/module/product/product.view/getProductListing?product_id={product_id}',
            headers=headers
        )
        
        if response.status_code == 200:
            return response.json()
        return None
    
    def split_field(self, field_value, en_field_value=None):
        """拆分中英文字段"""
        if en_field_value and en_field_value.strip():
            zh = field_value.strip() if field_value else ""
            en = en_field_value.strip()
            return zh, en
        
        if not field_value or field_value.strip() == "":
            return "", ""
        
        if '/' in field_value:
            parts = field_value.split('/', 1)
            zh = parts[0].strip() if len(parts) > 0 else ""
            en = parts[1].strip() if len(parts) > 1 else ""
            return zh, en
        
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
        """将ERP产品数据转换为MongoDB格式"""
        info = product_info.get('info', {})
        declaration = info.get('product_declaration_list', {})
        clearance = info.get('product_clearance_list', {})
        special_attr = info.get('special_attr', [])
        spec_info = info.get('spec_info', {})
        
        electrified, magnetic = self.parse_special_attr(special_attr)
        
        material_zh, material_en = self.split_field(
            clearance.get('customs_clearance_material', ''),
            clearance.get('customs_clearance_en_material', '')
        )
        
        use_zh, use_en = self.split_field(
            clearance.get('customs_clearance_usage', ''),
            clearance.get('customs_clearance_en_usage', '')
        )
        
        weight_str = spec_info.get('cg_product_net_weight', '0')
        try:
            weight = float(weight_str)
        except:
            weight = 0.0
        
        document = {
            "msku": link_info.get('msku', ''),
            "productNameZh": declaration.get('customs_export_name', ''),
            "productNameEn": declaration.get('customs_import_name', ''),
            "price": declaration.get('customs_import_price', ''),
            "brand": info.get('brand_name', '') or "无",
            "model": info.get('model', '') or "无",
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
            "X_ROW_K": "",
            "created_at": ""
        }
        
        return document
    
    def check_existing_mskus(self):
        """获取现有的MSKU列表（参考web_ticket实现）"""
        with db_operation_lock:
            existing_mskus = set(doc['msku'] for doc in self.collection.find({}, {'msku': 1}))
        return existing_mskus
    
    def batch_insert_documents(self, documents):
        """批量插入文档（参考web_ticket实现）"""
        if not documents:
            return 0
        
        with db_operation_lock:
            # 再次检查MSKU是否存在（防止并发插入）
            current_mskus = set(doc['msku'] for doc in self.collection.find({}, {'msku': 1}))
            documents_to_insert = [doc for doc in documents if doc.get('msku') not in current_mskus]
            
            if documents_to_insert:
                self.collection.insert_many(documents_to_insert)
                return len(documents_to_insert)
        
        return 0
    
    def sync_products(self, limit=None):
        """同步产品数据（参考web_ticket实现）"""
        print("\n" + "=" * 70)
        print("开始同步ERP产品数据到MongoDB")
        print("=" * 70)
        
        try:
            # 步骤1: 获取产品列表
            print("\n[步骤1] 获取产品列表...")
            products = self.get_product_list(limit=limit)
            
            # 步骤2: 获取现有MSKU列表
            print("\n[步骤2] 获取现有MSKU列表...")
            existing_mskus = self.check_existing_mskus()
            print(f"  数据库中已有 {len(existing_mskus)} 个MSKU")
            
            stats = {
                'total': len(products),
                'processed': 0,
                'inserted': 0,
                'skipped': 0,
                'errors': 0,
                'msku_count': 0
            }
            
            # 步骤3: 处理产品并收集待插入文档
            print(f"\n[步骤3] 开始处理 {stats['total']} 个产品...")
            documents_to_insert = []
            
            for idx, product in enumerate(products, 1):
                product_id = product.get('id')
                sku = product.get('sku', 'Unknown')
                
                print(f"\n--- [{idx}/{stats['total']}] 处理产品: {sku} (ID: {product_id}) ---")
                
                try:
                    # 获取产品详情
                    print(f"  → 获取产品详情...")
                    detail = self.get_product_detail(product_id)
                    if not detail or detail.get('code') != 1:
                        print(f"  ❌ 获取产品详情失败")
                        stats['errors'] += 1
                        continue
                    
                    # 获取产品链接
                    print(f"  → 获取产品链接...")
                    links = self.get_product_links(product_id)
                    if not links or links.get('code') != 1:
                        print(f"  ❌ 获取产品链接失败")
                        stats['errors'] += 1
                        continue
                    
                    link_data = links.get('data', [])
                    if not link_data:
                        print(f"  ⚠️  该产品没有MSKU链接")
                        stats['skipped'] += 1
                        continue
                    
                    print(f"  ✓ 找到 {len(link_data)} 个MSKU")
                    
                    # 处理每个MSKU
                    for link in link_data:
                        msku = link.get('msku', '')
                        stats['msku_count'] += 1
                        
                        if msku in existing_mskus:
                            print(f"    ⊗ MSKU {msku} 已存在，跳过")
                            stats['skipped'] += 1
                        else:
                            print(f"    → 转换MSKU: {msku}")
                            document = self.transform_product_data(detail, link)
                            documents_to_insert.append(document)
                            print(f"    ✓ 已加入待插入列表")
                    
                    stats['processed'] += 1
                    
                except Exception as e:
                    print(f"  ❌ 处理出错: {str(e)}")
                    stats['errors'] += 1
                
                time.sleep(0.3)
            
            # 步骤4: 批量插入数据
            print(f"\n[步骤4] 批量插入数据...")
            if documents_to_insert:
                print(f"  准备插入 {len(documents_to_insert)} 条记录...")
                inserted_count = self.batch_insert_documents(documents_to_insert)
                stats['inserted'] = inserted_count
                print(f"  ✓ 成功插入 {inserted_count} 条记录")
            else:
                print(f"  没有新数据需要插入")
            
            # 步骤5: 输出统计信息
            print("\n" + "=" * 70)
            print("同步完成！统计信息：")
            print("=" * 70)
            print(f"总产品数:     {stats['total']}")
            print(f"成功处理:     {stats['processed']}")
            print(f"总MSKU数:     {stats['msku_count']}")
            print(f"插入记录:     {stats['inserted']}")
            print(f"跳过记录:     {stats['skipped']}")
            print(f"错误数量:     {stats['errors']}")
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
