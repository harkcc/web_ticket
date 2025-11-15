#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
导出ERP产品数据到Excel表格
"""

import requests
import login
import time
import pandas as pd
from datetime import datetime
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed


class ERPDataExporter:
    def __init__(self, token, max_retries=3):
        self.token = token
        self.max_retries = max_retries
    
    def request_with_retry(self, request_func, *args, **kwargs):
        """带重试机制的请求"""
        for attempt in range(self.max_retries):
            try:
                result = request_func(*args, **kwargs)
                return result
            except requests.exceptions.RequestException as e:
                if attempt < self.max_retries - 1:
                    wait_time = (attempt + 1) * 2
                    print(f"  ⚠ 网络错误，{wait_time}秒后重试 ({attempt + 1}/{self.max_retries}): {str(e)}")
                    time.sleep(wait_time)
                else:
                    print(f"  ✗ 重试{self.max_retries}次后仍失败: {str(e)}")
                    raise
        return None
    
    def print_progress_bar(self, current, total, prefix='', suffix='', length=50):
        """打印进度条"""
        percent = 100 * (current / float(total))
        filled_length = int(length * current // total)
        bar = '█' * filled_length + '-' * (length - filled_length)
        sys.stdout.write(f'\r{prefix} |{bar}| {percent:.1f}% {suffix}')
        sys.stdout.flush()
        if current == total:
            print()
    
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
            'is_related': 1,  # 已配对
            'offset': 0,
            'length': limit if limit else 500,
            'product_type': [1, 2],
            'req_time_sequence': '/api/product/lists$$17',
        }
        
        all_data = []
        offset = 0
        
        print("\n[步骤1] 获取产品列表...")
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
        
        print(f"  ✓ 共获取 {len(all_data)} 条产品")
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
    
    def extract_product_data(self, product_info, link_info):
        """提取产品数据为字典"""
        info = product_info.get('info', {})
        declaration = info.get('product_declaration_list', {})
        clearance = info.get('product_clearance_list', {})
        special_attr = info.get('special_attr', [])
        spec_info = info.get('spec_info', {})
        
        # 解析电磁属性
        electrified, magnetic = self.parse_special_attr(special_attr)
        
        # 解析材质
        material_zh, material_en = self.split_field(
            clearance.get('customs_clearance_material', ''),
            clearance.get('customs_clearance_en_material', '')
        )
        
        # 解析用途
        use_zh, use_en = self.split_field(
            clearance.get('customs_clearance_usage', ''),
            clearance.get('customs_clearance_en_usage', '')
        )
        
        # 获取重量（ERP中单位是克，转换为千克）
        weight_str = spec_info.get('cg_product_net_weight', '0')
        try:
            weight_g = float(weight_str)  # 原始重量（克）
            weight = weight_g * 0.001  # 转换为千克
        except (ValueError, TypeError):
            weight = 0.0
        
        # 处理品牌和型号（空值显示"无"）
        brand = info.get('brand_name', '')
        brand = brand.strip() if brand else ''
        brand = brand if brand else "无"
        
        model = info.get('model', '')
        model = model.strip() if model else ''
        model = model if model else "无"
        
        return {
            'MSKU': link_info.get('msku', ''),
            'SKU': info.get('sku', ''),
            '产品ID': info.get('id', ''),
            '产品名称(中文)': declaration.get('customs_export_name', ''),
            '产品名称(英文)': declaration.get('customs_import_name', ''),
            '申报价格': declaration.get('customs_import_price', ''),
            '品牌': brand,
            '型号': model,
            'HS编码': declaration.get('customs_declaration_hs_code', ''),
            'ASIN': link_info.get('asin', ''),
            '产品链接': link_info.get('asin_url', ''),
            '带电': electrified,
            '带磁': magnetic,
            '材质(中文)': material_zh,
            '材质(英文)': material_en,
            '用途(中文)': use_zh,
            '用途(英文)': use_en,
            '重量(kg)': weight,
            '创建时间': info.get('create_time', ''),
            '更新时间': info.get('update_time', ''),
        }
    
    def process_single_product(self, product):
        """处理单个产品（用于并发）"""
        product_id = product.get('id')
        sku = product.get('sku', 'Unknown')
        result = {
            'success': False,
            'sku': sku,
            'data': [],
            'error': None
        }
        
        try:
            # 添加小延时，避免并发请求过快
            time.sleep(0.1)
            
            # 获取产品详情
            detail = self.get_product_detail(product_id)
            if not detail or detail.get('code') != 1:
                result['error'] = '获取产品详情失败'
                return result
            
            # 获取产品链接
            links = self.get_product_links(product_id)
            if not links or links.get('code') != 1:
                result['error'] = '获取产品链接失败'
                return result
            
            link_data = links.get('data', [])
            if not link_data:
                result['error'] = '没有MSKU链接'
                return result
            
            # 提取每个MSKU的数据
            for link in link_data:
                data = self.extract_product_data(detail, link)
                result['data'].append(data)
            
            result['success'] = True
            return result
            
        except Exception as e:
            result['error'] = str(e)
            return result
    
    def export_to_excel(self, limit=None, use_concurrent=True, max_workers=5):
        """导出数据到Excel"""
        print("=" * 70)
        print("开始导出ERP产品数据到Excel")
        print("=" * 70)
        
        # 获取产品列表
        products = self.get_product_list(limit=limit)
        
        if not products:
            print("❌ 没有获取到产品数据")
            return
        
        # 收集所有数据
        all_data = []
        total_products = len(products)
        stats = {
            'success': 0,
            'failed': 0,
            'total_msku': 0
        }
        
        print(f"\n[步骤2] 处理产品数据...")
        
        if use_concurrent and total_products > 1:
            # 并发处理
            print(f"  ✓ 使用并发模式 ({max_workers} 个线程)")
            
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                # 提交所有任务
                future_to_product = {
                    executor.submit(self.process_single_product, product): product
                    for product in products
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
                            all_data.extend(result['data'])
                            stats['success'] += 1
                            stats['total_msku'] += len(result['data'])
                            
                            # 显示进度条
                            self.print_progress_bar(
                                completed,
                                total_products,
                                prefix='处理进度',
                                suffix=f'{sku} ✓ {len(result["data"])}个MSKU'
                            )
                        else:
                            stats['failed'] += 1
                            self.print_progress_bar(
                                completed,
                                total_products,
                                prefix='处理进度',
                                suffix=f'{sku} ✗ {result["error"]}'
                            )
                    except Exception as e:
                        stats['failed'] += 1
                        self.print_progress_bar(
                            completed,
                            total_products,
                            prefix='处理进度',
                            suffix=f'{sku} ✗ {str(e)}'
                        )
        else:
            # 串行处理
            print(f"  ✓ 使用串行模式")
            
            for idx, product in enumerate(products, 1):
                sku = product.get('sku', 'Unknown')
                
                # 显示进度条
                self.print_progress_bar(
                    idx,
                    total_products,
                    prefix='处理进度',
                    suffix=f'({idx}/{total_products}) {sku}'
                )
                
                result = self.process_single_product(product)
                if result['success']:
                    all_data.extend(result['data'])
                    stats['success'] += 1
                    stats['total_msku'] += len(result['data'])
                else:
                    stats['failed'] += 1
        
        # 转换为DataFrame
        print(f"\n\n[步骤3] 生成Excel文件...")
        df = pd.DataFrame(all_data)
        
        # 生成文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'msku_info_export_{timestamp}.xlsx'
        
        # 保存到Excel
        df.to_excel(filename, index=False, engine='openpyxl')
        
        # 统计信息
        print("\n" + "=" * 70)
        print("导出完成！")
        print("=" * 70)
        print(f"总产品数:       {total_products}")
        print(f"成功处理:       {stats['success']}")
        print(f"失败数量:       {stats['failed']}")
        print(f"总MSKU数:       {stats['total_msku']}")
        print(f"输出文件:       {filename}")
        print("=" * 70)
        
        # 显示字段统计
        print("\n[字段统计]")
        for col in df.columns:
            non_empty = df[col].notna().sum()
            has_value = (df[col] != '').sum() if df[col].dtype == 'object' else non_empty
            print(f"  {col:20s}: {has_value}/{len(df)} 条有值")
        
        return filename


def main():
    """主函数"""
    print("=" * 70)
    print("ERP产品数据导出工具")
    print("=" * 70)
    
    # 询问导出数量
    print("\n请选择导出模式：")
    print("1. 测试模式（导出前10个产品）")
    print("2. 部分导出（自定义数量）")
    print("3. 全量导出（所有产品）")
    
    choice = input("\n请输入选项 (1/2/3): ").strip()
    
    limit = None
    if choice == '1':
        limit = 10
        print(f"✓ 选择测试模式，将导出前 {limit} 个产品")
    elif choice == '2':
        limit_input = input("请输入要导出的产品数量: ").strip()
        try:
            limit = int(limit_input)
            print(f"✓ 将导出前 {limit} 个产品")
        except:
            print("❌ 输入无效，使用测试模式（10个产品）")
            limit = 10
    else:
        print("✓ 选择全量导出，将导出所有产品")
    
    # 询问是否使用并发
    print("\n是否使用并发模式？")
    print("1. 是（推荐，速度更快）")
    print("2. 否（串行模式）")
    
    concurrent_choice = input("\n请输入选项 (1/2，默认1): ").strip() or '1'
    use_concurrent = concurrent_choice == '1'
    
    max_workers = 5
    if use_concurrent:
        workers_input = input(f"并发线程数（默认{max_workers}，建议3-10）: ").strip()
        if workers_input:
            try:
                max_workers = int(workers_input)
                max_workers = max(1, min(max_workers, 10))  # 限制在1-10之间
            except:
                pass
        print(f"✓ 使用并发模式，线程数: {max_workers}")
    else:
        print("✓ 使用串行模式")
    
    # 登录
    print("\n[步骤0] 登录ERP系统...")
    token = login.run()
    print(f"✓ 登录成功")
    
    # 创建导出器
    exporter = ERPDataExporter(token, max_retries=3)
    
    # 导出数据
    try:
        filename = exporter.export_to_excel(
            limit=limit,
            use_concurrent=use_concurrent,
            max_workers=max_workers
        )
        print(f"\n✓ 数据已成功导出到: {filename}")
    except KeyboardInterrupt:
        print("\n\n⚠ 用户中断导出")
    except Exception as e:
        print(f"\n\n❌ 导出失败: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
