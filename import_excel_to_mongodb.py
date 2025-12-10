#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从Excel文件导入/更新产品数据到MongoDB
"""

import pandas as pd
from datetime import datetime
from db_utils import MongoDBClient
import sys


class ExcelToMongoImporter:
    """Excel数据导入MongoDB工具"""
    
    def __init__(self):
        self.client = MongoDBClient()
        self.collection = None
    
    # 需要从Excel同步的字段映射（Excel列名 -> MongoDB字段名）
    FIELD_MAPPING = {
        'MSKU': 'msku',
        '产品名称(中文)': 'productNameZh',
        '产品名称(英文)': 'productNameEn',
        '申报价格': 'price',
        '品牌': 'brand',
        '型号': 'model',
        'HS编码': 'HS',
        'ASIN': 'asin',
        '产品链接': 'productLink',
        '带电': 'electrified',
        '带磁': 'magnetic',
        '材质(中文)': 'materialZh',
        '材质(英文)': 'materialEn',
        '用途(中文)': 'useZh',
        '用途(英文)': 'useEn',
        '重量(kg)': 'weight',
    }
    
    # 本地维护的字段（不会被Excel覆盖）
    LOCAL_FIELDS = ['image_url', 'askPrice', 'outboundFee', 'putAwayFee', 'X_ROW_K', 'created_at']
    
    def connect_db(self):
        """连接数据库"""
        print("\n[步骤1] 连接MongoDB...")
        self.client.connect()
        self.collection = self.client.db['msku_info']
        print("  ✓ 数据库连接成功")
    
    def read_excel(self, excel_file):
        """读取Excel文件"""
        print(f"\n[步骤2] 读取Excel文件: {excel_file}")
        try:
            df = pd.read_excel(excel_file)
            print(f"  ✓ 读取成功，共 {len(df)} 条记录")
            print(f"  ✓ 列名: {df.columns.tolist()}")
            return df
        except Exception as e:
            print(f"  ✗ 读取失败: {str(e)}")
            return None
    
    def get_existing_data(self):
        """获取现有MSKU数据"""
        print("\n[步骤3] 获取现有MSKU数据...")
        existing_data = {}
        for doc in self.collection.find({}):
            msku = doc.get('msku')
            if msku:
                existing_data[msku] = doc
        print(f"  ✓ 数据库中现有 {len(existing_data)} 条MSKU记录")
        return existing_data
    
    def convert_row_to_doc(self, row):
        """将Excel行转换为MongoDB文档"""
        doc = {}
        
        for excel_col, mongo_field in self.FIELD_MAPPING.items():
            if excel_col in row.index:
                value = row[excel_col]
                
                # 处理NaN值
                if pd.isna(value):
                    doc[mongo_field] = ''
                else:
                    # 特殊处理weight字段
                    if mongo_field == 'weight':
                        try:
                            doc[mongo_field] = float(value)
                        except:
                            doc[mongo_field] = 0.0
                    # 特殊处理price字段
                    elif mongo_field == 'price':
                        doc[mongo_field] = str(value) if value != '' else ''
                    # 特殊处理HS编码（去除.0后缀）
                    elif mongo_field == 'HS':
                        hs_str = str(value).strip()
                        # 如果是浮点数格式，去掉.0
                        if hs_str.endswith('.0'):
                            hs_str = hs_str[:-2]
                        doc[mongo_field] = hs_str
                    else:
                        doc[mongo_field] = str(value).strip() if value != '' else ''
        
        return doc
    
    def compare_and_get_updates(self, excel_doc, existing_doc):
        """
        比较Excel数据与现有数据，返回需要更新的字段
        
        Args:
            excel_doc: 从Excel读取的数据
            existing_doc: MongoDB中现有的数据
        
        Returns:
            dict: 需要更新的字段，如果没有差异返回空字典
        """
        updates = {}
        
        # 只比较FIELD_MAPPING中定义的字段（排除msku）
        for mongo_field in self.FIELD_MAPPING.values():
            if mongo_field == 'msku':
                continue
            
            excel_value = excel_doc.get(mongo_field, '')
            existing_value = existing_doc.get(mongo_field, '')
            
            # 统一处理空值比较
            if excel_value is None:
                excel_value = ''
            if existing_value is None:
                existing_value = ''
            
            # 数值类型特殊处理
            if mongo_field == 'weight':
                try:
                    excel_float = float(excel_value) if excel_value != '' else 0.0
                    existing_float = float(existing_value) if existing_value != '' else 0.0
                    if abs(excel_float - existing_float) > 0.0001:
                        updates[mongo_field] = excel_float
                except (ValueError, TypeError):
                    if str(excel_value) != str(existing_value):
                        updates[mongo_field] = excel_value
            elif mongo_field == 'price':
                try:
                    excel_price = float(excel_value) if excel_value != '' else 0.0
                    existing_price = float(existing_value) if existing_value != '' else 0.0
                    if abs(excel_price - existing_price) > 0.001:
                        updates[mongo_field] = excel_value
                except (ValueError, TypeError):
                    if str(excel_value) != str(existing_value):
                        updates[mongo_field] = excel_value
            else:
                # 字符串比较
                if str(excel_value).strip() != str(existing_value).strip():
                    updates[mongo_field] = excel_value
        
        return updates
    
    def import_data(self, excel_file, dry_run=False):
        """
        导入数据到MongoDB
        
        Args:
            excel_file: Excel文件路径
            dry_run: 是否为试运行（只显示会做什么，不实际执行）
        """
        # 读取Excel
        df = self.read_excel(excel_file)
        if df is None:
            return
        
        # 连接数据库
        self.connect_db()
        
        # 获取现有数据
        existing_data = self.get_existing_data()
        
        # 统计信息
        stats = {
            'total': len(df),
            'new': 0,
            'updated': 0,
            'unchanged': 0,
            'errors': 0
        }
        
        new_documents = []
        updates_list = []
        
        print(f"\n[步骤4] 分析数据差异...")
        if dry_run:
            print("  ⚠ 试运行模式，不会实际修改数据库")
        
        # 处理每一行
        for idx, row in df.iterrows():
            msku = row.get('MSKU', '')
            if not msku or pd.isna(msku):
                print(f"  ⚠ 第{idx+1}行: MSKU为空，跳过")
                stats['errors'] += 1
                continue
            
            # 转换为MongoDB文档
            excel_doc = self.convert_row_to_doc(row)
            
            if msku in existing_data:
                # MSKU已存在，比较差异
                existing_doc = existing_data[msku]
                updates = self.compare_and_get_updates(excel_doc, existing_doc)
                
                if updates:
                    # 有差异，需要更新
                    updates_list.append((msku, updates))
                    stats['updated'] += 1
                    
                    # 显示差异详情
                    print(f"\n  [{idx+1}] MSKU: {msku} - 需要更新")
                    for field, new_value in updates.items():
                        old_value = existing_doc.get(field, '')
                        print(f"      {field}: '{old_value}' -> '{new_value}'")
                else:
                    # 无差异
                    stats['unchanged'] += 1
            else:
                # 新MSKU，需要插入
                new_documents.append(excel_doc)
                stats['new'] += 1
                print(f"  [{idx+1}] MSKU: {msku} - 新增")
        
        # 显示统计
        print("\n" + "=" * 70)
        print("数据分析完成")
        print("=" * 70)
        print(f"总记录数:       {stats['total']}")
        print(f"新增:           {stats['new']}")
        print(f"需要更新:       {stats['updated']}")
        print(f"无变化:         {stats['unchanged']}")
        print(f"错误:           {stats['errors']}")
        print("=" * 70)
        
        if dry_run:
            print("\n⚠ 试运行模式，未实际修改数据库")
            self.client.close()
            return stats
        
        # 执行插入
        if new_documents:
            print(f"\n[步骤5] 批量插入新数据...")
            print(f"  准备插入 {len(new_documents)} 条新记录...")
            
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            for doc in new_documents:
                doc['created_at'] = now
                doc['updated_at'] = now
                # 添加本地维护字段的默认值
                doc['image_url'] = ''
                doc['askPrice'] = ''
                doc['outboundFee'] = ''
                doc['putAwayFee'] = ''
                doc['X_ROW_K'] = ''
            
            try:
                result = self.collection.insert_many(new_documents)
                print(f"  ✓ 成功插入 {len(result.inserted_ids)} 条记录")
            except Exception as e:
                print(f"  ✗ 插入失败: {str(e)}")
        
        # 执行更新
        if updates_list:
            print(f"\n[步骤6] 批量更新现有数据...")
            print(f"  准备更新 {len(updates_list)} 条记录...")
            
            updated_count = 0
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            for msku, updates in updates_list:
                if updates:
                    # 添加updated_at时间戳
                    update_data = dict(updates)
                    update_data['updated_at'] = now
                    
                    try:
                        result = self.collection.update_one(
                            {'msku': msku},
                            {'$set': update_data}
                        )
                        
                        if result.matched_count > 0:
                            updated_count += 1
                            print(f"  ✓ 更新 {msku}: {list(updates.keys())}")
                    except Exception as e:
                        print(f"  ✗ 更新 {msku} 失败: {str(e)}")
            
            print(f"  ✓ 成功更新 {updated_count} 条记录")
        
        # 关闭数据库连接
        self.client.close()
        
        print("\n" + "=" * 70)
        print("导入完成！")
        print("=" * 70)
        
        return stats


def main():
    """主函数"""
    print("=" * 70)
    print("Excel数据导入MongoDB工具")
    print("=" * 70)
    
    # 获取Excel文件路径
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excel_file = input("\n请输入Excel文件路径: ").strip()
    
    if not excel_file:
        print("❌ 未指定Excel文件")
        return
    
    # 询问是否试运行
    print("\n请选择运行模式：")
    print("1. 试运行（只显示会做什么，不实际修改数据库）")
    print("2. 正式运行（实际修改数据库）")
    
    choice = input("\n请输入选项 (1/2): ").strip()
    dry_run = (choice == '1')
    
    if dry_run:
        print("\n✓ 选择试运行模式")
    else:
        print("\n⚠ 选择正式运行模式，将实际修改数据库")
        confirm = input("确认继续？(yes/no): ").strip().lower()
        if confirm not in ['yes', 'y']:
            print("❌ 已取消")
            return
    
    # 创建导入器
    importer = ExcelToMongoImporter()
    
    # 导入数据
    try:
        stats = importer.import_data(excel_file, dry_run=dry_run)
    except KeyboardInterrupt:
        print("\n\n⚠ 用户中断导入")
    except Exception as e:
        print(f"\n\n❌ 导入失败: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
