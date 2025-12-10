#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修复数据库中HS编码的.0后缀问题
"""

from db_utils import MongoDBClient
from datetime import datetime


def fix_hs_codes():
    """修复所有HS编码，去除.0后缀"""
    
    print("=" * 70)
    print("修复HS编码工具")
    print("=" * 70)
    
    # 连接数据库
    print("\n[步骤1] 连接数据库...")
    client = MongoDBClient()
    client.connect()
    collection = client.db['msku_info']
    print("  ✓ 连接成功")
    
    # 查找所有包含.0的HS编码
    print("\n[步骤2] 查找需要修复的HS编码...")
    
    # 使用正则表达式查找以.0结尾的HS编码
    docs_to_fix = list(collection.find({
        'HS': {'$regex': r'\.0$'}
    }))
    
    print(f"  ✓ 找到 {len(docs_to_fix)} 条需要修复的记录")
    
    if len(docs_to_fix) == 0:
        print("\n✓ 没有需要修复的记录")
        client.close()
        return
    
    # 显示前5个示例
    print("\n[示例] 前5条需要修复的记录:")
    for i, doc in enumerate(docs_to_fix[:5], 1):
        msku = doc.get('msku', 'Unknown')
        old_hs = doc.get('HS', '')
        new_hs = old_hs[:-2] if old_hs.endswith('.0') else old_hs
        print(f"  {i}. MSKU: {msku}")
        print(f"     HS: '{old_hs}' -> '{new_hs}'")
    
    if len(docs_to_fix) > 5:
        print(f"  ... 还有 {len(docs_to_fix) - 5} 条记录")
    
    # 询问是否继续
    print("\n" + "=" * 70)
    confirm = input(f"确认修复这 {len(docs_to_fix)} 条记录？(yes/no): ").strip().lower()
    
    if confirm not in ['yes', 'y']:
        print("❌ 已取消")
        client.close()
        return
    
    # 执行修复
    print("\n[步骤3] 执行修复...")
    fixed_count = 0
    error_count = 0
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    for doc in docs_to_fix:
        msku = doc.get('msku')
        old_hs = doc.get('HS', '')
        
        if old_hs.endswith('.0'):
            new_hs = old_hs[:-2]  # 去掉最后的 .0（2个字符：点和零）
            
            try:
                result = collection.update_one(
                    {'msku': msku},
                    {
                        '$set': {
                            'HS': new_hs,
                            'updated_at': now
                        }
                    }
                )
                
                if result.modified_count > 0:
                    fixed_count += 1
                    if fixed_count <= 10:  # 只显示前10条
                        print(f"  ✓ {msku}: '{old_hs}' -> '{new_hs}'")
                    elif fixed_count == 11:
                        print(f"  ... 继续修复中 ...")
            except Exception as e:
                error_count += 1
                print(f"  ✗ {msku}: 修复失败 - {str(e)}")
    
    # 关闭连接
    client.close()
    
    # 显示结果
    print("\n" + "=" * 70)
    print("修复完成！")
    print("=" * 70)
    print(f"总记录数:       {len(docs_to_fix)}")
    print(f"成功修复:       {fixed_count}")
    print(f"失败数量:       {error_count}")
    print("=" * 70)


if __name__ == "__main__":
    try:
        fix_hs_codes()
    except KeyboardInterrupt:
        print("\n\n⚠ 用户中断")
    except Exception as e:
        print(f"\n\n❌ 修复失败: {str(e)}")
        import traceback
        traceback.print_exc()
