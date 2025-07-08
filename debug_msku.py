#!/usr/bin/env python3
"""
MSKU查询诊断脚本
用于检查数据库中MSKU数据的情况
"""

from db_connector import MongoDBConnector

def check_msku_data():
    """检查MSKU数据"""
    # 这些是出现警告的MSKU
    problem_mskus = ['1809-Aa-0494-Dfba', '2203-Aa-0743-Beige-LC-DKfr']
    
    try:
        connector = MongoDBConnector()
        with connector as db:
            collection = db['msku_info']
            
            print("="*60)
            print("MSKU数据库查询诊断")
            print("="*60)
            
            # 1. 检查集合总数
            total_count = collection.count_documents({})
            print(f"📊 数据库中总共有 {total_count} 条MSKU记录")
            
            # 2. 查看几个示例记录
            print(f"\n📋 数据库中的示例记录:")
            sample_docs = list(collection.find().limit(3))
            for i, doc in enumerate(sample_docs, 1):
                print(f"  {i}. MSKU: {doc.get('msku', '未知')}")
                print(f"     产品名: {doc.get('productNameZh', '未知')}")
                
            # 3. 检查问题MSKU
            print(f"\n🔍 检查问题MSKU:")
            for msku in problem_mskus:
                print(f"\n  查询MSKU: {msku}")
                
                # 精确匹配
                exact_result = collection.find_one({'msku': msku})
                if exact_result:
                    print(f"    ✅ 精确匹配找到: {exact_result.get('productNameZh', '未知')}")
                else:
                    print(f"    ❌ 精确匹配未找到")
                    
                    # 尝试模糊匹配
                    fuzzy_results = list(collection.find({'msku': {'$regex': msku.replace('-', '.*'), '$options': 'i'}}).limit(5))
                    if fuzzy_results:
                        print(f"    🔍 相似的MSKU:")
                        for result in fuzzy_results:
                            print(f"      - {result.get('msku', '未知')}: {result.get('productNameZh', '未知')}")
                    else:
                        print(f"    🔍 没有找到相似的MSKU")
                        
            # 4. 检查MSKU字段的数据格式
            print(f"\n📝 MSKU字段格式分析:")
            pipeline = [
                {"$group": {"_id": None, "mskus": {"$push": "$msku"}}},
                {"$project": {"sample_mskus": {"$slice": ["$mskus", 10]}}}
            ]
            format_result = list(collection.aggregate(pipeline))
            if format_result and format_result[0].get('sample_mskus'):
                print("    前10个MSKU格式:")
                for msku in format_result[0]['sample_mskus']:
                    print(f"      - {msku}")
                    
    except Exception as e:
        print(f"❌ 检查失败: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_msku_data() 