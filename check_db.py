from db_utils import MongoDBClient

def check_database_structure():
    db_client = MongoDBClient()
    try:
        # 连接数据库
        db_client.connect()
        
        # 获取所有集合
        collections = db_client.db.list_collection_names()
        print("\n现有的集合:", collections)
        
        # 查看每个集合的数据结构
        for collection_name in collections:
            print(f"\n集合 '{collection_name}' 的数据示例:")
            # 获取一条示例数据
            sample = db_client.db[collection_name].find_one()
            if sample:
                print("数据结构:")
                for key, value in sample.items():
                    print(f"- {key}: {type(value).__name__} = {value}")
            else:
                print("集合为空")
                
            # 获取文档数量
            count = db_client.db[collection_name].count_documents({})
            print(f"总文档数: {count}")
            
    finally:
        db_client.close()

if __name__ == "__main__":
    check_database_structure()
