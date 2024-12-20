import pymongo
from sshtunnel import SSHTunnelForwarder
from db_config import SSH_CONFIG, MONGO_CONFIG

def test_mongodb_connection():
    """测试MongoDB连接"""
    try:
        # 创建SSH隧道
        with SSHTunnelForwarder(
            (SSH_CONFIG['ssh_host'], 22),
            ssh_username=SSH_CONFIG['ssh_username'],
            ssh_password=SSH_CONFIG['ssh_password'],
            remote_bind_address=(SSH_CONFIG['remote_bind_address'], SSH_CONFIG['remote_bind_port']),
            local_bind_address=(SSH_CONFIG['local_bind_address'], SSH_CONFIG['local_bind_port'])
        ) as tunnel:
            print("SSH隧道已建立")
            
            # 构建MongoDB URI
            mongo_uri = (
                f"mongodb://{MONGO_CONFIG['username']}:{MONGO_CONFIG['password']}@"
                f"{SSH_CONFIG['local_bind_address']}:{tunnel.local_bind_port}/"
                f"{MONGO_CONFIG['database']}?authSource={MONGO_CONFIG['auth_source']}"
            )
            
            # 连接MongoDB
            client = pymongo.MongoClient(mongo_uri)
            
            # 测试连接
            db = client[MONGO_CONFIG['database']]
            collections = db.list_collection_names()
            print("MongoDB连接成功!")
            print("数据库中的集合:", collections)
            
            # 测试查询
            msku_collection = db[MONGO_CONFIG['collections']['msku_info']]
            doc_count = msku_collection.count_documents({})
            print(f"msku_info集合中的文档数量: {doc_count}")
            
            # 尝试获取一条数据作为示例
            sample_doc = msku_collection.find_one({})
            if sample_doc:
                print("\n示例数据:")
                print(sample_doc)
            
            # 关闭连接
            client.close()
            return True
            
    except Exception as e:
        print("MongoDB连接失败:", str(e))
        return False

if __name__ == '__main__':
    print("开始测试MongoDB连接...")
    test_mongodb_connection()
