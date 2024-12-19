"""
数据库连接工具类
"""
from sshtunnel import SSHTunnelForwarder
from pymongo import MongoClient
from db_config import SSH_CONFIG, MONGO_CONFIG
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class MongoDBConnector:
    def __init__(self):
        self.tunnel = None
        self.client = None
        self.db = None

    def connect(self):
        try:
            # 创建SSH隧道
            self.tunnel = SSHTunnelForwarder(
                ssh_address_or_host=(SSH_CONFIG['ssh_host'], 22),
                ssh_username=SSH_CONFIG['ssh_username'],
                ssh_password=SSH_CONFIG['ssh_password'],
                remote_bind_address=(SSH_CONFIG['remote_bind_address'], SSH_CONFIG['remote_bind_port']),
                local_bind_address=(SSH_CONFIG['local_bind_address'], SSH_CONFIG['local_bind_port'])
            )
            
            # 启动SSH隧道
            self.tunnel.start()
            logger.info("SSH tunnel established successfully")

            # 构建MongoDB URI
            mongo_uri = (
                f"mongodb://{MONGO_CONFIG['username']}:{MONGO_CONFIG['password']}@"
                f"{SSH_CONFIG['local_bind_address']}:{self.tunnel.local_bind_port}/"
                f"{MONGO_CONFIG['database']}?authSource={MONGO_CONFIG['auth_source']}"
            )
            
            # 连接MongoDB
            self.client = MongoClient(mongo_uri)
            self.db = self.client[MONGO_CONFIG['database']]
            
            # 测试连接
            self.db.command('ping')
            logger.info("MongoDB connection established successfully")
            
            return self.db
            
        except Exception as e:
            logger.error(f"Error connecting to database: {str(e)}")
            self.close()
            raise

    def close(self):
        if self.client:
            self.client.close()
            logger.info("MongoDB connection closed")
        if self.tunnel and self.tunnel.is_active:
            self.tunnel.close()
            logger.info("SSH tunnel closed")

    def __enter__(self):
        return self.connect()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def print_collection_info(self):
        # 获取 msku_info 集合
        msku_collection = self.db[MONGO_CONFIG['collections']['msku_info']]
        
        # 获取文档数量
        doc_count = msku_collection.count_documents({})
        logger.info(f"msku_info集合中的文档数量: {doc_count}")
        
        # 打印所有文档
        for doc in msku_collection.find():
            logger.info(doc)

# 使用示例
def test_connection():
    try:
        connector = MongoDBConnector()
        connector.connect()
        # 调用打印集合信息的方法
        connector.print_collection_info()
        connector.close()
    except Exception as e:
        logger.error(f"Test connection failed: {str(e)}")

if __name__ == "__main__":
    test_connection()
