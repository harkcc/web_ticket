"""
数据库连接工具类
支持本地开发和服务器部署两种环境
"""
from sshtunnel import SSHTunnelForwarder
from pymongo import MongoClient
from db_config import SSH_CONFIG, MONGO_CONFIG, DEPLOY_ENV, get_mongo_config
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class MongoDBConnector:
    def __init__(self):
        self.tunnel = None
        self.client = None
        self.db = None
        self.config = get_mongo_config()

    def connect(self):
        try:
            if DEPLOY_ENV == 'development':
                # 开发环境：使用SSH隧道
                self._connect_via_ssh()
            else:
                # 生产环境：直接连接
                self._connect_direct()

            # 测试连接
            self.db.command('ping')
            logger.info(f"MongoDB connection established successfully in {DEPLOY_ENV} environment")
            return self.db

        except Exception as e:
            logger.error(f"Error connecting to database: {str(e)}")
            raise

    def _connect_via_ssh(self):
        """通过SSH隧道连接数据库（开发环境）"""
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

        # 构建MongoDB URI并连接
        self._connect_to_mongodb()

    def _connect_direct(self):
        """直接连接数据库（生产环境）"""
        self._connect_to_mongodb()

    def _connect_to_mongodb(self):
        """连接到MongoDB数据库"""
        if self.config['use_auth']:
            # 使用认证连接
            mongo_uri = (
                f"mongodb://{self.config['username']}:{self.config['password']}@"
                f"{self.config['host']}:{self.config['port']}/"
                f"{self.config['database']}?authSource={self.config['auth_source']}"
            )
        else:
            # 无认证连接
            mongo_uri = f"mongodb://{self.config['host']}:{self.config['port']}"
        
        # 连接MongoDB
        self.client = MongoClient(mongo_uri)
        self.db = self.client[self.config['database']]

    def close(self):
        """关闭数据库连接"""
        if self.client:
            self.client.close()
        if self.tunnel and DEPLOY_ENV == 'development':
            self.tunnel.close()
        logger.info("Database connection closed")

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

def test_connection():
    """测试数据库连接"""
    connector = MongoDBConnector()
    try:
        db = connector.connect()
        logger.info(f"Successfully connected to database in {DEPLOY_ENV} environment")
        # 测试集合访问
        collections = db.list_collection_names()
        logger.info(f"Available collections: {collections}")
    finally:
        connector.close()

if __name__ == "__main__":
    test_connection()
