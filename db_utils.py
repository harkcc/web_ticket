from pymongo import MongoClient
from sshtunnel import SSHTunnelForwarder
from .db_config import get_mongo_config, DEPLOY_ENV, SSH_CONFIG

class MongoDBClient:
    def __init__(self):
        self.config = get_mongo_config()
        self.client = None
        self.db = None
        self.tunnel = None
        
    def connect(self):
        """连接到MongoDB数据库"""
        if DEPLOY_ENV == 'development':
            # 开发环境：通过SSH隧道连接
            self.tunnel = SSHTunnelForwarder(
                (SSH_CONFIG['ssh_host'], 22),
                ssh_username=SSH_CONFIG['ssh_username'],
                ssh_password=SSH_CONFIG['ssh_password'],
                remote_bind_address=(SSH_CONFIG['remote_bind_address'], SSH_CONFIG['remote_bind_port']),
                local_bind_address=(SSH_CONFIG['local_bind_address'], SSH_CONFIG['local_bind_port'])
            )
            self.tunnel.start()
            
        # 构建MongoDB连接URI
        if self.config.get('use_auth'):
            uri = f"mongodb://{self.config['username']}:{self.config['password']}@{self.config['host']}:{self.config['port']}/{self.config['database']}?authSource={self.config['auth_source']}"
        else:
            uri = f"mongodb://{self.config['host']}:{self.config['port']}/{self.config['database']}"
            
        self.client = MongoClient(uri)
        self.db = self.client[self.config['database']]
        
    def close(self):
        """关闭数据库连接"""
        if self.client:
            self.client.close()
        if self.tunnel and self.tunnel.is_active:
            self.tunnel.stop()
            
    def insert_one(self, collection_name: str, data: dict):
        """插入单条数据
        
        Args:
            collection_name: 集合名称
            data: 要插入的数据字典
        """
        collection = self.db[collection_name]
        return collection.insert_one(data)
        
    def insert_many(self, collection_name: str, data_list: list):
        """插入多条数据
        
        Args:
            collection_name: 集合名称
            data_list: 要插入的数据列表
        """
        collection = self.db[collection_name]
        return collection.insert_many(data_list)
