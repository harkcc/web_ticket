"""数据库连接管理模块"""

from pymongo import MongoClient
from config import MONGO_CONFIG
from contextlib import contextmanager

class DatabaseManager:
    """MongoDB数据库管理类"""
    
    def __init__(self):
        """初始化数据库管理器"""
        self.config = MONGO_CONFIG
        self._client = None

    @property
    def client(self):
        """懒加载MongoDB客户端连接"""
        if self._client is None:
            self._client = MongoClient(
                host=self.config['host'],
                port=self.config['port'],
                username=self.config.get('username'),
                password=self.config.get('password'),
                authSource=self.config.get('authentication_source', 'admin')
            )
        return self._client

    @contextmanager
    def get_connection(self):
        """获取数据库连接
        
        使用示例:
        with db_manager.get_connection() as db:
            collection = db['collection_name']
            result = collection.find_one({"field": "value"})
        """
        try:
            db = self.client[self.config['database']]
            yield db
        except Exception as e:
            print(f"MongoDB连接错误: {e}")
            raise
        
class ProductRepository:
    """产品数据仓库"""
    
    def __init__(self):
        """初始化产品仓库"""
        self.db = DatabaseManager()
        self.collection_name = 'products'  # MongoDB集合名称

    def get_product_info(self, sku):
        """根据SKU获取产品信息
        
        Args:
            sku: 产品SKU
            
        Returns:
            dict: 产品信息字典，包含 msku, product_name, fnsku 等信息
        """
        with self.db.get_connection() as db:
            collection = db[self.collection_name]
            return collection.find_one({"sku": sku})

    def get_products_by_skus(self, skus):
        """批量获取产品信息
        
        Args:
            skus: SKU列表
            
        Returns:
            dict: {sku: product_info} 格式的字典
        """
        with self.db.get_connection() as db:
            collection = db[self.collection_name]
            products = collection.find({"sku": {"$in": skus}})
            return {product["sku"]: product for product in products}

    # def insert_product(self, product_data):
    #     """插入新产品
        
    #     Args:
    #         product_data: 产品数据字典
    #     """
    #     with self.db.get_connection() as db:
    #         collection = db[self.collection_name]
    #         return collection.insert_one(product_data)

    # def update_product(self, sku, update_data):
    #     """更新产品信息
        
    #     Args:
    #         sku: 产品SKU
    #         update_data: 要更新的数据字典
    #     """
    #     with self.db.get_connection() as db:
    #         collection = db[self.collection_name]
    #         return collection.update_one(
    #             {"sku": sku},
    #             {"$set": update_data}
    #         )


