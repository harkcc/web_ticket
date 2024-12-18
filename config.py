"""数据库配置文件"""

# MongoDB配置
MONGO_CONFIG = {
    'host': 'localhost',     # MongoDB服务器地址
    'port': 27017,          # MongoDB默认端口
    'database': 'lingxing_db',  # 数据库名称
    'username': 'your_username',  # 如果需要认证则配置
    'password': 'your_password',  # 如果需要认证则配置
    'authentication_source': 'admin'  # 认证数据库，默认是admin
}

# 其他配置
OTHER_CONFIG = {
    'debug': False,  # 生产环境关闭调试
    'log_level': 'INFO'  # 日志级别
}
