"""
数据库配置文件
支持本地开发和服务器部署两种环境
使用环境变量 DEPLOY_ENV 控制：
- DEPLOY_ENV=development: 本地开发环境（使用SSH隧道连接远程数据库）
- DEPLOY_ENV=production: 生产环境（直接连接本地数据库）
"""
import os

# 获取部署环境，默认为生产环境，便于上传部署，上传是记得切换
DEPLOY_ENV = os.getenv('DEPLOY_ENV', 'production')
# DEPLOY_ENV = os.getenv('DEPLOY_ENV', 'development')

# SSH配置（仅开发环境使用）
SSH_CONFIG = {
    'ssh_host': '59.110.218.83',  # SSH服务器地址
    'ssh_username': 'root',        # SSH用户名
    'ssh_password': 'Baitai2024',  # SSH密码
    'remote_bind_address': '127.0.0.1',  # MongoDB所在服务器地址
    'remote_bind_port': 27017,     # MongoDB端口
    'local_bind_address': '127.0.0.1',  # 本地绑定地址
    'local_bind_port': 27018       # 本地绑定端口
}

# MongoDB配置
MONGO_CONFIG = {
    # 开发环境配置（通过SSH隧道连接远程数据库）
    'development': {
        'username': 'root',
        'password': 'baitai123456',
        'host': '127.0.0.1',  # 通过SSH隧道访问
        'port': 27018,        # 使用SSH隧道转发的端口
        'database': 'SKU_INFO',
        'auth_source': 'admin',
        'use_auth': True,     # 使用认证
        'collections': {
            'msku_info': 'msku_info',
            'ssh': 'ssh'
        }
    },
    # 生产环境配置（直接连接本地MongoDB）
    'production': {
        'host': '127.0.0.1',     # 本地MongoDB地址
        'port': 27017,           # 默认MongoDB端口
        'database': 'SKU_INFO',
        'use_auth': False,       # 不使用认证
        'collections': {
            'msku_info': 'msku_info',
            'ssh': 'ssh'
        }
    }
}


# 获取当前环境的MongoDB配置
def get_mongo_config():
    return MONGO_CONFIG[DEPLOY_ENV]
