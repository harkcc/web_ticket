"""
数据库配置文件
"""

# SSH配置
SSH_CONFIG = {
    'ssh_host': '59.110.218.83',  # SSH服务器地址
    'ssh_username': 'root',        # SSH用户名
    'ssh_password': 'Baitai2024',  # SSH密码
    'remote_bind_address': '127.0.0.1',  # MongoDB所在服务器地址
    'remote_bind_port': 27017,     # MongoDB端口
    'local_bind_address': '127.0.0.1',  # 本地绑定地址
    'local_bind_port': 27017       # 本地绑定端口
}

# MongoDB配置
MONGO_CONFIG = {
    'username': 'root',
    'password': 'baitai123456',  # 修改为连接字符串中的密码
    'database': 'SKU_INFO',
    'auth_source': 'admin',      # 添加认证数据库
    'collections': {
        'msku_info': 'msku_info',
        'ssh': 'ssh'
    }
}
