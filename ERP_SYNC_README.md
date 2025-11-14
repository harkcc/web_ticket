# ERP数据同步功能说明

## 功能概述

在web_ticket系统中添加了"从ERP更新数据库"按钮，可以一键从领星ERP同步产品数据到MongoDB。

## 使用方式

### 1. 启动Web服务
```bash
cd /Users/chenminghui/py/lingxing_request/web_ticket
python3 web_ticket.py
```

### 2. 访问页面
打开浏览器访问：`http://localhost:5009`

### 3. 点击按钮
在首页顶部找到"从ERP更新数据库"按钮，点击即可开始同步。

## 功能特点

### ✅ 后台异步执行
- 点击按钮后立即返回，不阻塞页面
- 在后台线程中执行同步任务
- 实时显示同步进度

### ✅ 实时状态更新
- 登录ERP系统...
- 连接数据库...
- 获取产品数据...
- 同步完成！

### ✅ 完整的统计信息
同步完成后显示：
- 插入记录数
- 跳过记录数（已存在）
- 错误数量
- 总产品数
- 总MSKU数

### ✅ 安全机制
- 点击前需要确认
- 自动检查MSKU是否已存在
- 已存在的记录会跳过，不会重复插入
- 线程安全的数据库操作

## 技术实现

### 前端（index.html）
```javascript
// 按钮点击事件
document.getElementById('syncDatabaseBtn').addEventListener('click', async function() {
    // 1. 确认对话框
    // 2. 发送POST请求到 /sync_erp_database
    // 3. 获取task_id
    // 4. 轮询检查状态 /sync_status/{task_id}
    // 5. 显示结果通知
});
```

### 后端（web_ticket.py）
```python
@app.route('/sync_erp_database', methods=['POST'])
def sync_erp_database():
    # 1. 创建任务ID
    # 2. 初始化任务状态
    # 3. 启动后台线程
    # 4. 返回task_id

def sync_erp_data_task(task_id):
    # 1. 登录ERP
    # 2. 创建同步对象
    # 3. 执行同步
    # 4. 更新任务状态
```

### 同步模块（sync_erp_to_mongodb.py）
```python
class ERPProductSync:
    def sync_products(self, limit=None):
        # 1. 获取产品列表
        # 2. 获取现有MSKU
        # 3. 处理每个产品
        # 4. 批量插入数据
        # 5. 返回统计信息
```

## 数据流程

```
用户点击按钮
    ↓
前端发送请求
    ↓
后端创建任务
    ↓
后台线程执行
    ├─ 登录ERP
    ├─ 获取产品列表（已配对）
    ├─ 获取产品详情
    ├─ 获取MSKU链接
    ├─ 转换数据格式
    └─ 批量插入MongoDB
    ↓
前端轮询状态
    ↓
显示完成通知
```

## 字段映射

| MongoDB字段 | ERP字段 | 说明 |
|------------|---------|------|
| msku | MSKU | 主键 |
| productNameZh | customs_export_name | 中文产品名 |
| productNameEn | customs_import_name | 英文产品名 |
| price | customs_import_price | 价格 |
| brand | brand_name | 品牌（空则"无"） |
| model | model | 型号（空则"无"） |
| HS | customs_declaration_hs_code | HS编码 |
| asin | asin | ASIN |
| electrified | special_attr | 是否有电 |
| magnetic | special_attr | 是否带磁 |
| materialZh/En | customs_clearance_material | 材质 |
| useZh/En | customs_clearance_usage | 用途 |
| weight | cg_product_net_weight | 重量 |
| productLink | asin_url | 产品链接 |

## 注意事项

1. **首次同步**：可能需要较长时间（取决于产品数量）
2. **增量同步**：已存在的MSKU会自动跳过
3. **网络要求**：需要能访问领星ERP API
4. **数据库环境**：
   - development: 通过SSH隧道连接远程MongoDB
   - production: 直接连接本地MongoDB

## 文件说明

- `web_ticket.py` - Web服务主文件（添加了同步路由）
- `sync_erp_to_mongodb.py` - ERP同步核心模块
- `templates/index.html` - 前端页面（添加了同步按钮）
- `login.py` - ERP登录模块
- `db_utils.py` - MongoDB连接工具
- `db_config.py` - 数据库配置

## 故障排查

### 问题：按钮点击无反应
- 检查浏览器控制台是否有错误
- 检查网络请求是否成功

### 问题：同步失败
- 检查ERP登录是否正常
- 检查数据库连接是否正常
- 查看服务器日志获取详细错误信息

### 问题：数据未插入
- 检查MSKU是否已存在（会自动跳过）
- 检查数据格式是否正确
- 查看统计信息中的错误数量
