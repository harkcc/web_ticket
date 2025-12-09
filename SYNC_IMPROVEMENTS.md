# ERP产品同步脚本改进说明

## 参考web_ticket.py的实现方式

### ✅ 已实现的改进

#### 1. **线程安全 - 数据库操作锁**
```python
# 添加全局锁（与web_ticket保持一致）
db_operation_lock = threading.Lock()

# 在数据库操作时使用锁
with db_operation_lock:
    existing_mskus = set(doc['msku'] for doc in self.collection.find({}, {'msku': 1}))
```

#### 2. **批量插入 - 性能优化**
```python
# 改进前：逐条插入
for document in documents:
    self.collection.insert_one(document)  # N次数据库操作

# 改进后：批量插入
self.collection.insert_many(documents_to_insert)  # 1次数据库操作
```

#### 3. **二次检查 - 防止并发重复**
```python
# 步骤1：预先获取现有MSKU
existing_mskus = self.check_existing_mskus()

# 步骤2：处理时跳过已存在的
if msku in existing_mskus:
    skip...

# 步骤3：插入前再次检查（防止并发插入）
with db_operation_lock:
    current_mskus = set(doc['msku'] for doc in self.collection.find({}, {'msku': 1}))
    documents_to_insert = [doc for doc in documents if doc.get('msku') not in current_mskus]
```

#### 4. **资源管理 - finally块关闭连接**
```python
try:
    # 处理数据
    ...
finally:
    # 确保关闭数据库连接
    if hasattr(self, 'db_client'):
        self.db_client.close()
```

#### 5. **处理流程优化**
```python
# web_ticket的流程：
# 1. 连接数据库
# 2. 获取现有MSKU列表
# 3. 预处理数据（收集待插入文档）
# 4. 批量插入
# 5. 关闭连接

# 我们的实现完全遵循这个流程
```

### 📊 性能对比

| 指标 | 改进前 | 改进后 |
|------|--------|--------|
| 数据库操作次数 | N次（每个MSKU一次） | 2次（查询1次+插入1次） |
| 并发安全 | ❌ 无锁保护 | ✅ 线程锁保护 |
| 重复检查 | ❌ 单次检查 | ✅ 双重检查 |
| 资源释放 | ⚠️ 依赖__del__ | ✅ finally保证 |

### 🔍 与web_ticket的一致性

#### 相同点：
- ✅ 使用 `db_operation_lock` 线程锁
- ✅ 使用 `db_client.db['msku_info']` 访问集合
- ✅ 使用 `insert_many` 批量插入
- ✅ 二次检查防止重复
- ✅ finally块关闭连接

#### 差异点：
- web_ticket: 从Excel读取数据
- 我们的脚本: 从ERP API获取数据

但数据库操作逻辑完全一致！

### 🎯 使用方式

```bash
cd /Users/chenminghui/py/lingxing_request/web_ticket
python3 sync_erp_to_mongodb.py
```

选择模式：
1. 测试模式（10个产品）
2. 正式模式（所有产品）

### 📝 注意事项

1. **一对多处理**：1个SKU → 多个MSKU → 每个MSKU独立记录 ✅
2. **字段映射**：完全按照需求实现 ✅
3. **数据转换**：支持两种格式（分开/合并） ✅
4. **重量字段**：从spec_info.cg_product_net_weight获取 ✅
5. **线程安全**：使用锁保护数据库操作 ✅
6. **批量插入**：提高性能 ✅

---

## 2024-12 全量同步更新

### 🔄 同步策略变更

**旧策略**：仅插入（补数据）
- 只检查MSKU是否存在
- 已存在的MSKU直接跳过
- 不更新已有数据

**新策略**：全量同步（插入+更新）
- 新MSKU：插入新记录
- 已存在MSKU：对比差异，以ERP为准更新
- 本地字段不会被覆盖

### 📋 字段分类

#### ERP同步字段（以ERP为准）
```python
ERP_SYNC_FIELDS = [
    'productNameZh', 'productNameEn', 'price', 'brand', 'model', 'HS',
    'asin', 'electrified', 'magnetic', 'materialEn', 'materialZh',
    'productLink', 'useEn', 'useZh', 'weight'
]
```

#### 本地维护字段（不会被ERP覆盖）
```python
LOCAL_FIELDS = ['image_url', 'askPrice', 'outboundFee', 'putAwayFee', 'X_ROW_K', 'created_at']
```

### ⏰ 时间戳字段

| 字段 | 新增记录 | 更新记录 |
|------|----------|----------|
| `created_at` | 设置为当前时间 | 保持原值不变 |
| `updated_at` | 设置为当前时间 | 更新为当前时间 |

### 📊 新增统计信息

```
同步完成！统计信息：
======================================================================
总产品数:       100
处理产品数:     100
总MSKU数:       250
  - 新增:       10
  - 更新:       50
  - 无变化:     190
错误数量:       0
======================================================================
```

### 🔧 核心方法

1. **`get_existing_data()`**：获取现有MSKU完整数据
2. **`compare_and_get_updates()`**：对比ERP数据与现有数据
3. **`batch_update_documents()`**：批量更新文档
4. **`process_single_product_sync()`**：处理单个产品（支持插入和更新）
5. **`sync_products()`**：全量同步主函数

### 💡 设计优势

1. **数据安全**：本地维护的字段永远不会被意外清空
2. **性能更好**：只更新有变化的字段，减少数据库写入
3. **可追溯性**：`updated_at`记录最后同步时间
4. **增量同步友好**：后续ERP只推送变更字段时，逻辑天然兼容
5. **扩展性好**：将来新增本地字段，不需要修改同步逻辑
