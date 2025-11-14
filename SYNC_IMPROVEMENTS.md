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
