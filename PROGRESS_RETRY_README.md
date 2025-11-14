# 进度条和重试机制说明

## 🎯 新增功能

### **1. 进度条显示** 📊

#### 功能说明
在同步过程中实时显示进度条，让你清楚地看到当前进度。

#### 显示位置
- **预检查阶段**: 检查产品MSKU时显示进度
- **处理阶段**: 处理产品详情时显示进度

#### 进度条样式
```
预检查进度 |████████████████████████████----------| 65.3% (653/1000)
处理进度 |██████████████████████████████████████| 100.0% SKU-001 ✓ 新增2个MSKU
```

#### 技术实现
```python
def print_progress_bar(self, current, total, prefix='', suffix='', length=50):
    """打印进度条"""
    with progress_lock:
        percent = 100 * (current / float(total))
        filled_length = int(length * current // total)
        bar = '█' * filled_length + '-' * (length - filled_length)
        sys.stdout.write(f'\r{prefix} |{bar}| {percent:.1f}% {suffix}')
        sys.stdout.flush()
        if current == total:
            print()  # 完成后换行
```

---

### **2. 网络请求重试机制** 🔄

#### 功能说明
当网络请求失败时，自动重试最多3次，避免因临时网络问题导致同步失败。

#### 重试策略
- **最大重试次数**: 3次（可配置）
- **等待时间**: 递增等待（2秒、4秒、6秒）
- **适用范围**: 所有ERP API请求

#### 重试流程
```
第1次请求失败 → 等待2秒 → 第2次重试
第2次请求失败 → 等待4秒 → 第3次重试
第3次请求失败 → 等待6秒 → 第4次重试
第4次请求失败 → 报错退出
```

#### 日志输出
```
⚠ 网络错误，2秒后重试 (1/3): Connection timeout
⚠ 网络错误，4秒后重试 (2/3): Connection timeout
✗ 重试3次后仍失败: Connection timeout
```

#### 技术实现
```python
def request_with_retry(self, request_func, *args, **kwargs):
    """带重试机制的请求"""
    for attempt in range(self.max_retries):
        try:
            result = request_func(*args, **kwargs)
            return result
        except requests.exceptions.RequestException as e:
            if attempt < self.max_retries - 1:
                wait_time = (attempt + 1) * 2  # 递增等待
                print(f"\n  ⚠ 网络错误，{wait_time}秒后重试 ({attempt + 1}/{self.max_retries}): {str(e)}")
                time.sleep(wait_time)
            else:
                print(f"\n  ✗ 重试{self.max_retries}次后仍失败: {str(e)}")
                raise
    return None
```

---

## 🔧 使用方法

### **默认配置**
```python
# 自动使用3次重试
sync = ERPProductSync(token)
stats = sync.sync_products(limit=None)
```

### **自定义重试次数**
```python
# 设置5次重试
sync = ERPProductSync(token, max_retries=5)
stats = sync.sync_products(limit=None)
```

### **禁用重试（不推荐）**
```python
# 设置1次（不重试）
sync = ERPProductSync(token, max_retries=1)
stats = sync.sync_products(limit=None)
```

---

## 📊 进度条效果演示

### **预检查阶段**
```
[步骤2.5] 预检查产品MSKU...
预检查进度 |████████████----------| 40.0% (400/1000)
```

### **并发处理阶段**
```
[步骤3] 开始处理 100 个产品...
  使用 5 个线程并发处理...
处理进度 |████████████████------| 75.0% SKU-2503 ✓ 新增2个MSKU
```

### **串行处理阶段**
```
[步骤3] 开始处理 100 个产品...

--- [1/100] 处理产品: SKU-001 ---
  ✓ 新增2个MSKU, 跳过0个

--- [2/100] 处理产品: SKU-002 ---
  ✓ 新增1个MSKU, 跳过1个
```

---

## 🔍 重试机制详解

### **适用的API请求**
1. ✅ 获取产品列表 (`get_product_list`)
2. ✅ 获取产品详情 (`get_product_detail`)
3. ✅ 获取产品链接 (`get_product_links`)

### **捕获的异常类型**
- `requests.exceptions.RequestException` - 网络相关错误
  - `ConnectionError` - 连接错误
  - `Timeout` - 超时错误
  - `HTTPError` - HTTP错误
  - `TooManyRedirects` - 重定向过多

### **不会重试的情况**
- 业务逻辑错误（如：产品不存在）
- 数据格式错误
- 认证失败（token无效）

---

## ⚙️ 配置参数

### **初始化参数**
```python
ERPProductSync(
    token,              # ERP认证token（必填）
    max_retries=3       # 最大重试次数（默认3）
)
```

### **同步参数**
```python
sync.sync_products(
    limit=None,         # 限制产品数量
    use_concurrent=True,# 是否使用并发
    max_workers=5       # 并发线程数
)
```

---

## 📈 性能影响

### **进度条**
- **CPU开销**: 极小（<0.1%）
- **内存开销**: 可忽略
- **显示延迟**: 无

### **重试机制**
- **成功情况**: 无额外开销
- **失败情况**: 增加等待时间（2-6秒/次）
- **最坏情况**: 增加12秒（3次重试）

---

## 🎯 最佳实践

### **网络稳定时**
```python
# 使用默认配置
sync = ERPProductSync(token)
stats = sync.sync_products(limit=None)
```

### **网络不稳定时**
```python
# 增加重试次数
sync = ERPProductSync(token, max_retries=5)
stats = sync.sync_products(limit=None, max_workers=3)
```

### **调试模式**
```python
# 减少重试次数，快速失败
sync = ERPProductSync(token, max_retries=1)
stats = sync.sync_products(limit=10, use_concurrent=False)
```

---

## 🔧 线程安全

### **进度条锁**
```python
# 进度条使用独立的锁，避免并发显示混乱
progress_lock = threading.Lock()

with progress_lock:
    sys.stdout.write(f'\r{prefix} |{bar}| {percent:.1f}% {suffix}')
    sys.stdout.flush()
```

### **数据库操作锁**
```python
# 数据库操作使用独立的锁，保证数据一致性
db_operation_lock = threading.Lock()

with db_operation_lock:
    self.collection.insert_many(documents_to_insert)
```

---

## 📝 日志示例

### **完整同步日志**
```
======================================================================
开始同步ERP产品数据到MongoDB
✓ 使用并发模式 (最大5个线程)
======================================================================

[步骤1] 获取产品列表...
  正在获取产品列表，offset=0...
    获取到 500 条产品
  共获取 1000 条产品
  ✓ 已配对产品: 1000 个

[步骤2] 获取现有MSKU列表...
  数据库中已有 5000 个MSKU

[步骤2.5] 预检查产品MSKU...
预检查进度 |██████████████████████████████████████| 100.0% (1000/1000)
  ✓ 需要处理: 100 个产品
  ⊗ 跳过(全部MSKU已存在): 900 个产品

[步骤3] 开始处理 100 个产品...
  使用 5 个线程并发处理...
处理进度 |██████████████████████████████████████| 100.0% SKU-2503 ✓ 新增2个MSKU

[步骤4] 批量插入数据...
  准备插入 150 条记录...
  ✓ 成功插入 150 条记录

======================================================================
同步完成！统计信息：
======================================================================
总产品数:       1000
跳过产品数:     900 (全部MSKU已存在)
处理产品数:     100
总MSKU数:       200
新增MSKU:       150
跳过MSKU:       50 (已存在)
错误数量:       0
======================================================================

✓ 数据库连接已关闭
```

---

## 🚀 优势总结

### **进度条**
1. ✅ 实时反馈 - 随时了解同步进度
2. ✅ 预估时间 - 根据进度估算剩余时间
3. ✅ 问题定位 - 快速发现卡住的位置
4. ✅ 用户体验 - 不再盲目等待

### **重试机制**
1. ✅ 提高成功率 - 临时网络问题自动恢复
2. ✅ 减少人工干预 - 自动重试无需手动重启
3. ✅ 智能等待 - 递增等待时间避免过度请求
4. ✅ 详细日志 - 清楚记录重试过程

---

## 🔮 未来优化

1. **动态调整重试间隔** - 根据错误类型调整等待时间
2. **断点续传** - 记录失败位置，支持继续同步
3. **进度持久化** - 保存进度到文件，重启后恢复
4. **智能限流** - 根据服务器响应自动调整请求速度
5. **彩色进度条** - 使用颜色区分成功/失败/警告
