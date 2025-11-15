# 材质和用途字段解析逻辑说明

## 📋 三种数据情况

### **情况1: 分开存储（4个独立字段）**

ERP返回的数据：
```json
{
  "product_clearance_list": {
    "customs_clearance_material": "聚酯纤维",
    "customs_clearance_en_material": "polyester",
    "customs_clearance_usage": "保暖",
    "customs_clearance_en_usage": "Keep Warm"
  }
}
```

**处理逻辑：**
```python
# split_field 函数会检测到有独立的英文字段
material_zh, material_en = split_field("聚酯纤维", "polyester")
# 返回: ("聚酯纤维", "polyester")

use_zh, use_en = split_field("保暖", "Keep Warm")
# 返回: ("保暖", "Keep Warm")
```

**存入数据库：**
```json
{
  "materialZh": "聚酯纤维",
  "materialEn": "polyester",
  "useZh": "保暖",
  "useEn": "Keep Warm"
}
```

---

### **情况2: 复合存储（斜杠分割）**

ERP返回的数据：
```json
{
  "product_clearance_list": {
    "customs_clearance_material": "聚酯纤维/polyester",
    "customs_clearance_en_material": "",
    "customs_clearance_usage": "保暖/Keep Warm",
    "customs_clearance_en_usage": ""
  }
}
```

或者JSON转义的斜杠：
```json
{
  "product_clearance_list": {
    "customs_clearance_material": "聚酯纤维\\/polyester",
    "customs_clearance_en_material": "",
    "customs_clearance_usage": "保暖\\/Keep Warm",
    "customs_clearance_en_usage": ""
  }
}
```

**处理逻辑：**
```python
# JSON中的 \/ 会被自动解析为 /
# split_field 函数会检测到斜杠，进行拆分
material_zh, material_en = split_field("聚酯纤维/polyester", "")
# 返回: ("聚酯纤维", "polyester")

use_zh, use_en = split_field("保暖/Keep Warm", "")
# 返回: ("保暖", "Keep Warm")
```

**存入数据库：**
```json
{
  "materialZh": "聚酯纤维",
  "materialEn": "polyester",
  "useZh": "保暖",
  "useEn": "Keep Warm"
}
```

---

### **情况3: 真实缺失（字段为空）**

ERP返回的数据：
```json
{
  "product_clearance_list": {
    "customs_clearance_material": "",
    "customs_clearance_en_material": "",
    "customs_clearance_usage": "",
    "customs_clearance_en_usage": ""
  }
}
```

或者字段不存在：
```json
{
  "product_clearance_list": {}
}
```

**处理逻辑：**
```python
# split_field 函数会检测到字段为空
material_zh, material_en = split_field("", "")
# 返回: ("", "")

use_zh, use_en = split_field("", "")
# 返回: ("", "")
```

**存入数据库：**
```json
{
  "materialZh": "",
  "materialEn": "",
  "useZh": "",
  "useEn": ""
}
```

---

## 🔧 核心函数：`split_field`

```python
def split_field(self, field_value, en_field_value=None):
    """
    拆分中英文字段，支持三种情况：
    1. 分开存储：customs_clearance_material + customs_clearance_en_material
    2. 复合存储：customs_clearance_material = "中文/English" (斜杠分割)
    3. 真实缺失：字段为空或不存在
    
    Args:
        field_value: 中文字段值（可能包含英文）
        en_field_value: 英文字段值（如果分开存储）
    
    Returns:
        tuple: (中文, 英文)
    """
    # 情况1: 分开存储 - 如果有独立的英文字段，优先使用
    if en_field_value and en_field_value.strip():
        zh = field_value.strip() if field_value else ""
        en = en_field_value.strip()
        return zh, en
    
    # 情况3: 真实缺失 - 字段为空
    if not field_value or field_value.strip() == "":
        return "", ""
    
    # 情况2: 复合存储 - 包含斜杠分隔符
    if '/' in field_value:
        parts = field_value.split('/', 1)  # 只分割第一个斜杠
        zh = parts[0].strip() if len(parts) > 0 else ""
        en = parts[1].strip() if len(parts) > 1 else ""
        return zh, en
    
    # 其他情况: 只有中文，没有英文
    return field_value.strip(), ""
```

---

## 🔍 处理优先级

函数按以下优先级处理：

1. **优先级1**: 检查是否有独立的英文字段 → 情况1（分开存储）
2. **优先级2**: 检查中文字段是否为空 → 情况3（真实缺失）
3. **优先级3**: 检查是否包含斜杠 → 情况2（复合存储）
4. **优先级4**: 只有中文，没有英文

---

## 📊 实际测试示例

### **测试1: 分开存储**
```python
result = split_field("聚酯纤维", "polyester")
print(result)  # ('聚酯纤维', 'polyester')
```

### **测试2: 斜杠分割**
```python
result = split_field("聚酯纤维/polyester", "")
print(result)  # ('聚酯纤维', 'polyester')
```

### **测试3: JSON转义斜杠**
```python
# JSON: "聚酯纤维\\/polyester"
# Python解析后: "聚酯纤维/polyester"
result = split_field("聚酯纤维/polyester", "")
print(result)  # ('聚酯纤维', 'polyester')
```

### **测试4: 真实缺失**
```python
result = split_field("", "")
print(result)  # ('', '')
```

### **测试5: 只有中文**
```python
result = split_field("聚酯纤维", "")
print(result)  # ('聚酯纤维', '')
```

---

## ⚠️ 关于 JSON 转义

### **JSON 中的斜杠转义**
- JSON 标准允许对斜杠进行转义：`\/`
- Python 的 `json` 模块会自动处理：
  ```python
  import json
  s = json.loads('"聚酯纤维\\/polyester"')
  print(s)  # 输出: 聚酯纤维/polyester
  ```
- `requests.json()` 也会自动处理，无需额外代码

### **不需要特殊处理**
- ✅ `split_field` 函数直接检查 `/` 即可
- ✅ 不需要检查 `\/`
- ✅ JSON 解析器已经处理了转义

---

## 🎯 总结

### **逻辑正确性**
✅ 当前的 `split_field` 函数已经正确处理了所有三种情况

### **优先级合理**
✅ 优先使用独立的英文字段（最可靠）
✅ 其次处理斜杠分割（兼容性）
✅ 最后返回空值（真实缺失）

### **无需额外处理**
✅ JSON 转义由解析器自动处理
✅ 不需要手动处理 `\/`

### **如果数据库中有空值**
可能的原因：
1. ERP 数据本身就是空的（情况3）
2. 产品未填写材质和用途信息
3. 旧数据（在修复逻辑之前同步的）

建议：重新同步这些产品的数据
