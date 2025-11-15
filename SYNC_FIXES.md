# sync_erp_to_mongodb.py 修复记录

根据 `export_erp_to_excel.py` 的正确处理逻辑，对 `sync_erp_to_mongodb.py` 进行了以下修复：

## 修复内容

### 1. 重量处理修复 ✅

**问题**：原代码没有进行单位转换，直接使用克作为单位

**修复前**：
```python
weight_str = spec_info.get('cg_product_net_weight', '0')
try:
    weight = float(weight_str)
except:
    weight = 0.0
```

**修复后**：
```python
# 获取重量（ERP中单位是克，转换为千克）
weight_str = spec_info.get('cg_product_net_weight', '0')
try:
    weight_g = float(weight_str)  # 原始重量（克）
    weight = weight_g * 0.001  # 转换为千克
except (ValueError, TypeError):
    weight = 0.0
```

**改进点**：
- ✅ 添加单位转换：克 → 千克（× 0.001）
- ✅ 添加清晰的注释说明
- ✅ 使用更明确的变量名（weight_g）
- ✅ 更精确的异常处理（ValueError, TypeError）

---

### 2. 品牌和型号处理修复 ✅

**问题**：原代码使用 `or "无"` 无法处理只有空格的字符串

**修复前**：
```python
"brand": info.get('brand_name', '') or "无",
"model": info.get('model', '') or "无",
```

**修复后**：
```python
# 处理品牌和型号（空值显示"无"）
brand = info.get('brand_name', '')
brand = brand.strip() if brand else ''
brand = brand if brand else "无"

model = info.get('model', '')
model = model.strip() if model else ''
model = model if model else "无"

document = {
    ...
    "brand": brand,
    "model": model,
    ...
}
```

**改进点**：
- ✅ 先 strip() 去除空格
- ✅ 再判断是否为空
- ✅ 正确处理 None、空字符串、只有空格等情况

**测试用例**：
| 输入值 | 原逻辑输出 | 新逻辑输出 |
|--------|-----------|-----------|
| `None` | "无" | "无" ✅ |
| `""` | "无" | "无" ✅ |
| `"   "` | `"   "` ❌ | "无" ✅ |
| `"Apple"` | "Apple" | "Apple" ✅ |
| `"  Nike  "` | "  Nike  " | "Nike" ✅ |

---

### 3. 材质和用途拆分处理 ✅

**状态**：已经正确实现，无需修改

`split_field()` 方法已经正确处理了三种情况：
1. ✅ 分开存储：`customs_clearance_material` + `customs_clearance_en_material`
2. ✅ 复合存储：`customs_clearance_material = "中文/English"`
3. ✅ 真实缺失：字段为空

---

## 修复对比总结

| 项目 | export_erp_to_excel.py | sync_erp_to_mongodb.py | 状态 |
|------|----------------------|----------------------|------|
| 重量单位转换 | ✅ 克→千克（×0.001） | ✅ 已修复 | ✅ |
| 品牌空值处理 | ✅ strip + 判空 | ✅ 已修复 | ✅ |
| 型号空值处理 | ✅ strip + 判空 | ✅ 已修复 | ✅ |
| 材质拆分 | ✅ 支持三种情况 | ✅ 已实现 | ✅ |
| 用途拆分 | ✅ 支持三种情况 | ✅ 已实现 | ✅ |
| 异常处理 | ✅ 精确捕获 | ✅ 已修复 | ✅ |

---

## 数据一致性保证

修复后，两个文件的数据处理逻辑完全一致：

1. **重量字段**：统一使用千克(kg)作为单位
2. **品牌/型号**：空值统一显示"无"，自动去除空格
3. **材质/用途**：统一的中英文拆分逻辑
4. **异常处理**：统一使用精确的异常类型捕获

---

## 修复时间

2025-11-15 17:40

## 修复人员

Cascade AI Assistant
