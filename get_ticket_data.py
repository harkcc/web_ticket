import pandas as pd
from typing import Dict, List, Optional
from dataclasses import dataclass, field


@dataclass
class PackingListItem:
    """装箱单中的单个商品信息"""
    sequence_no: int  # 序号
    msku: str  # MSKU
    fnsku: str  # FNSKU
    product_name: str  # 品名
    sku: str  # SKU
    quantity: int  # 发货数量
    box_quantities: Dict[int, int]  # 每箱数量 {箱号: 数量}
    box_original_values: Dict[int, str] = field(default_factory=dict)  # 存储特殊格式的原始值

    def __post_init__(self):
        if self.box_original_values is None:
            self.box_original_values = {}


class PackingListBox:
    """装箱单中的箱子信息"""

    def __init__(self, box_number: int):
        self.box_number = box_number
        self.items: List[PackingListItem] = []
        self.length: Optional[float] = None
        self.width: Optional[float] = None
        self.height: Optional[float] = None
        self.weight: Optional[float] = None

    def add_item(self, item: PackingListItem):
        """添加商品到箱子"""
        self.items.append(item)

    def set_dimensions(self, length: float, width: float, height: float):
        """设置箱子尺寸"""
        self.length = length
        self.width = width
        self.height = height

    def set_weight(self, weight: float):
        """设置箱子重量"""
        self.weight = weight


class PackingListProcessor:
    """领星装箱单处理器"""

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.shipment_id: Optional[str] = None
        self.boxes: Dict[int, PackingListBox] = {}
        self.items: List[PackingListItem] = []

    def process(self):
        """处理装箱单"""
        try:
            # 读取Excel文件
            print(f"Reading Excel file: {self.file_path}")
            df = pd.read_excel(self.file_path)

            if df.empty:
                raise ValueError("Excel file is empty")

            # 获取Shipment ID
            try:
                self.shipment_id = str(df.iloc[0, 1])  # 第一行第二列
                if pd.isna(self.shipment_id):
                    raise ValueError("Shipment ID is empty")
            except Exception as e:
                raise ValueError(f"Invalid Shipment ID: {str(e)}")

            print(f"Found Shipment ID: {self.shipment_id}")

            # 找到最后一个产品的行号
            last_product_row = None
            for i in range(len(df)):
                try:
                    first_col_value = df.iloc[i, 0]  # 使用 iloc 直接获取第一列的值
                    # 尝试将第一列转换为浮点数，然后检查是否为正数
                    if pd.isna(first_col_value):
                        continue
                    value = float(first_col_value)
                    if value > 0:
                        last_product_row = i  # 直接使用循环索引
                except (ValueError, TypeError):
                    continue

            if last_product_row is None:
                raise ValueError("No product data found")

            print(f"Last product row: {last_product_row}")

            # 处理商品信息（从第三行开始）
            row_count = 0
            for _, row in df.iloc[2:last_product_row + 1].iterrows():
                if pd.isna(row[0]):  # 如果序号为空，跳过
                    continue

                row_count += 1
                msku = row[1]  # 第二列是MSKU
                fnsku = row[2]  # 第三列是FNSKU
                product_name = row[3]  # 第四列是品名
                sku = row[4]  # 第五列是SKU
                total_quantity = row[5]  # 第六列是发货总数量

                if pd.isna(msku) or pd.isna(total_quantity):
                    continue

                print(f"Processing row {row_count}: SKU={sku}, Total Quantity={total_quantity}")

                # 创建商品信息
                item = PackingListItem(
                    sequence_no=int(row[0]),
                    msku=row[1],
                    fnsku=row[2],
                    product_name=row[3],
                    sku=row[4],
                    quantity=int(total_quantity),
                    box_quantities={}
                )

                # 处理每箱数量
                for col_idx, col in enumerate(df.columns[6:], start=6):
                    quantity = row[col_idx]
                    if not pd.isna(quantity) and quantity > 0:
                        box_number = col_idx - 5  # 箱号从1开始
                        item.box_quantities[box_number] = int(quantity)
                        print(f"  - Box {box_number}: {quantity} units")

                        # 确保箱子对象存在
                        if box_number not in self.boxes:
                            self.boxes[box_number] = PackingListBox(box_number)
                            print(f"  Created new box: Box {box_number}")
                        self.boxes[box_number].add_item(item)

                self.items.append(item)

            # 读取箱子尺寸信息
            # 箱子信息在最后一个产品后两行开始
            box_info_start = int(last_product_row) + 2  # 确保是整数运算

            # 对每个箱子
            for box_number in self.boxes:
                box = self.boxes[box_number]
                col_idx = 5 + box_number  # 箱子列索引（第1箱从第7列开始）

                # 获取箱子信息（按顺序：重量、长、宽、高）
                # 重量
                weight = df.iloc[box_info_start, col_idx]
                if not pd.isna(weight):
                    box.weight = float(weight)
                    print(f"Box {box_number} weight: {box.weight} kg")

                # 长度
                length = df.iloc[box_info_start + 1, col_idx]
                if not pd.isna(length):
                    box.length = float(length)
                    print(f"Box {box_number} length: {box.length} cm")

                # 宽度
                width = df.iloc[box_info_start + 2, col_idx]
                if not pd.isna(width):
                    box.width = float(width)
                    print(f"Box {box_number} width: {box.width} cm")

                # 高度
                height = df.iloc[box_info_start + 3, col_idx]
                if not pd.isna(height):
                    box.height = float(height)
                    print(f"Box {box_number} height: {box.height} cm")

            print(f"\nProcessing complete:")
            print(f"- Total products: {len(self.items)}")
            print(f"- Total boxes: {len(self.boxes)}")
            for box_number, box in self.boxes.items():
                print(f"Box {box_number}:")
                print(f"  - Dimensions: {box.length}x{box.width}x{box.height} cm")
                print(f"  - Weight: {box.weight} kg")
                print(f"  - Items: {len(box.items)}")

            return self.boxes

        except Exception as e:
            print(f"Error processing file: {str(e)}")
            return None

    def get_box_count(self) -> int:
        """获取箱子总数"""
        return len(self.boxes)

    def get_total_quantity(self) -> int:
        """获取总发货数量"""
        return sum(item.quantity for item in self.items)

    def get_box(self, box_number: int) -> Optional[PackingListBox]:
        """获取指定箱号的箱子信息"""
        return self.boxes.get(box_number)

    def get_all_boxes(self) -> List[PackingListBox]:
        """获取所有箱子信息"""
        return list(self.boxes.values())

    def get_all_items(self) -> List[PackingListItem]:
        """获取所有商品信息"""
        return self.items


class SimplePackingListProcessor:
    """简单装箱单处理器"""

    # 箱子规格映射表
    BOX_SPECS = {
        "中号": {"length": 51, "width": 41, "height": 41, "weight": 1.102},
        "1号": {"length": 54, "width": 30, "height": 38, "weight": 0.843},
        "2号": {"length": 54, "width": 24, "height": 30, "weight": 0.631},
        "3号": {"length": 44, "width": 22, "height": 28, "weight": 0.458},
        "4号": {"length": 36, "width": 20, "height": 24, "weight": 0.283},
        "定制49": {"length": 50, "width": 50, "height": 40, "weight": 1.2},
        "定制55": {"length": 56, "width": 46, "height": 51, "weight": 1.6},
        "搬家大": {"length": 61, "width": 41, "height": 51, "weight": 1.45},
        "定制64": {"length": 65, "width": 40, "height": 45, "weight": 1.4},
        "圣诞树": {"length": 91, "width": 49.5, "height": 37, "weight": 1.7},
        "定制59": {"length": 59, "width": 48, "height": 39, "weight": 1.35}
    }

    def __init__(self, file_path: str):
        """初始化处理器

        Args:
            file_path: Excel文件路径
        """
        self.file_path = file_path
        self.shipment_id: Optional[str] = None
        self.boxes: Dict[int, PackingListBox] = {}
        self.items: List[PackingListItem] = []

    def _parse_box_dimensions(self, box_spec):
        """解析箱子规格，返回长宽高的元组"""
        try:
            # 处理 "定制59" 这样的格式
            if box_spec.startswith('定制'):
                number = int(box_spec[2:])  # 提取数字部分
                # 根据箱号返回预定义的尺寸
                if number == 59:
                    return (59, 48, 39)
                elif number == 53:
                    return (53, 40, 35)
                # 可以添加更多的箱型
            
            # 处理 "59*48*39" 这样的格式
            if '*' in box_spec:
                dimensions = box_spec.split('*')
                if len(dimensions) == 3:
                    return tuple(int(d.strip()) for d in dimensions)
            
            print(f"Warning: Unknown box specification format: {box_spec}")
            return None
        except Exception as e:
            print(f"Error parsing box dimensions: {str(e)}")
            return None

    def process(self, template_name=None):
        """处理装箱单
        
        Args:
            template_name: 模板名称，用于特殊处理某些模板
        """
        print(f"\n========== 开始处理装箱单 ==========")
        print(f"模板名称: {template_name}")
        try:
            print(f"Reading Excel file: {self.file_path}")
            df = pd.read_excel(self.file_path, header=None)
            print(f"\n=== Excel文件基本信息 ===")
            print(f"总行数: {len(df)}")
            print(f"总列数: {len(df.columns)}")
            print("前5行数据预览:")
            print(df.head())

            # 如果是依诺达模板，删除第一列
            if template_name and "依诺达" in template_name:
                df = df.iloc[:, 1:]

            # 获取Shipment ID（第1行第2列）
            try:
                self.shipment_id = str(df.iloc[0, 1])  # 第1行第2列
                if pd.isna(self.shipment_id):
                    raise ValueError("Shipment ID is empty")
            except Exception as e:
                raise ValueError(f"Invalid Shipment ID: {str(e)}")

            print(f"Found Shipment ID: {self.shipment_id}")

            # 查找箱号行
            box_number_index = None
            box_columns = {}  # 用于存储箱号对应的列索引

            print("\n=== 查找箱号行 ===")
            for idx, row in df.iterrows():
                if not pd.isna(row.iloc[1]) and str(row.iloc[1]).strip() == "箱号":
                    box_number_index = idx
                    # 查找每个箱子的列索引
                    for col_idx in range(2, len(row)):  # 从第3列开始查找箱号
                        if not pd.isna(row.iloc[col_idx]):
                            try:
                                box_number = int(row.iloc[col_idx])
                                box_columns[col_idx] = box_number
                                print(f"找到箱号 {box_number} 在列 {col_idx}")
                            except (ValueError, TypeError):
                                continue
                    break

            if box_number_index is None:
                raise ValueError("未找到箱号行")
            print(f"找到箱号行，索引为: {box_number_index}")

            # 处理箱子规格
            print("\n=== 处理箱子规格 ===")
            box_spec_row = None
            box_weight_row = None

            # 查找箱规和重量行
            for idx in range(box_number_index - 2, box_number_index):
                row = df.iloc[idx]
                if not pd.isna(row.iloc[1]):
                    if str(row.iloc[1]).strip() == "箱规":
                        box_spec_row = idx
                    elif str(row.iloc[1]).strip() == "重量（kg）":
                        box_weight_row = idx

            # 处理每个箱子的规格和重量
            for col_idx, box_number in box_columns.items():
                print(f"箱子 {box_number}:")
                
                # 获取箱规
                box_spec = None
                if box_spec_row is not None:
                    box_spec = str(df.iloc[box_spec_row, col_idx]).strip() if not pd.isna(df.iloc[box_spec_row, col_idx]) else None
                    print(f"  - 箱规: {box_spec}")

                # 获取重量
                box_weight = None
                if box_weight_row is not None:
                    try:
                        box_weight = float(df.iloc[box_weight_row, col_idx]) if not pd.isna(df.iloc[box_weight_row, col_idx]) else None
                        print(f"  - 重量: {box_weight}kg")
                    except (ValueError, TypeError):
                        print(f"  - 重量转换失败")

                # 创建箱子对象
                box = PackingListBox(box_number)
                if box_spec:
                    print(f"  - 匹配到箱规: {box_spec}")
                    dimensions = self._parse_box_dimensions(box_spec)
                    if dimensions:
                        box.length, box.width, box.height = dimensions
                        print(f"  - 尺寸: {box.length}x{box.width}x{box.height}")
                if box_weight:
                    box.weight = box_weight
                self.boxes[box_number] = box

            if not self.boxes:
                raise ValueError("未找到有效的箱子信息")

            # 数据从箱号行后一行开始（使用索引）
            data_start_row = box_number_index + 1
            print(f"\n=== 数据范围信息 ===")
            print(f"箱号行索引: {box_number_index}")
            print(f"数据开始行索引: {data_start_row}")

            # 遍历所有可能的数据行
            valid_rows = []
            for idx in range(data_start_row, len(df)):
                row_data = df.iloc[idx]
                first_col = str(row_data.iloc[0]).strip() if not pd.isna(row_data.iloc[0]) else ""
                
                # 显示行数据用于调试
                print(f"\n检查行索引 {idx}:")
                print(f"第一列值: '{first_col}'")
                print(f"完整行数据: {row_data.to_dict()}")
                
                if first_col:  # 如果第一列不为空，说明是有效的数据行
                    print(f"发现有效数据行，索引: {idx}")
                    valid_rows.append(idx)

            if not valid_rows:
                raise ValueError("未找到有效的数据行")

            print(f"\n=== 有效数据行索引: {valid_rows} ===")

            # 处理商品信息
            row_count = 0
            item_count = 0  # 添加实际商品计数
            
            for idx in range(data_start_row, len(df)):
                row = df.iloc[idx]
                row_count += 1
                print(f"\n=== 处理第 {row_count} 行 ===")
                
                # 检查SKU是否为空或者只包含空白字符
                sku = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
                if not sku:
                    print(f"跳过空行 {row_count}")
                    continue

                print(f"处理商品: SKU={sku}")
                item_count += 1

                # 创建商品对象
                item = PackingListItem(
                    sequence_no=item_count,
                    msku=sku,
                    fnsku="",  # 暂时为空
                    product_name="",  # 暂时为空
                    sku=sku,
                    quantity=0,  # 稍后更新
                    box_quantities={}
                )

                # 尝试获取总数量
                total_quantity = None
                try:
                    if not pd.isna(row.iloc[1]):  # 第2列是总数量
                        total_quantity = int(float(str(row.iloc[1]).strip()))
                        print(f"找到总数量: {total_quantity}")
                except (ValueError, TypeError) as e:
                    print(f"Warning: Invalid total quantity in row {row_count}: {str(e)}")

                # 处理每个箱子中的数量
                box_total = 0
                print(f"\n=== 处理商品 {sku} 的箱子数量 ===")
                print(f"box_columns: {box_columns}")
                print(f"当前行完整数据: {row.to_dict()}")
                
                for col_idx, box_number in box_columns.items():
                    try:
                        print(f"\n检查箱子 {box_number} (列 {col_idx}):")
                        raw_value = row.iloc[col_idx]
                        print(f"原始值: {raw_value} (类型: {type(raw_value)})")
                        
                        if not pd.isna(raw_value):
                            value = str(raw_value).strip()
                            print(f"处理值: {value}")
                            
                            # 处理数量
                            try:
                                if ' ' in value:
                                    # 处理 "A1 70" 格式
                                    print(f"发现特殊格式: {value}")
                                    prefix, number = value.split(' ', 1)
                                    print(f"  - 前缀: {prefix}")
                                    print(f"  - 数量: {number}")
                                    quantity = int(number)
                                    original_value = value  # 保存完整的原始值
                                else:
                                    # 普通数字格式
                                    quantity = int(float(value))
                                    original_value = str(quantity)
                                
                                print(f"  - 处理后数量: {quantity}")
                                
                                if quantity > 0:
                                    print(f"添加到箱子 {box_number}: {quantity} (原始值: {original_value})")
                                    item.box_quantities[box_number] = quantity  # 存储数字形式的数量
                                    if not hasattr(item, 'box_original_values'):
                                        item.box_original_values = {}
                                    item.box_original_values[box_number] = original_value  # 存储原始值
                                    
                                    if box_number not in self.boxes:
                                        print(f"创建新箱子: {box_number}")
                                        self.boxes[box_number] = PackingListBox(box_number)
                                    self.boxes[box_number].add_item(item)
                                    box_total += quantity
                                    print(f"当前箱子总数: {box_total}")
                                    print(f"存储的原始值: {item.box_original_values}")
                            except (ValueError, TypeError) as e:
                                print(f"  - 数量转换失败: {str(e)}")
                                continue
                        else:
                            print(f"  - 跳过空值")
                            
                    except Exception as e:
                        print(f"Warning: Error processing box {box_number}: {str(e)}")
                        continue

                # 更新总数量
                print(f"\n=== 更新总数量 ===")
                print(f"总数量: {total_quantity}, 箱子总数: {box_total}")
                item.quantity = total_quantity if total_quantity is not None else box_total
                if total_quantity is not None and total_quantity != box_total:
                    print(f"Warning: Total quantity ({total_quantity}) doesn't match sum of box quantities ({box_total})")

                self.items.append(item)

            if not self.items:
                raise ValueError("未找到有效的商品信息")

            print(f"\nProcessing complete:")
            print(f"- Total products: {len(self.items)}")
            print(f"- Total boxes: {len(self.boxes)}")
            for box_number, box in self.boxes.items():
                box_type = "Unknown"
                print(f"Box {box_number} (Type: {box_type}):")
                print(f"  - Dimensions: {box.length}x{box.width}x{box.height} cm")
                print(f"  - Weight: {box.weight} kg")
                print(f"  - Items: {len(box.items)}")

            return self.boxes

        except Exception as e:
            print(f"Error processing file: {str(e)}")
            return None

    def get_box_count(self) -> int:
        """获取箱子总数"""
        return len(self.boxes)

    def get_total_quantity(self) -> int:
        """获取总发货数量"""
        return sum(item.quantity for item in self.items)

    def get_box(self, box_number: int) -> Optional[PackingListBox]:
        """获取指定箱号的箱子信息"""
        return self.boxes.get(box_number)

    def get_all_boxes(self) -> List[PackingListBox]:
        """获取所有箱子信息"""
        return list(self.boxes.values())

    def get_all_items(self) -> List[PackingListItem]:
        """获取所有商品信息"""
        return self.items