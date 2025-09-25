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

    def _detect_box_start_column(self, df):
        """检测箱子数据的起始列索引"""
        try:
            # 检查表头，寻找包含"箱"字的列
            header_row = df.iloc[1]  # 第二行通常是表头
            
            for col_idx, header in enumerate(header_row):
                if pd.notna(header) and ("第" in str(header) and "箱" in str(header)):
                    print(f"找到箱子列标题: {header} 在第 {col_idx + 1} 列")
                    return col_idx
            
            # 如果没找到标准的箱子列标题，检查是否有"单箱数量"和"箱数"列
            for col_idx, header in enumerate(header_row):
                if pd.notna(header) and "单箱数量" in str(header):
                    print(f"检测到单箱单个SKU格式，单箱数量列在第 {col_idx + 1} 列")
                    # 单箱数量后面应该是箱数，再后面才是箱子数据
                    return col_idx + 2  # 跳过"单箱数量"和"箱数"两列
            
            # 默认情况下，假设第7列开始是箱子数据（普通格式）
            print("使用默认箱子起始列: 6 (第7列)")
            return 6
            
        except Exception as e:
            print(f"检测箱子起始列时出错: {str(e)}，使用默认值")
            return 6

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
            
            # 检测装箱单类型（普通 vs 单箱单个SKU）
            box_start_col = self._detect_box_start_column(df)
            print(f"检测到箱子数据起始列: {box_start_col}")

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

                # 处理每箱数量（使用动态检测的起始列）
                for col_idx, col in enumerate(df.columns[box_start_col:], start=box_start_col):
                    quantity = row[col_idx]
                    if not pd.isna(quantity) and quantity > 0:
                        box_number = col_idx - box_start_col + 1  # 箱号从1开始
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
                col_idx = box_start_col + box_number - 1  # 箱子列索引（动态计算）

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
        "5号": {"length": 29, "width": 17, "height": 19, "weight": 0.283},
        "定制49": {"length": 50, "width": 50, "height": 40, "weight": 1.2},
        "定制55": {"length": 56, "width": 46, "height": 51, "weight": 1.6},
        "搬家大": {"length": 61, "width": 41, "height": 51, "weight": 1.45},
        "定制64": {"length": 65, "width": 40, "height": 45, "weight": 1.5},
        "圣诞树": {"length": 91, "width": 49.5, "height": 37, "weight": 1.7},
        "定制59": {"length": 59, "width": 48, "height": 39, "weight": 1.4},
        "定制53": {"length": 53, "width": 33, "height": 43, "weight": 1.35},
        "定制69": {"length": 69, "width": 42, "height": 48, "weight": 1.55},
    }

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.shipment_id: Optional[str] = None
        self.boxes: Dict[int, PackingListBox] = {}
        self.items: List[PackingListItem] = []

    def _parse_box_dimensions(self, box_spec: str):
        """解析箱子规格字符串"""
        try:
            # 首先检查是否在预定义的规格映射表中
            if box_spec in self.BOX_SPECS:
                spec = self.BOX_SPECS[box_spec]
                return spec["length"], spec["width"], spec["height"], spec["weight"]
            
            print(f"Warning: Unknown box specification format: {box_spec}")
            return None
        except Exception as e:
            print(f"Error parsing box dimensions: {str(e)}")
            return None

    def process(self, template_name=None):
        """处理装箱单"""
        try:
            # 读取Excel文件
            print(f"Reading Excel file: {self.file_path}")
            df = pd.read_excel(self.file_path)

            if df.empty:
                raise ValueError("Excel file is empty")

            # 获取Shipment ID
            try:
                self.shipment_id = str(df.iloc[0, 1])  # 第1行第2列
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
