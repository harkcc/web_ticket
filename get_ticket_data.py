import pandas as pd
from typing import Dict, List, Optional
from dataclasses import dataclass


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

    def process(self, template_name=None):
        """处理装箱单
        
        Args:
            template_name: 模板名称，用于特殊处理某些模板
        """
        try:
            print(f"Reading Excel file: {self.file_path}")
            df = pd.read_excel(self.file_path)

            if df.empty:
                raise ValueError("Excel file is empty")

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

            # 读取箱子信息
            box_number_index = None
            box_columns = {}
            box_types = {}

            # 找到箱号所在行
            for index, row in enumerate(df.iterrows()):
                if str(row[1].iloc[1]).strip() == '箱号':  # 第2列是否为"箱号"
                    box_number_index = index
                    break

            if box_number_index is None:
                raise ValueError("未找到箱号行")

            print(f"Found box number row at index {box_number_index}")

            # 从箱号行开始处理每一列
            valid_columns = []  # 存储有效的箱子列
            for col in range(2, len(df.columns)):  # 从第3列开始
                box_number_str = str(df.iloc[box_number_index, col]).strip() if not pd.isna(df.iloc[box_number_index, col]) else ""
                if box_number_str and box_number_str.isdigit():
                    valid_columns.append(col)
                    box_number = int(box_number_str)
                    box_type = str(df.iloc[box_number_index - 1, col]).strip() if not pd.isna(df.iloc[box_number_index - 1, col]) else ""
                    
                    # 查找匹配的箱规
                    matched_box_type = None
                    for spec_type in self.BOX_SPECS:
                        if spec_type.strip() in box_type or box_type in spec_type.strip():
                            matched_box_type = spec_type
                            break

                    if matched_box_type:
                        try:
                            # 获取重量
                            weight_str = str(df.iloc[box_number_index - 2, col]).strip() if not pd.isna(df.iloc[box_number_index - 2, col]) else ""
                            box_weight = float(weight_str) if weight_str and weight_str != '0' else None

                            # 创建箱子对象并设置规格
                            self.boxes[box_number] = PackingListBox(box_number)
                            specs = self.BOX_SPECS[matched_box_type]
                            self.boxes[box_number].length = specs["length"]
                            self.boxes[box_number].width = specs["width"]
                            self.boxes[box_number].height = specs["height"]
                            self.boxes[box_number].weight = box_weight if box_weight is not None else specs["weight"]

                            box_columns[col] = box_number
                            box_types[box_number] = matched_box_type
                            print(f"Processed box {box_number} (Type: {matched_box_type}, Weight: {self.boxes[box_number].weight}kg)")

                        except Exception as e:
                            print(f"Warning: Invalid box information in column {col+1}: {str(e)}")
                            continue

            if not box_columns:
                raise ValueError("未找到有效的箱子信息")

            # 数据从箱号行后两行开始
            data_start_row = box_number_index + 2
            data_end_row = None

            # 查找数据结束位置
            for i in range(data_start_row, len(df)):
                first_col = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ""
                if data_end_row is None and not first_col:
                    data_end_row = i - 1
                    break

            if data_end_row is None:
                data_end_row = len(df) - 1

            print(f"Data range: row {data_start_row + 1} to {data_end_row + 1}")

            # 处理商品信息
            row_count = 0
            item_count = 0  # 添加实际商品计数
            for _, row in df.iloc[data_start_row:data_end_row + 1].iterrows():
                row_count += 1
                
                # 检查SKU是否为空或者只包含空白字符
                sku = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
                if not sku:
                    print(f"跳过空行 {row_count}")
                    continue

                try:
                    item_count += 1  # 只有在有效商品时才增加计数

                    # 尝试获取总数量
                    total_quantity = None
                    try:
                        if not pd.isna(row.iloc[1]):  # 第2列是总数量
                            total_quantity = int(float(str(row.iloc[1]).strip()))
                    except (ValueError, TypeError):
                        print(f"Warning: Invalid quantity in row {row_count}, using sum of box quantities")

                    print(f"Processing row {row_count}: SKU={sku}")

                    # 创建商品信息
                    item = PackingListItem(
                        sequence_no=item_count,  # 使用实际商品计数作为序号
                        msku=str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else "",  # SKU作为MSKU
                        fnsku="",  # FNSKU置空
                        product_name="",  # 产品名称置空
                        sku=sku,
                        quantity=0,  # 先设为0，后面再更新
                        box_quantities={}
                    )

                    # 处理每个箱子中的数量
                    box_total = 0
                    for col, box_number in box_columns.items():
                        try:
                            if not pd.isna(row.iloc[col]):
                                quantity = int(float(str(row.iloc[col]).strip()))
                                if quantity > 0:
                                    item.box_quantities[box_number] = quantity
                                    self.boxes[box_number].add_item(item)
                                    box_total += quantity
                                    print(f"  - Box {box_number}: {quantity} units")
                        except (ValueError, TypeError) as e:
                            print(f"Warning: Invalid quantity in row {row_count}, box {box_number}: {str(e)}")
                            continue

                    # 更新总数量
                    item.quantity = total_quantity if total_quantity is not None else box_total
                    if total_quantity is not None and total_quantity != box_total:
                        print(f"Warning: Total quantity ({total_quantity}) doesn't match sum of box quantities ({box_total})")

                    self.items.append(item)

                except Exception as e:
                    print(f"Warning: Error processing row {row_count}: {str(e)}")
                    continue

            if not self.items:
                raise ValueError("未找到有效的商品信息")

            print(f"\nProcessing complete:")
            print(f"- Total products: {len(self.items)}")
            print(f"- Total boxes: {len(self.boxes)}")
            for box_number, box in self.boxes.items():
                box_type = box_types.get(box_number, "Unknown")
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