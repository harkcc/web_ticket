from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import os
import traceback
import json
from datetime import datetime
from db_connector import MongoDBConnector
from io import BytesIO


class ProcessingError(Exception):
    """处理错误的自定义异常类"""
    pass


class InvoiceGenerator:
    def __init__(self, upload_folder, output_folder, db_connector=None, image_folder=None):
        """
        初始化发票生成器
        :param upload_folder: 上传文件夹路径
        :param output_folder: 输出文件夹路径
        :param db_connector: 数据库连接器（可选）
        :param image_folder: 图片文件夹路径（可选）
        """
        self.db_connector = db_connector if db_connector else MongoDBConnector()
        self.upload_folder = upload_folder
        self.output_folder = output_folder
        # 获取当前文件所在的目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.image_folder = os.path.join(current_dir, '图片测试')  # 图片文件夹路径

    def _get_template_handler(self, template_path):
        """根据模板文件名选择对应的处理方法"""
        template_name = os.path.basename(template_path).lower()
        if "叮铛卡航限时达" in template_name:
            return self._fill_dingdang_template
        elif "林道美森限时快船" in template_name:
            return self._fill_ldmsxsd_template
        elif "法国" in template_name:
            return self._fill_france_template
        elif "英国" in template_name:
            return self._fill_uk_template
        else:
            return self._fill_default_template

    def generate_invoice(self, template_path, box_data, output_path, code=None, address_info=None):
        """
        生成发票
        :param template_path: 模板文件路径
        :param box_data: 箱子数据
        :param output_path: 输出文件路径
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :return: (success, message)
        """
        try:
            print(f"\n=== 开始生成发票 ===")
            print(f"模板文件: {template_path}")
            print(f"输出路径: {output_path}")
            print(f"编码: {code}")
            print(f"箱子数据: {box_data}")
            print(f"地址信息: {address_info}")

            # 使用openpyxl加载模板
            print(f"正在加载模板文件...")
            wb = load_workbook(template_path)
            print(f"成功加载模板文件，工作表: {wb.sheetnames}")

            # 获取对应的模板处理方法
            template_handler = self._get_template_handler(template_path)
            if template_handler:
                # 调用模板处理方法
                template_handler(wb, box_data, code, address_info)
            else:
                # 使用默认模板处理方法
                self._fill_default_template(wb, box_data, code, address_info)

            # 保存文件
            print(f"正在保存文件到: {output_path}")
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            wb.save(output_path)
            print(f"文件保存成功")

            return True, "发票生成成功"

        except Exception as e:
            error_msg = f"生成发票时发生错误: {str(e)}\n{traceback.format_exc()}"
            print(error_msg)
            return False, error_msg

    def _get_product_info(self, msku, db=None):
        """
        从MongoDB获取产品信息
        :param msku: 产品的MSKU
        :param db: 数据库连接（可选）
        :return: 包含产品信息的字典
        """
        try:
            if db is None:
                # 如果没有传入db连接，创建新的连接
                with self.db_connector as db:
                    return self._get_product_info(msku, db)
            
            # 使用传入的db连接
            collection = db['msku_info']
            product = collection.find_one({'msku': msku})
            
            if product:
                return {
                    'cn_name': product.get('productNameZh', ''),
                    'en_name': product.get('productNameEn', ''),
                    'en_usage': product.get('useEn', ''),
                    'ch_usage':product.get('useZh', ''),
                    'material_en': product.get('materialEn', ''),
                    'material_cn': product.get('materialZh', ''),
                    'hs_code': product.get('HS', ''),
                    'usage_en': product.get('useEn', ''),
                    'usage_cn': product.get('useEn', ''),
                    'brand': product.get('brand', ''),
                    'model': product.get('model', ''),
                    'link': product.get('productLink', ''),
                    'price':product.get('askprice', ''),
                    'electrified':product.get('electrified', ''),
                    'magnetic':product.get('magnetic', ''),
                    'weight':product.get('weight', ''),
                }
            return None
        except Exception as e:
            print(f"Error fetching product info for MSKU {msku}: {str(e)}")
            return None

    def _set_cell_value(self, sheet, row, column, value, style_info):
        """
        设置单元格的值和样式
        :param sheet: 工作表对象
        :param row: 行号
        :param column: 列号
        :param value: 单元格值
        :param style_info: 样式信息
        """
        cell = sheet.cell(row=row, column=column)
        cell.value = value
        cell.font = style_info['font']
        cell.border = style_info['border']
        cell.alignment = style_info['alignment']

    def _fill_dingdang_template(self, wb, box_data, code=None, address_info=None):
        """
        填充叮铛卡航限时达模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        """
        with self.db_connector as db:
            try:
                sheet = wb['模板']  # 获取模板工作表

                print("开始写入模版信息")

                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }

                # 在第一行B列填充编码
                if code:
                    cell = sheet.cell(row=1, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=9)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info']
                    try:
                        # 填充收件人信息
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=4, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                        # 填充地址信息
                        address_parts = []
                        if 'addressLine1' in address_info_detail:
                            address_parts.append(address_info_detail['addressLine1'])
                        if 'city' in address_info_detail:
                            address_parts.append(address_info_detail['city'])
                        if 'stateOrProvinceCode' in address_info_detail:
                            address_parts.append(address_info_detail['stateOrProvinceCode'])
                        if 'postalCode' in address_info_detail:
                            address_parts.append(address_info_detail['postalCode'])
                        if 'countryCode' in address_info_detail:
                            address_parts.append(address_info_detail['countryCode'])

                        # 填充地址信息
                        if 'addressLine1' in address_info_detail:
                            cell = sheet.cell(row=6, column=2)  # B3单元格
                            cell.value = address_info_detail['addressLine1']

                        # 城市
                        if 'city' in address_info_detail:
                            cell = sheet.cell(row=8, column=2)  # B4单元格
                            cell.value = address_info_detail['city']

                        # 省
                        if 'stateOrProvinceCode' in address_info_detail:
                            cell = sheet.cell(row=5, column=2)  # B5单元格
                            cell.value = address_info_detail['stateOrProvinceCode']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=10, column=2)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=11, column=2)  # B7单元格
                            cell.value = address_info_detail['countryCode']

                        if address_parts:
                            cell = sheet.cell(row=6, column=2)  # B3单元格
                            cell.value = ', '.join(address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=16, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 先解除所有合并的单元格
                print(f"正在解除合并单元格...")
                merged_ranges = list(sheet.merged_cells.ranges)
                for merged_range in merged_ranges:
                    try:
                        sheet.unmerge_cells(str(merged_range))
                    except:
                        pass
                print(f"合并单元格解除完成")

                # 填充数据
                row_num = 18  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始

                # 遍历每个箱子
                for box_number, box in box_data.items():
                    print(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        price = product_info.get('price', 0)
                        total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                        if product_info:
                            item.product_name = product_info.get('cn_name', item.product_name)

                        # 设置单元格值和样式
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (3,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (4, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            (5, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (6, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (7, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (8, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (9, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_ch', '' ))if product_info else ''),    # 用途 (H列)
                            (10, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (11, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (12, product_info.get('link', '') if product_info else ''),
                            (13, ''),  # 图片列 (M列)
                            (15, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                            (17, box.length if box.length is not None else ""),  # 长度 (Q列)
                            (18, box.width if box.width is not None else ""),    # 宽度 (R列)
                            (19, box.height if box.height is not None else "")   # 高度 (S列)
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)
                            
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"M{row_num}"  # 图片列（第13列）
                                self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

            finally:
                # 确保连接在最后关闭
                pass

    def _fill_ldmsxsd_template(self, wb, box_data, code=None, address_info=None):
        """填充林道美森限时快船模板"""
        sheet = wb['模板']  # 获取模板工作表

        # 定义样式信息
        style_info = {
            'font': Font(name='Arial', size=10),
            'border': Border(left=Side(border_style='thin'),
                             right=Side(border_style='thin'),
                             top=Side(border_style='thin'),
                             bottom=Side(border_style='thin')),
            'alignment': Alignment(horizontal='center', vertical='center')
        }

        # 在第一行B列填充编码
        if code:
            cell = sheet.cell(row=1, column=2)  # B列是第2列
            cell.value = code
            cell.font = Font(name='Arial', size=9)

        # 先解除所有合并的单元格
        print(f"正在解除合并单元格...")
        merged_ranges = list(sheet.merged_cells.ranges)
        for merged_range in merged_ranges:
            try:
                sheet.unmerge_cells(str(merged_range))
            except:
                pass
        print(f"合并单元格解除完成")

        # 填充数据
        row_num = 18  # 从第18行开始填充

        # 遍历每个箱子
        for box_number, box in box_data.items():
            print(f"处理箱子 {box_number}")

            # 遍历箱子中的每个产品
            for item in box.items:
                # 货箱编号 (A列)
                cell = sheet.cell(row=row_num, column=1)
                cell.value = box_number
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 重量 (B列)
                cell = sheet.cell(row=row_num, column=2)
                cell.value = box.weight if box.weight is not None else ""
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 品名 (C列)
                cell = sheet.cell(row=row_num, column=3)
                cell.value = item.product_name
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 数量 (D列)
                cell = sheet.cell(row=row_num, column=4)
                cell.value = item.box_quantities.get(box_number, 0)
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 长度 (E列)
                cell = sheet.cell(row=row_num, column=5)
                cell.value = box.length if box.length is not None else ""
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 宽度 (F列)
                cell = sheet.cell(row=row_num, column=6)
                cell.value = box.width if box.width is not None else ""
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 高度 (G列)
                cell = sheet.cell(row=row_num, column=7)
                cell.value = box.height if box.height is not None else ""
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                row_num += 1

            # 在每个箱子的产品列表后添加一个空行
            row_num += 1

        return row_num

    def _fill_france_template(self, wb, box_data, code=None, address_info=None):
        """填充法国模板"""
        sheet = wb['模板']  # 获取模板工作表

        # 定义样式信息
        style_info = {
            'font': Font(name='Arial', size=10),
            'border': Border(left=Side(border_style='thin'),
                             right=Side(border_style='thin'),
                             top=Side(border_style='thin'),
                             bottom=Side(border_style='thin')),
            'alignment': Alignment(horizontal='center', vertical='center')
        }

        # 先解除所有合并的单元格
        print(f"正在解除合并单元格...")
        merged_ranges = list(sheet.merged_cells.ranges)
        for merged_range in merged_ranges:
            try:
                sheet.unmerge_cells(str(merged_range))
            except Exception as e:
                print(f"Warning: Failed to unmerge cells {str(merged_range)}: {str(e)}")
                continue

        # 重置数据区域的单元格
        for row in sheet.iter_rows(min_row=18, max_row=sheet.max_row):
            for cell in row:
                try:
                    cell.value = None
                except Exception as e:
                    print(f"Warning: Failed to reset cell {cell.coordinate}: {str(e)}")

        # 填充数据
        row_num = 18  # 起始行号
        for box_number, box in box_data.items():
            # 填充箱子数据
            for item in box['array']:
                # 从数据库获取产品信息
                product_info = self._get_product_info(item['sku'])
                if product_info:
                    item.update(product_info)
                
                # 设置单元格值
                sheet.cell(row=row_num, column=1, value=box_number)
                sheet.cell(row=row_num, column=2, value=box['box_spec'])
                sheet.cell(row=row_num, column=3, value=box['weight'])
                sheet.cell(row=row_num, column=4, value=item['sku'])
                sheet.cell(row=row_num, column=5, value=item.get('cn_name', ''))
                sheet.cell(row=row_num, column=6, value=item.get('en_name', ''))
                sheet.cell(row=row_num, column=7, value=item['quantity'])
                sheet.cell(row=row_num, column=8, value=item['price'])
                sheet.cell(row=row_num, column=9, value=item['total_value'])
                sheet.cell(row=row_num, column=10, value=item.get('material', ''))
                sheet.cell(row=row_num, column=11, value=item.get('hs_code', ''))
                sheet.cell(row=row_num, column=12, value=item.get('usage', ''))
                sheet.cell(row=row_num, column=13, value=item.get('brand', ''))
                sheet.cell(row=row_num, column=14, value=item.get('model', ''))
                sheet.cell(row=row_num, column=15, value=item.get('link', ''))

                # 应用样式
                for col in range(1, 16):
                    cell = sheet.cell(row=row_num, column=col)
                    cell.font = style_info['font']
                    cell.border = style_info['border']
                    cell.alignment = style_info['alignment']

                row_num += 1

        return row_num

    def _fill_uk_template(self, wb, box_data, code=None, address_info=None):
        """填充英国模板"""
        sheet = wb['模板']  # 获取模板工作表

        # 定义样式信息
        style_info = {
            'font': Font(name='Arial', size=10),
            'border': Border(left=Side(border_style='thin'),
                             right=Side(border_style='thin'),
                             top=Side(border_style='thin'),
                             bottom=Side(border_style='thin')),
            'alignment': Alignment(horizontal='center', vertical='center')
        }

        # 先解除所有合并的单元格
        print(f"正在解除合并单元格...")
        merged_ranges = list(sheet.merged_cells.ranges)
        for merged_range in merged_ranges:
            try:
                sheet.unmerge_cells(str(merged_range))
            except Exception as e:
                print(f"Warning: Failed to unmerge cells {str(merged_range)}: {str(e)}")
                continue

        # 重置数据区域的单元格
        for row in sheet.iter_rows(min_row=18, max_row=sheet.max_row):
            for cell in row:
                try:
                    cell.value = None
                except Exception as e:
                    print(f"Warning: Failed to reset cell {cell.coordinate}: {str(e)}")

        # 填充数据
        row_num = 18  # 起始行号
        for box_number, box in box_data.items():
            # 填充箱子数据
            for item in box['array']:
                # 从数据库获取产品信息
                product_info = self._get_product_info(item['sku'])
                if product_info:
                    item.update(product_info)
                
                # 设置单元格值
                sheet.cell(row=row_num, column=1, value=box_number)
                sheet.cell(row=row_num, column=2, value=box['box_spec'])
                sheet.cell(row=row_num, column=3, value=box['weight'])
                sheet.cell(row=row_num, column=4, value=item['sku'])
                sheet.cell(row=row_num, column=5, value=item.get('cn_name', ''))
                sheet.cell(row=row_num, column=6, value=item.get('en_name', ''))
                sheet.cell(row=row_num, column=7, value=item['quantity'])
                sheet.cell(row=row_num, column=8, value=item['price'])
                sheet.cell(row=row_num, column=9, value=item['total_value'])
                sheet.cell(row=row_num, column=10, value=item.get('material', ''))
                sheet.cell(row=row_num, column=11, value=item.get('hs_code', ''))
                sheet.cell(row=row_num, column=12, value=item.get('usage', ''))
                sheet.cell(row=row_num, column=13, value=item.get('brand', ''))
                sheet.cell(row=row_num, column=14, value=item.get('model', ''))
                sheet.cell(row=row_num, column=15, value=item.get('link', ''))

                # 应用样式
                for col in range(1, 16):
                    cell = sheet.cell(row=row_num, column=col)
                    cell.font = style_info['font']
                    cell.border = style_info['border']
                    cell.alignment = style_info['alignment']

                row_num += 1

        return row_num

    def _fill_default_template(self, wb, box_data, code=None, address_info=None):
        """填充默认模板"""
        sheet = wb['模板']  # 获取模板工作表

        # 定义样式信息
        style_info = {
            'font': Font(name='Arial', size=10),
            'border': Border(left=Side(border_style='thin'),
                             right=Side(border_style='thin'),
                             top=Side(border_style='thin'),
                             bottom=Side(border_style='thin')),
            'alignment': Alignment(horizontal='center', vertical='center')
        }

        # 先解除所有合并的单元格
        print(f"正在解除合并单元格...")
        merged_ranges = list(sheet.merged_cells.ranges)
        for merged_range in merged_ranges:
            try:
                sheet.unmerge_cells(str(merged_range))
            except Exception as e:
                print(f"Warning: Failed to unmerge cells {str(merged_range)}: {str(e)}")
                continue

        # 重置数据区域的单元格
        for row in sheet.iter_rows(min_row=18, max_row=sheet.max_row):
            for cell in row:
                try:
                    cell.value = None
                except Exception as e:
                    print(f"Warning: Failed to reset cell {cell.coordinate}: {str(e)}")

        # 填充数据
        row_num = 18  # 起始行号
        for box_number, box in box_data.items():
            # 填充箱子数据
            for item in box['array']:
                # 从数据库获取产品信息
                product_info = self._get_product_info(item['sku'])
                if product_info:
                    item.update(product_info)
                
                # 设置单元格值
                sheet.cell(row=row_num, column=1, value=box_number)
                sheet.cell(row=row_num, column=2, value=box['box_spec'])
                sheet.cell(row=row_num, column=3, value=box['weight'])
                sheet.cell(row=row_num, column=4, value=item['sku'])
                sheet.cell(row=row_num, column=5, value=item.get('cn_name', ''))
                sheet.cell(row=row_num, column=6, value=item.get('en_name', ''))
                sheet.cell(row=row_num, column=7, value=item['quantity'])
                sheet.cell(row=row_num, column=8, value=item['price'])
                sheet.cell(row=row_num, column=9, value=item['total_value'])
                sheet.cell(row=row_num, column=10, value=item.get('material', ''))
                sheet.cell(row=row_num, column=11, value=item.get('hs_code', ''))
                sheet.cell(row=row_num, column=12, value=item.get('usage', ''))
                sheet.cell(row=row_num, column=13, value=item.get('brand', ''))
                sheet.cell(row=row_num, column=14, value=item.get('model', ''))
                sheet.cell(row=row_num, column=15, value=item.get('link', ''))

                # 应用样式
                for col in range(1, 16):
                    cell = sheet.cell(row=row_num, column=col)
                    cell.font = style_info['font']
                    cell.border = style_info['border']
                    cell.alignment = style_info['alignment']

                row_num += 1

        return row_num

    def insert_centered_image(self, worksheet, cell_address, image_path, fixed_width=None, fixed_height=None):
        """
        在指定的单元格中插入居中的图片
        :param worksheet: 工作表对象
        :param cell_address: 单元格地址（例如'A1'）
        :param image_path: 图片文件路径
        :param fixed_width: 固定宽度（可选）
        :param fixed_height: 固定高度（可选）
        :return: 是否成功插入图片
        """
        try:
            # 读取图片
            img = PILImage.open(image_path)
            
            # 获取单元格的宽度和高度（以像素为单位）
            column_width = worksheet.column_dimensions[cell_address[0]].width
            row_height = worksheet.row_dimensions[int(cell_address[1:])].height
            
            # 如果没有指定宽度和高度，使用单元格的大小
            if fixed_width is None:
                fixed_width = column_width * 7  # 转换为像素
            if fixed_height is None:
                fixed_height = row_height * 1.5  # 转换为像素
                
            # 获取原始图片尺寸
            original_width, original_height = img.size
            
            # 计算缩放比例
            width_ratio = fixed_width / original_width
            height_ratio = fixed_height / original_height
            scale = min(width_ratio, height_ratio)
            
            # 计算新的尺寸
            new_width = int(original_width * scale)
            new_height = int(original_height * scale)
            
            # 调整图片大小
            img = img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
            
            # 将图片保存到BytesIO对象
            img_byte_arr = BytesIO()
            img.save(img_byte_arr, format=img.format if img.format else 'PNG')
            img_byte_arr.seek(0)
            
            # 创建Excel图片对象
            xl_img = XLImage(img_byte_arr)
            xl_img.anchor = cell_address
            worksheet.add_image(xl_img)
            
            return True
        except Exception as e:
            print(f"插入图片时发生错误: {str(e)}")
            return False

    def insert_product_image(self, worksheet, cell_address, msku, image_folder, fixed_width=None, fixed_height=None):
        """
        在Excel工作表中插入产品图片
        :param worksheet: openpyxl工作表对象
        :param cell_address: 单元格地址
        :param msku: 产品MSKU
        :param image_folder: 图片文件夹路径
        :param fixed_width: 固定宽度（可选）
        :param fixed_height: 固定高度（可选）
        """
        try:
            # 构建图片文件路径
            image_path = os.path.join(image_folder, f"{msku}.jpg")
            print(f"尝试加载图片: {image_path}")
            
            # 检查图片文件是否存在
            if not os.path.exists(image_path):
                print(f"图片文件不存在: {image_path}")
                return False
                
            # 插入图片
            return self.insert_centered_image(worksheet, cell_address, image_path, fixed_width, fixed_height)
        except Exception as e:
            print(f"处理产品图片时发生错误: {str(e)}")
            return False
