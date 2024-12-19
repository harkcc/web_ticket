import os
import traceback
from turtle import pen
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment


class ProcessingError(Exception):
    """处理错误的自定义异常类"""
    pass

class InvoiceGenerator:
    def __init__(self, upload_folder, output_folder):
        self.upload_folder = upload_folder
        self.output_folder = output_folder

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
            wb = openpyxl.load_workbook(template_path)
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

    def _fill_dingdang_template(self, wb, box_data, code=None, address_info=None):
        """
        填充叮铛卡航限时达模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        """
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
                    # cell.font = style_info['font']
                    # cell.border = style_info['border']
                    # cell.alignment = style_info['alignment']

                # 城市
                if 'city' in address_info_detail:
                    cell = sheet.cell(row=8, column=2)  # B4单元格
                    cell.value = address_info_detail['city']
                    # cell.font = style_info['font']
                    # cell.border = style_info['border']
                    # cell.alignment = style_info['alignment']

                # 省
                if 'stateOrProvinceCode' in address_info_detail:
                    cell = sheet.cell(row=5, column=2)  # B5单元格
                    cell.value = address_info_detail['stateOrProvinceCode']
                    # cell.font = style_info['font']
                    # cell.border = style_info['border']
                    # cell.alignment = style_info['alignment']

                #邮政编码
                if 'postalCode' in address_info_detail:
                    cell = sheet.cell(row=10, column=2)  # B6单元格
                    cell.value = address_info_detail['postalCode']
                    # cell.font = style_info['font']
                    # cell.border = style_info['border']
                    # cell.alignment = style_info['alignment']

                #国家代码
                if 'countryCode' in address_info_detail:
                    cell = sheet.cell(row=11, column=2)  # B7单元格
                    cell.value = address_info_detail['countryCode']
                    # cell.font = style_info['font']
                    # cell.border = style_info['border']
                    # cell.alignment = style_info['alignment']

                if address_parts:
                    cell = sheet.cell(row=6, column=2)  # B3单元格
                    cell.value = ', '.join(address_parts)
            except Exception as e:
                print(f"填充地址信息时发生错误: {str(e)}")
                # 继续处理其他数据

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
                cell = sheet.cell(row=row_num, column=4)
                cell.value = item.product_name
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 数量 (F列)
                cell = sheet.cell(row=row_num, column=6)
                cell.value = item.box_quantities.get(box_number, 0)
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 长度 (Q列)
                cell = sheet.cell(row=row_num, column=17)  # Q是第17列
                cell.value = box.length if box.length is not None else ""
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 宽度 (R列)
                cell = sheet.cell(row=row_num, column=18)  # R是第18列
                cell.value = box.width if box.width is not None else ""
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 高度 (S列)
                cell = sheet.cell(row=row_num, column=19)  # S是第19列
                cell.value = box.height if box.height is not None else ""
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                row_num += 1

            # 在每个箱子的产品列表后添加一个空行,暂时不用
            # row_num += 1

        return row_num

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
                # 设置单元格值
                sheet.cell(row=row_num, column=1, value=box_number)
                sheet.cell(row=row_num, column=2, value=box['box_spec'])
                sheet.cell(row=row_num, column=3, value=box['weight'])
                sheet.cell(row=row_num, column=4, value=item['sku'])
                sheet.cell(row=row_num, column=5, value=item['cn_name'])
                sheet.cell(row=row_num, column=6, value=item['en_name'])
                sheet.cell(row=row_num, column=7, value=item['quantity'])
                sheet.cell(row=row_num, column=8, value=item['price'])
                sheet.cell(row=row_num, column=9, value=item['total_value'])
                sheet.cell(row=row_num, column=10, value=item['material'])
                sheet.cell(row=row_num, column=11, value=item['hs_code'])
                sheet.cell(row=row_num, column=12, value=item['usage'])
                sheet.cell(row=row_num, column=13, value=item['brand'])
                sheet.cell(row=row_num, column=14, value=item['model'])
                sheet.cell(row=row_num, column=15, value=item['link'])

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
                # 设置单元格值
                sheet.cell(row=row_num, column=1, value=box_number)
                sheet.cell(row=row_num, column=2, value=box['box_spec'])
                sheet.cell(row=row_num, column=3, value=box['weight'])
                sheet.cell(row=row_num, column=4, value=item['sku'])
                sheet.cell(row=row_num, column=5, value=item['cn_name'])
                sheet.cell(row=row_num, column=6, value=item['en_name'])
                sheet.cell(row=row_num, column=7, value=item['quantity'])
                sheet.cell(row=row_num, column=8, value=item['price'])
                sheet.cell(row=row_num, column=9, value=item['total_value'])
                sheet.cell(row=row_num, column=10, value=item['material'])
                sheet.cell(row=row_num, column=11, value=item['hs_code'])
                sheet.cell(row=row_num, column=12, value=item['usage'])
                sheet.cell(row=row_num, column=13, value=item['brand'])
                sheet.cell(row=row_num, column=14, value=item['model'])
                sheet.cell(row=row_num, column=15, value=item['link'])

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
                # 设置单元格值
                sheet.cell(row=row_num, column=1, value=box_number)
                sheet.cell(row=row_num, column=2, value=box['box_spec'])
                sheet.cell(row=row_num, column=3, value=box['weight'])
                sheet.cell(row=row_num, column=4, value=item['sku'])
                sheet.cell(row=row_num, column=5, value=item['cn_name'])
                sheet.cell(row=row_num, column=6, value=item['en_name'])
                sheet.cell(row=row_num, column=7, value=item['quantity'])
                sheet.cell(row=row_num, column=8, value=item['price'])
                sheet.cell(row=row_num, column=9, value=item['total_value'])
                sheet.cell(row=row_num, column=10, value=item['material'])
                sheet.cell(row=row_num, column=11, value=item['hs_code'])
                sheet.cell(row=row_num, column=12, value=item['usage'])
                sheet.cell(row=row_num, column=13, value=item['brand'])
                sheet.cell(row=row_num, column=14, value=item['model'])
                sheet.cell(row=row_num, column=15, value=item['link'])

                # 应用样式
                for col in range(1, 16):
                    cell = sheet.cell(row=row_num, column=col)
                    cell.font = style_info['font']
                    cell.border = style_info['border']
                    cell.alignment = style_info['alignment']

                row_num += 1

        return row_num
