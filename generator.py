import os
import traceback
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
        if "叮铛" in template_name:
            return self._fill_dingdang_template
        elif "德国" in template_name:
            return self._fill_germany_template
        elif "法国" in template_name:
            return self._fill_france_template
        elif "英国" in template_name:
            return self._fill_uk_template
        else:
            return self._fill_default_template

    def generate_invoice(self, template_path, box_data, output_path, code=None):
        """生成发票"""
        try:
            print(f"\n=== 开始生成发票 ===")
            print(f"模板文件: {template_path}")
            print(f"输出路径: {output_path}")
            print(f"编码: {code}")
            print(f"箱子数据: {box_data}")

            # 使用openpyxl加载模板
            print(f"正在加载模板文件...")
            wb = openpyxl.load_workbook(template_path)
            print(f"成功加载模板文件，工作表: {wb.sheetnames}")

            # 获取并使用对应的模板处理方法
            template_handler = self._get_template_handler(template_path)
            template_handler(wb, box_data)

            # 保存文件
            print(f"正在保存文件到: {output_path}")
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            wb.save(output_path)
            print(f"文件保存成功")

            return True, "发票生成成功"

        except Exception as e:
            error_msg = f"生成发票时发生错误: {str(e)}"
            print(error_msg)
            traceback.print_exc()
            return False, error_msg

    def _fill_dingdang_template(self, wb, box_data):
        """填充叮铛卡航限时达模板"""
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
            except:
                pass
        print(f"合并单元格解除完成")

        # 填充数据
        row_num = 1  # 从第一行开始填充
        for box_number, box in box_data.items():
            print(f"处理箱子 {box_number}")

            # 填充箱子信息
            for item in box.items:
                # 序号
                cell = sheet.cell(row=row_num, column=1)
                cell.value = row_num
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # MSKU
                cell = sheet.cell(row=row_num, column=2)
                cell.value = item.msku
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 品名
                cell = sheet.cell(row=row_num, column=3)
                cell.value = item.product_name
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                # 数量
                cell = sheet.cell(row=row_num, column=4)
                cell.value = item.box_quantities.get(box_number, 0)
                cell.font = style_info['font']
                cell.border = style_info['border']
                cell.alignment = style_info['alignment']

                row_num += 1

            # 在每个箱子后面添加一个空行
            row_num += 1

        return row_num

    def _fill_germany_template(self, wb, box_data):
        """填充德国模板"""
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

    def _fill_france_template(self, wb, box_data):
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

    def _fill_uk_template(self, wb, box_data):
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

    def _fill_default_template(self, wb, box_data):
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