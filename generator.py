from numpy import add
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import os
import traceback
import json
from datetime import datetime
from db_connector import MongoDBConnector
from io import BytesIO
from openpyxl.utils import get_column_letter


class ProcessingError(Exception):
    """处理错误的自定义异常类"""
    pass


def template_handler(keyword):
    """模板处理器装饰器"""
    def decorator(func):
        func._template_keyword = keyword
        return func
    return decorator


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
        self.image_folder = os.path.join(current_dir, '产品图片(1)')  # 图片文件夹路径
        
        # 定义模板配置，只列出不需要编码的模板
        self.template_config = {
            "依诺达": {"requires_code": False}  # 不需要编码的模板
        }
        
        # 初始化模板处理器字典
        self._template_handlers = {}
        # 注册所有带有_template_keyword属性的方法
        print("开始注册模板处理器...")
        for name in dir(self):
            method = getattr(self, name)
            if hasattr(method, '_template_keyword'):
                keyword = method._template_keyword
                self._template_handlers[keyword] = method
                print(f"注册模板处理器: {name} -> {keyword}")
        print(f"已注册的模板处理器: {list(self._template_handlers.keys())}")

    @template_handler("叮铛卡航限时达")
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
                        # 填充收件人信息，这里收件人和
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=4, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=3, column=2)  # B2单元格
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

                # 检查所有产品的电磁属性
                has_electric = False
                has_magnetic = False
                for box in box_data.values():
                    for item in box.items:
                        product_info = self._get_product_info(item.msku, db)
                        if product_info:
                            if product_info.get('electrified', '') == '是':
                                has_electric = True
                            if product_info.get('magnetic', '') == '是':
                                has_magnetic = True
                            if has_electric and has_magnetic:
                                break
                    if has_electric and has_magnetic:
                        break

                # 在表格顶部添加电磁属性标记
                if has_electric:
                    cell = sheet.cell(row=1, column=6)  # F列第1行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)
                
                if has_magnetic:
                    cell = sheet.cell(row=2, column=6)  # F列第2行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)

                # 填充数据
                row_num = 18  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始

                # 遍历每个箱子

                sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    print(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            print(f"警告: 未找到产品 {item.msku} 的信息")
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 设置单元格值和样式ç
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (3,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (4, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            (5, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (6, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (7, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (8, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (9, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (10, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (11, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (12, product_info.get('link', '') if product_info else ''),
                            (14, ''),  # 图片列 (N列)
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
                                image_cell = f"N{row_num}"  # 图片列（第14列）
                                self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("顺丰")
    def _fill_sf_template(self, wb, box_data, code=None, address_info=None):
        """填充顺丰模板"""
        """
        填充顺丰模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        """
        with self.db_connector as db:
            try:
                sheet = wb['Sheet1']  # 获取模板工作表
                # box_Reference_id = ''  # 在方法开始时就初始化
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

                self.unmerge_cells_in_range(sheet, 2, 4, 2, 9)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info']
                    # if 'seller_info' in address_info:
                    #     box_Reference_id = address_info['seller_info']['amazonReferenceId']
                    try:
                        # 填充收件人信息
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=2, column=2)  
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=3, column=2)  
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

                        if address_parts:
                            cell = sheet.cell(row=4, column=2)  
                            cell.value = ', '.join(address_parts)
  
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                # 填充数据
                row_num = 12  
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[12].height

                # 将box_data按箱号排序
                sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    print(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # print(f"产品信息：{product_info}")
                        # price = product_info.get('price', 0)
                        # total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                        price = 0
                        total_price = 0
                        # if product_info:
                        #     item.product_name = product_info.get('cn_name', item.product_name)


                        # print(f"产品信息：{product_info}")
                        if product_info is not None:
                            item.product_name = product_info.get('cn_name', item.product_name)
                            print(f"产品信息：{product_info}")
                        else:
                        # 处理未找到产品信息的情况
                            print(f"未找到产品信息，MSKU: {item.msku}")
                            item.product_name = "需要补数据"  # 可以设置一个默认值
                        
                        # box_number_str = code+f"{box_number:05d}" 
                        box_number_str = code+'00000'+str(box_number)
                        # Reference_id = ''  # 初始化为None
                        # if box_Reference_id:
                        #     Reference_id = box_Reference_id
                        
                        # 设置单元格值和样式
                        cell_data = [
                            # 基本信息
                            (1, box_number_str),                                   # 货箱编号
                            # (2, Reference_id),                                    # 参考编号
                            (3, item.msku),                                       # 商品编码
                            # 产品名称信息
                            (4, product_info.get('en_name', '') if product_info else ''),                 # 英文名称
                            (5, product_info.get('cn_name', '') if product_info else ''),                 # 中文名称
                            (6, product_info.get('brand', '') if product_info else ''),                   # 品牌
                            (7, product_info.get('model', '') if product_info else ''),                   # 型号
                            
                            # 产品材料和用途
                            (8, product_info.get('material_cn', '') if product_info else ''),             # 中文材料
                            (9, product_info.get('material_en', '') if product_info else ''),             # 英文材料
                            (10, str(product_info.get('usage_en', '') + 
                                   product_info.get('usage_cn', '')) if product_info else ''),            # 用途
                            (11, '纸箱'),                                         # 包装类型
                            
                            # 产品规格信息
                            (12, product_info.get('hs_code', '') if product_info else ''),                # HS编码
                            (13, item.box_quantities.get(box_number, 0)),         # 数量
                            
                            # 箱体信息
                            (16, box.length if box.length is not None else ""),   # 长度
                            (17, box.width if box.width is not None else ""),     # 宽度
                            (18, box.height if box.height is not None else ""),   # 高度
                            (19, box.weight if box.weight is not None else ""),   # 重量
                            
                            # 其他信息
                            (20, product_info.get('link', '') if product_info else ''),                   # 链接
                            (21, '')                                              # 图片占位
                        ]
                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                        
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"U{row_num}"  # 图片列（第14列）
                                self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1
                
                self.merge_cells_in_range(sheet, 2, 2, 2, 9)
                self.merge_cells_in_range(sheet, 3, 3, 2, 9)
                self.merge_cells_in_range(sheet, 4, 4, 2, 9)

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("依诺达")
    def _fill_ynd_template(self, wb, box_data, code=None, address_info=None):
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
                self.unmerge_cells_in_range(sheet, 15, 15, 2, 4)
                self.unmerge_cells_in_range(sheet, 1, 1, 6, 8)
                self.unmerge_cells_in_range(sheet, 2, 2, 6, 8)
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

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=15, column=2)  #填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 检查所有产品的电磁属性
                has_electric = False
                has_magnetic = False
                for box in box_data.values():
                    for item in box.items:
                        product_info = self._get_product_info(item.msku, db)
                        if product_info:
                            if product_info.get('electrified', '') == '是':
                                has_electric = True
                            if product_info.get('magnetic', '') == '是':
                                has_magnetic = True
                            if has_electric and has_magnetic:
                                break
                    if has_electric and has_magnetic:
                        break

                # 在表格顶部添加电磁属性标记
                if has_electric:
                    cell = sheet.cell(row=1, column=6)  # F列第1行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)
                
                if has_magnetic:
                    cell = sheet.cell(row=2, column=6)  # F列第2行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)

                # 填充数据
                row_num = 17  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[17].height

                # 遍历每个箱子
                sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    print(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            print(f"警告: 未找到产品 {item.msku} 的信息")
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 设置单元格值和样式ç
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, box.weight if box.weight is not None else ""),  # 重量 
                            (3, box.length if box.length is not None else ""),  # 长度 
                            (4, box.width if box.width is not None else ""),    # 宽度 
                            (5, box.height if box.height is not None else "") ,  # 高度 
                            (6, item.msku),                    
                            (7,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (8, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            # (5, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (10, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (11, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (13, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (12, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (14, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (15, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            (16, product_info.get('link', '') if product_info else ''),
                            (17, ''),  # 图片列 (N列)
                            # (15, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                          
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"Q{row_num}"  # 图片列（第14列）
                                self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

                self.merge_cells_in_range(sheet, 15, 15, 2, 4)
                self.merge_cells_in_range(sheet, 1, 1, 6, 8)
                self.merge_cells_in_range(sheet, 2, 2, 6, 8)


            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("叮铛(美洲)")
    def _fill_ddmz_template(self, wb, box_data, code=None, address_info=None):
        """
        填充叮铛卡航限时达模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        """
        with self.db_connector as db:
            try:
                sheet = wb['清关发票']  # 获取模板工作表

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

                self.unmerge_cells_in_range(sheet, 3, 3, 8, 15)
                self.unmerge_cells_in_range(sheet, 4, 4, 8, 15)
                self.unmerge_cells_in_range(sheet, 5, 5, 8, 11)
                self.unmerge_cells_in_range(sheet, 5, 5, 13, 15)
                self.unmerge_cells_in_range(sheet, 6, 6, 8, 15)
                self.unmerge_cells_in_range(sheet, 7, 7, 8, 11)
                self.unmerge_cells_in_range(sheet, 7, 7,13, 15)
                self.unmerge_cells_in_range(sheet, 8, 8,13, 15)

                # 在第一行B列填充编码
                if code:
                    cell = sheet.cell(row=3, column=2)  # B列是第2列
                    cell.value = code
                    cell.font = Font(name='Arial', size=9)

                # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info']
                    try:
                        # 填充收件人信息，这里收件人和
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=3, column=8)  # B2单元格
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=4, column=8)  # B2单元格
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=8, column=13)  # B2单元格
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
                        # if 'addressLine1' in address_info_detail:
                        #     cell = sheet.cell(row=6, column=2)  # B3单元格
                        #     cell.value = address_info_detail['addressLine1']

                        # 城市
                        if 'city' in address_info_detail:
                            cell = sheet.cell(row=7, column=8)  # B4单元格
                            cell.value = address_info_detail['city']

                        #省份
                        if 'stateOrProvinceCode' in address_info_detail:
                            cell = sheet.cell(row=7, column=13)  # B5单元格
                            cell.value = address_info_detail['stateOrProvinceCode']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=5, column=13)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=5, column=8)  # B7单元格
                            cell.value = address_info_detail['countryCode']

                        if address_parts:
                            cell = sheet.cell(row=6, column=8)  # B3单元格
                            cell.value = ', '.join(address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=3, column=6)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")


                # 填充数据
                row_num = 12  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[12].height
                Reference_id = ''

                # 遍历每个箱子
                sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    print(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            print(f"警告: 未找到产品 {item.msku} 的信息")
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        Reference_id = address_info['address_info'].get('amazonReferenceId','')

                        # 设置单元格值和样式
                        cell_data = [
                            (1, code+f"{box_number:05d}"),                    
                            # (2, box.weight if box.weight is not None else ""),  
                            (2,Reference_id),  
                            (3,f"{box.length}*{box.width}*{box.height}"),  
                            (4,box_number), 
                            (5,box.weight), 
                            (6, product_info.get('hs_code', '') if product_info else ''),  
                            (8,product_info.get('en_name', '') if product_info else ''),  
                            (7, product_info.get('cn_name', '') if product_info else ''), 
                            (9, item.box_quantities.get(box_number, 0)),
                            (10,''),
                            (11, product_info.get('brand', '') if product_info else ''),    
                            (12, product_info.get('model', '') if product_info else ''),  
                            (13, str(product_info.get('material_cn', '')) if product_info else ''),  
                            (14, str(product_info.get('usage_cn', '')+product_info.get('usage_en', '' ))if product_info else ''),    # 用途 (H列)
                            (15, ''),  
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height

                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"O{row_num}"  # 图片列（第14列）
                                self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1
                
                self.merge_cells_in_range(sheet, 3, 3, 8, 15)
                self.merge_cells_in_range(sheet, 4, 4, 8, 15)
                self.merge_cells_in_range(sheet, 5, 5, 8, 11)
                self.merge_cells_in_range(sheet, 5, 5, 13, 15)
                self.merge_cells_in_range(sheet, 6, 6, 8, 15)
                self.merge_cells_in_range(sheet, 7, 7, 8, 11)
                self.merge_cells_in_range(sheet, 7, 7,13, 15)
                self.merge_cells_in_range(sheet, 8, 8,13, 15)
                
            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("UPS(美洲)")
    def _fill_ups_template(self, wb, box_data, code=None, address_info=None):
        """
        填充UPS美洲模板
        :param wb: 工作簿对象
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        """
        with self.db_connector as db:
            try:
                sheet = wb['发票']  # 获取发票工作表
                
                # 定义样式信息
                style_info = {
                    'font': Font(name='Arial', size=10),
                    'border': Border(left=Side(border_style='thin'),
                                     right=Side(border_style='thin'),
                                     top=Side(border_style='thin'),
                                     bottom=Side(border_style='thin')),
                    'alignment': Alignment(horizontal='center', vertical='center')
                }


                # 解除合并单元格
                print("正在解除合并单元格...")
                # self.unmerge_cells_in_range(sheet, 13, 16, 1, 15)
                self.unmerge_cells_in_range(sheet, 4, 4, 1, 3)
                self.unmerge_cells_in_range(sheet, 7, 11, 1, 3)
                self.unmerge_cells_in_range(sheet, 7, 11, 4, 15)
                
                row_num = 13  # 从第13行开始填充数据
                total_quantity = 0
                total_amount = 0
                
                
                if code:
                    cell = sheet.cell(row=4, column=1)  # B列是第2列
                    cell.value = f"运单号码:{code}"
                    cell.font = Font(name='Arial', size=12)

                    cell_another = sheet.cell(row=7, column=1)  # B列是第2列
                    cell_another.value = f"FBA编号:{code}"
                    cell_another.font = Font(name='Arial', size=12)
                
                  # 如果有地址信息，填充到相应的单元格
                if address_info:
                    address_info_detail = address_info['address_info']
                    try:

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

                        if address_parts:
                            cell = sheet.cell(row=7, column=4)  # B3单元格
                            cell.value = ', '.join(address_parts)
                            cell.font = Font(name='Arial', size=12)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")
            
                # 遍历每个箱子
                sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                # row_height = sheet.row_dimensions[12].height
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    if not box.items:
                        continue
                    
                    start_row = row_num  # 记录当前箱子的起始行
                    box_items_count = len(box.items)
                    
                    for item in box.items:
                        # 获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            print(f"警告: 未找到产品 {item.msku} 的信息")
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        
                        # 累计总数和总金额
                        total_quantity += item.box_quantities.get(box_number, 0)
                        total_amount += total_price
                        
                        # 设置单元格值
                        cell_data = [
                            (1, f"FBA{box_number}" if item == box.items[0] else ""),  # FBA号,只在第一行显示
                            (2, box_number if item == box.items[0] else ""),  # 箱号,只在第一行显示
                            (3, product_info.get('cn_name', '') if product_info else ''),  # 中文品名
                            (4, product_info.get('en_name', '') if product_info else ''),  # 英文品名
                            (5, price),  # 单价
                            (6, item.box_quantities.get(box_number, 0)),  # 数量
                            (7, total_price),  # 总价
                            (8, f"{product_info.get('material_cn', '')}/{product_info.get('material_en', '')}" if product_info else ''),  # 材质
                            (9, f"{product_info.get('usage_cn', '')}/{product_info.get('usage_en', '')}" if product_info else ''),  # 用途
                            (10, box.weight if item == box.items[0] else ""),  # 毛重,只在第一行显示
                            (11, box.length if item == box.items[0] else ""),  # 长,只在第一行显示
                            (12, box.width if item == box.items[0] else ""),  # 宽,只在第一行显示
                            (13, box.height if item == box.items[0] else ""),  # 高,只在第一行显示
                            (14, product_info.get('brand', '') if product_info else ''),  # 品牌
                            (15, product_info.get('hs_code', '') if product_info else '')  # HS编码
                        ]
                        
                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)
                        
                        row_num += 1
                    
                    # 如果这个箱子有多个产品,需要合并单元格
                    if box_items_count > 1:
                        merge_columns = [1, 2, 10, 11, 12, 13]  # 需要合并的列
                        for col in merge_columns:
                            self.merge_cells_in_range(sheet, start_row, row_num-1, col, col)

                # 添加总计行
                total_row = row_num
                self._set_cell_value(sheet, total_row, 1, "总件数", style_info)
                self._set_cell_value(sheet, total_row, 2, len(box_data), style_info)
                self._set_cell_value(sheet, total_row, 6, total_quantity, style_info)
                self._set_cell_value(sheet, total_row, 7, total_amount, style_info)
                self._set_cell_value(sheet, total_row, 9, "总重", style_info)
                
                # 添加Made in China
                made_in_row = total_row + 1
                self._set_cell_value(sheet, made_in_row, 1, "Made in China", style_info)
                
                # 添加日期
                date_row = made_in_row
                current_date = datetime.now().strftime("%Y.%m.%d")
                self._set_cell_value(sheet, date_row, 11, "DATE", style_info)
                self._set_cell_value(sheet, date_row, 12, f"签字日期:{current_date}", style_info)

                # 设置行高
                # for row in range(13, row_num):
                #     sheet.row_dimensions[row].height = 
                
                # self.merge_cells_in_range(sheet, 13, 16, 1, 15)
                self.merge_cells_in_range(sheet, 4, 4, 1, 3)
                self.merge_cells_in_range(sheet, 7, 11, 1, 3)
                self.merge_cells_in_range(sheet, 7, 11, 4, 15)

            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise

    @template_handler("林道")
    def _fill_lindao_template(self, wb, box_data, code=None, address_info=None):
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
                self.unmerge_cells_in_range(sheet, 3, 3, 2, 4)
                self.unmerge_cells_in_range(sheet, 4, 4, 2, 4)
                self.unmerge_cells_in_range(sheet, 5, 5, 2, 4)
                self.unmerge_cells_in_range(sheet, 6, 6, 2, 4)
                self.unmerge_cells_in_range(sheet, 9, 9, 2, 4)
                self.unmerge_cells_in_range(sheet, 11, 11, 2, 4)
                self.unmerge_cells_in_range(sheet, 12, 12, 2, 4)
                self.unmerge_cells_in_range(sheet, 15, 15, 2, 4)

                self.unmerge_cells_in_range(sheet, 13, 13, 6, 8)
                self.unmerge_cells_in_range(sheet, 14, 14, 6, 8)
                self.unmerge_cells_in_range(sheet, 15, 15, 6, 8)
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
                        # 填充收件人信息，这里收件人和
                        if 'name' in address_info_detail:
                            cell = sheet.cell(row=4, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=3, column=2)  # B2单元格
                            cell.value = address_info_detail['name']

                            cell = sheet.cell(row=5, column=2)  # B2单元格
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
                            cell = sheet.cell(row=9, column=2)  # B4单元格
                            cell.value = address_info_detail['city']

                        #邮政编码
                        if 'postalCode' in address_info_detail:
                            cell = sheet.cell(row=11, column=2)  # B6单元格
                            cell.value = address_info_detail['postalCode']

                        #国家代码
                        if 'countryCode' in address_info_detail:
                            cell = sheet.cell(row=12, column=2)  # B7单元格
                            cell.value = address_info_detail['countryCode']

                        if address_parts:
                            cell = sheet.cell(row=7, column=2)  # B3单元格
                            cell.value = ', '.join(address_parts)
                    except Exception as e:
                        print(f"填充地址信息时发生错误: {str(e)}")

                try:
                    total_boxes = len(box_data.keys())
                    cell = sheet.cell(row=15, column=2)  # 在第7行B列填充箱数
                    cell.value = str(total_boxes)
                    cell.font = Font(name='Arial', size=9)
                except Exception as e:
                    print(f"填充箱数时发生错误: {str(e)}")

                # 检查所有产品的电磁属性
                has_electric = False
                has_magnetic = False
                for box in box_data.values():
                    for item in box.items:
                        product_info = self._get_product_info(item.msku, db)
                        if product_info:
                            if product_info.get('electrified', '') == '是':
                                has_electric = True
                            if product_info.get('magnetic', '') == '是':
                                has_magnetic = True
                            if has_electric and has_magnetic:
                                break
                    if has_electric and has_magnetic:
                        break

                # 在表格顶部添加电磁属性标记
                if has_electric:
                    cell = sheet.cell(row=1, column=6)  # F列第1行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)
                
                if has_magnetic:
                    cell = sheet.cell(row=2, column=6)  # F列第2行
                    cell.value = "是"
                    cell.font = Font(name='Arial', size=9)

                # 填充数据
                row_num = 17  # 从第18行开始填充
                index = 1    # 添加序号计数器，从1开始
                row_height = sheet.row_dimensions[17].height

                # 遍历每个箱子

                sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))
                
                # 遍历排序后的箱子
                for box_number, box in sorted_boxes:
                    print(f"处理箱子 {box_number}")

                    # 遍历箱子中的每个产品
                    for item in box.items:
                        # 从数据库获取产品信息
                        product_info = self._get_product_info(item.msku, db)
                        # 处理产品信息为None的情况
                        if product_info is None:
                            print(f"警告: 未找到产品 {item.msku} 的信息")
                            price = 0
                            total_price = 0
                        else:
                            price = product_info.get('price', 0)
                            total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0
                            item.product_name = product_info.get('cn_name', item.product_name)
                        Reference_id = address_info['address_info'].get('amazonReferenceId','')
                        
                        # 设置单元格值和样式
                        cell_data = [
                            (1, box_number),                    # 货箱编号 (A列)
                            (2, code),
                            (3,Reference_id), 
                            (4, box.weight if box.weight is not None else ""),  # 重量 (B列)
                            (5, box.length if box.length is not None else ""),  # 长度 (Q列)
                            (6, box.width if box.width is not None else ""),    # 宽度 (R列)
                            (7, box.height if box.height is not None else ""),   # 高度 (S列)
                            (8,product_info.get('en_name', '') if product_info else ''),  # 链接 (D列)
                            (9, product_info.get('cn_name', '') if product_info else ''),  # 链接 (D列)
                            (10, product_info.get('price', '') if product_info else ''),   # 仅在总价格大于0时填入
                            (11,"美元"),  
                            (12, item.box_quantities.get(box_number, 0)),  # 数量 (F列)
                            (13, str(product_info.get('material_en', '')+'/'+product_info.get('material_cn', '')) if product_info else ''),  # 材料 (D列) 
                            (14, product_info.get('hs_code', '') if product_info else ''),  # HS编码 (G列)
                            (15, str(product_info.get('usage_en', '')+'/'+product_info.get('usage_cn', '' ))if product_info else ''),    # 用途 (H列)
                            (16, product_info.get('brand', '') if product_info else ''),    # 品牌 (I列)
                            (17, product_info.get('model', '') if product_info else ''),   # 型号 (J列)
                            # (12, product_info.get('link', '') if product_info else ''),
                            (18, ''),  # 图片列 (N列)
                            # (15, total_price if total_price > 0 else ""),  # 仅在总价格大于0时填入
                            # (24,item.SKU)
                        ]

                        # 批量设置单元格值和样式
                        for column, value in cell_data:
                            self._set_cell_value(sheet, row_num, column, value, style_info)

                        sheet.row_dimensions[row_num].height = row_height
                        # 插入产品图片
                        if item.msku and hasattr(self, 'image_folder'):
                            try:
                                image_cell = f"R{row_num}" 
                                self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
                            except Exception as e:
                                print(f"插入图片时发生错误: {str(e)}")

                        row_num += 1

                self.merge_cells_in_range(sheet, 3, 3, 2, 4)
                self.merge_cells_in_range(sheet, 4, 4, 2, 4)
                self.merge_cells_in_range(sheet, 5, 5, 2, 4)
                self.merge_cells_in_range(sheet, 6, 6, 2, 4)
                self.merge_cells_in_range(sheet, 9, 9, 2, 4)
                self.merge_cells_in_range(sheet, 11, 11, 2, 4)
                self.merge_cells_in_range(sheet, 12, 12, 2, 4)
                self.merge_cells_in_range(sheet, 15, 15, 2, 4)
                self.merge_cells_in_range(sheet, 13, 13, 6, 8)
                self.merge_cells_in_range(sheet, 14, 14, 6, 8)
                self.merge_cells_in_range(sheet, 15, 15, 6, 8)
            except Exception as e:
                print(f"填充模板时发生错误: {str(e)}")
                raise


    # @template_handler("林道")
    # def _fill_lindao_template(self, wb, box_data, code=None, address_info=None):
        # """
        # 填充林道模板
        # :param wb: 工作簿对象
        # :param box_data: 箱子数据
        # :param code: 编码（可选）
        # :param address_info: 地址信息（可选）
        # """
        # with self.db_connector as db:
        #     try:
        #         sheet = wb['模板']  # 获取模板工作表
        #         print("开始写入林道模版信息")

        #         # 记录第17行的格式信息
        #         row_height = sheet.row_dimensions[17].height
        #         cell_border = Border(left=Side(border_style='thin'),
        #                            right=Side(border_style='thin'),
        #                            top=Side(border_style='thin'),
        #                            bottom=Side(border_style='thin'))
        #         cell_alignment = Alignment(horizontal='center', vertical='center')
                
        #         # 获取原始单元格样式
        #         style_cell = sheet.cell(row=17, column=1)
        #         # 创建新的填充样式
        #         from openpyxl.styles import PatternFill, Color
                
        #         # 创建填充样式
        #         try:
        #             # 尝试获取原始填充样式的颜色
        #             if style_cell.fill and style_cell.fill.fgColor and style_cell.fill.fgColor.rgb:
        #                 fg_color = Color(rgb=style_cell.fill.fgColor.rgb)
        #             else:
        #                 fg_color = Color(rgb='FFFFFF')
                        
        #             if style_cell.fill and style_cell.fill.bgColor and style_cell.fill.bgColor.rgb:
        #                 bg_color = Color(rgb=style_cell.fill.bgColor.rgb)
        #             else:
        #                 bg_color = Color(rgb='FFFFFF')
                        
        #             pattern_type = style_cell.fill.patternType if style_cell.fill else 'solid'
                    
        #             cell_fill = PatternFill(
        #                 patternType=pattern_type,
        #                 fgColor=fg_color,
        #                 bgColor=bg_color
        #             )
        #         except Exception as e:
        #             print(f"创建填充样式时发生错误: {str(e)}")
        #             # 使用默认填充样式
        #             cell_fill = PatternFill(
        #                 patternType='solid',
        #                 fgColor=Color(rgb='FFFFFF'),
        #                 bgColor=Color(rgb='FFFFFF')
        #             )

        #         cell_font = Font(name=style_cell.font.name, 
        #                        size=style_cell.font.size, 
        #                        bold=style_cell.font.bold,
        #                        italic=style_cell.font.italic,
        #                        vertAlign=style_cell.font.vertAlign, 
        #                        color=style_cell.font.color)

        #         # 清理原有数据
        #         # 解除合并单元格
        #         merge_ranges = [
        #             'B17:B28', 'C17:C28', 'D17:D28', 'E17:E28',
        #             'B29:B41', 'C29:C41', 'D29:D41', 'E29:E41',
        #             'B42:B54', 'C42:C54', 'D42:D54', 'E42:E54',
        #             'B55:B63', 'C55:C63', 'D55:D63', 'E55:E63'
        #         ]
        #         for ranges in merge_ranges:
        #             if ranges in sheet.merged_cells:
        #                 sheet.unmerge_cells(ranges)

        #         # 重置行高
        #         for r in range(17, 64):
        #             sheet.row_dimensions[r].height = sheet.row_dimensions[70].height

        #         # 在第一行B列填充编码
        #         if code:
        #             cell = sheet.cell(row=1, column=2)
        #             cell.value = code
        #             cell.font = Font(name='Arial', size=9)

        #         # 填充地址信息
        #         if address_info:
        #             address_info_detail = address_info['address_info']
        #             try:
        #                 # 收件人信息
        #                 if 'name' in address_info_detail:
        #                     for row in [3, 4]:
        #                         cell = sheet.cell(row=row, column=2)
        #                         cell.value = address_info_detail['name']

        #                 # 地址信息
        #                 if 'addressLine1' in address_info_detail:
        #                     cell = sheet.cell(row=6, column=2)
        #                     cell.value = address_info_detail['addressLine1']

        #                 if 'city' in address_info_detail:
        #                     cell = sheet.cell(row=8, column=2)
        #                     cell.value = address_info_detail['city']

        #                 if 'postalCode' in address_info_detail:
        #                     cell = sheet.cell(row=10, column=2)
        #                     cell.value = address_info_detail['postalCode']

        #                 if 'countryCode' in address_info_detail:
        #                     cell = sheet.cell(row=11, column=2)
        #                     cell.value = address_info_detail['countryCode']

        #             except Exception as e:
        #                 print(f"填充地址信息时发生错误: {str(e)}")

        #         # 填充箱数
        #         try:
        #             total_boxes = len(box_data.keys())
        #             cell = sheet.cell(row=16, column=2)
        #             cell.value = str(total_boxes)
        #             cell.font = Font(name='Arial', size=9)
        #         except Exception as e:
        #             print(f"填充箱数时发生错误: {str(e)}")

        #         # 填充数据
        #         row_num = 17
        #         last_box_number = None
        #         merge_start_row = row_num

        #         # 对箱子进行排序
        #         sorted_boxes = sorted(box_data.items(), key=lambda x: int(x[0]))

        #         # 遍历每个箱子
        #         for box_number, box in sorted_boxes:
        #             print(f"处理箱子 {box_number}")
                    
        #             # 如果是新的箱子，需要处理上一个箱子的合并单元格
        #             if last_box_number is not None and last_box_number != box_number:
        #                 if row_num > merge_start_row:
        #                     # 合并上一个箱子的单元格
        #                     for col in range(2, 6):  # B到E列
        #                         sheet.merge_cells(start_row=merge_start_row, 
        #                                        start_column=col, 
        #                                        end_row=row_num-1, 
        #                                        end_column=col)
        #                 merge_start_row = row_num

        #             # 遍历箱子中的每个产品
        #             for item in box.items:
        #                 # 获取产品信息
        #                 product_info = self._get_product_info(item.msku, db)
                        
        #                 # 处理产品信息为None的情况
        #                 if product_info is None:
        #                     print(f"警告: 未找到产品 {item.msku} 的信息")
        #                     price = 0
        #                     total_price = 0
        #                 else:
        #                     price = product_info.get('price', 0)
        #                     total_price = float(price) * item.box_quantities.get(box_number, 0) if price else 0

        #                 # 设置单元格值
        #                 cell_data = [
        #                     (1, box_number),  # 箱号
        #                     (2, box.weight if box.weight is not None else ""),  # 重量
        #                     (3, box.length if box.length is not None else ""),  # 长
        #                     (4, box.width if box.width is not None else ""),   # 宽
        #                     (5, box.height if box.height is not None else ""), # 高
        #                     (7, product_info.get('cn_name', '')),  # 中文品名
        #                     (6, product_info.get('en_name', '')),  # 英文品名
        #                     (8, product_info.get('price', '')),    # 单价
        #                     (10, product_info.get('material_en', '')), # 英文材质
        #                     (12, product_info.get('usage_en', '')),   # 英文用途
        #                     (14, product_info.get('model', '')),      # 型号
        #                     (11, product_info.get('hs_code', '')),    # HS编码
        #                     (17, product_info.get('link', '')),       # 链接
        #                     (13, product_info.get('brand', '')),      # 品牌
        #                     (18, product_info.get('weight', '')),     # 单重
        #                     (9, item.box_quantities.get(box_number, 0)), # 数量
        #                     (16, total_price if total_price > 0 else ""), # 总价
        #                     (20, item.msku)  # SKU
        #                 ]

        #                 # 设置单元格值和样式
        #                 for col, value in cell_data:
        #                     cell = sheet.cell(row=row_num, column=col)
        #                     cell.value = value
        #                     cell.border = cell_border
        #                     cell.alignment = cell_alignment
        #                     cell.font = cell_font
        #                     # 只为特定列设置填充色
        #                     if col in [1, 17]:  # A列和Q列
        #                         cell.fill = cell_fill

        #                 # 插入产品图片
        #                 try:
        #                     image_cell = f"O{row_num}"
        #                     self.insert_product_image(sheet, image_cell, item.msku, self.image_folder)
        #                 except Exception as e:
        #                     print(f"插入图片时发生错误: {str(e)}")

        #                 row_num += 1
                    
        #             last_box_number = box_number

        #         # 处理最后一个箱子的合并单元格
        #         if row_num > merge_start_row:
        #             for col in range(2, 6):  # B到E列
        #                 sheet.merge_cells(start_row=merge_start_row, 
        #                                start_column=col, 
        #                                end_row=row_num-1, 
        #                                end_column=col)

        #         # 设置标题行
        #         sheet.cell(row=16, column=20, value="产品SKU")

        #         # 设置特殊列的填充色
        #         for row in range(17, row_num):
        #             for col in [1, 17]:  # A列和Q列
        #                 cell = sheet.cell(row=row, column=col)
        #                 cell.fill = cell_fill

        #         print("林道模板填充完成")

        #     except Exception as e:
        #         print(f"填充模板时发生错误: {str(e)}")
        #         traceback.print_exc()
        #         raise

    def _fill_default_template(self, wb, box_data, code=None, address_info=None):
        """默认的模板处理方法"""
        raise ProcessingError("未找到匹配的模板处理方法，请确保模板文件名包含正确的关键字")

    def generate_invoice(self, template_path, box_data,code=None, address_info=None):
        """
        生成发票
        :param template_path: 模板文件路径
        :param box_data: 箱子数据
        :param code: 编码（可选）
        :param address_info: 地址信息（可选）
        :return: 生成的发票文件路径
        """
        try:
            print(f"开始处理模板文件: {template_path}")
            if not os.path.exists(template_path):
                raise ProcessingError(f"模板文件不存在: {template_path}")

            # 获取当前时间戳
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            
            # 检查产品的电磁属性
            has_electric = False
            has_magnetic = False
            for box in box_data.values():
                for item in box.items:
                    product_info = self._get_product_info(item.msku, self.db_connector)
                    if product_info:
                        if product_info.get('electrified', '') == '是':
                            has_electric = True
                        if product_info.get('magnetic', '') == '是':
                            has_magnetic = True
                        if has_electric and has_magnetic:
                            break
                if has_electric and has_magnetic:
                    break
            
            # 构建文件名后缀
            suffix = ""
            if has_electric:
                suffix += "_带电"
            if has_magnetic:
                suffix += "_带磁"

            # 确保输出目录存在
            if not os.path.exists(self.output_folder):
                os.makedirs(self.output_folder)

            # 构建输出文件路径
            if address_info:
                try:
                    shipment_name = address_info['address_info']['shipmentName']
                    # 使用正则表达式提取数据
                    time, logistics, number = self.extract_data(shipment_name)
                    
                    if logistics is not None:
                        logistics = logistics.replace('-', '')  # 去掉破折号
                    else:
                        raise ValueError('Logistics cannot be None')

                    # 获取code
                    code_suffix = f"-{code}" if code else ""

                    output_filename = f'百泰{code_suffix}-{time}-{logistics}票-{number}件-{address_info["seller_info"]["country_name"]}-发票装箱单.xlsx'
                    # 替换任何可能导致路径问题的字符
                    output_filename = "".join(c for c in output_filename if c not in r'<>:"/\|?*')
                    output_path = os.path.join(self.output_folder, output_filename)

                except (TypeError, ValueError) as e:
                    # 处理异常的情况
                    print(f"Error occurred: {e}. Using default filename.")
                    output_filename = f"{timestamp}{suffix}.xlsx"
                    output_path = os.path.join(self.output_folder, output_filename)
            else:
                output_filename = f"{timestamp}{suffix}.xlsx"
                output_path = os.path.join(self.output_folder, output_filename)

            # 使用openpyxl加载模板
            print(f"正在加载模板文件...")
            wb = load_workbook(template_path)
            print(f"成功加载模板文件，工作表: {wb.sheetnames}")

            # 获取对应的模板处理方法
            template_handler = self._get_template_handler(template_path)
            if template_handler is None:
                raise ProcessingError(f"未找到对应的模板处理方法: {template_path}")

            # 处理模板
            template_handler(wb, box_data, code, address_info)

            # 保存文件
            wb.save(output_path)
            print(f"发票已生成: {output_path}")

            return output_path

        except Exception as e:
            error_msg = f"生成发票时发生错误: {str(e)}"
            print(error_msg)
            traceback.print_exc()
            raise ProcessingError(error_msg)

    def _get_template_handler(self, template_path):
        """根据模板文件名选择对应的处理方法"""
        try:
            template_name = os.path.basename(template_path).lower()
            # print(f"正在查找模板处理器，模板路径: {template_path}")
            # print(f"模板文件名: {template_name}")
            # print(f"已注册的处理器: {self._template_handlers}")
            
            for keyword, handler in self._template_handlers.items():
                # print(f"检查关键字: {keyword}, 类型: {type(keyword)}")
                # print(f"模板名称: {template_name}, 类型: {type(template_name)}")
                keyword_lower = keyword.lower()
                if keyword_lower in template_name:
                    print(f"找到匹配的处理器: {handler.__name__}")
                    return handler.__get__(self, type(self))
            print(f"未找到匹配的处理器，可用的关键字: {list(self._template_handlers.keys())}")
            return self._fill_default_template
        except Exception as e:
            print(f"模板处理器匹配过程中出错: {str(e)}")
            return self._fill_default_template

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
                    'usage_cn': product.get('useZh', ''),
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
            image_path_jpg = os.path.join(image_folder, f"{msku}.jpg")
            image_path_png = os.path.join(image_folder, f"{msku}.png")
            print(f"尝试加载图片: {image_path_jpg}")
            
            # 检查JPEG图片文件是否存在
            if os.path.exists(image_path_jpg):
                return self.insert_centered_image(worksheet, cell_address, image_path_jpg, fixed_width, fixed_height)
            elif os.path.exists(image_path_png):
                print(f"尝试加载PNG图片: {image_path_png}")
                return self.insert_centered_image(worksheet, cell_address, image_path_png, fixed_width, fixed_height)
            else:
                print(f"图片文件不存在: {image_path_jpg} 和 {image_path_png}")
                return False
        except Exception as e:
            print(f"处理产品图片时发生错误: {str(e)}")
            return False

    def extract_data(self, ticket_str):
        import re
        # 定义正则表达式，修改为提取斜杠后的数字
        pattern = r'(?P<time>\d{4}\.\d{2}\.\d{2})-(?P<logistics>[^-]+(?:-[^-]+)*)-(?P<number>\d+)/(\d+)'
        match = re.search(pattern, ticket_str)
        if match:
            # 使用第四个捕获组（斜杠后的数字）作为件数
            return match.group('time'), match.group('logistics').replace('-', ''), match.group(4)
        else:
            return None, None, None

    def unmerge_cells_in_range(self, sheet, start_row, end_row, start_col, end_col):
        """
        解除指定范围内的所有合并单元格。

        :param sheet: 要操作的工作表对象
        :param start_row: 起始行
        :param end_row: 结束行
        :param start_col: 起始列
        :param end_col: 结束列
        """
        print("正在解除合并单元格...")
        merged_ranges = list(sheet.merged_cells.ranges)
        for merged_range in merged_ranges:
            min_row, min_col, max_row, max_col = merged_range.bounds
            # 检查合并单元格是否在指定范围内
            if (min_row >= start_row and max_row <= end_row and
                min_col >= start_col and max_col <= end_col):
                try:
                    sheet.unmerge_cells(str(merged_range))
                except Exception as e:
                    print(f"解除合并单元格时发生错误: {str(e)}")
        print("合并单元格解除完成")


    def merge_cells_in_range(self, sheet, start_row, end_row, start_col, end_col):
        """
        合并指定区域内的单元格。

        :param sheet: 要操作的工作表对象
        :param start_row: 起始行
        :param end_row: 结束行
        :param start_col: 起始列
        :param end_col: 结束列
        """
        try:
            # 获取合并区域的范围字符串
            merge_range = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
            print(f"正在合并单元格区域: {merge_range}")
            
            # 合并单元格
            sheet.merge_cells(merge_range)
            
            # 设置合并后的单元格样式（可选）
            merged_cell = sheet.cell(row=start_row, column=start_col)
            merged_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            print(f"单元格合并完成: {merge_range}")
        except Exception as e:
            print(f"合并单元格时发生错误: {str(e)}")