from PIL import Image
from flask import Flask, render_template, request, jsonify, send_file, make_response, session
from werkzeug.utils import secure_filename
from io import BytesIO
import os
import threading
from queue import Queue
import time
from datetime import datetime, timedelta
import json
import shutil
import pandas as pd
from generator import InvoiceGenerator, ProcessingError
from get_ticket_data_fixed import PackingListProcessor, SimplePackingListProcessor
from STA_data import get_address_info
from db_utils import MongoDBClient
from image_extractor import ImageExtractor
import traceback
import logging
import numpy as np
import requests
import zipfile
from io import BytesIO
from login import run as get_token

# 配置日志
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(levelname)s - %(message)s')

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'invoice_files', 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'invoice_files', 'output')
app.config['TEMPLATE_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), '表格模版')
app.config['HISTORY_FILE'] = 'history.json'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['MAX_HISTORY_DAYS'] = 90  # 历史记录保留天数

# 确保必要的目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 初始化发票生成器
invoice_generator = InvoiceGenerator(app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'])

# 创建任务队列和状态字典
task_queue = Queue()
task_status = {}
task_lock = threading.Lock()

# 添加数据库操作锁
db_operation_lock = threading.Lock()

# 数据字段映射定义
DATA_FIELD_MAPPING = {
    "msku": "MSKU",
    "productNameZh": "中文品名",
    "productNameEn": "英文品名",
    "price": "价格",
    "brand": "品牌",
    "model": "型号",
    "HS": "海关编码",
    "image_url": "图片链接",
    "asin": "ASIN",
    "askPrice": "",  # 留空
    "electrified": "电",
    "magnetic": "磁",
    "materialEn": "英文材质",
    "materialZh": "中文材质",
    "outboundFee": "出库手续费",
    "productLink": "销售链接",
    "putAwayFee": "上架手续费",
    "useEn": "英文用途",
    "useZh": "中文用途",
    "weight": "重量",
    "X_ROW_K": "",  # 留空
    "created_at": "创建时间"
}

# 反向映射（Excel列名到数据库字段）
EXCEL_TO_DB_MAPPING = {v: k for k, v in DATA_FIELD_MAPPING.items() if v}


# ==================== 领星API相关函数 ====================

def request_web_download_packing_list(token, shipment_ids, need_down_image=0):
    """
    下载FBA货件装箱清单压缩包
    
    :param token: 认证token
    :param shipment_ids: 货件ID，可以是单个ID(str/int)或多个ID的列表
    :param need_down_image: 是否需要下载图片，0=不需要，1=需要
    :return: 压缩包的二进制数据(bytes)
    """
    # 处理shipment_ids参数
    if isinstance(shipment_ids, (list, tuple)):
        shipment_ids_str = ','.join(str(sid) for sid in shipment_ids)
    else:
        shipment_ids_str = str(shipment_ids)
    
    headers = {
        'AK-Client-Type': 'web',
        'AK-Origin': 'https://erp.lingxing.com',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
        'Referer': 'https://erp.lingxing.com/erp/msupply/fbaCargo',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36',
        'X-AK-Company-Id': '901217529031491584',
        'X-AK-ENV-KEY': 'SAAS-101',
        'X-AK-PLATFORM': '1',
        'X-AK-Request-Source': 'erp',
        'X-AK-Zid': '10330128',
        'auth-token': token,
        'sec-ch-ua': '"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
    }
    
    url = f'https://erp.lingxing.com/api/fba_shipment/batchDownloadPackingList?shipmentIds={shipment_ids_str}&need_down_image={need_down_image}&req_time_sequence=%2Fapi%2Ffba_shipment%2FcheckShipmentPackingListStatus$$1'
    
    response = requests.get(url, headers=headers)
    
    # 检查响应状态
    if response.status_code == 200:
        return response.content  # 返回二进制数据
    else:
        raise Exception(f"下载失败，状态码: {response.status_code}, 响应: {response.text[:200]}")


def extract_excel_from_zip(zip_data_or_path, output_path=None):
    """
    直接从ZIP中提取Excel文件（简化版）
    
    :param zip_data_or_path: ZIP二进制数据(bytes)或ZIP文件路径(str)
    :param output_path: 输出Excel文件路径，默认自动生成
    :return: 提取的Excel文件路径
    """
    # 判断输入类型
    if isinstance(zip_data_or_path, bytes):
        zip_file = zipfile.ZipFile(BytesIO(zip_data_or_path))
    else:
        zip_file = zipfile.ZipFile(zip_data_or_path)
    
    # 查找Excel文件
    excel_file = None
    for file_info in zip_file.filelist:
        if file_info.filename.lower().endswith(('.xlsx', '.xls')):
            excel_file = file_info.filename
            break
    
    if not excel_file:
        zip_file.close()
        raise Exception("ZIP中未找到Excel文件")
    
    # 确定输出路径
    if output_path is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        ext = os.path.splitext(excel_file)[1]
        output_path = f'packing_list_{timestamp}{ext}'
    
    # 提取Excel文件
    with zip_file.open(excel_file) as source:
        with open(output_path, 'wb') as target:
            target.write(source.read())
    
    zip_file.close()
    
    print(f"Excel文件已提取: {os.path.abspath(output_path)}")
    return output_path


def download_packing_list_excel_simple(token, shipment_ids, output_path=None, need_down_image=0):
    """
    下载FBA装箱清单并直接提取Excel（简化版，适用于确定ZIP里就是Excel的情况）
    
    :param token: 认证token
    :param shipment_ids: 货件ID
    :param output_path: 输出Excel文件路径，默认自动生成
    :param need_down_image: 是否需要下载图片
    :return: Excel文件路径
    """
    print(f"正在下载货件 {shipment_ids} 的装箱清单...")
    
    # 下载ZIP数据
    zip_data = request_web_download_packing_list(token, shipment_ids, need_down_image)
    print(f"下载成功，文件大小: {len(zip_data) / (1024 * 1024):.2f} MB")
    
    # 直接提取Excel
    excel_path = extract_excel_from_zip(zip_data, output_path)
    
    return excel_path


def request_web_FBA_shipment_num(token, shipment_id):
    """
    通过货件编码查询货件的内部ID
    
    :param token: 认证token
    :param shipment_id: 货件编码（如 FBA193LX49TS）
    :return: 货件内部ID，失败返回2
    """
    # 获取当前日期
    today = datetime.now()
    # 开始日期：3个月前
    start_date = (today - timedelta(days=100)).strftime('%Y-%m-%d')
    # 结束日期：3个月后
    end_date = (today + timedelta(days=100)).strftime('%Y-%m-%d')
    
    headers = {
        'AK-Client-Type': 'web',
        'AK-Origin': 'https://erp.lingxing.com',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/json;charset=UTF-8',
        'Origin': 'https://erp.lingxing.com',
        'Pragma': 'no-cache',
        'Referer': 'https://erp.lingxing.com/erp/msupply/fbaCargo',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36',
        'X-AK-Company-Id': '901217529031491584',
        'X-AK-ENV-KEY': 'SAAS-101',
        'X-AK-Language': 'zh',
        'X-AK-PLATFORM': '1',
        'X-AK-Request-Id': '144118e5-1876-4cf5-9917-4f8b882d5fed',
        'X-AK-Request-Source': 'erp',
        'X-AK-Uid': '10431785',
        'X-AK-Version': '3.7.1.3.0.128',
        'X-AK-Zid': '10330128',
        'auth-token': token,
        'sec-ch-ua': '"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
    }

    json_data = {
        'search_field_time': 'create_date',
        'is_sta': '',
        'is_awd': '',
        'ship_mode': '',
        'is_closed': '',
        'step': [],
        'application_diff': '',
        'received_diff': '',
        'application_received_diff': '',
        'is_relate_packing_task_sn': '',
        'has_shipto_address': '',
        'is_shipto_diff': '',
        'is_add_tracking': '',
        'delivery_order_status': [],
        'is_transparency': '',
        'is_print_transparency': '',
        'box_type': '',
        'is_uploaded_box': '',
        'sta_transportation_mode': '',
        'create_uids': [],
        'is_store_diff': '',
        'is_update_shipment_tracking_no': '',
        'search_field': 'shipment_id',
        'search_value': shipment_id,
        'shipment_status': [],
        'is_relate_shipment': '',
        'start_date': start_date,
        'end_date': end_date,
        'seniorSearchList': [],
        'shipment_type': [],
        'offset': 0,
        'length': 500,
        'req_time_sequence': '/api/fba_shipment/showShipment_v2$$7',
    }

    try:
        response = requests.post(
            'https://erp.lingxing.com/api/fba_shipment/showShipment_v2',
            headers=headers,
            json=json_data,
            timeout=30  # 添加超时设置
        )
        response.raise_for_status()  # 检查HTTP错误
        data = response.json()
        
        if data.get('code') == 1 and data.get('data', {}).get('list'):
            return data['data']['list'][0]['id']
        return 2  # 默认错误码
    except (requests.RequestException, ValueError, KeyError, IndexError) as e:
        print(f"请求失败: {str(e)}")
        return 2

# ==================== 领星API相关函数结束 ====================


def clean_old_files():
    """清理旧文件和历史记录"""
    try:
        current_time = datetime.now()
        cutoff_date = current_time - timedelta(days=app.config['MAX_HISTORY_DAYS'])

        # 加载历史记录
        history = load_history()
        new_history = []
        files_to_keep = set()

        # 遍历历史记录
        for record in history:
            try:
                record_time = datetime.strptime(record['timestamp'], "%Y%m%d_%H%M%S")
                if record_time > cutoff_date:
                    new_history.append(record)
                    if 'output_file' in record:
                        files_to_keep.add(record['output_file'])
            except (ValueError, KeyError):
                continue

        # 保存更新后的历史记录
        save_history(new_history)

        # 清理结果文件夹中的旧文件
        for filename in os.listdir(app.config['OUTPUT_FOLDER']):
            if filename not in files_to_keep:
                try:
                    os.remove(os.path.join(app.config['OUTPUT_FOLDER'], filename))
                except OSError:
                    continue

        # 清理上传文件夹
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            try:
                if os.path.getctime(file_path) < cutoff_date.timestamp():
                    os.remove(file_path)
            except OSError:
                continue

    except Exception as e:
        print(f"清理文件时发生错误: {str(e)}")


def load_history():
    """加载处理历史记录"""
    try:
        if os.path.exists(app.config['HISTORY_FILE']):
            with open(app.config['HISTORY_FILE'], 'r', encoding='utf-8') as f:
                return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    return []


def save_history(history):
    """保存处理历史记录"""
    try:
        with open(app.config['HISTORY_FILE'], 'w', encoding='utf-8') as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存历史记录时发生错误: {str(e)}")


def process_task(task_info):
    """处理任务"""
    task_id = task_info['task_id']

    try:
        with task_lock:
            task_status[task_id]['status'] = 'processing'
            task_status[task_id]['message'] = 'Processing started'

        # 根据文件格式选择处理器并处理
        if task_info.get('is_simple_format', False):
            processor = SimplePackingListProcessor(task_info['files'])
            template_type = task_info.get('template_type', '')
            box_data = processor.process(template_name=template_type)
            shipment_id = processor.shipment_id  # 获取shipment_id
        else:
            processor = PackingListProcessor(task_info['files'])
            box_data = processor.process()
            shipment_id = None

        if not box_data:
            raise ProcessingError("处理装箱单失败")

        # 根据模板类型决定是否需要处理编码
        template_type = task_info.get('template_type', '')
        code = task_info.get('code')
        address_info = None

        # 检查模板是否需要编码
        template_config = invoice_generator.template_config.get(template_type, {})
        requires_code = template_config.get('requires_code', True)  # 默认需要编码

        if requires_code:
            if not code:
                print(f"警告：模板 {template_type} 需要编码，但未提供编码")
            else:
                try:
                    address_info = get_address_info(code)
                    if address_info:
                        print(f"获取到地址信息: {address_info}")
                    else:
                        print(f"未能获取到地址信息，将继续生成发票")
                except Exception as e:
                    print(f"获取地址信息时发生错误: {str(e)}，将继续生成发票")
                    # 记录错误但不影响发票生成
                    pass
        else:
            print(f"模板 {template_type} 不需要编码，跳过地址信息获取")

        # 生成发票
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], f"{task_info['template_type']}.xlsx")
        try:
            output_paths = []
            
            # 检查是否为多地址
            if address_info and isinstance(address_info, list):
                # 多地址情况：创建专门的文件夹来组织多个文件
                print(f"检测到多地址情况，共 {len(address_info)} 个地址")
                
                # 创建多地址文件夹
                multi_folder_name = f"{code}_多地址发票_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                multi_folder_path = os.path.join(app.config['OUTPUT_FOLDER'], multi_folder_name)
                os.makedirs(multi_folder_path, exist_ok=True)
                print(f"创建多地址文件夹: {multi_folder_path}")
                
                # 临时修改输出文件夹到多地址专用文件夹
                original_output_folder = invoice_generator.output_folder
                invoice_generator.output_folder = multi_folder_path
                
                try:
                    for i, addr in enumerate(address_info):
                        print(f"正在处理第 {i+1}/{len(address_info)} 个地址...")
                        output_path = invoice_generator.generate_invoice(template_path, box_data, code, addr, shipment_id=shipment_id)
                        if output_path:
                            # 修复路径问题：确保记录的是相对于多地址文件夹的文件名，而不是完整路径
                            relative_filename = os.path.basename(output_path)
                            output_paths.append(relative_filename)
                            print(f"第 {i+1} 个地址的发票生成成功: {relative_filename}")
                        else:
                            print(f"第 {i+1} 个地址的发票生成失败")
                finally:
                    # 恢复原始输出文件夹
                    invoice_generator.output_folder = original_output_folder
                
                if output_paths:
                    print(f"多地址发票生成完成，共生成 {len(output_paths)} 个文件")
                    print(f"所有文件已保存到文件夹: {multi_folder_name}")
                    
                    with task_lock:
                        task_status[task_id]['status'] = 'completed'
                        task_status[task_id]['message'] = f'Processing completed - {len(output_paths)} files generated in folder: {multi_folder_name}'
                        task_status[task_id]['output_files'] = [os.path.basename(path) for path in output_paths]
                        task_status[task_id]['output_folder'] = multi_folder_name  # 新增：记录文件夹名称
                        # 为了兼容性，设置文件夹作为主输出
                        task_status[task_id]['output_file'] = multi_folder_name
                else:
                    # 如果没有生成任何文件，删除空文件夹
                    try:
                        os.rmdir(multi_folder_path)
                        print(f"删除空文件夹: {multi_folder_path}")
                    except:
                        pass
                    raise ProcessingError("所有地址的发票生成都失败")
            else:
                # 单地址情况：保持原有逻辑
                output_path = invoice_generator.generate_invoice(template_path, box_data, code, address_info, shipment_id=shipment_id)
                if output_path:
                    output_paths.append(output_path)
                    print(f"发票生成成功: {output_path}")
                    with task_lock:
                        task_status[task_id]['status'] = 'completed'
                        task_status[task_id]['message'] = 'Processing completed'
                        task_status[task_id]['output_file'] = os.path.basename(output_path)
                else:
                    raise ProcessingError("发票生成失败")
        except Exception as e:
            error_msg = f"处理任务时发生错误: {str(e)}"
            print(error_msg)
            with task_lock:
                task_status[task_id]['status'] = 'error'
                task_status[task_id]['message'] = error_msg
                task_status[task_id]['error'] = str(e)

        history = load_history()
        history_record = {
            'task_id': task_id,
            'type': 'packing_list',
            'timestamp': datetime.now().strftime("%Y%m%d_%H%M%S"),
            'input_file': os.path.basename(task_info['files']),
            'output_file': task_status[task_id].get('output_file'),
            'code_input': code,
            'template_name': task_info.get('template_type', ''),
            'status': task_status[task_id]['status'],
            'result_file': task_status[task_id].get('output_file')
        }
        
        # 如果是多地址情况，记录所有生成的文件
        if 'output_files' in task_status[task_id]:
            history_record['output_files'] = task_status[task_id]['output_files']
            history_record['files_count'] = len(task_status[task_id]['output_files'])

        # 如果获取地址信息失败，记录到历史记录中
        if code and not address_info:
            history_record['address_info_status'] = 'failed'

        history.append(history_record)
        save_history(history)

    except Exception as e:
        error_msg = f"Error processing task {task_id}: {str(e)}"
        print(error_msg)
        with task_lock:
            task_status[task_id]['status'] = 'error'
            task_status[task_id]['message'] = error_msg
            task_status[task_id]['error'] = str(e)

        # 在发生错误时也保存到历史记录
        history = load_history()
        history.append({
            'task_id': task_id,
            'type': 'packing_list',
            'timestamp': datetime.now().strftime("%Y%m%d_%H%M%S"),
            'input_file': os.path.basename(task_info['files']),
            'code_input': task_info.get('code', ''),
            'template_name': task_info.get('template_type', ''),
            'status': 'failed',
            'error': str(e)
        })
        save_history(history)


def process_worker():
    """处理任务队列的工作线程"""
    while True:
        try:
            task_info = task_queue.get()
            if task_info is None:
                break

            process_task(task_info)
        except Exception as e:
            print(f"工作线程出错: {str(e)}")
        finally:
            task_queue.task_done()


# 启动工作线程
NUM_WORKER_THREADS = 3
worker_threads = []
for _ in range(NUM_WORKER_THREADS):
    t = threading.Thread(target=process_worker, daemon=True)
    t.start()
    worker_threads.append(t)


@app.route('/')
def index():
    """渲染主页"""
    return render_template('index.html')


@app.route('/msku_edit')
def msku_edit():
    """渲染主页"""
    return render_template('msku_edit.html')


@app.route('/history')
def get_history():
    """获取处理历史记录"""
    clean_old_files()  # 清理旧文件
    history = load_history()

    # 获取查询参数
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    code = request.args.get('code')

    # 过滤记录
    filtered_history = []
    for record in history:
        # 时间过滤
        record_date = record['timestamp'].split('_')[0]  # 获取日期部分
        if start_date and record_date < start_date:
            continue
        if end_date and record_date > end_date:
            continue

        # 编码过滤
        if code and code.lower() not in record.get('code_input', '').lower():
            continue

        filtered_history.append(record)

    # 按时间戳降序排序
    filtered_history.sort(key=lambda x: x['timestamp'], reverse=True)

    return jsonify(filtered_history)


@app.route('/download/<filename>')
def download_file(filename):
    """下载处理结果文件或文件夹"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        print(f"下载请求: {filename}")
        print(f"完整路径: {file_path}")
        print(f"路径存在: {os.path.exists(file_path)}")
        print(f"是否为目录: {os.path.isdir(file_path)}")
        
        # 检查是否为文件夹（多地址情况）
        if os.path.isdir(file_path):
            print(f"检测到多地址文件夹，开始创建ZIP")
            # 列出文件夹内容
            try:
                folder_contents = os.listdir(file_path)
                print(f"文件夹内容: {folder_contents}")
            except Exception as e:
                print(f"无法读取文件夹内容: {e}")
            
            # 创建ZIP文件
            import zipfile
            zip_filename = f"{filename}.zip"
            zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                file_count = 0
                for root, dirs, files in os.walk(file_path):
                    for file in files:
                        file_path_in_zip = os.path.join(root, file)
                        # 在ZIP中保持相对路径结构
                        arcname = os.path.relpath(file_path_in_zip, file_path)
                        zipf.write(file_path_in_zip, arcname)
                        file_count += 1
                        print(f"添加文件到ZIP: {arcname}")
            
            print(f"创建ZIP文件: {zip_path}, 包含 {file_count} 个文件")
            
            def remove_zip():
                """下载完成后删除临时ZIP文件"""
                try:
                    if os.path.exists(zip_path):
                        os.remove(zip_path)
                        print(f"删除临时ZIP文件: {zip_path}")
                except:
                    pass
            
            # 使用Flask的after_request来在响应发送后删除临时文件
            response = send_file(
                zip_path,
                as_attachment=True,
                download_name=zip_filename
            )
            
            # 注册清理函数
            @response.call_on_close
            def cleanup():
                remove_zip()
            
            return response
        else:
            # 普通文件下载
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename
            )
    except Exception as e:
        return jsonify({'error': f'下载文件失败: {str(e)}'}), 404


@app.route('/upload', methods=['POST'])
def upload():
    """处理文件上传"""
    try:
        print("\n=== 开始处理上传请求 ===")
        print(f"请求表单数据: {request.form}")
        print(f"请求文件: {request.files}")

        # 获取模板类型
        template_type = request.form.get('template_type', 'dingdang')  # 默认使用叮铛模板
        print(f"模板类型: {template_type}")

        # 获取编码（可选）
        code = request.form.get('code', '')
        print(f"编码: {code}")

        # 检查是否有文件上传
        packing_list = request.files.get('packing_list')
        invoice_info = request.files.get('invoice_info')

        if not packing_list and not invoice_info:
            print("错误：没有上传任何文件")
            return jsonify({'error': '请至少上传一个文件'}), 400

        # 确保上传目录存在
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

        # 保存并处理文件
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        file_path = None
        is_simple_format = False

        if packing_list and packing_list.filename:
            # 处理领星装箱单
            _, ext = os.path.splitext(packing_list.filename)
            if not ext:
                ext = '.xlsx'
            filename = f"{timestamp}_packing_list{ext}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            packing_list.save(file_path)
            is_simple_format = False
        elif invoice_info and invoice_info.filename:
            # 处理简单格式装箱单
            _, ext = os.path.splitext(invoice_info.filename)
            if not ext:
                ext = '.xlsx'
            filename = f"{timestamp}_invoice_info{ext}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            invoice_info.save(file_path)
            is_simple_format = True

        # 生成任务ID
        task_id = datetime.now().strftime("%Y%m%d%H%M%S")
        print(f"任务ID: {task_id}")

        # 创建任务信息
        task_info = {
            'task_id': task_id,
            'template_path': None,
            'files': file_path,
            'code': code,
            'template_type': template_type,
            'is_simple_format': is_simple_format
        }
        print(f"任务信息: {task_info}")

        # 初始化任务状态
        with task_lock:
            task_status[task_id] = {
                'status': 'pending',
                'created_at': datetime.now().strftime('%Y%m%d_%H%M%S')
            }

        # 将任务添加到队列
        task_queue.put(task_info)
        print(f"任务已添加到队列")
        print("=== 上传处理完成 ===\n")

        return jsonify({
            'success': True,
            'message': '文件已上传，正在处理中',
            'task_id': task_id
        })

    except Exception as e:
        error_msg = f"处理上传请求时出错: {str(e)}"
        print(error_msg)
        return jsonify({'error': error_msg}), 500


@app.route('/auto_fetch_shipment', methods=['POST'])
def auto_fetch_shipment():
    """通过货件编码自动获取装箱单并处理"""
    try:
        print("\n=== 开始处理自动获取货件请求 ===")
        
        # 获取参数
        shipment_code = request.form.get('shipment_code', '').strip()
        template_type = request.form.get('template_type', 'dingdang')
        code = request.form.get('code', '')  # 地址编码（可选）
        
        print(f"货件编码: {shipment_code}")
        print(f"模板类型: {template_type}")
        print(f"地址编码: {code}")
        
        # 验证必填参数
        if not shipment_code:
            return jsonify({'error': '请输入货件编码'}), 400
        
        if not template_type:
            return jsonify({'error': '请选择模板类型'}), 400
        
        # 获取 token
        print("正在获取登录 token...")
        try:
            token = get_token()
            print("Token 获取成功")
        except Exception as e:
            error_msg = f"获取 token 失败: {str(e)}"
            print(error_msg)
            return jsonify({'error': error_msg}), 500
        
        # 步骤1: 通过货件编码获取内部 ID
        print(f"正在查询货件 {shipment_code} 的内部 ID...")
        try:
            shipment_id = request_web_FBA_shipment_num(token, shipment_code)
            if shipment_id == 2:  # 错误码
                error_msg = f"未找到货件 {shipment_code}，请检查货件编码是否正确"
                print(error_msg)
                return jsonify({'error': error_msg}), 404
            print(f"获取到内部 ID: {shipment_id}")
        except Exception as e:
            error_msg = f"查询货件 ID 失败: {str(e)}"
            print(error_msg)
            return jsonify({'error': error_msg}), 500
        
        # 步骤2: 下载装箱单 Excel
        print(f"正在下载货件 {shipment_id} 的装箱单...")
        try:
            # 生成临时文件路径
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            filename = f"{timestamp}_auto_fetch_{shipment_code}.xlsx"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            
            # 下载并保存 Excel
            excel_path = download_packing_list_excel_simple(
                token, 
                shipment_id, 
                output_path=file_path
            )
            
            if not excel_path or not os.path.exists(excel_path):
                error_msg = "下载装箱单失败"
                print(error_msg)
                return jsonify({'error': error_msg}), 500
            
            print(f"装箱单下载成功: {excel_path}")
        except Exception as e:
            error_msg = f"下载装箱单失败: {str(e)}"
            print(error_msg)
            traceback.print_exc()
            return jsonify({'error': error_msg}), 500
        
        # 步骤3: 创建处理任务（复用现有逻辑）
        task_id = datetime.now().strftime("%Y%m%d%H%M%S")
        print(f"任务ID: {task_id}")
        
        # 创建任务信息（使用领星装箱单格式）
        task_info = {
            'task_id': task_id,
            'template_path': None,
            'files': excel_path,
            'code': code,
            'template_type': template_type,
            'is_simple_format': False  # 领星装箱单是详细格式
        }
        print(f"任务信息: {task_info}")
        
        # 初始化任务状态
        with task_lock:
            task_status[task_id] = {
                'status': 'pending',
                'created_at': datetime.now().strftime('%Y%m%d_%H%M%S'),
                'shipment_code': shipment_code,
                'auto_fetch': True  # 标记为自动获取
            }
        
        # 将任务添加到队列
        task_queue.put(task_info)
        print(f"任务已添加到队列")
        print("=== 自动获取处理完成 ===\n")
        
        return jsonify({
            'success': True,
            'message': f'货件 {shipment_code} 的装箱单已自动获取，正在处理中',
            'task_id': task_id,
            'shipment_id': shipment_id
        })
        
    except Exception as e:
        error_msg = f"自动获取货件时出错: {str(e)}"
        print(error_msg)
        traceback.print_exc()
        return jsonify({'error': error_msg}), 500


@app.route('/api/generate_invoice', methods=['POST'])
def generate_invoice():
    try:
        # 获取JSON数据
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400

        # 生成任务ID
        task_id = f"invoice_{int(time.time())}"

        # 保存JSON数据到临时文件
        json_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}.json")
        with open(json_file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # 初始化任务状态
        with task_lock:
            task_status[task_id] = {
                'status': 'pending',
                'message': 'Task queued',
                'output_file': None,
                'error': None
            }

        # 创建任务信息
        task_info = {
            'task_id': task_id,
            'type': 'invoice',
            'input_file': json_file_path,
            'timestamp': datetime.now().strftime("%Y%m%d_%H%M%S")
        }

        # 将任务添加到队列
        task_queue.put(task_info)

        return jsonify({
            'task_id': task_id,
            'status': 'pending',
            'message': 'Task queued successfully'
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/status/<task_id>')
def get_status(task_id):
    """获取任务处理状态"""
    with task_lock:
        task = task_status.get(task_id)
        if not task:
            return jsonify({'error': '任务不存在'}), 404

        if task['status'] == 'completed':
            # 检查output_file是否存在且有效
            output_file = task.get('output_file')
            if not output_file:
                return jsonify({
                    'status': 'failed',
                    'error': '输出文件不存在',
                    'message': '处理失败'
                })

            # 如果任务完成，返回文件下载链接
            return jsonify({
                'status': 'completed',
                'download_url': f'/download/{os.path.basename(output_file)}',
                'message': '处理完成'
            })
        elif task['status'] in ['failed', 'error']:
            return jsonify({
                'status': 'failed',
                'error': task.get('error', '处理失败'),
                'message': '处理失败'
            })
        elif task['status'] == 'processing':
            return jsonify({
                'status': 'processing',
                'message': '正在处理中'
            })
        else:
            # 未知状态当作失败处理
            return jsonify({
                'status': 'failed',
                'error': '未知状态',
                'message': '处理失败'
            })


@app.route('/api/get_msku_info/', methods=['POST'])
def get_msku_info():
    try:
        with invoice_generator.db_connector as db:
            collection = db['msku_info']
            page = request.json.get('page', 1)
            page_size = request.json.get('pageSize', 50)
            filters = request.json.get('filters', {})
            products = collection.find(filters).sort('_id',-1).skip((page - 1) * page_size).limit(page_size)
            results = []
            for i in products:
                results.append({
                    field[0]: i.get(field[0], None) for field in FIELDS
                })
            count = collection.count_documents(filters)
            print(count)
        return jsonify(status='success', data=results, total=count)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/save_msku_info/', methods=['POST'])
def save_msku_info():
    try:
        with invoice_generator.db_connector as db:
            collection = db['msku_info']
            data = request.json
            collection.update_one({
                "msku": data.get("msku")},
                {'$set': data},
                upsert=True)
        return jsonify(status='success')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/images/<string:msku>', methods=['GET'])
def show_image(msku):
    try:
        image_path_jpg = os.path.join(invoice_generator.image_folder, f"{msku}.jpg")
        image_path_png = os.path.join(invoice_generator.image_folder, f"{msku}.png")
        if os.path.exists(image_path_jpg):
            file_min = Image.open(image_path_jpg)

        elif os.path.exists(image_path_png):
            file_min = Image.open(image_path_jpg)
        else:
            image_data = open("static/no.png", "rb").read()
            response = make_response(image_data)
            response.headers['Content-Type'] = 'image/jpg'
            return response
        # 获取原图尺寸
        w, h = file_min.size
        # 计算压缩比
        bili = int(w / 300)
        if bili == 0:
            bili = 1
        # 按比例对宽高压缩
        file_min.thumbnail((w // bili, h // bili))
        bytesIO = BytesIO()
        file_min.save(bytesIO, format='PNG')
        response = make_response(bytesIO.getvalue())
        response.headers['Content-Type'] = 'image/jpg'
        return response
    except Exception as e:
        image_data = open("static/no.png", "rb").read()
        response = make_response(image_data)
        response.headers['Content-Type'] = 'image/jpg'
        return response

@app.route('/api/upload/', methods=['POST'])
def upload_image():
    if 'file' not in request.files:
        print("no file")
    file = request.files['file']
    msku = request.form['msku']
    filename = f"{msku}.jpg"
    file.save(os.path.join(invoice_generator.image_folder, filename))
    return {"code": 200, "name": filename,
            "url": f"/api/images/{msku}"}, 200
    # return {"code": 200, "name": filename,
    #         "url": f"https://em-erp-1252538772.cos.ap-nanjing.myqcloud.com/{filename}"},  200

def transform_data(row):
    """转换Excel行数据为MongoDB文档格式"""
    def convert_to_none(value):
        """将空值转换为None"""
        if pd.isna(value):  # 检查是否为NaN
            return None
        # 处理numpy数值类型
        if isinstance(value, (np.integer, np.floating)):
            return value.item()  # 转换为Python原生类型
        return value

    document = {
        'msku': str(convert_to_none(row['MSKU'])),  # 确保MSKU是字符串类型
        'productNameZh': convert_to_none(row['中文品名']),
        'productNameEn': convert_to_none(row['英文品名']),
        'price': convert_to_none(row['价格']),
        'materialZh': convert_to_none(row['中文材质']),
        'materialEn': convert_to_none(row['英文材质']),
        'useZh': convert_to_none(row['中文用途']),
        'useEn': convert_to_none(row['英文用途']),
        'model': convert_to_none(row['型号']),
        'HS': convert_to_none(row['海关HS编码']),
        'productLink': convert_to_none(row['商品链接']),
        'electrified': convert_to_none(row['是否带电']),
        'magnetic': convert_to_none(row['是否带磁']),
        'brand': convert_to_none(row['品牌']),
        'weight': convert_to_none(row['重量']),
        'asin': convert_to_none(row['ASIN']),
        'putAwayFee': convert_to_none(row['上架手续费']),
        'outboundFee': convert_to_none(row['出库手续费'])
    }
    
    # 添加创建时间
    document['created_at'] = datetime.now()
    
    return document

def process_excel_import(task_id, file_path):
    """处理Excel导入任务"""
    try:
        # 更新任务状态
        with task_lock:
            task_status[task_id] = {
                'status': 'processing',
                'progress': 0,
                'message': '正在读取Excel文件...',
                'success_count': 0,
                'skip_count': 0,
                'error_count': 0,
                'error_records': []
            }

        logging.info(f'开始处理任务 {task_id}')
        
        # 读取Excel文件
        try:
            df = pd.read_excel(file_path)
            logging.info(f'成功读取Excel文件，共 {len(df)} 行数据')
        except Exception as e:
            logging.error(f'读取Excel文件失败: {str(e)}')
            logging.error(traceback.format_exc())
            raise

        total_records = len(df)

        # 创建MongoDB连接
        try:
            db_client = MongoDBClient()
            db_client.connect()
            logging.info('成功连接到数据库')
        except Exception as e:
            logging.error(f'数据库连接失败: {str(e)}')
            logging.error(traceback.format_exc())
            raise

        try:
            # 获取现有的msku列表
            with db_operation_lock:
                existing_mskus = set(doc['msku'] for doc in db_client.db['msku_info'].find({}, {'msku': 1}))
                logging.info(f'获取到 {len(existing_mskus)} 个现有MSKU')

            # 预处理数据
            documents_to_insert = []
            for index, row in df.iterrows():
                try:
                    msku = str(row['MSKU'])
                    if msku in existing_mskus:
                        task_status[task_id]['skip_count'] += 1
                        logging.info(f'跳过已存在的MSKU: {msku}')
                    else:
                        document = transform_data(row)
                        documents_to_insert.append(document)
                        task_status[task_id]['success_count'] += 1
                        logging.info(f'处理MSKU成功: {msku}')

                except Exception as e:
                    task_status[task_id]['error_count'] += 1
                    error_msg = f'处理MSKU时出错: {str(e)}'
                    task_status[task_id]['error_records'].append({
                        'msku': row.get('MSKU', 'Unknown'),
                        'error': error_msg
                    })
                    logging.error(error_msg)
                    logging.error(traceback.format_exc())

                # 更新进度
                progress = int((index + 1) / total_records * 100)
                task_status[task_id]['progress'] = progress
                task_status[task_id]['message'] = f'已处理 {index + 1}/{total_records} 条记录'

            # 批量插入数据
            if documents_to_insert:
                with db_operation_lock:
                    try:
                        # 再次检查MSKU是否存在
                        current_mskus = set(doc['msku'] for doc in db_client.db['msku_info'].find({}, {'msku': 1}))
                        documents_to_insert = [doc for doc in documents_to_insert if doc['msku'] not in current_mskus]
                        
                        if documents_to_insert:
                            db_client.db['msku_info'].insert_many(documents_to_insert)
                            logging.info(f'成功插入 {len(documents_to_insert)} 条数据')
                    except Exception as e:
                        logging.error(f'批量插入数据失败: {str(e)}')
                        logging.error(traceback.format_exc())
                        raise

        finally:
            db_client.close()
            logging.info('数据库连接已关闭')

        # 完成处理
        task_status[task_id]['status'] = 'completed'
        task_status[task_id]['message'] = '导入完成'
        logging.info(f'任务 {task_id} 处理完成')

    except Exception as e:
        error_msg = f'导入失败: {str(e)}'
        logging.error(error_msg)
        logging.error(traceback.format_exc())
        task_status[task_id]['status'] = 'error'
        task_status[task_id]['message'] = error_msg

    finally:
        # 清理临时文件
        try:
            os.remove(file_path)
            logging.info(f'临时文件 {file_path} 已删除')
        except Exception as e:
            logging.error(f'删除临时文件失败: {str(e)}')


@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    """处理Excel文件上传"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400

        # 转换为小写后检查扩展名，支持.xlsx, .xls, .XLSX, .XLS等
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({'error': '请上传Excel文件（.xlsx 或 .xls 格式）'}), 400

        # 保存文件
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        temp_filename = f'import_{timestamp}_{filename}'
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], temp_filename)
        file.save(file_path)

        # 创建任务
        task_id = f'import_{timestamp}'
        threading.Thread(target=process_excel_import, args=(task_id, file_path)).start()

        return jsonify({'task_id': task_id})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/import_status/<task_id>')
def import_status(task_id):
    """获取导入任务状态"""
    with task_lock:
        if task_id not in task_status:
            return jsonify({'error': '任务不存在'}), 404
        
        status_data = task_status[task_id].copy()
        
        # 如果任务已完成，清理状态数据
        if status_data['status'] in ['completed', 'error']:
            task_status.pop(task_id, None)
        
        return jsonify(status_data)


# 图片上传功能路由
@app.route('/extract_images', methods=['POST'])
def extract_images():
    """处理图片上传任务"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400

        # 转换为小写后检查扩展名，支持.xlsx, .xls, .XLSX, .XLS等
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({'error': '请上传Excel文件（.xlsx 或 .xls 格式）'}), 400

        # 保存文件
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        temp_filename = f'images_{timestamp}_{filename}'
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], temp_filename)
        file.save(file_path)

        # 创建任务
        task_id = f'images_{timestamp}'
        with task_lock:
            task_status[task_id] = {
                'status': 'processing',
                'progress': 0,
                'message': '准备处理...',
                'timestamp': timestamp
            }

        # 启动后台线程处理任务
        threading.Thread(target=process_image_extraction, args=(task_id, file_path)).start()

        return jsonify({'task_id': task_id})

    except Exception as e:
        logging.error(f"图片上传任务创建失败: {str(e)}")
        logging.error(traceback.format_exc())
        return jsonify({'error': str(e)}), 500


def process_image_extraction(task_id, file_path):
    """处理图片上传任务"""
    try:
        # 初始化图片提取器
        image_extractor = ImageExtractor(app.config['UPLOAD_FOLDER'])
        
        # 获取任务状态引用
        with task_lock:
            if task_id not in task_status:
                task_status[task_id] = {
                    'status': 'processing',
                    'progress': 0,
                    'message': '准备处理...',
                }
        
        # 执行图片提取
        result = image_extractor.extract_images_from_excel(file_path, task_status[task_id])
        
        # 更新任务状态
        with task_lock:
            if result['success']:
                task_status[task_id]['status'] = 'completed'
                task_status[task_id]['progress'] = 100
                task_status[task_id]['message'] = '处理完成'
                task_status[task_id]['success_count'] = result['success_count']
                task_status[task_id]['error_count'] = result['error_count']
            else:
                task_status[task_id]['status'] = 'error'
                task_status[task_id]['message'] = result.get('error', '处理失败')
    
    except Exception as e:
        logging.error(f"图片上传处理失败: {str(e)}")
        logging.error(traceback.format_exc())
        with task_lock:
            if task_id in task_status:
                task_status[task_id]['status'] = 'error'
                task_status[task_id]['message'] = f'处理失败: {str(e)}'
    
    finally:
        # 清理临时文件
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logging.info(f'临时文件 {file_path} 已删除')
        except Exception as e:
            logging.error(f'删除临时文件失败: {str(e)}')


@app.route('/extract_images_status/<task_id>')
def extract_images_status(task_id):
    """获取图片上传任务状态"""
    with task_lock:
        if task_id not in task_status:
            return jsonify({'error': '任务不存在'}), 404
        
        status_data = task_status[task_id].copy()
        
        # 如果任务已完成，清理状态数据
        if status_data['status'] in ['completed', 'error']:
            task_status.pop(task_id, None)
        
        return jsonify(status_data)


@app.route('/export_data')
def export_data():
    """导出数据库数据为Excel文件"""
    try:
        # 创建 MongoDB连接
        db_client = MongoDBClient()
        db_client.connect()
        
        try:
            # 查询所有MSKU数据
            cursor = db_client.db['msku_info'].find({})
            
            # 创建Excel工作簿
            from openpyxl import Workbook
            
            wb = Workbook()
            ws = wb.active
            ws.title = "MSKU数据"
            
            # 写入第一行：数据库字段名（隐藏的元数据）
            db_fields = list(DATA_FIELD_MAPPING.keys())
            ws.append(db_fields)
            
            # 写入第二行：中文列名（用户看到的表头）
            chinese_headers = [DATA_FIELD_MAPPING[field] for field in db_fields]
            ws.append(chinese_headers)
            
            # 写入数据行
            row_count = 0
            for doc in cursor:
                row_data = []
                for field in db_fields:
                    value = doc.get(field, "")
                    
                    # 特殊字段处理
                    if field == "created_at":
                        if isinstance(value, datetime):
                            value = value.strftime('%Y-%m-%d %H:%M:%S')
                        elif value:
                            value = str(value)
                        else:
                            value = ""
                    elif field in ["askPrice", "X_ROW_K"]:
                        value = ""  # 留空字段
                    else:
                        value = str(value) if value is not None else ""
                    
                    row_data.append(value)
                
                ws.append(row_data)
                row_count += 1
            
            # 调整列宽
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # 最大宽度50
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # 保存文件
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'msku_export_{timestamp}.xlsx'
            file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
            wb.save(file_path)
            
            logging.info(f'成功导出 {row_count} 条MSKU记录到 {filename}')
            
            # 返回文件下载
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
        finally:
            db_client.close()
            
    except Exception as e:
        logging.error(f'数据导出失败: {str(e)}')
        logging.error(traceback.format_exc())
        return jsonify({'error': f'导出失败: {str(e)}'}), 500


@app.route('/upload_update', methods=['POST'])
def upload_update():
    """处理数据更新上传"""
    try:
        logging.info(f"收到上传请求，request.files: {list(request.files.keys())}")
        
        if 'file' not in request.files:
            logging.error("请求中没有file字段")
            return jsonify({'error': '没有上传文件'}), 400

        file = request.files['file']
        logging.info(f"文件对象: {file}, 文件名: {file.filename}")
        
        if file.filename == '':
            logging.error("文件名为空")
            return jsonify({'error': '没有选择文件'}), 400

        # 转换为小写后检查扩展名，支持.xlsx, .xls, .XLSX, .XLS等
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            logging.error(f"文件格式不支持: {file.filename}")
            return jsonify({'error': '请上传Excel文件（.xlsx 或 .xls 格式）'}), 400

        # 保存文件
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        temp_filename = f'update_{timestamp}_{filename}'
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], temp_filename)
        file.save(file_path)

        # 创建任务
        task_id = f'update_{timestamp}'
        with task_lock:
            task_status[task_id] = {
                'status': 'processing',
                'progress': 0,
                'message': '准备处理...',
                'success_count': 0,
                'update_count': 0,
                'insert_count': 0,
                'error_count': 0,
                'error_records': [],
                'timestamp': timestamp
            }

        # 启动后台线程处理任务
        threading.Thread(target=process_data_update, args=(task_id, file_path)).start()

        return jsonify({'task_id': task_id})

    except Exception as e:
        logging.error(f"数据更新任务创建失败: {str(e)}")
        logging.error(traceback.format_exc())
        return jsonify({'error': str(e)}), 500


def process_data_update(task_id, file_path):
    """处理数据更新任务"""
    try:
        logging.info(f'开始处理数据更新任务 {task_id}')
        
        # 更新任务状态
        with task_lock:
            task_status[task_id]['message'] = '正在读取Excel文件...'
        
        # 读取Excel文件
        try:
            df = pd.read_excel(file_path)
            logging.info(f'成功读取Excel文件，共 {len(df)} 行数据')
        except Exception as e:
            logging.error(f'读取Excel文件失败: {str(e)}')
            raise
        
        # 验证Excel格式
        if len(df) < 1:
            raise ValueError('Excel文件格式错误：至少需要1行数据')
        
        # 创建字段映射（使用DataFrame的列名）
        field_mapping = {}
        for i, chinese_name in enumerate(df.columns):
            if pd.notna(chinese_name) and chinese_name in DATA_FIELD_MAPPING.values():
                # 找到对应的数据库字段名
                for db_field, display_name in DATA_FIELD_MAPPING.items():
                    if display_name == chinese_name:
                        field_mapping[i] = db_field
                        break
        
        logging.info(f'识别到字段映射: {field_mapping}')
        
        # 检查是否包含必要字段
        if 'msku' not in field_mapping.values():
            raise ValueError('Excel文件格式错误：缺少MSKU字段')
        
        # 获取数据行（所有行都是数据）
        data_rows = df
        total_records = len(data_rows)
        
        if total_records == 0:
            raise ValueError('Excel文件中没有数据行')
        
        # 创建MongoDB连接
        db_client = MongoDBClient()
        db_client.connect()
        
        # 创建备份（用于回滚）
        backup_id = f'backup_{task_id}'
        backup_data = []
        
        # 准备事务操作列表
        operations = []
        
        try:
            with db_operation_lock:
                # 第一阶段：数据验证和操作准备
                for index, row in data_rows.iterrows():
                    try:
                        # 解析数据
                        document = {}
                        msku = None
                        
                        for col_idx, db_field in field_mapping.items():
                            value = row.iloc[col_idx] if col_idx < len(row) else None
                            
                            # 跳过留空字段
                            if db_field in ["askPrice", "X_ROW_K"]:
                                continue
                            
                            # 使用与现有批量上传相同的数据处理方式
                            def convert_to_none(val):
                                """将空值转换为None"""
                                if pd.isna(val):  # 检查是否为NaN
                                    return None
                                # 处理numpy数值类型
                                if isinstance(val, (np.integer, np.floating)):
                                    return val.item()  # 转换为Python原生类型
                                return val
                            
                            # 特殊处理created_at字段
                            if db_field == "created_at":
                                if value and str(value).strip():
                                    try:
                                        # 尝试解析时间格式
                                        if isinstance(value, str):
                                            value = datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
                                        elif not isinstance(value, datetime):
                                            value = datetime.now()
                                    except:
                                        value = datetime.now()
                                else:
                                    value = datetime.now()
                            else:
                                # 其他字段使用统一的转换方式
                                value = convert_to_none(value)
                            
                            document[db_field] = value
                            
                            if db_field == "msku":
                                msku = value
                        
                        if not msku:
                            task_status[task_id]['error_count'] += 1
                            task_status[task_id]['error_records'].append({
                                'row': index + 3,  # Excel行号
                                'error': 'MSKU不能为空'
                            })
                            continue
                        
                        # 查找现有记录
                        existing_doc = db_client.db['msku_info'].find_one({'msku': msku})
                        
                        if existing_doc:
                            # 准备更新操作
                            document['updated_at'] = datetime.now()
                            if 'created_at' not in document:
                                document['created_at'] = existing_doc.get('created_at', datetime.now())
                            
                            operations.append({
                                'type': 'update',
                                'msku': msku,
                                'document': document,
                                'backup_data': {
                                    'operation': 'update',
                                    'msku': msku,
                                    'original_data': existing_doc
                                }
                            })
                            
                        else:
                            # 准备插入操作
                            document['created_at'] = document.get('created_at', datetime.now())
                            document['updated_at'] = datetime.now()
                            
                            operations.append({
                                'type': 'insert',
                                'msku': msku,
                                'document': document,
                                'backup_data': {
                                    'operation': 'insert',
                                    'msku': msku,
                                    'original_data': None
                                }
                            })
                        
                    except Exception as e:
                        task_status[task_id]['error_count'] += 1
                        error_msg = f'处理第{index + 3}行数据时出错: {str(e)}'
                        task_status[task_id]['error_records'].append({
                            'row': index + 3,
                            'msku': row.iloc[0] if len(row) > 0 else 'Unknown',
                            'error': error_msg
                        })
                        continue
            
            # 第二阶段：执行所有操作（不使用事务，因为单机MongoDB不支持）
            try:
                # 执行所有操作
                for op in operations:
                    if op['type'] == 'update':
                        # 执行更新操作
                        result = db_client.db['msku_info'].replace_one(
                            {'msku': op['msku']}, 
                            op['document']
                        )
                        if result.matched_count == 0:
                            raise Exception(f"更新MSKU {op['msku']} 失败：记录不存在")
                        
                        task_status[task_id]['update_count'] += 1
                        logging.info(f'更新MSKU: {op["msku"]}')
                        
                    elif op['type'] == 'insert':
                        # 执行插入操作
                        result = db_client.db['msku_info'].insert_one(
                            op['document']
                        )
                        if not result.inserted_id:
                            raise Exception(f"插入MSKU {op['msku']} 失败")
                        
                        # 记录插入的ID用于回滚
                        op['backup_data']['inserted_id'] = result.inserted_id
                        task_status[task_id]['insert_count'] += 1
                        logging.info(f'插入新MSKU: {op["msku"]}')
                
                # 更新进度
                with task_lock:
                    task_status[task_id]['progress'] = 90
                    task_status[task_id]['message'] = '正在保存备份数据...'
                
                # 保存备份数据
                if backup_data:
                    db_client.db['backup_data'].insert_one({
                        'backup_id': backup_id,
                        'task_id': task_id,
                        'timestamp': datetime.now(),
                        'backup_data': backup_data,
                        'status': 'completed'
                    })
                    logging.info(f'备份数据已保存，备份ID: {backup_id}')
                
                logging.info('数据操作执行成功')
            
            except Exception as e:
                # 不需要回滚，因为没有使用事务
                logging.error(f'数据操作执行失败: {str(e)}')
                raise e
        
        finally:
            db_client.close()
        
        # 完成处理
        task_status[task_id]['status'] = 'completed'
        task_status[task_id]['message'] = '数据更新完成'
        task_status[task_id]['backup_id'] = backup_id
        logging.info(f'数据更新任务 {task_id} 处理完成')
        
    except Exception as e:
        error_msg = f'数据更新失败: {str(e)}'
        logging.error(error_msg)
        logging.error(traceback.format_exc())
        task_status[task_id]['status'] = 'error'
        task_status[task_id]['message'] = error_msg
    
    finally:
        # 清理临时文件
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logging.info(f'临时文件 {file_path} 已删除')
        except Exception as e:
            logging.error(f'删除临时文件失败: {str(e)}')


@app.route('/update_status/<task_id>')
def update_status(task_id):
    """获取数据更新任务状态"""
    with task_lock:
        if task_id not in task_status:
            return jsonify({'error': '任务不存在'}), 404
        
        status_data = task_status[task_id].copy()
        
        # 如果任务已完成，不立即清理状态数据（保留用于回滚）
        return jsonify(status_data)


@app.route('/sync_erp_database', methods=['POST'])
def sync_erp_database():
    """从ERP同步产品数据到MongoDB"""
    task_id = datetime.now().strftime('%Y%m%d%H%M%S')
    
    with task_lock:
        task_status[task_id] = {
            'status': 'processing',
            'progress': 0,
            'message': '正在初始化...',
            'inserted': 0,
            'skipped': 0,
            'errors': 0
        }
    
    # 在后台线程中执行同步
    thread = threading.Thread(target=sync_erp_data_task, args=(task_id,))
    thread.daemon = True
    thread.start()
    
    return jsonify({'task_id': task_id})


def sync_erp_data_task(task_id):
    """ERP数据同步任务（后台执行）"""
    try:
        # 导入同步模块
        from sync_erp_to_mongodb import ERPProductSync
        import login
        
        # 更新状态
        with task_lock:
            task_status[task_id]['message'] = '正在登录ERP系统...'
        
        # 登录
        token = login.run()
        
        # 更新状态
        with task_lock:
            task_status[task_id]['message'] = '正在连接数据库...'
        
        # 创建同步对象
        sync = ERPProductSync(token)
        
        # 更新状态
        with task_lock:
            task_status[task_id]['message'] = '正在获取产品数据...'
        
        # 执行同步（不限制数量，同步所有产品）
        stats = sync.sync_products(limit=None)
        
        # 更新最终状态
        with task_lock:
            task_status[task_id].update({
                'status': 'completed',
                'progress': 100,
                'message': '同步完成',
                'inserted': stats['inserted'],
                'skipped': stats['skipped'],
                'errors': stats['errors'],
                'total': stats['total'],
                'processed': stats['processed'],
                'msku_count': stats['msku_count']
            })
        
        logging.info(f'ERP数据同步完成: {stats}')
        
    except Exception as e:
        error_msg = f'同步失败: {str(e)}'
        logging.error(error_msg)
        logging.error(traceback.format_exc())
        
        with task_lock:
            task_status[task_id].update({
                'status': 'error',
                'message': error_msg,
                'error': str(e)
            })


@app.route('/sync_status/<task_id>')
def sync_status(task_id):
    """获取ERP同步任务状态"""
    with task_lock:
        if task_id not in task_status:
            return jsonify({'error': '任务不存在'}), 404
        
        return jsonify(task_status[task_id])


if __name__ == '__main__':
    os.makedirs(invoice_generator.image_folder,exist_ok=True)
    FIELDS = [
        ["msku", 0],  # sku必有
        ["productNameZh", 1],
        ["productNameEn", 2],
        ["price", 3],
        ["materialZh", 4],
        ["materialEn", 5],
        ["useZh", 6],
        ["useEn", 7],
        ["model", 8],
        ["HS", 9],
        ["productLink", 10],
        ["electrified", 11],
        ["magnetic", 12],
        ["brand", 13],
        ["weight", 14],
        ["asin", 15],
        ["putAwayFee", 16],
        ["outboundFee", 17]
    ]
    app.run(host="0.0.0.0", port=5009, debug=False)
