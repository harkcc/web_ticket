from PIL import Image
from flask import Flask, render_template, request, jsonify, send_file, make_response
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
from get_ticket_data import PackingListProcessor, SimplePackingListProcessor
from STA_data import get_address_info
from db_utils import MongoDBClient
import traceback
import logging
import numpy as np

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
            output_path = invoice_generator.generate_invoice(template_path, box_data, code, address_info, shipment_id=shipment_id)
            if output_path:
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
            'output_file': os.path.basename(output_path) if 'output_file' in task_status[task_id] else None,
            'code_input': code,
            'template_name': task_info.get('template_type', ''),
            'status': task_status[task_id]['status'],
            'result_file': os.path.basename(output_path) if 'output_file' in task_status[task_id] else None
        }

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
    """下载处理结果文件"""
    try:
        return send_file(
            os.path.join(app.config['OUTPUT_FOLDER'], filename),
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
                    msku = row['MSKU']
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

        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': '请上传Excel文件'}), 400

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
    app.run(host="0.0.0.0", port=5009, debug=True)
