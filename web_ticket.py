from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import threading
from queue import Queue
import time
from datetime import datetime, timedelta
import json
import shutil
import pandas as pd
from generator import InvoiceGenerator, ProcessingError
from get_ticket_data import PackingListProcessor

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'invoice_files', 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'invoice_files', 'output')
app.config['TEMPLATE_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), '表格模版')
app.config['HISTORY_FILE'] = 'history.json'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['MAX_HISTORY_DAYS'] = 30  # 历史记录保留天数

# 确保必要的目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 初始化发票生成器
invoice_generator = InvoiceGenerator(app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'])

# 创建任务队列和状态字典
task_queue = Queue()
task_status = {}
task_lock = threading.Lock()

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
        
        # 处理装箱单处理任务
        processor = PackingListProcessor(task_info['files'])
        box_data = processor.process()
        
        # 生成发票
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], f"{task_info['template_type']}.xlsx")
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{task_id}.xlsx")
        success, message = invoice_generator.generate_invoice(template_path, box_data, output_path, task_info.get('code'))
        
        if not success:
            raise ProcessingError(message)
        
        output_file = output_path
        
        with task_lock:
            task_status[task_id]['status'] = 'completed'
            task_status[task_id]['message'] = 'Processing completed'
            task_status[task_id]['output_file'] = os.path.basename(output_file)
        
        history = load_history()
        history.append({
            'task_id': task_id,
            'type': 'packing_list',
            'timestamp': datetime.now().strftime("%Y%m%d_%H%M%S"),
            'input_file': os.path.basename(task_info['files']),
            'output_file': os.path.basename(output_file),
            'code_input': task_info.get('code', ''),  # 添加编码
            'template_name': task_info.get('template_type', ''),  # 添加模板名称
            'status': 'completed',  # 添加状态
            'result_file': os.path.basename(output_file)  # 用于下载链接
        })
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
        if 'file' not in request.files:
            print("错误：没有上传文件")
            return jsonify({'error': '没有上传文件'}), 400
            
        file = request.files['file']
        if not file or not file.filename:
            print("错误：没有选择文件")
            return jsonify({'error': '没有选择文件'}), 400
            
        print(f"上传的文件: {file.filename}")
        print(f"文件类型: {file.content_type}")
        
        # 获取文件扩展名
        _, ext = os.path.splitext(file.filename)
        if not ext:
            ext = '.xlsx'  # 默认扩展名
        
        # 使用时间戳和编码生成新文件名
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        filename = f"{timestamp}_{code if code else 'upload'}{ext}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        print(f"将保存文件到: {file_path}")
        
        # 确保上传目录存在
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        
        # 保存上传的文件
        file.save(file_path)
        print(f"文件已保存，检查文件是否存在: {os.path.exists(file_path)}")
        if os.path.exists(file_path):
            print(f"文件大小: {os.path.getsize(file_path)} bytes")
        else:
            error_msg = "文件保存失败"
            print(error_msg)
            return jsonify({'error': error_msg}), 500
        
        # 根据模板类型选择模板文件
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], f'{template_type}.xlsx')
        print(f"使用模板: {template_path}")
        print(f"模板文件是否存在: {os.path.exists(template_path)}")
        
        # 验证模板文件是否存在
        if not os.path.exists(template_path):
            error_msg = f"模板文件不存在: {template_path}"
            print(error_msg)
            return jsonify({'error': error_msg}), 500
            
        # 生成任务ID
        task_id = datetime.now().strftime("%Y%m%d%H%M%S")
        print(f"任务ID: {task_id}")
        
        # 创建任务信息
        task_info = {
            'task_id': task_id,
            'template_path': template_path,
            'files': file_path,
            'code': code,
            'template_type': template_type
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
        elif task['status'] == 'failed':
            return jsonify({
                'status': 'failed',
                'error': task.get('error', '处理失败'),
                'message': '处理失败'
            })
        else:
            return jsonify({
                'status': 'processing',
                'message': '正在处理中'
            })

if __name__ == '__main__':
    app.run(debug=True, port=5009)
