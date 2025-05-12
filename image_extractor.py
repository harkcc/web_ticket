import pandas as pd
import os
import requests
import paramiko
from io import BytesIO
from PIL import Image
from db_config import SSH_CONFIG
import logging
import traceback
from datetime import datetime

# 配置日志
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(levelname)s - %(message)s')

class ImageExtractor:
    """图片提取器，用于从Excel文件中提取图片URL，下载图片并上传到远程服务器"""
    
    def __init__(self, upload_folder=None):
        """初始化图片提取器
        
        Args:
            upload_folder: 上传文件的临时存储目录
        """
        self.upload_folder = upload_folder
        self.remote_image_folder = "/root/web_ticket/web_ticket/产品图片(1)"
    
    def upload_to_remote(self, sftp, local_data, remote_path):
        """上传数据到远程服务器"""
        try:
            # 确保远程目录存在
            remote_dir = os.path.dirname(remote_path)
            try:
                sftp.stat(remote_dir)
            except FileNotFoundError:
                # 递归创建目录
                current_dir = ''
                for dir_part in remote_dir.split('/'):
                    if dir_part:
                        current_dir += '/' + dir_part
                        try:
                            sftp.stat(current_dir)
                        except FileNotFoundError:
                            sftp.mkdir(current_dir)

            # 使用文件对象上传数据
            with BytesIO(local_data) as fh:
                sftp.putfo(fh, remote_path)
            return True
        except Exception as e:
            logging.error(f"上传文件时出错: {str(e)}")
            return False

    def check_image_format(self, image_data):
        """检查图片真实格式，返回格式类型"""
        # 检查WebP格式
        if image_data[:4] == b'RIFF' and b'WEBP' in image_data[:12]:
            return 'webp'
        # 检查PNG格式
        elif image_data[:8] == b'\x89PNG\r\n\x1a\n':
            return 'png'
        # 检查JPG格式
        elif image_data[:3] == b'\xff\xd8\xff':
            return 'jpg'
        else:
            return 'unknown'

    def convert_to_jpg(self, image_data):
        """将图片数据转换为JPG格式"""
        try:
            # 使用BytesIO读取图片数据
            with BytesIO(image_data) as input_buffer:
                img = Image.open(input_buffer)
                
                # 转换为RGB模式（如果是RGBA，去除透明通道）
                if img.mode == 'RGBA':
                    # 创建白色背景
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    # 将图片粘贴到白色背景上
                    background.paste(img, mask=img.split()[3])  # 3 是 alpha 通道
                    img = background
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # 保存为JPG格式
                output_buffer = BytesIO()
                img.save(output_buffer, 'JPEG', quality=95)
                output_buffer.seek(0)
                return output_buffer.read()
        except Exception as e:
            logging.error(f"转换图片格式时出错: {str(e)}")
            return None

    def download_and_upload_image(self, url, remote_path, sftp):
        """下载图片并上传到远程服务器"""
        try:
            # 下载图片
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                image_data = response.content
                
                # 检查图片真实格式，无论扩展名是什么
                image_format = self.check_image_format(image_data)
                logging.info(f"检测到图片真实格式: {image_format}")
                
                # 如果是WebP格式，转换为JPG
                if image_format == 'webp':
                    logging.info(f"检测到WebP格式图片，正在转换为JPG格式...")
                    converted_data = self.convert_to_jpg(image_data)
                    if converted_data:
                        image_data = converted_data
                        logging.info("WebP转换为JPG成功")
                    else:
                        logging.warning("WebP转换为JPG失败，将使用原始图片")
                
                # 上传到远程服务器
                return self.upload_to_remote(sftp, image_data, remote_path)
            logging.error(f"下载图片失败，状态码: {response.status_code}")
            return False
        except Exception as e:
            logging.error(f"下载图片时出错: {str(e)}")
            return False

    def extract_images_from_excel(self, file_path, task_status=None):
        """从Excel文件中提取图片URL，下载图片并上传到远程服务器
        
        Args:
            file_path: Excel文件路径
            task_status: 任务状态字典，用于更新进度
            
        Returns:
            dict: 包含处理结果的字典
        """
        try:
            # 读取Excel文件
            logging.info("正在读取Excel文件...")
            df = pd.read_excel(file_path, sheet_name="产品")
            
            # 更新任务状态
            if task_status:
                task_status['message'] = "正在分析Excel文件..."
                task_status['progress'] = 5

            # 获取图片列名
            image_column = None
            for col in df.columns:
                if '图片' in col or 'image' in col.lower():
                    image_column = col
                    break

            if not image_column:
                error_msg = "未找到图片列！"
                logging.error(error_msg)
                if task_status:
                    task_status['status'] = 'error'
                    task_status['message'] = error_msg
                return {'success': False, 'error': error_msg}

            # 获取MSKU列名
            msku_column = None
            for col in df.columns:
                if 'msku' in col.lower():
                    msku_column = col
                    break

            if not msku_column:
                error_msg = "未找到MSKU列！"
                logging.error(error_msg)
                if task_status:
                    task_status['status'] = 'error'
                    task_status['message'] = error_msg
                return {'success': False, 'error': error_msg}

            logging.info(f"找到图片列: {image_column}")
            logging.info(f"找到MSKU列: {msku_column}")
            
            # 更新任务状态
            if task_status:
                task_status['message'] = "正在连接远程服务器..."
                task_status['progress'] = 10

            # 连接远程服务器
            logging.info("正在连接远程服务器...")
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            try:
                ssh.connect(
                    hostname=SSH_CONFIG['ssh_host'],
                    username=SSH_CONFIG['ssh_username'],
                    password=SSH_CONFIG['ssh_password']
                )
                sftp = ssh.open_sftp()

                # 处理每一行
                success_count = 0
                error_count = 0
                total_rows = len(df)
                
                # 更新任务状态
                if task_status:
                    task_status['message'] = f"开始处理，共 {total_rows} 个产品..."
                    task_status['progress'] = 15
                
                for index, row in df.iterrows():
                    # 计算进度
                    progress = 15 + int(85 * (index / total_rows))
                    if task_status:
                        task_status['progress'] = progress
                        task_status['message'] = f"正在处理 {index+1}/{total_rows}..."
                    
                    msku = str(row[msku_column]).strip()
                    image_url = row[image_column]
                    
                    if pd.isna(image_url) or not isinstance(image_url, str):
                        logging.info(f"跳过 {msku}: 没有图片URL")
                        continue

                    try:
                        # 构建远程路径
                        remote_path = f"{self.remote_image_folder}/{msku}.jpg"
                        
                        # 下载并上传图片
                        if self.download_and_upload_image(image_url, remote_path, sftp):
                            success_count += 1
                            logging.info(f"成功上传图片: {msku}")
                        else:
                            error_count += 1
                            logging.error(f"上传失败: {msku}")
                        
                    except Exception as e:
                        error_count += 1
                        logging.error(f"处理 {msku} 时出错: {str(e)}")

                # 更新任务状态
                if task_status:
                    task_status['status'] = 'completed'
                    task_status['progress'] = 100
                    task_status['message'] = "处理完成"
                    task_status['success_count'] = success_count
                    task_status['error_count'] = error_count

                logging.info(f"\n处理完成!")
                logging.info(f"成功上传: {success_count} 张图片")
                logging.info(f"失败: {error_count} 张图片")
                
                return {
                    'success': True, 
                    'success_count': success_count, 
                    'error_count': error_count
                }

            except Exception as e:
                error_msg = f"连接远程服务器时出错: {str(e)}"
                logging.error(error_msg)
                logging.error(traceback.format_exc())
                if task_status:
                    task_status['status'] = 'error'
                    task_status['message'] = error_msg
                return {'success': False, 'error': error_msg}
            finally:
                if 'sftp' in locals():
                    sftp.close()
                if 'ssh' in locals():
                    ssh.close()
                    
        except Exception as e:
            error_msg = f"处理Excel文件时出错: {str(e)}"
            logging.error(error_msg)
            logging.error(traceback.format_exc())
            if task_status:
                task_status['status'] = 'error'
                task_status['message'] = error_msg
            return {'success': False, 'error': error_msg}
