#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Web应用服务器
整合配置编辑和Excel处理功能
"""

import os
import sys
import json
import webbrowser
import http.server
import socketserver
import threading
import tempfile
import shutil
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote
from io import BytesIO
import email
from email.parser import BytesParser

from main import ExcelPriceUpdater


class WebAppHandler(http.server.SimpleHTTPRequestHandler):
    """Web应用HTTP处理器"""
    
    config_file = Path("config.json")
    temp_dir = Path(tempfile.gettempdir()) / "excel_updater"
    
    def __init__(self, *args, **kwargs):
        # 确保临时目录存在
        self.temp_dir.mkdir(exist_ok=True)
        super().__init__(*args, directory=os.getcwd(), **kwargs)
    
    def do_GET(self):
        """处理GET请求"""
        parsed_path = urlparse(self.path)
        
        # API: 获取配置
        if parsed_path.path == '/api/config':
            self.handle_get_config()
            return
        
        # API: 获取可用地域
        if parsed_path.path == '/api/regions':
            self.handle_get_regions()
            return
        
        # API: 下载处理后的文件
        if parsed_path.path.startswith('/api/download/'):
            filename = parsed_path.path.replace('/api/download/', '')
            self.handle_download_file(unquote(filename))
            return
        
        # 默认首页
        if parsed_path.path == '/':
            self.path = '/index.html'
        
        # 静态文件
        super().do_GET()
    
    def do_POST(self):
        """处理POST请求"""
        parsed_path = urlparse(self.path)
        
        # API: 保存配置
        if parsed_path.path == '/api/config':
            self.handle_save_config()
            return
        
        # API: 处理Excel文件
        if parsed_path.path == '/api/process':
            self.handle_process_excel()
            return
        
        self.send_error(404, "Not Found")
    
    def handle_get_config(self):
        """获取配置"""
        self.send_response(200)
        self.send_header('Content-type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.wfile.write(json.dumps(config, ensure_ascii=False).encode('utf-8'))
            else:
                self.wfile.write(b'{}')
        except Exception as e:
            self.wfile.write(json.dumps({'error': str(e)}, ensure_ascii=False).encode('utf-8'))
    
    def handle_save_config(self):
        """保存配置"""
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        
        try:
            config = json.loads(post_data.decode('utf-8'))
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({
                'success': True,
                'message': '配置已保存'
            }, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({
                'success': False,
                'error': str(e)
            }, ensure_ascii=False).encode('utf-8'))
    
    def handle_get_regions(self):
        """获取可用地域列表"""
        try:
            updater = ExcelPriceUpdater()
            regions = list(updater.price_columns.keys())
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({
                'regions': regions
            }, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({
                'error': str(e)
            }, ensure_ascii=False).encode('utf-8'))
    
    def handle_process_excel(self):
        """处理Excel文件"""
        try:
            # 获取content-type和boundary
            content_type = self.headers.get('content-type', '')
            if not content_type.startswith('multipart/form-data'):
                raise ValueError('Invalid content type')
            
            # 提取boundary
            boundary = content_type.split('boundary=')[1].strip()
            
            # 读取POST数据
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            
            # 解析multipart数据
            parts = post_data.split(('--' + boundary).encode())
            
            file_data = None
            filename = None
            regions = None
            
            for part in parts:
                if b'Content-Disposition' in part:
                    # 解析disposition头
                    lines = part.split(b'\r\n')
                    for i, line in enumerate(lines):
                        if b'Content-Disposition' in line:
                            disposition = line.decode('utf-8')
                            
                            # 提取文件
                            if 'filename=' in disposition:
                                filename = disposition.split('filename=')[1].strip('"')
                                # 文件内容在空行之后
                                content_start = part.find(b'\r\n\r\n') + 4
                                content_end = len(part) - 2  # 去掉结尾的\r\n
                                file_data = part[content_start:content_end]
                            
                            # 提取地域信息
                            elif 'name="regions"' in disposition:
                                content_start = part.find(b'\r\n\r\n') + 4
                                content_end = len(part) - 2
                                regions_str = part[content_start:content_end].decode('utf-8')
                                regions = json.loads(regions_str)
            
            if not file_data or not filename or not regions:
                raise ValueError('Missing file or regions data')
            
            # 保存上传的文件
            temp_input = self.temp_dir / filename
            with open(temp_input, 'wb') as f:
                f.write(file_data)
            
            # 处理文件
            print(f"Processing file: {temp_input}")
            print(f"Regions: {regions}")
            
            updater = ExcelPriceUpdater()
            success = updater.update_prices(
                str(temp_input),
                regions,
                output_suffix='_updated'
            )
            
            output_file = temp_input.parent / f"{temp_input.stem}_updated{temp_input.suffix}"
            
            if success and output_file.exists():
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'success': True,
                    'output_file': output_file.name,
                    'updated_count': 0  # TODO: 从updater获取实际更新数量
                }, ensure_ascii=False).encode('utf-8'))
                
                # 删除输入文件
                temp_input.unlink()
            else:
                raise Exception('Processing failed')
                
        except Exception as e:
            print(f"Error processing Excel: {e}")
            import traceback
            traceback.print_exc()
            
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({
                'success': False,
                'error': str(e)
            }, ensure_ascii=False).encode('utf-8'))
    
    def handle_download_file(self, filename):
        """下载文件"""
        file_path = self.temp_dir / filename
        
        if not file_path.exists():
            self.send_error(404, "File not found")
            return
        
        try:
            self.send_response(200)
            self.send_header('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.end_headers()
            
            with open(file_path, 'rb') as f:
                self.wfile.write(f.read())
            
            # 删除临时文件
            file_path.unlink()
        except Exception as e:
            print(f"Error downloading file: {e}")
            self.send_error(500, "Internal Server Error")
    
    def log_message(self, format, *args):
        """自定义日志格式"""
        return  # 静默模式


def start_server(port=8800):
    """启动Web服务器"""
    try:
        with socketserver.TCPServer(("", port), WebAppHandler) as httpd:
            print("="*60)
            print("Excel价格批量更新系统已启动")
            print("="*60)
            print(f"\n🌐 访问地址: http://localhost:{port}")
            print(f"\n功能：")
            print(f"  📈 价格更新 - 批量处理Excel文件")
            print(f"  ⚙️  配置管理 - 可视化编辑价格配置")
            print(f"\n按 Ctrl+C 停止服务器\n")
            print("="*60)
            
            # 在新线程中打开浏览器
            def open_browser():
                import time
                time.sleep(1)
                webbrowser.open(f'http://localhost:{port}')
            
            threading.Thread(target=open_browser, daemon=True).start()
            
            # 启动服务器
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n\n服务器已停止")
    except OSError as e:
        if e.errno == 48:  # Address already in use
            print(f"\n错误：端口 {port} 已被占用")
            print("请尝试：")
            print(f"  1. 关闭占用端口 {port} 的程序")
            print(f"  2. 或者使用其他端口")
        else:
            print(f"\n错误：{e}")


if __name__ == '__main__':
    start_server()

