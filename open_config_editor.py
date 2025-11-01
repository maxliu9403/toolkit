#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打开配置编辑器的启动脚本
"""

import os
import sys
import json
import webbrowser
import http.server
import socketserver
import threading
from pathlib import Path
from urllib.parse import urlparse, parse_qs


def load_config():
    """加载配置文件"""
    config_file = Path("config.json")
    if not config_file.exists():
        print("创建默认配置文件...")
        default_config = {
            "Nike Air force 1": {
                "hk": [550, 580, 10],
                "sg": [70, 85, 5],
                "my": [50, 60, 10]
            },
            "New Balance NB 327": {
                "hk": [480, 510, 10],
                "sg": [75, 90, 5],
                "my": [60, 70, 10]
            }
        }
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=2, ensure_ascii=False)
        print(f"✓ 已创建 {config_file}")
    
    return config_file


class ConfigHandler(http.server.SimpleHTTPRequestHandler):
    """自定义HTTP处理器"""
    
    config_file = Path("config.json")
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=os.getcwd(), **kwargs)
    
    def do_GET(self):
        """处理GET请求"""
        parsed_path = urlparse(self.path)
        
        # API端点：获取配置
        if parsed_path.path == '/api/config':
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
            return
        
        # 静态文件
        super().do_GET()
    
    def do_POST(self):
        """处理POST请求"""
        parsed_path = urlparse(self.path)
        
        # API端点：保存配置
        if parsed_path.path == '/api/config':
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
                self.wfile.write(json.dumps({'success': True, 'message': '配置已保存'}, ensure_ascii=False).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))
            return
        
        self.send_response(404)
        self.end_headers()
    
    def do_OPTIONS(self):
        """处理OPTIONS请求（CORS预检）"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
    
    def log_message(self, format, *args):
        """禁用日志输出"""
        pass


def start_server(port=8800):
    """启动本地服务器"""
    handler = ConfigHandler
    httpd = socketserver.TCPServer(("", port), handler)
    print(f"✓ 本地服务器已启动: http://localhost:{port}")
    
    # 在新线程中运行服务器
    server_thread = threading.Thread(target=httpd.serve_forever)
    server_thread.daemon = True
    server_thread.start()
    
    return httpd


def main():
    """主函数"""
    print("=" * 60)
    print("Excel价格配置编辑器启动器")
    print("=" * 60)
    print()
    
    # 检查HTML文件是否存在
    html_file = Path("config_editor.html")
    if not html_file.exists():
        print(f"✗ 错误: 找不到 {html_file}")
        print("   请确保 config_editor.html 文件存在")
        return 1
    
    # 加载或创建配置文件
    config_file = load_config()
    print(f"✓ 配置文件: {config_file}")
    
    # 启动本地服务器
    try:
        httpd = start_server(8800)
        port = 8800
    except OSError:
        print("✗ 端口 8800 被占用，尝试使用其他端口...")
        try:
            httpd = start_server(8801)
            port = 8801
        except OSError:
            print("✗ 端口 8801 也被占用，请关闭其他应用后重试")
            return 1
    
    # 打开浏览器
    url = f"http://localhost:{port}/config_editor.html"
    
    print(f"✓ 正在打开浏览器...")
    print()
    print("提示: 按 Ctrl+C 停止服务器")
    print("=" * 60)
    
    webbrowser.open(url)
    
    try:
        # 保持服务器运行
        while True:
            import time
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n停止服务器...")
        httpd.shutdown()
        print("✓ 服务器已停止")
        return 0


if __name__ == "__main__":
    sys.exit(main())
