#!/usr/bin/env python3

import sys
import subprocess
import os

def main():
    print("=== Markdown转Word Web转换器 ===")
    print()
    
    # 检查依赖
    print("正在检查依赖...")
    try:
        import flask
        import markdown2
        import docx
        print("✓ 依赖已安装")
    except ImportError:
        print("✗ 缺少依赖，正在安装...")
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✓ 依赖安装完成")
    
    print()
    print("正在启动Web服务器...")
    print("访问地址: http://localhost:5000")
    print("按 Ctrl+C 停止服务器")
    print()
    
    # 启动Flask应用
    os.system(f"{sys.executable} web_app.py")

if __name__ == "__main__":
    main()