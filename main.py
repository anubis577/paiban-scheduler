#!/usr/bin/env python3
"""启动排班管理程序"""
import sys
import os

# 获取程序所在目录（兼容打包后运行）
if getattr(sys, 'frozen', False):
    # 打包后的可执行文件
    app_dir = os.path.dirname(sys.executable)
else:
    # 开发环境：脚本所在目录
    app_dir = os.path.dirname(os.path.abspath(__file__))

os.chdir(app_dir)
sys.path.insert(0, app_dir)

from main_UI import main

if __name__ == "__main__":
    main()