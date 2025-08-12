#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具软件启动脚本
"""

import sys
import os

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from main import ExcelTool
    print("正在启动Excel工具软件...")
    app = ExcelTool()
    app.run()
except ImportError as e:
    print(f"导入错误：{e}")
    print("提示：若缺少 tkinter（图形界面库），可安装系统包 python3-tk，或使用命令行模式：python cli.py --help")
    print("请确保已安装所需依赖：pip install -r requirements.txt")
except Exception as e:
    print(f"启动失败：{e}")
    print("若在无图形环境，请改用命令行模式：python cli.py --help")
    print("请检查Python环境和依赖包是否正确安装")