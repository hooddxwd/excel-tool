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
    print("请确保已安装所需依赖：pip install -r requirements.txt")
except Exception as e:
    print(f"启动失败：{e}")
    print("请检查Python环境和依赖包是否正确安装")