#!/bin/bash
# Excel工具软件安装脚本

echo "=== Excel工具软件安装脚本 ==="
echo "正在检查Python环境..."

# 检查Python版本
if command -v python3 &> /dev/null; then
    PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
    echo "检测到Python版本: $PYTHON_VERSION"
else
    echo "错误：未检测到Python3，请先安装Python3"
    exit 1
fi

# 检查pip
if command -v pip3 &> /dev/null; then
    echo "检测到pip3"
else
    echo "错误：未检测到pip3，请先安装pip3"
    exit 1
fi

echo "正在安装依赖包..."
pip3 install -r requirements.txt

if [ $? -eq 0 ]; then
    echo "依赖包安装成功！"
else
    echo "依赖包安装失败，请检查网络连接或手动安装"
    exit 1
fi

echo "正在创建示例数据..."
python3 create_sample_data.py

echo "安装完成！"
echo ""
echo "使用方法："
echo "1. 运行软件：python3 run.py"
echo "2. 或者直接运行：python3 main.py"
echo ""
echo "示例文件已创建："
echo "- sample_file_a.xlsx"
echo "- sample_file_b.xlsx"
echo ""
echo "请参考README.md了解详细使用说明"