#!/bin/bash
# Excel工具软件安装脚本（使用本地虚拟环境，含回退方案）
set -e

echo "=== Excel工具软件安装脚本 ==="
echo "正在检查Python环境..."

# 检查Python3
if command -v python3 &> /dev/null; then
    PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
    echo "检测到Python版本: $PYTHON_VERSION"
else
    echo "错误：未检测到Python3，请先安装Python3"
    exit 1
fi

USE_VENV=0

# 尝试创建虚拟环境
if [ ! -d ".venv" ]; then
  echo "正在创建虚拟环境 .venv ..."
  if python3 -m venv .venv 2>/dev/null; then
    USE_VENV=1
  else
    echo "提示：未能创建虚拟环境，可能缺少 python3-venv。将使用 --user 模式安装依赖。"
  fi
else
  USE_VENV=1
fi

if [ "$USE_VENV" -eq 1 ]; then
  # shellcheck disable=SC1091
  source .venv/bin/activate || USE_VENV=0
fi

# 安装依赖
if [ "$USE_VENV" -eq 1 ]; then
  echo "已启用虚拟环境：.venv"
  python -m pip install -U pip
  pip install -r requirements.txt
else
  echo "使用 --user 模式安装依赖到当前用户目录 (~/.local)"
  python3 -m pip install -U --user pip
  python3 -m pip install --user -r requirements.txt
fi

echo "正在创建示例数据..."
python3 create_sample_data.py || true

echo "安装完成！"
echo ""
echo "使用方法："
if [ "$USE_VENV" -eq 1 ]; then
  echo "1) 激活环境：source .venv/bin/activate"
  echo "2) 启动软件：python run.py  或  python main.py"
  echo "   退出环境：deactivate"
else
  echo "1) （已使用 --user 安装）确保 ~/.local/bin 在 PATH 中"
  echo "2) 启动软件：python3 run.py  或  python3 main.py"
fi

echo ""
echo "示例文件（如依赖安装成功）会创建在项目根目录："
echo "- sample_file_a.xlsx"
echo "- sample_file_b.xlsx"
echo ""
echo "注意：若图形界面无法启动，可能是缺少 Tk（tkinter）。请在系统中安装 python3-tk 后重试。"