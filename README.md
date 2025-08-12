# excel-tool
局域网 Excel 便捷工具软件

一、项目概述
本项目专注于局域网环境，打造一款操作简便的 Excel 工具软件。专为不熟悉 Excel 复杂命令的用户设计，无需连接外网即可运行，帮助用户轻松完成各类 Excel 数据处理任务。

二、当前版本功能（v1.0.0）
- 核心功能：列对比（对比两个 Excel 文件指定列的差异，差异项标红并可保存）
- 数据去重（按列去重，支持保留第一条或最后一条，展示去重指标并可保存）
- 格式调整（日期统一为 YYYY-MM-DD、数字保留两位小数、文本去除首尾空格）

三、扩展功能（计划中）
- 数据拆分与合并：按列值拆分为多个文件，或将多个同结构文件按列合并
- 快速报表生成：基于模板快速生成销售/库存/绩效等可编辑报表
- 常用格式批量模板：更多预设格式模板与预览

四、安装与启动
方式一：使用安装脚本（推荐）
```bash
chmod +x install.sh
./install.sh
```
脚本将自动创建本地虚拟环境 `.venv`，安装依赖并生成示例数据。

方式二：手动安装
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python create_sample_data.py
```

启动软件
```bash
# 已激活 .venv 后：
python run.py     # 或
python main.py
```

五、基本操作示例（以列对比为例）
1. 打开“列对比功能”标签页
2. 选择文件 A 与文件 B，点击“加载文件”
3. 分别选择要对比的列，点击“开始对比”
4. 查看结果并可“保存结果”

六、注意事项
- 支持 .xls 与 .xlsx，建议单个文件不超过 50MB
- 处理过程中若异常中断，已处理数据不会自动保存，请及时保存结果
- 图形界面依赖系统的 Tk（tkinter）。如在无图形环境（例如服务器）运行，建议在带桌面的系统上启动

七、示例数据
安装脚本或命令会在项目根目录生成：
- `sample_file_a.xlsx`
- `sample_file_b.xlsx`

八、文档与支持
- 详细使用说明见 `USER_MANUAL.md`
- 依赖列表见 `requirements.txt`
- 无图形环境可使用命令行模式，参见下文
 
 九、命令行模式（无图形环境）
- 列对比：
```bash
python cli.py compare --file-a sample_file_a.xlsx --col-a 姓名 \
                     --file-b sample_file_b.xlsx --col-b 姓名 \
                     --output compare_result.xlsx
```
- 数据去重：
```bash
python cli.py dedup --file sample_file_a.xlsx --column 姓名 --keep first \
                    --output dedup_result.xlsx
```
- 格式调整：
```bash
python cli.py format --file sample_file_a.xlsx --column 入职日期 --type date \
                     --output format_result.xlsx
```

十、开发计划（概要）
- 第1阶段：完成列对比功能（已完成）
- 第2阶段：实现数据去重、格式调整（已完成）
- 第3阶段：实现快速报表、数据拆分与合并（进行中/计划中）
- 第4阶段：整体测试与界面优化
