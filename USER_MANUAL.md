# Excel工具软件用户手册

## 软件概述

Excel工具软件是一款专为局域网环境设计的Excel数据处理工具，无需连接外网即可运行。软件提供了多种实用的Excel处理功能，帮助用户轻松完成各类数据处理任务。

## 系统要求

- 操作系统：Windows 7/8/10/11, Linux, macOS
- Python 3.7 或更高版本
- 内存：至少 2GB RAM
- 磁盘空间：至少 100MB 可用空间

## 安装说明

### 方法一：使用安装脚本（推荐）

1. 下载软件包到本地
2. 打开终端或命令提示符
3. 进入软件目录
4. 运行安装脚本：
   ```bash
   chmod +x install.sh
   ./install.sh
   ```

### 方法二：手动安装

1. 确保已安装Python 3.7+
2. 安装依赖包：
   ```bash
   pip install -r requirements.txt
   ```
3. 创建示例数据：
   ```bash
   python create_sample_data.py
   ```

## 启动软件

运行以下命令启动软件：
```bash
python run.py
```

或者直接运行：
```bash
python main.py
```

## 功能说明

### 1. 列对比功能

**功能描述：** 对比两个Excel文件中指定列的数据，找出差异并标记。

**操作步骤：**
1. 点击"列对比功能"标签页
2. 选择文件A和文件B
3. 点击"加载文件"按钮
4. 在下拉菜单中选择要对比的列
5. 点击"开始对比"按钮
6. 查看对比结果
7. 点击"保存结果"保存处理后的文件

**结果说明：**
- A文件中独有的数据会在原位置标红显示
- B文件中独有的数据会添加到A文件末尾并标红显示
- 保存的Excel文件中，差异数据会以红色背景标记

### 2. 数据去重功能

**功能描述：** 对Excel文件中的指定列进行去重处理，删除重复数据。

**操作步骤：**
1. 点击"数据去重"标签页
2. 选择要去重的Excel文件
3. 点击"加载文件"按钮
4. 选择要去重的列
5. 选择保留策略（保留第一条或最后一条）
6. 点击"开始去重"按钮
7. 查看去重结果
8. 点击"保存结果"保存去重后的文件

**结果说明：**
- 显示原始数据量、去重后数据量、删除的重复数据量
- 计算去重率
- 保存去重后的Excel文件

### 3. 格式调整功能

**功能描述：** 对Excel文件中的数据进行格式统一调整。

**操作步骤：**
1. 点击"格式调整"标签页
2. 选择要调整格式的Excel文件
3. 点击"加载文件"按钮
4. 选择目标列
5. 选择格式类型：
   - 日期格式：统一为YYYY-MM-DD格式
   - 数字格式：保留两位小数
   - 文本格式：去除首尾空格
6. 点击"应用格式"按钮
7. 查看调整结果
8. 点击"保存结果"保存调整后的文件

## 使用技巧

### 文件格式支持
- 支持.xlsx和.xls格式的Excel文件
- 建议文件大小不超过50MB以保证处理效率

### 数据处理建议
- 对比功能建议选择唯一性较高的列（如ID、姓名等）
- 去重功能可根据业务需求选择保留策略
- 格式调整功能会修改原始数据，建议先备份

### 性能优化
- 大文件处理时请耐心等待
- 处理过程中不要关闭软件
- 如遇异常中断，可重新启动软件

## 常见问题

### Q: 软件无法启动怎么办？
A: 请检查Python环境是否正确安装，依赖包是否完整安装。

### Q: 文件加载失败怎么办？
A: 请检查文件格式是否正确，文件是否损坏，文件路径是否包含特殊字符。

### Q: 对比结果不准确怎么办？
A: 请检查选择的对比列是否正确，数据格式是否一致。

### Q: 处理大文件时软件卡死怎么办？
A: 请等待处理完成，或尝试处理较小的文件。建议单个文件不超过50MB。

### Q: 保存文件失败怎么办？
A: 请检查保存路径是否有写入权限，磁盘空间是否充足。

## 技术支持

如遇到问题，请检查：
1. Python版本是否符合要求
2. 依赖包是否正确安装
3. 文件格式是否正确
4. 系统权限是否足够

## 更新日志

### v1.0.0 (当前版本)
- 实现列对比功能
- 实现数据去重功能
- 实现格式调整功能
- 支持Excel文件读写
- 提供图形化用户界面

## 免责声明

本软件仅供学习和办公使用，请勿用于商业用途。使用本软件处理的数据请自行备份，开发者不承担因使用本软件造成的数据丢失责任。