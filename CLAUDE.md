# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个使用 Python 脚本来分发财务考核数据的项目,你需要协助用户来编写 Python 脚本，实现分发数据的目的。

## Python 程序开发

### 环境设置
使用 conda 管理虚拟环境，环境名称：`tiantian_excel_test`（与目录名相同）。

当某个处理流程需要重复使用时：

1. 先用 Claude Code 完成手工处理，确认逻辑正确
2. 将处理逻辑改写为 Python 脚本
3. 使用 `openpyxl` 或 `pandas` 处理 Excel 文件

常用命令：

```bash
# 创建 conda 虚拟环境
conda create -n tiantian_excel_test python=3.11

# 激活环境
conda activate tiantian_excel_test

# 安装依赖
conda install -c conda-forge openpyxl pandas
# 或使用 pip
pip install openpyxl pandas

# 运行 Python 脚本
python script.py

# 导出依赖列表
pip freeze > requirements.txt
```
### 代码规范

- 所有 Python 脚本均需符合 PEP 8 规范
- 函数和变量命名应使用下划线命名法（snake_case）
- 类名应使用驼峰命名法（CamelCase）
- 所有代码均需添加必要的注释，包括函数、类、模块等

### 设计模式
在设计时，请加载以下skill,要求每一个逻辑代码都尽量可以复用：
- /python-design-patterns 
- /solid 

## 数据说明
- 原始数据为：data/试验.xlsx，数据中包含多个 sheet，每个 sheet 里面是不同的财务内容：
1. 有的是考核内容
2. 有的是各种费用支出
3. 但是每个表里面都包含一个列包含“科室”字样，这个列是用来标识数据所属的科室的。

## 重要提示

- 处理财务数据时需特别注意数据准确性，多次验证
- 敏感财务数据不要提交到版本控制
- 使用 xlsx skill 时，明确说明需要的处理逻辑

## 技术注意事项

### Excel 公式处理

当使用 `openpyxl` 读取 Excel 文件时，如果单元格包含**公式**而非具体数值，需要注意：

1. **默认读取方式**：`load_workbook('file.xlsx')` 会读取公式本身（如 `=SUMIF(...)`），而不是计算结果
2. **读取计算值**：使用 `load_workbook('file.xlsx', data_only=True)` 可以读取公式计算后的实际数值
3. **使用场景**：复制数据到新文件时，通常需要读取计算值而非公式，因为目标文件可能缺少原公式引用的其他 sheet

**示例代码**：
```python
from openpyxl import load_workbook

# 读取计算后的值（推荐用于数据复制场景）
wb = load_workbook('data.xlsx', data_only=True)
ws = wb['Sheet1']

# 获取单元格的实际数值
value = ws['A1'].value
```
