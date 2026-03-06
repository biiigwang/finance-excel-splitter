# 甜甜的财务数据拆分工具 (SweetyFinanceExcelSplitter)

## 功能说明

本工具用于将包含多个 sheet 的财务考核 Excel 数据，按科室拆分为多个独立的 Excel 文件。每个科室的文件中包含所有 sheet，但只保留该科室的数据。

### 主要功能

- 自动识别所有 sheet 中的"科室"列
- 支持多种科室列表头（"科室"、"绩效科室"等）
- 自动收集所有科室名称
- 为每个科室生成包含所有 sheet 的独立 Excel 文件
- **保留原 Excel 的格式和样式**（可选择）
- **移除空白子表功能**（可选）
- **三种样式模式**：统一样式、保留原样式、无样式
- 提供两种使用方式：图形化界面 (GUI) 和命令行 (CLI)

---

## 文件说明

| 文件/目录 | 说明 |
|----------|------|
| `gui_app.py` | GUI 主程序（推荐日常使用） |
| `split_all_departments.py` | 命令行版本（适合批处理/自动化使用） |
| `requirements.txt` | Python 依赖包列表 |
| `build_local.py` | 本地构建脚本 |
| `build_macos.py` | macOS 打包脚本 |
| `build_windows.py` | Windows 打包脚本 |
| `core/` | 核心处理模块 |
| `data/` | 输入数据目录（不提交到 git） |
| `output/` | 输出结果目录（不提交到 git） |

---

## 使用方法（运行源码）

### 环境要求

- Python 3.8 或更高版本（推荐 3.9 以兼顾 Windows 7 支持）
- 依赖包：`openpyxl`, `pandas`

### 安装依赖

方式一：使用 requirements.txt（推荐）
```bash
pip install -r requirements.txt
```

方式二：手动安装
```bash
pip install openpyxl pandas
```

### 1. GUI 版本（推荐）

图形化界面，适合日常使用：

```bash
python gui_app.py
```

使用步骤：
1. 点击"浏览..."选择要处理的 Excel 文件
2. 选择输出目录（默认为程序所在目录的 output 文件夹）
3. 可选：勾选"移除空白子表"移除没有数据的 sheet
4. 可选：选择样式模式（统一样式、保留原样式、无样式）
5. 点击"开始处理"

### 2. 命令行版本 (CLI)

适合批处理或自动化脚本：

```bash
# 使用默认路径
python split_all_departments.py

# 指定输入和输出路径
python split_all_departments.py -i /path/to/input.xlsx -o /path/to/output_dir

# 或使用长参数
python split_all_departments.py --input /path/to/input.xlsx --output /path/to/output_dir
```

查看帮助：
```bash
python split_all_departments.py -h
```

---

## 使用方法（使用打包后的程序）

### macOS

1. 双击 `SweetyFinanceExcelSplitter.app`
2. 点击"浏览..."选择要处理的 Excel 文件
3. 选择输出目录（默认为程序所在目录的 output 文件夹）
4. 可选：勾选"移除空白子表"移除没有数据的 sheet
5. 可选：选择样式模式
6. 点击"开始处理"

### Windows

1. 双击 `SweetyFinanceExcelSplitter.exe`
2. 点击"浏览..."选择要处理的 Excel 文件
3. 选择输出目录（默认为程序所在目录的 output 文件夹）
4. 可选：勾选"移除空白子表"移除没有数据的 sheet
5. 可选：选择样式模式
6. 点击"开始处理"

---

## 打包说明

### 前置要求

安装 PyInstaller：

```bash
pip install pyinstaller
```

### 方式一：使用 GitHub Actions（推荐）

无需本地环境，直接在 GitHub 上构建：

1. 推送代码到 GitHub
2. 在 Actions 页面手动触发，或推送 tag（如 `v1.0.0`）
3. 下载构建结果

详细说明请查看 `GITHUB_ACTIONS使用说明.md`

### 方式二：本地构建

#### 打包 macOS 版本

在 macOS 系统上运行：

```bash
python build_macos.py
```

输出：`dist/SweetyFinanceExcelSplitter.app`

#### 打包 Windows 版本

**注意：必须在 Windows 系统（或虚拟机）上运行！

```bash
python build_windows.py
```

输出：`dist/SweetyFinanceExcelSplitter.exe`

### 支持的系统

- macOS 10.14 及更高版本
- Windows 7, Windows 10, Windows 11

---

## 兼容性设计

本工具具有良好的扩展性，支持以下场景无需修改代码：

- ✅ 新增任意 sheet，只要包含"科室"列就能自动处理
- ✅ 在现有 sheet 中新增科室行，自动识别并生成对应文件
- ✅ 科室列表头名称变化（如"科室"→"所属科室"），正则仍能匹配
- ✅ 新增 sheet 不需要在代码中添加配置

---

## 项目结构

```
finance-excel-splitter/
├── gui_app.py                    # GUI 主程序
├── split_all_departments.py      # 命令行版本
├── build_local.py              # 本地构建脚本
├── build_macos.py              # macOS 打包脚本
├── build_windows.py            # Windows 打包脚本
├── README.md                   # 本文档
├── CLAUDE.md                   # Claude Code 项目指导
├── GITHUB_ACTIONS使用说明.md   # GitHub Actions 使用说明
├── core/                      # 核心模块
│   ├── __init__.py
│   ├── sheet_structure.py        # Sheet 结构数据类
│   ├── style_utils.py            # 样式复制工具
│   ├── sheet_analyzer.py         # Sheet 分析器
│   ├── department_collector.py   # 科室收集器
│   ├── sheet_filter.py           # 数据筛选器
│   └── workbook_builder.py       # 工作簿构建器
├── .github/
│   └── workflows/
│       ├── build-all.yml         # 同时构建 Windows + macOS
│       └── build-windows.yml     # 仅构建 Windows
├── data/                       # 输入数据（不提交到 git）
│   └── .gitkeep
└── output/                     # 输出结果（不提交到 git）
    └── .gitkeep
```

---

## 常见问题

### Q: "实发绩效"列显示 #REF! 怎么办？

A: 这是原始 Excel 文件的公式引用问题。请在 Excel 中打开原始文件，确认公式能正常计算后，重新保存再使用本工具处理。

### Q: 可以处理 .xls 格式吗？

A: 目前仅支持 .xlsx 和 .xlsm 格式。如需处理 .xls，请先在 Excel 中另存为 .xlsx 格式。

### Q: 新增的科室能自动识别吗？

A: 可以！工具会自动收集所有 sheet 中的科室，无需修改代码。

### Q: GUI 版本和 CLI 版本有什么区别？

A:
- **GUI 版本**：图形化界面，操作直观，适合日常使用
- **CLI 版本**：命令行界面，适合批处理或自动化脚本

两者使用的是相同的核心处理模块，功能完全一致。

---

## 技术栈

- **GUI 框架**: Tkinter（Python 内置，跨平台）
- **Excel 处理**: openpyxl
- **打包工具**: PyInstaller

---

## 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件。
