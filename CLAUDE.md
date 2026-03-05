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

## GitHub Actions 构建最佳实践

### 问题与解决方案总结

在构建多平台发布版本时，遇到以下问题并逐步解决：

1. **Windows 构建产物上传失败（中文文件名编码问题）**
   - 原因：`fs.readdirSync('dist')` 在 Windows 环境下读取中文文件名时出现问题
   - 解决：使用 `glob` 库进行文件匹配，使用 `glob.sync('dist/**/*.exe')` 获取完整路径

2. **glob 变量名冲突**
   - 原因：在同一作用域内声明 `const glob = require('glob')` 后，又使用 `glob.sync()` 导致变量名冲突
   - 解决：将导入的模块重命名为 `globPattern` 避免冲突：`const globPattern = require('glob')`

3. **Windows 产物名称不一致**
   - 原因：不同 Python 版本生成的产物名称不同（win7 vs win10+）
   - 解决：使用 glob 模式匹配而非硬编码文件名

### Workflow 关键配置

```yaml
# 安装 glob 库
- name: Install glob for file matching
  run: npm install glob

# 上传时使用完整路径
- name: Upload to Release
  uses: actions/github-script@v7
  with:
    script: |
      const fs = require('fs');
      const path = require('path');
      const globPattern = require('glob');

      // 使用 glob 获取完整路径
      const files = globPattern.sync('dist/**/*.exe');

      for (const file of files) {
        const fileName = path.basename(file);
        await github.rest.repos.uploadReleaseAsset({
          owner: context.repo.owner,
          repo: context.repo.repo,
          release_id: ${{ needs.create-release.outputs.release_id }},
          name: fileName,
          data: fs.readFileSync(file)
        });
      }
```

### 调试技巧

1. **查看构建日志**：使用 `gh run view <run_id> --log-failed` 查看失败日志
2. **本地测试 Workflow**：使用 `workflow_dispatch` 手动触发构建
3. **使用artifact验证**：先通过 `actions/upload-artifact` 上传到 artifact，确认文件存在后再上传到 Release
4. **添加调试输出**：在脚本中添加 `console.log` 输出中间变量值
