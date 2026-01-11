# CLAUDE.md

本文件为 Claude Code (claude.ai/code) 在此代码库中工作时提供指导。

## 项目概述

本代码库包含一个用于药品注册质量文档编写的 Claude Code 技能（RA质量文档编写）。该技能协助完成：
- 从 Word 文档中提取质量标准表格
- 将 markdown 格式的质量标准数据填充到 Word 文档表格中
- 智能提取和总结检验方法文档，生成符合注册要求的标准格式
- 自动校验文档间的一致性

## 目录结构

```
ra-agent-skills/
├── .claude/
│   ├── settings.local.json          # Claude Code 权限配置
│   └── skills/
│       └── ra-doc-assit/            # 主技能目录
│           ├── SKILL.md             # 技能定义和说明
│           └── scripts/
│               ├── extract_quality_standards.py    # 从 Word 文档提取表格
│               └── fill_quality_standards.py       # 填充 Word 文档表格
├── doc/
│   └── example/                     # 示例 Word 文档
├── docs/                            # 文档目录
│   └── WINDOWS_INSTALLATION.md      # Windows 安装详细指南
├── tests/                           # 测试套件
│   ├── test_extract_quality_standards.py  # 提取功能测试
│   ├── test_fill_quality_standards.py     # 填充功能测试
│   ├── conftest.py                  # Pytest 配置
│   ├── requirements.txt             # 测试依赖
│   ├── README.md                    # 测试文档
│   ├── run_tests.sh                 # 测试运行脚本
│   └── verify_tests.py              # 环境验证脚本
├── pytest.ini                       # Pytest 配置文件
├── .coveragerc                      # 覆盖率配置
├── README-WINDOWS.md                # Windows 快速安装指南
├── setup.bat                        # Windows 自动安装脚本
├── CLAUDE.md                        # 本文件
└── .github/workflows/test.yml       # CI/CD 配置
```

## 依赖项

Python 脚本需要 `python-docx` 库来操作 Word 文档。使用以下命令安装：

```bash
pip install python-docx
```

## Windows 环境安装

🪟 Windows 用户请参考详细的安装指南：[**README-WINDOWS.md**](README-WINDOWS.md)

**快速开始**（3 步）：

```powershell
# 1. 运行自动安装脚本
右键点击 setup.bat -> "以管理员身份运行"

# 2. 配置 API 密钥
setx ANTHROPIC_API_KEY "sk-ant-你的API密钥"

# 3. 重启终端并使用
claude
```

**详细指南包含**：
- ✅ 完整的安装步骤（自动/手动两种方式）
- ✅ Claude Code 和 API 配置详解
- ✅ Git 和 Python 环境配置
- ✅ 常见问题解决方案
- ✅ 更新和维护指南
- ✅ 性能优化和安全建议

## 核心功能

### 提取质量标准表格

`extract_quality_standards.py` 脚本从 Word 文档中提取质量标准表格，具有以下关键特性：

- **自动表格检测**：搜索"4.3 检验项目、方法和标准"章节，识别包含质量关键词的表格
- **格式保留**：将 Word 的上标/下标格式转换为 Unicode 字符（⁰¹²³, ₀₁₂₃）用于文本表示
- **Markdown 输出**：以 markdown 格式返回提取的表格

主要函数：
- `extract_quality_standards_table_from_docx(docx_path: str)`：主提取函数
- `extract_text_with_formatting(cell)`：在提取时保留上标/下标格式
- `format_as_markdown_table(table_data, target_columns)`：转换为 markdown 格式

### 填充 Word 文档表格

`fill_quality_standards.py` 脚本使用 markdown 数据填充 Word 文档表格：

- **智能表格检测**：通过关键词匹配自动查找质量标准表格
- **单元格合并**：自动合并'类型'和'检验项目'列中的重复单元格
- **格式还原**：将 Unicode 上标/下标转换回 Word 格式
- **灵活输入**：接受来自字符串或文件的 markdown

主要函数：
- `fill_quality_standards_from_markdown(doc_path, output_path, markdown_content, table_index, auto_merge)`：从 markdown 字符串填充的主函数
- `fill_quality_standards_from_file(doc_path, output_path, markdown_file_path, table_index, auto_merge)`：从 markdown 文件填充
- `fill_quality_standards_inplace(doc_path, markdown_content, table_index, auto_merge)`：就地修改文档
- `auto_merge_duplicate_cells(table, target_columns)`：自动合并包含重复内容的单元格

### 列结构

质量标准表格遵循以下列结构：
- `类型`
- `检验项目`
- `检验方法`
- `质量标准`

### 总结检验方法内容

通过智能分析标准操作规程（SOP）文档，自动提取并归纳检验方法的核心内容，生成符合药品注册要求的标准格式总结。

## 代码使用指南

### 添加日志记录

两个脚本都使用 Python 的 logging 模块。配置日志级别以进行调试：

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

### 使用示例文档测试

`doc/example/` 目录中提供了示例 Word 文档：
- `例子-原液质量标准.docx` - 原液质量标准示例
- `例子-检验标准操作规程.docx` - 检验标准操作规程示例（可用于方法总结）
- `模板-质量标准.docx` - 质量标准模板
- `模板-分析方法.docx` - 分析方法模板

### 处理上标/下标

代码通过 Unicode 字符映射处理化学公式和科学记数法：

**提取**（Word → 文本）：
- Word 上标 → Unicode：⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ
- Word 下标 → Unicode：₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑᵢₒᵤₓₕₖₗₘₙₚₛₜ

**填充**（文本 → Word）：
- Unicode 或 ^/_ 符号 → Word 格式
- 示例：`H₂O` → H 带下标 2，O

### 错误处理

两个脚本都包含全面的错误处理：
- 文件存在性验证
- 格式化失败时的优雅回退
- 详细的日志和追踪信息
- 针对常见问题的清晰错误消息

## 技能调用

当用户请求以下帮助时，此技能会自动被调用：
- 从 Word 文档提取质量标准表格
- 将质量标准数据填充到 Word 模板中
- 总结检验方法内容（从 SOP 文档生成标准格式的方法概要）
- 校验药品注册文档之间的一致性
- 处理中文药品质量文档

技能名称为 `ra-doc-assit`，可通过 Claude Code 的技能系统调用。

## 测试

项目包含完整的单元测试套件，位于 `tests/` 目录。

### 快速开始

```bash
# 安装测试依赖
pip install -r tests/requirements.txt

# 验证测试环境
python tests/verify_tests.py

# 运行所有测试
pytest

# 或使用测试脚本
./tests/run_tests.sh

# 生成覆盖率报告
./tests/run_tests.sh coverage
```

### 测试文档

详细的测试文档和使用指南请参阅：
- [tests/README.md](tests/README.md) - 测试套件指南
- [tests/TEST_SUMMARY.md](tests/TEST_SUMMARY.md) - 测试总结
- [tests/SAMPLE_TEST_USAGE.md](tests/SAMPLE_TEST_USAGE.md) - 使用示例

### 测试覆盖

- **提取功能测试** (test_extract_quality_standards.py)
  - 文本提取（含上标/下标格式）
  - Markdown 表格格式化
  - Word 文档表格提取
  - 错误处理

- **填充功能测试** (test_fill_quality_standards.py)
  - Markdown 表格解析
  - Word 格式还原
  - 表格操作（清除、插入、合并）
  - 文档填充（多种方式）
  - 自动合并单元格
