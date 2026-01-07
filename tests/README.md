# RA 文档技能 - 测试指南

本目录包含 RA (Regulatory Affairs) 质量文档编写技能的完整单元测试套件。

## 目录结构

```
tests/
├── __init__.py                      # 测试包初始化
├── conftest.py                      # Pytest 配置和共享 fixtures
├── test_extract_quality_standards.py # 提取质量标准表格的测试
├── test_fill_quality_standards.py    # 填充质量标准表格的测试
├── requirements.txt                  # 测试依赖项
├── fixtures/                         # 测试数据目录
└── reports/                          # 测试报告目录
    ├── htmlcov/                      # HTML 覆盖率报告
    └── coverage.xml                  # XML 覆盖率报告
```

## 安装依赖

在运行测试之前，请安装所需的依赖项：

```bash
# 使用 pip 安装
pip install -r tests/requirements.txt

# 或使用项目根目录的安装命令
pip install pytest pytest-cov python-docx
```

## 运行测试

### 基本测试运行

```bash
# 运行所有测试
pytest

# 运行特定测试文件
pytest tests/test_extract_quality_standards.py
pytest tests/test_fill_quality_standards.py

# 运行特定测试类
pytest tests/test_extract_quality_standards.py::TestExtractTextWithFormatting

# 运行特定测试方法
pytest tests/test_extract_quality_standards.py::TestExtractTextWithFormatting::test_extract_plain_text
```

### 详细输出

```bash
# 显示详细输出
pytest -v

# 显示更详细的输出（包括 print 语句）
pytest -vv -s
```

### 覆盖率报告

```bash
# 生成覆盖率报告（终端）
pytest --cov=.claude/skills/ra-doc-assit/scripts --cov-report=term-missing

# 生成 HTML 覆盖率报告
pytest --cov=.claude/skills/ra-doc-assit/scripts --cov-report=html

# 生成 XML 覆盖率报告（用于 CI/CD）
pytest --cov=.claude/skills/ra-doc-assit/scripts --cov-report=xml
```

报告将保存在 `tests/reports/htmlcov/` 目录中。

### 运行特定标记的测试

```bash
# 仅运行单元测试
pytest -m unit

# 仅运行集成测试
pytest -m integration

# 排除慢速测试
pytest -m "not slow"
```

### 并行运行测试（需要安装 pytest-xdist）

```bash
pip install pytest-xdist

# 使用所有可用的 CPU 核心
pytest -n auto

# 使用指定数量的进程
pytest -n 4
```

## 测试内容

### test_extract_quality_standards.py

测试从 Word 文档中提取质量标准表格的功能：

- **TestExtractTextWithFormatting**: 测试从单元格中提取文本并保留上标/下标格式
- **TestFormatAsMarkdownTable**: 测试将表格数据转换为 Markdown 格式
- **TestExtractQualityStandardsTable**: 测试从 Word 文档中提取质量标准表格
- **TestExtractQualityStandardsTableFromDocx**: 测试主入口函数

### test_fill_quality_standards.py

测试将质量标准数据填充到 Word 文档表格的功能：

- **TestParseMarkdownTableFromString**: 测试从字符串解析 Markdown 表格
- **TestParseMarkdownTableFromFile**: 测试从文件解析 Markdown 表格
- **TestRestoreFormattingToCell**: 测试将 Unicode 格式还原为 Word 格式
- **TestClearTableContent**: 测试清除表格内容
- **TestInsertTableRows**: 测试插入表格行
- **TestFindQualityStandardsTable**: 测试查找质量标准表格
- **TestMergeCellsInColumn**: 测试合并单元格
- **TestAutoMergeDuplicateCells**: 测试自动合并重复单元格
- **TestFillWordDocumentTable**: 测试填充 Word 文档表格
- **TestFillQualityStandardsFromMarkdown**: 测试从 Markdown 填充
- **TestFillQualityStandardsFromFile**: 测试从文件填充
- **TestFillQualityStandardsInplace**: 测试就地修改文档

## 测试 Fixture

项目提供了以下共享 fixtures（在 `conftest.py` 中）：

- `test_data_dir`: 测试数据目录路径
- `example_doc_dir`: 示例文档目录路径
- `temp_output_dir`: 临时输出目录
- `sample_markdown_table`: 示例 Markdown 表格
- `sample_markdown_with_formatting`: 带有 Unicode 格式的示例 Markdown

## 持续集成

测试配置包含 CI/CD 系统所需的文件：

### pytest.ini
Pytest 的主要配置文件，包含：
- 测试发现模式
- 覆盖率配置
- 日志设置
- 标记定义

### .coveragerc
覆盖率工具的配置文件，定义：
- 源代码路径
- 排除的文件和行
- 报告格式

## 测试最佳实践

1. **运行测试前确保依赖已安装**: 使用 `pip install -r tests/requirements.txt`
2. **定期运行完整测试套件**: 确保所有功能正常工作
3. **查看覆盖率报告**: 目标是保持 80% 以上的代码覆盖率
4. **添加新功能时同时添加测试**: 保持测试套件的完整性
5. **使用描述性的测试名称**: 使测试失败时容易识别问题

## 常见问题

### Q: 测试失败提示找不到模块
A: 确保 scripts 目录在 Python 路径中。测试文件会自动添加路径，但如果手动运行脚本，需要设置 PYTHONPATH。

### Q: 覆盖率报告不准确
A: 清理 `.pytest_cache` 和 `__pycache__` 目录，然后重新运行测试。

### Q: 测试运行缓慢
A: 使用 `pytest -n auto` 并行运行测试，或使用 `-m "not slow"` 排除慢速测试。

### Q: 想要使用示例文档测试
A: 示例文档位于 `doc/example/` 目录，可以使用 `example_doc_dir` fixture 访问。

## 贡献指南

当添加新功能或修复错误时：

1. 为新功能添加测试
2. 确保所有现有测试通过
3. 保持或提高代码覆盖率
4. 更新相关文档

## 联系方式

如有问题或建议，请在项目仓库中提交 Issue。
