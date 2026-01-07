# 测试使用示例

本文档提供了使用测试套件的实际示例。

## 目录

1. [基本测试运行](#基本测试运行)
2. [测试特定功能](#测试特定功能)
3. [生成覆盖率报告](#生成覆盖率报告)
4. [使用示例文档测试](#使用示例文档测试)
5. [调试测试](#调试测试)

## 基本测试运行

### 运行所有测试

```bash
# 使用 pytest
pytest

# 或使用提供的脚本
./tests/run_tests.sh

# 或
./tests/run_tests.sh all
```

### 运行特定测试文件

```bash
# 测试提取功能
pytest tests/test_extract_quality_standards.py

# 测试填充功能
pytest tests/test_fill_quality_standards.py
```

## 测试特定功能

### 测试文本提取功能

```bash
# 测试纯文本提取
pytest tests/test_extract_quality_standards.py::TestExtractTextWithFormatting::test_extract_plain_text -v

# 测试上标提取
pytest tests/test_extract_quality_standards.py::TestExtractTextWithFormatting::test_extract_superscript_numbers -v

# 测试下标提取
pytest tests/test_extract_quality_standards.py::TestExtractTextWithFormatting::test_extract_subscript_numbers -v
```

### 测试 Markdown 表格解析

```bash
# 测试简单表格解析
pytest tests/test_fill_quality_standards.py::TestParseMarkdownTableFromString::test_parse_simple_table -v

# 测试多行表格
pytest tests/test_fill_quality_standards.py::TestParseMarkdownTableFromString::test_parse_with_multiple_rows -v
```

### 测试格式化功能

```bash
# 测试 Unicode 上标还原
pytest tests/test_fill_quality_standards.py::TestRestoreFormattingToCell::test_restore_superscript_unicode -v

# 测试 Unicode 下标还原
pytest tests/test_fill_quality_standards.py::TestRestoreFormattingToCell::test_restore_subscript_unicode -v
```

## 生成覆盖率报告

### 生成终端覆盖率报告

```bash
pytest --cov=.claude/skills/ra-doc-assit/scripts --cov-report=term-missing
```

输出示例：
```
Name                                                           Stmts   Miss  Cover   Missing
-------------------------------------------------------------------------------------------
.claude/skills/ra-doc-assit/scripts/extract_quality_standards.py      120     15    88%   23-27, 45-49
.claude/skills/ra-doc-assit/scripts/fill_quality_standards.py         180     20    89%   56-60, 78-82
-------------------------------------------------------------------------------------------
TOTAL                                                             300     35    88%
```

### 生成 HTML 覆盖率报告

```bash
# 使用 pytest
pytest --cov=.claude/skills/ra-doc-assit/scripts --cov-report=html

# 或使用提供的脚本
./tests/run_tests.sh coverage

# 在浏览器中打开报告
open tests/reports/htmlcov/index.html  # macOS
xdg-open tests/reports/htmlcov/index.html  # Linux
start tests/reports/htmlcov/index.html  # Windows
```

## 使用示例文档测试

项目包含了一些示例文档，可以用于集成测试：

### 使用示例文档运行测试

```python
import pytest
import os
from docx import Document

def test_with_example_document():
    """使用实际的示例文档进行测试"""
    example_dir = "doc/example"
    doc_path = os.path.join(example_dir, "例子-原液质量标准.docx")

    # 测试提取功能
    from extract_quality_standards import extract_quality_standards_table_from_docx
    result = extract_quality_standards_table_from_docx(doc_path)

    assert "| 类型 |" in result
    assert "| 检验项目 |" in result
```

### 使用示例模板进行填充测试

```python
def test_fill_example_template():
    """使用示例模板进行填充测试"""
    import tempfile
    from fill_quality_standards import fill_quality_standards_from_markdown

    template_path = "doc/example/模板-质量标准.docx"
    markdown_content = """| 类型 | 检验项目 | 检验方法 | 质量标准 |
| --- | --- | --- | --- |
| 鉴别 | 外观 | 目视 | 澄清 |
"""

    with tempfile.NamedTemporaryFile(suffix='.docx') as output:
        result = fill_quality_standards_from_markdown(
            template_path,
            output.name,
            markdown_content
        )

        assert "Successfully" in result
        assert os.path.exists(output.name)
```

## 调试测试

### 显示详细输出

```bash
# 显示 print 输出
pytest -v -s

# 显示更详细的输出
pytest -vv -s
```

### 使用 pdb 调试器

```bash
# 在失败时进入 pdb
pytest --pdb

# 在测试开始时进入 pdb
pytest --trace
```

### 只运行失败的测试

```bash
# 运行上次失败的测试
pytest --lf

# 先运行失败的测试，然后运行其余的
pytest --ff
```

### 显示本地变量

```bash
pytest -l
```

## 常见使用场景

### 场景 1：开发新功能时运行相关测试

```bash
# 假设你修改了提取功能
# 只运行提取相关的测试
pytest tests/test_extract_quality_standards.py -v

# 如果你想看到代码覆盖
pytest tests/test_extract_quality_standards.py --cov=.claude/skills/ra-doc-assit/scripts/extract_quality_standards
```

### 场景 2：提交代码前运行完整测试

```bash
# 运行所有测试并生成覆盖率报告
pytest --cov=.claude/skills/ra-doc-assit/scripts --cov-report=html --cov-report=term

# 或使用脚本
./tests/run_tests.sh coverage
```

### 场景 3：CI/CD 环境中运行测试

```bash
# 快速运行测试，不生成覆盖率报告
pytest -q

# 只运行单元测试
pytest -m unit -q
```

### 场景 4：调试特定测试

```bash
# 运行特定测试并显示详细输出
pytest tests/test_fill_quality_standards.py::TestFillWordDocumentTable::test_fill_table_successfully -vv -s

# 使用 pdb 调试
pytest tests/test_fill_quality_standards.py::TestFillWordDocumentTable::test_fill_table_successfully --pdb
```

## 测试 Fixture 使用示例

项目提供了多个有用的 fixtures：

```python
import pytest

def test_using_sample_markdown(sample_markdown_table):
    """使用示例 Markdown 表格 fixture"""
    print(sample_markdown_table)
    assert '| 类型 |' in sample_markdown_table

def test_using_temp_dir(temp_output_dir):
    """使用临时输出目录 fixture"""
    import os
    print(f"临时目录: {temp_output_dir}")
    assert os.path.exists(temp_output_dir)

def test_using_example_docs(example_doc_dir):
    """使用示例文档目录 fixture"""
    import os
    print(f"示例文档目录: {example_doc_dir}")
    assert os.path.exists(example_doc_dir)
    assert '例子-原液质量标准.docx' in os.listdir(example_doc_dir)
```

## 持续集成示例

### GitHub Actions 工作流

项目包含了预配置的 GitHub Actions 工作流 (`.github/workflows/test.yml`)：

```yaml
# 运行测试
- name: Run tests
  run: |
    pytest --cov=.claude/skills/ra-doc-assit/scripts \
           --cov-report=xml \
           --verbose
```

### 本地模拟 CI 环境

```bash
# 使用多个 Python 版本测试
pyenv versions  # 列出已安装的 Python 版本

# 在 Python 3.8 中测试
pyenv shell 3.8.18
pytest

# 在 Python 3.11 中测试
pyenv shell 3.11.0
pytest
```

## 性能测试

### 测量测试执行时间

```bash
# 使用 pytest-timeout 插件
pip install pytest-timeout
pytest --timeout=10

# 使用 pytest-benchmark 测量性能
pip install pytest-benchmark
pytest --benchmark-only
```

### 并行运行测试

```bash
# 安装 pytest-xdist
pip install pytest-xdist

# 使用所有 CPU 核心
pytest -n auto

# 使用指定数量的进程
pytest -n 4
```

## 更多资源

- [Pytest 文档](https://docs.pytest.org/)
- [pytest-cov 文档](https://pytest-cov.readthedocs.io/)
- [python-docx 文档](https://python-docx.readthedocs.io/)
- [项目测试 README](./README.md)
