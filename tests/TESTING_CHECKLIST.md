# 测试套件文件清单

## 目录结构

```
tests/
├── __init__.py                           # 测试包初始化
├── conftest.py                           # Pytest 配置和共享 fixtures
├── requirements.txt                      # 测试依赖项
├── README.md                             # 测试指南文档
├── TEST_SUMMARY.md                       # 测试套件总结
├── SAMPLE_TEST_USAGE.md                  # 测试使用示例
├── TESTING_CHECKLIST.md                  # 本文件
├── run_tests.sh                          # 快速测试运行脚本
├── verify_tests.py                       # 测试环境验证脚本
├── .gitignore                            # Git 忽略规则
│
├── test_extract_quality_standards.py     # 提取功能测试
│   ├── TestExtractTextWithFormatting     # 文本提取测试
│   ├── TestFormatAsMarkdownTable         # Markdown 格式化测试
│   ├── TestExtractQualityStandardsTable  # 表格提取测试
│   └── TestExtractQualityStandardsTableFromDocx  # 主函数测试
│
├── test_fill_quality_standards.py        # 填充功能测试
│   ├── TestParseMarkdownTableFromString  # Markdown 解析测试
│   ├── TestParseMarkdownTableFromFile    # 文件解析测试
│   ├── TestRestoreFormattingToCell       # 格式还原测试
│   ├── TestClearTableContent             # 表格清除测试
│   ├── TestInsertTableRows               # 行插入测试
│   ├── TestFindQualityStandardsTable     # 表格查找测试
│   ├── TestMergeCellsInColumn            # 单元格合并测试
│   ├── TestAutoMergeDuplicateCells       # 自动合并测试
│   ├── TestFillWordDocumentTable         # 文档填充测试
│   ├── TestFillQualityStandardsFromMarkdown  # Markdown 填充测试
│   ├── TestFillQualityStandardsFromFile  # 文件填充测试
│   └── TestFillQualityStandardsInplace   # 就地修改测试
│
├── fixtures/                             # 测试数据目录（可扩展）
│   └── (可以添加示例文档和测试数据)
│
└── reports/                              # 测试报告目录
    ├── htmlcov/                          # HTML 覆盖率报告
    │   └── index.html                    # 主报告文件
    └── coverage.xml                      # XML 覆盖率报告

项目根目录配置文件:
├── pytest.ini                            # Pytest 主配置文件
├── .coveragerc                           # 覆盖率工具配置
└── .github/workflows/
    └── test.yml                          # CI/CD 工作流配置
```

## 文件说明

### 核心测试文件

| 文件 | 行数 | 测试类 | 测试数 | 描述 |
|------|------|--------|--------|------|
| test_extract_quality_standards.py | ~350 | 4 | ~20 | 测试提取功能 |
| test_fill_quality_standards.py | ~600 | 12 | ~40 | 测试填充功能 |

### 配置文件

| 文件 | 用途 |
|------|------|
| pytest.ini | Pytest 配置（测试发现、标记、日志等） |
| .coveragerc | 覆盖率报告配置 |
| requirements.txt | Python 依赖项列表 |
| conftest.py | 共享 fixtures 和配置 |

### 文档文件

| 文件 | 内容 |
|------|------|
| README.md | 详细的测试指南 |
| TEST_SUMMARY.md | 测试套件总结 |
| SAMPLE_TEST_USAGE.md | 使用示例和场景 |
| TESTING_CHECKLIST.md | 本文件 - 文件清单 |

### 脚本文件

| 文件 | 用途 | 运行方式 |
|------|------|----------|
| run_tests.sh | 快速运行测试 | `./tests/run_tests.sh [选项]` |
| verify_tests.py | 验证测试环境 | `python tests/verify_tests.py` |

### 其他文件

| 文件 | 用途 |
|------|------|
| __init__.py | 标识为 Python 包 |
| .gitignore | 排除测试输出和临时文件 |

## 快速参考

### 运行测试的命令

```bash
# 所有测试
pytest

# 详细输出
pytest -v

# 覆盖率报告
pytest --cov=.claude/skills/ra-doc-assit/scripts --cov-report=html

# 特定测试文件
pytest tests/test_extract_quality_standards.py

# 使用脚本
./tests/run_tests.sh
./tests/run_tests.sh coverage
./tests/run_tests.sh extract
./tests/run_tests.sh fill
```

### 测试环境验证

```bash
# 检查环境
python tests/verify_tests.py

# 安装依赖
pip install -r tests/requirements.txt

# 检查版本
pytest --version
python --version
```

## 依赖项

主要测试依赖：

- pytest >= 7.4.0
- pytest-cov >= 4.1.0
- pytest-mock >= 3.11.1
- python-docx >= 0.8.11

可选依赖：

- pytest-xdist (并行测试)
- pytest-html (HTML 报告)
- pytest-json-report (JSON 报告)

## 下一步

- [ ] 运行 `python tests/verify_tests.py` 验证环境
- [ ] 安装依赖: `pip install -r tests/requirements.txt`
- [ ] 运行测试: `pytest` 或 `./tests/run_tests.sh`
- [ ] 查看覆盖率: `pytest --cov` 或 `./tests/run_tests.sh coverage`
- [ ] 阅读文档: [tests/README.md](README.md)

## 维护清单

- [ ] 定期更新依赖项版本
- [ ] 添加新功能时添加对应测试
- [ ] 保持测试覆盖率 > 80%
- [ ] 更新测试文档
- [ ] 清理临时文件和报告
- [ ] 检查 CI/CD 状态

## 联系和支持

如有问题或建议，请：
1. 查看 [tests/README.md](README.md)
2. 查看 [tests/SAMPLE_TEST_USAGE.md](SAMPLE_TEST_USAGE.md)
3. 在项目仓库提交 Issue
