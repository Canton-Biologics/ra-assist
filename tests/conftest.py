"""
Pytest configuration and fixtures for RA documentation tests
"""

import pytest
import os
import sys
import tempfile
from pathlib import Path

# Add scripts directory to Python path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
scripts_dir = os.path.join(project_root, '.claude', 'skills', 'ra-doc-assit', 'scripts')
sys.path.insert(0, scripts_dir)


@pytest.fixture(scope="session")
def test_data_dir():
    """Get the test data directory"""
    return os.path.join(os.path.dirname(__file__), 'fixtures')


@pytest.fixture(scope="session")
def example_doc_dir():
    """Get the example documents directory"""
    return os.path.join(os.path.dirname(__file__), '..', 'doc', 'example')


@pytest.fixture
def temp_output_dir():
    """Create a temporary directory for test outputs"""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


@pytest.fixture
def sample_markdown_table():
    """Provide a sample markdown table for testing"""
    return """| 类型 | 检验项目 | 检验方法 | 质量标准 |
| --- | --- | --- | --- |
| 鉴别 | 外观 | 目视 | 澄清、无色液体 |
| 鉴别 | pH值 | pH计 | 6.0-8.0 |
| 检查 | 相关物质 | HPLC | 单一杂质≤1.0%，总杂质≤3.0% |
| 检查 | 含量 | HPLC | 95.0%-105.0% |
"""


@pytest.fixture
def sample_markdown_with_formatting():
    """Provide markdown table with Unicode formatting for testing"""
    return """| 检验项目 | 公式 | 标准 |
| --- | --- | --- |
| 分子量 | C₁₂H₂₂O₁₁ | 342.30 |
| 浓度 | H²SO₄ | ≥98% |
| 体积 | m³ | 1.5-2.0 |
"""
