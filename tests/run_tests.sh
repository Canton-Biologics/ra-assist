#!/bin/bash

# 测试运行脚本
# 快速运行 RA 文档技能的测试套件

set -e

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}======================================${NC}"
echo -e "${GREEN}RA 文档技能 - 测试运行脚本${NC}"
echo -e "${GREEN}======================================${NC}"
echo ""

# 检查 Python 是否安装
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}错误: 未找到 Python3${NC}"
    exit 1
fi

# 检查是否在项目根目录
if [ ! -f "pytest.ini" ]; then
    echo -e "${RED}错误: 请在项目根目录运行此脚本${NC}"
    exit 1
fi

echo -e "${YELLOW}步骤 1: 检查依赖...${NC}"
if ! python3 -c "import pytest" 2>/dev/null; then
    echo -e "${YELLOW}安装测试依赖...${NC}"
    pip install -q -r tests/requirements.txt
fi
echo -e "${GREEN}✓ 依赖检查完成${NC}"
echo ""

# 创建报告目录
mkdir -p tests/reports

# 解析命令行参数
TEST_TYPE=${1:-all}

case $TEST_TYPE in
    all)
        echo -e "${YELLOW}运行所有测试...${NC}"
        pytest -v
        ;;
    extract)
        echo -e "${YELLOW}运行提取功能测试...${NC}"
        pytest tests/test_extract_quality_standards.py -v
        ;;
    fill)
        echo -e "${YELLOW}运行填充功能测试...${NC}"
        pytest tests/test_fill_quality_standards.py -v
        ;;
    coverage)
        echo -e "${YELLOW}生成测试覆盖率报告...${NC}"
        pytest --cov=.claude/skills/ra-doc-assit/scripts --cov-report=html --cov-report=term
        echo ""
        echo -e "${GREEN}覆盖率报告已生成: tests/reports/htmlcov/index.html${NC}"
        ;;
    fast)
        echo -e "${YELLOW}快速测试（跳过慢速测试）...${NC}"
        pytest -v -m "not slow"
        ;;
    parallel)
        echo -e "${YELLOW}并行运行测试...${NC}"
        if ! python3 -c "import xdist" 2>/dev/null; then
            echo -e "${YELLOW}安装 pytest-xdist...${NC}"
            pip install -q pytest-xdist
        fi
        pytest -v -n auto
        ;;
    *)
        echo -e "${RED}用法: $0 [all|extract|fill|coverage|fast|parallel]${NC}"
        echo ""
        echo "选项:"
        echo "  all       - 运行所有测试（默认）"
        echo "  extract   - 仅运行提取功能测试"
        echo "  fill      - 仅运行填充功能测试"
        echo "  coverage  - 生成测试覆盖率报告"
        echo "  fast      - 快速测试（跳过慢速测试）"
        echo "  parallel  - 并行运行所有测试"
        exit 1
        ;;
esac

echo ""
echo -e "${GREEN}======================================${NC}"
echo -e "${GREEN}测试完成！${NC}"
echo -e "${GREEN}======================================${NC}"
