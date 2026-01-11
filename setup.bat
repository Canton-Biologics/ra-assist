@echo off
chcp 65001 >nul
echo ========================================
echo   RA 质量文档编写 - Windows 环境设置
echo ========================================
echo.

REM 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 警告: 建议以管理员身份运行此脚本
    echo 右键点击此文件 -^> "以管理员身份运行"
    echo.
    pause
)

echo [1/7] 检查 Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [✗] 错误: Python 未安装或未添加到 PATH
    echo.
    echo 请访问 https://www.python.org/downloads/ 下载 Python 3.8+
    echo 安装时务必勾选 "Add Python to PATH"
    pause
    exit /b 1
)
echo [√] Python 已安装
python --version
echo.

echo [2/7] 检查 pip...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [✗] 错误: pip 未安装
    pause
    exit /b 1
)
echo [√] pip 已安装
echo.

echo [3/7] 安装核心依赖...
echo 正在安装 python-docx...
pip install python-docx
if %errorlevel% neq 0 (
    echo [✗] 错误: python-docx 安装失败
    pause
    exit /b 1
)
echo [√] python-docx 安装成功
echo.

echo [4/7] 验证依赖...
python -c "import docx; print('[√] python-docx 导入成功')" 2>nul
if %errorlevel% neq 0 (
    echo [✗] 错误: python-docx 导入失败
    pause
    exit /b 1
)
echo.

echo [5/7] 检查 Git...
git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] 警告: Git 未安装
    echo Git 用于版本控制和更新功能
    echo 请访问 https://git-scm.com/download/win 下载安装
    echo.
) else (
    echo [√] Git 已安装
    git --version
)
echo.

echo [6/7] 检查 Claude Code...
claude --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] 警告: Claude Code 未安装
    echo Claude Code 用于使用 RA 质量文档编写技能
    echo.
    echo 安装方法 1 (使用 npm):
    echo   npm install -g @anthropic-ai/claude-code
    echo.
    echo 安装方法 2 (使用 pip):
    echo   pip install claude-code
    echo.
    set /p continue="是否继续? (Y/N): "
    if /i not "%continue%"=="Y" (
        pause
        exit /b 1
    )
) else (
    echo [√] Claude Code 已安装
    claude --version
)
echo.

echo [7/7] 运行测试验证...
if exist "tests\verify_tests.py" (
    echo 正在运行环境验证...
    python tests\verify_tests.py
    if %errorlevel% neq 0 (
        echo [!] 警告: 环境验证失败
        echo 但这不影响核心功能的使用
        echo.
    ) else (
        echo [√] 环境验证通过
    )
) else (
    echo [!] 警告: 未找到测试脚本
    echo 请确保在项目根目录运行此脚本
    echo.
)
echo.

echo ========================================
echo          环境设置完成！
echo ========================================
echo.
echo 下一步操作:
echo.
echo 1. 配置 API 密钥（如未配置）:
echo    setx ANTHROPIC_API_KEY "你的API密钥"
echo.
echo 2. 重启终端使环境变量生效
echo.
echo 3. 开始使用:
echo    cd %CD%
echo    claude
echo.
echo 4. 更新功能:
echo    git pull origin main
echo    pip install --upgrade python-docx
echo.
echo ========================================
pause
