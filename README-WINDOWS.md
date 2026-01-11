# Windows 环境安装指南

欢迎使用 RA 质量文档编写技能！本指南将帮助你在 Windows 环境下完成完整的安装配置。

## 📋 目录

- [系统要求](#系统要求)
- [快速开始](#快速开始)
- [详细安装步骤](#详细安装步骤)
- [配置 API](#配置-api)
- [验证安装](#验证安装)
- [更新和维护](#更新和维护)
- [常见问题](#常见问题)
- [下一步](#下一步)

---

## 系统要求

### 必需软件

| 软件 | 最低版本 | 推荐版本 | 下载地址 |
|------|---------|---------|---------|
| **Windows** | Windows 10 | Windows 11 | - |
| **Python** | 3.8 | 3.11+ | [python.org](https://www.python.org/downloads/) |
| **Git** | 2.0 | 最新版 | [git-scm.com](https://git-scm.com/download/win) |

### 可选软件

| 软件 | 用途 | 下载地址 |
|------|------|---------|
| **Node.js** | 安装 Claude Code (npm 方式) | [nodejs.org](https://nodejs.org/) |
| **VS Code** | 代码编辑器 | [code.visualstudio.com](https://code.visualstudio.com/) |

---

## 快速开始

### 方式一：自动安装（推荐）⚡

1. **下载项目**
   ```powershell
   git clone https://github.com/Canton-Biologics/ra-assist.git
   cd ra-assist
   ```
2. **运行自动安装脚本**
   - 右键点击 `setup.bat`
   - 选择"**以管理员身份运行**"
   - 等待安装完成

3. **配置 API 密钥**
   ```powershell
   setx ANTHROPIC_API_KEY "sk-ant-你的API密钥"
   ```

4. **重启终端并开始使用**
   ```powershell
   claude
   ```

✅ 完成！开始使用技能吧。

### 方式二：手动安装

如果自动安装失败，请参考下面的[详细安装步骤](#详细安装步骤)。

---

## 详细安装步骤

### 步骤 1: 安装 Python

1. 访问 [Python 官网](https://www.python.org/downloads/)
2. 下载 **Windows installer (64-bit)**
3. **⚠️ 重要**: 运行安装程序时，务必勾选 **"Add Python to PATH"**
4. 点击 "Install Now" 完成安装

**验证安装**：
```powershell
python --version
# 应显示: Python 3.x.x
```

### 步骤 2: 安装 Git

1. 访问 [Git 下载页面](https://git-scm.com/download/win)
2. 下载并运行安装程序
3. 使用默认设置，一路点击 "Next"

**验证安装**：
```powershell
git --version
# 应显示: git version 2.x.x
```

### 步骤 3: 克隆项目

```powershell
# 克隆仓库到本地
git clone https://github.com/Canton-Biologics/ra-assist.git 

# 进入项目目录
cd ra-assist
```

### 步骤 4: 安装 Python 依赖

```powershell
# 安装核心依赖
pip install python-docx

# （可选）安装测试依赖
pip install -r tests/requirements.txt
```

**验证安装**：
```powershell
python -c "import docx; print('✅ python-docx 安装成功')"
```

### 步骤 5: 安装 Claude Code

**方式 A: 使用 npm**（需要先安装 Node.js）
```powershell
npm install -g @anthropic-ai/claude-code
```

**方式 B: 使用 pip**
```powershell
pip install claude-code
```

**验证安装**：
```powershell
claude --version
```

---

## 配置 API

### 获取 API 密钥

1. 访问 [Anthropic Console](https://console.anthropic.com/)
2. 注册账号（如需要）
3. 在 API Keys 部分创建新的 API 密钥
4. 复制密钥（格式：`sk-ant-xxxxx`）

### 配置 Anthropic API

**方法 1: 环境变量**（推荐）

```powershell
# 设置用户环境变量
setx ANTHROPIC_API_KEY "sk-ant-你的API密钥"

# 重启终端使环境变量生效
```

**方法 2: 配置文件**

1. 创建配置目录：
   ```powershell
   mkdir %USERPROFILE%\.claude
   ```

2. 创建配置文件：
   ```powershell
   notepad %USERPROFILE%\.claude\config.json
   ```

3. 添加以下内容：
   ```json
   {
     "apiKey": "sk-ant-你的API密钥",
     "baseUrl": "https://api.anthropic.com"
   }
   ```

### 配置 GLM API（可选）

如果使用智谱 GLM 或其他兼容服务：

```json
{
  "apiKey": "你的GLM_API密钥",
  "baseUrl": "https://open.bigmodel.cn/api/paas/v4/",
  "model": "glm-4"
}
```

### 测试 API 连接

```powershell
claude ask "Hello, can you hear me?"
```

---

## 验证安装

运行环境验证脚本：

```powershell
python tests/verify_tests.py
```

应该看到以下输出：
```
✓ 所有检查通过！测试环境已就绪。
```

---

## 更新和维护

### 更新代码

```powershell
# 查看当前分支
git branch

# 拉取最新代码
git pull origin main

# 或使用 rebase 保持干净的提交历史
git pull --rebase origin main
```

### 更新依赖

```powershell
# 更新核心依赖
pip install --upgrade python-docx

# 更新测试依赖
pip install --upgrade -r tests/requirements.txt

# 更新 Claude Code
npm update -g @anthropic-ai/claude-code
# 或
pip install --upgrade claude-code
```

### 查看已安装的包

```powershell
pip list
```

---
