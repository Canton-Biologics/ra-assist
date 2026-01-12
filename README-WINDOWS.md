# Windows 环境安装指南

欢迎使用 RA 质量文档编写技能！本指南将帮助你在 Windows 环境下完成完整的安装配置。

## 📋 目录

- [系统要求](#系统要求)
- [前提条件](#前提条件)
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

**⚠️ 注意**：Git 是必需的前提条件，详见[前提条件](#前提条件)章节。

### 可选软件

| 软件 | 用途 | 下载地址 |
|------|------|---------|
| **Node.js** | 安装 Claude Code (npm 方式) | [nodejs.org](https://nodejs.org/zh-cn/download) Windows安装程序msi |
| **VS Code** | 代码编辑器 | [code.visualstudio.com](https://code.visualstudio.com/) |

授权npm脚本运行权限
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```
---

## 前提条件

在开始安装之前，请确保已经安装了 **Git**。

### 安装 Git

Git 是版本控制工具，是使用本项目的**必需前提条件**。

1. 访问 [Git 官网下载页面](https://git-scm.com/download/win)
2. 下载 Windows 安装程序
3. 运行安装程序，使用默认设置，一路点击 "Next"
4. 安装完成后重启终端（PowerShell 或 CMD）

**验证 Git 是否已正确安装**：
```powershell
git --version
# 应显示: git version 2.x.x
```

如果看到版本号，说明 Git 安装成功。

### 关于 GitHub 访问

克隆 GitHub 仓库时，根据你的网络环境和仓库权限，有两种方式：

**方式 1: HTTPS（推荐用于公共仓库）**
```powershell
git clone https://github.com/Canton-Biologics/ra-assist.git
```

**方式 2: 使用 Personal Access Token（推荐用于私有仓库或二次验证）**

如果仓库是私有的，或者你的账号启用了二次验证（2FA），需要使用 Personal Access Token：

1. **创建 GitHub Personal Access Token**：
   - 登录 GitHub，访问 https://github.com/settings/tokens
   - 点击 "Generate new token" → "Generate new token (classic)"
   - 设置 token 名称（如 "ra-assist"）
   - 选择权限：勾选 `repo`（完整仓库访问权限）
   - 点击 "Generate token"
   - **⚠️ 重要**: 复制生成的 token（格式：`ghp_xxxxx`），此 token 只显示一次！

2. **使用 Token 克隆仓库**：
   ```powershell
   git clone https://YOUR_TOKEN@github.com/Canton-Biologics/ra-assist.git
   ```
   将 `YOUR_TOKEN` 替换为你刚才复制的 token。

   示例：
   ```powershell
   git clone https://ghp_1234567890abcdef@github.com/Canton-Biologics/ra-assist.git
   ```

3. **保存 Token（可选，避免重复输入）**：
   Windows 可以使用 Git Credential Manager 来保存凭据：
   ```powershell
   git config --global credential.helper manager-core
   ```
   第一次输入 token 后，Git 会记住它，后续操作不需要再次输入。

**安全提示**：
- ⚠️ 不要将 token 提交到代码库或分享给他人
- ⚠️ Token 具有仓库完整访问权限，请妥善保管
- ⚠️ 定期更新 token 以提高安全性

---

## 快速开始

### 方式一：自动安装（推荐）⚡

1. **下载项目**

   选择适合你的方式克隆仓库：

   **公共访问**：
   ```powershell
   git clone https://github.com/Canton-Biologics/ra-assist.git
   cd ra-assist
   ```

   **使用 Token**（如需要）：
   ```powershell
   git clone https://YOUR_TOKEN@github.com/Canton-Biologics/ra-assist.git
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

1. 访问 [Python 官网](https://www.python.org/downloads/windows/)
2. 下载 **Windows installer (64-bit)**
3. **⚠️ 重要**: 运行安装程序时，务必勾选 **"Add Python to PATH"**
4. 点击 "Install Now" 完成安装

**验证安装**：
```powershell
python --version
# 应显示: Python 3.x.x
```

### 步骤 2: 克隆项目

**前提**：确保已安装 Git（见[前提条件](#前提条件)）。

选择适合你的方式克隆仓库：

**公共访问**：
```powershell
# 克隆仓库到本地
git clone https://github.com/Canton-Biologics/ra-assist.git

# 进入项目目录
cd ra-assist
```

**使用 Personal Access Token**（如需要）：
```powershell
# 克隆仓库到本地（将 YOUR_TOKEN 替换为你的 token）
git clone https://YOUR_TOKEN@github.com/Canton-Biologics/ra-assist.git

# 进入项目目录
cd ra-assist
```

### 步骤 3: 安装 Python 依赖

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

### 步骤 4: 安装 Claude Code

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

如果使用智谱 GLM API：https://docs.bigmodel.cn/cn/coding-plan/tool/claude：

```json
# 编辑或新增 `settings.json` 文件
# MacOS & Linux 为 `~/.claude/settings.json`
# Windows 为`用户目录/.claude/settings.json`
# 新增或修改里面的 env 字段
# 注意替换里面的 `your_zhipu_api_key` 为您上一步获取到的 API Key
{
  "env": {
    "ANTHROPIC_AUTH_TOKEN": "your_zhipu_api_key",
    "ANTHROPIC_BASE_URL": "https://open.bigmodel.cn/api/anthropic",
    "API_TIMEOUT_MS": "3000000",
    "CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC": 1
  }
}
# 再编辑或新增 `.claude.json` 文件
# MacOS & Linux 为 `~/.claude.json`
# Windows 为`用户目录/.claude.json`
# 新增 `hasCompletedOnboarding` 参数
{
  "hasCompletedOnboarding": true
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
