# Windows 环境安装指南

欢迎使用 RA 质量文档编写技能！本指南将帮助你在 Windows 环境下完成完整的安装配置。

## 📋 目录

- [系统要求](#系统要求)
- [快速开始](#快速开始)
- [安装步骤](#安装步骤)
- [配置 API](#安装项目)
- [更新和维护](#更新和维护)

---

## 系统要求

### 必需软件

| 软件 | 最低版本 | 推荐版本 | 下载地址 |
|------|---------|---------|---------|
| **Windows** | Windows 10 | Windows 11 | - |
| **Python** | 3.8 | 3.11+ | [python.org](https://www.python.org/downloads/) |
| **Git** | 2.0 | 最新版 | [git-scm.com](https://git-scm.com/download/win) |
| **Node.js** | 14.x | 最新版 | [nodejs.org](https://nodejs.org/zh-cn/download) Windows安装程序msi |
| **Claude Code** | 最新版 | - | [npm](https://www.npmjs.com/package/@anthropic-ai/claude-code) |

**⚠️ 注意**：Git 是必需的前提条件，详见[前提条件](#前提条件)章节。

### 可选软件

| 软件 | 用途 | 下载地址 |
|------|------|---------|
| **VS Code** | 代码编辑器 | [code.visualstudio.com](https://code.visualstudio.com/) |


---


## 安装步骤

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


### 安装 Python

1. 访问 [Python 官网](https://www.python.org/downloads/windows/)
2. 下载 **Windows installer (64-bit)**
3. **⚠️ 重要**: 运行安装程序时，务必勾选 **"Add Python to PATH"**
4. 点击 "Install Now" 完成安装

**验证安装**：
```powershell
python --version
# 应显示: Python 3.x.x
```

### 安装 Node.js

Node.js 是运行 JavaScript 代码的平台，用于执行本项目的脚本。

1. 访问 [Node.js 官网](https://nodejs.org/zh-cn/download)
2. 下载 Windows 安装程序msi
3. 运行安装程序，使用默认设置，一路点击 "Next"
4. 安装完成后重启终端（PowerShell 或 CMD）

授权npm脚本运行权限
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

**验证 Node.js 是否已正确安装**：
```powershell
node --version
```
应显示 Node.js 的版本号。

### 安装Claude Code

Claude Code 是一个命令行工具，用于与 Claude 模型交互。

1. **安装 Claude Code**
   ```powershell
   npm install -g @anthropic-ai/claude-code
   ```

2. **验证安装**
   ```powershell
   claude --version
   ```
   应显示 Claude Code 的版本号。

---

### 配置 API 密钥 (智谱GLM)

参考：使用智谱 GLM API：https://docs.bigmodel.cn/cn/coding-plan/tool/claude：

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

## 安装项目
### 克隆项目

克隆 GitHub 仓库时，根据你的网络环境和仓库权限，有两种方式：

**方式 1: HTTPS（推荐用于公共仓库）**
```powershell
git clone https://github.com/Canton-Biologics/ra-assist.git
```
会弹出一个窗口，要求输入 GitHub 用户名和密码。


**方式 2： 使用Token**（如需要）：
```powershell
git clone https://YOUR_TOKEN@github.com/Canton-Biologics/ra-assist.git
```
---


### 安装 Python 依赖

```powershell
# 安装核心依赖
pip install python-docx

```

**验证安装**：
```powershell
python -c "import docx; print('✅ python-docx 安装成功')"
```

### 启动claude code

```powershell
cd ra-assist
claude
```

---


## 更新和维护

### 更新脚本

```powershell
# 拉取最新代码
git pull origin main

# 或使用 rebase 保持干净的提交历史
git pull --rebase origin main
```

## 反馈问题

如果在使用过程中遇到问题，请通过 GitHub Issues 反馈：

1. 访问项目仓库：[Canton-Biologics/ra-assist](https://github.com/Canton-Biologics/ra-assist)
2. 点击 "Issues" 标签
3. 点击 "New Issue"
4. 描述问题详情，包括复现步骤和预期行为
5. 提交 Issue

我们会尽快回复并解决您的问题。