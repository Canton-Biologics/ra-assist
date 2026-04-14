# RA 质量文档编写助手 - Skill 打包说明

## 概述

此 Skill 已整合为独立可分发的版本，所有核心代码都包含在 `scripts/regulatory_core/` 目录中，可以无需外部依赖直接使用。

## 目录结构

```
ra-doc-assist/
├── SKILL.md                  # Skill 定义和说明
├── REFERENCE.md              # 参考资料
├── README.md                 # 本文件
├── input/                    # 默认模版与 SOP（可替换）
└── scripts/
    ├── regulatory_core/      # 核心模块
    │   ├── sop_extractor.py
    │   ├── analysis_integrator.py
    │   └── method_section_order.py
    ├── examples/             # `-c` 映射配置示例
    ├── integrate_sop_method.py
    ├── refine_extracted.py
    ├── ra_compliance.py
    ├── update_toc.py
    ├── extract_quality_standards.py
    └── fill_quality_standards.py
```

## 打包分发

### 方法一：打包为 ZIP 文件

```bash
# 在 skill 目录的父目录执行
cd ~/.claude/skills
zip -r ra-doc-assist.zip ra-doc-assist/

# 分发给其他人，他们只需要解压到：
# ~/.claude/skills/ra-doc-assist/
```

### 方法二：使用 Git（推荐）

```bash
cd ~/.claude/skills/ra-doc-assist
git init
git add .
git commit -m "Initial release"

# 推送到 GitHub 或其他 Git 仓库
git remote add origin <your-repo-url>
git push -u origin main

# 其他人通过 git clone 获取
cd ~/.claude/skills
git clone <your-repo-url> ra-doc-assist
```

### 方法三：直接复制目录

```bash
# Windows
xcopy "C:\Users\wenjian.bao\.claude\skills\ra-doc-assist" "目标路径\ra-doc-assist" /E /I

# Linux/Mac
cp -r ~/.claude/skills/ra-doc-assist /path/to/destination/
```

## 安装（接收者）

### 1. 将 skill 文件放入正确位置

**Windows:**
```
C:\Users\<用户名>\.claude\skills\ra-doc-assist\
```

**Linux/Mac:**
```
~/.claude/skills/ra-doc-assist/
```

### 2. 安装依赖

```bash
pip install python-docx
```

### 3. 测试安装

```bash
cd ~/.claude/skills/ra-doc-assist/scripts
python integrate_sop_method.py --help
```

## 使用示例

### 通过 Claude Code 使用

在 Claude Code 中直接描述需求即可（见同目录 **`SKILL.md`** 的「执行铁律」与 CLI 流程）。

### 直接使用脚本（三步）

```bash
cd ~/.claude/skills/ra-doc-assist/scripts

# 1) 提取
python integrate_sop_method.py -t "../input/32s42-分析方法-模板文件.docx" \
  -c examples/methods.single.example.json --extract-only -o ../output/extracted.json

# 2) 规则精简
python refine_extracted.py ../output/extracted.json ../output/refined.json

# 3) 写回 Word
python integrate_sop_method.py -t "../input/32s42-分析方法-模板文件.docx" \
  --from-json ../output/refined.json
```

## 依赖说明

### 必需依赖

- `python-docx`: 用于读取和写入 Word 文档

```bash
pip install python-docx
```

### 可选依赖

如果需要使用完整功能（如 LLM 智能提取），需要额外安装：

```bash
pip install openai anthropic
```

## 核心模块说明

### regulatory_core/sop_extractor.py

从 SOP 第四章「程序」提取结构化小节（原理、材料、步骤、可接受标准等），供整合流程使用。

### regulatory_core/analysis_integrator.py

负责 **`--extract-only` 生成 JSON** 与 **`--from-json` 写回 Word**（含表格/样式锚点逻辑）。

### regulatory_core/method_section_order.py

方法章节 **物理重排**、**`--method-order`** 与 **`new_sections` 骨架插入**。

## 版本信息

- **当前版本**: 2.0 (独立打包版)
- **最后更新**: 2026-02-09
- **主要特性**:
  - 完全独立，无需外部项目依赖
  - 支持肽图 RP-UPLC 方法
  - 简化的模块结构

## 更新日志

### v2.0 (2026-02-09)
- ✅ 将核心代码整合到 `scripts/regulatory_core/`
- ✅ 添加肽图 RP-UPLC 方法支持
- ✅ 更新所有脚本使用本地模块
- ✅ 移除外部项目依赖
- ✅ 更新文档说明

### v1.0 (初始版本)
- SEC-HPLC 和 SoloVPE 方法支持
- 依赖外部 regulatory_converter 项目

## 常见问题

### Q: Skill 放在哪个目录？
A: 不同系统的位置：
- Windows: `C:\Users\<用户名>\.claude\skills\`
- Linux/Mac: `~/.claude/skills/`

### Q: 需要安装 regulatory_converter 项目吗？
A: 不需要。v2.0 版本已完全独立，核心代码包含在 skill 内部。

### Q: 如何添加新的检验方法？
A: 优先使用 **`integrate_sop_method.py --merge-sops` / `--infer-section` / `--method-order`**（见 **`SKILL.md`**）；确需改解析规则时再改 `regulatory_core` 模块。

### Q: 脚本找不到模块怎么办？
A: 确保脚本正确添加了 `regulatory_core` 目录到 Python 路径。脚本已自动处理，直接在 scripts 目录运行即可。

## 技术支持

如遇到问题，请检查：
1. Python 版本 >= 3.7
2. 已安装 python-docx
3. 文件路径正确（支持相对路径和绝对路径）
4. SOP 文档格式符合标准

## 许可证

内部使用，请勿外传。
