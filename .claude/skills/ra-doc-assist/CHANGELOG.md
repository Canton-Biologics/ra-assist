# Skill 更新说明 - 通用转换器

## 版本信息
- **版本号**: v2.1
- **更新日期**: 2026-02-09
- **更新内容**: 新增通用转换器，支持任意检验方法

---

## 🎯 主要改进

### 问题
之前的版本需要为每种检验方法编写专门的转换代码：
- 每增加一种 SOP，就要在 `method_converter.py` 中添加新的 `convert_xxx()` 方法
- 需要在 `sop_extractor.py` 中添加对应的 `extract_xxx()` 方法
- 修改代码后需要测试和维护

### 解决方案
新增 **通用转换器（UniversalMethodConverter）**，实现：
- ✅ 无需为每种方法编写代码
- ✅ 自动识别 SOP 中的章节内容
- ✅ 支持任意检验方法
- ✅ 使用配置文件批量转换

---

## 🚀 新增功能

### 1. 通用转换器脚本
**文件**: `scripts/convert_any_method.py`

```bash
# 转换单个方法
python convert_any_method.py --template 模板.docx --sop SOP.docx --method "章节名称"

# 批量转换
python convert_any_method.py --template 模板.docx --config methods.txt
```

### 2. 通用转换器核心类
**文件**: `scripts/regulatory_core/universal_converter.py`

主要类：`UniversalMethodConverter`

核心方法：
- `extract_sop_content()` - 通用内容提取
- `format_content_for_output()` - 标准格式化
- `convert()` - 单方法转换
- `convert_multi()` - 多方法批量转换

### 3. 配置文件支持
**文件**: `scripts/methods_example.txt`

格式：`SOP路径:模板章节名称:方法类型`

示例：
```bash
D:/docs/纯度SOP.docx:纯度（SEC-HPLC）:SEC-HPLC
D:/docs/肽图SOP.docx:肽图:RP-UPLC
D:/docs/电泳SOP.docx:纯度（CE-SDS）:CE-SDS
```

---

## 📊 两种转换器对比

| 特性 | 通用转换器 | 专用转换器 |
|------|-----------|-----------|
| **脚本** | `convert_any_method.py` | `convert_sop_to_method.py` |
| **适用性** | ✅ 任意方法 | ⚠️ 仅限预定义方法 |
| **新增方法** | ✅ 无需修改代码 | ❌ 需要编写代码 |
| **配置** | 命令行/配置文件 | 需指定方法类型 |
| **灵活性** | ✅ 高 | ⚠️ 固定格式 |
| **准确度** | ⚠️ 依赖关键词 | ✅ 针对性优化 |

---

## 💡 使用示例

### 添加新的检验方法（无需修改代码）

**场景**: 需要新增"毛细管电泳"方法

**使用通用转换器**：
```bash
# 直接使用，无需修改代码
python convert_any_method.py \
  --template 模板.docx \
  --sop 毛细管电泳SOP.docx \
  --method "毛细管电泳"
```

**如果使用专用转换器**（旧方法）：
1. 在 `sop_extractor.py` 添加 `extract_ce()` 方法
2. 在 `method_converter.py` 添加 `convert_ce()` 方法
3. 在 `convert_sop_to_method.py` 添加 `--ce` 参数
4. 测试新代码
5. 分发更新

---

## 📁 文件结构

```
scripts/
├── regulatory_core/
│   ├── __init__.py                    # 导出 UniversalMethodConverter
│   ├── sop_extractor.py               # SOP 内容提取
│   ├── method_converter.py            # 专用转换器
│   └── universal_converter.py         # ✨ 新增：通用转换器
├── convert_sop_to_method.py           # 专用转换器脚本
├── convert_any_method.py              # ✨ 新增：通用转换器脚本
├── methods_example.txt                # ✨ 新增：配置文件示例
├── extract_quality_standards.py
└── fill_quality_standards.py
```

---

## 🔧 技术细节

### 通用转换器工作原理

1. **章节识别**
   - 使用预定义的关键词列表识别各章节
   - 支持中英文关键词
   - 可扩展的关键词配置

2. **内容提取**
   - 基于关键词范围提取内容
   - 自动跳过章节标题
   - 提取表格数据（如色谱条件）

3. **格式化**
   - 按照医药注册标准格式组织内容
   - 6 个核心章节：原理、设备材料、操作步骤、系统适用性、结果计算、合格标准
   - 自动应用文档样式

4. **章节定位**
   - 在模板中查找匹配的章节名称
   - 支持模糊匹配
   - 自动识别章节范围

### 章节关键词映射

```python
SECTION_MAPPING = {
    '原理': ['实验原理', '原理', 'Principle', ...],
    '设备材料试剂': ['实验材料', '试剂与材料', '仪器与设备', ...],
    '样品处理': ['样品处理', '样品制备', '溶液配制', ...],
    '操作步骤': ['操作步骤', '测定法', '检验步骤', ...],
    '色谱条件': ['色谱条件', '仪器条件', ...],
    '系统适用性': ['系统适用性', 'System Suitability'],
    '结果计算': ['结果计算', '计算', '数据处理', ...],
    '可接受标准': ['可接受标准', '合格标准', ...],
}
```

---

## ⚠️ 注意事项

### 通用转换器的限制

1. **依赖章节关键词**
   - SOP 文档需要有清晰的章节标题
   - 章节标题需要包含识别关键词
   - 不规范的 SOP 可能识别不准确

2. **格式化相对通用**
   - 使用标准格式，不针对特定方法优化
   - 某些特殊情况可能需要手动调整

3. **色谱条件提取**
   - 依赖表格格式
   - 非表格格式的条件可能无法提取

### 何时使用专用转换器

- 需要针对特定方法优化内容格式
- SOP 文档格式不规范
- 需要特殊的章节处理逻辑
- 对准确性要求极高

---

## 📝 更新日志

### v2.1 (2026-02-09)
- ✅ 新增通用转换器 `UniversalMethodConverter`
- ✅ 新增通用转换脚本 `convert_any_method.py`
- ✅ 支持配置文件批量转换
- ✅ 支持任意检验方法，无需修改代码
- ✅ 添加 `--list-sections` 参数列出模板章节
- ✅ 更新文档，添加对比说明

### v2.0 (2026-02-09)
- ✅ 核心代码整合到 `regulatory_core/`
- ✅ 支持肽图 RP-UPLC 方法
- ✅ 移除外部项目依赖
- ✅ Skill 可独立分发

### v1.0 (初始版本)
- SEC-HPLC 和 SoloVPE 方法支持
- 依赖外部 regulatory_converter 项目

---

## 🎓 最佳实践

### 1. 优先使用通用转换器
除非有特殊需求，否则优先使用 `convert_any_method.py`

### 2. 使用配置文件管理批量转换
创建配置文件记录所有需要转换的方法，便于维护和重用

### 3. 先用 --list-sections 检查模板
在转换前，使用 `--list-sections` 确认模板中的章节名称

### 4. 保留专用转换器用于已优化方法
对于 SEC-HPLC、SoloVPE、肽图等已优化的方法，可以继续使用专用转换器

---

## 🔮 未来改进方向

- [ ] 支持自定义章节关键词配置
- [ ] 支持 LLM 辅助内容识别和提取
- [ ] 支持更多文档格式（PDF、Excel）
- [ ] 图形化配置界面
- [ ] 转换结果预览功能
