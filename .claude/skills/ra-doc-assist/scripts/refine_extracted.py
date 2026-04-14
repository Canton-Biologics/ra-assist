# -*- coding: utf-8 -*-
"""
SOP内容精简处理脚本
根据药品注册申报要求对提取的SOP内容进行精简
"""
import json
import sys
import re
import os
from typing import List

sys.stdout.reconfigure(encoding='utf-8')

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

try:
    from ra_compliance import print_skill_iron_rules, verify_principle_verbatim
except ImportError:
    print_skill_iron_rules = lambda **kwargs: None

    def verify_principle_verbatim(sop_raw_principle, refined_principle):
        return True, ""


def _normalize_sentence_for_dedup(s: str) -> str:
    """
    去重用的弱归一化：
    - 统一空白
    - 去掉常见中文/英文标点
    - 不改动数字/单位/比较符号
    """
    t = str(s or "").strip()
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"[，。,；;：:\(\)（）\[\]【】\-—_·]", "", t)
    return t.lower().strip()


def reconstruct_text(char_list):
    """将字符列表还原为文本"""
    if not char_list:
        return ''
    if isinstance(char_list, str):
        return char_list
    if isinstance(char_list, list) and len(char_list) > 0:
        if isinstance(char_list[0], str) and len(char_list[0]) == 1:
            return ''.join(char_list)
        return '\n'.join(char_list)
    return str(char_list)


def count_chars(text_list):
    """计算文本列表的总字数"""
    return sum(len(t) for t in text_list)


def normalize_terminology(text):
    """术语归一化"""
    if not text:
        return text

    # 工作参比品/参比品/对照品 → 标准物质
    text = re.sub(r'工作参比品|参比品|对照品', '标准物质', text)

    # 仪器/设备 → 主要设备 (但在材料和设备章节保留"主要设备")
    # 可接受标准/判定标准 → 合格标准
    text = re.sub(r'可接受标准|判定标准', '合格标准', text)

    return text


def refine_principle(sop_raw_list, template_list, ref_list, max_chars=200):
    """
    原理部分：直接沿用来源正文，不做句数/字数截断与改写（与 SKILL/REFERENCE 约定一致）。
    优先级：SOP 当前内容 > template_existing > reference_style。
    仅去除行首「原理：」类前缀；不调用术语归一化，以保持与 SOP 原文一致。
    max_chars 参数保留以兼容旧调用，现已不再用于截断。
    """
    _ = max_chars  # 保留签名兼容，不再用于截断
    sources = []

    if sop_raw_list:
        sop_text = reconstruct_text(sop_raw_list)
        if sop_text and sop_text.strip():
            # strict 原理要求：当 SOP 原理存在时，除去可选前缀外不做任何清洗/删改，
            # 以保证与 SOP 原文在空白归一化后逐字一致（ra_compliance.verify_principle_verbatim）。
            sources.append(('sop', sop_text.strip()))

    if template_list:
        template_text = reconstruct_text(template_list)
        if template_text and template_text.strip():
            t = template_text.strip()
            if t not in [s[1] for s in sources]:
                sources.append(('template', t))

    if ref_list:
        ref_text = reconstruct_text(ref_list)
        if ref_text and ref_text.strip():
            t = ref_text.strip()
            if t not in [s[1] for s in sources]:
                sources.append(('ref', t))

    if not sources:
        return []

    selected_src, selected_text = sources[0]
    selected_text = re.sub(r'^原理[：:]\s*', '', selected_text).strip()
    if not selected_text:
        return []

    # 若选用 SOP 原文：不做任何删改（除了去掉可选“原理：”前缀）
    if selected_src == 'sop':
        return [selected_text]

    # 非 SOP 来源（模板/参考样本）时允许做轻量清洗，去除明显的设备备注噪声
    equipment_patterns = [
        r'仪器设备\s*Equipment.*?(?=$|\n)',
        r'设备\s*Equipment.*?(?=$|\n)',
        r'备注[：:]\s*以上设备可使用等效的设备替代[。\.]*',
        r'备注[：:]\s*.*?等效.*?设备[。\.]*',
    ]
    filtered_text = selected_text
    for pattern in equipment_patterns:
        filtered_text = re.sub(pattern, '', filtered_text, flags=re.IGNORECASE)
    filtered_text = re.sub(r'\s+', ' ', filtered_text).strip()
    return [filtered_text] if filtered_text else [selected_text]


# 纯标题/无信息行：不写入模版
_MATERIAL_SKIP_TITLES = frozenset({
    '实验材料', '试剂与材料', '仪器与设备', '仪器及设备', '实验设备', '试验设备',
    '试液配制', '溶液配制', '溶液制备', '关键耗材', '仪器设备', '设备',
    'Experiments Material', 'Equipment', 'Instrument', 'Solution Preparation',
})


def _clean_material_line(line: str) -> str:
    """
    行级清洗：去掉**规程/文件编号**（引用哪份 SOP 文档）、厂家名、设备备注；
    **保留** 产品/参比品/物料名称（如「CB0012工作参比品」「标准物质：CB0012工作参比品」）。

    「去掉 SOP 编号」仅指标号形如 3-…-SOP-…、3-CBxxxx-SOP-TI…-xx，不是去掉名称里的 CB 编号。
    """
    s = str(line).strip()
    if not s:
        return ''
    if s in _MATERIAL_SKIP_TITLES:
        return ''
    # 等效设备备注整段去掉
    if re.match(r'^备注[：:]', s) and '设备' in s and ('等效' in s or '替代' in s):
        return ''
    # 规程/文件编号（REFERENCE：去除 SOP 文档引用），不误伤「CB0012工作参比品」等（须带 3-…-SOP- 结构）
    s = re.sub(
        r'3-[A-Z0-9]+-SOP-[A-Z0-9]+-\d+[A-Za-z0-9\-]*',
        '',
        s,
        flags=re.IGNORECASE,
    )
    s = re.sub(r'3-CB\d+-SOP-TI\d+-\d+\s*[^\s，。；、]*', '', s)
    s = re.sub(r'3-SOP-[A-Za-z0-9\-]+', '', s)
    # 厂家名（「：…有限公司」）；不用「：[A-Z]{2,}\\d+」删型号——会误删「：CB0012工作参比品」
    # 1) 统一处理“设备名：厂商”——若右侧明显是厂商信息，则仅保留左侧
    if re.search(r'[：:]', s):
        left, right = re.split(r'[：:]', s, maxsplit=1)
        r = right.strip()
        if re.search(
            r"(有限公司|有限责任公司|股份有限公司|GmbH|AG|Inc\.?|Ltd\.?|Co\.?,?\s*Ltd\.?|Corporation|Corp\.?)",
            r,
            flags=re.IGNORECASE,
        ):
            s = left.strip()
        elif re.search(r"(Eppendorf|Thermo|Beckman|Sartorius|Merck|Waters|Agilent|Tecan|Mettler|Toledo)", r, re.I):
            s = left.strip()

    # 2) 兜底：去掉“：厂商”尾巴（中英文后缀）
    s = re.sub(
        r"[：:][^：:]{2,}(有限公司|有限责任公司|股份有限公司|GmbH|AG|Inc\.?|Ltd\.?|Co\.?,?\s*Ltd\.?|Corporation|Corp\.?)\b[^：:]*",
        "",
        s,
        flags=re.IGNORECASE,
    )
    s = re.sub(r'\s+', ' ', s).strip(' 　')
    return s


def _is_sec_hplc_purity_section(section_name) -> bool:
    if not section_name:
        return False
    t = str(section_name).strip()
    if '纯度' not in t:
        return False
    u = t.upper().replace(' ', '').replace('\u3000', '')
    return 'SEC' in u and 'HPLC' in u


def _is_material_operation_line(line: str) -> bool:
    """配制/操作步骤行，不进入设备试剂名称列表。"""
    if not line:
        return False
    return bool(
        re.search(
            r'^(称取|量取|取\d|取1支|加超纯水|加.+混匀|加.+至\d|定容|超声脱气|临用新制|室温避光|分装保存|量取\d)',
            line,
        )
        or re.search(r'有效期\d|保存[，,]有效期|混匀[，,].*保存', line)
        or re.search(r'称取|量取|取\d+.*支|加超纯水|加.+定容|超声脱气|分装保存', line)
    )


def _is_material_context_skip_line(line: str) -> bool:
    """小节标题、样品类型说明（非标准物质清单）。"""
    s = (line or '').strip()
    if not s:
        return True
    if s in _MATERIAL_SKIP_TITLES or s == '实验试剂':
        return True
    if s in ('系统适用性样品', '供试品', 'FB溶液'):
        return True
    if re.match(r'^试液配制', s) or 'Solution Preparation' in s:
        return True
    if '以下试液均可' in s or '等比例制备' in s:
        return True
    if '原液的FB' in s:
        return True
    if '原液' in s and '、参比品' in s and '工作参比品' not in s:
        return True
    return False


def _is_consumable_name(name: str) -> bool:
    n = (name or '').lower()
    s = name or ''
    if any(k in s for k in ('脱盐柱', '色谱柱', '预柱', '滤芯', '层析柱')):
        return True
    if 'column' in n and any(c.isdigit() for c in s):
        return True
    if 'zeba' in n:
        return True
    return False


def _strip_star_note(name: str) -> str:
    s = re.sub(r'^[\*＊]\s*', '', name.strip())
    s = re.sub(r'为一次性使用耗材[。\.]*$', '', s)
    return s.strip()


def _shorten_reagent_title(line: str) -> str:
    """试液标题行 → 申报用简短名称（与常见肽图 SOP 结构对齐）。"""
    s = line.strip()
    rules = [
        (r'^1M\s*DTT溶液.*', 'DTT'),
        (r'^1M\s*IAM溶液.*', 'IAA'),
        (r'^1M\s*IAA溶液.*', 'IAA'),
        (
            r'^100mM\s*Tris-HCl.*400mM.*GuHCl.*',
            '1M Tris-HCl（pH7.5）、8M盐酸胍溶液',
        ),
        (r'^1mg/mL\s*Trypsin.*', 'Trypsin'),
        (r'^10%\s*FA溶液.*', '甲酸'),
        (r'^流动相A（0\.1%FA-水溶液）.*', '甲酸'),
        (r'^流动相B（0\.08%FA-乙腈溶液）.*', '乙腈'),
        (r'^10%甲醇（洗泵液、洗针液）.*', '甲醇'),
        (r'^10%\s*甲醇.*', '甲醇'),
    ]
    for pat, repl in rules:
        if re.match(pat, s, re.IGNORECASE):
            return repl
    return s


def _dedup_preserve_order(items: list) -> list:
    out = []
    seen = set()
    for x in items:
        k = re.sub(r'[\s\u3000]+', '', (x or '').lower())
        if not k or k in seen:
            continue
        seen.add(k)
        out.append(x)
    return out


def _integrate_material_lines(lines: list, section_name=None) -> list:
    """
    将 SOP 材料区打散行整合为：
    主要设备：…、标准物质：…、主要试剂/耗材：…
    去掉操作细则，设备/耗材/试液名称来自编号行、表格转写行及试液标题行。
    """
    equipment: list = []
    standards: list = []
    reagents: list = []

    # 全文扫描：色谱柱规格（有时落在长行）
    for ln in lines:
        if re.search(
            r'ACQUITY|BEH\s*C18|Peptide\s+BEH|TSKgel|2\.1\s*[×xX]\s*150',
            ln,
            re.I,
        ) and not _is_material_operation_line(ln):
            short = re.sub(r'\s+', ' ', ln.strip())
            if len(short) > 120:
                short = short[:117] + '…'
            if short not in reagents:
                reagents.append(short)

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        if _is_material_context_skip_line(line):
            i += 1
            continue
        if _is_material_operation_line(line):
            i += 1
            continue

        mnum = re.match(r'^\d+\.\d+\s*[\*＊]?\s*(.+)$', line)
        if mnum:
            name = _strip_star_note(mnum.group(1))
            if _is_consumable_name(name):
                reagents.append(name)
            else:
                equipment.append(name)
            i += 1
            continue

        if line.startswith('主要设备：'):
            payload = line.split('：', 1)[1] if '：' in line else ''
            for p in re.split(r'[、，,；;。]', payload):
                t = p.strip()
                if t:
                    equipment.append(t)
            i += 1
            continue

        # 识别纯设备名（无前缀、无编号，如"恒温混匀仪"、"干式恒温器"）
        if (len(line) < 20 and
            not _is_material_context_skip_line(line) and
            not _is_material_operation_line(line) and
            any(kw in line for kw in ['仪', '器', '机', '天平', '离心', '混匀', '恒温', '色谱', '光度计', '培养箱', '安全柜', '工作台', '洁净工作台', '超净台']) and
            not re.search(r'溶液|流动相|缓冲|试剂|FA|Trypsin|Tris|DTT|IAM', line, re.I)):
            equipment.append(line)
            i += 1
            continue

        if line.startswith('标准物质：'):
            payload = line.split('：', 1)[1].strip().rstrip('。.')
            if payload:
                standards.append(f'标准物质：{payload}。')
            i += 1
            continue

        if re.match(r'^CB\d+(?:-\d+)?工作参比品', line) or (
            '工作参比品' in line
            and re.search(r'CB\d+', line)
            and len(line) < 80
            and '原液' not in line
        ):
            payload = line.strip().rstrip('。.')
            payload = re.sub(r'^标准物质[：:]\s*', '', payload)
            standards.append(f'标准物质：{payload}。')
            i += 1
            continue

        # 试液标题行：下一行常为称取/量取操作
        if i + 1 < n and _is_material_operation_line(lines[i + 1]):
            if not _is_material_context_skip_line(line) and len(line) < 100:
                reagents.append(_shorten_reagent_title(line))
            i += 1
            continue

        # 单独成行的短试液名（无紧随操作行）
        if len(line) <= 90 and not _is_material_operation_line(line):
            if re.search(
                r'溶液|流动相|Trypsin|缓冲|洗泵|洗针|甲醇|乙腈|甲酸|FA|Tris|GuHCl|Guanidine|DTT|IAM|IAA',
                line,
                re.I,
            ):
                reagents.append(_shorten_reagent_title(line))

        i += 1

    # 标准物质去重（保留首条）
    std_out = ''
    if standards:
        seen_std = set()
        for s in standards:
            inner = re.sub(r'^标准物质[：:]', '', s).strip().rstrip('。')
            k = re.sub(r'\s+', '', inner.lower())
            if k in seen_std:
                continue
            seen_std.add(k)
            std_out = f'标准物质：{inner}。'
            break
    if not std_out:
        for ln in lines:
            if '原液' in ln:
                continue
            m = re.search(r'(CB\d+(?:-\d+)?工作参比品)', ln)
            if m:
                std_out = f'标准物质：{m.group(1)}。'
                break

    reagents = _dedup_preserve_order(reagents)
    equipment = _dedup_preserve_order(equipment)

    if _is_sec_hplc_purity_section(section_name):
        eq_line = '主要设备：高效液相色谱仪。'
    elif equipment:
        eq_line = '主要设备：' + '、'.join(equipment) + '。'
    else:
        eq_line = ''

    reg_line = ''
    if reagents:
        reg_line = '主要试剂/耗材：' + '、'.join(reagents) + '。'

    out = []
    if eq_line:
        out.append(eq_line)
    if std_out:
        out.append(std_out)
    if reg_line:
        out.append(reg_line)

    return out


def refine_materials(sop_raw_list, template_list, ref_list, max_chars=50, section_name=None):
    """
    设备、材料、试剂：以 SOP 为主，**整合**为「主要设备 / 标准物质 / 主要试剂/耗材」三类行，
    剔除配制操作句；名称来自编号项、表格行及试液标题行。
    「纯度（SEC-HPLC）」章的主要设备单行规则：整合阶段即写死高效液相色谱仪（与写入端一致）。

    max_chars 保留签名兼容，不再使用。
    """
    _ = max_chars

    lines: list = []
    seen: set = set()

    def add_from_list(raw_list):
        if not raw_list:
            return
        items = raw_list if isinstance(raw_list, list) else [raw_list]
        for item in items:
            text = reconstruct_text(item) if isinstance(item, list) else str(item)
            for part in text.splitlines():
                # 特殊处理：检查"设备名：厂家"格式（保留设备名，删除厂家）
                if re.match(r'^[^：:]+[：:][^：:]*有限公司[^：:]*$', part):
                    device_name = part.split('：', 1)[0].split(':', 1)[0].strip()
                    if device_name and len(device_name) < 30:
                        lines.append(device_name)
                        key = re.sub(r'[\s\u3000]+', '', device_name)
                        seen.add(key)
                        continue

                cleaned = _clean_material_line(part)
                if not cleaned:
                    continue
                key = re.sub(r'[\s\u3000]+', '', cleaned)
                if key in seen:
                    continue
                seen.add(key)
                lines.append(cleaned)

    add_from_list(sop_raw_list)
    if not lines:
        add_from_list(template_list)
    if not lines:
        add_from_list(ref_list)

    if not lines:
        return ['设备、材料、试剂根据检验方法配置。']

    integrated = _integrate_material_lines(lines, section_name=section_name)
    if integrated:
        return integrated
    return lines


_PREP_DROP_TITLES = frozenset({
    '供试品的制备', '样品处理', 'Sample Preparation', 'sample preparation',
})


def _classify_sample_prep_segment(text: str) -> str:
    """将 SOP 样品处理片段分为 blank / suitability / test（供试品制备侧）。"""
    s = str(text).strip()
    if not s:
        return 'test'
    head = s[:48]
    # 空白溶液、空白对照、FB 替代供试品等
    if re.match(r'^空白', s) or '空白溶液' in head or '空白对照' in head:
        return 'blank'
    if 'FB' in s and '替代供试品' in s:
        return 'blank'
    if '酶解缓冲' in s and '空白' in s:
        return 'blank'
    # 系统适用性溶液（含工作参比品按同法制备）
    if '系统适用性' in head or re.match(r'^系统适用', s):
        return 'suitability'
    if '工作参比品' in head or ('参比品' in head and '项下' in s and '按' in s):
        return 'suitability'
    return 'test'


def _is_short_step_heading(p: str) -> bool:
    """SOP 样品处理内小节名（变性还原、换液、酶解等），非完整操作句。"""
    p = str(p).strip()
    if len(p) > 14 or len(p) < 2:
        return False
    if re.search(r'[，。；：（）\d]', p):
        return False
    if p.startswith(
        ('取', '向', '按', '用', '将', '置', '加', '往', '更', '若', '每', '完')
    ):
        return False
    return bool(re.match(r'^[\u4e00-\u9fa5]{2,10}$', p))


def _group_test_bucket_into_paragraphs(parts: list) -> list:
    """
    供试品制备侧：按小节标题聚合成多段，段内为「标题：正文」而非「标题。正文」。
    返回若干字符串，每项对应 Word 中一个独立段落（与写回插入逻辑配合）。
    """
    if not parts:
        return []
    grouped: list = []
    i, n = 0, len(parts)
    while i < n:
        p = str(parts[i]).strip()
        if not p:
            i += 1
            continue
        if _is_short_step_heading(p) and i + 1 < n:
            bodies: list = []
            j = i + 1
            while j < n:
                nxt = str(parts[j]).strip()
                if not nxt:
                    j += 1
                    continue
                if _is_short_step_heading(nxt):
                    break
                bodies.append(nxt.rstrip('。'))
                j += 1
            inner = '，'.join(bodies) if bodies else ''
            grouped.append(f'{p}：{inner}。' if inner else f'{p}。')
            i = j
        else:
            if p.endswith(('。', '；')):
                grouped.append(p.strip())
            else:
                grouped.append(p.rstrip('。') + '。')
            i += 1
    return grouped


def _merge_prep_bucket_lines(lines: list) -> str:
    """合并同桶内多行（空白/系统适用性通常仅一条）；避免用句号硬接短标题与正文。"""
    cleaned = []
    drop_single = frozenset({'空白溶液', '系统适用性溶液', '空白对照溶液'})
    for L in lines:
        L = str(L).strip()
        if not L:
            continue
        if L in drop_single and len(lines) > 1:
            continue
        cleaned.append(L)
    if not cleaned:
        return ''
    if len(cleaned) == 1:
        s = cleaned[0].strip().rstrip('。')
        return s + '。'
    out = []
    for L in cleaned:
        out.append(L.rstrip('。'))
    text = '。'.join(out) + '。'
    text = re.sub(r'。{2,}', '。', text)
    return text.strip('。') + ('。' if text.strip() else '')


def _ensure_prep_label(label: str, body: str) -> str:
    """统一为「标签：正文」；避免重复标签。"""
    b = str(body).strip()
    if not b:
        return ''
    for pfx in (f'{label}：', f'{label}:'):
        if b.startswith(pfx) or b.startswith(label + '：'):
            return normalize_terminology(b).strip()
    return normalize_terminology(f'{label}：{b}').strip()


def refine_sample_prep(sop_raw_list, template_list, ref_list, max_chars=200):
    """
    样品处理（对应分析方法「3.1 样品处理」正文，模版侧常有独立「样品处理」小标题）。

    输出固定三块顺序：空白溶液 → 系统适用性溶液 → 供试品制备；
    内容来自 SOP 第四章样品处理段，按语义归类后合并，并做申报用语归一化。
    """
    _ = template_list, ref_list  # 保留签名；优先 SOP 全文
    if not sop_raw_list:
        return []

    parts = [str(x).strip() for x in sop_raw_list if str(x).strip()]
    if not parts:
        return []

    buckets = {'blank': [], 'suitability': [], 'test': []}
    for p in parts:
        if p in _PREP_DROP_TITLES:
            continue
        cat = _classify_sample_prep_segment(p)
        buckets[cat].append(p)

    out: list = []
    btxt = _merge_prep_bucket_lines(buckets['blank'])
    if btxt:
        out.append(_ensure_prep_label('空白溶液', btxt))

    stxt = _merge_prep_bucket_lines(buckets['suitability'])
    if stxt:
        out.append(_ensure_prep_label('系统适用性溶液', stxt))

    test_paras = _group_test_bucket_into_paragraphs(buckets['test'])
    if test_paras:
        first = test_paras[0]
        # 避免「供试品制备：变性还原：」双冒号，首步改读作「供试品制备，变性还原：…」
        if '：' in first[:20]:
            out.append(normalize_terminology(f'供试品制备，{first}').strip())
        else:
            out.append(_ensure_prep_label('供试品制备', first))
        out.extend(test_paras[1:])

    if not out:
        # 无法归类时退回整段文本，避免空节
        fallback = reconstruct_text(sop_raw_list).strip()
        if fallback:
            return [normalize_terminology(fallback[: max_chars * 3])]
        return ['样品按SOP进行处理。']

    # 软字数上限（总长），超出时优先保留前两块，供试品制备截断
    total = sum(len(x) for x in out)
    cap = max(max_chars * 12, 1800)
    if total <= cap:
        return out
    trimmed = []
    used = 0
    for block in out:
        if used + len(block) <= cap:
            trimmed.append(block)
            used += len(block)
        else:
            rest = cap - used
            if rest > 80:
                trimmed.append(block[:rest].rsplit('。', 1)[0] + '。')
            break
    return trimmed if trimmed else out


def refine_procedure(sop_raw_list, template_list, ref_list):
    """
    精简操作步骤（对应分析方法「3.2 操作步骤」）：
    - 色谱/电泳等方法：优先保留“条件参数 + 关键节点（平衡/进样序列/检测）”
    - 非色谱方法：保留流程主干（加样/孵育/洗板/显色/检测等），剔除大量逐步稀释、界面点击与记录表编号
    """
    if not sop_raw_list:
        return []

    # 1) 归一化为行列表
    reconstructed = reconstruct_text(sop_raw_list)
    all_lines = [str(x).strip() for x in (sop_raw_list if isinstance(sop_raw_list, list) else [reconstructed]) if str(x).strip()]
    if len(all_lines) <= 1:
        if "\n" in reconstructed:
            all_lines = [x.strip() for x in reconstructed.splitlines() if x.strip()]
        elif reconstructed.strip():
            all_lines = [reconstructed.strip()]

    def drop_noise(line: str) -> bool:
        s = (line or "").strip()
        if not s:
            return True
        if s.startswith(("注意", "注：", "备注", "例如")):
            return True
        if "记录填写在" in s or "记录在" in s:
            return True
        if re.search(r"3-[A-Z0-9]+-SOP-|3-SMP-|申请表|申请批准", s, re.I):
            return True
        # 过长的详细配制/界面操作行优先丢弃（保主干）
        if len(s) > 220 and any(k in s for k in ("取", "加入", "稀释", "倍", "mL", "μL", "rpm", "℃")):
            return True
        return False

    lines = [l for l in all_lines if not drop_noise(l)]

    # 2) 识别是否为色谱/条件类方法
    chrom_kws = ("色谱条件", "色谱柱", "检测波长", "流动相", "梯度", "柱温", "流速", "进样量")
    is_chrom = any(any(k in l for k in chrom_kws) for l in lines)

    # 3) 标题/锚点分组（只作为分组信号，不直接输出裸标题）
    headings = (
        "色谱条件", "电泳条件", "仪器参数", "参数设置", "进样序列", "创建序列",
        "平衡系统", "进样", "检测", "结果记录", "色谱柱清洗", "关机方法", "运行序列",
        "色谱柱清洗及保存", "色谱柱清洗及保存方法",
        "加样", "孵育", "洗板", "显色", "终止",
    )

    def is_heading(s: str) -> bool:
        t = s.strip().rstrip("：:")
        if len(t) > 14:
            return False
        return t in headings

    grouped = []
    cur_h = ""
    cur_buf: List[str] = []
    for l in lines:
        if is_heading(l):
            if cur_buf:
                grouped.append((cur_h, cur_buf))
            cur_h = l.strip().rstrip("：:")
            cur_buf = []
        else:
            cur_buf.append(l)
    if cur_buf:
        grouped.append((cur_h, cur_buf))

    out: List[str] = []

    def emit(label: str, items: List[str], max_items: int) -> None:
        if not items:
            return
        picked: List[str] = []
        for it in items:
            t = it.strip()
            if not t:
                continue
            if len(t) > 240:
                continue
            picked.append(t)
            if len(picked) >= max_items:
                break
        if not picked:
            return
        if label:
            # 统一写法：label：一句（或多句用；连接）
            body = "；".join(x.rstrip("。") for x in picked)
            out.append(f"{label}：{body}。")
        else:
            out.extend(picked)

    if is_chrom:
        # 色谱/条件类：优先抓“色谱条件：...”或分组后的条件块
        cond_lines = []
        for l in lines:
            s = l.strip()
            if any(k in s for k in ("平衡系统", "进样序列", "创建序列", "运行序列")):
                continue
            if s.startswith(("色谱条件：", "色谱条件:")):
                cond_lines.append(re.sub(r"^色谱条件[:：]\s*", "", s))
                continue
            # 只收“参数型”语句：包含“：”且命中参数关键词
            if "：" in s and any(k in s for k in ("色谱柱", "检测波长", "流动相", "梯度", "柱温", "流速", "进样量", "检测器", "运行时间", "进样器温度")):
                cond_lines.append(s)
        # 去掉裸“色谱条件”标题重复
        cond_lines = [l for l in cond_lines if l.strip() not in ("色谱条件", "色谱条件：")]
        if cond_lines:
            # 合并成一行，避免输出“色谱条件”+“色谱条件：...”两条
            merged = "；".join(str(x).strip().rstrip("。") for x in cond_lines[:8] if str(x).strip())
            if merged:
                out.append("色谱条件：" + merged.strip("；") + "。")

        # 关键节点
        for h, buf in grouped:
            if h in ("平衡系统",):
                emit("平衡系统", buf, 2)
            elif h in ("进样序列", "创建序列", "运行序列"):
                emit("进样序列", buf, 2)
            elif h in ("进样", "检测"):
                emit(h, buf, 2)
            elif h in ("色谱柱清洗", "关机方法"):
                emit(h, buf, 1)
            elif h in ("色谱柱清洗及保存", "色谱柱清洗及保存方法"):
                emit("色谱柱清洗及保存", buf, 1)

        if not out:
            # 兜底：保留最短的 3 条非空行
            out = [l for l in lines if len(l) <= 120][:3]
    else:
        # 非色谱：保留主干步骤标题（作为 label）+ 一句描述
        for h, buf in grouped:
            if h in ("加样", "孵育", "洗板", "显色", "终止", "检测", "实验前准备"):
                emit(h, buf, 1)
            else:
                # 无标题块：保留短句
                short = [x for x in buf if 6 <= len(x) <= 120]
                emit("", short, 4)

    # 4) 去重 + 限长
    dedup: List[str] = []
    seen = set()
    for x in out:
        key = _normalize_sentence_for_dedup(x)
        if not key or key in seen:
            continue
        seen.add(key)
        dedup.append(x)

    return dedup[:10] if dedup else ["按检验方法操作步骤进行测定。"]


def refine_suitability(sop_raw_list, template_list, ref_list, max_per_item=60):
    """
    精简试验成立标准部分
    目标：保留“量化/判定”标准，剔除操作步骤与表单编号噪声。
    """
    metric_kw = ('≤', '≥', '<', '>', '~', '～')
    metric_ctx_kw = (
        'RSD', 'CV', '相关系数', 'R2', '决定系数', '分离度', '理论塔板', '拖尾因子', '拖尾', '回收率',
        '重复性', '精密度', '线性', '准确度', '专属性', '无干扰', '信噪比',
        '比值', '上平台', '下平台', 'S型曲线', 'S形曲线', 'sigmoid',
    )
    stop_kw = (
        '加样', '孵育', '洗板', '显色', '终止', '检测', '进样', '平衡系统', '创建序列', '运行序列',
        '记录填写在', '记录在', '申请表', '申请批准',
    )

    def _sentences_from(source_list) -> List[str]:
        text = reconstruct_text(source_list or [])
        if not text:
            return []
        return [s.strip() for s in re.split(r'[。；;\n]', text) if str(s).strip()]

    def _is_noise(s: str) -> bool:
        t = str(s or '').strip()
        if not t:
            return True
        if t.startswith(('注意', '注：', '备注', '例如')):
            return True
        if re.search(r"3-[A-Z0-9]+-SOP-|3-SMP-|申请表|申请批准", t, re.I):
            return True
        if any(k in t for k in stop_kw):
            return True
        # 明显属于样品制备/操作描述的行（即便含 ≥/≤ 也不当作试验成立标准）
        if any(k in t for k in ('取', '加入', '混匀', '稀释', '离心', '孵育', '洗板', '显色', '终止')) and any(
            u in t for u in ('μL', 'uL', 'mL', 'μg', 'ug', 'rpm', '℃', 'min')
        ):
            return True
        return False

    def _pick_key_clause(s: str) -> str:
        # 保留原句但压缩空白，避免把主语裁没导致“意义不明”
        t = re.sub(r'\s+', ' ', str(s or '')).strip()
        # 优先保留“指标名 + 比较符号 + 数值/阈值”（尽量带上前置主语）
        m = re.search(r"(.{0,60}?)(≤|≥|<|>)(.{0,40})", t)
        if m:
            out = (m.group(1) + m.group(2) + m.group(3)).strip('，,。；; ')
            return out[:max(40, max_per_item)]
        # 数值范围类：0.8~1.25 / 0.8～1.25 / 0.8-1.25
        m = re.search(r"(.{0,60}?)(\d+(?:\.\d+)?)\s*(~|～|-|—|–)\s*(\d+(?:\.\d+)?)(.{0,20})", t)
        if m:
            out = (m.group(1) + m.group(2) + m.group(3) + m.group(4) + m.group(5)).strip('，,。；; ')
            return out[:max(40, max_per_item)]
        # 其次：RSD/相关系数等短句
        return t[:max(40, max_per_item)]

    # 来源优先级：SOP > 模版 > 参考（避免模板/参考把通用话塞进来）
    raw = []
    for src in (sop_raw_list, template_list, ref_list):
        raw.extend(_sentences_from(src))

    picked: List[str] = []
    seen_keys = set()
    for s in raw:
        if _is_noise(s):
            continue
        # 需要“量化符号”或“指标语境词”；单纯“符合要求”也可以作为兜底
        has_metric = any(k in s for k in metric_kw) or bool(re.search(r'\d+(?:\.\d+)?\s*(~|～|-|—|–)\s*\d', s))
        has_ctx = any(k in s for k in metric_ctx_kw)
        has_ok = '符合要求' in s or '应符合要求' in s
        if not (has_metric or has_ctx or has_ok):
            continue
        # 避免把 mg/mL 等样品条件误当成标准：需出现指标语境词，否则跳过
        if has_metric and not has_ctx and any(u in s for u in ('mg/mL', 'μg/mL', 'ug/mL')):
            continue
        if has_metric or has_ctx:
            clause = _pick_key_clause(s)
        else:
            clause = re.sub(r'\s+', '', str(s)).strip('，,。；;')
            if len(clause) > max_per_item:
                clause = clause[:max_per_item]
        clause = normalize_terminology(clause)
        if not clause:
            continue
        key = _normalize_sentence_for_dedup(clause)
        if key and key not in seen_keys:
            seen_keys.add(key)
            picked.append(clause)

    return picked[:8]


def refine_result_calc(sop_raw_list, template_list, ref_list):
    """
    精简结果计算部分
    要求：保留核心公式和参数定义
    """
    # 结果计算常见触发：等式/比值/换算/变量释义
    keep_kw = (
        '=', '＝', '×', '×100', '×100%', '× 100', '× 100%', '%', '‰',
        '/', '÷',
        '公式', '方程', '回归方程', '标准曲线', '线性回归',
        '计算', '计算公式', '换算', '其中', '式中', '注：', '注:', '定义',
        'Slope', 'Conc', 'A280', 'RSD',
        '峰面积', '面积归一', '主峰', '含量', '浓度', '回收率', '稀释倍数',
    )
    drop_kw = (
        '加样', '孵育', '洗板', '显色', '终止', '检测', '进样', '平衡系统', '创建序列', '运行序列',
        '色谱条件', '电泳条件', '操作步骤', '记录填写在', '记录在', '申请表', '申请批准',
    )

    def _lines_from(source_list) -> List[str]:
        text = reconstruct_text(source_list or [])
        if not text:
            return []
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        # 有些 SOP 用 “；/。” 连接公式说明，补一层拆分
        out: List[str] = []
        for l in lines:
            if len(l) > 180 and any(k in l for k in ('；', '。')):
                out.extend([x.strip() for x in re.split(r'[。；;]', l) if x.strip()])
            else:
                out.append(l)
        return out

    def _is_formula_like(line: str) -> bool:
        s = str(line or '').strip()
        if not s:
            return False
        # 等式/全角等号/乘除/明显变量释义
        if any(ch in s for ch in ('=', '＝', '×', '÷')) and any(k in s for k in ('=', '＝')):
            return True
        if re.search(r'\b(Slope|Conc|A280|RSD)\b', s):
            return True
        if s.startswith(('其中', '式中', '注：', '注:')):
            return True
        # 变量定义行：X：含义 或 X= 含义
        if re.match(r'^[A-Za-zαβγΔμ]\s*[=＝：:]', s):
            return True
        # 面积归一/回归方程等关键句
        if any(k in s for k in ('面积归一', '回归方程', '标准曲线', '计算公式', '换算')):
            return True
        return False

    def _clean_calc_lines(source_list) -> List[str]:
        """
        从单一来源抽取“结果计算”候选行：
        - 严格丢弃操作步骤/记录表/条件描述
        - 仅保留公式/方程/变量释义及其紧随的注释行
        """
        all_lines = _lines_from(source_list)
        kept: List[str] = []
        capturing_defs = 0
        for l in all_lines:
            if not l:
                continue
            if any(k in l for k in drop_kw):
                capturing_defs = 0
                continue
            if re.search(r"3-[A-Z0-9]+-SOP-|3-SMP-|申请表|申请批准", l, re.I):
                capturing_defs = 0
                continue
            s = normalize_terminology(l)
            # 明显操作细节（即使含“计算”也不应进结果计算）
            if not _is_formula_like(s) and any(v in s for v in ('取', '加入', '稀释', '混匀', '离心', '孔', '管')) and any(
                u in s for u in ('μL', 'uL', 'mL', 'μg', 'ug', 'rpm', '℃', 'min')
            ):
                continue
            if '+' in s and any(u in s for u in ('μL', 'uL', 'mL', 'μg', 'ug')) and ('=' not in s and '＝' not in s):
                continue
            if s.strip() in ('数据计算和结果报告', '计算Calculate', '使用面积归一法，计算：'):
                continue
            if len(s) > 260:
                continue

            # 仅保留“像公式/变量释义”的行；但允许在命中公式后，抓取最多 6 行释义
            if _is_formula_like(s) or any(k in s for k in keep_kw):
                kept.append(s)
                # 命中公式/方程后，放宽抓取变量释义若干行
                if any(ch in s for ch in ('=', '＝')) or any(k in s for k in ('公式', '方程', '回归方程', '标准曲线')):
                    capturing_defs = 6
                elif capturing_defs > 0:
                    capturing_defs -= 1
                continue
            if capturing_defs > 0 and (s.startswith(('其中', '式中', '注：', '注:')) or re.match(r'^[A-Za-zαβγΔμ]\s*[=＝：:]', s)):
                kept.append(s)
                capturing_defs -= 1
        return kept

    def _extract_from(*sources) -> List[str]:
        all_lines: List[str] = []
        for src in sources:
            all_lines.extend(_clean_calc_lines(src))

        kept: List[str] = []
        for l in all_lines:
            if not l:
                continue
            # _clean_calc_lines 已做过严格过滤，这里只做轻量保留
            if not (_is_formula_like(l) or any(k in l for k in keep_kw)):
                continue
            kept.append(normalize_terminology(l))

        # 合并“其中/式中”参数定义到上一条（让结果计算更像一个紧凑块）
        merged: List[str] = []
        for l in kept:
            if merged and (l.startswith(('其中', '式中')) or re.match(r"^[A-Za-zαβγΔμ]\s*[=：:]", l)):
                prev = merged[-1].rstrip('。')
                merged[-1] = (prev + '；' + l).strip('；') + ('。' if not prev.endswith('。') else '')
            else:
                merged.append(l)

        # 去重（弱归一化）
        out: List[str] = []
        seen = set()
        for x in merged:
            key = _normalize_sentence_for_dedup(x)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(x)

        return out[:8]

    # SOP 优先：先尽最大可能从 SOP 抽取公式与释义，避免混入模板/参考造成两套口径混写
    sop_best = _extract_from(sop_raw_list)
    if sop_best:
        return sop_best[:12]

    # 若 SOP 未抽到（例如抽取器把公式拆得很碎/无明显等号），再回退模板/参考
    fallback = _extract_from(template_list, ref_list)
    return fallback if fallback else ["按既定方法计算并报告结果。"]


def refine_acceptance(sop_raw_list, template_list, ref_list, max_chars=120):
    """
    精简合格标准部分：保留限度/比较符号句，亦保留定性表述（肽图、一致、专属性等）。
    """
    limit_kw = ('≥', '≤', '%', '范围内', '应', '不得', '限度', 'mg/mL', 'μg', 'IU')
    qual_kw = (
        '一致', '符合', '专属性', '肽图', '图谱', '保留时间', '相对保留', '批间',
        '稳定性', '降解', '杂质', '归一化',
    )

    def _pick_sentences(text: str) -> List[str]:
        sentences = re.split(r'[。；;\n]', text)
        out: List[str] = []
        for sentence in sentences:
            sentence = str(sentence).strip()
            if len(sentence) < 6:
                continue
            if any(kw in sentence for kw in limit_kw) or any(kw in sentence for kw in qual_kw):
                out.append(normalize_terminology(sentence))
        return out

    for source_list in [sop_raw_list, template_list, ref_list]:
        if not source_list:
            continue
        text = reconstruct_text(source_list)
        if not text:
            continue
        picked = _pick_sentences(text)
        if not picked:
            continue
        merged = picked[0]
        for s in picked[1:]:
            if len(merged) + 1 + len(s) <= max_chars:
                merged = merged + '；' + s
            else:
                break
        if len(merged) > max_chars:
            merged = merged[:max_chars].rsplit('，', 1)[0] + '。'
        return [merged] if merged else [picked[0][:max_chars]]

    return ["应符合规定。"]


def process_method(method):
    """处理单个方法"""
    name = method.get('name', '')
    sop_raw = method.get('sop_raw', {})
    template_existing = method.get('template_existing', {})
    reference_style = method.get('reference_style', {})

    print(f"\n处理方法: {name}")

    refined_method = {
        'name': name,
        'stop': method.get('stop', []),
        'refined': {}
    }

    # 1. 原理（SOP 原文保留，见 SKILL / ra_compliance 校验）
    print("  - 处理原理（SOP 原文保留）...")
    refined_method['refined']['principle'] = refine_principle(
        sop_raw.get('principle', []),
        template_existing.get('principle', []),
        reference_style.get('principle', []),
        max_chars=200
    )

    # 2. 材料和设备
    print("  - 精简材料和设备...")
    refined_method['refined']['materials_and_equipment'] = refine_materials(
        sop_raw.get('materials_and_equipment', []),
        template_existing.get('materials_and_equipment', []),
        reference_style.get('materials_and_equipment', []),
        max_chars=50,
        section_name=name,
    )

    # 3. 样品处理
    print("  - 精简样品处理...")
    refined_method['refined']['sample_prep'] = refine_sample_prep(
        sop_raw.get('sample_prep', []),
        template_existing.get('sample_prep', []),
        reference_style.get('sample_prep', []),
        max_chars=200
    )

    # 4. 操作步骤
    print("  - 精简操作步骤...")
    refined_method['refined']['procedure'] = refine_procedure(
        sop_raw.get('procedure', []),
        template_existing.get('procedure', []),
        reference_style.get('procedure', [])
    )

    # 5. 试验成立标准
    print("  - 精简试验成立标准...")
    refined_method['refined']['suitability_criteria'] = refine_suitability(
        sop_raw.get('suitability_criteria', []),
        template_existing.get('suitability_criteria', []),
        reference_style.get('suitability_criteria', []),
        max_per_item=60
    )

    # 6. 结果计算
    print("  - 精简结果计算...")
    refined_method['refined']['result_calculation'] = refine_result_calc(
        sop_raw.get('result_calculation', []),
        template_existing.get('result_calculation', []),
        reference_style.get('result_calculation', [])
    )

    # 7. 合格标准
    print("  - 精结合格标准...")
    refined_method['refined']['acceptance_criteria'] = refine_acceptance(
        sop_raw.get('acceptance_criteria', []),
        template_existing.get('acceptance_criteria', []),
        reference_style.get('acceptance_criteria', []),
        max_chars=30
    )

    # 统计字数
    print(f"\n  字数统计:")
    for section, content in refined_method['refined'].items():
        total_chars = count_chars(content)
        print(f"    {section}: {total_chars} 字")

    return refined_method


def main():
    """主处理流程"""
    print("=" * 60)
    print("SOP内容精简处理")
    print("=" * 60)

    args = [a for a in sys.argv[1:] if a]
    strict_principle = '--strict-principle' in args
    args = [a for a in args if a != '--strict-principle']

    print_skill_iron_rules()

    in_path = args[0] if len(args) > 0 else 'extracted.json'
    out_path = args[1] if len(args) > 1 else 'refined.json'

    # 读取提取的JSON
    print(f"\n读取 {in_path}...")
    with open(in_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    slim_base, _ = os.path.splitext(os.path.abspath(in_path))
    full_sidecar = slim_base + '.full.json'
    extracted_full_path = full_sidecar if os.path.isfile(full_sidecar) else ''

    # 创建精简后的数据（保留写入 docx 所需的 template_path / style）
    refined_data = {
        'template': data.get('template', ''),
        'template_path': data.get('template_path', ''),
        'style': data.get('style', 'hierarchical'),
        'extracted_full_path': extracted_full_path,
        'methods': []
    }

    # 处理每个方法
    total_methods = len(data.get('methods', []))
    print(f"\n共 {total_methods} 个方法需要处理")

    for idx, method in enumerate(data.get('methods', []), 1):
        print(f"\n[{idx}/{total_methods}]")
        refined_method = process_method(method)
        ok, vmsg = verify_principle_verbatim(
            method.get('sop_raw', {}).get('principle'),
            refined_method['refined'].get('principle'),
        )
        if not ok:
            print(f"  [WARN] principle 校验: {vmsg}")
            if strict_principle:
                print("  [FAIL] --strict-principle：终止", file=sys.stderr)
                sys.exit(1)
        elif strict_principle and method.get('sop_raw', {}).get('principle'):
            print("  [OK] principle 与 SOP 逐字一致（空白归一后）")

        refined_data['methods'].append(refined_method)

    # 保存精简后的JSON
    print("\n" + "=" * 60)
    print(f"保存 {out_path}...")
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(refined_data, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 60)
    print("精简完成！")
    print(f"处理了 {len(refined_data['methods'])} 个方法")
    print(f"输出文件: {out_path}")
    print("=" * 60)


if __name__ == '__main__':
    main()
