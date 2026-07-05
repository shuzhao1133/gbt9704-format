#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
format_fix.py — 政府报告 Word 格式一键修复（GB/T 9704 适用部分）

用法：
    python3 format_fix.py 文稿.docx -out 终稿.docx [--no-layout]

缺省行为：
  1. 套用机构排版规范（与 GB/T 9704 一致）：页边距上37/下35/左28/右26mm；
     题目方正小标宋简体二号不加粗居中；正文仿宋 3号（16pt）首行缩进2字符；
     一级标题黑体、二级楷体加粗、三级/四级仿宋不加粗，均3号；章级标题
     （第X章/前言等）黑体三号居中；标题行距固定33磅、正文28磅；数字英文
     Times New Roman；页码按 GB/T 9704 7.5 自动生成（宋体4号"— N —"奇右偶左）。
     加 --no-layout 则保留原版式，只修文字格式硬伤。
  2. 文字格式硬伤（正文+表格逐格）：删多余空格（保留英文单词间空格）；
     中文语境半角标点转全角；全角字母数字转半角；.../。。。→……；重复标点折叠；
     序号写法 "1、"→"1."、"（一）、"→"（一）"、"一，"→"一、"、"(1)"→"（1）"；
     短标题末尾句号删除。
  3. 只改格式不改内容；序号跳号/缺失只提示不改动。

输出：终稿 .docx + 终端分类计数汇总（无修改记录文件）。
"""
import argparse
import re
import sys
from collections import Counter
from copy import deepcopy

try:
    from docx import Document
    from docx.shared import Pt, Mm
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("需要 python-docx：pip3 install python-docx")

CJK = lambda c: '一' <= c <= '鿿' or c in '。，、；：？！（）《》【】""' "‘’—…·〔〕"
SPACES = ' 　\t '
CN_DIGITS = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}
CN_YEAR = {'〇': '0', '○': '0', '零': '0', '一': '1', '二': '2', '三': '3', '四': '4',
           '五': '5', '六': '6', '七': '7', '八': '8', '九': '9'}
HALF2FULL = {',': '，', ';': '；', ':': '：', '?': '？', '!': '！', '(': '（', ')': '）'}
FW_ALNUM = {chr(0xFF10 + i): chr(0x30 + i) for i in range(10)}
FW_ALNUM.update({chr(0xFF21 + i): chr(0x41 + i) for i in range(26)})
FW_ALNUM.update({chr(0xFF41 + i): chr(0x61 + i) for i in range(26)})

PAT_L1 = re.compile(r'^([一二三四五六七八九十]{1,3})([、，,\.])')
PAT_L2 = re.compile(r'^（([一二三四五六七八九十]{1,3})）([、，,\.。]?)')
PAT_L3 = re.compile(r'^(\d{1,2})([\.、，])')
PAT_L4 = re.compile(r'^[（(](\d{1,2})[）)]([、，\.]?)')


def cn2int(s):
    if s == '十':
        return 10
    v = 0
    if '十' in s:
        a, _, b = s.partition('十')
        v = (CN_DIGITS.get(a, 1) if a else 1) * 10 + (CN_DIGITS.get(b, 0) if b else 0)
    else:
        for ch in s:
            v = v * 10 + CN_DIGITS.get(ch, 0)
    return v


def iter_blocks(doc):
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield 'p', Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield 'tbl', Table(child, doc)


CHAP_RE = re.compile(r'^第[一二三四五六七八九十百]+[章篇]')
SPECIAL_HEADINGS = {'前言', '序言', '目录', '附录', '参考文献', '后记',
                    '引言', '结语', '结束语', '摘要'}


def is_chapter_heading(text):
    """第X章/第X篇，或前言/目录/附录等无序号篇章标题（黑体三号居中一层）。"""
    t = text.strip()
    if not t or len(t) > 30 or t.endswith(('。', '，', '；', '：', '、')):
        return False
    if CHAP_RE.match(t):
        return True
    return re.sub(r'\s+', '', t) in SPECIAL_HEADINGS


def transform_text(text, keep_inner_spaces=False):
    """返回 (新文本, Counter{规则: 次数}, [提示])。keep_inner_spaces=True 时保留段内
    空格（用于'第一章 总则''前  言'等篇章标题，其空格是标题内分隔而非多余空格）。"""
    n = len(text)
    out = list(text)
    stats = Counter()
    hints = []

    def ctx(i, d):
        j = i + d
        while 0 <= j < n and text[j] in SPACES:
            j += d
        return text[j] if 0 <= j < n else ''

    for i, c in enumerate(text):
        if c in FW_ALNUM:
            out[i] = FW_ALNUM[c]
            stats['全角字符转半角'] += 1

    for m in re.finditer(r'\.{3,}|。{3,}', text):
        s, e = m.span()
        out[s] = '……'
        for k in range(s + 1, e):
            out[k] = ''
        stats['省略号规范'] += 1

    # 2b. 汉字日期→阿拉伯（GB/T 9704 7.3.5.4 成文日期用阿拉伯数字标全）
    for m in re.finditer(r'[〇○零一二三四五六七八九]{2,4}年([一二三四五六七八九十]{1,3})月'
                         r'(?:([一二三四五六七八九十]{1,3})日)?', text):
        y = ''.join(CN_YEAR.get(c, '') for c in m.group(0).split('年')[0])
        if len(y) < 2:
            continue
        rep = f'{y}年{cn2int(m.group(1))}月'
        if m.group(2):
            rep += f'{cn2int(m.group(2))}日'
        out[m.start()] = rep
        for k in range(m.start() + 1, m.end()):
            out[k] = ''
        stats['汉字日期转阿拉伯'] += 1

    # 2c. 月日不编虚位：07月05日 → 7月5日（10月/20日 不受影响）
    for m in re.finditer(r'(?<![\d])0([1-9])\s*[月日]', text):
        if out[m.start()] == '0':
            out[m.start()] = ''
            stats['月日虚位删除'] += 1

    # 2d. 千分位逗号删除（GB/T 15835 分节用千分空不用逗号；仅中文语境段落）
    if any(CJK(c) for c in text):
        for m in re.finditer(r'(?<=\d),(?=\d{3}(?!\d))', text):
            if out[m.start()] == ',':
                out[m.start()] = ''
                stats['千分位逗号删除'] += 1

    for i, c in enumerate(text):
        if out[i] != c:
            continue
        if c in HALF2FULL:
            prev, nxt = ctx(i, -1), ctx(i, 1)
            if c == '(' and nxt and CJK(nxt):
                out[i] = '（'
            elif c == ')' and prev and CJK(prev):
                out[i] = '）'
            elif c in ',;:?!':
                if c == ',' and prev.isdigit() and nxt.isdigit():
                    continue
                if prev and CJK(prev):
                    out[i] = HALF2FULL[c]
            if out[i] != c:
                stats['半角标点转全角'] += 1
        elif c == '.':
            prev, nxt = ctx(i, -1), ctx(i, 1)
            if prev and CJK(prev) and (not nxt or CJK(nxt)):
                out[i] = '。'
                stats['半角标点转全角'] += 1

    i = 0
    while i < n - 1:
        c = text[i]
        if c in '。，、；：？！' and text[i + 1] == c and out[i] == c and out[i + 1] == c:
            j = i + 1
            while j < n and text[j] == c:
                out[j] = ''
                j += 1
            stats['重复标点折叠'] += 1
            i = j
        else:
            i += 1

    # 4b. 书名号/引号并列之间不加顿号：《A》、《B》→《A》《B》；"A"、"B"→"A""B"
    for i in range(1, n - 1):
        if text[i] == '、' and out[i] == '、':
            if (text[i - 1] == '》' and text[i + 1] == '《') or \
               (text[i - 1] == '”' and text[i + 1] == '“'):
                out[i] = ''
                stats['并列书名号/引号间顿号删除'] += 1

    # 4c. 发文字号：括号→六角〔〕；顺序号不加"第"、不编虚位（GB/T 9704 7.2.5）
    for m in re.finditer(r'([〔\[（(])((?:19|20|21)\d{2})([〕\]）)])(\s*)(第?)(0*)(?=\d{1,4}\s*号)',
                         text):
        if m.group(1) != '〔' and out[m.start(1)] == m.group(1):
            out[m.start(1)], out[m.start(3)] = '〔', '〕'
            stats['发文字号六角括号'] += 1
        if m.group(5) and out[m.start(5)] == '第':
            out[m.start(5)] = ''
            stats['发文序号"第"字删除'] += 1
        if m.group(6):
            for k in range(m.start(6), m.end(6)):
                out[k] = ''
            stats['发文序号虚位删除'] += 1

    # 4c2. 百分数范围两端都写%（GB/T 15835：63%～68%）：5～8% → 5%～8%
    for m in re.finditer(r'(?<![\d.%])(\d+(?:\.\d+)?)([～—~\-])(?=\d+(?:\.\d+)?%)', text):
        e = m.end(1) - 1
        if out[e] == text[e]:
            out[e] = text[e] + '%'
            stats['百分数范围补前一个%'] += 1
        if m.group(2) in '~-' and out[m.start(2)] == m.group(2):
            out[m.start(2)] = '～'
            stats['范围号规范为～'] += 1

    # 4c3. 半角波浪号：数字/百分号语境 5%~8% → 5%～8%
    for i, c in enumerate(text):
        if c == '~' and out[i] == '~':
            prev, nxt = ctx(i, -1), ctx(i, 1)
            if (prev.isdigit() or prev == '%' or CJK(prev)) and (nxt.isdigit() or CJK(nxt)):
                out[i] = '～'
                stats['范围号规范为～'] += 1

    # 4c4. "X月X号"→"X月X日"（有月份前缀才改；号楼/号线/号文件等不受影响）
    for m in re.finditer(r'月\s*\d{1,2}(号)(?![楼栋线路院区])', text):
        i = m.start(1)
        if out[i] == '号':
            out[i] = '日'
            stats['"号"改"日"'] += 1

    # 4d. 年份起止范围用一字线：2023-2025年 → 2023—2025年（号码/复合名词不动）
    for m in re.finditer(r'(?:19|20|21)\d{2}\s*([-‐‑–])\s*(?=(?:19|20|21)\d{2}\s*(?:年|[）)]))', text):
        i = m.start(1)
        if out[i] == text[i]:
            out[i] = '—'
            stats['年份范围改一字线'] += 1

    # 4e. 中文段落的直双引号→弯引号（成对时按开/闭交替转换；落单只提示）
    if any(CJK(c) for c in text):
        qpos = [i for i, c in enumerate(text) if c == '"' and out[i] == '"']
        if qpos:
            if len(qpos) % 2 == 0:
                for k, i in enumerate(qpos):
                    out[i] = '“' if k % 2 == 0 else '”'
                stats['直引号转弯引号'] += len(qpos)
            else:
                hints.append(f'直双引号有 {len(qpos)} 个（奇数，无法配对），请人工核对：'
                             f'「{text[:24]}…」')

    def keep_space(i):
        p = i - 1
        while p >= 0 and text[p] in SPACES:
            p -= 1
        j = i + 1
        while j < n and text[j] in SPACES:
            j += 1
        prev = text[p] if p >= 0 else ''
        nxt = text[j] if j < n else ''
        if not prev or not nxt:
            return False
        if CJK(prev) or CJK(nxt):
            return False
        return bool(re.match(r'[\w.,;:)%]', prev)) and bool(re.match(r'[\w(]', nxt))

    if not keep_inner_spaces:
        run_start = None
        for i in range(n + 1):
            is_sp = i < n and text[i] in SPACES
            if is_sp and run_start is None:
                run_start = i
            elif not is_sp and run_start is not None:
                if keep_space(run_start):
                    out[run_start] = ' '
                    for k in range(run_start + 1, i):
                        out[k] = ''
                    if text[run_start:i] != ' ':
                        stats['多空格并一'] += 1
                else:
                    for k in range(run_start, i):
                        out[k] = ''
                    stats['删除多余空格'] += 1
                run_start = None

    new = ''.join(out)

    m = PAT_L1.match(new)
    if m and m.group(2) != '、':
        new = new[:m.start(2)] + '、' + new[m.end(2):]
        stats['序号写法修正'] += 1
    m = PAT_L2.match(new)
    if m and m.group(2):
        new = new[:m.start(2)] + new[m.end(2):]
        stats['序号写法修正'] += 1
    m = PAT_L3.match(new)
    if m and m.group(2) != '.':
        new = new[:m.start(2)] + '.' + new[m.end(2):]
        stats['序号写法修正'] += 1
    m = PAT_L4.match(new)
    if m:
        fixed = f'（{m.group(1)}）'
        if new[:m.end()] != fixed:
            new = fixed + new[m.end():]
            stats['序号写法修正'] += 1

    if (PAT_L1.match(new) or PAT_L2.match(new)) and len(new) <= 30 and new.endswith('。') \
            and new.count('。') == 1:
        new = new[:-1]
        stats['标题末句号删除'] += 1

    # 8. 附件说明格式（GB/T 9704 7.3.4）："附件1：xx" → "附件：1.xx"；名称末尾不加标点
    m = re.match(r'^附件\s*(\d{1,2})\s*[：:.、]\s*', new)
    if m:
        new = f'附件：{m.group(1)}.' + new[m.end():]
        stats['附件说明格式修正'] += 1
    if re.match(r'^附件：', new) and len(new) <= 60 and new[-1] in '。；，、.;,':
        new = new[:-1]
        stats['附件名末尾标点删除'] += 1

    return new, stats, hints


def apply_to_runs(para, new_text):
    runs = [r for r in para.runs]
    if not runs:
        return
    old = ''.join(r.text for r in runs)
    if old == new_text:
        return
    a, b = old, new_text
    p = 0
    while p < len(a) and p < len(b) and a[p] == b[p]:
        p += 1
    s = 0
    while s < len(a) - p and s < len(b) - p and a[len(a) - 1 - s] == b[len(b) - 1 - s]:
        s += 1
    pos = 0
    written = False
    for r in runs:
        rl = len(r.text)
        r_start, r_end = pos, pos + rl
        if r_end <= p or r_start >= len(a) - s:
            pos = r_end
            continue
        if not written:
            keep_head = a[r_start:p] if r_start < p else ''
            tail_start = max(len(a) - s, r_start)
            r.text = keep_head + b[p:len(b) - s] + a[tail_start:r_end]
            written = True
        else:
            head = max(r_start, p)
            tail = min(r_end, len(a) - s)
            r.text = a[r_start:head] + a[tail:r_end]
        pos = r_end


def detect_heading(text):
    m = re.match(r'^([一二三四五六七八九十]{1,3})、', text)
    if m:
        return 1, cn2int(m.group(1))
    m = re.match(r'^（([一二三四五六七八九十]{1,3})）', text)
    if m:
        return 2, cn2int(m.group(1))
    m = re.match(r'^(\d{1,2})[\.、]', text)
    if m:
        return 3, int(m.group(1))
    m = re.match(r'^[（(](\d{1,2})[）)]', text)
    if m:
        return 4, int(m.group(1))
    return None, None


def set_eastasia(run, name):
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = rpr.makeelement(qn('w:rFonts'), {})
        rpr.insert(0, rf)
    rf.set(qn('w:eastAsia'), name)
    run.font.name = 'Times New Roman'


def apply_page_numbers(doc):
    """GB/T 9704 7.5：4号半角宋体阿拉伯数字，左右各一条一字线（— 4 —），
    一字线上缘距版心下边缘 7mm；单页码居右空一字，双页码居左空一字，连续编排。
    版心下边缘=页高297-下边距35=262mm，页码顶距页底 297-262-7=28mm。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement

    doc.settings.odd_and_even_pages_header_footer = True

    def styled(run):
        run.font.name = '宋体'
        set_eastasia(run, '宋体')
        run.font.name = '宋体'      # set_eastasia 会把西文字体设为 Times，页码须半角宋体
        run.font.size = Pt(14)
        run.font.bold = False
        return run

    def build(footer, odd):
        footer.is_linked_to_previous = False
        for extra in footer.paragraphs[1:]:
            extra._element.getparent().remove(extra._element)
        p = footer.paragraphs[0]
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
        p.paragraph_format.line_spacing = None
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if odd else WD_ALIGN_PARAGRAPH.LEFT
        if not odd:
            styled(p.add_run('　'))       # 双页码居左空一字
        styled(p.add_run('— '))
        fld = styled(p.add_run(''))
        for typ, instr in [('begin', None), (None, ' PAGE '), ('separate', None),
                           (None, None), ('end', None)]:
            if typ:
                e = OxmlElement('w:fldChar')
                e.set(qn('w:fldCharType'), typ)
                fld._element.append(e)
            elif instr:
                e = OxmlElement('w:instrText')
                e.set(qn('xml:space'), 'preserve')
                e.text = instr
                fld._element.append(e)
            else:
                e = OxmlElement('w:t')
                e.text = '1'
                fld._element.append(e)
        styled(p.add_run(' —'))
        if odd:
            styled(p.add_run('　'))       # 单页码居右空一字

    for sec in doc.sections:
        sec.footer_distance = Mm(28)
        build(sec.footer, odd=True)
        build(sec.even_page_footer, odd=False)


def clean_decorations(run):
    """公文黑字白底：清除文字颜色（统一黑）、突出显示、下划线、字符底纹。"""
    from docx.shared import RGBColor
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.highlight_color = None
    run.font.underline = False
    rpr = run._element.rPr
    if rpr is not None:
        shd = rpr.find(qn('w:shd'))
        if shd is not None:
            rpr.remove(shd)


def clean_para_shading(para):
    ppr = para._p.pPr
    if ppr is not None:
        shd = ppr.find(qn('w:shd'))
        if shd is not None:
            ppr.remove(shd)


def apply_layout(doc):
    """机构排版规范（2026-07 版，与 GB/T 9704 一致处从略）：
    题目=方正小标宋简体二号不加粗居中；一级黑体、二级楷体、三级/四级仿宋，
    全部不加粗，均三号；正文仿宋三号缩进2字符；标题行距33磅、正文28磅；
    全文黑字白底（清页面背景、段落底纹、高亮、彩字、下划线）；表格内字体统一
    仿宋黑字（字号、表头加粗保留原样）。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    # 页面背景色清除（w:document/w:background）
    bg = doc.element.find(qn('w:background'))
    if bg is not None:
        doc.element.remove(bg)
    for sec in doc.sections:
        sec.top_margin, sec.bottom_margin = Mm(37), Mm(35)
        sec.left_margin, sec.right_margin = Mm(28), Mm(26)
    # 2026-07-04 用户定版：用新版字体名"楷体/仿宋"（所有现代 Windows 自带，
    # 免安装即正确渲染），不用"楷体_GB2312/仿宋_GB2312"（需公文字体包，缺字体
    # 的机器会退化成宋体假扮）。两代字体同源同形，外观几乎无差别。
    # 2026-07-04 用户再定版：二级标题楷体不加粗（覆盖规范图片的"加粗"）
    FONTS = {1: ('黑体', False), 2: ('楷体', False), 3: ('仿宋', False),
             4: ('仿宋', False), None: ('仿宋', False)}
    title_pending = True   # 文档开头第一个非序号、非句子的段落视为题目
    for kind, block in iter_blocks(doc):
        if kind == 'tbl':
            # 表格自动调整（安全子集）：字体统一仿宋、黑字、去高亮/下划线/底纹；
            # 字号与表头加粗保留原样，单元格背景（表格样式）不动
            for row in block.rows:
                for cell in row.cells:
                    for cp in cell.paragraphs:
                        clean_para_shading(cp)
                        for r in cp.runs:
                            set_eastasia(r, '仿宋')
                            clean_decorations(r)
            continue
        if kind != 'p' or not block.text.strip():
            continue
        clean_para_shading(block)
        text = block.text.strip()
        level, _ = detect_heading(text)
        if title_pending:
            title_pending = False
            if not level and len(text) <= 50 and not text.endswith(('。', '，', '；', '：')):
                for r in block.runs:
                    set_eastasia(r, '方正小标宋简体')
                    r.font.size = Pt(22)
                    r.font.bold = False
                    clean_decorations(r)
                pf = block.paragraph_format
                pf.line_spacing = Pt(33)
                pf.space_before = pf.space_after = Pt(0)
                block.alignment = WD_ALIGN_PARAGRAPH.CENTER
                continue
        if is_chapter_heading(text):
            # 第X章 / 前言目录等：黑体三号居中，行距33，不加首行缩进
            for r in block.runs:
                set_eastasia(r, '黑体')
                r.font.size = Pt(16)
                r.font.bold = False
                clean_decorations(r)
            block.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = block.paragraph_format
            pf.line_spacing = Pt(33)
            pf.space_before = pf.space_after = Pt(0)
            block._p.get_or_add_pPr().get_or_add_ind().set(qn('w:firstLineChars'), '0')
            continue
        name, bold = FONTS.get(level, FONTS[None])
        for r in block.runs:
            set_eastasia(r, name)
            r.font.size = Pt(16)
            r.font.bold = bold
            clean_decorations(r)
        pf = block.paragraph_format
        pf.line_spacing = Pt(33) if level else Pt(28)
        pf.space_before = pf.space_after = Pt(0)   # 段间不留距，撑满版心（5.2.3）
        # 7.3.3 每个自然段左空二字——各级序数标题与正文一律首行缩进 2 字符
        ind = block._p.get_or_add_pPr().get_or_add_ind()
        ind.set(qn('w:firstLineChars'), '200')


def check_sequence(headings):
    """返回跳号/重号提示（不修改文档）。"""
    counters = {1: 0, 2: 0, 3: 0, 4: 0}
    hints = []
    for level, num, text in headings:
        for l in range(level + 1, 5):
            counters[l] = 0
        expect = counters[level] + 1
        if num == counters[level]:
            hints.append(f'序号重号：「{text[:20]}」')
        elif num > expect and not (num == 1 and counters[level] == 0):
            hints.append(f'序号跳号：「{text[:20]}」前缺第 {expect} 个同级序号')
        counters[level] = num
    return hints


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('docx')
    ap.add_argument('-out', '--out', required=True, help='终稿输出路径')
    ap.add_argument('--no-layout', action='store_true', help='保留原版式，只修文字格式硬伤')
    a = ap.parse_args()

    doc = Document(a.docx)
    total = Counter()
    headings, hints, to_delete = [], [], []
    first_heading_seen = False
    title_checked = False
    prev_p = None
    for kind, block in iter_blocks(doc):
        if kind == 'p':
            t = block.text
            if not t.strip():
                # 空白段落：正文区（首个标题之后）全删；封面区保留；
                # 含图片/分节符/分页符的不删（GB 公文段间不空行，靠行距）
                if first_heading_seen and not block._p.xpath(
                        './/w:drawing | .//w:pict | .//w:object | .//w:sectPr'
                        ' | .//w:br[@w:type="page"]'):
                    to_delete.append(block)
                prev_p = block._p
                continue
            level, num = detect_heading(t)
            if level or is_chapter_heading(t):
                first_heading_seen = True
            if level:
                headings.append((level, num, t))
            if not title_checked and not level and not is_chapter_heading(t):
                title_checked = True
                if len(t.strip()) > 20 and len(t.strip()) <= 50 \
                        and not t.strip().endswith(('。', '，', '；')):
                    hints.append(f'题目约超过一行（{len(t.strip())}字），请人工回行并'
                                 '排成梯形或菱形、词意完整（GB/T 9704 7.3.1）')
            if re.match(r'^附件\s*\d{0,2}\s*$', t.strip()):
                has_break = block._p.xpath('./w:pPr/w:pageBreakBefore') or (
                    prev_p is not None and prev_p.xpath('.//w:br[@w:type="page"]'))
                if not has_break:
                    hints.append(f'「{t.strip()}」未另面编排：附件应另起一页，"附件"'
                                 '及顺序号3号黑体顶格、标题居中于第三行（7.3.7），请人工分页')
            new, stats, hs = transform_text(t, keep_inner_spaces=is_chapter_heading(t))
            total += stats
            hints += hs
            if new != t:
                apply_to_runs(block, new)
            prev_p = block._p
        else:
            for row in block.rows:
                for cell in row.cells:
                    for cp in cell.paragraphs:
                        if not cp.text:
                            continue
                        new, stats, hs = transform_text(cp.text)
                        total += stats
                        hints += hs
                        if new != cp.text:
                            apply_to_runs(cp, new)

    for p in to_delete:
        p._p.getparent().remove(p._p)
    if to_delete:
        total['空白段落删除'] += len(to_delete)

    if not a.no_layout:
        apply_layout(doc)
        apply_page_numbers(doc)
    doc.save(a.out)

    n = sum(total.values())
    detail = '、'.join(f'{k} {v}' for k, v in total.most_common()) or '无'
    layout_msg = '版式已按机构规范统一（页边距、题目小标宋、标题黑/楷/仿宋、' \
        '正文仿宋3号、标题行距33磅、正文28磅、首行缩进；黑字白底：已清页面背景/' \
        '段落底纹/高亮/彩字/下划线；表格字体统一仿宋；页码按GB/T 9704 7.5：' \
        '宋体4号"— N —"，距版心下缘7mm，单页居右、双页居左各空一字）' \
        if not a.no_layout else '版式保留原样'
    print(f'已输出终稿：{a.out}')
    print(f'文字格式修改共 {n} 处：{detail}。{layout_msg}。')
    for h in check_sequence(headings) + hints:
        print(f'提示（未改动，需人工处理）：{h}')


if __name__ == '__main__':
    main()
