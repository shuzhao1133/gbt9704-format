#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
format_fix.py — 政府报告 Word 格式一键修复（GB/T 9704 适用部分）

用法：
    python3 format_fix.py 文稿.docx -out 终稿.docx [--no-layout]
    python3 format_fix.py 文稿.docx --review [--no-layout]

缺省行为：
  1. 套用机构排版规范（与 GB/T 9704 一致）：页边距上37/下35/左28/右26mm；
     题目方正小标宋简体二号不加粗居中；正文仿宋_GB2312 3号（16pt）首行缩进2字符；
     一级标题黑体、二级楷体_GB2312、三级/四级仿宋_GB2312，全部不加粗，均3号；章级标题
     （第X章/前言等）黑体三号居中；标题行距固定32磅、正文28磅；数字英文
     Times New Roman；页码按 GB/T 9704 7.5 自动生成（宋体4号"— N —"奇右偶左）。
     加 --no-layout 则保留原版式，只修文字格式硬伤。
  2. 段落分类（v2.0）：封面（首个标题之前，题目除外）与目录（SDT 内容控件/
     TOC 样式段）完全不处理；各级真标题（一、（一）与短标题式 1.（1））顶格
     不缩进；"1.《文件》"式清单与正文缩进2字符。加 --indent-headings 恢复
     红头公文式全段缩进（GB/T 9704 7.3.3）。
  3. 文字格式硬伤（正文+表格逐格）：删多余空格（保留英文单词间空格）；
     中文语境半角标点转全角；全角字母数字转半角；.../。。。→……；重复标点折叠；
     序号写法 "1、"→"1."、"（一）、"→"（一）"、"一，"→"一、"、"(1)"→"（1）"；
     短标题末尾句号删除。
  4. 只改格式不改内容；明显的序号问题直接修（v2.1 用户定版）：
     数字+书名号缺点号（"23 《x》"→"23.《x》"）自动补；同级序号明显跳号
     （缺口≤2，如一、二、四）自动前移改号并级联。重号/乱序/大缺口/越级
     仍只提示——该改号还是补内容，机器无法判断。
  5. --review 审查模式：输出 原名_GBT9704批注版.docx 单文件——全部自动修复
     直接落在文中；敏感自动修复（补点、改号）各留一条【GB/T9704已修复】批注
     供核对；机器不敢改的留【GB/T9704检查】批注（署名"GB/T9704格式审查"）。

输出：终稿 .docx + 终端分类计数汇总（无修改记录文件）。
"""
import argparse
import os
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

# 序号缺标点（v2.1 用户定版分两档）：
#   自动补点——数字+书名号（"23 《零售业态分类》""23《零售业态分类》"→"23.《…"），
#   段首数字紧跟书名号几乎不可能是别的意思；
#   只批注——数字+空格+汉字/引号、汉字序数+空格（歧义更高，人工确认）。
# 量词排除：段首"5 个项目""30 万平方米"是计量表述不是序号，不误报
_MEASURE = '个万亿千百年月日家项条只台套批期次层间处元吨米km%％'
PAT_NUM_BOOK = re.compile(r'^(\d{1,3})[ 　\t]*(?=《)')
PAT_NUM_NODOT = re.compile(r'^(\d{1,3})([ 　\t]+)(?=["“一-鿿])(?![' + _MEASURE + '])')
PAT_CN_NODUN = re.compile(r'^([一二三四五六七八九十]{1,3})([ 　\t]+)'
                          r'(?=[一-鿿《"“])(?![' + _MEASURE + '])')

COMMENT_AUTHOR = 'GB/T9704格式审查'
COMMENT_INITIALS = '格'


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


CHAP_RE = re.compile(r'^第[一二三四五六七八九十百]+(?:[章篇编]|部分)')
SPECIAL_HEADINGS = {'前言', '序言', '目录', '附录', '参考文献', '后记',
                    '引言', '结语', '结束语', '摘要'}


def is_chapter_heading(text):
    """第X章/第X篇/第X编/第X部分，或前言/目录/附录等无序号篇章标题（黑体三号居中一层）。"""
    t = text.strip()
    if not t or len(t) > 30 or t.endswith(('。', '，', '；', '：', '、')):
        return False
    if CHAP_RE.match(t):
        return True
    return re.sub(r'\s+', '', t) in SPECIAL_HEADINGS


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


def is_true_heading(level, text):
    """标题/清单二分（v2.0，问题1）：一、（一）恒为标题；1.（1）按标题特征判定。
    清单特征（按正文待遇，缩进2字符、行距28磅）：序号后紧跟《（文件清单），
    或以句读结尾/段内含句号（"27.其他……。"、连排式"1.××。正文……"），或超30字。
    依据：机构文档实测——一、（一）与"1、零售商业发展现状"顶格，
    "1.《城乡规划法》""27.其他……。"与正文同缩进。"""
    if level in (1, 2):
        return True
    t = text.strip()
    m = re.match(r'^[（(]?\d{1,2}[）)]?[\.、，]?\s*', t)
    rest = t[m.end():] if m else t
    if rest.startswith('《'):
        return False
    if len(t) > 30:
        return False
    if t.endswith(('。', '；', '，', '：', '、', '.', ';', ',')) or '。' in t:
        return False
    return True


def classify_blocks(doc):
    """段落分类（v2.0，问题1）。返回 (blocks, roles, first_heading_idx)。
    roles 与 blocks 对齐，取值：
      title/cover/cover_table —— 封面区（首个标题之前）；封面除题目外完全不处理
      toc —— TOC 样式段（目录本体在 SDT 内容控件里时不进 blocks，天然跳过）
      chapter/heading/list/body/table/empty
    """
    blocks = list(iter_blocks(doc))
    fh = None
    for i, (kind, b) in enumerate(blocks):
        if kind != 'p':
            continue
        t = b.text.strip()
        if t and (is_chapter_heading(t) or detect_heading(t)[0]):
            fh = i
            break
    # 封面区仅在"章级标题开篇"的文档形态（规划/报告书：前言/目录/第X章在前）启用；
    # 直接以"一、"开篇的短报告没有封面，题目之后即正文（保持 v1.3 行为）
    cover_mode = fh is not None and is_chapter_heading(blocks[fh][1].text.strip())
    roles = []
    seen_first = False
    for i, (kind, b) in enumerate(blocks):
        if kind == 'tbl':
            roles.append('cover_table' if cover_mode and i < fh else 'table')
            continue
        t = b.text.strip()
        if not t:
            roles.append('empty')
            continue
        try:
            style = (b.style.name or '').lower()
        except Exception:
            style = ''
        if style.startswith('toc'):
            roles.append('toc')
            continue
        is_first = not seen_first
        seen_first = True
        if is_first and len(t) <= 50 and not detect_heading(t)[0] \
                and not is_chapter_heading(t) \
                and not t.endswith(('。', '，', '；', '：')):
            roles.append('title')
            continue
        if cover_mode and i < fh:
            roles.append('cover')
            continue
        if is_chapter_heading(t):
            roles.append('chapter')
            continue
        level, _ = detect_heading(t)
        if level:
            roles.append('heading' if is_true_heading(level, t) else 'list')
        else:
            roles.append('body')
    return blocks, roles, fh


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
    for m in re.finditer(r'…+', text):          # 单字符省略号…规范为……（两字符六点）
        if len(m.group()) != 2:
            out[m.start()] = '……'
            for k in range(m.start() + 1, m.end()):
                out[k] = ''
            stats['省略号规范'] += 1

    # 2a. 双尖括号errata：<<通知>> → 《通知》
    for m in re.finditer(r'<<([^<>]{1,40}?)>>', text):
        if any(CJK(c) for c in m.group(1)):
            out[m.start()], out[m.start() + 1] = '《', ''
            out[m.end() - 2], out[m.end() - 1] = '》', ''
            stats['书名号规范（<<>>转《》）'] += 1

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
            if c == '(' and ((nxt and CJK(nxt)) or (prev and CJK(prev))):
                out[i] = '（'      # 括号内侧或外侧邻汉字即转全角，杜绝全半角混用
            elif c == ')' and ((prev and CJK(prev)) or (nxt and CJK(nxt))):
                out[i] = '）'
            elif c in ',;:?!':
                if c == ',' and prev.isdigit() and nxt.isdigit():
                    continue          # 千分位由 2d 规则处理
                if (prev and CJK(prev)) or (nxt and CJK(nxt)):
                    out[i] = HALF2FULL[c]   # 前后任一侧邻汉字即属中文语境
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
    #     v2.0 补：全角连字符"－"（U+FF0D）同样纳入（实测"（2026－2030）"漏网）
    for m in re.finditer(r'(?:19|20|21)\d{2}\s*([-‐‑–－])\s*(?=(?:19|20|21)\d{2}\s*(?:年|[）)]))', text):
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

    # 序号补下脚点（v2.1）：段首"数字+书名号"缺点号直接修，"23 《x》/23《x》"→"23.《x》"
    m = PAT_NUM_BOOK.match(new)
    if m:
        new = m.group(1) + '.' + new[m.end():]
        stats['序号补下脚点'] += 1

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


def transform_protect_ordinal(text):
    """序号缺标点段（"23 《零售业态分类》"）：序号与间隔空格原样保留（既是证据也
    留给人工补标点），只对其后的内容做常规修复。返回 (新文本, stats, hints)。"""
    stripped = text.lstrip(SPACES)
    pre = Counter()
    if stripped != text:
        pre['删除多余空格'] += 1
    m = PAT_NUM_NODOT.match(stripped) or PAT_CN_NODUN.match(stripped)
    if not m:
        return transform_text(text)
    prefix = stripped[:m.end()]
    new, stats, hints = transform_text(stripped[m.end():])
    return prefix + new, stats + pre, hints


def detect_missing_ordinal_punct(text):
    """序号缺标点检测——只批注不改的歧义情形（数字+书名号已由 transform_text
    自动补点，不在此列）。返回批注文本或 None。"""
    t = text.strip()
    if PAT_NUM_BOOK.match(t):
        return None      # 走自动修复通道
    m = PAT_NUM_NODOT.match(t)
    if m:
        rest = t[m.end():]
        return (f'数字序号后疑似缺少下脚点，建议检查是否应调整为：'
                f'{m.group(1)}. {rest[:20]}')
    m = PAT_CN_NODUN.match(t)
    if m:
        rest = t[m.end():]
        if len(t) <= 30 and '。' not in t:
            return (f'汉字序数后疑似缺少顿号，建议检查是否应调整为：'
                    f'{m.group(1)}、{rest[:20]}')
    return None


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
    return


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


def enable_update_fields(doc):
    """文档含 TOC 域时置 w:updateFields，让 Word/WPS 打开时自动刷新目录页码
    （排版后分页移动，目录里的旧页码需要重算）。"""
    from docx.oxml import OxmlElement
    has_toc = any(e.text and 'TOC' in e.text
                  for e in doc.element.body.iter(qn('w:instrText')))
    if not has_toc:
        return False
    s = doc.settings.element
    if s.find(qn('w:updateFields')) is None:
        e = OxmlElement('w:updateFields')
        e.set(qn('w:val'), 'true')
        s.append(e)
    return True


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


def _set_line_fixed(para, pt):
    """行距固定值（Word 里显示"固定值 N磅"）。python-docx 对原规则为"最小值"
    (atLeast) 的段落赋 Length 时只改数值不改规则，须显式置 lineRule=exact 兜底。"""
    from docx.enum.text import WD_LINE_SPACING
    pf = para.paragraph_format
    pf.line_spacing = Pt(pt)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY


def _set_first_line(para, chars):
    """firstLineChars 以字符数控制缩进；置 0 时同时清掉绝对值 w:firstLine，
    防止旧的磅值缩进（如 406400 EMU）继续生效。"""
    ind = para._p.get_or_add_pPr().get_or_add_ind()
    ind.set(qn('w:firstLineChars'), chars)
    if chars == '0' and ind.get(qn('w:firstLine')) is not None:
        del ind.attrib[qn('w:firstLine')]


def apply_layout(doc, blocks, roles, indent_headings=False):
    """机构排版规范（2026-07 版，与 GB/T 9704 一致处从略）：
    题目=方正小标宋简体二号不加粗居中；一级黑体、二级楷体_GB2312、
    三级/四级仿宋_GB2312，全部不加粗，均三号；正文仿宋_GB2312三号缩进2字符；
    标题行距32磅、正文28磅；
    全文黑字白底（清页面背景、段落底纹、高亮、彩字、下划线）；表格内字体统一
    仿宋黑字（字号、表头加粗保留原样）。
    v2.0（问题1）：封面/目录完全不处理；真标题顶格不缩进（indent_headings=True
    恢复 7.3.3 红头式全缩进）；"1.《文件》"式清单按正文待遇。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    # 页面背景色清除（w:document/w:background）
    bg = doc.element.find(qn('w:background'))
    if bg is not None:
        doc.element.remove(bg)
    for sec in doc.sections:
        sec.top_margin, sec.bottom_margin = Mm(37), Mm(35)
        sec.left_margin, sec.right_margin = Mm(28), Mm(26)
    # 2026-07-12 用户定版（覆盖 2026-07-04 的"新版字体名"路线）：楷体/仿宋一律用
    # GB2312 版字体名"楷体_GB2312/仿宋_GB2312"（机构规范原文写法；政府机器与 WPS
    # 普遍自带，个人机器缺字体仅影响本机预览）。黑体/宋体/方正小标宋无 GB2312 变体。
    # 2026-07-04 用户定版：二级标题楷体不加粗（覆盖规范图片的"加粗"）
    FONTS = {1: ('黑体', False), 2: ('楷体_GB2312', False), 3: ('仿宋_GB2312', False),
             4: ('仿宋_GB2312', False), None: ('仿宋_GB2312', False)}
    for (kind, block), role in zip(blocks, roles):
        if role == 'table':
            # 表格自动调整（安全子集）：字体统一仿宋、黑字、去高亮/下划线/底纹；
            # 字号与表头加粗保留原样，单元格背景（表格样式）不动
            for row in block.rows:
                for cell in row.cells:
                    for cp in cell.paragraphs:
                        clean_para_shading(cp)
                        for r in cp.runs:
                            set_eastasia(r, '仿宋_GB2312')
                            clean_decorations(r)
            continue
        if role in ('empty', 'cover', 'cover_table', 'toc'):
            continue      # 封面/目录不处理（问题1）
        clean_para_shading(block)
        text = block.text.strip()
        if role == 'title':
            for r in block.runs:
                set_eastasia(r, '方正小标宋简体')
                r.font.size = Pt(22)
                r.font.bold = False
                clean_decorations(r)
            _set_line_fixed(block, 32)
            pf = block.paragraph_format
            pf.space_before = pf.space_after = Pt(0)
            block.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_first_line(block, '0')
            continue
        if role == 'chapter':
            # 第X章 / 前言目录等：黑体三号居中，行距32，不加首行缩进
            for r in block.runs:
                set_eastasia(r, '黑体')
                r.font.size = Pt(16)
                r.font.bold = False
                clean_decorations(r)
            block.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_line_fixed(block, 32)
            pf = block.paragraph_format
            pf.space_before = pf.space_after = Pt(0)
            _set_first_line(block, '0')
            continue
        level, _ = detect_heading(text)
        if role == 'heading':
            name, bold = FONTS.get(level, FONTS[None])
            spacing = 32
            # 问题1定版：真标题顶格不缩进（机构惯例）；--indent-headings 恢复
            # GB/T 9704 7.3.3 红头公文式"每个自然段左空二字"
            first_line = '200' if indent_headings else '0'
        else:      # body / list（"1.《文件》"式清单按正文待遇）
            name, bold = FONTS[None]
            spacing = 28
            first_line = '200'
        for r in block.runs:
            set_eastasia(r, name)
            r.font.size = Pt(16)
            r.font.bold = bold
            clean_decorations(r)
        _set_line_fixed(block, spacing)
        pf = block.paragraph_format
        pf.space_before = pf.space_after = Pt(0)   # 段间不留距，撑满版心（5.2.3）
        _set_first_line(block, first_line)
        # 对齐（机构规范"四、段落"原文，2026-07-12）：标题左对齐；正文/清单两端
        # 对齐——但仅在原为左对齐或未设时改，居中/右对齐段（图题表题、落款）
        # 是刻意编排，保留
        if role == 'heading':
            block.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif block.alignment in (None, WD_ALIGN_PARAGRAPH.LEFT,
                                 WD_ALIGN_PARAGRAPH.JUSTIFY):
            block.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def audit_layout(blocks, roles):
    """--no-layout --review 组合：只查不改，按机构规范粗核字体/字号/行距/缩进，
    偏差按类别汇总（只统计显式设置且不符的，样式继承值不误报）。返回汇总行列表。"""
    EXPECT_FONT = {'title': '方正小标宋简体', 'chapter': '黑体', 1: '黑体',
                   2: '楷体_GB2312', 3: '仿宋_GB2312', 4: '仿宋_GB2312',
                   'body': '仿宋_GB2312', 'list': '仿宋_GB2312'}
    EXPECT_SIZE = {'title': 22}
    dev = {}

    def note(cat, sample):
        if cat not in dev:
            dev[cat] = [0, sample]
        dev[cat][0] += 1

    for (kind, block), role in zip(blocks, roles):
        if kind != 'p' or role in ('empty', 'cover', 'cover_table', 'toc', 'table'):
            continue
        text = block.text.strip()
        level, _ = detect_heading(text)
        key = level if role == 'heading' else role
        exp_font = EXPECT_FONT.get(key)
        exp_size = EXPECT_SIZE.get(key, 16)
        for r in block.runs:
            rpr = r._element.rPr
            rf = rpr.find(qn('w:rFonts')) if rpr is not None else None
            ea = rf.get(qn('w:eastAsia')) if rf is not None else None
            if ea and exp_font and ea != exp_font:
                note(f'{role_label(role, level)}字体应为{exp_font}',
                     f'「{text[:14]}」现为{ea}')
                break
        for r in block.runs:
            if r.font.size is not None and abs(r.font.size.pt - exp_size) > 0.6:
                note(f'{role_label(role, level)}字号应为{exp_size}磅',
                     f'「{text[:14]}」现为{r.font.size.pt:g}磅')
                break
        ls = block.paragraph_format.line_spacing
        exp_ls = 32 if role in ('title', 'chapter', 'heading') else 28
        if ls is not None and hasattr(ls, 'pt') and abs(ls.pt - exp_ls) > 0.6:
            note(f'{role_label(role, level)}行距应为固定{exp_ls}磅',
                 f'「{text[:14]}」现为{ls.pt:g}磅')
    return [f'{cat}：{cnt} 处（如{sample}）' for cat, (cnt, sample) in dev.items()]


def role_label(role, level):
    if role == 'heading' and level:
        return f'{"一二三四"[level - 1]}级标题'
    return {'title': '题目', 'chapter': '章级标题', 'body': '正文', 'list': '清单段'}.get(role, role)


REVIEW_CAP = 8      # 同类批注上限，防刷屏


def review_checks(text):
    """标点/数字用法审查项（问题4，只批注不改）。返回 [(类别, 批注文本)]。"""
    out = []
    m = re.search(r'(?<=[一-鿿])[〔\[（(](\d{1,3})[〕\]）)]\s*第?0*\d{1,4}\s*号', text)
    if m:
        out.append(('发文字号年份',
                    f'发文字号年份应用4位全称（GB/T 9704 7.2.5），"〔{m.group(1)}〕"疑似缩写，请核对'))
    m = re.search(r'(?<![“"])(十[三四五六]五)(?![”"])', text)
    if m:
        out.append(('五年规划专名引号',
                    f'五年规划专名建议加引号："{m.group(1)}"（GB/T 15835），请核对'))
    if re.search(r'上世纪|[一二三四五六七八九]十年代', text):
        out.append(('世纪年代数字用法',
                    '世纪、年代宜用阿拉伯数字（GB/T 15835），如"20世纪80年代"，请核对'))
    m = re.search(r'[约近]\s*\d+(?:\.\d+)?[^。；，]{0,8}?(左右|上下)', text)
    if m:
        out.append(('约数叠用',
                    f'约数表述疑似叠用（"约/近…{m.group(1)}"），宜保留其一（GB/T 15835），请核对'))
    return out


def int2cn(n):
    """1—99 → 汉字数字（一、二…十、十一…九十九）。"""
    U = '一二三四五六七八九'
    if not 1 <= n <= 99:
        return str(n)
    if n < 10:
        return U[n - 1]
    tens, ones = divmod(n, 10)
    s = ('' if tens == 1 else U[tens - 1]) + '十'
    return s + (U[ones - 1] if ones else '')


ORD_MARK = {1: lambda n: f'{int2cn(n)}、', 2: lambda n: f'（{int2cn(n)}）',
            3: lambda n: f'{n}.', 4: lambda n: f'（{n}）'}
ORD_HEAD_RE = {1: re.compile(r'^[一二三四五六七八九十]{1,3}、'),
               2: re.compile(r'^（[一二三四五六七八九十]{1,3}）'),
               3: re.compile(r'^\d{1,2}\.'),
               4: re.compile(r'^[（(]\d{1,2}[）)]')}
RENUMBER_GAP_MAX = 2   # 缺口≤2 视为"明显跳号"自动前移改号；更大缺口疑似结构缺失只提示


def fix_sequence(items):
    """序号检查 + 明显跳号自动改号（v2.1 用户定版：明显问题直接修）。
    items: ('chapter',) 或 ('ord', level, num, para, is_head)。
    章级标题重置各级计数器；跳号缺口≤RENUMBER_GAP_MAX 自动前移改号（级联），
    每处附【已修复】说明供核对；重号/乱序/越级/大缺口只提示——该改哪个号、
    还是缺了一节内容，机器无法判断。
    返回 (fixed_log, notes)，均为 [(para, 文本)]。"""
    counters = {1: 0, 2: 0, 3: 0, 4: 0}
    fixed, notes = [], []
    for it in items:
        if it[0] == 'chapter':
            counters = {1: 0, 2: 0, 3: 0, 4: 0}
            continue
        _, level, num, para, is_head = it
        text = para.text.strip()
        for l in range(level + 1, 5):
            counters[l] = 0
        expect = counters[level] + 1
        if is_head and level >= 2 and counters[level - 1] == 0:
            notes.append((para, f'层级疑似越级：「{text[:20]}」的上一级序号尚未出现'
                                '（若为并列清单可忽略）'))
        if num == expect:
            pass
        elif num <= counters[level]:
            notes.append((para, f'序号重号/乱序：「{text[:20]}」同级序号此前已用到 '
                                f'{counters[level]}'))
        elif num - expect <= RENUMBER_GAP_MAX:
            m = ORD_HEAD_RE[level].match(text)
            if m:
                new_mark = ORD_MARK[level](expect)
                apply_to_runs(para, new_mark + text[m.end():])
                fixed.append((para, f'序号跳号已自动改号：「{m.group()}」→「{new_mark}」'
                                    f'（{text[m.end():][:14]}）。若此处实为缺失一节内容、'
                                    '或正文有交叉引用旧序号，请恢复原号并人工处理'))
                num = expect
            else:
                notes.append((para, f'序号跳号：「{text[:20]}」前缺第 {expect} 个同级序号'))
        else:
            notes.append((para, f'序号跳号：「{text[:20]}」前缺第 {expect} 个同级序号'
                                '（缺口较大，疑似结构缺失，未自动改号）'))
        counters[level] = max(counters[level], num)
    return fixed, notes


def add_review_comments(doc, entries):
    """批注版落批注。entries: [(para|None, 完整批注文本)]——调用方自带
    【GB/T9704已修复】/【GB/T9704检查】前缀。para 为 None 或已被删除时挂到
    第一个非空段。批注锚定段首 run。"""
    first = None
    for kind, b in iter_blocks(doc):
        if kind == 'p' and b.text.strip() and b.runs:
            first = b
            break
    added = 0
    for para, text in entries:
        target = para
        if target is None or target._p.getparent() is None or not target.runs:
            target = first
        if target is None:
            continue
        doc.add_comment(runs=[target.runs[0]], text=text,
                        author=COMMENT_AUTHOR, initials=COMMENT_INITIALS)
        added += 1
    return added


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('docx')
    ap.add_argument('-out', '--out', help='终稿输出路径（--review 时可省略，按规则命名）')
    ap.add_argument('--no-layout', action='store_true', help='保留原版式，只修文字格式硬伤')
    ap.add_argument('--review', action='store_true',
                    help='审查模式：输出批注版单文件（自动修复落文+已修复/检查批注）')
    ap.add_argument('--indent-headings', action='store_true',
                    help='恢复红头公文式全段缩进（GB/T 9704 7.3.3，各级标题也左空二字）')
    a = ap.parse_args()
    if not a.review and not a.out:
        ap.error('未指定 -out（非 --review 模式必填）')

    doc = Document(a.docx)
    blocks, roles, fh = classify_blocks(doc)
    total = Counter()
    notes, fixed_log, to_delete, captions = [], [], [], []
    prev_p = None
    title_checked = False
    for idx, ((kind, block), role) in enumerate(zip(blocks, roles)):
        if kind == 'p':
            t = block.text
            if role == 'empty':
                # 空白段落：正文区（首个标题之后）全删；封面区保留；
                # 含图片/分节符/分页符的不删（GB 公文段间不空行，靠行距）
                prev_el, next_el = block._p.getprevious(), block._p.getnext()
                between_tables = (prev_el is not None and prev_el.tag.endswith('}tbl')
                                  and next_el is not None and next_el.tag.endswith('}tbl'))
                if fh is not None and idx > fh and not between_tables and not block._p.xpath(
                        './/w:drawing | .//w:pict | .//w:object | .//w:sectPr'
                        ' | .//w:br[@w:type="page"]'):
                    to_delete.append(block)   # 两表之间的空段保留，否则删除后表格会合并
                prev_p = block._p
                continue
            if role in ('cover', 'toc'):
                prev_p = block._p
                continue      # 封面/目录：文字与版式一概不动（问题1）
            mbook = PAT_NUM_BOOK.match(t.strip())
            if mbook:
                old_pref = t.strip()[:mbook.end()] + '《'
                fixed_log.append((block, f'序号已自动补下脚点：「{old_pref}」→'
                                  f'「{mbook.group(1)}.《」，请顺带核对'))
            if role == 'title' or (not title_checked and role == 'body'):
                title_checked = True
                if len(t.strip()) > 20 and len(t.strip()) <= 50 \
                        and not t.strip().endswith(('。', '，', '；')):
                    notes.append((block, f'题目约超过一行（{len(t.strip())}字），请人工回行并'
                                  '排成梯形或菱形、词意完整（GB/T 9704 7.3.1）'))
            if re.match(r'^附件\s*\d{0,2}\s*$', t.strip()):
                has_break = block._p.xpath('./w:pPr/w:pageBreakBefore') or (
                    prev_p is not None and prev_p.xpath('.//w:br[@w:type="page"]'))
                if not has_break:
                    notes.append((block, f'「{t.strip()}」未另面编排：附件应另起一页，"附件"'
                                  '及顺序号3号黑体顶格、标题居中于第三行（7.3.7），请人工分页'))
            cm = re.match(r'^(表|图)\s*(\d+|[一二三四五六七八九十]+)', t.strip())
            if cm:
                captions.append((cm.group(1), cm.group(2), t.strip()[:20], block))
            if t.count('《') != t.count('》'):
                notes.append((block, f'书名号不配对（《×{t.count("《")} vs 》×{t.count("》")}）：'
                              f'「{t.strip()[:24]}…」'))
            miss = detect_missing_ordinal_punct(t)
            if miss:
                notes.append((block, miss))
            for cat, msg in review_checks(t):
                notes.append((block, (cat, msg)))
            new, stats, hs = transform_protect_ordinal(t) if miss else transform_text(
                t, keep_inner_spaces=(role == 'chapter') or bool(cm))  # 表题/图题空格保留
            total += stats
            notes += [(block, h) for h in hs]
            if new != t:
                apply_to_runs(block, new)
            prev_p = block._p
        else:
            if role == 'cover_table':
                continue      # 封面表格（如审批表）不处理
            for row in block.rows:
                for cell in row.cells:
                    for cp in cell.paragraphs:
                        if not cp.text:
                            continue
                        new, stats, hs = transform_text(cp.text)
                        total += stats
                        notes += [(cp, h) for h in hs]
                        if new != cp.text:
                            apply_to_runs(cp, new)

    for p in to_delete:
        p._p.getparent().remove(p._p)
    if to_delete:
        total['空白段落删除'] += len(to_delete)

    # 图表编号统一性检查（只提示不改：改号涉及正文引用联动）
    for kind_c in ('表', '图'):
        items = [(num, txt, para) for k, num, txt, para in captions if k == kind_c]
        if not items:
            continue
        styles = {'阿拉伯' if num.isdigit() else '汉字' for num, _, _ in items}
        if len(styles) > 1:
            notes.append((items[0][2], f'{kind_c}题编号体例混用（{kind_c}1 与 {kind_c}一 并存），'
                          '建议统一为阿拉伯数字并全文连续编号'))
        nums = [(int(n), para) for n, _, para in items if n.isdigit()]
        for i2, (v, para) in enumerate(nums, 1):
            if v != i2:
                notes.append((para, f'{kind_c}题编号疑似断档/乱序：期望{kind_c}{i2}，'
                              f'实际出现{kind_c}{v}，请人工核对全文{kind_c}号与正文引用'))
                break

    # 序号检查与明显跳号自动改号（第二遍收集：补点后的"23.《"清单项已纳入计数，
    # 依据清单缺点号补齐后不再误报"缺第23条"）
    seq_items = []
    for (kind, block), role in zip(blocks, roles):
        if kind != 'p' or role in ('empty', 'cover', 'toc', 'title'):
            continue
        if block._p.getparent() is None:
            continue      # 已删除的空段
        t2 = block.text.strip()
        if not t2:
            continue
        if role == 'chapter':
            seq_items.append(('chapter',))
            continue
        lv, num2 = detect_heading(t2)
        if lv:
            seq_items.append(('ord', lv, num2, block, role == 'heading'))
    seq_fixed, seq_notes = fix_sequence(seq_items)
    if seq_fixed:
        total['序号跳号自动改号'] += len(seq_fixed)
    fixed_log += seq_fixed
    notes += seq_notes

    # 落款成文日期编排检查（检测到文末日期落款才提示；封面日期不在此列）
    tail = [b for (k, b), r in zip(blocks, roles)
            if k == 'p' and r not in ('empty',) and b.text.strip()
            and b._p.getparent() is not None][-2:]
    for b in tail:
        if re.fullmatch(r'\d{4}年\d{1,2}月(\d{1,2}日)?', b.text.strip()):
            notes.append((b, '文末日期若为成文日期落款，应右空四字编排（GB/T 9704 7.3.5.2），请核对'))
            break

    # 同类审查批注限流：超过 REVIEW_CAP 条的类别只保留前几条并加尾注
    capped, cat_count = [], Counter()
    for para, item in notes:
        if isinstance(item, tuple):
            cat, msg = item
            cat_count[cat] += 1
            if cat_count[cat] > REVIEW_CAP:
                continue
            if cat_count[cat] == REVIEW_CAP:
                msg += '（同类问题较多，此后不再逐条批注，请全文排查）'
            capped.append((para, msg))
        else:
            capped.append((para, item))
    notes = capped

    if not a.no_layout:
        apply_layout(doc, blocks, roles, indent_headings=a.indent_headings)
        apply_page_numbers(doc)
        enable_update_fields(doc)
    elif a.review:
        for line in audit_layout(blocks, roles):
            notes.append((None, f'版式偏差（--no-layout 未修复）：{line}'))

    base = os.path.splitext(a.docx)[0]
    if a.review:
        out_path = a.out or f'{base}_GBT9704批注版.docx'
        add_review_comments(
            doc, [(p, f'【GB/T9704已修复】{h}') for p, h in fixed_log]
                 + [(p, f'【GB/T9704检查】{h}') for p, h in notes])
    else:
        out_path = a.out
    doc.save(out_path)

    n = sum(total.values())
    detail = '、'.join(f'{k} {v}' for k, v in total.most_common()) or '无'
    layout_msg = '版式已按机构规范统一（页边距、题目小标宋、标题黑体/楷体_GB2312/' \
        '仿宋_GB2312、正文仿宋_GB2312 3号、标题行距32磅、正文28磅；真标题顶格、' \
        '正文与清单缩进2字符；封面与目录未改动；黑字白底：已清页面背景/段落底纹/' \
        '高亮/彩字/下划线；表格字体统一仿宋_GB2312；页码按GB/T 9704 7.5：' \
        '宋体4号"— N —"，距版心下缘7mm，单页居右、双页居左各空一字）' \
        if not a.no_layout else '版式保留原样'
    if not a.no_layout and a.indent_headings:
        layout_msg = layout_msg.replace('真标题顶格、正文与清单缩进2字符',
                                        '全段左空二字（--indent-headings）')
    print(f'已输出{"批注版" if a.review else "终稿"}：{out_path}')
    if a.review:
        print(f'批注共 {len(fixed_log) + len(notes)} 条：已修复说明 {len(fixed_log)}、'
              f'需人工确认 {len(notes)}')
    print(f'文字格式修改共 {n} 处：{detail}。{layout_msg}。')
    for para, h in fixed_log:
        print(f'已自动修复（请顺带核对）：{h}')
    for para, h in notes:
        print(f'提示（未改动，需人工处理）：{h}')


if __name__ == '__main__':
    main()
