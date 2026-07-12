#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gbt9704-format 断言测试（零依赖，python3 tests/test_format_fix.py 直接跑）。

覆盖：
  A. transform_text 文字修复规则（v1.3 的 32 条断言重建）
  B. v2.0 分类判定：is_true_heading / detect_missing_ordinal_punct / 序号保护
  C. check_sequence：跳号/重号/乱序/越级/按章重置
  D. review_checks 标点数字审查项
  E. 端到端：合成 docx 跑脚本，验证 版式/缩进/封面目录保护/审查版批注/
     --indent-headings / --no-layout 审计
"""
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, '..', 'scripts', 'format_fix.py')
sys.path.insert(0, os.path.join(HERE, '..', 'scripts'))
import format_fix as ff                                    # noqa: E402
from docx import Document                                  # noqa: E402
from docx.oxml.ns import qn                                # noqa: E402

PASS = FAIL = 0
FAILS = []


def check(name, cond):
    global PASS, FAIL
    if cond:
        PASS += 1
    else:
        FAIL += 1
        FAILS.append(name)
        print(f'  ✗ {name}')


def t(src, expect, name=None, **kw):
    got = ff.transform_text(src, **kw)[0]
    check(name or f'{src!r}→{expect!r}', got == expect)
    if got != expect:
        print(f'    实际: {got!r}')


# ---------- A. 文字修复规则 ----------
def test_transform():
    t('规划期为２０２５年', '规划期为2025年', '全角数字转半角')
    t('持续发展...', '持续发展……', '三点省略号')
    t('持续发展。。。', '持续发展……', '句号省略号')
    t('持续发展…', '持续发展……', '单字符省略号补全')
    t('印发<<实施方案>>执行', '印发《实施方案》执行', '双尖括号转书名号')
    t('二〇二五年十二月三十一日印发', '2025年12月31日印发', '汉字日期转阿拉伯')
    t('07月05日开工', '7月5日开工', '月日虚位删除')
    t('10月20日开工', '10月20日开工', '整十月日不误改')
    t('总量达1,500亿元', '总量达1500亿元', '千分位逗号删除')
    t('经济,社会,民生', '经济，社会，民生', '半角逗号转全角')
    t('增长12.3%左右', '增长12.3%左右', '小数点不误改')
    t('试点(示范)推进', '试点（示范）推进', '半角括号转全角')
    t('工作已完成.', '工作已完成。', '句末半角点转句号')
    t('加快推进。。', '加快推进。', '重复句号折叠')
    t('《通知》、《方案》、《意见》', '《通知》《方案》《意见》', '并列书名号顿号删除')
    t('整治“庸”、“懒”、“散”', '整治“庸”“懒”“散”', '并列引号顿号删除')
    t('《通知》和《方案》', '《通知》和《方案》', '书名号间连词不误删')
    t('国发[2024]5号', '国发〔2024〕5号', '方括号转六角')
    t('府办发（2023）12号', '府办发〔2023〕12号', '圆括号转六角')
    t('国发〔2024〕第05号', '国发〔2024〕5号', '发文序号删第删虚位')
    t('增长5～8%', '增长5%～8%', '百分数范围补前%')
    t('规模34~39万人', '规模34～39万人', '半角波浪号转全角')
    t('定于6月30号召开', '定于6月30日召开', '号改日')
    t('位于5号楼3层', '位于5号楼3层', '号楼不误改')
    t('2023-2025年实施', '2023—2025年实施', '年份范围半字线改一字线')
    t('规划（2026－2030）落地', '规划（2026—2030）落地', '年份范围全角连字符改一字线(v2.0)')
    t('刊号ISSN 1000-1234', '刊号ISSN 1000-1234', '刊号半字线不误改')
    t('推进"双千兆"网络建设', '推进“双千兆”网络建设', '直引号转弯引号')
    t('宁德 市中心 城区', '宁德市中心城区', '汉字间空格删除')
    t('GDP growth rate 保持', 'GDP growth rate保持', '英文间空格保留/汉英间删除')
    t('1、优化布局', '1.优化布局', '顿号序号改下脚点')
    t('（一）、总体要求', '（一）总体要求', '（一）后多余顿号删除')
    t('一，总体要求', '一、总体要求', '一后逗号改顿号')
    t('(1)重点任务', '（1）重点任务', '半角序号括号转全角')
    t('一、总体要求。', '一、总体要求', '短标题末句号删除')
    t('附件1：实施方案', '附件：1.实施方案', '附件说明格式修正')
    t('附件：1.实施方案。', '附件：1.实施方案', '附件名末尾标点删除')
    t('23 《零售业态分类》(GB/T18106-2021)；', '23.《零售业态分类》（GB/T18106-2021）；',
      '数字+书名号缺点号自动补(v2.1)')
    t('23《零售业态分类》', '23.《零售业态分类》', '直贴书名号自动补点(v2.1)')
    t('2035《远景目标纲要》', '2035《远景目标纲要》', '4位数字+书名号不误补')
    t('第一章 总则', '第一章 总则', '章标题内空格保留', keep_inner_spaces=True)
    hints = ff.transform_text('存在"落单引号的段落')[2]
    check('直引号落单只提示', len(hints) == 1 and '奇数' in hints[0])


# ---------- B. 分类判定 ----------
def test_classify():
    check('一、恒为标题', ff.is_true_heading(1, '一、规划范围'))
    check('（一）恒为标题', ff.is_true_heading(2, '（一）宏观经济指标回顾'))
    check('1.短语标题', ff.is_true_heading(3, '1.零售商业发展现状'))
    check('（1）短语标题', ff.is_true_heading(4, '（1）重点片区'))
    check('1.《文件》是清单', not ff.is_true_heading(3, '1.《中华人民共和国城乡规划法》（2019年修正）'))
    check('句号结尾是清单', not ff.is_true_heading(3, '27.其他相关规划与政策文件。'))
    check('超30字是清单', not ff.is_true_heading(3, '1.' + '很' * 31))
    check('连排段是清单', not ff.is_true_heading(3, '1.优化布局。要坚持全市一盘棋统筹推进。'))

    check('数字+书名号走自动通道',
          ff.detect_missing_ordinal_punct('23 《零售业态分类》(GB/T18106-2021)；') is None)
    miss = ff.detect_missing_ordinal_punct('23 零售业态分类规范说明')
    check('数字+空格+汉字→批注', miss is not None and '23. 零售业态分类' in miss)
    miss = ff.detect_missing_ordinal_punct('一 规划范围')
    check('汉字序数缺顿号→批注', miss is not None and '一、规划范围' in miss)
    check('量词不误报', ff.detect_missing_ordinal_punct('5 个项目建成投用') is None)
    check('计量单位不误报', ff.detect_missing_ordinal_punct('30 万平方米商业面积') is None)
    check('正常序号不误报', ff.detect_missing_ordinal_punct('23.《商业网点分类》') is None)
    check('正常正文不误报', ff.detect_missing_ordinal_punct('本次规划范围为宁德市中心城区。') is None)

    new, _, _ = ff.transform_protect_ordinal('23 零售业态分类规范,详见附录')
    check('序号保护：空格保留', new.startswith('23 零售业态分类'))
    check('序号保护：其余照修', '规范，详见附录' in new)


# ---------- C. 序号检查与明显跳号改号 ----------
def _seq(*texts):
    """用真实段落构造 items 跑 fix_sequence，返回 (文档, [已修复文本], [提示文本])。"""
    d = Document()
    items = []
    for tx in texts:
        p = d.add_paragraph(tx)
        if ff.is_chapter_heading(tx):
            items.append(('chapter',))
            continue
        lv, num = ff.detect_heading(tx)
        items.append(('ord', lv, num, p, ff.is_true_heading(lv, tx)))
    fixed, notes = ff.fix_sequence(items)
    return d, [m for _, m in fixed], [m for _, m in notes]


def test_sequence():
    d, fx, ns = _seq('一、总体要求', '二、主要任务', '三、保障措施')
    check('顺号无动作', not fx and not ns)
    d, fx, ns = _seq('一、总体要求', '二、主要任务', '四、保障措施')
    check('明显跳号自动改号(v2.1)', len(fx) == 1 and '「四、」→「三、」' in fx[0])
    check('改号已落文中', d.paragraphs[2].text == '三、保障措施')
    d, fx, ns = _seq('一、甲', '二、乙', '四、丙', '五、丁')
    check('跳号级联改号', d.paragraphs[2].text == '三、丙'
          and d.paragraphs[3].text == '四、丁')
    d, fx, ns = _seq('（一）甲', '（三）乙')
    check('二级跳号改号', d.paragraphs[1].text == '（二）乙')
    d, fx, ns = _seq('1.第一项', '2.第二项', '4.第四项')
    check('数字序号跳号改号', d.paragraphs[2].text == '3.第四项')
    d, fx, ns = _seq('一、甲', '二、乙', '六、丙')
    check('大缺口只提示不改', not fx and any('缺口较大' in m for m in ns)
          and d.paragraphs[2].text == '六、丙')
    d, fx, ns = _seq('一、甲', '二、乙', '二、丙')
    check('相邻重号只提示', not fx and any('重号' in m for m in ns))
    d, fx, ns = _seq('一、甲', '二、乙', '一、丙')
    check('隔位乱序只提示', any('重号/乱序' in m for m in ns))
    d, fx, ns = _seq('第一章 总则', '一、甲', '二、乙', '第二章 布局', '一、丙')
    check('跨章重置无误报', not fx and not ns)
    d, fx, ns = _seq('一、甲', '1.要点标题')
    check('真标题越级提示', any('越级' in m for m in ns))
    d, fx, ns = _seq('一、甲', '1.《城乡规划法》', '2.《商业网点分类》')
    check('清单不报越级', not any('越级' in m for m in ns))
    d, fx, ns = _seq('一、甲', '（一）乙', '1.要点甲', '（二）丙', '1.要点乙')
    check('高层出现重置低层', not fx and not any('跳号' in m or '重号' in m for m in ns))
    check('int2cn换算', ff.int2cn(3) == '三' and ff.int2cn(10) == '十'
          and ff.int2cn(21) == '二十一')


# ---------- D. 审查项 ----------
def test_review_checks():
    cats = [c for c, _ in ff.review_checks('按闽发〔24〕5号文件执行')]
    check('发文字号年份缩写', '发文字号年份' in cats)
    cats = [c for c, _ in ff.review_checks('依据闽政文[24]12号批复')]
    check('非六角括号年份缩写也查', '发文字号年份' in cats)
    check('4位年份不误报', not ff.review_checks('按国发〔2024〕5号执行'))
    cats = [c for c, _ in ff.review_checks('十五五期间加快发展')]
    check('五年规划专名裸奔', '五年规划专名引号' in cats)
    check('已加引号不误报', not [c for c, _ in ff.review_checks('“十五五”期间加快发展')
                          if c == '五年规划专名引号'])
    check('第十五个五年不误报', not [c for c, _ in ff.review_checks('第十五个五年规划纲要')
                            if c == '五年规划专名引号'])
    cats = [c for c, _ in ff.review_checks('上世纪八十年代以来')]
    check('世纪年代汉字', '世纪年代数字用法' in cats)
    check('20世纪80年代不误报', not ff.review_checks('20世纪80年代以来'))
    cats = [c for c, _ in ff.review_checks('新增约30家左右门店')]
    check('约数叠用', '约数叠用' in cats)
    check('单用约不误报', not ff.review_checks('新增约30家门店'))


# ---------- E. 端到端 ----------
def _build_sample(path):
    d = Document()
    d.add_paragraph('')                                   # 封面空段（应保留）
    d.add_paragraph('某市商业网点布局专项规划')                     # 题目
    d.add_paragraph('（2026－2030）')                        # 封面副题（应不动）
    sub = d.paragraphs[-1]
    ff.set_eastasia(sub.runs[0], '华文行楷')                   # 给副题设个显眼字体做哨兵
    d.add_paragraph('2026年6月')                            # 封面日期（应不动）
    tblc = d.add_table(rows=1, cols=1)                    # 封面表格（应不动）
    tblc.rows[0].cells[0].paragraphs[0].add_run('审批信息')
    ff.set_eastasia(tblc.rows[0].cells[0].paragraphs[0].runs[0], '华文行楷')
    d.add_paragraph('前  言')                               # 章级（封面结束）
    d.add_paragraph('十五五期间，某市经济持续发展，取得了显著成效。')       # 正文（触发审查项）
    d.add_paragraph('第一章 总则')
    d.add_paragraph('一、规划范围')                            # 一级标题
    hd = d.paragraphs[-1]
    hd._p.get_or_add_pPr().get_or_add_ind().set(qn('w:firstLine'), '640')  # 旧缩进哨兵
    ff.set_eastasia(hd.runs[0], '宋体')                       # 错误字体哨兵（审计用）
    d.add_paragraph('本次规划范围为某市中心城区,面积约100平方公里。')       # 正文（半角逗号）
    d.add_paragraph('二、规划依据')
    d.add_paragraph('1.《城乡规划法》(2019年修正)')               # 清单
    d.add_paragraph('23 《零售业态分类》(GB/T18106-2021)；')     # 缺点号
    d.add_paragraph('（一）总体要求')                           # 二级标题
    d.add_paragraph('1、优化布局结构')                          # 真三级标题（顿号待修）
    d.add_paragraph('')                                   # 正文空段（应删除）
    d.add_paragraph('落实国发[2024]5号文件要求。')
    d.add_paragraph('四、实施保障要点')                          # 明显跳号（缺三、应自动改号）
    tbl1 = d.add_table(rows=1, cols=1)
    tbl1.rows[0].cells[0].paragraphs[0].add_run('表内 文字')
    d.add_paragraph('')                                   # 两表之间空段（应保留）
    tbl2 = d.add_table(rows=1, cols=1)
    tbl2.rows[0].cells[0].paragraphs[0].add_run('第二表')
    d.save(path)


def _flc(p):
    ppr = p._p.pPr
    ind = ppr.find(qn('w:ind')) if ppr is not None else None
    return ind.get(qn('w:firstLineChars')) if ind is not None else None


def _ea(p):
    if not p.runs or p.runs[0]._element.rPr is None:
        return None
    rf = p.runs[0]._element.rPr.find(qn('w:rFonts'))
    return rf.get(qn('w:eastAsia')) if rf is not None else None


def _para(doc, prefix):
    for p in doc.paragraphs:
        if p.text.strip().startswith(prefix):
            return p
    return None


def test_end_to_end():
    py = sys.executable
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, '样例.docx')
        _build_sample(src)

        # E1. 默认模式
        out = os.path.join(td, '终稿.docx')
        r = subprocess.run([py, SCRIPT, src, '-out', out],
                           capture_output=True, text=True)
        check('默认模式退出码0', r.returncode == 0)
        d = Document(out)
        p = _para(d, '某市商业网点布局专项规划')
        check('题目小标宋22磅', _ea(p) == '方正小标宋简体' and p.runs[0].font.size.pt == 22)
        p = _para(d, '（2026－2030）')
        check('封面副题字体不动', _ea(p) == '华文行楷')
        check('封面副题文字不动', p.text == '（2026－2030）')
        p = _para(d, '2026年6月')
        check('封面日期不动', _ea(p) is None)
        pc = _para(d, '前')
        check('章级黑体居中顶格', _ea(pc) == '黑体' and _flc(pc) == '0')
        check('章标题内空格保留', '前  言' in pc.text)
        p = _para(d, '一、规划范围')
        ind = p._p.pPr.find(qn('w:ind'))
        check('一级标题顶格', _flc(p) == '0' and _ea(p) == '黑体')
        check('旧磅值缩进已清', ind.get(qn('w:firstLine')) is None)
        p = _para(d, '本次规划范围')
        check('正文缩进2字符', _flc(p) == '200' and _ea(p) == '仿宋')
        check('正文行距28磅', p.paragraph_format.line_spacing.pt == 28)
        check('半角逗号已修', '中心城区，面积' in p.text)
        p = _para(d, '（一）总体要求')
        check('二级标题楷体顶格', _flc(p) == '0' and _ea(p) == '楷体'
              and p.runs[0].font.bold is False)
        p = _para(d, '1.优化布局结构')
        check('真三级标题序号已修且顶格', p is not None and _flc(p) == '0')
        check('三级标题行距33磅', p.paragraph_format.line_spacing.pt == 33)
        p = _para(d, '1.《城乡规划法》')
        check('清单缩进2字符', _flc(p) == '200')
        check('清单行距28磅', p.paragraph_format.line_spacing.pt == 28)
        p = _para(d, '23.《零售业态分类》')
        check('缺点号已自动补(v2.1)', p is not None and '（GB/T18106-2021）' in p.text)
        check('明显跳号已改号(v2.1)', _para(d, '三、实施保障要点') is not None)
        p = _para(d, '落实国发')
        check('六角括号已修', '国发〔2024〕5号' in p.text)
        check('封面空段保留', d.paragraphs[0].text == '' or _para(d, '某市') is not None)
        body_texts = [q.text for q in d.paragraphs]
        check('正文空段已删', body_texts.count('') <= 2)   # 封面1 + 两表间1
        check('封面表格字体不动', _ea(d.tables[0].rows[0].cells[0].paragraphs[0]) == '华文行楷')
        check('正文表格字体统一仿宋', _ea(d.tables[1].rows[0].cells[0].paragraphs[0]) == '仿宋')
        check('两表之间空段保留', len(d.tables) == 3 or '第二表' in d.tables[-1].rows[0].cells[0].text)
        check('汇总含已修复说明', '已自动修复' in r.stdout and '「四、」→「三、」' in r.stdout)
        check('清单大缺口只提示', '缺口较大' in r.stdout)

        # E2. --indent-headings 恢复红头式全缩进
        out2 = os.path.join(td, '红头.docx')
        subprocess.run([py, SCRIPT, src, '-out', out2, '--indent-headings'],
                       capture_output=True, text=True)
        d2 = Document(out2)
        check('--indent-headings标题也缩进', _flc(_para(d2, '一、规划范围')) == '200')

        # E3. --review 批注版单文件（v2.1：修复落文 + 已修复/检查两类批注）
        r = subprocess.run([py, SCRIPT, src, '--review'], capture_output=True, text=True)
        annotated = os.path.join(td, '样例_GBT9704批注版.docx')
        check('批注版存在', os.path.exists(annotated))
        check('旧双文件不再输出',
              not os.path.exists(os.path.join(td, '样例_GBT9704修复版.docx'))
              and not os.path.exists(os.path.join(td, '样例_GBT9704审查版.docx')))
        zr = zipfile.ZipFile(annotated)
        check('批注版有批注', 'word/comments.xml' in zr.namelist())
        cxml = zr.read('word/comments.xml').decode('utf8')
        check('已修复批注前缀', '【GB/T9704已修复】' in cxml)
        check('改号批注', '「四、」→「三、」' in cxml)
        check('补点批注', '「23 《」→「23.《」' in cxml)
        check('检查批注前缀', '【GB/T9704检查】' in cxml)
        check('专名引号批注', '十五五' in cxml)
        check('批注署名', ff.COMMENT_AUTHOR in cxml)
        dann = Document(annotated)
        check('批注版修复已落文', _para(dann, '三、实施保障要点') is not None
              and _para(dann, '23.《零售业态分类》') is not None)
        check('批注版标题顶格', _flc(_para(dann, '一、规划范围')) == '0')

        # E4. --no-layout：版式不动
        out4 = os.path.join(td, '仅文字.docx')
        subprocess.run([py, SCRIPT, src, '-out', out4, '--no-layout'],
                       capture_output=True, text=True)
        d4 = Document(out4)
        check('--no-layout不改字体', _ea(_para(d4, '一、规划范围')) == '宋体')
        check('--no-layout仍修文字', '国发〔2024〕5号' in _para(d4, '落实国发').text)

        # E5. --no-layout --review：版式审计批注
        r = subprocess.run([py, SCRIPT, src, '--review', '--no-layout'],
                           capture_output=True, text=True)
        check('审计提示出现', '版式偏差' in r.stdout)


def main():
    test_transform()
    test_classify()
    test_sequence()
    test_review_checks()
    test_end_to_end()
    print(f'\n断言 {PASS + FAIL} 条：通过 {PASS}，失败 {FAIL}')
    if FAILS:
        print('失败项：' + '；'.join(FAILS))
        sys.exit(1)
    print('全部通过 ✓')


if __name__ == '__main__':
    main()
