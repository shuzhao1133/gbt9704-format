#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把「已在 clean 版改掉」的批注在批注版里标灰。

用途：本 skill 固定产出两个版本——批注版（修改前，全量批注）与 clean 版
（已自动修改格式与错字）。凡是 clean 版真改掉的问题，其对应批注要在批注版
里标出来，人一眼就能分清「已处理」和「待你决定」。

标灰采取两种手段同时施加（用户定版）：
  1. 批注锚定的原文加灰色底纹（w:shd fill=D9D9D9）——正文里一眼可见；
  2. 该批注的作者名改为「已处理」——审阅窗格里可按作者筛选。
     （Word 的批注颜色由作者自动分配，无法逐条设色，改作者名是唯一可行的
      区分办法。）

用法：
    python3 mark_resolved.py 批注版.docx -out 输出.docx --keywords 关键词1 关键词2
    python3 mark_resolved.py 批注版.docx -out 输出.docx --ids 3 7 12
    python3 mark_resolved.py 批注版.docx --list        # 先列出全部批注及其编号

--keywords：批注正文包含任一关键词即视为已处理（如「错别字」「重复」）。
--ids     ：直接按 --list 给出的编号指定。
两者可同时使用，取并集。
"""
import argparse
import os
import re
import shutil
import zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(tag):
    return '{%s}%s' % (W, tag)

GREY = 'D9D9D9'
RESOLVED_AUTHOR = '已处理'


def read_parts(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    data = {n: z.read(n) for n in names}
    z.close()
    return names, data


def comment_text(el):
    return ''.join(t.text or '' for t in el.iter(q('t')))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('-out', default=None)
    ap.add_argument('--keywords', nargs='*', default=[])
    ap.add_argument('--ids', nargs='*', default=[])
    ap.add_argument('--list', action='store_true')
    ap.add_argument('--author', default=RESOLVED_AUTHOR)
    a = ap.parse_args()

    names, data = read_parts(a.docx)
    if 'word/comments.xml' not in data:
        print('该文件没有批注，无需处理。')
        return

    croot = etree.fromstring(data['word/comments.xml'])
    comments = croot.findall(q('comment'))

    if a.list:
        for c in comments:
            print('%s\t%s\t%s' % (c.get(q('id')), c.get(q('author')),
                                  comment_text(c)[:70]))
        return

    want_ids = set(str(x) for x in a.ids)
    for c in comments:
        txt = comment_text(c)
        if a.keywords and any(k in txt for k in a.keywords):
            want_ids.add(c.get(q('id')))

    if not want_ids:
        print('没有匹配到要标灰的批注（请用 --list 查看，或调整 --keywords/--ids）。')
        return

    # 1) 改作者名
    changed = 0
    for c in comments:
        if c.get(q('id')) in want_ids:
            c.set(q('author'), a.author)
            c.set(q('initials'), a.author[:1])
            changed += 1
    data['word/comments.xml'] = etree.tostring(
        croot, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 2) 锚定原文加灰底
    droot = etree.fromstring(data['word/document.xml'])
    body = droot.find(q('body'))
    order = list(body.iter())
    shaded = 0
    for cid in want_ids:
        start = end = None
        for el in droot.iter(q('commentRangeStart')):
            if el.get(q('id')) == cid:
                start = el
                break
        for el in droot.iter(q('commentRangeEnd')):
            if el.get(q('id')) == cid:
                end = el
                break
        if start is None or end is None:
            continue
        seen, collecting = [], False
        for el in droot.iter():
            if el is start:
                collecting = True
                continue
            if el is end:
                break
            if collecting and el.tag == q('r'):
                seen.append(el)
        for r in seen:
            rPr = r.find(q('rPr'))
            if rPr is None:
                rPr = etree.SubElement(r, q('rPr'))
                r.insert(0, rPr)
            shd = rPr.find(q('shd'))
            if shd is None:
                shd = etree.SubElement(rPr, q('shd'))
            shd.set(q('val'), 'clear')
            shd.set(q('color'), 'auto')
            shd.set(q('fill'), GREY)
            shaded += 1
    data['word/document.xml'] = etree.tostring(
        droot, xml_declaration=True, encoding='UTF-8', standalone=True)

    out = a.out or (os.path.splitext(a.docx)[0] + '-已标灰.docx')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    print('已标灰批注 %d 条，锚定原文加灰底 %d 处（作者名改为「%s」）：%s'
          % (changed, shaded, a.author, out))


if __name__ == '__main__':
    main()
