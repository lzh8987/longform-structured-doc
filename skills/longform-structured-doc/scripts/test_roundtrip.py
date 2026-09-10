# -*- coding: utf-8 -*-
"""
冒烟测试：验证 docs_to_md.py 与 assemble_docx.py 的核心路径。

运行：python scripts/test_roundtrip.py
"""

import os
import sys
import zipfile
import shutil
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import docs_to_md
import assemble_docx

TEST_DIR = os.path.join(os.path.dirname(HERE), '_test')


def test_docx_to_md():
    src = os.path.join(TEST_DIR, 't.docx')
    if not os.path.exists(src):
        print('[SKIP] 缺少 _test/t.docx')
        return
    md = docs_to_md.docx_to_md(src)
    assert '# 第一章' in md, md[:200]
    print('[OK] docx_to_md')


def _make_xlsx(path):
    """用标准 OOXML 命名空间临时构造一个最小 xlsx。"""
    S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    sst = ('<?xml version="1.0"?><sst xmlns="%s">'
           '<si><t>项目</t></si><si><t>流域面积</t></si></sst>' % S)
    sheet = ('<?xml version="1.0"?><worksheet xmlns="%s"><sheetData>'
             '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
             '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>100</v></c></row>'
             '</sheetData></worksheet>' % S)
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('xl/sharedStrings.xml', sst)
        z.writestr('xl/worksheets/sheet1.xml', sheet)


def test_xlsx_to_md():
    d = tempfile.mkdtemp()
    try:
        src = os.path.join(d, 't.xlsx')
        _make_xlsx(src)
        md = docs_to_md.xlsx_to_md(src)
        assert '|' in md and '流域面积' in md, md[:200]
        print('[OK] xlsx_to_md（标准命名空间）')
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_assemble():
    d = tempfile.mkdtemp()
    try:
        ch = os.path.join(d, 'c1.md')
        with open(ch, 'w', encoding='utf-8') as f:
            f.write(
                '# 第一章 概述\n\n正文。\n\n'
                '| 项目 | 数值 |\n| --- | --- |\n| 流域面积 | 100 |\n\n'
                '待定【待补充：金额】【待确认：高程】【需核实：出处】\n'
            )
        app = assemble_docx.collect_todo([ch])
        joined = '\n'.join(app[0][1])
        assert all(k in joined for k in ('待补充', '待确认', '需核实')), joined

        out = os.path.join(d, 'out.docx')
        assemble_docx.build_docx('标题', '副', '编制组', [ch], app, out,
                                 toc=True, page_numbers=True)
        z = zipfile.ZipFile(out)
        names = z.namelist()
        assert 'word/footer1.xml' in names, names
        doc = z.read('word/document.xml').decode('utf-8')
        assert '<w:tbl>' in doc and 'TOC' in doc and 'footerReference' in doc
        print('[OK] assemble_docx（表格/TOC/页码/三类空缺符）')
    finally:
        shutil.rmtree(d, ignore_errors=True)


if __name__ == '__main__':
    test_docx_to_md()
    test_xlsx_to_md()
    test_assemble()
    print('ALL PASS')
