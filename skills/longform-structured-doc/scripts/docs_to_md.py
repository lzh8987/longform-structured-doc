#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docs_to_md.py — 将办公/版式文档拆成 Markdown 文本。

支持：PDF / DOCX / XLSX / XLS / DOC  ->  .md

图片：DOCX 内嵌图、PDF 位图会提取到输出目录的 img/ 子目录（文件名带源文件前缀），
正文对应位置插入 `![图片](img/xxx)` 占位；矢量图（AutoCAD 导出等）提取不到时静默跳过。

设计原则（与 longform-structured-doc skill 的"环境坑"一致）：
- DOCX、XLSX 用标准库 zipfile + xml.etree.ElementTree 直解，零依赖；
- DOC 优先调用外部 antiword（项目实测最稳），不可用时退回纯 Python CFB 读取；
- PDF 用 pypdf；XLS 用 xlrd（二者需 pip 安装，缺失时给出明确提示）。

用法：
  python docs_to_md.py 输入1 [输入2 ...] [--out 输出目录] [--combine 合并.md] [--exts pdf,docx,xlsx,xls,doc]

  输入可以是文件或目录（目录递归收集匹配扩展名）。
  默认每个输入文件在同目录生成 <原名>.md；--out 指定集中输出目录；--combine 把所有结果拼成一个 md。
"""

import os
import re
import sys
import zipfile
import argparse
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

QW = lambda t: f'{{{W}}}{t}'
QS = lambda t: f'{{{S}}}{t}'
QR = lambda t: f'{{{R}}}{t}'
QA = lambda t: f'{{{A}}}{t}'


# ----------------------------------------------------------------------------
# 通用工具
# ----------------------------------------------------------------------------

def _table_to_md(table):
    """table: list[list[str]] -> markdown 表格字符串（首行作表头）。"""
    if not table:
        return ''
    ncol = max(len(r) for r in table)
    norm = [r + [''] * (ncol - len(r)) for r in table]
    head = norm[0]
    lines = ['| ' + ' | '.join(head) + ' |',
             '| ' + ' | '.join(['---'] * ncol) + ' |']
    for row in norm[1:]:
        lines.append('| ' + ' | '.join(row) + ' |')
    return '\n'.join(lines)


def _docx_text(el):
    """抽取 w:p / w:tc 下所有 w:t 的文本。"""
    return ''.join(t.text or '' for t in el.iter(QW('t')))


# ----------------------------------------------------------------------------
# DOCX：zipfile + document.xml（标准库 XML 解析，零依赖）
# ----------------------------------------------------------------------------

def _resolve_zip_path(base, target):
    """把 rels 里的 Target（相对 base 目录）解析为 zip 内路径。"""
    target = target.replace('\\', '/')
    if target.startswith('/'):
        return target.lstrip('/')
    stack = []
    for p in (base + '/' + target).split('/'):
        if p in ('', '.'):
            continue
        if p == '..':
            if stack:
                stack.pop()
        else:
            stack.append(p)
    return '/'.join(stack)


def _docx_extract_images(z, path, img_dir):
    """提取 DOCX 内嵌图片到 img_dir，返回 {rid: 'img/<文件名>'}。"""
    os.makedirs(img_dir, exist_ok=True)
    rels_path = 'word/_rels/document.xml.rels'
    img_map = {}
    if rels_path not in z.namelist():
        return img_map
    try:
        relroot = ET.fromstring(z.read(rels_path))
    except Exception:
        return img_map
    prefix = os.path.splitext(os.path.basename(path))[0]
    rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    for rel in relroot.findall('{%s}Relationship' % rel_ns):
        rtype = rel.get('Type', '')
        if 'image' not in rtype:
            continue
        rid = rel.get('Id')
        target = rel.get('Target', '')
        if not rid or not target:
            continue
        media_path = _resolve_zip_path('word', target)
        if media_path not in z.namelist():
            continue
        ext = os.path.splitext(target)[1].lower() or '.png'
        fname = '%s__%s%s' % (prefix, rid.replace('rId', 'image'), ext)
        try:
            with open(os.path.join(img_dir, fname), 'wb') as f:
                f.write(z.read(media_path))
            img_map[rid] = 'img/' + fname
        except Exception:
            continue
    return img_map


def _docx_para_mixed(el, img_map):
    """按文档顺序抽取段落内文本与图片占位，返回字符串。"""
    parts = []
    for node in el.iter():
        if node.tag == QW('t'):
            parts.append(node.text or '')
        elif node.tag == QA('blip'):
            rid = node.get(QR('embed')) or node.get(QR('link'))
            if rid in img_map:
                parts.append('\n\n![图片](%s)\n\n' % img_map[rid])
    return ''.join(parts)


# 中文编号模式：Word 样式缺失时的标题层级兜底（仅限短段落，避免误判正文）
CN_HEADING_PATTERNS = [
    (1, r'^第[一二三四五六七八九十百千万\d]+[章节篇部分]'),
    (2, r'^[一二三四五六七八九十]+、'),
    (3, r'^[（(][一二三四五六七八九十]+[）)]'),
    (4, r'^\d+(\.\d+){0,3}\s+\S'),
]


def _cn_heading_level(line):
    if len(line) > 30:
        return 0
    for lvl, pat in CN_HEADING_PATTERNS:
        if re.match(pat, line):
            return lvl
    return 0


def docx_to_md(path, img_dir=None):
    z = zipfile.ZipFile(path)
    root = ET.fromstring(z.read('word/document.xml'))
    body = root.find(QW('body'))
    if body is None:
        return ''
    img_map = _docx_extract_images(z, path, img_dir) if img_dir else {}
    out = []
    for child in body:
        if child.tag == QW('p'):
            style = ''
            ppr = child.find(QW('pPr'))
            if ppr is not None:
                pstyle = ppr.find(QW('pStyle'))
                if pstyle is not None:
                    style = pstyle.get(QW('val')) or ''
            line = _docx_para_mixed(child, img_map).strip()
            if not line:
                continue
            h = _docx_style_to_heading(style, line)
            out.append(('#' * h + ' ' + line) if h else line)
        elif child.tag == QW('tbl'):
            rows = []
            for tr in child.findall(QW('tr')):
                row = [_docx_text(tc).strip() for tc in tr.findall(QW('tc'))]
                if any(c for c in row):
                    rows.append(row)
            if rows:
                out.append(_table_to_md(rows))
    return '\n\n'.join(out)


def _docx_style_to_heading(style, line=''):
    if style:
        if style.lower() == 'title':
            return 1
        m = re.match(r'Heading(\d)', style, re.I)
        if m:
            lvl = int(m.group(1))
            return lvl if 1 <= lvl <= 6 else 0
    return _cn_heading_level(line)


# ----------------------------------------------------------------------------
# XLSX：zipfile + sharedStrings/sheet XML（标准库 XML 解析，零依赖）
# ----------------------------------------------------------------------------

def xlsx_to_md(path, img_dir=None):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    shared = []
    if 'xl/sharedStrings.xml' in names:
        sroot = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in sroot.findall(QS('si')):
            shared.append(''.join(t.text or '' for t in si.iter(QS('t'))))
    sheets = sorted([n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)])
    out = []
    for sf in sheets:
        sroot = ET.fromstring(z.read(sf))
        rows = []
        for row in sroot.iter(QS('row')):
            cells = []
            for c in row.findall(QS('c')):
                t = c.get('t')  # 无前缀属性，用裸名
                v = c.find(QS('v'))
                if t == 's' and v is not None:
                    try:
                        val = shared[int(v.text)]
                    except (ValueError, IndexError):
                        val = v.text or ''
                else:
                    val = v.text if v is not None else ''
                cells.append(val or '')
            if any(x.strip() for x in cells):
                rows.append(cells)
        if rows:
            out.append(_table_to_md(rows))
    return '\n\n'.join(out)


# ----------------------------------------------------------------------------
# XLS：xlrd（需安装）
# ----------------------------------------------------------------------------

def xls_to_md(path, img_dir=None):
    try:
        import xlrd
    except ImportError:
        return '[XLS 解析需要 xlrd：pip install xlrd；或转存为 xlsx 后用 xlsx_to_md]'
    wb = xlrd.open_workbook(path)
    out = []
    for sn in wb.sheet_names():
        ws = wb.sheet_by_name(sn)
        table = []
        for r in range(ws.nrows):
            row = []
            for c in range(ws.ncols):
                v = ws.cell_value(r, c)
                row.append('' if v == '' else str(v))
            if any(x.strip() for x in row):
                table.append(row)
        if table:
            out.append(f'### 工作表: {sn}\n\n' + _table_to_md(table))
    return '\n\n'.join(out)


# ----------------------------------------------------------------------------
# PDF：pypdf
# ----------------------------------------------------------------------------

def _pdf_extract_page_images(pg, prefix, idx, img_dir):
    """提取单页位图到 img_dir，返回图片占位 markdown。"""
    try:
        images = pg.images
    except Exception:
        return ''
    marks = []
    for j, img in enumerate(images):
        try:
            data = img.data
            name = getattr(img, 'name', None) or 'image'
            ext = os.path.splitext(name)[1].lower() or '.png'
            fname = '%s__p%d_%d%s' % (prefix, idx + 1, j + 1, ext)
            with open(os.path.join(img_dir, fname), 'wb') as f:
                f.write(data)
            marks.append('![图片](img/%s)' % fname)
        except Exception:
            continue
    return '\n\n'.join(marks)


def pdf_to_md(path, img_dir=None):
    try:
        from pypdf import PdfReader
    except ImportError:
        return '[PDF 解析需要 pypdf：pip install pypdf]'
    r = PdfReader(path)
    prefix = os.path.splitext(os.path.basename(path))[0]
    parts = []
    for i, pg in enumerate(r.pages):
        try:
            parts.append(pg.extract_text() or '')
        except Exception as e:
            parts.append(f'[{e}]')
        if img_dir:
            marks = _pdf_extract_page_images(pg, prefix, i, img_dir)
            if marks:
                parts.append(marks)
    return '\n\n'.join(parts)


# ----------------------------------------------------------------------------
# DOC：antiword 优先，否则纯 Python CFB 读取
# ----------------------------------------------------------------------------

def doc_to_md(path, img_dir=None):
    # 部分 .doc 实为改名的 OOXML（docx），先检测并回退
    if zipfile.is_zipfile(path):
        try:
            z = zipfile.ZipFile(path)
            if 'word/document.xml' in z.namelist():
                return docx_to_md(path, img_dir)
        except Exception:
            pass
    try:
        import shutil
        import subprocess
        if shutil.which('antiword'):
            r = subprocess.run(['antiword', path], capture_output=True, text=True, timeout=30)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout
    except Exception:
        pass
    try:
        return _doc_cfb_text(path)
    except Exception as e:
        return f'[DOC 解析失败: {e}；建议安装 antiword 后用 `antiword 文件.doc` 提取文本]'


def _doc_cfb_text(path):
    """纯 Python 解析 Word97 OLE2 复合文档，提取正文（支持简单与多片段情形）。"""
    data = open(path, 'rb').read()
    if data[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        raise ValueError('不是 Compound File(OLE2) 格式')
    ss = 1 << int.from_bytes(data[0x1E:0x20], 'little')       # sector size
    nfat = int.from_bytes(data[0x28:0x2C], 'little')
    dirstart = int.from_bytes(data[0x30:0x34], 'little')

    # FAT：头部 109 项 + DIFAT 链
    fat = [int.from_bytes(data[0x4C + i * 4:0x50 + i * 4], 'little') for i in range(109)]
    difat_start = int.from_bytes(data[0x44:0x48], 'little')
    seen = set()
    s = difat_start
    while s not in (0xFFFFFFFE, 0xFFFFFFFF) and s >= 0 and s not in seen:
        seen.add(s)
        sector = data[ss + s * ss: ss + (s + 1) * ss]
        for j in range(0, len(sector) - 4, 4):
            ent = int.from_bytes(sector[j:j + 4], 'little')
            if ent in (0xFFFFFFFE, 0xFFFFFFFF):
                break
            fat.append(ent)
        nxt = int.from_bytes(sector[-4:], 'little')
        if nxt in (0xFFFFFFFE, 0xFFFFFFFF):
            break
        s = nxt

    def chain(start):
        res = b''
        seen2 = set()
        cur = start
        while cur not in (0xFFFFFFFE, 0xFFFFFFFF) and cur >= 0 and cur not in seen2:
            seen2.add(cur)
            res += data[ss + cur * ss: ss + (cur + 1) * ss]
            if cur >= len(fat):
                break
            cur = fat[cur]
        return res

    # 目录：找 WordDocument 流
    dir_bytes = chain(dirstart)
    wd_start = wd_size = None
    for off in range(0, len(dir_bytes) - 128, 128):
        entry = dir_bytes[off:off + 128]
        obj_type = entry[0x42]
        if obj_type != 2:  # 仅看 stream
            continue
        name_len = int.from_bytes(entry[0x40:0x42], 'little')
        name = entry[:name_len - 2].decode('utf-16-le', errors='ignore')
        if name == 'WordDocument':
            wd_start = int.from_bytes(entry[0x74:0x78], 'little')
            wd_size = int.from_bytes(entry[0x78:0x7C], 'little')
            break
    if wd_start is None:
        raise ValueError('未找到 WordDocument 流')
    wd = chain(wd_start)[:wd_size]

    # FIB
    nFib = int.from_bytes(wd[0x02:0x04], 'little')
    fcMin = int.from_bytes(wd[0x18:0x1C], 'little')
    ccpText = int.from_bytes(wd[0x4C:0x50], 'little')

    # 多片段情形（nFib >= 0x00D9 时有 piece table）
    if nFib >= 0x00D9 and len(wd) >= 0x01AA:
        fcClx = int.from_bytes(wd[0x01A2:0x01A6], 'little')
        ccbClx = int.from_bytes(wd[0x01A6:0x01AA], 'little')
        clx = wd[fcClx:fcClx + ccbClx]
        if clx[:1] == b'\x01':  # Pcdt
            lcb = int.from_bytes(clx[1:5], 'little')
            plc = clx[5:5 + lcb]
            n = (lcb - 4) // 12
            if n > 0:
                cps = [int.from_bytes(plc[i * 4:i * 4 + 4], 'little') for i in range(n + 1)]
                pieces = []
                ok = True
                for i in range(n):
                    pcd_off = 4 * (n + 1) + i * 8
                    fc = int.from_bytes(plc[pcd_off + 2:pcd_off + 6], 'little')
                    cbeg = cps[i]
                    cend = cps[i + 1]
                    start = fc & 0x3FFFFFFF
                    is_uni = bool(fc & 0x40000000)
                    length = (cend - cbeg) * (2 if is_uni else 1)
                    chunk = wd[start:start + length]
                    try:
                        text = chunk.decode('utf-16-le' if is_uni else 'cp1252', errors='ignore')
                    except Exception:
                        ok = False
                        break
                    pieces.append(text)
                if ok:
                    return ''.join(pieces)

    # 简单情形（nFib == 0xC1 或 piece table 不可用）：整段 UTF-16LE
    if ccpText > 0:
        return wd[fcMin:fcMin + ccpText * 2].decode('utf-16-le', errors='ignore')
    raise ValueError('无法从 FIB 定位文本')


# ----------------------------------------------------------------------------
# 调度
# ----------------------------------------------------------------------------

DISPATCH = {
    '.pdf': pdf_to_md,
    '.docx': docx_to_md,
    '.xlsx': xlsx_to_md,
    '.xls': xls_to_md,
    '.doc': doc_to_md,
}


def convert_one(path, img_dir=None):
    ext = os.path.splitext(path)[1].lower()
    fn = DISPATCH.get(ext)
    if not fn:
        return None, f'[跳过不支持的扩展名: {ext}]'
    try:
        if img_dir and ext in ('.docx', '.pdf', '.doc'):
            os.makedirs(img_dir, exist_ok=True)
        return ext, fn(path, img_dir)
    except Exception as e:
        return ext, f'[{ext} 解析异常: {e}]'


def collect(inputs, exts):
    files = []
    for it in inputs:
        if os.path.isdir(it):
            for root, _, fs in os.walk(it):
                for f in fs:
                    if os.path.splitext(f)[1].lower() in exts:
                        files.append(os.path.join(root, f))
        elif os.path.isfile(it):
            files.append(it)
    return sorted(set(files))


def main():
    ap = argparse.ArgumentParser(description='将 PDF/DOCX/XLSX/XLS/DOC 拆成 Markdown 文本')
    ap.add_argument('inputs', nargs='+', help='文件或目录')
    ap.add_argument('--out', help='输出目录（默认与源文件同目录）')
    ap.add_argument('--combine', help='合并输出到单个 .md 文件')
    ap.add_argument('--exts', default='pdf,docx,xlsx,xls,doc',
                    help='递归目录时收集的扩展名，逗号分隔')
    args = ap.parse_args()

    exts = ['.' + e.strip().lower().lstrip('.') for e in args.exts.split(',') if e.strip()]
    files = collect(args.inputs, exts)
    if not files:
        print('没有匹配到任何文件。')
        return

    combine_img_dir = None
    if args.combine:
        combine_img_dir = os.path.join(os.path.dirname(os.path.abspath(args.combine)) or '.', 'img')

    blocks = []
    for fp in files:
        if args.combine:
            img_dir = combine_img_dir
        else:
            out_dir = args.out or os.path.dirname(fp)
            img_dir = os.path.join(out_dir, 'img')
        ext, md = convert_one(fp, img_dir)
        if md is None:
            print(f'[SKIP] {fp}')
            continue
        print(f'[OK] {fp}  ({len(md)} 字符)')
        if args.combine:
            blocks.append(f'# {os.path.basename(fp)}\n\n{md}')
        else:
            os.makedirs(out_dir, exist_ok=True)
            out_path = os.path.join(out_dir, os.path.splitext(os.path.basename(fp))[0] + '.md')
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(md)
            print(f'      -> {out_path}')

    if args.combine and blocks:
        os.makedirs(os.path.dirname(os.path.abspath(args.combine)) or '.', exist_ok=True)
        with open(args.combine, 'w', encoding='utf-8') as f:
            f.write('\n\n\n'.join(blocks))
        print(f'[合并] -> {args.combine}')


if __name__ == '__main__':
    main()
