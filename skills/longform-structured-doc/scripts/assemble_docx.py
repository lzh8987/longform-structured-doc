# -*- coding: utf-8 -*-
"""
assemble_docx.py — 把「有序 Markdown 章节 + 附录」打包成合规 Word 文档。

基于 python-docx 生成，确保 Word/WPS 都能正常打开。

特性：
  - 自动映射 Markdown 标题：#→H1 … ####→H4；引用块(>) 作为缩进段落；
    代码块(```) 转为等宽段落；表格行(|...|) 转为真 Word 表格（带边框、表头加灰底）。
  - 自动抽取全文【待补充 / 待确认 / 需核实】生成「附录B 空缺事项汇总」（按三类分列）。
  - 可选 --toc 生成目录域、--page-numbers 生成页脚页码。
  - 内嵌 A4 页边距与中文字体样式（正文宋体、标题黑体，ascii 用 Times New Roman）。

用法：
  python assemble_docx.py --title "标题" --subtitle "副标题" --author "编制组" \
      --chapters ch1.md ch2.md --out out.docx

  # 或用 manifest.json：
  python assemble_docx.py --manifest manifest.json

manifest.json 字段：
  {
    "title": "...", "subtitle": "...", "author": "...",
    "chapters": ["ch1.md", ...],
    "source_docs": ["1. ...", ...],   # 附录A 依据文件清单
    "refs": ["1. ...", ...],          # 附录C 参考文献
    "extra_appendices": [             # 可选：自定义附录
      {"title": "附录D XX", "lines": ["...", "..."]}
    ]
  }
"""
import re
import json
import argparse
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BODY_FONT = "宋体"
HEADING_FONT = "黑体"
ASCII_FONT = "Times New Roman"


def set_run_font(run, font_name=BODY_FONT, size=12, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.font.bold = bold


def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def add_toc(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("目录")
    set_run_font(run, HEADING_FONT, 16, bold=True)

    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fld_char1 = OxmlElement('w:fldChar')
    fld_char1.set(qn('w:fldCharType'), 'begin')
    instr_text = OxmlElement('w:instrText')
    instr_text.set(qn('xml:space'), 'preserve')
    instr_text.text = ' TOC \\o "1-3" \\h \\z \\u '
    fld_char2 = OxmlElement('w:fldChar')
    fld_char2.set(qn('w:fldCharType'), 'separate')
    placeholder = OxmlElement('w:t')
    placeholder.text = '（右键「更新域」生成目录）'
    fld_char3 = OxmlElement('w:fldChar')
    fld_char3.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)
    run._r.append(placeholder)
    run._r.append(fld_char3)


def add_footer_page_numbers(doc):
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld_char1 = OxmlElement('w:fldChar')
    fld_char1.set(qn('w:fldCharType'), 'begin')
    instr_text = OxmlElement('w:instrText')
    instr_text.set(qn('xml:space'), 'preserve')
    instr_text.text = ' PAGE '
    fld_char2 = OxmlElement('w:fldChar')
    fld_char2.set(qn('w:fldCharType'), 'separate')
    placeholder = OxmlElement('w:t')
    placeholder.text = '1'
    fld_char3 = OxmlElement('w:fldChar')
    fld_char3.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)
    run._r.append(placeholder)
    run._r.append(fld_char3)
    set_run_font(run, ASCII_FONT, 10)


def parse_md_table(blk):
    lines = [l.strip() for l in blk.strip().split("\n") if l.strip()]
    if len(lines) < 2:
        return None
    if not (lines[0].startswith("|") and lines[1].startswith("|")):
        return None

    def split_row(l):
        l = l.strip()
        if l.startswith("|"):
            l = l[1:]
        if l.endswith("|"):
            l = l[:-1]
        return [c.strip() for c in l.split("|")]

    sep = split_row(lines[1])
    if not all(re.match(r"^:?-{2,}:?$", c) for c in sep if c != ""):
        return None
    rows = [split_row(lines[0])]
    for l in lines[2:]:
        if l.startswith("|"):
            rows.append(split_row(l))
    return rows


def add_markdown_block(doc, blk):
    blk = blk.strip("\n")
    if not blk.strip():
        return
    lines = blk.split("\n")
    first = lines[0]

    if first.startswith("# "):
        p = doc.add_heading(first[2:].strip(), level=1)
        for run in p.runs:
            set_run_font(run, HEADING_FONT, 16, bold=True)
    elif first.startswith("## "):
        p = doc.add_heading(first[3:].strip(), level=2)
        for run in p.runs:
            set_run_font(run, HEADING_FONT, 14, bold=True)
    elif first.startswith("### "):
        p = doc.add_heading(first[4:].strip(), level=3)
        for run in p.runs:
            set_run_font(run, HEADING_FONT, 12, bold=True)
    elif first.startswith("#### "):
        p = doc.add_paragraph()
        run = p.add_run(first[5:].strip())
        set_run_font(run, HEADING_FONT, 12, bold=True)
    elif first.startswith(">"):
        joined = " ".join(l.lstrip("> ").strip() for l in lines)
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        run = p.add_run(joined.strip())
        set_run_font(run, BODY_FONT, 12)
    elif first.startswith("```"):
        if blk.count("```") >= 2:
            end = blk.rfind("```")
            inner = blk[3:end].strip()
            for line in inner.split("\n"):
                p = doc.add_paragraph()
                run = p.add_run(line)
                set_run_font(run, "Courier New", 10)
    elif first.startswith("|"):
        rows = parse_md_table(blk)
        if rows:
            ncol = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=ncol)
            table.style = 'Table Grid'
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    text = rows[i][j] if j < len(rows[i]) else ""
                    cell.text = text
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            set_run_font(run, BODY_FONT, 10, bold=(i == 0))
                    if i == 0:
                        set_cell_shading(cell, "D9D9D9")
        else:
            p = doc.add_paragraph()
            run = p.add_run(blk.replace("\n", "  |  "))
            set_run_font(run, BODY_FONT, 12)
    else:
        joined = "".join(lines)
        p = doc.add_paragraph()
        run = p.add_run(joined)
        set_run_font(run, BODY_FONT, 12)


def build_docx(title, subtitle, author, chapter_paths, appendices, out_path,
               toc=False, page_numbers=False):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    styles = doc.styles
    for style_name in ['Normal', 'Body Text']:
        try:
            style = styles[style_name]
            style.font.name = ASCII_FONT
            style._element.rPr.rFonts.set(qn('w:eastAsia'), BODY_FONT)
            style.font.size = Pt(12)
        except KeyError:
            pass

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    set_run_font(run, HEADING_FONT, 22, bold=True)

    if subtitle:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(subtitle)
        set_run_font(run, BODY_FONT, 14)

    if toc:
        add_toc(doc)

    for p in chapter_paths:
        with open(p, encoding="utf-8") as f:
            txt = f.read()
        for blk in re.split(r"\n\s*\n", txt):
            add_markdown_block(doc, blk)

    for ap_title, ap_lines in appendices:
        p = doc.add_heading(ap_title, level=1)
        for run in p.runs:
            set_run_font(run, HEADING_FONT, 16, bold=True)
        for ln in ap_lines:
            p = doc.add_paragraph()
            run = p.add_run(ln)
            set_run_font(run, BODY_FONT, 12)

    if page_numbers:
        add_footer_page_numbers(doc)

    doc.core_properties.title = title
    doc.core_properties.author = author or ""

    doc.save(out_path)
    print("[OK] docx written -> %s" % out_path)


TODO_PATTERNS = [
    ("待补充", r"【待补充[^】]{0,200}】"),
    ("待确认", r"【待确认[^】]{0,200}】"),
    ("需核实", r"【需核实[^】]{0,200}】"),
]


def collect_todo(chapter_paths):
    found = {kind: [] for kind, _ in TODO_PATTERNS}
    for p in chapter_paths:
        with open(p, encoding="utf-8") as f:
            txt = f.read()
        for kind, pat in TODO_PATTERNS:
            for m in re.finditer(pat, txt):
                t = m.group(0)
                if t not in found[kind]:
                    found[kind].append(t)
    lines = []
    for kind, _ in TODO_PATTERNS:
        if found[kind]:
            lines.append("【%s】" % kind)
            lines.extend("• " + t for t in found[kind])
    return [("附录B 空缺事项汇总", lines or ["（无）"])]


def main():
    ap = argparse.ArgumentParser(description="Markdown 章节 → Word docx 组装器")
    ap.add_argument("--manifest", help="manifest.json 路径（优先）")
    ap.add_argument("--title", default="文档标题")
    ap.add_argument("--subtitle", default="")
    ap.add_argument("--author", default="编制组")
    ap.add_argument("--chapters", nargs="+", help="有序章节 md 文件列表")
    ap.add_argument("--out", default="output.docx")
    ap.add_argument("--toc", action="store_true", help="在封面后插入目录域（Word 中更新域生成）")
    ap.add_argument("--page-numbers", action="store_true", help="页脚插入页码")
    args = ap.parse_args()

    source_docs = []
    refs = []
    extra = []
    toc = args.toc
    page_numbers = args.page_numbers

    if args.manifest:
        with open(args.manifest, encoding="utf-8") as f:
            m = json.load(f)
        title = m.get("title", args.title)
        subtitle = m.get("subtitle", args.subtitle)
        author = m.get("author", args.author)
        chapters = m.get("chapters", [])
        source_docs = m.get("source_docs", [])
        refs = m.get("refs", [])
        extra = m.get("extra_appendices", [])
        out = m.get("out", args.out)
        toc = m.get("toc", args.toc)
        page_numbers = m.get("page_numbers", args.page_numbers)
    else:
        title, subtitle, author = args.title, args.subtitle, args.author
        chapters = args.chapters or []
        out = args.out

    if not chapters:
        ap.error("必须提供 --chapters 或 manifest.chapters")

    appendices = []
    if source_docs:
        appendices.append(("附录A 依据文件与资料清单", source_docs))
    appendices.extend(collect_todo(chapters))
    if refs:
        appendices.append(("附录C 参考文献", refs))
    appendices.extend((e["title"], e["lines"]) for e in extra)

    build_docx(title, subtitle, author, chapters, appendices, out,
               toc=toc, page_numbers=page_numbers)


if __name__ == "__main__":
    main()
