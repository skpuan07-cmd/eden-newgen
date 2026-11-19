"""
Generate a comprehensive .docx documenting the skpuan07-cmd/eden repository:
system structure, data structure, database schema, AI modules, modifications.
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# Configuration
# ============================================================
OUT_PATH = "/home/z/my-project/download/Sahana_Eden_skpuan07_Repo_Analysis.docx"

FONT_THAI    = "Loma"
FONT_HEADING = "Loma"
FONT_BODY    = "Loma"
FONT_CODE    = "Liberation Mono"

# Colors
C_PRIMARY    = RGBColor(0x1E, 0x3A, 0x8A)
C_ACCENT     = RGBColor(0x0E, 0x74, 0x90)
C_DARK       = RGBColor(0x0F, 0x17, 0x2A)
C_SECOND     = RGBColor(0x47, 0x55, 0x69)
C_CODE_BG    = "F1F5F9"
C_CODE_BORDER= "94A3B8"
C_TABLE_HEAD = "1E3A8A"
C_TABLE_ALT  = "EFF6FF"
C_NOTE_BG    = "FEF3C7"
C_NOTE_BORDER= "D97706"
C_AI_BG      = "EDE9FE"   # light purple for AI (used as bg only)
C_AI_BORDER  = "7C3AED"   # purple (string for shading); for run color use RGBColor
C_RED        = RGBColor(0xDC, 0x26, 0x26)
C_GREEN      = RGBColor(0x16, 0xA3, 0x4A)

# ============================================================
# Helpers (re-used pattern)
# ============================================================
def set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            spec = kwargs[edge]
            el = tcBorders.find(qn(f"w:{edge}"))
            if el is None:
                el = OxmlElement(f"w:{edge}")
                tcBorders.append(el)
            el.set(qn("w:val"), spec.get("val", "single"))
            el.set(qn("w:sz"), str(spec.get("sz", 8)))
            el.set(qn("w:color"), spec.get("color", "auto"))

def set_paragraph_shading(paragraph, hex_color):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)

def set_paragraph_border(paragraph, **kwargs):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            spec = kwargs[edge]
            el = pBdr.find(qn(f"w:{edge}"))
            if el is None:
                el = OxmlElement(f"w:{edge}")
                pBdr.append(el)
            el.set(qn("w:val"), spec.get("val", "single"))
            el.set(qn("w:sz"), str(spec.get("sz", 8)))
            el.set(qn("w:color"), spec.get("color", "auto"))
            el.set(qn("w:space"), str(spec.get("space", 4)))

def add_run(paragraph, text, font_name=FONT_BODY, size_pt=10.5,
            bold=False, italic=False, color=None):
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    return run

def set_paragraph_spacing(paragraph, before_pt=0, after_pt=6, line_pt=None):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    if line_pt is not None:
        pf.line_spacing = Pt(line_pt)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    else:
        pf.line_spacing = 1.3

def add_para(doc, text="", size=10.5, bold=False, italic=False,
             color=None, align=None, before=0, after=6, line_pt=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    set_paragraph_spacing(p, before_pt=before, after_pt=after, line_pt=line_pt)
    if text:
        add_run(p, text, font_name=FONT_BODY, size_pt=size,
                bold=bold, italic=italic, color=color)
    return p

def add_heading_h1(doc, text, num=None):
    full = f"บทที่ {num}   {text}" if num else text
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.keep_with_next = True
    for r in list(p.runs):
        r.text = ""
    add_run(p, full, font_name=FONT_HEADING, size_pt=18,
            bold=True, color=C_PRIMARY)
    set_paragraph_border(p, bottom={"val": "single", "sz": 12, "color": "1E3A8A", "space": 4})

def add_heading_h2(doc, text):
    p = doc.add_paragraph(style="Heading 2")
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    for r in list(p.runs):
        r.text = ""
    add_run(p, text, font_name=FONT_HEADING, size_pt=14, bold=True, color=C_DARK)

def add_heading_h3(doc, text):
    p = doc.add_paragraph(style="Heading 3")
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    for r in list(p.runs):
        r.text = ""
    add_run(p, text, font_name=FONT_HEADING, size_pt=12, bold=True, color=C_ACCENT)

def add_code_block(doc, code, language="python"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.left_indent = Cm(0.4)
    p.paragraph_format.right_indent = Cm(0.4)
    set_paragraph_shading(p, C_CODE_BG)
    set_paragraph_border(p,
        left={"val": "single", "sz": 24, "color": "7C3AED", "space": 6},
        top={"val": "single", "sz": 4,  "color": "94A3B8", "space": 2},
        bottom={"val": "single", "sz": 4,  "color": "94A3B8", "space": 2},
        right={"val": "single", "sz": 4,  "color": "94A3B8", "space": 2},
    )
    lines = code.rstrip("\n").split("\n")
    for i, line in enumerate(lines):
        r = p.add_run(line)
        r.font.name = FONT_CODE
        r.font.size = Pt(9)
        rPr = r._r.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"), FONT_CODE)
        rFonts.set(qn("w:hAnsi"), FONT_CODE)
        rFonts.set(qn("w:cs"), FONT_CODE)
        rFonts.set(qn("w:eastAsia"), FONT_CODE)
        if i < len(lines) - 1:
            r.add_break()
    tag = doc.add_paragraph()
    tag.paragraph_format.space_before = Pt(0)
    tag.paragraph_format.space_after = Pt(8)
    set_paragraph_spacing(tag, before_pt=0, after_pt=8)
    add_run(tag, f"[ {language} ]", font_name=FONT_BODY, size_pt=8,
            italic=True, color=C_SECOND)
    tag.paragraph_format.left_indent = Cm(0.4)

def add_note_box(doc, title, body, color=C_NOTE_BORDER, bg=C_NOTE_BG):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.right_indent = Cm(0.3)
    set_paragraph_shading(p, bg)
    set_paragraph_border(p,
        left={"val": "single", "sz": 24, "color": color, "space": 6},
        top={"val": "single", "sz": 4, "color": color, "space": 2},
        bottom={"val": "single", "sz": 4, "color": color, "space": 2},
        right={"val": "single", "sz": 4, "color": color, "space": 2},
    )
    add_run(p, title + "  ", font_name=FONT_HEADING, size_pt=10,
            bold=True, color=RGBColor.from_string(color))
    add_run(p, body, font_name=FONT_BODY, size_pt=10, color=C_DARK)

def add_bullet(doc, text, level=0, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.7 + level*0.6)
    p.paragraph_format.space_after = Pt(3)
    set_paragraph_spacing(p, before_pt=0, after_pt=3)
    add_run(p, text, font_name=FONT_BODY, size_pt=size, color=C_DARK)

def add_table(doc, headers, rows, col_widths_cm=None, header_bg=C_TABLE_HEAD,
              alt_bg=C_TABLE_ALT, font_size=9.5, header_size=9.5,
              first_col_bold=False):
    n_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)

    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_shading(cell, header_bg)
        set_cell_border(cell,
            top={"val": "single", "sz": 6, "color": header_bg},
            bottom={"val": "single", "sz": 6, "color": header_bg},
            left={"val": "single", "sz": 4, "color": header_bg},
            right={"val": "single", "sz": 4, "color": header_bg})
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, h, font_name=FONT_HEADING, size_pt=header_size,
                bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if r_idx % 2 == 1:
                set_cell_shading(cell, alt_bg)
            set_cell_border(cell,
                top={"val": "single", "sz": 4, "color": "CBD5E1"},
                bottom={"val": "single", "sz": 4, "color": "CBD5E1"},
                left={"val": "single", "sz": 4, "color": "CBD5E1"},
                right={"val": "single", "sz": 4, "color": "CBD5E1"})
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.2
            lines = cell_text.split("\n") if isinstance(cell_text, str) else [str(cell_text)]
            for li, line in enumerate(lines):
                if li > 0:
                    p.add_run().add_break()
                is_first_col = c_idx == 0
                add_run(p, line, font_name=FONT_BODY, size_pt=font_size,
                        bold=is_first_col and first_col_bold,
                        color=C_DARK if not (is_first_col and first_col_bold) else C_PRIMARY)
    trPr = hdr._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader)
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        cantSplit = OxmlElement("w:cantSplit")
        trPr.append(cantSplit)

    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(4)
    return table

# ============================================================
# Document construction
# ============================================================
doc = Document()

section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width  = Cm(21.0)
section.top_margin    = Cm(2.2)
section.bottom_margin = Cm(2.2)
section.left_margin   = Cm(2.3)
section.right_margin  = Cm(2.3)
section.header_distance = Cm(1.0)
section.footer_distance = Cm(1.0)

normal = doc.styles["Normal"]
normal.font.name = FONT_BODY
normal.font.size = Pt(10.5)
normal.paragraph_format.line_spacing = 1.3
normal.paragraph_format.space_after = Pt(6)
rPr = normal.element.get_or_add_rPr()
rFonts = rPr.find(qn("w:rFonts"))
if rFonts is None:
    rFonts = OxmlElement("w:rFonts")
    rPr.append(rFonts)
rFonts.set(qn("w:ascii"), FONT_BODY)
rFonts.set(qn("w:hAnsi"), FONT_BODY)
rFonts.set(qn("w:cs"), FONT_BODY)
rFonts.set(qn("w:eastAsia"), FONT_BODY)

# ============================================================
# COVER
# ============================================================
for _ in range(3):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, after_pt=0)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
add_run(p, "REPOSITORY ANALYSIS DOCUMENT", font_name=FONT_HEADING,
        size_pt=11, bold=True, color=C_ACCENT)

pdiv = doc.add_paragraph()
pdiv.alignment = WD_ALIGN_PARAGRAPH.CENTER
pdiv.paragraph_format.space_after = Pt(30)
set_paragraph_border(pdiv, bottom={"val": "single", "sz": 18, "color": C_TABLE_HEAD, "space": 1})

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
add_run(p, "Sahana Eden", font_name=FONT_HEADING, size_pt=38,
        bold=True, color=C_PRIMARY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(8)
add_run(p, "วิเคราะห์โครงสร้างระบบ โครงสร้างข้อมูล และฐานข้อมูล",
        font_name=FONT_HEADING, size_pt=20, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
add_run(p, "พร้อม AI Modules Extension Analysis",
        font_name=FONT_BODY, size_pt=13, italic=True, color=RGBColor(0x7C, 0x3A, 0xED))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(40)
add_run(p, "Repository: github.com/skpuan07-cmd/eden  (Eden v6.2 fork + AI modules)",
        font_name=FONT_BODY, size_pt=11, color=C_SECOND)

info = doc.add_table(rows=6, cols=2)
info.alignment = WD_TABLE_ALIGNMENT.CENTER
info.autofit = False
info_data = [
    ("Repository",   "https://github.com/skpuan07-cmd/eden.git"),
    ("Fork of",      "Sahana Eden v6.2 (Eden ASP branch)"),
    ("Base commit",  "5b8b8e4ed (2026-09-08) — upstream Sahana Eden"),
    ("Latest commit","63ad591 (Merge PR #3 feature/ai-modules)"),
    ("Total files",  "8,114 (vs 8,502 standard Eden — delta: AI additions)"),
    ("Repo size",    "140 MB"),
]
for i, (k, v) in enumerate(info_data):
    row = info.rows[i]
    row.cells[0].width = Cm(4.5)
    row.cells[1].width = Cm(11.0)
    for j, txt in enumerate((k, v)):
        cell = row.cells[j]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_shading(cell, "F8FAFC" if j == 0 else "FFFFFF")
        set_cell_border(cell,
            top={"val": "single", "sz": 4, "color": "CBD5E1"},
            bottom={"val": "single", "sz": 4, "color": "CBD5E1"},
            left={"val": "single", "sz": 4, "color": "CBD5E1"},
            right={"val": "single", "sz": 4, "color": "CBD5E1"})
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        add_run(p, txt, font_name=FONT_BODY, size_pt=10.5,
                bold=(j == 0), color=C_PRIMARY if j == 0 else C_DARK)

for _ in range(2):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(30)
add_run(p, "วันที่ศึกษา: 19 พฤศจิกายน 2026  ·  วัตถุประสงค์: เพื่อทำงานต่อไป",
        font_name=FONT_BODY, size_pt=10, italic=True, color=C_SECOND)

doc.add_page_break()

# ============================================================
# TOC
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_after = Pt(16)
set_paragraph_border(p, bottom={"val": "single", "sz": 12, "color": C_TABLE_HEAD, "space": 4})
add_run(p, "สารบัญ", font_name=FONT_HEADING, size_pt=20, bold=True, color=C_PRIMARY)

toc_items = [
    ("บทที่ 1",  "ภาพรวม Repository และความแตกต่างจาก Eden มาตรฐาน"),
    ("บทที่ 2",  "สถาปัตยกรรมระบบ (Base Eden)"),
    ("บทที่ 3",  "โครงสร้างไดเรกทอรี"),
    ("บทที่ 4",  "AI Modules — โครงสร้างและฟีเจอร์ที่เพิ่มเข้ามา"),
    ("บทที่ 5",  "โครงสร้างฐานข้อมูล — Standard Eden Tables (388 ตาราง)"),
    ("บทที่ 6",  "โครงสร้างฐานข้อมูล — AI Module Tables (9 ตารางใหม่)"),
    ("บทที่ 7",  "Data Model — การนิยามตารางและความสัมพันธ์"),
    ("บทที่ 8",  "Templates และ Pre-populate Data"),
    ("บทที่ 9",  "Modifications ที่แก้ไขจาก Eden มาตรฐาน (Patches)"),
    ("บทที่ 10", "Endpoints Reference — URL ทั้งหมดของ AI Modules"),
    ("บทที่ 11", "PDPA Compliance & Security"),
    ("บทที่ 12", "การติดตั้งและรันระบบ"),
    ("บทที่ 13", "สรุปและขั้นตอนการทำงานต่อ"),
]
for num, title in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(15.5), WD_ALIGN_PARAGRAPH.RIGHT, 2)
    add_run(p, num + "  ", font_name=FONT_HEADING, size_pt=11, bold=True, color=C_PRIMARY)
    add_run(p, title, font_name=FONT_BODY, size_pt=11, color=C_DARK)

# ============================================================
# บทที่ 1: ภาพรวม Repository
# ============================================================
add_heading_h1(doc, "ภาพรวม Repository และความแตกต่างจาก Eden มาตรฐาน", num=1)

add_para(doc,
    "Repository นี้เป็น fork ของ Sahana Eden v6.2 (Eden ASP branch) ที่ผู้พัฒนา "
    "skpuan07-cmd ได้ต่อยอดโดยการเพิ่ม AI modules ทั้งหมด 10 ตัว เพื่อใช้งานใน "
    "บริบทการจัดการภัยพิบัติในประเทศไทย โดยใช้ z-ai-web-dev-sdk (GLM-4-plus / GLM-5.3-flash) "
    "เป็น LLM หลัก และ OpenThai 2.1 (ของ iApp Technology) เป็น OCR engine เฉพาะ "
    "สำหรับบัตรประชาชนไทย การวิเคราะห์ครั้งนี้ทำเพื่อให้ผู้พัฒนาคนถัดไปเข้าใจ "
    "โครงสร้างทั้งหมดของระบบและสามารถทำงานต่อได้ทันที",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "1.1  Git History — 7 commits ทั้งหมด")

add_table(doc,
    headers=["Commit", "วันที่", "คำอธิบาย"],
    rows=[
        ["63ad591", "2026-11-18", "Merge PR #3: feature/ai-modules (docs)"],
        ["cc91268", "2026-11-18", "docs(ai): add AI modules documentation (5 ไฟล์ .md)"],
        ["4a36e09", "2026-11-18", "Merge PR #2: feature/ai-modules"],
        ["7720f12", "2026-11-18", "Add AI modules: chatbot, OCR, analytics, imagery, translate, NLQ, KB, voice, moderation"],
        ["666a2d2", "2026-11-17", "Merge PR #1: codespace-refactored-space-system"],
        ["c30e8d8", "2026-11-17", "Pending changes exported from your codespace"],
        ["f350471", "2026-09-08", "Initial commit (snapshot ของ Eden v6.2 จาก upstream)"],
    ],
    col_widths_cm=[2.0, 2.5, 11.5],
    first_col_bold=True)

add_heading_h2(doc, "1.2  สิ่งที่เพิ่มเข้ามาใน Fork นี้")

add_table(doc,
    headers=["ประเภท", "จำนวน", "รายละเอียด"],
    rows=[
        ["AI Controllers (ใหม่)",         "10 ไฟล์",  "ai.py + ai_chat.py + ai_ocr.py + ai_analytics.py + ai_imagery.py + ai_translate.py + ai_query.py + ai_kb.py + ai_voice.py + ai_moderation.py"],
        ["AI Views (ใหม่)",                "10 ไฟล์",  "views/ai*/index.html — UI สำหรับแต่ละ controller"],
        ["AI Database Tables (ใหม่)",      "9 ตาราง", "ai_chat_log, ai_knowledge_chunk, ai_ocr_job, ai_translation_job, ai_moderation_queue, ai_forecast, ai_query_log, ai_voice_log, ai_imagery_job"],
        ["AI Framework Modules (ใหม่)",    "2 ไฟล์",   "modules/s3/ai_client.py (wrapper รอบ z-ai CLI) + modules/s3/ai_ocr_openthai.py (iApp OCR client)"],
        ["AI Model Definition (ใหม่)",     "1 ไฟล์",   "models/zz_ai_tables.py — นิยามตาราง AI ทั้ง 9"],
        ["Documentation (ใหม่)",            "5 ไฟล์",   "docs/ai/{README, INSTALL_SUMMARY, SYSTEM_DOCUMENTATION, RECOMMENDATIONS, index}.md (~2,483 บรรทัด)"],
        ["Modified files",                  "4 ไฟล์",   "models/00_settings.py, modules/core/aaa/auth.py, modules/templates/default/config.py, views/default/index.html"],
    ],
    col_widths_cm=[3.5, 2.0, 10.5],
    first_col_bold=True)

add_heading_h2(doc, "1.3  เปรียบเทียบกับ Eden มาตรฐาน")

add_table(doc,
    headers=["รายการ", "Eden มาตรฐาน v6.2", "Fork นี้", "Delta"],
    rows=[
        ["Total files",       "8,502",  "8,114",  "-388 (ลบไฟล์ที่ไม่จำเป็น เช่น .git/)"],
        ["Controllers",       "44",     "54",     "+10 (AI modules)"],
        ["Database tables",   "388",    "397",    "+9 (AI tables)"],
        ["Modules (s3db/)",   "41",     "41",     "0 (ไม่แก้ data model เดิม)"],
        ["Languages",         "21",     "21",     "0 (รองรับ ar, bs, fr, de, es, ... ฯลฯ)"],
        ["Templates",         "default + อื่นๆ", "default", "ใช้ template default เป็นหลัก"],
    ],
    col_widths_cm=[3.5, 3.5, 3.5, 5.5],
    first_col_bold=True)

# ============================================================
# บทที่ 2: สถาปัตยกรรม
# ============================================================
add_heading_h1(doc, "สถาปัตยกรรมระบบ (Base Eden)", num=2)

add_heading_h2(doc, "2.1  Stack เทคโนโลยี")

add_table(doc,
    headers=["ชั้น", "เทคโนโลยี", "เวอร์ชั่น / รายละเอียด"],
    rows=[
        ["Web Framework",       "web2py",                "3.3.3-stable (commit 9005306)"],
        ["Language",            "Python",                "3.11+ (tested on 3.12)"],
        ["ORM",                 "PyDAL",                 "submodule ของ web2py"],
        ["Template engine",     "YATL",                  "built-in ของ web2py"],
        ["Database",            "SQLite (default)",       "รองรับ PostgreSQL/MySQL ด้วย"],
        ["HTTP Server",         "Rocket3",                "built-in ของ web2py"],
        ["Reverse Proxy",       "Caddy",                  "สำหรับ HTTPS termination (ในเอกสาร deploy)"],
        ["Frontend",            "jQuery + Foundation 5",  "OpenLayers, ExtJS, DataTables, D3.js"],
        ["AI Engine (ใหม่)",    "z-ai-web-dev-sdk",      "GLM-4-plus / GLM-5.3-flash (LLM, Vision, ASR, TTS)"],
        ["Thai OCR (ใหม่)",     "OpenThai 2.1 (iApp)",   "สำหรับบัตรประชาชนไทย — ต้องมี IAPP_API_KEY"],
        ["Forecast (ใหม่)",     "Facebook Prophet",       "pip install prophet pandas"],
    ],
    col_widths_cm=[3.0, 4.5, 8.5],
    first_col_bold=True)

add_heading_h2(doc, "2.2  สถาปัตยกรรมการ Deploy (จาก SYSTEM_DOCUMENTATION.md)")

add_code_block(doc, """┌─────────────────┐    HTTPS    ┌─────────────────┐    HTTP    ┌─────────────────┐
│                 │ ─────────→ │                 │ ────────→ │                 │
│   Web Browser   │            │   Caddy (port   │           │  web2py Rocket  │
│   (user)        │ ←───────── │   81, TLS)      │ ←──────── │  (port 3000)    │
│                 │            │                 │           │                 │
└─────────────────┘            └─────────────────┘           └─────────────────┘
                                                                     │
                                                                     │ SQLite
                                                                     ▼
                                                            ┌─────────────────┐
                                                            │  storage.db     │
                                                            │  (file-based)   │
                                                            └─────────────────┘

URL pattern:  https://preview-chat-<chat_id>.space-z.ai/eden/...""", language="text")

add_heading_h2(doc, "2.3  S3 Framework Layers (สืบทอดจาก Eden มาตรฐาน)")

add_para(doc,
    "Fork นี้ใช้ S3 framework เดิมทั้งหมด — DataModel, CRUDResource, CRUDRequest, "
    "BasicCRUD, RESTful ฯลฯ โดยไม่มีการแก้ไข core framework "
    "สิ่งที่เพิ่มเข้ามาเป็นเพียง extension layer ที่อยู่เหนือ S3 framework "
    "ทำให้สามารถ merge กับ upstream ได้ง่ายในอนาคต",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_table(doc,
    headers=["Layer", "Components", "หน้าที่"],
    rows=[
        ["Foundation",       "web2py + PyDAL",                    "HTTP routing, ORM, session, auth"],
        ["S3 Framework",     "DataModel, CRUDResource, CRUDRequest", "Lazy model loading, REST dispatch, hooks"],
        ["Eden Modules",     "41 s3db/*.py files",                 "domain data models (org, pr, inv, hrm, ...)"],
        ["AI Extension",     "controllers/ai*.py + modules/s3/ai_*", "AI layer ที่เพิ่มใน fork นี้"],
        ["Templates",        "modules/templates/default/",         "deployment config + prepopulate data"],
        ["Presentation",     "views/*.html + static/",              "YATL templates + jQuery + Foundation"],
    ],
    col_widths_cm=[3.0, 5.5, 7.5],
    first_col_bold=True)

# ============================================================
# บทที่ 3: โครงสร้างไดเรกทอรี
# ============================================================
add_heading_h1(doc, "โครงสร้างไดเรกทอรี", num=3)

add_code_block(doc, """eden/
├── ABOUT                                # คำอธิบายโปรเจกต์
├── LICENSE                              # MIT license
├── README.md                            # คำอธิบายสั้น
├── VERSION                              # 6.2-1-g5b8b8e4ed
├── _version.py                          # __version__ = "6.2"
├── __init__.py
├── requirements.txt                     # python-dateutil, lxml, requests
├── optional_requirements.txt            # openpyxl, geopy, shapely, Pillow, ...
│
├── controllers/                         # 54 ไฟล์ (44 มาตรฐาน + 10 AI)
│   ├── org.py, pr.py, inv.py, hrm.py    # ← Eden modules มาตรฐาน
│   ├── project.py, req.py, asset.py     # ← Eden modules มาตรฐาน
│   ├── cap.py, cms.py, dvr.py, msg.py   # ← Eden modules มาตรฐาน
│   ├── ai.py                            # ★ AI landing page (ใหม่)
│   ├── ai_chat.py                       # ★ AI Chatbot (ใหม่)
│   ├── ai_ocr.py                        # ★ AI OCR + data extraction (ใหม่)
│   ├── ai_analytics.py                  # ★ AI Predictive Analytics (ใหม่)
│   ├── ai_imagery.py                    # ★ AI Image Recognition (ใหม่)
│   ├── ai_translate.py                  # ★ AI Translation (ใหม่)
│   ├── ai_query.py                      # ★ AI Natural Language Query (ใหม่)
│   ├── ai_kb.py                         # ★ AI Knowledge Base admin (ใหม่)
│   ├── ai_voice.py                      # ★ AI Voice Interface (ใหม่)
│   └── ai_moderation.py                 # ★ AI Content Moderation (ใหม่)
│
├── models/                              # web2py model run-order
│   ├── 000_1st_run.py                   # import core as s3base
│   ├── 00_db.py                         # current.db = DAL(...)
│   ├── 00_settings.py                   # ★ MODIFIED: HTTPS session secure cookie + Thai lang
│   ├── 00_tables.py
│   ├── 00_utils.py                      # current.crud_controller = crud_controller
│   ├── tasks.py                         # scheduler tasks (maintenance, ...)
│   ├── zz_ai_tables.py                  # ★ NEW: 9 AI database tables (zz_ = runs last)
│   ├── zz_last.py
│   └── zzz_1st_run.py
│
├── modules/                             # Python framework code
│   ├── s3cfg.py                         # S3Config (deployment_settings)
│   ├── s3dal.py                         # PyDAL compat shim
│   ├── s3db/                            # 41 DataModel modules (org, pr, inv, ...)
│   ├── s3/                              # ★ NEW: AI-specific s3 subpackage
│   │   ├── ai_client.py                 # ★ AIClient wrapper รอบ z-ai CLI
│   │   └── ai_ocr_openthai.py           # ★ OpenThai 2.1 (iApp) OCR client
│   ├── core/                            # S3 framework (controller, model, resource, ...)
│   │   ├── controller.py                # CRUDRequest
│   │   ├── model/datamodel.py           # DataModel base
│   │   ├── resource/                    # CRUDResource + filter + components
│   │   ├── methods/                     # BasicCRUD, RESTful, S3Report, S3Map, ...
│   │   ├── formats/                     # S3XML, PDF, XLS, SHP writers
│   │   ├── filters/                     # URL → Query parsers
│   │   ├── aaa/                         # ★ MODIFIED: auth.py (str() coercion fix)
│   │   │   └── auth.py
│   │   ├── gis/, msg/, sync/, tools/, ui/
│   │   └── ...
│   └── templates/                       # deployment skins
│       └── default/                     # ★ MODIFIED: config.py (register 'ai' module)
│           ├── __init__.py
│           ├── config.py                # ★ Module registration (line 1346: ai module_type=7)
│           ├── 000_config.py            # template สำหรับสร้าง models/000_config.py
│           ├── Demo/                    # prepopulate CSV data
│           │   ├── auth_roles.csv       # ระบบ system roles (CR_ADMIN, EVENT_ADMIN, ...)
│           │   ├── cr_shelter_type.csv
│           │   ├── security_staff_type.csv
│           │   ├── security_zone_type.csv
│           │   └── vehicle_type.csv
│           ├── DefaultItems.csv
│           ├── DefaultSkillList.csv
│           ├── StandardItems.csv
│           └── ...
│
├── views/                               # YATL templates
│   ├── default/
│   │   └── index.html                   # ★ MODIFIED: AI panel on home page
│   ├── ai/                              # ★ NEW: AI landing page
│   │   └── index.html
│   ├── ai_chat/                         # ★ NEW
│   │   └── index.html
│   ├── ai_ocr/                          # ★ NEW
│   ├── ai_analytics/                    # ★ NEW
│   ├── ai_imagery/                      # ★ NEW
│   ├── ai_translate/                    # ★ NEW
│   ├── ai_query/                        # ★ NEW
│   ├── ai_kb/                           # ★ NEW
│   ├── ai_voice/                        # ★ NEW
│   └── ai_moderation/                   # ★ NEW
│
├── static/                              # JS, CSS, SCSS, images
├── languages/                           # i18n: 21 ภาษา (ar, bs, de, el, es, fa, fr, ...)
├── private/                             # scripts
├── cron/                                # scheduled tasks
├── tests/                               # unit tests
└── docs/                                # developer documentation (Sphinx)
    ├── source/                          # source of readthedocs
    └── ai/                              # ★ NEW: AI modules documentation
        ├── README.md                    # คู่มือ 10 use cases (414 บรรทัด)
        ├── INSTALL_SUMMARY.md           # รายงานการติดตั้ง (138 บรรทัด)
        ├── SYSTEM_DOCUMENTATION.md      # เอกสารระบบฉบับเต็ม (873 บรรทัด)
        ├── RECOMMENDATIONS.md           # คำแนะนำ AI/Cyber/PDPA (977 บรรทัด)
        └── index.md                     # index สำหรับ doc nav (81 บรรทัด)""", language="text")

# ============================================================
# บทที่ 4: AI Modules
# ============================================================
add_heading_h1(doc, "AI Modules — โครงสร้างและฟีเจอร์ที่เพิ่มเข้ามา", num=4)

add_para(doc,
    "Fork นี้เพิ่ม AI modules ทั้งหมด 10 ตัว (Use Cases 1-10 จากเอกสาร RECOMMENDATIONS.md) "
    "ที่ทำงานบน S3 framework เดิมของ Eden โดยใช้ z-ai-web-dev-sdk เป็น AI engine หลัก "
    "และ OpenThai 2.1 สำหรับ Thai OCR เฉพาะ ทุก module มี PDPA consent flow "
    "และ audit logging ในตาราง ai_* ของตัวเอง",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "4.1  ภาพรวม AI Modules ทั้ง 10 ตัว")

add_table(doc,
    headers=["#", "Use Case", "Controller", "DB Table", "Auth", "PDPA Consent"],
    rows=[
        ["1",  "AI Chatbot (RAG)",            "ai_chat.py",       "ai_chat_log",         "anonymous + login",  "บังคับ"],
        ["2",  "AI OCR + Data Extraction",    "ai_ocr.py",        "ai_ocr_job",          "login",              "บังคับ (ข้อมูลในรูปอาจ sensitive)"],
        ["3",  "AI Predictive Analytics",     "ai_analytics.py",  "ai_forecast",         "ADMIN",              "ไม่ (admin ops)"],
        ["4",  "AI Image Recognition",        "ai_imagery.py",    "ai_imagery_job",      "login",              "ไม่"],
        ["5",  "AI Translation (UI strings)", "ai_translate.py",  "ai_translation_job",  "ADMIN",              "ไม่ (UI text เท่านั้น)"],
        ["6",  "AI Natural Language Query",   "ai_query.py",      "ai_query_log",        "ADMIN",              "ไม่ (admin ops)"],
        ["7",  "AI Code Assistance",          "(docs only)",      "—",                   "—",                  "—"],
        ["8",  "RAG Knowledge Base admin",    "ai_kb.py",         "ai_knowledge_chunk",  "ADMIN",              "ไม่"],
        ["9",  "AI Voice Interface",          "ai_voice.py",      "ai_voice_log",        "login",              "ไม่ (transcript ไม่ใช่ PII)"],
        ["10", "AI Content Moderation",       "ai_moderation.py", "ai_moderation_queue", "auto (hook)",        "ไม่"],
        ["11", "AI Menu Integration",         "ai.py",            "—",                   "login",              "—"],
    ],
    col_widths_cm=[0.7, 3.5, 2.5, 3.0, 2.3, 4.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "4.2  AI Client Module (modules/s3/ai_client.py)")

add_para(doc,
    "เป็น Python wrapper รอบ z-ai CLI ที่ใช้ subprocess เรียก binary ที่ /usr/local/bin/z-ai "
    "รองรับ 4 capabilities หลัก: (1) chat — LLM completion, (2) describe_image — Vision, "
    "(3) transcribe_audio — ASR, (4) synthesize_speech — TTS "
    "มีการบังคับใช้ UTF-8 encoding ทุกครั้งเพื่อรองรับภาษาไทย "
    "และมี timeout ที่ 120 วินาที (configurable)",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''class AIClient:
    """Wrapper around the z-ai CLI for Sahana Eden."""

    def __init__(self, timeout=120):
        self.timeout = timeout
        self.cli_path = self._find_cli()

    @staticmethod
    def _find_cli():
        """Find the z-ai CLI binary."""
        for path in ["/usr/local/bin/z-ai", "/usr/bin/z-ai"]:
            if os.path.exists(path):
                return path
        import shutil
        cli = shutil.which("z-ai")
        if not cli:
            raise AIClientError(
                "z-ai CLI not found. Install z-ai-web-dev-sdk "
                "and ensure 'z-ai' is in PATH."
            )
        return cli

    # Main API:
    # client.chat(prompt, system=...) -> str
    # client.describe_image(image_path, prompt=...) -> str
    # client.transcribe_audio(audio_path) -> str
    # client.synthesize_speech(text, output_path=...) -> str''', language="python")

add_heading_h2(doc, "4.3  OpenThai OCR Module (modules/s3/ai_ocr_openthai.py)")

add_para(doc,
    "เป็น Thai-specific OCR client ที่เรียก iApp OpenThai 2.1 API "
    "URL: https://api.iapp.co.th — ต้องมี IAPP_API_KEY env var "
    "รองรับ endpoint เฉพาะบัตรประชาชนไทย: /thai-national-id-card/v3.5/front "
    "ที่ return structured fields โดยตรง (title, first_name_th, last_name_th, "
    "first_name_en, last_name_en, id_card_number, address, birth_date) "
    "สำหรับเอกสารอื่น ๆ จะ fall back ไปใช้ vision LLM แทน",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''OPENTHAI_ENDPOINTS = {
    "id_card": "/thai-national-id-card/v3.5/front",
    # Other doc types → not supported by OpenThai public endpoints yet;
    # controller will route them to vision LLM automatically.
}

class OpenThaiClient:
    def __init__(self, api_key=None, base_url=None, timeout=60):
        self.api_key = api_key or os.environ.get("IAPP_API_KEY") \\
                       or os.environ.get("OPENTHAI_API_KEY")
        if not self.api_key:
            raise OpenThaiError(
                "IAPP_API_KEY environment variable not set. "
                "Get one at https://iapp.co.th/ and export IAPP_API_KEY=..."
            )''', language="python")

# ============================================================
# บทที่ 5: Database — Standard Tables
# ============================================================
add_heading_h1(doc, "โครงสร้างฐานข้อมูล — Standard Eden Tables (388 ตาราง)", num=5)

add_para(doc,
    "ฐานข้อมูลของ Eden มาตรฐานประกอบด้วย 388 ตาราง (สร้างโดย DataModel subclasses "
    "ใน modules/s3db/*.py ทั้ง 41 ไฟล์) บน SQLite (default) หรือ PostgreSQL/MySQL "
    "ทุกตารางมี MetaFields อัตโนมัติ: uuid, mci, deleted, created_on, modified_on, "
    "owned_by_user, owned_by_role ฯลฯ",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "5.1  ตารางหลักในแต่ละ Module")

add_table(doc,
    headers=["Module", "ตารางหลัก", "จำนวน", "หน้าที่"],
    rows=[
        ["org",    "org_organisation, org_organisation_type, org_office, org_facility, org_region",                     "~30",  "องค์กร / สำนักงาน / สาขา"],
        ["pr",     "pr_person, pr_pentity, pr_address, pr_contact, pr_physical_description",                          "~40",  "ทะเบียนบุคคล"],
        ["hrm",    "hrm_human_resource, hrm_skill, hrm_certificate, hrm_training_event, hrm_programme",                "~30",  "บุคลากร / อาสาสมัคร / skill"],
        ["vol",    "vol_volunteer, vol_application, vol_assignment",                                                    "~15",  "อาสาสมัคร"],
        ["inv",    "inv_warehouse, inv_inv_item, inv_inv_track, inv_req_item, inv_send, inv_recv",                      "~40",  "คลังสินค้า / การเบิกจ่าย"],
        ["req",    "req_req, req_req_item, req_commit",                                                                 "~15",  "คำขอ / การจัดหา"],
        ["asset",  "asset_asset, asset_group, asset_item, asset_log",                                                   "~20",  "ครุภัณฑ์ / ทรัพย์สิน"],
        ["vehicle","vehicle_vehicle, vehicle_log, vehicle_status",                                                      "~10",  "ยานพาหนะ"],
        ["project","project_project, project_activity, project_beneficiary, project_location, project_sector",          "~30",  "โครงการ / กิจกรรม"],
        ["cr",     "cr_shelter, cr_shelter_registration, cr_shelter_unit, cr_shelter_type",                             "~15",  "ศูนย์พักพิง"],
        ["cms",    "cms_post, cms_post_series, cms_post_tag, cms_attachment",                                           "~12",  "ระบบเนื้อหา / ข่าวสาร"],
        ["cap",    "cap_alert, cap_info, cap_area, cap_resource",                                                       "~10",  "การแจ้งเตือนภัย (CAP)"],
        ["dvr",    "dvr_case, dvr_case_status, dvr_case_person",                                                        "~20",  "ทะเบียนผู้ประสบภัย (sensitive!)"],
        ["dvi",    "dvi_body, dvi_external, dvi_check_in, dvi_check_out",                                               "~15",  "สืบค้นคนหาย"],
        ["msg",    "msg_log, msg_channel, msg_subscription, msg_outbound",                                              "~12",  "SMS / Email / Twitter"],
        ["gis",    "gis_location, gis_config, gis_layer_feature, gis_layer_openstreetmap, gis_hierarchy",               "~30",  "ระบบสารสนเทศภูมิศาสตร์"],
        ["auth",   "auth_user, auth_group, auth_membership, auth_event, auth_consent, auth_organisation",               "~15",  "authentication / authorization"],
        ["sync",   "sync_repository, sync_log",                                                                          "~5",   "synchronization"],
        ["med",    "med_patient, med_patient_status",                                                                    "~10",  "ข้อมูลผู้ป่วย (sensitive!)"],
        ["budget", "budget_budget, budget_bundle, budget_kit, budget_kit_item",                                         "~10",  "งบประมาณ / kit / bundle"],
        ["อื่นๆ", "event, fire, hms, irs, supply, transport, member, mpr, doc, sit, stats, security, water, ...",       "~50",  "module อื่น ๆ"],
        ["scheduler","scheduler_task, scheduler_run, scheduler_worker, scheduler_task_deps",                           "~4",   "web2py scheduler"],
        ["s3 framework","s3_permission, s3_audit, s3_dashboard, s3_hierarchy, s3_import_job, ...",                      "~15",  "framework infrastructure"],
    ],
    col_widths_cm=[1.8, 8.5, 1.3, 5.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "5.2  ตาราง Sensitive (PDPA)"

)

add_note_box(doc, "ตารางที่เก็บข้อมูลส่วนบุคคล (PDPA-sensitive)",
    "ตารางต่อไปนี้เก็บข้อมูลส่วนบุคคลที่ PDPA บังคับ: "
    "auth_user (email, password hash), pr_person (ชื่อ-นามสกุล, ที่อยู่, เบอร์), "
    "dvr_case (ข้อมูลผู้ประสบภัย — ละเอียดมาก), "
    "med_patient (ข้อมูลสุขภาพ — sensitive สูงสุด), "
    "hrm_human_resource (เลขบัตรประชาชน, ที่อยู่), "
    "msg_log (content ของ SMS/Email), "
    "ai_chat_log (คำถาม-คำตอบ, ต้องมี consent), "
    "ai_ocr_job (รูปเอกสาร + extracted_data, ต้องมี consent), "
    "ai_voice_log (transcript เสียง). "
    "NLQ (Use Case 6) มี whitelist ห้าม query ตารางเหล่านี้",
    color="DC2626", bg="FEE2E2")

# ============================================================
# บทที่ 6: AI Tables
# ============================================================
add_heading_h1(doc, "โครงสร้างฐานข้อมูล — AI Module Tables (9 ตารางใหม่)", num=6)

add_para(doc,
    "ไฟล์ models/zz_ai_tables.py นิยามตาราง AI ทั้ง 9 ตัว โดยใช้ชื่อขึ้นต้นด้วย ai_ "
    "และใช้ zz_ prefix ในชื่อไฟล์เพื่อให้ web2py รันเป็นลำดับสุดท้าย "
    "(หลังจาก 00_*.py ที่สร้าง db, auth แล้ว) ทุกตารางมี base fields คือ "
    "uuid, mci, deleted, created_on, modified_on (mimics Eden's s3base_fields)",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "6.1  ตาราง ai_chat_log (Use Case 1: Chatbot)")

add_table(doc,
    headers=["Field", "Type", "รายละเอียด"],
    rows=[
        ["session_id",         "string(128)", "browser session ID"],
        ["user_id",            "reference auth_user", "NULL ถ้า anonymous"],
        ["question",           "text",        "คำถามของ user (max 1000 chars)"],
        ["answer",             "text",        "คำตอบจาก LLM"],
        ["context_used",       "text (JSON)", "retrieved chunks ที่ใช้ใน RAG"],
        ["sources",            "text (JSON)", "list ของ source URLs/IDs"],
        ["ip",                 "string(64)",  "client IP (สำหรับ rate limit)"],
        ["user_agent",         "string(256)", "browser UA"],
        ["model",              "string(64)",  "LLM model ที่ใช้ (เช่น glm-4-plus)"],
        ["tokens_prompt",      "integer",     "token count ของ prompt"],
        ["tokens_completion",  "integer",     "token count ของ completion"],
        ["latency_ms",         "integer",     "response time"],
        ["rating",             "integer",     "user feedback: -1, 0, +1"],
        ["feedback_comment",   "text",        "ความคิดเห็นเพิ่มเติม"],
        ["language",           "string(8)",   "detected language"],
        ["created_on",         "datetime",    "เวลาที่ถาม"],
        ["consent_given",      "boolean",     "★ PDPA consent flag — บังคับ"],
    ],
    col_widths_cm=[3.5, 3.0, 9.5],
    first_col_bold=True, font_size=9)

add_para(doc,
    "Retention policy: 90 วัน (auto-delete) — Rate limit: 60 requests/minute/IP",
    italic=True, color=C_SECOND)

add_heading_h2(doc, "6.2  ตาราง ai_knowledge_chunk (Use Case 1, 8: RAG)")

add_table(doc,
    headers=["Field", "Type", "รายละเอียด"],
    rows=[
        ["source_type",      "string(32)",   "cms_post, page, wiki, faq"],
        ["source_id",        "integer",      "FK to source table"],
        ["source_url",       "string(512)",  "URL ของ source (สำหรับ external)"],
        ["title",            "string(256)",  "title ของ chunk"],
        ["content",          "text",         "เนื้อหา chunk (chunked จาก original)"],
        ["chunk_index",      "integer",      "ลำดับ chunk ในเอกสารยาว"],
        ["language",         "string(8)",    "default 'en'"],
        ["embedding",        "text (JSON)",  "vector embeddings (list of floats) — สำหรับ pgvector upgrade"],
        ["embedding_model",  "string(64)",   "model ที่ใช้ embed"],
        ["tags",             "list:string",  "tags สำหรับ filter"],
        ["is_public",        "boolean",      "public chunk หรือไม่"],
        ["last_indexed",     "datetime",     "last re-embed timestamp"],
    ],
    col_widths_cm=[3.5, 3.0, 9.5],
    first_col_bold=True, font_size=9)

add_heading_h2(doc, "6.3  ตาราง ai_ocr_job (Use Case 2: OCR)")

add_table(doc,
    headers=["Field", "Type", "รายละเอียด"],
    rows=[
        ["user_id",          "reference auth_user", "user ที่สั่ง OCR"],
        ["document_type",    "string(32)",   "id_card, invoice, shelter_form, victim_report, other"],
        ["uploaded_file",    "upload",       "ไฟล์รูปที่อัปโหลด (ลบหลัง extraction)"],
        ["extracted_data",   "text (JSON)",  "structured output จาก OCR engine"],
        ["target_table",     "string(64)",   "ตาราง Eden ที่จะ insert (เช่น pr_person)"],
        ["target_record_id", "integer",      "record ID ที่ pre-fill"],
        ["status",           "string(16)",   "pending, done, failed"],
        ["error_message",    "text",         "error detail ถ้า failed"],
        ["consent_given",    "boolean",      "★ PDPA consent — บังคับ"],
        ["ip",               "string(64)",   "client IP"],
        ["model",            "string(64)",   "OCR engine (openthai-2.1 / glm-4-plus-vision)"],
    ],
    col_widths_cm=[3.5, 3.5, 9.0],
    first_col_bold=True, font_size=9)

add_heading_h2(doc, "6.4  ตารางที่เหลืออีก 6 ตาราง — สรุปโครงสร้าง")

add_table(doc,
    headers=["ตาราง", "Use Case", "Fields หลัก", "Retention"],
    rows=[
        ["ai_translation_job",    "5",  "source_lang, target_lang, string_key, source_text, translated_text, status (pending/reviewed/approved/rejected), reviewer_id, reviewer_comment, model", "ไม่กำหนด"],
        ["ai_moderation_queue",   "10", "source_table, source_id, content_preview, categories (JSON), scores (JSON), action (pending/approved/rejected/edited), reviewer_id, auto_action, model", "3 ปี"],
        ["ai_forecast",           "3",  "target_table, target_id, metric, forecast_date, forecast_value, confidence_lower, confidence_upper, horizon_days, model_version, generated_on", "ทับด้วย version ใหม่"],
        ["ai_query_log",          "6",  "user_id, question, interpreted_query, executed, result_count, answer_text, model, latency_ms, ip", "1 ปี"],
        ["ai_voice_log",          "9",  "user_id, audio_file, transcript, intent, response_text, response_audio_file, language, ip", "90 วัน"],
        ["ai_imagery_job",        "4",  "user_id, image_url, uploaded_file, imagery_type (satellite/drone/field_photo), analysis_type (flood_extent/building_damage/road_passability), result_json, result_summary, location_id, status, model", "ไม่กำหนด"],
    ],
    col_widths_cm=[3.2, 0.8, 8.5, 3.0],
    first_col_bold=True, font_size=8.5, header_size=9)

# ============================================================
# บทที่ 7: Data Model
# ============================================================
add_heading_h1(doc, "Data Model — การนิยามตารางและความสัมพันธ์", num=7)

add_heading_h2(doc, "7.1  Pattern การนิยาม AI Tables (web2py ตรงไปตรงมา)")

add_para(doc,
    "AI tables ใช้ web2py DAL แบบตรงไปตรงมา (db.define_table) ไม่ใช้ DataModel "
    "subclass ของ S3 framework เพราะ (1) ไม่ต้องการ lazy loading "
    "(2) ไม่ต้องการ super-entity/component graph ที่ซับซ้อน "
    "(3) ต้องการให้ตารางถูกสร้างทุกครั้งที่ first-run "
    "(4) ง่ายต่อการ maintain โดย developer ที่ไม่จำเป็นต้องเข้าใจ S3 framework ลึก",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''# ตัวอย่าง pattern จาก models/zz_ai_tables.py

def _base_fields():
    """Common base fields (mimics Eden's s3base_fields)"""
    from gluon import Field
    from datetime import datetime
    return [
        Field("uuid", length=128),
        Field("mci", "integer", default=0),
        Field("deleted", "boolean", default=False),
        Field("created_on", "datetime", default=datetime.utcnow),
        Field("modified_on", "datetime", update=datetime.utcnow),
    ]

db.define_table(
    "ai_chat_log",
    Field("session_id", length=128),
    Field("user_id", "reference auth_user"),  # NULL if anonymous
    Field("question", "text"),
    Field("answer", "text"),
    Field("context_used", "text"),  # JSON: retrieved chunks
    Field("sources", "text"),  # JSON: list of source URLs/IDs
    Field("consent_given", "boolean", default=False),  # PDPA
    *_base_fields(),
)
db.ai_chat_log.question.requires = IS_NOT_EMPTY()''', language="python")

add_heading_h2(doc, "7.2  ความสัมพันธ์ระหว่าง AI Tables และ Eden Tables")

add_table(doc,
    headers=["AI Table", "Eden Table FK", "ความสัมพันธ์"],
    rows=[
        ["ai_chat_log.user_id",         "auth_user.id",              "user ที่ถาม (NULL ถ้า anonymous)"],
        ["ai_ocr_job.user_id",          "auth_user.id",              "user ที่สั่ง OCR"],
        ["ai_ocr_job.target_table",     "(dynamic)",                 "ชื่อตาราง Eden ที่จะ insert (string, not FK)"],
        ["ai_ocr_job.target_record_id", "(dynamic)",                 "record ID ของ Eden table ที่ pre-fill"],
        ["ai_translation_job.reviewer_id","auth_user.id",             "reviewer ที่ตรวจ translation"],
        ["ai_moderation_queue.reviewer_id","auth_user.id",            "reviewer ที่ตรวจ moderation"],
        ["ai_moderation_queue.source_table","(dynamic)",              "ตาราง Eden ที่ content มาจาก (cms_post, msg_log)"],
        ["ai_moderation_queue.source_id","(dynamic)",                 "record ID ของ source"],
        ["ai_forecast.target_table",    "(dynamic)",                  "ตาราง Eden ที่ forecast (inv_inv_track, sit_presence)"],
        ["ai_forecast.target_id",       "(dynamic)",                  "record ID ของ target"],
        ["ai_imagery_job.location_id",  "gis_location.id (as integer)","พิกัดที่ผูกกับภาพ (integer เพื่อหลีก lazy-table)"],
        ["ai_query_log.user_id",        "auth_user.id",              "admin ที่ถาม"],
        ["ai_voice_log.user_id",        "auth_user.id",              "user ที่ใช้ voice"],
        ["ai_knowledge_chunk.source_type","(dynamic)",                 "ประเภท source: cms_post, page, wiki, faq"],
        ["ai_knowledge_chunk.source_id","(dynamic)",                  "record ID ของ source"],
    ],
    col_widths_cm=[5.5, 4.0, 6.5],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "7.3  การเชื่อมกับ S3 Framework")

add_para(doc,
    "AI controllers ไม่ได้ใช้ CRUDRequest หรือ BasicCRUD ของ S3 framework "
    "แต่ใช้ web2py standard controllers แบบธรรมดา (def index(), def ask(), ...) "
    "เพราะ AI endpoints ไม่ใช่ CRUD pattern แบบเดียวกับ Eden modules "
    "การ query ข้อมูล Eden จาก AI controllers ทำผ่าน current.db โดยตรง "
    "หรือผ่าน REST API call ภายใน (สำหรับ NLQ — Use Case 6)",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================================
# บทที่ 8: Templates + Prepopulate
# ============================================================
add_heading_h1(doc, "Templates และ Pre-populate Data", num=8)

add_heading_h2(doc, "8.1  Template Default (modules/templates/default/)")

add_table(doc,
    headers=["ไฟล์", "หน้าที่"],
    rows=[
        ["__init__.py",        "Python package marker"],
        ["config.py",          "★ MODIFIED — เพิ่ม 'ai' module registration (module_type=7)"],
        ["000_config.py",      "template สำหรับสร้าง models/000_config.py (คัดลอกตอน setup)"],
        ["Demo/auth_roles.csv","system roles pre-populate (CR_ADMIN, EVENT_ADMIN, HMS_ADMIN, MEMBER_ADMIN, ORG_ADMIN)"],
        ["Demo/cr_shelter_type.csv","ประเภทศูนย์พักพิง"],
        ["Demo/security_staff_type.csv","ประเภทเจ้าหน้าที่รักษาความปลอดภัย"],
        ["Demo/security_zone_type.csv","ประเภทเขตรักษาความปลอดภัย"],
        ["Demo/vehicle_type.csv","ประเภทยานพาหนะ"],
        ["Demo/tasks.cfg",     "scheduler task config"],
        ["DefaultItems.csv",   "รายการสินค้าเริ่มต้น"],
        ["DefaultSkillList.csv","รายการ skill เริ่มต้น"],
        ["DefaultSkillCompetency.csv","ระดับ competency"],
        ["DrivingSkillList.csv","ทักษะขับขี่"],
        ["LanguageSkillList.csv","ทักษะภาษา"],
        ["LanguageCompetency.csv","ระดับทักษะภาษา"],
        ["StandardItems.csv",  "รายการสินค้ามาตรฐาน"],
    ],
    col_widths_cm=[5.5, 10.5],
    first_col_bold=True)

add_heading_h2(doc, "8.2  การ Register 'ai' Module ใน Menu (modules/templates/default/config.py)")

add_code_block(doc, '''# บรรทัดที่ 1345-1350 ของ modules/templates/default/config.py

        # AI Assistant module — landing page linking to all 9 AI controllers
        ("ai", Storage(
            name_nice = T("AI Assistant"),
            #description = "AI-powered features: chatbot, OCR, analytics, imagery,
            #               translation, NLQ, voice, moderation",
            module_type = 7,    # 7th item in the menu (top-level, after gis)
        )),''', language="python")

add_heading_h2(doc, "8.3  Module Registration Order (default template)")

add_table(doc,
    headers=["ลำดับ Menu", "Module", "name_nice"],
    rows=[
        ["1",  "org",      "Organizations"],
        ["2",  "hrm, vol", "Staff / Volunteers"],
        ["4",  "inv",      "Warehouses"],
        ["5",  "asset",    "Assets"],
        ["6",  "gis",      "Map"],
        ["7",  "ai ★",     "AI Assistant (ใหม่)"],
        ["10", "pr, cms, doc", "Person Registry / Content / Documents"],
        ["—",  "default, admin, appadmin, errors, sync", "framework/ไม่แสดงใน menu"],
        ["—",  "msg, supply, req, project, stats, vehicle", "sub items หรือไม่แสดง"],
    ],
    col_widths_cm=[2.5, 5.5, 8.0],
    first_col_bold=True)

add_heading_h2(doc, "8.4  System Roles (Demo/auth_roles.csv)")

add_code_block(doc, '''uid,role,description,controller,function,table,uacl,oacl,entity,Notes
ORG_ADMIN,Organisation Administrator,,cr,,,"READ|CREATE|UPDATE|DELETE",,,
ORG_ADMIN,Organisation Administrator,,event,,,"READ|CREATE|UPDATE|DELETE",,,
ORG_ADMIN,Organisation Administrator,,hms,,,"READ|CREATE|UPDATE|DELETE",,,
ORG_ADMIN,Organisation Administrator,,member,,,"READ|CREATE|UPDATE|DELETE",,,
CR_ADMIN,Shelter Administrator,,cr,,,"READ|CREATE|UPDATE|DELETE",,,
EVENT_ADMIN,Event Administrator,,event,,,"READ|CREATE|UPDATE|DELETE",,,
HMS_ADMIN,Hospital Administrator,,hms,,,"READ|CREATE|UPDATE|DELETE",,,
MEMBER_ADMIN,Members Administrator,,member,,,"READ|CREATE|UPDATE|DELETE",,,',
    col_widths_cm=[3.0, 4.0, 4.5],
    first_col_bold=True)''', language="text")

# ============================================================
# บทที่ 9: Modifications
# ============================================================
add_heading_h1(doc, "Modifications ที่แก้ไขจาก Eden มาตรฐาน (Patches)", num=9)

add_para(doc,
    "Fork นี้แก้ไขไฟล์ของ Eden มาตรฐานเพียง 4 ไฟล์ (minimal patch) "
    "เพื่อให้สามารถ merge กับ upstream ได้ง่าย และลด conflict",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "9.1  Patch: models/00_settings.py")

add_para(doc,
    "เพิ่ม 2 ส่วน: (1) รองรับ Thai language ผ่าน locale th_TH "
    "(2) บังคับใช้ secure session cookie เมื่อเข้าผ่าน HTTPS",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''# บรรทัด 117-119 (ใหม่)
elif language == "th":
    s3.language = "th_TH"

# บรรทัด 220-223 (ใหม่)
# Use secure session cookie over HTTPs
if request.is_https:
    session.secure()''', language="python")

add_heading_h2(doc, "9.2  Patch: modules/core/aaa/auth.py")

add_para(doc,
    "แก้ Python 3 compatibility bug: subject/message arguments ของ email "
    "ต้อง coerce เป็น str() ก่อนส่งให้ mailer ไม่งั้น Storage object จะ trigger "
    "'ascii codec can't decode' error เมื่อมี non-ASCII characters",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''# บรรทัด 1633-1634 (เดิม → ใหม่)
- subject = messages.verify_email_subject % system,
- message = messages.verify_email % system,
+ subject = str(messages.verify_email_subject % system),
+ message = str(messages.verify_email % system),

# บรรทัด 1695-1696 (เดิม → ใหม่)
- subject = self.messages.reset_password_subject,
- message = message):
+ subject = str(self.messages.reset_password_subject),
+ message = message):''', language="diff")

add_heading_h2(doc, "9.3  Patch: modules/templates/default/config.py")

add_para(doc,
    "เพิ่ม 'ai' module registration ที่บรรทัดที่ 1345-1350 "
    "ทำให้ AI Assistant ปรากฏใน main menu เป็น item ที่ 7 "
    "(ระหว่าง Map กับ more ▼)",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "9.4  Patch: views/default/index.html")

add_para(doc,
    "เพิ่ม AI panel บน Eden home page (เฉพาะผู้ที่ login แล้ว) "
    "เป็น banner gradient สีม่วง (#667eea → #764ba2) พร้อมไอคอน 🤖 "
    "และปุ่ม 'เปิดใช้งาน AI →' ที่ลิงก์ไป /eden/ai/ "
    "รวมถึง quick links ไปยัง AI features ทั้ง 9 ตัว",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================================
# บทที่ 10: Endpoints Reference
# ============================================================
add_heading_h1(doc, "Endpoints Reference — URL ทั้งหมดของ AI Modules", num=10)

add_heading_h2(doc, "10.1  AI Landing (Use Case 11)")

add_table(doc,
    headers=["URL", "Method", "Auth", "คำอธิบาย"],
    rows=[
        ["/eden/ai/", "GET", "login", "AI features dashboard (9 cards)"],
    ],
    col_widths_cm=[3.5, 1.8, 1.8, 9.0],
    first_col_bold=True, font_size=9)

add_heading_h2(doc, "10.2  AI Chatbot (Use Case 1)")

add_table(doc,
    headers=["URL", "Method", "Auth", "คำอธิบาย"],
    rows=[
        ["/eden/ai_chat/",             "GET",  "anonymous + consent", "Chat UI"],
        ["/eden/ai_chat/ask",          "POST", "anonymous + consent", "ส่งคำถาม → รับคำตอบ (JSON)"],
        ["/eden/ai_chat/history",      "GET",  "login",               "ประวัติการสนทนา"],
        ["/eden/ai_chat/rate/<id>",    "POST", "login",               "Rate answer (+1/-1)"],
        ["/eden/ai_chat/clear",        "POST", "anonymous",            "เคลียร์ session"],
        ["/eden/ai_chat/widget",       "GET",  "anonymous",            "Embeddable widget (iframe)"],
        ["/eden/ai_chat/consent",      "GET",  "anonymous",            "Privacy notice (PDPA)"],
        ["/eden/ai_chat/consent_give", "POST", "anonymous",            "Record consent"],
        ["/eden/ai_chat/consent_withdraw","POST","login",              "Withdraw consent"],
    ],
    col_widths_cm=[4.5, 1.5, 3.0, 7.0],
    first_col_bold=True, font_size=9)

add_heading_h2(doc, "10.3  AI OCR (Use Case 2)")

add_table(doc,
    headers=["URL", "Method", "Auth", "คำอธิบาย"],
    rows=[
        ["/eden/ai_ocr/",         "GET",  "login + consent", "Upload UI"],
        ["/eden/ai_ocr/engines",  "GET",  "login",           "List available OCR engines"],
        ["/eden/ai_ocr/extract",  "POST", "login + consent", "Upload + extract (multipart/form-data)"],
        ["/eden/ai_ocr/jobs",     "GET",  "login",           "List user's OCR jobs"],
        ["/eden/ai_ocr/job/<id>", "GET",  "login",           "Get job detail"],
    ],
    col_widths_cm=[4.0, 1.5, 3.0, 7.5],
    first_col_bold=True, font_size=9)

add_heading_h2(doc, "10.4  AI Analytics, Imagery, KB, Moderation, Translate, Query, Voice (Use Cases 3-10)")

add_table(doc,
    headers=["Controller", "Endpoints", "Use Case"],
    rows=[
        ["ai_analytics.py",  "/eden/ai_analytics/  /eden/ai_analytics/forecast/<metric>  /eden/ai_analytics/data/<metric>.json  /eden/ai_analytics/run/<metric>",  "3"],
        ["ai_imagery.py",    "/eden/ai_imagery/  /eden/ai_imagery/analyze  /eden/ai_imagery/jobs  /eden/ai_imagery/job/<id>",                                       "4"],
        ["ai_translate.py",  "/eden/ai_translate/  /eden/ai_translate/missing/<lang>  /eden/ai_translate/translate_batch  /eden/ai_translate/review  /eden/ai_translate/approve  /eden/ai_translate/reject  /eden/ai_translate/apply_approved",  "5"],
        ["ai_query.py",      "/eden/ai_query/  /eden/ai_query/ask  /eden/ai_query/history",                                                                              "6"],
        ["ai_kb.py",         "/eden/ai_kb/  /eden/ai_kb/ingest_cms_posts  /eden/ai_kb/ingest_url  /eden/ai_kb/list_chunks  /eden/ai_kb/chunk/<id>  /eden/ai_kb/delete/<id>  /eden/ai_kb/rebuild_index",  "8"],
        ["ai_voice.py",      "/eden/ai_voice/  /eden/ai_voice/transcribe  /eden/ai_voice/speak  /eden/ai_voice/download  /eden/ai_voice/command  /eden/ai_voice/history",  "9"],
        ["ai_moderation.py", "/eden/ai_moderation/  /eden/ai_moderation/queue  /eden/ai_moderation/check  /eden/ai_moderation/approve  /eden/ai_moderation/reject  /eden/ai_moderation/history",  "10"],
    ],
    col_widths_cm=[3.5, 9.0, 1.0],
    first_col_bold=True, font_size=8.5, header_size=9)

# ============================================================
# บทที่ 11: PDPA + Security
# ============================================================
add_heading_h1(doc, "PDPA Compliance & Security", num=11)

add_heading_h2(doc, "11.1  ตาราง PDPA Compliance Matrix")

add_table(doc,
    headers=["Use Case", "เก็บ PII?", "Consent Required?", "Retention", "หมายเหตุ"],
    rows=[
        ["1 Chatbot",        "ใช่ (คำถาม-คำตอบ)",     "บังคับ",         "90 วัน",    "anonymous ใช้ได้ ต้อง consent ก่อน"],
        ["2 OCR",            "ใช่ (รูปเอกสาร + data)","บังคับ",         "1 ปี (audit)","รูปถูกลบหลัง extraction เก็บ JSON เท่านั้น"],
        ["3 Analytics",      "ไม่",                    "ไม่",           "ทับด้วย version","admin operation"],
        ["4 Imagery",        "ไม่",                    "ไม่",           "ไม่กำหนด",    "ภาพถ่ายดาวเทียม/drone"],
        ["5 Translation",    "ไม่",                    "ไม่",           "ไม่กำหนด",    "UI text strings เท่านั้น"],
        ["6 NLQ",            "อาจ (query result)",     "ไม่ (admin-only)","1 ปี",     "whitelist ห้าม query sensitive tables"],
        ["8 KB",             "ไม่ (public เท่านั้น)",   "ไม่",           "—",         "ห้าม index ข้อมูล sensitive"],
        ["9 Voice",          "อาจ (transcript)",       "ไม่ (admin ops)","90 วัน",   "transcript อาจมี PII"],
        ["10 Moderation",    "อาจ (content ที่ moderate)","ไม่ (auto)",  "3 ปี",      "PDPA violation detection รวม"],
    ],
    col_widths_cm=[2.5, 3.5, 2.8, 2.5, 4.7],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "11.2  Security Measures")

add_bullet(doc, "Rate limit: Chatbot 60 requests/minute/IP (others: implicit through admin-only)")
add_bullet(doc, "Authentication: Use Cases 2-10 บาง action ต้อง ADMIN membership; Use Case 1 anonymous + consent")
add_bullet(doc, "NLQ whitelist: อนุญาตเฉพาะ controllers org, cr, inv, project, cms, gis, hrm, req, stats — ห้าม auth_user, dvr_case, med_patient, msg_log")
add_bullet(doc, "OpenThai API key: โหลดจาก /etc/eden/env (root-only readable, never committed to repo)")
add_bullet(doc, "HTTPS session cookie: บังคับใช้ secure flag เมื่อ request.is_https (ใน models/00_settings.py)")
add_bullet(doc, "Audit log: ทุก AI call log model, latency, IP, consent flag")
add_bullet(doc, "File deletion: รูป OCR ถูกลบหลัง extraction (เก็บแค่ JSON output)")
add_bullet(doc, "PDPA violation detection: Moderation module ตรวจหาเลขบัตรประชาชน 13 หลัก + เบอร์มือถือไทยใน content")

# ============================================================
# บทที่ 12: Installation
# ============================================================
add_heading_h1(doc, "การติดตั้งและรันระบบ", num=12)

add_heading_h2(doc, "12.1  การติดตั้ง Base Eden (จาก docs/ai/INSTALL_SUMMARY.md)")

add_code_block(doc, '''# 1. Clone web2py
git clone https://github.com/web2py/web2py.git ~/web2py
cd ~/web2py
git reset --hard 9005306          # v3.3.3 stable
git submodule update --init --recursive

# 2. Clone Eden fork
git clone https://github.com/skpuan07-cmd/eden.git ~/eden

# 3. Symlink Eden เข้า web2py
cd ~/web2py/applications
ln -s ~/eden eden

# 4. คัดลอก config template
cd ~/eden
cp modules/templates/000_config.py models/000_config.py
# แก้ FINISHED_EDITING_CONFIG_FILE = True

# 5. First-time setup (สร้าง database + prepopulate)
cd ~/web2py
python web2py.py -S eden -M -R applications/eden/static/scripts/tools/noop.py

# 6. เริ่ม server
python web2py.py --no_gui -a <password>''', language="bash")

add_heading_h2(doc, "12.2  การติดตั้ง AI Dependencies")

add_code_block(doc, '''# 1. z-ai-web-dev-sdk (AI engine หลัก)
npm install -g z-ai-web-dev-sdk
# ตรวจสอบ:
which z-ai

# 2. Prophet (สำหรับ Use Case 3: Predictive Analytics)
pip install --break-system-packages prophet pandas

# 3. (Optional) PostgreSQL + pgvector (สำหรับ vector search จริง)
#    + pip install openai สำหรับ embeddings

# 4. OpenThai API key (สำหรับ Thai OCR)
#    สมัครที่ https://iapp.co.th/
#    เก็บ key ที่ /etc/eden/env (root-only readable):
echo "IAPP_API_KEY=your_key_here" | sudo tee /etc/eden/env
sudo chmod 600 /etc/eden/env''', language="bash")

add_heading_h2(doc, "12.3  การ Restart และตรวจสอบ")

add_code_block(doc, '''# Restart web2py
kill -TERM $(cat ~/eden-logs/web2py.pid)
python3 ~/scripts/launch_eden_daemon.py

# ตรวจสอบตาราง AI ถูกสร้าง
# เปิด: http://127.0.0.1:3000/eden/default/index
# login เป็น admin → ไปที่ appadmin → ดู ai_* tables

# ตรวจสอบ error tickets
ls -lt ~/eden/errors/ | head -5''', language="bash")

add_heading_h2(doc, "12.4  Demo Accounts (หลัง first-run)")

add_table(doc,
    headers=["Email", "Password", "Role"],
    rows=[
        ["admin@example.com",       "testing", "System Administrator"],
        ["normaluser@example.com",  "testing", "Authenticated (regular user)"],
    ],
    col_widths_cm=[5.5, 3.0, 7.5],
    first_col_bold=True)

# ============================================================
# บทที่ 13: สรุป
# ============================================================
add_heading_h1(doc, "สรุปและขั้นตอนการทำงานต่อ", num=13)

add_heading_h2(doc, "13.1  สรุปโครงสร้างระบบ")

add_para(doc,
    "Repository นี้เป็น fork ของ Sahana Eden v6.2 ที่เพิ่ม AI modules ทั้ง 10 ตัว "
    "โดยใช้ pattern ที่เรียบง่าย — เพิ่ม controllers/models/views/modules ใหม่ทั้งหมด "
    "และแก้ไขไฟล์เดิมเพียง 4 ไฟล์ (minimal patch) โครงสร้างฐานข้อมูลมาตรฐาน "
    "388 ตาราง + AI tables 9 ตารางใหม่ = 397 ตาราง ทั้งหมดอยู่บน SQLite (default) "
    "หรือ PostgreSQL/MySQL ได้",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "13.2  ขั้นตอนการทำงานต่อ (Recommendations)")

add_heading_h3(doc, "① ทำความเข้าใจ Base Eden ก่อน")
add_para(doc,
    "ศึกษา S3 framework architecture (DataModel, CRUDResource, CRUDRequest, BasicCRUD) "
    "จากเอกสารที่ผมเคยสร้างให้ (Sahana_Eden_S3_Framework_Architecture.docx) "
    "เพื่อเข้าใจว่า Eden modules มาตรฐานทำงานอย่างไร ก่อนเข้าใจ AI extensions",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "② ติดตั้งและรันระบบจริง")
add_para(doc,
    "ปฏิบัติตามบทที่ 12 ของเอกสารนี้ — ติดตั้ง web2py + Eden + AI dependencies "
    "รัน first-time setup script และเริ่ม server ทดสอบว่าทุก module ทำงานได้ "
    "ก่อนเริ่ม extend",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "③ ศึกษา AI Controllers เรียบตามลำดับ Use Cases")
add_para(doc,
    "แนะนำให้ไล่จาก Use Case 1 (Chatbot) → 2 (OCR) → 8 (KB) → 3 (Analytics) "
    "→ 6 (NLQ) → 4 (Imagery) → 5 (Translate) → 9 (Voice) → 10 (Moderation) "
    "→ 11 (Landing) เพราะ Use Cases แรก ๆ ค่อนข้าง self-contained "
    "Use Case 6 (NLQ) ต้องการความเข้าใจ Eden REST API ก่อน",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "④ อ่าน Documentation ใน docs/ai/")
add_table(doc,
    headers=["ไฟล์", "จำนวนบรรทัด", "เนื้อหา"],
    rows=[
        ["README.md",              "414",  "คู่มือ 10 use cases พร้อม endpoint table"],
        ["INSTALL_SUMMARY.md",     "138",  "รายงานการติดตั้ง + สถาปัตยกรรม"],
        ["SYSTEM_DOCUMENTATION.md","873",  "เอกสารระบบฉบับเต็ม (deploy, schema, patches)"],
        ["RECOMMENDATIONS.md",     "977",  "คำแนะนำ AI/Cyber/PDPA + Roadmap"],
        ["index.md",               "81",   "index สำหรับ doc navigation"],
    ],
    col_widths_cm=[5.0, 2.5, 8.5],
    first_col_bold=True)

add_heading_h3(doc, "⑤ ส่วนที่ควรพัฒนาต่อ")

add_bullet(doc, "Vector search จริง — ปัจจุบัน RAG ใช้ keyword + TF-IDF matching (placeholder) ควร upgrade เป็น pgvector")
add_bullet(doc, "Scheduler integration — เพิ่ม forecast_all() ใน models/tasks.py รันทุกคืน 02:00")
add_bullet(doc, "Moderation hooks — เชื่อม ai_moderation เข้ากับ controllers/cms.py หลัง post creation")
add_bullet(doc, "Line OA / Telegram integration — เชื่อม ai_chat เข้ากับ webhook ของ msg module")
add_bullet(doc, "Tests — เพิ่ม unit tests ใน tests/ สำหรับ AI controllers ทั้ง 10")
add_bullet(doc, "OpenThai endpoint อื่น ๆ — ปัจจุบันรองรับเฉพาะบัตรประชาชน ควรเพิ่ม invoice/receipt OCR")
add_bullet(doc, "Performance monitoring — เพิ่ม dashboard สำหรับดู latency/token usage/cost")
add_bullet(doc, "Multi-tenancy — ปัจจุบัน AI modules ใช้ shared knowledge base ควรแยก per-organisation")

add_heading_h2(doc, "13.3  ไฟล์เอกสารที่เกี่ยวข้อง")

add_table(doc,
    headers=["ไฟล์", "หน้าที่"],
    rows=[
        ["Sahana_Eden_skpuan07_Repo_Analysis.docx",      "เอกสารฉบับนี้ (analysis)"],
        ["Sahana_Eden_S3_Framework_Architecture.docx",   "S3 framework deep dive (สร้างก่อนหน้า)"],
        ["Sahana_Eden_Installation_README.md",            "คู่มือติดตั้ง base Eden"],
        ["docs/ai/SYSTEM_DOCUMENTATION.md",              "เอกสารระบบของ AI modules (ใน repo)"],
        ["docs/ai/RECOMMENDATIONS.md",                   "คำแนะนำ AI/Cyber/PDPA (ใน repo)"],
        ["docs/ai/README.md",                            "คู่มือ 10 use cases (ใน repo)"],
        ["docs/ai/INSTALL_SUMMARY.md",                   "รายงานการติดตั้ง (ใน repo)"],
    ],
    col_widths_cm=[7.0, 9.0],
    first_col_bold=True)

# Footer
section = doc.sections[0]
footer = section.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fp.text = ""
add_run(fp, "Sahana Eden (skpuan07-cmd fork) — Repository Analysis  ·  Page ",
        font_name=FONT_BODY, size_pt=9, color=C_SECOND)
fldChar1 = OxmlElement('w:fldChar')
fldChar1.set(qn('w:fldCharType'), 'begin')
instrText = OxmlElement('w:instrText')
instrText.text = 'PAGE'
fldChar2 = OxmlElement('w:fldChar')
fldChar2.set(qn('w:fldCharType'), 'end')
run = fp.add_run()
run.font.name = FONT_BODY
run.font.size = Pt(9)
run.font.color.rgb = C_SECOND
run._r.append(fldChar1)
run._r.append(instrText)
run._r.append(fldChar2)

os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
doc.save(OUT_PATH)
print(f"✓ Saved: {OUT_PATH}")
print(f"  Size: {os.path.getsize(OUT_PATH):,} bytes")
