"""
Generate the S3 Framework Architecture Analysis .docx document.
Style: Technical clean. Language: Thai (descriptions) + English (technical).
Target audience: Developers using Sahana Eden in production.
Length: ~20-30 pages.
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from docx.enum.section import WD_SECTION

# ============================================================
# Configuration
# ============================================================
OUT_PATH = "/home/z/my-project/download/Sahana_Eden_S3_Framework_Architecture.docx"
ASSETS   = "/home/z/my-project/scripts/assets"

# Fonts: Loma for Thai, Liberation Mono for code
FONT_THAI      = "Loma"
FONT_HEADING   = "Loma"
FONT_BODY      = "Loma"
FONT_CODE      = "Liberation Mono"

# Colors
C_PRIMARY     = RGBColor(0x1E, 0x3A, 0x8A)   # deep blue
C_ACCENT      = RGBColor(0x0E, 0x74, 0x90)   # teal
C_DARK        = RGBColor(0x0F, 0x17, 0x2A)   # near-black slate
C_SECOND      = RGBColor(0x47, 0x55, 0x69)   # slate
C_CODE_BG     = "F1F5F9"                       # light gray-blue (hex string for shading)
C_CODE_BORDER = "94A3B8"
C_TABLE_HEAD  = "1E3A8A"
C_TABLE_ALT   = "EFF6FF"
C_NOTE_BG     = "FEF3C7"                       # amber tint
C_NOTE_BORDER = "D97706"
C_RED         = RGBColor(0xDC, 0x26, 0x26)
C_GREEN       = RGBColor(0x16, 0xA3, 0x4A)

# ============================================================
# Helpers
# ============================================================
def set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    """kwargs: top, bottom, left, right — each is dict {sz, color, val}"""
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
            bold=False, italic=False, color=None, font_thai=True):
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    # Set East Asian font (Thai) via rPr
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    if font_thai:
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
             color=None, align=None, before=0, after=6, line_pt=None,
             font=FONT_BODY, indent_first=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    set_paragraph_spacing(p, before_pt=before, after_pt=after, line_pt=line_pt)
    if indent_first is not None:
        p.paragraph_format.first_line_indent = Cm(indent_first)
    if text:
        add_run(p, text, font_name=font, size_pt=size,
                bold=bold, italic=italic, color=color)
    return p

def add_heading_h1(doc, text, num=None):
    """Chapter heading: blue bar + number + title."""
    full = f"บทที่ {num}   {text}" if num else text
    p = doc.add_paragraph(style="Heading 1")
    p.style = doc.styles["Heading 1"]
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.keep_with_next = True
    # clear default runs
    for r in list(p.runs):
        r.text = ""
    add_run(p, full, font_name=FONT_HEADING, size_pt=18,
            bold=True, color=C_PRIMARY)
    # Bottom border
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
    """Code block with gray background and left border. Single paragraph."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.left_indent = Cm(0.4)
    p.paragraph_format.right_indent = Cm(0.4)
    set_paragraph_shading(p, C_CODE_BG)
    set_paragraph_border(p,
        left={"val": "single", "sz": 24, "color": "1E3A8A", "space": 6},
        top={"val": "single", "sz": 4,  "color": C_CODE_BORDER, "space": 2},
        bottom={"val": "single", "sz": 4,  "color": C_CODE_BORDER, "space": 2},
        right={"val": "single", "sz": 4,  "color": C_CODE_BORDER, "space": 2},
    )
    # Replace \n with line break runs
    lines = code.rstrip("\n").split("\n")
    for i, line in enumerate(lines):
        r = p.add_run(line)
        r.font.name = FONT_CODE
        r.font.size = Pt(9)
        # set east-asia too in case there's CJK
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
    # Tiny language tag
    tag = doc.add_paragraph()
    tag.paragraph_format.space_before = Pt(0)
    tag.paragraph_format.space_after = Pt(8)
    set_paragraph_spacing(tag, before_pt=0, after_pt=8)
    add_run(tag, f"[ {language} ]", font_name=FONT_BODY, size_pt=8,
            italic=True, color=C_SECOND)
    tag.paragraph_format.left_indent = Cm(0.4)

def add_inline_code(paragraph, code):
    """Inline monospace code in a normal paragraph."""
    r = paragraph.add_run(code)
    r.font.name = FONT_CODE
    r.font.size = Pt(10)
    rPr = r._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), FONT_CODE)
    rFonts.set(qn("w:hAnsi"), FONT_CODE)
    rFonts.set(qn("w:cs"), FONT_CODE)
    rFonts.set(qn("w:eastAsia"), FONT_CODE)
    # light gray highlight
    rpr = r._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F1F5F9")
    rpr.append(shd)
    return r

def add_note_box(doc, title, body, color=C_NOTE_BORDER, bg=C_NOTE_BG):
    """Admonition-style callout box."""
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
    """Build a styled table."""
    n_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)

    # Header row
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
        # clear default paragraph and set custom
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, h, font_name=FONT_HEADING, size_pt=header_size,
                bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    # Body rows
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
            # split on \n for multi-line cells
            lines = cell_text.split("\n") if isinstance(cell_text, str) else [str(cell_text)]
            for li, line in enumerate(lines):
                if li > 0:
                    p.add_run().add_break()
                is_first_col = c_idx == 0
                add_run(p, line, font_name=FONT_BODY, size_pt=font_size,
                        bold=is_first_col and first_col_bold,
                        color=C_DARK if not (is_first_col and first_col_bold) else C_PRIMARY)
    # Mark header row to repeat on each page
    trPr = hdr._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader)
    # cantSplit on all rows
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        cantSplit = OxmlElement("w:cantSplit")
        trPr.append(cantSplit)

    # spacer paragraph after table
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(4)
    return table

def add_image(doc, path, width_cm=16, caption=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True if caption else False
    r = p.add_run()
    r.add_picture(path, width=Cm(width_cm))
    if caption:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_after = Pt(10)
        add_run(cp, caption, font_name=FONT_BODY, size_pt=9,
                italic=True, color=C_SECOND)

# ============================================================
# Document construction
# ============================================================
doc = Document()

# ----- Page setup: A4, 2.2cm margins -----
section = doc.sections[0]
section.page_height = Cm(29.7)
section.page_width  = Cm(21.0)
section.top_margin    = Cm(2.2)
section.bottom_margin = Cm(2.2)
section.left_margin   = Cm(2.3)
section.right_margin  = Cm(2.3)
section.header_distance = Cm(1.0)
section.footer_distance = Cm(1.0)

# ----- Default style: Loma, 10.5pt, 1.3 spacing -----
normal = doc.styles["Normal"]
normal.font.name = FONT_BODY
normal.font.size = Pt(10.5)
normal.paragraph_format.line_spacing = 1.3
normal.paragraph_format.space_after = Pt(6)
# set east-asia font for default style
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
# COVER PAGE
# ============================================================
# Top spacer
for _ in range(3):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, after_pt=0)

# Top label
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
add_run(p, "DEVELOPER REFERENCE HANDBOOK", font_name=FONT_HEADING,
        size_pt=11, bold=True, color=C_ACCENT)

# horizontal divider
pdiv = doc.add_paragraph()
pdiv.alignment = WD_ALIGN_PARAGRAPH.CENTER
pdiv.paragraph_format.space_after = Pt(30)
set_paragraph_border(pdiv, bottom={"val": "single", "sz": 18, "color": C_TABLE_HEAD, "space": 1})

# Big title (TH)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
add_run(p, "Sahana Eden", font_name=FONT_HEADING, size_pt=38,
        bold=True, color=C_PRIMARY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(24)
add_run(p, "วิเคราะห์สถาปัตยกรรม S3 Framework เชิงลึก",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

# Subtitle
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(40)
add_run(p, "S3Resource, S3CRUD, REST Pattern & Refactor Notes (Eden ASP)",
        font_name=FONT_BODY, size_pt=12, italic=True, color=C_SECOND)

# Info table
info = doc.add_table(rows=4, cols=2)
info.alignment = WD_TABLE_ALIGNMENT.CENTER
info.autofit = False
info_data = [
    ("Repository",   "github.com/sahana/eden  (master branch)"),
    ("Branch",       "Eden ASP  (since May 2025)"),
    ("Version ref.", "6.2  (Sep 2026)"),
    ("Document",     "Technical Architecture Reference  ·  Nov 2026"),
]
for i, (k, v) in enumerate(info_data):
    row = info.rows[i]
    row.cells[0].width = Cm(4.5)
    row.cells[1].width = Cm(10.5)
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

# Bottom note
for _ in range(2):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(40)
add_run(p, "⚠ หมายเหตุ: หลังการ refactor Eden ASP (พ.ค. 2025) ชื่อคลาสสาธารณะเปลี่ยนไป\n"
           "S3Model → DataModel  ·  S3Resource → CRUDResource  ·  S3CRUD → BasicCRUD  ·  S3RESTController → CRUDRequest",
        font_name=FONT_BODY, size_pt=9.5, italic=True, color=C_RED)

# Page break to start TOC
doc.add_page_break()

# ============================================================
# TABLE OF CONTENTS (manual TOC since python-docx lacks field TOC)
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_after = Pt(16)
set_paragraph_border(p, bottom={"val": "single", "sz": 12, "color": C_TABLE_HEAD, "space": 4})
add_run(p, "สารบัญ", font_name=FONT_HEADING, size_pt=20, bold=True, color=C_PRIMARY)

toc_items = [
    ("บทที่ 1", "ภาพรวมและความสำคัญ"),
    ("บทที่ 2", "สถาปัตยกรรม S3 Framework"),
    ("บทที่ 3", "DataModel — การนิยาม Data Model"),
    ("บทที่ 4", "CRUDRequest — REST Dispatcher"),
    ("บทที่ 5", "CRUDResource — Resource Abstraction Layer"),
    ("บทที่ 6", "BasicCRUD — Standard CRUD Methods"),
    ("บทที่ 7", "REST Request Lifecycle"),
    ("บทที่ 8", "Content Negotiation"),
    ("บทที่ 9", "Default Methods และ Custom REST Methods"),
    ("บทที่ 10", "Hooks & Extensibility"),
    ("บทที่ 11", "File Map โครงสร้าง Repository"),
    ("บทที่ 12", "การเปลี่ยนแปลง Eden ASP Refactor"),
    ("บทที่ 13", "สรุป Design Philosophy"),
]
for num, title in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(15.5), WD_ALIGN_PARAGRAPH.RIGHT, 2)  # leader=dots
    add_run(p, num + "  ", font_name=FONT_HEADING, size_pt=11, bold=True, color=C_PRIMARY)
    add_run(p, title, font_name=FONT_BODY, size_pt=11, color=C_DARK)

# ============================================================
# บทที่ 1: ภาพรวมและความสำคัญ
# ============================================================
add_heading_h1(doc, "ภาพรวมและความสำคัญ", num=1)

add_para(doc,
    "Sahana Eden เป็น Rapid Application Development (RAD) Kit สำหรับสร้างเว็บแอปพลิเคชัน "
    "ด้านการจัดการภัยพิบัติและมนุษยธรรม พัฒนาโดย Sahana Software Foundation (SSF) "
    "องค์กรไม่แสวงหากำไรในรัฐแคลิฟอร์เนีย เริ่มต้นจากการตอบสนองสึนามิในศรีลังกาเมื่อปี 2004 "
    "และได้รับการนำไปใช้ในกว่า 30 ประเทศทั่วโลก ตั้งแต่เหตุการณ์สึนามิในมหาสมุทรอินเดีย, "
    "แผ่นดินไหวในปากีสถาน (2005), ประเทศฟิลิปปินส์ (2006), อินโดนีเซีย (2006) "
    "จนถึงแผ่นดินไหวในเฮติ (2010) ทำให้ Eden กลายเป็นหนึ่งในแพลตฟอร์มโอเพนซอร์ส "
    "ที่ได้รับความนิยมมากที่สุดในแวดวงมนุษยธรรม",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "ทำไมต้องเข้าใจ S3 Framework")

add_para(doc,
    "แกนกลางของ Eden คือ S3 framework ซึ่งเป็นชั้น abstraction ที่สร้างอยู่เหนือ web2py/PyDAL "
    "การจะ contribute โค้ด, สร้าง module ใหม่, หรือแม้แต่ติดตั้งและปรับแต่ง Eden ให้เข้ากับงาน "
    "ล้วนต้องเข้าใจ S3 framework เป็นพื้นฐาน เพราะทุก URL endpoint, ทุก data table, ทุก hook "
    "และทุก format output ที่ Eden รองรับ ล้วนผ่านกลไกของ S3 framework ทั้งสิ้น",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "สิ่งที่ผู้อ่านจะได้จากเอกสารนี้")

add_para(doc,
    "เอกสารนี้อธิบาย S3 framework ในระดับซอร์สโค้ด โดยอ้างอิงโค้ดจริงจาก master branch "
    "ของ repository github.com/sahana/eden เนื้อหาประกอบด้วยโครงสร้างคลาสหลัก, "
    "วงจรอายุของ request หนึ่งรายการตั้งแต่เข้าระบบจนส่ง response ตาราง content negotiation "
    "ที่ครอบคลุมทุก extension, ระบบ hooks ที่ขยายได้, และแผนผังไฟล์ใน repository "
    "เพื่อให้ developer สามารถนำไปประยุกต์ใช้ในการเขียน controller ใหม่, นิยาม DataModel "
    "หรือเพิ่ม custom REST method ได้ทันที",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_note_box(doc, "⚠ ข้อควรทราบสำคัญ",
    "หลังการ refactor ครั้งใหญ่ในสาขา Eden ASP (พฤษภาคม 2025) คลาส public หลักของ S3 framework "
    "ถูกเปลี่ยนชื่อ คำว่า S3Model, S3Resource, S3CRUD, S3RESTController ในเอกสารเก่าและบล็อก "
    "เป็นเพียงชื่อเดิมที่ยังถูกใช้ในเอกสารประกอบ ส่วนชื่อจริงในโค้ดปัจจุบันคือ "
    "DataModel, CRUDResource, BasicCRUD และ CRUDRequest ตามลำดับ การค้นหา "
    "S3RESTController ด้วย grep ใน repo ทั้งหมดจะได้ผลเป็นศูนย์ เพราะคลาสนี้ไม่มีอยู่จริงในปัจจุบัน")

# ============================================================
# บทที่ 2: สถาปัตยกรรม S3 Framework
# ============================================================
add_heading_h1(doc, "สถาปัตยกรรม S3 Framework", num=2)

add_para(doc,
    "S3 framework ของ Sahana Eden เป็นชั้น abstraction ที่สร้างอยู่บน web2py และ PyDAL "
    "ไม่ใช่การแทนที่ web2py แต่เป็นการเพิ่มความสามารถในระดับสูงขึ้นไป ได้แก่ "
    "lazy model loading, resource/component abstraction, REST dispatch, "
    "hooks, authorization, audit และ multi-format serialization "
    "การออกแบบนี้ทำให้ Eden สามารถรองรับ use case ที่หลากหลายในงานมนุษยธรรม "
    "โดยใช้ codebase เดียวกัน ตั้งแต่ case management, inventory tracking, "
    "volunteer management ไปจนถึง disaster victim identification",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_image(doc, os.path.join(ASSETS, "diagram_architecture.png"),
          width_cm=16.5,
          caption="รูปที่ 2.1  ผังความสัมพันธ์ระหว่างคลาสหลักของ S3 framework และตำแหน่งใน repository")

add_heading_h2(doc, "2.1  Foundation: web2py + PyDAL")

add_para(doc,
    "current.db ที่ใช้ทั่วทั้ง Eden เป็น PyDAL DAL ปกติ สร้างขึ้นในไฟล์ models/00_db.py "
    "เป็น database abstraction layer ของ web2py ที่รองรับหลาย database backend "
    "(MySQL, PostgreSQL, SQLite, MSSQL ฯลฯ) CRUDResource.table ที่ใช้ใน S3 framework "
    "ก็คือ PyDAL Table จริง ๆ ส่วน DataModel.define_table() เป็น wrapper ที่เรียก "
    "db.define_table() ของ PyDAL โดยตรง แต่เพิ่ม meta-fields (uuid, created_on, "
    "modified_on, owned_by_user ฯลฯ) เข้าไปอัตโนมัติ",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "2.2  S3 Framework Layer")

add_para(doc,
    "ชั้น S3 framework อยู่ใน package modules/core/ และถูก import เป็น namespace "
    "s3base ในไฟล์ models/000_1st_run.py ทำให้สามารถเข้าถึงคลาสทั้งหมดได้ผ่าน "
    "s3base.CRUDResource(), s3base.DataModel ฯลฯ โครงสร้างภายใน modules/core/ "
    "แบ่งเป็น subpackage ตามหน้าที่ ดังที่แสดงในรูปที่ 2.1 ด้านบน "
    "โดยแต่ละส่วนมีความรับผิดชอบแยกจากกันชัดเจน และสามารถพัฒนาต่อยอดได้อย่างอิสระ",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "คลาสหลัก 5 ตัวที่ developer ต้องรู้")

add_table(doc,
    headers=["คลาส (ปัจจุบัน)", "อดีตชื่อ", "ไฟล์", "หน้าที่หลัก"],
    rows=[
        ["DataModel",        "S3Model",           "core/model/datamodel.py",  "Base class สำหรับนิยามตาราง + lazy loader + hooks"],
        ["CRUDResource",     "S3Resource",        "core/resource/resource.py", "ห่อ PyDAL Table + filter + components พร้อม CRUD methods"],
        ["CRUDRequest",      "S3RESTController",  "core/controller.py",       "Parse URL, dispatch ไปยัง handler ที่เหมาะสม"],
        ["BasicCRUD",        "S3CRUD",            "core/methods/crud.py",     "create/read/update/delete/select handler"],
        ["RESTful",          "—",                 "core/methods/rest.py",     "XML/S3JSON/GeoJSON tree API (REST verb-style)"],
    ],
    col_widths_cm=[3.2, 3.0, 4.5, 5.5],
    first_col_bold=True)

add_heading_h3(doc, "คลาสสนับสนุนที่สำคัญ")

add_table(doc,
    headers=["คลาส", "ไฟล์", "หน้าที่"],
    rows=[
        ["CRUDMethod",         "core/methods/base.py",        "Base class ของทุก method handler"],
        ["S3ResourceFilter",   "core/resource/rfilter.py",    "สร้าง DAL Query จาก URL vars"],
        ["S3ResourceField",    "core/resource/query.py",      "เลือก field + สร้าง JOIN"],
        ["S3ResourceTree",     "core/resource/rtb.py",        "สร้าง XML tree สำหรับ export"],
        ["S3Components",       "core/resource/components.py", "จัดการ component graph (master + sub-records)"],
        ["ResourceData",       "core/resource/data.py",      "ดึง row จริงจาก database"],
        ["DataExporter",      "core/resource/exporter.py",   "Facade ส่งต่อไปยัง formats/* writers"],
        ["S3Config",           "modules/s3cfg.py",            "current.deployment_settings"],
        ["S3DAL",              "modules/s3dal.py",            "compat shim 124 บรรทัด (ไม่ใช่ DAL จริง!)"],
    ],
    col_widths_cm=[3.8, 4.7, 7.5],
    first_col_bold=True)

add_note_box(doc, "คำเตือนเรื่อง S3DAL",
    "ชื่อ S3DAL ฟังดูเหมือนเป็น Data Access Layer ที่ extend PyDAL แต่จริง ๆ แล้วเป็น "
    "compatibility shim ขนาด 124 บรรทัดเท่านั้น ทำหน้าที่จัดการ breaking changes ของ PyDAL API "
    "ระหว่างเวอร์ชัน (เช่น adapter.INVERT vs dialect.invert, table._dalname vs table._ot) "
    "ส่วน DAL จริงที่ใช้ใน Eden คือ PyDAL DAL ธรรมดาใน models/00_db.py "
    "งาน pagination, filtering, joins ที่ซับซ้อนอยู่ใน ResourceData และ S3ResourceFilter ไม่ใช่ใน S3DAL",
    color="DC2626", bg="FEE2E2")

# ============================================================
# บทที่ 3: DataModel
# ============================================================
add_heading_h1(doc, "DataModel — การนิยาม Data Model", num=3)

add_para(doc,
    "DataModel (อดีตชื่อ S3Model) เป็น base class สำหรับนิยาม data model ใน Eden "
    "แต่ละ domain ของ Eden (org, pr, inv, hrm, project, ...) มี DataModel subclass "
    "ของตัวเองอยู่ใน modules/s3db/ เช่น modules/s3db/org.py บรรจุ OrgOrganisationModel, "
    "OrgSiteModel, OrgFacilityModel ฯลฯ การเข้าถึงตารางใน Eden ทำผ่าน "
    "current.s3db.<tablename> ซึ่ง trigger การ lazy-load ตารางนั้น ๆ ขึ้นมาเมื่อต้องการใช้ "
    "ทำให้ Eden ไม่ต้องโหลดตารางทั้งหมด 41 ไฟล์ขึ้นมาตั้งแต่ต้น และลด memory footprint ได้มาก",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "3.1  โครงสร้าง DataModel subclass")

add_para(doc,
    "ตัวอย่างต่อไปนี้คือส่วนต้นของ OrgOrganisationModel ใน modules/s3db/org.py "
    "ซึ่งเป็นตัวอย่างที่เข้าใจง่ายที่สุด เพราะเป็น model ขนาดเล็กและมีโครงสร้างชัดเจน "
    "จุดที่สังเกตคือการประกาศ names tuple ซึ่งเป็นสัญญาว่า class นี้จะ export อะไรบ้าง "
    "ทั้งตาราง (org_organisation_type, org_organisation) และ symbol ที่ใช้ซ้ำ "
    "เช่น FieldTemplate (org_organisation_type_id) และ Represent class",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''class OrgOrganisationModel(DataModel):
    """ Organisations """

    names = ("org_organisation_type",
             "org_organisation_type_id",
             "org_region", "org_region_country", "org_region_id",
             "org_organisation",
             "org_organisation_crud_fields",
             "org_organisation_id",
             "org_organisation_represent",
             )   # ← ประกาศทุก symbol ที่จะ export

    def model(self):
        T = current.T
        db = current.db
        add_components = self.add_components
        configure = self.configure
        define_table = self.define_table

        tablename = "org_organisation_type"
        define_table(tablename,
                     Field("name", length=128, notnull=True, unique=True,
                           label=T("Name"),
                           requires=[IS_NOT_EMPTY(), IS_LENGTH(128),
                                     IS_NOT_IN_DB(db, "org_organisation_type.name")]),
                     Field("parent", "reference org_organisation_type",
                           ondelete="RESTRICT",
                           requires=IS_EMPTY_OR(IS_IN_DB(db, "%s.id" % tablename))),
                     CommentsField(),
                     )
        # MetaFields (uuid, created_on, modified_on, owned_by_user, ...)
        # ถูกแปะอัตโนมัติโดย define_table()''',
    language="python")

add_para(doc,
    "จุดสำคัญที่ทำให้ DataModel แตกต่างจากการ define_table แบบเดิมของ web2py คือ "
    "(1) การประกาศ names tuple เป็นสัญญาของ class ว่าจะ export อะไร "
    "(2) define_table() เป็น wrapper ที่เพิ่ม MetaFields อัตโนมัติ "
    "(3) มี method model() สำหรับ module ที่เปิดใช้งาน และ defaults() สำหรับ module ที่ปิด "
    "(4) แต่ละ class จะถูก instantiate เพียงครั้งเดียวเมื่อมีการเรียกใช้ครั้งแรก (lazy)",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "3.2  Lazy Loader — หัวใจของ DataModel")

add_para(doc,
    "กลไกที่ทำให้ Eden ไม่ต้องโหลดตารางทั้งหมดขึ้นมาตั้งแต่ start คือ __getattr__ "
    "ใน DataModel เมื่อมีการเข้าถึง current.s3db.org_organisation ครั้งแรก "
    "Python จะเรียก __getattr__ ซึ่งไปค้นหาว่า tablename นั้นอยู่ใน names "
    "ของ DataModel subclass ตัวไหน จากนั้นจึง instantiate class นั้น "
    "ซึ่ง trigger การเรียก model() และสร้างตารางจริงใน database "
    "มีกลไกป้องกัน cycle/deadlock ผ่าน LOCK/LOAD markers",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''class DataModel:
    """ Base class for data models """
    _edenmodel = True
    LOCK = "eden_model_lock"
    LOAD = "eden_model_load"

    def __init__(self, module=None):
        self.cache = (current.cache.ram, 60)
        self.classes = {}
        self._customised = {}
        if not hasattr(current, "model"):
            current.model = {"config": {}, "components": {},
                             "methods": {}, "cmethods": {}, "hierarchies": {}}
        if module in mandatory_models or \\
           current.deployment_settings.has_module(module):
            env = self.model()       # subclass นิยามตารางที่นี่
        else:
            env = self.defaults()    # module ที่ปิด → โหลดเฉพาะ globals ขั้นต่ำ

    def __getattr__(self, name):      # MODEL AUTO-LOADER
        table = self.table(name, DEFAULT)
        if table is DEFAULT:
            raise AttributeError("undefined table: %s" % name)
        return table

    @staticmethod
    def define_table(tablename, *fields, meta=True, **args):
        db = current.db
        if hasattr(db, tablename):
            table = getattr(db, tablename)   # idempotent
        else:
            if meta:
                fields = fields + MetaFields()
            table = db.define_table(tablename, *fields, **args)
        return table

    @classmethod
    def table(cls, tablename, default=None, db_only=False):
        # 1. response.s3[tablename]  → 2. s3db.classes[tablename]
        # 3. db.tablename (โหลดแล้ว?) → 4. หา DataModel subclass ที่ names มี tablename
        #    แล้ว instantiate → เรียก model() → ตารางถูกสร้างขึ้น
        ...''',
    language="python")

add_heading_h2(doc, "3.3  MetaFields และ FieldTemplate")

add_para(doc,
    "ทุกตารางที่สร้างผ่าน DataModel.define_table() จะได้รับ MetaFields "
    "โดยอัตโนมัติ ซึ่งประกอบด้วย uuid (string สำหรับ sync ข้าม instances), "
    "created_on, modified_on, owned_by_user, owned_by_role, deleted_vk "
    "(soft delete), mci (Master Client Index สำหรับ deduplication) "
    "และ fields อื่น ๆ ที่จำเป็นสำหรับ multi-tenant + sync "
    "ส่วน FieldTemplate เป็น reusable field definition ที่สามารถแชร์ข้าม model ได้ "
    "เช่น organisation_type_id เป็น FieldTemplate ที่ใช้ในหลายตาราง",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================================
# บทที่ 4: CRUDRequest
# ============================================================
add_heading_h1(doc, "CRUDRequest — REST Dispatcher", num=4)

add_para(doc,
    "CRUDRequest (อดีตชื่อ S3RESTController) เป็นคลาสที่รับผิดชอบการแปล URL "
    "เป็น action ที่จะทำ และ dispatch ไปยัง handler ที่เหมาะสม "
    "เป็นคลาสที่ผู้พัฒนาแทบไม่ต้องเรียกโดยตรง เพราะ controller ทุกตัวใน Eden "
    "เรียกผ่าน helper crud_controller() ซึ่งเป็น wrapper ของ CRUDRequest",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "4.1  Controller แบบ Eden — บางที่สุด")

add_para(doc,
    "ลักษณะเด่นของ controller ใน Eden คือมันบางมาก แต่ละฟังก์ชันมักมีเพียง 1 บรรทัด "
    "เพราะ URL routing คือ API ของ Eden ทั้งระบบ การกำหนดพฤติกรรมเฉพาะ "
    "จะทำผ่าน prep/postp hooks หรือ customise_resource() ใน model แทน "
    "ตัวอย่างต่อไปนี้คือส่วนหนึ่งของ controllers/org.py",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''module = request.controller

def organisation():
    """ RESTful CRUD controller """
    # Defined in the Model for use from Multiple Controllers for unified menus
    return s3db.org_organisation_controller()

def office():
    """ RESTful CRUD controller """
    return s3db.org_office_controller()

def facility_type():
    """ รูปแบบ zero-arg: อนุมาน prefix/name จาก URL """
    return crud_controller()

def org_search():
    """ จำกัดเฉพาะ search_ac method (autocomplete) """
    s3.prep = lambda r: r.method == "search_ac"
    return crud_controller(module, "organisation")''',
    language="python")

add_heading_h2(doc, "4.2  การ Parse URL")

add_para(doc,
    "เมื่อ request เข้ามา CRUDRequest.__parse() จะแยก URL ออกเป็นส่วนประกอบ "
    "5 อย่าง คือ id, component_name, component_id, method และ representation "
    "โดย representation มาจาก extension ของ URL เช่น .csv, .json, .xml "
    "และ http method มาจาก HTTP verb (GET/POST/PUT/DELETE) "
    "ตัวอย่างเช่น URL /eden/org/organisation/123/branch/45.csv จะถูก parse เป็น "
    "id=123, component_name='branch', component_id=45, method=None, "
    "representation='csv', http='GET'",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "4.3  CRUDRequest.__call__ — หัวใจของ Dispatcher")

add_para(doc,
    "เมื่อ CRUDRequest ถูกสร้างขึ้นและถูกเรียกด้วย r(**attr) มันจะทำงาน "
    "ตามลำดับขั้นตอนต่อไปนี้ (1) เรียก prep hook ถ้ามี (2) ตั้งค่า view "
    "และ Content-Type header (3) หา handler ที่เหมาะสมตามลำดับความสำคัญ "
    "(4) เรียก handler แล้วรับ output (5) เรียก postp hook เพื่อปรับ output "
    "ก่อนส่งให้ view โค้ดต่อไปนี้คือส่วนสำคัญของ CRUDRequest.__call__",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''def __call__(self, **attr):
    response = current.response
    s3 = response.s3
    self.next = None
    bypass = False; output = None
    representation = self.representation

    # 1) PRE-PROCESS HOOK
    preprocess = s3.get("prep") if s3 else None
    if preprocess:
        pre = preprocess(self)
        representation = self.representation
        if not pre: self.error(400, current.ERROR.BAD_REQUEST)
        elif isinstance(pre, dict):
            bypass = pre.get("bypass", False)
            output = pre.get("output")

    # 2) VIEW + CONTENT-TYPE
    if representation not in ("html", "popup"):
        response.view = "xml.html"
    response.headers["Content-Type"] = \\
        s3.content_type.get(representation, "text/html")

    # 3) RESOLVE HANDLER (ตามลำดับความสำคัญ)
    custom_action = self.custom_action or \\
                   current.s3db.get_method(self.tablename,
                                           component=self.component_name,
                                           method=self.method)
    http, method = self.http, self.method
    if not bypass:
        if custom_action:
            handler = custom_action
        else:
            handler = self.default_methods.get(method)
            if not handler:
                m = "import" if http in ("PUT", "POST") else None
                if not method and (http not in ("GET", "POST") or \\
                                   self.transformable(method=m)):
                    handler = RESTful
                elif http in HTTP_METHODS:
                    handler = BasicCRUD
                else:
                    self.error(405, current.ERROR.BAD_METHOD)
        if isinstance(handler, type):
            handler = handler()    # lazy instantiation
        output = handler(self, **attr)

    # 4) POST-PROCESS HOOK
    postprocess = s3.get("postp") if s3 else None
    if postprocess is not None:
        output = postprocess(self, output)
    if output is not None and isinstance(output, dict):
        output["r"] = self
    return output''',
    language="python")

add_heading_h2(doc, "4.4  ลำดับการ Resolve Handler")

add_para(doc,
    "การหา handler ที่จะใช้จะพิจารณาตามลำดับต่อไปนี้ ซึ่งเป็นกลไกที่ทำให้ "
    "Eden ยืดหยุ่นมาก เพราะสามารถ override พฤติกรรมเริ่มต้นได้ทุกระดับ "
    "ตั้งแต่เพิ่ม custom method ใหม่ ไปจนถึงเปลี่ยน handler ของ method ที่มีอยู่",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_bullet(doc, "① custom_action จาก s3db.get_method() — ลงทะเบียนผ่าน set_method()")
add_bullet(doc, "② default_methods[method] — method มาตรฐาน เช่น report, map, profile, hierarchy")
add_bullet(doc, "③ RESTful — ถ้าไม่มี method และ request เป็น transformable หรือเป็น PUT/POST ที่ไม่ใช่ form")
add_bullet(doc, "④ BasicCRUD — สำหรับ GET/POST ทั่วไปที่ใช้ form-based interaction")
add_bullet(doc, "⑤ 405 BAD_METHOD — ถ้าไม่ตรงเงื่อนไขใด ๆ ข้างต้น")

# ============================================================
# บทที่ 5: CRUDResource
# ============================================================
add_heading_h1(doc, "CRUDResource — Resource Abstraction Layer", num=5)

add_para(doc,
    "CRUDResource (อดีตชื่อ S3Resource) เป็นคลาสที่ห่อหุ้ม PyDAL Table "
    "พร้อมด้วย filter, components, link tables และ super-entity relationships "
    "เปิดเผย method มาตรฐาน เช่น count, select, load, insert, update, delete "
    "และ export_xml, import_xml ที่ทำงานภายใต้กฎของ S3 framework โดยอัตโนมัติ "
    "เช่น authorization, audit logging, ownership/ACL และ onaccept hooks "
    "ทำให้ developer ไม่ต้องจัดการเรื่องเหล่านี้เองในทุก ๆ ครั้ง",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "5.1  การสร้าง CRUDResource")

add_para(doc,
    "CRUDResource ถูกสร้างโดย CRUDRequest เมื่อรับ URL เข้ามา ตัวอย่างด้านล่าง "
    "แสดง __init__ และ method ที่ใช้บ่อยที่สุด คือ select() และ count() "
    "สังเกตว่าทั้งสอง method จะใช้ S3ResourceFilter เพื่อสร้าง DAL Query "
    "แล้วจึง delegate ไปยัง ResourceData สำหรับดึง row จริงจาก database",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''class CRUDResource:
    """ A 'resource' = a set of records in a table + their component references.

        resource = CRUDResource(table)
        …extended standard methods to access and manipulate data while
        respecting current authorization and other S3 framework rules.
    """

    def __init__(self, tablename, id=None, prefix=None, uid=None,
                 filter=None, vars=None, parent=None, linked=None,
                 linktable=None, alias=None, components=None,
                 filter_component=None, include_deleted=False,
                 approved=True, unapproved=False, extra_filters=None):
        self.table = s3db[tablename]                       # lazy-load model
        auth.permission.set_default_approver(self.table)
        self.accessible_query = auth.s3_accessible_query  # authz hook
        self.rfilter = None                                # S3ResourceFilter
        self._rows = None; self._length = None
        self.components = S3Components(self, components)   # component graph
        if parent is None:
            self.build_query(id=id, uid=uid, filter=filter, vars=vars,
                             extra_filters=extra_filters,
                             filter_component=filter_component)

    def select(self, fields, start=0, limit=None, left=None,
               orderby=None, groupby=None, distinct=False, virtual=True,
               count=False, getids=False, as_rows=False,
               represent=False, …):
        data = ResourceData(self, fields, start=start, limit=limit,
                            left=left, orderby=orderby, groupby=groupby,
                            distinct=distinct, virtual=virtual, count=count,
                            getids=getids, as_rows=as_rows,
                            represent=represent, …)
        return data.rows if as_rows else data

    def count(self, left=None, distinct=False):
        if self.rfilter is None: self.build_query()
        if self._length is None:
            self._length = self.rfilter.count(left=left, distinct=distinct)
        return self._length''',
    language="python")

add_heading_h2(doc, "5.2  insert() — Authorization + Audit + Ownership + Hooks")

add_para(doc,
    "method insert() ของ CRUDResource ไม่ได้แค่เรียก table.insert() ของ PyDAL "
    "แต่ทำงาน 5 ขั้นตอน (1) ตรวจสอบ authorization ผ่าน auth.s3_has_permission() "
    "(2) เรียก table.insert() จริง (3) เขียน audit log (4) อัปเดต super-entity "
    "และ ownership/ACL (5) เรียก onaccept hook ทำให้ทุก record ที่ถูก insert "
    "ผ่าน Eden มี audit trail, ownership, และ trigger hooks ที่ต้องการได้โดยอัตโนมัติ",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''def insert(self, **fields):
    table = self.table
    tablename = self.tablename
    # 1) Authorization
    if not current.auth.s3_has_permission("create", tablename):
        raise S3PermissionError("Operation not permitted: INSERT INTO %s" \\
                                % tablename)
    # 2) Insert
    record_id = self.table.insert(**fields)
    if record_id:
        # 3) Audit log
        current.audit("create", self.prefix, self.name, record=record_id)
        record = Storage(fields); record.id = record_id
        s3db = current.s3db
        # 4) Super-entity propagation + ownership/ACL
        s3db.update_super(table, record)
        auth = current.auth
        auth.s3_set_record_owner(table, record_id)
        auth.s3_make_session_owner(table, record_id)
        # 5) onaccept hook
        s3db.onaccept(tablename, record, method="create")
    return record_id''',
    language="python")

add_heading_h2(doc, "5.3  Component Graph — master + sub-records")

add_para(doc,
    "ลักษณะเฉพาะของ CRUDResource คือการรองรับ component relationship "
    "เช่น org_organisation (master) มี component ชื่อ branch (org_organisation ที่เป็น branch ของอีกราย) "
    "หรือ inv_inv_item (master) มี component ชื่อ stock_inv_item (รายการ stock movement) "
    "URL /eden/org/organisation/123/branch/45.csv จึงเข้าถึง branch รหัส 45 ของ organisation รหัส 123 ได้ "
    "โดย CRUDResource จะสร้าง S3Components object เก็บกราฟของ component ทั้งหมดที่เป็นไปได้ "
    "และอนุญาตให้เข้าถึง component เฉพาะที่กำหนดได้ผ่าน r.component",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================================
# บทที่ 6: BasicCRUD
# ============================================================
add_heading_h1(doc, "BasicCRUD — Standard CRUD Methods", num=6)

add_para(doc,
    "BasicCRUD (อดีตชื่อ S3CRUD) เป็น CRUDMethod subclass ที่รับผิดชอบ "
    "operation มาตรฐาน ได้แก่ create, read, update, delete, select (list), "
    "validate, review โดย method apply_method() จะ dispatch ไปยัง method ย่อย "
    "ตาม r.method และ r.http พารามิเตอร์ นี่คือ handler เริ่มต้นสำหรับ "
    "URL ที่ไม่มี method พิเศษและเป็น GET/POST ธรรมดา",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "6.1  apply_method() — Method Dispatch")

add_code_block(doc, '''class BasicCRUD(CRUDMethod):
    def apply_method(self, r, **attr):
        self.form = self.resource.get_config("crud_form") or DefaultForm()
        method = self.method
        if r.http == "DELETE" or method == "delete":
            output = self.delete(r, **attr)
        elif method == "create":
            output = self.create(r, **attr)
        elif method == "read":
            output = self.read(r, **attr)
        elif method == "update":
            output = self.update(r, **attr)
        elif method == "list":
            output = self.select(r, **attr)
        elif method in ("datatable", "datatable_f"):
            output = self.select(r, list_type="datatable", …)
        elif method in ("datalist", "datalist_f"):
            output = self.select(r, list_type="datalist", …)
        elif method == "validate":
            output = self.validate(r, **attr)
        elif method == "review":
            output = self.review(r, **attr) if r.record \\
                     else self.unapproved(r, **attr)
        else:
            r.error(405, current.ERROR.BAD_METHOD)
        return output''',
    language="python")

add_heading_h2(doc, "6.2  create() — อ่าน Hooks จาก Config")

add_para(doc,
    "ภายใน method create() จะอ่าน hooks จาก config ที่ลงทะเบียนไว้กับตาราง "
    "ผ่าน DataModel.configure() ส่งให้ form builder เพื่อ execute ในจังหวะที่เหมาะสม "
    "การแยก create_onvalidation ออกจาก onvalidation ทั่วไป ทำให้สามารถเขียน hook "
    "ที่ทำงานเฉพาะตอน create ได้โดยไม่กระทบ update หรือ delete",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''def create(self, r, **attr):
    …
    onvalidation = get_config("create_onvalidation") or \\
                   get_config("onvalidation")
    onaccept     = get_config("create_onaccept") or \\
                   get_config("onaccept")
    …
    output["form"] = self.form(request=request, resource=resource,
                                data=self.data, record_id=original,
                                from_table=from_table,
                                onvalidation=onvalidation,
                                onaccept=onaccept,
                                link=link, hierarchy=hierarchy,
                                message=message,
                                subheadings=subheadings,
                                format=representation)''',
    language="python")

add_heading_h2(doc, "6.3  CRUDMethod — Base Class")

add_para(doc,
    "CRUDMethod เป็น base class ของทุก method handler ใน Eden มี __call__ "
    "ที่จะเรียก apply_method() ใน subclass นี่คือสถาปัตยกรรมที่ทำให้ "
    "การเพิ่ม method ใหม่เป็นเรื่องง่าย เพียงแค่ subclass CRUDMethod "
    "และ implement apply_method() ก็สามารถลงทะเบียนเป็น default method ใหม่ได้ "
    "ผ่าน CRUDRequest.default_methods หรือเป็น custom method ของตารางใดตารางหนึ่ง",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================================
# บทที่ 7: REST Request Lifecycle
# ============================================================
add_heading_h1(doc, "REST Request Lifecycle", num=7)

add_para(doc,
    "เพื่อให้เห็นภาพรวมของการทำงานทั้งระบบ บทนี้อธิบายวงจรอายุของ request หนึ่งรายการ "
    "ตั้งแต่เข้าระบบจนส่ง response ออกไป ตัวอย่าง URL ที่ใช้คือ "
    "GET /eden/org/organisation/123/branch/45.csv?organisation_type.name=NGO "
    "ซึ่งเป็น request ที่ครอบคลุมทุกส่วนของ S3 framework ทั้ง master record, "
    "component, query filter และ format negotiation",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_image(doc, os.path.join(ASSETS, "diagram_lifecycle.png"),
          width_cm=15,
          caption="รูปที่ 7.1  วงจรอายุของ request หนึ่งรายการ ตั้งแต่ HTTP request จนถึง response")

add_heading_h2(doc, "7.1  ขั้นตอนที่ 1 — Routing")

add_para(doc,
    "เมื่อ HTTP request เข้ามา web2py router จะส่งไปยัง controllers/org.py::organisation() "
    "ซึ่งมีเพียง 1 บรรทัด คือ return s3db.org_organisation_controller() "
    "function นี้ถูกนิยามใน model (modules/s3db/org.py) เพื่อให้สามารถเรียกจาก controller "
    "หลายตัวได้ (เช่น menu ของ organisation อาจปรากฏในทุก module) "
    "ภายใน org_organisation_controller() จะเรียก crud_controller(prefix='org', "
    "resourcename='organisation', **attr) ซึ่งเป็น helper ที่ instantiate CRUDRequest",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "7.2  ขั้นตอนที่ 2 — Parse URL")

add_para(doc,
    "CRUDRequest.__parse() จะแยก URL ออกเป็นส่วนประกอบ สำหรับ URL "
    "/eden/org/organisation/123/branch/45.csv จะได้ผลลัพธ์ดังนี้ "
    "id=123 (record ID ของ master), component_name='branch' (ชื่อ component), "
    "component_id=45 (record ID ของ component), method=None (ไม่มี method ใน URL), "
    "representation='csv' (มาจาก extension), http='GET' (HTTP verb) "
    "ส่วน query string ?organisation_type.name=NGO จะถูก parse แยกไว้ใน request.vars",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "7.3  ขั้นตอนที่ 3 — Instantiate CRUDResource")

add_para(doc,
    "CRUDRequest สร้าง CRUDResource('org_organisation', id=123) "
    "ซึ่ง trigger lazy-load ของ OrgOrganisationModel เพื่อ define table ถ้ายังไม่ถูก define "
    "จากนั้น build_query() จะสร้าง S3ResourceFilter ซึ่งแปล query string "
    "?organisation_type.name=NGO ให้เป็น DAL Query พร้อม JOIN ที่จำเป็น "
    "และ S3Components จะโหลด component 'branch' เตรียมไว้ให้",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "7.4  ขั้นตอนที่ 4 — PRE-PROCESS Hook (prep)")

add_para(doc,
    "ก่อนเรียก handler ระบบจะเรียก prep hook ถ้ามี ซึ่งติดตั้งผ่าน response.s3.prep "
    "hook นี้สามารถ bypass request (return False เพื่อส่ง 400), "
    "inject output (return dict ที่มี 'output' key), หรือแค่ modify request object "
    "เช่นเปลี่ยน representation หรือเพิ่ม filter เพิ่มเติม "
    "นี่คือกลไกหลักที่ controller ใช้ปรับพฤติกรรมของ endpoint โดยไม่ต้องเขียน handler เอง",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "7.5  ขั้นตอนที่ 5 — Resolve Handler")

add_para(doc,
    "ระบบจะหา handler ตามลำดับความสำคัญ (อธิบายในบทที่ 4) "
    "สำหรับ request ตัวอย่างนี้ ไม่มี custom method และไม่มี default method "
    "จึงตัดสินใจใช้ BasicCRUD (เพราะเป็น GET request แบบธรรมดา) "
    "BasicCRUD.apply_method() จะเลือก read() เพราะมี id และไม่ใช่ create/update/delete "
    "และ representation='csv' จะถูกใช้ใน read() เพื่อส่งต่อไปยัง DataExporter.csv(resource)",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "7.6  ขั้นตอนที่ 6 — Execute Handler")

add_para(doc,
    "handler ที่ถูกเลือก (BasicCRUD.read) จะทำงาน โดยใช้ CRUDResource "
    "เพื่อ query ข้อมูลจาก database ผ่าน count() และ select() "
    "ซึ่ง delegate ไปยัง ResourceData และ S3ResourceFilter "
    "จากนั้น DataExporter.csv(resource) จะแปลง result เป็น CSV string "
    "ผ่าน XLSWriter หรือ CSV writer ใน modules/core/formats/",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "7.7  ขั้นตอนที่ 7 — POST-PROCESS Hook (postp)")

add_para(doc,
    "หลัง handler ทำงานเสร็จ ระบบจะเรียก postp hook ถ้ามี "
    "ซึ่งสามารถ modify output ก่อนส่งให้ view ได้ "
    "เช่นเพิ่มข้อมูลใน response, แปลง format, หรือ cache "
    "จากนั้น output['r'] = self เพื่อให้ view เข้าถึง request object ได้ "
    "สุดท้าย view template (xml.html สำหรับ non-HTML representation) "
    "จะ render output เป็น response body และส่ง HTTP response กลับไป "
    "พร้อม Content-Type header ที่ตรงกับ representation",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================================
# บทที่ 8: Content Negotiation
# ============================================================
add_heading_h1(doc, "Content Negotiation", num=8)

add_para(doc,
    "S3 framework รองรับ content negotiation ผ่าน URL extension "
    "เช่น .json, .xml, .csv, .xls, .pdf, .kml, .cap ฯลฯ "
    "extension นี้จะถูก parse โดย s3_get_extension() ใน CRUDRequest.__parse() "
    "และใช้ทั้งในการ dispatch handler และตั้งค่า Content-Type header "
    "ตารางด้านล่างสรุป extension ที่รองรับทั้งหมด พร้อม handler ที่ใช้ "
    "และ output method ที่สร้าง response",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "8.1  ตาราง Extension → Handler → Output")

add_table(doc,
    headers=["URL extension", "Handler ที่ถูก dispatch", "Output method", "Content-Type"],
    rows=[
        [".html / (none)",  "BasicCRUD",               "create/read/update/select view",  "text/html"],
        [".json",           "BasicCRUD (read/select)",  "DataExporter.json()",             "application/json"],
        [".s3json",         "RESTful",                  "resource.export_xml(as_json=True)", "application/json"],
        [".geojson",        "RESTful",                  "export_xml + maxbounds",          "application/json"],
        [".xml",            "RESTful",                  "resource.export_xml() + XSLT",     "text/xml"],
        [".csv",            "BasicCRUD",                "DataExporter.csv()",              "text/csv"],
        [".s3csv (PUT)",    "RESTful",                  "resource.import_xml(source_type=\"csv\")", "JSON result"],
        [".xls",            "BasicCRUD",                "DataExporter.xls() → XLSWriter",  "Excel MIME"],
        [".xlsx",           "BasicCRUD",                "DataExporter.xlsx() → XLSXWriter", "Excel MIME"],
        [".pdf",            "BasicCRUD",                "DataExporter.pdf() → PDFWriter", "application/pdf"],
        [".card",           "BasicCRUD",                "DataExporter.pdfcard()",          "application/pdf"],
        [".shp",            "BasicCRUD",                "DataExporter.shp() → SHPWriter",  "shapefile"],
        [".svg",            "BasicCRUD",                "DataExporter.svg() → SVGWriter",  "image/svg+xml"],
        [".rss / .georss",  "RESTful + XSLT",           "export_xml + GeoRSS stylesheet",  "application/rss+xml"],
        [".kml",            "RESTful + XSLT",           "export_xml + KML stylesheet",     "application/vnd.google-earth.kml+xml"],
        [".gpx / .osm",     "RESTful + XSLT",           "export_xml + stylesheet",          "XML"],
        [".cap",            "RESTful + XSLT",           "CAP stylesheet",                  "XML"],
        [".aadata",         "BasicCRUD",                "DataTable server-side JSON",      "application/json"],
        [".plain",          "BasicCRUD",                "plain.html view",                 "text/html"],
        [".dl",             "BasicCRUD",                "download-list view",              "text/html"],
    ],
    col_widths_cm=[3.0, 3.8, 5.5, 3.7],
    font_size=8.5, header_size=9)

add_heading_h2(doc, "8.2  ความแตกต่างระหว่าง BasicCRUD และ RESTful")

add_para(doc,
    "จากตารางข้างต้น จะเห็นว่ามี handler 2 ตัวที่ใช้บ่อย คือ BasicCRUD และ RESTful "
    "ความแตกต่างคือ BasicCRUD ใช้สำหรับ interactive form-based operations "
    "(HTML views, form submission) และ structured exports (CSV, XLS, PDF, SHP) "
    "ในขณะที่ RESTful ใช้สำหรับ tree-structured API ที่ส่งออกเป็น S3XML format "
    "ซึ่งสามารถแปลงผ่าน XSLT เป็น KML, CAP, RSS ฯลฯ ได้ "
    "การเลือกใช้ขึ้นอยู่กับว่า extension นั้นเป็น transformable format "
    "(ต้องผ่าน XSLT) หรือเปล่า ซึ่งตรวจสอบโดย CRUDRequest.transformable()",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_note_box(doc, "S3XML — รูปแบบ canonical ของ Eden",
    "S3XML เป็น XML format เฉพาะของ Eden ที่ใช้เป็น canonical representation "
    "ของ resource ทุกชนิด สามารถ export เป็น JSON หรือ XML ได้ และ import กลับเข้าระบบได้ "
    "การส่งออกเป็น .kml หรือ .cap ทำได้โดย export เป็น S3XML ก่อน แล้วจึง apply XSLT stylesheet "
    "เพื่อแปลงเป็น format ที่ต้องการ ทำให้ Eden สามารถรองรับ format ใหม่ได้ง่าย "
    "เพียงแค่เพิ่ม XSLT stylesheet โดยไม่ต้องเขียน exporter ใหม่",
    color="0E7490", bg="ECFDF5")

# ============================================================
# บทที่ 9: Default Methods + Custom REST Methods
# ============================================================
add_heading_h1(doc, "Default Methods และ Custom REST Methods", num=9)

add_para(doc,
    "นอกจาก CRUD พื้นฐานแล้ว S3 framework มาพร้อมกับ method มาตรฐานอีกกว่า 20 ตัว "
    "ที่ลงทะเบียนใน CRUDRequest.default_methods และสามารถเรียกผ่าน URL "
    "?method=<name> หรือเป็น path segment เช่น /eden/org/organisation/report "
    "นอกจากนี้ยังสามารถเพิ่ม custom REST method ของตารางใดตารางหนึ่งได้ "
    "ผ่าน s3db.set_method()",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "9.1  Default Methods Table")

add_table(doc,
    headers=["Method name", "Handler class", "Use case"],
    rows=[
        ["create",                "BasicCRUD",             "สร้าง record ใหม่ผ่าน form"],
        ["read",                  "BasicCRUD",             "แสดง record เดียวแบบ read-only"],
        ["update",                "BasicCRUD",             "แก้ไข record ผ่าน form"],
        ["delete",                "BasicCRUD",             "ลบ record (soft delete)"],
        ["list",                  "BasicCRUD",             "แสดงรายการแบบ list view"],
        ["datatable",             "BasicCRUD",             "AJAX-backed DataTable list"],
        ["datalist",              "BasicCRUD",             "Card-style data list"],
        ["report",                "S3Report",              "Aggregated / pivot reports"],
        ["map",                   "S3Map",                 "Geographic visualization"],
        ["profile",              "S3Profile",              "Single-record detail view"],
        ["hierarchy",            "S3HierarchyCRUD",        "Tree-view of hierarchical data"],
        ["summary",              "S3Summary",              "Dashboard summary"],
        ["organize",             "S3Organizer",            "Bulk update via checkboxes"],
        ["timeplot",             "TimePlot",               "Time-series visualization"],
        ["merge / deduplicate",  "S3Merge",                "Merge duplicate records"],
        ["xform",                "S3XForms",               "XForms mobile data entry"],
        ["filter",               "FilterManager",           "Saved filter UI"],
        ["columns",              "ColumnConfigManager",     "Configure visible columns"],
        ["fields",               "RESTful",                 "Schema introspection"],
        ["options",              "RESTful",                 "Dropdown options"],
        ["import",               "SpreadsheetImporter",     "Bulk CSV/XLS import"],
        ["cms",                  "S3CMS",                   "Content blocks"],
        ["compose",              "S3Compose",               "Compose multiple records"],
        ["search_ac",            "(function)",              "Autocomplete endpoint"],
    ],
    col_widths_cm=[3.5, 4.0, 8.5],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "9.2  การลงทะเบียน Custom REST Method")

add_para(doc,
    "เมื่อต้องการเพิ่ม endpoint ใหม่สำหรับตารางใดตารางหนึ่ง เช่น /eden/pr/person/5/check_in "
    "สามารถทำได้โดยเรียก s3db.set_method() ใน model ของตารางนั้น "
    "เมื่อ CRUDRequest ทำงาน จะเรียก s3db.get_method() เพื่อหา custom method ก่อนเสมอ "
    "ถ้าเจอ จะใช้ action นี้แทน BasicCRUD/RESTful โค้ดด้านล่างแสดงทั้ง set และ get",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''@classmethod
def set_method(cls, tablename, component=None, method=None, action=None):
    """ Configure a URL method for a table (or component in the table's context). """
    methods  = current.model["methods"]
    cmethods = current.model["cmethods"]
    if not method: raise SyntaxError("No method specified")
    if not component:
        methods.setdefault(method, {})[tablename] = action
    else:
        cmethods.setdefault(method, {})\\
                .setdefault(component, {})[tablename] = action

@classmethod
def get_method(cls, tablename, component=None, method=None):
    """ Lookup (called by CRUDRequest before falling back to defaults). """
    if component:
        cmethods = current.model["cmethods"]
        if method in cmethods and component in cmethods[method] \\
           and tablename in cmethods[method][component]:
            return cmethods[method][component][tablename]
        hook = cls.get_component(tablename, component)
        tablename = hook["tablename"] if hook else None
    methods = current.model["methods"]
    return methods.get(method, {}).get(tablename) if method else None''',
    language="python")

add_heading_h3(doc, "ตัวอย่างการใช้งานจริง")

add_code_block(doc, '''def check_in_func(r, **attr):
    """ Handler for check_in method """
    person_id = r.id
    # … business logic …
    return {}

# registers  /eden/pr/person/5/check_in
s3db.set_method("pr_person", method="check_in", action=check_in_func)''',
    language="python")

# ============================================================
# บทที่ 10: Hooks & Extensibility
# ============================================================
add_heading_h1(doc, "Hooks & Extensibility", num=10)

add_para(doc,
    "S3 framework มีระบบ hooks ที่ครอบคลุม ทำให้สามารถปรับพฤติกรรมของระบบได้ "
    "โดยไม่ต้องแก้ไขโค้ดหลัก hooks แบ่งเป็น 3 ระดับ คือ "
    "(1) request-level hooks (prep, postp) ที่ติดตั้งผ่าน response.s3 "
    "(2) record-level hooks (onvalidation, onaccept, ondelete) ที่ลงทะเบียนผ่าน configure() "
    "(3) method-specific hooks (create_onvalidation, update_onaccept) ที่ทำงานเฉพาะ operation",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "10.1  Request-level Hooks (prep / postp)")

add_para(doc,
    "prep และ postp เป็น hooks ที่ติดตั้งได้ใน controller ผ่าน response.s3 "
    "prep ทำงานก่อน handler ถูกเรียก สามารถ bypass request, inject output "
    "หรือ modify request object ได้ postp ทำงานหลัง handler ทำงานเสร็จ "
    "สามารถ modify output ก่อนส่งให้ view ได้ hooks เหล่านี้เป็นวิธีหลัก "
    "ที่ controller ใช้ปรับพฤติกรรมของ endpoint โดยไม่ต้องเขียน handler เอง",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''def org_search():
    """ จำกัดเฉพาะ search_ac method (autocomplete) """
    s3.prep = lambda r: r.method == "search_ac"
    return crud_controller(module, "organisation")

def custom_endpoint():
    """ ใช้ prep เพื่อ inject output และ bypass handler """
    def prep(r):
        if r.http == "GET":
            return {"output": {"items": [...]}, "bypass": True}
        return None  # proceed normally
    s3.prep = prep
    return crud_controller("org", "organisation")''',
    language="python")

add_heading_h2(doc, "10.2  Record-level Hooks (onvalidation / onaccept)")

add_para(doc,
    "onvalidation และ onaccept เป็น hooks ที่ทำงานในระดับ record "
    "onvalidation ทำงานก่อน record ถูก insert/update เพื่อ validate ข้อมูลเพิ่มเติม "
    "onaccept ทำงานหลัง record ถูก insert/update/delete เพื่อ trigger side effects "
    "เช่น ส่ง notification, อัปเดต cache, sync ข้อมูล ฯลฯ "
    "hooks เหล่านี้ลงทะเบียนผ่าน DataModel.configure() ใน model",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''@classmethod
def onaccept(cls, table, record, method="create"):
    tablename = table._tablename if hasattr(table, "_tablename") else table
    onaccept = cls.get_config(tablename, "%s_onaccept" % method,
                              cls.get_config(tablename, "onaccept"))
    if onaccept:
        if "vars" not in record:
            record = Storage(vars=Storage(record), errors=Storage())
        callback(onaccept, record, tablename=tablename)

@classmethod
def onvalidation(cls, table, record, method="create"):
    onvalidation = cls.get_config(tablename, "%s_onvalidation" % method,
                                  cls.get_config(tablename, "onvalidation"))
    if "vars" not in record:
        record = Storage(vars=Storage(record), errors=Storage())
    if onvalidation:
        callback(onvalidation, record, tablename=tablename)
    return record.errors''',
    language="python")

add_heading_h2(doc, "10.3  รายการ Hooks ทั้งหมด")

add_table(doc,
    headers=["Hook name", "Level", "Trigger เมื่อ", "ลงทะเบียนผ่าน"],
    rows=[
        ["prep",                       "Request", "ก่อน handler ทำงาน",            "response.s3.prep"],
        ["postp",                      "Request", "หลัง handler ทำงาน",            "response.s3.postp"],
        ["onvalidation",               "Record",  "ก่อน insert/update (ทั่วไป)",    "configure(onvalidation=...)"],
        ["create_onvalidation",        "Record",  "ก่อน create เท่านั้น",          "configure(create_onvalidation=...)"],
        ["update_onvalidation",        "Record",  "ก่อน update เท่านั้น",          "configure(update_onvalidation=...)"],
        ["onaccept",                   "Record",  "หลัง insert/update (ทั่วไป)",   "configure(onaccept=...)"],
        ["create_onaccept",            "Record",  "หลัง create เท่านั้น",          "configure(create_onaccept=...)"],
        ["update_onaccept",            "Record",  "หลัง update เท่านั้น",          "configure(update_onaccept=...)"],
        ["ondelete",                   "Record",  "หลัง delete",                   "configure(ondelete=...)"],
        ["onvalidate_import",          "Import",  "ระหว่าง import pipeline",       "configure(onvalidate_import=...)"],
        ["customise_resource",         "Resource", "เมื่อ CRUDResource ถูกสร้าง",  "settings.customise_resource()"],
        ["customise_controller",        "Controller","เมื่อ controller ถูกเรียก",    "settings.customise_controller()"],
    ],
    col_widths_cm=[3.5, 2.0, 5.5, 5.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "10.4  add_custom_callback — การ Chain Hooks")

add_para(doc,
    "นอกจากการตั้งค่า hook แล้ว DataModel ยังมี add_custom_callback() "
    "ที่ chain hook ใหม่เข้ากับ chain ที่มีอยู่โดยไม่ทับซ้อน "
    "ทำให้ template หลายตัวสามารถเพิ่ม hook ไปยังตารางเดียวกันได้ "
    "โดยไม่ต้องกังวลเรื่องการ override ซึ่งเป็นสิ่งสำคัญใน multi-template deployment",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================================
# บทที่ 11: File Map
# ============================================================
add_heading_h1(doc, "File Map โครงสร้าง Repository", num=11)

add_para(doc,
    "เพื่อให้ developer สามารถนำทางใน repository ได้สะดวก "
    "บทนี้สรุปโครงสร้างไฟล์และโฟลเดอร์หลักของ Eden พร้อมคำอธิบาย "
    "สังเกตว่าโค้ดของ S3 framework แท้ทั้งหมดอยู่ใน modules/core/ "
    "ส่วน data model ของแต่ละ domain อยู่ใน modules/s3db/ "
    "และ controller บาง ๆ อยู่ใน controllers/",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "11.1  modules/core/ — S3 Framework")

add_table(doc,
    headers=["Path", "หน้าที่"],
    rows=[
        ["controller.py",            "CRUDRequest, crud_controller (~1400 บรรทัด)"],
        ["model/datamodel.py",       "DataModel base + lazy loader + hooks (~1840 บรรทัด)"],
        ["model/fields.py",           "MetaFields, FieldTemplate ฯลฯ"],
        ["model/dynamic.py",          "DynamicTableModel (runtime-defined tables)"],
        ["model/options.py",          "WorkflowOptions"],
        ["resource/resource.py",      "CRUDResource (~2920 บรรทัด)"],
        ["resource/data.py",          "ResourceData (actual select/extraction, ~1428 บรรทัด)"],
        ["resource/query.py",         "S3ResourceQuery, S3Joins, S3URLQuery (~3245 บรรทัด)"],
        ["resource/rfilter.py",       "S3ResourceFilter"],
        ["resource/rtb.py",            "S3ResourceTree (tree export)"],
        ["resource/components.py",    "S3Components"],
        ["resource/exporter.py",       "DataExporter facade → formats/*"],
        ["resource/importer.py",       "S3XML/JSON import pipeline (~103 KB)"],
        ["resource/delete.py, merge.py","DeleteProcess, MergeProcess"],
        ["methods/base.py",            "CRUDMethod base class"],
        ["methods/crud.py",            "BasicCRUD (~3367 บรรทัด)"],
        ["methods/rest.py",            "RESTful (~525 บรรทัด)"],
        ["methods/hcrud.py",           "S3HierarchyCRUD"],
        ["methods/report.py",          "S3Report"],
        ["methods/profile.py",         "S3Profile"],
        ["methods/mapview.py",         "S3Map"],
        ["methods/summary.py",         "S3Summary"],
        ["methods/timeplot.py",        "TimePlot"],
        ["methods/merge.py",           "S3Merge"],
        ["methods/organizer.py",       "S3Organizer"],
        ["methods/mobile.py",          "S3MobileCRUD"],
        ["methods/xforms.py",          "S3XForms"],
        ["methods/select.py",          "Select"],
        ["methods/filtermgr.py",       "FilterManager"],
        ["methods/colmgr.py",          "ColumnConfigManager"],
        ["methods/cms.py",             "S3CMS"],
        ["methods/compose.py",         "S3Compose"],
        ["methods/grouped.py",         "S3GroupedItemsReport"],
        ["methods/spreadsheet.py",     "SpreadsheetImporter"],
        ["formats/xml.py",             "S3XML, S3XMLFormat, S3EntityResolver (canonical codec)"],
        ["formats/pdf.py",             "PDFWriter, EdenDocTemplate"],
        ["formats/card.py",            "PDFCardWriter, PDFCardLayout"],
        ["formats/xls.py",             "XLSWriter"],
        ["formats/xlsx.py",            "XLSXWriter, XLSXPivotTableWriter"],
        ["formats/shp.py",             "SHPWriter"],
        ["formats/svg.py",             "SVGWriter"],
        ["filters/",                   "OptionsFilter, HierarchyFilter, DateFilter, LocationFilter, ..."],
        ["aaa/",                       "AuthS3, permissions, audit, lock, oauth, masterkey, consent"],
        ["gis/",                       "Mapping (current.gis)"],
        ["msg/",                       "Messaging (current.msg)"],
        ["sync/",                      "Synchronization (current.sync) + adapters/"],
        ["tools/",                     "IS_ONE_OF, S3Hierarchy, s3_get_extension, ..."],
        ["ui/",                        "DataTable, S3DataList, layouts, menus, widgets"],
    ],
    col_widths_cm=[5.0, 11.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "11.2  modules/s3db/ — 41 DataModel Modules")

add_para(doc,
    "แต่ละไฟล์ใน modules/s3db/ บรรจุ DataModel subclass ของ domain หนึ่ง "
    "โดยทั่วไปจะมี controller คู่กันใน controllers/ เช่น "
    "controllers/org.py ↔ modules/s3db/org.py ไฟล์ s3.py เป็นของ framework เอง "
    "บรรจุ S3HierarchyModel, S3DashboardModel, S3ImportJobModel, S3DynamicTablesModel "
    "ส่วนไฟล์อื่น ๆ เป็นของ domain ต่าง ๆ",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "11.3  โครงสร้าง Top-level")

add_code_block(doc, '''eden/
├── controllers/      ← 44+ controller files (1 file = 1 module, mostly thin)
├── models/           ← web2py run-order:
│   ├── 000_1st_run.py    → import core as s3base, instantiate S3Config
│   ├── 00_db.py          → create current.db = DAL(...)  (plain PyDAL!)
│   ├── 00_settings.py    → response.s3.content_type, json_formats, csv_formats
│   ├── 00_utils.py       → current.crud_controller = crud_controller
│   ├── 00_tables.py
│   ├── tasks.py
│   └── zz_last.py, zzz_1st_run.py
├── modules/
│   ├── core/         ← S3 framework proper (see table above)
│   ├── s3db/         ← 41 domain DataModel modules
│   ├── templates/    ← deployment skins
│   ├── s3cfg.py      ← S3Config (deployment_settings), ~5800 บรรทัด
│   ├── s3dal.py      ← compat shim (124 บรรทัด)
│   ├── s3log.py
│   ├── s3migration.py
│   └── s3theme.py
├── views/            ← HTML views + DataSeriesView
├── static/           ← JS, CSS, SCSS, images
├── languages/        ← i18n
├── tests/            ← unit tests
├── cron/             ← scheduled tasks
├── private/          ← scripts/ระบบ
├── docs/             ← developer handbook (Sphinx) → readthedocs
├── requirements.txt
└── optional_requirements.txt''',
    language="text")

# ============================================================
# บทที่ 12: Eden ASP Refactor
# ============================================================
add_heading_h1(doc, "การเปลี่ยนแปลง Eden ASP Refactor", num=12)

add_para(doc,
    "ในเดือนพฤษภาคม 2025 Sahana Eden ได้ทำการ refactor ครั้งใหญ่ "
    "โดยย้ายสาขาหลักมาเป็น Eden ASP (Application Service Provider) "
    "ซึ่งเน้น case management สำหรับผู้รับความช่วยเหลือรายบุคคล "
    "และ multi-tenant architecture สาขาเดิมถูกเก็บเป็น Eden Legacy "
    "การ refactor นี้นำมาซึ่งการเปลี่ยนแปลงหลายอย่างที่ developer "
    "ที่คุ้นเคยกับ Eden รุ่นเก่าต้องทำความเข้าใจ",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "12.1  ตารางเปรียบเทียบชื่อเก่า-ใหม่")

add_table(doc,
    headers=["อดีตชื่อ (Legacy)", "ชื่อปัจจุบัน (Eden ASP)", "หมายเหตุ"],
    rows=[
        ["S3Model",            "DataModel",       "Base class สำหรับนิยาม data model"],
        ["S3Resource",         "CRUDResource",    "ห่อหุ้ม Table + filter + components"],
        ["S3CRUD",             "BasicCRUD",       "create/read/update/delete handler"],
        ["S3RESTController",   "CRUDRequest",     "URL parser + handler dispatcher"],
        ["s3xrc",              "(removed)",       "ถูกแทนด้วย prep/postp + set_method + configure"],
        ["S3DAL (as a DAL)",   "S3DAL (compat shim)", "ไม่ใช่ DAL จริง — เป็นเพียง compat adapter"],
        ["modules/s3/*.py",     "modules/core/",   "โครงสร้างโฟลเดอร์ถูก reorganize"],
        ["Eden main branch",   "Eden ASP branch", "เน้น case management + multi-tenancy"],
        ["Eden legacy",        "Eden Legacy (archived)", "เก็บไว้อ้างอิง, ไม่พัฒนาต่อ"],
    ],
    col_widths_cm=[4.0, 4.0, 8.0],
    first_col_bold=True, font_size=9, header_size=9.5)

add_heading_h2(doc, "12.2  การเปลี่ยนแปลงอื่น ๆ ที่สำคัญ")

add_bullet(doc, "ลบ Facebook และ Twitter integration ที่ล้าสมัยออก (controllers และ models)")
add_bullet(doc, "เปลี่ยนชื่อ concept 'obstable' เป็น 'data series' เพื่อความชัดเจน")
add_bullet(doc, "เพิ่ม measurement forms ใน DataSeriesView สำหรับบันทึกค่า measurement")
add_bullet(doc, "ลดความซับซ้อนของ code base เพื่อให้ professional team build/maintain ได้ง่ายขึ้น")
add_bullet(doc, "เปลี่ยน focus จาก self-deployment ไปเป็น foundation สำหรับ professional teams")
add_bullet(doc, "กลายเป็น open project อย่างเต็มตัวในปี 2024 ภายใต้ MIT license")

add_heading_h2(doc, "12.3  สิ่งที่หายไปใน Eden ASP")

add_para(doc,
    "ด้วยการ refactor ครั้งนี้ สิ่งที่ developer รุ่นเก่าอาจคุ้นเคยแต่หายไปใน Eden ASP "
    "ได้แก่ s3xrc object ที่เคยเป็น REST hook registry กลายเป็นสามส่วนแยกกัน "
    "คือ response.s3.prep/postp สำหรับ request-level hooks, "
    "s3db.set_method/get_method สำหรับ custom REST methods "
    "และ s3db.configure สำหรับ onvalidation/onaccept hooks "
    "การแยกนี้ทำให้แต่ละส่วนมีหน้าที่ชัดเจนขึ้น และลด coupling ระหว่าง component",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ============================================================
# บทที่ 13: Design Philosophy
# ============================================================
add_heading_h1(doc, "สรุป Design Philosophy", num=13)

add_para(doc,
    "S3 framework เป็น resource-oriented REST layer ที่สร้างอยู่บน web2py/PyDAL "
    "การออกแบบมีหลักการสำคัญ 6 ข้อที่ทำให้ Eden สามารถรองรับ use case ที่หลากหลาย "
    "ในงานมนุษยธรรมได้ด้วย codebase เดียว หลักการเหล่านี้คือสิ่งที่ developer "
    "ควรเข้าใจและปฏิบัติตามเมื่อ extend Eden หรือ contribute โค้ดกลับเข้าระบบ",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "① Convention over Configuration")

add_para(doc,
    "controller แทบทุกฟังก์ชันใน Eden คือ return crud_controller() "
    "URL routing คือ API ของ Eden ทั้งระบบ ไม่ต้องเขียน route definition "
    "หรือ controller method เอง ทำให้โค้ดสั้น อ่านง่าย และสม่ำเสมอ "
    "การปรับพฤติกรรมทำผ่าน hooks และ configuration ไม่ใช่ผ่านการเขียน controller ใหม่",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "② Lazy, On-demand Model Loading")

add_para(doc,
    "current.s3db.<tablename> auto-instantiate DataModel subclass ที่ถูกต้อง "
    "ทีละตัวเมื่อมีการเรียกใช้ครั้งแรก พร้อม cycle/deadlock detection ผ่าน LOCK/LOAD markers "
    "ทำให้ Eden ไม่ต้องโหลดตารางทั้ง 41 domain ขึ้นมาตั้งแต่ start "
    "ลด memory footprint และ start time ได้มาก โดยเฉพาะใน deployment ที่เปิดเพียงบาง module",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "③ Resource Abstraction")

add_para(doc,
    "CRUDResource ห่อ PyDAL Table + filter (S3ResourceFilter) + components (S3Components) "
    "+ link tables + super-entity relationships เปิดเผย method count/select/load/"
    "insert/update/delete/export_xml/import_xml ที่โปร่งใสในการ apply authorization, "
    "audit, ownership และ onaccept hooks ทำให้ทุก operation ผ่าน Eden มี audit trail "
    "และ security check โดยอัตโนมัติ โดย developer ไม่ต้องจัดการเอง",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "④ Pluggable Method Handlers")

add_para(doc,
    "ทุก URL method คือ CRUDMethod subclass ทั้ง create/read/update/delete พื้นฐาน "
    "และ report/map/profile/hierarchy/organize รวมถึง custom method ที่ลงทะเบียนผ่าน set_method "
    "CRUDRequest หา handler ตามลำดับ get_method → default_methods → BasicCRUD/RESTful "
    "ทำให้การเพิ่ม method ใหม่เป็นเรื่องง่าย เพียง subclass CRUDMethod "
    "และลงทะเบียนในตำแหน่งที่เหมาะสม",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "⑤ Uniform Content Negotiation")

add_para(doc,
    "URL extension กำหนดทั้ง handler branch และ Content-Type header "
    "format writers แยกอยู่ใน modules/core/formats/ โดย XML/S3JSON เป็น canonical format "
    "ที่รองรับ XSLT transform ทำให้เพิ่ม format ใหม่ได้ง่ายเพียงแค่เขียน XSLT stylesheet "
    "โดยไม่ต้องเขียน exporter ใหม่ และ format ทั้งหมด share authorization และ filter logic เดียวกัน",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "⑥ Pervasive Hooks")

add_para(doc,
    "prep/postp ระดับ request, onvalidation/onaccept/ondelete ระดับ record "
    "create_onvalidation/update_onaccept ระดับ method-specific "
    "และ set_method สำหรับ endpoint ใหม่ทั้งหมด เก็บใน registry เดียวกัน (current.model) "
    "ทำให้สามารถปรับพฤติกรรมของระบบได้ทุกระดับ โดยไม่ต้องแก้ไขโค้ดหลัก "
    "และ template หลายตัวสามารถ stack hooks บนตารางเดียวกันได้ผ่าน add_custom_callback",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "แหล่งข้อมูลเพิ่มเติม")

add_table(doc,
    headers=["แหล่งข้อมูล", "URL"],
    rows=[
        ["GitHub Repository",   "https://github.com/sahana/eden"],
        ["Developer Handbook",   "https://eden-asp.readthedocs.io/en/latest/"],
        ["Project Wiki",         "https://eden.sahanafoundation.org"],
        ["Mailing List",         "https://groups.google.com/g/eden-asp"],
        ["Sahana Foundation",    "https://sahanafoundation.org"],
    ],
    col_widths_cm=[5.0, 11.0],
    first_col_bold=True)

add_para(doc,
    "เอกสารนี้อ้างอิงจาก source code จริงใน master branch ของ repository "
    "ณ เดือนพฤศจิกายน 2026 ผู้อ่านควรตรวจสอบ commit ล่าสุดเสมอ "
    "เนื่องจาก Eden เป็นโครงการที่พัฒนาอยู่อย่างต่อเนื่อง และอาจมีการ refactor "
    "หรือเปลี่ยนแปลง API ในอนาคต หากพบข้อขัดแย้งระหว่างเอกสารนี้กับ source code "
    "ให้เชื่อ source code เป็นอันดับแรกเสมอ",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY, italic=True, color=C_SECOND)

# ============================================================
# Footer with page numbers
# ============================================================
section = doc.sections[0]
footer = section.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fp.text = ""
add_run(fp, "Sahana Eden — S3 Framework Architecture  ·  Page ",
        font_name=FONT_BODY, size_pt=9, color=C_SECOND)
# Add PAGE field
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

# ============================================================
# Save
# ============================================================
os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
doc.save(OUT_PATH)
print(f"✓ Saved: {OUT_PATH}")
print(f"  Size: {os.path.getsize(OUT_PATH):,} bytes")
