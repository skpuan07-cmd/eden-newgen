"""
Generate comprehensive .docx covering:
Part A: web2py → FastAPI migration path with code examples
Part B: Python vs Go vs Rust vs TypeScript comparison for every use case
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_PATH = "/home/z/my-project/download/Sahana_Eden_Migration_And_Language_Comparison.docx"

FONT_THAI    = "Loma"
FONT_HEADING = "Loma"
FONT_BODY    = "Loma"
FONT_CODE    = "Liberation Mono"

C_PRIMARY    = RGBColor(0x1E, 0x3A, 0x8A)
C_ACCENT     = RGBColor(0x0E, 0x74, 0x90)
C_DARK       = RGBColor(0x0F, 0x17, 0x2A)
C_SECOND     = RGBColor(0x47, 0x55, 0x69)
C_RED        = RGBColor(0xDC, 0x26, 0x26)
C_GREEN      = RGBColor(0x16, 0xA3, 0x4A)
C_AI         = RGBColor(0x7C, 0x3A, 0xED)
C_AMBER      = RGBColor(0xD9, 0x77, 0x06)

# ============================================================
# Helpers (reused pattern)
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
    set_paragraph_shading(p, "F1F5F9")
    set_paragraph_border(p,
        left={"val": "single", "sz": 24, "color": "1E3A8A", "space": 6},
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

def add_note_box(doc, title, body, color="D97706", bg="FEF3C7"):
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

def add_bullet(doc, text, level=0, size=10.5, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.7 + level*0.6)
    p.paragraph_format.space_after = Pt(3)
    set_paragraph_spacing(p, before_pt=0, after_pt=3)
    if bold_prefix:
        add_run(p, bold_prefix, font_name=FONT_BODY, size_pt=size, bold=True, color=C_PRIMARY)
        add_run(p, text, font_name=FONT_BODY, size_pt=size, color=C_DARK)
    else:
        add_run(p, text, font_name=FONT_BODY, size_pt=size, color=C_DARK)

def add_table(doc, headers, rows, col_widths_cm=None, header_bg="1E3A8A",
              alt_bg="EFF6FF", font_size=9.5, header_size=9.5,
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
add_run(p, "TECHNICAL ANALYSIS DOCUMENT", font_name=FONT_HEADING,
        size_pt=11, bold=True, color=C_ACCENT)

pdiv = doc.add_paragraph()
pdiv.alignment = WD_ALIGN_PARAGRAPH.CENTER
pdiv.paragraph_format.space_after = Pt(30)
set_paragraph_border(pdiv, bottom={"val": "single", "sz": 18, "color": "1E3A8A", "space": 1})

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
add_run(p, "Sahana Eden", font_name=FONT_HEADING, size_pt=38,
        bold=True, color=C_PRIMARY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(8)
add_run(p, "Migration Path และ Language Comparison",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
add_run(p, "Part A: web2py → FastAPI Migration (พร้อม code examples)",
        font_name=FONT_BODY, size_pt=13, italic=True, color=C_AI)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(40)
add_run(p, "Part B: Python vs Go vs Rust vs TypeScript — เปรียบเทียบสำหรับทุก use case",
        font_name=FONT_BODY, size_pt=13, italic=True, color=C_AI)

info = doc.add_table(rows=6, cols=2)
info.alignment = WD_TABLE_ALIGNMENT.CENTER
info.autofit = False
info_data = [
    ("Repository",     "github.com/skpuan07-cmd/eden"),
    ("Base system",    "Sahana Eden v6.2 + AI modules (10 use cases)"),
    ("Stack ปัจจุบัน",  "web2py 3.3.3 (legacy) + Python 3.11+"),
    ("เป้าหมาย Part A", "Migrate web2py → FastAPI + SQLAlchemy 2.0 + Pydantic v2"),
    ("เป้าหมาย Part B", "เลือกภาษาที่เหมาะสมสำหรับแต่ละ use case"),
    ("วันที่",          "19 พฤศจิกายน 2026"),
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
add_run(p, "เอกสารฉบับนี้ประกอบด้วย 2 ส่วนที่เสริมกัน — migration strategy "
           "และการเลือกภาษาตาม use case เพื่อตัดสินใจ architecture ระยะยาว",
        font_name=FONT_BODY, size_pt=10, italic=True, color=C_SECOND)

doc.add_page_break()

# ============================================================
# TOC
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_after = Pt(16)
set_paragraph_border(p, bottom={"val": "single", "sz": 12, "color": "1E3A8A", "space": 4})
add_run(p, "สารบัญ", font_name=FONT_HEADING, size_pt=20, bold=True, color=C_PRIMARY)

toc_items = [
    ("Part A",     "Migration Path: web2py → FastAPI"),
    ("บทที่ 1",    "ทำไมต้อง migrate จาก web2py"),
    ("บทที่ 2",    "Architecture Comparison"),
    ("บทที่ 3",    "Component Mapping (web2py → FastAPI equivalents)"),
    ("บทที่ 4",    "Migration Strategy — Strangler Pattern"),
    ("บทที่ 5",    "Code Examples: Controller Migration"),
    ("บทที่ 6",    "Code Examples: Model & Database Layer"),
    ("บทที่ 7",    "Code Examples: Authentication & Authorization"),
    ("บทที่ 8",    "Code Examples: Templates & Frontend Integration"),
    ("บทที่ 9",    "Coexistence Strategy (รันทั้งสองระบบชั่วคราว)"),
    ("บทที่ 10",   "Migration Phases Roadmap"),
    ("Part B",     "Language Comparison: Python vs Go vs Rust vs TypeScript"),
    ("บทที่ 11",   "ภาพรวมการเปรียบเทียบ 4 ภาษา"),
    ("บทที่ 12",   "Use Case Matrix (14 use cases)"),
    ("บทที่ 13",   "เจาะลึก: AI/ML Stack"),
    ("บทที่ 14",   "เจาะลึก: Web Backend & Microservices"),
    ("บทที่ 15",   "เจาะลึก: Frontend & Mobile"),
    ("บทที่ 16",   "เจาะลึก: Systems Programming & CLI"),
    ("บทที่ 17",   "เจาะลึก: Data Science & Real-time Streaming"),
    ("บทที่ 18",   "สรุปและคำแนะนำสำหรับ Sahana Eden"),
]
for num, title in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(15.5), WD_ALIGN_PARAGRAPH.RIGHT, 2)
    add_run(p, num + "  ", font_name=FONT_HEADING, size_pt=11, bold=True, color=C_PRIMARY)
    add_run(p, title, font_name=FONT_BODY, size_pt=11, color=C_DARK)

doc.add_page_break()

# ============================================================
# PART A — DIVIDER
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(80)
add_run(p, "PART A", font_name=FONT_HEADING, size_pt=42, bold=True, color=C_PRIMARY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(20)
add_run(p, "Migration Path: web2py → FastAPI",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "การวิเคราะห์และย้ายระบบจาก web2py 3.3.3 (legacy) ไป FastAPI + SQLAlchemy 2.0 + Pydantic v2 ทีละส่วน "
           "ด้วย Strangler Pattern เพื่อลด risk และคง business continuity",
        font_name=FONT_BODY, size_pt=11, italic=True, color=C_SECOND)

doc.add_page_break()

# ============================================================
# บทที่ 1: ทำไมต้อง migrate
# ============================================================
add_heading_h1(doc, "ทำไมต้อง migrate จาก web2py", num=1)

add_para(doc,
    "web2py เป็น full-stack web framework ที่พัฒนาโดย Massimo Di Pierro ในปี 2007 "
    "มีจุดเด่นคือง่ายต่อการเริ่มต้นและมี built-in ทุกอย่าง (ORM, auth, admin, scheduler) "
    "แต่หลังจากเวอร์ชั่น 3.3.3 (ต้นปี 2024) framework เข้าสู่ maintenance mode "
    "Massimo หันไปพัฒนา py4web (successor) ส่วน Sahana Eden ที่ใช้ web2py มาตั้งแต่ต้น "
    "ยังคงติดอยู่กับ framework ที่ไม่ได้รับการพัฒนาแบบ active แล้ว",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "1.1  ปัญหาของ web2py ในปัจจุบัน")

add_table(doc,
    headers=["ปัญหา", "รายละเอียด", "ผลกระทบ"],
    rows=[
        ["EOL status",            "เวอร์ชั่น 3.3.3 เป็น stable สุดท้าย (2024), no active development", "ไม่มี new features, security patches ช้า"],
        ["Python compatibility",  "เข้ากับ Python 3.13+ ได้ยาก (free-threaded mode ไม่ support)", "ติดขัด Python รุ่นใหม่"],
        ["Async support",         "ไม่ได้ native — Rocket3 HTTP server เป็น sync",                 "AI streaming (SSE) ทำไม่ได้จริง"],
        ["Documentation",         "เว็บ doc ลดลง, community เล็กลง",                              "Developer ใหม่หาความช่วยเหลือยาก"],
        ["Hiring",                "Niche skill — หา dev ที่รู้ web2py ยาก",                       "เพิ่มความเสี่ยงต่อ bus factor"],
        ["Modern patterns",       "ไม่มี dependency injection, no OpenAPI, no async/await",      "น้อยกว่า frameworks สมัยใหม่มาก"],
        ["Performance",           "Sync + GIL — throughput ต่ำกว่า uvicorn 5-10x",                "Scale ได้น้อยกว่า"],
        ["PyDAL ORM",             "ยัง maintain แต่ legacy patterns",                             "ไม่ทันสมัยเท่า SQLAlchemy 2.0"],
        ["Template engine (YATL)","เก่า ไม่มี IDE support, no syntax highlighting",              "DX ต่ำ"],
        ["Security updates",      "CVE patches ช้า (community-driven)",                           "Risk ด้าน security"],
        ["Type hints integration","รองรับแบบ limited",                                            "IDE/mypy ช่วยได้น้อย"],
        ["Deployment",            "Rocket3 + manual setup, ไม่ native WSGI/ASGI",                 "Docker/K8s deployment ยากกว่า"],
    ],
    col_widths_cm=[3.0, 7.5, 5.5],
    first_col_bold=True, font_size=9, header_size=9.5)

add_heading_h2(doc, "1.2  ทำไมเลือก FastAPI เป็นเป้าหมาย")

add_table(doc,
    headers=["คุณสมบัติ", "FastAPI", "เปรียบเทียบ web2py"],
    rows=[
        ["Performance",            "เร็วที่สุดใน Python web frameworks (Starlette + Pydantic)", "5-10x เร็วกว่า web2py Rocket3"],
        ["Async native",           "async/await throughout, ASGI",                  "Sync only"],
        ["Type hints",             "Required — Pydantic v2 validation",            "Optional"],
        ["OpenAPI/Swagger",        "Auto-generated docs at /docs",                  "ไม่มี"],
        ["Dependency injection",   "First-class — clean architecture",              "Globals (current.db, current.auth)"],
        ["Community",              "ใหญ่และ active (Sebastián Ramírez + Microsoft)", "เล็กและลดลง"],
        ["Ecosystem",              "Modern: SQLAlchemy 2.0, Pydantic v2, uvicorn",  "PyDAL + YATL"],
        ["Learning resources",     "Doc ดีเยี่ยม + คอร์สเยอะ",                       "Niche"],
        ["Hiring",                 "Popular — หา dev ง่าย",                          "Niche"],
        ["Testing",                "pytest + httpx AsyncClient (modern)",           "web2py shell"],
        ["Deployment",             "Docker + uvicorn + gunicorn (standard)",         "Rocket3 + manual"],
        ["WebSocket / SSE",        "Native support",                                 "ไม่ native"],
        ["Security updates",       "Active maintenance, fast CVE response",         "ช้า"],
    ],
    col_widths_cm=[3.0, 6.5, 6.5],
    first_col_bold=True, font_size=9, header_size=9.5)

add_note_box(doc, "สำคัญ: Migrate framework ไม่ใช่ migrate ภาษา",
    "web2py → FastAPI เป็นการเปลี่ยน framework เท่านั้น ยังใช้ Python เหมือนเดิม "
    "นี่คือจุดเด่น — ทีมเดิมที่คุ้น Python สามารถ migrate ได้โดยไม่ต้องเรียนภาษาใหม่ "
    "เฉพาะ framework pattern และ libraries (SQLAlchemy 2.0, Pydantic v2)",
    color="0E7490", bg="ECFDF5")

# ============================================================
# บทที่ 2: Architecture Comparison
# ============================================================
add_heading_h1(doc, "Architecture Comparison", num=2)

add_heading_h2(doc, "2.1  web2py Architecture (ปัจจุบัน)")

add_code_block(doc, '''┌──────────────────────────────────────────────────────────────────┐
│ Browser                                                            │
└──────────────────────────────────────────────────────────────────┘
                              │ HTTP
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Caddy (reverse proxy + HTTPS)                                      │
└──────────────────────────────────────────────────────────────────┘
                              │ HTTP (port 3000)
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Rocket3 (web2py built-in HTTP server) — SYNC                      │
│  • Single-threaded async (limited)                                │
│  • GIL-bound                                                      │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ web2py Application (eden/)                                         │
│                                                                   │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │ Models (run alphabetically on each request!)                 │ │
│  │   000_1st_run.py → 00_db.py → 00_settings.py → ...           │ │
│  │   → 00_utils.py → zz_ai_tables.py → zz_last.py               │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                              │                                    │
│                              ▼                                    │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │ Controllers (controllers/*.py)                               │ │
│  │   def index(): return crud_controller()                      │ │
│  │   — Thin, but coupled to global state                        │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                              │                                    │
│                              ▼                                    │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │ Views (YATL templates)                                       │ │
│  │   {{extend 'layout.html'}}                                   │ │
│  │   {{=form}}                                                  │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                                                                   │
│  Global state: current.db, current.auth, current.response, ...   │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ PyDAL ORM → SQLite / PostgreSQL                                   │
└──────────────────────────────────────────────────────────────────┘''', language="text")

add_heading_h2(doc, "2.2  FastAPI Architecture (เป้าหมาย)")

add_code_block(doc, '''┌──────────────────────────────────────────────────────────────────┐
│ Browser                                                            │
└──────────────────────────────────────────────────────────────────┘
                              │ HTTPS
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Caddy / Traefik (reverse proxy + HTTPS + rate limit)               │
└──────────────────────────────────────────────────────────────────┘
                              │ HTTP (port 8000)
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Uvicorn / Gunicorn (ASGI server) — ASYNC                          │
│  • Multi-process + multi-worker                                    │
│  • asyncio event loop                                             │
│  • Native WebSocket / SSE support                                 │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ FastAPI Application                                                │
│                                                                    │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ Middleware (CORS, Auth, Session, Rate Limit, Logging)       │  │
│  └────────────────────────────────────────────────────────────┘  │
│                              │                                    │
│                              ▼                                    │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ Dependency Injection                                        │  │
│  │   DBSessionDep = Annotated[AsyncSession, Depends(...)]      │  │
│  │   CurrentUserDep = Annotated[User, Depends(get_current)]    │  │
│  └────────────────────────────────────────────────────────────┘  │
│                              │                                    │
│                              ▼                                    │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ Routers (app/api/routers/*.py)                              │  │
│  │   @router.get("/org/{id}")                                  │  │
│  │   async def get_org(id: int, db: DBSessionDep): ...         │  │
│  └────────────────────────────────────────────────────────────┘  │
│                              │                                    │
│                              ▼                                    │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ Schemas (Pydantic v2) + Models (SQLAlchemy 2.0)             │  │
│  │   Separate concerns: validation vs ORM                       │  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                    │
│  OpenAPI auto-docs at /docs (Swagger UI) + /redoc                  │
└──────────────────────────────────────────────────────────────────┘
                              │ asyncpg (async)
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ PostgreSQL 16 + pgvector (with async connection pool)             │
└──────────────────────────────────────────────────────────────────┘''', language="text")

add_heading_h2(doc, "2.3  Key Architectural Differences")

add_table(doc,
    headers=["Aspect", "web2py", "FastAPI"],
    rows=[
        ["Execution model",    "Sync (Rocket3)",                        "Async (ASGI / uvicorn)"],
        ["State",              "Global (current.*)",                    "Dependency injection"],
        ["Models",             "PyDAL Table (declarative)",             "SQLAlchemy 2.0 Mapped classes"],
        ["Validation",         "PyDAL validators (IS_NOT_EMPTY etc.)",  "Pydantic v2 schemas"],
        ["Routing",            "URL → controller function",             "Decorator-based (@router.get)"],
        ["Templates",          "YATL (built-in)",                       "Jinja2 (or HTMX for SPA-like)"],
        ["Auth",               "Built-in Auth (gluon.tools)",           "FastAPI Security + custom deps"],
        ["OpenAPI",            "ไม่มี",                                  "Auto-generated"],
        ["Migration",          "auto-migrate (โหมด dev)",               "Alembic (explicit, reproducible)"],
        ["Sessions",           "File/DB-based sessions",                "Starlette SessionMiddleware"],
        ["Background jobs",    "Scheduler (built-in)",                  "Celery / RQ / arq"],
        ["Testing",            "web2py shell + doctests",               "pytest + httpx.AsyncClient"],
        ["Deployment",         "Rocket3 (single process)",              "uvicorn (multi-worker)"],
        ["Hot reload",         "Built-in (dev mode)",                   "uvicorn --reload"],
    ],
    col_widths_cm=[3.0, 5.5, 7.5],
    first_col_bold=True)

# ============================================================
# บทที่ 3: Component Mapping
# ============================================================
add_heading_h1(doc, "Component Mapping (web2py → FastAPI equivalents)", num=3)

add_heading_h2(doc, "3.1  Project Structure Comparison")

add_code_block(doc, '''# ─── web2py (ปัจจุบัน) ────────────────────────────────
eden/
├── controllers/
│   ├── org.py         (def organisation(): return crud_controller())
│   ├── pr.py
│   ├── inv.py
│   ├── ai_chat.py
│   └── ...
├── models/
│   ├── 000_1st_run.py     ← run on EVERY request (web2py pattern)
│   ├── 00_db.py           ← db = DAL(...)
│   ├── 00_settings.py     ← global config
│   ├── 00_utils.py        ← helper globals
│   ├── zz_ai_tables.py    ← AI tables defined here
│   └── zz_last.py
├── views/
│   ├── default/index.html
│   ├── org/organisation.html
│   └── layout.html        ← YATL master template
├── modules/
│   ├── s3cfg.py
│   ├── s3db/              ← DataModel subclasses (41 files)
│   └── s3/
│       ├── ai_client.py
│       └── ai_ocr_openthai.py
├── static/                ← JS/CSS/images
├── languages/             ← i18n (.py files)
├── private/
├── tests/
└── databases/             ← SQLite file

# ─── FastAPI (target) ──────────────────────────────────────
eden_fastapi/
├── app/
│   ├── __init__.py
│   ├── main.py             ← FastAPI() instance, lifespan, routers
│   ├── config.py           ← Pydantic Settings (env vars)
│   │
│   ├── api/
│   │   ├── __init__.py
│   │   ├── dependencies/
│   │   │   ├── __init__.py
│   │   │   ├── core.py     ← DBSessionDep, PaginatorDep
│   │   │   ├── auth.py     ← CurrentUserDep, AdminUserDep
│   │   │   └── pagination.py
│   │   ├── routers/
│   │   │   ├── __init__.py
│   │   │   ├── org.py      ← @router.get("/org/{id}")
│   │   │   ├── pr.py
│   │   │   ├── inv.py
│   │   │   ├── ai_chat.py
│   │   │   ├── ai_ocr.py
│   │   │   └── ...
│   │   └── middleware/
│   │       ├── auth.py     ← AuthMiddleware
│   │       ├── rate_limit.py
│   │       └── logging.py
│   │
│   ├── models/             ← SQLAlchemy 2.0 (DB schema)
│   │   ├── __init__.py     ← Base = declarative_base()
│   │   ├── base.py         ← TimestampMixin, UUIDMixin
│   │   ├── org.py          ← class Organisation(Base): ...
│   │   ├── pr.py
│   │   ├── inv.py
│   │   ├── auth.py
│   │   └── ai.py           ← AI tables (ai_chat_log etc.)
│   │
│   ├── schemas/            ← Pydantic v2 (validation)
│   │   ├── org.py          ← OrganisationCreate, OrganisationRead, ...
│   │   ├── pr.py
│   │   ├── auth.py         ← UserRead, Token
│   │   └── ai.py
│   │
│   ├── crud/               ← Data access layer
│   │   ├── base.py         ← CRUDBase[ModelType, SchemaType]
│   │   ├── org.py          ← CRUDOrganisation
│   │   ├── pr.py
│   │   └── ai.py
│   │
│   ├── core/
│   │   ├── database.py     ← DatabaseSessionManager (async)
│   │   ├── security.py     ← JWT, password hashing
│   │   ├── config.py       ← pydantic-settings
│   │   └── exceptions.py
│   │
│   ├── services/           ← Business logic
│   │   ├── ai_client.py    ← AsyncAIClient (aiohttp)
│   │   ├── ai_rag.py       ← VectorStore (pgvector)
│   │   ├── ai_ocr.py       ← OpenThaiClient
│   │   └── ...
│   │
│   └── templates/          ← Jinja2 (if server-rendered)
│       ├── base.html
│       ├── org/
│       └── default/
│
├── alembic/                ← Database migrations
│   ├── env.py
│   ├── versions/
│   └── alembic.ini
│
├── tests/
│   ├── conftest.py
│   ├── unit/
│   └── integration/
│
├── static/                 ← JS/CSS (Tailwind + HTMX)
├── pyproject.toml          ← ruff, black, mypy, pytest config
├── Dockerfile
├── docker-compose.yml
└── requirements.txt''', language="text")

add_heading_h2(doc, "3.2  Component Mapping Table")

add_table(doc,
    headers=["Concept", "web2py", "FastAPI equivalent"],
    rows=[
        ["Application entry",      "web2py.py + applications/eden/",              "app/main.py: app = FastAPI()"],
        ["HTTP server",            "Rocket3 (built-in)",                           "uvicorn / gunicorn"],
        ["URL routing",            "controllers/<module>.py → /eden/<module>",     "@router.get('/path')"],
        ["Controller function",    "def organisation(): return crud_controller()", "async def list_orgs(db: DBSessionDep): ..."],
        ["DB connection",          "current.db (global)",                          "AsyncSession via Depends(get_db_session)"],
        ["ORM",                    "PyDAL Table + DataModel",                       "SQLAlchemy 2.0 Mapped classes"],
        ["Validation",             "Field requires (IS_NOT_EMPTY etc.)",           "Pydantic v2 schemas"],
        ["Auth",                   "current.auth (gluon.tools.Auth)",              "FastAPI Security + JWT + custom deps"],
        ["Sessions",               "session (Storage object)",                     "SessionMiddleware (signed cookies)"],
        ["Templates",              "YATL ({{=var}})",                              "Jinja2 ({{ var }}) or HTMX"],
        ["i18n",                   "languages/<lang>.py (dict)",                    "Babel + gettext (.po files)"],
        ["Static files",           "static/ (served by web2py)",                    "StaticFiles mount"],
        ["Config",                 "models/000_config.py + modules/s3cfg.py",       "pydantic-settings (env-driven)"],
        ["Migrations",             "migrate=True (auto on startup)",                "Alembic (explicit)"],
        ["Background jobs",        "scheduler (built-in)",                          "Celery + Redis"],
        ["Email",                  "gluon.tools.Mail",                              "aiosmtplib / FastMail"],
        ["Cache",                  "cache.ram / cache.disk",                        "Redis (aiocache)"],
        ["OpenAPI docs",           "ไม่มี",                                          "Auto at /docs"],
        ["Error handling",         "tickets ใน errors/ folder",                     "@app.exception_handler + Sentry"],
        ["Logging",                "logging module (basic)",                        "structlog + OpenTelemetry"],
        ["Testing",                "doctests + manual",                             "pytest + httpx.AsyncClient"],
        ["CLI",                    "python web2py.py -S eden -M -R ...",             "typer / click commands"],
        ["Hot reload",             "--reload flag",                                 "uvicorn --reload"],
        ["Cron jobs",              "cron/crontab",                                  "Celery Beat / APScheduler"],
    ],
    col_widths_cm=[3.0, 5.0, 8.0],
    first_col_bold=True, font_size=9, header_size=9.5)

# ============================================================
# บทที่ 4: Strangler Pattern
# ============================================================
add_heading_h1(doc, "Migration Strategy — Strangler Pattern", num=4)

add_para(doc,
    "Strangler Fig Pattern (Martin Fowler) เป็นวิธี migration ที่ค่อย ๆ แทนที่ระบบเดิม "
    "ทีละส่วน โดยใช้ reverse proxy (Caddy/Nginx) เป็น facade ที่ route request "
    "ไปยัง web2py (เดิม) หรือ FastAPI (ใหม่) ตาม path — ทำให้สามารถ migrate ได้ "
    "โดยไม่ต้อง rewrite ทั้งหมดในครั้งเดียว ลด risk และคง business continuity",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "4.1  Strangler Pattern Architecture")

add_code_block(doc, '''┌─────────────────────────────────────────────────────────────┐
│ Browser                                                      │
└─────────────────────────────────────────────────────────────┘
                            │ HTTPS
                            ▼
┌─────────────────────────────────────────────────────────────┐
│ Caddy (reverse proxy + facade)                               │
│                                                              │
│  Routes by path prefix:                                      │
│    /eden/api/v2/*         → FastAPI (port 8001)  [NEW]       │
│    /eden/ai_chat/*        → FastAPI (port 8001)  [NEW]       │
│    /eden/ai_ocr/*         → FastAPI (port 8001)  [NEW]       │
│    /eden/ai_query/*       → FastAPI (port 8001)  [NEW]       │
│    /eden/ai_kb/*          → FastAPI (port 8001)  [NEW]       │
│    /eden/* (อื่นๆ)         → web2py (port 8000)   [LEGACY]    │
│                                                              │
│  → ค่อยๆ ขยายส่วนของ FastAPI จน web2py เหลือน้อยจนลบได้       │
└─────────────────────────────────────────────────────────────┘
              │                              │
              ▼                              ▼
   ┌────────────────────┐         ┌────────────────────┐
   │ FastAPI (port 8001)│         │ web2py (port 8000) │
   │  [NEW system]      │         │  [LEGACY]          │
   │                    │         │                    │
   │  • AI modules      │         │  • org/pr/inv/hrm  │
   │  • api/v2/*        │         │  • project/req/... │
   │  • async endpoints │         │  • S3 CRUD         │
   └────────────────────┘         └────────────────────┘
              │                              │
              └──────────┬───────────────────┘
                         ▼
              ┌────────────────────┐
              │ PostgreSQL 16      │
              │ (shared DB)        │
              └────────────────────┘
                         │
                         ▼
              ┌────────────────────┐
              │ Redis (cache/jobs) │
              └────────────────────┘''', language="text")

add_heading_h2(doc, "4.2  Caddyfile สำหรับ Strangler Routing")

add_code_block(doc, '''# Caddyfile — Strangler Pattern
{$EDEN_DOMAIN} {
    # 1. NEW: FastAPI endpoints (migrated)
    @fastapi {
        path /eden/api/v2/*
        path /eden/ai_chat/*
        path /eden/ai_ocr/*
        path /eden/ai_query/*
        path /eden/ai_kb/*
        path /eden/ai_voice/*
        path /eden/ai_analytics/*
        path /eden/ai_imagery/*
        path /eden/ai_translate/*
        path /eden/ai_moderation/*
        path /eden/health
        path /eden/docs/*
        path /eden/openapi.json
    }
    reverse_proxy @fastapi fastapi:8001 {
        header_up X-Real-IP {remote_host}
        header_up X-Forwarded-Proto {scheme}
    }
    
    # 2. LEGACY: web2py (everything else)
    reverse_proxy web2py:8000 {
        header_up X-Real-IP {remote_host}
        header_up X-Forwarded-Proto {scheme}
    }
    
    # Security headers (applies to both)
    header {
        Content-Security-Policy "..."
        Strict-Transport-Security "max-age=31536000; includeSubDomains"
        X-Frame-Options "DENY"
    }
}''', language="text")

add_heading_h2(doc, "4.3  Migration Order (Priority)")

add_table(doc,
    headers=["Phase", "Migrate อะไร", "เหตุผล", "Risk"],
    rows=[
        ["1",  "AI modules (10 controllers)",      "ใหม่ทั้งหมด — ง่ายสุด ไม่มี legacy coupling",           "ต่ำ"],
        ["2",  "Health endpoint + metrics",         "Standalone — ไม่กระทบ business logic",                  "ต่ำ"],
        ["3",  "Auth + User management (/eden/default/user/*)", "Foundation สำหรับ endpoints อื่น",        "กลาง"],
        ["4",  "Read-only API (/eden/api/v2/*)",    "เริ่ม expose REST API สำหรับ external consumers",        "กลาง"],
        ["5",  "CMS module (/eden/cms/*)",          "Self-contained — ไม่พึ่ง controllers อื่น",               "กลาง"],
        ["6",  "ORG module (/eden/org/*)",          "Foundation สำหรับ inv/asset/req ที่อ้าง org_organisation","กลาง-สูง"],
        ["7",  "PR module (/eden/pr/*)",            "Sensitive (PDPA) — ต้องระวัง consent + encryption",      "สูง"],
        ["8",  "HRM module (/eden/hrm/*)",          "ข้อมูลบุคลากร — sensitive",                              "สูง"],
        ["9",  "INV module (/eden/inv/*)",          "Complex — warehouse + items + shipments",                "สูง"],
        ["10", "REQ, ASSET, VEHICLE",                "Module ย่อย — ทำหลัง foundation เสร็จ",                  "กลาง"],
        ["11", "PROJECT, STATS",                    "Module ย่อย",                                            "กลาง"],
        ["12", "GIS + CAP",                         "GIS complex — อาจเก็บ web2py ไว้นาน",                     "สูง"],
        ["13", "DVR (ผู้ประสบภัย — sensitive)",     "PDPA-critical — ต้อง audit + test อย่างละเอียด",         "สูงมาก"],
        ["14", "Sync, Admin, Appadmin",             "Internal tools — migrate ท้ายสุด",                       "กลาง"],
        ["15", "Decommission web2py",                "Shutdown web2py process — migration เสร็จ",              "—"],
    ],
    col_widths_cm=[1.0, 4.0, 7.5, 2.0],
    first_col_bold=True, font_size=9, header_size=9.5)

# ============================================================
# บทที่ 5: Controller Migration Examples
# ============================================================
add_heading_h1(doc, "Code Examples: Controller Migration", num=5)

add_heading_h2(doc, "5.1  Example: ai_chat controller (web2py → FastAPI)")

add_heading_h3(doc, "BEFORE — web2py (controllers/ai_chat.py)")

add_code_block(doc, '''# controllers/ai_chat.py (web2py — current)
"""
AI Chatbot controller — Use Case 1: AI Chatbot for citizens and volunteers
"""
import json
import time

def index():
    """Main chat UI."""
    if request.vars.consent == "yes":
        session.ai_chat_consent = True
    consent_given = session.ai_chat_consent or False
    
    response.title = T("AI Assistant")
    return dict(
        consent_given=consent_given,
        chat_endpoint=URL("ai_chat", "ask"),
        rate_endpoint=URL("ai_chat", "rate"),
    )

def ask():
    """POST endpoint: receive a question, return an answer (JSON)."""
    from s3.ai_client import AIClient, AIClientError
    
    if request.env.request_method != "POST":
        response.view = "generic.json"
        return dict(error="POST required")
    
    question = (request.post_vars.question or "").strip()
    if not question:
        response.headers["Content-Type"] = "application/json"
        return json.dumps({"error": "missing question"})
    
    if len(question) > 1000:
        response.headers["Content-Type"] = "application/json"
        return json.dumps({"error": "question too long (max 1000 chars)"})
    
    # Rate limit: simple per-IP throttle (60/min)
    client_ip = request.env.remote_addr or "unknown"
    one_min_ago = request.now - datetime.timedelta(seconds=60)
    recent_count = db(
        (db.ai_chat_log.ip == client_ip) &
        (db.ai_chat_log.created_on > one_min_ago)
    ).count()
    if recent_count > 60:
        response.headers["Content-Type"] = "application/json"
        return json.dumps({"error": "rate limit exceeded (60/min)"})
    
    # PDPA consent check
    consent_given = bool(session.ai_chat_consent)
    if not consent_given:
        response.headers["Content-Type"] = "application/json"
        return json.dumps({
            "error": "consent_required",
            "message": "PDPA consent required",
        })
    
    # RAG: retrieve relevant knowledge chunks
    t0 = time.time()
    chunks = _retrieve_relevant_chunks(question, top_k=5)
    
    # Call LLM
    try:
        client = AIClient(timeout=120)
        context = "\\n\\n".join([c.content for c in chunks])
        answer = client.chat(
            prompt=question,
            system=f"You are a Sahana Eden assistant. Use this context:\\n{context}",
        )
    except AIClientError as e:
        response.headers["Content-Type"] = "application/json"
        return json.dumps({"error": str(e)})
    
    latency_ms = int((time.time() - t0) * 1000)
    
    # Log to DB
    db.ai_chat_log.insert(
        session_id=session.id,
        user_id=session.auth_user.id if session.auth else None,
        question=question,
        answer=answer,
        context_used=json.dumps([{"id": c.id, "title": c.title} for c in chunks]),
        ip=client_ip,
        model="glm-4-plus",
        latency_ms=latency_ms,
        consent_given=True,
    )
    db.commit()
    
    response.headers["Content-Type"] = "application/json"
    return json.dumps({
        "answer": answer,
        "sources": [{"id": c.id, "title": c.title} for c in chunks],
        "latency_ms": latency_ms,
    })

def history():
    """User's conversation history."""
    if not session.auth:
        return dict(error="login required")
    logs = db(db.ai_chat_log.user_id == session.auth_user.id).select(
        orderby=~db.ai_chat_log.created_on,
        limitby=(0, 50),
    )
    return dict(logs=logs.as_list())''', language="python")

add_heading_h3(doc, "AFTER — FastAPI (app/api/routers/ai_chat.py)")

add_code_block(doc, '''# app/api/routers/ai_chat.py (FastAPI — target)
"""
AI Chatbot router — Use Case 1: AI Chatbot for citizens and volunteers

Endpoints:
- GET  /eden/ai_chat/        → Chat UI (HTML)
- POST /eden/ai_chat/ask     → Submit question, get answer (JSON)
- GET  /eden/ai_chat/history → User's conversation history
- POST /eden/ai_chat/rate    → Rate answer (+1/-1)
- GET  /eden/ai_chat/stream  → SSE streaming endpoint
"""
from fastapi import APIRouter, Depends, HTTPException, Request, Response
from fastapi.responses import JSONResponse, StreamingResponse, HTMLResponse
from sqlalchemy import select, desc
from sqlalchemy.ext.asyncio import AsyncSession
from typing import Annotated, Optional
import json
import time
import asyncio

from app.api.dependencies.core import DBSessionDep, RateLimiterDep
from app.api.dependencies.auth import CurrentUserOptional
from app.core.config import settings
from app.models.ai import AIChatLog
from app.schemas.ai import (
    ChatAskRequest,
    ChatAskResponse,
    ChatHistoryItem,
    RateRequest,
)
from app.services.ai_client import AsyncAIClient
from app.services.ai_rag import VectorStore

router = APIRouter(prefix="/eden/ai_chat", tags=["AI Chat"])


# ─── Chat UI (HTML page) ────────────────────────────────────────────────
@router.get("/", response_class=HTMLResponse)
async def chat_ui(request: Request):
    """Render chat UI (Jinja2 template)."""
    consent_given = request.session.get("ai_chat_consent", False)
    return request.app.state.templates.TemplateResponse(
        "ai_chat/index.html",
        {"request": request, "consent_given": consent_given},
    )


# ─── Submit question ─────────────────────────────────────────────────────
@router.post("/ask", response_model=ChatAskResponse)
async def ask(
    body: ChatAskRequest,                    # ← Pydantic validation
    request: Request,
    db: DBSessionDep,                         # ← Injected DB session
    rate_limiter: RateLimiterDep,             # ← Rate limit dependency
    current_user: CurrentUserOptional,        # ← Optional auth (None if anon)
):
    """
    Submit a question, get an answer (JSON).
    
    - **question**: max 1000 chars
    - Requires PDPA consent (recorded in session)
    - Rate limited: 60 req/min/IP
    """
    # Pydantic ทำ validation แล้ว — ไม่ต้องเช็ค length เอง
    
    # Rate limit check
    client_ip = request.client.host
    await rate_limiter.check_or_raise(
        key=f"ai_chat:{client_ip}",
        limit=60,
        window=60,
    )
    
    # PDPA consent check
    consent_given = request.session.get("ai_chat_consent", False)
    if not consent_given:
        raise HTTPException(
            status_code=403,
            detail={
                "error": "consent_required",
                "message": "PDPA consent required",
                "consent_url": "/eden/ai_chat/?show_consent=1",
            },
        )
    
    # RAG: retrieve relevant chunks
    t0 = time.time()
    vector_store = VectorStore()
    await vector_store.connect()
    chunks = await vector_store.search(
        query_embedding=await _embed_question(body.question),
        top_k=5,
    )
    
    # Async LLM call
    try:
        client = AsyncAIClient(
            api_key=settings.zai_api_key,
            timeout=120,
        )
        context = "\\n\\n".join([c["content"] for c in chunks])
        answer = await client.chat(
            prompt=body.question,
            system=f"You are a Sahana Eden assistant. Context:\\n{context}",
        )
    except Exception as e:
        raise HTTPException(
            status_code=502,
            detail=f"AI service unavailable: {e}",
        )
    
    latency_ms = int((time.time() - t0) * 1000)
    
    # Log to DB (async)
    log_entry = AIChatLog(
        session_id=request.session.get("session_id"),
        user_id=current_user.id if current_user else None,
        question=body.question,
        answer=answer,
        context_used=json.dumps([
            {"id": c["id"], "title": c["title"]} for c in chunks
        ]),
        ip=client_ip,
        model="glm-4-plus",
        latency_ms=latency_ms,
        consent_given=True,
    )
    db.add(log_entry)
    await db.commit()
    await db.refresh(log_entry)
    
    return ChatAskResponse(
        answer=answer,
        sources=[{"id": c["id"], "title": c["title"]} for c in chunks],
        latency_ms=latency_ms,
        log_id=log_entry.id,
    )


# ─── Streaming chat (SSE) ────────────────────────────────────────────────
@router.post("/stream")
async def stream(
    body: ChatAskRequest,
    request: Request,
    db: DBSessionDep,
    current_user: CurrentUserOptional,
):
    """Stream chat tokens via Server-Sent Events."""
    consent_given = request.session.get("ai_chat_consent", False)
    if not consent_given:
        raise HTTPException(status_code=403, detail="consent_required")
    
    async def event_generator():
        client = AsyncAIClient(api_key=settings.zai_api_key)
        try:
            async for token in client.chat_stream(body.question):
                yield f"data: {json.dumps({'token': token})}\\n\\n"
            yield f"data: {json.dumps({'done': True})}\\n\\n"
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)})}\\n\\n"
    
    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",  # disable nginx buffering
        },
    )


# ─── Conversation history ────────────────────────────────────────────────
@router.get("/history", response_model=list[ChatHistoryItem])
async def history(
    db: DBSessionDep,
    current_user: CurrentUserOptional = None,    # ← required auth
    limit: int = 50,
):
    """Get user's conversation history (requires login)."""
    if not current_user:
        raise HTTPException(status_code=401, detail="login required")
    
    stmt = (
        select(AIChatLog)
        .where(AIChatLog.user_id == current_user.id)
        .order_by(desc(AIChatLog.created_on))
        .limit(limit)
    )
    result = await db.execute(stmt)
    logs = result.scalars().all()
    return [ChatHistoryItem.model_validate(log) for log in logs]


# ─── Rate answer ─────────────────────────────────────────────────────────
@router.post("/rate/{log_id}")
async def rate(
    log_id: int,
    body: RateRequest,
    db: DBSessionDep,
    current_user: CurrentUserOptional = None,
):
    """Rate an answer (+1 or -1)."""
    if not current_user:
        raise HTTPException(status_code=401, detail="login required")
    
    stmt = select(AIChatLog).where(
        AIChatLog.id == log_id,
        AIChatLog.user_id == current_user.id,
    )
    result = await db.execute(stmt)
    log = result.scalar_one_or_none()
    if not log:
        raise HTTPException(status_code=404, detail="log not found")
    
    if body.rating not in (-1, 0, 1):
        raise HTTPException(status_code=422, detail="rating must be -1, 0, or 1")
    
    log.rating = body.rating
    if body.comment:
        log.feedback_comment = body.comment
    await db.commit()
    
    return {"status": "ok"}


async def _embed_question(question: str) -> list[float]:
    """Generate embedding for question."""
    from app.services.ai_client import AsyncAIClient
    client = AsyncAIClient(api_key=settings.zai_api_key)
    return await client.embed(question)''', language="python")

add_heading_h3(doc, "Pydantic Schema (app/schemas/ai.py)")

add_code_block(doc, '''# app/schemas/ai.py — Pydantic v2 schemas (validation)

from pydantic import BaseModel, Field, ConfigDict
from datetime import datetime
from typing import Optional


class ChatAskRequest(BaseModel):
    """Request body for POST /eden/ai_chat/ask."""
    question: str = Field(
        ...,
        min_length=1,
        max_length=1000,
        description="User's question (max 1000 chars)",
        examples=["ศูนย์พักพิงใกล้ฉันอยู่ที่ไหน?"],
    )
    language: Optional[str] = Field(
        default=None,
        max_length=8,
        description="Detected language (auto if None)",
    )


class ChatAskResponse(BaseModel):
    """Response for POST /eden/ai_chat/ask."""
    answer: str = Field(..., description="AI-generated answer")
    sources: list[dict] = Field(
        default_factory=list,
        description="Source chunks used in RAG",
    )
    latency_ms: int = Field(..., description="Response time in ms")
    log_id: int = Field(..., description="ID in ai_chat_log for rate()")


class ChatHistoryItem(BaseModel):
    """Single chat history entry."""
    model_config = ConfigDict(from_attributes=True)  # ← from ORM object
    
    id: int
    question: str
    answer: str
    created_on: datetime
    rating: Optional[int] = None


class RateRequest(BaseModel):
    """Request body for POST /eden/ai_chat/rate/{log_id}."""
    rating: int = Field(..., ge=-1, le=1, description="-1, 0, or +1")
    comment: Optional[str] = Field(default=None, max_length=500)''', language="python")

add_heading_h3(doc, "สิ่งที่ดีขึ้นใน FastAPI version")

add_bullet(doc, "Type-safe — request/response validated อัตโนมัติด้วย Pydantic", bold_prefix="✓ ")
add_bullet(doc, "Async — DB queries + LLM calls ไม่ block thread", bold_prefix="✓ ")
add_bullet(doc, "Dependency injection — DBSessionDep, CurrentUserDep (no global state)", bold_prefix="✓ ")
add_bullet(doc, "OpenAPI auto-docs — /eden/docs แสดง schema + try-it-out", bold_prefix="✓ ")
add_bullet(doc, "Native SSE streaming — StreamingResponse + async generator", bold_prefix="✓ ")
add_bullet(doc, "Rate limit as dependency — reusable, testable", bold_prefix="✓ ")
add_bullet(doc, "Error handling — HTTPException with structured detail", bold_prefix="✓ ")
add_bullet(doc, "Testable — pytest + httpx.AsyncClient", bold_prefix="✓ ")

# ============================================================
# บทที่ 6: Model & Database Layer
# ============================================================
add_heading_h1(doc, "Code Examples: Model & Database Layer", num=6)

add_heading_h2(doc, "6.1  BEFORE — web2py DataModel (modules/s3db/org.py)")

add_code_block(doc, '''# modules/s3db/org.py (web2py — current)

class OrgOrganisationModel(DataModel):
    """Organisations data model."""

    names = ("org_organisation_type",
             "org_organisation_type_id",
             "org_organisation",
             "org_organisation_id",
             "org_organisation_represent",)

    def model(self):
        T = current.T
        db = current.db
        define_table = self.define_table

        # Organisation Type
        tablename = "org_organisation_type"
        define_table(tablename,
            Field("name", length=128, notnull=True, unique=True,
                  label=T("Name"),
                  requires=[IS_NOT_EMPTY(),
                            IS_LENGTH(128),
                            IS_NOT_IN_DB(db, "org_organisation_type.name")]),
            Field("parent", "reference org_organisation_type",
                  ondelete="RESTRICT",
                  requires=IS_EMPTY_OR(
                      IS_IN_DB(db, "%s.id" % tablename))),
            CommentsField(),
        )

        # Organisation
        tablename = "org_organisation"
        define_table(tablename,
            Field("name", length=128, notnull=True, unique=True,
                  label=T("Name"),
                  requires=[IS_NOT_EMPTY(),
                            IS_LENGTH(128)]),
            Field("acronym", length=8),
            Field("type", "reference org_organisation_type"),
            Field("website", requires=IS_EMPTY_OR(IS_URL())),
            Field("email", requires=IS_EMPTY_OR(IS_EMAIL())),
            Field("phone", length=32),
            Field("country", "reference gis_location"),
            Field("deleted", "boolean", default=False),
            CommentsField(),
        )''', language="python")

add_heading_h2(doc, "6.2  AFTER — SQLAlchemy 2.0 (app/models/org.py)")

add_code_block(doc, '''# app/models/org.py (SQLAlchemy 2.0 — target)

from sqlalchemy import String, Integer, Boolean, ForeignKey, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship
from typing import Optional
from datetime import datetime

from app.models.base import Base, TimestampMixin, UUIDMixin, SoftDeleteMixin


class OrganisationType(Base, UUIDMixin, TimestampMixin, SoftDeleteMixin):
    """Type of organisation (e.g., NGO, Government, Private)."""
    __tablename__ = "org_organisation_type"
    
    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), unique=True, nullable=False, index=True)
    parent_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_organisation_type.id", ondelete="RESTRICT"),
        nullable=True,
    )
    
    # Self-referential relationship
    parent: Mapped[Optional["OrganisationType"]] = relationship(
        remote_side="OrganisationType.id",
        back_populates="children",
    )
    children: Mapped[list["OrganisationType"]] = relationship(
        back_populates="parent",
    )
    
    # Reverse relationship
    organisations: Mapped[list["Organisation"]] = relationship(
        back_populates="type",
    )


class Organisation(Base, UUIDMixin, TimestampMixin, SoftDeleteMixin):
    """Organisation registered in Eden."""
    __tablename__ = "org_organisation"
    
    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, index=True)
    acronym: Mapped[Optional[str]] = mapped_column(String(8), nullable=True)
    type_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("org_organisation_type.id"),
        nullable=True,
        index=True,
    )
    website: Mapped[Optional[str]] = mapped_column(String(512), nullable=True)
    email: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    phone: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    country_id: Mapped[Optional[int]] = mapped_column(
        ForeignKey("gis_location.id"),
        nullable=True,
    )
    
    # Relationships
    type: Mapped[Optional["OrganisationType"]] = relationship(
        back_populates="organisations",
    )
    offices: Mapped[list["Office"]] = relationship(
        back_populates="organisation",
        cascade="all, delete-orphan",
    )
    staff: Mapped[list["HumanResource"]] = relationship(
        back_populates="organisation",
    )


# app/models/base.py — mixins
from sqlalchemy import DateTime, Boolean, String
from sqlalchemy.orm import Mapped, mapped_column
from datetime import datetime
from sqlalchemy.orm import DeclarativeBase


class Base(DeclarativeBase):
    """SQLAlchemy 2.0 declarative base."""
    pass


class UUIDMixin:
    """Adds uuid field for sync (compatible with Eden's MetaFields)."""
    uuid: Mapped[str] = mapped_column(String(128), unique=True, index=True)


class TimestampMixin:
    """Adds created_on and modified_on (with auto-update)."""
    created_on: Mapped[datetime] = mapped_column(
        DateTime,
        default=datetime.utcnow,
    )
    modified_on: Mapped[datetime] = mapped_column(
        DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
    )


class SoftDeleteMixin:
    """Adds deleted flag (Eden's soft-delete pattern)."""
    deleted: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
        nullable=False,
        index=True,
    )''', language="python")

add_heading_h2(doc, "6.3  Database Session (app/core/database.py)")

add_code_block(doc, '''# app/core/database.py — Async SQLAlchemy session manager

import contextlib
from typing import Any, AsyncIterator
from sqlalchemy.ext.asyncio import (
    AsyncConnection,
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)
from sqlalchemy.orm import declarative_base

from app.core.config import settings

Base = declarative_base()


class DatabaseSessionManager:
    """Manages async engine + session factory."""
    
    def __init__(self, host: str, engine_kwargs: dict[str, Any] = None):
        self._engine = create_async_engine(host, **(engine_kwargs or {}))
        self._sessionmaker = async_sessionmaker(
            autocommit=False,
            bind=self._engine,
            expire_on_commit=False,
        )
    
    async def close(self):
        if self._engine is None:
            raise Exception("DatabaseSessionManager is not initialized")
        await self._engine.dispose()
        self._engine = None
        self._sessionmaker = None
    
    @contextlib.asynccontextmanager
    async def session(self) -> AsyncIterator[AsyncSession]:
        """Get a session — auto-rollback on exception."""
        if self._sessionmaker is None:
            raise Exception("DatabaseSessionManager is not initialized")
        async_session = self._sessionmaker()
        try:
            yield async_session
        except Exception:
            await async_session.rollback()
            raise
        finally:
            await async_session.close()


# Singleton instance
sessionmanager = DatabaseSessionManager(
    host=settings.database_url,
    engine_kwargs={
        "echo": settings.echo_sql,
        "pool_size": 10,
        "max_overflow": 20,
        "pool_pre_ping": True,  # detect stale connections
    },
)


async def get_db_session() -> AsyncIterator[AsyncSession]:
    """FastAPI dependency: provides an async DB session."""
    async with sessionmanager.session() as session:
        yield session''', language="python")

add_heading_h2(doc, "6.4  Dependency Injection (app/api/dependencies/core.py)")

add_code_block(doc, '''# app/api/dependencies/core.py — Reusable dependencies

from typing import Annotated
from fastapi import Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db_session

# Type alias for use in route signatures
DBSessionDep = Annotated[AsyncSession, Depends(get_db_session)]

# Usage in routers:
#   async def get_org(org_id: int, db: DBSessionDep):
#       ...''', language="python")

add_heading_h2(doc, "6.5  CRUD Layer (app/crud/org.py)")

add_code_block(doc, '''# app/crud/org.py — Data access layer (async)

from sqlalchemy import select, func
from sqlalchemy.ext.asyncio import AsyncSession
from typing import Optional, Sequence

from app.models.org import Organisation, OrganisationType
from app.schemas.org import OrganisationCreate, OrganisationUpdate


async def get_organisation(
    db: AsyncSession,
    org_id: int,
) -> Optional[Organisation]:
    """Get a single organisation by ID."""
    stmt = select(Organisation).where(
        Organisation.id == org_id,
        Organisation.deleted == False,
    )
    result = await db.execute(stmt)
    return result.scalar_one_or_none()


async def get_organisations(
    db: AsyncSession,
    *,
    skip: int = 0,
    limit: int = 25,
    type_id: Optional[int] = None,
    search: Optional[str] = None,
) -> Sequence[Organisation]:
    """Get paginated list of organisations with optional filters."""
    stmt = select(Organisation).where(Organisation.deleted == False)
    
    if type_id:
        stmt = stmt.where(Organisation.type_id == type_id)
    
    if search:
        stmt = stmt.where(
            Organisation.name.ilike(f"%{search}%")
        )
    
    stmt = stmt.offset(skip).limit(limit).order_by(Organisation.name)
    result = await db.execute(stmt)
    return result.scalars().all()


async def count_organisations(
    db: AsyncSession,
    *,
    type_id: Optional[int] = None,
    search: Optional[str] = None,
) -> int:
    """Count organisations matching filters (for pagination metadata)."""
    stmt = select(func.count(Organisation.id)).where(
        Organisation.deleted == False
    )
    if type_id:
        stmt = stmt.where(Organisation.type_id == type_id)
    if search:
        stmt = stmt.where(Organisation.name.ilike(f"%{search}%"))
    result = await db.execute(stmt)
    return result.scalar_one()


async def create_organisation(
    db: AsyncSession,
    org_in: OrganisationCreate,
    *,
    user_id: Optional[int] = None,
) -> Organisation:
    """Create a new organisation."""
    org = Organisation(
        name=org_in.name,
        acronym=org_in.acronym,
        type_id=org_in.type_id,
        website=org_in.website,
        email=org_in.email,
        phone=org_in.phone,
    )
    db.add(org)
    await db.flush()  # get ID without committing
    return org


async def update_organisation(
    db: AsyncSession,
    org: Organisation,
    org_in: OrganisationUpdate,
) -> Organisation:
    """Update an organisation (partial update)."""
    update_data = org_in.model_dump(exclude_unset=True)
    for field, value in update_data.items():
        setattr(org, field, value)
    db.add(org)
    await db.flush()
    return org


async def soft_delete_organisation(
    db: AsyncSession,
    org: Organisation,
) -> None:
    """Soft-delete an organisation (set deleted=True)."""
    org.deleted = True
    db.add(org)
    await db.flush()''', language="python")

add_heading_h2(doc, "6.6  Router with CRUD (app/api/routers/org.py)")

add_code_block(doc, '''# app/api/routers/org.py — REST API for organisations

from fastapi import APIRouter, Depends, HTTPException, Query, status
from typing import Annotated, Optional

from app.api.dependencies.core import DBSessionDep
from app.api.dependencies.auth import CurrentUserDep, AdminUserDep
from app.crud import org as crud_org
from app.schemas.org import (
    OrganisationRead,
    OrganisationCreate,
    OrganisationUpdate,
    OrganisationList,
)

router = APIRouter(prefix="/eden/org", tags=["Organisations"])


@router.get("/", response_model=OrganisationList)
async def list_organisations(
    db: DBSessionDep,
    current_user: CurrentUserDep,         # requires auth
    skip: int = Query(0, ge=0),
    limit: int = Query(25, ge=1, le=100),
    type_id: Optional[int] = None,
    search: Optional[str] = None,
):
    """List organisations (paginated, filterable)."""
    orgs = await crud_org.get_organisations(
        db, skip=skip, limit=limit, type_id=type_id, search=search
    )
    total = await crud_org.count_organisations(
        db, type_id=type_id, search=search
    )
    return OrganisationList(
        items=[OrganisationRead.model_validate(o) for o in orgs],
        total=total,
        skip=skip,
        limit=limit,
    )


@router.get("/{org_id}", response_model=OrganisationRead)
async def get_organisation(
    org_id: int,
    db: DBSessionDep,
    current_user: CurrentUserDep,
):
    """Get a single organisation by ID."""
    org = await crud_org.get_organisation(db, org_id)
    if not org:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Organisation {org_id} not found",
        )
    return OrganisationRead.model_validate(org)


@router.post(
    "/",
    response_model=OrganisationRead,
    status_code=status.HTTP_201_CREATED,
)
async def create_organisation(
    org_in: OrganisationCreate,
    db: DBSessionDep,
    current_user: AdminUserDep,    # requires ADMIN role
):
    """Create a new organisation (admin only)."""
    org = await crud_org.create_organisation(db, org_in)
    await db.commit()              # commit in router, not CRUD
    await db.refresh(org)
    return OrganisationRead.model_validate(org)


@router.patch("/{org_id}", response_model=OrganisationRead)
async def update_organisation(
    org_id: int,
    org_in: OrganisationUpdate,
    db: DBSessionDep,
    current_user: AdminUserDep,
):
    """Update an organisation (admin only, partial update)."""
    org = await crud_org.get_organisation(db, org_id)
    if not org:
        raise HTTPException(status_code=404, detail="Not found")
    
    org = await crud_org.update_organisation(db, org, org_in)
    await db.commit()
    await db.refresh(org)
    return OrganisationRead.model_validate(org)


@router.delete("/{org_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_organisation(
    org_id: int,
    db: DBSessionDep,
    current_user: AdminUserDep,
):
    """Soft-delete an organisation (admin only)."""
    org = await crud_org.get_organisation(db, org_id)
    if not org:
        raise HTTPException(status_code=404, detail="Not found")
    
    await crud_org.soft_delete_organisation(db, org)
    await db.commit()''', language="python")

# ============================================================
# บทที่ 7: Auth
# ============================================================
add_heading_h1(doc, "Code Examples: Authentication & Authorization", num=7)

add_heading_h2(doc, "7.1  BEFORE — web2py Auth (built-in)")

add_code_block(doc, '''# web2py — built-in auth (gluon.tools.Auth)
# ไม่ต้องเขียน code มาก — ใช้ current.auth ได้เลย

# In controllers/default.py:
def user():
    return dict(form=auth())

# Check if logged in:
if auth.user:
    user_id = auth.user.id
    email = auth.user.email

# Check role:
if auth.has_membership("ADMIN"):
    ...

# Login required:
@auth.requires_login()
def secret():
    return dict(message="only for logged-in users")

# Role required:
@auth.requires_membership("ADMIN")
def admin_only():
    ...

# Permission check:
if auth.s3_has_permission("create", "org_organisation"):
    ...

# Limitation:
# - Sync only (block thread during DB query)
# - Global state (current.auth)
# - Hard to test (no DI)
# - No JWT / OAuth2 native (must implement manually)''', language="python")

add_heading_h2(doc, "7.2  AFTER — FastAPI Security (JWT + Dependencies)")

add_code_block(doc, '''# app/core/security.py — JWT + password hashing

from datetime import datetime, timedelta, timezone
from typing import Optional
import jwt
from passlib.context import CryptContext

from app.core.config import settings

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")


def verify_password(plain: str, hashed: str) -> bool:
    return pwd_context.verify(plain, hashed)


def hash_password(password: str) -> str:
    return pwd_context.hash(password)


def create_access_token(
    subject: str,
    expires_minutes: int = 60 * 24,  # 24h default
    extra_claims: dict = None,
) -> str:
    expire = datetime.now(timezone.utc) + timedelta(minutes=expires_minutes)
    to_encode = {
        "sub": subject,
        "exp": expire,
        "iat": datetime.now(timezone.utc),
        **(extra_claims or {}),
    }
    return jwt.encode(
        to_encode,
        settings.jwt_secret_key,
        algorithm=settings.jwt_algorithm,
    )


def decode_token(token: str) -> Optional[dict]:
    try:
        return jwt.decode(
            token,
            settings.jwt_secret_key,
            algorithms=[settings.jwt_algorithm],
        )
    except jwt.PyJWTError:
        return None


# ─────────────────────────────────────────────────────────────────────────
# app/api/dependencies/auth.py — Auth dependencies

from fastapi import Depends, HTTPException, status
from fastapi.security import (
    OAuth2PasswordBearer,
    OAuth2PasswordRequestForm,
)
from sqlalchemy.ext.asyncio import AsyncSession
from typing import Annotated, Optional

from app.api.dependencies.core import DBSessionDep
from app.core.security import verify_password, create_access_token, decode_token
from app.models.auth import User
from app.schemas.auth import Token

# OAuth2 with password flow
oauth2_scheme = OAuth2PasswordBearer(
    tokenUrl="/eden/auth/token",
    auto_error=False,
)


async def get_current_user(
    db: DBSessionDep,
    token: Annotated[Optional[str], Depends(oauth2_scheme)],
) -> Optional[User]:
    """Get current user from JWT token (None if not authenticated)."""
    if not token:
        return None
    
    payload = decode_token(token)
    if not payload:
        return None
    
    user_id = payload.get("sub")
    if not user_id:
        return None
    
    from sqlalchemy import select
    stmt = select(User).where(User.id == int(user_id), User.deleted == False)
    result = await db.execute(stmt)
    return result.scalar_one_or_none()


# Optional auth — None if not logged in
CurrentUserOptional = Annotated[Optional[User], Depends(get_current_user)]


async def get_current_user_required(
    current: CurrentUserOptional,
) -> User:
    """Require authentication — raise 401 if not logged in."""
    if not current:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Not authenticated",
            headers={"WWW-Authenticate": "Bearer"},
        )
    return current


# Required auth
CurrentUserDep = Annotated[User, Depends(get_current_user_required)]


async def require_admin(
    current: CurrentUserDep,
) -> User:
    """Require ADMIN role."""
    if not any(role.name == "ADMIN" for role in current.roles):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Admin access required",
        )
    return current


AdminUserDep = Annotated[User, Depends(require_admin)]


# ─────────────────────────────────────────────────────────────────────────
# app/api/routers/auth.py — Login endpoint

from fastapi import APIRouter
from sqlalchemy import select

router = APIRouter(prefix="/eden/auth", tags=["Authentication"])


@router.post("/token", response_model=Token)
async def login_for_access_token(
    form_data: Annotated[OAuth2PasswordRequestForm, Depends()],
    db: DBSessionDep,
):
    """OAuth2 password flow — exchange email/password for JWT."""
    stmt = select(User).where(
        User.email == form_data.username,
        User.deleted == False,
    )
    result = await db.execute(stmt)
    user = result.scalar_one_or_none()
    
    if not user or not verify_password(form_data.password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password",
            headers={"WWW-Authenticate": "Bearer"},
        )
    
    # Update last login
    user.last_login_on = datetime.utcnow()
    db.add(user)
    await db.commit()
    
    access_token = create_access_token(
        subject=str(user.id),
        extra_claims={
            "email": user.email,
            "roles": [r.name for r in user.roles],
        },
    )
    return Token(
        access_token=access_token,
        token_type="bearer",
        expires_in=60 * 24,
    )''', language="python")

add_heading_h2(doc, "7.3  Auth Comparison")

add_table(doc,
    headers=["Aspect", "web2py Auth", "FastAPI Security"],
    rows=[
        ["State",          "Global (current.auth)",                  "Dependency injection"],
        ["Token type",     "Session cookie (signed)",                "JWT (stateless)"],
        ["Async",          "Sync (block thread)",                     "Async throughout"],
        ["OAuth2",         "Manual implementation",                   "Native (OAuth2PasswordBearer)"],
        ["JWT",            "ไม่มี built-in",                          "python-jose / PyJWT"],
        ["Password hash",  "PBKDF2 (built-in)",                       "passlib (bcrypt/argon2)"],
        ["Roles",          "auth_group + auth_membership",            "User.roles relationship"],
        ["Permissions",    "s3_permission table",                     "RBAC dependency"],
        ["Testing",        "Hard (session-based)",                    "Easy (override dependency)"],
        ["OpenAPI",        "ไม่แสดง auth requirements",                "Auto-shows security schemes"],
        ["2FA",            "ไม่มี built-in",                          "pyotp integration easy"],
        ["Rate limit",     "Manual in controller",                    "slowapi dependency"],
    ],
    col_widths_cm=[3.0, 5.0, 8.0],
    first_col_bold=True)

# ============================================================
# บทที่ 8: Templates
# ============================================================
add_heading_h1(doc, "Code Examples: Templates & Frontend Integration", num=8)

add_heading_h2(doc, "8.1  BEFORE — web2py YATL Template")

add_code_block(doc, '''<!-- views/org/organisation.html (web2py YATL — current) -->
{{extend 'layout.html'}}

{{block section_class}}section-width-1{{end}}

<h1>{{=T("Organisations")}}</h1>

{{if messages:}}
<div class="alert alert-success">
  {{for msg in messages:}}
    <p>{{=msg}}</p>
  {{pass}}
</div>
{{pass}}

<div class="row">
  <div class="small-12 columns">
    <a href="{{=URL('org', 'organisation', args='create')}}"
       class="button primary">
      {{=T("Create Organisation")}}
    </a>
  </div>
</div>

<table class="data-table">
  <thead>
    <tr>
      <th>{{=T("Name")}}</th>
      <th>{{=T("Type")}}</th>
      <th>{{=T("Website")}}</th>
      <th>{{=T("Actions")}}</th>
    </tr>
  </thead>
  <tbody>
    {{for org in organisations:}}
    <tr>
      <td>{{=org.name}}</td>
      <td>{{=org.type.name if org.type else '-'}}</td>
      <td>
        {{if org.website:}}
          <a href="{{=org.website}}" target="_blank">{{=org.website}}</a>
        {{pass}}
      </td>
      <td>
        <a href="{{=URL('org', 'organisation', args=[org.id, 'read'])}}"
           class="button small">Open</a>
      </td>
    </tr>
    {{pass}}
  </tbody>
</table>

<script>
$(document).ready(function() {
  $('.data-table').DataTable({
    pageLength: 25,
    serverSide: true,
    ajax: '{{=URL("org", "organisation", extension="aadata")}}'
  });
});
</script>''', language="html")

add_heading_h2(doc, "8.2  AFTER — Jinja2 + HTMX (Server-rendered)")

add_code_block(doc, '''<!-- app/templates/org/list.html (Jinja2 — target) -->
{% extends "base.html" %}
{% from "macros/pagination.html" import pagination %}

{% block title %}{{ _("Organisations") }}{% endblock %}

{% block content %}
<h1>{{ _("Organisations") }}</h1>

{% if messages %}
<div class="alert alert-success">
  {% for msg in messages %}
    <p>{{ msg }}</p>
  {% endfor %}
</div>
{% endif %}

<div class="row mb-3">
  <div class="col">
    <a href="{{ url_for('org.create_organisation_form') }}"
       class="btn btn-primary">
      {{ _("Create Organisation") }}
    </a>
  </div>
</div>

<!-- HTMX-powered search (no JS) -->
<input type="text"
       name="search"
       placeholder="{{ _('Search...') }}"
       hx-get="/eden/api/v2/org/organisations"
       hx-target="#org-table"
       hx-trigger="keyup changed delay:300ms"
       hx-include="[name='type_id']"
       hx-select="tbody tr"
       class="form-control mb-3">

<table id="org-table" class="table table-striped">
  <thead>
    <tr>
      <th>{{ _("Name") }}</th>
      <th>{{ _("Type") }}</th>
      <th>{{ _("Website") }}</th>
      <th>{{ _("Actions") }}</th>
    </tr>
  </thead>
  <tbody>
    {% for org in organisations %}
    <tr>
      <td>{{ org.name }}</td>
      <td>{{ org.type.name if org.type else '-' }}</td>
      <td>
        {% if org.website %}
          <a href="{{ org.website }}" target="_blank">{{ org.website }}</a>
        {% endif %}
      </td>
      <td>
        <a href="{{ url_for('org.get_organisation', org_id=org.id) }}"
           class="btn btn-sm btn-primary">Open</a>
      </td>
    </tr>
    {% else %}
    <tr>
      <td colspan="4" class="text-center">
        {{ _("No organisations found") }}
      </td>
    </tr>
    {% endfor %}
  </tbody>
</table>

{{ pagination(total=total, skip=skip, limit=limit) }}
{% endblock %}


<!-- app/templates/base.html — Jinja2 master template -->
<!DOCTYPE html>
<html lang="{{ current_language }}">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{% block title %}{{ _("Sahana Eden") }}{% endblock %}</title>
  
  <!-- Tailwind CSS -->
  <link rel="stylesheet" href="/eden/static/dist/main.css">
  
  <!-- HTMX -->
  <script src="https://unpkg.com/htmx.org@2.0.0"></script>
  
  <!-- Alpine.js (for reactive widgets) -->
  <script defer src="https://unpkg.com/alpinejs@3.14.0/dist/cdn.min.js"></script>
</head>
<body>
  {% include "partials/header.html" %}
  
  <main class="container mx-auto px-4 py-8">
    {% block content %}{% endblock %}
  </main>
  
  {% include "partials/footer.html" %}
</body>
</html>''', language="html")

add_heading_h2(doc, "8.3  FastAPI: Jinja2 + HTMX Integration")

add_code_block(doc, '''# app/api/routers/pages.py — Server-rendered pages with Jinja2

from fastapi import APIRouter, Depends, Request, Query
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from typing import Optional

from app.api.dependencies.core import DBSessionDep
from app.api.dependencies.auth import CurrentUserOptional
from app.crud import org as crud_org

router = APIRouter(tags=["Pages"])
templates = Jinja2Templates(directory="app/templates")


@router.get("/eden/org/organisations", response_class=HTMLResponse)
async def list_organisations_page(
    request: Request,
    db: DBSessionDep,
    current_user: CurrentUserOptional,
    skip: int = Query(0, ge=0),
    limit: int = Query(25, ge=1, le=100),
    search: Optional[str] = None,
):
    """Server-rendered list of organisations."""
    orgs = await crud_org.get_organisations(
        db, skip=skip, limit=limit, search=search
    )
    total = await crud_org.count_organisations(db, search=search)
    
    return templates.TemplateResponse(
        "org/list.html",
        {
            "request": request,
            "organisations": orgs,
            "total": total,
            "skip": skip,
            "limit": limit,
            "search": search,
            "current_user": current_user,
        },
    )''', language="python")

# ============================================================
# บทที่ 9: Coexistence
# ============================================================
add_heading_h1(doc, "Coexistence Strategy (รันทั้งสองระบบชั่วคราว)", num=9)

add_heading_h2(doc, "9.1  Shared Database Strategy")

add_para(doc,
    "ในระหว่าง migration ทั้ง web2py และ FastAPI ใช้ database เดียวกัน "
    "ตารางเดิมที่สร้างโดย PyDAL สามารถอ่าน/เขียนโดย SQLAlchemy 2.0 ได้ "
    "เพราะ schema เหมือนกัน (เป็น SQL table ปกติ) แค่ต้องระวัง:",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_bullet(doc, "อย่า migrate schema (Alembic) จนกว่า web2py จะถูก decommission", bold_prefix="⚠ ")
add_bullet(doc, "web2py's auto-migrate (migrate=True) ต้องปิด — ใช้ Alembic อย่างเดียว", bold_prefix="⚠ ")
add_bullet(doc, "UUID fields ต้อง compatible — ทั้งคู่ใช้ string 128 chars", bold_prefix="✓ ")
add_bullet(doc, "MetaFields (created_on, modified_on, deleted) ต้องมีใน SQLAlchemy models", bold_prefix="✓ ")
add_bullet(doc, "Soft-delete pattern เหมือนกัน — deleted=True/False", bold_prefix="✓ ")

add_heading_h2(doc, "9.2  Shared Auth Strategy")

add_para(doc,
    "ปัญหาใหญ่: web2py ใช้ session cookie, FastAPI ใช้ JWT "
    "ถ้า user login ผ่าน web2py แล้วไปใช้ FastAPI endpoint ต้อง re-login "
    "วิธีแก้:",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''# Option A: Dual-session bridge (recommended for transition)

# web2py: after login, also set a JWT cookie
# (modify controllers/default.py user() function)

def user():
    form = auth()
    if form.accepted and auth.user:
        # Generate JWT token for FastAPI
        from s3.fastapi_bridge import create_jwt_for_user
        jwt_token = create_jwt_for_user(auth.user.id)
        response.cookies["eden_jwt"] = jwt_token
        response.cookies["eden_jwt"]["path"] = "/"
        response.cookies["eden_jwt"]["httponly"] = True
        response.cookies["eden_jwt"]["secure"] = request.is_https
    return dict(form=form)


# Option B: SSO via OAuth2 (longer term)

# web2py: act as OAuth2 provider
# FastAPI: act as OAuth2 client/resource server
# User logs in once at web2py, gets OAuth2 token, FastAPI validates it


# Option C: Migrate auth FIRST (Phase 3 of roadmap)

# Migrate /eden/default/user/* to FastAPI early
# Both web2py and FastAPI validate against auth_user table
# Set both session cookie AND JWT on login
# Eventually decommission web2py session cookie''', language="python")

add_heading_h2(doc, "9.3  Monitoring Coexistence")

add_code_block(doc, '''# docker-compose.yml — both systems running together

services:
  web2py:
    image: eden-web2py:latest
    ports: ["8000"]            # internal only
    environment:
      EDEN_ENV: production
      DATABASE_URL: postgresql://eden:${POSTGRES_PASSWORD}@postgres/eden
    depends_on: [postgres, redis]
    volumes:
      - eden_uploads:/app/web2py/applications/eden/uploads
  
  fastapi:
    image: eden-fastapi:latest
    ports: ["8001"]            # internal only
    environment:
      DATABASE_URL: postgresql+asyncpg://eden:${POSTGRES_PASSWORD}@postgres/eden
      REDIS_URL: redis://redis:6379/0
      JWT_SECRET_KEY: ${JWT_SECRET_KEY}
    depends_on: [postgres, redis]
    volumes:
      - eden_uploads:/app/uploads    # shared!
  
  caddy:
    image: caddy:2-alpine
    ports: ["80:80", "443:443"]
    volumes:
      - ./Caddyfile:/etc/caddy/Caddyfile:ro
    depends_on: [web2py, fastapi]
  
  postgres:
    image: postgres:16-alpine
    environment:
      POSTGRES_DB: eden
      POSTGRES_USER: eden
      POSTGRES_PASSWORD: ${POSTGRES_PASSWORD}
    volumes: [pg_data:/var/lib/postgresql/data]
  
  redis:
    image: redis:7-alpine
    volumes: [redis_data:/data]

# Health checks for both
#   GET /eden/health       → FastAPI (new)
#   GET /eden/appadmin/health  → web2py (legacy)
# Caddy returns 502 if either is down''', language="yaml")

# ============================================================
# บทที่ 10: Migration Phases
# ============================================================
add_heading_h1(doc, "Migration Phases Roadmap", num=10)

add_table(doc,
    headers=["Phase", "ระยะเวลา", "งาน", "ผลลัพธ์"],
    rows=[
        ["Phase 0: Setup",      "สัปดาห์ 1",    "สร้าง FastAPI skeleton, Dockerfile, docker-compose, CI/CD",                "fastapi container รันได้ (empty app)"],
        ["Phase 1: AI modules", "สัปดาห์ 2-4",  "Migrate 10 AI controllers → FastAPI routers + Pydantic schemas",            "AI endpoints ทำงานบน FastAPI"],
        ["Phase 2: Foundation", "เดือน 2",      "Auth (login + JWT), User model, Organisation CRUD",                          "Auth ทำงานบน FastAPI"],
        ["Phase 3: Read API",   "เดือน 3",      "Add /eden/api/v2/* read endpoints สำหรับ org/pr/inv/hrm",                  "OpenAPI docs ใช้งานได้"],
        ["Phase 4: Write API",  "เดือน 4-5",    "Add write endpoints (create/update/delete) + Pydantic schemas",              "Full CRUD สำหรับ core modules"],
        ["Phase 5: Frontend",   "เดือน 5-6",    "Migrate Jinja2 templates + HTMX (ทีละ module)",                              "UI ใช้ FastAPI ทั้งหมด"],
        ["Phase 6: Cleanup",    "เดือน 6+",    "Decommission web2py, ลบ Caddy strangler routes",                              "web2py shutdown"],
    ],
    col_widths_cm=[3.0, 2.0, 6.5, 4.5],
    first_col_bold=True, font_size=9, header_size=9.5)

add_heading_h2(doc, "10.1  Risk Mitigation")

add_table(doc,
    headers=["Risk", "Mitigation"],
    rows=[
        ["Database corruption",       "Backup ก่อน migrate, ใช้ Alembic offline mode สำหรับ verification"],
        ["Auth mismatch",             "Bridge cookies (Option A) ก่อน migrate auth fully"],
        ["Performance regression",    "Load test FastAPI version ก่อน cutover"],
        ["PDPA violation",            "Audit consent flow ทุก endpoint, encrypt sensitive columns"],
        ["Downtime",                  "Blue-green deploy — run both, switch Caddy routing atomically"],
        ["Rollback need",             "เก็บ web2py image ไว้ 30 วันหลัง decommission"],
        ["Data inconsistency",        "Single source of truth (PostgreSQL) — ไม่ใช้ cache ใน transition"],
        ["Developer confusion",       "Clear ADRs + runbooks + dev docs"],
        ["External API breakage",     "Keep /eden/* paths เดิม — ใช้ Caddy route ไม่ใช่ path change"],
        ["Test coverage gap",         "80% coverage gate ใน CI ก่อน migrate แต่ละ module"],
    ],
    col_widths_cm=[4.5, 11.5],
    first_col_bold=True)

doc.add_page_break()

# ============================================================
# PART B DIVIDER
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(80)
add_run(p, "PART B", font_name=FONT_HEADING, size_pt=42, bold=True, color=C_PRIMARY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(20)
add_run(p, "Language Comparison",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "Python vs Go vs Rust vs TypeScript — เปรียบเทียบสำหรับทุก use case",
        font_name=FONT_BODY, size_pt=13, italic=True, color=C_SECOND)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "อ้างอิงข้อมูลล่าสุดจาก blog/medium/Reddit/TIOBE (ปลาย 2026)",
        font_name=FONT_BODY, size_pt=10, color=C_SECOND)

doc.add_page_break()

# ============================================================
# บทที่ 11: ภาพรวม 4 ภาษา
# ============================================================
add_heading_h1(doc, "ภาพรวมการเปรียบเทียบ 4 ภาษา", num=11)

add_heading_h2(doc, "11.1  Quick Summary (ปลาย 2026)")

add_table(doc,
    headers=["ภาษา", "ปีที่สร้าง", "Creator", "Paradigm", "Compilation", "ปัจจุบัน TIOBE"],
    rows=[
        ["Python",     "1991", "Guido van Rossum",      "Multi-paradigm, dynamic",    "Interpreted",          "#1"],
        ["Go",         "2009", "Google (Robert Pike)",   "Concurrent, procedural",     "Compiled (static)",    "Top 10"],
        ["Rust",       "2010", "Mozilla (Graydon Hoare)","Memory-safe, multi-paradigm","Compiled (no GC)",     "Top 15"],
        ["TypeScript", "2012", "Microsoft (Anders H.)",  "Object, static types on JS", "Transpiled to JS",     "Top 5"],
    ],
    col_widths_cm=[2.5, 1.8, 4.0, 4.0, 3.0, 1.7],
    first_col_bold=True, font_size=9)

add_heading_h2(doc, "11.2  Performance Benchmarks (สัมพัทธ์)")

add_table(doc,
    headers=["ภาษา", "CPU-bound", "Memory usage", "Startup time", "Binary size"],
    rows=[
        ["C/C++ (baseline)",   "1.0x",   "1.0x",   "1.0x",   "1.0x"],
        ["Rust",               "1.0-1.2x","1.0-1.5x","1.2x",  "1.5x"],
        ["Go",                 "1.5-3x", "2-3x",   "1.5x",   "3-5x"],
        ["Java (JIT)",         "1.5-3x", "5-10x",  "10-20x", "— (JVM)"],
        ["C# (.NET)",          "2-4x",   "5-8x",   "10-15x", "— (runtime)"],
        ["JavaScript (V8)",    "2-5x",   "3-5x",   "5-10x",  "— (runtime)"],
        ["TypeScript→JS",      "2-5x",   "3-5x",   "5-10x",  "— (runtime)"],
        ["Python 3.12",        "10-30x", "5-15x",  "1x",     "— (runtime)"],
        ["Python + PyPy",      "3-7x",   "5-10x",  "5-10x",  "— (runtime)"],
    ],
    col_widths_cm=[3.0, 2.5, 2.5, 2.5, 2.5],
    first_col_bold=True, font_size=9)

add_note_box(doc, "หมายเหตุเรื่อง benchmark",
    "ตัวเลขเป็นสัมพัทธ์และขึ้นกับ use case จริง — ส่วนใหญ่ application เป็น I/O-bound "
    "(DB query, HTTP) ไม่ใช่ CPU-bound ดังนั้น performance difference มีผลน้อย "
    "ยกเว้นงาน performance-critical เช่น ML training, image processing, parsing",
    color="0E7490", bg="ECFDF5")

add_heading_h2(doc, "11.3  Development Speed (Productivity)")

add_table(doc,
    headers=["ภาษา", "Lines of code (สัมพัทธ์)", "Time-to-prototype", "Refactoring ease"],
    rows=[
        ["Python",     "1.0x (baseline)",   "1.0x (เร็วสุด)",       "★★★★★"],
        ["TypeScript", "1.2-1.5x",           "1.2x",                "★★★★☆"],
        ["Go",         "1.5-2x",             "1.5-2x",              "★★★☆☆"],
        ["Rust",       "2-3x (verbose)",    "2-3x",                "★★☆☆☆"],
    ],
    col_widths_cm=[2.5, 4.0, 4.0, 4.0],
    first_col_bold=True)

add_heading_h2(doc, "11.4  Ecosystem & Hiring")

add_table(doc,
    headers=["ภาษา", "Ecosystem size", "Annual new packages", "Job postings", "Avg salary (US)"],
    rows=[
        ["Python",     "PyPI: 500k+",      "100k+",            "สูงสุด (AI/data)",     "$110-180k"],
        ["TypeScript", "npm: 4M+ (with JS)","200k+",           "สูงมาก (frontend)",     "$120-200k"],
        ["Go",         "Mod: 200k+",        "30k+",             "สูง (DevOps/backend)",   "$140-220k"],
        ["Rust",       "Crates: 150k+",     "20k+",             "กลาง (เติบโตเร็ว)",      "$150-230k"],
    ],
    col_widths_cm=[2.5, 3.5, 3.0, 3.5, 3.5],
    first_col_bold=True, font_size=9)

# ============================================================
# บทที่ 12: Use Case Matrix
# ============================================================
add_heading_h1(doc, "Use Case Matrix (14 use cases)", num=12)

add_table(doc,
    headers=["Use Case", "Python", "Go", "Rust", "TypeScript", "แนะนำ"],
    rows=[
        ["AI / ML",                  "★★★★★", "★★",      "★★★",     "★★",      "Python"],
        ["Web Backend",              "★★★★",  "★★★★★",   "★★★★",    "★★★★★",   "แล้วแต่ context"],
        ["Microservices / API",      "★★★",   "★★★★★",   "★★★★",    "★★★★",    "Go"],
        ["Frontend (browser)",       "—",     "—",       "— (wasm)", "★★★★★",   "TypeScript"],
        ["Mobile App",               "★",     "—",       "—",        "★★★★",    "TypeScript (RN)"],
        ["Systems Programming",      "★★",    "★★★★",    "★★★★★",    "—",       "Rust"],
        ["DevOps / CLI tools",       "★★★",   "★★★★★",   "★★★★",    "★★★",     "Go"],
        ["Data Science / ETL",       "★★★★★", "★★",      "★★★",     "★★★",     "Python"],
        ["Real-time Streaming",      "★★★",   "★★★★",    "★★★★★",    "★★★★",    "Rust / Go"],
        ["GIS / Mapping (backend)",  "★★★★★", "★★★",     "★★",      "★★",      "Python"],
        ["Background Jobs",          "★★★★",  "★★★",     "★★",      "★★★",     "Python (Celery)"],
        ["High-traffic API Gateway", "★★",    "★★★★★",   "★★★★",    "★★★",     "Go"],
        ["Game Development",         "★★",    "★★",      "★★★★",    "★★★",     "Rust / TS"],
        ["Enterprise Apps",          "★★★",   "★★★",     "★★",      "★★★★",    "Java/TS"],
    ],
    col_widths_cm=[4.0, 2.0, 2.0, 2.0, 2.5, 3.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "12.1  Scoring Legend")
add_bullet(doc, "★★★★★ = Excellent choice — best-in-class ecosystem, native support, mature")
add_bullet(doc, "★★★★ = Good choice — viable with minor trade-offs")
add_bullet(doc, "★★★ = Acceptable — works but not optimal")
add_bullet(doc, "★★ = Poor — limited ecosystem or painful")
add_bullet(doc, "★ = Bad — barely possible, avoid")
add_bullet(doc, "— = Not applicable (e.g., Python ไม่ใช่ใน browser)")

# ============================================================
# บทที่ 13: AI/ML Stack
# ============================================================
add_heading_h1(doc, "เจาะลึก: AI/ML Stack", num=13)

add_heading_h2(doc, "13.1  Ecosystem Comparison")

add_table(doc,
    headers=["Library / Tool", "Python", "Go", "Rust", "TypeScript"],
    rows=[
        ["PyTorch / TensorFlow",    "★★★★★ (native)",      "— (bindings)",   "★★★ (tch-rs)",     "★★ (TF.js)"],
        ["LangChain / LlamaIndex",  "★★★★★ (native)",      "★ (limited)",    "★ (limited)",      "★★★ (LangChain.js)"],
        ["Hugging Face transformers","★★★★★ (native)",     "—",              "★★ (candle)",      "★★ (transformers.js)"],
        ["OpenAI / Anthropic SDK",  "★★★★★ (official)",    "★★ (community)", "★★ (community)",   "★★★★ (official)"],
        ["scikit-learn",            "★★★★★ (native)",      "—",              "★★ (linfa)",       "—"],
        ["pandas / NumPy",          "★★★★★ (native)",      "★★★ (gota)",     "★★ (polars)",      "—"],
        ["Polars / DuckDB",         "★★★★★ (Python API)",  "—",              "★★★★★ (Rust native)","—"],
        ["Jupyter Notebook",        "★★★★★ (native)",      "—",              "★★ (evcxr)",       "★★ (nteract)"],
        ["MLflow / W&B",            "★★★★★ (native)",      "—",              "—",                "★★ (limited)"],
        ["Embedding models",        "★★★★★ (sentence-transformers)","—",     "★★ (candle)",      "★★ (transformers.js)"],
    ],
    col_widths_cm=[3.5, 3.0, 2.5, 3.0, 3.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "13.2  คำแนะนำสำหรับ AI/ML")

add_para(doc,
    "Python เป็นเจ้าตลาด AI/ML ที่แทบจะ monopoly — ไม่มีภาษาอื่นที่ทัดเทียม "
    "ในแง่ของ library ecosystem, community, และเครื่องมือ "
    "Go และ Rust มี performance ดีกว่าแต่ ecosystem น้อยกว่ามาก "
    "TypeScript ใช้ได้สำหรับ frontend AI (เช่น transformers.js ใน browser) "
    "แต่ production AI ยังเป็น Python",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "Hybrid AI Stack (Recommended)")

add_code_block(doc, '''# Production AI architecture (hybrid)

┌──────────────────────────────────────────────────────────────────┐
│ Frontend (TypeScript)                                              │
│  • transformers.js สำหรับ small model in browser                  │
│  • Call backend API สำหรับ heavy inference                         │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Backend API (Python — FastAPI)                                     │
│  • LangChain orchestration                                          │
│  • Pydantic schema validation                                       │
│  • async LLM calls (aiohttp)                                        │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ AI Libraries (Python)                                              │
│  • PyTorch (training)                                               │
│  • Hugging Face transformers (inference)                           │
│  • sentence-transformers (embeddings)                              │
│  • LangChain (RAG / agents)                                        │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Performance Layer (Rust)                                           │
│  • Polars (DataFrame — เร็วกว่า pandas 10-30x)                     │
│  • Candle (ML inference — เร็วกว่า PyTorch สำหรับ small models)   │
│  • Ruff, uv (tooling — เร็วกว่า flake8/pip 100x)                  │
│  • ใช้ผ่าน Python API — เขียน Rust, เรียกจาก Python                 │
└──────────────────────────────────────────────────────────────────┘''', language="text")

# ============================================================
# บทที่ 14: Web Backend & Microservices
# ============================================================
add_heading_h1(doc, "เจาะลึก: Web Backend & Microservices", num=14)

add_heading_h2(doc, "14.1  Framework Comparison")

add_table(doc,
    headers=["Aspect", "FastAPI (Py)", "Go (Fiber/Gin)", "Rust (Axum)", "TS (NestJS)"],
    rows=[
        ["Performance (req/s)",   "50k-100k",   "100k-200k",  "200k-500k",  "50k-100k"],
        ["Startup time",          "~500ms",     "~50ms",      "~20ms",      "~500ms"],
        ["Memory per request",    "5-15MB",     "1-3MB",      "0.5-2MB",    "5-10MB"],
        ["Hot reload",            "uvicorn --reload","air",    "cargo watch","nodemon"],
        ["Async model",           "asyncio",    "goroutines", "tokio",      "async/await"],
        ["Type safety",           "Type hints", "Native",     "Native",     "Native"],
        ["OpenAPI/Swagger",       "Auto",       "Manual",     "utoipa",     "Auto"],
        ["Microservices",         "Good",       "Excellent",  "Excellent",  "Good"],
        ["gRPC",                  "grpclib",    "grpc-go",    "tonic",      "@grpc/*"],
        ["WebSocket",             "Native",     "Native",     "tokio-tungstenite","ws"],
        ["DB drivers",            "asyncpg",    "pgx",        "sqlx",       "pg / prisma"],
        ["ORM",                   "SQLAlchemy 2.0","sqlx/gorm","sea-orm",    "Prisma/TypeORM"],
    ],
    col_widths_cm=[3.0, 3.5, 3.0, 3.0, 3.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "14.2  When to Use Each")

add_bullet(doc, "Use FastAPI เมื่อ: prototype เร็ว, AI integration heavy, team เป็น Python", bold_prefix="▸ ")
add_bullet(doc, "Use Go เมื่อ: high-traffic microservices, DevOps tools, CLI", bold_prefix="▸ ")
add_bullet(doc, "Use Rust เมื่อ: performance-critical (latency < 1ms), embedded, systems", bold_prefix="▸ ")
add_bullet(doc, "Use NestJS เมื่อ: full-stack JS team, real-time apps, enterprise structure", bold_prefix="▸ ")

add_heading_h2(doc, "14.3  Recommendation สำหรับ Sahana Eden")

add_code_block(doc, '''# Sahana Eden backend architecture (recommended)

┌─────────────────────────────────────────────────────────────────┐
│ API Gateway (Go — Caddy or custom)                                │
│  • TLS termination                                                 │
│  • Rate limiting                                                   │
│  • Auth validation (JWT)                                           │
│  • หน้าที่: high-throughput, low-latency                            │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│ Application Layer (Python — FastAPI)                              │
│  • Business logic                                                  │
│  • AI integration (LangChain, pgvector)                           │
│  • CRUD operations                                                 │
│  • Background job dispatch                                         │
│  • หน้าที่: feature-rich, AI ecosystem                              │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│ Performance Layer (Rust extensions via PyO3)                     │
│  • Vector search (ถ้าต้องการ < 10ms)                              │
│  • Image processing                                                │
│  • Cryptographic operations                                        │
│  • หน้าที่: hot path optimization                                   │
└─────────────────────────────────────────────────────────────────┘''', language="text")

# ============================================================
# บทที่ 15: Frontend & Mobile
# ============================================================
add_heading_h1(doc, "เจาะลึก: Frontend & Mobile", num=15)

add_heading_h2(doc, "15.1  Frontend Framework Comparison")

add_table(doc,
    headers=["Aspect", "React+TS", "Vue+TS", "Svelte+TS", "HTMX+Alpine"],
    rows=[
        ["Learning curve",        "Steep",        "Moderate",       "Easy",          "Very easy"],
        ["Bundle size",           "~50KB (gzip)", "~30KB (gzip)",   "~10KB (gzip)",  "~14KB (HTMX)"],
        ["Performance",           "Excellent",    "Excellent",      "Best",          "Server-dependent"],
        ["Ecosystem",             "Largest",      "Large",          "Growing",       "Small"],
        ["TypeScript support",    "★★★★★",        "★★★★★",          "★★★★★",         "★★★"],
        ["Server-side rendering", "Next.js",      "Nuxt",           "SvelteKit",     "Native (server-rendered)"],
        ["Mobile (RN)",           "★★★★★",        "NativeScript",   "Svelte Native",  "—"],
        ["Best for",              "Complex SPAs", "Mid-size apps",  "Performance",   "Server-rendered"],
    ],
    col_widths_cm=[3.0, 3.0, 3.0, 3.0, 3.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "15.2  Mobile App Comparison")

add_table(doc,
    headers=["Framework", "Language", "Performance", "Native feel", "Code reuse with web"],
    rows=[
        ["React Native",       "TypeScript", "★★★★",   "★★★★",   "★★★★ (share logic)"],
        ["Flutter",            "Dart",       "★★★★★", "★★★★★", "★★ (separate UI)"],
        ["Native (Swift/Kotlin)","Swift/Kotlin","★★★★★","★★★★★","—"],
        ["Capacitor (Ionic)",  "TypeScript", "★★★",    "★★★",    "★★★★★ (web code)"],
        ["Kivy (Python)",      "Python",     "★★",     "★★",     "—"],
    ],
    col_widths_cm=[3.0, 3.0, 2.5, 2.5, 4.0],
    first_col_bold=True, font_size=9, header_size=9)

add_heading_h2(doc, "15.3  Recommendation สำหรับ Sahana Eden")

add_para(doc,
    "Frontend ของ Eden ควร migrate จาก jQuery+Foundation ไปยัง HTMX+Alpine.js "
    "สำหรับ 90% ของหน้า (server-rendered, declarative) "
    "และ React/TypeScript สำหรับ complex widgets (เช่น map editor, data grid ที่ interactive มาก) "
    "สำหรับ mobile app แนะนำ React Native เพราะ share logic กับ web",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''# Sahana Eden frontend architecture (recommended)

┌──────────────────────────────────────────────────────────────────┐
│ Web App                                                            │
│                                                                    │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ 90% of pages: HTMX + Alpine.js + Tailwind                    │  │
│  │  • Server-rendered (FastAPI + Jinja2)                        │  │
│  │  • Progressive enhancement                                   │  │
│  │  • Small bundle, fast first paint                            │  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                    │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │ 10% complex widgets: React + TypeScript                      │  │
│  │  • Map editor (OpenLayers → MapLibre GL)                     │  │
│  │  • Data grid (AG Grid)                                       │  │
│  │  • Dashboard charts (Recharts/D3)                            │  │
│  │  • Loaded lazily (code-split)                                │  │
│  └────────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────────┐
│ Mobile App (Phase 4)                                              │
│                                                                    │
│  React Native + TypeScript                                         │
│  • Share business logic (TypeScript) กับ web                      │
│  • Offline-first architecture                                      │
│  • Push notifications สำหรับ CAP alerts                             │
│  • ส่งต่อไปยัง FastAPI backend                                      │
└──────────────────────────────────────────────────────────────────┘''', language="text")

# ============================================================
# บทที่ 16: Systems & CLI
# ============================================================
add_heading_h1(doc, "เจาะลึก: Systems Programming & CLI", num=16)

add_heading_h2(doc, "16.1  Systems Programming Comparison")

add_table(doc,
    headers=["Aspect", "Rust", "Go", "C/C++", "Python"],
    rows=[
        ["Memory safety",      "★★★★★ (compile-time)", "★★★★ (GC)", "★★ (manual)", "★★★ (GC)"],
        ["Performance",        "★★★★★",                "★★★★",       "★★★★★",        "★★"],
        ["Concurrency safety", "★★★★★ (Send/Sync)",    "★★★★ (CSP)", "★★ (manual)", "★★ (GIL)"],
        ["Startup time",       "★★★★★ (20ms)",         "★★★★ (50ms)","★★★★★",        "★★★ (500ms)"],
        ["Binary size",        "★★★ (1.5x)",           "★★ (3-5x)",  "★★★★★",        "— (runtime)"],
        ["Cross-compilation",  "★★★★",                 "★★★★★",      "★★",           "★★"],
        ["Ecosystem",          "Crates: 150k+",        "Mod: 200k+",  "Huge",         "PyPI: 500k+"],
        ["Learning curve",     "Steep (borrow checker)","Easy",      "Hard",         "Very easy"],
        ["Best for",           "Systems, embedded, perf","Services, CLI","Legacy systems","Tooling, scripts"],
    ],
    col_widths_cm=[3.0, 3.5, 3.0, 2.5, 3.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "16.2  CLI Tools Comparison")

add_code_block(doc, '''# CLI tool comparison (concrete example: a tool that fetches + parses JSON)

# ─── Python (click) ──────────────────────────────
import click
import requests
import json

@click.command()
@click.option("--url", required=True)
def fetch(url):
    """Fetch and parse JSON from URL."""
    r = requests.get(url)
    data = r.json()
    click.echo(json.dumps(data, indent=2))

if __name__ == "__main__":
    fetch()
# Run: python fetch.py --url http://api.example.com
# Pros: 10 lines, easy to write
# Cons: requires Python runtime, slow startup (~500ms)

# ─── Go (cobra) ───────────────────────────────────
package main

import (
    "encoding/json"
    "fmt"
    "io"
    "net/http"
    "github.com/spf13/cobra"
)

func main() {
    var url string
    cmd := &cobra.Command{
        Use: "fetch",
        Run: func(cmd *cobra.Command, args []string) {
            resp, _ := http.Get(url)
            defer resp.Body.Close()
            body, _ := io.ReadAll(resp.Body)
            var data interface{}
            json.Unmarshal(body, &data)
            b, _ := json.MarshalIndent(data, "", "  ")
            fmt.Println(string(b))
        },
    }
    cmd.Flags().StringVar(&url, "url", "", "URL to fetch")
    cmd.MarkFlagRequired("url")
    cmd.Execute()
}
# Build: go build -o fetch
# Run: ./fetch --url http://api.example.com
# Pros: single binary, fast startup (~50ms), cross-compile easy
# Cons: verbose

# ─── Rust (clap) ──────────────────────────────────
use clap::Parser;
use reqwest;
use serde_json::Value;

#[derive(Parser)]
struct Args {
    #[arg(long)]
    url: String,
}

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse();
    let resp = reqwest::get(&args.url).await?;
    let text = resp.text().await?;
    let data: Value = serde_json::from_str(&text)?;
    println!("{}", serde_json::to_string_pretty(&data)?);
    Ok(())
}
# Build: cargo build --release
# Run: ./fetch --url http://api.example.com
# Pros: fastest, single binary, memory-safe
# Cons: longest to write, compile time''', language="python")

add_heading_h2(doc, "16.3  Recommendation สำหรับ CLI tools ของ Sahana Eden")

add_bullet(doc, "Python scripts สำหรับ internal ops (one-off, no distribution needed)")
add_bullet(doc, "Go สำหรับ CLI tools ที่ distribute ให้ user (eden-cli, eden-migrate, eden-backup)")
add_bullet(doc, "Rust สำหรับ performance-critical CLI (เช่น vector index builder)")

# ============================================================
# บทที่ 17: Data Science & Streaming
# ============================================================
add_heading_h1(doc, "เจาะลึก: Data Science & Real-time Streaming", num=17)

add_heading_h2(doc, "17.1  Data Science Stack")

add_table(doc,
    headers=["Tool", "Python", "Go", "Rust", "TypeScript"],
    rows=[
        ["pandas (DataFrame)",       "★★★★★",   "★★ (gota)",      "—",              "—"],
        ["Polars (fast DataFrame)",   "★★★★★ (API)","—",            "★★★★★ (native)", "—"],
        ["DuckDB (analytics DB)",     "★★★★★ (API)","—",            "★★★★★ (native)", "—"],
        ["NumPy (arrays)",            "★★★★★",   "★★ (gonum)",     "★★ (ndarray)",   "—"],
        ["scikit-learn (ML)",         "★★★★★",   "—",              "★★ (linfa)",     "—"],
        ["Matplotlib / Plotly",       "★★★★★",   "★★ (goterm)",    "★★ (plotters)",  "★★★ (Plotly.js)"],
        ["Jupyter Notebook",          "★★★★★",   "—",              "★★ (evcxr)",     "★★ (nteract)"],
        ["Apache Spark",              "★★★★★ (PySpark)","★★ (GrokSpark)","★★★ (datafusion)","—"],
        ["Apache Arrow",              "★★★★★ (pyarrow)","★★","★★★★ (native)","★★ (arrow-js)"],
    ],
    col_widths_cm=[3.0, 3.0, 3.0, 3.5, 2.5],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "17.2  Real-time Streaming")

add_table(doc,
    headers=["Aspect", "Python (asyncio)", "Go", "Rust (tokio)", "TypeScript"],
    rows=[
        ["Concurrency model",   "async/await (single thread)","goroutines (true parallelism)","tokio (true parallelism)","async/await (event loop)"],
        ["Throughput",          "10-50k msg/s",         "100k-1M msg/s",      "1M-5M msg/s",        "10-100k msg/s"],
        ["Latency p99",         "5-50ms",                "1-10ms",             "0.1-1ms",            "5-50ms"],
        ["Memory per conn",     "5-15MB",                "1-3MB",              "0.5-2MB",            "5-10MB"],
        ["Kafka client",        "aiokafka",              "confluent-kafka-go", "rdkafka",            "kafkajs"],
        ["WebSocket",           "★★★★",                  "★★★★★",              "★★★★★",              "★★★★"],
        ["SSE",                 "★★★★",                  "★★★★",               "★★★★",               "★★★★★"],
        ["Best for",            "ML/data pipelines",    "High-throughput",   "Ultra-low latency", "Browser/Web"],
    ],
    col_widths_cm=[3.0, 3.5, 3.0, 3.0, 2.5],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "17.3  Recommendation สำหรับ Sahana Eden Data Pipeline")

add_code_block(doc, '''# Sahana Eden data pipeline (recommended)

┌──────────────────────────────────────────────────────────────────┐
│ Ingestion Layer (Python — async)                                  │
│  • FastAPI endpoints รับ data                                       │
│  • Celery workers สำหรับ batch                                     │
│  • aiokafka สำหรับ streaming                                        │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Processing Layer (Python + Rust)                                  │
│  • Python: business logic, AI inference                            │
│  • Rust (via PyO3): Polars สำหรับ large DataFrames                 │
│  • Rust (via PyO3): DataFusion สำหรับ SQL queries on Parquet       │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Storage Layer                                                     │
│  • PostgreSQL (transactional data)                                 │
│  • pgvector (embeddings)                                          │
│  • DuckDB (analytics queries)                                      │
│  • S3 / MinIO (Parquet files สำหรับ historical)                    │
└──────────────────────────────────────────────────────────────────┘

# Pattern: เขียน Rust, เรียกจาก Python — ได้ performance โดยไม่ต้องเปลี่ยน stack
# ตัวอย่าง: Polars (DataFrame) เป็น Rust ที่มี Python API''', language="text")

# ============================================================
# บทที่ 18: สรุป
# ============================================================
add_heading_h1(doc, "สรุปและคำแนะนำสำหรับ Sahana Eden", num=18)

add_heading_h2(doc, "18.1  Final Stack Recommendation")

add_table(doc,
    headers=["Layer", "ภาษา / Framework", "เหตุผล"],
    rows=[
        ["AI / ML",                     "Python (FastAPI + LangChain + PyTorch)",   "Ecosystem monopoly — ไม่มีทางเลือกอื่น"],
        ["Web Backend (CRUD)",          "Python (FastAPI + SQLAlchemy 2.0)",         "Async, type-safe, OpenAPI, migration ง่าย"],
        ["API Gateway / Reverse Proxy", "Go (Caddy) หรือ Rust (Traefik)",             "High throughput, low memory"],
        ["Frontend (ส่วนใหญ่)",          "TypeScript + HTMX + Alpine.js + Tailwind",  "Server-rendered, progressive, small bundle"],
        ["Frontend (complex widgets)",  "TypeScript + React",                        "Map editor, dashboards, charts"],
        ["Mobile App",                  "TypeScript + React Native",                  "Share logic กับ web, offline-first"],
        ["CLI Tools",                   "Go",                                         "Single binary, fast startup, easy distribution"],
        ["Performance Layer",           "Rust (via PyO3)",                            "Vector search, image processing, crypto"],
        ["Background Jobs",             "Python (Celery) + Redis",                    "AI ecosystem integration"],
        ["Database",                    "PostgreSQL 16 + pgvector",                   "Production-grade, vector search"],
        ["Cache / Pub-Sub",             "Redis 7",                                    "Sessions, rate limit, job queue"],
        ["DevOps Tooling",              "Python + Go (mixed)",                        "Python สำหรับ scripts, Go สำหรับ CLI"],
    ],
    col_widths_cm=[3.5, 4.5, 8.0],
    first_col_bold=True)

add_heading_h2(doc, "18.2  Migration Strategy Summary")

add_para(doc,
    "Sahana Eden ควร migrate จาก web2py → FastAPI ทีละส่วน (Strangler Pattern) "
    "ภายใน 6 เดือน — เริ่มจาก AI modules (ง่ายสุด) ตามด้วย auth/org/inv และ decommission web2py ในท้ายที่สุด "
    "ไม่ต้องเปลี่ยนภาษา — ยังเป็น Python แต่เปลี่ยน framework และ libraries "
    "ส่วน frontend/mobile ใช้ TypeScript, performance-critical ใช้ Rust (ผ่าน PyO3)",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "18.3  Key Takeaways")

add_bullet(doc, "Python ยังเหมาะกับ AI/ML, data science, web backend ของ Eden (ถูกแล้ว)", bold_prefix="✓ ")
add_bullet(doc, "web2py → FastAPI migration ทีละส่วนด้วย Strangler Pattern (6 เดือน)", bold_prefix="② ")
add_bullet(doc, "Frontend ใช้ TypeScript (HTMX 90% + React 10% สำหรับ complex)", bold_prefix="③ ")
add_bullet(doc, "Mobile app ใช้ React Native + TypeScript (Phase 4)", bold_prefix="④ ")
add_bullet(doc, "Performance-critical ใช้ Rust ผ่าน PyO3 (ไม่ต้องเปลี่ยน Python)", bold_prefix="⑤ ")
add_bullet(doc, "API Gateway / CLI ใช้ Go (single binary, fast startup)", bold_prefix="⑥ ")
add_bullet(doc, "Database ใช้ PostgreSQL 16 + pgvector (แทน SQLite)", bold_prefix="⑦ ")
add_bullet(doc, "อย่าลืม ADRs + runbooks + tests — not just code migration", bold_prefix="⑧ ")

add_heading_h2(doc, "18.4  ไฟล์เอกสารที่เกี่ยวข้อง")

add_table(doc,
    headers=["ไฟล์", "เนื้อหา"],
    rows=[
        ["Sahana_Eden_Migration_And_Language_Comparison.docx", "เอกสารฉบับนี้ (Part A + Part B)"],
        ["Sahana_Eden_Migration_And_Language_Comparison.pdf",  "PDF version"],
        ["Sahana_Eden_Modernization_Recommendations.docx",     "Modernization roadmap (10 หมวด)"],
        ["Sahana_Eden_skpuan07_Repo_Analysis.docx",             "วิเคราะห์โครงสร้าง repo"],
        ["Sahana_Eden_S3_Framework_Architecture.docx",          "S3 framework deep dive"],
        ["Sahana_Eden_Installation_README.md",                   "คู่มือติดตั้ง"],
    ],
    col_widths_cm=[7.0, 9.0],
    first_col_bold=True)

# Footer
section = doc.sections[0]
footer = section.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fp.text = ""
add_run(fp, "Sahana Eden — Migration & Language Comparison  ·  Page ",
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
