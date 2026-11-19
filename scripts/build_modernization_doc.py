"""
Generate .docx summarizing modernization recommendations for Sahana Eden fork.
Style: Technical clean (consistent with previous docs).
Includes: 10 categories, code examples, priority matrix.
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# Configuration
# ============================================================
OUT_PATH = "/home/z/my-project/download/Sahana_Eden_Modernization_Recommendations.docx"

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
add_run(p, "MODERNIZATION ROADMAP", font_name=FONT_HEADING,
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
add_run(p, "คำแนะนำการปรับปรุงระบบให้ทันสมัย",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
add_run(p, "Modernization, Dev Experience & Deployment Recommendations",
        font_name=FONT_BODY, size_pt=13, italic=True, color=C_AI)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(40)
add_run(p, "10 หมวด · Code Examples · Priority Matrix (5 Phases)",
        font_name=FONT_BODY, size_pt=11, color=C_SECOND)

info = doc.add_table(rows=6, cols=2)
info.alignment = WD_TABLE_ALIGNMENT.CENTER
info.autofit = False
info_data = [
    ("Repository",     "github.com/skpuan07-cmd/eden"),
    ("Base system",    "Sahana Eden v6.2 + AI modules (10 use cases)"),
    ("Stack ปัจจุบัน",  "web2py 3.3.3 + SQLite + jQuery + Foundation 5"),
    ("เป้าหมาย",        "Production-ready, maintainable, deployable ภายใน 6 เดือน"),
    ("ผู้อ่าน",         "Developer / DevOps / Tech Lead"),
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
add_run(p, "เอกสารฉบับนี้เสริมกับ docs/ai/RECOMMENDATIONS.md ที่มีอยู่ (AI/Cyber/PDPA) "
           "โดยเน้น modernization ของ stack, dev experience และ deployment",
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
    ("บทที่ 1",  "ภาพรวมและ Priority Matrix"),
    ("บทที่ 2",  "Containerization & Infrastructure as Code"),
    ("บทที่ 3",  "CI/CD Pipeline"),
    ("บทที่ 4",  "Database Modernization"),
    ("บทที่ 5",  "Frontend Modernization"),
    ("บทที่ 6",  "Code Quality & Type Safety"),
    ("บทที่ 7",  "Observability & Monitoring"),
    ("บทที่ 8",  "Async I/O และ Streaming"),
    ("บทที่ 9",  "Security Hardening (เพิ่มจากที่มี)"),
    ("บทที่ 10", "API Modernization"),
    ("บทที่ 11", "Documentation & Developer Onboarding"),
    ("บทที่ 12", "สรุป Priority Matrix และ Roadmap"),
]
for num, title in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(15.5), WD_ALIGN_PARAGRAPH.RIGHT, 2)
    add_run(p, num + "  ", font_name=FONT_HEADING, size_pt=11, bold=True, color=C_PRIMARY)
    add_run(p, title, font_name=FONT_BODY, size_pt=11, color=C_DARK)

# ============================================================
# บทที่ 1: ภาพรวม
# ============================================================
add_heading_h1(doc, "ภาพรวมและ Priority Matrix", num=1)

add_para(doc,
    "เอกสารฉบับนี้เสนอคำแนะนำการปรับปรุง Sahana Eden fork (skpuan07-cmd/eden) "
    "ให้ทันสมัย สะดวกต่อการพัฒนา และพร้อม deploy ใน production "
    "เนื้อหาแบ่งเป็น 10 หมวด ครอบคลุม containerization, CI/CD, database, "
    "frontend, code quality, observability, async I/O, security, API และ documentation "
    "แต่ละหมวดมี code examples ที่ใช้งานได้จริงและ priority matrix แบ่งเป็น 5 phases",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_note_box(doc, "เอกสารที่เกี่ยวข้อง",
    "เอกสารฉบับนี้เสริมกับ docs/ai/RECOMMENDATIONS.md (977 บรรทัด) ที่มีอยู่ใน repo "
    "ซึ่งเน้น AI, Cyber Security และ PDPA — เอกสารนี้เน้ม modernization ของ stack "
    "และ deployment แทน เพื่อให้ครอบคลุมทุกมิติของการพัฒนาระบบ",
    color="0E7490", bg="ECFDF5")

add_heading_h2(doc, "1.1  สถานะปัจจุบันของระบบ")

add_table(doc,
    headers=["ส่วน", "ปัจจุบัน", "เป้าหมายหลัง modernize"],
    rows=[
        ["Deployment",     "Manual (git clone + symlink + restart)", "Docker Compose / Kubernetes"],
        ["Database",       "SQLite (file-based)",                     "PostgreSQL 16 + pgvector"],
        ["Frontend",       "jQuery + Foundation 5 + ExtJS + OpenLayers","HTMX + Tailwind + MapLibre GL"],
        ["AI engine",      "subprocess.run(z-ai CLI) — sync/block",   "Async aiohttp + streaming SSE"],
        ["Tests",          "ไม่มี (เฉพาะ framework tests)",            "pytest + 80% coverage สำหรับ AI"],
        ["CI/CD",          "Manual PR merge",                        "GitHub Actions: lint+test+build+deploy"],
        ["Logging",        "web2py debug toolbar",                    "structlog + OpenTelemetry + Sentry"],
        ["API docs",       "ไม่มี",                                    "OpenAPI 3.0 + Swagger UI"],
        ["Secrets",        "/etc/eden/env (manual)",                  "Vault / AWS Secrets Manager"],
        ["Observability",  "Error tickets ใน errors/",                "Prometheus + Grafana + Sentry"],
    ],
    col_widths_cm=[2.5, 5.5, 8.0],
    first_col_bold=True)

add_heading_h2(doc, "1.2  Priority Matrix ภาพรวม (5 Phases)")

add_table(doc,
    headers=["Phase", "ระยะเวลา", "เป้าหมาย", "งานหลัก"],
    rows=[
        ["Phase 1 — Quick Wins",    "1-2 สัปดาห์",  "Setup ใหม่ใช้เวลา < 5 นาที, มี CI",  "Dockerfile, docker-compose, pre-commit, GitHub Actions, type hints, health endpoint"],
        ["Phase 2 — Foundation",    "1 เดือน",      "Production-grade infrastructure",   "PostgreSQL + pgvector, test suite 80%, Alembic migrations, Sentry + Prometheus, OpenAPI spec"],
        ["Phase 3 — Modernization", "2-3 เดือน",    "Modern stack, async I/O",           "Async AI client (aiohttp), streaming chat (SSE), Celery + Redis, HTMX, MapLibre GL"],
        ["Phase 4 — Scale",         "3-6 เดือน",    "Kubernetes + mobile",               "Helm chart, K8s workers, CDN, PWA offline, React Native mobile app"],
        ["Phase 5 — Compliance",    "ongoing",      "Security audit + PDPA",             "2FA admin, column encryption, WAF + CSP, penetration testing, DPIA"],
    ],
    col_widths_cm=[3.5, 2.5, 4.0, 6.0],
    first_col_bold=True, font_size=9, header_size=9.5)

# ============================================================
# บทที่ 2: Containerization
# ============================================================
add_heading_h1(doc, "Containerization & Infrastructure as Code", num=2)

add_heading_h2(doc, "2.1  สถานะปัจจุบัน")
add_bullet(doc, "ไม่มี Dockerfile ใน repo (มีเพียง .dockerignore)")
add_bullet(doc, "ติดตั้ง manual ผ่าน git clone + git reset --hard 9005306 + symlink + คัดลอก config")
add_bullet(doc, "Caddy reverse proxy config อยู่นอก repo")
add_bullet(doc, "Setup environment ใหม่ใช้เวลา 30+ นาที และ error-prone")

add_heading_h2(doc, "2.2  Dockerfile (Multi-stage Build)")

add_code_block(doc, '''# Dockerfile — multi-stage build
# ใช้ขนาด image เล็ก (~250 MB) และ build เร็ว

FROM python:3.12-slim AS base
RUN apt-get update && apt-get install -y --no-install-recommends \\
    libxml2-dev libxslt1-dev libldap2-dev libsasl2-dev libssl-dev \\
    libjpeg-dev zlib1g-dev libpq-dev gcc git && \\
    rm -rf /var/lib/apt/lists/*

# Stage 1: build web2py (cached layer)
FROM base AS web2py-builder
WORKDIR /build
RUN git clone https://github.com/web2py/web2py.git && \\
    cd web2py && \\
    git reset --hard 9005306 && \\
    git submodule update --init --recursive

# Stage 2: runtime image
FROM base AS runtime
WORKDIR /app

# Copy web2py framework
COPY --from=web2py-builder /build/web2py /app/web2py

# Copy Eden application (from current directory = repo root)
COPY . /app/web2py/applications/eden

# Install Python dependencies
RUN pip install --no-cache-dir \\
    -r /app/web2py/applications/eden/requirements.txt \\
    -r /app/web2py/applications/eden/optional_requirements.txt \\
    prophet pandas openpyxl

# Create config from template (if not exists)
RUN cp -n /app/web2py/applications/eden/modules/templates/000_config.py \\
          /app/web2py/applications/eden/models/000_config.py && \\
    sed -i 's/FINISHED_EDITING_CONFIG_FILE = False/\\
            FINISHED_EDITING_CONFIG_FILE = True/' \\
            /app/web2py/applications/eden/models/000_config.py

# Create uploads directory
RUN mkdir -p /app/web2py/applications/eden/uploads/ai_voice \\
    && chown -R www-data:www-data /app/web2py/applications/eden/uploads

EXPOSE 8000
USER www-data

ENV EDEN_PORT=8000 \\
    EDEN_IP=0.0.0.0 \\
    PYTHONUNBUFFERED=1 \\
    PYTHONDONTWRITEBYTECODE=1

CMD ["python", "web2py.py", "--no_gui", \\
     "-a", "${EDEN_ADMIN_PW}", \\
     "--ip", "0.0.0.0", "--port", "8000"]''', language="dockerfile")

add_heading_h2(doc, "2.3  docker-compose.yml (Dev + Production)")

add_code_block(doc, '''# docker-compose.yml
# Usage:
#   Development: docker compose up
#   Production:  docker compose -f docker-compose.prod.yml up -d

services:
  eden:
    build: .
    ports:
      - "8000:8000"
    env_file: .env
    depends_on:
      postgres:
        condition: service_healthy
      redis:
        condition: service_started
    volumes:
      - eden_uploads:/app/web2py/applications/eden/uploads
      - eden_errors:/app/web2py/applications/eden/errors
      - eden_databases:/app/web2py/applications/eden/databases
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:8000/eden/health"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 60s

  postgres:
    image: postgres:16-alpine
    environment:
      POSTGRES_DB: eden
      POSTGRES_USER: eden
      POSTGRES_PASSWORD: ${POSTGRES_PASSWORD}
      POSTGRES_INITDB_ARGS: "--encoding=UTF-8"
    volumes:
      - pg_data:/var/lib/postgresql/data
      - ./scripts/init-pgvector.sql:/docker-entrypoint-initdb.d/init-pgvector.sql
    healthcheck:
      test: ["CMD-SHELL", "pg_isready -U eden"]
      interval: 10s
      timeout: 5s
      retries: 5
    restart: unless-stopped

  redis:
    image: redis:7-alpine
    command: redis-server --appendonly yes --maxmemory 256mb --maxmemory-policy allkeys-lru
    volumes:
      - redis_data:/data
    restart: unless-stopped

  caddy:
    image: caddy:2-alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./Caddyfile:/etc/caddy/Caddyfile:ro
      - caddy_data:/data
      - caddy_config:/config
    depends_on: [eden]
    restart: unless-stopped

volumes:
  eden_uploads:
  eden_errors:
  eden_databases:
  pg_data:
  redis_data:
  caddy_data:
  caddy_config:''', language="yaml")

add_heading_h2(doc, "2.4  .env File (Template)")

add_code_block(doc, '''# .env (DO NOT commit — add to .gitignore)
# Copy from .env.example and fill in real values

# Database
POSTGRES_HOST=postgres
POSTGRES_DB=eden
POSTGRES_USER=eden
POSTGRES_PASSWORD=change_me_to_strong_password

# web2py admin
EDEN_ADMIN_PW=change_me_to_strong_password

# AI APIs
IAPP_API_KEY=your_openThai_key_here
ZAI_API_KEY=your_z_ai_key_here

# Optional: Sentry, Redis, etc.
SENTRY_DSN=
REDIS_URL=redis://redis:6379/0

# Encryption key for sensitive columns (generate: python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())")
EDEN_ENCRYPTION_KEY=''', language="bash")

add_heading_h2(doc, "2.5  Caddyfile (Reverse Proxy + HTTPS)")

add_code_block(doc, '''# Caddyfile — automatic HTTPS with Let's Encrypt

{$EDEN_DOMAIN} {
    # Reverse proxy to web2py
    reverse_proxy eden:8000 {
        header_up X-Real-IP {remote_host}
        header_up X-Forwarded-For {remote_host}
        header_up X-Forwarded-Proto {scheme}
    }

    # Security headers
    header {
        Content-Security-Policy "default-src 'self'; script-src 'self' 'unsafe-inline' cdnjs.cloudflare.com; img-src 'self' data: https:; style-src 'self' 'unsafe-inline' cdnjs.cloudflare.com; font-src 'self' cdnjs.cloudflare.com; connect-src 'self'; frame-ancestors 'none';"
        X-Frame-Options "DENY"
        X-Content-Type-Options "nosniff"
        Referrer-Policy "strict-origin-when-cross-origin"
        Strict-Transport-Security "max-age=31536000; includeSubDomains; preload"
        Permissions-Policy "geolocation=(), microphone=(), camera=()"
    }

    # Compression
    encode gzip zstd

    # Rate limiting (chatbot endpoint)
    @chatbot path /eden/ai_chat/ask
    rate_limit @chatbot 60r/m

    # Logging
    log {
        output file /data/access.log
        format json
    }
}

# HTTP → HTTPS redirect
http://{$EDEN_DOMAIN} {
    redir https://{$EDEN_DOMAIN}{uri} permanent
}''', language="text")

add_heading_h2(doc, "2.6  Kubernetes Helm Chart (Phase 4)")

add_para(doc,
    "สำหรับ production scale แนะนำ Helm chart ที่ deploy ไปยัง Kubernetes cluster "
    "โดยมี features: HorizontalPodAutoscaler, PodDisruptionBudget, NetworkPolicy, "
    "PV/PVC สำหรับ uploads/databases, และ readiness/liveness probes",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''# helm/eden/values.yaml
replicaCount: 3
image:
  repository: ghcr.io/skpuan07-cmd/eden
  tag: latest
  pullPolicy: IfNotPresent

resources:
  requests: { cpu: 200m, memory: 512Mi }
  limits:   { cpu: 1,    memory: 1Gi }

autoscaling:
  enabled: true
  minReplicas: 3
  maxReplicas: 10
  targetCPUUtilizationPercentage: 70

postgresql:
  enabled: true
  primary:
    persistence:
      size: 50Gi
      storageClass: fast-ssd

redis:
  enabled: true
  architecture: standalone
  master:
    persistence:
      size: 5Gi

ingress:
  enabled: true
  className: nginx
  annotations:
    cert-manager.io/cluster-issuer: letsencrypt-prod
  hosts:
    - host: eden.example.com
      paths:
        - path: /
          pathType: Prefix
  tls:
    - secretName: eden-tls
      hosts: [eden.example.com]''', language="yaml")

# ============================================================
# บทที่ 3: CI/CD
# ============================================================
add_heading_h1(doc, "CI/CD Pipeline", num=3)

add_heading_h2(doc, "3.1  GitHub Actions Workflow")

add_code_block(doc, '''# .github/workflows/ci.yml
name: CI/CD

on:
  push:
    branches: [main, develop]
  pull_request:
    branches: [main]

env:
  PYTHON_VERSION: "3.12"
  NODE_VERSION: "20"

jobs:
  # ─────────────────────────────────────────────────────────
  lint:
    name: Lint & Type Check
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "${{ env.PYTHON_VERSION }}" }
      - run: pip install ruff black mypy
      - name: Ruff (lint)
        run: ruff check controllers/ modules/s3/ models/ tests/
      - name: Black (format check)
        run: black --check controllers/ modules/s3/ models/ tests/
      - name: Mypy (type check)
        run: mypy modules/s3/ai_client.py modules/s3/ai_ocr_openthai.py \\
                   controllers/ai_*.py

  # ─────────────────────────────────────────────────────────
  test:
    name: Test Suite
    runs-on: ubuntu-latest
    needs: lint
    services:
      postgres:
        image: postgres:16-alpine
        env:
          POSTGRES_DB: eden_test
          POSTGRES_PASSWORD: test
        ports: ["5432:5432"]
        options: >-
          --health-cmd pg_isready
          --health-interval 10s
          --health-timeout 5s
          --health-retries 5
      redis:
        image: redis:7-alpine
        ports: ["6379:6379"]
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "${{ env.PYTHON_VERSION }}" }
      - name: Install deps
        run: |
          pip install -r requirements.txt \\
                      -r optional_requirements.txt \\
                      pytest pytest-cov pytest-asyncio \\
                      factory-boy responses
      - name: Run tests
        env:
          EDEN_TEST_DB: postgres://postgres:test@localhost/eden_test
          REDIS_URL: redis://localhost:6379/0
        run: pytest tests/ --cov=modules/s3 --cov-report=xml --cov-report=html
      - name: Upload coverage
        uses: codecov/codecov-action@v4
        with: { files: ./coverage.xml }

  # ─────────────────────────────────────────────────────────
  security:
    name: Security Scan
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: snyk/actions/python@master
        with:
          command: test
          args: --severity-threshold=high
        env:
          SNYK_TOKEN: ${{ secrets.SNYK_TOKEN }}
      - uses: github/codeql-action/init@v3
        with:
          languages: python
      - uses: github/codeql-action/analyze@v3
      - name: pip-audit
        run: |
          pip install pip-audit
          pip-audit -r requirements.txt -r optional_requirements.txt

  # ─────────────────────────────────────────────────────────
  build:
    name: Build Docker Image
    needs: [lint, test, security]
    runs-on: ubuntu-latest
    permissions:
      contents: read
      packages: write
    steps:
      - uses: actions/checkout@v4
      - uses: docker/setup-buildx-action@v3
      - uses: docker/login-action@v3
        with:
          registry: ghcr.io
          username: ${{ github.actor }}
          password: ${{ secrets.GITHUB_TOKEN }}
      - uses: docker/build-push-action@v5
        with:
          context: .
          push: true
          tags: |
            ghcr.io/${{ github.repository }}:${{ github.sha }}
            ghcr.io/${{ github.repository }}:latest
          cache-from: type=gha
          cache-to: type=gha,mode=max

  # ─────────────────────────────────────────────────────────
  deploy-staging:
    name: Deploy to Staging
    if: github.ref == 'refs/heads/develop'
    needs: build
    runs-on: ubuntu-latest
    environment: staging
    steps:
      - uses: actions/checkout@v4
      - name: Deploy via SSH
        uses: appleboy/ssh-action@v1
        with:
          host: ${{ secrets.STAGING_HOST }}
          username: ${{ secrets.STAGING_USER }}
          key: ${{ secrets.STAGING_SSH_KEY }}
          script: |
            cd /opt/eden
            docker compose pull
            docker compose up -d --no-deps eden
            docker image prune -f

  # ─────────────────────────────────────────────────────────
  deploy-production:
    name: Deploy to Production
    if: github.ref == 'refs/heads/main'
    needs: build
    runs-on: ubuntu-latest
    environment: production
    steps:
      - uses: actions/checkout@v4
      - name: Deploy via Helm
        run: |
          helm upgrade --install eden ./helm/eden \\
            --namespace eden --create-namespace \\
            --set image.tag=${{ github.sha }} \\
            --values ./helm/eden/values.prod.yaml
        env:
          KUBECONFIG: ${{ secrets.KUBECONFIG }}''', language="yaml")

add_heading_h2(doc, "3.2  Pre-commit Hooks")

add_code_block(doc, '''# .pre-commit-config.yaml
# Install: pip install pre-commit && pre-commit install

repos:
  # Python linting & formatting
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.5.0
    hooks:
      - id: ruff
        args: [--fix]
      - id: ruff-format

  # Type checking
  - repo: https://github.com/pre-commit/mirrors-mypy
    rev: v1.10.0
    hooks:
      - id: mypy
        additional_dependencies: [types-requests, types-PyYAML]
        files: ^modules/s3/ai_

  # Secret detection
  - repo: https://github.com/Yelp/detect-secrets
    rev: v1.5.0
    hooks:
      - id: detect-secrets
        args: [--baseline, .secrets.baseline]

  # General checks
  - repo: https://github.com/pre-commit/pre-commit-hooks
    rev: v4.6.0
    hooks:
      - id: trailing-whitespace
      - id: end-of-file-fixer
      - id: check-yaml
      - id: check-merge-conflict
      - id: check-added-large-files
        args: [--maxkb=500]
      - id: check-case-conflict
      - id: requirements-txt-fixer

  # Dockerfile linting
  - repo: https://github.com/hadolint/hadolint
    rev: v2.12.0
    hooks:
      - id: hadolint''', language="yaml")

add_heading_h2(doc, "3.3  pyproject.toml (Centralized Config)")

add_code_block(doc, '''# pyproject.toml
[project]
name = "sahana-eden"
version = "6.2"
requires-python = ">=3.11"

[tool.ruff]
line-length = 100
target-version = "py311"
extend-exclude = ["modules/s3db", "static", "languages"]

[tool.ruff.lint]
select = [
    "E",   # pycodestyle errors
    "W",   # pycodestyle warnings
    "F",   # pyflakes
    "I",   # isort
    "B",   # flake8-bugbear
    "C4",  # flake8-comprehensions
    "UP",  # pyupgrade
    "RUF", # ruff-specific
]
ignore = ["E501"]  # handled by formatter

[tool.ruff.lint.per-file-ignores]
"tests/*" = ["B011"]  # assert False in tests is OK

[tool.black]
line-length = 100
target-version = ["py311"]

[tool.mypy]
python_version = "3.11"
strict_optional = true
ignore_missing_imports = true
warn_redundant_casts = true
warn_unused_ignores = true
exclude = "modules/s3db/|modules/core/"

[tool.pytest.ini_options]
minversion = "7.0"
testpaths = ["tests"]
addopts = [
    "--cov=modules/s3",
    "--cov-report=html",
    "--cov-report=term-missing",
    "--cov-fail-under=80",
    "--strict-markers",
]
markers = [
    "slow: marks tests as slow",
    "integration: requires external services",
    "ai: tests AI modules (may incur API costs)",
]''', language="toml")

# ============================================================
# บทที่ 4: Database
# ============================================================
add_heading_h1(doc, "Database Modernization", num=4)

add_heading_h2(doc, "4.1  สถานะปัจจุบัน vs เป้าหมาย")

add_table(doc,
    headers=["Aspect", "ปัจจุบัน (SQLite)", "เป้าหมาย (PostgreSQL 16)"],
    rows=[
        ["Concurrency",      "Single writer",                    "Multi-writer + MVCC"],
        ["Vector search",    "ไม่ได้ (placeholder JSON)",          "pgvector + ivfflat index"],
        ["Backup",           "Copy .db file",                     "pg_dump + point-in-time recovery"],
        ["Replication",      "ไม่ได้",                              "Streaming replication + read replicas"],
        ["FTS",              "LIKE query",                        "pg_trgm + full-text search"],
        ["JSON columns",     "Text field",                        "JSONB with indexing"],
        ["Size limit",       "~140 GB practical",                 "Virtually unlimited"],
        ["Production-ready", "ไม่",                                "ใช่"],
    ],
    col_widths_cm=[3.0, 5.5, 7.5],
    first_col_bold=True)

add_heading_h2(doc, "4.2  PostgreSQL Configuration")

add_code_block(doc, '''# models/000_config.py (modified for PostgreSQL)
import os

# Database settings — PostgreSQL for production
settings.database.db_type = "postgres"
settings.database.host = os.environ.get("POSTGRES_HOST", "localhost")
settings.database.port = int(os.environ.get("POSTGRES_PORT", 5432))
settings.database.database = os.environ.get("POSTGRES_DB", "eden")
settings.database.username = os.environ.get("POSTGRES_USER", "eden")
settings.database.password = os.environ.get("POSTGRES_PASSWORD", "")

# Connection pooling
settings.database.pool_size = 10
settings.database.max_connections = 100

# Enable migrate only in dev
if os.environ.get("EDEN_ENV") == "production":
    settings.base.migrate = False
    settings.base.fake_migrate = False''', language="python")

add_heading_h2(doc, "4.3  pgvector Setup (สำหรับ RAG จริง)")

add_code_block(doc, '''-- scripts/init-pgvector.sql
-- Run once on PostgreSQL 16+

-- Enable pgvector extension
CREATE EXTENSION IF NOT EXISTS vector;

-- Enable trigram (for fuzzy text search)
CREATE EXTENSION IF NOT EXISTS pg_trgm;

-- Enable unaccent (for accent-insensitive search)
CREATE EXTENSION IF NOT EXISTS unaccent;

-- Migrate ai_knowledge_chunk.embedding from JSON text to vector
ALTER TABLE ai_knowledge_chunk 
  ADD COLUMN embedding_vec vector(1536);

-- Create HNSW index (faster than ivfflat for cosine similarity)
CREATE INDEX idx_knowledge_embedding 
  ON ai_knowledge_chunk 
  USING hnsw (embedding_vec vector_cosine_ops)
  WITH (m = 16, ef_construction = 64);

-- GIN index for fast text search on tags
CREATE INDEX idx_knowledge_tags 
  ON ai_knowledge_chunk 
  USING GIN (tags);

-- Updated_at trigger
CREATE OR REPLACE FUNCTION update_modified_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.modified_on = NOW();
    RETURN NEW;
END;
$$ language 'plpgsql';

CREATE TRIGGER update_ai_knowledge_modified
  BEFORE UPDATE ON ai_knowledge_chunk
  FOR EACH ROW EXECUTE FUNCTION update_modified_column();''', language="sql")

add_heading_h2(doc, "4.4  Vector Search Query (แทน TF-IDF)")

add_code_block(doc, '''# modules/s3/ai_rag.py (NEW — vector search implementation)
import asyncpg
import os
from typing import List, Dict

class VectorStore:
    """pgvector-backed vector store for RAG."""

    def __init__(self):
        self.conn = None

    async def connect(self):
        self.conn = await asyncpg.connect(
            host=os.environ["POSTGRES_HOST"],
            database=os.environ["POSTGRES_DB"],
            user=os.environ["POSTGRES_USER"],
            password=os.environ["POSTGRES_PASSWORD"],
        )

    async def search(
        self,
        query_embedding: List[float],
        top_k: int = 5,
        threshold: float = 0.7,
    ) -> List[Dict]:
        """Find similar chunks using cosine similarity."""
        # pgvector <=> operator = cosine distance
        rows = await self.conn.fetch("""
            SELECT 
                id, title, content, source_type, source_id, source_url,
                1 - (embedding_vec <=> $1::vector) AS similarity
            FROM ai_knowledge_chunk
            WHERE 
                deleted = false 
                AND is_public = true
                AND 1 - (embedding_vec <=> $1::vector) > $2
            ORDER BY embedding_vec <=> $1::vector
            LIMIT $3
        """, str(query_embedding), threshold, top_k)
        
        return [dict(r) for r in rows]

    async def upsert_chunk(
        self, chunk_id: int, embedding: List[float]
    ):
        """Update embedding for an existing chunk."""
        await self.conn.execute("""
            UPDATE ai_knowledge_chunk
            SET embedding_vec = $1::vector,
                last_indexed = NOW()
            WHERE id = $2
        """, str(embedding), chunk_id)''', language="python")

add_heading_h2(doc, "4.5  Alembic Migrations")

add_code_block(doc, '''# migrations/env.py
from alembic import context
from sqlalchemy import engine_from_config, pool
import os

config = context.config

# Override URL from env
config.set_main_option(
    "sqlalchemy.url",
    f"postgresql+psycopg2://{os.environ['POSTGRES_USER']}:"
    f"{os.environ['POSTGRES_PASSWORD']}@"
    f"{os.environ['POSTGRES_HOST']}/"
    f"{os.environ['POSTGRES_DB']}"
)

def run_migrations_offline():
    context.configure(url=config.get_main_option("sqlalchemy.url"))
    with context.begin_transaction():
        context.run_migrations()

def run_migrations_online():
    connectable = engine_from_config(
        config.get_section(config.config_ini_section),
        prefix="sqlalchemy.",
        poolclass=pool.NullPool,
    )
    with connectable.connect() as connection:
        context.configure(connection=connection)
        with context.begin_transaction():
            context.run_migrations()

if context.is_offline_mode():
    run_migrations_offline()
else:
    run_migrations_online()

# Usage:
# alembic revision --autogenerate -m "add ai tables"
# alembic upgrade head
# alembic downgrade -1''', language="python")

add_heading_h2(doc, "4.6  Automated Backup Strategy")

add_code_block(doc, '''#!/bin/bash
# scripts/backup-postgres.sh
# Cron: 0 * * * * /opt/eden/scripts/backup-postgres.sh

set -e
BACKUP_DIR=/backups/postgres
RETENTION_HOURLY=24    # keep 24 hourly backups
RETENTION_DAILY=7      # keep 7 daily
RETENTION_WEEKLY=4     # keep 4 weekly

mkdir -p $BACKUP_DIR/{hourly,daily,weekly}

TIMESTAMP=$(date +%Y%m%d-%H%M%S)
DAILY_TS=$(date +%Y%m%d)
WEEKLY_TS=$(date +%Y%U)

# Hourly backup
pg_dump -h $POSTGRES_HOST -U $POSTGRES_USER $POSTGRES_DB | \\
  gzip > $BACKUP_DIR/hourly/eden-$TIMESTAMP.sql.gz

# Daily snapshot (at 02:00)
if [ $(date +%H) -eq "02" ]; then
    cp $BACKUP_DIR/hourly/eden-$TIMESTAMP.sql.gz \\
       $BACKUP_DIR/daily/eden-$DAILY_TS.sql.gz
fi

# Weekly snapshot (Sundays)
if [ $(date +%u) -eq "7" ] && [ $(date +%H) -eq "03" ]; then
    cp $BACKUP_DIR/hourly/eden-$TIMESTAMP.sql.gz \\
       $BACKUP_DIR/weekly/eden-$WEEKLY_TS.sql.gz
fi

# Cleanup
find $BACKUP_DIR/hourly -mtime +$RETENTION_HOURLY -delete
find $BACKUP_DIR/daily  -mtime +$RETENTION_DAILY  -delete
find $BACKUP_DIR/weekly -mtime +$((RETENTION_WEEKLY*7)) -delete

# Upload to S3 (optional)
aws s3 sync $BACKUP_DIR s3://eden-backups/ --delete \\
  --exclude "*" --include "*.sql.gz"

echo "[$(date)] Backup complete: eden-$TIMESTAMP.sql.gz"''', language="bash")

# ============================================================
# บทที่ 5: Frontend
# ============================================================
add_heading_h1(doc, "Frontend Modernization", num=5)

add_heading_h2(doc, "5.1  สถานะปัจจุบัน")

add_table(doc,
    headers=["Library", "เวอร์ชั่นปัจจุบัน", "สถานะ", "ทางเลือกทันสมัย"],
    rows=[
        ["jQuery",       "1.x",   "Deprecated, EOL",     "HTMX + Alpine.js"],
        ["Foundation",   "5",     "Legacy, EOL",         "Tailwind CSS"],
        ["ExtJS",        "6",     "Commercial, heavy",   "AG Grid Community"],
        ["DataTables",   "1.x",   "Active, but legacy",  "TanStack Table"],
        ["OpenLayers",   "6.x",   "Active",              "MapLibre GL JS"],
        ["D3.js",        "7.x",   "Active",              "เก็บไว้ (ดีแล้ว)"],
        ["YATL",         "—",     "web2py built-in",     "เก็บไว้ (incremental migration)"],
    ],
    col_widths_cm=[2.5, 2.5, 3.5, 7.5],
    first_col_bold=True)

add_heading_h2(doc, "5.2  Incremental Migration Strategy")

add_para(doc,
    "ไม่แนะนำให้ rewrite ทั้งหมด — ใช้ strategy 'strangler pattern' ย้ายทีละ module "
    "จาก jQuery → HTMX + Alpine.js โดย Eden modules ที่ใช้ S3 framework อยู่แล้ว "
    "สามารถค่อย ๆ แทนที่ frontend ได้โดยไม่กระทบ backend",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "ตัวอย่าง: HTMX แทน AJAX ของ jQuery")

add_code_block(doc, '''<!-- OLD: jQuery AJAX (controllers/ai_chat.py view) -->
<button onclick="sendQuestion()">Send</button>
<script>
function sendQuestion() {
    $.ajax({
        url: "/eden/ai_chat/ask",
        method: "POST",
        data: { question: $("#chat-input").val() },
        success: function(response) {
            $("#chat-messages").append(
                "<div><strong>You:</strong> " + response.question + "</div>" +
                "<div><strong>Bot:</strong> " + response.answer + "</div>"
            );
        }
    });
}
</script>

<!-- NEW: HTMX (declarative, no JS) -->
<button hx-post="/eden/ai_chat/ask"
        hx-include="[name='question']"
        hx-target="#chat-messages"
        hx-swap="beforeend"
        hx-indicator="#loading">
    Send
</button>
<span id="loading" class="htmx-indicator">⏳</span>''', language="html")

add_heading_h3(doc, "ตัวอย่าง: Alpine.js แทน jQuery DOM")

add_code_block(doc, '''<!-- OLD: jQuery -->
<div id="chat">
    <input id="input" type="text" />
    <button onclick="send()">Send</button>
</div>
<script>
let messages = [];
function send() {
    messages.push({ text: $("#input").val(), role: "user" });
    render();
}
function render() {
    $("#chat").html(messages.map(m => `<p>${m.role}: ${m.text}</p>`).join(""));
}
</script>

<!-- NEW: Alpine.js (reactive, no manual render) -->
<div x-data="{ messages: [], input: '' }">
    <template x-for="m in messages" :key="m.id">
        <p x-text="m.role + ': ' + m.text"></p>
    </template>
    <input x-model="input" type="text" @keyup.enter="messages.push({ id: Date.now(), role: 'user', text: input }); input = ''" />
</div>''', language="html")

add_heading_h2(doc, "5.3  Tailwind CSS Setup")

add_code_block(doc, '''# package.json
{
  "name": "eden-frontend",
  "scripts": {
    "dev": "tailwindcss -i ./static/src/main.css -o ./static/dist/main.css --watch",
    "build": "tailwindcss -i ./static/src/main.css -o ./static/dist/main.css --minify"
  },
  "devDependencies": {
    "tailwindcss": "^3.4.0",
    "@tailwindcss/forms": "^0.5.7"
  }
}

# tailwind.config.js
module.exports = {
  content: [
    "./views/**/*.html",
    "./static/src/**/*.js",
  ],
  theme: {
    extend: {
      colors: {
        eden: {
          primary: "#1E3A8A",
          accent: "#0E7490",
          dark: "#0F172A",
        }
      }
    }
  },
  plugins: [require("@tailwindcss/forms")],
}''', language="json")

add_heading_h2(doc, "5.4  PWA (Progressive Web App) สำหรับ Offline Use")

add_para(doc,
    "Critical สำหรับ disaster management — อาสาสมัครในพื้นที่ที่อินเทอร์เน็ตไม่ stable "
    "PWA ทำให้ Eden ใช้งานได้ offline และ sync เมื่อกลับ online",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_code_block(doc, '''// static/sw.js — Service Worker
const CACHE_NAME = "eden-v1";
const STATIC_ASSETS = [
  "/eden/default/index",
  "/eden/static/dist/main.css",
  "/eden/static/dist/app.js",
];

// Install: cache critical assets
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS))
  );
});

// Activate: cleanup old caches
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
});

// Fetch: network-first, fallback to cache
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Cache successful responses
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => 
            cache.put(event.request, clone)
          );
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});

// Background Sync: queue mutations when offline
self.addEventListener("sync", (event) => {
  if (event.tag === "eden-sync") {
    event.waitUntil(syncPendingMutations());
  }
});''', language="javascript")

# ============================================================
# บทที่ 6: Code Quality
# ============================================================
add_heading_h1(doc, "Code Quality & Type Safety", num=6)

add_heading_h2(doc, "6.1  Type Hints สำหรับ AI Modules")

add_code_block(doc, '''# modules/s3/ai_client.py (with type hints)
from __future__ import annotations
from pathlib import Path
from typing import Optional, Literal, AsyncIterator
from dataclasses import dataclass

ModelName = Literal["glm-4-plus", "glm-5.3-flash", "glm-4v-plus"]
Language = Literal["th", "en", "vi", "km", "lo", "my", "id", "ms", "zh-cn", "ja", "ko"]

@dataclass
class ChatResponse:
    content: str
    model: str
    tokens_prompt: int
    tokens_completion: int
    latency_ms: int

@dataclass  
class OCRResult:
    extracted_data: dict[str, str]
    confidence: float
    engine: str  # "openthai" | "glm-vision"
    raw_response: dict

class AIClient:
    """Wrapper around the z-ai CLI for Sahana Eden."""

    def __init__(
        self,
        timeout: int = 120,
        cli_path: Optional[Path] = None,
    ) -> None:
        self.timeout = timeout
        self.cli_path = cli_path or self._find_cli()

    def chat(
        self,
        prompt: str,
        *,
        system: Optional[str] = None,
        model: ModelName = "glm-4-plus",
        max_tokens: int = 2000,
        temperature: float = 0.7,
    ) -> ChatResponse:
        """Send a chat completion request."""
        ...

    def describe_image(
        self,
        image_path: Path,
        *,
        prompt: str = "Describe this image in detail.",
        model: ModelName = "glm-4v-plus",
    ) -> str:
        """Analyze an image using vision model."""
        ...

    async def chat_stream(
        self,
        prompt: str,
        *,
        model: ModelName = "glm-4-plus",
    ) -> AsyncIterator[str]:
        """Stream chat completion tokens (for SSE endpoints)."""
        ...''', language="python")

add_heading_h2(doc, "6.2  Test Suite Structure")

add_code_block(doc, '''tests/
├── conftest.py                          # shared fixtures
├── unit/
│   ├── ai/
│   │   ├── test_ai_client.py            # AIClient unit tests
│   │   ├── test_ai_ocr_openthai.py      # OpenThai client tests
│   │   ├── test_ai_rag.py               # vector search tests
│   │   └── test_vector_store.py
│   └── ...
├── integration/
│   ├── ai/
│   │   ├── test_ai_chat_endpoint.py     # full HTTP flow
│   │   ├── test_ai_ocr_flow.py          # upload → extract → pre-fill
│   │   ├── test_ai_query_whitelist.py   # PDPA whitelist enforcement
│   │   ├── test_ai_moderation.py        # auto-moderation pipeline
│   │   └── test_ai_voice.py             # ASR + TTS
│   └── ...
└── fixtures/
    ├── sample_id_card.jpg               # Thai ID card sample
    ├── sample_invoice.jpg               # invoice sample
    ├── sample_satellite_image.tif       # flood imagery
    └── sample_audio_thai.wav            # Thai voice sample

# tests/conftest.py
import pytest
from unittest.mock import AsyncMock, patch
from pathlib import Path

@pytest.fixture
def mock_ai_client():
    """Mock AIClient to avoid real API calls."""
    with patch("modules.s3.ai_client.AIClient") as mock:
        client = mock.return_value
        client.chat.return_value = "Mocked AI response"
        client.describe_image.return_value = "Mocked image description"
        yield client

@pytest.fixture
def sample_id_card_path() -> Path:
    return Path(__file__).parent / "fixtures" / "sample_id_card.jpg"

@pytest.fixture
def db_session():
    """Rollback-per-test database session."""
    # ... use transactional fixtures
    yield session
    session.rollback()

# tests/integration/ai/test_ai_chat_endpoint.py
import pytest

@pytest.mark.asyncio
async def test_chat_requires_consent(client, mock_ai_client):
    """PDPA: chat must require consent."""
    response = await client.post("/eden/ai_chat/ask", 
                                 data={"question": "test"})
    assert response.status_code == 200
    data = response.json()
    assert data["error"] == "consent_required"

@pytest.mark.asyncio
async def test_chat_rate_limit(client, mock_ai_client):
    """Rate limit: 60 requests/minute/IP."""
    # Make 60 requests
    for i in range(60):
        response = await client.post("/eden/ai_chat/ask",
                                     data={"question": f"q{i}"},
                                     cookies={"session": "test"})
        assert response.status_code == 200
    
    # 61st should fail
    response = await client.post("/eden/ai_chat/ask",
                                 data={"question": "overflow"},
                                 cookies={"session": "test"})
    data = response.json()
    assert "rate limit" in data["error"]''', language="python")

add_heading_h2(doc, "6.3  Coverage Targets")

add_table(doc,
    headers=["Module", "Coverage Target", "หมายเหตุ"],
    rows=[
        ["modules/s3/ai_client.py",           "95%", "Core AI wrapper"],
        ["modules/s3/ai_ocr_openthai.py",     "85%", "External API client"],
        ["controllers/ai_chat.py",            "90%", "Public endpoint + PDPA"],
        ["controllers/ai_ocr.py",             "85%", "Sensitive data flow"],
        ["controllers/ai_query.py",           "95%", "PDPA whitelist critical"],
        ["controllers/ai_moderation.py",      "85%", "Auto-moderation logic"],
        ["controllers/ai_analytics.py",       "70%", "Prophet integration (slow tests)"],
        ["modules/s3db/ (Eden base)",         "60%", "Existing tests"],
        ["Overall",                            "80%", "Coverage gate in CI"],
    ],
    col_widths_cm=[5.5, 3.0, 7.5],
    first_col_bold=True)

# ============================================================
# บทที่ 7: Observability
# ============================================================
add_heading_h1(doc, "Observability & Monitoring", num=7)

add_heading_h2(doc, "7.1  Structured Logging (structlog)")

add_code_block(doc, '''# modules/s3/logging_config.py
import structlog
import logging
import sys

def configure_logging(env: str = "development"):
    """Configure structured logging for the application."""
    
    if env == "production":
        # JSON output for log aggregation (ELK, Loki, etc.)
        processors = [
            structlog.contextvars.merge_contextvars,
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.add_logger_context,
            structlog.processors.CallsiteParameterAdder(
                parameters=[
                    structlog.processors.CallsiteParameter.FILENAME,
                    structlog.processors.CallsiteParameter.FUNC_NAME,
                    structlog.processors.CallsiteParameter.LINENO,
                ]
            ),
            structlog.processors.StackInfoRenderer(),
            structlog.processors.format_exc_info,
            structlog.processors.JSONRenderer(),
        ]
    else:
        # Pretty console output for dev
        processors = [
            structlog.processors.TimeStamper(fmt="%H:%M:%S"),
            structlog.dev.ConsoleRenderer(colors=True),
        ]
    
    structlog.configure(
        processors=processors,
        wrapper_class=structlog.make_filtering_bound_logger(logging.INFO),
        context_class=dict,
        logger_factory=structlog.PrintLoggerFactory(),
        cache_logger_on_first_use=True,
    )

# Usage in AI controllers:
log = structlog.get_logger()

def ask():
    log.info("ai_chat_request",
             user_id=session.auth_user.id if session.auth else None,
             question_length=len(question),
             model="glm-4-plus",
             consent_given=consent_given)
    
    t0 = time.time()
    response = AIClient().chat(question)
    latency_ms = int((time.time() - t0) * 1000)
    
    log.info("ai_chat_response",
             latency_ms=latency_ms,
             tokens=response.tokens_completion,
             rating=None)''', language="python")

add_heading_h2(doc, "7.2  OpenTelemetry Instrumentation")

add_code_block(doc, '''# modules/s3/telemetry.py
from opentelemetry import trace
from opentelemetry.sdk.trace import TracerProvider
from opentelemetry.sdk.trace.export import BatchSpanProcessor
from opentelemetry.exporter.otlp.proto.grpc.trace_exporter import (
    OTLPSpanExporter,
)
from opentelemetry.instrumentation.web2py import Web2PyInstrumentor
from opentelemetry.instrumentation.psycopg2 import Psycopg2Instrumentor

def setup_telemetry(service_name: str = "eden", otlp_endpoint: str = "http://otel-collector:4317"):
    provider = TracerProvider()
    provider.add_span_processor(
        BatchSpanProcessor(
            OTLPSpanExporter(endpoint=otlp_endpoint)
        )
    )
    trace.set_tracer_provider(provider)
    
    # Auto-instrument web2py + PostgreSQL
    Web2PyInstrumentor().instrument()
    Psycopg2Instrumentor().instrument()

# Usage in AI controllers:
tracer = trace.get_tracer(__name__)

def ask(question):
    with tracer.start_as_current_span("ai_chat.ask") as span:
        span.set_attribute("ai.model", "glm-4-plus")
        span.set_attribute("ai.question.length", len(question))
        span.set_attribute("ai.consent_given", consent_given)
        
        # RAG retrieval
        with tracer.start_as_current_span("rag.search"):
            chunks = vector_store.search(embedding)
            span.set_attribute("rag.chunks_found", len(chunks))
        
        # LLM call
        with tracer.start_as_current_span("llm.chat"):
            response = AIClient().chat(question, context=chunks)
            span.set_attribute("llm.tokens", response.tokens_completion)
        
        return response''', language="python")

add_heading_h2(doc, "7.3  Prometheus Metrics")

add_code_block(doc, '''# modules/s3/metrics.py
from prometheus_client import Counter, Histogram, Gauge, CollectorRegistry

REGISTRY = CollectorRegistry()

# AI module metrics
ai_requests_total = Counter(
    "eden_ai_requests_total",
    "Total AI API requests",
    ["controller", "endpoint", "status"],
    registry=REGISTRY,
)

ai_latency_seconds = Histogram(
    "eden_ai_latency_seconds",
    "AI response latency",
    ["controller", "endpoint"],
    buckets=(0.1, 0.5, 1, 2, 5, 10, 30, 60, 120),
    registry=REGISTRY,
)

ai_tokens_used = Counter(
    "eden_ai_tokens_total",
    "Total tokens consumed",
    ["model", "type"],  # type: prompt|completion
    registry=REGISTRY,
)

ai_active_jobs = Gauge(
    "eden_ai_active_jobs",
    "Currently running AI jobs",
    ["controller"],
    registry=REGISTRY,
)

# Database metrics
db_connections_active = Gauge(
    "eden_db_connections_active",
    "Active database connections",
    registry=REGISTRY,
)

# Usage in ai_chat.py:
ai_requests_total.labels(controller="ai_chat", endpoint="ask", status="200").inc()
with ai_latency_seconds.labels(controller="ai_chat", endpoint="ask").time():
    response = AIClient().chat(question)
ai_tokens_used.labels(model="glm-4-plus", type="completion").inc(response.tokens_completion)

# Expose at /metrics endpoint
# controllers/metrics.py
def index():
    from prometheus_client import generate_latest
    response.headers["Content-Type"] = "text/plain; version=0.0.4"
    return generate_latest()''', language="python")

add_heading_h2(doc, "7.4  Sentry Error Tracking")

add_code_block(doc, '''# modules/s3/sentry.py
import sentry_sdk
from sentry_sdk.integrations.redis import RedisIntegration
from sentry_sdk.integrations.psycopg2 import Psycopg2Integration

def init_sentry(dsn: str, environment: str = "production"):
    sentry_sdk.init(
        dsn=dsn,
        environment=environment,
        release=f"eden@{__version__}",
        traces_sample_rate=0.1,  # 10% performance monitoring
        profiles_sample_rate=0.1,
        integrations=[
            RedisIntegration(),
            Psycopg2Integration(),
        ],
        before_send=filter_sensitive_data,
    )

def filter_sensitive_data(event, hint):
    """Remove PII before sending to Sentry."""
    if "request" in event:
        # Strip auth headers
        event["request"].get("headers", {}).pop("Authorization", None)
        event["request"].get("headers", {}).pop("Cookie", None)
    
    if "request" in event and "data" in event["request"]:
        data = event["request"]["data"]
        if isinstance(data, dict):
            # Mask password fields
            for key in ["password", "secret", "api_key", "token"]:
                if key in data:
                    data[key] = "[REDACTED]"
            # Mask email fields
            if "email" in data:
                data["email"] = "[REDACTED]"
    
    return event''', language="python")

add_heading_h2(doc, "7.5  Health Endpoint")

add_code_block(doc, '''# controllers/health.py
"""Health check endpoints for Kubernetes probes."""

def index():
    """Liveness + readiness probe."""
    from s3.metrics import REGISTRY
    import psutil
    import os
    
    checks = {}
    
    # Database
    try:
        current.db._adapter.connection.execute("SELECT 1")
        checks["database"] = "ok"
    except Exception as e:
        checks["database"] = f"fail: {e}"
    
    # Redis
    try:
        import redis
        r = redis.from_url(os.environ.get("REDIS_URL"))
        r.ping()
        checks["redis"] = "ok"
    except Exception as e:
        checks["redis"] = f"fail: {e}"
    
    # AI engine (lightweight ping)
    try:
        from s3.ai_client import AIClient
        client = AIClient(timeout=5)
        # Just check CLI exists, don't make real call
        checks["ai_engine"] = "ok" if client.cli_path else "fail: CLI not found"
    except Exception as e:
        checks["ai_engine"] = f"fail: {e}"
    
    # Disk space (uploads directory)
    disk = psutil.disk_usage("/app/web2py/applications/eden/uploads")
    checks["disk_free_gb"] = f"{disk.free / (1024**3):.1f} GB"
    
    # Memory
    mem = psutil.virtual_memory()
    checks["memory_percent"] = f"{mem.percent}%"
    
    overall = "ok" if all(
        v == "ok" for k, v in checks.items() 
        if isinstance(v, str) and not v.startswith("fail")
    ) else "degraded"
    
    response.headers["Content-Type"] = "application/json"
    return dict(
        status=overall,
        version=__version__,
        timestamp=request.now.isoformat(),
        checks=checks,
    )''', language="python")

# ============================================================
# บทที่ 8: Async I/O
# ============================================================
add_heading_h1(doc, "Async I/O และ Streaming", num=8)

add_heading_h2(doc, "8.1  ปัญหาปัจจุบัน")
add_bullet(doc, "AIClient._run() ใช้ subprocess.run() (sync, block) — ทุก request รอ LLM response เต็มก่อนตอบ")
add_bullet(doc, "poor UX สำหรับ chatbot — user ต้องรอ 5-30 วินาทีโดยไม่เห็น progress")
add_bullet(doc, "z-ai CLI invocation มี overhead ~200ms ต่อ call")
add_bullet(doc, "ไม่รองรับ concurrent requests (web2py Rocket3 multi-threaded แต่ sync code block thread)")

add_heading_h2(doc, "8.2  Async AI Client (aiohttp)")

add_code_block(doc, '''# modules/s3/ai_client_async.py
import aiohttp
import asyncio
from typing import AsyncIterator, Optional
from dataclasses import dataclass

@dataclass
class AsyncAIClient:
    """Async wrapper around z-ai HTTP API (replaces subprocess CLI)."""
    
    api_key: str
    base_url: str = "https://api.z.ai/v1"
    timeout: int = 120
    
    async def chat(
        self,
        prompt: str,
        *,
        system: Optional[str] = None,
        model: str = "glm-4-plus",
        max_tokens: int = 2000,
        temperature: float = 0.7,
    ) -> str:
        messages = []
        if system:
            messages.append({"role": "system", "content": system})
        messages.append({"role": "user", "content": prompt})
        
        async with aiohttp.ClientSession(
            timeout=aiohttp.ClientTimeout(total=self.timeout)
        ) as session:
            async with session.post(
                f"{self.base_url}/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": model,
                    "messages": messages,
                    "max_tokens": max_tokens,
                    "temperature": temperature,
                },
            ) as resp:
                resp.raise_for_status()
                data = await resp.json()
                return data["choices"][0]["message"]["content"]
    
    async def chat_stream(
        self,
        prompt: str,
        *,
        model: str = "glm-4-plus",
    ) -> AsyncIterator[str]:
        """Stream tokens as they arrive (for SSE)."""
        async with aiohttp.ClientSession() as session:
            async with session.post(
                f"{self.base_url}/chat/completions",
                headers={"Authorization": f"Bearer {self.api_key}"},
                json={
                    "model": model,
                    "messages": [{"role": "user", "content": prompt}],
                    "stream": True,  # ← enable streaming
                },
            ) as resp:
                async for line in resp.content:
                    if line.startswith(b"data: "):
                        chunk = line[6:].strip()
                        if chunk == b"[DONE]":
                            break
                        import json
                        data = json.loads(chunk)
                        delta = data["choices"][0].get("delta", {}).get("content", "")
                        if delta:
                            yield delta''', language="python")

add_heading_h2(doc, "8.3  SSE Streaming Endpoint (Chatbot UX)")

add_code_block(doc, '''# controllers/ai_chat.py — streaming endpoint

def stream():
    """SSE endpoint: stream chat tokens as they arrive."""
    import asyncio
    from s3.ai_client_async import AsyncAIClient
    import os
    
    question = request.post_vars.question
    if not question:
        return dict(error="missing question")
    
    # PDPA consent check (same as ask())
    if not session.ai_chat_consent:
        return dict(error="consent_required")
    
    response.headers["Content-Type"] = "text/event-stream"
    response.headers["Cache-Control"] = "no-cache"
    response.headers["Connection"] = "keep-alive"
    
    async def generate():
        client = AsyncAIClient(api_key=os.environ["ZAI_API_KEY"])
        try:
            async for token in client.chat_stream(question):
                # SSE format: data: <json>\\n\\n
                yield f"data: {json.dumps({'token': token})}\\n\\n"
            
            # Send completion event
            yield f"data: {json.dumps({'done': True})}\\n\\n"
        
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)})}\\n\\n"
    
    # Run async generator in sync web2py context
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    
    def sync_gen():
        try:
            async_gen = generate()
            while True:
                try:
                    chunk = loop.run_until_complete(async_gen.__anext__())
                    yield chunk
                except StopAsyncIteration:
                    break
        finally:
            loop.close()
    
    return sync_gen()''', language="python")

add_heading_h2(doc, "8.4  Celery + Redis (Background Jobs)")

add_code_block(doc, '''# celery_app.py
from celery import Celery
import os

app = Celery(
    "eden",
    broker=os.environ.get("REDIS_URL", "redis://redis:6379/0"),
    backend=os.environ.get("REDIS_URL", "redis://redis:6379/0"),
)

app.conf.update(
    task_serializer="json",
    result_serializer="json",
    accept_content=["json"],
    timezone="Asia/Bangkok",
    enable_utc=True,
    task_track_started=True,
    task_time_limit=30 * 60,  # 30 min hard limit
    task_soft_time_limit=25 * 60,
    worker_prefetch_multiplier=1,  # fair scheduling
    worker_max_tasks_per_child=100,  # restart workers periodically
)

# tasks/ai_tasks.py
@app.task(bind=True)
def run_forecast(self, metric: str):
    """Long-running forecast (Use Case 3)."""
    from controllers.ai_analytics import generate_forecast
    try:
        result = generate_forecast(metric)
        return {"status": "ok", "data": result}
    except Exception as e:
        self.update_state(state="FAILURE", meta={"error": str(e)})
        raise

@app.task(bind=True)
def process_ocr_batch(self, job_ids: list[int]):
    """Batch OCR processing."""
    from controllers.ai_ocr import process_job
    completed = 0
    for job_id in job_ids:
        process_job(job_id)
        completed += 1
        self.update_state(
            state="PROGRESS",
            meta={"current": completed, "total": len(job_ids)}
        )
    return {"completed": completed}

@app.task
def ingest_kb_chunk(chunk_id: int):
    """Generate embedding for a single KB chunk."""
    from controllers.ai_kb import embed_chunk
    embed_chunk(chunk_id)

# Scheduler: run nightly at 02:00
app.conf.beat_schedule = {
    "forecast-inventory-daily": {
        "task": "tasks.ai_tasks.run_forecast",
        "schedule": crontab(hour=2, minute=0),
        "args": ("inventory_demand",),
    },
    "forecast-shelter-daily": {
        "task": "tasks.ai_tasks.run_forecast",
        "schedule": crontab(hour=2, minute=15),
        "args": ("shelter_occupancy",),
    },
    "cleanup-chat-logs": {
        "task": "tasks.maintenance.cleanup_old_logs",
        "schedule": crontab(hour=3, minute=0),
        "args": (90,),  # 90 days retention
    },
}''', language="python")

# ============================================================
# บทที่ 9: Security
# ============================================================
add_heading_h1(doc, "Security Hardening (เพิ่มจากที่มี)", num=9)

add_heading_h2(doc, "9.1  CSP Headers (Caddy)")

add_code_block(doc, '''# Caddyfile (security headers section)

{$EDEN_DOMAIN} {
    header {
        # Content Security Policy — restrict sources
        Content-Security-Policy """
            default-src 'self';
            script-src 'self' 'unsafe-inline' cdnjs.cloudflare.com;
            style-src 'self' 'unsafe-inline' cdnjs.cloudflare.com fonts.googleapis.com;
            img-src 'self' data: https: *.tile.openstreetmap.org *.tile.openstreetmap.fr;
            font-src 'self' cdnjs.cloudflare.com fonts.gstatic.com;
            connect-src 'self' https://api.z.ai https://api.iapp.co.th;
            frame-ancestors 'none';
            form-action 'self';
            base-uri 'self';
            object-src 'none';
        """
        
        X-Frame-Options "DENY"
        X-Content-Type-Options "nosniff"
        Referrer-Policy "strict-origin-when-cross-origin"
        Strict-Transport-Security "max-age=31536000; includeSubDomains; preload"
        Permissions-Policy "geolocation=(), microphone=(self), camera=()"
        Cross-Origin-Opener-Policy "same-origin"
        Cross-Origin-Embedder-Policy "require-corp"
    }
}''', language="text")

add_heading_h2(doc, "9.2  2FA for Admin Accounts")

add_code_block(doc, '''# modules/s3/two_factor.py
import pyotp
import qrcode
import io
import base64
from gluon import current

class TwoFactorAuth:
    """Time-based One-Time Password (TOTP) for admin accounts."""

    @staticmethod
    def generate_secret(user_id: int) -> str:
        """Generate TOTP secret for user."""
        secret = pyotp.random_base32()
        db = current.db
        db.auth_user[user_id].update_record(totp_secret=secret)
        return secret
    
    @staticmethod
    def get_qr_code_data_uri(user_email: str, secret: str) -> str:
        """Generate QR code for authenticator app."""
        totp = pyotp.TOTP(secret)
        uri = totp.provisioning_uri(
            name=user_email,
            issuer_name="Sahana Eden"
        )
        img = qrcode.make(uri)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}"
    
    @staticmethod
    def verify(user_id: int, code: str) -> bool:
        """Verify TOTP code (±1 window for clock skew)."""
        db = current.db
        user = db.auth_user[user_id]
        if not user or not user.totp_secret:
            return False
        totp = pyotp.TOTP(user.totp_secret)
        return totp.verify(code, valid_window=1)  # ±30s

# controllers/default.py — modify user/login
def user():
    if request.args(0) == "login" and request.env.request_method == "POST":
        # After password auth succeeds, check if 2FA required
        user = auth.user
        if user and user.totp_secret:
            if not request.post_vars.totp_code:
                # Show 2FA form
                return dict(two_factor_required=True)
            if not TwoFactorAuth.verify(user.id, request.post_vars.totp_code):
                return dict(error="Invalid 2FA code")''', language="python")

add_heading_h2(doc, "9.3  Column-Level Encryption")

add_code_block(doc, '''# modules/s3/encryption.py
from cryptography.fernet import Fernet
import os
from typing import Optional

class FieldEncryptor:
    """Encrypt sensitive fields at rest."""
    
    def __init__(self):
        key = os.environ.get("EDEN_ENCRYPTION_KEY")
        if not key:
            raise RuntimeError(
                "EDEN_ENCRYPTION_KEY environment variable required"
            )
        self._fernet = Fernet(key.encode() if isinstance(key, str) else key)
    
    def encrypt(self, plaintext: Optional[str]) -> Optional[str]:
        if plaintext is None:
            return None
        return self._fernet.encrypt(plaintext.encode()).decode()
    
    def decrypt(self, ciphertext: Optional[str]) -> Optional[str]:
        if ciphertext is None:
            return None
        return self._fernet.decrypt(ciphertext.encode()).decode()

# Usage in models/zz_ai_tables.py:
from s3.encryption import FieldEncryptor
enc = FieldEncryptor()

# After table definition, add encrypted accessors:
db.ai_ocr_job.extracted_data.encrypt = enc.encrypt
db.ai_ocr_job.extracted_data.decrypt = enc.decrypt

# In dvr_case (sensitive victim data):
db.dvr_case.sensitive_notes.encrypt = enc.encrypt
db.dvr_case.sensitive_notes.decrypt = enc.decrypt

# In auth_user (Thai national ID):
db.auth_user.national_id.encrypt = enc.encrypt
db.auth_user.national_id.decrypt = enc.decrypt

# Migration: encrypt existing data
def migrate_encrypt(tablename, fieldname, batch_size=1000):
    """Encrypt existing plaintext data in batches."""
    db = current.db
    rows = db(db[tablename].id > 0).select(
        db[tablename].id, db[tablename][fieldname]
    )
    for row in rows:
        if row[fieldname] and not row[fieldname].startswith("gAAAAA"):  # not encrypted
            encrypted = enc.encrypt(row[fieldname])
            db[tablename][row.id].update(**{fieldname: encrypted})''', language="python")

add_heading_h2(doc, "9.4  Audit Log Immutability")

add_code_block(doc, '''# modules/s3/audit.py
"""Send audit logs to syslog (immutable) in addition to DB."""

import logging
import logging.handlers
import json
from gluon import current

# Configure syslog handler (write-once, admin can't modify)
syslog_handler = logging.handlers.SysLogHandler(
    address="/dev/log",
    facility=logging.handlers.SysLogHandler.LOG_LOCAL0,
)
syslog_handler.setFormatter(
    logging.Formatter("eden[%(process)d]: %(message)s")
)

audit_logger = logging.getLogger("eden.audit")
audit_logger.addHandler(syslog_handler)
audit_logger.setLevel(logging.INFO)

def audit(action: str, controller: str, function: str, 
          record_id: int = None, **extra):
    """Log audit event to BOTH database AND syslog."""
    # 1. Database (existing behavior)
    current.audit(action, controller, function, record=record_id)
    
    # 2. Syslog (immutable, can't be modified by admin)
    log_entry = {
        "action": action,
        "controller": controller,
        "function": function,
        "record_id": record_id,
        "user_id": current.auth.user.id if current.auth.user else None,
        "user_email": current.auth.user.email if current.auth.user else None,
        "ip": current.request.env.remote_addr,
        "timestamp": current.request.now.isoformat(),
        **extra,
    }
    audit_logger.info(json.dumps(log_entry))

# Forward to external SIEM (Splunk, ELK) via syslog
# Configure in /etc/rsyslog.conf:
# local0.*  @@siem.example.com:514''', language="python")

add_heading_h2(doc, "9.5  Dependency Scanning")

add_code_block(doc, '''# scripts/check-deps.sh
#!/bin/bash
# Run weekly: checks for known vulnerabilities

set -e

echo "=== pip-audit ==="
pip install --quiet pip-audit
pip-audit -r requirements.txt -r optional_requirements.txt

echo ""
echo "=== npm audit (for frontend) ==="
if [ -f package.json ]; then
    npm audit --omit=dev || true
fi

echo ""
echo "=== Docker image scan (Trivy) ==="
if command -v trivy &> /dev/null; then
    trivy image ghcr.io/skpuan07-cmd/eden:latest --severity HIGH,CRITICAL
fi

echo ""
echo "=== Secrets in code (detect-secrets) ==="
detect-secrets scan --all-files --exclude-files '.git/.*' \\
    --exclude-files 'static/.*' --baseline .secrets.baseline''', language="bash")

# ============================================================
# บทที่ 10: API Modernization
# ============================================================
add_heading_h1(doc, "API Modernization", num=10)

add_heading_h2(doc, "10.1  OpenAPI 3.0 Specification")

add_code_block(doc, '''# modules/s3/openapi.py
"""Auto-generate OpenAPI 3.0 spec from Eden controllers."""

from openapi_pydantic import (
    OpenAPI, Info, PathItem, Operation, Response,
    Parameter, RequestBody, MediaType, Schema,
)

def build_spec() -> OpenAPI:
    return OpenAPI(
        info=Info(
            title="Sahana Eden API",
            version="6.2",
            description="REST API for Sahana Eden with AI extensions",
            contact={"email": "admin@example.com"},
            license={"name": "MIT", "url": "https://opensource.org/licenses/MIT"},
        ),
        servers=[
            {"url": "/eden", "description": "Relative to current host"},
        ],
        paths={
            # AI Chatbot endpoints
            "/ai_chat/ask": PathItem(
                post=Operation(
                    tags=["AI Chat"],
                    summary="Submit a question to the AI chatbot",
                    description="Requires PDPA consent. Rate limited: 60 req/min/IP.",
                    requestBody=RequestBody(
                        required=True,
                        content={
                            "application/json": MediaType(
                                schema=Schema(
                                    type="object",
                                    properties={
                                        "question": Schema(
                                            type="string", 
                                            maxLength=1000,
                                            example="ศูนย์พักพิงใกล้ฉันอยู่ที่ไหน?"
                                        ),
                                    },
                                    required=["question"],
                                )
                            )
                        },
                    ),
                    responses={
                        "200": Response(
                            description="Successful response",
                            content={
                                "application/json": MediaType(
                                    schema=Schema(
                                        type="object",
                                        properties={
                                            "answer": Schema(type="string"),
                                            "sources": Schema(type="array", items=Schema(type="string")),
                                            "latency_ms": Schema(type="integer"),
                                        },
                                    )
                                )
                            },
                        ),
                        "429": Response(description="Rate limit exceeded"),
                        "403": Response(description="PDPA consent required"),
                    },
                    security=[{"sessionAuth": []}],
                )
            ),
            # ... other endpoints
        },
        components={
            "securitySchemes": {
                "sessionAuth": {
                    "type": "apiKey",
                    "in": "cookie",
                    "name": "session_id_eden",
                },
                "apiKey": {
                    "type": "apiKey",
                    "in": "header",
                    "name": "X-API-Key",
                },
            }
        },
    )

# controllers/api_docs.py
def openapi_json():
    """Serve OpenAPI spec at /eden/api_docs/openapi_json"""
    from s3.openapi import build_spec
    response.headers["Content-Type"] = "application/json"
    return build_spec().json(by_alias=True, exclude_none=True)

def swagger():
    """Swagger UI at /eden/api_docs/swagger"""
    response.view = "api_docs/swagger.html"
    return dict(spec_url=URL("api_docs", "openapi_json"))

def redoc():
    """ReDoc UI at /eden/api_docs/redoc"""
    response.view = "api_docs/redoc.html"
    return dict(spec_url=URL("api_docs", "openapi_json"))''', language="python")

add_heading_h2(doc, "10.2  API Versioning Strategy")

add_code_block(doc, '''# URL pattern:
# /eden/api/v1/org/organisation.json   — stable v1
# /eden/api/v2/org/organisation.json   — breaking changes
# /eden/org/organisation.json          — legacy (current behavior, no version prefix)

# routers.py — add API version router
routers = dict(
    BASE=dict(
        default_application="eden",
        applications=["eden"],
    ),
    eden=dict(
        controllers=[
            # Standard Eden controllers (legacy URLs)
            "default", "org", "pr", "hrm", "inv", "project", "asset",
            "req", "cap", "cms", "dvr", "gis", "msg", "vehicle",
            # AI controllers
            "ai", "ai_chat", "ai_ocr", "ai_analytics", "ai_imagery",
            "ai_translate", "ai_query", "ai_kb", "ai_voice", "ai_moderation",
            # New API v1/v2 controllers
            "api_v1", "api_v2", "api_docs",
        ],
    ),
)

# controllers/api_v1.py
def index():
    """API v1 — stable, backward-compatible."""
    response.headers["Content-Type"] = "application/json"
    return dict(
        version="1.0",
        status="stable",
        endpoints=[
            "/eden/api/v1/org/organisation",
            "/eden/api/v1/pr/person",
            "/eden/api/v1/inv/warehouse",
            # ...
        ],
    )

def org_organisation():
    """Versioned organization endpoint."""
    # Wrap existing S3 CRUD with versioned response format
    from gluon import current
    current.response.s3.prep = lambda r: r.method == "read"
    return current.crud_controller("org", "organisation")''', language="python")

add_heading_h2(doc, "10.3  SDK Generation")

add_code_block(doc, '''# scripts/generate-sdks.sh
#!/bin/bash
# Generate client SDKs from OpenAPI spec

set -e

# Download spec
curl -s http://localhost:8000/eden/api_docs/openapi_json > /tmp/eden-openapi.json

# Generate Python SDK
npx @openapitools/openapi-generator-cli generate \\
    -i /tmp/eden-openapi.json \\
    -g python \\
    -o sdk-python/ \\
    --additional-properties=packageName=eden_client

# Generate TypeScript SDK
npx @openapitools/openapi-generator-cli generate \\
    -i /tmp/eden-openapi.json \\
    -g typescript-fetch \\
    -o sdk-ts/ \\
    --additional-properties=supportsES6=true

# Generate Java SDK (for Android app)
npx @openapitools/openapi-generator-cli generate \\
    -i /tmp/eden-openapi.json \\
    -g java \\
    -o sdk-java/ \\
    --additional-properties=library=retrofit2

echo "✓ SDKs generated in sdk-python/, sdk-ts/, sdk-java/"''', language="bash")

# ============================================================
# บทที่ 11: Documentation
# ============================================================
add_heading_h1(doc, "Documentation & Developer Onboarding", num=11)

add_heading_h2(doc, "11.1  MkDocs Material (แทน Sphinx)")

add_code_block(doc, '''# mkdocs.yml
site_name: Sahana Eden Documentation
site_description: Open-source disaster management platform with AI extensions
site_url: https://skpuan07-cmd.github.io/eden/

theme:
  name: material
  features:
    - navigation.tabs
    - navigation.sections
    - navigation.expand
    - navigation.top
    - search.highlight
    - search.suggest
    - content.code.copy
    - content.code.annotate
  palette:
    - scheme: default
      primary: indigo
      toggle:
        icon: material/toggle-switch-off-outline
        name: Switch to dark mode
    - scheme: slate
      toggle:
        icon: material/toggle-switch
        name: Switch to light mode

plugins:
  - search
  - mkdocstrings:
      handlers:
        python:
          options:
            show_source: false
            show_root_heading: true
            show_signature_annotations: true
            separate_signature: true
  - git-revision-date-localized:
      type: datetime

markdown_extensions:
  - admonition
  - pymdownx.details
  - pymdownx.superfences:
      custom_fences:
        - name: mermaid
          class: mermaid
          format: !!python/name:pymdownx.superfences.fence_code_format
  - pymdownx.highlight:
      anchor_linenums: true
  - pymdownx.inlinehilite
  - pymdownx.tabbed:
      alternate_style: true
  - toc:
      permalink: true

nav:
  - Home: index.md
  - Getting Started:
      - Installation: getting-started/installation.md
      - Quick Start: getting-started/quickstart.md
      - Configuration: getting-started/configuration.md
  - Architecture:
      - Overview: architecture/overview.md
      - S3 Framework: architecture/s3-framework.md
      - Database Schema: architecture/database.md
      - AI Modules: architecture/ai-modules.md
  - AI Modules:
      - Overview: ai/README.md
      - Chatbot: ai/chatbot.md
      - OCR: ai/ocr.md
      - Analytics: ai/analytics.md
      - Imagery: ai/imagery.md
      - Translation: ai/translation.md
      - NLQ: ai/nlq.md
      - Knowledge Base: ai/kb.md
      - Voice: ai/voice.md
      - Moderation: ai/moderation.md
  - API Reference: api-reference.md
  - Deployment:
      - Docker: deployment/docker.md
      - Kubernetes: deployment/kubernetes.md
      - Production: deployment/production.md
  - Development:
      - Contributing: development/contributing.md
      - Testing: development/testing.md
      - Code Style: development/code-style.md''', language="yaml")

add_heading_h2(doc, "11.2  Architecture Decision Records (ADRs)")

add_code_block(doc, '''# docs/adr/0001-record-architecture-decisions.md
# ADR 0001: Record Architecture Decisions

Date: 2026-11-19

## Status
Accepted

## Context
We need to record architectural decisions made in this project.

## Decision
We will use Architecture Decision Records (ADRs) to record decisions.

ADRs will be stored in `docs/adr/` and numbered sequentially.

## Consequences
- Future developers can understand WHY decisions were made
- Easy to revisit decisions when requirements change
- Lightweight (Markdown only, no tooling required)

---

# docs/adr/0002-use-postgres-with-pgvector.md
# ADR 0002: Use PostgreSQL with pgvector for Production

Date: 2026-11-19

## Status
Accepted

## Context
Current SQLite setup doesn't support:
- Concurrent writes (single writer)
- Vector similarity search (placeholder JSON embeddings)
- Production-grade backup/replication

## Decision
Migrate to PostgreSQL 16 + pgvector extension.

## Consequences
- **Positive**: Production-ready, vector search, replication, FTS
- **Negative**: Requires PostgreSQL operational expertise
- **Migration plan**: Use Alembic for schema migrations, dual-write during transition

---

# docs/adr/0003-adopt-htmx-over-react.md
# ADR 0003: Adopt HTMX over React for Frontend

Date: 2026-11-19

## Status
Accepted

## Context
Current frontend uses jQuery (EOL) + Foundation 5 (legacy).
Options:
1. React/Vue SPA — full rewrite, complex
2. HTMX + Alpine.js — incremental, simpler

## Decision
Adopt HTMX + Alpine.js + Tailwind CSS.

## Consequences
- **Positive**: Incremental migration, no SPA complexity, smaller bundle
- **Negative**: Less interactive than SPA for complex UIs
- **Mitigation**: Use HTMX for 90% of UIs, React only for complex widgets (e.g., map)''', language="markdown")

add_heading_h2(doc, "11.3  Runbooks (Operations)")

add_table(doc,
    headers=["Runbook", "สถานการณ์", "ขั้นตอนหลัก"],
    rows=[
        ["deployment.md",        "Deploy new version",                "1. CI/CD triggers  2. Helm upgrade  3. Health check  4. Smoke test  5. Monitor"],
        ["rollback.md",          "Rollback bad deployment",            "1. Identify bad version  2. helm rollback  3. Verify  4. Notify users"],
        ["database-restore.md",  "Restore DB from backup",             "1. Stop Eden  2. Drop current DB  3. pg_restore  4. Run migrations  5. Restart"],
        ["ai-engine-down.md",    "z-ai API unavailable",                "1. Detect via health check  2. Switch to fallback model  3. Notify users  4. Monitor"],
        ["high-load.md",         "Traffic spike during disaster",      "1. HPA scales pods  2. Enable read replicas  3. Cache aggressively  4. Throttle non-critical"],
        ["pdpa-data-breach.md",  "Suspected data breach",              "1. Contain  2. Assess scope  3. Notify PDPA officer  4. Notify affected users  5. Document"],
        ["disk-full.md",         "Uploads directory full",             "1. Identify large files  2. Move to S3  3. Cleanup  4. Alert on threshold"],
        ["cert-expired.md",      "TLS certificate expiring",           "1. Caddy auto-renews  2. If failed, manual certbot  3. Verify HTTPS"],
    ],
    col_widths_cm=[3.5, 4.0, 8.5],
    first_col_bold=True, font_size=9)

add_heading_h2(doc, "11.4  VS Code Dev Container")

add_code_block(doc, '''// .devcontainer/devcontainer.json
{
  "name": "Sahana Eden Development",
  "image": "mcr.microsoft.com/devcontainers/python:3.12",
  
  "features": {
    "ghcr.io/devcontainers/features/docker-outside-of-docker:1": {},
    "ghcr.io/devcontainers/features/node:1": { "version": "20" },
    "ghcr.io/devcontainers/features/postgresql:1": { "version": "16" }
  },
  
  "postCreateCommand": [
    "pip install --user -r requirements.txt -r optional_requirements.txt",
    "pip install --user pytest pytest-cov ruff black mypy pre-commit",
    "pre-commit install",
    "git submodule update --init --recursive"
  ],
  
  "forwardPorts": [8000, 5432, 6379],
  
  "portsAttributes": {
    "8000": { "label": "Eden web", "onAutoForward": "openBrowser" },
    "5432": { "label": "PostgreSQL" },
    "6379": { "label": "Redis" }
  },
  
  "customizations": {
    "vscode": {
      "extensions": [
        "ms-python.python",
        "ms-python.vscode-pylance",
        "charliermarsh.ruff",
        "ms-python.black-formatter",
        "ms-toolsai.jupyter",
        "redhat.vscode-yaml",
        "github.copilot"
      ],
      "settings": {
        "python.defaultInterpreterPath": "/usr/local/bin/python",
        "python.linting.enabled": true,
        "python.formatting.provider": "black",
        "editor.formatOnSave": true,
        "editor.rulers": [100]
      }
    }
  }
}''', language="json")

# ============================================================
# บทที่ 12: Priority Matrix + Roadmap
# ============================================================
add_heading_h1(doc, "สรุป Priority Matrix และ Roadmap", num=12)

add_heading_h2(doc, "12.1  Detailed Priority Matrix")

add_table(doc,
    headers=["#", "งาน", "Phase", "Effort", "Impact", "Priority"],
    rows=[
        ["1",  "Dockerfile + docker-compose",                 "1", "S (1 วัน)",   "สูง",    "★★★★★"],
        ["2",  "Pre-commit hooks (ruff/black/mypy)",            "1", "S",          "สูง",    "★★★★★"],
        ["3",  "GitHub Actions CI (lint+test+build)",          "1", "M (3 วัน)",   "สูง",    "★★★★★"],
        ["4",  "Type hints สำหรับ AI modules",                 "1", "M",          "กลาง",   "★★★★☆"],
        ["5",  "Health endpoint + structured logging",         "1", "S",          "กลาง",   "★★★★☆"],
        ["6",  "Migrate SQLite → PostgreSQL 16",                "2", "L (1 สัปดาห์)", "สูงมาก", "★★★★★"],
        ["7",  "Test suite + 80% coverage สำหรับ AI",          "2", "L",          "สูง",    "★★★★★"],
        ["8",  "Alembic migrations",                            "2", "M",          "กลาง",   "★★★★☆"],
        ["9",  "Sentry + Prometheus + Grafana",                 "2", "M",          "สูง",    "★★★★☆"],
        ["10", "OpenAPI spec + Swagger UI",                     "2", "M",          "กลาง",   "★★★★☆"],
        ["11", "Async AI client (aiohttp)",                     "3", "L",          "สูง",    "★★★★☆"],
        ["12", "Streaming chat (SSE)",                          "3", "M",          "สูง",    "★★★★☆"],
        ["13", "Celery + Redis สำหรับ background jobs",        "3", "L",          "สูง",    "★★★★☆"],
        ["14", "HTMX migration (incremental)",                  "3", "L",          "กลาง",   "★★★☆☆"],
        ["15", "MapLibre GL แทน OpenLayers",                    "3", "M",          "กลาง",   "★★★☆☆"],
        ["16", "Kubernetes Helm chart",                         "4", "L",          "สูง",    "★★★★☆"],
        ["17", "CDN + image optimization",                     "4", "S",          "กลาง",   "★★★☆☆"],
        ["18", "PWA offline support",                           "4", "L",          "สูงมาก", "★★★★★"],
        ["19", "Mobile app (React Native)",                     "4", "XL",         "สูงมาก", "★★★★☆"],
        ["20", "2FA for admin accounts",                        "5", "M",          "สูงมาก", "★★★★★"],
        ["21", "Column-level encryption",                       "5", "M",          "สูง",    "★★★★☆"],
        ["22", "WAF + CSP headers",                             "5", "S",          "สูง",    "★★★★☆"],
        ["23", "Penetration testing",                            "5", "L",          "สูง",    "★★★★☆"],
        ["24", "PDPA DPIA documentation",                       "5", "M",          "สูง",    "★★★★☆"],
        ["25", "Audit log immutability (syslog)",               "5", "S",          "สูง",    "★★★★☆"],
    ],
    col_widths_cm=[0.7, 6.5, 1.0, 1.8, 1.5, 1.7],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "12.2  Roadmap Timeline")

add_code_block(doc, '''┌────────────────────────────────────────────────────────────────────────┐
│  Phase 1: Quick Wins         สัปดาห์ที่ 1-2   ▓▓▓▓▓▓▓▓▓▓                │
│  ─────────────────────────────────────────                              │
│  • Dockerfile + docker-compose                                           │
│  • Pre-commit hooks + GitHub Actions CI                                  │
│  • Type hints + health endpoint                                          │
│                                                                          │
│  Phase 2: Foundation          เดือนที่ 1     ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓  │
│  ─────────────────────────────────────────                              │
│  • PostgreSQL 16 + pgvector migration                                    │
│  • Test suite (pytest, 80% coverage)                                     │
│  • Alembic migrations                                                    │
│  • Sentry + Prometheus + Grafana                                         │
│  • OpenAPI spec + Swagger UI                                             │
│                                                                          │
│  Phase 3: Modernization       เดือนที่ 2-4   ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓  │
│  ─────────────────────────────────────────                              │
│  • Async AI client (aiohttp แทน subprocess)                              │
│  • Streaming chat (SSE)                                                  │
│  • Celery + Redis (forecast jobs, OCR batch)                             │
│  • HTMX migration (ทีละ module)                                          │
│  • MapLibre GL แทน OpenLayers                                            │
│                                                                          │
│  Phase 4: Scale               เดือนที่ 5-7   ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓  │
│  ─────────────────────────────────────────                              │
│  • Kubernetes Helm chart                                                 │
│  • CDN + image optimization                                              │
│  • PWA (offline support สำหรับพื้นที่ภัยพิบัติ)                            │
│  • Mobile app (React Native สำหรับอาสาสมัครสนาม)                       │
│                                                                          │
│  Phase 5: Compliance          ongoing      ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓  │
│  ─────────────────────────────────────────                              │
│  • 2FA admin + column encryption                                         │
│  • WAF + CSP headers                                                     │
│  • Penetration testing                                                   │
│  • PDPA DPIA + audit log immutability                                    │
└────────────────────────────────────────────────────────────────────────┘''', language="text")

add_heading_h2(doc, "12.3  Quick Wins (ทำได้ทันที — สัปดาห์แรก)")

add_table(doc,
    headers=["งาน", "ไฟล์ที่สร้างใหม่", "ไฟล์ที่แก้", "ผลลัพธ์"],
    rows=[
        ["Dockerfile",            "Dockerfile",                        "—",                        "docker build ได้ image เดียวที่รันได้"],
        ["docker-compose",        "docker-compose.yml, .env.example",  "—",                        "docker compose up ได้ระบบเต็ม (eden+pg+redis+caddy)"],
        ["Caddyfile",             "Caddyfile",                         "—",                        "auto HTTPS + security headers"],
        ["Pre-commit",            ".pre-commit-config.yaml",           "—",        "auto lint+format ทุก commit"],
        ["pyproject.toml",        "pyproject.toml",                    "—",                        "centralized ruff/black/mypy/pytest config"],
        ["GitHub Actions CI",     ".github/workflows/ci.yml",          "—",        "auto lint+test+build ทุก push"],
        ["Type hints",            "—",                                 "modules/s3/ai_client.py",  "better IDE support + mypy checks"],
        ["Health endpoint",       "controllers/health.py",             "—",                        "/eden/health สำหรับ K8s probes"],
        ["Structured logging",    "modules/s3/logging_config.py",      "controllers/ai_chat.py",   "JSON logs สำหรับ log aggregation"],
    ],
    col_widths_cm=[3.0, 5.5, 4.5, 5.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "12.4  Decision Matrix — Phase 2 (สูงสุด 5 อันดับแรกที่ควรทำ)")

add_table(doc,
    headers=["อันดับ", "งาน", "เหตุผล", "Risk ถ้าไม่ทำ"],
    rows=[
        ["1", "PostgreSQL + pgvector",       "Production-grade DB, vector search จริง",  "SQLite corrupt ได้ง่ายใน production, no vector search"],
        ["2", "Test suite 80% coverage",      "กัน regression, build confidence",          "Bug หลุดไป production บ่อย"],
        ["3", "Sentry error tracking",        "เห็น error real-time ไม่ใช่ดู errors/ ทีละไฟล์", "User พบ error ก่อน developer"],
        ["4", "Alembic migrations",           "Schema changes reproducible",              "Manual schema drift ระหว่าง env"],
        ["5", "OpenAPI spec",                 "API documentation + SDK generation",       "External integrators ต้องอ่านโค้ด"],
    ],
    col_widths_cm=[1.0, 4.5, 6.0, 5.5],
    first_col_bold=True)

add_heading_h2(doc, "12.5  สรุปและขั้นตอนถัดไป")

add_para(doc,
    "เอกสารฉบับนี้เสนอคำแนะนำการปรับปรุง Sahana Eden fork ใน 10 หมวด "
    "พร้อม code examples ที่ใช้งานได้จริง และ priority matrix แบ่งเป็น 5 phases "
    "รวมทั้งหมด 25 งาน ครอบคลุม containerization, CI/CD, database, frontend, "
    "code quality, observability, async I/O, security, API และ documentation "
    "โดยเน้น 'Quick Wins' ที่ทำได้ใน 1-2 สัปดาห์แรกเพื่อให้เห็นผลทันที "
    "และ 'Foundation' ในเดือนแรกเพื่อสร้าง infrastructure ที่ production-ready",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h3(doc, "ขั้นตอนถัดไปที่แนะนำ")

add_bullet(doc, "เริ่มจาก Phase 1 (Quick Wins) ทันที — 5 งานแรกทำได้ภายในสัปดาห์", bold_prefix="① ")
add_bullet(doc, "วางแผน Phase 2 (Foundation) อย่างละเอียด — เป็นพื้นฐานของทุกอย่างถัดไป", bold_prefix="② ")
add_bullet(doc, "อย่างลืมว่า docs/ai/RECOMMENDATIONS.md มีอยู่แล้ว — เอกสารนี้เสริม ไม่ทับซ้อน", bold_prefix="③ ")
add_bullet(doc, "ใช้ ADR pattern บันทึกการตัดสินใจทุกครั้ง — เพื่อ audit trail และ future reference", bold_prefix="④ ")
add_bullet(doc, "ทำ security audit + PDPA DPIA เร็วที่สุด — เป็น compliance requirement", bold_prefix="⑤ ")

add_heading_h3(doc, "ไฟล์เอกสารที่เกี่ยวข้อง")

add_table(doc,
    headers=["ไฟล์", "เนื้อหา"],
    rows=[
        ["Sahana_Eden_Modernization_Recommendations.docx",  "เอกสารฉบับนี้"],
        ["Sahana_Eden_Modernization_Recommendations.pdf",   "PDF version"],
        ["Sahana_Eden_skpuan07_Repo_Analysis.docx",          "วิเคราะห์โครงสร้าง repo"],
        ["Sahana_Eden_S3_Framework_Architecture.docx",       "S3 framework deep dive"],
        ["Sahana_Eden_Installation_README.md",                "คู่มือติดตั้ง base Eden"],
        ["docs/ai/RECOMMENDATIONS.md (ใน repo)",              "AI/Cyber/PDPA recommendations (977 บรรทัด)"],
        ["docs/ai/SYSTEM_DOCUMENTATION.md (ใน repo)",         "เอกสารระบบฉบับเต็ม (873 บรรทัด)"],
    ],
    col_widths_cm=[7.0, 9.0],
    first_col_bold=True)

# Footer
section = doc.sections[0]
footer = section.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fp.text = ""
add_run(fp, "Sahana Eden Modernization Roadmap  ·  Page ",
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
