"""
Generate comprehensive .docx for:
1. System relationship structure inspection
2. Strategic improvement plan (step-by-step, phased)
3. AI-powered registration features for staff/volunteers in disaster zones
4. Change log of all improvements
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_PATH = "/home/z/my-project/download/Sahana_Eden_Strategic_Improvement_Plan.docx"

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
add_run(p, "STRATEGIC IMPROVEMENT PLAN", font_name=FONT_HEADING,
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
add_run(p, "แผนปรับปรุงเชิงกลยุทธ์อย่างเป็นขั้นตอน",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
add_run(p, "ตรวจสอบโครงสร้างความสัมพันธ์ · เพิ่ม AI ลงทะเบียนเจ้าหน้าที่/จิตอาสา · บันทึกการปรับปรุง",
        font_name=FONT_BODY, size_pt=13, italic=True, color=C_AI)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(40)
add_run(p, "Phased Roadmap · System Relationship Map · AI-Powered Registration · Change Log",
        font_name=FONT_BODY, size_pt=11, color=C_SECOND)

info = doc.add_table(rows=6, cols=2)
info.alignment = WD_TABLE_ALIGNMENT.CENTER
info.autofit = False
info_data = [
    ("Repository",          "github.com/skpuan07-cmd/eden"),
    ("Base version",        "Sahana Eden v6.2 + AI modules (10 use cases)"),
    ("Focus",               "การนำข้อมูลเข้าระบบสะดวกขึ้น (staff + volunteer registration)"),
    ("Approach",            "Strangler Pattern — ทีละส่วน 6 เดือน"),
    ("เป้าหมาย",             "ลดเวลาลงทะเบียน 70% + เพิ่มช่องทางนำข้อมูลเข้า 5 รูปแบบ"),
    ("วันที่",                "19 พฤศจิกายน 2026"),
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
add_run(p, "เอกสารฉบับนี้รวม 4 ส่วน: ตรวจสอบระบบ · แผนปรับปรุง · "
           "ออกแบบ AI Registration · บันทึกการปรับปรุง",
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
    ("Part A",     "ตรวจสอบโครงสร้างความสัมพันธ์ของระบบ"),
    ("บทที่ 1",    "ภาพรวมโครงสร้าง Eden Modules"),
    ("บทที่ 2",    "Entity Relationship Map (ผังความสัมพันธ์)"),
    ("บทที่ 3",    "Pain Points ของ Data Entry Flow ปัจจุบัน"),
    ("Part B",     "แผนปรับปรุงเชิงกลยุทธ์ (Phased Plan)"),
    ("บทที่ 4",    "หลักการและเป้าหมายเชิงกลยุทธ์"),
    ("บทที่ 5",    "Phase 1: Foundation (เดือน 1)"),
    ("บทที่ 6",    "Phase 2: AI Registration MVP (เดือน 2-3)"),
    ("บทที่ 7",    "Phase 3: Smart Forms + Bulk Import (เดือน 4)"),
    ("บทที่ 8",    "Phase 4: Mobile PWA + Messaging Integration (เดือน 5)"),
    ("บทที่ 9",    "Phase 5: Production Migration (เดือน 6)"),
    ("Part C",     "AI-Powered Registration Features (Concrete Design)"),
    ("บทที่ 10",   "Feature 1: OCR-Based Staff/Volunteer Registration"),
    ("บทที่ 11",   "Feature 2: Voice-Based Registration (Thai)"),
    ("บทที่ 12",   "Feature 3: SMS/Line/Telegram Bot Registration"),
    ("บทที่ 13",   "Feature 4: Bulk Import via Natural Language"),
    ("บทที่ 14",   "Feature 5: Smart Disaster Event Hub"),
    ("Part D",     "บันทึกการปรับปรุง (Change Log)"),
    ("บทที่ 15",   "Change Log Format และ ADRs"),
    ("บทที่ 16",   "สรุปและขั้นตอนถัดไป"),
]
for num, title in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(15.5), WD_ALIGN_PARAGRAPH.RIGHT, 2)
    add_run(p, num + "  ", font_name=FONT_HEADING, size_pt=11, bold=True, color=C_PRIMARY)
    add_run(p, title, font_name=FONT_BODY, size_pt=11, color=C_DARK)

doc.add_page_break()

# ============================================================
# PART A DIVIDER
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(80)
add_run(p, "PART A", font_name=FONT_HEADING, size_pt=42, bold=True, color=C_PRIMARY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(20)
add_run(p, "ตรวจสอบโครงสร้างความสัมพันธ์ของระบบ",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "วิเคราะห์โครงสร้าง modules, entity relationships และ pain points ของ data entry flow",
        font_name=FONT_BODY, size_pt=11, italic=True, color=C_SECOND)

doc.add_page_break()

# ============================================================
# บทที่ 1: ภาพรวมโครงสร้าง Eden Modules
# ============================================================
add_heading_h1(doc, "ภาพรวมโครงสร้าง Eden Modules", num=1)

add_para(doc,
    "Sahana Eden ประกอบด้วย 41 domain modules ที่เชื่อมโยงกันผ่าน foreign keys "
    "และ component relationships ในแบบ master-detail "
    "การทำความเข้าใจความสัมพันธ์เหล่านี้เป็นพื้นฐานสำคัญก่อนออกแบบ AI features "
    "เพื่อให้ AI ช่วย auto-fill ข้อมูลที่เกี่ยวข้องข้าม modules ได้อย่างถูกต้อง",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "1.1  โมดูลหลักที่เกี่ยวข้องกับการลงทะเบียนเจ้าหน้าที่/จิตอาสา")

add_table(doc,
    headers=["Module", "Controller", "ตารางหลัก", "หน้าที่ในงานลงทะเบียน"],
    rows=[
        ["auth",      "default",     "auth_user",                  "บัญชีผู้ใช้ (email/password) — login"],
        ["pr",        "pr",          "pr_person, pr_pentity",      "ทะเบียนบุคคล (ชื่อ-นามสกุล, วันเกิด, เพศ, ที่อยู่)"],
        ["org",       "org",         "org_organisation, org_office, org_site", "องค์กร/สำนักงาน/สาขา ที่เจ้าหน้าที่สังกัด"],
        ["hrm",       "hrm",         "hrm_human_resource, hrm_skill, hrm_certificate", "ทรัพยากรบุคคล (เจ้าหน้าที่ + อาสาสมัคร + skill)"],
        ["vol",       "vol",         "vol_volunteer, vol_application, vol_assignment", "อาสาสมัครเฉพาะ (แยกจาก hrm)"],
        ["gis",       "gis",         "gis_location, gis_hierarchy", "ตำแหน่งที่อยู่ — ใช้กับที่อยู่เจ้าหน้าที่ + พื้นที่ภัยพิบัติ"],
        ["event",     "event",       "event_event, event_incident", "เหตุการณ์ภัยพิบัติ — อ้างอิงเมื่อลงทะเบียนในพื้นที่"],
        ["cr",        "cr",          "cr_shelter, cr_shelter_registration", "ศูนย์พักพิง — เจ้าหน้าที่ประจำศูนย์"],
        ["dvr",       "dvr",         "dvr_case, dvr_case_person",  "ผู้ประสบภัย (sensitive — PDPA)"],
        ["inv",       "inv",         "inv_warehouse, inv_inv_item", "คลังสินค้า — เจ้าหน้าที่โลจิสติกส์"],
        ["asset",     "asset",       "asset_asset, asset_item",    "ครุภัณฑ์ — ผู้รับผิดชอบ"],
        ["req",       "req",         "req_req",                    "คำขอ — ผู้ส่งคำขอ"],
        ["cap",       "cap",         "cap_alert",                  "การแจ้งเตือนภัย"],
        ["cms",       "cms",         "cms_post",                   "ข่าวสาร/ประกาศ"],
        ["msg",       "msg",         "msg_log, msg_channel",       "SMS/Email/Telegram (ใช้ส่ง OTP + การยืนยัน)"],
    ],
    col_widths_cm=[1.5, 1.5, 5.5, 7.5],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "1.2  สถิติโมดูลที่ใช้ในการลงทะเบียน")

add_table(doc,
    headers=["ตัวชี้วัด", "ค่า", "หมายเหตุ"],
    rows=[
        ["Total modules ที่เกี่ยวข้อง", "15", "จาก 41 โมดูลทั้งหมด"],
        ["ตารางที่เกี่ยวกับบุคคล",     "20+", "pr_person, hrm_human_resource, vol_volunteer, dvr_case_person, ..."],
        ["Foreign keys ที่ cross-module", "60+", "organisation_id, site_id, person_id, location_id, ..."],
        ["Component relationships",   "100+", "add_components() calls ใน modules/s3db/*.py"],
        ["Form fields เฉลี่ยต่อการลงทะเบียน", "25-40", "เจ้าหน้าที่ใหม่ต้องกรอก ~30 fields"],
        ["เวลาเฉลี่ยลงทะเบียน 1 คน (manual)", "8-15 นาที", "เป้าหมาย: ลดเหลือ 2-5 นาทีด้วย AI"],
    ],
    col_widths_cm=[5.0, 3.0, 8.0],
    first_col_bold=True)

# ============================================================
# บทที่ 2: Entity Relationship Map
# ============================================================
add_heading_h1(doc, "Entity Relationship Map (ผังความสัมพันธ์)", num=2)

add_heading_h2(doc, "2.1  ผังความสัมพันธ์หลัก (Entity Relationship Diagram)")

add_code_block(doc, '''┌────────────────────────────────────────────────────────────────────────┐
│                         AUTHENTICATION LAYER                              │
│                                                                            │
│                        auth_user (login account)                           │
│                              │                                            │
│                              │ 1:1                                         │
│                              ▼                                            │
│  ┌─────────────────────────────────────────────────────────────────────┐ │
│  │                    pr_person (Person Registry)                       │ │
│  │                    ─────────────────────                            │ │
│  │  • id, pe_id (super entity)                                          │ │
│  │  • first_name, middle_name, last_name                               │ │
│  │  • gender, date_of_birth                                            │ │
│  │  • location_id ──────► gis_location                                 │ │
│  │  • pe_label (ID tag)                                                 │ │
│  │                                                                      │ │
│  │  Components: pr_address, pr_contact, pr_physical_description         │ │
│  └─────────────────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────────────────┘
                  │
                  │ 1:1 (via person_id)
                  ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                       HRM LAYER (Staff + Volunteers)                      │
│                                                                            │
│  ┌────────────────────────────────────────────────────────────────────┐ │
│  │              hrm_human_resource (the registration record)           │ │
│  │              ──────────────────────────────────────                  │ │
│  │  • id, person_id ──────► pr_person                                  │ │
│  │  • organisation_id ────► org_organisation                           │ │
│  │  • site_id ───────────► org_site (office/facility/shelter)          │ │
│  │  • type (staff/volunteer/trainee)                                   │ │
│  │  • job_title_id ──────► hrm_job_title                               │ │
│  │  • status (active/inactive)                                         │ │
│  │  • start_date, end_date                                             │ │
│  │                                                                     │ │
│  │  Components:                                                        │ │
│  │    • hrm_human_resource_skill (M:N with hrm_skill)                  │ │
│  │    • hrm_human_resource_certificate                                 │ │
│  │    • hrm_human_resource_training                                    │ │
│  │    • hrm_human_resource_department                                  │ │
│  └────────────────────────────────────────────────────────────────────┘ │
│                                                                            │
│  ┌────────────────────────┐   ┌────────────────────────┐                 │
│  │  vol_volunteer          │   │  hrm_skill             │                 │
│  │  (separate from hrm)    │   │  (skill catalog)       │                 │
│  │  • human_resource_id ──►│   │                        │                 │
│  │  • application status   │   └────────────────────────┘                 │
│  │  • cluster_id           │                                              │
│  └────────────────────────┘                                              │
└──────────────────────────────────────────────────────────────────────────┘
                  │
                  │ organisation_id / site_id
                  ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                  ORGANIZATION LAYER (where they work)                     │
│                                                                            │
│  ┌────────────────────────┐   ┌────────────────────────┐                 │
│  │  org_organisation       │   │  org_site (super)       │                │
│  │  • name, acronym        │   │  • name                 │                 │
│  │  • type_id ─► org_type  │   │  • location_id ─► gis   │                │
│  │  • parent_id (self-ref) │   │  • organisation_id ──►  │                │
│  └────────────────────────┘   └────────┬───────────────┘                 │
│            ▲                            │                                  │
│            │            ┌───────────────┼─────────────┐                  │
│            │            │               │             │                  │
│            │   ┌────────▼──┐  ┌────────▼────┐  ┌──────▼─────┐           │
│            │   │ org_office │  │ org_facility │  │ cr_shelter │           │
│            │   │            │  │              │  │ (ศูนย์พักพิง)│           │
│            │   └────────────┘  └──────────────┘  └────────────┘           │
│            │                                                                │
└────────────┼──────────────────────────────────────────────────────────────┘
             │
             │ location_id (geographic)
             ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                       GIS LAYER (Locations)                               │
│                                                                            │
│  ┌────────────────────────┐                                                │
│  │  gis_location           │   • Hierarchical: country → province →       │
│  │  • name                 │     district → subdistrict → village         │
│  │  • parent_id (self-ref) │   • lat, lon                                 │
│  │  • level (L0-L5)        │   • Used by: addresses, shelters, offices,   │
│  │  • lat, lon             │     warehouses, event locations              │
│  └────────────────────────┘                                                │
└──────────────────────────────────────────────────────────────────────────┘

─ Disaster Context ──────────────────────────────────────────────────────────
┌──────────────────────────────────────────────────────────────────────────┐
│  event_event (Disaster Event Hub)                                         │
│  ├── event_incident (specific incidents)                                 │
│  ├── event_event_location (areas affected)                              │
│  └── event_sitrep (situation reports)                                    │
└──────────────────────────────────────────────────────────────────────────┘
         │
         │ (currently loose coupling — should be tightened)
         ▼
   All other entities (org, hrm, inv, cr, dvr, asset) SHOULD reference
   event_event to scope operations during disaster response''', language="text")

add_heading_h2(doc, "2.2  Critical Foreign Keys (Cross-Module)")

add_table(doc,
    headers=["Child Table", "Field", "Parent Table", "ความหมาย"],
    rows=[
        ["hrm_human_resource", "person_id",        "pr_person",           "ตัวบุคคลที่ลงทะเบียน"],
        ["hrm_human_resource", "organisation_id",  "org_organisation",    "องค์กรที่สังกัด"],
        ["hrm_human_resource", "site_id",          "org_site",            "สาขา/สำนักงาน/ศูนย์พักพิง ที่ประจำ"],
        ["hrm_human_resource", "job_title_id",     "hrm_job_title",       "ตำแหน่งงาน"],
        ["vol_volunteer",      "human_resource_id","hrm_human_resource",  "เชื่อม volunteer เข้ากับ HR record"],
        ["vol_volunteer",      "cluster_id",       "vol_cluster",         "กลุ่มอาสาสมัคร (เช่น ตามพื้นที่)"],
        ["cr_shelter",         "site_id",          "org_site",            "ศูนย์พักพิง = site ประเภท shelter"],
        ["cr_shelter_registration", "shelter_id",  "cr_shelter",          "การลงทะเบียนผู้อพยพ"],
        ["cr_shelter_registration", "person_id",   "pr_person",           "ตัวบุคคลที่ลงทะเบียน"],
        ["dvr_case",           "person_id",        "pr_person",           "ผู้ประสบภัย (sensitive)"],
        ["dvr_case",           "shelter_id",       "cr_shelter",          "ศูนย์พักพิงที่รับ"],
        ["inv_warehouse",      "site_id",          "org_site",            "คลังสินค้า = site"],
        ["inv_inv_item",       "site_id",          "org_site",            "สินค้าอยู่ที่ site"],
        ["asset_asset",        "site_id",          "org_site",            "ครุภัณฑ์ประจำที่"],
        ["asset_asset",        "assigned_to_id",   "hrm_human_resource",  "ผู้รับผิดชอบครุภัณฑ์"],
        ["req_req",            "site_id",          "org_site",            "คำขอจาก site"],
        ["req_req",            "person_id",        "pr_person",           "ผู้ส่งคำขอ"],
        ["gis_location",       "parent_id",        "gis_location",        "ตำแหน่งลำดับชั้น (self-ref)"],
        ["pr_person",          "location_id",      "gis_location",        "ที่อยู่บุคคล"],
        ["event_event_location","location_id",    "gis_location",        "พื้นที่เกิดเหตุ"],
    ],
    col_widths_cm=[3.5, 3.0, 3.5, 6.0],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "2.3  Component Graph (Master-Detail Relationships)")

add_para(doc,
    "S3 framework มี concept 'component' — เป็น master-detail relationship "
    "ที่ access ผ่าน URL เช่น /eden/pr/person/5/contact (component 'contact' ของ person 5) "
    "ทำให้สามารถ fetch/modify sub-records ผ่าน master record ได้",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_table(doc,
    headers=["Master Table", "Component (Detail)", "URL Pattern", "จำนวน fields"],
    rows=[
        ["pr_person",         "pr_address",                "/eden/pr/person/5/address",          "8 fields"],
        ["pr_person",         "pr_contact",                "/eden/pr/person/5/contact",          "5 fields"],
        ["pr_person",         "pr_physical_description",   "/eden/pr/person/5/physical_description", "12 fields"],
        ["hrm_human_resource","hrm_human_resource_skill",  "/eden/hrm/staff/5/skill",            "4 fields"],
        ["hrm_human_resource","hrm_human_resource_certificate", "/eden/hrm/staff/5/certificate", "5 fields"],
        ["hrm_human_resource","hrm_human_resource_training",    "/eden/hrm/staff/5/training",    "6 fields"],
        ["cr_shelter",        "cr_shelter_registration",   "/eden/cr/shelter/3/registration",    "10 fields"],
        ["org_organisation",  "org_organisation_user",     "/eden/org/organisation/2/user",      "5 fields"],
        ["inv_warehouse",     "inv_inv_item",              "/eden/inv/warehouse/1/inv_item",     "12 fields"],
    ],
    col_widths_cm=[3.5, 4.5, 5.5, 2.5],
    first_col_bold=True, font_size=8.5, header_size=9)

# ============================================================
# บทที่ 3: Pain Points
# ============================================================
add_heading_h1(doc, "Pain Points ของ Data Entry Flow ปัจจุบัน", num=3)

add_heading_h2(doc, "3.1  ขั้นตอนการลงทะเบียนเจ้าหน้าที่/จิตอาสาปัจจุบัน")

add_code_block(doc, '''ขั้นตอนปัจจุบัน (manual, 8-15 นาทีต่อคน):

1. Login เป็น admin → ไปที่ /eden/hrm/staff → คลิก "Create Staff"
2. กรอกฟอร์ม pr_person (8 fields):
   □ ชื่อ, นามสกุล, ชื่อกลาง
   □ เพศ, วันเกิด
   □ เลือก location (tree dropdown — ช้า)
   □ ID tag number

3. กรอกฟอร์ม hrm_human_resource (12 fields):
   □ เลือกองค์กร (autocomplete — ช้าถ้ามี org เยอะ)
   □ เลือก site (autocomplete)
   □ เลือก job title
   □ วันเริ่มงาน
   □ status

4. เพิ่ม components (optional แต่บ่อย):
   □ Skills (multi-select)
   □ Certificates
   □ Contact info (โทร, email)
   □ Address

5. กด Save → รอ 3-5 วินาที
6. ถ้าต้องการสร้างบัญชี login → ไป /eden/default/user/register อีกครั้ง

ปัญหา:
• Field เยอะ → เจ้าหน้าที่ที่รับสมัครในสนามพิมพ์ไม่ทัน
• Location dropdown ช้า (tree traversal ทุกครั้ง)
• Organization autocomplete ช้า (search ใน db ทุก keystroke)
• ต้องกรอกหลาย form แยกกัน → ข้อมูลไม่ครบ
• ไม่มี OCR → กรอกจากบัตรประชาชนเอง
• ไม่มี voice → พิมพ์ด้วยมือในสนามยาก
• ไม่มี bulk import → ทีละคน
• Offline ไม่ได้ → ถ้าอินเทอร์เน็ตไม่ดีในพื้นที่ภัยพิบัติ''', language="text")

add_heading_h2(doc, "3.2  Pain Points Matrix")

add_table(doc,
    headers=["Pain Point", "ผลกระทบ", "ความถี่", "Severity"],
    rows=[
        ["Field เยอะ (30+ ต่อ record)",      "เสียเวลา 8-15 นาที/คน",            "ทุกครั้ง",            "สูง"],
        ["Location dropdown ช้า",              "รอ 5-10 วินาทีต่อเลือก",            "ทุก record",         "กลาง"],
        ["Org autocomplete ช้า",               "รอ 1-3 วินาทีต่อ keystroke",        "ทุก record",         "กลาง"],
        ["หลาย form แยกกัน",                  "ข้อมูลไม่ครบ, ต้องกรอกซ้ำ",          "ทุกครั้ง",            "สูง"],
        ["ไม่มี OCR",                          "กรอกจากบัตรเอง → typo",            "ทุกครั้ง",            "สูง"],
        ["ไม่มี voice input",                  "พิมพ์ในสนามยาก (มือเปียกฝน)",      "ทุกครั้งในสนาม",     "สูง"],
        ["ไม่มี bulk import",                  "100 อาสาสมัคร = 15+ ชั่วโมง",      "ภัยพิบัติใหญ่",      "สูงมาก"],
        ["Offline ไม่ได้",                     "อินเทอร์เน็ตไม่ดี → กรอกไม่ได้",    "พื้นที่ห่างไกล",     "สูงมาก"],
        ["ไม่มี PDPA consent flow",            "เสี่ยง legal",                     "ทุกครั้ง",            "กลาง"],
        ["ไม่มี duplicate detection",          "กรอกซ้ำซ้อน",                      "บ่อย",               "กลาง"],
        ["Photo ID card ไม่ถูกเก็บ/verify",   "ไม่ audit trail",                  "—",                  "กลาง"],
        ["Skill/Certificate กรอก manual",     "ไม่ verify",                       "บ่อย",               "กลาง"],
        ["Multi-language form (Thai/English)", "UI ไม่รองรับ field ไทย-อังกฤษ",   "พื้นที่ต่างจังหวัด", "กลาง"],
        ["ไม่มง mobile-responsive",            "ต้องใช้ laptop",                   "สนาม",               "สูง"],
        ["ไม่มง event-scoped registration",    "ไม่รู้ว่าอาสาสมัครคนนี้อยู่ event ไหน","ภัยพิบัติใหญ่",    "กลาง"],
    ],
    col_widths_cm=[4.5, 5.0, 3.0, 2.5],
    first_col_bold=True, font_size=8.5, header_size=9)

add_heading_h2(doc, "3.3  เป้าหมายการปรับปรุง (Quantified)")

add_table(doc,
    headers=["Metric", "ปัจจุบัน", "เป้าหมายหลัง Phase 5", "วิธีวัด"],
    rows=[
        ["เวลาลงทะเบียน 1 คน",         "8-15 นาที",      "2-5 นาที",        "Stopwatch ใน field test"],
        ["ช่องทางนำข้อมูลเข้า",          "1 (manual form)","5 (form, OCR, voice, SMS, bulk)", "นับ input channels"],
        ["Error rate (typo)",            "~15%",          "< 3%",            "Audit 100 records"],
        ["Offline capability",           "ไม่ได้",         "PWA + sync",      "Test ในพื้นที่ no-signal"],
        ["Bulk import (100 คน)",         "15+ ชั่วโมง",   "10 นาที",         "Benchmark import"],
        ["PDPA consent compliance",      "ไม่มี",          "100%",            "Audit consent logs"],
        ["Duplicate detection",          "ไม่มี",          "auto-warn",       "Test with seeded dups"],
        ["Mobile-responsive",            "ไม่",            "PWA installable", "Lighthouse audit"],
        ["Field auto-fill จาก OCR",      "ไม่มี",          "ID card → 8 fields","OCR test บัตร 100 ใบ"],
        ["Voice registration (Thai)",    "ไม่มี",          "ทำได้",            "ASR accuracy test"],
    ],
    col_widths_cm=[4.0, 3.5, 4.5, 4.0],
    first_col_bold=True, font_size=8.5, header_size=9)

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
add_run(p, "แผนปรับปรุงเชิงกลยุทธ์ (Phased Plan)",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "5 phases ใน 6 เดือน — เน้น AI-powered registration เป็นหัวใจ",
        font_name=FONT_BODY, size_pt=11, italic=True, color=C_SECOND)

doc.add_page_break()

# ============================================================
# บทที่ 4: หลักการและเป้าหมาย
# ============================================================
add_heading_h1(doc, "หลักการและเป้าหมายเชิงกลยุทธ์", num=4)

add_heading_h2(doc, "4.1  หลักการ 5 ข้อ")

add_bullet(doc, "ทีละส่วน (Strangler Pattern) — ไม่ rewrite ทั้งหมด, คง business continuity", bold_prefix="① Incremental: ")
add_bullet(doc, "เริ่มจาก pain point สูงสุดก่อน (Registration) — ค่อยไป frontend, mobile", bold_prefix="② Pain-first: ")
add_bullet(doc, "AI ช่วย pre-fill + validate — ไม่ใช่ AI ตัดสินใจแทน (human-in-the-loop)", bold_prefix="③ AI-assist: ")
add_bullet(doc, "Offline-first สำหรับพื้นที่ภัยพิบัติ — Sync เมื่อกลับ online", bold_prefix="④ Offline-capable: ")
add_bullet(doc, "PDPA consent ทุก channel — audit log + retention policy", bold_prefix="⑤ PDPA-native: ")

add_heading_h2(doc, "4.2  เป้าหมายเชิงกลยุทธ์ (6 เดือน)")

add_table(doc,
    headers=["เป้าหมาย", "ค่าเป้าหมาย", "Phase ที่บรรลุ"],
    rows=[
        ["ลดเวลาลงทะเบียน staff/volunteer",      "8-15 นาที → 2-5 นาที",          "Phase 2"],
        ["รองรับ input channels 5 รูปแบบ",         "form + OCR + voice + SMS + bulk", "Phase 4"],
        ["Offline-first PWA",                       "ทำงานได้ไม่มีอินเทอร์เน็ต",        "Phase 4"],
        ["Migrate web2py → FastAPI 50%",            "AI modules + auth + org + hrm",   "Phase 5"],
        ["Database migration SQLite → PostgreSQL",  "Production-ready",                 "Phase 1"],
        ["AI accuracy (OCR บัตร ปชช.)",             "> 95% field accuracy",             "Phase 2"],
        ["AI accuracy (Voice Thai ASR)",            "> 90% word accuracy",              "Phase 2"],
        ["Test coverage",                            "80% สำหรับ AI modules",            "Phase 1"],
        ["PDPA consent capture rate",                "100%",                             "Phase 2"],
        ["Duplicate detection precision",            "> 90%",                            "Phase 3"],
    ],
    col_widths_cm=[6.5, 5.5, 4.0],
    first_col_bold=True)

add_heading_h2(doc, "4.3  Phase Overview")

add_table(doc,
    headers=["Phase", "ระยะเวลา", "Theme", "ผลลัพธ์หลัก"],
    rows=[
        ["1",  "เดือน 1",    "Foundation — infrastructure พื้นฐาน",     "Docker, PostgreSQL, FastAPI skeleton, CI/CD"],
        ["2",  "เดือน 2-3",  "AI Registration MVP — OCR + Voice",       "Feature 1 + Feature 2 (บัตรประชาชน + voice Thai)"],
        ["3",  "เดือน 4",    "Smart Forms + Bulk Import",                "Feature 3 + Feature 4 (SMS + NLQ bulk)"],
        ["4",  "เดือน 5",    "Mobile PWA + Messaging Integration",       "Feature 5 + PWA + Line/Telegram"],
        ["5",  "เดือน 6",    "Production Migration + Decommission",      "FastAPI 50%, web2py legacy ลดลง"],
    ],
    col_widths_cm=[1.0, 2.5, 5.0, 7.5],
    first_col_bold=True)

# ============================================================
# บทที่ 5: Phase 1
# ============================================================
add_heading_h1(doc, "Phase 1: Foundation (เดือน 1)", num=5)

add_heading_h2(doc, "5.1  เป้าหมาย Phase 1")
add_bullet(doc, "ตั้ง infrastructure พื้นฐานที่ production-ready")
add_bullet(doc, "Migrate database SQLite → PostgreSQL 16 + pgvector")
add_bullet(doc, "สร้าง FastAPI skeleton สำหรับ endpoints ใหม่")
add_bullet(doc, "ตั้ง CI/CD + Docker + observability")

add_heading_h2(doc, "5.2  Tasks ใน Phase 1")

add_table(doc,
    headers=["Task ID", "งาน", "Deliverable", "วันทำการ"],
    rows=[
        ["P1.1", "Dockerfile + docker-compose.yml",            "eden image + PostgreSQL + Redis + Caddy",  "2"],
        ["P1.2", "Migrate SQLite → PostgreSQL 16",              "pg_dump → pg_restore + verification script", "2"],
        ["P1.3", "Enable pgvector + migrate embeddings schema",  "ai_knowledge_chunk.embedding_vec vector(1536)", "1"],
        ["P1.4", "FastAPI skeleton (app/main.py + config.py)",   "FastAPI container รันพร้อม OpenAPI /docs", "2"],
        ["P1.5", "GitHub Actions CI (lint+test+build)",          "Auto CI ทุก push", "2"],
        ["P1.6", "Pre-commit hooks (ruff/black/mypy)",            "Auto format + lint ทุก commit", "1"],
        ["P1.7", "Health endpoint + structured logging",          "/eden/health + JSON logs", "1"],
        ["P1.8", "Sentry + Prometheus + Grafana setup",           "Error tracking + dashboards", "2"],
        ["P1.9", "Alembic migrations setup",                       "alembic upgrade head ทำงาน", "1"],
        ["P1.10","Test suite scaffold (pytest + httpx)",           "tests/ directory + conftest.py", "2"],
        ["P1.11","Documentation site (MkDocs Material)",            "docs site on GitHub Pages", "2"],
    ],
    col_widths_cm=[1.2, 6.0, 6.0, 2.0],
    first_col_bold=True, font_size=9, header_size=9.5)

add_heading_h2(doc, "5.3  Phase 1 Exit Criteria")
add_bullet(doc, "✓ docker compose up รันระบบเต็มได้ภายใน 5 นาที")
add_bullet(doc, "✓ PostgreSQL 16 + pgvector ใช้งานได้ พร้อม vector search")
add_bullet(doc, "✓ FastAPI container รัน + OpenAPI docs ที่ /eden/docs")
add_bullet(doc, "✓ CI auto-run ทุก push, lint+test+build")
add_bullet(doc, "✓ /eden/health endpoint ตอบ 200 OK")
add_bullet(doc, "✓ Sentry รับ error, Prometheus scrape /metrics")
add_bullet(doc, "✓ Alembic migration สำหรับ schema ปัจจุบัน")
add_bullet(doc, "✓ Test coverage > 30% (foundation only)")

# ============================================================
# บทที่ 6: Phase 2
# ============================================================
add_heading_h1(doc, "Phase 2: AI Registration MVP (เดือน 2-3)", num=6)

add_heading_h2(doc, "6.1  เป้าหมาย Phase 2")
add_bullet(doc, "ลดเวลาลงทะเบียนเจ้าหน้าที่/จิตอาสาจาก 8-15 นาที → 2-5 นาที")
add_bullet(doc, "เพิ่ม 2 input channels: OCR (บัตรประชาชน) + Voice (Thai)")
add_bullet(doc, "Auto-fill form จาก OCR/Voice → user แค่ตรวจและกด Save")
add_bullet(doc, "PDPA consent flow ครบทุก channel")

add_heading_h2(doc, "6.2  Tasks ใน Phase 2")

add_table(doc,
    headers=["Task ID", "งาน", "Deliverable", "วันทำการ"],
    rows=[
        ["P2.1",  "FastAPI router: /eden/api/v2/register/staff",     "POST endpoint สร้าง staff record", "3"],
        ["P2.2",  "FastAPI router: /eden/api/v2/register/volunteer", "POST endpoint สร้าง volunteer", "2"],
        ["P2.3",  "Feature 1: OCR-based registration",                "อัปโหลดบัตรประชาชน → pre-fill 8 fields", "5"],
        ["P2.4",  "Feature 2: Voice-based registration (Thai)",       "พูด → transcribe → pre-fill form", "5"],
        ["P2.5",  "Smart Form UI (HTMX + Alpine.js)",                  "Form ที่ auto-fill + validate real-time", "3"],
        ["P2.6",  "PDPA consent capture + audit log",                  "Consent checkbox + log in ai_chat_log style", "2"],
        ["P2.7",  "Duplicate detection (person)",                      "Check ชื่อ+นามสกุล+วันเกิด ก่อน create", "2"],
        ["P2.8",  "Location autocomplete (Typeahead)",                  "Fast search gis_location tree", "2"],
        ["P2.9",  "Org autocomplete (cached)",                         "Redis cache ชื่อองค์กร", "1"],
        ["P2.10", "Test suite (pytest) — 80% coverage AI modules",     "Unit + integration tests", "3"],
        ["P2.11", "Field test ในพื้นที่จริง (1 วัน)",                   "Real volunteers + ID cards", "1"],
    ],
    col_widths_cm=[1.2, 6.0, 6.0, 2.0],
    first_col_bold=True, font_size=9, header_size=9.5)

add_heading_h2(doc, "6.3  Phase 2 Exit Criteria")
add_bullet(doc, "✓ POST /eden/api/v2/register/staff ทำงาน — สร้าง pr_person + hrm_human_resource + auth_user")
add_bullet(doc, "✓ OCR บัตรประชาชนไทย — accuracy > 95% (test 100 ใบ)")
add_bullet(doc, "✓ Voice registration ภาษาไทย — accuracy > 90% (test 50 recordings)")
add_bullet(doc, "✓ Smart form auto-fill จาก OCR — ลด manual input จาก 30 fields → 5 (เช็ค + แก้)")
add_bullet(doc, "✓ PDPA consent log 100%")
add_bullet(doc, "✓ Duplicate detection — warn ถ้าเจอคนเดียวกัน")
add_bullet(doc, "✓ Test coverage > 80% สำหรับ AI modules")
add_bullet(doc, "✓ Field test จริง 1 วัน — ลงทะเบียน 20 อาสาสมัครจริง")

# ============================================================
# บทที่ 7: Phase 3
# ============================================================
add_heading_h1(doc, "Phase 3: Smart Forms + Bulk Import (เดือน 4)", num=7)

add_heading_h2(doc, "7.1  เป้าหมาย Phase 3")
add_bullet(doc, "เพิ่ม 2 input channels: SMS/Line bot + Natural Language bulk import")
add_bullet(doc, "Smart Disaster Event Hub — ผูกทุก registration เข้ากับ event")
add_bullet(doc, "Bulk import 100 อาสาสมัคร ใน 10 นาที")

add_heading_h2(doc, "7.2  Tasks ใน Phase 3")

add_table(doc,
    headers=["Task ID", "งาน", "Deliverable", "วันทำการ"],
    rows=[
        ["P3.1",  "Feature 3: SMS/Line/Telegram bot registration",      "Webhook รับข้อความ → parse → create record", "5"],
        ["P3.2",  "Feature 4: NLQ-based bulk import",                     "Paste Excel → NLQ → bulk create", "4"],
        ["P3.3",  "Excel/CSV import template",                            "Template + parser + validation", "3"],
        ["P3.4",  "Bulk import with progress tracking",                   "Celery job + WebSocket progress", "3"],
        ["P3.5",  "Feature 5: Disaster Event Hub",                        "event_event → link orgs/staff/shelters", "5"],
        ["P3.6",  "Event-scoped registration (event_id field)",           "เลือก event ก่อน register", "2"],
        ["P3.7",  "Auto-suggest related resources",                       "register volunteer → suggest nearby shelter", "3"],
        ["P3.8",  "Smart form context awareness",                          "Form ปรับ field ตาม event type", "3"],
        ["P3.9",  "Field test ในพื้นที่ 2 — bulk + event",                 "Test ทีมตอบโต้ภัยพิบัติจริง", "1"],
    ],
    col_widths_cm=[1.2, 6.0, 6.0, 2.0],
    first_col_bold=True, font_size=9, header_size=9.5)

add_heading_h2(doc, "7.3  Phase 3 Exit Criteria")
add_bullet(doc, "✓ SMS bot: ส่ง 'REGISTER ชื่อ นามสกุล วันเกิด' → สร้าง record")
add_bullet(doc, "✓ Line OA bot: ส่งรูปบัตร → auto-extract → สร้าง record")
add_bullet(doc, "✓ Bulk import Excel 100 rows ใน < 10 นาที (with progress)")
add_bullet(doc, "✓ NLQ: 'นำเข้าอาสาสมัครจากไฟล์ volunteers.xlsx' → bulk import")
add_bullet(doc, "✓ Disaster Event Hub: สร้าง event → เชื่อม orgs/staff/shelters")
add_bullet(doc, "✓ Auto-suggest: ลงทะเบียน → suggest nearby shelter/skills")

# ============================================================
# บทที่ 8: Phase 4
# ============================================================
add_heading_h1(doc, "Phase 4: Mobile PWA + Messaging Integration (เดือน 5)", num=8)

add_heading_h2(doc, "8.1  เป้าหมาย Phase 4")
add_bullet(doc, "PWA installable + offline-first — ทำงานในพื้นที่ no-signal")
add_bullet(doc, "Background sync — บันทึก offline, sync เมื่อกลับ online")
add_bullet(doc, "Push notifications — CAP alerts ไปยัง mobile")
add_bullet(doc, "Native mobile app (React Native) — Phase 5 ขยายต่อ")

add_heading_h2(doc, "8.2  Tasks ใน Phase 4")

add_table(doc,
    headers=["Task ID", "งาน", "Deliverable", "วันทำการ"],
    rows=[
        ["P4.1",  "PWA manifest + Service Worker",                    "installable + offline cache", "3"],
        ["P4.2",  "IndexedDB for offline storage",                     "บันทึก records offline", "3"],
        ["P4.3",  "Background Sync API",                                "Auto-sync เมื่อ online", "3"],
        ["P4.4",  "Mobile-responsive UI (Tailwind)",                    "Lighthouse > 90", "3"],
        ["P4.5",  "Push notifications (CAP alerts)",                    "Web Push API + FCM", "3"],
        ["P4.6",  "QR code check-in สำหรับ staff/volunteer",            "สแกน QR → check-in ที่ event", "2"],
        ["P4.7",  "Camera integration (HTML5)",                         "ถ่ายบัตร/ถ่ายภาพ field", "2"],
        ["P4.8",  "Field test ในพื้นที่ 3 — PWA + offline",              "จำลองพื้นที่ no-signal", "1"],
    ],
    col_widths_cm=[1.2, 6.0, 6.0, 2.0],
    first_col_bold=True, font_size=9, header_size=9.5)

add_heading_h2(doc, "8.3  Phase 4 Exit Criteria")
add_bullet(doc, "✓ PWA installable (Add to Home Screen)")
add_bullet(doc, "✓ Offline-first — ลงทะเบียนได้ไม่มีอินเทอร์เน็ต")
add_bullet(doc, "✓ Background sync — auto-sync เมื่อกลับ online")
add_bullet(doc, "✓ Push notifications — CAP alert ส่งไป mobile")
add_bullet(doc, "✓ Lighthouse score > 90 (Performance, Accessibility, Best Practices, SEO)")
add_bullet(doc, "✓ QR check-in — สแกน 1 วินาที ต่อคน")

# ============================================================
# บทที่ 9: Phase 5
# ============================================================
add_heading_h1(doc, "Phase 5: Production Migration (เดือน 6)", num=9)

add_heading_h2(doc, "9.1  เป้าหมาย Phase 5")
add_bullet(doc, "Migrate web2py → FastAPI 50% (auth + org + hrm + AI modules)")
add_bullet(doc, "Decommission web2py สำหรับ endpoints ที่ migrate แล้ว")
add_bullet(doc, "Production deploy — Kubernetes + Helm")
add_bullet(doc, "Security audit + penetration test")

add_heading_h2(doc, "9.2  Tasks ใน Phase 5")

add_table(doc,
    headers=["Task ID", "งาน", "Deliverable", "วันทำการ"],
    rows=[
        ["P5.1",  "Migrate auth module (login/JWT/2FA)",                "FastAPI auth + JWT", "4"],
        ["P5.2",  "Migrate org module (CRUD + OpenAPI)",                 "/eden/api/v2/org/* endpoints", "5"],
        ["P5.3",  "Migrate hrm module (staff CRUD)",                     "/eden/api/v2/hrm/* endpoints", "5"],
        ["P5.4",  "Migrate pr module (person registry)",                 "/eden/api/v2/pr/* endpoints", "3"],
        ["P5.5",  "Caddy strangler routing",                              "Route /api/v2/* → FastAPI", "1"],
        ["P5.6",  "Kubernetes Helm chart",                                 "helm install eden", "3"],
        ["P5.7",  "2FA for admin accounts",                                "TOTP + QR code", "2"],
        ["P5.8",  "Column-level encryption (PDPA)",                       "Fernet สำหรับ sensitive fields", "2"],
        ["P5.9",  "Penetration testing (external)",                        "OWASP Top 10 audit", "3"],
        ["P5.10", "PDPA DPIA documentation",                               "DPIA report", "2"],
        ["P5.11", "Production deploy + monitoring",                        "Live on K8s", "2"],
        ["P5.12", "Decommission web2py (50% endpoints)",                  "Caddy routes ลบ web2py paths", "1"],
    ],
    col_widths_cm=[1.2, 6.0, 6.0, 2.0],
    first_col_bold=True, font_size=9, header_size=9.5)

add_heading_h2(doc, "9.3  Phase 5 Exit Criteria")
add_bullet(doc, "✓ /eden/api/v2/auth/* + org/* + hrm/* + pr/* ทำงานบน FastAPI")
add_bullet(doc, "✓ JWT auth + 2FA สำหรับ admin")
add_bullet(doc, "✓ Kubernetes production cluster running")
add_bullet(doc, "✓ Pen test passed (no critical vulnerabilities)")
add_bullet(doc, "✓ PDPA DPIA complete")
add_bullet(doc, "✓ web2py ให้บริการเฉพาะ endpoints ที่ยังไม่ migrate")
add_bullet(doc, "✓ Sentry + Prometheus + Grafana monitoring production")

doc.add_page_break()

# ============================================================
# PART C DIVIDER
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(80)
add_run(p, "PART C", font_name=FONT_HEADING, size_pt=42, bold=True, color=C_PRIMARY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(20)
add_run(p, "AI-Powered Registration Features (Concrete Design)",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "5 features ที่ออกแบบเพื่อแก้ pain points ในการลงทะเบียน",
        font_name=FONT_BODY, size_pt=11, italic=True, color=C_SECOND)

doc.add_page_break()

# ============================================================
# บทที่ 10: Feature 1 - OCR Registration
# ============================================================
add_heading_h1(doc, "Feature 1: OCR-Based Staff/Volunteer Registration", num=10)

add_heading_h2(doc, "10.1  Use Case Description")

add_para(doc,
    "เจ้าหน้าที่รับสมัครในพื้นที่ถ่ายรูปบัตรประชาชนไทยของอาสาสมัคร "
    "ระบบจะ OCR ด้วย OpenThai 2.1 (iApp) หรือ vision LLM (fallback) "
    "ดึงข้อมูล 8 fields (ชื่อ-นามสกุน ไทย/อังกฤษ, เลขบัตร, ที่อยู่, วันเกิด) "
    "auto-fill form → user ตรวจและแก้ → กด Save → สร้าง pr_person + hrm_human_resource",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "10.2  OCR Flow Architecture")

add_code_block(doc, '''┌──────────────┐
│ Mobile/Web   │   User photographs ID card
│ Camera       │   → upload to /eden/api/v2/register/ocr
└──────┬───────┘
       │ multipart/form-data
       ▼
┌──────────────────────────────────────────────────────────────┐
│ FastAPI Endpoint                                                │
│   POST /eden/api/v2/register/ocr                                 │
│   • Verify PDPA consent (request body)                          │
│   • Check rate limit (30 uploads/min/IP)                        │
│   • Save image to temp storage (auto-delete after extraction)   │
└──────┬─────────────────────────────────────────────────────────┘
       │
       ▼
┌──────────────────────────────────────────────────────────────┐
│ OCR Engine Selection (auto)                                     │
│   if document_type == "id_card" and IAPP_API_KEY:              │
│     → OpenThaiClient.extract_id_card(image)                     │
│   else:                                                          │
│     → AIClient.describe_image(image, prompt="Extract fields")   │
└──────┬─────────────────────────────────────────────────────────┘
       │ structured JSON
       ▼
┌──────────────────────────────────────────────────────────────┐
│ Field Validation                                                │
│   • Validate เลขบัตร 13 หลัก (checksum)                        │
│   • Parse วันเกิด (Buddhist → Gregorian)                        │
│   • Normalize ที่อยู่ (split เลขที่/ตำบล/อำเภอ/จังหวัด)            │
│   • Geocode ที่อยู่ → gis_location.id                            │
└──────┬─────────────────────────────────────────────────────────┘
       │ pre-fill data
       ▼
┌──────────────────────────────────────────────────────────────┐
│ Smart Form (HTMX + Alpine.js)                                   │
│   • Display extracted fields (editable)                         │
│   • Highlight low-confidence fields (yellow)                    │
│   • Auto-suggest organization (if disaster event active)        │
│   • Check duplicate (name + DOB + ID)                           │
│   • On submit → POST /eden/api/v2/register/staff                │
└────────────────────────────────────────────────────────────────┘
       │
       ▼
┌──────────────────────────────────────────────────────────────┐
│ Audit Log (ai_ocr_job table)                                    │
│   • user_id, document_type, status, consent_given               │
│   • extracted_data (encrypted with Fernet)                      │
│   • Image deleted after extraction                              │
└────────────────────────────────────────────────────────────────┘''', language="text")

add_heading_h2(doc, "10.3  Code: OCR Endpoint (FastAPI)")

add_code_block(doc, '''# app/api/routers/register_ocr.py

from fastapi import APIRouter, UploadFile, File, Form, HTTPException, Depends
from fastapi.responses import JSONResponse
from typing import Optional
import io
import base64
import json

from app.api.dependencies.core import DBSessionDep, RateLimiterDep
from app.api.dependencies.auth import CurrentUserDep
from app.services.ai_ocr_openthai import OpenThaiClient, OpenThaiError
from app.services.ai_client import AsyncAIClient
from app.services.geocoder import ThaiGeocoder
from app.core.security import FieldEncryptor
from app.models.ai import AIOcrJob
from app.schemas.register import OCRResponse, RegisterStaffRequest

router = APIRouter(prefix="/eden/api/v2/register", tags=["Registration"])


@router.post("/ocr", response_model=OCRResponse)
async def extract_id_card(
    file: UploadFile = File(...),
    document_type: str = Form("id_card"),
    consent: bool = Form(..., description="PDPA consent (required)"),
    db: DBSessionDep = None,
    rate_limiter: RateLimiterDep = None,
    current_user: CurrentUserDep = None,
):
    """
    Extract structured data from ID card image.
    
    - **file**: Image file (jpg/png)
    - **document_type**: 'id_card' (default), 'volunteer_id', 'staff_id'
    - **consent**: PDPA consent — must be true
    
    Returns structured fields for pre-filling registration form.
    Image is deleted after extraction (no PII stored at-rest beyond JSON output).
    """
    # PDPA consent check
    if not consent:
        raise HTTPException(
            status_code=403,
            detail="PDPA consent required to process ID card image",
        )
    
    # Rate limit: 30 uploads/min/IP
    await rate_limiter.check_or_raise(
        key=f"ocr:{current_user.id if current_user else 'anon'}",
        limit=30,
        window=60,
    )
    
    # Read image
    image_bytes = await file.read()
    if len(image_bytes) > 5 * 1024 * 1024:  # 5MB max
        raise HTTPException(
            status_code=413,
            detail="Image too large (max 5MB)",
        )
    
    # OCR: try OpenThai first, fall back to vision LLM
    extracted = {}
    engine = ""
    
    try:
        # Primary: OpenThai 2.1 for Thai ID card
        if document_type == "id_card":
            client = OpenThaiClient()  # reads IAPP_API_KEY from env
            extracted = await client.extract_id_card(image_bytes)
            engine = "openthai-2.1"
        else:
            raise OpenThaiError("not id_card type")
    
    except (OpenThaiError, Exception) as e:
        # Fallback: vision LLM
        client = AsyncAIClient(api_key=settings.zai_api_key)
        prompt = """Extract these fields from the Thai national ID card image:
        - title_th, first_name_th, last_name_th
        - first_name_en, last_name_en
        - id_card_number (13 digits)
        - address (Thai)
        - birth_date (Buddhist year)
        Return as JSON."""
        
        # Convert image to base64 for vision API
        image_b64 = base64.b64encode(image_bytes).decode()
        response = await client.describe_image(
            image_b64=image_b64,
            prompt=prompt,
            model="glm-4v-plus",
        )
        try:
            extracted = json.loads(response)
            engine = "glm-4v-plus"
        except json.JSONDecodeError:
            raise HTTPException(
                status_code=502,
                detail="OCR failed — could not parse response",
            )
    
    # Validate + normalize extracted data
    validated = await _validate_extracted(extracted)
    
    # Geocode address (if present)
    location_id = None
    if validated.get("address"):
        geocoder = ThaiGeocoder()
        location_id = await geocoder.geocode_thai_address(validated["address"])
    
    # Check duplicate (เลขบัตร or name+DOB)
    duplicate = await _check_duplicate(db, validated)
    
    # Log to ai_ocr_job (encrypted)
    encryptor = FieldEncryptor()
    job = AIOcrJob(
        user_id=current_user.id if current_user else None,
        document_type=document_type,
        extracted_data=encryptor.encrypt(json.dumps(extracted)),  # encrypted
        target_table="pr_person",
        status="done",
        consent_given=True,
        ip=current_user.ip if hasattr(current_user, "ip") else None,
        model=engine,
    )
    db.add(job)
    await db.commit()
    
    # Image is NOT stored — only extracted JSON
    
    return OCRResponse(
        fields=validated,
        location_id=location_id,
        duplicate_warning=duplicate,
        engine=engine,
        confidence_scores=_get_confidence_scores(extracted),
        log_id=job.id,
    )


async def _validate_extracted(data: dict) -> dict:
    """Validate and normalize extracted fields."""
    validated = {}
    
    # เลขบัตร 13 หลัก — validate checksum
    id_num = data.get("id_card_number", "").replace(" ", "").replace("-", "")
    if len(id_num) == 13 and _validate_thai_id(id_num):
        validated["national_id"] = id_num
    
    # ชื่อ-นามสกุล
    validated["first_name_th"] = data.get("first_name_th", "")
    validated["last_name_th"] = data.get("last_name_th", "")
    validated["first_name_en"] = data.get("first_name_en", "")
    validated["last_name_en"] = data.get("last_name_en", "")
    
    # วันเกิด (Buddhist → Gregorian)
    birth_date = data.get("birth_date")
    if birth_date:
        validated["date_of_birth"] = _buddhist_to_gregorian(birth_date)
    
    # ที่อยู่
    validated["address"] = data.get("address", "")
    
    return validated


def _validate_thai_id(id_num: str) -> bool:
    """Validate Thai national ID checksum (13 digits)."""
    if len(id_num) != 13 or not id_num.isdigit():
        return False
    total = 0
    for i in range(12):
        total += int(id_num[i]) * (13 - i)
    check = (11 - (total % 11)) % 10
    return check == int(id_num[12])


def _buddhist_to_gregorian(date_str: str) -> str:
    """Convert Buddhist year (2550) to Gregorian (2007)."""
    # Implementation depends on input format
    # e.g., "25 มี.ค. 2535" → "1992-03-25"
    ...


async def _check_duplicate(db, fields: dict) -> Optional[dict]:
    """Check if person already exists."""
    from sqlalchemy import select, or_, and_
    from app.models.pr import Person
    
    if fields.get("national_id"):
        stmt = select(Person).where(
            Person.national_id == fields["national_id"],
            Person.deleted == False,
        )
        result = await db.execute(stmt)
        if existing := result.scalar_one_or_none():
            return {
                "type": "exact_id_match",
                "existing_id": existing.id,
                "name": f"{existing.first_name} {existing.last_name}",
            }
    
    # Fuzzy match: name + DOB
    if fields.get("first_name_th") and fields.get("last_name_th") and fields.get("date_of_birth"):
        stmt = select(Person).where(
            and_(
                Person.first_name.ilike(fields["first_name_th"]),
                Person.last_name.ilike(fields["last_name_th"]),
                Person.date_of_birth == fields["date_of_birth"],
                Person.deleted == False,
            )
        )
        result = await db.execute(stmt)
        if existing := result.scalar_one_or_none():
            return {
                "type": "name_dob_match",
                "existing_id": existing.id,
                "name": f"{existing.first_name} {existing.last_name}",
            }
    
    return None''', language="python")

add_heading_h2(doc, "10.4  Smart Form UI (HTMX)")

add_code_block(doc, '''<!-- app/templates/register/staff_form.html -->
{% extends "base.html" %}

{% block content %}
<h1>ลงทะเบียนเจ้าหน้าที่/อาสาสมัคร</h1>

<!-- Step 1: Upload ID card -->
<div x-data="{ ocrDone: false, extracting: false, fields: {} }">
  
  <div class="card mb-4">
    <h3>ขั้นตอนที่ 1: ถ่ายรูปบัตรประชาชน</h3>
    
    <!-- PDPA consent -->
    <label class="flex items-center gap-2 mb-3">
      <input type="checkbox" x-model="consent" required>
      <span class="text-sm">
        ยินยอมให้ประมวลผลรูปบัตรประชาชน (PDPA) — 
        รูปจะถูกลบหลัง extraction
      </span>
    </label>
    
    <!-- Camera/upload -->
    <input type="file" 
           accept="image/*" 
           capture="environment"
           @change="uploadCard($event)"
           class="mb-2">
    
    <div x-show="extracting" class="text-blue-600">
      ⏳ กำลังประมวลผล OCR...
    </div>
    
    <!-- Preview extracted fields -->
    <div x-show="ocrDone" class="mt-4">
      <h4>ข้อมูลที่ดึงได้ (ตรวจสอบและแก้ไข้ได้):</h4>
      
      <div class="grid grid-cols-2 gap-3">
        <div>
          <label>ชื่อ (ไทย)</label>
          <input type="text" x-model="fields.first_name_th" 
                 :class="fields.first_name_th_confidence < 0.8 ? 'bg-yellow-100' : ''">
        </div>
        <div>
          <label>นามสกุล (ไทย)</label>
          <input type="text" x-model="fields.last_name_th">
        </div>
        <!-- ... more fields ... -->
      </div>
      
      <div x-show="duplicate_warning" class="bg-red-100 p-3 mt-3">
        ⚠️ พบข้อมูลซ้ำ: <span x-text="duplicate_warning?.name"></span>
        (<a :href="'/eden/pr/person/' + duplicate_warning?.existing_id">เปิดดู</a>)
      </div>
    </div>
  </div>
  
  <!-- Step 2: Additional fields (manual) -->
  <div x-show="ocrDone" class="card mb-4">
    <h3>ขั้นตอนที่ 2: ข้อมูลเพิ่มเติม</h3>
    
    <div class="grid grid-cols-2 gap-3">
      <div>
        <label>องค์กร</label>
        <select name="organisation_id" 
                hx-get="/eden/api/v2/org/search"
                hx-trigger="keyup changed delay:300ms"
                hx-target="#org-suggestions">
        </select>
        <div id="org-suggestions"></div>
      </div>
      
      <div>
        <label>ตำแหน่ง</label>
        <select name="job_title_id">
          {% for job in job_titles %}
          <option value="{{ job.id }}">{{ job.name }}</option>
          {% endfor %}
        </select>
      </div>
      
      <div>
        <label>วันเริ่มงาน</label>
        <input type="date" name="start_date" value="{{ today }}">
      </div>
      
      <div>
        <label>ประเภท</label>
        <select name="type">
          <option value="staff">เจ้าหน้าที่</option>
          <option value="volunteer">อาสาสมัคร</option>
          <option value="trainee">ฝึกงาน</option>
        </select>
      </div>
    </div>
  </div>
  
  <!-- Step 3: Skills (optional) -->
  <div x-show="ocrDone" class="card mb-4">
    <h3>ขั้นตอนที่ 3: ทักษะ (ถ้ามี)</h3>
    <input type="text" placeholder="ค้นหาทักษะ..."
           hx-get="/eden/api/v2/hrm/skills/search"
           hx-trigger="keyup changed delay:300ms"
           hx-target="#skill-list">
    <div id="skill-list"></div>
  </div>
  
  <!-- Submit -->
  <button type="submit" 
          x-show="ocrDone"
          hx-post="/eden/api/v2/register/staff"
          hx-include="form"
          hx-target="#result"
          class="btn btn-primary">
    บันทึกการลงทะเบียน
  </button>
  
  <div id="result"></div>
</div>

<script>
async function uploadCard(event) {
  const file = event.target.files[0];
  if (!file) return;
  
  extracting.value = true;
  
  const formData = new FormData();
  formData.append('file', file);
  formData.append('document_type', 'id_card');
  formData.append('consent', 'true');
  
  try {
    const response = await fetch('/eden/api/v2/register/ocr', {
      method: 'POST',
      body: formData,
    });
    
    if (response.ok) {
      const data = await response.json();
      fields.value = data.fields;
      duplicate_warning.value = data.duplicate_warning;
      ocrDone.value = true;
    } else {
      alert('OCR failed: ' + (await response.text()));
    }
  } catch (e) {
    alert('Error: ' + e.message);
  } finally {
    extracting.value = false;
  }
}
</script>
{% endblock %}''', language="html")

# ============================================================
# บทที่ 11: Voice Registration
# ============================================================
add_heading_h1(doc, "Feature 2: Voice-Based Registration (Thai)", num=11)

add_heading_h2(doc, "11.1  Use Case Description")

add_para(doc,
    "เจ้าหน้าที่ในสนาม (มือเปียกฝน, สวมถุงมือ) ไม่สะดวกพิมพ์ "
    "พูดง่ายกว่า: 'ฉันชื่อสมชาย ใจดี เกิด 15 มีนาคม 2535 อาสาสมัครจากมูลนิธิกระจกเงา' "
    "ระบบ transcribe ด้วย ASR (z-ai) → parse intent → pre-fill form "
    "user ตรวจและแก้ → Save",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "11.2  Voice Registration Flow")

add_code_block(doc, '''┌──────────────┐
│ Mobile/Web   │   User clicks "Record" button, speaks
│ Microphone   │   → records 5-30 seconds
└──────┬───────┘
       │ audio/webm (or wav)
       ▼
┌──────────────────────────────────────────────────────────────┐
│ FastAPI Endpoint                                                │
│   POST /eden/api/v2/register/voice                              │
│   • Verify PDPA consent                                         │
│   • Check rate limit (10/min — voice is heavier)               │
│   • Save audio to temp (auto-delete after ASR)                 │
└──────┬─────────────────────────────────────────────────────────┘
       │
       ▼
┌──────────────────────────────────────────────────────────────┐
│ ASR (z-ai transcribe_audio)                                     │
│   • Auto-detect language (Thai expected)                        │
│   • Return transcript text                                      │
│   • Confidence score per segment                                │
└──────┬─────────────────────────────────────────────────────────┘
       │ transcript text (Thai)
       ▼
┌──────────────────────────────────────────────────────────────┐
│ Intent Extraction (LLM)                                         │
│   Prompt: "Extract registration fields from Thai text:"         │
│   {transcript}                                                  │
│   Return JSON: {                                                │
│     "first_name": "...",                                         │
│     "last_name": "...",                                          │
│     "date_of_birth": "YYYY-MM-DD",                              │
│     "organisation": "...",                                      │
│     "role": "volunteer|staff",                                  │
│     "skills": ["..."],                                           │
│   }                                                              │
└──────┬─────────────────────────────────────────────────────────┘
       │ structured data
       ▼
┌──────────────────────────────────────────────────────────────┐
│ Smart Form (same as Feature 1)                                  │
│   • Display extracted fields                                    │
│   • User can re-record specific fields                          │
│   • On submit → POST /eden/api/v2/register/staff                │
└────────────────────────────────────────────────────────────────┘''', language="text")

add_heading_h2(doc, "11.3  Code: Voice Endpoint (FastAPI)")

add_code_block(doc, '''# app/api/routers/register_voice.py

from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from typing import Optional
import json

from app.api.dependencies.core import DBSessionDep, RateLimiterDep
from app.api.dependencies.auth import CurrentUserDep
from app.services.ai_client import AsyncAIClient
from app.services.ai_voice import ThaiNameParser, DateParser
from app.models.ai import AIVoiceLog
from app.schemas.register import VoiceResponse

router = APIRouter(prefix="/eden/api/v2/register", tags=["Voice Registration"])


@router.post("/voice", response_model=VoiceResponse)
async def voice_register(
    audio: UploadFile = File(...),
    consent: bool = Form(...),
    db: DBSessionDep = None,
    rate_limiter: RateLimiterDep = None,
    current_user: CurrentUserDep = None,
):
    """
    Register staff/volunteer via voice (Thai).
    
    Audio format: WAV, MP3, WebM
    Returns extracted fields for form pre-fill.
    """
    if not consent:
        raise HTTPException(status_code=403, detail="PDPA consent required")
    
    await rate_limiter.check_or_raise(
        key=f"voice:{current_user.id if current_user else 'anon'}",
        limit=10,
        window=60,
    )
    
    audio_bytes = await audio.read()
    if len(audio_bytes) > 5 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="Audio too large (max 5MB)")
    
    client = AsyncAIClient(api_key=settings.zai_api_key)
    
    # Step 1: ASR
    transcript = await client.transcribe_audio(audio_bytes)
    
    # Step 2: Intent extraction via LLM
    extraction_prompt = f"""You are a registration assistant for Sahana Eden disaster relief system.
Extract registration fields from this Thai voice transcript.

Transcript: "{transcript}"

Return JSON with these fields (use null if not mentioned):
{{
  "first_name_th": "ชื่อไทย",
  "last_name_th": "นามสกุลไทย",
  "first_name_en": "English first name",
  "last_name_en": "English last name",
  "date_of_birth": "YYYY-MM-DD (convert Buddhist year if needed)",
  "organisation": "name of organisation mentioned",
  "role": "staff|volunteer|trainee",
  "skills": ["list of skills mentioned"],
  "phone": "phone number if mentioned",
  "email": "email if mentioned"
}}

Only include fields that were clearly mentioned. Be conservative — better to return null than guess."""
    
    response = await client.chat(
        prompt=extraction_prompt,
        system="You are a precise Thai-language data extraction assistant.",
        model="glm-4-plus",
        temperature=0.1,  # low temperature for precision
    )
    
    try:
        extracted = json.loads(response)
    except json.JSONDecodeError:
        # Try to extract JSON from response
        import re
        match = re.search(r'\\{[^{}]+\\}', response, re.DOTALL)
        if match:
            extracted = json.loads(match.group())
        else:
            raise HTTPException(
                status_code=502,
                detail="Could not parse voice extraction response",
            )
    
    # Validate Thai name + parse date
    if extracted.get("first_name_th"):
        extracted["first_name_th"] = ThaiNameParser.normalize(extracted["first_name_th"])
    
    if extracted.get("date_of_birth"):
        extracted["date_of_birth"] = DateParser.parse_thai_date(extracted["date_of_birth"])
    
    # Check duplicate
    duplicate = await _check_duplicate(db, extracted)
    
    # Log to ai_voice_log (transcript + extracted)
    log = AIVoiceLog(
        user_id=current_user.id if current_user else None,
        transcript=transcript,
        intent="register",
        response_text=json.dumps(extracted, ensure_ascii=False),
        language="th",
        ip=current_user.ip if hasattr(current_user, "ip") else None,
    )
    db.add(log)
    await db.commit()
    
    # Audio is NOT stored permanently
    
    return VoiceResponse(
        transcript=transcript,
        fields=extracted,
        duplicate_warning=duplicate,
        log_id=log.id,
    )


# Thai utilities
class ThaiNameParser:
    @staticmethod
    def normalize(name: str) -> str:
        """Normalize Thai name (remove titles, extra spaces)."""
        titles = ["นาย", "นาง", "นางสาว", "เด็กชาย", "เด็กหญิง"]
        for title in titles:
            if name.startswith(title):
                return name[len(title):].strip()
        return name.strip()


class DateParser:
    THAI_MONTHS = {
        "ม.ค.": 1, "มี.ค.": 2, "เม.ย.": 4, "พ.ค.": 5, "มิ.ย.": 6,
        "ก.ค.": 7, "ส.ค.": 8, "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12,
        "มกราคม": 1, "กุมภาพันธ์": 2, "มีนาคม": 3, "เมษายน": 4,
        "พฤษภาคม": 5, "มิถุนายน": 6, "กรกฎาคม": 7, "สิงหาคม": 8,
        "กันยายน": 9, "ตุลาคม": 10, "พฤศจิกายน": 11, "ธันวาคม": 12,
    }
    
    @classmethod
    def parse_thai_date(cls, date_str: str) -> Optional[str]:
        """Parse Thai date string to YYYY-MM-DD (Gregorian)."""
        import re
        from datetime import date
        
        # Try formats: "15 มี.ค. 2535", "15/3/2535", "15-03-2535"
        # Buddhist year (2535) → Gregorian (1992): subtract 543
        
        for pattern, fmt in [
            (r'(\\d{1,2})\\s+([\\u0E00-\\u0E7F\\.]+)\\s+(\\d{4})', 'thai_month'),
            (r'(\\d{1,2})/(\\d{1,2})/(\\d{4})', 'numeric'),
            (r'(\\d{4})-(\\d{2})-(\\d{2})', 'iso'),
        ]:
            match = re.match(pattern, date_str.strip())
            if match:
                if fmt == 'thai_month':
                    day, month_str, year = int(match.group(1)), match.group(2), int(match.group(3))
                    month = cls.THAI_MONTHS.get(month_str)
                    if month:
                        if year > 2500:  # Buddhist
                            year -= 543
                        return f"{year:04d}-{month:02d}-{day:02d}"
                elif fmt == 'numeric':
                    day, month, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    if year > 2500:
                        year -= 543
                    return f"{year:04d}-{month:02d}-{day:02d}"
                elif fmt == 'iso':
                    return date_str  # already ISO
        
        return None''', language="python")

# ============================================================
# บทที่ 12: SMS/Line/Telegram
# ============================================================
add_heading_h1(doc, "Feature 3: SMS/Line/Telegram Bot Registration", num=12)

add_heading_h2(doc, "12.1  Use Case Description")

add_para(doc,
    "อาสาสมัครส่ง SMS หรือข้อความ Line/Telegram สั้น ๆ "
    "เช่น 'REGISTER สมชาย ใจดี 15/03/2535 มูลนิธิกระจกเงา' "
    "หรือส่งรูปบัตรประชาชนผ่าน Line → bot auto-respond + ข้อมูลเสร็จสมบูรณ์ "
    "เหมาะกับผู้ที่ไม่สะดวกใช้เว็บ หรืออยู่ในพื้นที่ที่ mobile signal แย่",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "12.2  Multi-Channel Bot Architecture")

add_code_block(doc, '''┌──────────────────────────────────────────────────────────────────┐
│ Channel Adapters (รับ message จาก platform ต่าง ๆ)                 │
│                                                                    │
│  ┌────────────┐  ┌────────────┐  ┌────────────┐  ┌────────────┐  │
│  │ SMS        │  │ Line OA    │  │ Telegram   │  │ Web Form   │  │
│  │ (Twilio)   │  │ Webhook    │  │ Bot API    │  │            │  │
│  └─────┬──────┘  └─────┬──────┘  └─────┬──────┘  └─────┬──────┘  │
└────────┼───────────────┼───────────────┼───────────────┼──────────┘
         │               │               │               │
         ▼               ▼               ▼               ▼
┌──────────────────────────────────────────────────────────────────┐
│ Unified Message Queue (Redis Streams)                              │
│   • Normalize: { channel, sender_id, text, image?, timestamp }   │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Message Processor (Celery worker)                                  │
│   1. Detect intent (REGISTER / QUERY / HELP)                       │
│   2. If REGISTER: extract fields (text or image OCR)               │
│   3. Validate + check duplicate                                    │
│   4. If missing fields: ask follow-up via same channel             │
│   5. If complete: create record + reply confirmation               │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Eden Database (FastAPI + SQLAlchemy)                               │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Reply to user via original channel                                 │
│   "✅ ลงทะเบียนสำเร็จ! ID: 12345 / ดูรายละเอียด: eden.example/person/12345" │
└──────────────────────────────────────────────────────────────────┘''', language="text")

add_heading_h2(doc, "12.3  Code: Line Webhook (FastAPI)")

add_code_block(doc, '''# app/api/routers/line_webhook.py

from fastapi import APIRouter, Request, HTTPException, BackgroundTasks
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, ImageMessage,
    TextSendMessage, TemplateSendMessage, ConfirmTemplate,
    PostbackAction,
)
import os
import json

router = APIRouter(prefix="/eden/webhooks/line", tags=["Line OA"])

line_bot_api = LineBotApi(os.environ["LINE_CHANNEL_ACCESS_TOKEN"])
handler = WebhookHandler(os.environ["LINE_CHANNEL_SECRET"])


@router.post("/")
async def line_webhook(request: Request, background_tasks: BackgroundTasks):
    """Receive Line webhook events."""
    signature = request.headers.get("X-Line-Signature", "")
    body = await request.body()
    
    try:
        handler.handle(body.decode(), signature)
    except InvalidSignatureError:
        raise HTTPException(status_code=403, detail="Invalid signature")
    
    return "OK"


@handler.add(MessageEvent, message=TextMessage)
def handle_text_message(event):
    """Process text messages from users."""
    text = event.message.text.strip()
    user_id = event.source.user_id
    
    if text.upper().startswith("REGISTER"):
        # Process registration in background
        import asyncio
        asyncio.create_task(_process_registration_text(user_id, text))
    
    elif text.upper().startswith("QUERY"):
        # NLQ query
        asyncio.create_task(_process_query(user_id, text[6:]))
    
    elif text.upper() == "HELP":
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=(
                "คำสั่งที่ใช้ได้:\\n"
                "REGISTER [ชื่อ] [นามสกุล] [วันเกิด DD/MM/YYYY] [องค์กร]\\n"
                "QUERY [คำถาม]\\n"
                "หรือส่งรูปบัตรประชาชนเพื่อลงทะเบียนอัตโนมัติ"
            ))
    
    else:
        # Free text — try intent detection
        asyncio.create_task(_detect_intent_and_respond(user_id, text))


@handler.add(MessageEvent, message=ImageMessage)
def handle_image_message(event):
    """Process image (ID card) from users."""
    message_id = event.message.id
    
    # Get image content
    image_content = line_bot_api.get_message_content(message_id)
    image_bytes = image_content.content
    
    # Process in background (async OCR)
    import asyncio
    asyncio.create_task(_process_id_card_image(
        user_id=event.source.user_id,
        image_bytes=image_bytes,
        channel="line",
    ))


async def _process_registration_text(user_id: str, text: str):
    """Parse 'REGISTER ชื่อ นามสกุล DD/MM/YYYY องค์กร' format."""
    from app.services.register_parser import RegistrationParser
    
    try:
        fields = RegistrationParser.parse_sms_format(text[8:].strip())
        
        # Validate required fields
        missing = RegistrationParser.get_missing_fields(fields)
        if missing:
            line_bot_api.push_message(
                user_id,
                TextSendMessage(text=(
                    f"ข้อมูลยังไม่ครบ ขา: {', '.join(missing)}\\n"
                    f"รูปแบบ: REGISTER [ชื่อ] [นามสกุล] [วันเกิด DD/MM/YYYY] [องค์กร]"
                ))
            )
            return
        
        # Check duplicate
        from app.crud.person import check_duplicate
        from app.core.database import sessionmanager
        
        async with sessionmanager.session() as db:
            duplicate = await check_duplicate(db, fields)
            if duplicate:
                line_bot_api.push_message(
                    user_id,
                    TextSendMessage(text=(
                        f"⚠️ พบข้อมูลซ้ำกับ: {duplicate['name']}\\n"
                        f"หากไม่ใช่คนเดียวกัน กรุณาเพิ่มรหัสบัตรประชาชน"
                    ))
                )
                return
            
            # Create record
            from app.crud.staff import create_staff_from_fields
            staff = await create_staff_from_fields(db, fields, source_channel="line")
            await db.commit()
            
            line_bot_api.push_message(
                user_id,
                TextSendMessage(text=(
                    f"✅ ลงทะเบียนสำเร็จ!\\n"
                    f"ID: {staff.id}\\n"
                    f"ชื่อ: {staff.person.first_name} {staff.person.last_name}\\n"
                    f"องค์กร: {staff.organisation.name if staff.organisation else '-'}\\n"
                    f"ดูรายละเอียด: https://eden.example.org/eden/hrm/staff/{staff.id}"
                ))
    
    except Exception as e:
        line_bot_api.push_message(
            user_id,
            TextSendMessage(text=f"เกิดข้อผิดพลาด: {str(e)}"),
        )''', language="python")

# ============================================================
# บทที่ 13: NLQ Bulk Import
# ============================================================
add_heading_h1(doc, "Feature 4: Bulk Import via Natural Language", num=13)

add_heading_h2(doc, "13.1  Use Case Description")

add_para(doc,
    "เจ้าหน้าที่มี Excel/CSV ที่มี 100+ อาสาสมัคร — ไม่ต้อง manual import "
    "แค่พูดหรือพิมพ์: 'นำเข้าอาสาสมัครจากไฟล์ volunteers.xlsx' "
    "หรือ paste ตาราง: 'สมชาย ใจดี 15/03/2535 มูลนิธิกระจกเงา\\nสมหญิง รักไทย 20/06/2540 ...' "
    "AI parse → validate → bulk create พร้อม progress bar",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "13.2  NLQ Bulk Import Flow")

add_code_block(doc, '''┌──────────────────────────────────────────────────────────────────┐
│ User input (3 ways):                                              │
│                                                                    │
│  A. NLQ + file: "นำเข้าอาสาสมัครจากไฟล์ volunteers.xlsx"            │
│     + upload file                                                  │
│                                                                    │
│  B. Paste table:                                                   │
│     "สมชาย ใจดี 15/03/2535 มูลนิธิกระจกเงา"                         │
│     "สมหญิง รักไทย 20/06/2540 มูลนิธิกระจกเงา"                     │
│     "..."                                                          │
│                                                                    │
│  C. Excel template download + upload                               │
│     → /eden/api/v2/register/template.xlsx                          │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ FastAPI Endpoint                                                   │
│   POST /eden/api/v2/register/bulk                                  │
│   • Accept file OR pasted text                                     │
│   • Admin only (require_admin dependency)                          │
│   • Create Celery job — return job_id immediately                  │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Celery Worker (background)                                         │
│   1. Parse file (openpyxl/pandas) or text                          │
│   2. For each row:                                                 │
│      - Validate fields                                             │
│      - Check duplicate                                             │
│      - Create pr_person + hrm_human_resource                       │
│      - Track success/skip/error                                    │
│   3. Update progress via WebSocket (or SSE)                        │
│   4. Generate report (CSV with results)                            │
└──────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌──────────────────────────────────────────────────────────────────┐
│ Real-time UI (WebSocket/SSE)                                       │
│   "กำลังนำเข้า: 45/100 (สำเร็จ 40, ซ้ำ 3, error 2)"                │
│   "เสร็จสิ้น! ดาวน์โหลดรายงาน: import_report_20261119.csv"        │
└──────────────────────────────────────────────────────────────────┘''', language="text")

add_heading_h2(doc, "13.3  Code: Bulk Import Endpoint + Celery Task")

add_code_block(doc, '''# app/api/routers/register_bulk.py

from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from fastapi.responses import StreamingResponse
import io
import csv
import asyncio

from app.api.dependencies.core import DBSessionDep
from app.api.dependencies.auth import AdminUserDep
from app.tasks.celery_app import bulk_register_task
from app.schemas.register import BulkImportResponse, BulkImportStatus

router = APIRouter(prefix="/eden/api/v2/register", tags=["Bulk Import"])


@router.post("/bulk", response_model=BulkImportResponse)
async def start_bulk_import(
    file: UploadFile = File(None),
    pasted_text: str = None,
    current_user: AdminUserDep = None,
):
    """
    Start bulk import (admin only).
    
    Returns job_id immediately — use /bulk/status/{job_id} for progress.
    """
    if not file and not pasted_text:
        raise HTTPException(status_code=400, detail="file or pasted_text required")
    
    # Read data
    if file:
        content = await file.read()
        filename = file.filename
    else:
        content = pasted_text.encode()
        filename = "pasted.txt"
    
    # Dispatch Celery task
    task = bulk_register_task.delay(
        user_id=current_user.id,
        filename=filename,
        content_base64=content.decode('latin-1'),  # safe transport
    )
    
    return BulkImportResponse(
        job_id=task.id,
        status="queued",
        estimated_seconds=len(content) // 100,  # rough estimate
    )


@router.get("/bulk/status/{job_id}", response_model=BulkImportStatus)
async def bulk_status(job_id: str, current_user: AdminUserDep = None):
    """Get bulk import progress."""
    from celery.result import AsyncResult
    from app.tasks.celery_app import app as celery_app
    
    result = AsyncResult(job_id, app=celery_app)
    
    if result.state == "PENDING":
        return BulkImportStatus(state="pending", progress=0)
    elif result.state == "PROGRESS":
        return BulkImportStatus(
            state="progress",
            progress=result.info.get("current", 0) / result.info.get("total", 1),
            current=result.info.get("current"),
            total=result.info.get("total"),
            success=result.info.get("success", 0),
            duplicates=result.info.get("duplicates", 0),
            errors=result.info.get("errors", 0),
        )
    elif result.state == "SUCCESS":
        info = result.info
        return BulkImportStatus(
            state="done",
            progress=1.0,
            total=info["total"],
            success=info["success"],
            duplicates=info["duplicates"],
            errors=info["errors"],
            report_url=f"/eden/api/v2/register/bulk/report/{job_id}",
        )
    else:  # FAILURE
        return BulkImportStatus(state="failed", error=str(result.info))


@router.get("/bulk/report/{job_id}")
async def bulk_report(job_id: str, current_user: AdminUserDep = None):
    """Download CSV report of bulk import."""
    from celery.result import AsyncResult
    from app.tasks.celery_app import app as celery_app
    
    result = AsyncResult(job_id, app=celery_app)
    if result.state != "SUCCESS":
        raise HTTPException(status_code=404, detail="Report not ready")
    
    info = result.info
    rows = info["rows"]  # list of {row_num, status, message}
    
    # Generate CSV
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["row", "status", "name", "message", "created_id"])
    for row in rows:
        writer.writerow([
            row["row_num"], row["status"], row.get("name", ""),
            row.get("message", ""), row.get("created_id", ""),
        ])
    
    output.seek(0)
    return StreamingResponse(
        io.BytesIO(output.getvalue().encode()),
        media_type="text/csv",
        headers={
            "Content-Disposition": f"attachment; filename=import_report_{job_id}.csv"
        },
    )


# ─────────────────────────────────────────────────────────────────────────
# Celery task

# app/tasks/bulk_register.py
from app.tasks.celery_app import app
from app.core.database import sessionmanager
from app.crud.staff import create_staff_from_fields
from app.crud.person import check_duplicate
import openpyxl
import csv
import io
import base64

@app.task(bind=True)
def bulk_register_task(self, user_id: int, filename: str, content_base64: str):
    """Bulk register staff/volunteers from Excel or CSV."""
    
    content = base64.b64decode(content_base64.encode())
    
    # Parse based on extension
    rows_data = []
    if filename.endswith(".xlsx"):
        wb = openpyxl.load_workbook(io.BytesIO(content))
        ws = wb.active
        headers = [str(cell.value).lower().strip() if cell.value else "" for cell in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_dict = dict(zip(headers, row))
            rows_data.append(row_dict)
    elif filename.endswith(".csv"):
        text = content.decode("utf-8")
        reader = csv.DictReader(io.StringIO(text))
        for row in reader:
            rows_data.append({k.lower().strip(): v for k, v in row.items()})
    else:
        # Plain text — parse line by line (space-separated)
        text = content.decode("utf-8")
        for line in text.strip().split("\\n"):
            parts = line.split()
            if len(parts) >= 4:
                rows_data.append({
                    "first_name": parts[0],
                    "last_name": parts[1],
                    "date_of_birth": parts[2],
                    "organisation": parts[3],
                })
    
    total = len(rows_data)
    success = 0
    duplicates = 0
    errors = 0
    result_rows = []
    
    import asyncio
    loop = asyncio.new_event_loop()
    
    async def process_all():
        nonlocal success, duplicates, errors
        
        async with sessionmanager.session() as db:
            for i, row_data in enumerate(rows_data, 1):
                row_result = {"row_num": i, "name": ""}
                
                try:
                    # Normalize fields
                    fields = _normalize_row(row_data)
                    row_result["name"] = f"{fields.get('first_name_th', '')} {fields.get('last_name_th', '')}"
                    
                    # Check duplicate
                    dup = await check_duplicate(db, fields)
                    if dup:
                        duplicates += 1
                        row_result.update(
                            status="duplicate",
                            message=f"Duplicate of {dup['name']} (ID: {dup['existing_id']})",
                        )
                    else:
                        # Create record
                        staff = await create_staff_from_fields(
                            db, fields, source_channel="bulk_import", user_id=user_id
                        )
                        await db.flush()  # get ID without commit
                        success += 1
                        row_result.update(
                            status="created",
                            created_id=staff.id,
                            message="OK",
                        )
                
                except Exception as e:
                    errors += 1
                    row_result.update(status="error", message=str(e))
                
                result_rows.append(row_result)
                
                # Update progress every 5 rows
                if i % 5 == 0 or i == total:
                    self.update_state(
                        state="PROGRESS",
                        meta={
                            "current": i,
                            "total": total,
                            "success": success,
                            "duplicates": duplicates,
                            "errors": errors,
                        },
                    )
            
            await db.commit()
    
    loop.run_until_complete(process_all())
    loop.close()
    
    return {
        "total": total,
        "success": success,
        "duplicates": duplicates,
        "errors": errors,
        "rows": result_rows,
    }


def _normalize_row(row: dict) -> dict:
    """Normalize row dict to standard fields."""
    return {
        "first_name_th": row.get("first_name") or row.get("ชื่อ") or row.get("name"),
        "last_name_th": row.get("last_name") or row.get("นามสกุล") or row.get("surname"),
        "date_of_birth": row.get("date_of_birth") or row.get("วันเกิด") or row.get("dob"),
        "organisation": row.get("organisation") or row.get("องค์กร") or row.get("org"),
        "phone": row.get("phone") or row.get("เบอร์"),
        "email": row.get("email"),
    }''', language="python")

# ============================================================
# บทที่ 14: Disaster Event Hub
# ============================================================
add_heading_h1(doc, "Feature 5: Smart Disaster Event Hub", num=14)

add_heading_h2(doc, "14.1  Use Case Description")

add_para(doc,
    "ปัจจุบัน Eden มี event module แต่ coupling หลวม — registrations ไม่ถูกผูกกับ event "
    "ทำให้ดู 'ใน event นี้มี staff กี่คน, อาสาสมัครกี่คน, ศูนย์พักพิงกี่แห่ง' ไม่ได้ "
    "Feature นี้เพิ่ม event-scoped context ให้ทุก registration "
    "พร้อม auto-suggest related resources",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "14.2  Event Hub Data Model")

add_code_block(doc, '''┌──────────────────────────────────────────────────────────────────┐
│  event_event (Disaster Event Hub — NEW enhanced role)             │
│  ─────────────────────────────────────                            │
│  • id, name (e.g., "อุทกภัยเหนือ 2567")                             │
│  • event_type_id ───► event_event_type                            │
│  • start_date, end_date                                           │
│  • status (active/closed/archived)                                │
│  • severity (1-5)                                                 │
│                                                                    │
│  ★ NEW fields:                                                    │
│  • default_organisation_id ─► org_organisation                    │
│  • default_location_id ─────► gis_location (center of disaster)   │
│  • cap_alert_template_id ───► cap_alert                           │
└──────────────────────────────────────────────────────────────────┘
                              │
       ┌──────────┬──────────┼──────────┬──────────┐
       │          │          │          │          │
       ▼          ▼          ▼          ▼          ▼
┌────────────┐ ┌────────────┐ ┌────────────┐ ┌────────────┐ ┌────────────┐
│ hrm_human  │ │ cr_shelter │ │ inv_ware   │ │ asset_asset│ │ req_req    │
│ _resource  │ │            │ │ house      │ │            │ │            │
│            │ │            │ │            │ │            │ │            │
│ ★ NEW:     │ │ ★ NEW:     │ │ ★ NEW:     │ │ ★ NEW:     │ │ ★ NEW:     │
│ event_id   │ │ event_id   │ │ event_id   │ │ event_id   │ │ event_id   │
└────────────┘ └────────────┘ └────────────┘ └────────────┘ └────────────┘
                              │
                              ▼
       When user creates any record during active event:
       1. Form auto-fills event_id (current active event)
       2. UI shows: "ลงทะเบียนใน event: อุทกภัยเหนือ 2567"
       3. Auto-suggest: nearby shelters, available staff, etc.''', language="text")

add_heading_h2(doc, "14.3  Code: Event Hub Service")

add_code_block(doc, '''# app/services/event_hub.py

from sqlalchemy import select, and_
from sqlalchemy.ext.asyncio import AsyncSession
from typing import Optional
from datetime import datetime

from app.models.event import Event
from app.models.org import Organisation
from app.models.cr import Shelter
from app.models.hrm import HumanResource
from app.models.gis import Location


class EventHub:
    """Smart Disaster Event Hub — central context for registrations."""
    
    def __init__(self, db: AsyncSession):
        self.db = db
    
    async def get_active_event(self) -> Optional[Event]:
        """Get the currently active disaster event."""
        stmt = select(Event).where(
            and_(
                Event.status == "active",
                Event.start_date <= datetime.utcnow(),
                (Event.end_date.is_(None)) | (Event.end_date >= datetime.utcnow()),
            )
        ).order_by(Event.start_date.desc()).limit(1)
        result = await self.db.execute(stmt)
        return result.scalar_one_or_none()
    
    async def get_event_summary(self, event_id: int) -> dict:
        """Get summary of resources for an event."""
        # Count staff assigned to event
        staff_count = await self.db.scalar(
            select(func.count(HumanResource.id)).where(
                HumanResource.event_id == event_id
            )
        )
        
        # Count shelters active in event
        shelters_count = await self.db.scalar(
            select(func.count(Shelter.id)).where(
                Shelter.event_id == event_id
            )
        )
        
        # Total people in shelters
        shelter_population = await self.db.scalar(
            select(func.sum(Shelter.population)).where(
                Shelter.event_id == event_id
            )
        )
        
        return {
            "event_id": event_id,
            "staff_count": staff_count or 0,
            "shelters_count": shelters_count or 0,
            "shelter_population": shelter_population or 0,
        }
    
    async def suggest_resources_for_registration(
        self, event_id: int, person_location_id: Optional[int] = None
    ) -> dict:
        """Suggest related resources when registering a person in event."""
        suggestions = {}
        
        # Suggest nearby shelters (within 50km)
        if person_location_id:
            nearby_shelters = await self._find_nearby_shelters(
                event_id, person_location_id, radius_km=50
            )
            suggestions["nearby_shelters"] = nearby_shelters[:3]
        
        # Suggest organisations involved in event
        orgs = await self._get_event_organisations(event_id)
        suggestions["organisations"] = orgs[:5]
        
        # Suggest skills needed (based on event type)
        event = await self.db.get(Event, event_id)
        if event:
            needed_skills = await self._get_needed_skills(event.event_type_id)
            suggestions["needed_skills"] = needed_skills
        
        return suggestions
    
    async def _find_nearby_shelters(
        self, event_id: int, location_id: int, radius_km: float
    ) -> list[dict]:
        """Find shelters within radius using PostGIS."""
        stmt = """
            SELECT s.id, s.name, s.population, s.capacity,
                   ST_Distance(l.geom, loc.geom) / 1000 as distance_km
            FROM cr_shelter s
            JOIN gis_location l ON s.site_id = l.id
            CROSS JOIN gis_location loc
            WHERE loc.id = :location_id
              AND s.event_id = :event_id
              AND s.deleted = false
              AND ST_DWithin(l.geom, loc.geom, :radius_meters)
            ORDER BY distance_km
            LIMIT 5
        """
        result = await self.db.execute(
            stmt,
            {
                "location_id": location_id,
                "event_id": event_id,
                "radius_meters": radius_km * 1000,
            }
        )
        return [dict(r) for r in result]
    
    async def _get_event_organisations(self, event_id: int) -> list[dict]:
        """Get organisations that have staff/resources in event."""
        stmt = """
            SELECT DISTINCT o.id, o.name, o.acronym,
                   COUNT(h.id) as staff_count
            FROM org_organisation o
            JOIN hrm_human_resource h ON h.organisation_id = o.id
            WHERE h.event_id = :event_id AND h.deleted = false
            GROUP BY o.id, o.name, o.acronym
            ORDER BY staff_count DESC
        """
        result = await self.db.execute(stmt, {"event_id": event_id})
        return [dict(r) for r in result]


# Usage in registration flow:
#   event_hub = EventHub(db)
#   active_event = await event_hub.get_active_event()
#   if active_event:
#       suggestions = await event_hub.suggest_resources_for_registration(
#           active_event.id, person.location_id
#       )
#       return {"suggestions": suggestions, "event": active_event}''', language="python")

doc.add_page_break()

# ============================================================
# PART D DIVIDER
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(80)
add_run(p, "PART D", font_name=FONT_HEADING, size_pt=42, bold=True, color=C_PRIMARY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(20)
add_run(p, "บันทึกการปรับปรุง (Change Log)",
        font_name=FONT_HEADING, size_pt=22, bold=True, color=C_DARK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, "ADRs + Change Log format — track ทุกการปรับปรุงของระบบ",
        font_name=FONT_BODY, size_pt=11, italic=True, color=C_SECOND)

doc.add_page_break()

# ============================================================
# บทที่ 15: Change Log Format
# ============================================================
add_heading_h1(doc, "Change Log Format และ ADRs", num=15)

add_heading_h2(doc, "15.1  Change Log Format (Keep a Changelog)")

add_code_block(doc, '''# CHANGELOG.md

# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- **Feature 1: OCR-based staff/volunteer registration** (`POST /eden/api/v2/register/ocr`)
  - OpenThai 2.1 primary OCR engine for Thai ID cards
  - Vision LLM fallback for non-ID documents
  - Auto-fill 8 fields from ID card
  - PDPA consent flow + audit log
  - Thai ID checksum validation
  - Duplicate detection (national ID + name+DOB)
  - Geocode address → gis_location
- **Feature 2: Voice-based registration (Thai)** (`POST /eden/api/v2/register/voice`)
  - ASR via z-ai (auto-detect Thai)
  - LLM intent extraction → structured fields
  - Thai date parser (Buddhist → Gregorian)
  - Thai name normalization (remove titles)
- **Feature 3: SMS/Line/Telegram bot registration**
  - Line OA webhook (`POST /eden/webhooks/line`)
  - SMS via Twilio webhook
  - Telegram bot webhook
  - Unified message queue (Redis Streams)
  - Multi-turn conversation for missing fields
- **Feature 4: Bulk import via NLQ**
  - Excel/CSV/paste-text import
  - Celery background processing
  - Real-time progress via WebSocket
  - CSV report download
- **Feature 5: Disaster Event Hub**
  - `event_id` field added to: hrm_human_resource, cr_shelter, inv_warehouse, asset_asset, req_req
  - Auto-suggest nearby shelters (PostGIS)
  - Auto-suggest organisations in event
  - Event dashboard (staff count, shelter population)
- **Foundation**: Dockerfile, docker-compose.yml, Caddyfile, GitHub Actions CI
- **Foundation**: PostgreSQL 16 + pgvector migration
- **Foundation**: FastAPI skeleton + OpenAPI docs at /eden/docs
- **Foundation**: Alembic migrations, pre-commit hooks, pyproject.toml
- **Foundation**: Sentry + Prometheus + Grafana + structured logging
- **Foundation**: Health endpoint at /eden/health

### Changed
- Migrated `/eden/ai_chat/*` from web2py to FastAPI (strangler pattern)
- Migrated `/eden/ai_ocr/*` from web2py to FastAPI
- Migrated `/eden/ai_query/*` from web2py to FastAPI
- Migrated `/eden/ai_kb/*` from web2py to FastAPI
- Migrated `/eden/ai_voice/*` from web2py to FastAPI
- Migrated `/eden/ai_analytics/*` from web2py to FastAPI
- Migrated `/eden/ai_imagery/*` from web2py to FastAPI
- Migrated `/eden/ai_translate/*` from web2py to FastAPI
- Migrated `/eden/ai_moderation/*` from web2py to FastAPI
- Migrated `/eden/auth/*` from web2py session auth to FastAPI JWT
- Migrated `/eden/org/*` from web2py to FastAPI (CRUD + OpenAPI)
- Migrated `/eden/hrm/staff/*` from web2py to FastAPI
- Migrated `/eden/pr/person/*` from web2py to FastAPI
- Replaced jQuery AJAX with HTMX (incremental)
- Replaced Foundation 5 with Tailwind CSS
- Replaced OpenLayers with MapLibre GL JS
- Replaced PyDAL with SQLAlchemy 2.0
- Replaced YATL templates with Jinja2
- Replaced web2py Auth with FastAPI Security (JWT + OAuth2)
- Replaced SQLite with PostgreSQL 16
- Replaced subprocess.run() with aiohttp (async AI client)
- Replaced manual scheduler with Celery + Redis
- Replaced manual forms with Pydantic v2 schemas

### Deprecated
- web2py controllers for migrated endpoints (to be removed in Phase 6)
- SQLite database (replaced by PostgreSQL)
- jQuery 1.x (replaced by HTMX + Alpine.js)
- Foundation 5 (replaced by Tailwind)
- YATL templates (replaced by Jinja2)

### Removed
- Twitter integration (already removed in Eden v6.2)
- Facebook integration (already removed in Eden v6.2)

### Security
- Added CSP headers via Caddy
- Added 2FA (TOTP) for admin accounts
- Added column-level encryption for sensitive fields (Fernet)
- Added audit log immutability (syslog forwarding)
- Added rate limiting: 60/min for chatbot, 30/min for OCR, 10/min for voice
- Added dependency scanning (pip-audit, Snyk, CodeQL) in CI
- Added secrets detection (detect-secrets) pre-commit hook
- Added PDPA consent flow for all AI features
- Added Thai ID checksum validation
- Added audit trail for OCR jobs (encrypted extracted_data)

### PDPA Compliance
- All AI features require explicit consent (chatbot, OCR, voice)
- Sensitive data encrypted at-rest (Fernet)
- Image data deleted after extraction (no PII in image storage)
- Audit logs forwarded to syslog (immutable)
- Retention: ai_chat_log 90d, ai_voice_log 90d, ai_query_log 1y, ai_ocr_job 1y
- NLQ whitelist blocks sensitive tables (auth_user, dvr_case, med_patient)

## [6.2.1] - 2026-11-19

### Added
- Initial FastAPI skeleton (app/main.py, app/core/database.py)
- Dockerfile + docker-compose.yml
- GitHub Actions CI workflow
- Pre-commit hooks (ruff, black, mypy, detect-secrets)
- pyproject.toml (centralized config)
- Health endpoint
- Structured logging (structlog)

## [6.2.0] - 2026-09-08
- Upstream Sahana Eden v6.2 (Eden ASP branch)

## [6.2.0-ai.1] - 2026-11-18
### Added
- AI modules: chatbot, OCR, analytics, imagery, translate, NLQ, KB, voice, moderation
- AI tables: ai_chat_log, ai_knowledge_chunk, ai_ocr_job, ai_forecast, ai_imagery_job,
  ai_translation_job, ai_query_log, ai_voice_log, ai_moderation_queue
- AI client wrapper (modules/s3/ai_client.py)
- OpenThai OCR client (modules/s3/ai_ocr_openthai.py)
- AI documentation (docs/ai/*.md)''', language="markdown")

add_heading_h2(doc, "15.2  Architecture Decision Records (ADRs)")

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

---

# docs/adr/0002-migrate-web2py-to-fastapi.md
# ADR 0002: Migrate web2py to FastAPI

Date: 2026-11-19

## Status
Accepted

## Context
web2py 3.3.3 is in maintenance mode. No new features, slow security patches.
Async support is poor. Developer hiring is difficult (niche skill).

## Decision
Migrate from web2py to FastAPI using the Strangler Fig Pattern.

Migration order:
1. AI modules (newest, least coupled)
2. Auth + User management
3. Read-only API (/eden/api/v2/*)
4. Write API (org, pr, hrm, inv)
5. Frontend (Jinja2 + HTMX)
6. Decommission web2py

## Consequences
- **Positive**: Modern async framework, OpenAPI docs auto-generated,
  better hiring, faster CI/CD
- **Negative**: 6-month migration effort, dual-system operation period
- **Mitigation**: Strangler Pattern ensures business continuity

---

# docs/adr/0003-adopt-postgresql-with-pgvector.md
# ADR 0003: Use PostgreSQL with pgvector

Date: 2026-11-19

## Status
Accepted

## Context
SQLite doesn't support:
- Concurrent writes (single writer)
- Vector similarity search (placeholder JSON embeddings)
- Production-grade backup/replication

## Decision
Migrate to PostgreSQL 16 + pgvector extension.

## Consequences
- **Positive**: Production-ready, vector search, replication, FTS
- **Negative**: Requires PostgreSQL operational expertise
- **Migration**: Use Alembic for schema migrations, dual-write during transition

---

# docs/adr/0004-ai-assisted-registration.md
# ADR 0004: AI-Assisted Registration (Human-in-the-Loop)

Date: 2026-11-19

## Status
Accepted

## Context
Manual registration of staff/volunteers takes 8-15 minutes per person.
In disaster zones, this delays response. Field workers need faster input.

## Decision
Implement AI-assisted registration with human-in-the-loop:
1. OCR extracts data from ID card → pre-fill form
2. Voice transcribes Thai speech → pre-fill form
3. SMS/Line/Telegram bot → parse → create record
4. Bulk import via NLQ
5. Smart form: user reviews and edits before submit

AI does NOT auto-create records without human review.

## Consequences
- **Positive**: 70% reduction in registration time, multi-channel input,
  PDPA-compliant (consent + audit)
- **Negative**: AI accuracy not 100% (mitigated by human review)
- **Risk**: PDPA — mitigated by consent flow, encrypted storage, image deletion

---

# docs/adr/0005-disaster-event-hub.md
# ADR 0005: Disaster Event Hub

Date: 2026-11-19

## Status
Accepted

## Context
Eden's event module has loose coupling with registrations. Cannot answer
"How many staff in event X?" without complex queries.

## Decision
Add `event_id` foreign key to: hrm_human_resource, cr_shelter,
inv_warehouse, asset_asset, req_req.

Create EventHub service for:
- Auto-suggest resources during registration
- Event dashboard (counts of staff, shelters, population)
- PostGIS-powered nearby shelter suggestions

## Consequences
- **Positive**: Event-scoped operations, better situational awareness
- **Negative**: Schema migration required (event_id columns)
- **Migration**: Add nullable event_id columns, backfill from event_event_location

---

# docs/adr/0006-pwa-offline-first.md
# ADR 0006: PWA Offline-First for Field Operations

Date: 2026-11-19

## Status
Accepted

## Context
Disaster zones often have unreliable internet. Field workers cannot
register people without connectivity.

## Decision
Implement PWA with:
- Service Worker for offline caching
- IndexedDB for offline record storage
- Background Sync API for auto-sync when online
- Installable (Add to Home Screen)

## Consequences
- **Positive**: Works offline, faster, installable
- **Negative**: Conflict resolution needed (mitigated by last-write-wins + manual review)''', language="markdown")

# ============================================================
# บทที่ 16: สรุป
# ============================================================
add_heading_h1(doc, "สรุปและขั้นตอนถัดไป", num=16)

add_heading_h2(doc, "16.1  สรุปแผนปรับปรุง")

add_para(doc,
    "เอกสารฉบับนี้วางแผนปรับปรุง Sahana Eden fork ใน 5 phases ภายใน 6 เดือน "
    "โดยเน้นการเพิ่ม AI สำหรับการลงทะเบียนเจ้าหน้าที่/อิตอาสาในพื้นที่ภัยพิบัติ "
    "เป็นหัวใจ — ลดเวลาจาก 8-15 นาทีเหลือ 2-5 นาที และเพิ่มช่องทางนำข้อมูลเข้า "
    "จาก 1 เป็น 5 (form, OCR, voice, SMS/bot, bulk)",
    align=WD_ALIGN_PARAGRAPH.JUSTIFY)

add_heading_h2(doc, "16.2  Phase Summary")

add_table(doc,
    headers=["Phase", "ระยะเวลา", "Features หลัก", "ลดเวลาลงทะเบียน"],
    rows=[
        ["1", "เดือน 1",    "Foundation (Docker, PostgreSQL, FastAPI, CI/CD)",     "—"],
        ["2", "เดือน 2-3",  "OCR + Voice registration",                              "8-15 นาที → 5-8 นาที"],
        ["3", "เดือน 4",    "SMS bot + Bulk import + Event Hub",                     "→ 3-5 นาที"],
        ["4", "เดือน 5",    "PWA offline + Push + QR check-in",                      "→ 2-3 นาที (offline)"],
        ["5", "เดือน 6",    "Migrate 50% web2py → FastAPI + K8s + Security audit", "—"],
    ],
    col_widths_cm=[1.2, 2.5, 8.0, 4.0],
    first_col_bold=True)

add_heading_h2(doc, "16.3  Key Improvements Summary")

add_bullet(doc, "ลดเวลาลงทะเบียน 70% (8-15 นาที → 2-5 นาที)", bold_prefix="① ")
add_bullet(doc, "5 input channels: form + OCR + voice + SMS/bot + bulk", bold_prefix="② ")
add_bullet(doc, "Offline-first PWA สำหรับพื้นที่ภัยพิบัติ", bold_prefix="③ ")
add_bullet(doc, "Disaster Event Hub ผูกทุก registration เข้ากับ event", bold_prefix="④ ")
add_bullet(doc, "PDPA compliance 100% (consent + audit + encryption)", bold_prefix="⑤ ")
add_bullet(doc, "Migrate web2py → FastAPI 50% (Strangler Pattern)", bold_prefix="⑥ ")
add_bullet(doc, "PostgreSQL 16 + pgvector (แทน SQLite)", bold_prefix="⑦ ")
add_bullet(doc, "Production-ready (Docker + K8s + CI/CD + observability)", bold_prefix="⑧ ")

add_heading_h2(doc, "16.4  ไฟล์เอกสารที่เกี่ยวข้อง")

add_table(doc,
    headers=["ไฟล์", "เนื้อหา"],
    rows=[
        ["Sahana_Eden_Strategic_Improvement_Plan.docx", "เอกสารฉบับนี้ (Part A-D)"],
        ["Sahana_Eden_Strategic_Improvement_Plan.pdf",  "PDF version"],
        ["Sahana_Eden_Migration_And_Language_Comparison.docx", "Migration + Language"],
        ["Sahana_Eden_Modernization_Recommendations.docx",     "Modernization roadmap"],
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
add_run(fp, "Sahana Eden — Strategic Improvement Plan  ·  Page ",
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
