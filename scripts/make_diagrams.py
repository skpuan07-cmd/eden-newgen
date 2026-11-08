"""Generate architecture and REST request lifecycle diagrams for the S3 Framework doc."""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle
from matplotlib.lines import Line2D

# Use Loma for Thai + DejaVu Sans for English/code
import matplotlib.font_manager as fm
fm.fontManager.addfont("/usr/share/fonts/opentype/tlwg/Loma.otf")
fm.fontManager.addfont("/usr/share/fonts/opentype/tlwg/Loma-Bold.otf")
plt.rcParams["font.family"] = ["Loma", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

OUT_DIR = "/home/z/my-project/scripts/assets"
os.makedirs(OUT_DIR, exist_ok=True)

# Palette
C_BG       = "#FFFFFF"
C_DARK     = "#1E3A8A"   # deep blue
C_PRIMARY  = "#2563EB"   # blue
C_ACCENT   = "#0E7490"   # teal
C_SECOND   = "#475569"   # slate
C_LIGHT    = "#EFF6FF"   # pale blue
C_LIGHT2   = "#F0FDF4"   # pale green
C_LIGHT3   = "#FEF3C7"   # pale amber
C_GRAY     = "#F1F5F9"
C_BORDER   = "#94A3B8"

def box(ax, x, y, w, h, text, fill=C_LIGHT, border=C_PRIMARY, text_color="#0F172A",
        fontsize=9, bold=False, radius=0.02):
    """Draw a rounded box with centered text."""
    p = FancyBboxPatch((x, y), w, h,
                       boxstyle=f"round,pad=0.005,rounding_size={radius}",
                       linewidth=1.2, edgecolor=border, facecolor=fill, zorder=2)
    ax.add_patch(p)
    ax.text(x + w/2, y + h/2, text,
            ha="center", va="center",
            fontsize=fontsize, color=text_color,
            fontweight="bold" if bold else "normal",
            zorder=3, wrap=True)

def arrow(ax, x1, y1, x2, y2, color=C_SECOND, style="->", lw=1.2):
    a = FancyArrowPatch((x1, y1), (x2, y2),
                        arrowstyle=style, mutation_scale=12,
                        color=color, lw=lw, zorder=1)
    ax.add_patch(a)

def label(ax, x, y, text, fontsize=8, color=C_SECOND, ha="center", bold=False):
    ax.text(x, y, text, ha=ha, va="center",
            fontsize=fontsize, color=color,
            fontweight="bold" if bold else "normal", zorder=4)

# ============================================================
# Diagram 1: S3 Framework Class Hierarchy
# ============================================================
fig, ax = plt.subplots(figsize=(11.5, 8.5), dpi=150)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.set_aspect("equal")
ax.axis("off")
fig.patch.set_facecolor(C_BG)

# Title
ax.text(50, 97, "S3 Framework — ผังความสัมพันธ์ระหว่างคลาสหลัก",
        ha="center", va="center", fontsize=14, fontweight="bold", color=C_DARK)
ax.text(50, 93.5, "Built on top of web2py / PyDAL",
        ha="center", va="center", fontsize=9, color=C_SECOND, style="italic")

# Layer 1: web2py (top)
box(ax, 5, 84, 90, 6,
    "web2py / PyDAL (External Framework)   ·   DAL  ·  Table  ·  Field  ·  Query  ·  Row  ·  Auth  ·  current",
    fill="#F8FAFC", border=C_SECOND, text_color=C_SECOND, fontsize=9, bold=False, radius=0.05)

arrow(ax, 50, 84, 50, 80, color=C_SECOND, lw=1.5)

# Layer 2: S3 Framework banner
box(ax, 5, 75, 90, 5,
    "S3 Framework  (modules/core/, imported as s3base)",
    fill=C_DARK, border=C_DARK, text_color="#FFFFFF", fontsize=11, bold=True, radius=0.05)

# Layer 3: 5 columns of major subsystems
y_sub = 60
sub_h = 13
sub_w = 17

cols = [
    (5,  "model/\ndatamodel.py",   "DataModel\n(อดีต S3Model)\n\n• define_table()\n• lazy loader\n• configure()\n• onvalidation\n• onaccept\n• set_method()", C_LIGHT, C_PRIMARY),
    (24, "resource/\nresource.py",  "CRUDResource\n(อดีต S3Resource)\n\n• count()/select()\n• insert()/update()\n• delete()\n• export_xml()\n• import_xml()\n• components", C_LIGHT2, "#16A34A"),
    (43, "controller.py",          "CRUDRequest\n(อดีต S3RESTController)\n\n• __parse() URL\n• __call__ dispatch\n• default_methods\n• prep / postp\n• Content-Type", C_LIGHT3, "#D97706"),
    (62, "methods/\ncrud.py",      "BasicCRUD\n(อดีต S3CRUD)\n\n• apply_method()\n• create / read\n• update / delete\n• select / list\n• validate\n• review", C_LIGHT, C_PRIMARY),
    (81, "methods/\nrest.py",      "RESTful\n\n• get_tree()\n• put_tree()\n• get_fields()\n• get_options()\n\n(XML / S3JSON /\n GeoJSON API)", C_LIGHT2, "#16A34A"),
]

for x, title, body, fill, border in cols:
    # title bar
    box(ax, x, y_sub + sub_h - 3.5, sub_w, 3.5, title,
        fill=border, border=border, text_color="#FFFFFF", fontsize=8.5, bold=True, radius=0.04)
    # body
    box(ax, x, y_sub, sub_w, sub_h - 3.5, body,
        fill=fill, border=border, text_color="#0F172A", fontsize=7.5, radius=0.04)

# arrow from S3 framework banner to each subsystem
for x, *_ in cols:
    arrow(ax, 50, 75, x + sub_w/2, y_sub + sub_h, color=C_SECOND, lw=1)

# Layer 4: supporting subsystems
y_sup = 44
sup_h = 9
sup_w = 21
support = [
    (5,  "formats/",   "FormatWriter/Parser\n• xml.py (S3XML)\n• pdf.py / card.py\n• xls.py / xlsx.py\n• shp.py / svg.py",   C_LIGHT,  C_PRIMARY),
    (28, "filters/",   "URL → Query parsers\n• OptionsFilter\n• HierarchyFilter\n• DateFilter\n• LocationFilter",        C_LIGHT,  C_PRIMARY),
    (51, "aaa/",       "Auth / Permissions\n• AuthS3\n• audit / lock\n• oauth\n• masterkey / consent",                  C_LIGHT3, "#D97706"),
    (74, "gis/ + msg/", "Mapping + Messaging\n• current.gis\n• current.msg\n• SMS / Email / Twitter",                    C_LIGHT2, "#16A34A"),
]
for x, title, body, fill, border in support:
    box(ax, x, y_sup + sup_h - 2.5, sup_w, 2.5, title,
        fill=border, border=border, text_color="#FFFFFF", fontsize=8, bold=True, radius=0.04)
    box(ax, x, y_sup, sup_w, sup_h - 2.5, body,
        fill=fill, border=border, text_color="#0F172A", fontsize=7.5, radius=0.04)

# Layer 5: bottom — DataModel modules (s3db)
y_db = 26
box(ax, 5, y_db + 12, 90, 3, "modules/s3db/*.py  —  41 DataModel modules (one per domain)",
    fill=C_ACCENT, border=C_ACCENT, text_color="#FFFFFF", fontsize=9.5, bold=True, radius=0.04)

db_modules = ["org.py", "pr.py", "inv.py", "hrm.py", "project.py", "req.py", "cr.py",
              "cap.py", "dvr.py", "supply.py", "asset.py", "...41 total"]
db_x = 6
db_w = (90 - 5 - (len(db_modules)-1)*0.3) / len(db_modules)
for i, m in enumerate(db_modules):
    x = 5 + i * (db_w + 0.3)
    box(ax, x, y_db, db_w, 4, m, fill=C_LIGHT2, border="#16A34A",
        text_color="#0F172A", fontsize=7, radius=0.04)

# Layer 6: controllers (web2py thin layer)
y_ctrl = 16
box(ax, 5, y_ctrl + 4, 90, 4,
    "controllers/*.py  —  Thin web2py controllers (1 function per URL endpoint, ~1 line each)",
    fill=C_SECOND, border=C_SECOND, text_color="#FFFFFF", fontsize=9, bold=True, radius=0.04)
box(ax, 5, y_ctrl, 90, 4,
    "def organisation(): return s3db.org_organisation_controller()\n        ↓  crud_controller(prefix, resourcename)",
    fill=C_GRAY, border=C_BORDER, text_color="#0F172A", fontsize=8, radius=0.04)

# Layer 7: configuration / settings
y_cfg = 4
box(ax, 5, y_cfg, 44, 8,
    "modules/s3cfg.py\nS3Config\n→ current.deployment_settings\n\nmodels/000_config.py\n→ server-side configuration",
    fill=C_LIGHT, border=C_PRIMARY, text_color="#0F172A", fontsize=8, radius=0.04)
box(ax, 51, y_cfg, 44, 8,
    "modules/s3dal.py\nS3DAL  (compat shim — NOT a real DAL!)\nRe-exports PyDAL objects; handles\nversion-specific API differences",
    fill="#FEE2E2", border="#DC2626", text_color="#0F172A", fontsize=8, radius=0.04)

# Arrows: controllers → CRUDRequest
arrow(ax, 50, y_ctrl, 50, y_sub + sub_h + 0.5, color=C_SECOND, lw=1.2)
label(ax, 51, y_ctrl + 5.5, "routes via", fontsize=7, color=C_SECOND, ha="left")

# Arrows: CRUDRequest → CRUDResource (instantiation)
arrow(ax, 51.5, y_sub, 32.5, y_sub, color=C_PRIMARY, lw=1.2, style="->")
label(ax, 42, y_sub + 1, "instantiates", fontsize=7, color=C_PRIMARY)

# Arrows: CRUDRequest → BasicCRUD/RESTful (handler dispatch)
arrow(ax, 62, y_sub, 71, y_sub, color=C_PRIMARY, lw=1.2, style="->")
label(ax, 66.5, y_sub + 1, "dispatch", fontsize=7, color=C_PRIMARY)

# Arrows: CRUDResource → DataModel (lazy table loader)
arrow(ax, 14, y_sub, 14, y_db + 4, color="#16A34A", lw=1.2, style="->")
label(ax, 16, 41, "s3db.<table>\n(lazy)", fontsize=7, color="#16A34A", ha="left")

# Arrows: BasicCRUD → DataModel (read hooks)
arrow(ax, 70, y_sup + sup_h, 30, y_db + 4, color="#16A34A", lw=1.0, style="->")
label(ax, 56, 45, "configure()/get_config()\nread hooks", fontsize=7, color="#16A34A")

# Arrows: BasicCRUD → CRUDResource (operate on resource)
arrow(ax, 71.5, y_sub, 32.5, y_sub - 2, color=C_PRIMARY, lw=1.0, style="->")
label(ax, 56, y_sub - 1.5, "operate on", fontsize=7, color=C_PRIMARY)

# Arrows: BasicCRUD → formats/ (export)
arrow(ax, 67, y_sup + sup_h, 15.5, y_sup + sup_h, color=C_PRIMARY, lw=1.0, style="->")
label(ax, 40, y_sup + sup_h + 1, "export via DataExporter", fontsize=7, color=C_PRIMARY)

# Arrows: Resource → filters/
arrow(ax, 28, y_sup + sup_h, 38.5, y_sup + sup_h, color="#16A34A", lw=1.0, style="->")
label(ax, 33, y_sup + sup_h + 1, "build_query", fontsize=7, color="#16A34A")

# Footer / legend
ax.text(50, 0.8,
        "อดีตใช้ชื่อ: S3Model / S3Resource / S3CRUD / S3RESTController   →   ปัจจุบัน: DataModel / CRUDResource / BasicCRUD / CRUDRequest",
        ha="center", va="center", fontsize=8, color="#DC2626", style="italic")

plt.savefig(os.path.join(OUT_DIR, "diagram_architecture.png"),
            dpi=150, bbox_inches="tight", facecolor=C_BG, pad_inches=0.3)
plt.close()
print("✓ diagram_architecture.png")

# ============================================================
# Diagram 2: REST Request Lifecycle (flowchart)
# ============================================================
fig, ax = plt.subplots(figsize=(10, 11), dpi=150)
ax.set_xlim(0, 100)
ax.set_ylim(0, 110)
ax.axis("off")
fig.patch.set_facecolor(C_BG)

ax.text(50, 107, "REST Request Lifecycle — การไหลของ Request หนึ่งรายการ",
        ha="center", va="center", fontsize=13, fontweight="bold", color=C_DARK)
ax.text(50, 103.5,
        "Example:  GET /eden/org/organisation/123/branch/45.csv?organisation_type.name=NGO",
        ha="center", va="center", fontsize=9, color=C_SECOND, style="italic",
        family="Loma")

# Flow boxes (vertical pipeline) on left
flow = [
    (96, 8, "HTTP request เข้าระบบ",   C_LIGHT3, "#D97706"),
    (88, 8, "web2py router → controllers/org.py::organisation()",  C_GRAY, C_BORDER),
    (80, 8, "controller เรียก crud_controller('org', 'organisation')", C_LIGHT, C_PRIMARY),
    (72, 8, "CRUDRequest.__parse()  →  แยก URL เป็น\nid=123, component='branch', component_id=45,\nmethod=None, representation='csv', http='GET", C_LIGHT, C_PRIMARY),
    (63, 8, "CRUDResource('org_organisation', id=123, ...)\n├─ build_query() → S3ResourceFilter → DAL Query\n└─ components = S3Components(['branch'])",  C_LIGHT2, "#16A34A"),
    (54, 8, "PRE-PROCESS: prep = s3.prep(r)\nสามารถ bypass / inject output / return False (400)",  C_LIGHT3, "#D97706"),
    (45, 8, "RESOLVE HANDLER (ตามลำดับ):\n① s3db.get_method() — custom REST method?\n② default_methods[method] (report/map/profile/...)?\n③ RESTful (ถ้าไม่มี method และ transformable)\n④ BasicCRUD (สำหรับ GET/POST ทั่วไป)\n⑤ 405 BAD_METHOD",  C_LIGHT3, "#D97706"),
    (35, 8, "handler(r, **attr)  →  เช่น BasicCRUD.apply_method()\n├─ create() / read() / update() / delete() / select()\n└─ read() + rep='csv' → DataExporter.csv(resource)",  C_LIGHT, C_PRIMARY),
    (26, 8, "CRUDResource internals:\n• count() → rfilter.count() → db(query).count()\n• select() → ResourceData → db(query).select()\n• insert()/update() → table.insert() + audit +\n                       ownership + onaccept hook",  C_LIGHT2, "#16A34A"),
    (17, 8, "POST-PROCESS: postp = s3.postp(r, output)\noutput['r'] = self  (ส่งให้ view)",  C_LIGHT3, "#D97706"),
    (9,  8, "view template (xml.html สำหรับ non-HTML)",  C_GRAY, C_BORDER),
    (2,  8, "HTTP 200 OK   Content-Type: text/csv",  C_LIGHT3, "#D97706"),
]

for y, h, text, fill, border in flow:
    box(ax, 8, y, 60, h, text, fill=fill, border=border, fontsize=8, radius=0.04)

# Vertical arrows
for i in range(len(flow) - 1):
    y1 = flow[i][0]
    y2 = flow[i+1][0] + flow[i+1][1]
    arrow(ax, 38, y1, 38, y2, color=C_SECOND, lw=1.5)

# Side panel: hooks called at each step (right side)
side_x = 72
side_w = 26

# Heading
box(ax, side_x, 96, side_w, 5, "Hooks & Extension Points\nCalled at Each Stage",
    fill=C_DARK, border=C_DARK, text_color="#FFFFFF", fontsize=9, bold=True, radius=0.04)

# Hook annotations (right side, aligned with corresponding stage)
hooks = [
    (87, "settings.customise_controller()\n(template customization)", C_LIGHT3, "#D97706"),
    (75, "—", C_GRAY, C_BORDER),
    (66, "r.customise_resource()\n(per-table customization)", C_LIGHT, C_PRIMARY),
    (57, "PREP HOOK  ✓\ns3.prep = lambda r: ...\n→ dict(bypass=True, output=...)\n→ False  (return 400)\n→ None  (proceed)", C_LIGHT3, "#D97706"),
    (48, "s3db.get_method()\n(custom REST method lookup)", C_LIGHT, C_PRIMARY),
    (40, "BASICCRUD\n→ create(): reads\n  create_onvalidation\n  create_onaccept\n  onvalidation / onaccept", C_LIGHT, C_PRIMARY),
    (31, "CRUDRESOURCE\n→ insert(): runs\n  • authz check\n  • audit('create', ...)\n  • update_super()\n  • set_record_owner()\n  • onaccept(method='create')", C_LIGHT2, "#16A34A"),
    (22, "POSTP HOOK  ✓\ns3.postp = lambda r, out: ...\n→ modify output before view", C_LIGHT3, "#D97706"),
    (13, "—", C_GRAY, C_BORDER),
    (4, "response.headers['Content-Type']\n= s3.content_type['csv']", C_GRAY, C_BORDER),
]
for y, text, fill, border in hooks:
    h = 5
    box(ax, side_x, y, side_w, h, text, fill=fill, border=border, fontsize=7.5, radius=0.04)

# Connect arrows from main flow to hook panel
for y, _, _, _ in hooks:
    if y == 87 or y == 4 or y == 13:
        continue  # skip "—" entries
    arrow(ax, 68, y + 2.5, side_x, y + 2.5, color=C_SECOND, lw=0.8, style="-")
    arrow(ax, 68, y + 2.5, side_x - 0.3, y + 2.5, color=C_SECOND, lw=0.8, style="->")

plt.savefig(os.path.join(OUT_DIR, "diagram_lifecycle.png"),
            dpi=150, bbox_inches="tight", facecolor=C_BG, pad_inches=0.3)
plt.close()
print("✓ diagram_lifecycle.png")

print("Done.")
