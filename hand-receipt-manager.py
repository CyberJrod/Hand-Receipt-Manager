"""
Hand Receipt Manager (DA Form 2062, DEC 2023)
Author: Joseph Rodrigues
Contact: (224)769-0899

• Inventory with scan/manual entry + CSV/XLSX import/export
• Issue and Return on separate tabs (Issue does NOT generate a PDF)
• Issued Items tab: see per-custodian equipment, edit metadata (Issued by / Contact),
  and Generate 2062 using stored metadata
• DA 2062 overlay writes:
    – Page 1 header: FROM name, TO name (no labels), and Contact (below TO)
    – All pages: page fraction (e.g., 1/3, 2/3, ...)
    – Table: Item Description (col c) and QTY AUTH (col g)
• Row packing: 10 serials per row as two printed lines (4 + 6).
  Second line is drawn inside the same row using a small vertical offset.

Place DA2062_flat.pdf next to this file (flattened, non-XFA).
Install deps:
  python -m pip install --upgrade pypdf reportlab pandas openpyxl cryptography
"""

import os
import io
import sys
import re
import json
import sqlite3
import traceback
from dataclasses import dataclass, asdict
from datetime import datetime
from collections import defaultdict

# -------------- Startup guard --------------
def _startup_error_dialog():
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        r = _tk.Tk(); r.withdraw()
        _mb.showerror("Startup error", traceback.format_exc())
        try: r.destroy()
        except Exception: pass
    except Exception:
        print(traceback.format_exc(), file=sys.stderr)

# -------------- GUI imports --------------
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, simpledialog
except Exception:
    _startup_error_dialog()
    raise

# -------------- PDF / Excel deps --------------
try:
    from pypdf import PdfReader as PyPdfReader, PdfWriter as PyPdfWriter
except Exception:
    messagebox.showerror("Missing dependency", "Install pypdf:\n\npython -m pip install pypdf")
    raise

try:
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.utils import simpleSplit
except Exception:
    messagebox.showerror("Missing dependency", "Install reportlab:\n\npython -m pip install reportlab")
    raise

try:
    import pandas as pd
except Exception:
    messagebox.showerror("Missing dependency",
                         "Excel import needs 'pandas' + 'openpyxl':\n\npython -m pip install pandas openpyxl")
    raise

# -------------- Config --------------
DB_FILE = "inventory.db"
TEMPLATE_PDF = "DA2062_flat.pdf"  # flattened, non-XFA

STATUS_ON_HAND = "On Hand"
STATUS_ISSUED  = "Issued"

# Pack 10 serials per logical row (4 on line 1, 6 on line 2)
SERIALS_FIRST_LINE = 4
SERIALS_SECOND_LINE = 6
SERIALS_PER_ROW = SERIALS_FIRST_LINE + SERIALS_SECOND_LINE

# -------------- Layout calibration --------------
LAYOUT_FILE = "da2062_layout.json"

@dataclass
class LayoutConfig:
    # Fonts
    font_name: str = "Helvetica"
    font_size: float = 9.0
    font_size_hdr: float = 10.0

    # Header coordinates (Page 1 only for names/contact)
    x_from: float = 260.0    # inside the FROM box; name only (no label)
    y_from: float = 590.0
    x_to: float   = 710.0    # inside the TO box; name only (no label)
    y_to: float   = 590.0
    to_contact_offset: float = 12.0  # drop below TO for contact line

    # Page fraction (e.g., 1/2)
    x_page_right: float = 575.0
    y_identifier: float = 717.0

    # Table columns / rows
    item_desc_x: float = 226.0
    qty_auth_x: float  = 591.0
    item_start_y_first: float = 493.0
    item_start_y_next: float  = 640.0
    line_spacing: float = 23.0
    second_line_offset: float = 11.0
    rows_first: int = 16
    rows_next: int  = 20

def load_layout() -> "LayoutConfig":
    if os.path.exists(LAYOUT_FILE):
        try:
            with open(LAYOUT_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return LayoutConfig(**{**asdict(LayoutConfig()), **data})
        except Exception:
            return LayoutConfig()
    return LayoutConfig()

def save_layout(cfg: "LayoutConfig"):
    with open(LAYOUT_FILE, "w", encoding="utf-8") as f:
        json.dump(asdict(cfg), f, indent=2)

LCFG = load_layout()

# -------------- DB --------------
def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    # inventory
    cur.execute("""
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model TEXT NOT NULL,
            category TEXT NOT NULL,
            box_no TEXT,
            serial TEXT UNIQUE NOT NULL,
            asset_tag TEXT,
            status TEXT NOT NULL,
            custodian TEXT,
            updated_at TEXT NOT NULL
        )
    """)
    # issues (history)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS issues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            issue_dt TEXT NOT NULL,
            issued_from TEXT NOT NULL,
            issued_to TEXT NOT NULL,
            doc_no TEXT,
            remarks TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS issue_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            issue_id INTEGER NOT NULL,
            model TEXT NOT NULL,
            category TEXT NOT NULL,
            serial TEXT NOT NULL,
            asset_tag TEXT,
            FOREIGN KEY(issue_id) REFERENCES issues(id)
        )
    """)
    # NEW: custodian meta (persist contact and issued_from)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS custodian_meta (
            custodian TEXT PRIMARY KEY,
            contact TEXT,
            issued_from TEXT,
            updated_at TEXT NOT NULL
        )
    """)
    conn.commit(); conn.close()

def _table_columns(conn, table):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    return {row[1] for row in cur.fetchall()}

def migrate_db():
    conn = sqlite3.connect(DB_FILE)
    try:
        # make sure issues has doc_no & remarks
        cols = _table_columns(conn, "issues")
        cur = conn.cursor()
        if "doc_no" not in cols:
            cur.execute("ALTER TABLE issues ADD COLUMN doc_no TEXT")
        if "remarks" not in cols:
            cur.execute("ALTER TABLE issues ADD COLUMN remarks TEXT")
        conn.commit()
    finally:
        conn.close()

def with_conn(fn):
    def wrap(*a, **k):
        conn = sqlite3.connect(DB_FILE)
        try:
            out = fn(conn, *a, **k)
            conn.commit()
            return out
        finally:
            conn.close()
    return wrap

# inventory ops
@with_conn
def db_add_item(conn, model, category, box_no, serial, asset_tag):
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO inventory (model, category, box_no, serial, asset_tag, status, custodian, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (model.strip(), category.strip(), (box_no or "").strip() or None,
          serial.strip(), (asset_tag or "").strip() or None,
          STATUS_ON_HAND, None, now_iso()))
    return cur.lastrowid

@with_conn
def db_remove_by_serial(conn, serial):
    cur = conn.cursor()
    cur.execute("DELETE FROM inventory WHERE serial=?", (serial.strip(),))
    return cur.rowcount

@with_conn
def db_list_inventory(conn):
    cur = conn.cursor()
    cur.execute("""
        SELECT id, model, category, COALESCE(box_no,''), serial, COALESCE(asset_tag,''), status, COALESCE(custodian,''), updated_at
        FROM inventory
        ORDER BY model, serial
    """)
    return cur.fetchall()

# CSV/Excel
@with_conn
def db_export_csv(conn, path):
    import csv
    cur = conn.cursor()
    cur.execute("""
        SELECT model, category, COALESCE(box_no,''), serial, COALESCE(asset_tag,''), status, COALESCE(custodian,''), updated_at
        FROM inventory ORDER BY model, serial
    """)
    rows = cur.fetchall()
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Model","Category","Box #","Serial Number","Asset Tag #","Status","Custodian","Updated At"])
        for r in rows: w.writerow(r)

@with_conn
def db_import_csv(conn, path):
    import csv
    cur = conn.cursor(); added = 0
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        r = csv.DictReader(f)
        required = {"Model","Category","Serial Number"}
        if not required.issubset(set(r.fieldnames or [])):
            raise ValueError(f"CSV must include at least {required}")
        for row in r:
            model = (row.get("Model") or "").strip()
            category = (row.get("Category") or "").strip()
            serial = (row.get("Serial Number") or "").strip()
            if not (model and category and serial): continue
            box = (row.get("Box #") or "").strip() or None
            asset = (row.get("Asset Tag #") or "").strip() or None
            try:
                cur.execute("""
                    INSERT INTO inventory (model, category, box_no, serial, asset_tag, status, custodian, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (model, category, box, serial, asset, STATUS_ON_HAND, None, now_iso()))
                added += 1
            except sqlite3.IntegrityError:
                pass
    return added

@with_conn
def db_import_excel(conn, path):
    df = pd.read_excel(path, engine="openpyxl")
    cols = [str(c).strip() for c in df.columns]
    colmap = {}
    for c in cols:
        lc = c.lower()
        if lc == "model": colmap["Model"] = c
        elif lc in ("category","cat"): colmap["Category"] = c
        elif lc in ("serial number","serial","s/n","sn","s\\n"): colmap["Serial Number"] = c
        elif lc in ("box #","box","box number","box no","boxno"): colmap["Box #"] = c
        elif lc in ("asset tag #","asset tag","asset","asset#","asset id"): colmap["Asset Tag #"] = c
    required = {"Model","Category","Serial Number"}
    if not required.issubset(colmap):
        raise ValueError(f"Excel must include: {sorted(required)} (found {sorted(colmap.keys())})")
    cur = conn.cursor(); added = 0
    for _, r in df.iterrows():
        model = str(r[colmap["Model"]]).strip() if pd.notna(r[colmap["Model"]]) else ""
        category = str(r[colmap["Category"]]).strip() if pd.notna(r[colmap["Category"]]) else ""
        serial = str(r[colmap["Serial Number"]]).strip() if pd.notna(r[colmap["Serial Number"]]) else ""
        if not (model and category and serial): continue
        box = (str(r[colmap["Box #"]]).strip() if "Box #" in colmap and pd.notna(r[colmap["Box #"]]) else None)
        asset = (str(r[colmap["Asset Tag #"]]).strip() if "Asset Tag #" in colmap and pd.notna(r[colmap["Asset Tag #"]]) else None)
        try:
            cur.execute("""
                INSERT INTO inventory (model, category, box_no, serial, asset_tag, status, custodian, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (model, category, box, serial, asset, STATUS_ON_HAND, None, now_iso()))
            added += 1
        except sqlite3.IntegrityError:
            pass
    return added

# lookups
@with_conn
def db_find_onhand_by_serials(conn, serials):
    if not serials: return []
    q = ",".join("?" for _ in serials)
    cur = conn.cursor()
    cur.execute(f"""
        SELECT id, model, category, COALESCE(asset_tag,''), serial
        FROM inventory
        WHERE serial IN ({q}) AND status=?
    """, (*[s.strip() for s in serials], STATUS_ON_HAND))
    return cur.fetchall()

@with_conn
def db_find_issued_by_serials(conn, serials):
    if not serials: return []
    q = ",".join("?" for _ in serials)
    cur = conn.cursor()
    cur.execute(f"""
        SELECT id, model, category, COALESCE(asset_tag,''), serial, COALESCE(custodian,'')
        FROM inventory
        WHERE serial IN ({q}) AND status=?
    """, (*[s.strip() for s in serials], STATUS_ISSUED))
    return cur.fetchall()

# custodian meta
@with_conn
def db_upsert_custodian_meta(conn, custodian, contact, issued_from):
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO custodian_meta (custodian, contact, issued_from, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(custodian) DO UPDATE SET
            contact=excluded.contact,
            issued_from=excluded.issued_from,
            updated_at=excluded.updated_at
    """, (custodian.strip(), (contact or "").strip(), (issued_from or "").strip(), now_iso()))

@with_conn
def db_get_custodian_meta(conn, custodian):
    cur = conn.cursor()
    cur.execute("SELECT contact, issued_from FROM custodian_meta WHERE custodian=?", (custodian.strip(),))
    row = cur.fetchone()
    if not row: return {"contact":"", "issued_from":""}
    return {"contact": row[0] or "", "issued_from": row[1] or ""}

# issue/return
@with_conn
def db_mark_issued(conn, items, issued_from, issued_to):
    cur = conn.cursor()
    issue_dt = now_iso()
    cur.execute("""
        INSERT INTO issues (issue_dt, issued_from, issued_to)
        VALUES (?, ?, ?)
    """, (issue_dt, issued_from.strip(), issued_to.strip()))
    issue_id = cur.lastrowid
    for it in items:
        cur.execute("""
            INSERT INTO issue_items (issue_id, model, category, serial, asset_tag)
            VALUES (?, ?, ?, ?, ?)
        """, (issue_id, it["model"], it["category"], it["serial"], it.get("asset_tag") or None))
        cur.execute("""
            UPDATE inventory SET status=?, custodian=?, updated_at=? WHERE serial=?
        """, (STATUS_ISSUED, issued_to.strip(), now_iso(), it["serial"]))
    return issue_id

@with_conn
def db_mark_returned(conn, serials):
    cur = conn.cursor(); updated = 0
    for s in serials:
        cur.execute("""
            UPDATE inventory
               SET status=?, custodian=NULL, updated_at=?
             WHERE serial=? AND status=?
        """, (STATUS_ON_HAND, now_iso(), s.strip(), STATUS_ISSUED))
        updated += cur.rowcount
    return updated

@with_conn
def db_distinct_custodians_with_counts(conn):
    cur = conn.cursor()
    cur.execute("""
        SELECT COALESCE(custodian,''), COUNT(*)
        FROM inventory
        WHERE status=?
        GROUP BY COALESCE(custodian,'')
        ORDER BY LOWER(COALESCE(custodian,'')) ASC
    """, (STATUS_ISSUED,))
    return cur.fetchall()

@with_conn
def db_list_issued_by_custodian(conn, custodian):
    cur = conn.cursor()
    cur.execute("""
        SELECT model, category, COALESCE(asset_tag,''), serial, updated_at
        FROM inventory
        WHERE status=? AND COALESCE(custodian,'')=?
        ORDER BY model, serial
    """, (STATUS_ISSUED, custodian.strip()))
    return cur.fetchall()

# -------------- Utils --------------
def sanitize_serials_blob(text):
    parts = []
    for line in str(text).splitlines():
        for chunk in str(line).split(","):
            v = chunk.strip()
            if v: parts.append(v)
    seen, out = set(), []
    for p in parts:
        if p not in seen:
            seen.add(p); out.append(p)
    return out

def chunk_list(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

def build_rows_grouped_by_model(items):
    """
    Build logical rows grouped by model:
      – Up to 10 serials per row
      – Line 1: 'Model - S/N: first 4'
      – Line 2 (optional): 'S/N: next up to 6'
      – QTY (col g) printed once per row
    """
    groups = defaultdict(list)
    for it in items:
        tag = (it.get("asset_tag") or "").strip()
        s = it["serial"].strip() + (f" [AT:{tag}]" if tag else "")
        groups[(it["model"], it["category"])].append(s)

    rows = []
    for (model, _cat), serials in groups.items():
        serials = sorted(serials)
        for start in range(0, len(serials), SERIALS_PER_ROW):
            pack = serials[start:start+SERIALS_PER_ROW]
            count = len(pack)
            if count <= SERIALS_FIRST_LINE:
                l1 = f"{model} - S/N: {', '.join(pack)}"
                rows.append({"l1": l1, "l2": None, "qty": count})
            else:
                first = pack[:SERIALS_FIRST_LINE]
                rest  = pack[SERIALS_FIRST_LINE:SERIALS_FIRST_LINE+SERIALS_SECOND_LINE]
                rows.append({"l1": f"{model} - S/N: {', '.join(first)}",
                             "l2": f"S/N: {', '.join(rest)}" if rest else None,
                             "qty": count})
    return rows

def sanitize_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"[^\w\s.-]+", "", name)
    name = re.sub(r"\s+", "_", name)
    return name or "Unknown"

# -------------- PDF overlay --------------
def _template_reader():
    if not os.path.exists(TEMPLATE_PDF):
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PDF}")
    r = PyPdfReader(TEMPLATE_PDF)
    if getattr(r, "is_encrypted", False):
        try:
            r.decrypt("")
        except Exception:
            pwd = simpledialog.askstring("PDF Password","Enter template password (blank for none):", show='*')
            if pwd is None:
                raise RuntimeError("PDF is encrypted and no password was provided.")
            if not r.decrypt(pwd):
                raise RuntimeError("Could not decrypt the template with the provided password.")
    return r

def _draw_header(c, meta, page_no, of_pages, first_page: bool):
    """
    Page 1: draw FROM name and TO name (no labels), plus Contact line below TO.
    All pages: draw page fraction at top-right calibrated position.
    """
    c.setFont(LCFG.font_name, LCFG.font_size_hdr)
    if first_page:
        # Names only (labels already on template)
        if meta.get("issued_from"):
            c.drawString(LCFG.x_from, LCFG.y_from, meta["issued_from"])
        if meta.get("issued_to"):
            c.drawString(LCFG.x_to,   LCFG.y_to,   meta["issued_to"])
        contact = (meta.get("to_contact") or "").strip()
        if contact:
            max_w = 300
            lines = simpleSplit(f"Contact: {contact}", LCFG.font_name, LCFG.font_size_hdr, max_w)
            y = LCFG.y_to - LCFG.to_contact_offset
            for ln in lines[:2]:
                c.drawString(LCFG.x_to, y, ln)
                y -= (LCFG.font_size_hdr + 2)

    # Page fraction (move with calibration if it touches the TO panel)
    c.drawRightString(LCFG.x_page_right, LCFG.y_identifier, f"{page_no}/{of_pages}")

def _draw_rows(c, rows, start_y, rows_cap):
    y = start_y
    max_w = 430
    for r in rows[:rows_cap]:
        # line 1
        c.setFont(LCFG.font_name, LCFG.font_size)
        l1w = simpleSplit(r["l1"], LCFG.font_name, LCFG.font_size, max_w)
        c.drawString(LCFG.item_desc_x, y, l1w[0])
        if len(l1w) > 1:
            c.setFont(LCFG.font_name, LCFG.font_size - 1)
            c.drawString(LCFG.item_desc_x, y - (LCFG.font_size - 1) - 1, l1w[1])
            c.setFont(LCFG.font_name, LCFG.font_size)

        # qty on line 1 only
        c.drawRightString(LCFG.qty_auth_x, y, str(r["qty"]))

        # line 2 (inside same row)
        if r["l2"]:
            c.setFont(LCFG.font_name, LCFG.font_size)
            l2y = y - LCFG.second_line_offset
            l2w = simpleSplit(r["l2"], LCFG.font_name, LCFG.font_size, max_w)
            c.drawString(LCFG.item_desc_x, l2y, l2w[0])
            if len(l2w) > 1:
                c.setFont(LCFG.font_name, LCFG.font_size - 1)
                c.drawString(LCFG.item_desc_x, l2y - (LCFG.font_size - 1) - 1, l2w[1])
                c.setFont(LCFG.font_name)

        y -= LCFG.line_spacing

def _make_overlay_page(meta, rows, page_no, of_pages, first_page=True):
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=letter)
    _draw_header(c, meta, page_no, of_pages, first_page=first_page)
    start_y = LCFG.item_start_y_first if first_page else LCFG.item_start_y_next
    cap = LCFG.rows_first if first_page else LCFG.rows_next
    _draw_rows(c, rows, start_y, cap)
    c.showPage()
    c.save()
    buf.seek(0)
    return buf

def _merge_overlay(base_page, overlay_buf):
    overlay_reader = PyPdfReader(overlay_buf)
    overlay_page = overlay_reader.pages[0]
    base_page.merge_page(overlay_page)

def render_2062_overlay(output_path, meta, rows):
    reader = _template_reader()
    if len(reader.pages) == 0:
        raise RuntimeError("Template has no pages.")
    tpl_first = reader.pages[0]
    tpl_next  = reader.pages[1] if len(reader.pages) > 1 else tpl_first

    first_rows = rows[:LCFG.rows_first]
    rest = rows[LCFG.rows_first:]
    next_groups = list(chunk_list(rest, LCFG.rows_next))
    total_pages = 1 + len(next_groups)

    writer = PyPdfWriter()

    # page 1
    writer.add_page(tpl_first)
    page = writer.pages[-1]
    buf = _make_overlay_page(meta, first_rows, 1, total_pages, first_page=True)
    _merge_overlay(page, buf)

    # continuation pages
    for idx, grp in enumerate(next_groups, start=2):
        writer.add_page(tpl_next)
        page = writer.pages[-1]
        buf = _make_overlay_page(meta, grp, idx, total_pages, first_page=False)
        _merge_overlay(page, buf)

    with open(output_path, "wb") as f:
        writer.write(f)

# -------------- GUI --------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hand Receipt Manager (DA 2062 - DEC 2023)")
        self.geometry("1320x900")

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True)
        self.inv_frame = ttk.Frame(nb)
        self.issue_frame = ttk.Frame(nb)
        self.return_frame = ttk.Frame(nb)
        self.issued_frame = ttk.Frame(nb)
        self.calib_frame = ttk.Frame(nb)
        nb.add(self.inv_frame, text="Modify Inventory")
        nb.add(self.issue_frame, text="Issue")
        nb.add(self.return_frame, text="Return")
        nb.add(self.issued_frame, text="Issued Items")
        nb.add(self.calib_frame, text="Calibration")

        self._build_inventory_tab()
        self._build_issue_tab()
        self._build_return_tab()
        self._build_issued_tab()
        self._build_calibration_tab()

        self.refresh_inventory()
        self.refresh_issued_lists()

    # ----- Inventory tab -----
    def _build_inventory_tab(self):
        frm = self.inv_frame
        lf = ttk.LabelFrame(frm, text="Add Item"); lf.pack(side="top", fill="x", padx=8, pady=8)

        self.model_var = tk.StringVar()
        self.category_var = tk.StringVar(value="Laptop")
        self.box_var = tk.StringVar()
        self.serial_var = tk.StringVar()
        self.asset_var = tk.StringVar()

        r = 0
        ttk.Label(lf, text="Model*").grid(row=r, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(lf, textvariable=self.model_var, width=28).grid(row=r, column=1, sticky="w")
        ttk.Label(lf, text="Category*").grid(row=r, column=2, sticky="w", padx=6)
        ttk.Combobox(lf, textvariable=self.category_var,
                     values=["Laptop","IP Phone","Tablet","Monitor","Headset","Other"], width=26)\
            .grid(row=r, column=3, sticky="w")
        r += 1
        ttk.Label(lf, text="Box #").grid(row=r, column=0, sticky="w", padx=6)
        ttk.Entry(lf, textvariable=self.box_var, width=28).grid(row=r, column=1, sticky="w")
        ttk.Label(lf, text="Serial Number*").grid(row=r, column=2, sticky="w", padx=6)
        ttk.Entry(lf, textvariable=self.serial_var, width=28).grid(row=r, column=3, sticky="w")
        r += 1
        ttk.Label(lf, text="Asset Tag #").grid(row=r, column=0, sticky="w", padx=6)
        ttk.Entry(lf, textvariable=self.asset_var, width=28).grid(row=r, column=1, sticky="w")
        ttk.Button(lf, text="Add to Inventory", command=self.add_item)\
            .grid(row=r, column=3, sticky="e", pady=4)

        tbl = ttk.Frame(frm); tbl.pack(side="top", fill="both", expand=True, padx=8, pady=8)
        cols = ("id","Model","Category","Box #","Serial","Asset Tag #","Status/Custodian","Updated")
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings", selectmode="extended")
        for c in cols:
            self.tree.heading(c, text=c); self.tree.column(c, width=140, anchor="w")
        self.tree.column("id", width=50, anchor="center")
        self.tree.column("Status/Custodian", width=240)
        vsb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True); vsb.pack(side="right", fill="y")

        btns = ttk.Frame(frm); btns.pack(side="bottom", fill="x", padx=8, pady=8)
        ttk.Button(btns, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=4)
        ttk.Button(btns, text="Export CSV", command=self.export_csv).pack(side="left", padx=4)
        ttk.Button(btns, text="Import CSV", command=self.import_csv).pack(side="left", padx=4)
        ttk.Button(btns, text="Import Excel (.xlsx)", command=self.import_excel).pack(side="left", padx=4)
        ttk.Button(btns, text="Refresh", command=self.refresh_inventory).pack(side="left", padx=4)

    def add_item(self):
        model = self.model_var.get().strip()
        category = self.category_var.get().strip()
        serial = self.serial_var.get().strip()
        if not model or not category or not serial:
            messagebox.showwarning("Missing", "Model, Category, and Serial Number are required.")
            return
        try:
            db_add_item(model, category, self.box_var.get(), serial, self.asset_var.get())
            self.model_var.set(""); self.box_var.set(""); self.serial_var.set(""); self.asset_var.set("")
            self.refresh_inventory()
        except sqlite3.IntegrityError:
            messagebox.showerror("Duplicate", "That serial number already exists.")

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel: return
        if not messagebox.askyesno("Confirm", "Remove selected item(s) from inventory?"): return
        removed = 0
        for iid in sel:
            vals = self.tree.item(iid, "values"); serial = vals[4]
            removed += db_remove_by_serial(serial)
        self.refresh_inventory()
        messagebox.showinfo("Done", f"Removed {removed} item(s).")

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV","*.csv")],
                                            title="Export Inventory to CSV")
        if not path: return
        db_export_csv(path)
        messagebox.showinfo("Exported", f"Inventory exported to:\n{path}")

    def import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV","*.csv")], title="Import Inventory from CSV")
        if not path: return
        try:
            added = db_import_csv(path)
            self.refresh_inventory()
            messagebox.showinfo("Imported", f"Imported {added} item(s).")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))

    def import_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")], title="Import Inventory from Excel")
        if not path: return
        try:
            added = db_import_excel(path)
            self.refresh_inventory()
            messagebox.showinfo("Imported", f"Imported {added} item(s) from Excel.")
        except Exception as e:
            messagebox.showerror("Excel Import Error", str(e))

    def refresh_inventory(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for (id_, model, cat, box, serial, asset, status, cust, updated) in db_list_inventory():
            scell = status if status == STATUS_ON_HAND else f"{status} to {cust}"
            self.tree.insert("", "end", values=(id_, model, cat, box, serial, asset, scell, updated))

    # ----- Issue tab -----
    def _build_issue_tab(self):
        frm = self.issue_frame
        lf = ttk.LabelFrame(frm, text="Issue Equipment"); lf.pack(side="top", fill="x", padx=8, pady=8)

        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()
        self.to_contact_var = tk.StringVar()

        r = 0
        ttk.Label(lf, text="From:").grid(row=r, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(lf, textvariable=self.from_var, width=40).grid(row=r, column=1, sticky="w")
        ttk.Label(lf, text="To (Person/Unit):").grid(row=r, column=2, sticky="w")
        ttk.Entry(lf, textvariable=self.to_var, width=40).grid(row=r, column=3, sticky="w")
        r += 1
        ttk.Label(lf, text="To (Contact Information):").grid(row=r, column=2, sticky="w")
        ttk.Entry(lf, textvariable=self.to_contact_var, width=40).grid(row=r, column=3, sticky="w")
        r += 1

        ttk.Label(lf, text="Scan/Enter Serials (newline or comma separated):").grid(row=r, column=0, columnspan=4, sticky="w", padx=6)
        r += 1
        self.serials_text_issue = tk.Text(lf, height=8)
        self.serials_text_issue.grid(row=r, column=0, columnspan=4, sticky="we", padx=6, pady=4)
        r += 1

        ttk.Button(lf, text="Validate On-Hand", command=self.validate_issue_serials).grid(row=r, column=0, sticky="w", padx=6, pady=4)
        ttk.Button(lf, text="Issue", command=self.issue_only).grid(row=r, column=1, sticky="w", padx=6)
        ttk.Button(lf, text="Clear Serials", command=lambda: self.serials_text_issue.delete("1.0","end")).grid(row=r, column=2, sticky="w", padx=6)

        lo = ttk.LabelFrame(frm, text="Validation Output")
        lo.pack(side="top", fill="both", expand=True, padx=8, pady=8)
        self.output_issue = tk.Text(lo, height=12); self.output_issue.pack(fill="both", expand=True)

    def append_issue_output(self, msg):
        self.output_issue.insert("end", msg + "\n"); self.output_issue.see("end")

    def validate_issue_serials(self):
        serials = sanitize_serials_blob(self.serials_text_issue.get("1.0","end"))
        if not serials:
            messagebox.showwarning("No Serials", "Please scan or enter serial numbers.")
            return
        onhand = db_find_onhand_by_serials(serials)
        found = {s for (_,_,_,_,s) in onhand}
        missing = [s for s in serials if s not in found]
        self.output_issue.delete("1.0","end")
        self.append_issue_output(f"Requested serials: {len(serials)}")
        self.append_issue_output(f"On-hand & available: {len(onhand)}")
        if missing:
            self.append_issue_output(f"Not available / not found ({len(missing)}): {', '.join(missing)}")
        else:
            self.append_issue_output("All requested serials are available.")

    def issue_only(self):
        issued_from = self.from_var.get().strip()
        issued_to   = self.to_var.get().strip()
        to_contact  = self.to_contact_var.get().strip()
        if not issued_from or not issued_to:
            messagebox.showwarning("Missing", "Please provide both 'From' and 'To (Person/Unit)'.")
            return

        serials = sanitize_serials_blob(self.serials_text_issue.get("1.0","end"))
        if not serials:
            messagebox.showwarning("No Serials", "Please scan or enter serial numbers.")
            return

        onhand = db_find_onhand_by_serials(serials)
        found = {s for (_id, _m, _c, _a, s) in onhand}
        missing = [s for s in serials if s not in found]
        if missing:
            messagebox.showerror("Cannot Issue", "Not on hand / not found:\n" + ", ".join(missing))
            return

        items = [{"model": m, "category": c, "serial": s, "asset_tag": a}
                 for (_id, m, c, a, s) in onhand]

        # Update meta store for this custodian
        db_upsert_custodian_meta(issued_to, to_contact, issued_from)

        # Mark issued
        db_mark_issued(items, issued_from, issued_to)
        self.refresh_inventory()
        self.refresh_issued_lists()

        messagebox.showinfo("Issued", f"Issued {len(items)} item(s) to {issued_to}.\n"
                                      f"You can generate a 2062 later from the 'Issued Items' tab.")
        self.serials_text_issue.delete("1.0","end"); self.output_issue.delete("1.0","end")

    # ----- Return tab -----
    def _build_return_tab(self):
        frm = self.return_frame
        lr = ttk.LabelFrame(frm, text="Return Equipment"); lr.pack(side="top", fill="x", padx=8, pady=8)

        ttk.Label(lr, text="Scan/Enter Serials to Return (newline or comma separated):").grid(row=0, column=0, columnspan=3, sticky="w", padx=6)
        self.return_text = tk.Text(lr, height=8)
        self.return_text.grid(row=1, column=0, columnspan=3, sticky="we", padx=6, pady=4)

        ttk.Button(lr, text="Validate Issued", command=self.validate_return_serials).grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Button(lr, text="Mark Returned", command=self.mark_returned).grid(row=2, column=1, sticky="w", padx=6, pady=4)
        ttk.Button(lr, text="Clear", command=lambda: self.return_text.delete("1.0","end")).grid(row=2, column=2, sticky="w", padx=6, pady=4)

        lo = ttk.LabelFrame(frm, text="Validation Output")
        lo.pack(side="top", fill="both", expand=True, padx=8, pady=8)
        self.output_return = tk.Text(lo, height=12); self.output_return.pack(fill="both", expand=True)

    def append_return_output(self, msg):
        self.output_return.insert("end", msg + "\n"); self.output_return.see("end")

    def validate_return_serials(self):
        serials = sanitize_serials_blob(self.return_text.get("1.0", "end"))
        if not serials:
            messagebox.showwarning("No Serials", "Please scan or enter serial numbers to validate.")
            return

        issued = db_find_issued_by_serials(serials)
        issued_found = {s for (_, _, _, _, s, _) in issued}
        missing = [s for s in serials if s not in issued_found]

        self.output_return.delete("1.0", "end")
        self.append_return_output(f"Entered serials: {len(serials)}")
        self.append_return_output(f"Currently issued: {len(issued)}")

        if issued:
            details = [f"{s} (to {cust})" for (_, _, _, _, s, cust) in issued]
            self.append_return_output("Details:\n  " + "\n  ".join(details))

        if missing:
            self.append_return_output(
                f"Not currently issued / not found ({len(missing)}): {', '.join(missing)}"
            )

    def mark_returned(self):
        serials = sanitize_serials_blob(self.return_text.get("1.0","end"))
        if not serials:
            messagebox.showwarning("No Serials", "Please scan or enter serial numbers to return."); return
        updated = db_mark_returned(serials)
        self.refresh_inventory()
        self.refresh_issued_lists()
        messagebox.showinfo("Returned", f"Marked {updated} item(s) as returned (On Hand).")
        self.return_text.delete("1.0","end")
        self.output_return.delete("1.0","end")

    # ----- Issued Items tab -----
    def _build_issued_tab(self):
        frm = self.issued_frame

        top = ttk.Frame(frm); top.pack(fill="x", padx=8, pady=6)
        ttk.Label(top, text="Filter custodians:").pack(side="left")
        self.filter_custodian_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.filter_custodian_var, width=30).pack(side="left", padx=6)
        ttk.Button(top, text="Apply", command=self.refresh_custodian_list).pack(side="left")
        ttk.Button(top, text="Clear", command=lambda: (self.filter_custodian_var.set(""), self.refresh_custodian_list())).pack(side="left", padx=4)

        body = ttk.Frame(frm); body.pack(fill="both", expand=True, padx=8, pady=8)

        # left: custodians
        left = ttk.Frame(body); left.pack(side="left", fill="y", padx=(0,8))
        ttk.Label(left, text="Custodians (To:)").pack(anchor="w")
        self.cust_list = ttk.Treeview(left, columns=("Custodian","Count"), show="headings", height=22, selectmode="browse")
        self.cust_list.heading("Custodian", text="Custodian")
        self.cust_list.heading("Count", text="# Items")
        self.cust_list.column("Custodian", width=260, anchor="w")
        self.cust_list.column("Count", width=80, anchor="center")
        vsb_l = ttk.Scrollbar(left, orient="vertical", command=self.cust_list.yview)
        self.cust_list.configure(yscrollcommand=vsb_l.set)
        self.cust_list.pack(side="left", fill="y"); vsb_l.pack(side="left", fill="y", padx=(0,6))
        self.cust_list.bind("<<TreeviewSelect>>", lambda e: self.show_items_for_selected_custodian())

        # right: details + meta editor
        right = ttk.Frame(body); right.pack(side="left", fill="both", expand=True)

        # items table
        ttk.Label(right, text="Items issued to selected custodian").pack(anchor="w")
        cols = ("Model","Category","Asset Tag #","Serial","Updated")
        self.cust_items = ttk.Treeview(right, columns=cols, show="headings", selectmode="extended")
        for c in cols:
            self.cust_items.heading(c, text=c)
            self.cust_items.column(c, width=150 if c!="Updated" else 170, anchor="w")
        vsb_r = ttk.Scrollbar(right, orient="vertical", command=self.cust_items.yview)
        self.cust_items.configure(yscrollcommand=vsb_r.set)
        self.cust_items.pack(side="top", fill="both", expand=True); vsb_r.pack(side="right", fill="y")

        # meta editor
        meta = ttk.LabelFrame(frm, text="Custodian metadata (used for 2062 header)")
        meta.pack(fill="x", padx=8, pady=6)
        self.meta_from_var = tk.StringVar()
        self.meta_contact_var = tk.StringVar()
        ttk.Label(meta, text="Issued by (From):").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(meta, textvariable=self.meta_from_var, width=50).grid(row=0, column=1, sticky="w")
        ttk.Label(meta, text="To (Contact information):").grid(row=0, column=2, sticky="w", padx=12)
        ttk.Entry(meta, textvariable=self.meta_contact_var, width=50).grid(row=0, column=3, sticky="w")
        ttk.Button(meta, text="Save metadata", command=self.save_selected_meta).grid(row=0, column=4, sticky="w", padx=8)

        # bottom buttons
        bottom = ttk.Frame(frm); bottom.pack(fill="x", padx=8, pady=6)
        ttk.Button(bottom, text="Refresh", command=self.refresh_issued_lists).pack(side="left")
        ttk.Button(bottom, text="Generate 2062 (Selected Custodian)", command=self.generate_2062_for_selected).pack(side="left", padx=8)

    def refresh_custodian_list(self):
        for i in self.cust_list.get_children(): self.cust_list.delete(i)
        filt = self.filter_custodian_var.get().strip().lower()
        rows = db_distinct_custodians_with_counts()
        for cust, cnt in rows:
            cust_disp = cust if cust else "(Unassigned)"
            if filt and filt not in cust_disp.lower(): continue
            self.cust_list.insert("", "end", values=(cust, cnt))
        # auto-select first row
        kids = self.cust_list.get_children()
        if kids and not self.cust_list.selection():
            self.cust_list.selection_set(kids[0])
        self.show_items_for_selected_custodian()

    def show_items_for_selected_custodian(self):
        for i in self.cust_items.get_children(): self.cust_items.delete(i)
        sel = self.cust_list.selection()
        if not sel:
            self.meta_from_var.set(""); self.meta_contact_var.set("")
            return
        custodian = self.cust_list.item(sel[0], "values")[0]
        rows = db_list_issued_by_custodian(custodian)
        for (m, c, a, s, upd) in rows:
            self.cust_items.insert("", "end", values=(m, c, a, s, upd))
        # load meta into editor
        meta = db_get_custodian_meta(custodian)
        self.meta_from_var.set(meta.get("issued_from",""))
        self.meta_contact_var.set(meta.get("contact",""))

    def save_selected_meta(self):
        sel = self.cust_list.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select a custodian on the left first.")
            return
        custodian = self.cust_list.item(sel[0], "values")[0]
        db_upsert_custodian_meta(custodian, self.meta_contact_var.get(), self.meta_from_var.get())
        messagebox.showinfo("Saved", f"Metadata saved for:\n{custodian}")

    def refresh_issued_lists(self):
        self.refresh_custodian_list()

    def generate_2062_for_selected(self):
        sel = self.cust_list.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select a custodian in the list first.")
            return
        custodian = self.cust_list.item(sel[0], "values")[0]
        items = db_list_issued_by_custodian(custodian)
        if not items:
            messagebox.showwarning("No items", f"No items currently issued to {custodian}.")
            return

        meta = db_get_custodian_meta(custodian)
        issued_from = meta.get("issued_from","") or simpledialog.askstring("From", "Enter FROM (issuing unit/person):", parent=self) or ""
        to_contact  = meta.get("contact","") or simpledialog.askstring("Contact Info", f"Enter contact info for {custodian} (optional):", parent=self) or ""
        # persist any newly-entered values
        db_upsert_custodian_meta(custodian, to_contact, issued_from)

        # Convert rows to overlay items
        to_items = [{"model": m, "category": c, "asset_tag": a, "serial": s}
                    for (m, c, a, s, _upd) in items]
        rows = build_rows_grouped_by_model(to_items)

        to_name = sanitize_filename(custodian)[:60]
        default_name = f"DA2062_{to_name}_{datetime.now().strftime('%Y%m%d')}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                            filetypes=[("PDF","*.pdf")],
                                            initialfile=default_name,
                                            title="Save DA Form 2062")
        if not path:
            return

        hdr = {"issued_from": issued_from, "issued_to": custodian, "to_contact": to_contact}
        try:
            render_2062_overlay(path, hdr, rows)
            messagebox.showinfo("Saved", f"DA Form 2062 saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("PDF Error", str(e))

    # ----- Calibration tab -----
    def _build_calibration_tab(self):
        frm = self.calib_frame
        info = ttk.LabelFrame(frm, text="How calibration works")
        info.pack(fill="x", padx=10, pady=8)
        ttk.Label(info, text=(
            "Use a real 2062 or a quick test to see where text lands.\n"
            "Move text DOWN by decreasing Y values. Row height controls spacing between rows.\n"
            "Second line offset controls the small drop for line 2 inside the same row.\n"
            "If a black square appears near TO, move the page fraction using 'Header: Page right X' or 'baseline Y'."
        )).pack(anchor="w", padx=8, pady=6)

        grid = ttk.LabelFrame(frm, text="Overlay settings (points)")
        grid.pack(fill="both", expand=True, padx=10, pady=8)

        self.vars = {}
        fields = [
            ("item_desc_x", "Item description X (col c)"),
            ("qty_auth_x",  "QTY AUTH X (col g)"),
            ("item_start_y_first", "First page row-1 Y"),
            ("item_start_y_next",  "Next pages row-1 Y"),
            ("line_spacing", "Row height"),
            ("second_line_offset", "Second line offset (Y)"),
            ("x_from",  "Header: FROM name X"),
            ("y_from",  "Header: FROM name Y"),
            ("x_to",    "Header: TO name X"),
            ("y_to",    "Header: TO name Y"),
            ("to_contact_offset", "Header: TO contact offset"),
            ("y_identifier", "Header: Page fraction baseline Y"),
            ("x_page_right", "Header: Page right X"),
            ("font_size", "Body font size"),
            ("font_size_hdr", "Header font size"),
        ]
        for i, (key, label) in enumerate(fields):
            ttk.Label(grid, text=label).grid(row=i, column=0, sticky="w", padx=6, pady=4)
            init = getattr(LCFG, key)
            var = tk.DoubleVar(value=float(init))
            self.vars[key] = var
            ttk.Spinbox(grid, textvariable=var, from_=0, to=1000, increment=0.5, width=10)\
                .grid(row=i, column=1, sticky="w", padx=6, pady=4)

        # rows per page
        ttk.Label(grid, text="Rows first page (logical rows)").grid(row=len(fields), column=0, sticky="w", padx=6, pady=4)
        self.vars["rows_first"] = tk.IntVar(value=int(LCFG.rows_first))
        ttk.Spinbox(grid, textvariable=self.vars["rows_first"], from_=1, to=30, increment=1, width=10)\
            .grid(row=len(fields), column=1, sticky="w", padx=6, pady=4)

        ttk.Label(grid, text="Rows next pages (logical rows)").grid(row=len(fields)+1, column=0, sticky="w", padx=6, pady=4)
        self.vars["rows_next"] = tk.IntVar(value=int(LCFG.rows_next))
        ttk.Spinbox(grid, textvariable=self.vars["rows_next"], from_=1, to=30, increment=1, width=10)\
            .grid(row=len(fields)+1, column=1, sticky="w", padx=6, pady=4)

        btns = ttk.Frame(frm); btns.pack(fill="x", padx=10, pady=8)
        ttk.Button(btns, text="Save", command=self.save_calibration).pack(side="left", padx=4)
        ttk.Button(btns, text="Reset Defaults", command=self.reset_calibration).pack(side="left", padx=4)

    def save_calibration(self):
        global LCFG
        for k, var in self.vars.items():
            val = var.get()
            try:
                if isinstance(getattr(LCFG, k), int):
                    setattr(LCFG, k, int(val))
                else:
                    setattr(LCFG, k, float(val))
            except Exception:
                setattr(LCFG, k, val)
        LCFG.rows_first = int(self.vars["rows_first"].get())
        LCFG.rows_next  = int(self.vars["rows_next"].get())
        save_layout(LCFG)
        messagebox.showinfo("Saved", f"Calibration saved to {LAYOUT_FILE}.")

    def reset_calibration(self):
        global LCFG
        LCFG = LayoutConfig()
        save_layout(LCFG)
        for k, var in self.vars.items():
            var.set(getattr(LCFG, k))
        messagebox.showinfo("Reset", "Calibration reset to defaults.")

# -------------- Main --------------
if __name__ == "__main__":
    try:
        init_db()
        migrate_db()
        app = App()
        app.mainloop()
    except Exception:
        _startup_error_dialog()
        raise