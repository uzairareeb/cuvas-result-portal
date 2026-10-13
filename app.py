"""
CUVAS PhD Entry Test Result Portal
----------------------------------
Admin uploads the entry-test result list (CSV/Excel).
Students enter their CNIC -> see their Entry Test Result Card -> download PDF.

Run locally:  python3 app.py        (uses SQLite, results.db)
Deploy on Vercel: set DATABASE_URL env var to a Postgres URL
  (e.g. Neon via Vercel's one-click Postgres integration).
"""
import csv
import io
import os
import re
import sqlite3
from datetime import datetime
from functools import wraps

from flask import (Flask, flash, g, redirect, render_template, request,
                   send_file, session, url_for)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "results.db")
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
USE_PG = bool(DATABASE_URL)

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, "templates"),
            static_folder=os.path.join(BASE_DIR, "static"))
app.secret_key = os.environ.get("PORTAL_SECRET", "cuvas-result-portal-dev-key")
app.config["ADMIN_PASSWORD"] = os.environ.get("ADMIN_PASSWORD", "cuvas123")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB
ALLOWED_EXT = {"csv", "xlsx", "xls"}

COLUMN_MAP = {
    "name": ["name", "candidate name", "candidate", "student name", "applicant name"],
    "father_name": ["father name", "father_name", "father", "fathers name", "s/o d/o"],
    "cnic": ["cnic", "cnic no", "cnic number", "nic", "national id"],
    "roll_no": ["roll no", "roll_no", "roll number", "rollnumber", "seat no"],
    "discipline": ["discipline", "applied degree", "degree", "program", "programme",
                 "department", "subject", "field"],
    "test_date": ["test date", "test_date", "held on", "held_on", "date", "exam date"],
    "obtained": ["obtained", "obtained marks", "marks obtained", "marks", "score"],
    "total": ["total", "total marks", "max marks", "out of"],
    "merit_no": ["merit no", "merit_no", "merit number", "merit", "rank"],
}

UNI_NAME = "Cholistan University of Veterinary and Animal Sciences"
UNI_SHORT = "CUVAS, Bahawalpur"
PORTAL_TITLE = "PhD Entry Test Result Portal"


# ---------------- database (SQLite locally, Postgres on Vercel) ----------------
class DB:
    """Tiny wrapper: same `?` placeholders work on sqlite3 and psycopg2."""

    def __init__(self):
        if USE_PG:
            import psycopg2
            self.conn = psycopg2.connect(DATABASE_URL)
            self.pg = True
        else:
            self.conn = sqlite3.connect(DB_PATH)
            self.conn.row_factory = sqlite3.Row
            self.pg = False
        self._cur = None

    def _sql(self, sql):
        return sql.replace("?", "%s") if self.pg else sql

    def execute(self, sql, params=()):
        if self.pg:
            import psycopg2.extras
            cur = self.conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(self._sql(sql), params)
            self._cur = cur
            return cur
        return self.conn.execute(sql, params)

    def fetchone(self, sql, params=()):
        cur = self.execute(sql, params)
        row = cur.fetchone()
        if self.pg:
            cur.close()
        return row

    def fetchall(self, sql, params=()):
        cur = self.execute(sql, params)
        rows = cur.fetchall()
        if self.pg:
            cur.close()
        return rows

    def commit(self):
        self.conn.commit()

    def close(self):
        try:
            if self._cur is not None:
                self._cur.close()
        except Exception:
            pass
        self.conn.close()


def get_db():
    if "db" not in g:
        g.db = DB()
    return g.db


@app.teardown_appcontext
def close_db(_exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = DB()
    if USE_PG:
        db.execute("""
            CREATE TABLE IF NOT EXISTS candidates (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                father_name TEXT,
                cnic TEXT NOT NULL UNIQUE,
                roll_no TEXT,
                discipline TEXT,
                test_date TEXT,
                obtained DOUBLE PRECISION,
                total DOUBLE PRECISION,
                merit_no TEXT,
                uploaded_at TEXT
            )
        """)
        db.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                   " ON CONFLICT(key) DO NOTHING", ("passing_percent", "60"))
        db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                   " ON CONFLICT(key) DO NOTHING", ("published", "1"))
        # Migrate the old 50% default to the new 60% default (custom values untouched).
        db.execute("UPDATE settings SET value='60' WHERE key='passing_percent' AND value='50'")
    else:
        db.execute("""
            CREATE TABLE IF NOT EXISTS candidates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                father_name TEXT,
                cnic TEXT NOT NULL UNIQUE,
                roll_no TEXT,
                discipline TEXT,
                test_date TEXT,
                obtained REAL,
                total REAL,
                merit_no TEXT,
                uploaded_at TEXT
            )
        """)
        db.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        db.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('passing_percent', '60')")
        db.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('published', '1')")
        # Migrate the old 50% default to the new 60% default (custom values untouched).
        db.execute("UPDATE settings SET value='60' WHERE key='passing_percent' AND value='50'")
    db.commit()
    db.close()


_db_ready = False


@app.before_request
def _ensure_db():
    # On Vercel there is no startup hook; create tables lazily on first request.
    global _db_ready
    if not _db_ready:
        try:
            init_db()
        except Exception:
            # Concurrent cold starts can race on CREATE TABLE; the loser
            # gets a duplicate-object error which is safe to ignore.
            pass
        _db_ready = True


# Vercel forwards every request to the serverless function (api/index.py),
# so Flask may see PATH_INFO like "/api/index" or "/api/index/result/..."
# instead of the real route. Strip that prefix so routing works the same
# locally and on Vercel. No-op when the prefix is absent.
_VERCEL_PREFIX = "/api/index"
_orig_wsgi_app = app.wsgi_app


def _strip_vercel_prefix(environ, start_response):
    path = environ.get("PATH_INFO", "") or ""
    if path == _VERCEL_PREFIX:
        environ["PATH_INFO"] = "/"
    elif path.startswith(_VERCEL_PREFIX + "/"):
        environ["PATH_INFO"] = path[len(_VERCEL_PREFIX):]
    return _orig_wsgi_app(environ, start_response)


app.wsgi_app = _strip_vercel_prefix


def get_setting(key, default=""):
    row = get_db().fetchone("SELECT value FROM settings WHERE key=?", (key,))
    if row is None:
        return default
    return row["value"] if isinstance(row, dict) else row[0]


def set_setting(key, value):
    db = get_db()
    if USE_PG:
        db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                   " ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value",
                   (key, value))
    else:
        db.execute("INSERT OR REPLACE INTO settings(key, value) VALUES (?, ?)", (key, value))
    db.commit()


def norm_cnic(raw):
    return re.sub(r"\D", "", raw or "")


def fmt_num(x):
    """Show whole numbers without decimals (50 not 50.0)."""
    try:
        f = float(x)
    except (TypeError, ValueError):
        return "" if x is None else str(x)
    return str(int(f)) if f.is_integer() else str(round(f, 2))


def calc_result(obtained, total, passing_percent):
    """Passing is a PERCENTAGE of total marks (default 60%).
    Returns (passing_marks_display, status)."""
    try:
        total_f = float(total)
        obt_f = float(obtained)
    except (TypeError, ValueError):
        total_f, obt_f = 0.0, 0.0
    try:
        pct_f = float(passing_percent)
    except (TypeError, ValueError):
        pct_f = 60.0
    passing_marks = round(total_f * pct_f / 100, 2)
    status = "Qualified" if (total_f > 0 and obt_f >= passing_marks) else "Not Qualified"
    return fmt_num(passing_marks), status


_DATE_FORMATS = ("%d-%b-%Y", "%d %b %Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y",
                 "%m/%d/%Y", "%b %d, %Y", "%d.%m.%Y", "%Y/%m/%d")


def fmt_cnic(cnic):
    """3610346954971 -> 36103-4695497-1 for display."""
    d = norm_cnic(cnic)
    if len(d) == 13:
        return f"{d[:5]}-{d[5:12]}-{d[12]}"
    return cnic or ""


def parse_date(s):
    s = (s or "").strip()
    for f in _DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    return None


def add_one_year(d):
    try:
        return d.replace(year=d.year + 1)
    except ValueError:  # Feb 29 -> Feb 28
        return d.replace(year=d.year + 1, day=28)


def fmt_date(d):
    return d.strftime("%b %d, %Y") if d else ""


def build_card(r, passing_pct):
    """Assemble the result-card fields from a candidate record."""
    held = parse_date(r.get("test_date"))
    valid = add_one_year(held) if held else None
    passing_marks, status = calc_result(r.get("obtained"), r.get("total"), passing_pct)
    return {
        "name": r.get("name", ""),
        "father_name": r.get("father_name", ""),
        "cnic": fmt_cnic(r.get("cnic", "")),
        "cnic_raw": norm_cnic(r.get("cnic", "")),
        "degree": r.get("discipline", ""),
        "held_on": fmt_date(held) if held else (r.get("test_date") or ""),
        "total": fmt_num(r.get("total")),
        "obtained": fmt_num(r.get("obtained")),
        "passing_marks": passing_marks,
        "status": status,
        "valid_till": fmt_date(valid),
        "footer_year": str(valid.year) if valid else str(datetime.now().year),
    }


# ---------------- student side ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if get_setting("published", "1") != "1":
        return render_template("unpublished.html", uni=UNI_NAME, title=PORTAL_TITLE)
    error = None
    if request.method == "POST":
        cnic = norm_cnic(request.form.get("cnic", ""))
        if len(cnic) != 13:
            error = "Please enter a valid 13-digit CNIC number."
        else:
            return redirect(url_for("result", cnic=cnic))
    return render_template("index.html", uni=UNI_NAME, short=UNI_SHORT,
                           title=PORTAL_TITLE, error=error)


@app.route("/result/<cnic>")
def result(cnic):
    cnic = norm_cnic(cnic)
    row = get_db().fetchone("SELECT * FROM candidates WHERE cnic=?", (cnic,))
    if not row:
        return render_template("notfound.html", uni=UNI_NAME, title=PORTAL_TITLE), 404
    r = dict(row)
    passing_pct = float(get_setting("passing_percent", "60") or 60)
    card = build_card(r, passing_pct)
    return render_template("result.html", uni=UNI_NAME, short=UNI_SHORT,
                           title=PORTAL_TITLE, card=card)


@app.route("/result/<cnic>/pdf")
def result_pdf(cnic):
    cnic = norm_cnic(cnic)
    row = get_db().fetchone("SELECT * FROM candidates WHERE cnic=?", (cnic,))
    if not row:
        return "Result not found.", 404
    r = dict(row)
    passing_pct = float(get_setting("passing_percent", "60") or 60)
    card = build_card(r, passing_pct)
    buf = build_pdf(card)
    filename = f"CUVAS_AdmissionTest_Result_{card['cnic_raw']}.pdf"
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


# ---------------- PDF generation ----------------
def build_pdf(card):
    """Result card PDF matching the official CUVAS template:
    logo header, Admission Test (Postgraduate), Held On, candidate fields,
    photo placeholder, marks table with Valid Till + Result Status, blue footer.
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    GREEN = colors.HexColor("#0b7a3e")
    BLUE = colors.HexColor("#232a7a")
    DARK = colors.HexColor("#111111")
    GREY = colors.HexColor("#555555")
    LBL_BG = colors.HexColor("#eef0fa")
    LINE = colors.HexColor("#333333")

    # Card border
    bx, bw = 30, W - 60
    btop, bbot = H - 52, 52
    c.setStrokeColor(DARK)
    c.setLineWidth(1.2)
    c.rect(bx, bbot, bw, btop - bbot, stroke=1, fill=0)

    # Header: logo box + university name
    lx, ly, lw, lh = bx + 16, btop - 58, 96, 44
    c.setFillColor(BLUE)
    c.rect(lx, ly, lw, lh, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 19)
    c.drawCentredString(lx + lw / 2, ly + 17, "CUVAS")
    c.setFont("Helvetica-Bold", 12.5)
    c.setFillColor(DARK)
    c.drawString(lx + lw + 12, ly + 16,
                 "Cholistan University of Veterinary & Animal Sciences, Bahawalpur")

    # Green rule
    c.setStrokeColor(GREEN)
    c.setLineWidth(2)
    c.line(bx + 16, ly - 12, bx + bw - 16, ly - 12)

    # Title block
    cy = W / 2
    c.setFillColor(DARK)
    c.setFont("Helvetica", 12)
    c.drawCentredString(cy, ly - 40, "CUVAS - Admission Test (Postgraduate)")
    c.setFillColor(GREEN)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(cy, ly - 62, f"Held On:  {card['held_on']}")

    # Candidate fields (two columns; right side reserved for the photo box)
    fy = ly - 108
    photo_w = 102
    fields_right = bx + bw - 16 - photo_w - 14
    col1_x = bx + 16
    col2_x = col1_x + (fields_right - col1_x) / 2
    val_dx = 100
    c.setFont("Helvetica-Bold", 10.5)
    c.setFillColor(DARK)
    c.drawString(col1_x, fy, "Name:")
    c.drawString(col2_x, fy, "Father's Name:")
    c.drawString(col1_x, fy - 34, "CNIC:")
    c.drawString(col2_x, fy - 34, "Applied Degree:")
    c.setFont("Helvetica", 10.5)
    c.drawString(col1_x + val_dx, fy, card["name"])
    c.drawString(col2_x + val_dx, fy, card["father_name"])
    c.drawString(col1_x + val_dx, fy - 34, card["cnic"])
    c.drawString(col2_x + val_dx, fy - 34, card["degree"])

    # Photo placeholder (silhouette), vertically centered on the field rows
    px, pw, ph = fields_right + 14, photo_w, 112
    py = fy - 72
    c.saveState()
    clip = c.beginPath()
    clip.roundRect(px, py, pw, ph, 2)
    c.clipPath(clip, stroke=0, fill=0)
    c.setFillColor(colors.HexColor("#e4e4e4"))
    c.rect(px, py, pw, ph, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#8d8d8d"))
    scx = px + pw / 2
    c.circle(scx, py + 72, 18, stroke=0, fill=1)        # head
    c.ellipse(scx - 32, py + 2, 64, 56, stroke=0, fill=1)  # shoulders
    c.restoreState()
    c.setStrokeColor(colors.HexColor("#b5b5b5"))
    c.setLineWidth(1)
    c.roundRect(px, py, pw, ph, 2, stroke=1, fill=0)

    # Marks table
    tx, tw = bx + 16, bw - 32
    ty, rh = fy - 108, 30
    # Marks table: 6 columns; row 2 merges the last cell for Result Status.
    cols = [88, 78, 108, 62, 72, 91]  # sums to 499
    rows_spec = [
        [("Total Marks:", card["total"], 2),
         ("Obtained Marks:", card["obtained"], 2),
         ("Valid Till:", card["valid_till"], 2)],
        [("Passing Marks:", card["passing_marks"], 2),
         ("Result Status:", card["status"], 4)],
    ]
    c.setLineWidth(1)
    for row_i, cells in enumerate(rows_spec):
        ry = ty - row_i * rh
        x = tx
        ci = 0
        for lab, val, ncols in cells:
            lab_w = cols[ci]
            cw = sum(cols[ci:ci + ncols])
            val_w = cw - lab_w
            # label cell
            c.setFillColor(LBL_BG)
            c.rect(x, ry - rh, lab_w, rh, stroke=0, fill=1)
            c.setFillColor(DARK)
            c.setFont("Helvetica-Bold", 10.5)
            c.drawCentredString(x + lab_w / 2, ry - 20, lab)
            # value cell
            c.setFillColor(colors.white)
            c.rect(x + lab_w, ry - rh, val_w, rh, stroke=0, fill=1)
            if lab == "Result Status:":
                c.setFillColor(GREEN if val == "Qualified" else colors.HexColor("#b3261e"))
                c.setFont("Helvetica-Bold", 11)
            else:
                c.setFillColor(DARK)
                c.setFont("Helvetica", 10.5)
            c.drawCentredString(x + lab_w + val_w / 2, ry - 20, str(val))
            # cell borders: one rect per label cell and value cell
            # (no internal dividers inside merged cells)
            c.setStrokeColor(LINE)
            c.rect(x, ry - rh, lab_w, rh, stroke=1, fill=0)
            c.rect(x + lab_w, ry - rh, val_w, rh, stroke=1, fill=0)
            x += cw
            ci += ncols

    # Blue footer band
    fh = 34
    c.setFillColor(BLUE)
    c.rect(bx + 14, bbot + 22, bw - 28, fh, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica", 10)
    c.drawCentredString(cy, bbot + 34,
                        f"\u00a9 {card['footer_year']}   Cholistan University of Veterinary & Animal Sciences, Bahawalpur")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


# ---------------- admin side ----------------
def admin_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if not session.get("admin"):
            return redirect(url_for("admin_login"))
        return f(*a, **kw)
    return wrapper


@app.route("/admin", methods=["GET", "POST"])
def admin_login():
    if session.get("admin"):
        return redirect(url_for("admin_dashboard"))
    error = None
    if request.method == "POST":
        if request.form.get("password") == app.config["ADMIN_PASSWORD"]:
            session["admin"] = True
            return redirect(url_for("admin_dashboard"))
        error = "Wrong password."
    return render_template("admin_login.html", uni=UNI_NAME, title=PORTAL_TITLE, error=error)


@app.route("/admin/logout")
def admin_logout():
    session.pop("admin", None)
    return redirect(url_for("admin_login"))


@app.route("/admin/dashboard")
@admin_required
def admin_dashboard():
    rows = get_db().fetchall("SELECT * FROM candidates ORDER BY name")
    rows = [dict(r) for r in rows]
    passing = get_setting("passing_percent", "50")
    published = get_setting("published", "1") == "1"
    return render_template("admin.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           rows=rows, passing=passing, published=published,
                           total=len(rows))


@app.route("/admin/settings", methods=["POST"])
@admin_required
def admin_settings():
    set_setting("passing_percent", request.form.get("passing_percent", "50"))
    set_setting("published", "1" if request.form.get("published") else "0")
    flash("Settings saved.")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/clear", methods=["POST"])
@admin_required
def admin_clear():
    db = get_db()
    db.execute("DELETE FROM candidates")
    db.commit()
    flash("All result records deleted.")
    return redirect(url_for("admin_dashboard"))


def header_key(header):
    h = re.sub(r"\s+", " ", str(header or "").strip().lower())
    for field, aliases in COLUMN_MAP.items():
        if h in aliases:
            return field
    return None


def read_upload(file_storage):
    """Read CSV/Excel straight from the upload stream (no disk writes,
    so it works on serverless hosts like Vercel)."""
    filename = file_storage.filename or ""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext == "csv":
        text = io.TextIOWrapper(file_storage.stream, encoding="utf-8-sig")
        return list(csv.DictReader(text))
    data = file_storage.read()
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.active
    data_rows = list(ws.values)
    if not data_rows:
        return []
    headers = [str(h) if h is not None else "" for h in data_rows[0]]
    return [dict(zip(headers, row)) for row in data_rows[1:] if any(row)]


@app.route("/admin/upload", methods=["POST"])
@admin_required
def admin_upload():
    file = request.files.get("file")
    if not file or not file.filename:
        flash("Please choose a file to upload.")
        return redirect(url_for("admin_dashboard"))
    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in ALLOWED_EXT:
        flash("Only CSV or Excel (.xlsx/.xls) files are allowed.")
        return redirect(url_for("admin_dashboard"))

    try:
        raw_rows = read_upload(file)
    except Exception as e:
        flash(f"Could not read file: {e}")
        return redirect(url_for("admin_dashboard"))
    if not raw_rows:
        flash("No data rows found in the file.")
        return redirect(url_for("admin_dashboard"))

    first = raw_rows[0]
    mapping = {}
    for hdr in first.keys():
        k = header_key(hdr)
        if k and k not in mapping:
            mapping[hdr] = k

    if "name" not in mapping.values() or "cnic" not in mapping.values():
        flash("File must contain at least a Name column and a CNIC column.")
        return redirect(url_for("admin_dashboard"))

    db = get_db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    imported, skipped = 0, 0
    cols = ("name", "father_name", "cnic", "roll_no", "discipline",
            "test_date", "obtained", "total", "merit_no", "uploaded_at")
    placeholders = ", ".join(["?"] * len(cols))
    if USE_PG:
        upsert = (f"INSERT INTO candidates ({', '.join(cols)}) VALUES ({placeholders}) "
                  "ON CONFLICT(cnic) DO UPDATE SET "
                  + ", ".join(f"{c}=EXCLUDED.{c}" for c in cols if c != "cnic"))
    else:
        upsert = f"INSERT OR REPLACE INTO candidates ({', '.join(cols)}) VALUES ({placeholders})"
    for rr in raw_rows:
        rec = {}
        for hdr, val in rr.items():
            field = mapping.get(hdr)
            if field:
                rec[field] = str(val).strip() if val is not None else ""
        cnic = norm_cnic(rec.get("cnic", ""))
        if not rec.get("name") or len(cnic) != 13:
            skipped += 1
            continue
        values = (rec.get("name"), rec.get("father_name"), cnic, rec.get("roll_no"),
                  rec.get("discipline"), rec.get("test_date"),
                  rec.get("obtained") or None, rec.get("total") or None,
                  rec.get("merit_no"), now)
        try:
            db.execute(upsert, values)
            imported += 1
        except Exception:
            skipped += 1
    db.commit()
    flash(f"Upload complete: {imported} records imported"
          + (f", {skipped} rows skipped (bad CNIC or missing name)." if skipped else "."))
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/template")
@admin_required
def admin_template():
    """Download a sample Excel file with the exact columns the upload expects."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    headers = ["Name", "Father's Name", "CNIC", "Applied Degree",
               "Held On", "Total Marks", "Obtained Marks"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0B6B3A")
        cell.alignment = Alignment(horizontal="center")
    ws.append(["Manzar Hayat", "Khizar Hayat", "36103-4695497-1",
               "PhD Zoology", "15-Sep-2025", 50, 45])
    ws.append(["Ayesha Khan", "Muhammad Khan", "35202-1234567-1",
               "PhD Biochemistry", "15-Sep-2025", 100, 72])
    widths = [22, 22, 20, 22, 16, 14, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name="cuvas_result_upload_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=True)
