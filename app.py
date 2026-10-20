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
from datetime import datetime, date
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
    "father_name": ["father name", "father's name", "father_name", "father", "fathers name", "s/o d/o"],
    "cnic": ["cnic", "cnic no", "cnic number", "nic", "national id"],
    "program": ["program", "programme", "degree program", "program level", "level"],
    "roll_no": ["roll no", "roll_no", "roll number", "rollnumber", "seat no"],
    "discipline": ["discipline", "applied degree", "degree",
                 "department", "subject", "field"],
    "test_date": ["test date", "test_date", "held on", "held_on", "date", "exam date"],
    "obtained": ["obtained", "obtained marks", "marks obtained", "marks", "score"],
    "total": ["total", "total marks", "max marks", "out of"],
    "merit_no": ["merit no", "merit_no", "merit number", "merit", "rank"],
}

UNI_NAME = "Cholistan University of Veterinary and Animal Sciences"
UNI_SHORT = "CUVAS, Bahawalpur"
PORTAL_TITLE = "Entry Test Result Portal"


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
        db.execute("ALTER TABLE candidates ADD COLUMN IF NOT EXISTS program TEXT DEFAULT 'PhD'")
        db.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                   " ON CONFLICT(key) DO NOTHING", ("published", "1"))
        # Migrate the old 50% default to the new 60% default (custom values untouched).
        db.execute("UPDATE settings SET value='60' WHERE key='passing_percent' AND value='50'")
        # Program-aware passing percentages (PhD 60%, MPhil 50%).
        old = db.fetchone("SELECT value FROM settings WHERE key='passing_percent'")
        phd_default = old["value"] if old else "60"
        db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                   " ON CONFLICT(key) DO NOTHING", ("phd_passing_percent", phd_default))
        db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                   " ON CONFLICT(key) DO NOTHING", ("mphil_passing_percent", "50"))
        db.execute("DELETE FROM settings WHERE key='passing_percent'")
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
        cols = [c[1] for c in db.fetchall("PRAGMA table_info(candidates)")]
        if "program" not in cols:
            db.execute("ALTER TABLE candidates ADD COLUMN program TEXT DEFAULT 'PhD'")
        db.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('published', '1')")
        # Program-aware passing percentages (PhD 60%, MPhil 50%).
        old = db.fetchone("SELECT value FROM settings WHERE key='passing_percent'")
        phd_default = old["value"] if old else "60"
        db.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('phd_passing_percent', ?)", (phd_default,))
        db.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('mphil_passing_percent', '50')")
        db.execute("DELETE FROM settings WHERE key='passing_percent'")
        # Migrate the old 50% default to the new 60% default (custom values untouched).
        db.execute("UPDATE settings SET value='60' WHERE key='phd_passing_percent' AND value='50'")
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


def norm_program(val):
    """Canonical program name: 'MPhil' or 'PhD' (default)."""
    v = str(val or "").strip().lower().replace(".", "")
    if v.startswith("mphil") or v.startswith("m phil"):
        return "MPhil"
    return "PhD"


def program_passing_pct(program):
    """Passing percentage for a program: PhD 60%, MPhil 50% (both adjustable)."""
    if (program or "PhD") == "MPhil":
        try:
            return float(get_setting("mphil_passing_percent", "50") or 50)
        except (TypeError, ValueError):
            return 50.0
    try:
        return float(get_setting("phd_passing_percent", "60") or 60)
    except (TypeError, ValueError):
        return 60.0


_DATE_FORMATS = ("%d-%b-%Y", "%d %b %Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y",
                 "%m/%d/%Y", "%b %d, %Y", "%d.%m.%Y", "%Y/%m/%d",
                 "%Y-%m-%d %H:%M:%S", "%d-%b-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S")


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


def build_card(r):
    """Assemble the result-card fields from a candidate record."""
    held = parse_date(r.get("test_date"))
    valid = add_one_year(held) if held else None
    pct = program_passing_pct(r.get("program"))
    passing_marks, status = calc_result(r.get("obtained"), r.get("total"), pct)
    obtained_s = fmt_num(r.get("obtained"))
    try:
        obt_f, tot_f = float(r.get("obtained")), float(r.get("total"))
        obtained_pct = (fmt_num(round(obt_f / tot_f * 100, 2)) + "%") if tot_f > 0 else ""
    except (TypeError, ValueError):
        obtained_pct = ""
    obtained_display = f"{obtained_s} ({obtained_pct})" if obtained_pct else obtained_s
    return {
        "name": r.get("name", ""),
        "father_name": r.get("father_name", ""),
        "cnic": fmt_cnic(r.get("cnic", "")),
        "cnic_raw": norm_cnic(r.get("cnic", "")),
        "degree": r.get("discipline", ""),
        "held_on": fmt_date(held) if held else (r.get("test_date") or ""),
        "total": fmt_num(r.get("total")),
        "obtained": obtained_s,
        "obtained_display": obtained_display,
        "passing_marks": passing_marks,
        "passing_pct": fmt_num(pct) + "%",
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
    # Per-program announcement banners (set in admin → Upload Results).
    notices = []
    for prog, msg_key, date_key in (("PhD", "phd_notice", "phd_notice_date"),
                                    ("MPhil", "mphil_notice", "mphil_notice_date")):
        msg = get_setting(msg_key, "").strip()
        if msg:
            d = parse_date(get_setting(date_key, "").strip())
            notices.append({"program": prog, "message": msg,
                            "date": fmt_date(d) if d else ""})
    return render_template("index.html", uni=UNI_NAME, short=UNI_SHORT,
                           title=PORTAL_TITLE, error=error, notices=notices)


@app.route("/result/<cnic>")
def result(cnic):
    cnic = norm_cnic(cnic)
    row = get_db().fetchone("SELECT * FROM candidates WHERE cnic=?", (cnic,))
    if not row:
        return render_template("notfound.html", uni=UNI_NAME, title=PORTAL_TITLE), 404
    r = dict(row)
    card = build_card(r)
    return render_template("result.html", uni=UNI_NAME, short=UNI_SHORT,
                           title=PORTAL_TITLE, card=card)


@app.route("/result/<cnic>/pdf")
def result_pdf(cnic):
    cnic = norm_cnic(cnic)
    row = get_db().fetchone("SELECT * FROM candidates WHERE cnic=?", (cnic,))
    if not row:
        return "Result not found.", 404
    r = dict(row)
    card = build_card(r)
    buf = build_pdf(card)
    filename = f"CUVAS_AdmissionTest_Result_{card['cnic_raw']}.pdf"
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


# ---------------- PDF generation ----------------
def build_pdf(card):
    """Result card PDF matching the web card exactly.

    The page itself is card-sized (no A4 whitespace): same layout,
    same proportions as the on-screen result card.
    """
    from reportlab.lib import colors
    from reportlab.pdfgen import canvas

    SOLID = colors.HexColor("#2d3290")
    DARK = colors.HexColor("#111111")
    LBL_BG = colors.HexColor("#eef0fb")
    LINE = colors.HexColor("#333333")
    RED = colors.HexColor("#b3261e")

    W, H = 760, 470  # card-sized page: same look as the web card
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(W, H))

    def band(x, y, w, h):
        c.setFillColor(SOLID)
        c.rect(x, y, w, h, stroke=0, fill=1)

    # Card border
    bx, bw = 14, W - 28
    btop, bbot = H - 14, 14
    c.setStrokeColor(DARK)
    c.setLineWidth(1.2)
    c.rect(bx, bbot, bw, btop - bbot, stroke=1, fill=0)

    # Header: official logo + university name
    lh = 46
    lw = lh * 270 / 148
    lx, ly = bx + 16, btop - 16 - lh
    logo_path = os.path.join(BASE_DIR, "static", "img", "cuvas-logo.png")
    if os.path.exists(logo_path):
        c.drawImage(logo_path, lx, ly, lw, lh, preserveAspectRatio=True)
    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(DARK)
    c.drawString(lx + lw + 12, ly + 17,
                 "Cholistan University of Veterinary & Animal Sciences, Bahawalpur")

    # Solid rule
    band(bx + 16, ly - 12, bw - 32, 3)

    # Title block
    cx = bx + bw / 2
    c.setFillColor(DARK)
    c.setFont("Helvetica", 12.5)
    c.drawCentredString(cx, ly - 42, "CUVAS - Admission Test (Postgraduate)")
    c.setFillColor(SOLID)
    c.setFont("Helvetica-Bold", 12.5)
    c.drawCentredString(cx, ly - 64, f"Held On:  {card['held_on']}")

    # Candidate fields (full width, two columns)
    fy = ly - 106
    col1_x = bx + 16
    col2_x = bx + 16 + (bw - 32) / 2
    val_dx = 104
    c.setFillColor(DARK)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(col1_x, fy, "Name:")
    c.drawString(col2_x, fy, "Father's Name:")
    c.drawString(col1_x, fy - 34, "CNIC:")
    c.drawString(col2_x, fy - 34, "Applied Degree:")
    c.setFont("Helvetica", 11)
    c.drawString(col1_x + val_dx, fy, card["name"])
    c.drawString(col2_x + val_dx, fy, card["father_name"])
    c.drawString(col1_x + val_dx, fy - 34, card["cnic"])
    c.drawString(col2_x + val_dx, fy - 34, card["degree"])

    # Marks table
    tx, tw = bx + 16, bw - 32
    ty, rh = fy - 66, 30
    pairs = [  # (label, label_w, value_w)
        [("Total Marks:", 124, 110), ("Obtained Marks:", 150, 86),
         ("Valid Till:", 100, 130)],
        [("Passing %:", 124, 110), ("Result Status:", 150, 316)],
    ]
    values = {"Total Marks:": card["total"], "Obtained Marks:": card["obtained_display"],
              "Valid Till:": card["valid_till"], "Passing %:": card["passing_pct"],
              "Result Status:": card["status"]}
    c.setLineWidth(1)
    for row_i, cells in enumerate(pairs):
        ry = ty - row_i * rh
        x = tx
        for lab, lab_w, val_w in cells:
            val = values[lab]
            c.setFillColor(LBL_BG)
            c.rect(x, ry - rh, lab_w, rh, stroke=0, fill=1)
            c.setFillColor(DARK)
            c.setFont("Helvetica-Bold", 11)
            c.drawCentredString(x + lab_w / 2, ry - 20, lab)
            c.setFillColor(colors.white)
            c.rect(x + lab_w, ry - rh, val_w, rh, stroke=0, fill=1)
            if lab == "Result Status:":
                c.setFillColor(SOLID if val == "Qualified" else RED)
                c.setFont("Helvetica-Bold", 11.5)
            else:
                c.setFillColor(DARK)
                c.setFont("Helvetica", 11)
            c.drawCentredString(x + lab_w + val_w / 2, ry - 20, str(val))
            c.setStrokeColor(LINE)
            c.rect(x, ry - rh, lab_w, rh, stroke=1, fill=0)
            c.rect(x + lab_w, ry - rh, val_w, rh, stroke=1, fill=0)
            x += lab_w + val_w

    # Disclaimer note (italic, centered) between table and footer
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    note_style = ParagraphStyle("note", fontName="Helvetica-Oblique", fontSize=10.5,
                                leading=14.5, alignment=TA_CENTER, textColor=DARK)
    note = Paragraph("NOTE: This is online generated result card. Errors and Omissions "
                     "are expected. If you have any query please contact at "
                     "das@cuvas.edu.pk.", note_style)
    _nw, nh = note.wrap(tw - 20, 120)
    note.drawOn(c, tx + 10, 140 - nh)

    # Solid footer band
    fh = 36
    band(bx, bbot, bw, fh)
    c.setFillColor(colors.white)
    c.setFont("Helvetica", 10.5)
    c.drawCentredString(cx, bbot + 13,
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
    db = get_db()
    rows = [dict(r) for r in db.fetchall("SELECT * FROM candidates")]
    total = len(rows)
    phd = mphil = qcount = 0
    for r in rows:
        prog = r.get("program") or "PhD"
        r["program"] = prog
        if prog == "MPhil":
            mphil += 1
        else:
            phd += 1
        _, r["status"] = calc_result(r.get("obtained"), r.get("total"),
                                     program_passing_pct(prog))
        if r["status"] == "Qualified":
            qcount += 1
    breakdown = []
    for prog in ("PhD", "MPhil"):
        pr = [r for r in rows if r["program"] == prog]
        pq = sum(1 for r in pr if r["status"] == "Qualified")
        breakdown.append({"program": prog, "total": len(pr), "qualified": pq,
                          "rate": round(pq / len(pr) * 100, 1) if pr else 0})
    recent = sorted(rows, key=lambda r: r.get("uploaded_at") or "", reverse=True)[:8]
    return render_template("admin_dashboard.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="dashboard", total=total, phd=phd, mphil=mphil,
                           qcount=qcount, nqcount=total - qcount,
                           published=get_setting("published", "1") == "1",
                           phd_pct=get_setting("phd_passing_percent", "60"),
                           mphil_pct=get_setting("mphil_passing_percent", "50"),
                           breakdown=breakdown, recent=recent)


@app.route("/admin/upload")
@admin_required
def admin_upload_page():
    return render_template("admin_upload.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="upload",
                           phd_pct=get_setting("phd_passing_percent", "60"),
                           mphil_pct=get_setting("mphil_passing_percent", "50"),
                           published=get_setting("published", "1") == "1",
                           phd_notice=get_setting("phd_notice", ""),
                           phd_notice_date=get_setting("phd_notice_date", ""),
                           mphil_notice=get_setting("mphil_notice", ""),
                           mphil_notice_date=get_setting("mphil_notice_date", ""))


def _clean_pct(v):
    s = str(v)
    return s.rstrip("0").rstrip(".") if "." in s else s


@app.route("/admin/settings", methods=["POST"])
@admin_required
def admin_settings():
    try:
        phd = float(request.form.get("phd_passing_percent", "60"))
        mphil = float(request.form.get("mphil_passing_percent", "50"))
    except (TypeError, ValueError):
        flash("Passing percentages must be numbers.")
        return redirect(url_for("admin_upload_page"))
    if not (1 <= phd <= 100 and 1 <= mphil <= 100):
        flash("Passing percentages must be between 1 and 100.")
        return redirect(url_for("admin_upload_page"))
    set_setting("phd_passing_percent", _clean_pct(phd))
    set_setting("mphil_passing_percent", _clean_pct(mphil))
    set_setting("published", "1" if request.form.get("published") else "0")
    set_setting("phd_notice", request.form.get("phd_notice", "").strip())
    set_setting("phd_notice_date", request.form.get("phd_notice_date", "").strip())
    set_setting("mphil_notice", request.form.get("mphil_notice", "").strip())
    set_setting("mphil_notice_date", request.form.get("mphil_notice_date", "").strip())
    flash("Settings saved.")
    return redirect(url_for("admin_upload_page"))


def _filtered_students(q, program, degree, status):
    """Students matching the records-page filters, with program + status computed."""
    db = get_db()
    clauses, params = [], []
    if q:
        like = f"%{q}%"
        clauses.append("(name LIKE ? OR cnic LIKE ? OR discipline LIKE ? OR father_name LIKE ?)")
        params += [like] * 4
    if program in ("PhD", "MPhil"):
        clauses.append("program=?")
        params.append(program)
    if degree and degree != "all":
        clauses.append("discipline=?")
        params.append(degree)
    sql = "SELECT * FROM candidates"
    if clauses:
        sql += " WHERE " + " AND ".join(clauses)
    sql += " ORDER BY name"
    rows = [dict(r) for r in db.fetchall(sql, params)]
    for r in rows:
        r["program"] = r.get("program") or "PhD"
        _, r["status"] = calc_result(r.get("obtained"), r.get("total"),
                                     program_passing_pct(r["program"]))
    if status in ("Qualified", "Not Qualified"):
        rows = [r for r in rows if r["status"] == status]
    return rows


def _degree_options():
    """Distinct applied-degree values for the degree filter."""
    db = get_db()
    rows = db.fetchall("SELECT DISTINCT discipline FROM candidates "
                       "WHERE discipline IS NOT NULL AND TRIM(discipline) != '' "
                       "ORDER BY discipline")
    return [r["discipline"] for r in rows]


@app.route("/admin/students")
@admin_required
def admin_students():
    q = request.args.get("q", "").strip()
    program = request.args.get("program", "all")
    degree = request.args.get("degree", "all")
    status = request.args.get("status", "all")
    page = request.args.get("page", 1, type=int) or 1
    per_page = 15

    rows = _filtered_students(q, program, degree, status)

    total = len(rows)
    pages = max(1, (total + per_page - 1) // per_page)
    page = min(max(page, 1), pages)
    page_rows = rows[(page - 1) * per_page:page * per_page]
    return render_template("admin_students.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="students", rows=page_rows, page=page, pages=pages,
                           total=total, q=q, program=program, degree=degree,
                           degrees=_degree_options(), status=status,
                           per_page=per_page)


@app.route("/admin/export")
@admin_required
def admin_export():
    """One-click Excel export of the currently filtered student records."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    q = request.args.get("q", "").strip()
    program = request.args.get("program", "all")
    degree = request.args.get("degree", "all")
    status = request.args.get("status", "all")
    rows = _filtered_students(q, program, degree, status)

    wb = Workbook()
    ws = wb.active
    ws.title = "Students"
    headers = ["Name", "Father's Name", "CNIC", "Program", "Applied Degree",
               "Held On", "Total Marks", "Obtained Marks", "Passing Marks",
               "Result Status", "Valid Till"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2D3290")
        cell.alignment = Alignment(horizontal="center")
    for r in rows:
        held = parse_date(r.get("test_date"))
        valid = add_one_year(held) if held else None
        pm, _ = calc_result(r.get("obtained"), r.get("total"),
                            program_passing_pct(r["program"]))
        ws.append([r.get("name"), r.get("father_name"), r.get("cnic"),
                   r.get("program"), r.get("discipline"),
                   fmt_date(held) if held else (r.get("test_date") or ""),
                   r.get("total"), r.get("obtained"), pm, r.get("status"),
                   fmt_date(valid) if valid else ""])
    widths = [22, 22, 20, 12, 24, 16, 14, 16, 14, 16, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name="cuvas_students_export.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/admin/export-pdf")
@admin_required
def admin_export_pdf():
    """Professional PDF list of the currently filtered student records."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, Image as RLImage)

    q = request.args.get("q", "").strip()
    program = request.args.get("program", "all")
    degree = request.args.get("degree", "all")
    status = request.args.get("status", "all")
    rows = _filtered_students(q, program, degree, status)

    INDIGO = colors.HexColor("#2d3290")
    DARK = colors.HexColor("#1a1f36")
    MUTED = colors.HexColor("#667085")
    OK = colors.HexColor("#12805c")
    BAD = colors.HexColor("#b3261e")

    buf = io.BytesIO()
    page = landscape(A4)
    doc = SimpleDocTemplate(buf, pagesize=page, leftMargin=14 * mm,
                            rightMargin=14 * mm, topMargin=14 * mm,
                            bottomMargin=14 * mm, title="CUVAS Student Result List")
    avail = page[0] - 28 * mm

    title_style = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=15,
                                 textColor=DARK, alignment=TA_LEFT, leading=18)
    sub_style = ParagraphStyle("s", fontName="Helvetica", fontSize=10,
                               textColor=MUTED, alignment=TA_LEFT, leading=13)
    filt_style = ParagraphStyle("f", fontName="Helvetica", fontSize=9,
                                textColor=DARK, alignment=TA_LEFT, leading=12)
    cell_style = ParagraphStyle("c", fontName="Helvetica", fontSize=8,
                                textColor=DARK, leading=10)
    head_style = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=8.5,
                                textColor=colors.white, leading=11)

    story = []
    # Header: logo + university name
    logo_path = os.path.join(BASE_DIR, "static", "img", "cuvas-logo.png")
    if os.path.exists(logo_path):
        logo = RLImage(logo_path, width=30 * mm, height=30 * mm * 148 / 270)
        head_tbl = Table([[logo, Paragraph(
            "Cholistan University of Veterinary &amp; Animal Sciences, Bahawalpur<br/>"
            "<font size=11 color=\"#2d3290\"><b>Entry Test Result Portal — Student Result List</b></font>",
            title_style)]], colWidths=[32 * mm, avail - 32 * mm])
        head_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                      ("LEFTPADDING", (1, 0), (1, 0), 6)]))
        story.append(head_tbl)
    else:
        story.append(Paragraph(
            "Cholistan University of Veterinary &amp; Animal Sciences, Bahawalpur",
            title_style))
        story.append(Paragraph("Entry Test Result Portal — Student Result List",
                               sub_style))
    story.append(Spacer(1, 4 * mm))

    # Filter summary
    filt_bits = []
    if program != "all":
        filt_bits.append(f"Program: <b>{program}</b>")
    if degree != "all":
        filt_bits.append(f"Degree: <b>{degree}</b>")
    if status != "all":
        filt_bits.append(f"Status: <b>{status}</b>")
    if q:
        filt_bits.append(f"Search: <b>{q}</b>")
    filt_line = " &nbsp;|&nbsp; ".join(filt_bits) if filt_bits else "All records"
    story.append(Paragraph(
        f"{filt_line} &nbsp;&nbsp;•&nbsp;&nbsp; Total: <b>{len(rows)}</b> student(s)"
        f" &nbsp;&nbsp;•&nbsp;&nbsp; Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}",
        filt_style))
    story.append(Spacer(1, 4 * mm))

    # Data table
    headers = ["#", "Name", "Father's Name", "CNIC", "Program", "Applied Degree",
               "Total", "Obtained", "Passing", "Status", "Held On"]
    col_w = [10 * mm, 36 * mm, 32 * mm, 30 * mm, 18 * mm, 42 * mm, 15 * mm,
             17 * mm, 15 * mm, 22 * mm, 28 * mm]
    data = [[Paragraph(h, head_style) for h in headers]]
    for i, r in enumerate(rows, start=1):
        held = parse_date(r.get("test_date"))
        pm, _ = calc_result(r.get("obtained"), r.get("total"),
                            program_passing_pct(r["program"]))
        st = r.get("status", "")
        st_color = "#12805c" if st == "Qualified" else "#b3261e"
        data.append([
            Paragraph(str(i), cell_style),
            Paragraph(str(r.get("name") or ""), cell_style),
            Paragraph(str(r.get("father_name") or ""), cell_style),
            Paragraph(str(r.get("cnic") or ""), cell_style),
            Paragraph(str(r.get("program") or ""), cell_style),
            Paragraph(str(r.get("discipline") or ""), cell_style),
            Paragraph(str(r.get("total") or ""), cell_style),
            Paragraph(str(r.get("obtained") or ""), cell_style),
            Paragraph(str(pm), cell_style),
            Paragraph(f'<font color="{st_color}"><b>{st}</b></font>', cell_style),
            Paragraph(fmt_date(held) if held else str(r.get("test_date") or ""),
                      cell_style),
        ])
    tbl = Table(data, colWidths=col_w, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), INDIGO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d0d5e8")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f6fc")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)

    def _footer(canv, _doc):
        canv.saveState()
        canv.setFont("Helvetica", 8)
        canv.setFillColor(MUTED)
        canv.drawString(14 * mm, 10 * mm,
                        "Cholistan University of Veterinary & Animal Sciences, Bahawalpur")
        canv.drawRightString(page[0] - 14 * mm, 10 * mm, f"Page {_doc.page}")
        canv.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name="cuvas_students_list.pdf",
                     mimetype="application/pdf")


@app.route("/admin/students/<int:cid>/edit", methods=["GET", "POST"])
@admin_required
def admin_edit_student(cid):
    db = get_db()
    row = db.fetchone("SELECT * FROM candidates WHERE id=?", (cid,))
    if not row:
        flash("Student not found.")
        return redirect(url_for("admin_students"))
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        father = request.form.get("father_name", "").strip()
        cnic = norm_cnic(request.form.get("cnic", ""))
        program = norm_program(request.form.get("program"))
        discipline = request.form.get("discipline", "").strip()
        test_date = request.form.get("test_date", "").strip()
        try:
            total = float(request.form.get("total") or 0)
            obtained = float(request.form.get("obtained") or 0)
        except ValueError:
            flash("Marks must be numbers.")
            return redirect(url_for("admin_edit_student", cid=cid))
        if not name or len(cnic) != 13:
            flash("Name and a valid 13-digit CNIC are required.")
            return redirect(url_for("admin_edit_student", cid=cid))
        dup = db.fetchone("SELECT id FROM candidates WHERE cnic=? AND id<>?", (cnic, cid))
        if dup:
            flash("Another student already uses this CNIC.")
            return redirect(url_for("admin_edit_student", cid=cid))
        db.execute("UPDATE candidates SET name=?, father_name=?, cnic=?, program=?,"
                   " discipline=?, test_date=?, obtained=?, total=? WHERE id=?",
                   (name, father, cnic, program, discipline, test_date,
                    obtained or None, total or None, cid))
        db.commit()
        flash("Student record updated.")
        return redirect(url_for("admin_students"))
    r = dict(row)
    r["program"] = r.get("program") or "PhD"
    _, r["status"] = calc_result(r.get("obtained"), r.get("total"),
                                 program_passing_pct(r["program"]))
    return render_template("admin_edit.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="students", student=r,
                           phd_pct=get_setting("phd_passing_percent", "60"),
                           mphil_pct=get_setting("mphil_passing_percent", "50"))


@app.route("/admin/students/<int:cid>/delete", methods=["POST"])
@admin_required
def admin_delete_student(cid):
    db = get_db()
    db.execute("DELETE FROM candidates WHERE id=?", (cid,))
    db.commit()
    flash("Student record deleted.")
    return redirect(url_for("admin_students"))


@app.route("/admin/clear", methods=["POST"])
@admin_required
def admin_clear():
    db = get_db()
    db.execute("DELETE FROM candidates")
    db.commit()
    flash("All result records deleted.")
    return redirect(url_for("admin_students"))


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
        return redirect(url_for("admin_upload_page"))
    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in ALLOWED_EXT:
        flash("Only CSV or Excel (.xlsx/.xls) files are allowed.")
        return redirect(url_for("admin_upload_page"))

    try:
        raw_rows = read_upload(file)
    except Exception as e:
        flash(f"Could not read file: {e}")
        return redirect(url_for("admin_upload_page"))
    if not raw_rows:
        flash("No data rows found in the file.")
        return redirect(url_for("admin_upload_page"))

    first = raw_rows[0]
    mapping = {}
    for hdr in first.keys():
        k = header_key(hdr)
        if k and k not in mapping:
            mapping[hdr] = k

    if "name" not in mapping.values() or "cnic" not in mapping.values():
        flash("File must contain at least a Name column and a CNIC column.")
        return redirect(url_for("admin_upload_page"))

    db = get_db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    imported, skipped = 0, 0
    cols = ("name", "father_name", "cnic", "program", "roll_no", "discipline",
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
                # Excel often gives real datetime objects (with 00:00:00 time);
                # store them as clean dates so no timestamp leaks onto the card.
                if isinstance(val, datetime):
                    val = val.strftime("%d-%b-%Y")
                elif isinstance(val, date):
                    val = val.strftime("%d-%b-%Y")
                if field == "program":
                    val = norm_program(val)
                rec[field] = str(val).strip() if val is not None else ""
        cnic = norm_cnic(rec.get("cnic", ""))
        if not rec.get("name") or len(cnic) != 13:
            skipped += 1
            continue
        values = (rec.get("name"), rec.get("father_name"), cnic,
                  rec.get("program") or "PhD", rec.get("roll_no"),
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
    return redirect(url_for("admin_upload_page"))


@app.route("/admin/template")
@admin_required
def admin_template():
    """Download a sample Excel file with the exact columns the upload expects."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    headers = ["Name", "Father's Name", "CNIC", "Program", "Applied Degree",
               "Held On", "Total Marks", "Obtained Marks"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2D3290")
        cell.alignment = Alignment(horizontal="center")
    ws.append(["Manzar Hayat", "Khizar Hayat", "36103-4695497-1", "PhD",
               "PhD Zoology", "15-Sep-2025", 50, 45])
    ws.append(["Ayesha Khan", "Muhammad Khan", "35202-1234567-1", "MPhil",
               "MPhil Biochemistry", "15-Sep-2025", 100, 72])
    widths = [22, 22, 20, 12, 22, 16, 14, 16]
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
