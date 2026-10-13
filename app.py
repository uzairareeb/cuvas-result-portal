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
    "discipline": ["discipline", "department", "subject", "program", "programme", "field"],
    "test_date": ["test date", "test_date", "date", "exam date"],
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
                   " ON CONFLICT(key) DO NOTHING", ("passing_percent", "50"))
        db.execute("INSERT INTO settings(key, value) VALUES (?, ?)"
                   " ON CONFLICT(key) DO NOTHING", ("published", "1"))
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
        db.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('passing_percent', '50')")
        db.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('published', '1')")
    db.commit()
    db.close()


_db_ready = False


@app.before_request
def _ensure_db():
    # On Vercel there is no startup hook; create tables lazily on first request.
    global _db_ready
    if not _db_ready:
        init_db()
        _db_ready = True


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


def calc_result(obtained, total, passing_percent):
    try:
        pct = (float(obtained) / float(total)) * 100 if float(total) else 0.0
    except (TypeError, ValueError):
        pct = 0.0
    status = "PASS" if pct >= float(passing_percent) else "FAIL"
    return round(pct, 2), status


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
    passing = float(get_setting("passing_percent", "50") or 50)
    pct, status = calc_result(r["obtained"], r["total"], passing)
    return render_template("result.html", uni=UNI_NAME, short=UNI_SHORT,
                           title=PORTAL_TITLE, r=r, pct=pct, status=status,
                           passing=passing)


@app.route("/result/<cnic>/pdf")
def result_pdf(cnic):
    cnic = norm_cnic(cnic)
    row = get_db().fetchone("SELECT * FROM candidates WHERE cnic=?", (cnic,))
    if not row:
        return "Result not found.", 404
    r = dict(row)
    passing = float(get_setting("passing_percent", "50") or 50)
    pct, status = calc_result(r["obtained"], r["total"], passing)
    buf = build_pdf(r, pct, status, passing)
    filename = f"CUVAS_PhD_EntryTest_Result_{r['roll_no'] or cnic}.pdf"
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


# ---------------- PDF generation ----------------
def build_pdf(r, pct, status, passing):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    GREEN = colors.HexColor("#0b6b3a")
    LIGHT = colors.HexColor("#eef7f1")
    DARK = colors.HexColor("#1a1a1a")
    RED = colors.HexColor("#b3261e")

    c.setFillColor(GREEN)
    c.rect(0, H - 58 * mm, W, 58 * mm, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(W / 2, H - 22 * mm, UNI_NAME.upper())
    c.setFont("Helvetica", 10)
    c.drawCentredString(W / 2, H - 30 * mm, UNI_SHORT + "  |  Directorate of Advanced Studies")
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(W / 2, H - 42 * mm, "PhD ENTRY TEST RESULT CARD")
    if r.get("test_date"):
        c.setFont("Helvetica", 9)
        c.drawCentredString(W / 2, H - 50 * mm, f"Test Date: {r['test_date']}")

    y = H - 72 * mm
    x0, box_w = 22 * mm, W - 44 * mm
    rows = [
        ("Candidate Name", r.get("name", "")),
        ("Father Name", r.get("father_name", "")),
        ("CNIC", r.get("cnic", "")),
        ("Roll No", r.get("roll_no", "")),
        ("Discipline / Department", r.get("discipline", "")),
        ("Marks Obtained", f"{r.get('obtained','')} / {r.get('total','')}"),
        ("Percentage", f"{pct}%"),
        ("Passing Criteria", f"{passing}%"),
    ]
    if r.get("merit_no"):
        rows.append(("Merit No", r["merit_no"]))

    c.setFillColor(LIGHT)
    c.roundRect(x0, y - len(rows) * 10 * mm - 8 * mm, box_w,
                len(rows) * 10 * mm + 8 * mm, 4 * mm, stroke=0, fill=1)
    for i, (label, value) in enumerate(rows):
        ry = y - i * 10 * mm
        c.setFillColor(colors.HexColor("#555555"))
        c.setFont("Helvetica", 10)
        c.drawString(x0 + 8 * mm, ry, label)
        c.setFillColor(DARK)
        c.setFont("Helvetica-Bold", 10.5)
        c.drawString(x0 + 62 * mm, ry, str(value))
        if i < len(rows) - 1:
            c.setStrokeColor(colors.HexColor("#d9e8de"))
            c.line(x0 + 6 * mm, ry - 3.5 * mm, x0 + box_w - 6 * mm, ry - 3.5 * mm)

    banner_y = y - len(rows) * 10 * mm - 26 * mm
    banner_color = GREEN if status == "PASS" else RED
    c.setFillColor(banner_color)
    c.roundRect(x0, banner_y, box_w, 20 * mm, 4 * mm, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(W / 2, banner_y + 7.5 * mm, status)

    c.setFillColor(colors.HexColor("#777777"))
    c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, 22 * mm,
                        "This is a computer-generated result card. Verify with the Directorate of Advanced Studies, CUVAS.")
    c.drawCentredString(W / 2, 16 * mm,
                        f"Generated on {datetime.now().strftime('%d-%b-%Y %I:%M %p')}")
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


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=True)
