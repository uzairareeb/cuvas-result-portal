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
import hashlib
import hmac
import io
import json
import os
import re
import sqlite3
from datetime import datetime, date, timedelta
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
                 "department", "subject", "field", "degree/discipline",
                 "discipline/degree", "degree / discipline"],
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
    # ---- admin portal upgrade tables (both backends) ----
    idcol = "SERIAL PRIMARY KEY" if USE_PG else "INTEGER PRIMARY KEY AUTOINCREMENT"
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS audit_log (
            id {idcol},
            ts TEXT NOT NULL,
            action TEXT NOT NULL,
            detail TEXT
        )
    """)
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS backups (
            id {idcol},
            ts TEXT NOT NULL,
            reason TEXT NOT NULL,
            count INTEGER NOT NULL,
            data TEXT NOT NULL
        )
    """)
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS uploads (
            id {idcol},
            ts TEXT NOT NULL,
            filename TEXT NOT NULL,
            imported INTEGER NOT NULL DEFAULT 0,
            updated INTEGER NOT NULL DEFAULT 0,
            skipped INTEGER NOT NULL DEFAULT 0
        )
    """)
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS upload_staging (
            id {idcol},
            batch TEXT NOT NULL,
            row_no INTEGER NOT NULL,
            status TEXT NOT NULL,
            reason TEXT,
            data TEXT
        )
    """)
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS lookups (
            id {idcol},
            ts TEXT NOT NULL,
            cnic TEXT NOT NULL,
            found INTEGER NOT NULL DEFAULT 0
        )
    """)
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS login_attempts (
            id {idcol},
            ts_epoch DOUBLE PRECISION NOT NULL,
            ip TEXT NOT NULL
        )
    """)
    # ---- test sessions: sessions table + candidates.session_id ----
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS sessions (
            id {idcol},
            name TEXT NOT NULL UNIQUE,
            created_at TEXT,
            is_active INTEGER NOT NULL DEFAULT 0
        )
    """)
    if USE_PG:
        db.execute("ALTER TABLE candidates ADD COLUMN IF NOT EXISTS session_id INTEGER")
        db.execute("ALTER TABLE candidates DROP CONSTRAINT IF EXISTS candidates_cnic_key")
        db.execute("CREATE UNIQUE INDEX IF NOT EXISTS candidates_cnic_session_uidx"
                   " ON candidates(cnic, session_id)")
        db.execute("ALTER TABLE upload_staging ADD COLUMN IF NOT EXISTS warning TEXT")
    else:
        cols = [c[1] for c in db.fetchall("PRAGMA table_info(candidates)")]
        if "session_id" not in cols:
            db.execute("ALTER TABLE candidates ADD COLUMN session_id INTEGER")
            cols.append("session_id")
        srow = db.fetchone("SELECT sql FROM sqlite_master WHERE type='table'"
                           " AND name='candidates'")
        csql = (srow["sql"] if isinstance(srow, dict) else srow[0]) or ""
        if "cnic TEXT NOT NULL UNIQUE" in csql:
            # Rebuild: UNIQUE(cnic) alone blocks the same CNIC in a later session.
            keep = [c for c in ("id", "name", "father_name", "cnic", "program",
                                "roll_no", "discipline", "test_date", "obtained",
                                "total", "merit_no", "uploaded_at", "session_id")
                    if c in cols]
            db.execute("""
                CREATE TABLE candidates_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    father_name TEXT,
                    cnic TEXT NOT NULL,
                    program TEXT DEFAULT 'PhD',
                    roll_no TEXT,
                    discipline TEXT,
                    test_date TEXT,
                    obtained REAL,
                    total REAL,
                    merit_no TEXT,
                    uploaded_at TEXT,
                    session_id INTEGER,
                    UNIQUE(cnic, session_id)
                )
            """)
            collist = ", ".join(keep)
            db.execute(f"INSERT INTO candidates_new ({collist}) SELECT {collist}"
                       " FROM candidates")
            db.execute("DROP TABLE candidates")
            db.execute("ALTER TABLE candidates_new RENAME TO candidates")
        scols = [c[1] for c in db.fetchall("PRAGMA table_info(upload_staging)")]
        if "warning" not in scols:
            db.execute("ALTER TABLE upload_staging ADD COLUMN warning TEXT")
    # One default session for all pre-session records; exactly one active.
    srow = db.fetchone("SELECT id FROM sessions ORDER BY id LIMIT 1")
    if not srow:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        if USE_PG:
            db.execute("INSERT INTO sessions(name, created_at, is_active)"
                       " VALUES (?, ?, 1) RETURNING id", ("Entry Test 2025", now))
            sid = db.fetchone("SELECT id FROM sessions WHERE is_active=1")["id"]
        else:
            cur = db.execute("INSERT INTO sessions(name, created_at, is_active)"
                             " VALUES (?, ?, 1)", ("Entry Test 2025", now))
            sid = cur.lastrowid
    else:
        sid = srow["id"] if isinstance(srow, dict) else srow[0]
    db.execute("UPDATE candidates SET session_id=? WHERE session_id IS NULL", (sid,))
    arow = db.fetchone("SELECT id FROM sessions WHERE is_active=1 LIMIT 1")
    if not arow:
        db.execute("UPDATE sessions SET is_active=1 WHERE id=?", (sid,))
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


# ---------------- admin helpers: audit, backups, security ----------------
def _now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def log_audit(action, detail=""):
    """Record an admin action. Never let logging break the main flow."""
    try:
        db = get_db()
        db.execute("INSERT INTO audit_log(ts, action, detail) VALUES (?, ?, ?)",
                   (_now_str(), action, detail))
        db.commit()
    except Exception:
        pass


def make_backup(reason):
    """Snapshot ALL candidate records (JSON) so a destructive action can be
    undone from Admin -> Backups. Returns the number of records saved."""
    try:
        db = get_db()
        rows = [dict(r) for r in db.fetchall("SELECT * FROM candidates")]
        db.execute("INSERT INTO backups(ts, reason, count, data) VALUES (?, ?, ?, ?)",
                   (_now_str(), reason, len(rows), json.dumps(rows)))
        db.commit()
        return len(rows)
    except Exception:
        return 0


def obtained_pct(obtained, total):
    """Obtained marks as a percentage of total (None when not computable)."""
    try:
        t, o = float(total), float(obtained)
    except (TypeError, ValueError):
        return None
    if t <= 0:
        return None
    return round(o / t * 100, 1)


# ---------------- test sessions ----------------
def get_sessions():
    try:
        return [dict(r) for r in get_db().fetchall(
            "SELECT * FROM sessions ORDER BY id")]
    except Exception:
        return []


def active_session():
    """The session students see by default (exactly one is active)."""
    rows = get_sessions()
    for s in rows:
        if s.get("is_active"):
            return s
    return rows[0] if rows else None


def session_name(sid):
    for s in get_sessions():
        if s["id"] == sid:
            return s["name"]
    return ""


def get_or_create_session(name):
    """Find a session by name or create it; a NEW session becomes the active
    one (the previous session stays as an archive students can select)."""
    db = get_db()
    name = (name or "").strip() or "Entry Test"
    row = db.fetchone("SELECT * FROM sessions WHERE name=?", (name,))
    if row:
        return dict(row), False
    db.execute("UPDATE sessions SET is_active=0")
    now = _now_str()
    if USE_PG:
        db.execute("INSERT INTO sessions(name, created_at, is_active)"
                   " VALUES (?, ?, 1)", (name, now))
    else:
        db.execute("INSERT INTO sessions(name, created_at, is_active)"
                   " VALUES (?, ?, 1)", (name, now))
    db.commit()
    return dict(db.fetchone("SELECT * FROM sessions WHERE name=?", (name,))), True


# ---------------- QR verification ----------------
def verify_url(cnic, sid=None):
    base = request.host_url
    if base.startswith("http://") and "localhost" not in base and "127.0.0.1" not in base:
        base = "https://" + base[len("http://"):]
    url = f"{base}verify/{cnic}"
    if sid:
        url += f"?s={sid}"
    return url


def make_qr_png(url):
    import qrcode
    qr = qrcode.QRCode(box_size=6, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#1a1f36", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def _client_ip():
    fwd = request.headers.get("X-Forwarded-For", "")
    if fwd:
        return fwd.split(",")[0].strip()
    return request.remote_addr or "unknown"


LOGIN_MAX_FAILS = 5
LOGIN_WINDOW_MIN = 15


def _login_locked(ip):
    """Locked when LOGIN_MAX_FAILS failed attempts happened inside the window."""
    try:
        db = get_db()
        cutoff = (datetime.now() - timedelta(minutes=LOGIN_WINDOW_MIN)).timestamp()
        row = db.fetchone("SELECT COUNT(*) AS c FROM login_attempts "
                          "WHERE ip=? AND ts_epoch>=?", (ip, cutoff))
        return (row["c"] if isinstance(row, dict) else row[0]) >= LOGIN_MAX_FAILS
    except Exception:
        return False


def _record_login_fail(ip):
    try:
        db = get_db()
        db.execute("INSERT INTO login_attempts(ts_epoch, ip) VALUES (?, ?)",
                   (datetime.now().timestamp(), ip))
        db.commit()
    except Exception:
        pass


def _clear_login_fails(ip):
    try:
        db = get_db()
        db.execute("DELETE FROM login_attempts WHERE ip=?", (ip,))
        db.commit()
    except Exception:
        pass


def _password_ok(pw):
    """The in-app password (hashed, set from Admin -> Change Password) wins;
    otherwise fall back to the ADMIN_PASSWORD environment value."""
    stored = ""
    try:
        stored = get_setting("admin_password_hash", "")
    except Exception:
        stored = ""
    if stored:
        digest = hashlib.sha256((pw or "").encode("utf-8")).hexdigest()
        return hmac.compare_digest(digest, stored)
    return hmac.compare_digest(str(pw or ""), str(app.config["ADMIN_PASSWORD"]))


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
        "session_id": r.get("session_id"),
        "session_name": session_name(r.get("session_id")) if r.get("session_id") else "",
        "verify_url": verify_url(norm_cnic(r.get("cnic", "")), r.get("session_id")),
        "status": status,
        "valid_till": fmt_date(valid),
        "footer_year": str(valid.year) if valid else str(datetime.now().year),
    }


# ---------------- student side ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if get_setting("published", "1") != "1":
        ann = None
        if get_setting("announce_enabled", "0") == "1":
            raw = get_setting("announce_at", "").strip()
            when = None
            try:
                when = datetime.strptime(raw, "%Y-%m-%dT%H:%M")
            except (TypeError, ValueError):
                when = None
            if when is not None:
                ann = {"title": get_setting("announce_title", "").strip()
                       or "Result Announcement",
                       "message": get_setting("announce_message", "").strip(),
                       "iso": raw + ":00+05:00",
                       "when_disp": when.strftime("%d %B %Y, %I:%M %p")}
        return render_template("unpublished.html", uni=UNI_NAME,
                               title=PORTAL_TITLE, announce=ann)
    error = None
    if request.method == "POST":
        cnic = norm_cnic(request.form.get("cnic", ""))
        if len(cnic) != 13:
            error = "Please enter a valid 13-digit CNIC number."
        else:
            sid = request.form.get("session", type=int)
            return redirect(url_for("result", cnic=cnic, s=sid))
    # Per-program announcement banners (set in admin → Upload Results).
    notices = []
    for prog, msg_key, date_key in (("PhD", "phd_notice", "phd_notice_date"),
                                    ("MPhil", "mphil_notice", "mphil_notice_date")):
        msg = get_setting(msg_key, "").strip()
        if msg:
            d = parse_date(get_setting(date_key, "").strip())
            notices.append({"program": prog, "message": msg,
                            "date": fmt_date(d) if d else ""})
    sess = get_sessions()
    act = active_session()
    return render_template("index.html", uni=UNI_NAME, short=UNI_SHORT,
                           title=PORTAL_TITLE, error=error, notices=notices,
                           sessions=sess, active_id=act["id"] if act else None)


@app.route("/result/<cnic>")
def result(cnic):
    cnic = norm_cnic(cnic)
    sid = request.args.get("s", type=int)
    if not sid:
        act = active_session()
        sid = act["id"] if act else None
    db = get_db()
    # Strict: a result is only ever shown from the selected session. A record
    # in another session must NOT appear here (the student picks that session).
    row = db.fetchone("SELECT * FROM candidates WHERE cnic=? AND session_id=?",
                      (cnic, sid))
    try:  # lookup analytics for the admin Lookup Log (never breaks the page)
        db = get_db()
        db.execute("INSERT INTO lookups(ts, cnic, found) VALUES (?, ?, ?)",
                   (_now_str(), cnic, 1 if row else 0))
        db.commit()
    except Exception:
        pass
    if not row:
        return render_template("notfound.html", uni=UNI_NAME, title=PORTAL_TITLE,
                               cnic_disp=fmt_cnic(cnic),
                               sess_name=session_name(sid) if sid else ""), 404
    r = dict(row)
    card = build_card(r)
    return render_template("result.html", uni=UNI_NAME, short=UNI_SHORT,
                           title=PORTAL_TITLE, card=card)


@app.route("/result/<cnic>/pdf")
def result_pdf(cnic):
    cnic = norm_cnic(cnic)
    sid = request.args.get("s", type=int)
    if not sid:
        act = active_session()
        sid = act["id"] if act else None
    db = get_db()
    row = db.fetchone("SELECT * FROM candidates WHERE cnic=? AND session_id=?",
                      (cnic, sid))
    if not row:
        return "Result not found.", 404
    r = dict(row)
    card = build_card(r)
    buf = build_pdf(card)
    filename = f"CUVAS_AdmissionTest_Result_{card['cnic_raw']}.pdf"
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


@app.route("/qr/<cnic>")
def qr_code(cnic):
    """QR code image encoding this record's public verification URL."""
    cnic = norm_cnic(cnic)
    sid = request.args.get("s", type=int)
    if not sid:
        act = active_session()
        sid = act["id"] if act else None
    row = get_db().fetchone("SELECT session_id FROM candidates WHERE cnic=?"
                            " AND session_id=?", (cnic, sid))
    if row:
        sid = row["session_id"] if isinstance(row, dict) else row[0]
    resp = send_file(make_qr_png(verify_url(cnic, sid)), mimetype="image/png")
    resp.headers["Cache-Control"] = "public, max-age=86400"
    return resp


@app.route("/verify/<cnic>")
def verify(cnic):
    """Public verification page a QR scan lands on: confirms whether this
    CNIC has an official record and shows the same summary as the card."""
    cnic = norm_cnic(cnic)
    sid = request.args.get("s", type=int)
    db = get_db()
    row = None
    if sid:
        # The QR names its session: verify strictly against that session.
        row = db.fetchone("SELECT * FROM candidates WHERE cnic=? AND session_id=?",
                          (cnic, sid))
    else:
        row = db.fetchone("SELECT * FROM candidates WHERE cnic=?"
                          " ORDER BY session_id DESC", (cnic,))
    card = None
    if row:
        card = build_card(dict(row))
    return render_template("verify.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           card=card, cnic_disp=fmt_cnic(cnic)), (200 if card else 404)


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

    # QR verification code (bottom-right, clear of the note and the footer)
    try:
        from reportlab.lib.utils import ImageReader
        qs = 46
        qx, qy = bx + bw - 16 - qs, bbot + 44
        c.drawImage(ImageReader(make_qr_png(card["verify_url"])), qx, qy, qs, qs)
        c.setFillColor(DARK)
        c.setFont("Helvetica", 8.5)
        c.drawRightString(qx - 8, qy + qs / 2 - 3, "Scan to verify this result")
    except Exception:
        pass

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
        ip = _client_ip()
        if _login_locked(ip):
            error = ("Too many failed attempts. Please wait "
                     f"{LOGIN_WINDOW_MIN} minutes and try again.")
            return render_template("admin_login.html", uni=UNI_NAME,
                                   title=PORTAL_TITLE, error=error)
        if _password_ok(request.form.get("password")):
            session["admin"] = True
            _clear_login_fails(ip)
            log_audit("Admin login", f"IP {ip}")
            return redirect(url_for("admin_dashboard"))
        _record_login_fail(ip)
        log_audit("Failed login attempt", f"IP {ip}")
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
    act = active_session()
    if act:
        rows = [dict(r) for r in db.fetchall(
            "SELECT * FROM candidates WHERE session_id=?", (act["id"],))]
    else:
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

    # ---- insights: discipline stats, marks distribution, top scorers ----
    for r in rows:
        r["pct"] = obtained_pct(r.get("obtained"), r.get("total"))
    disc = {}
    for r in rows:
        d = (r.get("discipline") or "—").strip() or "—"
        s = disc.setdefault(d, {"degree": d, "total": 0, "qualified": 0})
        s["total"] += 1
        if r["status"] == "Qualified":
            s["qualified"] += 1
    disc_stats = sorted(disc.values(), key=lambda s: -s["total"])[:8]
    for s in disc_stats:
        s["rate"] = round(s["qualified"] / s["total"] * 100, 1) if s["total"] else 0
    buckets = [("<40%", 0, 40), ("40–49%", 40, 50), ("50–59%", 50, 60),
               ("60–69%", 60, 70), ("70–79%", 70, 80), ("80–89%", 80, 90),
               ("90–100%", 90, 100.01)]
    dist = []
    for label, lo, hi in buckets:
        n = sum(1 for r in rows if r["pct"] is not None and lo <= r["pct"] < hi)
        dist.append({"label": label, "count": n})
    dist_max = max((d["count"] for d in dist), default=0) or 1
    for d in dist:
        d["width"] = round(d["count"] / dist_max * 100, 1)
    top_scorers = sorted([r for r in rows if r["pct"] is not None],
                         key=lambda r: (-r["pct"], r.get("name") or ""))[:5]
    pass_rate = round(qcount / total * 100, 1) if total else 0
    # Donut geometry: r=52, circumference ~= 326.7
    donut_dash = round(pass_rate / 100 * 326.7, 1)
    try:
        lk = db.fetchone("SELECT COUNT(*) AS c FROM lookups")
        lk_total = lk["c"] if isinstance(lk, dict) else lk[0]
        lkf = db.fetchone("SELECT COUNT(*) AS c FROM lookups WHERE found=0")
        lk_fails = lkf["c"] if isinstance(lkf, dict) else lkf[0]
    except Exception:
        lk_total, lk_fails = 0, 0
    return render_template("admin_dashboard.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="dashboard", total=total, phd=phd, mphil=mphil,
                           qcount=qcount, nqcount=total - qcount,
                           published=get_setting("published", "1") == "1",
                           phd_pct=get_setting("phd_passing_percent", "60"),
                           mphil_pct=get_setting("mphil_passing_percent", "50"),
                           breakdown=breakdown, recent=recent,
                           disc_stats=disc_stats, dist=dist,
                           top_scorers=top_scorers, pass_rate=pass_rate,
                           donut_dash=donut_dash, lk_total=lk_total,
                           lk_fails=lk_fails,
                           session_name=act["name"] if act else "")


@app.route("/admin/upload")
@admin_required
def admin_upload_page():
    try:
        history = [dict(r) for r in get_db().fetchall(
            "SELECT * FROM uploads ORDER BY id DESC LIMIT 10")]
    except Exception:
        history = []
    act = active_session()
    return render_template("admin_upload.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="upload", sessions=get_sessions(),
                           active_session_name=act["name"] if act else "",
                           phd_pct=get_setting("phd_passing_percent", "60"),
                           mphil_pct=get_setting("mphil_passing_percent", "50"),
                           published=get_setting("published", "1") == "1",
                           phd_notice=get_setting("phd_notice", ""),
                           phd_notice_date=get_setting("phd_notice_date", ""),
                           mphil_notice=get_setting("mphil_notice", ""),
                           mphil_notice_date=get_setting("mphil_notice_date", ""),
                           announce_enabled=get_setting("announce_enabled", "0") == "1",
                           announce_title=get_setting("announce_title", ""),
                           announce_message=get_setting("announce_message", ""),
                           announce_at=get_setting("announce_at", ""),
                           history=history)


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
        flash("Passing percentages must be numbers.", "err")
        return redirect(url_for("admin_upload_page"))
    if not (1 <= phd <= 100 and 1 <= mphil <= 100):
        flash("Passing percentages must be between 1 and 100.", "err")
        return redirect(url_for("admin_upload_page"))
    set_setting("phd_passing_percent", _clean_pct(phd))
    set_setting("mphil_passing_percent", _clean_pct(mphil))
    set_setting("published", "1" if request.form.get("published") else "0")
    set_setting("phd_notice", request.form.get("phd_notice", "").strip())
    set_setting("phd_notice_date", request.form.get("phd_notice_date", "").strip())
    set_setting("mphil_notice", request.form.get("mphil_notice", "").strip())
    set_setting("mphil_notice_date", request.form.get("mphil_notice_date", "").strip())
    set_setting("announce_enabled", "1" if request.form.get("announce_enabled") else "0")
    set_setting("announce_title", request.form.get("announce_title", "").strip())
    set_setting("announce_message", request.form.get("announce_message", "").strip())
    set_setting("announce_at", request.form.get("announce_at", "").strip())
    log_audit("Settings updated",
              f"PhD passing {_clean_pct(phd)}%, MPhil passing {_clean_pct(mphil)}%, "
              f"portal {'published' if request.form.get('published') else 'hidden'}")
    flash("Settings saved.")
    return redirect(url_for("admin_upload_page"))


@app.route("/admin/password", methods=["GET", "POST"])
@admin_required
def admin_password():
    if request.method == "POST":
        current = request.form.get("current", "")
        new = request.form.get("new", "")
        confirm = request.form.get("confirm", "")
        if not _password_ok(current):
            flash("Current password is wrong.", "err")
            return redirect(url_for("admin_password"))
        if len(new) < 6:
            flash("New password must be at least 6 characters.", "err")
            return redirect(url_for("admin_password"))
        if new != confirm:
            flash("New passwords do not match.", "err")
            return redirect(url_for("admin_password"))
        set_setting("admin_password_hash",
                    hashlib.sha256(new.encode("utf-8")).hexdigest())
        _clear_login_fails(_client_ip())
        log_audit("Admin password changed", "")
        flash("Admin password changed.")
        return redirect(url_for("admin_dashboard"))
    return render_template("admin_password.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="password")


def _filtered_students(q, program, degree, status, session_id=None):
    """Students matching the records-page filters, with program + status computed."""
    db = get_db()
    clauses, params = [], []
    if q:
        like = f"%{q}%"
        clauses.append("(name LIKE ? OR cnic LIKE ? OR discipline LIKE ? OR father_name LIKE ?)")
        params += [like] * 4
    if session_id:
        clauses.append("session_id=?")
        params.append(session_id)
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


def _session_arg_id():
    """Resolve the ?session= request arg: 'all' -> None, digits -> that id,
    anything else (incl. missing) -> the active session's id."""
    arg = request.args.get("session", "active")
    if arg == "all":
        return None
    if arg and arg.isdigit():
        return int(arg)
    act = active_session()
    return act["id"] if act else None


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
    act = active_session()
    sess_arg = request.args.get("session", "active")
    if sess_arg == "all":
        sess_id = None
    elif sess_arg.isdigit():
        sess_id = int(sess_arg)
    else:
        sess_id = act["id"] if act else None

    rows = _filtered_students(q, program, degree, status, session_id=sess_id)

    total = len(rows)
    pages = max(1, (total + per_page - 1) // per_page)
    page = min(max(page, 1), pages)
    page_rows = rows[(page - 1) * per_page:page * per_page]
    return render_template("admin_students.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="students", rows=page_rows, page=page, pages=pages,
                           total=total, q=q, program=program, degree=degree,
                           degrees=_degree_options(), status=status,
                           per_page=per_page, sessions=get_sessions(),
                           sess_arg=sess_arg)


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
    rows = _filtered_students(q, program, degree, status,
                              session_id=_session_arg_id())
    log_audit("Records exported", f"{len(rows)} record(s) to Excel")
    return _xlsx_response(rows, "cuvas_students_export.xlsx")


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
    rows = _filtered_students(q, program, degree, status,
                              session_id=_session_arg_id())

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
               "Total", "Obtained", "Passing %", "Status", "Held On"]
    col_w = [10 * mm, 36 * mm, 32 * mm, 30 * mm, 18 * mm, 39 * mm, 15 * mm,
             21 * mm, 17 * mm, 22 * mm, 28 * mm]
    data = [[Paragraph(h, head_style) for h in headers]]
    for i, r in enumerate(rows, start=1):
        held = parse_date(r.get("test_date"))
        opct = obtained_pct(r.get("obtained"), r.get("total"))
        ppct = program_passing_pct(r["program"])
        obt_disp = str(r.get("obtained") or "")
        if opct is not None and obt_disp:
            obt_disp = f"{obt_disp} ({fmt_num(opct)}%)"
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
            Paragraph(obt_disp, cell_style),
            Paragraph(fmt_num(ppct) + "%", cell_style),
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
        flash("Student not found.", "err")
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
            flash("Marks must be numbers.", "err")
            return redirect(url_for("admin_edit_student", cid=cid))
        if not name or len(cnic) != 13:
            flash("Name and a valid 13-digit CNIC are required.", "err")
            return redirect(url_for("admin_edit_student", cid=cid))
        dup = db.fetchone("SELECT id FROM candidates WHERE cnic=? AND id<>?", (cnic, cid))
        if dup:
            flash("Another student already uses this CNIC.", "err")
            return redirect(url_for("admin_edit_student", cid=cid))
        db.execute("UPDATE candidates SET name=?, father_name=?, cnic=?, program=?,"
                   " discipline=?, test_date=?, obtained=?, total=? WHERE id=?",
                   (name, father, cnic, program, discipline, test_date,
                    obtained or None, total or None, cid))
        db.commit()
        log_audit("Record edited", f"{name} ({fmt_cnic(cnic)}) — {obtained}/{total}")
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
    row = db.fetchone("SELECT name, cnic FROM candidates WHERE id=?", (cid,))
    label = f"{row['name']} ({fmt_cnic(row['cnic'])})" if row else f"id {cid}"
    make_backup(f"Before deleting {label}")
    db.execute("DELETE FROM candidates WHERE id=?", (cid,))
    db.commit()
    log_audit("Record deleted", label)
    flash("Student record deleted. A backup was saved under Admin → Backups.", "del")
    return redirect(url_for("admin_students"))


@app.route("/admin/clear", methods=["POST"])
@admin_required
def admin_clear():
    db = get_db()
    n = make_backup("Before deleting ALL records")
    db.execute("DELETE FROM candidates")
    db.commit()
    log_audit("All records deleted", f"{n} record(s) were backed up first")
    flash("All result records deleted. A backup was saved under Admin → Backups.", "del")
    return redirect(url_for("admin_students"))


# ---------------- admin: one-page summary report (PDF) ----------------
def build_summary_pdf(rows, sess_name):
    """One-page summary report PDF as a BytesIO (route + archives)."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.pdfgen import canvas as rl_canvas
    for r in rows:
        r["program"] = r.get("program") or "PhD"
        _, r["status"] = calc_result(r.get("obtained"), r.get("total"),
                                     program_passing_pct(r["program"]))
        r["pct"] = obtained_pct(r.get("obtained"), r.get("total"))
    total = len(rows)
    qual = sum(1 for r in rows if r["status"] == "Qualified")
    pass_rate = round(qual / total * 100, 1) if total else 0.0
    progs = []
    for prog in ("PhD", "MPhil"):
        pr = [r for r in rows if r["program"] == prog]
        pq = sum(1 for r in pr if r["status"] == "Qualified")
        progs.append({"program": prog, "total": len(pr), "qualified": pq,
                      "rate": round(pq / len(pr) * 100, 1) if pr else 0.0})
    degmap = {}
    for r in rows:
        key = ((r.get("discipline") or "—").strip() or "—", r["program"])
        d = degmap.setdefault(key, {"degree": key[0], "program": key[1],
                                    "total": 0, "qualified": 0})
        d["total"] += 1
        if r["status"] == "Qualified":
            d["qualified"] += 1
    degs = sorted(degmap.values(), key=lambda d: (-d["total"], d["degree"]))
    CAP = 16
    if len(degs) > CAP:
        rest = degs[CAP - 1:]
        degs = degs[:CAP - 1] + [{"degree": f"All other degrees ({len(rest)})",
                                  "program": "—",
                                  "total": sum(d["total"] for d in rest),
                                  "qualified": sum(d["qualified"] for d in rest)}]
    for d in degs:
        d["rate"] = round(d["qualified"] / d["total"] * 100, 1) if d["total"] else 0.0
    scored = [r for r in rows if r["pct"] is not None]
    top5 = sorted(scored, key=lambda r: (-r["pct"], -(r.get("obtained") or 0)))[:5]
    avg_pct = round(sum(r["pct"] for r in scored) / len(scored), 1) if scored else 0.0
    dates = sorted({fmt_date(parse_date(r.get("test_date") or ""))
                    for r in rows if parse_date(r.get("test_date") or "")})
    held = ", ".join(d for d in dates if d) or "—"

    INDIGO = colors.HexColor("#2d3290")
    INK = colors.HexColor("#1a1f36")
    MUTED = colors.HexColor("#667085")
    LINE = colors.HexColor("#d7dbea")
    OKC = colors.HexColor("#12805c")
    BADC = colors.HexColor("#b3261e")
    LIGHT = colors.HexColor("#eef0fb")

    buf = io.BytesIO()
    W, H = A4
    c = rl_canvas.Canvas(buf, pagesize=A4)
    M = 42            # margin
    CW = W - 2 * M    # content width

    # header band
    c.setFillColor(INDIGO)
    c.rect(0, H - 66, W, 66, stroke=0, fill=1)
    logo_path = os.path.join(BASE_DIR, "static", "img", "cuvas-logo.png")
    if os.path.exists(logo_path):
        try:
            c.setFillColor(colors.white)
            c.roundRect(M, H - 56, 46, 46, 6, stroke=0, fill=1)
            c.drawImage(logo_path, M + 4, H - 52, 38, 38,
                        preserveAspectRatio=True, mask="auto")
        except Exception:
            pass
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 12.5)
    c.drawString(M + 58, H - 34, "Cholistan University of Veterinary & Animal Sciences, Bahawalpur")
    c.setFont("Helvetica", 9.5)
    c.drawString(M + 58, H - 50, "Directorate of Advanced Studies & Research  •  Entry Test Result Portal")
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(W - M, H - 34, "SUMMARY REPORT")
    c.setFont("Helvetica", 8.5)
    c.drawRightString(W - M, H - 50, "For Official Use")

    y = H - 66 - 30
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 15.5)
    c.drawString(M, y, "Entry Test Result — Summary Report")
    y -= 16
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9.5)
    c.drawString(M, y, f"Session: {sess_name}    •    Test held: {held}    •    "
                 f"Passing criteria: PhD {fmt_num(program_passing_pct('PhD'))}%  •  "
                 f"MPhil {fmt_num(program_passing_pct('MPhil'))}%")
    y -= 10
    c.setStrokeColor(LINE); c.setLineWidth(1)
    c.line(M, y, W - M, y)

    # KPI boxes
    y -= 62
    kpis = [("Total Candidates", str(total), INK),
            ("Qualified", str(qual), OKC),
            ("Not Qualified", str(total - qual), BADC),
            ("Overall Pass Rate", f"{fmt_num(pass_rate)}%", INDIGO)]
    bw = (CW - 3 * 10) / 4
    for i, (label, val, col) in enumerate(kpis):
        x = M + i * (bw + 10)
        c.setFillColor(LIGHT)
        c.roundRect(x, y, bw, 48, 8, stroke=0, fill=1)
        c.setFillColor(col)
        c.setFont("Helvetica-Bold", 17)
        c.drawCentredString(x + bw / 2, y + 26, val)
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8)
        c.drawCentredString(x + bw / 2, y + 10, label.upper())

    # program split (left) + top scorers (right)
    y -= 28
    colw = (CW - 18) / 2
    c.setFillColor(INK); c.setFont("Helvetica-Bold", 11)
    c.drawString(M, y, "Program-wise Result")
    c.drawString(M + colw + 18, y, "Top Scorers")
    y -= 8
    # program table
    ty = y
    headers = ["Program", "Candidates", "Qualified", "Pass %"]
    widths = [colw * .34, colw * .24, colw * .22, colw * .20]
    c.setFillColor(INDIGO)
    c.roundRect(M, ty - 16, colw, 16, 3, stroke=0, fill=1)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 8.5)
    x = M + 6
    for htxt, wdt in zip(headers, widths):
        c.drawString(x, ty - 11.5, htxt); x += wdt
    ty -= 16
    body = progs + [{"program": "Overall", "total": total, "qualified": qual,
                     "rate": pass_rate}]
    for i, p in enumerate(body):
        ty -= 15
        if i % 2 == 0:
            c.setFillColor(colors.HexColor("#f6f7fc"))
            c.rect(M, ty - 3.5, colw, 15, stroke=0, fill=1)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold" if p["program"] == "Overall" else "Helvetica", 9)
        vals = [p["program"], str(p["total"]), str(p["qualified"]),
                f"{fmt_num(p['rate'])}%"]
        x = M + 6
        for v, wdt in zip(vals, widths):
            c.drawString(x, ty, v); x += wdt
    # top scorers list
    ty2 = y
    c.setFillColor(INDIGO)
    c.roundRect(M + colw + 18, ty2 - 16, colw, 16, 3, stroke=0, fill=1)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 8.5)
    c.drawString(M + colw + 24, ty2 - 11.5, "Candidate — Degree (Program)")
    c.drawRightString(M + colw + 18 + colw - 6, ty2 - 11.5, "Score")
    ty2 -= 16
    if not top5:
        c.setFillColor(MUTED); c.setFont("Helvetica", 9)
        c.drawString(M + colw + 24, ty2 - 11, "No scored records yet.")
    for i, r in enumerate(top5):
        ty2 -= 15
        if i % 2 == 0:
            c.setFillColor(colors.HexColor("#f6f7fc"))
            c.rect(M + colw + 18, ty2 - 3.5, colw, 15, stroke=0, fill=1)
        c.setFillColor(INK); c.setFont("Helvetica", 8.8)
        line = f"{i + 1}. {r.get('name') or ''} — {r.get('discipline') or ''} ({r['program']})"
        while c.stringWidth(line, "Helvetica", 8.8) > colw - 70 and len(line) > 10:
            line = line[:-2]
        c.drawString(M + colw + 24, ty2, line)
        c.setFont("Helvetica-Bold", 8.8)
        c.drawRightString(M + colw + 18 + colw - 6, ty2,
                          f"{fmt_num(r.get('obtained'))}/{fmt_num(r.get('total'))} ({fmt_num(r['pct'])}%)")
    y = min(ty, ty2) - 16
    c.setFillColor(MUTED); c.setFont("Helvetica", 9)
    c.drawString(M, y, f"Marks snapshot:  Highest {fmt_num(max((r['pct'] for r in scored), default=0))}%"
                 f"   •   Average {fmt_num(avg_pct)}%"
                 f"   •   Lowest {fmt_num(min((r['pct'] for r in scored), default=0))}%")

    # degree-wise table
    y -= 26
    c.setFillColor(INK); c.setFont("Helvetica-Bold", 11)
    c.drawString(M, y, "Degree-wise Result")
    y -= 8
    dwidths = [CW * .46, CW * .14, CW * .15, CW * .13, CW * .12]
    dheads = ["Degree / Discipline", "Program", "Candidates", "Qualified", "Pass %"]
    c.setFillColor(INDIGO)
    c.roundRect(M, y - 17, CW, 17, 3, stroke=0, fill=1)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 8.5)
    x = M + 6
    for htxt, wdt in zip(dheads, dwidths):
        c.drawString(x, y - 12, htxt); x += wdt
    y -= 17
    for i, d in enumerate(degs):
        y -= 15.5
        if i % 2 == 0:
            c.setFillColor(colors.HexColor("#f6f7fc"))
            c.rect(M, y - 3.5, CW, 15.5, stroke=0, fill=1)
        c.setFillColor(INK); c.setFont("Helvetica", 9)
        vals = [d["degree"], d["program"], str(d["total"]), str(d["qualified"]),
                f"{fmt_num(d['rate'])}%"]
        x = M + 6
        for v, wdt in zip(vals, dwidths):
            vv = v
            while c.stringWidth(vv, "Helvetica", 9) > wdt - 10 and len(vv) > 8:
                vv = vv[:-2]
            c.drawString(x, y, vv); x += wdt
    if not degs:
        y -= 15
        c.setFillColor(MUTED); c.setFont("Helvetica", 9)
        c.drawString(M + 6, y, "No records in this session yet.")

    # footer: signatures + generated line (fixed at page bottom)
    fy = 74
    c.setStrokeColor(LINE)
    c.line(M, fy + 26, M + 170, fy + 26)
    c.line(W - M - 190, fy + 26, W - M, fy + 26)
    c.setFillColor(INK); c.setFont("Helvetica", 9)
    c.drawString(M, fy + 12, "Prepared by (DAS Office)")
    c.drawRightString(W - M, fy + 12, "Director, Advanced Studies & Research")
    c.setFillColor(MUTED); c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, 40, f"Generated from the CUVAS Result Portal on "
                            f"{datetime.now().strftime('%d %b %Y, %I:%M %p')} — for official record only.")
    c.showPage()
    c.save()
    buf.seek(0)
    return buf


@app.route("/admin/summary-pdf")
@admin_required
def admin_summary_pdf():
    """One-page, print-ready summary of a whole entry test for the office file:
    KPIs, program split, degree-wise table, marks snapshot and top scorers."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.pdfgen import canvas as rl_canvas

    sid = _session_arg_id()
    act = active_session()
    sess_row = None
    for s in get_sessions():
        if s["id"] == sid:
            sess_row = s
    sess_name = sess_row["name"] if sess_row else (act["name"] if act else "")
    db = get_db()
    if sid:
        rows = [dict(r) for r in db.fetchall(
            "SELECT * FROM candidates WHERE session_id=?", (sid,))]
    else:
        rows = [dict(r) for r in db.fetchall("SELECT * FROM candidates")]
    buf = build_summary_pdf(rows, sess_name)
    log_audit("Summary report downloaded", f"session '{sess_name}' — {len(rows)} record(s)")
    fname = "CUVAS_Summary_Report_" + re.sub(r"[^A-Za-z0-9]+", "_", sess_name).strip("_") + ".pdf"
    return send_file(buf, mimetype="application/pdf", as_attachment=True,
                     download_name=fname)


# ---------------- admin: test sessions ----------------
@app.route("/admin/sessions")
@admin_required
def admin_sessions():
    db = get_db()
    sess = []
    for s in get_sessions():
        row = db.fetchone("SELECT COUNT(*) AS c FROM candidates WHERE session_id=?",
                          (s["id"],))
        s["count"] = row["c"] if isinstance(row, dict) else row[0]
        sess.append(s)
    return render_template("admin_sessions.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="sessions", sessions=sess)


@app.route("/admin/sessions/<int:sid>/activate", methods=["POST"])
@admin_required
def admin_sessions_activate(sid):
    db = get_db()
    row = db.fetchone("SELECT * FROM sessions WHERE id=?", (sid,))
    if not row:
        flash("Session not found.", "err")
        return redirect(url_for("admin_sessions"))
    db.execute("UPDATE sessions SET is_active=0")
    db.execute("UPDATE sessions SET is_active=1 WHERE id=?", (sid,))
    db.commit()
    log_audit("Session activated", dict(row)["name"])
    flash(f"Session '{dict(row)['name']}' is now the active session.")
    return redirect(url_for("admin_sessions"))


@app.route("/admin/sessions/<int:sid>/delete", methods=["POST"])
@admin_required
def admin_sessions_delete(sid):
    db = get_db()
    row = db.fetchone("SELECT COUNT(*) AS c FROM candidates WHERE session_id=?", (sid,))
    n = row["c"] if isinstance(row, dict) else row[0]
    if n:
        flash("Only empty sessions can be deleted — this one still has records.", "err")
        return redirect(url_for("admin_sessions"))
    srow = db.fetchone("SELECT * FROM sessions WHERE id=?", (sid,))
    db.execute("DELETE FROM sessions WHERE id=?", (sid,))
    db.commit()
    if srow:
        log_audit("Session deleted", dict(srow)["name"])
    flash("Session deleted.", "del")
    return redirect(url_for("admin_sessions"))


@app.route("/admin/sessions/<int:sid>/archive")
@admin_required
def admin_sessions_archive(sid):
    """Download the WHOLE session as one ZIP for permanent record keeping:
    students.xlsx, merit-list.pdf, summary-report.pdf and one PDF result
    card per candidate, plus a README."""
    import zipfile
    db = get_db()
    srow = db.fetchone("SELECT * FROM sessions WHERE id=?", (sid,))
    if not srow:
        flash("Session not found.", "err")
        return redirect(url_for("admin_sessions"))
    sess_name = dict(srow)["name"]
    rows = [dict(r) for r in db.fetchall(
        "SELECT * FROM candidates WHERE session_id=? ORDER BY name", (sid,))]
    xrows = _filtered_students("", "all", "all", "all", session_id=sid)
    mrows = _merit_rows("all", "all", session_id=sid)
    zipbuf = io.BytesIO()
    with zipfile.ZipFile(zipbuf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("students.xlsx", _xlsx_buf(xrows).getvalue())
        z.writestr("merit-list.pdf",
                   build_merit_pdf(mrows, "all", "all", sess_name).getvalue())
        z.writestr("summary-report.pdf",
                   build_summary_pdf([dict(r) for r in rows], sess_name).getvalue())
        for r in rows:
            card = build_card(dict(r))
            safe = re.sub(r"[^A-Za-z0-9]+", "_", r.get("name") or "student").strip("_")
            z.writestr(f"result-cards/{r.get('cnic')}_{safe}.pdf",
                       build_pdf(card).getvalue())
        z.writestr("README.txt",
                   "CUVAS Entry Test — Session Archive\n"
                   f"Session: {sess_name}\n"
                   f"Generated: {datetime.now().strftime('%d %b %Y, %I:%M %p')}\n"
                   f"Records: {len(rows)}\n\n"
                   "Contents:\n"
                   "- students.xlsx — all records of this session\n"
                   "- merit-list.pdf — merit list of this session\n"
                   "- summary-report.pdf — one-page summary report\n"
                   "- result-cards/ — one PDF result card per candidate\n")
    zipbuf.seek(0)
    log_audit("Session archive downloaded", f"'{sess_name}' — {len(rows)} record(s)")
    fname = ("CUVAS_Session_Archive_"
             + re.sub(r"[^A-Za-z0-9]+", "_", sess_name).strip("_") + ".zip")
    return send_file(zipbuf, mimetype="application/zip", as_attachment=True,
                     download_name=fname)


# ---------------- admin: global quick search ----------------
@app.route("/admin/quicksearch")
@admin_required
def admin_quicksearch():
    q = (request.args.get("q") or "").strip()
    if not q:
        return redirect(url_for("admin_students"))
    digits = norm_cnic(q)
    db = get_db()
    row = None
    if len(digits) == 13:
        row = db.fetchone("SELECT id FROM candidates WHERE cnic=?"
                          " ORDER BY session_id DESC", (digits,))
    if not row:
        rows = db.fetchall("SELECT id FROM candidates WHERE name LIKE ?",
                           (f"%{q}%",))
        if len(rows) == 1:
            row = rows[0]
    if row:
        rid = row["id"] if isinstance(row, dict) else row[0]
        return redirect(url_for("admin_edit_student", cid=rid))
    return redirect(url_for("admin_students", q=q, session="all"))


# ---------------- admin: activity log ----------------
@app.route("/admin/activity")
@admin_required
def admin_activity():
    rows = [dict(r) for r in get_db().fetchall(
        "SELECT * FROM audit_log ORDER BY id DESC LIMIT 200")]
    return render_template("admin_activity.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="activity", rows=rows)


# ---------------- admin: backups ----------------
BACKUP_COLS = ("name", "father_name", "cnic", "program", "roll_no", "discipline",
               "test_date", "obtained", "total", "merit_no", "uploaded_at", "session_id")


@app.route("/admin/backups")
@admin_required
def admin_backups():
    rows = [dict(r) for r in get_db().fetchall(
        "SELECT id, ts, reason, count FROM backups ORDER BY id DESC")]
    return render_template("admin_backups.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="backups", rows=rows)


@app.route("/admin/backups/<int:bid>/download")
@admin_required
def admin_backup_download(bid):
    row = get_db().fetchone("SELECT * FROM backups WHERE id=?", (bid,))
    if not row:
        flash("Backup not found.", "err")
        return redirect(url_for("admin_backups"))
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Name", "Father's Name", "CNIC", "Program", "Roll No",
                "Applied Degree", "Held On", "Total Marks", "Obtained Marks",
                "Merit No", "Uploaded At"])
    for rec in json.loads(dict(row)["data"] or "[]"):
        w.writerow([rec.get(c, "") for c in BACKUP_COLS])
    out = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(out, as_attachment=True,
                     download_name=f"backup_{bid}_{dict(row)['ts'][:10]}.csv",
                     mimetype="text/csv")


@app.route("/admin/backups/<int:bid>/restore", methods=["POST"])
@admin_required
def admin_backup_restore(bid):
    db = get_db()
    row = db.fetchone("SELECT * FROM backups WHERE id=?", (bid,))
    if not row:
        flash("Backup not found.", "err")
        return redirect(url_for("admin_backups"))
    make_backup(f"Before restoring backup #{bid}")
    recs = json.loads(dict(row)["data"] or "[]")
    db.execute("DELETE FROM candidates")
    act = active_session()
    act_id = act["id"] if act else None
    placeholders = ", ".join(["?"] * len(BACKUP_COLS))
    for rec in recs:
        db.execute(f"INSERT INTO candidates ({', '.join(BACKUP_COLS)}) "
                   f"VALUES ({placeholders})",
                   tuple((rec.get(c, "") or None) if c != "session_id"
                         else (rec.get("session_id") or act_id)
                         for c in BACKUP_COLS))
    db.commit()
    log_audit("Backup restored", f"backup #{bid} — {len(recs)} record(s)")
    flash(f"Backup #{bid} restored ({len(recs)} record(s)).")
    return redirect(url_for("admin_students"))


# ---------------- admin: lookup log ----------------
@app.route("/admin/lookups")
@admin_required
def admin_lookups():
    db = get_db()
    try:
        total = db.fetchone("SELECT COUNT(*) AS c FROM lookups")
        total = total["c"] if isinstance(total, dict) else total[0]
        fails = db.fetchone("SELECT COUNT(*) AS c FROM lookups WHERE found=0")
        fails = fails["c"] if isinstance(fails, dict) else fails[0]
        recent = [dict(r) for r in db.fetchall(
            "SELECT * FROM lookups ORDER BY id DESC LIMIT 60")]
        top = [dict(r) for r in db.fetchall(
            "SELECT cnic, COUNT(*) AS tries, MAX(ts) AS last FROM lookups "
            "WHERE found=0 GROUP BY cnic ORDER BY tries DESC LIMIT 20")]
    except Exception:
        total, fails, recent, top = 0, 0, [], []
    return render_template("admin_lookups.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="lookups", total=total, fails=fails,
                           recent=recent, top=top)


# ---------------- admin: merit list ----------------
def _merit_rows(program="all", degree="all", session_id=None):
    rows = _filtered_students("", program, degree, "all", session_id=session_id)
    for r in rows:
        r["pct"] = obtained_pct(r.get("obtained"), r.get("total"))
    rows.sort(key=lambda r: (-(r["pct"] if r["pct"] is not None else -1),
                             r.get("name") or ""))
    for i, r in enumerate(rows, start=1):
        r["rank"] = i
    return rows


@app.route("/admin/merit")
@admin_required
def admin_merit():
    program = request.args.get("program", "all")
    degree = request.args.get("degree", "all")
    rows = _merit_rows(program, degree, session_id=_session_arg_id())
    return render_template("admin_merit.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="merit", rows=rows, program=program,
                           degree=degree, degrees=_degree_options(),
                           sessions=get_sessions(),
                           sess_arg=request.args.get("session", "active"))


def build_merit_pdf(rows, program="all", degree="all", sess_name=""):
    """Merit list PDF as a BytesIO (shared by the route + session archives)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, Image as RLImage)

    INDIGO = colors.HexColor("#2d3290")
    DARK = colors.HexColor("#1a1f36")
    MUTED = colors.HexColor("#667085")

    buf = io.BytesIO()
    page = landscape(A4)
    doc = SimpleDocTemplate(buf, pagesize=page, leftMargin=14 * mm,
                            rightMargin=14 * mm, topMargin=14 * mm,
                            bottomMargin=14 * mm, title="CUVAS Merit List")
    avail = page[0] - 28 * mm
    title_style = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=15,
                                 textColor=DARK, alignment=TA_LEFT, leading=18)
    sub_style = ParagraphStyle("s", fontName="Helvetica", fontSize=10,
                               textColor=MUTED, alignment=TA_LEFT, leading=13)
    cell_style = ParagraphStyle("c", fontName="Helvetica", fontSize=8.5,
                                textColor=DARK, leading=11)
    head_style = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=9,
                                textColor=colors.white, leading=11)
    story = []
    logo_path = os.path.join(BASE_DIR, "static", "img", "cuvas-logo.png")
    if os.path.exists(logo_path):
        logo = RLImage(logo_path, width=30 * mm, height=30 * mm * 148 / 270)
        head_tbl = Table([[logo, Paragraph(
            "Cholistan University of Veterinary &amp; Animal Sciences, Bahawalpur<br/>"
            "<font size=11 color=\"#2d3290\"><b>Entry Test — Merit List</b></font>",
            title_style)]], colWidths=[32 * mm, avail - 32 * mm])
        head_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                      ("LEFTPADDING", (1, 0), (1, 0), 6)]))
        story.append(head_tbl)
    else:
        story.append(Paragraph(
            "Cholistan University of Veterinary &amp; Animal Sciences, Bahawalpur",
            title_style))
        story.append(Paragraph("Entry Test — Merit List", sub_style))
    bits = []
    if sess_name:
        bits.append(f"Session: <b>{sess_name}</b>")
    if program != "all":
        bits.append(f"Program: <b>{program}</b>")
    if degree != "all":
        bits.append(f"Degree: <b>{degree}</b>")
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(
        ((" &nbsp;|&nbsp; ".join(bits) + " &nbsp;&nbsp;•&nbsp;&nbsp;") if bits else "")
        + f"Total: <b>{len(rows)}</b> candidate(s) &nbsp;&nbsp;•&nbsp;&nbsp; "
        f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}", sub_style))
    story.append(Spacer(1, 4 * mm))
    headers = ["Rank", "Merit No", "Name", "Father's Name", "CNIC", "Program",
               "Applied Degree", "Total", "Obtained", "%", "Status"]
    col_w = [13 * mm, 18 * mm, 38 * mm, 34 * mm, 30 * mm, 18 * mm, 40 * mm,
             14 * mm, 17 * mm, 14 * mm, 24 * mm]
    data = [[Paragraph(h, head_style) for h in headers]]
    for r in rows:
        st = r.get("status", "")
        st_color = "#12805c" if st == "Qualified" else "#b3261e"
        data.append([
            Paragraph(str(r["rank"]), cell_style),
            Paragraph(str(r.get("merit_no") or ""), cell_style),
            Paragraph(str(r.get("name") or ""), cell_style),
            Paragraph(str(r.get("father_name") or ""), cell_style),
            Paragraph(str(r.get("cnic") or ""), cell_style),
            Paragraph(str(r.get("program") or ""), cell_style),
            Paragraph(str(r.get("discipline") or ""), cell_style),
            Paragraph(str(r.get("total") or ""), cell_style),
            Paragraph(str(r.get("obtained") or ""), cell_style),
            Paragraph((fmt_num(r["pct"]) + "%") if r["pct"] is not None else "",
                       cell_style),
            Paragraph(f'<font color="{st_color}"><b>{st}</b></font>', cell_style),
        ])
    tbl = Table(data, colWidths=col_w, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), INDIGO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d0d5e8")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f6fc")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
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
    return buf


@app.route("/admin/merit-pdf")
@admin_required
def admin_merit_pdf():
    """Printable merit list (sorted by obtained percentage), PDF."""
    program = request.args.get("program", "all")
    degree = request.args.get("degree", "all")
    rows = _merit_rows(program, degree, session_id=_session_arg_id())
    buf = build_merit_pdf(rows, program, degree)
    return send_file(buf, as_attachment=True,
                     download_name="cuvas_merit_list.pdf",
                     mimetype="application/pdf")


# ---------------- admin: bulk actions on student records ----------------
def _xlsx_buf(rows):
    """Students workbook as a BytesIO (shared by the export route + archives)."""
    return _xlsx_response(rows, "export.xlsx", _buf_only=True)


def _xlsx_response(rows, filename, _buf_only=False):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "Students"
    headers = ["Name", "Father's Name", "CNIC", "Program", "Applied Degree",
               "Held On", "Total Marks", "Obtained Marks", "Obtained %",
               "Passing %", "Result Status", "Valid Till"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2D3290")
        cell.alignment = Alignment(horizontal="center")
    for r in rows:
        held = parse_date(r.get("test_date"))
        valid = add_one_year(held) if held else None
        opct = obtained_pct(r.get("obtained"), r.get("total"))
        ppct = program_passing_pct(r["program"])
        ws.append([r.get("name"), r.get("father_name"), r.get("cnic"),
                   r.get("program"), r.get("discipline"),
                   fmt_date(held) if held else (r.get("test_date") or ""),
                   r.get("total"), r.get("obtained"),
                   (fmt_num(opct) + "%") if opct is not None else "",
                   fmt_num(ppct) + "%", r.get("status"),
                   fmt_date(valid) if valid else ""])
    widths = [22, 22, 20, 12, 24, 16, 14, 16, 12, 12, 16, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    if _buf_only:
        return buf
    return send_file(
        buf, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/admin/students/bulk", methods=["POST"])
@admin_required
def admin_students_bulk():
    ids = request.form.getlist("ids")
    ids = [int(i) for i in ids if str(i).isdigit()]
    action = request.form.get("action", "")
    if not ids:
        flash("No records selected.", "err")
        return redirect(url_for("admin_students"))
    db = get_db()
    rows = [dict(r) for r in db.fetchall(
        f"SELECT * FROM candidates WHERE id IN ({','.join(['?'] * len(ids))})", ids)]
    for r in rows:
        r["program"] = r.get("program") or "PhD"
        _, r["status"] = calc_result(r.get("obtained"), r.get("total"),
                                     program_passing_pct(r["program"]))
    if action == "export":
        log_audit("Bulk export", f"{len(rows)} record(s)")
        return _xlsx_response(rows, "cuvas_selected_students.xlsx")
    if action == "delete":
        make_backup(f"Before bulk delete ({len(rows)} record(s))")
        db.execute(f"DELETE FROM candidates WHERE id IN ({','.join(['?'] * len(ids))})",
                   ids)
        db.commit()
        log_audit("Bulk delete", f"{len(rows)} record(s)")
        flash(f"{len(rows)} record(s) deleted. A backup was saved under Admin → Backups.", "del")
        return redirect(url_for("admin_students"))
    flash("Unknown bulk action.", "err")
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
        flash("Please choose a file to upload.", "err")
        return redirect(url_for("admin_upload_page"))
    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in ALLOWED_EXT:
        flash("Only CSV or Excel (.xlsx/.xls) files are allowed.", "err")
        return redirect(url_for("admin_upload_page"))

    try:
        raw_rows = read_upload(file)
    except Exception as e:
        flash(f"Could not read file: {e}", "err")
        return redirect(url_for("admin_upload_page"))
    if not raw_rows:
        flash("No data rows found in the file.", "err")
        return redirect(url_for("admin_upload_page"))

    first = raw_rows[0]
    mapping = {}
    for hdr in first.keys():
        k = header_key(hdr)
        if k and k not in mapping:
            mapping[hdr] = k

    if "name" not in mapping.values() or "cnic" not in mapping.values():
        flash("File must contain at least a Name column and a CNIC column.", "err")
        return redirect(url_for("admin_upload_page"))

    # Stage every row first — nothing is saved until the admin reviews the
    # preview page and presses Confirm.
    db = get_db()
    sess_name = (request.form.get("session_name") or "").strip()
    act = active_session()
    if not sess_name:
        sess_name = act["name"] if act else "Entry Test"
    # CNIC existence is judged inside the target session only — the same CNIC
    # in an older session is a legitimately different record.
    target = db.fetchone("SELECT id FROM sessions WHERE name=?", (sess_name,))
    existing = set()
    if target:
        tid = target["id"] if isinstance(target, dict) else target[0]
        existing = {r["cnic"] for r in db.fetchall(
            "SELECT cnic FROM candidates WHERE session_id=?", (tid,))}
    # name -> CNICs map (existing records + this file) for duplicate-name warnings
    name_map = {}
    for r in db.fetchall("SELECT name, cnic FROM candidates"):
        key = re.sub(r"\s+", " ", (r["name"] or "").strip().lower())
        if key:
            name_map.setdefault(key, set()).add(r["cnic"])
    batch = datetime.now().strftime("%Y%m%d%H%M%S") + "-" + os.urandom(3).hex()
    db.execute("DELETE FROM upload_staging")  # one pending preview at a time
    staged = 0
    for row_no, rr in enumerate(raw_rows, start=2):  # row 1 = header in the file
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
        rec["cnic"] = cnic
        if not rec.get("name"):
            status, reason = "invalid", "Missing name"
        elif len(cnic) != 13:
            status, reason = "invalid", "CNIC is not 13 digits"
        elif cnic in existing:
            status, reason = "update", "CNIC already exists in this session — record will be updated"
        else:
            status, reason = "new", ""
            existing.add(cnic)
        # ---- non-blocking warnings (row still saves; admin just sees them) ----
        warns = []
        if status != "invalid":
            try:
                obt_f = float(rec.get("obtained") or "nan")
            except ValueError:
                obt_f = float("nan")
            try:
                tot_f = float(rec.get("total") or "nan")
            except ValueError:
                tot_f = float("nan")
            if obt_f == obt_f and tot_f == tot_f:  # both are real numbers
                if obt_f > tot_f:
                    warns.append("Obtained marks are MORE than total marks")
                elif tot_f > 0 and obt_f == tot_f:
                    warns.append("Full marks (100%) — please double-check")
                elif obt_f == 0:
                    warns.append("Zero marks — please double-check")
                elif obt_f < 0:
                    warns.append("Negative marks")
            if not (rec.get("father_name") or "").strip():
                warns.append("Father's name is empty")
            key = re.sub(r"\s+", " ", (rec.get("name") or "").strip().lower())
            if key and key in name_map and cnic not in name_map[key]:
                warns.append("Same name already exists with a different CNIC")
            if key:
                name_map.setdefault(key, set()).add(cnic)
        db.execute("INSERT INTO upload_staging(batch, row_no, status, reason, data, warning)"
                   " VALUES (?, ?, ?, ?, ?, ?)",
                   (batch, row_no, status, reason, json.dumps(rec), "; ".join(warns)))
        staged += 1
    db.commit()
    session["upload_batch"] = batch
    session["upload_file"] = file.filename
    session["upload_session"] = sess_name
    log_audit("Upload staged for preview", f"{file.filename} — {staged} row(s), session '{sess_name}'")
    return redirect(url_for("admin_upload_preview"))


@app.route("/admin/upload/preview")
@admin_required
def admin_upload_preview():
    batch = session.get("upload_batch")
    if not batch:
        flash("No upload is waiting for review.", "err")
        return redirect(url_for("admin_upload_page"))
    db = get_db()
    rows = []
    for s in db.fetchall("SELECT * FROM upload_staging WHERE batch=? ORDER BY row_no",
                         (batch,)):
        d = dict(s)
        d["rec"] = json.loads(d["data"] or "{}")
        rows.append(d)
    counts = {"new_count": 0, "update_count": 0, "invalid_count": 0}
    warn_count = 0
    for d in rows:
        counts[d["status"] + "_count"] = counts.get(d["status"] + "_count", 0) + 1
        if d.get("warning"):
            warn_count += 1
    return render_template("admin_preview.html", uni=UNI_NAME, title=PORTAL_TITLE,
                           active="upload", rows=rows, counts=counts,
                           warn_count=warn_count,
                           sess_name=session.get("upload_session", ""),
                           filename=session.get("upload_file", ""))


@app.route("/admin/upload/confirm", methods=["POST"])
@admin_required
def admin_upload_confirm():
    batch = session.get("upload_batch")
    if not batch:
        flash("No upload is waiting for review.", "err")
        return redirect(url_for("admin_upload_page"))
    db = get_db()
    staged = db.fetchall("SELECT * FROM upload_staging WHERE batch=? ORDER BY row_no",
                         (batch,))
    filename = session.get("upload_file", "upload")
    make_backup(f"Before upload: {filename}")
    sess, sess_created = get_or_create_session(session.get("upload_session", ""))
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    cols = ("name", "father_name", "cnic", "program", "roll_no", "discipline",
            "test_date", "obtained", "total", "merit_no", "uploaded_at", "session_id")
    placeholders = ", ".join(["?"] * len(cols))
    if USE_PG:
        upsert = (f"INSERT INTO candidates ({', '.join(cols)}) VALUES ({placeholders}) "
                  "ON CONFLICT(cnic, session_id) DO UPDATE SET "
                  + ", ".join(f"{c}=EXCLUDED.{c}" for c in cols
                              if c not in ("cnic", "session_id")))
    else:
        upsert = f"INSERT OR REPLACE INTO candidates ({', '.join(cols)}) VALUES ({placeholders})"
    imported = updated = skipped = 0
    for s in staged:
        d = dict(s)
        if d["status"] == "invalid":
            skipped += 1
            continue
        rec = json.loads(d["data"] or "{}")
        values = (rec.get("name"), rec.get("father_name"), rec.get("cnic"),
                  rec.get("program") or "PhD", rec.get("roll_no"),
                  rec.get("discipline"), rec.get("test_date"),
                  rec.get("obtained") or None, rec.get("total") or None,
                  rec.get("merit_no"), now, sess["id"])
        try:
            db.execute(upsert, values)
            if d["status"] == "update":
                updated += 1
            else:
                imported += 1
        except Exception:
            skipped += 1
    db.execute("INSERT INTO uploads(ts, filename, imported, updated, skipped)"
               " VALUES (?, ?, ?, ?, ?)", (_now_str(), filename, imported, updated, skipped))
    db.execute("DELETE FROM upload_staging WHERE batch=?", (batch,))
    db.commit()
    session.pop("upload_batch", None)
    session.pop("upload_file", None)
    session.pop("upload_session", None)
    log_audit("Upload applied", f"{filename} — {imported} new, {updated} updated, "
              f"{skipped} skipped, session '{sess['name']}'"
              f"{' (new session, now active)' if sess_created else ''} (backup saved first)")
    flash(f"Upload complete: {imported} new record(s), {updated} updated"
          + (f", {skipped} invalid row(s) skipped." if skipped else ".")
          + f" Session: {sess['name']}")
    return redirect(url_for("admin_upload_page"))


@app.route("/admin/upload/cancel", methods=["POST"])
@admin_required
def admin_upload_cancel():
    batch = session.get("upload_batch")
    if batch:
        db = get_db()
        db.execute("DELETE FROM upload_staging WHERE batch=?", (batch,))
        db.commit()
    session.pop("upload_batch", None)
    session.pop("upload_file", None)
    session.pop("upload_session", None)
    flash("Upload cancelled — nothing was saved.", "err")
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
