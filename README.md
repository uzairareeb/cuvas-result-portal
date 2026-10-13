# CUVAS PhD Entry Test Result Portal

Admin uploads the entry-test result list (CSV/Excel). Students enter their CNIC,
see their Entry Test Result Card, and download it as a PDF.

## Deploy on Render (recommended, free)

Netlify cannot host this app — it only runs static sites and JavaScript/TypeScript/Go
functions, not Python. Render runs the Flask app as-is:

1. Push this folder to a GitHub repo.
2. On [render.com](https://render.com) → New → **Blueprint** → select the repo
   (`render.yaml` is included, so the service is pre-configured).
3. Add env vars: `DATABASE_URL` (free Postgres from [Neon](https://neon.tech)),
   `ADMIN_PASSWORD`, and you're live. Tables are created automatically on first visit.

## Deploy on Vercel (also works)

Vercel's servers are stateless, so the portal uses **Postgres** there instead of
the local SQLite file. The free Neon Postgres one-click integration is the easiest:

1. Push this folder to a GitHub repo (see commands below).
2. On [vercel.com](https://vercel.com) → Add New → Project → Import the repo.
3. Storage tab → Connect → **Neon Postgres** (free) → Create. This sets `DATABASE_URL` for you.
4. Environment Variables → add:
   - `ADMIN_PASSWORD` = your chosen admin password
   - `PORTAL_SECRET` = any long random string
5. Deploy. Tables are created automatically on first visit.

Local runs keep using SQLite (`results.db`) — no `DATABASE_URL` needed.

## Push to GitHub

```bash
cd cuvas_result_portal
git init
git add .
git commit -m "CUVAS PhD entry test result portal"
# create an empty repo on github.com, then:
git remote add origin https://github.com/YOUR-USERNAME/cuvas-result-portal.git
git branch -M main
git push -u origin main
```

## How to run (on any computer with Python 3.10+)

```bash
cd cuvas_result_portal
pip install -r requirements.txt
python3 seed.py      # loads 10 sample candidates (optional, for testing)
python3 app.py
```

Then open:

- Student portal: http://localhost:5000
- Admin panel: http://localhost:5000/admin  (default password: `cuvas123`)

Change the admin password by setting the `ADMIN_PASSWORD` environment variable,
e.g. on Windows: `set ADMIN_PASSWORD=mynewpass` before running `python app.py`.

## Upload file format

CSV or Excel (.xlsx). Required columns: **Name**, **CNIC**.
Optional: Father Name, Roll No, Discipline, Test Date, Obtained, Total, Merit No.

Example:

| Name | Father Name | CNIC | Roll No | Discipline | Test Date | Obtained | Total | Merit No |
|---|---|---|---|---|---|---|---|---|
| Ahmed Raza | Muhammad Aslam | 3520212345671 | PHD-2026-001 | Veterinary Anatomy | 12-Sep-2026 | 78 | 100 | 3 |

Column headers are matched flexibly ("Candidate Name", "Roll Number", "Marks Obtained" all work).
CNICs may be written with or without dashes — students can type them either way.

## Admin features

- Upload / replace the result list (re-uploading the same CNIC updates that record)
- Set the passing percentage (default 50%)
- Publish / unpublish the portal (hide results before the official announcement)
- View all uploaded records, delete everything when starting a new test cycle

## Files

- `app.py` — the whole application
- `templates/` — pages (student search, result card, admin)
- `static/css/style.css` — green CUVAS theme
- `results.db` — SQLite database (created on first run)
- `seed.py` — sample data for testing
