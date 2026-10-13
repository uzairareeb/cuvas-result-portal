"""Seed sample admission-test data for testing. Run: python3 seed.py"""
import sqlite3, os
from app import init_db, norm_cnic

init_db()
db = sqlite3.connect(os.path.join(os.path.dirname(os.path.abspath(__file__)), "results.db"))

SAMPLE = [
    # name, father, cnic, degree, held_on, total, obtained
    ("Manzar Hayat", "Khizar Hayat", "3610346954971", "PhD Zoology", "15-Sep-2025", 50, 45),
    ("Ayesha Khan", "Muhammad Khan", "3520212345671", "PhD Biochemistry", "15-Sep-2025", 100, 72),
    ("Usman Tariq", "Tariq Mehmood", "3520234567893", "PhD Theriogenology", "15-Sep-2025", 50, 22),
    ("Fatima Raza", "Ali Raza", "3520245678904", "PhD Anatomy", "15-Sep-2025", 100, 58),
]

db.execute("DELETE FROM candidates")
for s in SAMPLE:
    db.execute("""INSERT INTO candidates
        (name, father_name, cnic, roll_no, discipline, test_date, obtained, total, merit_no, uploaded_at)
        VALUES (?,?,?,?,?,?,?,?,?, datetime('now'))""",
        (s[0], s[1], norm_cnic(s[2]), "", s[3], s[4], s[6], s[5], ""))
db.commit()
print(f"Seeded {len(SAMPLE)} sample candidates.")
print("Try CNIC: 36103-4695497-1 (Manzar Hayat, 45/50 Qualified)")
print("Try CNIC: 35202-3456789-3 (Usman Tariq, 22/50 Not Qualified)")
db.close()
