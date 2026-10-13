"""Seed sample PhD entry-test data for testing. Run: python3 seed.py"""
import sqlite3, os
from app import init_db, norm_cnic

init_db()
db = sqlite3.connect(os.path.join(os.path.dirname(os.path.abspath(__file__)), "results.db"))

SAMPLE = [
    ("Ahmed Raza", "Muhammad Aslam", "3520212345671", "PHD-2026-001", "Veterinary Anatomy", "12-Sep-2026", 78, 100, "3"),
    ("Fatima Khan", "Abdul Rehman Khan", "3520223456782", "PHD-2026-002", "Animal Nutrition", "12-Sep-2026", 64, 100, "11"),
    ("Usman Tariq", "Tariq Mehmood", "3520234567893", "PHD-2026-003", "Theriogenology", "12-Sep-2026", 42, 100, ""),
    ("Ayesha Siddiqui", "Muhammad Siddiq", "3520245678904", "PHD-2026-004", "Veterinary Pathology", "12-Sep-2026", 85, 100, "1"),
    ("Bilal Hussain", "Ghulam Hussain", "3520256789015", "PHD-2026-005", "Livestock Management", "12-Sep-2026", 55, 100, "18"),
    ("Sana Malik", "Malik Ashraf", "3520267890126", "PHD-2026-006", "Veterinary Microbiology", "12-Sep-2026", 71, 100, "6"),
    ("Imran Yousaf", "Muhammad Yousaf", "3520278901237", "PHD-2026-007", "Poultry Science", "12-Sep-2026", 38, 100, ""),
    ("Hina Shahid", "Shahid Mehmood", "3520289012348", "PHD-2026-008", "Veterinary Pharmacology", "12-Sep-2026", 66, 100, "9"),
    ("Kamran Ali", "Ali Nawaz", "3520290123459", "PHD-2026-009", "Animal Breeding & Genetics", "12-Sep-2026", 59, 100, "14"),
    ("Mahnoor Fatima", "Rashid Mehmood", "3520201234560", "PHD-2026-010", "Veterinary Parasitology", "12-Sep-2026", 81, 100, "2"),
]

db.execute("DELETE FROM candidates")
for s in SAMPLE:
    db.execute("""INSERT INTO candidates
        (name, father_name, cnic, roll_no, discipline, test_date, obtained, total, merit_no, uploaded_at)
        VALUES (?,?,?,?,?,?,?,?,?, datetime('now'))""",
        (s[0], s[1], norm_cnic(s[2]), s[3], s[4], s[5], s[6], s[7], s[8]))
db.commit()
print(f"Seeded {len(SAMPLE)} sample candidates.")
print("Try CNIC: 35202-1234567-1  (Ahmed Raza, PASS)")
print("Try CNIC: 35202-3456789-3  (Usman Tariq, FAIL)")
db.close()
