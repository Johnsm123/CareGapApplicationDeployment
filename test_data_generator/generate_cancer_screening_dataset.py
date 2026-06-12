"""
generate_cancer_screening_dataset.py
====================================

Generates a "perfect" cancer-screening patient dataset (.xlsx) for the
HEDIS care-gap application's bulk-upload endpoint
(POST /api/v1/members/bulk-upload).

Every member is engineered so the application's PURE-PYTHON detector
(`src.care_gap_agents.detect_care_gaps`) produces an exact, predictable
result — the right measures applicable, the right gaps open, and therefore
the right dashboard pill (Critical / Needs Attention / Compliant).

WHY THIS SCRIPT EXISTS
----------------------
The bulk-upload parser is strict about formats:

  * Extended fields are PIPE-delimited, records SEMICOLON-separated, e.g.
        FamilyHistory  -> "relation|alive|age|cond1,cond2;relation|alive|age|cond1"
        PastConditions -> "name|year|status|notes;..."
        Allergies      -> "substance|severity|reaction;..."
        Medications    -> "name|dose|started|purpose;..."
        Immunizations  -> "name|year;..."
  * PriorScreenings is MEASURE:DATE pairs, semicolon-separated, e.g.
        "BCS:2025-09-15;COL:2020-06-18"

A hand-edited sheet that uses colons inside FamilyHistory (e.g.
"Mother:Breast Cancer:55") is silently mis-parsed. This generator emits the
exact format the parser consumes, and computes every age / screening date
RELATIVE TO TODAY so the dataset is always inside the correct lookback window
no matter when you run it.

DETECTION MODEL (HEDIS MY2025 golden reference)
-----------------------------------------------
  Measure  Eligibility            Lookback              Prior-CPT that closes it
  BCS      Female, age 52-74      24 months             77067 (screening mammogram)
  CCS      Female, age 21-64      36 months (cytology)  88175 (liquid-based Pap)
  COL      Any gender, age 45-75  120 months (scope)    45378 (colonoscopy)

Default app scope (env CARE_GAP_ENABLED_MEASURES, default "BCS,CCS,COL").
Dashboard pills: Critical = open_gaps >= 3, Needs Attention = 1-2, Compliant = 0.

MEMBER MATRIX (built for the default BCS,CCS,COL scope)
------------------------------------------------------
  M1  Margaret Chen   F age 60  no priors                 -> BCS+CCS+COL open (3) -> CRITICAL
  M2  Sandra Patel    F age 60  BCS prior                 -> CCS+COL open (2)     -> NEEDS ATTENTION
  M3  Emily Rodriguez F age 35  no priors                 -> CCS open (1)         -> NEEDS ATTENTION
                                                              (age<45 -> COL N/A; age<52 -> BCS N/A)
  M4  Daniel Williams M age 55  no priors                 -> COL open (1)         -> NEEDS ATTENTION
                                                              (male -> BCS/CCS N/A)
  M5  Linda Johnson   F age 60  BCS+CCS+COL all priors    -> 0 open               -> COMPLIANT

Expected dashboard counts:  Critical=1  Needs Attention=3  Compliant=1  Total=5

USAGE
-----
    python generate_cancer_screening_dataset.py
    python generate_cancer_screening_dataset.py --out ../cancer_screening_test_members.xlsx
    python generate_cancer_screening_dataset.py --email you@example.com

By default the file is written next to this script in ./output/.
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import date, datetime, timedelta

import pandas as pd

# ── Column order: matches the bulk-upload parser's expected headers ──────────
COLUMNS = [
    "Name", "DOB", "Gender", "Email", "Phone", "PCPID", "PlanID", "ZIP",
    "ChronicConditions", "InsuranceType", "EnrollmentStart", "EnrollmentEnd",
    "PriorScreenings",
    "HeightCm", "WeightKg", "SmokingStatus", "AlcoholUse", "ExerciseFrequency",
    "DietType", "SleepHoursAvg", "StressLevel", "LifestyleNotes",
    "FamilyHistory", "PastConditions", "CurrentConditions", "Surgeries",
    "Allergies", "Medications", "Immunizations",
]

DEFAULT_EMAIL = "ajohnsm2020@gmail.com"


# ── Date helpers — everything is relative to "today" so the dataset never
#    drifts out of a lookback window. ──────────────────────────────────────────
def dob_for_age(age: int, today: date) -> str:
    """
    Return an ISO DOB (YYYY-MM-DD) that yields EXACTLY `age` years today,
    robust to whatever month the script is run in. We anchor to mid-current-
    month, drop the year back by `age`, then back off ~60 days so the birthday
    has already occurred this year.
    """
    anchor = date(today.year - age, today.month, 15)
    birth = anchor - timedelta(days=60)
    return birth.isoformat()


def months_ago(today: date, months: int) -> str:
    """Return an ISO date `months` calendar-months before today."""
    m = today.month - months
    y = today.year + (m - 1) // 12
    m = (m - 1) % 12 + 1
    day = min(today.day, 28)  # avoid month-length edge cases
    return date(y, m, day).isoformat()


def build_rows(today: date, email: str) -> list[dict]:
    enroll_start = date(today.year - 2, 1, 1).isoformat()   # 2 yrs continuous enrollment
    enroll_end = date(today.year, 12, 31).isoformat()

    # Prior-screening dates, each comfortably inside its lookback window.
    bcs_prior = months_ago(today, 6)    # within 24-mo BCS window
    ccs_prior = months_ago(today, 12)   # within 36-mo CCS window
    col_prior = months_ago(today, 36)   # within 120-mo COL window

    return [
        # ── M1 — CRITICAL: F60, no priors -> BCS + CCS + COL all open (3) ──────
        {
            "Name": "Margaret Chen", "DOB": dob_for_age(60, today), "Gender": "F",
            "Email": email, "Phone": "15551110001",
            "PCPID": "P1001", "PlanID": "PLAN-001", "ZIP": "10001",
            "ChronicConditions": "",
            "InsuranceType": "Commercial",
            "EnrollmentStart": enroll_start, "EnrollmentEnd": enroll_end,
            "PriorScreenings": "",
            "HeightCm": 162, "WeightKg": 65,
            "SmokingStatus": "Never", "AlcoholUse": "Occasional",
            "ExerciseFrequency": "2-3x/week", "DietType": "Balanced",
            "SleepHoursAvg": 7, "StressLevel": "Low",
            "LifestyleNotes": "Walks daily; no current concerns.",
            "FamilyHistory": "Mother|true|82|None;Father|false|78|Heart Disease",
            "PastConditions": "Seasonal allergies|2010|Resolved|Mild",
            "CurrentConditions": "",
            "Surgeries": "",
            "Allergies": "Penicillin|Mild|Rash",
            "Medications": "Multivitamin|1 tab|2018|General wellness",
            "Immunizations": "Influenza|2025;Tdap|2022",
        },
        # ── M2 — NEEDS ATTENTION: F60, BCS prior -> CCS + COL open (2) ────────
        {
            "Name": "Sandra Patel", "DOB": dob_for_age(60, today), "Gender": "F",
            "Email": email, "Phone": "15551110002",
            "PCPID": "P1002", "PlanID": "PLAN-001", "ZIP": "10002",
            "ChronicConditions": "",
            "InsuranceType": "Commercial",
            "EnrollmentStart": enroll_start, "EnrollmentEnd": enroll_end,
            "PriorScreenings": f"BCS:{bcs_prior}",
            "HeightCm": 158, "WeightKg": 62,
            "SmokingStatus": "Never", "AlcoholUse": "None",
            "ExerciseFrequency": "Daily", "DietType": "Mediterranean",
            "SleepHoursAvg": 8, "StressLevel": "Low",
            "LifestyleNotes": "Active retiree; volunteers weekly.",
            "FamilyHistory": "Mother|true|85|None;Sister|true|62|Breast Cancer",
            "PastConditions": "Vitamin D deficiency|2019|Resolved|Supplementation",
            "CurrentConditions": "",
            "Surgeries": "",
            "Allergies": "",
            "Medications": "Vitamin D3|2000 IU|2019|Bone health",
            "Immunizations": "Influenza|2025;Pneumococcal|2024;Shingles|2023",
        },
        # ── M3 — NEEDS ATTENTION: F35, no priors -> CCS open only (1) ─────────
        {
            "Name": "Emily Rodriguez", "DOB": dob_for_age(35, today), "Gender": "F",
            "Email": email, "Phone": "15551110003",
            "PCPID": "P1003", "PlanID": "PLAN-001", "ZIP": "10003",
            "ChronicConditions": "",
            "InsuranceType": "Commercial",
            "EnrollmentStart": enroll_start, "EnrollmentEnd": enroll_end,
            "PriorScreenings": "",
            "HeightCm": 165, "WeightKg": 60,
            "SmokingStatus": "Never", "AlcoholUse": "Occasional",
            "ExerciseFrequency": "3-4x/week", "DietType": "Balanced",
            "SleepHoursAvg": 7, "StressLevel": "Moderate",
            "LifestyleNotes": "Software engineer; sedentary work, runs on weekends.",
            "FamilyHistory": "Mother|true|65|None;Father|true|68|Hypertension",
            "PastConditions": "",
            "CurrentConditions": "",
            "Surgeries": "",
            "Allergies": "",
            "Medications": "",
            "Immunizations": "Influenza|2025;Tdap|2021;HPV|2010",
        },
        # ── M4 — NEEDS ATTENTION: M55, no priors -> COL open only (1) ─────────
        {
            "Name": "Daniel Williams", "DOB": dob_for_age(55, today), "Gender": "M",
            "Email": email, "Phone": "15551110004",
            "PCPID": "P1004", "PlanID": "PLAN-001", "ZIP": "10004",
            "ChronicConditions": "",
            "InsuranceType": "Commercial",
            "EnrollmentStart": enroll_start, "EnrollmentEnd": enroll_end,
            "PriorScreenings": "",
            "HeightCm": 178, "WeightKg": 82,
            "SmokingStatus": "Former", "AlcoholUse": "Moderate",
            "ExerciseFrequency": "2x/week", "DietType": "Mixed",
            "SleepHoursAvg": 6, "StressLevel": "Moderate",
            "LifestyleNotes": "Quit smoking 2018; occasional gym workouts.",
            "FamilyHistory": "Father|false|72|Colorectal Cancer;Mother|true|78|None",
            "PastConditions": "Tobacco use disorder|2018|Resolved|Quit successfully",
            "CurrentConditions": "",
            "Surgeries": "Appendectomy|2002|Uneventful",
            "Allergies": "",
            "Medications": "",
            "Immunizations": "Influenza|2025;Tdap|2020",
        },
        # ── M5 — COMPLIANT: F60, all three priors -> 0 open ───────────────────
        {
            "Name": "Linda Johnson", "DOB": dob_for_age(60, today), "Gender": "F",
            "Email": email, "Phone": "15551110005",
            "PCPID": "P1005", "PlanID": "PLAN-001", "ZIP": "10005",
            "ChronicConditions": "",
            "InsuranceType": "Commercial",
            "EnrollmentStart": enroll_start, "EnrollmentEnd": enroll_end,
            "PriorScreenings": f"BCS:{bcs_prior};CCS:{ccs_prior};COL:{col_prior}",
            "HeightCm": 160, "WeightKg": 63,
            "SmokingStatus": "Never", "AlcoholUse": "Occasional",
            "ExerciseFrequency": "Daily", "DietType": "Mediterranean",
            "SleepHoursAvg": 8, "StressLevel": "Low",
            "LifestyleNotes": "Retired teacher; engaged with primary care annually.",
            "FamilyHistory": "Mother|true|88|None;Father|true|85|None",
            "PastConditions": "Hypothyroidism|2010|Resolved|Levothyroxine",
            "CurrentConditions": "",
            "Surgeries": "Cataract|2023|Right eye",
            "Allergies": "Sulfa|Mild|Rash",
            "Medications": "Vitamin D3|1000 IU|2020|Bone health",
            "Immunizations": "Influenza|2025;Pneumococcal|2024;Shingles|2024;Tdap|2022",
        },
    ]


# ── Self-validation: replicate the app's eligibility + lookback logic so we
#    can PROVE each member lands in the intended pill before writing the file.
EXPECTED = {
    "Margaret Chen":   ("CRITICAL",        {"BCS", "CCS", "COL"}),
    "Sandra Patel":    ("NEEDS ATTENTION", {"CCS", "COL"}),
    "Emily Rodriguez": ("NEEDS ATTENTION", {"CCS"}),
    "Daniel Williams": ("NEEDS ATTENTION", {"COL"}),
    "Linda Johnson":   ("COMPLIANT",       set()),
}

# (age_min, age_max, gender_required, lookback_months) per the golden reference.
_MEASURE_RULES = {
    "BCS": (52, 74, "F", 24),
    "CCS": (21, 64, "F", 36),
    "COL": (45, 75, None, 120),
}


def _age_on(dob_iso: str, today: date) -> int:
    b = datetime.strptime(dob_iso, "%Y-%m-%d").date()
    return today.year - b.year - ((today.month, today.day) < (b.month, b.day))


def _predict(row: dict, today: date) -> tuple[str, set]:
    age = _age_on(row["DOB"], today)
    gender = row["Gender"][:1].upper()
    priors = {}
    for entry in (row["PriorScreenings"] or "").split(";"):
        entry = entry.strip()
        if ":" in entry:
            mid, dstr = entry.split(":", 1)
            priors[mid.strip().upper()] = datetime.strptime(dstr.strip()[:10], "%Y-%m-%d").date()
    open_gaps = set()
    for mid, (amin, amax, greq, lookback) in _MEASURE_RULES.items():
        if greq and gender != greq:
            continue                      # gender N/A
        if not (amin <= age <= amax):
            continue                      # age N/A
        prior = priors.get(mid)
        cutoff = today - timedelta(days=lookback * 30)
        if prior and prior >= cutoff:
            continue                      # compliant
        open_gaps.add(mid)
    n = len(open_gaps)
    pill = "CRITICAL" if n >= 3 else "NEEDS ATTENTION" if n >= 1 else "COMPLIANT"
    return pill, open_gaps


def validate(rows: list[dict], today: date) -> bool:
    ok = True
    print("\nSelf-validation (replicates app eligibility + lookback logic):")
    for row in rows:
        name = row["Name"]
        pill, gaps = _predict(row, today)
        exp_pill, exp_gaps = EXPECTED.get(name, (None, None))
        match = (pill == exp_pill and gaps == exp_gaps)
        ok = ok and match
        flag = "OK " if match else "!! "
        age = _age_on(row["DOB"], today)
        gaps_str = ", ".join(sorted(gaps)) or "none"
        print(f"  {flag}{name:<16} {row['Gender']} age {age:<3} "
              f"open=[{gaps_str:<18}] -> {pill}"
              + ("" if match else f"   EXPECTED {exp_pill} {sorted(exp_gaps)}"))
    print("  =>", "ALL MEMBERS MATCH EXPECTED DETECTION" if ok else "MISMATCH DETECTED")
    return ok


def main(argv: list[str] | None = None) -> int:
    here = os.path.dirname(os.path.abspath(__file__))
    default_out = os.path.join(here, "output", "cancer_screening_test_members.xlsx")

    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=default_out,
                    help="Output .xlsx path (default: ./output/cancer_screening_test_members.xlsx)")
    ap.add_argument("--email", default=DEFAULT_EMAIL,
                    help="Email address applied to every generated member.")
    ap.add_argument("--date", default=None,
                    help="Override 'today' as YYYY-MM-DD (for reproducible runs).")
    args = ap.parse_args(argv)

    today = datetime.strptime(args.date, "%Y-%m-%d").date() if args.date else date.today()

    rows = build_rows(today, args.email)

    if not validate(rows, today):
        print("\nABORTING: generated members do not match expected detection.", file=sys.stderr)
        return 1

    df = pd.DataFrame(rows, columns=COLUMNS)
    out = args.out
    os.makedirs(os.path.dirname(out), exist_ok=True)
    try:
        df.to_excel(out, index=False)
    except PermissionError:
        stamp = datetime.now().strftime("%H%M%S")
        out = out.replace(".xlsx", f"_{stamp}.xlsx")
        df.to_excel(out, index=False)
        print(f"\nNOTE: target was locked (open in Excel); wrote to {out} instead.")

    print(f"\nWrote {len(df)} members ('today' = {today.isoformat()}) to:\n  {out}")
    print("\nExpected dashboard pill counts: Critical=1  Needs Attention=3  Compliant=1  Total=5")
    print("Upload via the bulk-upload page or:")
    print(f'  curl -F "file=@{out}" <host>/api/v1/members/bulk-upload')
    return 0


if __name__ == "__main__":
    sys.exit(main())
