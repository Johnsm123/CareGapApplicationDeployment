# Cancer-Screening Test Dataset Generator

Generates a **perfect** patient dataset (`cancer_screening_test_members.xlsx`)
for the HEDIS care-gap app's bulk-upload endpoint
(`POST /api/v1/members/bulk-upload`).

Every member is engineered so the app's pure-Python detector
(`src/care_gap_agents.py → detect_care_gaps`) produces an **exact, predictable**
result: the right measures applicable, the right gaps open, and therefore the
right dashboard pill.

## Run

```bash
cd test_data_generator
python generate_cancer_screening_dataset.py
```

Output → `./output/cancer_screening_test_members.xlsx`.

Options:

| Flag | Purpose |
|------|---------|
| `--out PATH`     | Write somewhere else (e.g. `--out ../cancer_screening_test_members.xlsx`). |
| `--email ADDR`   | Email applied to every member (default `ajohnsm2020@gmail.com`). |
| `--date YYYY-MM-DD` | Override "today" for reproducible runs. |

Before writing, the script **self-validates** every member by replaying the
app's eligibility + lookback math and aborts if any member would not land in
its intended pill.

## Why a generator (and not a hand-edited sheet)

The bulk-upload parser is strict:

- **Extended fields are pipe-delimited, records semicolon-separated.**
  `FamilyHistory = "relation|alive|age|cond1,cond2;relation|alive|age|cond1"`.
  A sheet that writes `Mother:Breast Cancer:55` (colons) is silently mis-parsed.
- **`PriorScreenings` is `MEASURE:DATE` pairs**, e.g. `BCS:2025-09-15;COL:2020-06-18`.
- **Ages and screening dates must sit inside the lookback window.** The script
  computes both **relative to today**, so the dataset is always valid no matter
  when you run it.

## Detection model (HEDIS MY2025 golden reference)

| Measure | Eligibility | Lookback | Prior-CPT that closes it |
|---------|-------------|----------|--------------------------|
| **BCS** Breast Cancer Screening | Female, age 52–74 | 24 months | `77067` mammogram |
| **CCS** Cervical Cancer Screening | Female, age 21–64 | 36 months (cytology) | `88175` liquid-based Pap |
| **COL** Colorectal Cancer Screening | Any gender, age 45–75 | 120 months (colonoscopy) | `45378` colonoscopy |

- **Default app scope** is `BCS,CCS,COL` (env `CARE_GAP_ENABLED_MEASURES`).
- **Dashboard pills:** Critical = `open_gaps ≥ 3`, Needs Attention = `1–2`, Compliant = `0`.

## Member matrix

| # | Name | Profile | Priors | Open gaps | Pill |
|---|------|---------|--------|-----------|------|
| M1 | Margaret Chen | F, 60 | none | BCS + CCS + COL (3) | **Critical** |
| M2 | Sandra Patel | F, 60 | BCS | CCS + COL (2) | **Needs Attention** |
| M3 | Emily Rodriguez | F, 35 | none | CCS (1) — age<45 → COL N/A, age<52 → BCS N/A | **Needs Attention** |
| M4 | Daniel Williams | M, 55 | none | COL (1) — male → BCS/CCS N/A | **Needs Attention** |
| M5 | Linda Johnson | F, 60 | BCS + CCS + COL | none (0) | **Compliant** |

**Expected dashboard counts:** Critical = 1, Needs Attention = 3, Compliant = 1, Total = 5.

## Upload

Use the in-app bulk-upload page, or:

```bash
curl -F "file=@output/cancer_screening_test_members.xlsx" <host>/api/v1/members/bulk-upload
```

## Adding / changing members

Edit `build_rows()` and add the member's expected result to the `EXPECTED`
dict. Keep `_MEASURE_RULES` in sync with `src/hedis_golden_reference.py` if the
rulebook changes. The self-validation step will refuse to write a sheet whose
detection doesn't match `EXPECTED`, so a mistake fails loudly instead of
producing a misleading dataset.

> **Scope note:** this matrix is built for the default `BCS,CCS,COL` scope. If
> you run the app with `CARE_GAP_ENABLED_MEASURES` including `AAP` (Adults'
> Access, age 20+, 12-mo PCP visit), members without a recent `AAP:` prior will
> also show an AAP gap — extend `_MEASURE_RULES`, `EXPECTED`, and the rows to
> cover it.
```
