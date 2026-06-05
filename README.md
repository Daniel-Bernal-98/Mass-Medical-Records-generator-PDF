# Mass Medical Records Generator (PDF) — HIPAA-friendly workflow

GUI tool to generate Medical Records (MR) PDFs from a **claims CSV** and a **SOAP Notes PDF**, with consistent document ordering and an auto-generated **Table of Contents (TOC)**.

This project supports two workflows:

1. **Per-Claim PDFs**: generates **one PDF per `claim_id`** (recommended for claim-by-claim submissions).
2. **Bulk PDF**: generates **one single PDF for all `claim_id`s in the CSV** (recommended for a date range submission covering multiple claims).

> Notes:
> - This tool is designed to reduce manual effort and improve consistency. Always validate outputs before sending.
> - All documents are **PDF**, except the input **CSV**.

---

## Requirements

- Windows 10/11 recommended
- Python 3.10+ recommended
- Packages:
  - `pypdf`
  - `PyMuPDF` (imported as `fitz`)

Install:

```bash
pip install -r requirements.txt
```

Run:

```bash
python MR_APP.py
```

---

## Inputs (Files you select in the UI)

### Required (both modes)
- **Claims List (CSV)** — drives claim grouping, DOS/CPT lookup, and bulk claim list generation
- **Cover Letter (PDF)**
- **AOR - Consent Form (PDF)** *(mandatory)*
- **Base Docs (PDFs)** — select the PDFs that represent:
  - **TX** (Treatment plan)
  - **DX** (Diagnosis Report)
  - **Progress Report** (TOC label: Progress Notes)
- **SOAP Notes (PDF)** — large PDF containing notes for multiple dates (the tool extracts pages per DOS/CPT)
- **Output Folder**

### Optional
- **Payer Letter (PDF)**
- **Target List Report (PDF)** *(for now this is a PDF; Excel→PDF automation will be added later)*
- **DTT Trial Sheets (PDF)** *(for now it is appended as-is; automation can be added later)*
- **Behavior Reduction Report (PDF)**

---

## CSV format

### Required columns
- `claim_id` — used to group rows (multiple lines per claim are supported)
- `DOS` — Date of Service used to match SOAP pages
- `CPT` — used as a secondary filter for SOAP pages (fallback to DOS-only if no CPT match)

### Bulk-only column (required for the claim list table)
- `billed_amount`

### Recommended column (used for bulk output filename)
- `patient_name`

> The app also supports common header variants for claim and DOS (case-insensitive), but the recommended/standard headers are shown above.

---

## SOAP Notes matching rules

SOAP pages are selected per claim using:

1. **DOS match** (using the `SESSION DATE` text found inside SOAP pages)
2. If the claim has CPT(s) in the CSV:
   - the tool filters to pages containing **`CPT CODE`** and one of the claim CPT codes
3. If CPT-filtering finds no pages, it falls back to **DOS-only** pages

This preserves reliability even when some SOAP pages don’t contain CPT codes in a consistent way.

---

## Output naming (MM.DD.YY format)

All dates are formatted as: **`MM.DD.YY`**

### Per-Claim PDFs
One output per claim:

```
<claim_id>-<DOS>-sent-<today>.pdf
```

Example:

```
123456-04.15.26-sent-05.04.26.pdf
```

> DOS is taken from the CSV. If multiple DOS values exist for the same claim, the app logs a warning and uses the first value.

### Bulk PDF
One output for all claims in the CSV:

```
<patient_name>-<initialDOS>-<finalDOS>-sent-<today>.pdf
```

Example:

```
John_Smith-04.01.26-04.30.26-sent-05.04.26.pdf
```

- `initialDOS` = earliest DOS found in the CSV
- `finalDOS` = latest DOS found in the CSV
- If `patient_name` is missing, a fallback name may be used.

---

## PDF order and TOC labels

### Per-Claim PDFs (one document per `claim_id`)
Final PDF order:

1. Cover Letter
2. **Payer Letter** (optional)
3. **TX** *(TOC label: `Treatment plan`)*
4. **DX** *(TOC label: `Diagnosis Report`)*
5. **AOR - Consent Form** *(mandatory)*
6. **SOAP Notes** *(matched by DOS + CPT, fallback to DOS-only)*
7. **Progress Report** *(TOC label: `Progress Notes`)*
8. **Target List Report** (optional)
9. **DTT Trial Sheets** (optional)
10. **Behavior Reduction Report** (optional)

### Bulk PDF (one document for all `claim_id`s)
Final PDF order:

1. Cover Letter
2. **Payer Letter** (optional)
3. **Claim List** *(generated from CSV: `claim_id`, `DOS`, `billed_amount`)*
4. **TX** *(Treatment plan)*
5. **DX** *(Diagnosis Report)*
6. **AOR - Consent Form** *(mandatory)*
7. **SOAP Notes** *(appended for every claim_id in the CSV)*
8. **Progress Report** *(Progress Notes)*
9. **Target List Report** (optional)
10. **DTT Trial Sheets** (optional)
11. **Behavior Reduction Report** (optional)

---

## Handling duplicate `claim_id` rows (important)

Some CSV exports contain multiple lines per claim (same `claim_id`) for different CPT codes.

This app:
- groups rows by `claim_id`
- collects all CPT codes for that claim
- generates **one PDF per `claim_id`** in Per-Claim mode
- uses the same grouped claims to generate **one combined output** in Bulk mode

If a single `claim_id` contains multiple DOS values in the CSV, the app logs a warning and uses the first DOS.

---

## Troubleshooting

### “No SOAP pages matched”
- Confirm the DOS format in CSV is valid
- Confirm SOAP pages contain `SESSION DATE`
- Confirm CPT values exist in CSV column `CPT` (optional but improves accuracy)

### Bulk output name looks wrong
- Ensure the CSV includes `patient_name`
- Ensure DOS values are parseable so the app can compute initial/final DOS

---

## Project

### Entry point
Run the GUI:

```bash
python MR_APP.py
```

### Compliance note
This repository provides automation support for document assembly. It does **not** provide legal/compliance guarantees.  
Always follow your organization’s HIPAA policies and validate the output PDFs before sending.

### License & Usage Restrictions

This software is proprietary and developed for internal corporate use only.

Unauthorized use, reproduction, modification, distribution, or sharing of this software, in whole or in part, is strictly prohibited without prior written authorization from the repository owner.

No part of this software may be copied, reproduced, modified, distributed, or transmitted in any form or by any means without prior written permission.

Access to this codebase does not grant any rights or licenses for external use or distribution.

Any violation of these terms may result in legal action in accordance with applicable intellectual property laws.

---