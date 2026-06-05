# Development Guide

## Overview

Medical Records Generator is an internal desktop application designed to automate the creation of Medical Records packages for payer submissions.

The application processes claim information from CSV files, matches supporting documentation, and generates organized PDF packages ready for submission.

---

# Technology Stack

## Language

* Python 3.10+

## GUI

* Tkinter
* ttk

## PDF Processing

* pypdf
* PyMuPDF (fitz)

## OCR

* Windows OCR (WinRT)
* Pillow

## Data Processing

* pandas

---

# Project Structure

Current project structure:

```text
MR_APP/
│
├── assets/
│   └── app.ico
│
├── documentation/
│   ├── CHANGELOG.md
│   ├── DEVELOPMENT_GUIDE.md
│   ├── SOP_MR_GENERATION.md
│   └── RELEASE_CHECKLIST.md
│
├── MR_APP.py
├── README.md
├── requirements.txt
└── User Manual
```

---

# Core Functional Areas

## CSV Processing

Responsible for:

* Reading claim information
* Validating required columns
* Grouping claims
* Preparing claim metadata

---

## SOAP Matching

SOAP pages are matched using:

* DOS (Date of Service)
* CPT Codes

This allows the application to include only the relevant SOAP pages for each claim.

---

## OCR Processing

OCR is used to:

* Read payer letters
* Extract claim identifiers
* Associate payer correspondence with claims

Requirements:

* Windows OCR support
* English language OCR package installed

---

## PDF Assembly

Generated packages may include:

* Cover Letter
* Table of Contents
* AOR
* Base Documents
* SOAP Notes
* Payer Letters
* Optional Supporting Documents

---

# Build Instructions

## Create Virtual Environment

```bash
python -m venv .venv
```

## Install Dependencies

```bash
pip install -r requirements.txt
```

## Run Application

```bash
python MR_APP.py
```

---

# Portable Build

```bash
pyinstaller --noconfirm --clean --onefile --windowed --name "MassMedicalRecordsGenerator" --icon "assets\\app.ico" --add-data "assets\\app.ico;assets" MR_APP.py
```

Generated executable:

```text
dist/
└── MassMedicalRecordsGenerator.exe
```

---

# Maintenance Guidelines

Before modifying production logic:

1. Validate existing workflows.
2. Test both Individual and Bulk modes.
3. Verify SOAP matching accuracy.
4. Verify OCR functionality.
5. Verify PDF generation integrity.
6. Update CHANGELOG.md.
7. Update documentation when applicable.

---

# Future Improvements

Potential future enhancements:

* Modular code structure.
* Additional payer-specific automation.
* Enhanced OCR accuracy.
* Automated document validation.
* Extended reporting capabilities.
* Installer-based deployment.
