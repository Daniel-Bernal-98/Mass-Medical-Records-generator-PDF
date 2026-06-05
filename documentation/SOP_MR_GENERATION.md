# SOP - Medical Records Package Generation

## Purpose

To standardize the generation of Medical Records packages and ensure consistency, accuracy, and compliance during payer submissions.

---

# Scope

This procedure applies to all personnel responsible for preparing Medical Records packages using the Medical Records Generator application.

---

# Responsibilities

The operator is responsible for:

* Verifying source documentation.
* Ensuring claim information accuracy.
* Reviewing generated output.
* Reporting application issues.

---

# Required Inputs

Before processing, verify the availability of:

* Claims CSV
* Cover Letter
* AOR Document
* Base Documents
* Output Folder

Optional documentation:

* Payer Letters
* Target List
* DTT Documents
* Daily Behavior Data
* Daily Trial Counts
* Behavior Reduction Reports
* Filter File

---

# Procedure

## Step 1 - Validate Source Documents

Confirm:

* All PDFs are accessible.
* Files are complete and readable.
* CSV contains valid claim information.

---

## Step 2 - Launch Application

Open:

```text
MassMedicalRecordsGenerator.exe
```

Wait until the application is fully loaded.

---

## Step 3 - Load Required Documents

Select:

1. Claims CSV
2. Cover Letter
3. AOR
4. Base Documents
5. Output Folder

---

## Step 4 - Load Optional Documents

If available:

* Payer Letters
* Target Lists
* DTT Documents
* Additional supporting documentation

---

## Step 5 - Execute Processing

Choose one of:

### Generate PDFs

Creates one package per claim.

### Generate Bulk PDF

Creates a single consolidated package.

---

## Step 6 - Review Results

Verify:

* Claim ID
* DOS
* CPT matching
* SOAP pages
* Payer correspondence
* Table of Contents
* Document order

---

## Quality Control

Every generated package must be reviewed before submission.

Recommended checks:

* First page accuracy
* Correct claim information
* Proper SOAP inclusion
* No missing pages
* No duplicate pages

---

# Incident Handling

If processing fails:

1. Capture the error message.
2. Save execution logs.
3. Record claim identifiers.
4. Notify application support.

---

# Records Retention

Generated PDFs should be stored according to organizational document retention policies.
