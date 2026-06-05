"""
Mass Medical Records PDF Generator
==================================

GUI tool to generate:
1) One PDF per claim_id from a CSV list (even if the CSV contains multiple lines per claim).
2) Bulk mode: One single PDF for ALL claim_ids in the CSV (for a DOS date range workflow).

SOAP matching rules:
- Primary filter: DOS (SESSION DATE)
- Secondary filter: CPT (matches "CPT CODE" line on SOAP pages)
- If CPT filter finds no pages, fallback to DOS-only pages.

DOCUMENT ORDER / TOC
--------------------

Per-claim (one doc per claim_id) order:
  1) Cover Letter
  2) Payer Letter(s) (optional; OCR-matched per claim)
  3) TX (TOC label: Treatment plan)
  4) DX (TOC label: Diagnosis Report)
  5) AOR - Consent Form (mandatory)
  6) SOAP Notes (matched by DOS + CPT, fallback to DOS-only)
  7) Progress Report (TOC label: Progress Notes)
  8) Target List Report (mandatory, PDF)
  9) DTT Trial Sheets (mandatory, multiple PDFs supported)
  10) Daily Behavior Data (mandatory, multiple PDFs supported)
  11) Daily Trial Counts (mandatory, multiple PDFs supported)
  12) Behavior Reduction Report (mandatory, multiple PDFs supported)

Bulk (one doc for all claim_ids) order:
  1) Cover Letter (bulk mode: do NOT fill, use as-is)
  2) Payer Letter(s) (optional; ALL payer letters included once, not filtered)
  3) Claim List (generated from CSV: claim_id, DOS, billed_amount)
  4) TX (Treatment plan)
  5) DX (Diagnosis Report)
  6) AOR - Consent Form (mandatory)
  7) SOAP Notes (for each claim_id, appended in group order)
  8) Progress Report (Progress Notes)
  9) Target List Report (mandatory, PDF)
  10) DTT Trial Sheets (mandatory, multiple PDFs supported)
  11) Daily Behavior Data (mandatory, multiple PDFs supported)
  12) Daily Trial Counts (mandatory, multiple PDFs supported)
  13) Behavior Reduction Report (mandatory, multiple PDFs supported)

Filenames:
- Per-claim: claim_id-<DOS>-sent-<todays_date>.pdf
- Bulk: patient_name-<initialDOS>-<finalDOS>-sent-<todays_date>.pdf

All dates formatted: MM.DD.YY
"""

import csv
import os
import threading
import tempfile
import time
from datetime import datetime
import re
import sys

# OCR Imports (Windows OCR WinRT + PIL)
from PIL import Image

from winrt.windows.media.ocr import OcrEngine
from winrt.windows.globalization import Language
from winrt.windows.graphics.imaging import BitmapDecoder
from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject

import fitz  # PyMuPDF


# ----------------------------
# Settings
# ----------------------------
MAX_OUTPUT_BYTES = 5 * 1024 * 1024  # 5MB (only for individual doc compression, optional)


# ----------------------------
# Globals (app state)
# ----------------------------
csv_file = ""

# Shared inputs
cover_letter_file = ""
payer_letter_files = []  # OPTIONAL (multiple PDFs supported)

# Mandatory docs
aor_file = ""  # AOR - Consent Form
base_files = []  # Must include TX/DX/Progress (PDFs)
target_list_file = ""  # PDF (mandatory)
dtt_files = []  # multiple PDFs (mandatory)
daily_behavior_data_files = []  # multiple PDFs (mandatory)
daily_trial_counts_files = []  # multiple PDFs (mandatory)
behavior_reduction_files = []  # PDFs (mandatory, multiple supported)

# SOAP
filter_file = ""
output_dir = ""

cancel_operation = False
processing_thread = None

# Timer globals
start_time = None
timer_job = None

# UI globals (assigned in build_ui)
root = None
log_text = None

csv_label = None
cover_label = None
payer_letter_label = None
aor_label = None
base_label = None
target_list_label = None
dtt_label = None
daily_behavior_data_label = None
daily_trial_counts_label = None
behavior_reduction_label = None
filter_label = None
output_label = None

progress_bar = None
progress_label = None
elapsed_label = None
eta_label = None

generate_button = None
generate_bulk_button = None
cancel_button = None
reset_button = None


# ----------------------------
# Utility: UI logging
# ----------------------------
def ui_log(msg: str):
    """Log to console and to the GUI log window (if present)."""
    print(msg)
    global log_text
    if log_text is not None:
        log_text.configure(state="normal")
        log_text.insert("end", msg + "\n")
        log_text.see("end")
        log_text.configure(state="disabled")


def resource_path(relative_path: str) -> str:
    """Absolute path to resource (works in dev + PyInstaller onefile)."""
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, relative_path)


def set_windows_appusermodelid(appid: str):
    """Helps taskbar icon consistency on Windows."""
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appid)
    except Exception:
        pass


# ----------------------------
# Windows OCR (WinRT) for scanned payer letters
# ----------------------------
PAYER_OCR_REGION_REL = (0.60, 0.12, 0.98, 0.45)  # (x0, y0, x1, y1) each in [0..1]
CLAIM_ID_REGEX = re.compile(r"\bCLAIM\s*#\s*[:\-]?\s*([0-9]{8,20})\b", re.IGNORECASE)


def _render_payer_region_image(pdf_path: str, dpi_scale: float = 3.0) -> Image.Image:
    """
    Render first page and crop to the claim# region for OCR, returning a PIL Image.
    """
    doc = fitz.open(pdf_path)
    try:
        if doc.page_count == 0:
            raise ValueError("PDF has no pages")

        page = doc.load_page(0)
        rect = page.rect

        x0r, y0r, x1r, y1r = PAYER_OCR_REGION_REL
        clip = fitz.Rect(
            rect.x0 + rect.width * x0r,
            rect.y0 + rect.height * y0r,
            rect.x0 + rect.width * x1r,
            rect.y0 + rect.height * y1r,
        )

        mat = fitz.Matrix(dpi_scale, dpi_scale)  # ~216 DPI when scale = 3
        pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        return img
    finally:
        doc.close()


async def _winrt_ocr_image_async(pil_img: Image.Image) -> str:
    """
    Use Windows OCR to extract text from the PIL image. Returns extracted text.
    """
    import io

    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    png_bytes = buf.getvalue()

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(png_bytes)
    await writer.store_async()
    await writer.flush_async()
    writer.detach_stream()
    stream.seek(0)

    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()

    engine = OcrEngine.try_create_from_language(Language("en-US"))
    if engine is None:
        raise RuntimeError("Windows OCR not available or OCR Engine is None")

    result = await engine.recognize_async(bitmap)
    return (result.text or "").strip()


def ocr_claim_id_from_payer_letter(pdf_path: str):
    """
    Returns the claim_id extracted from the payer letter (page 1, claim block), or None.
    """
    try:
        img = _render_payer_region_image(pdf_path, dpi_scale=3.0)
    except Exception as e:
        ui_log(f"[WARN] Could not render payer letter for OCR '{pdf_path}': {e}")
        return None

    try:
        import asyncio

        text = asyncio.run(_winrt_ocr_image_async(img))
    except Exception as e:
        ui_log(f"[WARN] OCR failed for payer letter '{pdf_path}': {e}")
        return None

    m = CLAIM_ID_REGEX.search(text)
    if not m:
        return None
    return m.group(1).strip()


def build_payer_letter_index(payer_paths: list[str]) -> dict[str, list[str]]:
    """
    Build an index of claim_id -> list of payer letter paths that match that claim_id based on OCR.
    """
    index: dict[str, list[str]] = {}
    if not payer_paths:
        return index

    ui_log("[INFO] OCR indexing payer letter(s) by Claim ID (Windows OCR)...")
    for p in payer_paths:
        cid = ocr_claim_id_from_payer_letter(p)
        ui_log(f"[INFO] Payer OCR: {os.path.basename(p)} -> claim_id={cid or 'NONE'}")
        if cid:
            index.setdefault(cid, []).append(p)

    ui_log(f"[INFO] Payer OCR index ready. claim_id keys: {len(index)}")
    return index


# ----------------------------
# Date formatting helpers (MM.DD.YY)
# ----------------------------
def fmt_mmddyy_from_date(d: datetime.date) -> str:
    return f"{d.month:02d}.{d.day:02d}.{str(d.year)[-2:]}"


def fmt_mmddyy_from_any(dos_str: str) -> str:
    d = parse_dos_date(dos_str)
    if d:
        return fmt_mmddyy_from_date(d)
    s = str(dos_str or "").strip()
    s = s.replace("/", ".").replace("-", ".")
    return s


def today_mmddyy() -> str:
    now = datetime.now().date()
    return fmt_mmddyy_from_date(now)


# ----------------------------
# File size / compression helpers
# ----------------------------
def file_size_bytes(path: str) -> int:
    try:
        return os.path.getsize(path)
    except OSError:
        return 0


def compress_pdf_rasterized(input_pdf: str, output_pdf: str, dpi: int, jpg_quality: int = 75):
    src = fitz.open(input_pdf)
    out = fitz.open()

    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    try:
        for page in src:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = pix.tobytes("jpeg", jpg_quality=jpg_quality)

            new_page = out.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=img)

        out.save(output_pdf, deflate=True, garbage=4, clean=True)
    finally:
        out.close()
        src.close()


def compress_pdf_under_limit(input_pdf: str, max_bytes: int = MAX_OUTPUT_BYTES) -> str:
    size0 = file_size_bytes(input_pdf)
    if size0 <= max_bytes:
        return input_pdf

    ui_log(f"[INFO] Document over limit ({size0 / (1024 * 1024):.2f} MB). Compressing this document only...")

    attempts = [
        (180, 80),
        (160, 78),
        (140, 75),
        (120, 72),
        (110, 70),
        (100, 68),
    ]

    tmp_fd, tmp_out = tempfile.mkstemp(suffix=".pdf")
    os.close(tmp_fd)

    for dpi, q in attempts:
        try:
            compress_pdf_rasterized(input_pdf, tmp_out, dpi=dpi, jpg_quality=q)
            sz = file_size_bytes(tmp_out)
            ui_log(f"[INFO] Compression try dpi={dpi}, q={q} -> {sz / (1024 * 1024):.2f} MB")
            if sz <= max_bytes:
                ui_log("[OK] Individual document compressed under limit.")
                return tmp_out
        except Exception as e:
            ui_log(f"[WARN] Compression attempt failed (dpi={dpi}, q={q}): {e}")

    ui_log("[WARN] Could not compress under limit; using best-effort compressed output.")
    return tmp_out


# ----------------------------
# CSV helpers
# ----------------------------
def detect_csv_delimiter(path: str) -> str:
    with open(path, "r", encoding="utf-8", newline="") as f:
        sample_line = f.readline()
    return ";" if (";" in sample_line and sample_line.count(";") > sample_line.count(",")) else ","


def get_row_value_case_insensitive(row: dict, keys: list[str]) -> str:
    if not row:
        return ""
    for k in keys:
        if k in row and row.get(k) is not None:
            return str(row.get(k)).strip()
    lower_map = {str(k).lower(): k for k in row.keys()}
    for k in keys:
        lk = str(k).lower()
        if lk in lower_map:
            orig = lower_map[lk]
            val = row.get(orig)
            if val is not None:
                return str(val).strip()
    return ""


def extract_claim_id(row: dict) -> str:
    return get_row_value_case_insensitive(row, ["claim_id", "ClaimID", "Claim ID", "Claim #", "claim#", "claim_number"])


def extract_dos(row: dict) -> str:
    return get_row_value_case_insensitive(row, ["DOS", "dos", "Date of Service", "Service Date"])


def extract_patient_name(row: dict) -> str:
    return get_row_value_case_insensitive(row, ["patient_name", "Patient Name", "patient", "name"])


def extract_billed_amount(row: dict) -> str:
    return get_row_value_case_insensitive(row, ["billed_amount", "Billed Amount", "Billed", "Amount Billed", "amount"])


# ----------------------------
# Document role detection
# ----------------------------
def normalize_name_for_role(path: str) -> str:
    name = os.path.basename(path).lower()
    name = re.sub(r"\.pdf$", "", name)
    name = re.sub(r"[^a-z0-9]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def doc_role_from_filename(path: str) -> str:
    n = normalize_name_for_role(path)
    tokens = set(n.split())

    if "progress" in tokens or "prog" in tokens or ("progressreport" in n.replace(" ", "")):
        return "PROGRESS"

    if "dx" in tokens or n.startswith("dx "):
        return "DX"
    if "diagnosis" in tokens or "diag" in tokens:
        return "DX"

    if "tx" in tokens or n.startswith("tx "):
        return "TX"
    if "treatment" in tokens or "treat" in tokens:
        return "TX"

    return "OTHER"


def order_base_files_fixed(base_files_list):
    priority = {"TX": 0, "DX": 1, "PROGRESS": 2, "OTHER": 99}

    def key_func(path):
        role = doc_role_from_filename(path)
        return (priority.get(role, 99), os.path.basename(path).lower())

    return sorted(list(base_files_list), key=key_func)


def label_for_role(role: str, other_index: int | None = None) -> str:
    if role == "DX":
        return "Diagnosis Report"
    if role == "TX":
        return "Treatment plan"
    if role == "PROGRESS":
        return "Progress Notes"
    if other_index is None:
        return "Additional Document"
    return f"Additional Document {other_index}"


# ----------------------------
# TOC PDF generation
# ----------------------------
def create_toc_pdf(output_path, header_left, header_right, sections, page_size="letter"):
    doc = fitz.open()

    if page_size == "letter":
        width, height = 612, 792
    else:
        width, height = 595, 842

    page = doc.new_page(width=width, height=height)

    margin_x = 54
    y = 72

    page.insert_text((margin_x, y), "Content Table", fontsize=18, fontname="helv")
    y += 28

    page.insert_text((margin_x, y), str(header_left), fontsize=11, fontname="helv")
    y += 16
    page.insert_text((margin_x, y), str(header_right), fontsize=11, fontname="helv")
    y += 24

    page.insert_text((margin_x, y), "Document", fontsize=11, fontname="helv")
    page.insert_text((width - margin_x - 120, y), "Page", fontsize=11, fontname="helv")
    y += 10
    page.draw_line((margin_x, y), (width - margin_x, y), color=(0, 0, 0), width=1)
    y += 18

    for s in sections:
        label = s["label"]
        start_page = s.get("start_page")
        pages = s.get("pages", 0)
        page_str = "N/A" if (start_page is None or pages == 0) else str(start_page)

        page.insert_text((margin_x, y), label, fontsize=11, fontname="helv")
        page.insert_text((width - margin_x - 20, y), page_str, fontsize=11, fontname="helv")
        y += 18

    doc.save(output_path, deflate=True)
    doc.close()


# ----------------------------
# Cover letter: stamp fields + remove widgets (flatten)
# ----------------------------
def flatten_pdf_remove_widgets(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        annots = page.get("/Annots")
        if annots:
            kept = []
            for a in annots:
                aobj = a.get_object()
                if aobj.get("/Subtype") == "/Widget":
                    continue
                kept.append(a)
            if kept:
                page[NameObject("/Annots")] = kept
            else:
                page.pop(NameObject("/Annots"), None)

        writer.add_page(page)

    writer._root_object.pop(NameObject("/AcroForm"), None)

    with open(output_pdf, "wb") as f:
        writer.write(f)


def get_widget_positions(pdf_template_path):
    doc = fitz.open(pdf_template_path)
    positions = {}

    for page_index, page in enumerate(doc):
        widgets = page.widgets()
        if not widgets:
            continue
        for w in widgets:
            try:
                name = w.field_name
                if name and name not in positions:
                    positions[name] = (page_index, w.rect)
            except Exception:
                pass

    doc.close()
    return positions


def stamp_text_on_pdf(template_pdf, output_pdf, values, positions, font_size=10):
    doc = fitz.open(template_pdf)

    for field, value in values.items():
        if field not in positions:
            ui_log(f"[WARN] Field '{field}' not found in template widgets")
            continue

        page_index, rect = positions[field]
        page = doc[page_index]

        pad_x, pad_y = 1.5, 1.5
        r = fitz.Rect(rect.x0 + pad_x, rect.y0 + pad_y, rect.x1 - pad_x, rect.y1 - pad_y)

        page.insert_textbox(
            r,
            str(value),
            fontsize=font_size,
            fontname="helv",
            color=(0, 0, 0),
            align=0,
        )

    doc.save(output_pdf, deflate=True)
    doc.close()


def fill_and_flatten_cover(pdf_path, data_dict, output_path):
    temp1 = output_path + ".__stamped__.pdf"

    values = {
        "DOS": str(data_dict.get("DOS", "")),
        "claim_id": str(data_dict.get("claim_id", "")),
    }

    positions = get_widget_positions(pdf_path)
    stamp_text_on_pdf(pdf_path, temp1, values, positions, font_size=10)
    flatten_pdf_remove_widgets(temp1, output_path)

    try:
        os.remove(temp1)
    except OSError:
        pass

    return output_path


# ----------------------------
# SOAP indexing + CPT filtering
# ----------------------------
def parse_dos_date(dos_str):
    if not dos_str:
        return None

    s = str(dos_str).strip()
    m = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", s)
    if m:
        s = m.group(1)

    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    return None


def dos_to_soap_key(d):
    return f"{d.month:02d}-{d.day:02d}-{d.year:04d}"


def build_soap_index(soap_pdf_path):
    index = {}
    pattern = re.compile(
        r"\bSESSION\s*DATE\b\s*[:\-]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b",
        re.IGNORECASE,
    )

    doc = fitz.open(soap_pdf_path)
    try:
        for i in range(doc.page_count):
            page = doc.load_page(i)
            text = page.get_text("text") or ""
            if not text:
                continue

            t = " ".join(text.split())
            m = pattern.search(t)
            if not m:
                continue

            raw_date = m.group(1)
            d = parse_dos_date(raw_date)
            if not d:
                continue

            key = dos_to_soap_key(d)
            index.setdefault(key, []).append(i)
    finally:
        doc.close()

    return index


def normalize_cpt(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    digits = re.findall(r"\d+", s)
    return digits[0] if digits else ""


def extract_cpt_from_row(row: dict) -> str:
    return normalize_cpt(row.get("CPT"))


def soap_page_matches_any_cpt(soap_page_text: str, cpts: list[str]) -> bool:
    if not soap_page_text or not cpts:
        return False

    t = " ".join(soap_page_text.split())
    if "CPT CODE" not in t.upper():
        return False

    for c in cpts:
        if re.search(rf"\b{re.escape(c)}\b", t):
            return True
    return False


def match_soap_pages_by_dos_and_cpt(soap_pdf_path: str, dos_str: str, cpts: list[str], soap_index: dict) -> list[int]:
    d = parse_dos_date(dos_str)
    if not d:
        return []

    key = dos_to_soap_key(d)
    candidates = soap_index.get(key, [])
    if not candidates:
        return []

    if not cpts:
        return candidates

    doc = fitz.open(soap_pdf_path)
    try:
        matched = []
        for i in candidates:
            page = doc.load_page(i)
            text = page.get_text("text") or ""
            if soap_page_matches_any_cpt(text, cpts):
                matched.append(i)

        return matched if matched else candidates
    finally:
        doc.close()


# ----------------------------
# Helper: load optional single/multi PDFs
# ----------------------------
def load_optional_pdf(path: str, label: str, temp_paths_to_cleanup: list[str]):
    if not path:
        ui_log(f"[INFO] {label} not selected (optional).")
        return None

    try:
        compressed_path = compress_pdf_under_limit(path, max_bytes=MAX_OUTPUT_BYTES)
        if compressed_path != path:
            temp_paths_to_cleanup.append(compressed_path)

        r = PdfReader(compressed_path)
        ui_log(f"[INFO] Loaded {label} pages={len(r.pages)}" + (" (compressed)" if compressed_path != path else ""))
        return (path, r, len(r.pages))
    except Exception as e:
        ui_log(f"[WARN] Could not read {label} file '{path}': {e}")
        return None


def load_optional_pdfs(paths: list[str], label: str, temp_paths_to_cleanup: list[str]):
    loaded = []
    if not paths:
        ui_log(f"[INFO] {label} not selected (optional).")
        return loaded

    for p in paths:
        try:
            compressed_path = compress_pdf_under_limit(p, max_bytes=MAX_OUTPUT_BYTES)
            if compressed_path != p:
                temp_paths_to_cleanup.append(compressed_path)

            r = PdfReader(compressed_path)
            ui_log(
                f"[INFO] Loaded {label}: {os.path.basename(p)} pages={len(r.pages)}"
                + (" (compressed)" if compressed_path != p else "")
            )
            loaded.append((p, r, len(r.pages)))
        except Exception as e:
            ui_log(f"[WARN] Could not read {label} file '{p}': {e}")

    return loaded


# ----------------------------
# Claim List PDF (bulk mode)
# ----------------------------
def create_claim_list_pdf(output_path: str, patient_name: str, claims: list[dict], page_size="letter"):
    doc = fitz.open()
    if page_size == "letter":
        width, height = 612, 792
    else:
        width, height = 595, 842

    margin_x = 54
    y0 = 72
    line_h = 14

    page = doc.new_page(width=width, height=height)
    y = y0

    title = "Claim List"
    subtitle = f"Patient: {patient_name}" if patient_name else "Patient: (not provided)"
    page.insert_text((margin_x, y), title, fontsize=16, fontname="helv")
    y += 22
    page.insert_text((margin_x, y), subtitle, fontsize=10, fontname="helv")
    y += 18

    page.insert_text((margin_x, y), "Claim ID", fontsize=10, fontname="helv")
    page.insert_text((margin_x + 170, y), "DOS", fontsize=10, fontname="helv")
    page.insert_text((margin_x + 270, y), "Billed Amount", fontsize=10, fontname="helv")
    y += 8
    page.draw_line((margin_x, y), (width - margin_x, y), color=(0, 0, 0), width=1)
    y += 14

    for c in claims:
        if y > height - 72:
            page = doc.new_page(width=width, height=height)
            y = y0

        claim_id = str(c.get("claim_id") or "")
        dos = str(c.get("dos") or "")
        billed = str(c.get("billed_amount") or "")

        page.insert_text((margin_x, y), claim_id, fontsize=9, fontname="helv")
        page.insert_text((margin_x + 170, y), dos, fontsize=9, fontname="helv")
        page.insert_text((margin_x + 270, y), billed, fontsize=9, fontname="helv")
        y += line_h

    doc.save(output_path, deflate=True)
    doc.close()


# ----------------------------
# Build PDF per claim (per-claim mode)
# ----------------------------
def create_pdf_for_claim_v2(
    claim_id: str,
    representative_row: dict,
    dos_str: str,
    cpts: list[str],
    cover_letter_path: str,
    payer_letter_cached_list,
    payer_letter_index: dict,
    aor_cached,
    base_docs_cached,
    soap_reader,
    soap_pdf_path: str,
    soap_index: dict,
    target_list_cached,
    dtt_cached_list,
    daily_behavior_data_cached_list,
    daily_trial_counts_cached_list,
    behavior_reduction_cached_list,
    output_directory: str,
    progress_callback=None,
):
    global cancel_operation

    if cancel_operation:
        return

    ui_log(f">>> Creating PDF for claim_id={claim_id}, DOS={dos_str}, CPTs={cpts}")

    # 1) Cover letter -> per-claim: try fill
    temp_fd, temp_filled_path = tempfile.mkstemp(suffix=".pdf")
    os.close(temp_fd)

    writer = PdfWriter()

    try:
        fill_and_flatten_cover(cover_letter_path, representative_row, temp_filled_path)
        ui_log(f"[OK] Cover letter filled for claim {claim_id}")

        filled_reader = PdfReader(temp_filled_path)
        for page in filled_reader.pages:
            writer.add_page(page)

    except Exception as e:
        ui_log(f"[WARN] Could not fill cover letter: {e}. Using template.")
        reader = PdfReader(cover_letter_path)
        for page in reader.pages:
            writer.add_page(page)
    finally:
        try:
            os.remove(temp_filled_path)
        except OSError:
            pass

    if cancel_operation:
        return

    # 2) Pre-compute SOAP matched pages (needed for TOC pages count)
    matched_pages = match_soap_pages_by_dos_and_cpt(
        soap_pdf_path=soap_pdf_path,
        dos_str=dos_str,
        cpts=cpts,
        soap_index=soap_index,
    )
    ui_log(f"[OK] SOAP matched {len(matched_pages)} page(s) for DOS={dos_str} with CPT filter.")

    # 3) Build TOC sections with updated order
    cover_pages = len(writer.pages)
    toc_pages = 1
    next_page = cover_pages + toc_pages + 1

    sections = []

    # Payer letters matched by OCR claim_id
    matched_payer_paths = []
    if payer_letter_index:
        matched_payer_paths = payer_letter_index.get(str(claim_id).strip(), []) or []

    matched_payer_cached = []
    if payer_letter_cached_list and matched_payer_paths:
        want = set(matched_payer_paths)
        matched_payer_cached = [t for t in payer_letter_cached_list if t[0] in want]

    payer_pages = sum(n for (_p, _r, n) in (matched_payer_cached or []))
    if payer_pages:
        sections.append({"label": "Payer Letter", "start_page": next_page, "pages": payer_pages})
        next_page += payer_pages

    # Split base docs by role to control final order precisely.
    tx_docs = []
    dx_docs = []
    other_docs = []
    prog_docs = []
    for orig_path, reader_obj, n_pages, role in base_docs_cached:
        if role == "TX":
            tx_docs.append((orig_path, reader_obj, n_pages, role))
        elif role == "DX":
            dx_docs.append((orig_path, reader_obj, n_pages, role))
        elif role == "PROGRESS":
            prog_docs.append((orig_path, reader_obj, n_pages, role))
        else:
            other_docs.append((orig_path, reader_obj, n_pages, role))

    # TX
    for _orig_path, _reader_obj, n_pages, role in tx_docs:
        label = label_for_role(role)
        sections.append({"label": label, "start_page": next_page if n_pages else None, "pages": n_pages})
        next_page += n_pages

    # DX
    for _orig_path, _reader_obj, n_pages, role in dx_docs:
        label = label_for_role(role)
        sections.append({"label": label, "start_page": next_page if n_pages else None, "pages": n_pages})
        next_page += n_pages

    # AOR - Consent Form
    if aor_cached:
        _, _, n = aor_cached
        sections.append({"label": "AOR - Consent Form", "start_page": next_page if n else None, "pages": n})
        next_page += n

    # SOAP Notes
    sections.append({"label": "SOAP Notes", "start_page": next_page if len(matched_pages) else None, "pages": len(matched_pages)})
    next_page += len(matched_pages)

    # Progress Notes
    for _orig_path, _reader_obj, n_pages, role in prog_docs:
        label = label_for_role(role)
        sections.append({"label": label, "start_page": next_page if n_pages else None, "pages": n_pages})
        next_page += n_pages

    # Target List
    if target_list_cached:
        _, _, n = target_list_cached
        sections.append({"label": "Target List Report", "start_page": next_page if n else None, "pages": n})
        next_page += n

    # DTT (multiple PDFs)
    dtt_pages = sum(n for (_p, _r, n) in (dtt_cached_list or []))
    sections.append({"label": "DTT Trial Sheets", "start_page": next_page if dtt_pages else None, "pages": dtt_pages})
    next_page += dtt_pages

    # Daily Behavior Data (multiple PDFs)
    dbd_pages = sum(n for (_p, _r, n) in (daily_behavior_data_cached_list or []))
    sections.append({"label": "Daily Behavior Data", "start_page": next_page if dbd_pages else None, "pages": dbd_pages})
    next_page += dbd_pages

    # Daily Trial Counts (multiple PDFs)
    dtc_pages = sum(n for (_p, _r, n) in (daily_trial_counts_cached_list or []))
    sections.append({"label": "Daily Trial Counts", "start_page": next_page if dtc_pages else None, "pages": dtc_pages})
    next_page += dtc_pages

    # Behavior Reduction (multiple PDFs)
    br_pages = sum(n for (_p, _r, n) in (behavior_reduction_cached_list or []))
    sections.append({"label": "Behavior Reduction Report", "start_page": next_page if br_pages else None, "pages": br_pages})
    next_page += br_pages

    # Any other docs from base selection - keep last
    other_counter = 0
    for _orig_path, _reader_obj, n_pages, role in other_docs:
        other_counter += 1
        label = label_for_role("OTHER", other_index=other_counter)
        sections.append({"label": label, "start_page": next_page if n_pages else None, "pages": n_pages})
        next_page += n_pages

    # 4) Insert TOC
    toc_fd, toc_path = tempfile.mkstemp(suffix=".pdf")
    os.close(toc_fd)
    try:
        create_toc_pdf(
            output_path=toc_path,
            header_left=f"Claim ID: {claim_id}",
            header_right=f"DOS: {fmt_mmddyy_from_any(dos_str)}",
            sections=sections,
            page_size="letter",
        )
        toc_reader = PdfReader(toc_path)
        for p in toc_reader.pages:
            writer.add_page(p)
    finally:
        try:
            os.remove(toc_path)
        except OSError:
            pass

    # 5) Append docs in exact same order as TOC

    # Matched payer letters (per-claim)
    for _p, r, _n in (matched_payer_cached or []):
        for page in r.pages:
            writer.add_page(page)

    for _orig_path, r, _n, _role in tx_docs:
        for page in r.pages:
            writer.add_page(page)

    for _orig_path, r, _n, _role in dx_docs:
        for page in r.pages:
            writer.add_page(page)

    if aor_cached:
        _, r, _n = aor_cached
        for page in r.pages:
            writer.add_page(page)

    if soap_reader and matched_pages:
        for i in matched_pages:
            if cancel_operation:
                return
            writer.add_page(soap_reader.pages[i])

    for _orig_path, r, _n, _role in prog_docs:
        for page in r.pages:
            writer.add_page(page)

    if target_list_cached:
        _, r, _n = target_list_cached
        for page in r.pages:
            writer.add_page(page)

    for _p, r, _n in (dtt_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    for _p, r, _n in (daily_behavior_data_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    for _p, r, _n in (daily_trial_counts_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    for _p, r, _n in (behavior_reduction_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    for _orig_path, r, _n, _role in other_docs:
        for page in r.pages:
            writer.add_page(page)

    # 6) Write output with new naming
    dos_mmddyy = fmt_mmddyy_from_any(dos_str)
    today = today_mmddyy()
    out_name = f"{claim_id}-{dos_mmddyy}-sent-{today}.pdf"
    output_path = os.path.join(output_directory, out_name)

    with open(output_path, "wb") as f:
        writer.write(f)

    ui_log(f"[OK] Final PDF saved to: {output_path}")

    if progress_callback and not cancel_operation:
        progress_callback()


# ----------------------------
# Bulk PDF generation (single output for ALL claims)
# ----------------------------
def create_pdf_bulk_v2(
    patient_name: str,
    claims_order: list[dict],
    cover_letter_path: str,
    payer_letter_cached_list,
    claim_list_pdf_cached,
    aor_cached,
    base_docs_cached,
    soap_reader,
    soap_pdf_path: str,
    soap_index: dict,
    target_list_cached,
    dtt_cached_list,
    daily_behavior_data_cached_list,
    daily_trial_counts_cached_list,
    behavior_reduction_cached_list,
    output_directory: str,
):
    global cancel_operation
    if cancel_operation:
        return

    writer = PdfWriter()

    # 1) Cover letter (bulk mode: do NOT fill; use as-is)
    ui_log("[INFO] Bulk mode: using cover letter as-is (no filling).")
    r = PdfReader(cover_letter_path)
    for page in r.pages:
        writer.add_page(page)

    # Split base docs
    tx_docs = []
    dx_docs = []
    other_docs = []
    prog_docs = []
    for orig_path, reader_obj, n_pages, role in base_docs_cached:
        if role == "TX":
            tx_docs.append((orig_path, reader_obj, n_pages, role))
        elif role == "DX":
            dx_docs.append((orig_path, reader_obj, n_pages, role))
        elif role == "PROGRESS":
            prog_docs.append((orig_path, reader_obj, n_pages, role))
        else:
            other_docs.append((orig_path, reader_obj, n_pages, role))

    # Precompute SOAP pages total for TOC and later append
    soap_pages_by_claim = []
    total_soap_pages = 0
    for c in claims_order:
        if cancel_operation:
            return
        dos_str = c.get("dos_raw") or c.get("dos") or ""
        cpts = c.get("cpts") or []
        claim_id = c.get("claim_id") or ""

        matched = match_soap_pages_by_dos_and_cpt(
            soap_pdf_path=soap_pdf_path,
            dos_str=dos_str,
            cpts=cpts,
            soap_index=soap_index,
        )
        soap_pages_by_claim.append((claim_id, dos_str, matched))
        total_soap_pages += len(matched)

    # 2) Build TOC sections order for bulk
    cover_pages = len(writer.pages)
    toc_pages = 1
    next_page = cover_pages + toc_pages + 1
    sections = []

    # Payer letters in bulk: include ALL once (optional)
    payer_pages = sum(n for (_p, _r, n) in (payer_letter_cached_list or []))
    if payer_pages:
        sections.append({"label": "Payer Letters", "start_page": next_page, "pages": payer_pages})
        next_page += payer_pages

    if claim_list_pdf_cached:
        _, _, n = claim_list_pdf_cached
        sections.append({"label": "Claim List", "start_page": next_page if n else None, "pages": n})
        next_page += n

    for _orig_path, _r, n_pages, role in tx_docs:
        sections.append({"label": label_for_role(role), "start_page": next_page if n_pages else None, "pages": n_pages})
        next_page += n_pages

    for _orig_path, _r, n_pages, role in dx_docs:
        sections.append({"label": label_for_role(role), "start_page": next_page if n_pages else None, "pages": n_pages})
        next_page += n_pages

    if aor_cached:
        _, _, n = aor_cached
        sections.append({"label": "AOR - Consent Form", "start_page": next_page if n else None, "pages": n})
        next_page += n

    sections.append({"label": "SOAP Notes", "start_page": next_page if total_soap_pages else None, "pages": total_soap_pages})
    next_page += total_soap_pages

    for _orig_path, _r, n_pages, role in prog_docs:
        sections.append({"label": label_for_role(role), "start_page": next_page if n_pages else None, "pages": n_pages})
        next_page += n_pages

    if target_list_cached:
        _, _, n = target_list_cached
        sections.append({"label": "Target List Report", "start_page": next_page if n else None, "pages": n})
        next_page += n

    dtt_pages = sum(n for (_p, _r, n) in (dtt_cached_list or []))
    sections.append({"label": "DTT Trial Sheets", "start_page": next_page if dtt_pages else None, "pages": dtt_pages})
    next_page += dtt_pages

    dbd_pages = sum(n for (_p, _r, n) in (daily_behavior_data_cached_list or []))
    sections.append({"label": "Daily Behavior Data", "start_page": next_page if dbd_pages else None, "pages": dbd_pages})
    next_page += dbd_pages

    dtc_pages = sum(n for (_p, _r, n) in (daily_trial_counts_cached_list or []))
    sections.append({"label": "Daily Trial Counts", "start_page": next_page if dtc_pages else None, "pages": dtc_pages})
    next_page += dtc_pages

    br_pages = sum(n for (_p, _r, n) in (behavior_reduction_cached_list or []))
    sections.append({"label": "Behavior Reduction Report", "start_page": next_page if br_pages else None, "pages": br_pages})
    next_page += br_pages

    other_counter = 0
    for _orig_path, _r, n_pages, _role in other_docs:
        other_counter += 1
        sections.append(
            {"label": label_for_role("OTHER", other_index=other_counter), "start_page": next_page if n_pages else None, "pages": n_pages}
        )
        next_page += n_pages

    dos_dates = []
    for c in claims_order:
        d = parse_dos_date(c.get("dos_raw") or c.get("dos") or "")
        if d:
            dos_dates.append(d)

    if dos_dates:
        initial_d = min(dos_dates)
        final_d = max(dos_dates)
        initial_s = fmt_mmddyy_from_date(initial_d)
        final_s = fmt_mmddyy_from_date(final_d)
    else:
        initial_s = ""
        final_s = ""

    toc_fd, toc_path = tempfile.mkstemp(suffix=".pdf")
    os.close(toc_fd)
    try:
        create_toc_pdf(
            output_path=toc_path,
            header_left=f"Claim IDs: {len(claims_order)}",
            header_right=f"DOS Range: {initial_s} - {final_s}",
            sections=sections,
            page_size="letter",
        )
        toc_reader = PdfReader(toc_path)
        for p in toc_reader.pages:
            writer.add_page(p)
    finally:
        try:
            os.remove(toc_path)
        except OSError:
            pass

    # Append docs in order

    for _p, r, _n in (payer_letter_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    if claim_list_pdf_cached:
        _, r, _n = claim_list_pdf_cached
        for page in r.pages:
            writer.add_page(page)

    for _orig_path, r, _n, _role in tx_docs:
        for page in r.pages:
            writer.add_page(page)

    for _orig_path, r, _n, _role in dx_docs:
        for page in r.pages:
            writer.add_page(page)

    if aor_cached:
        _, r, _n = aor_cached
        for page in r.pages:
            writer.add_page(page)

    if soap_reader:
        for _claim_id, _dos_str, matched in soap_pages_by_claim:
            if cancel_operation:
                return
            for i in matched:
                if cancel_operation:
                    return
                writer.add_page(soap_reader.pages[i])

    for _orig_path, r, _n, _role in prog_docs:
        for page in r.pages:
            writer.add_page(page)

    if target_list_cached:
        _, r, _n = target_list_cached
        for page in r.pages:
            writer.add_page(page)

    for _p, r, _n in (dtt_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    for _p, r, _n in (daily_behavior_data_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    for _p, r, _n in (daily_trial_counts_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    for _p, r, _n in (behavior_reduction_cached_list or []):
        for page in r.pages:
            writer.add_page(page)

    for _orig_path, r, _n, _role in other_docs:
        for page in r.pages:
            writer.add_page(page)

    patient = (patient_name or "patient").strip()
    patient = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", patient)
    patient = re.sub(r"\s+", " ", patient).strip().replace(" ", "_")

    today = today_mmddyy()
    out_name = f"{patient}-{initial_s}-{final_s}-sent-{today}.pdf"
    output_path = os.path.join(output_directory, out_name)

    with open(output_path, "wb") as f:
        writer.write(f)

    ui_log(f"[OK] Bulk PDF saved to: {output_path}")


# ----------------------------
# Timer helpers
# ----------------------------
def format_seconds(seconds: float) -> str:
    seconds = max(0, int(seconds))
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h:02d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


def update_timer(total_items: int):
    global timer_job, start_time
    if start_time is None:
        return

    elapsed = time.perf_counter() - start_time
    elapsed_label.config(text=f"Time: {format_seconds(elapsed)}")

    done = int(progress_bar["value"])
    if done > 0:
        rate = elapsed / done
        remaining = (total_items - done) * rate
        eta_label.config(text=f"Remaining (est.): {format_seconds(remaining)}")
    else:
        eta_label.config(text="Remaining (est.): --:--")

    timer_job = root.after(500, update_timer, total_items)


def stop_timer_and_finalize():
    global timer_job, start_time
    if timer_job:
        root.after_cancel(timer_job)
        timer_job = None

    if start_time is not None:
        final_elapsed = time.perf_counter() - start_time
        elapsed_label.config(text=f"Time: {format_seconds(final_elapsed)}")
    start_time = None


# ----------------------------
# App actions (validation)
# ----------------------------
def validate_common_inputs(require_output=True) -> bool:
    if not csv_file:
        messagebox.showerror("Error", "Please select the Claims List (CSV).")
        return False
    if not cover_letter_file:
        messagebox.showerror("Error", "Please select the Cover Letter (PDF).")
        return False
    if not aor_file:
        messagebox.showerror("Error", "Please select the AOR - Consent Form (PDF). This file is mandatory.")
        return False
    if not base_files:
        messagebox.showerror("Error", "Please select the Base Docs (TX / DX / Progress) PDFs.")
        return False
    if not filter_file:
        messagebox.showerror("Error", "Please select the SOAP Notes PDF.")
        return False

    # Mandatory additional docs (all mandatory except payer letter)
    if not target_list_file:
        messagebox.showerror("Error", "Please select the Target List Report (PDF). This file is mandatory.")
        return False
    if not dtt_files:
        messagebox.showerror("Error", "Please select DTT Trial Sheets (PDFs). These files are mandatory.")
        return False
    if not daily_behavior_data_files:
        messagebox.showerror("Error", "Please select Daily Behavior Data (PDFs). These files are mandatory.")
        return False
    if not daily_trial_counts_files:
        messagebox.showerror("Error", "Please select Daily Trial Counts (PDFs). These files are mandatory.")
        return False
    if not behavior_reduction_files:
        messagebox.showerror("Error", "Please select Behavior Reduction Report PDF(s). These files are mandatory.")
        return False

    if require_output and not output_dir:
        messagebox.showerror("Error", "Please select the Output Folder.")
        return False
    return True


def validate_output_dir() -> bool:
    if not os.path.isdir(output_dir):
        messagebox.showerror("Error", f"Output folder is not valid: {output_dir}")
        return False
    if not os.access(output_dir, os.W_OK):
        messagebox.showerror("Permission Error", f"No write permission for folder:\n{output_dir}")
        return False
    return True


# ----------------------------
# Per-claim processing
# ----------------------------
def process_per_claim():
    global cancel_operation, processing_thread, start_time

    try:
        if not validate_common_inputs(require_output=True):
            return
        if not validate_output_dir():
            return

        os.makedirs(output_dir, exist_ok=True)

        start_time = time.perf_counter()
        cancel_operation = False

        generate_button.config(state="disabled")
        generate_bulk_button.config(state="disabled")
        cancel_button.config(state="normal")

        ui_log("=== Starting per-claim processing ===")
        processing_thread = threading.Thread(target=process_pdfs_per_claim, daemon=True)
        processing_thread.start()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")
        progress_label.config(text="Error")


def process_pdfs_per_claim():
    global cancel_operation

    base_docs_cached = []
    soap_index = {}
    soap_reader = None

    payer_letter_cached_list = []
    payer_letter_index = {}

    aor_cached = None
    target_list_cached = None
    dtt_cached_list = []
    daily_behavior_data_cached_list = []
    daily_trial_counts_cached_list = []
    behavior_reduction_cached_list = []

    temp_paths_to_cleanup = []

    try:
        ui_log("[INFO] Loading base PDFs (one-time)...")
        base_files_ordered = order_base_files_fixed(base_files)
        for fpath in base_files_ordered:
            if cancel_operation:
                return
            role = doc_role_from_filename(fpath)
            compressed_path = compress_pdf_under_limit(fpath, max_bytes=MAX_OUTPUT_BYTES)
            if compressed_path != fpath:
                temp_paths_to_cleanup.append(compressed_path)
            r = PdfReader(compressed_path)
            base_docs_cached.append((fpath, r, len(r.pages), role))
            ui_log(
                f"[INFO] Loaded {os.path.basename(fpath)} role={role} pages={len(r.pages)}"
                + (" (compressed)" if compressed_path != fpath else "")
            )

        # Optional payer letter(s)
        payer_letter_cached_list = load_optional_pdfs(list(payer_letter_files), "Payer Letter", temp_paths_to_cleanup)
        payer_letter_index = build_payer_letter_index(list(payer_letter_files))

        # Mandatory docs
        aor_cached = load_optional_pdf(aor_file, "AOR - Consent Form", temp_paths_to_cleanup)
        if not aor_cached:
            raise RuntimeError("AOR - Consent Form is mandatory but could not be loaded.")

        target_list_cached = load_optional_pdf(target_list_file, "Target List Report", temp_paths_to_cleanup)
        if not target_list_cached:
            raise RuntimeError("Target List Report is mandatory but could not be loaded.")

        dtt_cached_list = load_optional_pdfs(dtt_files, "DTT Trial Sheets", temp_paths_to_cleanup)
        if not dtt_cached_list:
            raise RuntimeError("DTT Trial Sheets are mandatory but could not be loaded.")

        daily_behavior_data_cached_list = load_optional_pdfs(daily_behavior_data_files, "Daily Behavior Data", temp_paths_to_cleanup)
        if not daily_behavior_data_cached_list:
            raise RuntimeError("Daily Behavior Data is mandatory but could not be loaded.")

        daily_trial_counts_cached_list = load_optional_pdfs(daily_trial_counts_files, "Daily Trial Counts", temp_paths_to_cleanup)
        if not daily_trial_counts_cached_list:
            raise RuntimeError("Daily Trial Counts is mandatory but could not be loaded.")

        behavior_reduction_cached_list = load_optional_pdfs(
            list(behavior_reduction_files),
            "Behavior Reduction Report",
            temp_paths_to_cleanup,
        )
        if not behavior_reduction_cached_list:
            raise RuntimeError("Behavior Reduction Report is mandatory but could not be loaded.")

        ui_log("[INFO] Indexing SOAP notes by SESSION DATE (one-time)...")
        soap_index = build_soap_index(filter_file)
        ui_log(f"[INFO] SOAP index ready. Unique dates: {len(soap_index)}")
        soap_reader = PdfReader(filter_file)

        groups = {}
        rows_read = 0
        delim = detect_csv_delimiter(csv_file)
        ui_log(f"Detected CSV delimiter: '{delim}'")

        with open(csv_file, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=delim)
            ui_log(f"CSV Headers: {reader.fieldnames}")

            for row in reader:
                rows_read += 1
                claim_id = extract_claim_id(row)
                if not claim_id:
                    ui_log(f"[WARN] Missing claim_id at row {rows_read}; skipping row.")
                    continue

                dos_raw = extract_dos(row)
                cpt = extract_cpt_from_row(row)

                g = groups.setdefault(claim_id, {"rows": [], "cpts": set(), "dos_set": set()})
                g["rows"].append(row)
                if dos_raw:
                    g["dos_set"].add(str(dos_raw).strip())
                if cpt:
                    g["cpts"].add(cpt)

        total_claims = len(groups)
        ui_log(f"[INFO] Grouped {total_claims} unique claim_id(s) from {rows_read} CSV row(s).")

        progress_bar["maximum"] = max(1, total_claims)
        progress_bar["value"] = 0
        progress_label.config(text="0%")
        update_timer(total_claims)

        created = 0

        for claim_id, g in groups.items():
            if cancel_operation:
                break

            representative_row = g["rows"][0]
            representative_row["claim_id"] = claim_id
            representative_row["DOS"] = extract_dos(representative_row) or ""

            cpts = sorted(g["cpts"])
            dos_list = sorted(g["dos_set"])
            dos_str = dos_list[0] if dos_list else (representative_row.get("DOS") or "")

            if len(dos_list) > 1:
                ui_log(f"[WARN] claim_id={claim_id} has multiple DOS values in CSV: {dos_list}. Using: {dos_str}")

            def update_progress():
                nonlocal created
                created += 1
                progress_bar["value"] = created
                pct = int((created / total_claims) * 100) if total_claims else 100
                progress_label.config(text=f"{pct}%")
                root.update_idletasks()

            create_pdf_for_claim_v2(
                claim_id=claim_id,
                representative_row=representative_row,
                dos_str=dos_str,
                cpts=cpts,
                cover_letter_path=cover_letter_file,
                payer_letter_cached_list=payer_letter_cached_list,
                payer_letter_index=payer_letter_index,
                aor_cached=aor_cached,
                base_docs_cached=base_docs_cached,
                soap_reader=soap_reader,
                soap_pdf_path=filter_file,
                soap_index=soap_index,
                target_list_cached=target_list_cached,
                dtt_cached_list=dtt_cached_list,
                daily_behavior_data_cached_list=daily_behavior_data_cached_list,
                daily_trial_counts_cached_list=daily_trial_counts_cached_list,
                behavior_reduction_cached_list=behavior_reduction_cached_list,
                output_directory=output_dir,
                progress_callback=update_progress,
            )

        if cancel_operation:
            ui_log("=== Cancelled ===")
            messagebox.showinfo("Cancelled", "Operation cancelled.")
            progress_label.config(text="Cancelled")
        else:
            ui_log("=== Completed ===")
            messagebox.showinfo("Success", f"PDFs created successfully!\n\nOutput folder: {output_dir}")
            progress_label.config(text="100%")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during processing:\n{e}")
        progress_label.config(text="Error")

    finally:
        for p in temp_paths_to_cleanup:
            try:
                os.remove(p)
            except OSError:
                pass

        generate_button.config(state="normal")
        generate_bulk_button.config(state="normal")
        cancel_button.config(state="disabled")
        cancel_operation = False
        stop_timer_and_finalize()


# ----------------------------
# Bulk processing
# ----------------------------
def process_bulk():
    global cancel_operation, processing_thread, start_time

    try:
        if not validate_common_inputs(require_output=True):
            return
        if not validate_output_dir():
            return

        os.makedirs(output_dir, exist_ok=True)

        start_time = time.perf_counter()
        cancel_operation = False

        generate_button.config(state="disabled")
        generate_bulk_button.config(state="disabled")
        cancel_button.config(state="normal")

        ui_log("=== Starting BULK processing ===")
        processing_thread = threading.Thread(target=process_pdfs_bulk, daemon=True)
        processing_thread.start()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")
        progress_label.config(text="Error")


def process_pdfs_bulk():
    global cancel_operation

    base_docs_cached = []
    soap_index = {}
    soap_reader = None

    payer_letter_cached_list = []
    aor_cached = None
    target_list_cached = None
    dtt_cached_list = []
    daily_behavior_data_cached_list = []
    daily_trial_counts_cached_list = []
    behavior_reduction_cached_list = []

    claim_list_pdf_cached = None

    temp_paths_to_cleanup = []
    temp_internal_paths = []

    try:
        ui_log("[INFO] Loading base PDFs (one-time)...")
        base_files_ordered = order_base_files_fixed(base_files)
        for fpath in base_files_ordered:
            if cancel_operation:
                return
            role = doc_role_from_filename(fpath)
            compressed_path = compress_pdf_under_limit(fpath, max_bytes=MAX_OUTPUT_BYTES)
            if compressed_path != fpath:
                temp_paths_to_cleanup.append(compressed_path)
            r = PdfReader(compressed_path)
            base_docs_cached.append((fpath, r, len(r.pages), role))
            ui_log(
                f"[INFO] Loaded {os.path.basename(fpath)} role={role} pages={len(r.pages)}"
                + (" (compressed)" if compressed_path != fpath else "")
            )

        # Optional payer letters (bulk includes ALL once)
        payer_letter_cached_list = load_optional_pdfs(list(payer_letter_files), "Payer Letter", temp_paths_to_cleanup)

        aor_cached = load_optional_pdf(aor_file, "AOR - Consent Form", temp_paths_to_cleanup)
        if not aor_cached:
            raise RuntimeError("AOR - Consent Form is mandatory but could not be loaded.")

        target_list_cached = load_optional_pdf(target_list_file, "Target List Report", temp_paths_to_cleanup)
        if not target_list_cached:
            raise RuntimeError("Target List Report is mandatory but could not be loaded.")

        dtt_cached_list = load_optional_pdfs(dtt_files, "DTT Trial Sheets", temp_paths_to_cleanup)
        if not dtt_cached_list:
            raise RuntimeError("DTT Trial Sheets are mandatory but could not be loaded.")

        daily_behavior_data_cached_list = load_optional_pdfs(daily_behavior_data_files, "Daily Behavior Data", temp_paths_to_cleanup)
        if not daily_behavior_data_cached_list:
            raise RuntimeError("Daily Behavior Data is mandatory but could not be loaded.")

        daily_trial_counts_cached_list = load_optional_pdfs(daily_trial_counts_files, "Daily Trial Counts", temp_paths_to_cleanup)
        if not daily_trial_counts_cached_list:
            raise RuntimeError("Daily Trial Counts is mandatory but could not be loaded.")

        behavior_reduction_cached_list = load_optional_pdfs(
            list(behavior_reduction_files),
            "Behavior Reduction Report",
            temp_paths_to_cleanup,
        )
        if not behavior_reduction_cached_list:
            raise RuntimeError("Behavior Reduction Report is mandatory but could not be loaded.")

        ui_log("[INFO] Indexing SOAP notes by SESSION DATE (one-time)...")
        soap_index = build_soap_index(filter_file)
        ui_log(f"[INFO] SOAP index ready. Unique dates: {len(soap_index)}")
        soap_reader = PdfReader(filter_file)

        groups = {}
        rows_read = 0
        delim = detect_csv_delimiter(csv_file)
        ui_log(f"Detected CSV delimiter: '{delim}'")

        claims_flat = []
        patient_name = ""

        with open(csv_file, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=delim)
            ui_log(f"CSV Headers: {reader.fieldnames}")

            for row in reader:
                rows_read += 1
                claim_id = extract_claim_id(row)
                if not claim_id:
                    ui_log(f"[WARN] Missing claim_id at row {rows_read}; skipping row.")
                    continue

                dos_raw = extract_dos(row)
                billed = extract_billed_amount(row)
                p_name = extract_patient_name(row)
                if p_name and not patient_name:
                    patient_name = p_name

                cpt = extract_cpt_from_row(row)

                g = groups.setdefault(
                    claim_id, {"rows": [], "cpts": set(), "dos_set": set(), "billed": "", "patient_name": ""}
                )
                g["rows"].append(row)
                if dos_raw:
                    g["dos_set"].add(str(dos_raw).strip())
                if cpt:
                    g["cpts"].add(cpt)
                if billed and not g["billed"]:
                    g["billed"] = billed
                if p_name and not g["patient_name"]:
                    g["patient_name"] = p_name

        for claim_id, g in groups.items():
            dos_list = sorted(g["dos_set"])
            dos_str = dos_list[0] if dos_list else ""
            claims_flat.append(
                {
                    "claim_id": claim_id,
                    "dos": fmt_mmddyy_from_any(dos_str),
                    "dos_raw": dos_str,
                    "billed_amount": g.get("billed") or "",
                    "cpts": sorted(g["cpts"]),
                    "_row": (g["rows"][0] if g["rows"] else {}),
                }
            )

        if not patient_name:
            for _cid, g in groups.items():
                if g.get("patient_name"):
                    patient_name = g["patient_name"]
                    break

        total_claims = len(groups)
        ui_log(f"[INFO] Grouped {total_claims} unique claim_id(s) from {rows_read} CSV row(s).")
        if total_claims == 0:
            raise RuntimeError("No valid claim_id entries found in CSV.")

        claim_list_fd, claim_list_path = tempfile.mkstemp(suffix=".pdf")
        os.close(claim_list_fd)
        temp_internal_paths.append(claim_list_path)

        create_claim_list_pdf(
            output_path=claim_list_path,
            patient_name=patient_name,
            claims=claims_flat,
            page_size="letter",
        )
        claim_list_pdf_cached = load_optional_pdf(claim_list_path, "Claim List", temp_paths_to_cleanup)
        if not claim_list_pdf_cached:
            raise RuntimeError("Could not create/read Claim List PDF for bulk mode.")

        progress_bar["maximum"] = 1
        progress_bar["value"] = 0
        progress_label.config(text="0%")
        update_timer(1)

        ui_log("[INFO] Building bulk PDF...")
        create_pdf_bulk_v2(
            patient_name=patient_name,
            claims_order=claims_flat,
            cover_letter_path=cover_letter_file,
            payer_letter_cached_list=payer_letter_cached_list,
            claim_list_pdf_cached=claim_list_pdf_cached,
            aor_cached=aor_cached,
            base_docs_cached=base_docs_cached,
            soap_reader=soap_reader,
            soap_pdf_path=filter_file,
            soap_index=soap_index,
            target_list_cached=target_list_cached,
            dtt_cached_list=dtt_cached_list,
            daily_behavior_data_cached_list=daily_behavior_data_cached_list,
            daily_trial_counts_cached_list=daily_trial_counts_cached_list,
            behavior_reduction_cached_list=behavior_reduction_cached_list,
            output_directory=output_dir,
        )

        progress_bar["value"] = 1
        progress_label.config(text="100%")

        if cancel_operation:
            ui_log("=== Cancelled ===")
            messagebox.showinfo("Cancelled", "Operation cancelled.")
            progress_label.config(text="Cancelled")
        else:
            ui_log("=== Completed BULK ===")
            messagebox.showinfo("Success", f"Bulk PDF created successfully!\n\nOutput folder: {output_dir}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during bulk processing:\n{e}")
        progress_label.config(text="Error")

    finally:
        for p in temp_paths_to_cleanup:
            try:
                os.remove(p)
            except OSError:
                pass
        for p in temp_internal_paths:
            try:
                os.remove(p)
            except OSError:
                pass

        generate_button.config(state="normal")
        generate_bulk_button.config(state="normal")
        cancel_button.config(state="disabled")
        cancel_operation = False
        stop_timer_and_finalize()


# ----------------------------
# Cancel / Reset
# ----------------------------
def cancel_process():
    global cancel_operation
    cancel_operation = True
    progress_label.config(text="Cancelling...")
    root.update_idletasks()


def reset_app():
    global csv_file, cover_letter_file, payer_letter_files, aor_file
    global base_files, target_list_file, dtt_files, daily_behavior_data_files, daily_trial_counts_files, behavior_reduction_files
    global filter_file, output_dir, cancel_operation

    if processing_thread and processing_thread.is_alive():
        messagebox.showinfo("Info", "Cannot reset while processing.")
        return

    csv_file = ""
    cover_letter_file = ""
    payer_letter_files = []
    aor_file = ""

    base_files = []
    target_list_file = ""
    dtt_files = []
    daily_behavior_data_files = []
    daily_trial_counts_files = []
    behavior_reduction_files = []

    filter_file = ""
    output_dir = ""
    cancel_operation = False

    for lbl in (
        csv_label,
        cover_label,
        payer_letter_label,
        aor_label,
        base_label,
        target_list_label,
        dtt_label,
        daily_behavior_data_label,
        daily_trial_counts_label,
        behavior_reduction_label,
        filter_label,
        output_label,
    ):
        if lbl is not None:
            lbl.config(text="")

    progress_bar["value"] = 0
    progress_label.config(text="0%")
    elapsed_label.config(text="Time: 00:00")
    eta_label.config(text="Remaining (est.): --:--")

    cancel_button.config(state="disabled")
    generate_button.config(state="normal")
    generate_bulk_button.config(state="normal")

    if log_text is not None:
        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.configure(state="disabled")


# ----------------------------
# File names display helper
# ----------------------------
def _display_name_for_file(path: str) -> str:
    return os.path.basename(path) if path else ""


def _display_name_for_dir(path: str) -> str:
    if not path:
        return ""
    base = os.path.basename(path.rstrip("\\/"))
    return base or path


# ----------------------------
# File pickers
# ----------------------------
def select_csv():
    global csv_file
    csv_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if csv_file:
        csv_label.config(text=_display_name_for_file(csv_file))


def select_cover_letter():
    global cover_letter_file
    cover_letter_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if cover_letter_file:
        cover_label.config(text=_display_name_for_file(cover_letter_file))


def select_payer_letter_files():
    global payer_letter_files
    payer_letter_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if payer_letter_files:
        payer_letter_label.config(
            text=_display_name_for_file(payer_letter_files[0])
            if len(payer_letter_files) == 1
            else f"{len(payer_letter_files)} file(s) selected"
        )
    else:
        payer_letter_label.config(text="")


def select_aor_file():
    global aor_file
    aor_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    aor_label.config(text=_display_name_for_file(aor_file) or "")


def select_base_files():
    global base_files
    base_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if base_files:
        base_label.config(
            text=_display_name_for_file(base_files[0]) if len(base_files) == 1 else f"{len(base_files)} file(s) selected"
        )
    else:
        base_label.config(text="")


def select_target_list_pdf():
    global target_list_file
    target_list_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    target_list_label.config(text=_display_name_for_file(target_list_file) or "")


def select_dtt_files():
    global dtt_files
    dtt_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if dtt_files:
        dtt_label.config(
            text=_display_name_for_file(dtt_files[0]) if len(dtt_files) == 1 else f"{len(dtt_files)} file(s) selected"
        )
    else:
        dtt_label.config(text="")


def select_daily_behavior_data_files():
    global daily_behavior_data_files
    daily_behavior_data_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if daily_behavior_data_files:
        daily_behavior_data_label.config(
            text=_display_name_for_file(daily_behavior_data_files[0])
            if len(daily_behavior_data_files) == 1
            else f"{len(daily_behavior_data_files)} file(s) selected"
        )
    else:
        daily_behavior_data_label.config(text="")


def select_daily_trial_counts_files():
    global daily_trial_counts_files
    daily_trial_counts_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if daily_trial_counts_files:
        daily_trial_counts_label.config(
            text=_display_name_for_file(daily_trial_counts_files[0])
            if len(daily_trial_counts_files) == 1
            else f"{len(daily_trial_counts_files)} file(s) selected"
        )
    else:
        daily_trial_counts_label.config(text="")


def select_behavior_reduction_files():
    global behavior_reduction_files
    behavior_reduction_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if behavior_reduction_files:
        behavior_reduction_label.config(
            text=_display_name_for_file(behavior_reduction_files[0])
            if len(behavior_reduction_files) == 1
            else f"{len(behavior_reduction_files)} file(s) selected"
        )
    else:
        behavior_reduction_label.config(text="")


def select_filter_file():
    global filter_file
    filter_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    filter_label.config(text=_display_name_for_file(filter_file) or "")


def select_output_dir():
    global output_dir
    output_dir = filedialog.askdirectory()
    output_label.config(text=_display_name_for_dir(output_dir) or "")


# ----------------------------
# UI helper: scrollable container
# ----------------------------
class VScrollableFrame(ttk.Frame):
    """A vertically scrollable frame (mouse wheel supported) for lots of widgets."""

    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.inner.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_frame_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.inner_id, width=event.width)

    def _on_mousewheel(self, event):
        try:
            x = self.winfo_pointerx() - self.winfo_rootx()
            y = self.winfo_pointery() - self.winfo_rooty()
            if 0 <= x <= self.winfo_width() and 0 <= y <= self.winfo_height():
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass


# ----------------------------
# UI builder
# ----------------------------
def build_ui():
    global root, log_text
    global csv_label, cover_label, payer_letter_label, aor_label, base_label
    global target_list_label, dtt_label, daily_behavior_data_label, daily_trial_counts_label, behavior_reduction_label
    global filter_label, output_label
    global progress_bar, progress_label, elapsed_label, eta_label
    global generate_button, generate_bulk_button, cancel_button, reset_button

    root = tk.Tk()
    set_windows_appusermodelid("DanielBernal.MassMedicalRecordsGenerator.1")

    icon_path = resource_path(os.path.join("assets", "app.ico"))
    try:
        root.iconbitmap(icon_path)
    except Exception as e:
        ui_log(f"[WARN] Could not set icon: {e}")

    root.title("Medical Records Package Compiler")
    root.geometry("1100x800")
    root.minsize(950, 750)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
    style.configure("Section.TLabelframe", padding=(12, 10))
    style.configure("TButton", padding=(10, 6))
    style.configure("Small.TLabel", font=("Segoe UI", 9), foreground="#444")
    style.configure("Path.TLabel", font=("Segoe UI", 9), foreground="#0B5394")

    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    container = ttk.Frame(root, padding=14)
    container.grid(row=0, column=0, sticky="nsew")

    container.grid_columnconfigure(0, weight=1)
    container.grid_rowconfigure(2, weight=1)
    container.grid_rowconfigure(4, weight=1)

    header = ttk.Label(container, text="Medical Records Package Compiler", style="Title.TLabel")
    header.grid(row=0, column=0, sticky="w")

    subtitle = ttk.Label(
        container,
        text="Per-Claim = one PDF per claim_id.  Bulk = one PDF for all claim_ids in the CSV.",
        style="Small.TLabel",
    )
    subtitle.grid(row=1, column=0, sticky="w", pady=(2, 10))

    # Scrollable Files section
    scroller = VScrollableFrame(container)
    scroller.grid(row=2, column=0, sticky="nsew")

    scroller.grid_rowconfigure(0, weight=1)
    scroller.grid_columnconfigure(0, weight=1)
    scroller.inner.grid_columnconfigure(0, weight=1)

    files_frame = ttk.Labelframe(scroller.inner, text="Files", style="Section.TLabelframe")
    files_frame.grid(row=0, column=0, sticky="nsew")

    files_frame.grid_columnconfigure(0, weight=0, minsize=290)
    files_frame.grid_columnconfigure(1, weight=0)
    files_frame.grid_columnconfigure(2, weight=1)

    def add_picker_row(r, label, button_text, command, assign_label_name):
        ttk.Label(files_frame, text=label).grid(row=r, column=0, sticky="w", padx=(0, 12), pady=6)

        btn = ttk.Button(files_frame, text=button_text, command=command)
        btn.grid(row=r, column=1, sticky="w", pady=6)

        lbl = ttk.Label(files_frame, text="", style="Path.TLabel", anchor="w", justify="left")
        lbl.grid(row=r, column=2, sticky="ew", pady=6, padx=(12, 0))

        globals()[assign_label_name] = lbl

    add_picker_row(0, "Claims List (CSV)", "Select", select_csv, "csv_label")
    add_picker_row(1, "Cover Letter (PDF)", "Select", select_cover_letter, "cover_label")
    add_picker_row(2, "Payer Letter (optional)", "Select", select_payer_letter_files, "payer_letter_label")
    add_picker_row(3, "AOR - Consent Form (mandatory)", "Select", select_aor_file, "aor_label")
    add_picker_row(4, "Base Docs (TX / DX / Progress) PDFs", "Select", select_base_files, "base_label")
    add_picker_row(5, "Target List Report (mandatory PDF)", "Select", select_target_list_pdf, "target_list_label")
    add_picker_row(6, "DTT Trial Sheets (mandatory PDFs)", "Select", select_dtt_files, "dtt_label")
    add_picker_row(7, "Daily Behavior Data (mandatory PDFs)", "Select", select_daily_behavior_data_files, "daily_behavior_data_label")
    add_picker_row(8, "Daily Trial Counts (mandatory PDFs)", "Select", select_daily_trial_counts_files, "daily_trial_counts_label")
    add_picker_row(9, "Behavior Reduction Report (mandatory PDFs)", "Select", select_behavior_reduction_files, "behavior_reduction_label")
    add_picker_row(10, "SOAP Notes (PDF)", "Select", select_filter_file, "filter_label")
    add_picker_row(11, "Output Folder", "Select", select_output_dir, "output_label")

    # Progress
    progress_frame = ttk.Labelframe(container, text="Progress", style="Section.TLabelframe")
    progress_frame.grid(row=3, column=0, sticky="ew", pady=(12, 0))
    progress_frame.grid_columnconfigure(0, weight=1)

    progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
    progress_bar.grid(row=0, column=0, sticky="ew", pady=(6, 6))

    progress_label = ttk.Label(progress_frame, text="0%", font=("Segoe UI", 10, "bold"))
    progress_label.grid(row=1, column=0, sticky="w")

    elapsed_label = ttk.Label(progress_frame, text="Time: 00:00")
    elapsed_label.grid(row=2, column=0, sticky="w")

    eta_label = ttk.Label(progress_frame, text="Remaining (est.): --:--")
    eta_label.grid(row=3, column=0, sticky="w")

    # Buttons
    button_frame = ttk.Frame(container)
    button_frame.grid(row=4, column=0, sticky="ew", pady=(10, 0))

    generate_button = ttk.Button(button_frame, text="Generate (Per-Claim)", command=process_per_claim)
    generate_button.pack(side="left", padx=5)

    generate_bulk_button = ttk.Button(button_frame, text="Generate (Bulk)", command=process_bulk)
    generate_bulk_button.pack(side="left", padx=5)

    cancel_button = ttk.Button(button_frame, text="Cancel", command=cancel_process, state="disabled")
    cancel_button.pack(side="left", padx=5)

    reset_button = ttk.Button(button_frame, text="Reset", command=reset_app)
    reset_button.pack(side="left", padx=5)

    # Activity Log
    ttk.Label(container, text="Activity Log:").grid(row=5, column=0, sticky="w", pady=(10, 0))

    log_text = tk.Text(container, height=8, bg="#1e1e1e", fg="#00ff00", font=("Courier New", 9))
    log_text.grid(row=6, column=0, sticky="nsew", pady=(5, 0))
    log_text.configure(state="disabled")

    container.grid_rowconfigure(6, weight=1)

    root.mainloop()


if __name__ == "__main__":
    build_ui()