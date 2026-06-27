"""
build_batch_excel.py
Creates an Excel file with one sheet per NotebookLM batch.
Each sheet contains the full metadata for every article in that batch,
matched from the original Scopus Excel using the DOI from the PDF filename.

Usage:
  1. Fill in BATCH_FOLDERS below with the path to each batch folder.
  2. Run: python build_batch_excel.py
  3. Output: batch_overview.xlsx
"""

import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
EXCEL_FILE = "Nets4Dem_ScopusReview_Final_GA_12 June.xlsx"
SHEETS_TO_SCAN = ["HIGH relevance", "MEDIUM screening"]
OUTPUT_FILE = "batch_overview.xlsx"

# List of (sheet_name_in_output, path_to_batch_folder)
# Edit these to match your actual folder paths and desired sheet names
BATCH_FOLDERS = [
    ("Batch 1", r"C:\Users\Administrator\Downloads\Livro\Batch1"),
    ("Batch 2", r"C:\Users\Administrator\Downloads\Livro\Batch2"),
    ("Batch 3", r"C:\Users\Administrator\Downloads\Livro\Batch3"),
    ("Batch 4", r"C:\Users\Administrator\Downloads\Livro\Batch4"),
    ("Batch 5", r"C:\Users\Administrator\Downloads\Livro\Batch5"),
]
# ───────────────────────────────────────────────────────────────────────────────


def sanitize_filename(doi: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", doi)


def build_doi_lookup(excel_file: str) -> dict:
    """
    Reads all sheets in the source Excel and builds a dict:
      sanitized_doi_filename -> {col_name: value, ...}
    """
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    lookup = {}

    for sheet_name in SHEETS_TO_SCAN:
        if sheet_name not in wb.sheetnames:
            print(f"  [WARN] Sheet '{sheet_name}' not found in source Excel")
            continue
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]

        doi_idx = None
        for i, h in enumerate(headers):
            if h and str(h).strip().upper() == "DOI":
                doi_idx = i
                break
        if doi_idx is None:
            print(f"  [WARN] No DOI column in '{sheet_name}'")
            continue

        for row in ws.iter_rows(min_row=2, values_only=True):
            raw_doi = row[doi_idx]
            if not raw_doi:
                continue
            doi = str(raw_doi).strip()
            for prefix in ("https://doi.org/", "http://doi.org/", "doi.org/", "DOI:", "doi:"):
                if doi.lower().startswith(prefix.lower()):
                    doi = doi[len(prefix):]
            doi = doi.strip()
            if not doi or doi.lower() == "nan":
                continue

            key = sanitize_filename(doi)
            record = {"_source_sheet": sheet_name}
            for col_name, value in zip(headers, row):
                if col_name:
                    record[str(col_name)] = value
            lookup[key] = record

    print(f"  Loaded {len(lookup)} articles from source Excel.")
    return lookup


def get_dois_in_folder(folder_path: str) -> list[str]:
    """Returns list of DOI keys (filename without .pdf) for all PDFs in folder."""
    if not os.path.isdir(folder_path):
        print(f"  [WARN] Folder not found: {folder_path}")
        return []
    keys = []
    for fname in sorted(os.listdir(folder_path)):
        if fname.lower().endswith(".pdf"):
            keys.append(fname[:-4])  # strip .pdf
    return keys


def write_sheet(ws, headers: list, rows: list):
    """Write headers + data rows to a worksheet with basic styling."""
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)

    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for row in rows:
        ws.append(row)

    # Auto-width (capped at 60)
    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(header))
        for row_cells in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row_cells:
                if cell.value:
                    max_len = max(max_len, min(len(str(cell.value)), 60))
        ws.column_dimensions[col_letter].width = max_len + 2

    ws.freeze_panes = "A2"


def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"[ERROR] Source Excel not found: {EXCEL_FILE}")
        return

    print(f"Reading source Excel: {EXCEL_FILE}")
    lookup = build_doi_lookup(EXCEL_FILE)

    # Determine all column headers from lookup (preserve insertion order)
    all_cols = []
    seen = set()
    for record in lookup.values():
        for k in record:
            if k != "_source_sheet" and k not in seen:
                all_cols.append(k)
                seen.add(k)

    output_headers = ["Source sheet"] + all_cols + ["PDF filename", "Matched"]

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)  # remove default empty sheet

    summary_rows = []

    for sheet_name, folder_path in BATCH_FOLDERS:
        print(f"\nProcessing batch '{sheet_name}' from {folder_path}")
        doi_keys = get_dois_in_folder(folder_path)
        print(f"  Found {len(doi_keys)} PDFs")

        rows = []
        matched = unmatched = 0
        for key in doi_keys:
            record = lookup.get(key)
            if record:
                row = [record.get("_source_sheet", "")]
                for col in all_cols:
                    row.append(record.get(col, ""))
                row.append(key + ".pdf")
                row.append("YES")
                matched += 1
            else:
                row = [""] * (len(all_cols) + 1)
                row.append(key + ".pdf")
                row.append("NO - not found in Excel")
                unmatched += 1
            rows.append(row)

        ws = out_wb.create_sheet(title=sheet_name[:31])  # Excel max 31 chars
        write_sheet(ws, output_headers, rows)
        print(f"  Matched: {matched}  |  Unmatched: {unmatched}")
        summary_rows.append((sheet_name, folder_path, len(doi_keys), matched, unmatched))

    # Summary sheet
    ws_sum = out_wb.create_sheet(title="Summary", index=0)
    sum_headers = ["Batch", "Folder", "Total PDFs", "Matched", "Unmatched"]
    write_sheet(ws_sum, sum_headers, summary_rows)

    out_wb.save(OUTPUT_FILE)
    print(f"\n{'='*60}")
    print(f"Output saved to: {os.path.abspath(OUTPUT_FILE)}")
    for name, folder, total, matched, unmatched in summary_rows:
        print(f"  {name}: {total} PDFs | {matched} matched | {unmatched} unmatched")


if __name__ == "__main__":
    main()
