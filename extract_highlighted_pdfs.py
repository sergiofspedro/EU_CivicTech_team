"""
extract_highlighted_pdfs.py
Reads batch_overview.xlsx (with your yellow highlights across Batch 1-5 sheets)
and copies the corresponding PDFs from their batch folders into a single
destination folder.

Usage:
  python extract_highlighted_pdfs.py
"""

import os
import shutil
import openpyxl

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
REVIEWED_EXCEL   = "batch_overview.xlsx"   # the file you highlighted
BATCH_SHEETS     = ["Batch 1", "Batch 2", "Batch 3", "Batch 4", "Batch 5"]
YELLOW_COLOR     = "FFFFFF00"

# Source folders for each batch (must match where the PDFs actually live)
BATCH_FOLDERS = {
    "Batch 1": r"C:\Users\Administrator\Downloads\Livro\HIGH relevance\1",
    "Batch 2": r"C:\Users\Administrator\Downloads\Livro\HIGH relevance\2",
    "Batch 3": r"C:\Users\Administrator\Downloads\Livro\HIGH relevance\3",
    "Batch 4": r"C:\Users\Administrator\Downloads\Livro\MEDIUM screening\4",
    "Batch 5": r"C:\Users\Administrator\Downloads\Livro\MEDIUM screening\5",
}

DEST_FOLDER = r"C:\Users\Administrator\Downloads\Livro\Selected"
# ───────────────────────────────────────────────────────────────────────────────


def row_color(ws_row) -> str | None:
    for cell in ws_row:
        fill = cell.fill
        if fill and fill.fgColor and fill.fgColor.type == "rgb":
            rgb = fill.fgColor.rgb
            if rgb and rgb not in ("00000000", "FF000000"):
                return rgb
    return None


def main():
    if not os.path.exists(REVIEWED_EXCEL):
        print(f"[ERROR] '{REVIEWED_EXCEL}' not found.")
        return

    os.makedirs(DEST_FOLDER, exist_ok=True)

    wb = openpyxl.load_workbook(REVIEWED_EXCEL, data_only=True)

    copied = missing = skipped_not_yellow = 0
    missing_list = []

    for sheet_name in BATCH_SHEETS:
        if sheet_name not in wb.sheetnames:
            print(f"[WARN] Sheet '{sheet_name}' not found — skipping.")
            continue
        ws = wb[sheet_name]
        headers = [c.value for c in ws[1]]
        pdf_idx = headers.index("PDF filename") if "PDF filename" in headers else None
        title_idx = headers.index("Title") if "Title" in headers else None
        if pdf_idx is None:
            print(f"[WARN] No 'PDF filename' column in '{sheet_name}' — skipping.")
            continue

        src_folder = BATCH_FOLDERS.get(sheet_name)
        sheet_count = 0

        for row in ws.iter_rows(min_row=2):
            color = row_color(row)
            if color != YELLOW_COLOR:
                continue

            pdf_name = row[pdf_idx].value
            if not pdf_name:
                continue
            pdf_name = str(pdf_name).strip()

            src_path = os.path.join(src_folder, pdf_name)
            dest_path = os.path.join(DEST_FOLDER, pdf_name)

            title = str(row[title_idx].value)[:60] if title_idx is not None and row[title_idx].value else pdf_name

            if not os.path.exists(src_path):
                print(f"  [MISSING] {sheet_name}: {pdf_name} — {title}")
                missing += 1
                missing_list.append((sheet_name, pdf_name, title))
                continue

            if os.path.exists(dest_path):
                print(f"  [SKIP] Already in destination: {pdf_name}")
            else:
                shutil.copy2(src_path, dest_path)
                print(f"  [OK] {sheet_name}: {pdf_name}")

            copied += 1
            sheet_count += 1

        print(f"{sheet_name}: {sheet_count} highlighted PDFs processed\n")

    print(f"{'='*60}")
    print(f"Total copied/found: {copied}")
    print(f"Missing source files: {missing}")
    print(f"Destination folder: {os.path.abspath(DEST_FOLDER)}")

    if missing_list:
        print(f"\nMissing files (source PDF not found in batch folder):")
        for sheet_name, pdf_name, title in missing_list:
            print(f"  {sheet_name}: {pdf_name} — {title}")


if __name__ == "__main__":
    main()
