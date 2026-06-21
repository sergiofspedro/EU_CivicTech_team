"""
PDF Downloader — Nets4Dem Scopus Review
Reads an Excel file, detects row highlight colors, and downloads PDFs via:
  Layer 1 - Unpaywall API (Open Access)
  Layer 2 - Direct DOI resolution via institutional VPN
"""

import os
import re
import time
import requests
import openpyxl

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
EXCEL_FILE      = "Nets4Dem_ScopusReview_Final_GA_12_June.xlsx"
DOI_COLUMN      = "DOI"
UNPAYWALL_EMAIL = "your_email@example.com"   # <-- change to your e-mail
DELAY_SECONDS   = 2

YELLOW_COLOR    = "FFFFFF00"                 # bright yellow → scrape
GREEN_COLORS    = {"FFC6EFCE", "FF92D050"}   # light/dark green → Future readings
SHEETS_TO_SCAN  = ["HIGH relevance", "MEDIUM screening"]
FUTURE_FOLDER   = "Future readings"
# ───────────────────────────────────────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


def clean_doi(raw) -> str | None:
    """Return a bare DOI string, or None if the value is empty/invalid."""
    if raw is None:
        return None
    doi = str(raw).strip()
    for prefix in ("https://doi.org/", "http://doi.org/", "doi.org/", "DOI:", "doi:"):
        if doi.lower().startswith(prefix.lower()):
            doi = doi[len(prefix):]
    doi = doi.strip()
    return doi if doi and doi.lower() != "nan" else None


def sanitize_filename(doi: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", doi)


def row_color(ws_row) -> str | None:
    """Return the fgColor RGB of the first colored cell in the row, or None."""
    for cell in ws_row:
        fill = cell.fill
        if fill and fill.fgColor and fill.fgColor.type == "rgb":
            rgb = fill.fgColor.rgb
            if rgb and rgb != "00000000" and rgb != "FF000000":
                return rgb
    return None


def collect_articles(wb) -> list[dict]:
    """
    Walk SHEETS_TO_SCAN, detect row colors, and build a list of
    {doi, folder} dicts for rows that need downloading.
    """
    articles = []
    for sheet_name in SHEETS_TO_SCAN:
        if sheet_name not in wb.sheetnames:
            print(f"[WARN] Sheet '{sheet_name}' not found — skipping.")
            continue
        ws = wb[sheet_name]

        # Find DOI column index from header row
        header = [cell.value for cell in ws[1]]
        try:
            doi_idx = header.index(DOI_COLUMN)
        except ValueError:
            print(f"[WARN] Column '{DOI_COLUMN}' not found in '{sheet_name}' — skipping.")
            continue

        yellow_count = green_count = 0
        for row in ws.iter_rows(min_row=2):
            color = row_color(row)
            if color == YELLOW_COLOR:
                doi = clean_doi(row[doi_idx].value)
                if doi:
                    articles.append({"doi": doi, "folder": sheet_name})
                    yellow_count += 1
            elif color in GREEN_COLORS:
                doi = clean_doi(row[doi_idx].value)
                if doi:
                    articles.append({"doi": doi, "folder": FUTURE_FOLDER})
                    green_count += 1

        print(f"  '{sheet_name}': {yellow_count} yellow (scrape) | {green_count} green (Future readings)")

    return articles


def try_unpaywall(doi: str) -> str | None:
    url = f"https://api.unpaywall.org/v2/{doi}?email={UNPAYWALL_EMAIL}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            return None
        data = resp.json()
        if not data.get("is_oa"):
            return None
        location = data.get("best_oa_location") or {}
        return location.get("url_for_pdf") or location.get("url")
    except Exception:
        return None


def download_bytes(url: str) -> bytes | None:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=30, allow_redirects=True)
        content_type = resp.headers.get("Content-Type", "")
        if resp.status_code == 200 and "pdf" in content_type.lower():
            return resp.content
        if resp.status_code == 200 and len(resp.content) > 10_000:
            return resp.content
        return None
    except Exception:
        return None


def download_via_vpn(doi: str) -> bytes | None:
    return download_bytes(f"https://doi.org/{doi}")


def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"[ERROR] Excel file '{EXCEL_FILE}' not found in current directory.")
        return

    print(f"Reading '{EXCEL_FILE}' ...")
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)

    print("Scanning sheets for highlighted rows ...")
    articles = collect_articles(wb)
    total = len(articles)
    print(f"\nTotal articles to download: {total}\n")

    # Create output folders
    folders_needed = {a["folder"] for a in articles}
    for folder in folders_needed:
        os.makedirs(folder, exist_ok=True)

    success = fail = skipped = 0

    for idx, art in enumerate(articles, start=1):
        doi    = art["doi"]
        folder = art["folder"]
        safe   = sanitize_filename(doi) + ".pdf"
        path   = os.path.join(folder, safe)

        print(f"[{idx}/{total}] {doi}")

        if os.path.exists(path):
            print(f"  [SKIP] Already exists.")
            skipped += 1
            continue

        pdf_bytes = source = None

        # Layer 1: Unpaywall
        try:
            pdf_url = try_unpaywall(doi)
            if pdf_url:
                pdf_bytes = download_bytes(pdf_url)
                if pdf_bytes:
                    source = "Unpaywall (OA)"
        except Exception as e:
            print(f"  [WARN] Unpaywall error: {e}")

        # Layer 2: VPN
        if not pdf_bytes:
            try:
                pdf_bytes = download_via_vpn(doi)
                if pdf_bytes:
                    source = "VPN"
            except Exception as e:
                print(f"  [WARN] VPN error: {e}")

        if pdf_bytes:
            with open(path, "wb") as f:
                f.write(pdf_bytes)
            print(f"  [OK] {source} → {path}")
            success += 1
        else:
            print(f"  [FAIL] Could not download.")
            fail += 1

        time.sleep(DELAY_SECONDS)

    print(f"\n{'='*60}")
    print(f"Done.  Success: {success}  |  Failed: {fail}  |  Skipped: {skipped}  |  Total: {total}")
    for folder in sorted(folders_needed):
        count = len([f for f in os.listdir(folder) if f.endswith(".pdf")])
        print(f"  {folder}/  → {count} PDFs")


if __name__ == "__main__":
    main()
