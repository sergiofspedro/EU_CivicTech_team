"""
PDF Downloader via DOI
Reads DOIs from an Excel file and downloads PDFs using:
  Layer 1 - Unpaywall API (Open Access)
  Layer 2 - Direct DOI resolution via institutional VPN
"""

import os
import re
import time
import requests
import pandas as pd

# ─── CONFIGURATION ────────────────────────────────────────────────────────────
EXCEL_FILE      = "artigos.xlsx"
DOI_COLUMN_NAME = "DOI"
OUTPUT_FOLDER   = "PDFs_Baixados"
UNPAYWALL_EMAIL = "your_email@example.com"   # <-- change to your e-mail
DELAY_SECONDS   = 2
# ──────────────────────────────────────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


def clean_doi(raw: str) -> str:
    """Strip whitespace and any leading URL prefix, return bare DOI."""
    doi = str(raw).strip()
    # remove common prefixes
    for prefix in ("https://doi.org/", "http://doi.org/", "doi.org/", "DOI:", "doi:"):
        if doi.startswith(prefix):
            doi = doi[len(prefix):]
    return doi.strip()


def sanitize_filename(doi: str) -> str:
    """Replace characters invalid in filenames with underscores."""
    return re.sub(r'[\\/:*?"<>|]', "_", doi)


def try_unpaywall(doi: str) -> str | None:
    """
    Query Unpaywall. Return a PDF URL if the article is OA, else None.
    """
    url = f"https://api.unpaywall.org/v2/{doi}?email={UNPAYWALL_EMAIL}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            return None
        data = resp.json()
        if not data.get("is_oa"):
            return None
        location = data.get("best_oa_location") or {}
        pdf_url = location.get("url_for_pdf") or location.get("url")
        return pdf_url if pdf_url else None
    except Exception:
        return None


def try_vpn_download(doi: str) -> bytes | None:
    """
    Follow https://doi.org/<doi> with browser headers.
    The university VPN IP grants access; we expect a PDF response.
    """
    doi_url = f"https://doi.org/{doi}"
    try:
        resp = requests.get(doi_url, headers=HEADERS, timeout=30, allow_redirects=True)
        content_type = resp.headers.get("Content-Type", "")
        if resp.status_code == 200 and "pdf" in content_type.lower():
            return resp.content
        return None
    except Exception:
        return None


def download_pdf(pdf_url: str) -> bytes | None:
    """Download bytes from a direct PDF URL."""
    try:
        resp = requests.get(pdf_url, headers=HEADERS, timeout=30, stream=True)
        content_type = resp.headers.get("Content-Type", "")
        if resp.status_code == 200 and "pdf" in content_type.lower():
            return resp.content
        # Some servers omit Content-Type; accept if response is large enough
        if resp.status_code == 200 and len(resp.content) > 10_000:
            return resp.content
        return None
    except Exception:
        return None


def main():
    # Ensure output folder exists
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Load Excel
    if not os.path.exists(EXCEL_FILE):
        print(f"[ERROR] Excel file '{EXCEL_FILE}' not found in current directory.")
        return

    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")

    if DOI_COLUMN_NAME not in df.columns:
        available = ", ".join(df.columns.tolist())
        print(f"[ERROR] Column '{DOI_COLUMN_NAME}' not found. Available columns: {available}")
        return

    # Filter valid DOI rows
    doi_series = df[DOI_COLUMN_NAME].dropna()
    dois = [clean_doi(d) for d in doi_series if str(d).strip() not in ("", "nan")]
    total = len(dois)
    print(f"Found {total} DOIs to process.\n")

    success_count = 0
    fail_count = 0

    for idx, doi in enumerate(dois, start=1):
        print(f"Processing {idx}/{total}: {doi} ...")
        safe_name = sanitize_filename(doi) + ".pdf"
        out_path  = os.path.join(OUTPUT_FOLDER, safe_name)

        if os.path.exists(out_path):
            print(f"  [SKIP] Already downloaded.")
            success_count += 1
            time.sleep(DELAY_SECONDS)
            continue

        pdf_bytes = None
        source    = None

        # ── Layer 1: Unpaywall ──────────────────────────────────────────────
        try:
            pdf_url = try_unpaywall(doi)
            if pdf_url:
                pdf_bytes = download_pdf(pdf_url)
                if pdf_bytes:
                    source = "Unpaywall (OA)"
        except Exception as e:
            print(f"  [WARN] Unpaywall error: {e}")

        # ── Layer 2: VPN / direct DOI resolution ───────────────────────────
        if not pdf_bytes:
            try:
                pdf_bytes = try_vpn_download(doi)
                if pdf_bytes:
                    source = "VPN (institutional access)"
            except Exception as e:
                print(f"  [WARN] VPN download error: {e}")

        # ── Save or log failure ─────────────────────────────────────────────
        if pdf_bytes:
            with open(out_path, "wb") as f:
                f.write(pdf_bytes)
            print(f"  [OK] Saved via {source} → {out_path}")
            success_count += 1
        else:
            print(f"  [FAIL] Could not download PDF for {doi}")
            fail_count += 1

        time.sleep(DELAY_SECONDS)

    print(f"\nDone. Success: {success_count}/{total}  |  Failed: {fail_count}/{total}")
    print(f"PDFs saved in: {os.path.abspath(OUTPUT_FOLDER)}")


if __name__ == "__main__":
    main()
