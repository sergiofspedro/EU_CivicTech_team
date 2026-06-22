"""
PDF Downloader — Nets4Dem Scopus Review
Reads an Excel file, detects row highlight colors, and downloads PDFs via
a 6-layer fallback pipeline:
  1. Unpaywall API
  2. Semantic Scholar
  3. OpenAlex
  4. Europe PMC
  5. Sci-Hub
  6. Direct DOI / VPN + HTML parsing
"""

import os
import re
import time
import requests
import openpyxl
from urllib.parse import urljoin, urlparse

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
EXCEL_FILE      = "Nets4Dem_ScopusReview_Final_GA_12 June.xlsx"
DOI_COLUMN      = "DOI"
UNPAYWALL_EMAIL = "sergiopedro@ua.pt"
DELAY_SECONDS   = 2

YELLOW_COLOR    = "FFFFFF00"
GREEN_COLORS    = {"FFC6EFCE", "FF92D050"}
SHEETS_TO_SCAN  = ["HIGH relevance", "MEDIUM screening"]
FUTURE_FOLDER   = "Future readings"

SCIHUB_MIRRORS  = [
    "https://sci-hub.se",
    "https://sci-hub.st",
    "https://sci-hub.ru",
]
# ───────────────────────────────────────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


# ── Helpers ────────────────────────────────────────────────────────────────────

def clean_doi(raw) -> str | None:
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
    for cell in ws_row:
        fill = cell.fill
        if fill and fill.fgColor and fill.fgColor.type == "rgb":
            rgb = fill.fgColor.rgb
            if rgb and rgb not in ("00000000", "FF000000"):
                return rgb
    return None


def fetch(url: str, timeout: int = 20) -> requests.Response | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        return r if r.status_code == 200 else None
    except Exception:
        return None


def is_pdf(resp: requests.Response) -> bool:
    ct = resp.headers.get("Content-Type", "")
    return "pdf" in ct.lower()


def pdf_from_response(resp: requests.Response) -> bytes | None:
    if is_pdf(resp) and len(resp.content) > 5_000:
        return resp.content
    return None


# ── Excel scanning ─────────────────────────────────────────────────────────────

def collect_articles(wb) -> list[dict]:
    articles = []
    for sheet_name in SHEETS_TO_SCAN:
        if sheet_name not in wb.sheetnames:
            print(f"  [WARN] Sheet '{sheet_name}' not found — skipping.")
            continue
        ws = wb[sheet_name]
        header = [cell.value for cell in ws[1]]
        try:
            doi_idx = header.index(DOI_COLUMN)
        except ValueError:
            print(f"  [WARN] Column '{DOI_COLUMN}' not in '{sheet_name}' — skipping.")
            continue

        yellow_count = green_count = 0
        for row in ws.iter_rows(min_row=2):
            color = row_color(row)
            doi = clean_doi(row[doi_idx].value)
            if not doi:
                continue
            if color == YELLOW_COLOR:
                articles.append({"doi": doi, "folder": sheet_name})
                yellow_count += 1
            elif color in GREEN_COLORS:
                articles.append({"doi": doi, "folder": FUTURE_FOLDER})
                green_count += 1

        print(f"  '{sheet_name}': {yellow_count} yellow | {green_count} green")
    return articles


# ── Download layers ────────────────────────────────────────────────────────────

def layer_unpaywall(doi: str) -> bytes | None:
    url = f"https://api.unpaywall.org/v2/{doi}?email={UNPAYWALL_EMAIL}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return None
        data = r.json()
        if not data.get("is_oa"):
            return None
        loc = data.get("best_oa_location") or {}
        pdf_url = loc.get("url_for_pdf") or loc.get("url")
        if pdf_url:
            resp = fetch(pdf_url, timeout=30)
            if resp:
                return pdf_from_response(resp)
    except Exception:
        pass
    return None


def layer_semantic_scholar(doi: str) -> bytes | None:
    url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}?fields=openAccessPdf"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return None
        data = r.json()
        oa = data.get("openAccessPdf") or {}
        pdf_url = oa.get("url")
        if pdf_url:
            resp = fetch(pdf_url, timeout=30)
            if resp:
                return pdf_from_response(resp)
    except Exception:
        pass
    return None


def layer_openalex(doi: str) -> bytes | None:
    url = f"https://api.openalex.org/works/https://doi.org/{doi}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return None
        data = r.json()
        oa = data.get("open_access") or {}
        pdf_url = oa.get("oa_url")
        if pdf_url:
            resp = fetch(pdf_url, timeout=30)
            if resp:
                return pdf_from_response(resp)
    except Exception:
        pass
    return None


def layer_europe_pmc(doi: str) -> bytes | None:
    search_url = (
        f"https://www.ebi.ac.uk/europepmc/webservices/rest/search"
        f"?query=DOI:{doi}&format=json&resultType=core"
    )
    try:
        r = requests.get(search_url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return None
        results = r.json().get("resultList", {}).get("result", [])
        if not results:
            return None
        pmcid = results[0].get("pmcid")
        if pmcid:
            pdf_url = f"https://europepmc.org/articles/{pmcid}/pdf/render"
            resp = fetch(pdf_url, timeout=30)
            if resp:
                return pdf_from_response(resp)
    except Exception:
        pass
    return None


_PDF_PATTERNS = [
    r'<meta[^>]+name=["\']citation_pdf_url["\'][^>]+content=["\']([^"\']+)["\']',
    r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+name=["\']citation_pdf_url["\']',
    r'href=["\']([^"\']*\.pdf[^"\']*)["\']',
    r'data-pdf-url=["\']([^"\']+)["\']',
    r'"pdfUrl"\s*:\s*"([^"]+)"',
    r'"pdf_url"\s*:\s*"([^"]+)"',
]


def find_pdf_url_in_html(html: str, base_url: str) -> str | None:
    for pattern in _PDF_PATTERNS:
        for m in re.findall(pattern, html, re.IGNORECASE):
            if any(x in m.lower() for x in ["thumb", "cover", "suppl", "fig", "icon"]):
                continue
            full = urljoin(base_url, m)
            if urlparse(full).scheme in ("http", "https"):
                return full
    return None


def layer_scihub(doi: str) -> bytes | None:
    for mirror in SCIHUB_MIRRORS:
        try:
            resp = requests.get(f"{mirror}/{doi}", headers=HEADERS,
                                timeout=30, allow_redirects=True)
            if resp.status_code != 200:
                continue
            pdf_url = find_pdf_url_in_html(resp.text, resp.url)
            for pattern in [r'iframe[^>]+src=["\']([^"\']+)["\']',
                             r'embed[^>]+src=["\']([^"\']+)["\']']:
                for m in re.findall(pattern, resp.text, re.IGNORECASE):
                    candidate = urljoin(resp.url, m)
                    if urlparse(candidate).scheme in ("http", "https"):
                        pdf_url = pdf_url or candidate
                        break
            if pdf_url:
                dl = fetch(pdf_url, timeout=60)
                if dl:
                    result = pdf_from_response(dl)
                    if result:
                        return result
        except Exception:
            continue
    return None


def layer_vpn_html(doi: str) -> bytes | None:
    resp = fetch(f"https://doi.org/{doi}", timeout=30)
    if not resp:
        return None
    if is_pdf(resp):
        return resp.content if len(resp.content) > 5_000 else None
    pdf_url = find_pdf_url_in_html(resp.text, resp.url)
    if pdf_url:
        dl = fetch(pdf_url, timeout=30)
        if dl:
            return pdf_from_response(dl)
    return None


# ── Pipeline ───────────────────────────────────────────────────────────────────

LAYERS = [
    ("Unpaywall",        layer_unpaywall),
    ("Semantic Scholar", layer_semantic_scholar),
    ("OpenAlex",         layer_openalex),
    ("Europe PMC",       layer_europe_pmc),
    ("Sci-Hub",          layer_scihub),
    ("VPN/HTML",         layer_vpn_html),
]


def download_pdf(doi: str) -> tuple[bytes | None, str | None]:
    for name, fn in LAYERS:
        try:
            result = fn(doi)
            if result:
                return result, name
        except Exception as e:
            print(f"    [{name}] error: {e}")
    return None, None


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"[ERROR] '{EXCEL_FILE}' not found in current directory.")
        return

    print(f"Reading '{EXCEL_FILE}' ...")
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)

    print("Scanning sheets ...")
    articles = collect_articles(wb)
    total = len(articles)
    print(f"\nTotal articles to process: {total}\n")

    folders_needed = {a["folder"] for a in articles}
    for folder in folders_needed:
        os.makedirs(folder, exist_ok=True)

    success = fail = skipped = 0
    failed_list = []  # collect (doi, folder) for failed downloads

    for idx, art in enumerate(articles, start=1):
        doi    = art["doi"]
        folder = art["folder"]
        path   = os.path.join(folder, sanitize_filename(doi) + ".pdf")

        print(f"[{idx}/{total}] {doi}")

        if os.path.exists(path):
            print(f"  [SKIP] Already exists.")
            skipped += 1
            continue

        pdf_bytes, source = download_pdf(doi)

        if pdf_bytes:
            with open(path, "wb") as f:
                f.write(pdf_bytes)
            print(f"  [OK] {source} → {path}")
            success += 1
        else:
            print(f"  [FAIL] All layers exhausted.")
            failed_list.append({"doi": doi, "folder": folder,
                                 "doi_url": f"https://doi.org/{doi}"})
            fail += 1

        time.sleep(DELAY_SECONDS)

    print(f"\n{'='*60}")
    print(f"Done.  OK: {success}  |  Failed: {fail}  |  Skipped: {skipped}  |  Total: {total}")
    for folder in sorted(folders_needed):
        count = len([f for f in os.listdir(folder) if f.endswith(".pdf")])
        print(f"  {folder}/  → {count} PDFs")

    # Write failed DOIs to CSV
    if failed_list:
        import csv
        csv_path = "failed_downloads.csv"
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["doi", "folder", "doi_url"])
            writer.writeheader()
            writer.writerows(failed_list)
        print(f"\nFailed DOIs saved to: {os.path.abspath(csv_path)}")


if __name__ == "__main__":
    main()
