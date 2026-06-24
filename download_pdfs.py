"""
PDF Downloader — Nets4Dem Scopus Review
Reads an Excel file, detects row highlight colors, and downloads PDFs via
a 6-layer fallback pipeline with improved extraction and diagnostics:
  1. Unpaywall API (all OA locations, not just best)
  2. Semantic Scholar
  3. OpenAlex (all locations array)
  4. Europe PMC
  5. Sci-Hub (improved protocol-relative URL handling)
  6. Direct DOI / VPN + HTML parsing

On completion writes:
  - failed_downloads.csv  (DOIs that failed with per-layer reason)
  - download_log.csv      (all results with source used)
"""

import os
import re
import csv
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
    "https://sci-hub.mksa.top",
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


def make_absolute(url: str, base: str) -> str | None:
    """Handle http://, https://, and protocol-relative // URLs."""
    if not url:
        return None
    if url.startswith("//"):
        parsed_base = urlparse(base)
        url = f"{parsed_base.scheme}:{url}"
    full = urljoin(base, url)
    return full if urlparse(full).scheme in ("http", "https") else None


def fetch(url: str, timeout: int = 20) -> tuple[requests.Response | None, str]:
    """Returns (response, error_reason). Response is None on failure."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        if r.status_code == 200:
            return r, ""
        return None, f"HTTP {r.status_code}"
    except requests.exceptions.Timeout:
        return None, "timeout"
    except requests.exceptions.ConnectionError:
        return None, "connection error"
    except Exception as e:
        return None, str(e)


def is_pdf(resp: requests.Response) -> bool:
    ct = resp.headers.get("Content-Type", "")
    return "pdf" in ct.lower()


def pdf_from_response(resp: requests.Response) -> bytes | None:
    if is_pdf(resp) and len(resp.content) > 5_000:
        return resp.content
    return None


def try_download_url(url: str, timeout: int = 30) -> tuple[bytes | None, str]:
    """Download a URL, return (bytes, reason). bytes is None on failure."""
    resp, err = fetch(url, timeout=timeout)
    if not resp:
        return None, err
    pdf = pdf_from_response(resp)
    if pdf:
        return pdf, ""
    ct = resp.headers.get("Content-Type", "unknown")
    return None, f"not a PDF (Content-Type: {ct}, size: {len(resp.content)})"


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

def layer_unpaywall(doi: str) -> tuple[bytes | None, str]:
    """Check ALL oa_locations, not just best_oa_location."""
    url = f"https://api.unpaywall.org/v2/{doi}?email={UNPAYWALL_EMAIL}"
    try:
        r, err = fetch(url, timeout=15)
        if not r:
            return None, f"API error: {err}"
        data = r.json()
        if not data.get("is_oa"):
            return None, "not OA"
        # Collect all candidate PDF URLs from every location
        locations = data.get("oa_locations") or []
        best = data.get("best_oa_location")
        if best:
            locations = [best] + [l for l in locations if l != best]
        for loc in locations:
            pdf_url = loc.get("url_for_pdf") or loc.get("url")
            if not pdf_url:
                continue
            pdf, reason = try_download_url(pdf_url)
            if pdf:
                return pdf, ""
        return None, "OA locations found but no PDF downloadable"
    except Exception as e:
        return None, str(e)


def layer_semantic_scholar(doi: str) -> tuple[bytes | None, str]:
    url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}?fields=openAccessPdf"
    try:
        r, err = fetch(url, timeout=15)
        if not r:
            return None, f"API error: {err}"
        data = r.json()
        oa = data.get("openAccessPdf") or {}
        pdf_url = oa.get("url")
        if not pdf_url:
            return None, "no OA PDF in response"
        pdf, reason = try_download_url(pdf_url)
        return (pdf, "") if pdf else (None, reason)
    except Exception as e:
        return None, str(e)


def layer_openalex(doi: str) -> tuple[bytes | None, str]:
    """Check open_access.oa_url AND every entry in the locations array."""
    url = f"https://api.openalex.org/works/https://doi.org/{doi}"
    try:
        r, err = fetch(url, timeout=15)
        if not r:
            return None, f"API error: {err}"
        data = r.json()

        candidate_urls = []
        oa = data.get("open_access") or {}
        if oa.get("oa_url"):
            candidate_urls.append(oa["oa_url"])

        # locations array — each may have a pdf_url
        for loc in data.get("locations") or []:
            pu = loc.get("pdf_url")
            lu = loc.get("landing_page_url")
            if pu:
                candidate_urls.append(pu)
            elif lu and loc.get("is_oa"):
                candidate_urls.append(lu)

        if not candidate_urls:
            return None, "no OA URLs in response"

        for cu in candidate_urls:
            pdf, reason = try_download_url(cu)
            if pdf:
                return pdf, ""
        return None, "URLs found but no PDF downloadable"
    except Exception as e:
        return None, str(e)


def layer_europe_pmc(doi: str) -> tuple[bytes | None, str]:
    search_url = (
        f"https://www.ebi.ac.uk/europepmc/webservices/rest/search"
        f"?query=DOI:{doi}&format=json&resultType=core"
    )
    try:
        r, err = fetch(search_url, timeout=15)
        if not r:
            return None, f"API error: {err}"
        results = r.json().get("resultList", {}).get("result", [])
        if not results:
            return None, "not found in Europe PMC"
        pmcid = results[0].get("pmcid")
        if not pmcid:
            return None, "found but no PMCID (not OA)"
        pdf_url = f"https://europepmc.org/articles/{pmcid}/pdf/render"
        pdf, reason = try_download_url(pdf_url)
        return (pdf, "") if pdf else (None, reason)
    except Exception as e:
        return None, str(e)


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
            full = make_absolute(m, base_url)
            if full:
                return full
    return None


def layer_scihub(doi: str) -> tuple[bytes | None, str]:
    """Try each mirror; handle protocol-relative (//host/path) iframe URLs."""
    last_reason = "no mirrors reachable"
    for mirror in SCIHUB_MIRRORS:
        try:
            resp, err = fetch(f"{mirror}/{doi}", timeout=30)
            if not resp:
                last_reason = f"{mirror}: {err}"
                continue

            # Sci-Hub serves PDF in an iframe/embed whose src may be:
            #   //cdn.../file.pdf  (protocol-relative)
            #   https://cdn.../file.pdf
            #   /tree/...pdf
            pdf_url = None
            for pattern in [
                r'<iframe[^>]+src=["\']([^"\']+)["\']',
                r'<embed[^>]+src=["\']([^"\']+)["\']',
                r'onclick=["\'][^"\']*location\.href\s*=\s*["\']([^"\']+)["\']',
            ]:
                for m in re.findall(pattern, resp.text, re.IGNORECASE):
                    candidate = make_absolute(m, resp.url)
                    if candidate:
                        pdf_url = candidate
                        break
                if pdf_url:
                    break

            # Fallback to generic HTML PDF patterns
            if not pdf_url:
                pdf_url = find_pdf_url_in_html(resp.text, resp.url)

            if pdf_url:
                pdf, reason = try_download_url(pdf_url, timeout=60)
                if pdf:
                    return pdf, ""
                last_reason = f"{mirror}: PDF URL found ({pdf_url}) but {reason}"
            else:
                last_reason = f"{mirror}: no PDF URL found in HTML"

        except Exception as e:
            last_reason = f"{mirror}: {e}"
            continue

    return None, last_reason


def layer_vpn_html(doi: str) -> tuple[bytes | None, str]:
    resp, err = fetch(f"https://doi.org/{doi}", timeout=30)
    if not resp:
        return None, f"doi.org unreachable: {err}"
    if is_pdf(resp):
        pdf = resp.content if len(resp.content) > 5_000 else None
        return (pdf, "") if pdf else (None, "response was PDF but too small")
    pdf_url = find_pdf_url_in_html(resp.text, resp.url)
    if not pdf_url:
        return None, f"landed on {resp.url} — no PDF link in HTML (JS-rendered?)"
    pdf, reason = try_download_url(pdf_url)
    return (pdf, "") if pdf else (None, f"PDF URL {pdf_url}: {reason}")


# ── Pipeline ───────────────────────────────────────────────────────────────────

LAYERS = [
    ("Unpaywall",        layer_unpaywall),
    ("Semantic Scholar", layer_semantic_scholar),
    ("OpenAlex",         layer_openalex),
    ("Europe PMC",       layer_europe_pmc),
    ("Sci-Hub",          layer_scihub),
    ("VPN/HTML",         layer_vpn_html),
]


def download_pdf(doi: str) -> tuple[bytes | None, str | None, dict]:
    """Returns (bytes, source_name, {layer: reason}) — reasons only for failures."""
    reasons = {}
    for name, fn in LAYERS:
        try:
            pdf, reason = fn(doi)
            if pdf:
                return pdf, name, reasons
            reasons[name] = reason or "no PDF"
        except Exception as e:
            reasons[name] = f"exception: {e}"
    return None, None, reasons


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    # Allow running only on failed_downloads.csv if it exists
    retry_mode = os.path.exists("failed_downloads.csv") and not any(
        os.path.exists(os.path.join(f, x))
        for f in [FUTURE_FOLDER, "HIGH relevance", "MEDIUM screening"]
        for x in [""] if os.path.isdir(os.path.join(".", f))
    )

    if retry_mode:
        print("Detected failed_downloads.csv — running in RETRY mode (only failed DOIs).")
        import csv as _csv
        with open("failed_downloads.csv", encoding="utf-8") as f:
            reader = _csv.DictReader(f)
            articles = [{"doi": row["doi"], "folder": row["folder"]} for row in reader]
    else:
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
    failed_list = []
    log_rows = []

    for idx, art in enumerate(articles, start=1):
        doi    = art["doi"]
        folder = art["folder"]
        path   = os.path.join(folder, sanitize_filename(doi) + ".pdf")

        print(f"[{idx}/{total}] {doi}")

        if os.path.exists(path):
            print(f"  [SKIP] Already exists.")
            skipped += 1
            log_rows.append({"doi": doi, "folder": folder, "status": "skip", "source": "", "reasons": ""})
            continue

        pdf_bytes, source, reasons = download_pdf(doi)

        if pdf_bytes:
            with open(path, "wb") as f:
                f.write(pdf_bytes)
            print(f"  [OK] {source} → {path}")
            success += 1
            log_rows.append({"doi": doi, "folder": folder, "status": "ok", "source": source, "reasons": ""})
        else:
            reason_str = " | ".join(f"{k}: {v}" for k, v in reasons.items())
            print(f"  [FAIL] {reason_str}")
            failed_list.append({"doi": doi, "folder": folder,
                                 "doi_url": f"https://doi.org/{doi}",
                                 "reasons": reason_str})
            log_rows.append({"doi": doi, "folder": folder, "status": "fail", "source": "", "reasons": reason_str})
            fail += 1

        time.sleep(DELAY_SECONDS)

    print(f"\n{'='*60}")
    print(f"Done.  OK: {success}  |  Failed: {fail}  |  Skipped: {skipped}  |  Total: {total}")
    for folder in sorted(folders_needed):
        if os.path.isdir(folder):
            count = len([f for f in os.listdir(folder) if f.endswith(".pdf")])
            print(f"  {folder}/  → {count} PDFs")

    import csv
    if failed_list:
        with open("failed_downloads.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["doi", "folder", "doi_url", "reasons"])
            writer.writeheader()
            writer.writerows(failed_list)
        print(f"\nFailed DOIs → failed_downloads.csv ({len(failed_list)} entries)")

    with open("download_log.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["doi", "folder", "status", "source", "reasons"])
        writer.writeheader()
        writer.writerows(log_rows)
    print(f"Full log    → download_log.csv")


if __name__ == "__main__":
    main()
