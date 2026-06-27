"""
search_grey_literature.py
Searches grey literature across multiple sources using cluster keyword sets,
downloads PDFs, deduplicates results, and exports a single Excel sheet.

Sources: CORE, OpenAlex, Semantic Scholar (incl. SSRN), arXiv, BASE
Output:  grey_literature/ folder + grey_literature.xlsx

Requirements: pip install requests openpyxl
VPN: run with university VPN active for maximum PDF access.
"""

import os
import re
import csv
import time
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from urllib.parse import urljoin, urlparse, quote_plus
import xml.etree.ElementTree as ET

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
CORE_API_KEY    = "pcjCa9P8V74fHKlGNT503RqD6nLSzwZg"
UNPAYWALL_EMAIL = "sergiopedro@ua.pt"
SCOPUS_EXCEL    = "Nets4Dem_ScopusReview_Final_GA_12 June.xlsx"
SCOPUS_SHEETS   = ["HIGH relevance", "MEDIUM screening"]
OUTPUT_FOLDER   = "grey literature"
OUTPUT_EXCEL    = "grey_literature.xlsx"
MIN_YEAR        = 2005
MAX_RESULTS_PER_SOURCE_PER_CLUSTER = 50   # cap per source per cluster
DELAY_SECONDS   = 1
# ───────────────────────────────────────────────────────────────────────────────

# ─── CLUSTER KEYWORD SETS ──────────────────────────────────────────────────────
# Each cluster: list of OR-groups. A result must match at least one term from
# each group. Simplified from the Scopus queries for API compatibility.
CLUSTERS = {
    "Cluster 1 – Network types": {
        "group1": [
            "network of networks", "multi-actor network", "civil society hub",
            "civic ecosystem", "civic alliance", "civic platform",
            "hub organisation", "umbrella organisation", "networked organisation",
            "network formalisation",
        ],
        "group2": [
            "civil society", "democratic governance", "democratic innovation",
            "participatory democracy", "deliberation",
        ],
    },
    "Cluster 2 – Democratic intensification": {
        "group1": [
            "democratic intensification", "democratic densification",
            "shrinking civic space", "democratic backsliding",
            "democratic erosion", "de-democratization", "autocratization",
            "illiberal turn", "democratic resilience", "democratic breakdown",
        ],
        "group2": [
            "civil society", "network", "alliance", "civic organisation",
        ],
    },
    "Cluster 3 – Network resilience": {
        "group1": [
            "network resilience", "network formation", "network transformation",
            "network governance", "collaborative governance", "brokerage",
            "bridging organisations", "structural holes", "knowledge brokering",
            "co-opetition", "coopetition", "collaborative competition",
        ],
        "group2": [
            "civil society", "democratic", "civic", "participation",
        ],
    },
    "Cluster 4 – Internal governance": {
        "group1": [
            "network membership", "civil society funding", "legal form",
            "network institutionalisation", "internal governance",
            "variable geometry", "network sustainability", "informal network",
            "network formalization", "funding dependency", "governance structure",
        ],
        "group2": [
            "civil society", "democratic", "civic",
        ],
    },
    "Cluster 5 – Transnational networks Europe": {
        "group1": [
            "transnational civil society", "European civil society",
            "cross-border network", "multi-level governance",
            "pan-European network", "EU-level civil society",
            "Open Government Partnership", "civic coalition",
        ],
        "group2": [
            "Europe", "European", "EU",
        ],
    },
    "Cluster 6 – Grey literature signals": {
        "group1": [
            "civic space monitoring", "civil society index",
            "everyday democracy index", "network impact evaluation",
            "CIVICUS monitor", "network self-assessment",
            "democratic innovation report", "NGO sustainability index",
            "civic participation report", "civil society shrinking",
        ],
        "group2": [],   # no second group needed — terms are specific enough
    },
}
# ───────────────────────────────────────────────────────────────────────────────

BASE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Accept-Language": "en-US,en;q=0.5",
}

SESSION = requests.Session()
SESSION.headers.update(BASE_HEADERS)


# ── Helpers ────────────────────────────────────────────────────────────────────

def normalise_doi(raw: str) -> str | None:
    if not raw:
        return None
    doi = str(raw).strip().lower()
    for prefix in ("https://doi.org/", "http://doi.org/", "doi.org/", "doi:", "doi "):
        if doi.startswith(prefix):
            doi = doi[len(prefix):]
    doi = doi.strip()
    return doi if doi and doi != "nan" else None


def normalise_title(title: str) -> str:
    return re.sub(r'\W+', ' ', str(title).lower()).strip()


def sanitize_filename(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", s)[:180]


def cluster_query_string(cluster: dict, max_terms: int = 6) -> str:
    """Build a simple AND query string from cluster groups."""
    parts = []
    for grp_terms in cluster.values():
        if grp_terms:
            sample = grp_terms[:max_terms]
            parts.append(" OR ".join(f'"{t}"' for t in sample))
    return " AND ".join(f"({p})" for p in parts)


def try_download_pdf(url: str, referer: str = "") -> bytes | None:
    if not url:
        return None
    try:
        hdrs = {**BASE_HEADERS, "Accept": "application/pdf,*/*", "Referer": referer}
        r = SESSION.get(url, headers=hdrs, timeout=40, allow_redirects=True)
        if r.status_code == 200 and "pdf" in r.headers.get("Content-Type", "").lower():
            if len(r.content) > 5_000:
                return r.content
    except Exception:
        pass
    return None


def save_pdf(pdf_bytes: bytes, doi: str, title: str) -> str:
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    key = sanitize_filename(doi if doi else normalise_title(title)[:80])
    path = os.path.join(OUTPUT_FOLDER, key + ".pdf")
    with open(path, "wb") as f:
        f.write(pdf_bytes)
    return path


# ── Load existing Scopus DOIs (to exclude duplicates) ─────────────────────────

def load_scopus_dois(excel_file: str) -> set:
    existing = set()
    if not os.path.exists(excel_file):
        return existing
    try:
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        for sheet in SCOPUS_SHEETS:
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            headers = [c.value for c in ws[1]]
            doi_idx = next((i for i, h in enumerate(headers)
                            if h and str(h).strip().upper() == "DOI"), None)
            if doi_idx is None:
                continue
            for row in ws.iter_rows(min_row=2, values_only=True):
                d = normalise_doi(row[doi_idx])
                if d:
                    existing.add(d)
    except Exception as e:
        print(f"  [WARN] Could not read Scopus Excel: {e}")
    return existing


# ── Search functions ───────────────────────────────────────────────────────────

def search_core(query: str, cluster_name: str) -> list[dict]:
    results = []
    url = "https://api.core.ac.uk/v3/search/works"
    params = {
        "q": query,
        "limit": MAX_RESULTS_PER_SOURCE_PER_CLUSTER,
        "offset": 0,
        "yearFrom": MIN_YEAR,
    }
    hdrs = {**BASE_HEADERS, "Authorization": f"Bearer {CORE_API_KEY}"}
    try:
        r = SESSION.get(url, headers=hdrs, params=params, timeout=20)
        if r.status_code != 200:
            return results
        data = r.json()
        for item in data.get("results") or []:
            year = item.get("yearPublished") or 0
            if year and int(year) < MIN_YEAR:
                continue
            doi = normalise_doi(item.get("doi") or "")
            results.append({
                "doi": doi or "",
                "title": item.get("title") or "",
                "year": year,
                "authors": "; ".join(
                    (a.get("name") or "") for a in (item.get("authors") or [])
                ),
                "source_db": "CORE",
                "cluster": cluster_name,
                "abstract": (item.get("abstract") or "")[:500],
                "landing_url": item.get("sourceFulltextUrls", [None])[0] or item.get("downloadUrl") or "",
                "pdf_url": item.get("downloadUrl") or "",
            })
    except Exception as e:
        print(f"    [CORE error] {e}")
    return results


def search_openalex(query: str, cluster_name: str) -> list[dict]:
    results = []
    url = "https://api.openalex.org/works"
    params = {
        "search": query,
        "filter": f"publication_year:>{MIN_YEAR - 1},type:article|preprint|report|book-chapter",
        "per-page": MAX_RESULTS_PER_SOURCE_PER_CLUSTER,
        "select": "doi,title,publication_year,authorships,primary_location,open_access,abstract_inverted_index",
    }
    try:
        r = SESSION.get(url, params=params, timeout=20)
        if r.status_code != 200:
            return results
        for item in r.json().get("results") or []:
            doi = normalise_doi(item.get("doi") or "")
            oa = item.get("open_access") or {}
            loc = item.get("primary_location") or {}
            authors = "; ".join(
                a.get("author", {}).get("display_name", "") or ""
                for a in (item.get("authorships") or [])[:5]
            )
            # reconstruct abstract from inverted index
            inv = item.get("abstract_inverted_index") or {}
            if inv:
                words = [""] * (max(max(v) for v in inv.values()) + 1)
                for word, positions in inv.items():
                    for pos in positions:
                        words[pos] = word
                abstract = " ".join(words)[:500]
            else:
                abstract = ""
            results.append({
                "doi": doi or "",
                "title": item.get("title") or "",
                "year": item.get("publication_year") or "",
                "authors": authors,
                "source_db": "OpenAlex",
                "cluster": cluster_name,
                "abstract": abstract,
                "landing_url": loc.get("landing_page_url") or "",
                "pdf_url": oa.get("oa_url") or loc.get("pdf_url") or "",
            })
    except Exception as e:
        print(f"    [OpenAlex error] {e}")
    return results


def search_semantic_scholar(query: str, cluster_name: str) -> list[dict]:
    results = []
    url = "https://api.semanticscholar.org/graph/v1/paper/search"
    params = {
        "query": query,
        "limit": MAX_RESULTS_PER_SOURCE_PER_CLUSTER,
        "fields": "externalIds,title,year,authors,openAccessPdf,abstract,venue",
    }
    try:
        r = SESSION.get(url, params=params, timeout=20)
        if r.status_code != 200:
            return results
        for item in r.json().get("data") or []:
            year = item.get("year") or 0
            if year and int(year) < MIN_YEAR:
                continue
            ext = item.get("externalIds") or {}
            doi = normalise_doi(ext.get("DOI") or "")
            oa = item.get("openAccessPdf") or {}
            authors = "; ".join(
                a.get("name", "") for a in (item.get("authors") or [])[:5]
            )
            results.append({
                "doi": doi or "",
                "title": item.get("title") or "",
                "year": year,
                "authors": authors,
                "source_db": "Semantic Scholar / SSRN",
                "cluster": cluster_name,
                "abstract": (item.get("abstract") or "")[:500],
                "landing_url": f"https://www.semanticscholar.org/paper/{item.get('paperId', '')}",
                "pdf_url": oa.get("url") or "",
            })
    except Exception as e:
        print(f"    [Semantic Scholar error] {e}")
    return results


def search_arxiv(query: str, cluster_name: str) -> list[dict]:
    results = []
    # Restrict to relevant categories
    cats = "cat:cs.CY OR cat:cs.SI OR cat:econ.GN OR cat:econ.PE"
    full_query = f"({query}) AND ({cats})"
    url = "http://export.arxiv.org/api/query"
    params = {
        "search_query": f"all:{full_query}",
        "max_results": MAX_RESULTS_PER_SOURCE_PER_CLUSTER,
        "sortBy": "relevance",
    }
    try:
        r = SESSION.get(url, params=params, timeout=20,
                        headers={**BASE_HEADERS, "Accept": "application/xml"})
        if r.status_code != 200:
            return results
        ns = {"a": "http://www.w3.org/2005/Atom"}
        root = ET.fromstring(r.text)
        for entry in root.findall("a:entry", ns):
            year_raw = (entry.findtext("a:published", "", ns) or "")[:4]
            year = int(year_raw) if year_raw.isdigit() else 0
            if year and year < MIN_YEAR:
                continue
            doi_tag = entry.find('a:link[@title="doi"]', ns)
            doi = normalise_doi(doi_tag.attrib.get("href", "") if doi_tag is not None else "")
            arxiv_id = entry.findtext("a:id", "", ns).split("/abs/")[-1]
            pdf_url = f"https://arxiv.org/pdf/{arxiv_id}"
            authors = "; ".join(
                a.findtext("a:name", "", ns)
                for a in entry.findall("a:author", ns)[:5]
            )
            results.append({
                "doi": doi or "",
                "title": entry.findtext("a:title", "", ns).replace("\n", " ").strip(),
                "year": year,
                "authors": authors,
                "source_db": "arXiv",
                "cluster": cluster_name,
                "abstract": entry.findtext("a:summary", "", ns).replace("\n", " ").strip()[:500],
                "landing_url": f"https://arxiv.org/abs/{arxiv_id}",
                "pdf_url": pdf_url,
            })
    except Exception as e:
        print(f"    [arXiv error] {e}")
    return results


def search_base(query: str, cluster_name: str) -> list[dict]:
    """Bielefeld Academic Search Engine — rich in grey literature."""
    results = []
    url = "https://api.base-search.net/cgi-bin/BaseHttpSearchInterface.fcgi"
    params = {
        "func": "PerformSearch",
        "query": f"dcterms.description:({query})",
        "hits": MAX_RESULTS_PER_SOURCE_PER_CLUSTER,
        "offset": 0,
        "format": "json",
        "boost": "oa",
    }
    try:
        r = SESSION.get(url, params=params, timeout=20)
        if r.status_code != 200:
            return results
        data = r.json()
        docs = data.get("response", {}).get("docs") or []
        for item in docs:
            year_raw = str(item.get("dcyear") or "")[:4]
            year = int(year_raw) if year_raw.isdigit() else 0
            if year and year < MIN_YEAR:
                continue
            doi = normalise_doi(item.get("dcdoi") or "")
            links = item.get("dclink") or []
            pdf_url = next((l for l in links if ".pdf" in l.lower()), "")
            landing = links[0] if links else ""
            authors_raw = item.get("dcauthor") or []
            authors = "; ".join(authors_raw[:5]) if isinstance(authors_raw, list) else str(authors_raw)
            results.append({
                "doi": doi or "",
                "title": (item.get("dctitle") or [""])[0] if isinstance(item.get("dctitle"), list)
                         else str(item.get("dctitle") or ""),
                "year": year,
                "authors": authors,
                "source_db": "BASE",
                "cluster": cluster_name,
                "abstract": str(item.get("dcdescription") or "")[:500],
                "landing_url": landing,
                "pdf_url": pdf_url,
            })
    except Exception as e:
        print(f"    [BASE error] {e}")
    return results


SEARCH_FUNCTIONS = [
    ("CORE",             search_core),
    ("OpenAlex",         search_openalex),
    ("Semantic Scholar", search_semantic_scholar),
    ("arXiv",            search_arxiv),
    ("BASE",             search_base),
]


# ── Deduplication ──────────────────────────────────────────────────────────────

def deduplicate(records: list[dict], scopus_dois: set) -> list[dict]:
    seen_dois = set(scopus_dois)
    seen_titles = set()
    unique = []
    for rec in records:
        doi = rec.get("doi", "").strip()
        title_key = normalise_title(rec.get("title", ""))
        if doi and doi in seen_dois:
            continue
        if title_key and title_key in seen_titles:
            continue
        if doi:
            seen_dois.add(doi)
        if title_key:
            seen_titles.add(title_key)
        unique.append(rec)
    return unique


# ── Excel output ───────────────────────────────────────────────────────────────

EXCEL_COLUMNS = [
    "title", "authors", "year", "doi", "cluster",
    "source_db", "abstract", "landing_url", "pdf_url", "pdf_downloaded",
]

def write_excel(records: list[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "grey literature"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    headers_display = [
        "Title", "Authors", "Year", "DOI", "Cluster",
        "Source DB", "Abstract", "Landing URL", "PDF URL", "PDF Downloaded",
    ]
    ws.append(headers_display)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for rec in records:
        ws.append([rec.get(c, "") for c in EXCEL_COLUMNS])

    col_widths = [60, 40, 6, 35, 30, 20, 80, 60, 60, 15]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    wb.save(OUTPUT_EXCEL)
    print(f"\nExcel saved: {os.path.abspath(OUTPUT_EXCEL)}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print(f"Loading existing Scopus DOIs from '{SCOPUS_EXCEL}' ...")
    scopus_dois = load_scopus_dois(SCOPUS_EXCEL)
    print(f"  {len(scopus_dois)} DOIs loaded (will be excluded from results).\n")

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    all_records = []

    for cluster_name, cluster in CLUSTERS.items():
        print(f"── {cluster_name} ──────────────────────────────────────────")
        query = cluster_query_string(cluster)
        print(f"  Query: {query[:120]}...")

        cluster_records = []
        for src_name, fn in SEARCH_FUNCTIONS:
            print(f"  Searching {src_name} ...", end=" ", flush=True)
            results = fn(query, cluster_name)
            print(f"{len(results)} results")
            cluster_records.extend(results)
            time.sleep(DELAY_SECONDS)

        all_records.extend(cluster_records)

    print(f"\nTotal raw results: {len(all_records)}")
    unique = deduplicate(all_records, scopus_dois)
    print(f"After deduplication (incl. removing Scopus hits): {len(unique)}")

    # Try to download PDFs
    print(f"\nAttempting PDF downloads into '{OUTPUT_FOLDER}/' ...")
    for i, rec in enumerate(unique, start=1):
        pdf_url = rec.get("pdf_url") or rec.get("landing_url") or ""
        doi = rec.get("doi", "")
        title = rec.get("title", f"record_{i}")
        filename_key = sanitize_filename(doi if doi else normalise_title(title)[:80])
        path = os.path.join(OUTPUT_FOLDER, filename_key + ".pdf")

        if os.path.exists(path):
            rec["pdf_downloaded"] = "YES (existing)"
            continue

        if not pdf_url:
            rec["pdf_downloaded"] = "NO - no URL"
            continue

        print(f"  [{i}/{len(unique)}] {title[:60]} ...", end=" ", flush=True)
        pdf = try_download_pdf(pdf_url, referer=f"https://doi.org/{doi}" if doi else pdf_url)

        # fallback: try Unpaywall if we have a DOI
        if not pdf and doi:
            try:
                r = SESSION.get(
                    f"https://api.unpaywall.org/v2/{doi}?email={UNPAYWALL_EMAIL}",
                    timeout=10
                )
                if r.status_code == 200:
                    data = r.json()
                    for loc in (data.get("oa_locations") or []):
                        u = loc.get("url_for_pdf") or loc.get("url")
                        if u:
                            pdf = try_download_pdf(u, referer=f"https://doi.org/{doi}")
                            if pdf:
                                break
            except Exception:
                pass

        if pdf:
            save_pdf(pdf, doi if doi else f"record_{i}", title)
            rec["pdf_downloaded"] = "YES"
            print("OK")
        else:
            rec["pdf_downloaded"] = "NO"
            print("FAIL")

        time.sleep(DELAY_SECONDS)

    write_excel(unique)

    downloaded = sum(1 for r in unique if r.get("pdf_downloaded", "").startswith("YES"))
    print(f"\n{'='*60}")
    print(f"Total unique results: {len(unique)}")
    print(f"PDFs downloaded:      {downloaded}")
    print(f"Output folder:        {os.path.abspath(OUTPUT_FOLDER)}")
    print(f"Output Excel:         {os.path.abspath(OUTPUT_EXCEL)}")


if __name__ == "__main__":
    main()
