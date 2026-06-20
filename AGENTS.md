# Livro-Redes-Democracia

## What this repo is

A single Python script that batch-downloads academic PDFs from DOIs listed in an Excel file, using Unpaywall (open access) then direct DOI resolution via institutional VPN.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\Activate
pip install -r requirements.txt
```

Dependencies: `pandas`, `openpyxl`, `requests` (pinned in [requirements.txt](requirements.txt)). The script explicitly uses `pd.read_excel(..., engine="openpyxl")`, so `openpyxl` is mandatory, not optional.

## Operation

1. Place `artigos.xlsx` in the repo root with a column named `DOI`
2. Edit `UNPAYWALL_EMAIL` in [download_pdfs.py:18](download_pdfs.py#L18) with your email
3. PDFs land in `PDFs_Baixados/` (auto-created)

```powershell
python download_pdfs.py
```

The script skips already-downloaded files (match by DOI → sanitized filename). `DELAY_SECONDS` (default 2, line 19) rate-limits requests between every DOI.

## Key facts for agents

- **No tests, no type checker, no linter config, no formatter** — zero tooling beyond pip
- **Single entrypoint**: [download_pdfs.py](download_pdfs.py)
- **Two download layers**: Unpaywall (works without VPN) → VPN / direct DOI resolution (requires university VPN connection)
- **`openpyxl` engine required**: `pd.read_excel(..., engine="openpyxl")` — do not switch to `xlrd`/`calamine` without testing
- **`.gitignore` exists** but only ignores `opencode.local.json` and `.opencode/local/` — `artigos.xlsx` and `PDFs_Baixados/` are **not** gitignored by default. Be aware when making commits.
- **`opencode.json` loads instructions** from `./tools/philosophy.md` (currently empty). No other shared instruction files are loaded.
- **Repo-local OpenCode config**: 5 project skills under `.opencode/skills/` (`code-philosophy`, `code-review`, `frontend-philosophy`, `plan-protocol`, `plan-review`). The `opencode.json` also defines restricted agents (`researcher`, `scribe`, `coder`, `reviewer`, etc.) with granular bash permissions.

## Download behaviour notes

- [`clean_doi()`](download_pdfs.py#L31) strips URL prefixes (`https://doi.org/`, etc.) and whitespace before processing
- [`sanitize_filename()`](download_pdfs.py#L41) replaces `\ / : * ? " < > |` with `_` for safe filenames
- [`download_pdf()`](download_pdfs.py#L81) accepts a response as PDF if `Content-Type` contains `pdf` **or** the body exceeds 10 KB — catches misconfigured servers

## Git

- Commits are fine to make freely
- **NEVER push, open a PR, or merge** without explicit confirmation
