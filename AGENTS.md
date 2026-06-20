# Livro-Redes-Democracia

## What this repo is

A single Python script that batch-downloads academic PDFs from DOIs listed in an Excel file, using Unpaywall (open access) then direct DOI resolution via institutional VPN.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\Activate
pip install -r requirements.txt
```

## Operation

1. Place `artigos.xlsx` in the repo root with a column named `DOI`
2. Edit `UNPAYWALL_EMAIL` in `download_pdfs.py:18` with your email
3. PDFs land in `PDFs_Baixados/` (auto-created)

```powershell
python download_pdfs.py
```

The script skips already-downloaded files (match by DOI → sanitized filename). `DELAY_SECONDS` (default 2) rate-limits requests.

## Key facts for agents

- No tests, no type checker, no linter config, no formatter
- Single entrypoint: `download_pdfs.py`
- Unpaywall layer works without VPN; VPN layer requires university VPN connection
- The `opencode.json` in this repo loads shared instructions from `../Claude/` — that external context is already available to OpenCode sessions
- `artigos.xlsx` and `PDFs_Baixados/` are both gitignored (`.gitignore` does not exist yet — create one if output files appear)

## Git

- Commits are fine to make freely
- NEVER push, open a PR, or merge without explicit confirmation
