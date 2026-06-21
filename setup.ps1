# ============================================================
# OpenCode Hackathon Setup Script â€” TEAM VERSION (tested)
# Repo: EU_CivicTech_team
#
# Prerequisites (do these first, manually):
#   1. Install OpenCode Desktop from opencode.ai
#   2. Get your own DeepSeek API key from platform.deepseek.com
#   3. Clone this repo and cd into it:
#        gh repo clone sergiofspedro/EU_CivicTech_team
#        cd EU_CivicTech_team
#   4. Allow script execution for this session:
#        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   5. Run this script:  .\setup.ps1
# ============================================================

Write-Host "=== Step 1: Install bun (required by ocx) ===" -ForegroundColor Cyan
powershell -c "irm bun.sh/install.ps1 | iex"
Write-Host "IMPORTANT: close and reopen this terminal now, then re-run this script." -ForegroundColor Red
Write-Host "(bun needs a fresh terminal session to be on PATH)" -ForegroundColor Yellow
$continue = Read-Host "Already restarted your terminal and bun --version works? (y/n)"
if ($continue -ne "y") {
    Write-Host "Please restart your terminal, cd back into the repo, and re-run .\setup.ps1" -ForegroundColor Red
    exit
}

Write-Host "`n=== Step 2: Install ocx (extension manager) ===" -ForegroundColor Cyan
npm install -g ocx

Write-Host "`n=== Step 3: Initialize ocx (global + local) ===" -ForegroundColor Cyan
ocx init --global
ocx init

Write-Host "`n=== Step 4: Install Oh My OpenCode (Ultimate) ===" -ForegroundColor Cyan
npx oh-my-openagent install

Write-Host "`n=== Step 5: Install kdco/workspace bundle (worktree + notify + delegation + planning) ===" -ForegroundColor Cyan
ocx registry add https://registry.kdco.dev --name kdco
ocx add kdco/workspace
Write-Host "Installing .opencode dependencies..." -ForegroundColor Cyan
Push-Location .opencode
npm install
Pop-Location

Write-Host "`n=== Step 6: Install Simple Memory ===" -ForegroundColor Cyan
if (Test-Path "$env:TEMP\simple-memory") { Remove-Item "$env:TEMP\simple-memory" -Recurse -Force }
git clone https://github.com/cnicolov/opencode-plugin-simple-memory.git $env:TEMP\simple-memory
New-Item -ItemType Directory -Force -Path ".opencode\plugins" | Out-Null
if (Test-Path ".opencode\plugins\simple-memory") { Remove-Item ".opencode\plugins\simple-memory" -Recurse -Force }
robocopy "$env:TEMP\simple-memory" ".opencode\plugins\simple-memory" /E /NFL /NDL /NJH /NJS
Remove-Item ".opencode\plugins\simple-memory\.git" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n=== Step 7: Install EnvSitter Guard ===" -ForegroundColor Cyan
if (Test-Path "$env:TEMP\envsitter-guard") { Remove-Item "$env:TEMP\envsitter-guard" -Recurse -Force }
git clone https://github.com/boxpositron/envsitter-guard.git $env:TEMP\envsitter-guard
if (Test-Path ".opencode\plugins\envsitter-guard") { Remove-Item ".opencode\plugins\envsitter-guard" -Recurse -Force }
robocopy "$env:TEMP\envsitter-guard" ".opencode\plugins\envsitter-guard" /E /NFL /NDL /NJH /NJS
Remove-Item ".opencode\plugins\envsitter-guard\.git" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n=== Step 8: Install opencode-firecrawl plugin + CLI ===" -ForegroundColor Cyan
if (Test-Path "$env:TEMP\opencode-firecrawl") { Remove-Item "$env:TEMP\opencode-firecrawl" -Recurse -Force }
git clone https://github.com/firecrawl/opencode-firecrawl.git $env:TEMP\opencode-firecrawl
if (Test-Path ".opencode\plugins\firecrawl") { Remove-Item ".opencode\plugins\firecrawl" -Recurse -Force }
robocopy "$env:TEMP\opencode-firecrawl" ".opencode\plugins\firecrawl" /E /NFL /NDL /NJH /NJS
Remove-Item ".opencode\plugins\firecrawl\.git" -Recurse -Force -ErrorAction SilentlyContinue
npm install -g firecrawl-cli
Write-Host "ACTION NEEDED: run 'firecrawl login --browser' OR set your own FIRECRAWL_API_KEY (firecrawl.dev)" -ForegroundColor Yellow

Write-Host "`n=== Step 9: Playwright MCP ===" -ForegroundColor Cyan
Write-Host "Already in opencode.json (committed to repo) â€” no action needed" -ForegroundColor Green

Write-Host "`n=== Step 10: Composio MCP ===" -ForegroundColor Cyan
Write-Host "ACTION NEEDED:" -ForegroundColor Yellow
Write-Host "  1. Create your own account at composio.dev"
Write-Host "  2. Get your API key from the dashboard"
Write-Host "  3. NEVER paste your key in chat/Slack - set it as an env var:"
Write-Host '       $env:COMPOSIO_API_KEY = "your-key-here"' -ForegroundColor Yellow
Write-Host "  4. Inside OpenCode, run /connect and select Composio"

Write-Host "`n=== Step 11: Context7 ===" -ForegroundColor Cyan
Write-Host "Included free inside Oh My OpenCode â€” no separate install" -ForegroundColor Green

Write-Host "`n=== DONE - now launch OpenCode Desktop ===" -ForegroundColor Green
Write-Host "Open the OpenCode Desktop app, use 'Open Folder' on this repo folder, then run:"
Write-Host "  /connect   -> enter YOUR OWN DeepSeek API key"
Write-Host "  /init      -> scan repo, merge with shared AGENTS.md"
Write-Host ""
Write-Host "Known non-blocking issue: /init may show a one-time Playwright JS parse error during repo scan. Safe to ignore." -ForegroundColor Yellow
