# ============================================================
# OpenCode Hackathon Setup Script — TEAM VERSION
# Repo: Livro-Redes-Democracia
#
# Prerequisites (do these first, manually):
#   1. Install OpenCode Desktop from opencode.ai
#   2. Get your own DeepSeek API key from platform.deepseek.com
#   3. Clone this repo and cd into it:
#        gh repo clone sergiofspedro/Livro-Redes-Democracia
#        cd Livro-Redes-Democracia
#   4. Run this script:  .\setup.ps1
# ============================================================

Write-Host "=== Step 1: Install ocx (extension manager) ===" -ForegroundColor Cyan
Write-Host "Required for opencode-notify and opencode-worktree"
npm install -g ocx

Write-Host "`n=== Step 2: Initialize ocx ===" -ForegroundColor Cyan
ocx init --global

Write-Host "`n=== Step 3: Install Oh My OpenCode (Ultimate) ===" -ForegroundColor Cyan
npx oh-my-openagent install

Write-Host "`n=== Step 4: Install opencode-worktree ===" -ForegroundColor Cyan
ocx add worktree --from kdcokenny

Write-Host "`n=== Step 5: Install opencode-notify ===" -ForegroundColor Cyan
ocx add notify --from kdcokenny

Write-Host "`n=== Step 6: Install Simple Memory ===" -ForegroundColor Cyan
git clone https://github.com/cnicolov/opencode-plugin-simple-memory.git $env:TEMP\simple-memory
New-Item -ItemType Directory -Force -Path ".opencode\plugins" | Out-Null
Copy-Item "$env:TEMP\simple-memory\*" ".opencode\plugins\simple-memory\" -Recurse -Force

Write-Host "`n=== Step 7: Install EnvSitter Guard ===" -ForegroundColor Cyan
git clone https://github.com/boxpositron/envsitter-guard.git $env:TEMP\envsitter-guard
Copy-Item "$env:TEMP\envsitter-guard\*" ".opencode\plugins\envsitter-guard\" -Recurse -Force

Write-Host "`n=== Step 8: Install opencode-firecrawl ===" -ForegroundColor Cyan
git clone https://github.com/firecrawl/opencode-firecrawl.git $env:TEMP\opencode-firecrawl
Copy-Item "$env:TEMP\opencode-firecrawl\*" ".opencode\plugins\firecrawl\" -Recurse -Force
Write-Host "ACTION NEEDED: run 'firecrawl login --browser' OR set your own FIRECRAWL_API_KEY (firecrawl.dev)" -ForegroundColor Yellow

Write-Host "`n=== Step 9: Playwright MCP ===" -ForegroundColor Cyan
Write-Host "Already in opencode.json (committed to repo) — no action needed" -ForegroundColor Green

Write-Host "`n=== Step 10: Composio MCP ===" -ForegroundColor Cyan
Write-Host "ACTION NEEDED: create your own account at composio.dev, then inside OpenCode run /connect" -ForegroundColor Yellow

Write-Host "`n=== Step 11: Context7 ===" -ForegroundColor Cyan
Write-Host "Included free inside Oh My OpenCode — no separate install" -ForegroundColor Green

Write-Host "`n=== DONE - now launch OpenCode ===" -ForegroundColor Green
Write-Host "Run: opencode"
Write-Host "Then inside the app:"
Write-Host "  /connect   -> enter YOUR OWN DeepSeek API key"
Write-Host "  /init      -> scan repo, merge with shared AGENTS.md"
