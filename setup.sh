#!/bin/bash
# ============================================================
# OpenCode Hackathon Setup Script — MACOS VERSION (UNTESTED)
# Repo: EU_CivicTech_team
#
# WARNING: this script has NOT been tested on a real Mac.
# It mirrors the validated Windows setup.ps1 logic but may
# need fixes. If you hit an error, fix it locally and please
# share the correction with the team.
#
# Prerequisites (do these first, manually):
#   1. Install OpenCode Desktop from opencode.ai (Mac version)
#   2. Get your own DeepSeek API key from platform.deepseek.com
#   3. Clone this repo and cd into it:
#        gh repo clone sergiofspedro/EU_CivicTech_team
#        cd EU_CivicTech_team
#   4. Make this script executable and run it:
#        chmod +x setup.sh
#        ./setup.sh
# ============================================================

set -e

echo "=== Step 1: Install bun (required by ocx) ==="
curl -fsSL https://bun.sh/install | bash
echo "IMPORTANT: close and reopen your terminal now, then re-run this script."
echo "(bun needs a fresh shell session to be on PATH)"
read -p "Already restarted your terminal and 'bun --version' works? (y/n) " continue_answer
if [ "$continue_answer" != "y" ]; then
    echo "Please restart your terminal, cd back into the repo, and re-run ./setup.sh"
    exit 1
fi

echo ""
echo "=== Step 2: Install ocx (extension manager) ==="
npm install -g ocx

echo ""
echo "=== Step 3: Initialize ocx (global + local) ==="
ocx init --global
ocx init

echo ""
echo "=== Step 4: Install Oh My OpenCode (Ultimate) ==="
npx oh-my-openagent install

echo ""
echo "=== Step 5: Install kdco/workspace bundle (worktree + notify + delegation + planning) ==="
ocx registry add https://registry.kdco.dev --name kdco
ocx add kdco/workspace
echo "Installing .opencode dependencies..."
(cd .opencode && npm install)

echo ""
echo "=== Step 6: Install Simple Memory ==="
rm -rf /tmp/simple-memory
git clone https://github.com/cnicolov/opencode-plugin-simple-memory.git /tmp/simple-memory
mkdir -p .opencode/plugins
rm -rf .opencode/plugins/simple-memory
cp -R /tmp/simple-memory .opencode/plugins/simple-memory
rm -rf .opencode/plugins/simple-memory/.git

echo ""
echo "=== Step 7: Install EnvSitter Guard ==="
rm -rf /tmp/envsitter-guard
git clone https://github.com/boxpositron/envsitter-guard.git /tmp/envsitter-guard
rm -rf .opencode/plugins/envsitter-guard
cp -R /tmp/envsitter-guard .opencode/plugins/envsitter-guard
rm -rf .opencode/plugins/envsitter-guard/.git

echo ""
echo "=== Step 8: Install opencode-firecrawl plugin + CLI ==="
rm -rf /tmp/opencode-firecrawl
git clone https://github.com/firecrawl/opencode-firecrawl.git /tmp/opencode-firecrawl
rm -rf .opencode/plugins/firecrawl
cp -R /tmp/opencode-firecrawl .opencode/plugins/firecrawl
rm -rf .opencode/plugins/firecrawl/.git
npm install -g firecrawl-cli
echo "ACTION NEEDED: run 'firecrawl login --browser' OR set your own FIRECRAWL_API_KEY (firecrawl.dev)"

echo ""
echo "=== Step 9: Playwright MCP ==="
echo "Already in opencode.json (committed to repo) — no action needed"

echo ""
echo "=== Step 10: Composio MCP ==="
echo "ACTION NEEDED:"
echo "  1. Create your own account at composio.dev"
echo "  2. Get your API key from the dashboard"
echo "  3. NEVER paste your key in chat/Slack - set it as an env var:"
echo '       export COMPOSIO_API_KEY="your-key-here"'
echo "  4. Inside OpenCode, run /connect and select Composio"

echo ""
echo "=== Step 11: Context7 ==="
echo "Included free inside Oh My OpenCode — no separate install"

echo ""
echo "=== DONE - now launch OpenCode Desktop ==="
echo "Open the OpenCode Desktop app, use 'Open Folder' on this repo folder, then run:"
echo "  /connect   -> enter YOUR OWN DeepSeek API key"
echo "  /init      -> scan repo, merge with shared AGENTS.md"
echo ""
echo "Known non-blocking issue: /init may show a one-time Playwright JS parse error during repo scan. Safe to ignore."
