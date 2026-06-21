# EU_CivicTech_team

EU Civic Tech Hackathon project â€” 22â€“23 June 2026.

## Team Setup â€” OpenCode Environment

This project uses [OpenCode Desktop](https://opencode.ai) as the shared AI coding agent, with a curated stack of plugins and MCPs for the hackathon.

### 1. Install OpenCode Desktop
Download from [opencode.ai](https://opencode.ai).

### 2. Get your own API key
Each team member needs their own DeepSeek API key from [platform.deepseek.com](https://platform.deepseek.com).
(Other providers also work â€” ask in the team channel before switching the project default.)

### 3. Clone this repo
```powershell
gh repo clone sergiofspedro/EU_CivicTech_team
cd EU_CivicTech_team
```

### 4. Allow script execution (this terminal session only)
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### 5. Run the setup script
```powershell
.\setup.ps1
```
This installs the full shared stack:
- **Oh My OpenCode** â€” multi-agent orchestration
- **kdco/workspace bundle** â€” worktree, notify, delegation, planning agents
- **Simple Memory** â€” git-committed shared team memory
- **EnvSitter Guard** â€” blocks agents from reading/editing `.env` secrets
- **opencode-firecrawl** â€” web scraping/crawling (e.g. EU open data sources)
- **Playwright MCP** â€” browser automation/testing
- **Context7 + Exa + gh_grep** â€” bundled free with Oh My OpenCode

**Important:** the script installs `bun` first and will ask you to restart your terminal partway through â€” this is expected, just restart, `cd` back into the repo, and re-run `.\setup.ps1`.

### 6. Manual steps after the script
Two things need one-time setup outside the script:

**Firecrawl:**
```powershell
firecrawl login --browser
```
or set your own key: `$env:FIRECRAWL_API_KEY = "your-key"` (get one at firecrawl.dev)

**Composio:**
1. Create your own free account at [composio.dev](https://composio.dev) (20K tool calls/month free, plenty for the hackathon)
2. Get your API key from the dashboard
3. **Never paste your API key in chat or Slack.** Set it as an environment variable instead:
   ```powershell
   $env:COMPOSIO_API_KEY = "your-key-here"
   ```
4. Inside OpenCode, run `/connect` and select Composio

### 7. Launch OpenCode and connect
Open the OpenCode Desktop app, use **Open Folder** to point it at this repo folder, then in the chat bar:
```
/connect     -> enter YOUR OWN DeepSeek API key
/init        -> scans the repo, merges with shared AGENTS.md rules
```

**Known issue (safe to ignore):** `/init` may show a one-time Playwright `SyntaxError` during the repo scan. It's non-blocking and doesn't affect the rest of the session.

### Project rules (automatic â€” see AGENTS.md)
- The agent will ask clarifying questions before ambiguous or non-trivial tasks
- The agent will propose alternative approaches before committing to one
- Local commits are fine; **pushing, opening PRs, or merging always require your explicit confirmation first**

### Shared vs personal config
- `opencode.json` (committed) â€” shared MCPs and agent permissions. Don't edit unless agreed with the team.
- `opencode.local.json` (gitignored) â€” your own personal overrides/extra instructions, optional, never shared.

### Security reminder
Never commit or paste API keys anywhere in chat, Slack, or this repo. All keys go in environment variables (`$env:VAR_NAME = "..."`) or the app's own `/connect` prompts.
