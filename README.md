# EU_CivicTech_team

EU Civic Tech Hackathon project — 22–23 June 2026.

## Team Setup — OpenCode Environment

This project uses [OpenCode Desktop](https://opencode.ai) as the shared AI coding agent, with a curated stack of plugins and MCPs for the hackathon.

### 1. Install OpenCode Desktop
Download from [opencode.ai](https://opencode.ai).

### 2. Get your own OpenRouter API key
Each team member needs their own [OpenRouter](https://openrouter.ai) account and API key. OpenRouter gives access to multiple models through a single key, letting you switch models per task without re-configuring.

**Recommended model per work mode:**

| Mode | Model | Why |
|---|---|---|
| **Plan mode** (architecture, scoping, design decisions) | `moonshotai/kimi-k2.7-code` | Strong long-horizon planning, native multi-turn reasoning, good MCP tool-use depth |
| **Build mode** (routine implementation, boilerplate) | `deepseek/deepseek-v4-flash` | Very low cost per call, fast — best for high-volume, straightforward coding tasks |
| **Debug mode** (tricky bugs, full-codebase reasoning) | `deepseek/deepseek-v4-pro` | Stronger reasoning, large context window — better at tracing issues across a whole repo |

You can switch models anytime inside OpenCode with `/models`. Pricing and exact model availability can shift — check [openrouter.ai/models](https://openrouter.ai/models) for current rates before the hackathon.

### 3. Clone this repo
```powershell
gh repo clone sergiofspedro/EU_CivicTech_team
cd EU_CivicTech_team
```

### 4. Run the setup script — choose your OS

This repo has **two setup scripts** — pick the one matching your operating system:

1. **`setup.ps1`** — for **Windows** (PowerShell). Tested and verified end-to-end.
2. **`setup.sh`** — for **macOS** (Bash). ⚠️ **Untested** — mirrors the validated Windows setup but has not been run on a real Mac. If you hit an error, please fix it locally and share the correction with the team (commit + push, or flag in the team channel).

**Windows (PowerShell):**
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\setup.ps1
```

**macOS (Bash):**
```bash
chmod +x setup.sh
./setup.sh
```

Both scripts install the full shared stack:
- **Oh My OpenCode** — multi-agent orchestration
- **kdco/workspace bundle** — worktree, notify, delegation, planning agents
- **Simple Memory** — git-committed shared team memory
- **EnvSitter Guard** — blocks agents from reading/editing `.env` secrets
- **opencode-firecrawl** — web scraping/crawling (e.g. EU open data sources)
- **Playwright MCP** — browser automation/testing
- **Context7 + Exa + gh_grep** — bundled free with Oh My OpenCode

**Important:** the script installs `bun` first and will ask you to restart your terminal partway through — this is expected, just restart, `cd` back into the repo, and re-run the setup script.

### 5. Manual steps after the script
Two things need one-time setup outside the script:

**Firecrawl:**
```powershell
firecrawl login --browser
```
or set your own key: `FIRECRAWL_API_KEY` (get one at firecrawl.dev)

**Composio:**
1. Create your own free account at [composio.dev](https://composio.dev) (20K tool calls/month free, plenty for the hackathon)
2. Get your API key from the dashboard
3. **Never paste your API key in chat or Slack.** Set it as an environment variable instead.
4. Inside OpenCode, run `/connect` and select Composio

### 6. Launch OpenCode and connect
Open the OpenCode Desktop app, use **Open Folder** to point it at this repo folder, then in the chat bar:
```
/connect     -> enter YOUR OWN OpenRouter API key
/models      -> select your model per the table above (or switch as needed)
/init        -> scans the repo, merges with shared AGENTS.md rules
```

**Known issue (safe to ignore):** `/init` may show a one-time Playwright `SyntaxError` during the repo scan. It's non-blocking and doesn't affect the rest of the session.

### Project rules (automatic — see AGENTS.md)
- The agent will ask clarifying questions before ambiguous or non-trivial tasks
- The agent will propose alternative approaches before committing to one
- Local commits are fine; **pushing, opening PRs, or merging always require your explicit confirmation first**

### Shared vs personal config
- `opencode.json` (committed) — shared MCPs and agent permissions. Don't edit unless agreed with the team.
- `opencode.local.json` (gitignored) — your own personal overrides/extra instructions, optional, never shared.

### Security reminder
Never commit or paste API keys anywhere in chat, Slack, or this repo. All keys go in environment variables or the app's own `/connect` prompts.
