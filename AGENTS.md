# EU_CivicTech_team

## What this repo is

A civic tech prototype built for the EU Civic Tech Hackathon (22-23 June 2026), under the European Democracy Shield initiative. The project addresses a gap in EU civic participation infrastructure - see tools/philosophy.md for the guiding principles and README.md for team setup instructions.

Note: this repo previously contained an unrelated PDF-downloading script (download_pdfs.py, Unpaywall/DOI-based). That was a different, earlier project and is not part of the current hackathon scope - disregard any references to it unless the file is still physically present and explicitly relevant.

## Setup

See README.md for the full team onboarding flow (OpenCode Desktop, API keys, setup.ps1, plugin stack).

## Project rules

- Ask clarifying questions before starting any non-trivial or ambiguous task.
- Propose alternative approaches with trade-offs before committing to one.
- Local git commits are fine to make freely.
- NEVER push to GitHub, open a pull request, or merge a pull request without explicit confirmation first.

## OpenCode environment

- opencode.json loads instructions from ./tools/philosophy.md (project principles) and defines MCPs (Playwright, Context7, Exa, gh_grep) plus restricted agents (researcher, scribe, coder, reviewer, plan, build, explore) with granular bash/file permissions.
- Plugin stack installed via setup.ps1: Oh My OpenCode, kdco/workspace bundle (worktree, notify, delegation, planning), Simple Memory, EnvSitter Guard, opencode-firecrawl, Playwright MCP, Composio MCP, Context7 (bundled).
- worktree plugin confirmed working - use it for parallel branch work across team members.
- .gitignore protects opencode.local.json (personal config overrides) from being committed.

## Known non-blocking issues

- /init may show a one-time Playwright SyntaxError during repo scan. Safe to ignore.

## Tech stack

(To be filled in once the team finalizes the hackathon concept and architecture.)
