# Project Philosophy - Livro-Redes-Democracia

## Context
This project is built for the EU Civic Tech Hackathon (22-22 June 2026), under the European Democracy Shield. It must demonstrate a working prototype addressing a real, named gap in EU civic participation, with potential for post-hackathon scaling.

## Core principles

**Simplicity over ambition.** A focused, convincing 48-hour demo beats an incomplete ambitious build. When in doubt, cut scope, not quality.

**Interoperability over isolation.** This tool should integrate with or complement existing EU civic tech infrastructure (Decidim, CONSUL, Adhocracy+, EDMO, EU Digital Identity Wallet) rather than reinventing it. Prefer adapters and plugins over standalone replacements.

**Digital sovereignty by design.** Where AI is used, prefer EU-hosted or open-weight models over foreign commercial APIs (OpenAI, Anthropic, Google), in line with the June 2026 EU Technological Sovereignty Package. If a foreign API is used for prototyping speed, flag it clearly as a placeholder to be replaced.

**Safety and trust are not optional.** Any feature touching citizen-submitted content, deliberation, or AI-generated output must consider misinformation, manipulation, and AI Act compliance (Article 50: transparency and detection for AI-generated content) from the start, not as an afterthought.

**Open by default.** Code, data schemas, and documentation should be open-source where possible, to support EU Civic Tech Hub adoption and avoid vendor lock-in.

## Working style

- Ask clarifying questions before ambiguous or non-trivial implementation decisions.
- Propose alternative approaches with trade-offs before committing to one, especially for architecture choices.
- Local commits are fine to make freely. Never push, open a PR, or merge without explicit confirmation.
- Favor working, demoable increments over theoretically complete but untested features.
