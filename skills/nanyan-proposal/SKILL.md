---
name: nanyan-proposal
description: "Generate the Nanyan Graduate Student Union rights-proposal report from a questionnaire Excel file. The plugin bundles the Exa MCP connection for independent public-account history searches."
---

Use the complete workflow in the plugin root `SKILL.md`, `references/AGENTS.md`, and `references/history-search.md`. When the host exposes the bundled `exa` MCP server, call `exa.web_search_exa` and `exa.web_fetch_exa` directly. If direct MCP tools are unavailable in a terminal-only run, use `python scripts/history_search.py` as the bundled fallback. Never copy historical evidence from earlier reports.
