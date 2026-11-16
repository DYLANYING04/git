# History Search

Use this procedure for every proposal, not just one representative issue.

## Health check

When this repository is installed as a Codex/ChatGPT plugin, use the bundled `exa` MCP server first. Its endpoint is declared in `.mcp.json`:

```json
{
  "mcpServers": {
    "exa": {
      "type": "streamable-http",
      "url": "https://mcp.exa.ai/mcp"
    }
  }
}
```

Exa documents this hosted endpoint as an anonymous, rate-limited option; sign-in or an API key is only needed for higher limits or authenticated Exa Agent features. If the host does not load plugin MCP servers, run the bundled route check:

```powershell
python scripts/history_search.py doctor
```

The script uses Exa through `mcporter` when available and falls back to `EXA_API_KEY`. It returns a failure instead of silently skipping history search.

## Search

Earlier report directories are not a history source for the current month. They may be read for layout, category, and wording reference only. Do not copy their `history_reference`, `history_sources`, screenshots, or historical conclusions into the current configuration.

Start with a concrete query such as:

```text
南燕研究生会 + core issue words + 提案
```

Add the facility, building, month, or year when needed. Prefer the exact object over broad terms such as only “食堂” or “宿舍”. Keep one search-log entry per proposal, even when no result is found.

```powershell
python scripts/history_search.py search --query "南燕研究生会 上床下桌 提案" --num-results 10
```

## Full text

When a URL or mirror URL is found, fetch the full text with the current Exa interface:

```powershell
python scripts/history_search.py fetch --url "ARTICLE_URL"
```

`exa.crawling_exa` is obsolete. Record the article month, title, item locator, evidence type, and URL in `history_sources`. A mirror can be used when its full text is clear, but the Word supplement contains months only and never article titles, item numbers, or URLs. The bundled route records the raw search/fetch response so the method used is auditable.

## Evidence filter

- Allowed: a verified proposal from the 12 months before the questionnaire month.
- Record but do not adopt: ordinary campus news, logistics meetings, notices, recruitment posts, search-only snippets, and articles outside the time window.
- Never infer a reply or proposal month from an ellipsis or an unrelated article.
