# History Search

Use this procedure for every proposal, not just one representative issue.

## Health check

Run `agent-reach doctor`. If it has no output or hangs, do not wait indefinitely. Run `mcporter list` and continue only when `exa` is healthy.

## Search

Earlier report directories are not a history source for the current month. They may be read for layout, category, and wording reference only. Do not copy their `history_reference`, `history_sources`, screenshots, or historical conclusions into the current configuration.

Start with a concrete query such as:

```text
南燕研究生会 + core issue words + 提案
```

Add the facility, building, month, or year when needed. Prefer the exact object over broad terms such as only “食堂” or “宿舍”. Keep one search-log entry per proposal, even when no result is found.

```powershell
mcporter call 'exa.web_search_exa(query: "南燕研究生会 上床下桌 提案", numResults: 10, includeDomains: ["mp.weixin.qq.com"])'
```

## Full text

When a URL or mirror URL is found, fetch the full text with the current Exa interface:

```powershell
mcporter call 'exa.web_fetch_exa(urls: ["ARTICLE_URL"])'
```

`exa.crawling_exa` is obsolete. Record the article month, title, item locator, evidence type, and URL in `history_sources`. A mirror can be used when its full text is clear, but the Word supplement contains months only and never article titles, item numbers, or URLs.

## Evidence filter

- Allowed: a verified proposal from the 12 months before the questionnaire month.
- Record but do not adopt: ordinary campus news, logistics meetings, notices, recruitment posts, search-only snippets, and articles outside the time window.
- Never infer a reply or proposal month from an ellipsis or an unrelated article.
