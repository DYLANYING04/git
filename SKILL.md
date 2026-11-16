---
name: nanyan-proposal
description: "Generate the Nanyan Graduate Student Union Standing Representatives Committee rights-proposal report from a questionnaire Excel file. Use this skill whenever a user provides a survey .xlsx, asks for a monthly rights proposal report, asks to reproduce the Nanyan proposal format, or asks to search past Nanyan proposal history. The skill is self-contained: use its bundled rules, scripts, template, validation, history-search workflow, image handling, and preview checks; do not substitute a generic document or spreadsheet workflow."
compatibility: "Requires Python 3.13+, the bundled Python packages in requirements.txt, LibreOffice, and pdftoppm for rendered previews. Internet history checks require an available Exa route or an Exa API key."
---

# Nanyan Proposal Report

This skill turns one questionnaire `.xlsx` into a month-isolated Word proposal report. Treat the bundled `references/AGENTS.md` as the complete project rulebook. Do not ask the user for a template, output directory, or previous configuration.

## Deterministic workflow

1. Run the bundled entry point with the user's Excel file:
   `python scripts/monthly_report.py --xlsx "questionnaire.xlsx"`
   If it returns `needs_editorial_config`, this is an internal scaffold state, not a deliverable. Do not stop or ask the user to edit JSON. Read the workbook, complete the month-specific `proposal-config.json` yourself, and run the entry point again.
2. Read the generated monthly `proposal-config.json` and inspect every non-empty response. Identify headers by name, normalize email domains, accept only valid email addresses or 11-digit mainland China mobile numbers, preserve source rows, split independent issues when they have different objects or measures or belong to different domains, and record any reused row in `allowed_reused_rows`.
3. Use the bundled `exa.web_search_exa` and `exa.web_fetch_exa` MCP tools when the plugin host exposes them. The plugin config points to Exa's hosted MCP server and does not require an API key for basic anonymous use. In a terminal-only run, use `python scripts/history_search.py doctor` followed by its `search` and `fetch` commands. No Agent Reach skill is required. Earlier reports may be used for style and category reference only; never copy their `history_reference`, `history_sources`, or historical conclusions into the current month. Record the query, time window, result status, URLs, evidence type, and adoption decision in `historical_search_log`. Never treat ordinary news, notices, meetings, prior report annotations, or out-of-window material as current historical evidence.
4. Fill every proposal's `measure_basis` with exact questionnaire quotations. Do not add measures that are absent from the questionnaire.
5. Re-run the same command. It creates the Word file, audit JSON, manifest, images, PDF, and page previews. The renderer clears stale preview files before conversion.
6. Open every generated page under `reports/YYYY-MM/working/preview/` in numeric order. Check for blank pages, clipped text, orphan headings, overlap, misplaced images, and unexpected highlights.

## Bundled resources

- `scripts/monthly_report.py`: month detection, source snapshot, build, validation, render, and manifest.
- `scripts/build_proposal.py`: Excel loading, image download/compression, Word generation, source checks, and audit.
- `scripts/validate_proposal.py`: content, language, styles, source, contact, image, and highlight checks.
- `templates/proposal-template.docx`: fixed approved layout.
- `references/AGENTS.md`: full Chinese rules and quality gates.
- `scripts/history_search.py`: bundled Exa health check, search, and full-text route.
- `references/history-search.md`: independent Exa and mirror-search procedure.
- `references/reproducibility.md`: frozen versus live search and release-archive guidance.

## Output contract

Every run writes only to `reports/YYYY-MM/`:

- final `.docx`
- same-name `.audit.json`
- `manifest.json`
- read-only Excel snapshot
- original and compressed image attachments
- PDF and one JPEG per page
- `proposal-config.json` containing source mappings and per-proposal search logs

## Reproducibility boundary

The skill pins document logic, template, rules, dependency versions, and output structure. Pixel-identical rendering also requires the same Python version, LibreOffice version, fonts, and PDF renderer. Use the bundled `Dockerfile` for the strongest cross-machine match; otherwise install the versions in `requirements.txt` and use the same LibreOffice/font environment.

## Privacy

Do not commit real questionnaires, contact information, generated reports, attachments, or monthly `reports/` directories to GitHub. Use `.gitignore` and keep only the template, code, rules, redacted fixtures, and tests in the public repository.
