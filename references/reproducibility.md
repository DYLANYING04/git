# Reproducibility Modes

## Frozen mode

Use a committed, redacted monthly configuration with:

- `source_rows`, `body`, `measure_basis`, and section ordering
- `history_reference` and `history_sources`
- `historical_search_log`
- the exact template hash and dependency lock

Running the same Excel snapshot, configuration, Python packages, fonts, LibreOffice, and renderer produces the same report structure and content. Keep private contacts and source workbooks outside the public repository.

## Live mode

When only a new Excel is supplied, history search uses the live web. Search ranking, article availability, and article mirrors can change. The report remains rule-compliant, but live search cannot promise byte-for-byte identical output. Save the generated configuration and audit JSON if the result must be reproduced later.

## Recommended release layout

Keep the public GitHub repository limited to the skill package. Keep a private companion archive for each released report containing the original Excel snapshot, monthly config, audit JSON, manifest, and fetched history evidence.
