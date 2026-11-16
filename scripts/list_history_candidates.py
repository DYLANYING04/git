from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


PERIOD = re.compile(r"^(20\d{2})-(0[1-9]|1[0-2])$")


def parse_period(value: str) -> int:
    match = PERIOD.fullmatch(value)
    if not match:
        raise ValueError("月份必须使用 YYYY-MM 格式")
    return int(match.group(1)) * 12 + int(match.group(2))


def main() -> None:
    parser = argparse.ArgumentParser(
        description="从本地往期报告配置中列出近一年历史依据候选"
    )
    parser.add_argument("--reports-root", type=Path, default=Path("reports"))
    parser.add_argument("--period", required=True, help="当前问卷月份 YYYY-MM")
    args = parser.parse_args()

    current_month = parse_period(args.period)
    candidates = []
    for config_path in sorted(args.reports_root.glob("*/proposal-config.json")):
        try:
            config = json.loads(config_path.read_text(encoding="utf-8"))
            report_period = int(config["year"]) * 12 + int(config["month"])
        except (OSError, ValueError, KeyError, json.JSONDecodeError):
            continue
        if not 1 <= current_month - report_period <= 12:
            continue
        for section in config.get("sections", []):
            for proposal in section.get("proposals", []):
                if not proposal.get("history_sources"):
                    continue
                candidates.append({
                    "report": str(config_path.resolve()),
                    "report_period": f"{config['year']}-{int(config['month']):02d}",
                    "proposal_title": proposal.get("title"),
                    "history_reference": proposal.get("history_reference"),
                    "history_sources": proposal.get("history_sources", []),
                })
    print(json.dumps(candidates, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
