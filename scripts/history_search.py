from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import requests


EXA_API = "https://api.exa.ai"
URL_PATTERN = re.compile(r"https?://[^\s\]\)]+")


def mcporter_executable() -> str | None:
    for name in ("mcporter.cmd", "mcporter", "mcporter.ps1"):
        found = shutil.which(name)
        if found:
            return found
    return None


def mcporter_call(expression: str) -> str:
    executable = mcporter_executable()
    if not executable:
        raise RuntimeError("未找到 mcporter；请配置 EXA_API_KEY 或安装可用的 Exa MCP 路由")
    completed = subprocess.run(
        [executable, "call", expression],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return completed.stdout


def exa_headers() -> dict[str, str]:
    key = os.environ.get("EXA_API_KEY", "").strip()
    if not key:
        raise RuntimeError("未找到 Exa 路由，也未设置 EXA_API_KEY")
    return {"x-api-key": key, "Content-Type": "application/json"}


def direct_search(query: str, num_results: int, domains: list[str]) -> dict:
    payload = {
        "query": query,
        "numResults": num_results,
        "contents": {"text": {"maxCharacters": 12000}},
    }
    if domains:
        payload["includeDomains"] = domains
    response = requests.post(
        f"{EXA_API}/search",
        headers=exa_headers(),
        json=payload,
        timeout=60,
    )
    response.raise_for_status()
    return {"method": "exa_api", "query": query, "response": response.json()}


def direct_fetch(urls: list[str]) -> dict:
    response = requests.post(
        f"{EXA_API}/contents",
        headers=exa_headers(),
        json={"urls": urls, "text": {"maxCharacters": 20000}},
        timeout=90,
    )
    response.raise_for_status()
    return {"method": "exa_api", "urls": urls, "response": response.json()}


def search(query: str, num_results: int, domains: list[str]) -> dict:
    expression = (
        "exa.web_search_exa("
        f"query: {json.dumps(query, ensure_ascii=False)}, "
        f"numResults: {num_results}, "
        f"includeDomains: {json.dumps(domains, ensure_ascii=False)})"
    )
    executable = mcporter_executable()
    if executable:
        raw = mcporter_call(expression)
        return {
            "method": "mcporter_exa",
            "query": query,
            "domains": domains,
            "urls": list(dict.fromkeys(URL_PATTERN.findall(raw))),
            "raw": raw,
        }
    return direct_search(query, num_results, domains)


def fetch(urls: list[str]) -> dict:
    expression = (
        "exa.web_fetch_exa("
        f"urls: {json.dumps(urls, ensure_ascii=False)})"
    )
    executable = mcporter_executable()
    if executable:
        return {
            "method": "mcporter_exa",
            "urls": urls,
            "raw": mcporter_call(expression),
        }
    return direct_fetch(urls)


def doctor() -> dict:
    executable = mcporter_executable()
    if executable:
        completed = subprocess.run(
            [executable, "list"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        lines = completed.stdout.splitlines()
        exa_lines = [line for line in lines if re.search(r"\bexa\b", line, re.I)]
        healthy = any(
            "offline" not in line.lower()
            and "auth required" not in line.lower()
            and "error" not in line.lower()
            for line in exa_lines
        )
        return {
            "ok": healthy,
            "route": "mcporter",
            "executable": executable,
            "exa_lines": exa_lines,
            "raw": completed.stdout,
        }
    return {
        "ok": bool(os.environ.get("EXA_API_KEY", "").strip()),
        "route": "exa_api" if os.environ.get("EXA_API_KEY", "").strip() else None,
        "reason": "未找到 mcporter，使用 EXA_API_KEY 直连" if os.environ.get("EXA_API_KEY", "").strip() else "未找到 mcporter 或 EXA_API_KEY",
    }


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="独立的南燕公众号历史检索路由")
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("doctor")
    search_parser = subparsers.add_parser("search")
    search_parser.add_argument("--query", required=True)
    search_parser.add_argument("--num-results", type=int, default=10)
    search_parser.add_argument("--include-domain", action="append", default=["mp.weixin.qq.com"])
    fetch_parser = subparsers.add_parser("fetch")
    fetch_parser.add_argument("--url", action="append", required=True)
    args = parser.parse_args()

    try:
        if args.command == "doctor":
            result = doctor()
        elif args.command == "search":
            result = search(args.query, args.num_results, args.include_domain)
        else:
            result = fetch(args.url)
    except (RuntimeError, requests.RequestException, subprocess.CalledProcessError) as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        raise SystemExit(1) from exc
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
