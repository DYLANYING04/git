from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from build_proposal import (
    build_document,
    download_response_images,
    load_responses,
    valid_contact,
    validate_config,
    write_audit,
)
from validate_proposal import validate_document


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_REPORTS_ROOT = ROOT / "reports"
DEFAULT_TEMPLATE = ROOT / "templates" / "proposal-template.docx"
PERIOD_PATTERN = re.compile(r"(?P<year>20\d{2})\s*年[^\\/]*?(?P<month>1[0-2]|[1-9])\s*月")
DIRECTORY_PERIOD_PATTERN = re.compile(r"(?<!\d)(?P<year>20\d{2})[-_.](?P<month>1[0-2]|0[1-9])(?=\D|$)")
DATE_TEXT_PATTERN = re.compile(r"(?<!\d)(?P<year>20\d{2})[-/.年](?P<month>1[0-2]|0?[1-9])(?:[-/.月]\d{1,2})?")


@dataclass(frozen=True)
class ReportPaths:
    root: Path
    source: Path
    config: Path
    output: Path
    audit: Path
    manifest: Path
    preview: Path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def parse_period(value: str) -> tuple[int, int]:
    match = re.fullmatch(r"(20\d{2})-(1[0-2]|0[1-9])", value.strip())
    if not match:
        raise ValueError("月份必须使用 YYYY-MM 格式，例如 2026-09")
    return int(match.group(1)), int(match.group(2))


def period_from_filename(path: Path) -> tuple[int, int] | None:
    match = PERIOD_PATTERN.search(path.stem)
    if not match:
        return None
    return int(match.group("year")), int(match.group("month"))


def period_from_parent(path: Path) -> tuple[int, int] | None:
    for parent in path.parents:
        match = DIRECTORY_PERIOD_PATTERN.search(parent.name)
        if match:
            return int(match.group("year")), int(match.group("month"))
    return None


def period_from_workbook(path: Path) -> tuple[int, int] | None:
    counts: Counter[tuple[int, int]] = Counter()
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    value = cell.value
                    if isinstance(value, (datetime, date)):
                        counts[(value.year, value.month)] += 1
                    elif isinstance(value, str):
                        match = DATE_TEXT_PATTERN.search(value)
                        if match:
                            counts[(int(match.group("year")), int(match.group("month")))] += 1
    finally:
        workbook.close()
    if not counts:
        return None
    return counts.most_common(1)[0][0]


def infer_period(path: Path, explicit: str | None = None) -> tuple[int, int, str]:
    if explicit:
        year, month = parse_period(explicit)
        return year, month, "explicit"

    filename_period = period_from_filename(path)
    workbook_period = period_from_workbook(path)
    parent_period = period_from_parent(path)
    if filename_period and workbook_period and filename_period != workbook_period:
        raise ValueError(
            "Excel 文件名与问卷填写日期的月份不一致："
            f"文件名为 {filename_period[0]}-{filename_period[1]:02d}，"
            f"填写日期主要为 {workbook_period[0]}-{workbook_period[1]:02d}"
        )
    if filename_period:
        return *filename_period, "filename"
    if workbook_period:
        return *workbook_period, "workbook"
    if parent_period:
        return *parent_period, "parent_directory"
    raise ValueError("无法从 Excel 文件名、填写日期或所在目录识别月份，请使用 --period YYYY-MM")


def report_paths(reports_root: Path, year: int, month: int) -> ReportPaths:
    root = reports_root / f"{year}-{month:02d}"
    basename = f"北京大学深圳研究生院研究生会常务代表委员会{year}年{month}月提案汇总"
    output = root / f"{basename}.docx"
    return ReportPaths(
        root=root,
        source=root / "source",
        config=root / "proposal-config.json",
        output=output,
        audit=output.with_suffix(".audit.json"),
        manifest=root / "manifest.json",
        preview=root / "working" / "preview",
    )


def snapshot_source(source: Path, destination_dir: Path) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    source_hash = sha256_file(source)
    destination = destination_dir / source.name
    if destination.exists():
        if sha256_file(destination) == source_hash:
            return destination
        destination = destination_dir / f"{source.stem}-{source_hash[:8]}{source.suffix}"
        if destination.exists() and sha256_file(destination) == source_hash:
            return destination
    shutil.copy2(source, destination)
    return destination


def skeleton_config(year: int, month: int) -> dict:
    return {
        "year": year,
        "month": month,
        "reply_department": "总务处",
        "email": "gsusz@pku.edu.cn",
        "deadline": f"{year}年【】月【】日（周【】）",
        "excluded_rows": [],
        "allowed_reused_rows": [],
        "historical_evidence_notes": [],
        "historical_search_log": [],
        "sections": [],
    }


def install_config(source: Path | None, destination: Path, year: int, month: int) -> tuple[Path, bool]:
    if source:
        source = source.resolve()
        supplied = json.loads(source.read_text(encoding="utf-8"))
        if (int(supplied.get("year", 0)), int(supplied.get("month", 0))) != (year, month):
            raise ValueError("配置月份与识别出的问卷月份不一致")
        if destination.exists() and sha256_file(destination) != sha256_file(source):
            raise ValueError(f"月份目录中已存在不同配置，为避免覆盖已停止：{destination}")
        if not destination.exists():
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
        return destination, False

    if destination.exists():
        return destination, False
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(
        json.dumps(skeleton_config(year, month), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return destination, True


def render_preview(docx: Path, preview_dir: Path) -> list[Path]:
    soffice = shutil.which("soffice") or shutil.which("soffice.com")
    pdftoppm = shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        missing = [name for name, path in (("LibreOffice", soffice), ("pdftoppm", pdftoppm)) if not path]
        raise ValueError(f"缺少 Word 逐页预览依赖：{', '.join(missing)}")
    preview_dir.mkdir(parents=True, exist_ok=True)
    # Remove stale PDF/page images so the manifest only describes this render.
    for existing in preview_dir.iterdir():
        if existing.is_file():
            existing.unlink()
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(preview_dir), str(docx)],
        check=True,
        capture_output=True,
        text=True,
    )
    pdf = preview_dir / f"{docx.stem}.pdf"
    if not pdf.exists():
        raise ValueError("LibreOffice 未生成预期的 PDF 预览")
    page_prefix = preview_dir / "page"
    subprocess.run(
        [pdftoppm, "-jpeg", "-r", "130", str(pdf), str(page_prefix)],
        check=True,
        capture_output=True,
        text=True,
    )
    pages = sorted(
        preview_dir.glob("page-*.jpg"),
        key=lambda path: int(re.search(r"page-(\d+)\.jpg$", path.name).group(1)),
    )
    if not pages:
        raise ValueError("未生成逐页图片预览")
    return pages


def write_manifest(paths: ReportPaths, payload: dict) -> None:
    paths.root.mkdir(parents=True, exist_ok=True)
    paths.manifest.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="仅凭一份问卷 Excel 建立当月目录并生成提案 Word")
    parser.add_argument("--xlsx", type=Path, required=True)
    parser.add_argument("--period", help="仅在无法自动识别时填写 YYYY-MM")
    parser.add_argument("--config", type=Path, help="可选：已整理好的当月提案配置")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--reports-root", type=Path, default=DEFAULT_REPORTS_ROOT)
    parser.add_argument("--skip-render", action="store_true", help="跳过 PDF 和逐页图片预览")
    args = parser.parse_args()

    xlsx = args.xlsx.resolve()
    template = args.template.resolve()
    if not xlsx.is_file():
        raise SystemExit(f"Excel 不存在：{xlsx}")
    if xlsx.suffix.lower() != ".xlsx":
        raise SystemExit("输入必须是 .xlsx 文件")
    if not template.is_file():
        raise SystemExit(f"固定模板不存在：{template}")

    try:
        year, month, period_source = infer_period(xlsx, args.period)
        paths = report_paths(args.reports_root.resolve(), year, month)
        source_snapshot = snapshot_source(xlsx, paths.source)
        config_path, created_skeleton = install_config(args.config, paths.config, year, month)
        config = json.loads(config_path.read_text(encoding="utf-8"))
        responses = load_responses(source_snapshot)
        base_manifest = {
            "period": f"{year}-{month:02d}",
            "period_source": period_source,
            "source": {
                "snapshot": str(source_snapshot.resolve()),
                "sha256": sha256_file(source_snapshot),
                "response_count": len(responses),
                "eligible_response_count": sum(valid_contact(item["contact"]) for item in responses.values()),
            },
            "template": {"path": str(template), "sha256": sha256_file(template)},
            "config": {"path": str(config_path.resolve()), "sha256": sha256_file(config_path)},
        }
        if created_skeleton or not config.get("sections"):
            write_manifest(paths, {**base_manifest, "status": "needs_editorial_config"})
            print(json.dumps({
                "status": "needs_editorial_config",
                "month_directory": str(paths.root.resolve()),
                "source_snapshot": str(source_snapshot.resolve()),
                "config": str(config_path.resolve()),
            }, ensure_ascii=False, indent=2))
            return

        if (int(config["year"]), int(config["month"])) != (year, month):
            raise ValueError("月份目录、Excel 和配置的年月不一致")
        covered = validate_config(config, responses)
        image_downloads = download_response_images(responses, set(covered), paths.output)
        build_document(template, config, responses, image_downloads, paths.output)
        audit_path = write_audit(config, responses, covered, image_downloads, paths.output)
        validation_errors = validate_document(config, responses, paths.output)
        if validation_errors:
            raise ValueError("Word 校验失败：" + "；".join(validation_errors))
        pages = [] if args.skip_render else render_preview(paths.output, paths.preview)
        audit = json.loads(audit_path.read_text(encoding="utf-8"))
        write_manifest(paths, {
            **base_manifest,
            "status": "generated_and_validated",
            "output": {
                "docx": str(paths.output.resolve()),
                "sha256": sha256_file(paths.output),
                "audit": str(audit_path.resolve()),
                "proposal_count": audit["proposal_count"],
                "embedded_image_count": audit["embedded_image_count"],
                "preview_page_count": len(pages),
            },
        })
        print(json.dumps({
            "status": "generated_and_validated",
            "month_directory": str(paths.root.resolve()),
            "docx": str(paths.output.resolve()),
            "audit": str(audit_path.resolve()),
            "manifest": str(paths.manifest.resolve()),
            "preview_pages": [str(page.resolve()) for page in pages],
        }, ensure_ascii=False, indent=2))
    except (OSError, ValueError, json.JSONDecodeError, subprocess.CalledProcessError) as exc:
        raise SystemExit(str(exc)) from exc


if __name__ == "__main__":
    main()
