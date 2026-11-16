from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
from collections import Counter
from pathlib import Path

import requests
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook
from PIL import Image, ImageOps, UnidentifiedImageError


EMPTY_VALUES = {"", "(空)", "空", "无", "暂无", "n/a", "none", "null"}
EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
PHONE_PATTERN = re.compile(r"^1[3-9]\d{9}$")
URL_PATTERN = re.compile(r'''https?://[^\s,，;；"'<>\)]+''', re.IGNORECASE)
MAX_IMAGE_BYTES = 25 * 1024 * 1024
SUPPORTED_IMAGE_FORMATS = {
    "JPEG": ".jpg",
    "PNG": ".png",
    "GIF": ".gif",
    "BMP": ".bmp",
    "TIFF": ".tiff",
}
DEFENSIVE_PHRASES = {
    "评估",
    "可行性",
    "在条件允许的情况下",
    "有待改善",
    "逐步完善",
    "酌情",
    "视情况",
}
UNSUPPORTED_MEASURES = {
    "临时警示",
    "故障反馈渠道",
    "夜间转介机制",
    "备用通行",
}
WEAK_PROBLEM_PHRASES = {
    "饭菜温度不足",
    "菜品温度不足",
}
TOPIC_GROUPS = {
    "食堂": ("食堂", "饭菜", "菜品", "早餐", "包子", "麻辣烫"),
    "网络": ("网络", "校园网", "pku网", "网速"),
    "教学": ("课程", "教学网", "课堂", "课件", "录像回放"),
    "住宿": ("宿舍", "床架", "电梯", "晾晒", "小飞虫"),
    "运动": ("健身房", "运动场", "器械", "晨练"),
}


def normalized(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def canonical_contact(value) -> str:
    """Normalize common contact forms without accepting malformed values."""
    contact = normalized(value)
    if EMAIL_PATTERN.fullmatch(contact):
        local, domain = contact.rsplit("@", 1)
        return f"{local}@{domain.lower()}"
    digits = re.sub(r"[\s-]", "", contact)
    if PHONE_PATTERN.fullmatch(digits):
        return digits
    return contact


def contact_error(value) -> str | None:
    contact = canonical_contact(value)
    if contact.lower() in EMPTY_VALUES:
        return "未提供联系方式"
    if EMAIL_PATTERN.fullmatch(contact):
        return None
    if PHONE_PATTERN.fullmatch(contact):
        return None
    return "联系方式格式无效（邮箱地址或11位中国大陆手机号码）"


def valid_contact(value) -> bool:
    return contact_error(value) is None


def topic_groups(text: str) -> set[str]:
    lowered = text.lower()
    groups = {
        group
        for group, terms in TOPIC_GROUPS.items()
        if any(term.lower() in lowered for term in terms)
    }
    # "宿舍网络" is one network issue; "宿舍" is only its location.
    if "网络" in groups:
        groups.discard("住宿")
    return groups


def extract_image_urls(cell):
    candidates = []
    if cell.hyperlink and cell.hyperlink.target:
        candidates.append(cell.hyperlink.target)
    value = normalized(cell.value)
    if value.lower() not in EMPTY_VALUES:
        candidates.extend(URL_PATTERN.findall(value))

    urls = []
    for candidate in candidates:
        match = URL_PATTERN.search(candidate)
        if not match:
            continue
        url = match.group(0).strip().strip('"\'')
        if url not in urls:
            urls.append(url)
    return urls


def locate_columns(sheet):
    columns = {}
    for cell in sheet[1]:
        text = normalized(cell.value)
        if "联系方式" in text:
            columns["contact"] = cell.column
        if "问题描述" in text:
            columns["problem"] = cell.column
        if "对于解决或改进上述反馈的问题" in text:
            columns["suggestion"] = cell.column
        if "相关图片" in text or "related pictures" in text.lower():
            columns["images"] = cell.column
        if text == "序号":
            columns["id"] = cell.column
    missing = {"contact", "problem", "id"} - set(columns)
    if missing:
        raise ValueError(f"问卷缺少必要字段：{', '.join(sorted(missing))}")
    return columns


def load_responses(path: Path):
    workbook = load_workbook(path, data_only=False, read_only=False)
    sheet = workbook.active
    columns = locate_columns(sheet)
    responses = {}
    for row in range(2, sheet.max_row + 1):
        source_id = sheet.cell(row, columns["id"]).value
        problem = normalized(sheet.cell(row, columns["problem"]).value)
        if source_id is None and not problem:
            continue
        responses[row] = {
            "source_id": source_id,
            "contact": canonical_contact(sheet.cell(row, columns["contact"]).value),
            "problem": problem,
            "suggestion": (
                normalized(sheet.cell(row, columns["suggestion"]).value)
                if "suggestion" in columns
                else ""
            ),
            "image_urls": (
                extract_image_urls(sheet.cell(row, columns["images"]))
                if "images" in columns
                else []
            ),
        }
    workbook.close()
    return responses


def download_response_images(responses, included_rows, output: Path):
    attachments_dir = output.parent / "attachments"
    original_dir = attachments_dir / "original"
    word_dir = attachments_dir / "word"
    downloads = {}
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 proposal-document-builder/1.0"})

    for row, item in responses.items():
        if row not in included_rows or not valid_contact(item["contact"]):
            continue
        for index, url in enumerate(item["image_urls"], start=1):
            try:
                response = session.get(url, timeout=30)
                response.raise_for_status()
            except requests.RequestException as exc:
                raise ValueError(f"Excel 第 {row} 行相关图片下载失败：{exc}") from exc

            content = response.content
            if not content:
                raise ValueError(f"Excel 第 {row} 行相关图片为空")
            if len(content) > MAX_IMAGE_BYTES:
                raise ValueError(
                    f"Excel 第 {row} 行相关图片超过 {MAX_IMAGE_BYTES // 1024 // 1024} MB 限制"
                )

            try:
                with Image.open(io.BytesIO(content)) as image:
                    image.verify()
                with Image.open(io.BytesIO(content)) as image:
                    image_format = image.format
                    width_px, height_px = image.size
            except (UnidentifiedImageError, OSError) as exc:
                raise ValueError(f"Excel 第 {row} 行链接未返回有效图片") from exc

            if image_format not in SUPPORTED_IMAGE_FORMATS:
                raise ValueError(
                    f"Excel 第 {row} 行图片格式 {image_format or '未知'} 无法插入 Word"
                )
            digest = hashlib.sha256(content).hexdigest()
            suffix = SUPPORTED_IMAGE_FORMATS[image_format]
            filename = f"excel-row-{row}-{index}-{digest[:12]}{suffix}"
            original_dir.mkdir(parents=True, exist_ok=True)
            word_dir.mkdir(parents=True, exist_ok=True)
            local_path = original_dir / filename
            local_path.write_bytes(content)

            embedded_path = word_dir / f"excel-row-{row}-{index}-{digest[:12]}-word.jpg"
            with Image.open(io.BytesIO(content)) as image:
                image = ImageOps.exif_transpose(image).convert("RGB")
                image.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
                image.save(embedded_path, "JPEG", quality=75, optimize=True, progressive=True)
                embedded_width_px, embedded_height_px = image.size
            downloads.setdefault(row, []).append(
                {
                    "source_url": url,
                    "local_path": str(local_path.resolve()),
                    "sha256": digest,
                    "content_type": response.headers.get("Content-Type"),
                    "format": image_format,
                    "width_px": width_px,
                    "height_px": height_px,
                    "embedded_path": str(embedded_path.resolve()),
                    "embedded_width_px": embedded_width_px,
                    "embedded_height_px": embedded_height_px,
                    "embedded_bytes": embedded_path.stat().st_size,
                }
            )
    return downloads


def clear_body(document: Document):
    body = document._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def set_run_font(run, size=12, bold=None, color=None, highlight=False):
    run.font.name = "Times New Roman"
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    if highlight:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW


def set_paragraph_format(
    paragraph,
    *,
    align=None,
    first_indent=None,
    left_indent=None,
    line_spacing=1.5,
    keep_with_next=False,
):
    paragraph.alignment = align
    fmt = paragraph.paragraph_format
    fmt.line_spacing = line_spacing
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.keep_with_next = keep_with_next
    if first_indent is not None:
        fmt.first_line_indent = Pt(first_indent)
    if left_indent is not None:
        fmt.left_indent = Pt(left_indent)


def add_text_paragraph(
    document,
    text="",
    *,
    style="Normal",
    size=12,
    bold=False,
    color=None,
    highlight=False,
    align=None,
    first_indent=None,
    left_indent=None,
    keep_with_next=False,
):
    paragraph = document.add_paragraph(style=style)
    set_paragraph_format(
        paragraph,
        align=align,
        first_indent=first_indent,
        left_indent=left_indent,
        keep_with_next=keep_with_next,
    )
    if text:
        run = paragraph.add_run(text)
        set_run_font(run, size=size, bold=bold, color=color, highlight=highlight)
    return paragraph


def add_email_hyperlink(paragraph, email):
    relationship_id = paragraph.part.relate_to(
        f"mailto:{email}",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    fonts.set(qn("w:eastAsia"), "宋体")
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "24")
    run_properties.extend([fonts, color, underline, size])
    text = OxmlElement("w:t")
    text.text = email
    run.extend([run_properties, text])
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_related_images(document, image_items):
    if not image_items:
        return
    section = document.sections[-1]
    max_width = min(Inches(4.9), section.page_width - section.left_margin - section.right_margin)
    max_height = Inches(6.2)
    for item in image_items:
        width_px = item["embedded_width_px"]
        height_px = item["embedded_height_px"]
        ratio = min(max_width / width_px, max_height / height_px)
        image_paragraph = document.add_paragraph(style="Normal")
        set_paragraph_format(image_paragraph, line_spacing=1.5)
        run = image_paragraph.add_run()
        run.add_picture(
            item["embedded_path"],
            width=int(width_px * ratio),
            height=int(height_px * ratio),
        )


def ensure_styles(document):
    styles = document.styles
    if "Heading 1" not in styles:
        styles.add_style("Heading 1", WD_STYLE_TYPE.PARAGRAPH)
    heading = styles["Heading 1"]
    heading.font.name = "Times New Roman"
    heading._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    heading.font.size = Pt(14)
    heading.font.bold = True
    heading.font.color.rgb = RGBColor(0x00, 0x7B, 0xB8)
    heading.paragraph_format.line_spacing = 1.5
    heading.paragraph_format.space_before = Pt(0)
    heading.paragraph_format.space_after = Pt(0)


def validate_rendered_language(document: Document):
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    violations = [
        phrase
        for phrase in sorted(DEFENSIVE_PHRASES | UNSUPPORTED_MEASURES)
        if phrase in text
    ]
    if violations:
        raise ValueError(f"Word 输出前语言检查失败：{', '.join(violations)}")


def proposal_contacts(proposal, responses):
    contacts = []
    for row in proposal["source_rows"]:
        if row not in responses:
            raise ValueError(f"配置引用了不存在的 Excel 行：{row}")
        contact = responses[row]["contact"]
        if not valid_contact(contact):
            raise ValueError(f"Excel 第 {row} 行未提供有效联系方式，不能收录")
        if contact not in contacts:
            contacts.append(contact)
    return contacts


def validate_config(config, responses):
    covered = Counter()
    manual_exclusions = {}
    for exclusion in config.get("excluded_rows", []):
        row = exclusion.get("row")
        reason = normalized(exclusion.get("reason"))
        if row not in responses:
            raise ValueError(f"人工排除项引用了不存在的 Excel 行：{row}")
        if not valid_contact(responses[row]["contact"]):
            raise ValueError(f"Excel 第 {row} 行本就没有有效联系方式，无需重复人工排除")
        if not reason:
            raise ValueError(f"Excel 第 {row} 行的人工排除项缺少原因")
        if row in manual_exclusions:
            raise ValueError(f"Excel 第 {row} 行被重复列入人工排除项")
        manual_exclusions[row] = reason

    proposals = [
        proposal
        for section in config["sections"]
        for proposal in section["proposals"]
    ]
    logs = config.get("historical_search_log")
    if not isinstance(logs, list):
        raise ValueError("配置缺少 historical_search_log 列表")
    proposal_titles = [proposal.get("title", "") for proposal in proposals]
    log_titles = [item.get("proposal_title") for item in logs if isinstance(item, dict)]
    if len(logs) != len(proposals) or sorted(log_titles) != sorted(proposal_titles):
        raise ValueError(
            "historical_search_log 必须为每条提案提供且仅提供一条记录"
        )
    required_log_fields = {
        "proposal_title",
        "query",
        "window",
        "result",
        "urls",
        "evidence_type",
        "adoption_decision",
    }
    for item in logs:
        if not isinstance(item, dict):
            raise ValueError("historical_search_log 中存在非对象记录")
        missing = sorted(required_log_fields - set(item))
        if missing:
            raise ValueError(
                "历史检索记录缺少字段：" + ", ".join(missing)
            )
        if not all(normalized(item.get(field)) for field in (
            "proposal_title", "query", "window", "result",
            "evidence_type", "adoption_decision"
        )):
            raise ValueError("历史检索记录的文字字段不能为空")
        if not isinstance(item.get("urls"), list):
            raise ValueError("历史检索记录的 urls 必须是列表")

    for section in config["sections"]:
        for proposal in section["proposals"]:
            rows = proposal.get("source_rows", [])
            if not rows:
                raise ValueError(f"提案“{proposal.get('title', '')}”未填写 source_rows")
            if len(rows) != len(set(rows)):
                raise ValueError(f"提案“{proposal['title']}”存在重复来源行")
            proposal_contacts(proposal, responses)
            covered.update(rows)
            if "反应" in proposal["body"]:
                raise ValueError(f"提案“{proposal['title']}”应使用“反映”")
            weak_phrases = [
                phrase for phrase in WEAK_PROBLEM_PHRASES if phrase in proposal["body"]
            ]
            if weak_phrases:
                raise ValueError(
                    f"提案“{proposal['title']}”含不准确的保温表述：{', '.join(weak_phrases)}；"
                    "请改为“保温不当”等直接表述"
                )
            groups = topic_groups(proposal["body"])
            if len(groups) > 1:
                raise ValueError(
                    f"提案“{proposal['title']}”混合了多个问题领域（{'、'.join(sorted(groups))}），"
                    "请拆成独立提案"
                )
            for phrase in sorted(DEFENSIVE_PHRASES):
                if phrase in proposal["body"]:
                    raise ValueError(
                        f"提案“{proposal['title']}”含防御性或模板化措辞“{phrase}”"
                    )
            for phrase in sorted(UNSUPPORTED_MEASURES):
                if phrase in proposal["body"]:
                    raise ValueError(
                        f"提案“{proposal['title']}”含禁止自行补充的措施“{phrase}”"
                    )

            measure_basis = proposal.get("measure_basis", [])
            if not measure_basis:
                raise ValueError(f"提案“{proposal['title']}”缺少措施原文依据")
            for basis in measure_basis:
                basis_row = basis.get("source_row")
                field = basis.get("field")
                quote = normalized(basis.get("quote"))
                if basis_row not in rows:
                    raise ValueError(
                        f"提案“{proposal['title']}”的措施依据行 {basis_row} 不在 source_rows 中"
                    )
                if field not in {"problem", "suggestion"}:
                    raise ValueError(
                        f"提案“{proposal['title']}”的措施依据字段必须是 problem 或 suggestion"
                    )
                if not quote or quote not in responses[basis_row][field]:
                    raise ValueError(
                        f"提案“{proposal['title']}”的措施依据不是 Excel 第 {basis_row} 行原文"
                    )
            if proposal.get("history_reference") and not proposal.get("history_sources"):
                raise ValueError(f"历史提案“{proposal['title']}”缺少 history_sources")
            if proposal.get("history_reference"):
                if re.search(r"第\s*\d+\s*条|《|》|回复", proposal["history_reference"]):
                    raise ValueError(
                        f"历史提案“{proposal['title']}”的 history_reference "
                        "只能写年月和“有提到相关问题”"
                    )
                history_pattern = (
                    r"\d{4}年\d{1,2}月提案"
                    r"(?:和\d{4}年\d{1,2}月提案)*"
                    r"(?:有)?提到(?:过)?相关问题"
                )
                if not re.fullmatch(history_pattern, proposal["history_reference"]):
                    raise ValueError(
                        f"历史提案“{proposal['title']}”的 history_reference 句式不符合规范"
                    )
            for source in proposal.get("history_sources", []):
                required = {
                    "platform", "account", "year", "month", "article_title", "item",
                    "query", "evidence_type", "locator"
                }
                missing_source_fields = sorted(required - set(source))
                if missing_source_fields:
                    raise ValueError(
                        f"历史提案“{proposal['title']}”的来源缺少字段："
                        f"{', '.join(missing_source_fields)}"
                    )
                source_month = int(source["year"]) * 12 + int(source["month"])
                current_month = int(config["year"]) * 12 + int(config["month"])
                month_age = current_month - source_month
                if not 1 <= month_age <= 12:
                    raise ValueError(
                        f"历史提案“{proposal['title']}”引用的"
                        f"{source['year']}年{source['month']}月不在近一年窗口内"
                    )

    allowed_reuse = {}
    for item in config.get("allowed_reused_rows", []):
        row = item.get("row")
        reason = normalized(item.get("reason"))
        if not reason:
            raise ValueError(f"Excel 第 {row} 行的拆分复用项缺少原因")
        allowed_reuse[row] = reason
    duplicates = {row for row, count in covered.items() if count > 1}
    unapproved_duplicates = sorted(duplicates - set(allowed_reuse))
    if unapproved_duplicates:
        raise ValueError(f"以下 Excel 来源行被重复计入多条提案：{unapproved_duplicates}")
    unused_approvals = sorted(set(allowed_reuse) - duplicates)
    if unused_approvals:
        raise ValueError(f"以下来源行登记了拆分复用，但实际未被复用：{unused_approvals}")
    overlap = sorted(set(covered) & set(manual_exclusions))
    if overlap:
        raise ValueError(f"以下 Excel 来源行同时被收录和人工排除：{overlap}")

    eligible = {row for row, item in responses.items() if valid_contact(item["contact"])}
    missing = sorted(eligible - set(covered) - set(manual_exclusions))
    if missing:
        raise ValueError(f"以下有效问卷既未映射到提案，也未登记人工排除原因：{missing}")
    return covered


def build_document(template: Path, config, responses, image_downloads, output: Path):
    document = Document(template)
    clear_body(document)
    ensure_styles(document)

    year = config["year"]
    month = config["month"]
    department = config["reply_department"]

    add_text_paragraph(
        document,
        "北京大学深圳研究生院研究生会常务代表委员会",
        size=18,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    add_text_paragraph(
        document,
        f"{year}年{month}月提案汇总",
        size=18,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    add_text_paragraph(document)
    add_text_paragraph(document, "各位老师好：")
    add_text_paragraph(
        document,
        "研究生会常务代表委员会近期根据学生反馈信息，汇总了近期的学生权益相关提案信息，希望获得学院相关部门的回复和解决，共同提升学生在校的生活质量，更好地为同学们服务。",
        first_indent=24,
    )

    deadline_paragraph = add_text_paragraph(document, first_indent=24)
    prefix = deadline_paragraph.add_run(
        "以下为具体问题及相关建议方案，请各位老师在回复一栏进行回复即可。请各位老师尽量于"
    )
    set_run_font(prefix)
    deadline = deadline_paragraph.add_run(config["deadline"])
    set_run_font(deadline, color=(0xFF, 0x00, 0x00))
    suffix = deadline_paragraph.add_run("前发送回本邮箱")
    set_run_font(suffix)
    add_email_hyperlink(deadline_paragraph, config["email"])
    period = deadline_paragraph.add_run("。")
    set_run_font(period)

    add_text_paragraph(document, "感谢各位老师对研究生会工作的大力支持！", first_indent=24)
    add_text_paragraph(
        document,
        "北京大学深圳研究生院研究生会",
        align=WD_ALIGN_PARAGRAPH.RIGHT,
    )
    add_text_paragraph(document, "常务代表委员会", align=WD_ALIGN_PARAGRAPH.RIGHT)
    add_text_paragraph(document, f"{year}年{month}月", align=WD_ALIGN_PARAGRAPH.RIGHT)

    for section in config["sections"]:
        heading = document.add_paragraph(section["title"], style="Heading 1")
        set_paragraph_format(heading, keep_with_next=True)
        for run in heading.runs:
            set_run_font(run, size=14, bold=True, color=(0x00, 0x7B, 0xB8))

        for index, proposal in enumerate(section["proposals"], start=1):
            add_text_paragraph(
                document,
                f"{index}.《{proposal['title']}》",
                bold=True,
                first_indent=21,
                left_indent=15,
                keep_with_next=True,
            )
            add_text_paragraph(
                document,
                "问题描述：",
                style="List Paragraph",
                bold=True,
                keep_with_next=True,
            )
            body_paragraph = add_text_paragraph(
                document,
                proposal["body"],
                style="List Paragraph",
                first_indent=24,
                highlight=bool(proposal.get("highlight")),
            )
            body_paragraph.paragraph_format.widow_control = True
            add_text_paragraph(
                document,
                f"【{department}】回复：",
                bold=True,
                first_indent=21,
                left_indent=15,
                keep_with_next=True,
            )
            contacts = "、".join(proposal_contacts(proposal, responses))
            supplement = f"【权益部补充】：提案人联系方式：{contacts}"
            if proposal.get("history_reference"):
                supplement += f"；{proposal['history_reference']}"
            has_images = any(image_downloads.get(row) for row in proposal["source_rows"])
            add_text_paragraph(
                document,
                supplement,
                bold=True,
                first_indent=21,
                left_indent=15,
                keep_with_next=has_images,
            )
            for source_row in proposal["source_rows"]:
                add_related_images(document, image_downloads.get(source_row, []))
            add_text_paragraph(document)

    output.parent.mkdir(parents=True, exist_ok=True)
    validate_rendered_language(document)
    document.save(output)


def write_audit(config, responses, covered, image_downloads, output: Path):
    eligible = sorted(row for row, item in responses.items() if valid_contact(item["contact"]))
    excluded_no_contact = [
        {
            "excel_row": row,
            "source_id": item["source_id"],
            "reason": contact_error(item["contact"]) or "联系方式无效",
        }
        for row, item in responses.items()
        if not valid_contact(item["contact"])
    ]
    excluded_manual = [
        {
            "excel_row": item["row"],
            "source_id": responses[item["row"]]["source_id"],
            "reason": item["reason"],
        }
        for item in config.get("excluded_rows", [])
    ]
    excluded = excluded_no_contact + excluded_manual
    proposals = []
    for section in config["sections"]:
        for proposal in section["proposals"]:
            proposal_images = [
                image
                for row in proposal["source_rows"]
                for image in image_downloads.get(row, [])
            ]
            proposals.append(
                {
                    "section": section["title"],
                    "title": proposal["title"],
                    "source_rows": proposal["source_rows"],
                    "source_count": len(proposal["source_rows"]),
                    "highlight": bool(proposal.get("highlight")),
                    "history_reference": proposal.get("history_reference"),
                    "history_sources": proposal.get("history_sources", []),
                    "measure_basis": proposal.get("measure_basis", []),
                    "image_count": len(proposal_images),
                    "image_source_rows": [
                        row for row in proposal["source_rows"] if image_downloads.get(row)
                    ],
                }
            )
    audit = {
        "period": f"{config['year']}-{config['month']:02d}",
        "source_response_count": len(responses),
        "eligible_response_count": len(eligible),
        "included_response_count": len(set(covered)),
        "excluded_response_count": len(excluded),
        "eligible_rows": eligible,
        "excluded": excluded,
        "allowed_reused_rows": config.get("allowed_reused_rows", []),
        "all_eligible_rows_accounted_for": set(eligible)
        <= set(covered) | {item["excel_row"] for item in excluded_manual},
        "proposal_count": len(proposals),
        "highlighted_proposal_count": sum(item["highlight"] for item in proposals),
        "downloaded_image_count": sum(len(items) for items in image_downloads.values()),
        "downloaded_images": [
            {"excel_row": row, **image}
            for row, images in sorted(image_downloads.items())
            for image in images
        ],
        "embedded_image_count": sum(item["image_count"] for item in proposals),
        "proposals": proposals,
        "historical_evidence_notes": config.get("historical_evidence_notes", []),
        "historical_search_log": config.get("historical_search_log", []),
    }
    audit_path = output.with_suffix(".audit.json")
    audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
    return audit_path


def main():
    parser = argparse.ArgumentParser(description="将权益问卷配置渲染为提案 Word")
    parser.add_argument("--xlsx", type=Path, required=True)
    parser.add_argument("--template", type=Path, required=True)
    parser.add_argument("--config", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    config = json.loads(args.config.read_text(encoding="utf-8"))
    responses = load_responses(args.xlsx)
    covered = validate_config(config, responses)
    image_downloads = download_response_images(responses, set(covered), args.output)
    build_document(args.template, config, responses, image_downloads, args.output)
    audit_path = write_audit(config, responses, covered, image_downloads, args.output)
    print(json.dumps({"docx": str(args.output.resolve()), "audit": str(audit_path.resolve())}, ensure_ascii=False))


if __name__ == "__main__":
    main()
