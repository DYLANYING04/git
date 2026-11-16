from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from build_proposal import (
    DEFENSIVE_PHRASES,
    UNSUPPORTED_MEASURES,
    load_responses,
    validate_config,
)


EMOTIONAL_PHRASES = {
    "太难吃",
    "修建皇宫",
    "钱没地方花",
    "最丑的院区",
    "不患寡而患不均",
    "很不友好",
}


def validate_document(config, responses, docx_path: Path):
    errors = []
    document = Document(docx_path)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)

    if "反应" in text:
        errors.append("正文存在“反应”，权益反馈语境应使用“反映”")
    for phrase in sorted(EMOTIONAL_PHRASES):
        if phrase in text:
            errors.append(f"正文残留情绪化原句：{phrase}")
    for phrase in sorted(DEFENSIVE_PHRASES):
        if phrase in text:
            errors.append(f"正文含防御性或模板化措辞：{phrase}")
    for phrase in sorted(UNSUPPORTED_MEASURES):
        if phrase in text:
            errors.append(f"正文含禁止自行补充的措施：{phrase}")

    expected_titles = []
    expected_contacts = []
    expected_highlighted_bodies = set()
    expected_image_count = 0
    for section in config["sections"]:
        if section["title"] not in text:
            errors.append(f"缺少栏目：{section['title']}")
        for proposal in section["proposals"]:
            expected_titles.append(proposal["title"])
            for row in proposal["source_rows"]:
                contact = responses[row]["contact"]
                if contact not in expected_contacts:
                    expected_contacts.append(contact)
            if proposal.get("highlight"):
                expected_highlighted_bodies.add(proposal["body"])
            for row in proposal["source_rows"]:
                image_count = len(responses[row].get("image_urls", []))
                expected_image_count += image_count

    for title in expected_titles:
        if f"《{title}》" not in text:
            errors.append(f"缺少提案标题：{title}")
    for contact in expected_contacts:
        if contact not in text:
            errors.append(f"缺少联系方式：{contact}")

    actual_highlighted_bodies = set()
    for paragraph in document.paragraphs:
        highlighted = [run for run in paragraph.runs if run.font.highlight_color == WD_COLOR_INDEX.YELLOW]
        if highlighted:
            if len(highlighted) != len(paragraph.runs):
                errors.append(f"存在部分高亮段落：{paragraph.text[:30]}")
            actual_highlighted_bodies.add(paragraph.text)
    if actual_highlighted_bodies != expected_highlighted_bodies:
        errors.append(
            "高亮正文与配置不一致："
            f"期望 {len(expected_highlighted_bodies)} 条，实际 {len(actual_highlighted_bodies)} 条"
        )

    actual_image_count = len(document.inline_shapes)
    if actual_image_count != expected_image_count:
        errors.append(
            f"相关图片数量不一致：期望 {expected_image_count} 张，实际 {actual_image_count} 张"
        )
    if "相关图片：" in text:
        errors.append("正式成稿不应显示技术性图片标签“相关图片：”")
    with ZipFile(docx_path) as archive:
        oversized_media = [
            (item.filename, item.file_size)
            for item in archive.infolist()
            if item.filename.startswith("word/media/") and item.file_size > 1024 * 1024
        ]
    for filename, size in oversized_media:
        errors.append(f"Word 内嵌图片未压缩：{filename}（{size} 字节）")

    title_pattern = re.compile(r"^(\d+)\.《.+》$")
    current_section = None
    expected_number = 1
    for paragraph in document.paragraphs:
        if paragraph.style and paragraph.style.name == "Heading 1":
            current_section = paragraph.text
            expected_number = 1
            continue
        match = title_pattern.match(paragraph.text)
        if match:
            number = int(match.group(1))
            if number != expected_number:
                errors.append(f"栏目“{current_section}”编号不连续：期望 {expected_number}，实际 {number}")
            expected_number += 1

    blank_tail = 0
    for paragraph in reversed(document.paragraphs):
        if paragraph.text.strip():
            break
        blank_tail += 1
    if blank_tail > 1:
        errors.append(f"文档末尾存在 {blank_tail} 个空段落，可能产生空白尾页")

    return errors


def main():
    parser = argparse.ArgumentParser(description="校验权益提案 Word")
    parser.add_argument("--xlsx", type=Path, required=True)
    parser.add_argument("--config", type=Path, required=True)
    parser.add_argument("--docx", type=Path, required=True)
    args = parser.parse_args()

    config = json.loads(args.config.read_text(encoding="utf-8"))
    responses = load_responses(args.xlsx)
    try:
        validate_config(config, responses)
    except ValueError as exc:
        raise SystemExit(f"配置校验失败：{exc}")

    errors = validate_document(config, responses, args.docx)
    if errors:
        print(json.dumps({"ok": False, "errors": errors}, ensure_ascii=False, indent=2))
        raise SystemExit(1)
    print(json.dumps({"ok": True, "errors": []}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
