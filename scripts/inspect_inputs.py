from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook


def enum_name(value):
    return getattr(value, "name", None) if value is not None else None


def length_pt(value):
    return round(value.pt, 3) if value is not None else None


def color_value(color):
    if color is None or color.type is None:
        return None
    return {
        "type": enum_name(color.type),
        "rgb": str(color.rgb) if color.rgb is not None else None,
        "theme": enum_name(color.theme_color),
    }


def inspect_workbook(path: Path):
    workbook = load_workbook(path, data_only=False)
    result = {
        "path": str(path),
        "properties": {
            "creator": workbook.properties.creator,
            "title": workbook.properties.title,
        },
        "sheets": [],
    }
    for sheet in workbook.worksheets:
        rows = []
        for row in sheet.iter_rows():
            values = [cell.value for cell in row]
            if any(value is not None for value in values):
                rows.append(
                    {
                        "row": row[0].row,
                        "values": values,
                        "styles": [cell.style_id for cell in row],
                    }
                )
        result["sheets"].append(
            {
                "title": sheet.title,
                "dimensions": sheet.calculate_dimension(),
                "merged_cells": [str(item) for item in sheet.merged_cells.ranges],
                "freeze_panes": str(sheet.freeze_panes) if sheet.freeze_panes else None,
                "rows": rows,
            }
        )
    return result


def inspect_document(path: Path, include_xml: bool = False):
    document = Document(path)
    paragraphs = []
    for index, paragraph in enumerate(document.paragraphs):
        fmt = paragraph.paragraph_format
        item = {
                "index": index,
                "text": paragraph.text,
                "style": paragraph.style.name if paragraph.style else None,
                "alignment": enum_name(paragraph.alignment),
                "format": {
                    "left_indent_pt": length_pt(fmt.left_indent),
                    "right_indent_pt": length_pt(fmt.right_indent),
                    "first_line_indent_pt": length_pt(fmt.first_line_indent),
                    "space_before_pt": length_pt(fmt.space_before),
                    "space_after_pt": length_pt(fmt.space_after),
                    "line_spacing": fmt.line_spacing,
                    "line_spacing_rule": enum_name(fmt.line_spacing_rule),
                    "keep_together": fmt.keep_together,
                    "keep_with_next": fmt.keep_with_next,
                    "page_break_before": fmt.page_break_before,
                    "widow_control": fmt.widow_control,
                },
                "runs": [
                    {
                        "text": run.text,
                        "bold": run.bold,
                        "italic": run.italic,
                        "underline": run.underline,
                        "font_name": run.font.name,
                        "font_size_pt": length_pt(run.font.size),
                        "font_color": color_value(run.font.color),
                        "highlight": enum_name(run.font.highlight_color),
                    }
                    for run in paragraph.runs
                ],
            }
        if include_xml:
            item["xml"] = paragraph._p.xml
        paragraphs.append(item)

    styles = []
    for style in document.styles:
        if style.type.name != "PARAGRAPH":
            continue
        pf = style.paragraph_format
        styles.append(
            {
                "name": style.name,
                "font_name": style.font.name,
                "font_size_pt": length_pt(style.font.size),
                "bold": style.font.bold,
                "color": color_value(style.font.color),
                "alignment": enum_name(pf.alignment),
                "first_line_indent_pt": length_pt(pf.first_line_indent),
                "space_before_pt": length_pt(pf.space_before),
                "space_after_pt": length_pt(pf.space_after),
                "line_spacing": pf.line_spacing,
                "line_spacing_rule": enum_name(pf.line_spacing_rule),
            }
        )

    sections = []
    for section in document.sections:
        sections.append(
            {
                "page_width_pt": length_pt(section.page_width),
                "page_height_pt": length_pt(section.page_height),
                "top_margin_pt": length_pt(section.top_margin),
                "bottom_margin_pt": length_pt(section.bottom_margin),
                "left_margin_pt": length_pt(section.left_margin),
                "right_margin_pt": length_pt(section.right_margin),
                "header_distance_pt": length_pt(section.header_distance),
                "footer_distance_pt": length_pt(section.footer_distance),
                "header": [p.text for p in section.header.paragraphs],
                "footer": [p.text for p in section.footer.paragraphs],
            }
        )

    return {
        "path": str(path),
        "core_properties": {
            "author": document.core_properties.author,
            "title": document.core_properties.title,
        },
        "sections": sections,
        "paragraphs": paragraphs,
        "tables": [
            [[cell.text for cell in row.cells] for row in table.rows]
            for table in document.tables
        ],
        "styles": styles,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", type=Path, required=True)
    parser.add_argument("--docx", type=Path, required=True)
    parser.add_argument("--xml", action="store_true")
    args = parser.parse_args()
    print(
        json.dumps(
            {
                "workbook": inspect_workbook(args.xlsx),
                "document": inspect_document(args.docx, include_xml=args.xml),
            },
            ensure_ascii=False,
            indent=2,
            default=str,
        )
    )


if __name__ == "__main__":
    main()
