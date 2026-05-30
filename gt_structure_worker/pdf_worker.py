from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import fitz

from gt_structure_worker.common import (
    CellItem,
    add_common_args,
    clean_text,
    first_non_blank,
    infer_note_from_heading,
    infer_note_from_path,
    iter_files,
    matrix_to_table_payload,
    merge_note_payloads,
    valid_note_candidate,
    write_json,
)

PDF_SUFFIXES = (".pdf",)


def page_heading_note(page: fitz.Page) -> tuple[str, str]:
    try:
        text = page.get_text("text") or ""
    except Exception:
        return "", ""
    lines = [clean_text(line) for line in text.splitlines() if clean_text(line)]
    scan_lines = lines[:90]
    for line in scan_lines:
        note_no, note_name = infer_note_from_heading(line, allow_space_separator=False, strict=True)
        if note_no:
            return note_no, note_name
    for idx, line in enumerate(scan_lines[:-1]):
        if not line.isdigit():
            continue
        note_no = line.lstrip("0") or line
        title = scan_lines[idx + 1]
        if valid_note_candidate(note_no, title, strict=True):
            return note_no, title
    return "", ""


def table_to_matrix(file_path: Path, page_no: int, table_index: int, table: Any) -> list[list[CellItem]]:
    try:
        extracted = table.extract()
    except Exception:
        return []
    matrix: list[list[CellItem]] = []
    for row_idx, row in enumerate(extracted or [], start=1):
        matrix_row: list[CellItem] = []
        for col_idx, value in enumerate(row or [], start=1):
            locator = f"{file_path.name}!p{page_no}!table{table_index}!r{row_idx}c{col_idx}"
            matrix_row.append(CellItem(value=value, locator=locator))
        matrix.append(matrix_row)
    return matrix


def text_lines_to_table(
    file_path: Path,
    page_no: int,
    lines: list[str],
    *,
    table_order: int,
) -> dict[str, Any] | None:
    rows = []
    cells = []
    columns = [{
        "columnKey": "文本内容",
        "columnPath": "文本内容",
        "columnLeaf": "文本内容",
        "columnOrder": 1,
        "sourceLocator": f"{file_path.name}!p{page_no}!text",
    }]
    for idx, line in enumerate(lines, start=1):
        text = clean_text(line)
        if not text:
            continue
        row_key = f"第{idx}行"
        rows.append({
            "rowKey": row_key,
            "rowPath": row_key,
            "rowLeaf": row_key,
            "rowOrder": idx,
            "sourceLocator": f"{file_path.name}!p{page_no}!line{idx}",
        })
        cells.append({
            "rowKey": row_key,
            "columnKey": "文本内容",
            "rowPath": row_key,
            "columnPath": "文本内容",
            "valueText": text,
            "normalizedValue": text,
            "sourceLocator": f"{file_path.name}!p{page_no}!line{idx}",
        })
    if not rows:
        return None
    return {
        "tableTitle": f"第{page_no}页文本段落",
        "tableOrder": table_order,
        "sourceLocator": f"{file_path.name}!p{page_no}!text",
        "rows": rows,
        "columns": columns,
        "cells": cells,
    }


def parse_pdf(
    file_path: Path,
    *,
    include_text_lines: bool = False,
    include_empty_cells: bool = False,
) -> list[dict[str, Any]]:
    file_note_no, file_note_name = infer_note_from_path(file_path)
    notes: dict[str, dict[str, Any]] = {}
    table_order_by_note: dict[str, int] = {}
    current_note_no = file_note_no
    current_note_name = file_note_name
    notes_started = bool(file_note_no)

    doc = fitz.open(file_path)
    try:
        for page_index in range(doc.page_count):
            page = doc.load_page(page_index)
            page_no = page_index + 1
            page_text = page.get_text("text") or ""
            if not notes_started and (
                "合并财务报表主要项目附注" in page_text
                or "财务报表主要项目附注" in page_text
                or "主要项目附注" in page_text
            ):
                notes_started = True
            if not notes_started:
                continue
            page_note_no, page_note_name = page_heading_note(page)
            if page_note_no:
                current_note_no = page_note_no
                current_note_name = first_non_blank(page_note_name, current_note_name)
            if not current_note_no:
                continue
            note_key = current_note_no or f"file:{file_path.stem}"
            if note_key not in notes:
                notes[note_key] = {
                    "noteNo": current_note_no,
                    "noteName": first_non_blank(current_note_name, file_path.stem),
                    "sourceLocator": f"{file_path.name}!p{page_no}",
                    "tables": [],
                }
                table_order_by_note[note_key] = 1

            table_finder = None
            try:
                table_finder = page.find_tables()
            except Exception:
                table_finder = None
            tables = list(getattr(table_finder, "tables", []) or [])

            for table_index, table in enumerate(tables, start=1):
                order = table_order_by_note[note_key]
                matrix = table_to_matrix(file_path, page_no, table_index, table)
                bbox = ""
                if getattr(table, "bbox", None):
                    bbox = ",".join(str(round(x, 2)) for x in table.bbox)
                payload = matrix_to_table_payload(
                    matrix,
                    default_table_title=f"第{page_no}页 表{table_index}",
                    table_order=order,
                    locator_prefix=f"{file_path.name}!p{page_no}!table{table_index}" + (f"!bbox={bbox}" if bbox else ""),
                    include_empty_cells=include_empty_cells,
                )
                if payload:
                    notes[note_key]["tables"].append(payload)
                    table_order_by_note[note_key] += 1

            if include_text_lines and not tables:
                text = page.get_text("text") or ""
                lines = [line for line in text.splitlines() if clean_text(line)]
                order = table_order_by_note[note_key]
                payload = text_lines_to_table(file_path, page_no, lines, table_order=order)
                if payload:
                    notes[note_key]["tables"].append(payload)
                    table_order_by_note[note_key] += 1
    finally:
        doc.close()

    return list(notes.values())


def build_pdf_payload(
    input_path: str | Path,
    *,
    include_text_lines: bool = False,
    include_empty_cells: bool = False,
) -> dict[str, Any]:
    files = iter_files(input_path, PDF_SUFFIXES)
    notes: list[dict[str, Any]] = []
    for file_path in files:
        notes.extend(parse_pdf(
            file_path,
            include_text_lines=include_text_lines,
            include_empty_cells=include_empty_cells,
        ))
    return {
        "side": "PDF",
        "sourceFilePath": str(Path(input_path)),
        "notes": merge_note_payloads(notes),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="把 PDF 财报附注拆成后端可导入的结构 JSON")
    add_common_args(parser)
    parser.add_argument("--include-text-lines", action="store_true", help="如果页面没有表格，把文本行作为文本段落表导入")
    args = parser.parse_args()

    payload = build_pdf_payload(
        args.input,
        include_text_lines=args.include_text_lines,
        include_empty_cells=args.include_empty_cells,
    )
    write_json(payload, args.output)
    print(f"PDF notes={len(payload['notes'])} output={args.output}")


if __name__ == "__main__":
    main()
