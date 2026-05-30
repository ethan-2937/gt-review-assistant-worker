from __future__ import annotations

import argparse
import json
import re
from collections import OrderedDict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

NOTE_PATTERNS = [
    re.compile(r"(?:附注|注释)\s*([0-9]{1,3})\s*[：:、.．\-－—_\s]*([^\\/\n\r]*)"),
    re.compile(r"^\s*([0-9]{1,3})\s*[：:、.．\-－—_\s]+([^\\/\n\r]+)"),
]
STRICT_NOTE_PATTERN = re.compile(r"^\s*(?:\(?([0-9]{1,3})\)?|附注\s*([0-9]{1,3}))\s*[：:、.．\-－—)]\s*([^\\/\n\r]+)")
LOOSE_NOTE_PATTERN = re.compile(r"^\s*([0-9]{1,3})\s*[：:、.．\-－—_\s]+\s*([^\\/\n\r]+)")
CHINESE_RE = re.compile(r"[\u4e00-\u9fff]")
NOTE_TITLE_BAD_TOKENS = ("亿元", "万元", "同比", "%", "％", "人民币", "本公司", "公司财务", "第", "页")
HEADER_HINTS = (
    "项目", "名称", "类别", "类型", "期末", "期初", "年末", "年初", "本期", "上期",
    "本年", "上年", "本集团", "金额", "比例", "账面", "余额", "发生额", "202", "单位",
)
ROW_STUB_HINTS = ("项目", "名称", "类别", "类型", "单位", "行次", "被投资单位", "债券名称", "证券名称")
NUMERIC_RE = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?$")


@dataclass
class CellItem:
    value: Any = None
    locator: str = ""


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Decimal):
        return format(value, "f")
    text = str(value)
    text = text.replace("\u00a0", " ").replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def key_text(value: Any) -> str:
    text = clean_text(value)
    text = text.replace("／", "/").replace("\\", "/")
    text = re.sub(r"\s*/\s*", "/", text)
    return text


def is_empty(value: Any) -> bool:
    return clean_text(value) == ""


def looks_numeric(value: Any) -> bool:
    text = clean_text(value)
    if text in {"", "-", "—", "--", "不适用"}:
        return False
    if isinstance(value, (int, float, Decimal)):
        return True
    return bool(NUMERIC_RE.match(text.replace("，", ",")))


def cell_value(item: Any) -> Any:
    if isinstance(item, dict):
        return item.get("value")
    if isinstance(item, CellItem):
        return item.value
    return item


def cell_locator(item: Any) -> str:
    if isinstance(item, dict):
        return clean_text(item.get("locator"))
    if isinstance(item, CellItem):
        return clean_text(item.locator)
    return ""


def row_values(row: list[Any]) -> list[str]:
    return [clean_text(cell_value(cell)) for cell in row]


def non_empty_count(row: list[Any]) -> int:
    return sum(1 for value in row_values(row) if value)


def trim_matrix(matrix: list[list[Any]]) -> list[list[Any]]:
    rows = [row for row in matrix if any(not is_empty(cell_value(cell)) for cell in row)]
    if not rows:
        return []
    max_len = max(len(row) for row in rows)
    padded = [row + [CellItem()] * (max_len - len(row)) for row in rows]
    used_cols = []
    for col_idx in range(max_len):
        if any(not is_empty(cell_value(row[col_idx])) for row in padded):
            used_cols.append(col_idx)
    return [[row[col_idx] for col_idx in used_cols] for row in padded]


def split_blocks(rows: list[list[Any]], blank_gap: int = 2) -> list[list[list[Any]]]:
    blocks: list[list[list[Any]]] = []
    current: list[list[Any]] = []
    blank_count = 0
    for row in rows:
        if any(not is_empty(cell_value(cell)) for cell in row):
            if blank_count >= blank_gap and current:
                blocks.append(current)
                current = []
            blank_count = 0
            current.append(row)
        else:
            blank_count += 1
            if current and blank_count < blank_gap:
                current.append(row)
    if current:
        blocks.append(current)
    return blocks


def infer_note_from_path(path: str | Path) -> tuple[str, str]:
    p = Path(path)
    candidates = [p.stem]
    candidates.extend(part for part in reversed(p.parts) if part not in {p.anchor, p.name})
    for candidate in candidates:
        note = infer_note_from_text(candidate)
        if note[0]:
            return note
    return "", p.stem


def valid_note_candidate(note_no: str, title: str, *, strict: bool = False) -> bool:
    try:
        number = int(note_no)
    except ValueError:
        return False
    if number <= 0 or number > 100:
        return False
    title = clean_text(title)
    if not title or len(title) > 50:
        return False
    if strict and not CHINESE_RE.search(title):
        return False
    if strict and any(token in title for token in NOTE_TITLE_BAD_TOKENS):
        return False
    if strict and re.search(r"\d", title):
        return False
    return True


def infer_note_from_heading(text: str, *, allow_space_separator: bool = False, strict: bool = True) -> tuple[str, str]:
    normalized = clean_text(text)
    if not normalized:
        return "", ""
    if "附注" in normalized or "注释" in normalized:
        for pattern in NOTE_PATTERNS[:1]:
            match = pattern.search(normalized)
            if match:
                note_no = match.group(1).lstrip("0") or match.group(1)
                title = clean_text(match.group(2))
                title = re.sub(r"\.(xlsx|xls|pdf)$", "", title, flags=re.IGNORECASE).strip(" -_—－")
                if valid_note_candidate(note_no, title, strict=False):
                    return note_no, title
    pattern = LOOSE_NOTE_PATTERN if allow_space_separator else STRICT_NOTE_PATTERN
    match = pattern.search(normalized)
    if not match:
        return "", ""
    if pattern is STRICT_NOTE_PATTERN:
        note_no = first_non_blank(match.group(1), match.group(2))
        title = clean_text(match.group(3))
    else:
        note_no = clean_text(match.group(1))
        title = clean_text(match.group(2))
    note_no = note_no.lstrip("0") or note_no
    title = re.sub(r"\.(xlsx|xls|pdf)$", "", title, flags=re.IGNORECASE).strip(" -_—－")
    if valid_note_candidate(note_no, title, strict=strict):
        return note_no, title
    return "", ""


def infer_note_from_text(text: str) -> tuple[str, str]:
    normalized = clean_text(text)
    for pattern in NOTE_PATTERNS:
        match = pattern.search(normalized)
        if match:
            note_no = match.group(1).lstrip("0") or match.group(1)
            title = clean_text(match.group(2))
            title = re.sub(r"\.(xlsx|xls|pdf)$", "", title, flags=re.IGNORECASE).strip(" -_—－")
            if valid_note_candidate(note_no, title, strict=False):
                return note_no, title
    return "", ""


def note_key(note: dict[str, Any]) -> str:
    return clean_text(note.get("noteNo")) or clean_text(note.get("noteName"))


def merge_note_payloads(notes: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: OrderedDict[str, dict[str, Any]] = OrderedDict()
    for note in notes:
        key = note_key(note)
        if not key:
            key = f"unknown_{len(merged) + 1}"
        if key not in merged:
            merged[key] = {
                "noteNo": clean_text(note.get("noteNo")),
                "noteName": clean_text(note.get("noteName")),
                "sourceLocator": clean_text(note.get("sourceLocator")),
                "tables": [],
            }
        target = merged[key]
        if not target.get("noteName") and note.get("noteName"):
            target["noteName"] = clean_text(note.get("noteName"))
        target["tables"].extend(note.get("tables") or [])
    return list(merged.values())


def dedupe_key(base: str, seen: dict[str, int]) -> str:
    cleaned = key_text(base) or "未命名"
    count = seen.get(cleaned, 0) + 1
    seen[cleaned] = count
    if count == 1:
        return cleaned
    return f"{cleaned}__{count}"


def first_non_blank(*values: Any) -> str:
    for value in values:
        text = clean_text(value)
        if text:
            return text
    return ""


def is_header_like(row: list[Any]) -> bool:
    values = row_values(row)
    filled = [value for value in values if value]
    if len(filled) < 2:
        return False
    joined = "/".join(filled)
    if any(hint in joined for hint in HEADER_HINTS):
        return True
    numeric_count = sum(1 for value in filled if looks_numeric(value))
    return numeric_count == 0 and len(filled) >= 2


def leading_title(matrix: list[list[Any]]) -> tuple[str, list[list[Any]]]:
    titles: list[str] = []
    idx = 0
    while idx < min(3, len(matrix)):
        row = matrix[idx]
        values = [value for value in row_values(row) if value]
        if 0 < len(values) <= 2 and not is_header_like(row) and not any(looks_numeric(v) for v in values):
            titles.append(" / ".join(values))
            idx += 1
            continue
        break
    return " / ".join(titles), matrix[idx:]


def detect_header_rows(matrix: list[list[Any]]) -> int:
    if not matrix:
        return 0
    count = 0
    for idx, row in enumerate(matrix[:4]):
        if is_header_like(row):
            count += 1
            continue
        if idx == 0 and non_empty_count(row) >= 2:
            count = 1
        break
    return max(count, 1 if len(matrix) > 1 else 0)


def detect_row_label_col(header_rows: list[list[Any]], data_rows: list[list[Any]]) -> int:
    if header_rows:
        max_cols = max(len(row) for row in header_rows)
        for col_idx in range(max_cols):
            header_text = "/".join(clean_text(cell_value(row[col_idx])) for row in header_rows if col_idx < len(row))
            if any(hint in header_text for hint in ROW_STUB_HINTS):
                return col_idx
    if data_rows:
        max_cols = max(len(row) for row in data_rows)
        for col_idx in range(max_cols):
            if any(col_idx < len(row) and not is_empty(cell_value(row[col_idx])) for row in data_rows):
                return col_idx
    return 0


def column_label(header_rows: list[list[Any]], col_idx: int) -> str:
    parts: list[str] = []
    for row in header_rows:
        if col_idx < len(row):
            text = clean_text(cell_value(row[col_idx]))
            if text and (not parts or parts[-1] != text):
                parts.append(text)
    return " / ".join(parts)


def matrix_to_table_payload(
    matrix: list[list[Any]],
    *,
    default_table_title: str,
    table_order: int,
    locator_prefix: str,
    include_empty_cells: bool = False,
) -> dict[str, Any] | None:
    matrix = trim_matrix(matrix)
    if len(matrix) < 2:
        return None

    title_from_rows, content = leading_title(matrix)
    table_title = first_non_blank(title_from_rows, default_table_title, f"表{table_order}")
    if len(content) < 2:
        return None

    header_count = detect_header_rows(content)
    header_rows = content[:header_count]
    data_rows = content[header_count:]
    if not data_rows:
        return None

    label_col = detect_row_label_col(header_rows, data_rows)
    max_cols = max(len(row) for row in content)

    value_cols: list[int] = []
    for col_idx in range(max_cols):
        if col_idx == label_col:
            continue
        has_header = bool(column_label(header_rows, col_idx))
        has_value = any(col_idx < len(row) and not is_empty(cell_value(row[col_idx])) for row in data_rows)
        if has_header or has_value:
            value_cols.append(col_idx)

    if not value_cols:
        return None

    rows: list[dict[str, Any]] = []
    row_key_seen: dict[str, int] = {}
    row_key_by_data_index: dict[int, str] = {}
    for data_idx, row in enumerate(data_rows, start=1):
        if label_col >= len(row):
            continue
        row_label = clean_text(cell_value(row[label_col]))
        has_value = any(col_idx < len(row) and not is_empty(cell_value(row[col_idx])) for col_idx in value_cols)
        if not row_label or not has_value:
            continue
        row_key = dedupe_key(row_label, row_key_seen)
        row_key_by_data_index[data_idx] = row_key
        rows.append({
            "rowKey": row_key,
            "rowPath": row_label,
            "rowLeaf": row_label.split("/")[-1].strip(),
            "rowOrder": data_idx,
            "sourceLocator": cell_locator(row[label_col]) or f"{locator_prefix}:r{data_idx + header_count}c{label_col + 1}",
        })

    columns: list[dict[str, Any]] = []
    col_key_seen: dict[str, int] = {}
    col_key_by_index: dict[int, str] = {}
    for order, col_idx in enumerate(value_cols, start=1):
        label = first_non_blank(column_label(header_rows, col_idx), f"第{col_idx + 1}列")
        col_key = dedupe_key(label, col_key_seen)
        col_key_by_index[col_idx] = col_key
        locator = ""
        for row in header_rows:
            if col_idx < len(row) and not is_empty(cell_value(row[col_idx])):
                locator = cell_locator(row[col_idx])
                break
        columns.append({
            "columnKey": col_key,
            "columnPath": label,
            "columnLeaf": label.split("/")[-1].strip(),
            "columnOrder": order,
            "sourceLocator": locator or f"{locator_prefix}:header:c{col_idx + 1}",
        })

    cells: list[dict[str, Any]] = []
    for data_idx, row in enumerate(data_rows, start=1):
        row_key = row_key_by_data_index.get(data_idx)
        if not row_key:
            continue
        row_label = next((r["rowPath"] for r in rows if r["rowKey"] == row_key), "")
        for col_idx in value_cols:
            value = cell_value(row[col_idx]) if col_idx < len(row) else None
            value_text = clean_text(value)
            if not include_empty_cells and not value_text:
                continue
            col_key = col_key_by_index[col_idx]
            col_label = next((c["columnPath"] for c in columns if c["columnKey"] == col_key), "")
            locator = cell_locator(row[col_idx]) if col_idx < len(row) else ""
            cells.append({
                "rowKey": row_key,
                "columnKey": col_key,
                "rowPath": row_label,
                "columnPath": col_label,
                "valueText": value_text,
                "normalizedValue": value_text,
                "sourceLocator": locator or f"{locator_prefix}:r{data_idx + header_count}c{col_idx + 1}",
            })

    if not rows and not columns and not cells:
        return None

    return {
        "tableTitle": table_title,
        "tableOrder": table_order,
        "sourceLocator": locator_prefix,
        "rows": rows,
        "columns": columns,
        "cells": cells,
    }


def write_json(payload: dict[str, Any], output: str | Path) -> None:
    path = Path(output)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def add_common_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--input", required=True, help="输入文件或目录")
    parser.add_argument("--output", required=True, help="输出 JSON 文件")
    parser.add_argument("--include-empty-cells", action="store_true", help="把空白单元格也写入 cells")


def iter_files(input_path: str | Path, suffixes: tuple[str, ...]) -> list[Path]:
    path = Path(input_path)
    if path.is_file():
        return [path]
    files: list[Path] = []
    for suffix in suffixes:
        files.extend(path.rglob(f"*{suffix}"))
    return sorted(file for file in files if not file.name.startswith("~$"))
