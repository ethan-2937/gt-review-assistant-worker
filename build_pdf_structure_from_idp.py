from __future__ import annotations

import argparse
import hashlib
import html.parser
import json
import mimetypes
import re
import uuid
from collections import defaultdict
from pathlib import Path
from typing import Any
from urllib import request as urlrequest
from urllib.error import HTTPError, URLError

from gt_structure_worker.common import (
    CellItem,
    clean_text,
    first_non_blank,
    infer_note_from_heading,
    infer_note_from_path,
    iter_files,
    matrix_to_table_payload,
    merge_note_payloads,
    write_json,
)

PDF_SUFFIXES = (".pdf",)
TABLE_TYPE_TOKENS = ("table", "grid")
TEXT_KEYS = ("text", "content", "value", "caption", "title", "name")
CELL_LIST_KEYS = ("cells", "table_cells", "tableCells")
ROW_MATRIX_KEYS = ("rows", "data", "table_data", "tableData")
HTML_KEYS = ("html", "table_html", "tableHtml")
MARKDOWN_KEYS = ("markdown", "md")


class SimpleTableHTMLParser(html.parser.HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.rows: list[list[str]] = []
        self._current_row: list[str] | None = None
        self._current_cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "tr":
            self._current_row = []
        elif tag in {"td", "th"} and self._current_row is not None:
            self._current_cell = []

    def handle_data(self, data: str) -> None:
        if self._current_cell is not None:
            self._current_cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in {"td", "th"} and self._current_row is not None and self._current_cell is not None:
            self._current_row.append(clean_text("".join(self._current_cell)))
            self._current_cell = None
        elif tag == "tr" and self._current_row is not None:
            if any(clean_text(cell) for cell in self._current_row):
                self.rows.append(self._current_row)
            self._current_row = None


def post_json(url: str, payload: dict[str, Any]) -> dict[str, Any]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urlrequest.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    try:
        with urlrequest.urlopen(req, timeout=120) as response:
            body = response.read().decode("utf-8")
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code}: {body}") from exc
    except URLError as exc:
        raise RuntimeError(f"backend request failed: {exc}") from exc
    return json.loads(body)


def maybe_import(payload: dict[str, Any], backend: str | None, project_id: str | None, *, replace_side: bool = False) -> None:
    if not backend or not project_id:
        return
    endpoint = "import" if replace_side else "import-notes"
    url = backend.rstrip("/") + f"/api/projects/{project_id}/structures/{endpoint}"
    result = post_json(url, payload)
    if not result.get("success"):
        raise RuntimeError(result.get("message") or "import failed")
    print(f"imported endpoint={endpoint} side={payload.get('side')} result={result.get('data')}")


def maybe_compare(backend: str | None, project_id: str | None) -> None:
    if not backend or not project_id:
        return
    url = backend.rstrip("/") + f"/api/projects/{project_id}/structures/compare"
    result = post_json(url, {})
    if not result.get("success"):
        raise RuntimeError(result.get("message") or "compare failed")
    print(f"compare result={result.get('data')}")


def bool_form(value: bool) -> str:
    return "true" if value else "false"


def build_multipart_body(
    fields: dict[str, Any],
    file_field: str,
    file_path: Path,
) -> tuple[bytes, str]:
    boundary = "----gt-review-idp-" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for key, value in fields.items():
        if value is None or value == "":
            continue
        chunks.append(f"--{boundary}\r\n".encode("utf-8"))
        chunks.append(f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode("utf-8"))
        chunks.append(str(value).encode("utf-8"))
        chunks.append(b"\r\n")

    content_type = mimetypes.guess_type(file_path.name)[0] or "application/pdf"
    chunks.append(f"--{boundary}\r\n".encode("utf-8"))
    chunks.append(
        (
            f'Content-Disposition: form-data; name="{file_field}"; '
            f'filename="{file_path.name}"\r\n'
            f"Content-Type: {content_type}\r\n\r\n"
        ).encode("utf-8")
    )
    chunks.append(file_path.read_bytes())
    chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def submit_idp_sync(
    *,
    idp_base: str,
    pdf_path: Path,
    page_range: str,
    timeout_seconds: int,
    engine: str,
    engine_param: str,
    tenant_id: str,
    authorization: str,
    include_blocks: bool = True,
) -> dict[str, Any]:
    url = idp_base.rstrip("/") + "/v1/tasks/sync"
    fields: dict[str, Any] = {
        "priority": "low",
        "timeout": timeout_seconds,
        "page_range": page_range,
        "include_blocks": bool_form(include_blocks),
        "strict_backend": bool_form(False),
    }
    if engine:
        fields[engine_param] = engine

    body, content_type = build_multipart_body(fields, "file", pdf_path)
    headers = {
        "Content-Type": content_type,
        "User-Agent": "gt-review-assistant-worker/1.0",
    }
    if tenant_id:
        headers["X-Tenant-Id"] = tenant_id
    if authorization:
        headers["Authorization"] = authorization

    req = urlrequest.Request(url, data=body, headers=headers, method="POST")
    try:
        with urlrequest.urlopen(req, timeout=timeout_seconds + 30) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"IDP HTTP {exc.code}: {raw}") from exc
    except URLError as exc:
        raise RuntimeError(f"IDP request failed: {exc}") from exc

    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"IDP returned non-JSON response prefix: {raw[:500]}") from exc


def get_idp_text(url: str, *, timeout_seconds: int = 120, proxy_download: bool = False) -> str:
    headers = {"User-Agent": "gt-review-assistant-worker/1.0"}
    if proxy_download:
        headers["X-Download-Proxy"] = "1"
    req = urlrequest.Request(url, headers=headers, method="GET")
    try:
        with urlrequest.urlopen(req, timeout=timeout_seconds) as response:
            return response.read().decode("utf-8", errors="replace")
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"IDP HTTP {exc.code}: {body}") from exc
    except URLError as exc:
        raise RuntimeError(f"IDP request failed: {exc}") from exc


def parse_jsonl_blocks(text: str) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            item = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(item, dict):
            blocks.append(item)
    return blocks


def fetch_idp_blocks_artifact(raw: dict[str, Any], idp_base: str, timeout_seconds: int) -> list[dict[str, Any]]:
    doc_id = clean_text(raw.get("doc_id"))
    task_id = clean_text(raw.get("task_id"))
    if not doc_id:
        return []

    suffix = f"?task_id={task_id}" if task_id else ""
    candidates = [
        f"{idp_base.rstrip('/')}/v1/documents/{doc_id}/files/blocks_jsonl{suffix}",
        f"{idp_base.rstrip('/')}/v1/documents/{doc_id}/files/blocks.jsonl{suffix}",
    ]
    for url in candidates:
        try:
            text = get_idp_text(url, timeout_seconds=timeout_seconds, proxy_download=True)
        except RuntimeError:
            continue
        blocks = parse_jsonl_blocks(text)
        if blocks:
            return blocks
    return []


def enrich_raw_with_artifacts(raw: dict[str, Any], args: argparse.Namespace, raw_path: Path) -> dict[str, Any]:
    if raw.get("blocks"):
        return raw
    blocks = fetch_idp_blocks_artifact(raw, args.idp_base, args.timeout)
    if not blocks:
        return raw
    raw["blocks"] = blocks
    raw.setdefault("artifact_fetch", {})["blocks_jsonl"] = {
        "status": "fetched",
        "block_count": len(blocks),
    }
    raw_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
    return raw


def cache_key(pdf_path: Path, page_range: str, engine: str, engine_param: str) -> str:
    stat = pdf_path.stat()
    seed = {
        "path": str(pdf_path.resolve()),
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
        "page_range": page_range,
        "engine": engine,
        "engine_param": engine_param,
    }
    return hashlib.sha256(json.dumps(seed, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:24]


def load_or_submit_idp(
    *,
    args: argparse.Namespace,
    pdf_path: Path,
    cache_dir: Path,
) -> tuple[dict[str, Any], Path, bool]:
    if args.raw_input:
        raw_path = Path(args.raw_input)
        return json.loads(raw_path.read_text(encoding="utf-8-sig")), raw_path, True

    key = cache_key(pdf_path, args.page_range, args.engine, args.engine_param)
    raw_path = cache_dir / f"{key}.raw.json"
    if raw_path.exists() and not args.force:
        return json.loads(raw_path.read_text(encoding="utf-8-sig")), raw_path, True

    if args.dry_run:
        print("DRY RUN: would submit one IDP task")
        print(f"idp_base={args.idp_base}")
        print(f"pdf={pdf_path}")
        print(f"page_range={args.page_range}")
        print(f"engine={args.engine} via {args.engine_param}")
        raise SystemExit(0)

    raw = submit_idp_sync(
        idp_base=args.idp_base,
        pdf_path=pdf_path,
        page_range=args.page_range,
        timeout_seconds=args.timeout,
        engine=args.engine,
        engine_param=args.engine_param,
        tenant_id=args.tenant_id,
        authorization=args.authorization,
        include_blocks=True,
    )
    raw_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
    return raw, raw_path, False


def scalar_text(value: Any) -> str:
    if isinstance(value, str):
        return clean_text(value)
    if isinstance(value, (int, float)):
        return clean_text(value)
    if isinstance(value, list):
        return clean_text(" ".join(scalar_text(item) for item in value))
    if isinstance(value, dict):
        for key in TEXT_KEYS:
            if key in value:
                text = scalar_text(value.get(key))
                if text:
                    return text
    return ""


def first_value(obj: dict[str, Any], keys: tuple[str, ...]) -> Any:
    for key in keys:
        if key in obj and obj.get(key) not in (None, ""):
            return obj.get(key)
    return None


def int_value(value: Any) -> int | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        return int(float(str(value)))
    except ValueError:
        return None


def page_number(obj: dict[str, Any], default: int = 1) -> int:
    zero_based = int_value(first_value(obj, ("page_index", "page_idx", "pageIndex")))
    if zero_based is not None:
        return zero_based + 1
    one_based = int_value(first_value(obj, ("page_no", "page_num", "pageNumber", "page")))
    if one_based is not None:
        return max(1, one_based)
    return default


def raw_page_count(raw: dict[str, Any]) -> int | None:
    candidates = [
        raw.get("page_count"),
        raw.get("pages"),
        raw.get("doc_tree", {}).get("metadata", {}).get("page_count"),
        raw.get("doc_tree", {}).get("processing", {}).get("file_info", {}).get("pages"),
        raw.get("doc_tree", {}).get("processing", {}).get("file_info", {}).get("page_count"),
    ]
    for candidate in candidates:
        count = int_value(candidate)
        if count is not None:
            return count
    return None


def page_range_highest(page_range: str) -> int | None:
    highest: int | None = None
    for part in page_range.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            _, right = part.split("-", 1)
            value = int_value(right.strip())
        else:
            value = int_value(part)
        if value is None:
            continue
        highest = value if highest is None else max(highest, value)
    return highest


def bbox_text(obj: dict[str, Any]) -> str:
    value = first_value(obj, ("bbox", "box", "rect", "position"))
    if isinstance(value, (list, tuple)):
        nums = []
        for part in value[:4]:
            try:
                nums.append(str(round(float(part), 2)))
            except (TypeError, ValueError):
                pass
        return ",".join(nums)
    if isinstance(value, dict):
        candidates = [value.get(k) for k in ("x0", "y0", "x1", "y1")]
        if all(v is not None for v in candidates):
            return ",".join(str(round(float(v), 2)) for v in candidates)
        x1y1x2y2 = [value.get(k) for k in ("x1", "y1", "x2", "y2")]
        if all(v is not None for v in x1y1x2y2):
            return ",".join(str(round(float(v), 2)) for v in x1y1x2y2)
        xywh = [value.get(k) for k in ("x", "y", "width", "height")]
        if all(v is not None for v in xywh):
            x, y, w, h = [float(v) for v in xywh]
            return ",".join(str(round(v, 2)) for v in (x, y, x + w, y + h))
    return ""


def is_table_candidate(obj: dict[str, Any]) -> bool:
    type_text = " ".join(
        clean_text(obj.get(key)).lower()
        for key in ("type", "block_type", "blockType", "category", "kind", "label", "layout_type")
        if obj.get(key) is not None
    )
    if any(token in type_text for token in TABLE_TYPE_TOKENS):
        return True
    if any(key in obj and isinstance(obj.get(key), list) for key in CELL_LIST_KEYS):
        return True
    if any(key in obj and looks_like_html_table(clean_text(obj.get(key))) for key in HTML_KEYS):
        return True
    if any(key in obj and looks_like_markdown_table(clean_text(obj.get(key))) for key in MARKDOWN_KEYS + TEXT_KEYS):
        return True
    return False


def walk_dicts(obj: Any) -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    if isinstance(obj, dict):
        found.append(obj)
        for value in obj.values():
            found.extend(walk_dicts(value))
    elif isinstance(obj, list):
        for item in obj:
            found.extend(walk_dicts(item))
    return found


def collect_page_text(raw: dict[str, Any]) -> dict[int, list[str]]:
    page_texts: dict[int, list[str]] = defaultdict(list)
    for obj in walk_dicts(raw):
        page = page_number(obj, default=0)
        if page <= 0:
            continue
        type_text = clean_text(first_value(obj, ("type", "block_type", "blockType", "category", "kind"))).lower()
        if "table" in type_text:
            continue
        text = scalar_text(obj)
        if text:
            page_texts[page].append(text)
    return page_texts


def infer_page_notes(raw: dict[str, Any], pdf_path: Path, forced_note_no: str, forced_note_name: str) -> dict[int, tuple[str, str]]:
    if forced_note_no:
        return defaultdict(lambda: (forced_note_no, forced_note_name or forced_note_no))

    file_note_no, file_note_name = infer_note_from_path(pdf_path)
    page_texts = collect_page_text(raw)
    page_notes: dict[int, tuple[str, str]] = {}
    current_note_no = file_note_no
    current_note_name = file_note_name
    for page in sorted(page_texts):
        for text in page_texts[page][:120]:
            for line in re.split(r"[\r\n]+", text):
                note_no, note_name = infer_note_from_heading(line, allow_space_separator=False, strict=True)
                if note_no:
                    current_note_no = note_no
                    current_note_name = first_non_blank(note_name, current_note_name)
                    break
            if current_note_no:
                break
        if current_note_no:
            page_notes[page] = (current_note_no, first_non_blank(current_note_name, pdf_path.stem))
    return page_notes


def looks_like_html_table(text: str) -> bool:
    lowered = text.lower()
    return "<table" in lowered and ("<tr" in lowered or "<td" in lowered or "<th" in lowered)


def html_to_matrix(text: str, pdf_path: Path, page: int, table_index: int) -> list[list[CellItem]]:
    parser = SimpleTableHTMLParser()
    parser.feed(text)
    matrix: list[list[CellItem]] = []
    for row_idx, row in enumerate(parser.rows, start=1):
        matrix.append([
            CellItem(value=value, locator=f"{pdf_path.name}!p{page}!idp!table{table_index}!r{row_idx}c{col_idx}")
            for col_idx, value in enumerate(row, start=1)
        ])
    return matrix


def looks_like_markdown_table(text: str) -> bool:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    pipe_lines = [line for line in lines if "|" in line]
    return len(pipe_lines) >= 2 and any(re.match(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?$", line) for line in pipe_lines[:3])


def markdown_to_matrix(text: str, pdf_path: Path, page: int, table_index: int) -> list[list[CellItem]]:
    rows: list[list[str]] = []
    for line in text.splitlines():
        stripped = line.strip()
        if "|" not in stripped:
            continue
        parts = [clean_text(part) for part in stripped.strip("|").split("|")]
        if not any(parts):
            continue
        if all(re.match(r"^:?-{2,}:?$", part.replace(" ", "")) for part in parts if part):
            continue
        rows.append(parts)
    matrix: list[list[CellItem]] = []
    for row_idx, row in enumerate(rows, start=1):
        matrix.append([
            CellItem(value=value, locator=f"{pdf_path.name}!p{page}!idp!table{table_index}!r{row_idx}c{col_idx}")
            for col_idx, value in enumerate(row, start=1)
        ])
    return matrix


def cell_list_to_matrix(cells: list[Any], pdf_path: Path, page: int, table_index: int) -> list[list[CellItem]]:
    parsed: list[tuple[int, int, str, str]] = []
    rows_seen: list[int] = []
    cols_seen: list[int] = []
    for cell in cells:
        if not isinstance(cell, dict):
            continue
        row = int_value(first_value(cell, ("row", "row_index", "row_idx", "rowIndex", "start_row", "row_start")))
        col = int_value(first_value(cell, ("col", "column", "column_index", "col_index", "colIndex", "start_col", "col_start")))
        if row is None or col is None:
            continue
        text = scalar_text(cell)
        bbox = bbox_text(cell)
        parsed.append((row, col, text, bbox))
        rows_seen.append(row)
        cols_seen.append(col)
    if not parsed:
        return []

    row_offset = 1 if min(rows_seen) == 0 else 0
    col_offset = 1 if min(cols_seen) == 0 else 0
    max_row = max(row + row_offset for row, _, _, _ in parsed)
    max_col = max(col + col_offset for _, col, _, _ in parsed)
    matrix = [[CellItem() for _ in range(max_col)] for _ in range(max_row)]
    for row, col, text, bbox in parsed:
        row_no = row + row_offset
        col_no = col + col_offset
        locator = f"{pdf_path.name}!p{page}!idp!table{table_index}!r{row_no}c{col_no}"
        if bbox:
            locator += f"!bbox={bbox}"
        matrix[row_no - 1][col_no - 1] = CellItem(value=text, locator=locator)
    return matrix


def rows_to_matrix(rows: list[Any], pdf_path: Path, page: int, table_index: int) -> list[list[CellItem]]:
    matrix: list[list[CellItem]] = []
    for row_idx, row in enumerate(rows, start=1):
        if isinstance(row, dict):
            values = row.get("cells") or row.get("columns") or row.get("values") or []
        else:
            values = row
        if not isinstance(values, list):
            continue
        matrix_row: list[CellItem] = []
        for col_idx, value in enumerate(values, start=1):
            text = scalar_text(value)
            bbox = bbox_text(value) if isinstance(value, dict) else ""
            locator = f"{pdf_path.name}!p{page}!idp!table{table_index}!r{row_idx}c{col_idx}"
            if bbox:
                locator += f"!bbox={bbox}"
            matrix_row.append(CellItem(value=text, locator=locator))
        matrix.append(matrix_row)
    return matrix


def table_matrix(obj: dict[str, Any], pdf_path: Path, page: int, table_index: int) -> list[list[CellItem]]:
    for key in CELL_LIST_KEYS:
        cells = obj.get(key)
        if isinstance(cells, list):
            matrix = cell_list_to_matrix(cells, pdf_path, page, table_index)
            if matrix:
                return matrix

    for key in ROW_MATRIX_KEYS:
        rows = obj.get(key)
        if isinstance(rows, list):
            matrix = rows_to_matrix(rows, pdf_path, page, table_index)
            if matrix:
                return matrix

    for key in HTML_KEYS:
        text = clean_text(obj.get(key))
        if looks_like_html_table(text):
            matrix = html_to_matrix(text, pdf_path, page, table_index)
            if matrix:
                return matrix

    for key in MARKDOWN_KEYS + TEXT_KEYS:
        text = clean_text(obj.get(key))
        if looks_like_markdown_table(text):
            matrix = markdown_to_matrix(text, pdf_path, page, table_index)
            if matrix:
                return matrix
    return []


def table_title(obj: dict[str, Any], page: int, table_index: int) -> str:
    return first_non_blank(
        scalar_text(obj.get("caption")),
        scalar_text(obj.get("title")),
        scalar_text(obj.get("name")),
        f"IDP page {page} table {table_index}",
    )


def table_locator(pdf_path: Path, obj: dict[str, Any], page: int, table_index: int) -> str:
    locator = f"{pdf_path.name}!p{page}!idp!table{table_index}"
    bbox = bbox_text(obj)
    if bbox:
        locator += f"!bbox={bbox}"
    block_id = first_value(obj, ("block_id", "blockId", "id"))
    if block_id:
        locator += f"!block={block_id}"
    return locator


def extract_table_candidates(raw: dict[str, Any]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    seen: set[int] = set()
    for obj in walk_dicts(raw):
        if id(obj) in seen:
            continue
        if is_table_candidate(obj):
            candidates.append(obj)
            seen.add(id(obj))
    return candidates


def idp_raw_to_payload(
    raw: dict[str, Any],
    pdf_path: Path,
    *,
    forced_note_no: str = "",
    forced_note_name: str = "",
    include_empty_cells: bool = False,
) -> tuple[dict[str, Any], dict[str, Any]]:
    page_notes = infer_page_notes(raw, pdf_path, forced_note_no, forced_note_name)
    file_note_no, file_note_name = infer_note_from_path(pdf_path)
    notes: dict[str, dict[str, Any]] = {}
    table_order_by_note: dict[str, int] = defaultdict(lambda: 1)
    candidates = extract_table_candidates(raw)
    converted_count = 0

    for table_index, obj in enumerate(candidates, start=1):
        page = page_number(obj, default=1)
        note_no, note_name = page_notes.get(page, ("", ""))
        note_no = first_non_blank(forced_note_no, note_no, file_note_no, "unknown")
        note_name = first_non_blank(forced_note_name, note_name, file_note_name, pdf_path.stem)
        note_key = note_no or note_name

        if note_key not in notes:
            notes[note_key] = {
                "noteNo": note_no,
                "noteName": note_name,
                "sourceLocator": f"{pdf_path.name}!p{page}",
                "tables": [],
            }

        matrix = table_matrix(obj, pdf_path, page, table_index)
        if not matrix:
            continue

        order = table_order_by_note[note_key]
        payload = matrix_to_table_payload(
            matrix,
            default_table_title=table_title(obj, page, table_index),
            table_order=order,
            locator_prefix=table_locator(pdf_path, obj, page, table_index),
            include_empty_cells=include_empty_cells,
        )
        if not payload:
            continue
        notes[note_key]["tables"].append(payload)
        table_order_by_note[note_key] += 1
        converted_count += 1

    payload = {
        "side": "PDF",
        "sourceFilePath": str(pdf_path),
        "notes": merge_note_payloads(notes.values()),
    }
    summary = {
        "tableCandidates": len(candidates),
        "convertedTables": converted_count,
        "noteCount": len(payload["notes"]),
        "rowCount": sum(len(table.get("rows") or []) for note in payload["notes"] for table in note.get("tables") or []),
        "columnCount": sum(len(table.get("columns") or []) for note in payload["notes"] for table in note.get("tables") or []),
        "cellCount": sum(len(table.get("cells") or []) for note in payload["notes"] for table in note.get("tables") or []),
    }
    return payload, summary


def merge_payloads(payloads: list[dict[str, Any]], source_path: str | Path) -> dict[str, Any]:
    notes: list[dict[str, Any]] = []
    for payload in payloads:
        notes.extend(payload.get("notes") or [])
    return {
        "side": "PDF",
        "sourceFilePath": str(source_path),
        "notes": merge_note_payloads(notes),
    }


def write_summary(output_path: Path, rows: list[dict[str, Any]], payload: dict[str, Any]) -> None:
    lines = [
        "# IDP PDF Structure Import Summary",
        "",
        "This run uses one IDP request at a time. No concurrent requests are used.",
        "",
        f"- output: `{output_path}`",
        f"- notes: {len(payload.get('notes') or [])}",
        f"- tables: {sum(len(note.get('tables') or []) for note in payload.get('notes') or [])}",
        f"- rows: {sum(len(table.get('rows') or []) for note in payload.get('notes') or [] for table in note.get('tables') or [])}",
        f"- columns: {sum(len(table.get('columns') or []) for note in payload.get('notes') or [] for table in note.get('tables') or [])}",
        f"- cells: {sum(len(table.get('cells') or []) for note in payload.get('notes') or [] for table in note.get('tables') or [])}",
        "",
        "## Files",
        "",
    ]
    for row in rows:
        lines.extend([
            f"### {row['pdf']}",
            "",
            f"- raw: `{row['rawPath']}`",
            f"- cacheHit: {row['cacheHit']}",
            f"- rawPageCount: {row.get('rawPageCount')}",
            f"- requestedPageRange: `{row.get('requestedPageRange')}`",
            f"- tableCandidates: {row['tableCandidates']}",
            f"- convertedTables: {row['convertedTables']}",
            f"- notes: {row['noteCount']}",
            f"- rows: {row['rowCount']}",
            f"- columns: {row['columnCount']}",
            f"- cells: {row['cellCount']}",
            "",
        ])
    output_path.write_text("\n".join(lines), encoding="utf-8")


def validate_non_empty_before_import(payload: dict[str, Any], rows: list[dict[str, Any]], page_range: str) -> None:
    note_count = len(payload.get("notes") or [])
    table_count = sum(len(note.get("tables") or []) for note in payload.get("notes") or [])
    if note_count and table_count:
        return

    hints = [
        "IDP did not produce any convertible PDF tables, so nothing was imported into the backend.",
        "Check idp_pdf_structure_summary.md and the raw JSON under idp_raw_cache.",
    ]
    for row in rows:
        page_count = row.get("rawPageCount")
        highest = page_range_highest(page_range or "")
        if page_count is not None and highest is not None and highest >= page_count:
            hints.append(
                f"Requested page_range `{page_range}` is outside this PDF's page_count={page_count}. "
                "For a split one-page PDF, use `--page-range 0` or `--allow-full-document`."
            )
            break
    hints.append("If this empty result came from cache, add `--force` after correcting the page range.")
    raise SystemExit("\n".join(hints))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Call IDP with one request at a time and convert PDF table blocks into structure JSON."
    )
    parser.add_argument("--idp-base", default="http://8.140.53.175:23035")
    parser.add_argument("--pdf-input", required=True, help="PDF file or directory")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--page-range", default="", help="IDP page_range, for example 159-165. Required unless --allow-full-document is used.")
    parser.add_argument("--allow-full-document", action="store_true", help="Allow parsing the whole PDF. Use carefully.")
    parser.add_argument("--engine", default="idp-v2.5-pro")
    parser.add_argument(
        "--engine-param",
        choices=["pipeline", "backend", "parser_name", "pipeline_name", "template"],
        default="pipeline",
        help="Which multipart field receives --engine. Change this if IDP routing expects another field.",
    )
    parser.add_argument("--timeout", type=int, default=600)
    parser.add_argument("--tenant-id", default="")
    parser.add_argument("--authorization", default="", help="Authorization header, if needed. Example: Bearer xxx")
    parser.add_argument("--note-no", default="", help="Force all converted tables into this note number.")
    parser.add_argument("--note-name", default="", help="Force all converted tables into this note name.")
    parser.add_argument("--include-empty-cells", action="store_true")
    parser.add_argument("--force", action="store_true", help="Ignore local raw-result cache and submit IDP again.")
    parser.add_argument("--dry-run", action="store_true", help="Print the single IDP request that would be submitted.")
    parser.add_argument("--raw-input", default="", help="Use an existing IDP raw JSON instead of submitting a request.")
    parser.add_argument("--backend", default="")
    parser.add_argument("--project-id", default="")
    parser.add_argument("--compare", action="store_true")
    parser.add_argument(
        "--allow-replace-side",
        action="store_true",
        help="Use full-side /structures/import instead of safe note-level /structures/import-notes.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.raw_input and not args.page_range and not args.allow_full_document:
        raise SystemExit("--page-range is required unless --allow-full-document is used. This prevents accidental full-PDF IDP jobs.")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    cache_dir = output_dir / "idp_raw_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = iter_files(args.pdf_input, PDF_SUFFIXES)
    if not pdf_files:
        raise SystemExit(f"No PDF files found: {args.pdf_input}")

    payloads: list[dict[str, Any]] = []
    summary_rows: list[dict[str, Any]] = []
    for pdf_path in pdf_files:
        raw, raw_path, cache_hit = load_or_submit_idp(args=args, pdf_path=pdf_path, cache_dir=cache_dir)
        raw = enrich_raw_with_artifacts(raw, args, raw_path)
        payload, summary = idp_raw_to_payload(
            raw,
            pdf_path,
            forced_note_no=args.note_no,
            forced_note_name=args.note_name,
            include_empty_cells=args.include_empty_cells,
        )
        payloads.append(payload)
        summary_rows.append({
            "pdf": str(pdf_path),
            "rawPath": str(raw_path),
            "cacheHit": cache_hit,
            "rawPageCount": raw_page_count(raw),
            "requestedPageRange": args.page_range,
            **summary,
        })
        print(
            f"IDP PDF file={pdf_path.name} cache={cache_hit} "
            f"tables={summary['convertedTables']}/{summary['tableCandidates']} "
            f"rows={summary['rowCount']} cells={summary['cellCount']}"
        )

    final_payload = merge_payloads(payloads, args.pdf_input)
    output_path = output_dir / "pdf_structure_idp.json"
    summary_path = output_dir / "idp_pdf_structure_summary.md"
    write_json(final_payload, output_path)
    write_summary(summary_path, summary_rows, final_payload)
    print(f"output json={output_path}")
    print(f"output summary={summary_path}")

    if args.backend and args.project_id:
        validate_non_empty_before_import(final_payload, summary_rows, args.page_range)
    maybe_import(final_payload, args.backend, args.project_id, replace_side=args.allow_replace_side)
    if args.compare:
        maybe_compare(args.backend, args.project_id)


if __name__ == "__main__":
    main()
