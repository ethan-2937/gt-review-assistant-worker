from __future__ import annotations

import argparse
import json
from pathlib import Path
from urllib import request as urlrequest
from urllib.error import HTTPError, URLError

from gt_structure_worker.excel_worker import build_excel_payload
from gt_structure_worker.pdf_worker import build_pdf_payload
from gt_structure_worker.common import write_json


def post_json(url: str, payload: dict) -> dict:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urlrequest.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    try:
        with urlrequest.urlopen(req, timeout=60) as response:
            body = response.read().decode("utf-8")
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code}: {body}") from exc
    except URLError as exc:
        raise RuntimeError(f"请求后端失败：{exc}") from exc
    return json.loads(body)


def maybe_import(payload: dict, backend: str | None, project_id: str | None) -> None:
    if not backend or not project_id:
        return
    url = backend.rstrip("/") + f"/api/projects/{project_id}/structures/import"
    result = post_json(url, payload)
    if not result.get("success"):
        raise RuntimeError(result.get("message") or "导入失败")
    print(f"imported side={payload.get('side')} result={result.get('data')}")


def maybe_compare(backend: str | None, project_id: str | None) -> None:
    if not backend or not project_id:
        return
    url = backend.rstrip("/") + f"/api/projects/{project_id}/structures/compare"
    result = post_json(url, {})
    if not result.get("success"):
        raise RuntimeError(result.get("message") or "生成差异失败")
    print(f"compare result={result.get('data')}")


def main() -> None:
    parser = argparse.ArgumentParser(description="统一生成 PDF/Excel 结构 JSON，也可直接导入后端")
    subparsers = parser.add_subparsers(dest="command", required=True)

    excel = subparsers.add_parser("excel", help="解析 Excel")
    excel.add_argument("--input", required=True)
    excel.add_argument("--output", required=True)
    excel.add_argument("--include-hidden", action="store_true")
    excel.add_argument("--include-empty-cells", action="store_true")
    excel.add_argument("--blank-row-gap", type=int, default=2)
    excel.add_argument("--backend", default="")
    excel.add_argument("--project-id", default="")

    pdf = subparsers.add_parser("pdf", help="解析 PDF")
    pdf.add_argument("--input", required=True)
    pdf.add_argument("--output", required=True)
    pdf.add_argument("--include-text-lines", action="store_true")
    pdf.add_argument("--include-empty-cells", action="store_true")
    pdf.add_argument("--backend", default="")
    pdf.add_argument("--project-id", default="")

    both = subparsers.add_parser("both", help="同时解析 Excel 和 PDF")
    both.add_argument("--excel-input", required=True)
    both.add_argument("--pdf-input", required=True)
    both.add_argument("--output-dir", required=True)
    both.add_argument("--include-hidden", action="store_true")
    both.add_argument("--include-text-lines", action="store_true")
    both.add_argument("--include-empty-cells", action="store_true")
    both.add_argument("--blank-row-gap", type=int, default=2)
    both.add_argument("--backend", default="")
    both.add_argument("--project-id", default="")
    both.add_argument("--compare", action="store_true", help="导入两侧后调用后端生成差异")

    args = parser.parse_args()

    if args.command == "excel":
        payload = build_excel_payload(
            args.input,
            include_hidden=args.include_hidden,
            include_empty_cells=args.include_empty_cells,
            blank_row_gap=args.blank_row_gap,
        )
        write_json(payload, args.output)
        print(f"EXCEL notes={len(payload['notes'])} output={args.output}")
        maybe_import(payload, args.backend, args.project_id)
        return

    if args.command == "pdf":
        payload = build_pdf_payload(
            args.input,
            include_text_lines=args.include_text_lines,
            include_empty_cells=args.include_empty_cells,
        )
        write_json(payload, args.output)
        print(f"PDF notes={len(payload['notes'])} output={args.output}")
        maybe_import(payload, args.backend, args.project_id)
        return

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    excel_payload = build_excel_payload(
        args.excel_input,
        include_hidden=args.include_hidden,
        include_empty_cells=args.include_empty_cells,
        blank_row_gap=args.blank_row_gap,
    )
    pdf_payload = build_pdf_payload(
        args.pdf_input,
        include_text_lines=args.include_text_lines,
        include_empty_cells=args.include_empty_cells,
    )
    excel_output = output_dir / "excel_structure.json"
    pdf_output = output_dir / "pdf_structure.json"
    write_json(excel_payload, excel_output)
    write_json(pdf_payload, pdf_output)
    print(f"EXCEL notes={len(excel_payload['notes'])} output={excel_output}")
    print(f"PDF notes={len(pdf_payload['notes'])} output={pdf_output}")

    maybe_import(excel_payload, args.backend, args.project_id)
    maybe_import(pdf_payload, args.backend, args.project_id)
    if args.compare:
        maybe_compare(args.backend, args.project_id)


if __name__ == "__main__":
    main()