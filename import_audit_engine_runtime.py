from __future__ import annotations

import argparse
import csv
import json
import platform
from pathlib import Path
from typing import Any
from urllib import request as urlrequest
from urllib.error import HTTPError, URLError


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\u00a0", " ").replace("\r", " ").replace("\n", " ")
    return " ".join(text.split()).strip()


def first_non_blank(*values: Any) -> str:
    for value in values:
        text = clean_text(value)
        if text:
            return text
    return ""


def item_key(table_title: str, row_key: str, row_path: str, row_label: str, column_key: str, column_path: str, column_label: str) -> str:
    table = first_non_blank(table_title, "default table")
    row = first_non_blank(row_key, row_path, row_label)
    column = first_non_blank(column_key, column_path, column_label)
    joined = f"{table}|{row}|{column}"
    return joined[:2000]


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
        raise RuntimeError(f"后端请求失败：{exc}") from exc
    return json.loads(body)


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def case_id_text(case: dict[str, Any]) -> str:
    return clean_text(case.get("case_id"))


def extract_clean_runtime_evidence(case: dict[str, Any]) -> dict[str, Any]:
    return (
        case.get("debug_counts", {})
        .get("direct_doc_materialization", {})
        .get("clean_runtime_evidence", {})
    ) or {}


def fact_to_item(case: dict[str, Any], fact: dict[str, Any], runtime_side: str, source_json_path: str) -> dict[str, Any]:
    table_title = clean_text(fact.get("table_context"))
    row_key = clean_text(fact.get("row_key"))
    row_label = clean_text(fact.get("row_label"))
    column_key = clean_text(fact.get("column_key"))
    column_label = clean_text(fact.get("column_label"))
    raw_value = clean_text(fact.get("raw_value"))
    value_signature = clean_text(fact.get("value_signature"))
    cell_coordinate = clean_text(fact.get("cell_coordinate"))
    locator_method = clean_text(fact.get("locator_method"))
    quote_text = clean_text(fact.get("quote_text"))
    locator = first_non_blank(cell_coordinate, locator_method)

    return {
        "caseId": case_id_text(case),
        "noteNo": case_id_text(case),
        "noteName": clean_text(case.get("case_name")),
        "level": "cell",
        "runtimeSide": runtime_side,
        "itemKey": item_key(table_title, row_key, "", row_label, column_key, "", column_label),
        "tableTitle": table_title,
        "tableProfileId": clean_text(fact.get("table_profile_id")),
        "rowKey": row_key,
        "rowLabel": row_label,
        "rowPath": row_label,
        "columnKey": column_key,
        "columnLabel": column_label,
        "columnPath": column_label,
        "valueText": raw_value,
        "normalizedValue": first_non_blank(value_signature, raw_value),
        "valueSignature": value_signature,
        "valueType": clean_text(fact.get("value_type")),
        "sourceLocator": locator,
        "cellCoordinate": cell_coordinate,
        "quoteText": quote_text,
        "sourceArtifact": "runtime_eval.json",
        "sourceJsonPath": source_json_path,
        "locatorMethod": locator_method,
        "confidenceLevel": "SAMPLE",
        "rawPayloadJson": json.dumps(fact, ensure_ascii=False),
    }


def extract_items(runtime_eval: dict[str, Any]) -> tuple[list[dict[str, Any]], int, int]:
    items: list[dict[str, Any]] = []
    source_count = 0
    target_count = 0
    for case_index, case in enumerate(runtime_eval.get("cases") or []):
        evidence = extract_clean_runtime_evidence(case)
        source_samples = evidence.get("source_fact_samples") or []
        target_samples = evidence.get("target_fact_samples") or []
        source_count += len(source_samples)
        target_count += len(target_samples)
        for idx, fact in enumerate(source_samples):
            if isinstance(fact, dict):
                items.append(fact_to_item(case, fact, "PDF", f"cases[{case_index}].debug_counts.direct_doc_materialization.clean_runtime_evidence.source_fact_samples[{idx}]"))
        for idx, fact in enumerate(target_samples):
            if isinstance(fact, dict):
                items.append(fact_to_item(case, fact, "EXCEL", f"cases[{case_index}].debug_counts.direct_doc_materialization.clean_runtime_evidence.target_fact_samples[{idx}]"))
    return items, source_count, target_count


def extract_structure_qa_counts(structure_qa_path: Path) -> dict[str, int]:
    if not structure_qa_path.exists():
        return {}
    payload = load_json(structure_qa_path)
    source_structured = 0
    target_structured = 0
    table_count = 0
    for case in payload.get("cases") or []:
        counts = case.get("counts") or {}
        tables = case.get("table_counts") or {}
        source_structured += int(counts.get("source:structured_cell_count") or 0)
        target_structured += int(counts.get("target:structured_cell_count") or 0)
        table_count += int(tables.get("table_count") or 0)
    return {
        "sourceStructuredCellCount": source_structured,
        "targetStructuredCellCount": target_structured,
        "tableCount": table_count,
    }


def build_payload(run_root: Path, run_type: str, run_status: str) -> dict[str, Any]:
    runtime_eval_path = run_root / "runtime_eval.json"
    if not runtime_eval_path.exists():
        raise FileNotFoundError(f"未找到 runtime_eval.json: {runtime_eval_path}")

    runtime_eval = load_json(runtime_eval_path)
    items, source_sample_count, target_sample_count = extract_items(runtime_eval)
    qa_counts = extract_structure_qa_counts(run_root / "structure_qa.json")

    return {
        "runKey": run_root.name,
        "runRoot": str(run_root),
        "runType": run_type,
        "runStatus": run_status,
        "datasetKey": clean_text(runtime_eval.get("dataset_key")),
        "versionLabel": clean_text(runtime_eval.get("version")),
        "caseCount": len(runtime_eval.get("cases") or []),
        "artifactCompleteness": "SAMPLE_ONLY",
        "confidenceLevel": "SAMPLE",
        "sourceHost": platform.node(),
        "sourceSampleCount": source_sample_count,
        "targetSampleCount": target_sample_count,
        "sourceStructuredCellCount": qa_counts.get("sourceStructuredCellCount"),
        "targetStructuredCellCount": qa_counts.get("targetStructuredCellCount"),
        "tableCount": qa_counts.get("tableCount"),
        "items": items,
    }


def write_outputs(payload: dict[str, Any], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = output_dir / "runtime_structure_payload.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    csv_path = output_dir / "runtime_structure_items.csv"
    fields = [
        "caseId", "noteNo", "runtimeSide", "level", "itemKey", "tableTitle", "rowKey", "rowLabel",
        "columnKey", "columnLabel", "valueText", "valueSignature", "cellCoordinate", "locatorMethod",
        "sourceJsonPath",
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for item in payload["items"]:
            writer.writerow({field: item.get(field, "") for field in fields})

    summary_path = output_dir / "runtime_import_summary.md"
    summary_path.write_text(
        "\n".join([
            "# Runtime Structure Import Summary",
            "",
            f"- runKey: `{payload.get('runKey')}`",
            f"- runRoot: `{payload.get('runRoot')}`",
            f"- datasetKey: `{payload.get('datasetKey')}`",
            f"- versionLabel: `{payload.get('versionLabel')}`",
            f"- artifactCompleteness: `{payload.get('artifactCompleteness')}`",
            f"- confidenceLevel: `{payload.get('confidenceLevel')}`",
            f"- sourceSampleCount: {payload.get('sourceSampleCount')}",
            f"- targetSampleCount: {payload.get('targetSampleCount')}",
            f"- sourceStructuredCellCount: {payload.get('sourceStructuredCellCount')}",
            f"- targetStructuredCellCount: {payload.get('targetStructuredCellCount')}",
            f"- tableCount: {payload.get('tableCount')}",
            f"- itemCount: {len(payload.get('items') or [])}",
            "",
            "注意：当前导入的是 runtime_eval.json 里的 fact samples，不代表完整 runtime 识别率。",
        ]),
        encoding="utf-8",
    )
    print(f"output json={json_path}")
    print(f"output csv={csv_path}")
    print(f"output summary={summary_path}")


def maybe_import(payload: dict[str, Any], backend: str, project_id: str) -> None:
    if not backend or not project_id:
        return
    url = backend.rstrip("/") + f"/api/projects/{project_id}/runtime-runs/import"
    result = post_json(url, payload)
    if not result.get("success"):
        raise RuntimeError(result.get("message") or "runtime 导入失败")
    print(f"import result={result.get('data')}")


def main() -> None:
    parser = argparse.ArgumentParser(description="从 audit-engine run root 抽取 runtime structure sample 并导入后端")
    parser.add_argument("--run-root", required=True, help="audit-engine 运行结果目录")
    parser.add_argument("--output-dir", required=True, help="输出 JSON/CSV/summary 的目录")
    parser.add_argument("--backend", default="", help="后端地址，例如 http://localhost:18081")
    parser.add_argument("--project-id", default="", help="gt-review-assistant 项目 ID")
    parser.add_argument("--run-type", default="smoke")
    parser.add_argument("--run-status", default="diagnostic")
    args = parser.parse_args()

    payload = build_payload(Path(args.run_root), args.run_type, args.run_status)
    write_outputs(payload, Path(args.output_dir))
    maybe_import(payload, args.backend, args.project_id)
    print(
        "RUNTIME items={items} source_samples={source} target_samples={target}".format(
            items=len(payload["items"]),
            source=payload.get("sourceSampleCount"),
            target=payload.get("targetSampleCount"),
        )
    )


if __name__ == "__main__":
    main()
