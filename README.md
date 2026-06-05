# GT Review Assistant Worker

这个目录负责把原始 Excel/PDF 自动拆成后端 `/structures/import` 能直接导入的 JSON。

## 输出 JSON 结构

输出是一个对象，不是单纯数组：

```json
{
  "side": "EXCEL",
  "sourceFilePath": "D:/data-annotation/2025.06/...",
  "notes": [
    {
      "noteNo": "36",
      "noteName": "租赁负债",
      "tables": [
        {
          "tableTitle": "租赁负债到期分析",
          "rows": [],
          "columns": [],
          "cells": []
        }
      ]
    }
  ]
}
```

这正好可以粘贴到前端“导入结构”弹窗，也可以用命令直接导入后端。

## 安装依赖

```powershell
cd D:\audit-engine\gt-review-assistant\worker
python -m pip install -r requirements.txt
```

如果本机没有全局 `python`，Codex 本地 runtime 可以用：

```powershell
C:\Users\23885\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
```

## 只生成 Excel JSON

```powershell
python build_structure_json.py excel `
  --input "D:\data-annotation\2025.06\2025上半年-附注分割版-标黄-仅2025-复核" `
  --output "D:\audit-engine\gt-review-assistant\workspace\202506_excel_structure.json"
```

## 只生成 PDF JSON

```powershell
python build_structure_json.py pdf `
  --input "D:\data-annotation\2025.06\pdf分割版目录" `
  --output "D:\audit-engine\gt-review-assistant\workspace\202506_pdf_structure.json"
```

如果某些 PDF 页面没有表格，但你也想把文本段落放进去：

```powershell
python build_structure_json.py pdf `
  --input "D:\data-annotation\2025.06\pdf分割版目录" `
  --output "D:\audit-engine\gt-review-assistant\workspace\202506_pdf_structure.json" `
  --include-text-lines
```

## 同时生成两侧 JSON

```powershell
python build_structure_json.py both `
  --excel-input "D:\data-annotation\2025.06\excel目录" `
  --pdf-input "D:\data-annotation\2025.06\pdf目录" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\202506_structure_json"
```

## 生成后直接导入后端

先确保后端已启动，并且已经在前端创建了项目，拿到 `projectId`。

```powershell
python build_structure_json.py both `
  --excel-input "D:\data-annotation\2025.06\excel目录" `
  --pdf-input "D:\data-annotation\2025.06\pdf目录" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\202506_structure_json" `
  --backend "http://localhost:18081" `
  --project-id 1 `
  --compare
```

执行后会：

1. 生成 `excel_structure.json`
2. 生成 `pdf_structure.json`
3. 导入后端
4. 调用后端重新生成 PDF vs Excel 差异

## 当前解析策略

### Excel

- 支持 `.xlsx` 和 `.xlsm`
- 支持输入单个文件或目录递归扫描
- 从文件名、父目录、sheet 名、前几行文本推断附注号和附注名称
- 根据空行拆分表/小节
- 自动识别行名列、列名、数据格
- 输出定位，例如 `4 衍生金融工具.xlsx!Sheet1!C26`

### PDF

- 使用 PyMuPDF 的 `page.find_tables()` 抽取表格
- 支持输入单个 PDF 或目录递归扫描
- 从文件名、父目录、页面前几行推断附注号和附注名称
- 输出定位，例如 `xxx.pdf!p3!table1!r5c2`
- 可选 `--include-text-lines` 把无表格页面作为文本段落导入

## 注意

这是一版 MVP worker，目标是先把结构可视化跑通。

PDF 表格解析天然比 Excel 难，如果某些 PDF 表格没有识别好，后续再针对具体附注补规则。



## Parse PDF With Company IDP (single request, no concurrency)

If the PyMuPDF PDF result shows row names like `row 1`, `row 2`, etc., it means the script did not really recover business row names. Use the company IDP service as the PDF structure source.

New script:

```text
build_pdf_structure_from_idp.py
```

Safety defaults:

- Submits one IDP task at a time. No concurrent requests are used.
- Requires `--page-range` by default to avoid accidental full-PDF jobs.
- Caches raw IDP results in `idp_raw_cache`. Re-running the same PDF + page range + engine does not call IDP again.
- Use `--force` only when you really want to submit again.
- Use `--dry-run` to print the request without calling IDP.
- Does not call pause_all, clear_all, delete, cancel, or any global task-control endpoint.
- If the sync response has empty `blocks`, it reads IDP `blocks_jsonl` artifact and converts tables from there.
- Backend import uses safe note-level `/structures/import-notes` by default, so other PDF notes are not deleted.
- Use `--allow-replace-side` only when you intentionally want to call full-side `/structures/import`.

Dry run first:

```powershell
cd D:\audit-engine\gt-review-assistant\worker

python build_pdf_structure_from_idp.py `
  --idp-base "http://8.140.53.175:23035" `
  --pdf-input "D:\path\to\202506-report.pdf" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\idp_parse_note36" `
  --page-range "159-165" `
  --note-no 36 `
  --note-name "lease liabilities" `
  --dry-run
```

Then remove `--dry-run` to submit one IDP task:

```powershell
python build_pdf_structure_from_idp.py `
  --idp-base "http://8.140.53.175:23035" `
  --pdf-input "D:\path\to\202506-report.pdf" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\idp_parse_note36" `
  --page-range "159-165" `
  --note-no 36 `
  --note-name "lease liabilities"
```

Import into local backend after parsing. This safely replaces only the note(s) in `pdf_structure_idp.json`:

```powershell
python build_pdf_structure_from_idp.py `
  --idp-base "http://8.140.53.175:23035" `
  --pdf-input "D:\path\to\202506-report.pdf" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\idp_parse_note36" `
  --page-range "159-165" `
  --note-no 36 `
  --note-name "lease liabilities" `
  --backend "http://localhost:8080" `
  --project-id 2 `
  --compare
```

Outputs:

```text
pdf_structure_idp.json
idp_pdf_structure_summary.md
idp_raw_cache/*.raw.json
```

If the IDP router expects the engine in a different form field, change this option:

```powershell
--engine-param backend
```

Allowed values: `pipeline`, `backend`, `parser_name`, `pipeline_name`, `template`.


Full-side replacement is still available, but use it carefully:

```powershell
python build_pdf_structure_from_idp.py `
  --idp-base "http://8.140.53.175:23035" `
  --pdf-input "D:\path\to\all-pdf-notes" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\idp_parse_all_pdf" `
  --allow-full-document `
  --backend "http://localhost:8080" `
  --project-id 2 `
  --compare `
  --allow-replace-side
```

## Apply Problem GT Proposal Safely

New script:

```text
apply_problem_gt_proposal.py
```

It reads the backend exported `problem_gt` proposal JSON and creates a new `final_gt` workbook copy with approved add rows appended. It never overwrites the original workbook and does not delete or rewrite existing rows.

Example:

```powershell
python apply_problem_gt_proposal.py `
  --gt-workbook "D:\data-annotation\final_gt_202506.xlsx" `
  --proposal-json "D:\audit-engine\gt-review-assistant\workspace\problem_gt_export_check\proposal.json" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\problem_gt_apply_check"
```

Outputs:

- `*_problem_gt_applied_YYYYMMDD-HHMMSS.xlsx`
- `applied_add_rows.csv`
- `skipped_duplicate_rows.csv`
- `excluded_not_applied.csv`
- `unresolved_not_applied.csv`
- `apply_problem_gt_proposal_summary.md`
- `apply_problem_gt_proposal_summary.json`

Safety defaults:

- Only appends `proposal.adds`.
- `excluded` rows are exported for evidence only; no delete is performed.
- `unresolved` rows are exported for follow-up only; they are not written to final GT.
- Suspected duplicate rows already present in final GT are skipped unless `--allow-duplicates` is used.

### Auto-download proposal from backend

Instead of manually exporting `proposal.json`, the apply script can now fetch it from the backend and then create the new workbook:

```powershell
python apply_problem_gt_proposal.py `
  --backend "http://localhost:8080" `
  --project-id 2 `
  --source-run-key "problem_gt_mvp_202506_v16_fast3" `
  --gt-workbook "D:\data-annotation\final_gt_202506.xlsx" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\problem_gt_apply_from_ui"
```

The downloaded proposal is saved as `problem_gt_proposal_<run>.json` in the output directory. The original workbook is still never overwritten.

## Generate Problem GT Source Previews

Create read-only PDF/Excel screenshots for the problem-GT review page:

```powershell
python generate_problem_gt_previews.py `
  --backend "http://localhost:8080" `
  --project-id 2 `
  --source-run-key "problem_gt_mvp_202506_v16_fast3" `
  --source-root "D:\data-annotation\2025.06\2025上半年-附注分割版-标黄-仅2025-复核" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\preview_assets"
```

Outputs are written under:

```text
D:\audit-engine\gt-review-assistant\workspace\preview_assets\problem_gt\<source_run_key>\
```

The frontend reads these files through `/preview-assets/problem_gt/<source_run_key>/<candidate_key>_pdf.png` and `_excel.png`.

Safety defaults:

- No source PDF/XLSX files are modified.
- No GT workbook is modified.
- Missing/weak locators are reported in `problem_gt_preview_manifest.csv`; they simply show as "not generated" in the UI.

### Preview accuracy notes

`generate_problem_gt_previews.py` now has two fallback modes to reduce manual source lookup:

- PDF: if the locator points only to a note PDF or is empty, the worker searches the split-note PDF for row/column/value text and crops the nearest matching area when possible.
- Excel: if there is no explicit cell reference, the worker searches worksheet rows by row label and numeric value tokens, then renders a nearby table screenshot.

The backend can also run this worker directly through the Problem GT page button `??????`. Manual PowerShell execution is still useful for debugging or custom output directories.
