# Changelog

## v1.5.2 - 2026-06-04

???????GT??????? worker?????????audit-engine runtime ???IDP PDF ?????????????

### Added

- ?? `import_audit_engine_runtime.py`????? audit-engine run root?? `runtime_eval.json` ?? source/target fact samples?
- runtime ?????? cell?row?column ????????????????????????
- ?? `build_pdf_structure_from_idp.py`???????????? IDP ?? PDF ???
- IDP worker ?? `--dry-run`????`blocks_jsonl` artifact ???note ??????
- ?? `generate_source_previews.py`?? locator ?? Excel ?????? PDF ?????
- ?? Pillow ??????? Excel ?????

### Changed

- runtime itemKey ?????????row ??+??column ??+??cell ??+?+??
- README ?? IDP PDF ??????????
- ??????????? `/structures/import-notes`?????????????? PDF ???

### Safety

- IDP worker ??????????
- ?????? page range???????? PDF?
- ??? pause_all?clear_all?delete?cancel ????????
- ?? worker ?????? JSON/CSV/summary??????????????

### Validation

- `python import_audit_engine_runtime.py --help` ???
- runtime case23 sample ??? 800 ? cell samples????? row/column/cell ????
- note 36 IDP ???????? 2 ???7 ??14 ?????
- ?? worker ??????? `/preview-assets/` ??? PNG ???


## v1.6 - 2026-06-05

??GT?? MVP worker?

### Added

- ??/?? `build_problem_gt_review_pack.py`???? GT????????????????GT????
- ?? Excel?HTML?CSV?manifest ????? payload?
- ?? `--backend` / `--project-id` ???????
- ?? `--backend-pdf-mode`?`--xlsx-pdf-mode` ?? PDF ???????

### Changed

- ???? XLSX XML reader ?? `final_gt_202506.xlsx`??? openpyxl ???? workbook ???
- ????????? PDF ?????????????????????

### Validation

- `python -m py_compile build_problem_gt_review_pack.py` ???
- ?????? `D:\data-annotation\problem_gt_mvp_202506_v16_fast3`?

## v1.9 - 2026-06-05

Problem GT review efficiency and safe write-back.

### Added

- Added `apply_problem_gt_proposal.py` to create a new final_gt workbook from exported problem-GT proposal JSON.
- The script appends approved add rows only, writes summary JSON/Markdown, and exports add/exclude/unresolved evidence CSV files.
- Duplicate fingerprints are skipped by default to reduce accidental double-counting.

### Safety

- Original `final_gt_202506.xlsx` is never overwritten.
- Excluded and unresolved proposal rows are recorded only; no delete/update is applied automatically.

## v2.0 - 2026-06-05

Problem GT proposal auto-download.

### Added

- `apply_problem_gt_proposal.py` can now download the proposal JSON directly from the backend with `--backend`, `--project-id`, and optional `--source-run-key`.
- The downloaded proposal is saved in the output directory before workbook generation.

Example:

```powershell
python apply_problem_gt_proposal.py `
  --backend "http://localhost:8080" `
  --project-id 2 `
  --source-run-key "problem_gt_mvp_202506_v16_fast3" `
  --gt-workbook "D:\data-annotation\final_gt_202506.xlsx" `
  --output-dir "D:\audit-engine\gt-review-assistant\workspace\problem_gt_apply_from_ui"
```

### Safety

- This keeps the same safe-write behavior: the original workbook is not overwritten.
- If `--proposal-json` is provided, the script uses the local proposal file as before.

## v2.1 - 2026-06-05

Problem GT source screenshot previews.

### Added

- Added `generate_problem_gt_previews.py` to generate read-only PNG screenshots for problem-GT candidates.
- Excel previews render the candidate cell area and highlight candidate cells when `excel_locator` contains addresses such as `Sheet6 D23; N23`.
- PDF previews render the located page when a page locator exists, otherwise the first page of the split-note PDF.
- Outputs are deterministic so the frontend can show them without a database change: `problem_gt/<run>/<candidateKey>_pdf.png` and `_excel.png`.

### Validation

- Generated previews for `problem_gt_mvp_202506_v16_fast3`: 315 candidates, 300 PDF previews, 137 Excel previews.

## v2.2 - 2026-06-05

Problem GT preview precision and backend task support.

### Added

- PDF previews now try to crop around matching row/column/value text instead of always rendering the whole page.
- When a PDF locator is empty, the worker searches the whole split-note PDF and crops the closest matching page when possible.
- Excel previews now fall back to fuzzy row/value lookup when a candidate has no explicit cell reference.
- Linux/Noto CJK font candidates were added so Docker-generated Excel screenshots can render Chinese text.

### Validation

- `python -m py_compile generate_problem_gt_previews.py` passed.
- Local preview check for `problem_gt_mvp_202506_v16_fast3` with `--limit 80`: 80 candidates, 29 Excel previews, 80 PDF previews.
- Full backend-generated preview run: 315 candidates, 137 Excel previews, 300 PDF previews, including 43 PDF crops found by whole-document search.

### Safety

- Preview generation is read-only. It writes PNG/CSV/JSON assets only and never modifies source PDF/XLSX files or final GT workbooks.
