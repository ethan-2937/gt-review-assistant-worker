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