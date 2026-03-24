---
name: "excel-to-markdown"
description: "Converts local Excel (.xlsx) files to Markdown tables. Invoke when a user asks to turn an Excel workbook/sheet into Markdown for Trae/LLM consumption."
---

# Excel to Markdown

## Goal

将本地 Excel（.xlsx）转换成适合 Trae/LLM 直接阅读与引用的 Markdown（pipe table），支持多 sheet、按空行自动分表、限制行列避免输出过大。

## When to Invoke

- 用户说“把这个 Excel 转成 Markdown / 贴给大模型用 / 给 agent 用”
- 用户给了一个 .xlsx 文件路径，希望你输出可直接粘贴的 Markdown 表格
- 用户希望指定某些 sheet、或只导出前 N 行/列

## How to Use (Repository Tooling)

本技能目录内置了可直接交付的脚本：`.trae/skills/excel-to-markdown/excel_to_markdown.py`。

### Basic

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx > out.md
```

### Select sheets

按名称：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx --sheet Sheet1 --sheet Sheet2
```

按序号（从 1 开始）：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx --sheet-index 1 --sheet-index 3
```

按正则匹配：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx --sheet-regex "日报|周报"
```

### Split multiple tables inside a sheet

当一个 sheet 中用空行分隔了多张表：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx --split-tables
```

可调整分隔判定（连续空行数）与最小表行数：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx --split-tables --blank-rows-gap 1 --min-table-rows 2
```

### Keep output small (recommended for chat)

限制行/列（裁剪空白后再限制）：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx --max-rows 60 --max-cols 12
```

### Headings control

多 sheet / 多表时默认会输出标题，便于引用；可关闭：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx --no-headings
```

## Output Expectations

- 默认导出“计算后的值”（Excel 里公式单元格会输出计算结果）
- 如果用户要看公式文本，使用 `--formulas`
- 单元格内换行会被转换为 `<br>`，并对 `|`、`\` 做转义以避免破坏表格

## Troubleshooting

- 报错 `unsupported_format: .xls`：先把 .xls 转成 .xlsx 再导出
- 报错 `sheet_too_large`：提高 `--max-cells` 或使用 `--sheet/--max-rows/--max-cols` 缩小范围
- 输出太长不适合直接粘贴：优先用 `--sheet`、`--split-tables`、`--max-rows/--max-cols`，或写到文件后只回传关键部分
