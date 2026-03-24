# excel-to-markdown (Trae Skill)

把本地 Excel（.xlsx）转换为适合 Trae / LLM / agent 直接消费的 Markdown 表格（pipe table）。支持多 sheet、按空行拆分多张表、限制导出规模以便在对话中传递。

## 目录结构

```text
.trae/skills/excel-to-markdown/
  SKILL.md
  excel_to_markdown.py
  README.md
```

## 功能特性

- 导出 `.xlsx` 为 Markdown pipe table
- 选择 sheet：按名称 / 序号 / 正则
- 自动分表：sheet 内按空行分隔多张表（可调阈值）
- 输出可控：限制行/列/单元格数，避免一次导出过大
- 单元格转义：处理 `|`、`\`、换行，避免破坏表格结构

## 环境要求

- Python 3
- 依赖：`openpyxl`

安装依赖：

```bash
python3 -m pip install openpyxl
```

## 快速开始

把 Excel 转成 Markdown（输出到 stdout）：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx > out.md
```

输出到文件：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx -o out.md
```

## 常用用法

### 选择特定 sheet

按名称（可重复）：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --sheet Sheet1 --sheet Sheet2
```

按序号（从 1 开始，可重复）：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --sheet-index 1 --sheet-index 3
```

按正则（匹配 sheet 名称）：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --sheet-regex "日报|周报"
```

### 一个 sheet 内拆分多张表

当一个 sheet 中用空行分隔了多块表格：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --split-tables
```

调整分表策略：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --split-tables --blank-rows-gap 1 --min-table-rows 2
```

### 控制输出规模（推荐用于对话/agent）

限制导出行/列（先裁剪四周空白，再截取）：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --max-rows 60 --max-cols 12
```

限制单表最大单元格数量（安全阈值，避免巨大表卡住）：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --max-cells 20000
```

### 标题控制

多 sheet / 多表时默认会输出标题（便于引用/定位）；可关闭：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --no-headings
```

指定标题层级：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py /path/to/file.xlsx \
  --heading-level 2
```

## 参数一览

运行 `-h` 查看完整帮助：

```bash
python3 .trae/skills/excel-to-markdown/excel_to_markdown.py -h
```

常用参数：

- `-o/--output`：输出到文件（默认 stdout）
- `--sheet`：按名称选择 sheet（可重复）
- `--sheet-index`：按序号选择 sheet（从 1 开始，可重复）
- `--sheet-regex`：按正则选择 sheet
- `--split-tables`：按空行拆分多张表
- `--blank-rows-gap`：连续空行达到多少行开始分表
- `--min-table-rows`：保留表格的最小行数
- `--max-rows/--max-cols`：裁剪后的最大行/列
- `--max-cells`：单表最大单元格数（超出直接报错）
- `--no-headings`：不输出标题
- `--heading-level`：标题的基础层级
- `--formulas`：输出公式文本（默认输出计算后的值）
- `--no-trim`：不裁剪四周空白行/列
- `--header-row`：指定表头行（used range 内，从 1 开始）
- `--no-header`：生成 Col1.. 的表头（不使用 Excel 表头行）

## 输出规则（重要）

- 默认 `data_only=True`：公式单元格输出计算结果；若希望导出公式本身，使用 `--formulas`
- 单元格内换行会转为 `<br>`，并对 `|`、`\` 做转义，以保证 Markdown 表格不被破坏
- 仅支持 `.xlsx`；`.xls` 请先转换为 `.xlsx`

## 在 Trae / agent 中使用

- 技能入口文件：[SKILL.md](./SKILL.md)
- 通常做法：让 agent 运行上面的命令，将输出的 Markdown 直接粘贴回对话或写入 `out.md` 后再摘取关键片段
- 推荐加上 `--max-rows/--max-cols` 来控制一次性输出长度

## 常见问题

- 报错 `Missing dependency: openpyxl`：执行 `python3 -m pip install openpyxl`
- 报错 `sheet_too_large`：提高 `--max-cells` 或用 `--sheet/--max-rows/--max-cols` 缩小范围
