# Table_combiner  CLI


A lightweight Python CLI tool to merge multiple table files from a folder into a single Excel workbook.

---

## ✨ Features

- Merge all tables under a folder into one Excel file.
- Sheet naming rules:
  - If a file has only one sheet → sheet name = the part after the last `_` in the filename (or the whole stem if no `_`).
  - If a file has multiple sheets → keep the original sheet names.
- Auto-fix illegal Excel characters (`: \ / ? * [ ]`).
- Enforce Excel’s 31-char limit with safe truncation.
- Ensure unique sheet names by appending `_1`, `_2`, … if needed.
- Auto-detect encoding for CSV (`utf-8` with fallback to `gbk`).
- Supports `.xlsx`, `.xls`, `.csv`, `.tsv`, `.txt`.

---

## 📦 Installation

Requires **Python 3.9+**.

```bash
git clone https://github.com/yourname/excel-table-merger.git
cd excel-table-merger
pip install -r requirements.txt
````

Or install dependencies manually:

```bash
pip install pandas openpyxl
```

---

## 🚀 Usage

```bash
pip install pandas openpyxl
# 运行
python combiner.py -i /your/folder/path -o "ABC.xlsx"
# 静默模式（仅警告/错误）
python combiner.py -i /your/folder/path -o "ABC.xlsx" -q

```

### Options

* `-i, --input-folder` → Folder containing table files.
* `-o, --output-name`  → Output Excel filename (saved into the same folder).
* `-q, --quiet`        → Suppress info logs, only warnings/errors shown.

---

## ⚠️ Notes

* Files are processed in alphabetical order by filename.
* Empty tables are written with a placeholder `(empty)` column.
* Only merges as separate sheets, no vertical concatenation.

---

## 📜 License

MIT License © 2025 \[Mochiao Chen]

