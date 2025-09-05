# 📝 md_to_docx_table

A Python utility to convert **Markdown (`.md`) files into Word (`.docx`) documents** with special attention to **robust table handling**.

Unlike basic converters, this tool:
- Ensures pipe-style Markdown tables render as real Word tables
- Handles blank lines before/inside tables
- Creates clean, formatted Word tables with:
  - Grid borders
  - Bold + centered header row
  - Compact spacing

---

## 🚀 Features

- ✅ Converts a single Markdown file into Word
- ✅ Outputs to the same folder, or to a custom output directory
- ✅ Overwrite protection (`--force` flag required)
- ✅ Supports:
  - Tables (including ones with extra blank lines)
  - Headers (`#`, `##`, `###`)
  - Bullet and numbered lists
  - Code blocks (monospace font)

---

## 📂 Project Structure

```
md_to_docx_table/
├── md_to_docx_table.py   # Main script (CLI app)
├── README.md             # This file
└── requirements.txt      # Dependencies (optional)
```

---

## 🔧 Prerequisites

Python 3.8+ and `pip`.

Install dependencies:

```bash
pip install python-docx markdown beautifulsoup4
```

*(or use `pip install -r requirements.txt` if you add one)*

---

## ▶️ Usage

### Basic conversion

```bash
python md_to_docx_table.py notes.md
```

This creates `notes.docx` alongside the input.

### Custom output name

```bash
python md_to_docx_table.py notes.md -o report.docx
```

### Output into a folder

```bash
python md_to_docx_table.py notes.md --out-dir converted_docs
```

If `converted_docs/` doesn’t exist, it will be created.

### Overwrite protection

By default, if the `.docx` already exists the script will refuse to overwrite it.  
Use `--force` to override:

```bash
python md_to_docx_table.py notes.md --force
```

---

## 📌 Example

**Input (`example.md`):**

```markdown
# Demo Document

Some intro text.

| Name  | Age | Role    |
|-------|-----|---------|
| Alice |  30 | Engineer|
| Bob   |  25 | Analyst |
```

**Output (`example.docx`):**

- Heading “Demo Document”
- Paragraph “Some intro text.”
- A Word table with a bold, centered header row and two data rows.

---

## 📄 License

This project is provided under the MIT License.

---

## 👨‍💻 Author

**Erick Perales**  
IT Architect, Cloud Migration Specialist  
[https://github.com/peralese](https://github.com/peralese)
