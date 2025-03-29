with open("/mnt/data/README.md", "w") as f:
    f.write("""# 📊 ExcelWriter)

**ExcelWriter** is a Python utility for easily creating, editing, and saving Excel `.xlsx` files using dictionaries and pandas DataFrames. It's ideal for automating data processing tasks with readable, testable Python code.

---

## 🚀 Features

- Create new Excel workbooks and sheets
- Write lists of dictionaries to Excel
- Write or append pandas DataFrames
- Edit specific Excel cells
- Save workbooks to custom file paths
- Automatically back up existing files
- Fully tested with `pytest` + `pytest-cov`

---

## 🛠 Requirements

- Python 3.10+
- `pandas`
- `openpyxl`
- `pytest`
- `pytest-cov`

---

## 🧪 Running Tests

Run all tests and show coverage using:

```bash
python run_tests.py
