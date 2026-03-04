# Sales & Inventory Lot Tool

Internal tool for processing monthly sales/inventory Excel reports by exploding lot numbers, allocating sales/costs/fees,
and generating derived tabs (real sales, samples, credit memos, summary).

### Setup

- **Create a virtual environment** (recommended):

```bash
cd sales-lot-tool
python -m venv .venv
# Windows PowerShell
.venv\Scripts\Activate.ps1
```

- **Install dependencies**:

```bash
pip install -r requirements.txt
```

### Streamlit web app (upload → download)

Run:

```bash
cd sales-lot-tool
streamlit run app.py
```

Then open the URL Streamlit prints (usually `http://localhost:8501`), upload your monthly Excel file
(`.xlsx`), and download the processed workbook with the additional sheets.

### Command-line usage (optional)

You can also run the processor directly on a file:

```bash
cd sales-lot-tool
python -m sales_tool.processor path\to\your_report.xlsx
```

This updates the file in-place, adding/replacing the derived sheets.

