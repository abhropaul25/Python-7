
# Renewable Tender → Excel Autofill (Starter Kit)

This starter kit learns your column schema directly from **8.1 Lakhwar_Master_v8_MLready_v2.xlsx** (detected master sheet: **Parameters_Master**),
then scans a folder of tender documents and writes extracted values into a new Excel.

## Files generated here
- `renewable_columns_master.json` — schema harvested from your workbook (sheet + columns).
- `tags_renewable.yaml` — starter regex rules for common renewable tags (solar/wind/hybrid/BESS). Edit/extend as needed.
- `renewable_excel_autofill.py` — the main Python script.
- this `README`

## Install (first time)
```bash
pip install pandas openpyxl python-docx PyPDF2 pyyaml
# optional (for stronger PDF coverage):
pip install pdfminer.six pdfplumber
```

## Prepare
- Put all tender files to parse in a folder, e.g. `./input_docs/` (PDF/DOCX/TXT/XLS/XLSX supported).
- If you have a **renewable**-specific template workbook, pass it with `--template`. If not given, the script creates a minimal workbook with the same master sheet and header as your source.

## Run
```bash
python renewable_excel_autofill.py --docs ./input_docs \\
  --out ./renewable_output.xlsx \\
  --columns ./renewable_columns_master.json \\
  --tags ./tags_renewable.yaml \\
  --template "/path/to/renewable_template.xlsx"
```

## How to teach the extractor
- Edit `tags_renewable.yaml`: add/adjust regex patterns to match your exact tender language. First match wins.
- Column mapping: the script maps tag keys to columns by normalised name; add aliases in `build_row()` if your column header differs.

## Extending beyond the Master sheet
- Duplicate the pattern for other sheets (e.g., "AI Tag Definitions", "Bidders Eligibility"). The current starter writes only to the detected master sheet for clarity.
- For cross-document aggregation (e.g., amendments vs. base tender), add a small state machine keyed by `tender_id` and `doc_type`.

## Safety notes
- This is a **starter**. Always spot-check the output before publishing or analytics.
- Keep a frozen copy of the source documents and the run logs for auditability.
