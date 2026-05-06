# Ship Duplicate Review MVP v3

This Streamlit app now supports **generic Excel workbooks** via column mapping while preserving Sprint 1 stable decisions/session behavior.

## Install
```bash
pip install -r requirements.txt
```

## Run
```bash
python3 -m streamlit run app.py
```

## Workflow
1. Upload workbook.
2. Choose sheet.
3. Choose a **primary deduplication column** (required).
4. Choose optional evidence columns (year/date, type/category, amount, unit, notes).
5. Review/accept safe auto-merges.
6. Review remaining manual candidates.
7. Save/load review session JSON.
8. Export logs and canonical mapping CSV files.

## Notes
- Original workbook is never overwritten.
- Column mapping defaults still match the original ship workbook when those columns exist.
- Session JSON includes column mapping config plus auto/manual decisions.
- Cleaned Excel export is **not** implemented in this sprint.
