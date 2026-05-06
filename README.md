# Ship Duplicate Review MVP v3

This project supports the duplicate-review workflow with a FastAPI backend, a React frontend shell, and a Streamlit fallback interface.

## Install backend dependencies
```bash
pip install -r requirements.txt
```

## Run backend
```bash
python3 -m uvicorn api:app --reload
```

## Frontend
```bash
cd frontend
npm install
npm run dev
```

## Streamlit fallback
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
- Cleaned Excel export is available through the app workflow and the FastAPI export endpoint.
