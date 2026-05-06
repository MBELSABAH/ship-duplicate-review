# Ship Duplicate Review MVP v3

This Streamlit app includes Sprint 1 v3 features:

1. **Stable decisions**
   - stable keys for auto groups and manual pair decisions
2. **Session save/load**
   - download review progress as JSON and reload it later
3. **Safe workflow**
   - original workbook is never overwritten
4. **Exportable logs/mapping**
   - export auto/manual decisions, merge history, and canonical mapping

## Install
```bash
pip install -r requirements.txt
```

## Run
```bash
python3 -m streamlit run app.py
```

## Workflow
1. Upload workbook
2. Review / accept safe auto-merges
3. Review remaining manual candidates
4. Check merge history
5. Undo anything suspicious
6. Export:
   - auto-merge decisions
   - manual review decisions
   - merge history
   - canonical mapping

## Important
This MVP does not overwrite the workbook directly. It exports the merge plan and logs so the process stays auditable.
