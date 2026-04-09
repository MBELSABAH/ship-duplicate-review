# Ship Duplicate Review MVP v2

This Streamlit app adds three important safeguards:

1. **Safe auto-merges**
   - conservative automatic merges for names that only differ by punctuation/case/spaces
2. **Merge history**
   - every currently active merge is listed with the reason it was merged
3. **Undo**
   - you can undo any active auto-merge or manual merge

## Install
```bash
pip3 install pandas openpyxl rapidfuzz streamlit
```

## Run
```bash
streamlit run mvp_ship_human_review_v2.py
```

If `streamlit` is not found:
```bash
python3 -m streamlit run mvp_ship_human_review_v2.py
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
