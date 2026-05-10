# Ship Duplicate Review

A Streamlit-based duplicate-review tool for messy historical spreadsheet data. The app helps detect, review, merge, undo, standardize, and export duplicate or near-duplicate entity names while keeping the workflow auditable and reversible.

The project was built for historical ship-ledger data, where the same vessel may appear under inconsistent spelling, punctuation, capitalization, or transcription variants. The current version also supports generic Excel workbooks through column mapping, so the review workflow can be applied to other text/entity columns too.

## Live demo

The public Streamlit app is linked in the repository **About** section.

> Do not upload private, restricted, or sensitive research data to a public deployment unless you have permission to do so.

## What it does

- Upload an Excel workbook (`.xlsx` or `.xls`)
- Select the worksheet to review
- Choose the primary column to deduplicate
- Map optional evidence columns such as year/date, type/category, amount, unit, and notes
- Generate conservative safe auto-merge groups
- Generate fuzzy manual-review candidates
- Review candidates as `merge`, `keep separate`, or `unsure`
- Hide or show reviewed candidates
- Undo saved auto/manual decisions, including direct row-selection undo in Merge History
- Save and reload review sessions as JSON
- Export a reviewed/auditable workbook that keeps original values and appends dedupe columns
- Export a standardized workbook that replaces the selected entity column with reviewed canonical values
- Preserve original Excel workbook formatting as much as possible during export
- Export audit files such as auto decisions, manual decisions, merge history, and canonical mapping CSVs

## Why this project matters

Historical datasets often contain inconsistent names and uncertain transcriptions. Blind automatic merging is risky because incorrect merges can damage the research dataset.

This tool uses a human-in-the-loop workflow:

1. automate only the safest obvious cases
2. surface ambiguous cases for review
3. show evidence before decisions are made
4. keep decisions traceable
5. allow undo and exportable audit records
6. produce both auditable and standardized outputs

## Export options

### Reviewed workbook

The reviewed workbook keeps the selected entity column unchanged and appends dedupe/audit columns such as:

- `dedupe_original_value`
- `dedupe_canonical_value`
- `dedupe_cluster_id`
- `dedupe_review_status`
- `dedupe_decision_source`
- `dedupe_score`
- `dedupe_reason`

Use this when you need a traceable research output.

### Standardized workbook

The standardized workbook replaces the selected entity column with reviewed canonical values and does not add audit columns.

Use this when you need a clean downstream-analysis file for filtering, pivot tables, or further processing.

## Tech stack

- Python
- Streamlit
- pandas
- openpyxl
- rapidfuzz
- FastAPI
- Uvicorn

## Repository structure

```text
ship-duplicate-review/
├── app.py              # Streamlit review interface
├── api.py              # Optional FastAPI backend
├── dedupe_engine.py    # Matching, scoring, export, and workbook logic
├── requirements.txt    # Python dependencies
└── README.md
```

## Install

Clone the repository:

```bash
git clone https://github.com/MBELSABAH/ship-duplicate-review.git
cd ship-duplicate-review
```

Create and activate a virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate
```

Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

## Run the Streamlit app locally

```bash
python3 -m streamlit run app.py
```

Then open the local Streamlit URL shown in the terminal.

## Run the optional FastAPI backend

```bash
python3 -m uvicorn api:app --reload
```

API docs will be available locally at:

```text
http://127.0.0.1:8000/docs
```

## Basic workflow

1. Upload a workbook.
2. Choose the sheet to review.
3. Select the primary column to deduplicate.
4. Map optional evidence columns.
5. Review safe auto-merge groups.
6. Review fuzzy manual candidates.
7. Use Merge History to inspect or undo decisions.
8. Save a review-session JSON if you want to continue later.
9. Export either a reviewed workbook, a standardized workbook, or audit CSV files.

## Notes on data safety

- The uploaded workbook is never overwritten.
- Exported workbooks are generated as separate files.
- Review decisions can be saved separately as JSON.
- Real research datasets should not be committed to the repository unless permission is explicitly granted.
- Public deployments should be used only with data that is safe to upload.

## Current status

This is an MVP research tool. The core workflow is functional and currently supports generic Excel column mapping, safe auto-merges, manual review, undo, session save/load, formatting-preserving workbook export, and Streamlit Cloud deployment.

Progress report: [MVP Progress Report](docs/mvp_progress_report.md)

Strong next improvements include:

- group-level review instead of only pair-level review
- persistent saved review projects
- reusable canonical dictionaries across future workbooks
- screenshots and demo data
- richer tests around export behavior and Streamlit UI interactions

## License

No formal license has been specified yet. Please contact the maintainer before reusing or redistributing this code.

## Suggested GitHub topics

```text
streamlit python data-cleaning fuzzy-matching historical-data digital-humanities entity-resolution excel fastapi
```
