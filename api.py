import io
import json
from datetime import date, datetime
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, Response

from dedupe_engine import (
    SHIP_DEFAULTS,
    active_manual_decisions_for_config,
    build_cleaned_workbook_bytes,
    build_merge_outputs,
    build_name_stats,
    build_safe_auto_groups,
    generate_candidate_pairs,
    preprocess_rows,
    resolved_names_from_auto,
)

app = FastAPI(title="Ship Duplicate Review API", version="0.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def parse_json_field(value: str | None, default: Any, field_name: str, required: bool = False):
    if value is None or value == "":
        if required:
            raise HTTPException(status_code=400, detail=f"Missing required JSON field: {field_name}")
        return default
    try:
        return json.loads(value)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid JSON for {field_name}: {exc.msg}") from exc


def json_safe_value(value: Any):
    if isinstance(value, dict):
        return {str(k): json_safe_value(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_safe_value(v) for v in value]
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    if pd.isna(value):
        return None
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def df_to_records(df: pd.DataFrame) -> list[dict[str, Any]]:
    return [{k: json_safe_value(v) for k, v in row.items()} for row in df.to_dict(orient="records")]


def recommend_column_config(available_cols: list[str], preview_df: pd.DataFrame) -> dict[str, Any]:
    text_candidates = [c for c in available_cols if preview_df[c].dtype == "object"]
    default_entity = SHIP_DEFAULTS["entity_column"] if SHIP_DEFAULTS["entity_column"] in available_cols else (text_candidates[0] if text_candidates else available_cols[0])
    config = {"entity_column": default_entity}
    for key in ["year_column", "type_column", "amount_column", "unit_column", "notes_column_1", "notes_column_2"]:
        config[key] = SHIP_DEFAULTS[key] if SHIP_DEFAULTS[key] in available_cols else None
    return config


def load_raw_df(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid sheet_name '{sheet_name}'.") from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to read workbook/sheet: {exc}") from exc


def validate_column_config(column_config: dict[str, Any], raw_df: pd.DataFrame):
    entity_column = column_config.get("entity_column")
    if not entity_column:
        raise HTTPException(status_code=400, detail="column_config.entity_column is required")
    if entity_column not in raw_df.columns.tolist():
        raise HTTPException(status_code=400, detail=f"entity_column '{entity_column}' not found in selected sheet")


@app.get("/")
def root():
    return {"status": "ok", "service": "ship-duplicate-review-api"}


@app.post("/workbook/sheets")
def workbook_sheets(file: UploadFile = File(...)):
    file_bytes = file.file.read()
    try:
        sheet_names = pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to read workbook: {exc}") from exc
    return {"sheet_names": sheet_names}


@app.post("/workbook/preview")
def workbook_preview(file: UploadFile = File(...), sheet_name: str = Form(...)):
    file_bytes = file.file.read()
    raw_df_preview = load_raw_df(file_bytes, sheet_name).head(50)
    available_cols = raw_df_preview.columns.tolist()
    if not available_cols:
        raise HTTPException(status_code=400, detail="Selected sheet has no columns")
    return {
        "columns": available_cols,
        "preview_rows": df_to_records(raw_df_preview.head(10)),
        "recommended_column_config": recommend_column_config(available_cols, raw_df_preview),
    }


@app.post("/dedupe/analyze")
def dedupe_analyze(
    file: UploadFile = File(...),
    sheet_name: str = Form(...),
    column_config: str = Form(...),
    fuzzy_threshold: int = Form(88),
    min_manual_score: float = Form(0.75),
    auto_status: str | None = Form(None),
    manual_decisions: str | None = Form(None),
):
    file_bytes = file.file.read()
    parsed_column_config = parse_json_field(column_config, {}, "column_config", required=True)
    parsed_auto_status = parse_json_field(auto_status, {}, "auto_status")
    parsed_manual_decisions = parse_json_field(manual_decisions, {}, "manual_decisions")

    raw_df = load_raw_df(file_bytes, sheet_name)
    validate_column_config(parsed_column_config, raw_df)

    rows_df = preprocess_rows(raw_df, parsed_column_config)
    stats_df = build_name_stats(rows_df)
    auto_groups_df = build_safe_auto_groups(stats_df, parsed_column_config["entity_column"])
    resolved_names = resolved_names_from_auto(auto_groups_df, parsed_auto_status)
    full_queue_df = generate_candidate_pairs(stats_df, entity_column=parsed_column_config["entity_column"], resolved_names=resolved_names, fuzzy_threshold=fuzzy_threshold)
    score_filtered_queue_df = full_queue_df[full_queue_df["score"] >= min_manual_score].reset_index(drop=True) if not full_queue_df.empty else full_queue_df.copy()
    active_manual_decisions = active_manual_decisions_for_config(parsed_manual_decisions, parsed_column_config)
    history_df, mapping_df = build_merge_outputs(stats_df, auto_groups_df, parsed_auto_status, active_manual_decisions)

    return JSONResponse({
        "summary": {
            "unique_raw_primary_values": len(stats_df),
            "safe_auto_groups": len(auto_groups_df),
            "accepted_auto_groups": sum(1 for v in parsed_auto_status.values() if v == "accepted"),
            "manual_queue": len(score_filtered_queue_df),
            "merged_names_now": len(mapping_df),
        },
        "auto_groups": df_to_records(auto_groups_df),
        "manual_queue": df_to_records(score_filtered_queue_df),
        "merge_history": df_to_records(history_df),
        "canonical_mapping": df_to_records(mapping_df),
    })


@app.post("/export/cleaned-workbook")
def export_cleaned_workbook(
    file: UploadFile = File(...),
    sheet_name: str = Form(...),
    column_config: str = Form(...),
    auto_status: str = Form("{}"),
    manual_decisions: str = Form("{}"),
    fuzzy_threshold: int = Form(88),
    min_manual_score: float = Form(0.75),
):
    file_bytes = file.file.read()
    parsed_column_config = parse_json_field(column_config, {}, "column_config", required=True)
    parsed_auto_status = parse_json_field(auto_status, {}, "auto_status")
    parsed_manual_decisions = parse_json_field(manual_decisions, {}, "manual_decisions")

    raw_df = load_raw_df(file_bytes, sheet_name)
    validate_column_config(parsed_column_config, raw_df)

    rows_df = preprocess_rows(raw_df, parsed_column_config)
    stats_df = build_name_stats(rows_df)
    auto_groups_df = build_safe_auto_groups(stats_df, parsed_column_config["entity_column"])
    resolved_names = resolved_names_from_auto(auto_groups_df, parsed_auto_status)
    full_queue_df = generate_candidate_pairs(stats_df, entity_column=parsed_column_config["entity_column"], resolved_names=resolved_names, fuzzy_threshold=fuzzy_threshold)
    score_filtered_queue_df = full_queue_df[full_queue_df["score"] >= min_manual_score].reset_index(drop=True) if not full_queue_df.empty else full_queue_df.copy()
    active_manual_decisions = active_manual_decisions_for_config(parsed_manual_decisions, parsed_column_config)
    history_df, mapping_df = build_merge_outputs(stats_df, auto_groups_df, parsed_auto_status, active_manual_decisions)

    workbook_out = build_cleaned_workbook_bytes(raw_df=raw_df, mapping_df=mapping_df, history_df=history_df, manual_decisions=active_manual_decisions, auto_groups_df=auto_groups_df, column_config=parsed_column_config, sheet_name=sheet_name, workbook_bytes=file_bytes, candidate_queue_df=score_filtered_queue_df)
    return Response(content=workbook_out, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="cleaned_{sheet_name}.xlsx"'})
