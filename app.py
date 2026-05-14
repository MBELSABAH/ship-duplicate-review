import io
import json
import hashlib
import html
from copy import deepcopy
from uuid import uuid4

import pandas as pd
import streamlit as st

from dedupe_engine import (
    active_manual_decisions_for_config,
    build_cleaned_workbook_bytes,
    build_manual_decision_record,
    build_merge_outputs,
    build_name_stats,
    build_safe_auto_groups,
    build_standardized_workbook_bytes,
    generate_candidate_pairs,
    make_download_bytes,
    now_iso,
    preprocess_rows,
    resolved_names_from_auto,
    normalize_column_config,
    set_to_text,
    SHIP_DEFAULTS,
)

APP_TITLE = "Duplicate Review MVP v3"
SESSION_SCHEMA_VERSION = 1
VALID_AUTO_STATUSES = {"accepted", "rejected"}
VALID_MANUAL_DECISIONS = {"merge", "keep_separate", "unsure"}
REQUIRED_MANUAL_DECISION_FIELDS = {"pair_key", "entity_column", "name_a", "name_b", "decision"}
AUTO_GROUP_COLUMNS = [
    "auto_group_id",
    "auto_group_key",
    "strict_name_key",
    "canonical_name",
    "member_count",
    "member_names",
    "members_list",
    "reasons",
    "confidence",
    "total_rows",
    "min_year",
    "max_year",
    "units",
    "vessel_types",
]

def init_state():
    defaults = {
        "auto_status": {},
        "manual_decisions": {},
        "auto_index": 0,
        "auto_open_group_key": None,
        "pair_index": 0,
        "column_config": {},
        "evidence_fields": [],
        "evidence_fields_fingerprint": None,
        "queue_settings_fingerprint": None,
        "pending_loaded_session": None,
        "load_session_feedback": None,
        "merge_history_undo_feedback": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if "pair_decisions" in st.session_state and st.session_state["pair_decisions"] and not st.session_state["manual_decisions"]:
        st.session_state["manual_decisions"] = {}

def reset_decision_state():
    st.session_state["auto_status"] = {}
    st.session_state["manual_decisions"] = {}
    st.session_state["auto_index"] = 0
    st.session_state["auto_open_group_key"] = None
    st.session_state["pair_index"] = 0
    st.session_state["queue_settings_fingerprint"] = None
    if "pair_decisions" in st.session_state:
        st.session_state["pair_decisions"] = {}

def has_saved_manual_decision(pair_key: str, entity_column: str) -> bool:
    record = st.session_state.get("manual_decisions", {}).get(pair_key)
    if not isinstance(record, dict):
        return False
    return record.get("entity_column") == entity_column and record.get("decision") in VALID_MANUAL_DECISIONS

def undo_manual_decision_for_pair(pair_key: str):
    st.session_state.get("manual_decisions", {}).pop(pair_key, None)
    st.session_state.pop(f"reviewer_comment_{pair_key}", None)

def mapping_fingerprint(sheet_name: str, available_cols: list[str], entity_column: str) -> str:
    payload = {
        "sheet": sheet_name,
        "columns": [str(c) for c in available_cols],
        "entity": entity_column,
    }
    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()

def config_fingerprint(column_config: dict) -> str:
    return hashlib.sha256(json.dumps(column_config, sort_keys=True).encode("utf-8")).hexdigest()

def new_factor_id() -> str:
    return f"E{uuid4().hex[:12]}"

def normalize_factor_rows_for_ui(evidence_fields: list[dict], available_cols: list[str], entity_column: str) -> list[dict]:
    out = []
    seen_ids = set()
    seen_columns = set()
    for raw_field in evidence_fields or []:
        if not isinstance(raw_field, dict):
            continue
        column = str(raw_field.get("column") or "").strip()
        if not column or column == entity_column or column not in available_cols:
            continue
        normalized_column_key = column.lower()
        if normalized_column_key in seen_columns:
            continue
        factor_id = str(raw_field.get("id") or "").strip()
        if not factor_id or factor_id in seen_ids:
            factor_id = new_factor_id()
        seen_ids.add(factor_id)
        seen_columns.add(normalized_column_key)
        out.append({
            "id": factor_id,
            "column": column,
            "kind": "auto",
            "weight": float(raw_field.get("weight", 0.08)),
            "enabled": True,
        })
    return out

def default_factor_fields(available_cols: list[str], entity_column: str) -> list[dict]:
    selectable_columns = [c for c in available_cols if c != entity_column]
    preferred_order = [
        "Type of Veseel",
        "Year",
        "Amount (primary)",
        "Unit (primary)",
    ]
    noisy_tokens = {"day", "month", "volume", "oid", "page"}
    picked = []
    for col in preferred_order:
        if col in selectable_columns:
            picked.append(col)
            break
    if not picked:
        for col in selectable_columns:
            lowered = str(col).lower().replace("_", " ").replace("-", " ")
            if not any(token in lowered.split() for token in noisy_tokens):
                picked.append(col)
                break
    out = []
    for col in picked:
        out.append({
            "id": new_factor_id(),
            "column": col,
            "kind": "auto",
            "weight": 0.08,
            "enabled": True,
        })
    return out

def ensure_auto_group_schema(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame(columns=AUTO_GROUP_COLUMNS)
    out = df.copy()
    for col in AUTO_GROUP_COLUMNS:
        if col not in out.columns:
            out[col] = pd.Series(dtype=object)
    return out[AUTO_GROUP_COLUMNS]

@st.cache_data(show_spinner=False)
def load_sheet_names(file_bytes: bytes):
    bio = io.BytesIO(file_bytes)
    return pd.ExcelFile(bio).sheet_names

@st.cache_data(show_spinner=False)
def build_base_data(file_bytes: bytes, sheet_name: str, column_config: dict):
    bio = io.BytesIO(file_bytes)
    raw_df = pd.read_excel(bio, sheet_name=sheet_name)
    rows_df = preprocess_rows(raw_df, column_config)
    stats_df = build_name_stats(rows_df)
    auto_groups_df = build_safe_auto_groups(stats_df, column_config["entity_column"])
    return raw_df, rows_df, stats_df, auto_groups_df

@st.cache_data(show_spinner=False)
def sheet_fingerprint(file_bytes: bytes, sheet_name: str) -> dict:
    raw_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
    row_count = int(len(raw_df))
    columns = [str(col) for col in raw_df.columns.tolist()]
    sample_rows_per_side = 25
    if row_count > sample_rows_per_side * 2:
        sample_df = pd.concat([raw_df.head(sample_rows_per_side), raw_df.tail(sample_rows_per_side)], ignore_index=True)
    else:
        sample_df = raw_df.copy()
    sample_df = sample_df.reindex(columns=raw_df.columns)
    normalized_sample = sample_df.where(sample_df.notna(), "__NA__").astype(str)
    sample_row_hashes = [int(x) for x in pd.util.hash_pandas_object(normalized_sample, index=False).tolist()]
    sample_hash = hashlib.sha256(",".join(str(x) for x in sample_row_hashes).encode("utf-8")).hexdigest()
    fingerprint_input = {
        "sheet_name": sheet_name,
        "row_count": row_count,
        "columns": columns,
        "sample_size": int(len(sample_df)),
        "sample_hash": sample_hash,
    }
    fingerprint = hashlib.sha256(
        json.dumps(fingerprint_input, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
    ).hexdigest()
    return {
        "algorithm": "sheet-v1-sha256",
        "sheet_name": sheet_name,
        "row_count": row_count,
        "columns": columns,
        "sample_size": int(len(sample_df)),
        "sample_hash": sample_hash,
        "fingerprint": fingerprint,
    }

def validate_session_payload(data: object) -> tuple[bool, str, dict]:
    if not isinstance(data, dict):
        return False, "Invalid session format. The JSON root must be an object.", {}
    column_config = data.get("column_config")
    auto_status = data.get("auto_status")
    manual_decisions = data.get("manual_decisions")
    if not isinstance(column_config, dict):
        return False, "Invalid session format. `column_config` must be an object.", {}
    if not isinstance(auto_status, dict):
        return False, "Invalid session format. `auto_status` must be an object.", {}
    if not isinstance(manual_decisions, dict):
        return False, "Invalid session format. `manual_decisions` must be an object.", {}
    for group_key, status in auto_status.items():
        if status not in VALID_AUTO_STATUSES:
            return False, f"Invalid session format. auto_status entry `{group_key}` has invalid value `{status}`.", {}
    for pair_key, record in manual_decisions.items():
        if not isinstance(record, dict):
            return False, f"Invalid session format. manual_decisions entry `{pair_key}` must be an object.", {}
        missing_fields = [field for field in REQUIRED_MANUAL_DECISION_FIELDS if field not in record]
        if missing_fields:
            return False, (
                f"Invalid session format. manual_decisions entry `{pair_key}` is missing required fields: "
                + ", ".join(sorted(missing_fields))
                + "."
            ), {}
        decision = record.get("decision")
        if decision not in VALID_MANUAL_DECISIONS:
            return False, f"Invalid session format. manual_decisions entry `{pair_key}` has invalid decision `{decision}`.", {}
        reviewer_comment = record.get("reviewer_comment", "")
        if reviewer_comment is not None and not isinstance(reviewer_comment, str):
            return False, f"Invalid session format. manual_decisions entry `{pair_key}` has invalid reviewer_comment type.", {}
    return True, "", data

def build_session_payload(
    source_filename: str,
    workbook_bytes: bytes,
    sheet_name: str,
    raw_df: pd.DataFrame,
    column_config: dict,
):
    fingerprint = sheet_fingerprint(workbook_bytes, sheet_name)
    return {
        "session_schema_version": SESSION_SCHEMA_VERSION,
        "app_version": "v3-stable-decisions-session",
        "saved_at": now_iso(),
        "source_filename": source_filename,
        "sheet_name": sheet_name,
        "sheet_row_count": int(len(raw_df)),
        "sheet_columns": [str(col) for col in raw_df.columns.tolist()],
        "sheet_fingerprint": fingerprint,
        "column_config": column_config,
        "auto_status": st.session_state.get("auto_status", {}),
        "manual_decisions": st.session_state.get("manual_decisions", {}),
    }

def show_summary_card(title: str, data: dict):
    st.markdown(f"### {title}")
    st.markdown(f"**Primary value:** {data.get('raw_name', '')}")
    st.markdown(f"**Cleaned value:** `{data.get('clean_name', '')}`")
    st.markdown(f"**Rows:** {data.get('row_count', 0)}")
    st.markdown(f"**Years:** {data.get('min_year', '')} → {data.get('max_year', '')}")
    st.markdown(f"**Type/category evidence:** {set_to_text(data.get('vessel_types', [])) or '—'}")
    st.markdown(f"**Units:** {set_to_text(data.get('units', [])) or '—'}")
    median_tons = data.get("median_tons")
    st.markdown(f"**Median cargo amount when unit=tons:** {median_tons if median_tons is not None else '—'}")
    if data.get("sample_notes"):
        st.markdown(f"**Sample notes:** {data['sample_notes']}")

def render_prominent_name(name: object, side_label: str, align: str = "left"):
    safe_side_label = html.escape(str(side_label))
    safe_name = html.escape(str(name or ""))
    st.markdown(
        (
            f"<div style='text-align:{align}; line-height:1.15;'>"
            f"<div style='font-size:0.85rem; font-weight:600; opacity:0.75;'>{safe_side_label}</div>"
            f"<div style='font-size:2.0rem; font-weight:700; margin-top:0.1rem;'>{safe_name}</div>"
            "</div>"
        ),
        unsafe_allow_html=True,
    )

def auto_group_names_for_display(group_row: pd.Series) -> list[str]:
    members_list = group_row.get("members_list", [])
    if isinstance(members_list, list):
        names = [str(name).strip() for name in members_list if str(name).strip()]
        if names:
            return names
    member_names = str(group_row.get("member_names", "") or "")
    return [part.strip() for part in member_names.split("|") if part.strip()]

def format_summary_value(value: object) -> str:
    if value is None:
        return "—"
    if isinstance(value, float):
        if pd.isna(value):
            return "—"
        if value.is_integer():
            return str(int(value))
        return f"{value:.2f}".rstrip("0").rstrip(".")
    if isinstance(value, int):
        return str(value)
    return str(value)

def summarize_group_factor(series: pd.Series, column_name: str) -> str:
    cleaned = series.map(lambda value: "" if pd.isna(value) else str(value).strip())
    total_count = len(cleaned)
    nonempty = cleaned[cleaned != ""]
    missing_count = total_count - len(nonempty)
    prefix = f"{column_name}: "

    if total_count == 0 or nonempty.empty:
        return f"{prefix}unavailable"

    lower_column_name = str(column_name).lower()
    if "year" in lower_column_name:
        year_values = pd.to_numeric(nonempty, errors="coerce").dropna()
        if not year_values.empty:
            min_year = int(year_values.min())
            max_year = int(year_values.max())
            year_text = str(min_year) if min_year == max_year else f"{min_year}–{max_year}"
        else:
            unique_values = sorted(set(nonempty.tolist()))
            if len(unique_values) == 1:
                year_text = unique_values[0]
            else:
                year_text = f"{unique_values[0]}–{unique_values[-1]}"
        if missing_count > 0:
            return f"{prefix}missing for some rows"
        return f"{prefix}{year_text}"

    numeric_values = pd.to_numeric(nonempty, errors="coerce")
    numeric_nonempty = numeric_values.dropna()
    if len(numeric_nonempty) == len(nonempty):
        min_value = numeric_nonempty.min()
        max_value = numeric_nonempty.max()
        if min_value == max_value:
            summary = f"all same ({format_summary_value(min_value)})"
        else:
            summary = f"range {format_summary_value(min_value)}–{format_summary_value(max_value)}"
    else:
        unique_values = sorted(set(nonempty.tolist()))
        if len(unique_values) == 1:
            summary = f"all same ({unique_values[0]})"
        elif len(unique_values) <= 3:
            summary = "values " + ", ".join(unique_values)
        else:
            summary = f"{len(unique_values)} distinct values"

    if missing_count > 0:
        summary += "; missing for some rows"
    return f"{prefix}{summary}"

def build_safe_group_evidence_summary(
    group_rows: pd.DataFrame,
    selected_factor_columns: list[str],
) -> str:
    parts = ["Names match after removing punctuation/case/spaces"]
    for factor_column in selected_factor_columns:
        if factor_column not in group_rows.columns:
            parts.append(f"{factor_column}: unavailable")
            continue
        parts.append(summarize_group_factor(group_rows[factor_column], factor_column))
    return "; ".join(parts)

def evidence_summary_from_scores(evidence_scores: object) -> str:
    if not isinstance(evidence_scores, list):
        return ""
    parts = []
    for item in evidence_scores:
        if not isinstance(item, dict):
            continue
        column = str(item.get("column") or "evidence").strip()
        reason = str(item.get("reason") or "").strip()
        if reason:
            if reason.lower().startswith(f"{column.lower()}:"):
                parts.append(reason)
            else:
                parts.append(f"{column}: {reason}")
            continue
        score_value = item.get("score")
        if isinstance(score_value, (int, float)):
            parts.append(f"{column}: score {score_value:.2f}")
        else:
            parts.append(column)
    return " | ".join(parts)

def app():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_state()

    st.title(APP_TITLE)
    st.caption("Duplicate-review workflow with suggested safe groups, manual review, history, and export.")

    with st.sidebar:
        st.subheader("Workbook")
        uploaded = st.file_uploader("Upload workbook", type=["xlsx", "xls"], key="upload_workbook")
        st.subheader("Matching")
        fuzzy_threshold = st.slider(
            "Name similarity cutoff",
            60,
            98,
            88,
            1,
            key="name_match_strictness_slider",
            help=(
                "Controls which name pairs enter the review queue. Higher = fewer candidates with more similar names. "
                "Lower = more candidates, more noise. Strong selected factors can still surface borderline names."
            ),
        )
        min_manual_score = st.slider(
            "Minimum final score",
            0.0,
            1.0,
            0.75,
            0.01,
            key="overall_evidence_threshold_slider",
            help=(
                "Filters generated candidates after name similarity and selected dedupe factors are scored together. "
                "Higher = only stronger overall matches are shown."
            ),
        )
        st.subheader("Review")
        hide_reviewed_candidates = st.checkbox(
            "Hide reviewed candidates",
            value=True,
            key="hide_reviewed_candidates_checkbox",
            help="When ON, reviewed items are hidden from the current queue; when OFF, reviewed items stay visible.",
        )
        sample_rows = st.slider("Evidence rows per side", 3, 12, 5, 1, key="evidence_rows_per_side_slider")
        st.subheader("Danger / reset")
        if st.button("Reset all decisions", width="stretch", key="reset_all_decisions_button"):
            reset_decision_state()
            st.rerun()

    if not uploaded:
        st.info("Upload the workbook to start.")
        st.stop()

    file_bytes = uploaded.getvalue()
    sheet_names = load_sheet_names(file_bytes)
    pending_loaded_session = st.session_state.pop("pending_loaded_session", None)
    if pending_loaded_session:
        pending_sheet = pending_loaded_session.get("sheet_name")
        if pending_sheet:
            if pending_sheet in sheet_names:
                st.session_state["sheet_select"] = pending_sheet
            else:
                st.session_state["load_session_feedback"] = (
                    "error",
                    f"This session was saved for sheet `{pending_sheet}`, but that sheet is not in the uploaded workbook. Session was not loaded.",
                )
                pending_loaded_session = None

    with st.sidebar:
        sheet_name = st.selectbox("Sheet", sheet_names, index=0, key="sheet_select")

    raw_df_preview = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, nrows=50)
    available_cols = raw_df_preview.columns.tolist()
    text_candidates = [c for c in available_cols if raw_df_preview[c].dtype == "object"]
    default_entity = SHIP_DEFAULTS["entity_column"] if SHIP_DEFAULTS["entity_column"] in available_cols else (text_candidates[0] if text_candidates else available_cols[0])
    if pending_loaded_session:
        saved_cfg = pending_loaded_session.get("column_config", {}) if isinstance(pending_loaded_session.get("column_config", {}), dict) else {}
        saved_cfg = normalize_column_config(saved_cfg, available_columns=available_cols)
        saved_entity_column = saved_cfg.get("entity_column")
        if not saved_entity_column or saved_entity_column not in available_cols:
            st.session_state["load_session_feedback"] = (
                "error",
                f"The saved primary/entity column `{saved_entity_column}` is missing from the uploaded sheet. Session was not loaded.",
            )
            pending_loaded_session = None
        else:
            saved_sheet_fingerprint = pending_loaded_session.get("sheet_fingerprint")
            has_saved_fingerprint = isinstance(saved_sheet_fingerprint, dict) and bool(saved_sheet_fingerprint.get("fingerprint"))
            fingerprint_is_valid = True
            older_session_warning = False
            if has_saved_fingerprint:
                current_sheet_fingerprint = sheet_fingerprint(file_bytes, sheet_name)
                fingerprint_is_valid = (
                    saved_sheet_fingerprint.get("fingerprint") == current_sheet_fingerprint.get("fingerprint")
                )
                saved_row_count = pending_loaded_session.get("sheet_row_count")
                if isinstance(saved_row_count, int):
                    fingerprint_is_valid = fingerprint_is_valid and saved_row_count == current_sheet_fingerprint.get("row_count")
                saved_columns = pending_loaded_session.get("sheet_columns")
                if isinstance(saved_columns, list):
                    normalized_saved_columns = [str(col) for col in saved_columns]
                    fingerprint_is_valid = fingerprint_is_valid and normalized_saved_columns == current_sheet_fingerprint.get("columns", [])
            else:
                older_session_warning = True

            if not fingerprint_is_valid:
                st.session_state["load_session_feedback"] = (
                    "error",
                    "This session does not appear to match the uploaded workbook/sheet. Session was not loaded.",
                )
                pending_loaded_session = None
            else:
                missing_evidence_cols = [f["column"] for f in saved_cfg.get("evidence_fields", []) if f.get("column") not in available_cols]
                saved_cfg["evidence_fields"] = [f for f in saved_cfg.get("evidence_fields", []) if f.get("column") in available_cols]
                st.session_state["entity_column_select"] = saved_entity_column
                st.session_state["evidence_fields"] = normalize_factor_rows_for_ui(
                    deepcopy(saved_cfg.get("evidence_fields", [])),
                    available_cols=available_cols,
                    entity_column=saved_entity_column,
                )
                st.session_state["evidence_fields_fingerprint"] = mapping_fingerprint(sheet_name, available_cols, saved_entity_column)
                st.session_state["column_config"] = saved_cfg
                st.session_state["auto_status"] = pending_loaded_session.get("auto_status", {})
                loaded_manual_decisions = pending_loaded_session.get("manual_decisions", {})
                for record in loaded_manual_decisions.values():
                    if isinstance(record, dict):
                        record.setdefault("reviewer_comment", "")
                st.session_state["manual_decisions"] = loaded_manual_decisions
                st.session_state["auto_index"] = 0
                st.session_state["pair_index"] = 0
                st.session_state["queue_settings_fingerprint"] = None
                feedback_messages = []
                if older_session_warning:
                    feedback_messages.append("This looks like an older session file without workbook fingerprint metadata. Loaded with caution.")
                if missing_evidence_cols:
                    feedback_messages.append(
                        "Loaded session. Some saved evidence columns are missing in the current sheet: "
                        + ", ".join(sorted(set(missing_evidence_cols)))
                        + ". Current mapping fallback was used where needed."
                    )
                if feedback_messages:
                    st.session_state["load_session_feedback"] = ("warning", " ".join(feedback_messages))
                else:
                    st.session_state["load_session_feedback"] = ("success", "Review session loaded.")

    saved_config = normalize_column_config(st.session_state.get("column_config", {}) or {}, available_columns=available_cols)
    with st.sidebar:
        st.subheader("Column mapping")
        entity_default = saved_config.get("entity_column") if saved_config.get("entity_column") in available_cols else default_entity
        entity_column = st.selectbox(
            "Select column to deduplicate",
            available_cols,
            index=available_cols.index(entity_default),
            key="entity_column_select",
        )

        current_mapping_fingerprint = mapping_fingerprint(sheet_name, available_cols, entity_column)
        if st.session_state.get("evidence_fields_fingerprint") != current_mapping_fingerprint:
            seeded = [f for f in saved_config.get("evidence_fields", []) if f.get("column") in available_cols and f.get("column") != entity_column]
            if not seeded:
                seeded = default_factor_fields(available_cols, entity_column)
            st.session_state["evidence_fields"] = normalize_factor_rows_for_ui(
                deepcopy(seeded),
                available_cols=available_cols,
                entity_column=entity_column,
            )
            st.session_state["evidence_fields_fingerprint"] = current_mapping_fingerprint

        st.markdown("**Dedupe based on:**")
        editable_fields = normalize_factor_rows_for_ui(
            deepcopy(st.session_state.get("evidence_fields", [])),
            available_cols=available_cols,
            entity_column=entity_column,
        )
        selectable_columns = [c for c in available_cols if c != entity_column]
        remove_idx = None
        for idx, field in enumerate(editable_fields):
            row_key = str(field.get("id", f"row-{idx}"))
            row_cols = st.columns([2.0, 0.5])
            column_value = field.get("column")
            if column_value not in selectable_columns and selectable_columns:
                column_value = selectable_columns[0]
            if selectable_columns:
                selected_column = row_cols[0].selectbox(
                    f"Column {idx + 1}",
                    selectable_columns,
                    index=selectable_columns.index(column_value) if column_value in selectable_columns else 0,
                    key=f"evidence_column_{row_key}",
                    label_visibility="collapsed",
                )
            else:
                selected_column = ""
            if row_cols[1].button("-", key=f"remove_evidence_{row_key}", width="stretch"):
                remove_idx = idx
            field.update({
                "column": selected_column,
                "kind": "auto",
                "enabled": True,
            })

        if remove_idx is not None:
            editable_fields.pop(remove_idx)
            st.session_state["evidence_fields"] = editable_fields
            st.rerun()

        if st.button(
            "+ Add factor",
            width="stretch",
            help="Add another column to consider when scoring possible matches.",
        ):
            if selectable_columns:
                existing_columns = {f.get("column") for f in editable_fields}
                next_column = next((c for c in selectable_columns if c not in existing_columns), selectable_columns[0])
                editable_fields.append({
                    "id": new_factor_id(),
                    "column": next_column,
                    "kind": "auto",
                    "weight": 0.08,
                    "enabled": True,
                })
                st.session_state["evidence_fields"] = normalize_factor_rows_for_ui(
                    editable_fields,
                    available_cols=available_cols,
                    entity_column=entity_column,
                )
            st.rerun()

    st.session_state["evidence_fields"] = normalize_factor_rows_for_ui(
        editable_fields,
        available_cols=available_cols,
        entity_column=entity_column,
    )
    column_config = normalize_column_config(
        {"entity_column": entity_column, "evidence_fields": st.session_state.get("evidence_fields", [])},
        available_columns=available_cols,
    )
    column_config["evidence_fields"] = [f for f in column_config["evidence_fields"] if f.get("column") != entity_column]
    st.session_state["column_config"] = column_config

    raw_df, rows_df, stats_df, auto_groups_df = build_base_data(file_bytes, sheet_name, column_config)
    auto_groups_df = ensure_auto_group_schema(auto_groups_df)
    load_session_feedback = st.session_state.pop("load_session_feedback", None)
    if load_session_feedback:
        level, message = load_session_feedback
        if level == "warning":
            st.warning(message)
        elif level == "error":
            st.error(message)
        else:
            st.success(message)

    if pd.api.types.is_numeric_dtype(raw_df[column_config["entity_column"]]):
        st.warning("This column looks numeric. The tool works best for text/entity columns such as names, places, categories, or labels.")
        st.info("Choose a text/entity column in Sidebar -> Column mapping to continue.")
        st.stop()

    resolved_names = resolved_names_from_auto(auto_groups_df, st.session_state["auto_status"])
    active_manual_decisions = active_manual_decisions_for_config(st.session_state["manual_decisions"], column_config)
    queue_settings_fingerprint = (
        sheet_name,
        config_fingerprint(column_config),
        fuzzy_threshold,
        min_manual_score,
        hide_reviewed_candidates,
    )
    if st.session_state.get("queue_settings_fingerprint") != queue_settings_fingerprint:
        st.session_state["pair_index"] = 0
        st.session_state["queue_settings_fingerprint"] = queue_settings_fingerprint

    full_queue_df = generate_candidate_pairs(
        stats_df,
        entity_column=column_config["entity_column"],
        resolved_names=resolved_names,
        fuzzy_threshold=fuzzy_threshold,
        column_config=column_config,
    )
    score_filtered_queue_df = full_queue_df[full_queue_df["score"] >= min_manual_score].reset_index(drop=True) if not full_queue_df.empty else full_queue_df.copy()
    visible_queue_df = score_filtered_queue_df.copy()
    if hide_reviewed_candidates and not visible_queue_df.empty:
        visible_queue_df = visible_queue_df[~visible_queue_df["pair_key"].isin(active_manual_decisions.keys())].reset_index(drop=True)
    reviewed_hidden_count = max(len(score_filtered_queue_df) - len(visible_queue_df), 0) if hide_reviewed_candidates else 0

    hidden_decision_count = len(st.session_state["manual_decisions"]) - len(active_manual_decisions)
    if hidden_decision_count > 0:
        st.info("Some saved decisions belong to a different primary column and are hidden from the current mapping.")
    history_df, mapping_df = build_merge_outputs(stats_df, auto_groups_df, st.session_state["auto_status"], active_manual_decisions)

    stat_lookup = stats_df.set_index("raw_name").to_dict("index")

    visible_auto_count = len(auto_groups_df[auto_groups_df["auto_group_key"].map(lambda gid: st.session_state["auto_status"].get(gid, "pending")) == "pending"]) if hide_reviewed_candidates else len(auto_groups_df)
    current_auto_group_keys = set(auto_groups_df["auto_group_key"].tolist()) if not auto_groups_df.empty else set()
    active_auto_decision_count = sum(
        1
        for gid, status in st.session_state["auto_status"].items()
        if gid in current_auto_group_keys and status in VALID_AUTO_STATUSES
    )
    active_decision_count = len(active_manual_decisions) + active_auto_decision_count
    merged_count = sum(1 for x in active_manual_decisions.values() if x.get("decision") == "merge")
    separate_count = sum(1 for x in active_manual_decisions.values() if x.get("decision") == "keep_separate")
    unsure_count = sum(1 for x in active_manual_decisions.values() if x.get("decision") == "unsure")
    dashboard1 = st.columns(5)
    dashboard1[0].metric("Unique primary values", len(stats_df))
    dashboard1[1].metric("Safe auto-groups visible / total", f"{visible_auto_count} / {len(auto_groups_df)}")
    dashboard1[2].metric("Manual candidates visible / generated", f"{len(visible_queue_df)} / {len(full_queue_df)}")
    dashboard1[3].metric("Decisions saved", active_decision_count)
    dashboard1[4].metric("Active merged names", len(mapping_df))
    dashboard2 = st.columns(3)
    dashboard2[0].metric("Merged", merged_count)
    dashboard2[1].metric("Kept separate", separate_count)
    dashboard2[2].metric("Unsure", unsure_count)

    review_history_records = []
    for _, row in auto_groups_df.iterrows():
        auto_decision_status = st.session_state["auto_status"].get(row["auto_group_key"], "pending")
        if auto_decision_status in VALID_AUTO_STATUSES:
            review_history_records.append({
                "decision_type": "auto_decision",
                "decision": auto_decision_status,
                "merge_source": "auto_group",
                "decision_id": row["auto_group_key"],
                "name_a": row.get("canonical_name", ""),
                "name_b": row.get("member_names", ""),
                "suggested_canonical": row.get("canonical_name", ""),
                "score": None,
                "reasons": row.get("reasons", ""),
                "evidence_summary": "",
                "reviewer_comment": "",
                "status": "active",
            })
    for pair_key, record in active_manual_decisions.items():
        evidence_summary = evidence_summary_from_scores(record.get("evidence_scores", []))
        review_history_records.append({
            "decision_type": "manual_decision",
            "decision": record.get("decision", ""),
            "merge_source": "manual_pair",
            "decision_id": pair_key,
            "name_a": record.get("name_a", ""),
            "name_b": record.get("name_b", ""),
            "suggested_canonical": record.get("suggested_canonical", ""),
            "score": record.get("score", None),
            "reasons": record.get("reasons", ""),
            "evidence_summary": evidence_summary,
            "reviewer_comment": record.get("reviewer_comment", ""),
            "status": "active",
        })
    review_history_df = pd.DataFrame(review_history_records)
    if not review_history_df.empty:
        review_history_df = review_history_df.sort_values(
            ["decision_type", "decision", "decision_id"],
            ascending=[True, True, True],
        ).reset_index(drop=True)

    tabs = st.tabs(["Safe Auto-Merges", "Manual Review Queue", "Review History", "Export"])

    with tabs[0]:
        st.subheader("Suggested safe groups")
        st.write("These are suggested safe groups based on strict-name matches. Nothing changes until you review and accept.")

        auto_preview_df = auto_groups_df.copy()
        auto_preview_df["status"] = auto_preview_df["auto_group_key"].map(lambda gid: st.session_state["auto_status"].get(gid, "pending"))
        if hide_reviewed_candidates:
            visible_auto_groups_df = auto_preview_df[auto_preview_df["status"] == "pending"].reset_index(drop=True)
        else:
            visible_auto_groups_df = auto_preview_df.reset_index(drop=True)
        reviewed_auto_hidden_count = len(auto_preview_df) - len(visible_auto_groups_df) if hide_reviewed_candidates else 0
        has_auto_groups = not auto_groups_df.empty
        auto_decision_count = len(st.session_state["auto_status"])
        selected_factor_columns = [f.get("column") for f in column_config.get("evidence_fields", []) if f.get("column")]
        selected_factor_columns = [column for column in selected_factor_columns if column in raw_df.columns and column != column_config["entity_column"]]
        selected_factor_columns = list(dict.fromkeys(selected_factor_columns))
        safe_overview_columns = [column_config["entity_column"]] + selected_factor_columns

        st.caption(f"Diagnostics — generated: {len(auto_preview_df)} | visible: {len(visible_auto_groups_df)} | reviewed hidden: {reviewed_auto_hidden_count}")
        with st.expander("Bulk review actions", expanded=True):
            actions = st.columns(3)
            if actions[0].button(
                "Review and accept all suggested safe groups",
                width="stretch",
                key="auto_bulk_accept_button",
                disabled=not has_auto_groups,
            ):
                for gid in auto_groups_df["auto_group_key"].tolist():
                    st.session_state["auto_status"][gid] = "accepted"
                st.session_state["auto_index"] = 0
                st.rerun()
            if actions[1].button(
                "Reject all suggested safe groups",
                width="stretch",
                key="auto_bulk_reject_button",
                disabled=not has_auto_groups,
            ):
                for gid in auto_groups_df["auto_group_key"].tolist():
                    st.session_state["auto_status"][gid] = "rejected"
                st.session_state["auto_index"] = 0
                st.rerun()
            if actions[2].button(
                "Clear all safe-group decisions",
                width="stretch",
                key="clear_all_auto_decisions_button",
                disabled=auto_decision_count == 0,
            ):
                st.session_state["auto_status"] = {}
                st.session_state["auto_index"] = 0
                st.rerun()
            if auto_decision_count == 0:
                st.info("No saved safe-group decisions to clear.")

        if visible_auto_groups_df.empty:
            st.info("No visible suggested safe groups under the current settings.")
            if hide_reviewed_candidates and reviewed_auto_hidden_count > 0:
                st.info("Some reviewed suggested safe groups are hidden. Turn off Hide reviewed candidates to inspect them.")
        else:
            visible_group_keys = visible_auto_groups_df["auto_group_key"].tolist()
            selected_open_group_key = st.session_state.get("auto_open_group_key")
            if selected_open_group_key and selected_open_group_key not in visible_group_keys:
                st.session_state["auto_open_group_key"] = None
                selected_open_group_key = None

            entity_values = raw_df[column_config["entity_column"]].fillna("").astype(str).str.strip()
            group_context_by_key = {}
            for _, group_row in visible_auto_groups_df.iterrows():
                group_key = group_row["auto_group_key"]
                group_id = group_row.get("auto_group_id", group_key)
                group_status = st.session_state["auto_status"].get(group_key, "pending")
                group_names = auto_group_names_for_display(group_row)
                if group_names:
                    group_mask = entity_values.isin(set(group_names))
                    group_rows = raw_df.loc[group_mask, safe_overview_columns].copy()
                else:
                    group_rows = pd.DataFrame(columns=safe_overview_columns)
                evidence_summary = build_safe_group_evidence_summary(group_rows, selected_factor_columns)
                member_count_value = group_row.get("member_count", len(group_names))
                if pd.notna(member_count_value):
                    try:
                        member_count = int(member_count_value)
                    except (TypeError, ValueError):
                        member_count = len(group_names)
                else:
                    member_count = len(group_names)
                group_context_by_key[group_key] = {
                    "group_id": group_id,
                    "status": group_status,
                    "group_names": group_names,
                    "group_rows": group_rows,
                    "evidence_summary": evidence_summary,
                    "canonical_name": group_row.get("canonical_name", ""),
                    "member_count": member_count,
                }

            st.markdown("#### Spreadsheet overview")
            st.caption("Shows only the primary column and selected `Dedupe based on` factor columns for each suggested safe group.")

            for group_key in visible_group_keys:
                context = group_context_by_key[group_key]
                with st.container(border=True):
                    group_block = st.columns([1.6, 6])
                    if group_block[0].button(
                        f"Open group {context['group_id']}",
                        width="stretch",
                        key=f"auto_open_group_{group_key}",
                    ):
                        st.session_state["auto_open_group_key"] = group_key
                        st.session_state["auto_index"] = visible_group_keys.index(group_key)
                        st.rerun()

                    with group_block[1]:
                        st.markdown(
                            f"**Group {context['group_id']}** | **Suggested canonical:** {context['canonical_name']} "
                            f"| **{context['member_count']} names** | **Status:** {context['status']}"
                        )
                        group_rows_for_overview = context["group_rows"]
                        if group_rows_for_overview.empty:
                            st.info("No source rows found for this group under the current primary column mapping.")
                        else:
                            st.dataframe(group_rows_for_overview[safe_overview_columns], width="stretch", hide_index=True)
                        st.caption(f"Why suggested: {context['evidence_summary']}")

            open_group_key = st.session_state.get("auto_open_group_key")
            if open_group_key and open_group_key in group_context_by_key:
                row = visible_auto_groups_df[visible_auto_groups_df["auto_group_key"] == open_group_key].iloc[0]
                group_context = group_context_by_key[open_group_key]
                position = visible_group_keys.index(open_group_key)
                group_names = group_context["group_names"]
                group_names_text = ", ".join(group_names) if group_names else "—"
                group_size_value = row.get("member_count")
                if pd.notna(group_size_value):
                    try:
                        group_size = int(group_size_value)
                    except (TypeError, ValueError):
                        group_size = len(group_names)
                else:
                    group_size = len(group_names)

                st.markdown("#### Selected group detail")
                with st.container(border=True):
                    st.markdown(f"### Suggested safe group {position + 1} of {len(visible_auto_groups_df)}")
                    summary = st.columns([1.2, 1.8, 2.8])
                    summary[0].markdown(f"**Status:** {group_context['status']}")
                    summary[1].markdown(f"**Suggested canonical:** {row['canonical_name']}")
                    summary[2].markdown(f"**{group_size} names in this suggested safe group**")
                    st.markdown(f"**Names in this group:** {group_names_text}")
                    st.markdown(f"**Reason / evidence summary:** {group_context['evidence_summary']}")
                    if row.get("reasons"):
                        st.caption(f"Strict match rationale: {row['reasons']}")

                detail_actions = st.columns(4)
                if detail_actions[0].button("Review and accept", width="stretch", key="auto_accept_selected_button"):
                    st.session_state["auto_status"][open_group_key] = "accepted"
                    if hide_reviewed_candidates:
                        st.session_state["auto_open_group_key"] = None
                    st.rerun()
                if detail_actions[1].button("Reject group", width="stretch", key="auto_reject_selected_button"):
                    st.session_state["auto_status"][open_group_key] = "rejected"
                    if hide_reviewed_candidates:
                        st.session_state["auto_open_group_key"] = None
                    st.rerun()
                if detail_actions[2].button("Undo", width="stretch", key="auto_undo_selected_button"):
                    st.session_state["auto_status"].pop(open_group_key, None)
                    st.rerun()
                if detail_actions[3].button("Back to overview", width="stretch", key="auto_back_to_overview_button"):
                    st.session_state["auto_open_group_key"] = None
                    st.rerun()

                with st.expander("Rows in this group (primary + selected dedupe-factor columns)", expanded=True):
                    group_rows_for_detail = group_context["group_rows"]
                    if group_rows_for_detail.empty:
                        st.info("No source rows found for this group under the current primary column mapping.")
                    else:
                        st.dataframe(group_rows_for_detail, width="stretch", hide_index=True)

                with st.expander("Additional group context", expanded=False):
                    detail = st.columns(2)
                    detail[0].metric("Member names", row["member_count"])
                    detail[1].metric("Total rows", row["total_rows"])
                    st.markdown(f"**Units:** {row['units'] or '—'}")
                    st.markdown(f"**Vessel types:** {row['vessel_types'] or '—'}")
                    st.caption(f"Evidence rows per side slider is set to {sample_rows}.")

    with tabs[1]:
        st.subheader("Manual review queue")
        st.write("These are remaining ambiguous candidates after accepted suggested safe groups are removed from manual review.")

        st.caption(f"Diagnostics — generated: {len(full_queue_df)} | score-filtered: {len(score_filtered_queue_df)} | visible: {len(visible_queue_df)} | reviewed hidden: {reviewed_hidden_count}")
        st.caption("Name similarity cutoff controls candidate generation. Minimum final score filters generated candidates after evidence scoring.")

        if visible_queue_df is None or visible_queue_df.empty:
            st.success("No visible manual-review candidates under the current settings.")
            if hide_reviewed_candidates and reviewed_hidden_count > 0:
                st.info("Some reviewed candidates are hidden. Turn off Hide reviewed candidates to inspect them.")
        else:
            idx = min(st.session_state["pair_index"], len(visible_queue_df) - 1)
            st.session_state["pair_index"] = idx
            row = visible_queue_df.iloc[idx]
            existing_decision_record = active_manual_decisions.get(row["pair_key"], {})
            decision = existing_decision_record.get("decision", "unreviewed")
            has_current_decision = has_saved_manual_decision(row["pair_key"], column_config["entity_column"])
            saved_reviewer_comment = existing_decision_record.get("reviewer_comment", "") or ""
            reviewer_comment_key = f"reviewer_comment_{row['pair_key']}"
            if reviewer_comment_key not in st.session_state:
                st.session_state[reviewer_comment_key] = saved_reviewer_comment

            with st.container(border=True):
                st.markdown(f"### Candidate {idx + 1} of {len(visible_queue_df)}")
                pair_names = st.columns([4, 1, 4])
                with pair_names[0]:
                    render_prominent_name(row["name_a"], "Side A", align="left")
                with pair_names[1]:
                    st.markdown("<div style='text-align:center; font-size:1.45rem; font-weight:700; margin-top:0.8rem;'>vs</div>", unsafe_allow_html=True)
                with pair_names[2]:
                    render_prominent_name(row["name_b"], "Side B", align="right")

                info = st.columns(3)
                info[0].markdown(f"**Score:** `{row['score']:.3f}`")
                info[1].markdown(f"**Decision status:** {decision}")
                info[2].markdown(f"**Suggested canonical:** {row['suggested_canonical']}")
                st.markdown(f"**Reasons:** {row['reasons']}")

            st.text_area(
                "Reviewer comment / reasoning",
                key=reviewer_comment_key,
                help="Optional note explaining why this pair should be merged, kept separate, or marked unsure.",
                placeholder=(
                    "e.g., same vessel type and cargo pattern, but transcription is uncertain\n"
                    "e.g., names are similar but years are too far apart\n"
                    "e.g., reviewer recognizes this as a common spelling variant"
                ),
            )
            current_reviewer_comment = st.session_state.get(reviewer_comment_key, "")

            nav = st.columns(6)
            if nav[0].button("Previous", width="stretch", key="manual_prev_button"):
                st.session_state["pair_index"] = max(st.session_state["pair_index"] - 1, 0)
                st.rerun()
            if nav[1].button("Merge", width="stretch", key="manual_merge_button"):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(
                    row,
                    "merge",
                    column_config["entity_column"],
                    reviewer_comment=current_reviewer_comment,
                )
                if not hide_reviewed_candidates:
                    st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(visible_queue_df) - 1)
                st.rerun()
            if nav[2].button("Keep separate", width="stretch", key="manual_keep_separate_button"):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(
                    row,
                    "keep_separate",
                    column_config["entity_column"],
                    reviewer_comment=current_reviewer_comment,
                )
                if not hide_reviewed_candidates:
                    st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(visible_queue_df) - 1)
                st.rerun()
            if nav[3].button("Unsure", width="stretch", key="manual_unsure_button"):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(
                    row,
                    "unsure",
                    column_config["entity_column"],
                    reviewer_comment=current_reviewer_comment,
                )
                if not hide_reviewed_candidates:
                    st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(visible_queue_df) - 1)
                st.rerun()
            if nav[4].button(
                "Undo",
                width="stretch",
                key="manual_undo_button",
                disabled=not has_current_decision,
                help="Undo this pair's saved decision (merge / keep separate / unsure).",
            ):
                undo_manual_decision_for_pair(row["pair_key"])
                st.rerun()
            if nav[5].button("Next", width="stretch", key="manual_next_button"):
                st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(visible_queue_df) - 1)
                st.rerun()
            if not has_current_decision:
                st.caption("No saved decision for this pair yet. Undo is enabled after Merge, Keep separate, or Unsure.")

            cols = st.columns(2)
            with cols[0]:
                show_summary_card("Side A", {"raw_name": row["name_a"], **stat_lookup[row["name_a"]]})
            with cols[1]:
                show_summary_card("Side B", {"raw_name": row["name_b"], **stat_lookup[row["name_b"]]})

            score_cols = st.columns(4)
            score_cols[0].metric("Raw name", f"{row['raw_name_score']:.3f}")
            score_cols[1].metric("Clean name", f"{row['clean_name_score']:.3f}")
            score_cols[2].metric("Name signal", f"{row.get('name_signal_score', 0.0):.3f}")
            score_cols[3].metric("Evidence signal", f"{row.get('evidence_signal_score', 0.5):.3f}")

            evidence_scores = row.get("evidence_scores", [])
            if isinstance(evidence_scores, list) and evidence_scores:
                evidence_df = pd.DataFrame(evidence_scores)
                evidence_df = evidence_df[["column", "kind", "score", "weight", "reason"]]
                st.markdown("#### Evidence score breakdown")
                st.dataframe(evidence_df, width="stretch", hide_index=True)

            selected_cols = [column_config["entity_column"]] + [f.get("column") for f in column_config.get("evidence_fields", []) if f.get("column")]
            display_columns = [c for c in selected_cols if c and c in raw_df.columns]
            display_columns = list(dict.fromkeys(display_columns))
            left_all_rows = raw_df[raw_df[column_config["entity_column"]].fillna("").astype(str).str.strip() == row["name_a"]][display_columns]
            right_all_rows = raw_df[raw_df[column_config["entity_column"]].fillna("").astype(str).str.strip() == row["name_b"]][display_columns]
            left_rows = left_all_rows.head(sample_rows)
            right_rows = right_all_rows.head(sample_rows)
            st.caption(
                f"Showing up to {sample_rows} evidence rows per side live. "
                f"Side A total rows: {len(left_all_rows)}. Side B total rows: {len(right_all_rows)}."
            )

            previews = st.columns(2)
            with previews[0]:
                with st.expander("Original rows for Side A", expanded=True):
                    st.dataframe(left_rows, width="stretch", hide_index=True)
            with previews[1]:
                with st.expander("Original rows for Side B", expanded=True):
                    st.dataframe(right_rows, width="stretch", hide_index=True)

    with tabs[2]:
        st.subheader("Review history and undo")
        st.write("All saved review decisions appear here. You can filter and undo selected decisions.")
        merge_history_undo_feedback = st.session_state.pop("merge_history_undo_feedback", None)
        if merge_history_undo_feedback:
            st.success(merge_history_undo_feedback)

        decision_filter = st.selectbox(
            "Decision filter",
            ["All decisions", "Merged", "Kept separate", "Unsure", "Auto decisions"],
            index=0,
            key="review_history_decision_filter",
        )
        if review_history_df.empty:
            st.info("No review decisions yet.")
        else:
            filter_map = {
                "Merged": lambda df: df[df["decision"] == "merge"],
                "Kept separate": lambda df: df[df["decision"] == "keep_separate"],
                "Unsure": lambda df: df[df["decision"] == "unsure"],
                "Auto decisions": lambda df: df[df["decision_type"] == "auto_decision"],
            }
            filtered_review_history_df = review_history_df.copy()
            if decision_filter in filter_map:
                filtered_review_history_df = filter_map[decision_filter](filtered_review_history_df).reset_index(drop=True)

            if filtered_review_history_df.empty:
                st.info("No decisions match this filter.")
            else:
                st.markdown("#### Active review history")
                selected_row_indexes = []
                selection_supported = True
                selection_response = None

                try:
                    selection_response = st.dataframe(
                        filtered_review_history_df,
                        width="stretch",
                        hide_index=True,
                        selection_mode="multi-row",
                        # Streamlit dataframe row selection currently requires rerun to refresh
                        # selected rows in session state.
                        on_select="rerun",
                        key="merge_history_selection_table",
                    )
                except TypeError:
                    selection_supported = False

                if selection_supported:
                    selected_rows = []

                    def extract_rows(selection_obj):
                        if selection_obj is None:
                            return []
                        if hasattr(selection_obj, "selection"):
                            selection_attr = selection_obj.selection
                            if hasattr(selection_attr, "rows"):
                                return list(selection_attr.rows or [])
                            if isinstance(selection_attr, dict):
                                return list(selection_attr.get("rows", []) or [])
                        if isinstance(selection_obj, dict):
                            return list(selection_obj.get("selection", {}).get("rows", []) or [])
                        return []

                    selected_rows = extract_rows(selection_response)
                    if not selected_rows:
                        selected_rows = extract_rows(st.session_state.get("merge_history_selection_table"))

                    unique_rows = sorted({int(row_idx) for row_idx in selected_rows if isinstance(row_idx, int) or str(row_idx).isdigit()})
                    selected_row_indexes = [row_idx for row_idx in unique_rows if 0 <= row_idx < len(filtered_review_history_df)]

                    undo_disabled = len(selected_row_indexes) == 0
                    if undo_disabled:
                        st.caption("Select one or more review-history rows to undo.")

                    if st.button(
                        "Undo selected decisions",
                        width="stretch",
                        key="undo_selected_merges_button",
                        disabled=undo_disabled,
                    ):
                        selected_records = filtered_review_history_df.iloc[selected_row_indexes].to_dict(orient="records")
                        decision_targets = {(record["decision_type"], record["decision_id"]) for record in selected_records}
                        for decision_type, decision_id in decision_targets:
                            if decision_type == "auto_decision":
                                st.session_state["auto_status"].pop(decision_id, None)
                            else:
                                st.session_state["manual_decisions"].pop(decision_id, None)
                        undone_count = len(decision_targets)
                        st.session_state["merge_history_undo_feedback"] = (
                            f"Undid {undone_count} selected decision."
                            if undone_count == 1
                            else f"Undid {undone_count} selected decisions."
                        )
                        st.session_state.pop("merge_history_selection_table", None)
                        st.rerun()
                else:
                    st.dataframe(filtered_review_history_df, width="stretch", hide_index=True, key="merge_history_fallback_table")
                    fallback_options = filtered_review_history_df["decision_id"].tolist()
                    selected_decision = st.selectbox("Select a decision to undo", fallback_options, key="selected_merge_to_undo")
                    if st.button("Undo selected decisions", width="stretch", key="undo_selected_merges_fallback_button"):
                        selected = filtered_review_history_df[filtered_review_history_df["decision_id"] == selected_decision].iloc[0]
                        if selected["decision_type"] == "auto_decision":
                            st.session_state["auto_status"].pop(selected_decision, None)
                        else:
                            st.session_state["manual_decisions"].pop(selected_decision, None)
                        st.session_state["merge_history_undo_feedback"] = "Undid 1 selected decision."
                        st.rerun()

        st.markdown("#### Current canonical mapping")
        st.dataframe(mapping_df, width="stretch", hide_index=True)
        st.markdown("#### Manual decision records")
        manual_decisions_df = pd.DataFrame(active_manual_decisions.values())
        if not manual_decisions_df.empty:
            manual_decisions_df = manual_decisions_df.copy()
            if "evidence_scores" in manual_decisions_df.columns:
                manual_decisions_df["evidence_summary"] = manual_decisions_df["evidence_scores"].apply(evidence_summary_from_scores)
                manual_decisions_df = manual_decisions_df.drop(columns=["evidence_scores"])
            display_columns = [
                "decision",
                "name_a",
                "name_b",
                "suggested_canonical",
                "score",
                "reasons",
                "evidence_summary",
                "reviewer_comment",
            ]
            available_display_columns = [col for col in display_columns if col in manual_decisions_df.columns]
            if available_display_columns:
                manual_decisions_df = manual_decisions_df[available_display_columns]
        st.dataframe(manual_decisions_df, width="stretch", hide_index=True)

    with tabs[3]:
        st.subheader("Export")
        auto_export = auto_groups_df.copy()
        if not auto_export.empty:
            auto_export["status"] = auto_export["auto_group_key"].map(lambda gid: st.session_state["auto_status"].get(gid, "pending"))
            auto_export["entity_column"] = column_config["entity_column"]
            auto_export["members"] = auto_export["members_list"].map(lambda x: " | ".join(x))
            auto_export["members_list"] = auto_export["members_list"].map(lambda x: " | ".join(x))
        pair_export = pd.DataFrame(st.session_state["manual_decisions"].values())
        if "reviewer_comment" not in pair_export.columns:
            pair_export["reviewer_comment"] = ""
        source_filename = getattr(uploaded, "name", "") or "uploaded_workbook.xlsx"
        session_payload = build_session_payload(
            source_filename=source_filename,
            workbook_bytes=file_bytes,
            sheet_name=sheet_name,
            raw_df=raw_df,
            column_config=column_config,
        )
        session_bytes = json.dumps(session_payload, indent=2).encode("utf-8")
        uploaded_session = st.file_uploader("Upload review session JSON", type=["json"], key="session_upload")
        if st.button("Load review session", width="content", key="load_review_session_button"):
            if uploaded_session is None:
                st.warning("Please upload a review session JSON first.")
            else:
                try:
                    data = json.loads(uploaded_session.getvalue().decode("utf-8"))
                except (json.JSONDecodeError, UnicodeDecodeError):
                    st.warning("Invalid JSON file. Please upload a valid review session JSON.")
                else:
                    is_valid, validation_message, normalized_session = validate_session_payload(data)
                    if not is_valid:
                        st.warning(validation_message)
                    else:
                        for record in normalized_session.get("manual_decisions", {}).values():
                            if isinstance(record, dict) and not isinstance(record.get("reviewer_comment", ""), str):
                                record["reviewer_comment"] = ""
                            elif isinstance(record, dict):
                                record.setdefault("reviewer_comment", "")
                        st.session_state["pending_loaded_session"] = normalized_session
                        st.rerun()

        base_name = getattr(uploaded, "name", "") or "cleaned_duplicate_review.xlsx"
        if base_name.lower().endswith(".xlsx"):
            cleaned_name = f"{base_name[:-5]}_reviewed.xlsx"
            standardized_name = f"{base_name[:-5]}_standardized.xlsx"
        else:
            cleaned_name = "cleaned_duplicate_review.xlsx"
            standardized_name = "standardized_duplicate_review.xlsx"
        cleaned_bytes = build_cleaned_workbook_bytes(
            raw_df=raw_df,
            mapping_df=mapping_df,
            history_df=history_df,
            manual_decisions=active_manual_decisions,
            auto_groups_df=auto_groups_df,
            column_config=column_config,
            sheet_name=sheet_name,
            workbook_bytes=file_bytes,
            candidate_queue_df=score_filtered_queue_df,
        )
        standardized_bytes = build_standardized_workbook_bytes(
            raw_df=raw_df,
            mapping_df=mapping_df,
            column_config=column_config,
            sheet_name=sheet_name,
            workbook_bytes=file_bytes,
        )
        st.markdown("#### Primary export")
        st.download_button("Download cleaned workbook", cleaned_bytes, cleaned_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch", key="download_cleaned_workbook")
        st.caption("Exports a reviewed copy of the workbook. The original uploaded file is not overwritten.")
        st.download_button("Download standardized workbook", standardized_bytes, standardized_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch", key="download_standardized_workbook")
        st.caption("This export replaces the selected entity column with reviewed canonical values only. It does not add audit columns.")
        with st.expander("Continue later", expanded=True):
            st.download_button(
                "Download review session JSON",
                session_bytes,
                "ship_review_session.json",
                "application/json",
                width="stretch",
                key="download_review_session_json",
            )
        with st.expander("Audit exports", expanded=True):
            st.download_button("auto decisions CSV", make_download_bytes(auto_export if not auto_export.empty else pd.DataFrame()), "ship_auto_merge_decisions.csv", "text/csv", width="stretch", key="download_auto_decisions_csv")
            st.download_button("manual decisions CSV", make_download_bytes(pair_export if not pair_export.empty else pd.DataFrame()), "ship_manual_review_decisions.csv", "text/csv", width="stretch", key="download_manual_decisions_csv")
            st.download_button("merge history CSV", make_download_bytes(history_df if not history_df.empty else pd.DataFrame()), "ship_merge_history.csv", "text/csv", width="stretch", key="download_merge_history_csv")
            st.download_button("canonical mapping CSV", make_download_bytes(mapping_df if not mapping_df.empty else pd.DataFrame()), "ship_canonical_mapping.csv", "text/csv", width="stretch", key="download_canonical_mapping_csv")

if __name__ == "__main__":
    app()
