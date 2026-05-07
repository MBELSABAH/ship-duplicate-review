import io
import json

import pandas as pd
import streamlit as st

from dedupe_engine import (
    active_manual_decisions_for_config,
    build_cleaned_workbook_bytes,
    build_manual_decision_record,
    build_merge_outputs,
    build_name_stats,
    build_safe_auto_groups,
    generate_candidate_pairs,
    make_download_bytes,
    now_iso,
    preprocess_rows,
    resolved_names_from_auto,
    set_to_text,
    SHIP_DEFAULTS,
)

APP_TITLE = "Duplicate Review MVP v3"
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
        "pair_index": 0,
        "column_config": {},
        "queue_settings_fingerprint": None,
        "pending_loaded_session": None,
        "load_session_feedback": None,
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
    st.session_state["pair_index"] = 0
    st.session_state["queue_settings_fingerprint"] = None
    if "pair_decisions" in st.session_state:
        st.session_state["pair_decisions"] = {}

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

def build_session_payload(sheet_name: str, column_config: dict):
    return {
        "app_version": "v3-stable-decisions-session",
        "saved_at": now_iso(),
        "sheet_name": sheet_name,
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

def app():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_state()

    st.title(APP_TITLE)
    st.caption("Duplicate-review workflow with safe auto-merges, manual review, history, and export.")

    with st.sidebar:
        st.subheader("Workbook")
        uploaded = st.file_uploader("Upload workbook", type=["xlsx", "xls"], key="upload_workbook")
        st.subheader("Matching")
        fuzzy_threshold = st.slider(
            "Name-match strictness",
            60,
            98,
            88,
            1,
            key="name_match_strictness_slider",
            help="Controls whether a pair is generated at all. Lower this if an expected pair is missing.",
        )
        min_manual_score = st.slider(
            "Overall evidence threshold",
            0.0,
            1.0,
            0.75,
            0.01,
            key="overall_evidence_threshold_slider",
            help="Filters generated pairs by combined evidence score. Lower this if generated pairs are hidden.",
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
        if st.button("Reset all decisions", use_container_width=True, key="reset_all_decisions_button"):
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
        if pending_sheet in sheet_names:
            st.session_state["sheet_select"] = pending_sheet

    with st.sidebar:
        sheet_name = st.selectbox("Sheet", sheet_names, index=0, key="sheet_select")

    raw_df_preview = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, nrows=50)
    available_cols = raw_df_preview.columns.tolist()
    text_candidates = [c for c in available_cols if raw_df_preview[c].dtype == "object"]
    default_entity = SHIP_DEFAULTS["entity_column"] if SHIP_DEFAULTS["entity_column"] in available_cols else (text_candidates[0] if text_candidates else available_cols[0])
    if pending_loaded_session:
        saved_cfg = pending_loaded_session.get("column_config", {}) if isinstance(pending_loaded_session.get("column_config", {}), dict) else {}
        missing_cols = [v for v in saved_cfg.values() if v and v not in available_cols]
        mapping_to_widget = {
            "entity_column": "entity_column_select",
            "year_column": "year_column_select",
            "type_column": "type_column_select",
            "amount_column": "amount_column_select",
            "unit_column": "unit_column_select",
            "notes_column_1": "notes_column_1_select",
            "notes_column_2": "notes_column_2_select",
        }
        for cfg_key, widget_key in mapping_to_widget.items():
            value = saved_cfg.get(cfg_key)
            fallback = default_entity if cfg_key == "entity_column" else "(None)"
            if value in available_cols:
                st.session_state[widget_key] = value
            else:
                st.session_state[widget_key] = fallback
        st.session_state["column_config"] = saved_cfg
        st.session_state["auto_status"] = pending_loaded_session.get("auto_status", {})
        st.session_state["manual_decisions"] = pending_loaded_session.get("manual_decisions", {})
        st.session_state["auto_index"] = 0
        st.session_state["pair_index"] = 0
        st.session_state["queue_settings_fingerprint"] = None
        if missing_cols:
            msg = f"Loaded session. Some saved columns are missing in the current sheet: {', '.join(sorted(set(missing_cols)))}. Current mapping fallback was used where needed."
            st.session_state["load_session_feedback"] = ("warning", msg)
        else:
            st.session_state["load_session_feedback"] = ("success", "Review session loaded.")

    saved_config = st.session_state.get("column_config", {}) or {}
    with st.sidebar:
        st.subheader("Column mapping")
        entity_default = saved_config.get("entity_column") if saved_config.get("entity_column") in available_cols else default_entity
        if pending_loaded_session:
            entity_column = st.selectbox("Primary column to deduplicate/review", available_cols, key="entity_column_select")
        else:
            entity_column = st.selectbox(
                "Primary column to deduplicate/review",
                available_cols,
                index=available_cols.index(entity_default),
                key="entity_column_select",
            )
        optional_options = ["(None)"] + available_cols
        def optional_select(label, cfg_key, widget_key):
            from_saved = saved_config.get(cfg_key)
            if from_saved in available_cols:
                default_val = from_saved
            else:
                default_val = SHIP_DEFAULTS[cfg_key] if SHIP_DEFAULTS[cfg_key] in available_cols else "(None)"
            if pending_loaded_session:
                return st.selectbox(label, optional_options, key=widget_key)
            return st.selectbox(label, optional_options, index=optional_options.index(default_val), key=widget_key)
        year_column = optional_select("Year/date column (optional)", "year_column", "year_column_select")
        type_column = optional_select("Type/category column (optional)", "type_column", "type_column_select")
        amount_column = optional_select("Amount column (optional)", "amount_column", "amount_column_select")
        unit_column = optional_select("Unit column (optional)", "unit_column", "unit_column_select")
        notes_column_1 = optional_select("Remarks/notes column 1 (optional)", "notes_column_1", "notes_column_1_select")
        notes_column_2 = optional_select("Remarks/notes column 2 (optional)", "notes_column_2", "notes_column_2_select")

    def none_if_placeholder(v):
        return None if v == "(None)" else v
    column_config = {
        "entity_column": entity_column,
        "year_column": none_if_placeholder(year_column),
        "type_column": none_if_placeholder(type_column),
        "amount_column": none_if_placeholder(amount_column),
        "unit_column": none_if_placeholder(unit_column),
        "notes_column_1": none_if_placeholder(notes_column_1),
        "notes_column_2": none_if_placeholder(notes_column_2),
    }
    st.session_state["column_config"] = column_config

    raw_df, rows_df, stats_df, auto_groups_df = build_base_data(file_bytes, sheet_name, column_config)
    auto_groups_df = ensure_auto_group_schema(auto_groups_df)
    load_session_feedback = st.session_state.pop("load_session_feedback", None)
    if load_session_feedback:
        level, message = load_session_feedback
        if level == "warning":
            st.warning(message)
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
        tuple(sorted(column_config.items())),
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
    merged_count = sum(1 for x in active_manual_decisions.values() if x.get("decision") == "merge")
    separate_count = sum(1 for x in active_manual_decisions.values() if x.get("decision") == "keep_separate")
    unsure_count = sum(1 for x in active_manual_decisions.values() if x.get("decision") == "unsure")
    dashboard1 = st.columns(5)
    dashboard1[0].metric("Unique primary values", len(stats_df))
    dashboard1[1].metric("Safe auto-groups visible / total", f"{visible_auto_count} / {len(auto_groups_df)}")
    dashboard1[2].metric("Manual candidates visible / generated", f"{len(visible_queue_df)} / {len(full_queue_df)}")
    dashboard1[3].metric("Decisions saved", len(st.session_state["manual_decisions"]) + len(st.session_state["auto_status"]))
    dashboard1[4].metric("Active merged names", len(mapping_df))
    dashboard2 = st.columns(3)
    dashboard2[0].metric("Merged", merged_count)
    dashboard2[1].metric("Kept separate", separate_count)
    dashboard2[2].metric("Unsure", unsure_count)

    tabs = st.tabs(["Safe Auto-Merges", "Manual Review Queue", "Merge History", "Export"])

    with tabs[0]:
        st.subheader("Safe auto-merges")
        st.write("These merges trigger only when names share the same strict key after removing punctuation, case, and spaces.")

        auto_preview_df = auto_groups_df.copy()
        auto_preview_df["status"] = auto_preview_df["auto_group_key"].map(lambda gid: st.session_state["auto_status"].get(gid, "pending"))
        if hide_reviewed_candidates:
            visible_auto_groups_df = auto_preview_df[auto_preview_df["status"] == "pending"].reset_index(drop=True)
        else:
            visible_auto_groups_df = auto_preview_df.reset_index(drop=True)
        reviewed_auto_hidden_count = len(auto_preview_df) - len(visible_auto_groups_df) if hide_reviewed_candidates else 0
        has_auto_groups = not auto_groups_df.empty
        auto_decision_count = len(st.session_state["auto_status"])

        st.caption(f"Diagnostics — generated: {len(auto_preview_df)} | visible: {len(visible_auto_groups_df)} | reviewed hidden: {reviewed_auto_hidden_count}")
        with st.expander("Bulk actions", expanded=True):
            actions = st.columns(3)
            if actions[0].button(
                "Accept all safe auto-merges",
                use_container_width=True,
                key="auto_bulk_accept_button",
                disabled=not has_auto_groups,
            ):
                for gid in auto_groups_df["auto_group_key"].tolist():
                    st.session_state["auto_status"][gid] = "accepted"
                st.session_state["auto_index"] = 0
                st.rerun()
            if actions[1].button(
                "Reject all safe auto-merges",
                use_container_width=True,
                key="auto_bulk_reject_button",
                disabled=not has_auto_groups,
            ):
                for gid in auto_groups_df["auto_group_key"].tolist():
                    st.session_state["auto_status"][gid] = "rejected"
                st.session_state["auto_index"] = 0
                st.rerun()
            if actions[2].button(
                "Clear all auto-merge decisions",
                use_container_width=True,
                key="clear_all_auto_decisions_button",
                disabled=auto_decision_count == 0,
            ):
                st.session_state["auto_status"] = {}
                st.session_state["auto_index"] = 0
                st.rerun()
            if auto_decision_count == 0:
                st.info("No saved auto-merge decisions to clear.")

        if visible_auto_groups_df.empty:
            st.info("No visible safe auto-groups under the current settings.")
            if hide_reviewed_candidates and reviewed_auto_hidden_count > 0:
                st.info("Some reviewed auto-groups are hidden. Turn off Hide reviewed candidates to inspect them.")
        else:
            idx = min(st.session_state["auto_index"], len(visible_auto_groups_df) - 1)
            st.session_state["auto_index"] = idx
            row = visible_auto_groups_df.iloc[idx]
            status = st.session_state["auto_status"].get(row["auto_group_key"], "pending")

            with st.container(border=True):
                info = st.columns([2, 1.5, 1.5, 2])
                info[0].markdown(f"### Group {idx + 1} of {len(visible_auto_groups_df)}")
                info[1].markdown(f"**Status:** {status}")
                info[2].markdown(f"**Canonical:** {row['canonical_name']}")
                info[3].markdown(f"**Reason:** {row['reasons']}")
            nav = st.columns(5)
            if nav[0].button("Previous", use_container_width=True, key="auto_prev_button"):
                st.session_state["auto_index"] = max(st.session_state["auto_index"] - 1, 0)
                st.rerun()
            if nav[1].button("Accept", use_container_width=True, key="auto_accept_button"):
                st.session_state["auto_status"][row["auto_group_key"]] = "accepted"
                if not hide_reviewed_candidates:
                    st.session_state["auto_index"] = min(st.session_state["auto_index"] + 1, len(visible_auto_groups_df) - 1)
                st.rerun()
            if nav[2].button("Reject", use_container_width=True, key="auto_reject_button"):
                st.session_state["auto_status"][row["auto_group_key"]] = "rejected"
                if not hide_reviewed_candidates:
                    st.session_state["auto_index"] = min(st.session_state["auto_index"] + 1, len(visible_auto_groups_df) - 1)
                st.rerun()
            if nav[3].button("Undo", use_container_width=True, key="auto_undo_button"):
                st.session_state["auto_status"].pop(row["auto_group_key"], None)
                st.rerun()
            if nav[4].button("Next", use_container_width=True, key="auto_next_button"):
                st.session_state["auto_index"] = min(st.session_state["auto_index"] + 1, len(visible_auto_groups_df) - 1)
                st.rerun()

            detail = st.columns(4)
            detail[0].metric("Member names", row["member_count"])
            detail[1].metric("Total rows", row["total_rows"])
            detail[2].metric("Min year", row["min_year"] if pd.notna(row["min_year"]) else "—")
            detail[3].metric("Max year", row["max_year"] if pd.notna(row["max_year"]) else "—")

            st.markdown(f"**Members:** {row['member_names']}")
            st.markdown(f"**Units:** {row['units'] or '—'}")
            st.markdown(f"**Vessel types:** {row['vessel_types'] or '—'}")
            st.caption(f"Evidence rows per side slider is set to {sample_rows}.")

            preview = auto_preview_df[["auto_group_id", "auto_group_key", "canonical_name", "member_count", "total_rows", "reasons", "status"]].copy()
            st.dataframe(preview, use_container_width=True, hide_index=True)

    with tabs[1]:
        st.subheader("Manual review queue")
        st.write("These are the remaining ambiguous candidates after any accepted safe auto-merges are removed from review.")

        st.caption(f"Diagnostics — generated: {len(full_queue_df)} | score-filtered: {len(score_filtered_queue_df)} | visible: {len(visible_queue_df)} | reviewed hidden: {reviewed_hidden_count}")
        st.caption("Name-match strictness controls candidate generation. Overall evidence threshold filters generated candidates.")

        if visible_queue_df is None or visible_queue_df.empty:
            st.success("No visible manual-review candidates under the current settings.")
            if hide_reviewed_candidates and reviewed_hidden_count > 0:
                st.info("Some reviewed candidates are hidden. Turn off Hide reviewed candidates to inspect them.")
        else:
            idx = min(st.session_state["pair_index"], len(visible_queue_df) - 1)
            st.session_state["pair_index"] = idx
            row = visible_queue_df.iloc[idx]
            decision = active_manual_decisions.get(row["pair_key"], {}).get("decision", "unreviewed")

            with st.container(border=True):
                info = st.columns([2, 1.3, 1.3, 2])
                info[0].markdown(f"### Candidate {idx + 1} of {len(visible_queue_df)}")
                info[1].markdown(f"**Score:** `{row['score']:.3f}`")
                info[2].markdown(f"**Decision:** {decision}")
                info[3].markdown(f"**Suggested canonical:** {row['suggested_canonical']}")

            nav = st.columns(6)
            if nav[0].button("Previous", use_container_width=True, key="manual_prev_button"):
                st.session_state["pair_index"] = max(st.session_state["pair_index"] - 1, 0)
                st.rerun()
            if nav[1].button("Merge", use_container_width=True, key="manual_merge_button"):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(row, "merge", column_config["entity_column"])
                if not hide_reviewed_candidates:
                    st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(visible_queue_df) - 1)
                st.rerun()
            if nav[2].button("Keep separate", use_container_width=True, key="manual_keep_separate_button"):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(row, "keep_separate", column_config["entity_column"])
                if not hide_reviewed_candidates:
                    st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(visible_queue_df) - 1)
                st.rerun()
            if nav[3].button("Unsure", use_container_width=True, key="manual_unsure_button"):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(row, "unsure", column_config["entity_column"])
                if not hide_reviewed_candidates:
                    st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(visible_queue_df) - 1)
                st.rerun()
            if nav[4].button("Undo", use_container_width=True, key="manual_undo_button"):
                st.session_state["manual_decisions"].pop(row["pair_key"], None)
                st.rerun()
            if nav[5].button("Next", use_container_width=True, key="manual_next_button"):
                st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(visible_queue_df) - 1)
                st.rerun()

            st.markdown("#### Why this pair was flagged")
            st.write(row["reasons"])

            cols = st.columns(2)
            with cols[0]:
                show_summary_card("Side A", {"raw_name": row["name_a"], **stat_lookup[row["name_a"]]})
            with cols[1]:
                show_summary_card("Side B", {"raw_name": row["name_b"], **stat_lookup[row["name_b"]]})

            scores = st.columns(6)
            scores[0].metric("Raw name", f"{row['raw_name_score']:.3f}")
            scores[1].metric("Clean name", f"{row['clean_name_score']:.3f}")
            scores[2].metric("Years", f"{row['year_score']:.3f}")
            scores[3].metric("Units", f"{row['unit_score']:.3f}")
            scores[4].metric("Type/category evidence", f"{row['type_score']:.3f}")
            scores[5].metric("Amount evidence", f"{row['cargo_amount_score']:.3f}")
            st.caption("Cargo amount is weak supporting evidence, not registered vessel tonnage.")

            selected_cols = [column_config["entity_column"], column_config["year_column"], column_config["type_column"], column_config["amount_column"], column_config["unit_column"], column_config["notes_column_1"], column_config["notes_column_2"]]
            id_candidates = ["OID", "Volume", "Page", "Day", "Month", "Year", "Entry no."]
            display_columns = [c for c in selected_cols + id_candidates if c and c in raw_df.columns]
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
                    st.dataframe(left_rows, use_container_width=True, hide_index=True)
            with previews[1]:
                with st.expander("Original rows for Side B", expanded=True):
                    st.dataframe(right_rows, use_container_width=True, hide_index=True)

    with tabs[2]:
        st.subheader("Merge history and undo")
        st.write("Every active merge appears here with the reason it was merged. You can undo anything from this list.")

        if history_df.empty:
            st.info("No active merges yet.")
        else:
            st.dataframe(history_df, use_container_width=True, hide_index=True)
            st.markdown("#### Active merge history")
            selected_merge = st.selectbox("Select a merge to undo", history_df["merge_id"].tolist(), key="selected_merge_to_undo")
            if st.button("Undo selected merge", use_container_width=True, key="undo_selected_merge_button"):
                selected = history_df[history_df["merge_id"] == selected_merge].iloc[0]
                if selected["merge_source"] == "auto_group":
                    st.session_state["auto_status"].pop(selected_merge, None)
                else:
                    st.session_state["manual_decisions"].pop(selected_merge, None)
                st.rerun()

        st.markdown("#### Current canonical mapping")
        st.dataframe(mapping_df, use_container_width=True, hide_index=True)
        st.markdown("#### Manual decision records")
        st.dataframe(pd.DataFrame(active_manual_decisions.values()), use_container_width=True, hide_index=True)

    with tabs[3]:
        st.subheader("Export")
        auto_export = auto_groups_df.copy()
        if not auto_export.empty:
            auto_export["status"] = auto_export["auto_group_key"].map(lambda gid: st.session_state["auto_status"].get(gid, "pending"))
            auto_export["entity_column"] = column_config["entity_column"]
            auto_export["members"] = auto_export["members_list"].map(lambda x: " | ".join(x))
            auto_export["members_list"] = auto_export["members_list"].map(lambda x: " | ".join(x))
        pair_export = pd.DataFrame(st.session_state["manual_decisions"].values())
        session_payload = build_session_payload(sheet_name, column_config)
        session_bytes = json.dumps(session_payload, indent=2).encode("utf-8")
        uploaded_session = st.file_uploader("Upload review session JSON", type=["json"], key="session_upload")
        if st.button("Load review session", use_container_width=False, key="load_review_session_button"):
            if uploaded_session is None:
                st.warning("Please upload a review session JSON first.")
            else:
                try:
                    data = json.loads(uploaded_session.getvalue().decode("utf-8"))
                except json.JSONDecodeError:
                    st.warning("Invalid JSON file. Please upload a valid review session JSON.")
                else:
                    if not isinstance(data, dict):
                        st.warning("Invalid session format. The JSON root must be an object.")
                    else:
                        st.session_state["pending_loaded_session"] = data
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
        standardized_bytes = build_cleaned_workbook_bytes(
            raw_df=raw_df,
            mapping_df=mapping_df,
            history_df=history_df,
            manual_decisions=active_manual_decisions,
            auto_groups_df=auto_groups_df,
            column_config=column_config,
            sheet_name=sheet_name,
            workbook_bytes=file_bytes,
            candidate_queue_df=score_filtered_queue_df,
            overwrite_entity_column=True,
        )
        st.markdown("#### Primary export")
        st.download_button("Download cleaned workbook", cleaned_bytes, cleaned_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="download_cleaned_workbook")
        st.caption("Exports a reviewed copy of the workbook. The original uploaded file is not overwritten.")
        st.download_button("Download standardized workbook", standardized_bytes, standardized_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="download_standardized_workbook")
        st.caption("This export replaces the selected entity column with reviewed canonical values while keeping audit columns.")
        with st.expander("Continue later", expanded=True):
            session_actions = st.columns([1, 1.2])
            session_actions[0].download_button("Download review session JSON", session_bytes, "ship_review_session.json", "application/json", use_container_width=True, key="download_review_session_json")
        with st.expander("Audit exports", expanded=True):
            st.download_button("auto decisions CSV", make_download_bytes(auto_export if not auto_export.empty else pd.DataFrame()), "ship_auto_merge_decisions.csv", "text/csv", use_container_width=True, key="download_auto_decisions_csv")
            st.download_button("manual decisions CSV", make_download_bytes(pair_export if not pair_export.empty else pd.DataFrame()), "ship_manual_review_decisions.csv", "text/csv", use_container_width=True, key="download_manual_decisions_csv")
            st.download_button("merge history CSV", make_download_bytes(history_df if not history_df.empty else pd.DataFrame()), "ship_merge_history.csv", "text/csv", use_container_width=True, key="download_merge_history_csv")
            st.download_button("canonical mapping CSV", make_download_bytes(mapping_df if not mapping_df.empty else pd.DataFrame()), "ship_canonical_mapping.csv", "text/csv", use_container_width=True, key="download_canonical_mapping_csv")

if __name__ == "__main__":
    app()
