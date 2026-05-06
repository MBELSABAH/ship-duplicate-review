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

def init_state():
    defaults = {
        "auto_status": {},
        "manual_decisions": {},
        "auto_index": 0,
        "pair_index": 0,
        "column_config": {},
        "queue_settings_fingerprint": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

@st.cache_data(show_spinner=False)
def load_sheet_names(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names

@st.cache_data(show_spinner=False)
def build_base_data(file_bytes: bytes, sheet_name: str, column_config: dict):
    raw_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
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

def app():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_state()
    st.title(APP_TITLE)

    with st.sidebar:
        st.header("Workbook")
        uploaded = st.file_uploader("Upload workbook", type=["xlsx", "xls"])

        st.header("Matching settings")
        fuzzy_threshold = st.slider("Name match strictness", 60, 98, 88, 1)
        min_manual_score = st.slider("Overall evidence threshold", 0.0, 1.0, 0.75, 0.01)

        st.header("Review behavior")
        hide_reviewed_candidates = st.checkbox("Hide reviewed candidates", value=True)
        sample_rows = st.slider("Original rows preview per side", 3, 12, 5, 1)

        st.header("Reset")
        if st.button("Reset all decisions", use_container_width=True):
            st.session_state["auto_status"] = {}
            st.session_state["manual_decisions"] = {}
            st.session_state["auto_index"] = 0
            st.session_state["pair_index"] = 0
            st.rerun()

    if not uploaded:
        st.info("Upload the workbook to start.")
        st.stop()

    file_bytes = uploaded.getvalue()
    sheet_names = load_sheet_names(file_bytes)
    with st.sidebar:
        sheet_name = st.selectbox("Sheet", sheet_names, index=0)
        st.header("Column mapping")

    raw_df_preview = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, nrows=50)
    available_cols = raw_df_preview.columns.tolist()
    text_candidates = [c for c in available_cols if raw_df_preview[c].dtype == "object"]
    default_entity = SHIP_DEFAULTS["entity_column"] if SHIP_DEFAULTS["entity_column"] in available_cols else (text_candidates[0] if text_candidates else available_cols[0])
    saved_config = st.session_state.get("column_config", {}) or {}

    with st.sidebar:
        entity_default = saved_config.get("entity_column") if saved_config.get("entity_column") in available_cols else default_entity
        entity_column = st.selectbox("Primary column", available_cols, index=available_cols.index(entity_default), key="entity_column_select")
        optional_options = ["(None)"] + available_cols
        def optional_select(label, cfg_key, widget_key):
            from_saved = saved_config.get(cfg_key)
            default_val = from_saved if from_saved in available_cols else (SHIP_DEFAULTS[cfg_key] if SHIP_DEFAULTS[cfg_key] in available_cols else "(None)")
            return st.selectbox(label, optional_options, index=optional_options.index(default_val), key=widget_key)
        year_column = optional_select("Year/date column", "year_column", "year_column_select")
        type_column = optional_select("Type/category column", "type_column", "type_column_select")
        amount_column = optional_select("Amount column", "amount_column", "amount_column_select")
        unit_column = optional_select("Unit column", "unit_column", "unit_column_select")
        notes_column_1 = optional_select("Notes column 1", "notes_column_1", "notes_column_1_select")
        notes_column_2 = optional_select("Notes column 2", "notes_column_2", "notes_column_2_select")

    none_if_placeholder = lambda v: None if v == "(None)" else v
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
    resolved_names = resolved_names_from_auto(auto_groups_df, st.session_state["auto_status"])
    active_manual_decisions = active_manual_decisions_for_config(st.session_state["manual_decisions"], column_config)
    full_queue_df = generate_candidate_pairs(stats_df, entity_column=column_config["entity_column"], resolved_names=resolved_names, fuzzy_threshold=fuzzy_threshold)
    score_filtered_queue_df = full_queue_df[full_queue_df["score"] >= min_manual_score].reset_index(drop=True) if not full_queue_df.empty else full_queue_df.copy()
    visible_queue_df = score_filtered_queue_df.copy()
    if hide_reviewed_candidates and not visible_queue_df.empty:
        visible_queue_df = visible_queue_df[~visible_queue_df["pair_key"].isin(active_manual_decisions.keys())].reset_index(drop=True)
    history_df, mapping_df = build_merge_outputs(stats_df, auto_groups_df, st.session_state["auto_status"], active_manual_decisions)

    auto_preview_df = auto_groups_df.copy()
    auto_preview_df["status"] = auto_preview_df["auto_group_key"].map(lambda gid: st.session_state["auto_status"].get(gid, "pending")) if not auto_preview_df.empty else pd.Series(dtype=str)
    visible_auto_groups_df = auto_preview_df[auto_preview_df["status"] == "pending"].reset_index(drop=True) if hide_reviewed_candidates else auto_preview_df.reset_index(drop=True)

    metrics = st.columns(5)
    metrics[0].metric("Unique primary values", len(stats_df))
    metrics[1].metric("Safe auto-groups visible/total", f"{len(visible_auto_groups_df)}/{len(auto_groups_df)}")
    metrics[2].metric("Manual candidates visible/generated", f"{len(visible_queue_df)}/{len(full_queue_df)}")
    metrics[3].metric("Manual decisions saved", len(active_manual_decisions))
    metrics[4].metric("Active merged names", len(mapping_df))

    tabs = st.tabs(["1) Safe Auto-Merges", "2) Manual Review Queue", "3) Merge History + Undo", "4) Export"])
    with tabs[0]:
        if not visible_auto_groups_df.empty:
            with st.expander("Bulk actions", expanded=False):
                a = st.columns(3)
                if a[0].button("Accept all visible safe auto-merges", use_container_width=True):
                    for gid in visible_auto_groups_df["auto_group_key"]: st.session_state["auto_status"][gid] = "accepted"; st.rerun()
                if a[1].button("Reject all visible safe auto-merges", use_container_width=True):
                    for gid in visible_auto_groups_df["auto_group_key"]: st.session_state["auto_status"][gid] = "rejected"; st.rerun()
                if a[2].button("Clear all visible auto decisions", use_container_width=True):
                    for gid in visible_auto_groups_df["auto_group_key"]: st.session_state["auto_status"].pop(gid, None); st.rerun()
            idx = min(st.session_state["auto_index"], len(visible_auto_groups_df)-1)
            row = visible_auto_groups_df.iloc[idx]
            st.markdown(f"### Group {idx+1} of {len(visible_auto_groups_df)}")
            st.markdown(f"**Canonical:** {row['canonical_name']}  ")
            st.markdown(f"**Members:** {row['member_names']}  ")
            st.markdown(f"**Row count:** {row['total_rows']}  ")
            st.markdown(f"**Year range:** {(row['min_year'] if pd.notna(row['min_year']) else '—')} → {(row['max_year'] if pd.notna(row['max_year']) else '—')}  ")
            st.markdown(f"**Reason:** {row['reasons']}")
            n=st.columns(5)
            if n[0].button("Previous", use_container_width=True): st.session_state["auto_index"]=max(idx-1,0); st.rerun()
            if n[1].button("Accept", use_container_width=True): st.session_state["auto_status"][row["auto_group_key"]]="accepted"; st.rerun()
            if n[2].button("Reject", use_container_width=True): st.session_state["auto_status"][row["auto_group_key"]]="rejected"; st.rerun()
            if n[3].button("Undo", use_container_width=True): st.session_state["auto_status"].pop(row["auto_group_key"],None); st.rerun()
            if n[4].button("Next", use_container_width=True): st.session_state["auto_index"]=min(idx+1,len(visible_auto_groups_df)-1); st.rerun()

    with tabs[1]:
        if not visible_queue_df.empty:
            idx=min(st.session_state["pair_index"],len(visible_queue_df)-1); row=visible_queue_df.iloc[idx]
            st.markdown(f"### Candidate {idx+1} of {len(visible_queue_df)}")
            st.markdown(f"**Score:** `{row['score']:.3f}`")
            st.markdown(f"**Suggested canonical:** {row['suggested_canonical']}")
            st.markdown(f"**Name A:** {row['name_a']}")
            st.markdown(f"**Name B:** {row['name_b']}")
            st.markdown("#### Concise reasons")
            st.write(row["reasons"])
            sc=st.columns(6)
            for i,(lbl,key) in enumerate([("Raw","raw_name_score"),("Clean","clean_name_score"),("Years","year_score"),("Units","unit_score"),("Type","type_score"),("Amount","cargo_amount_score")]): sc[i].metric(lbl,f"{row[key]:.3f}")
            nav=st.columns(6)
            if nav[0].button("Previous",use_container_width=True): st.session_state["pair_index"]=max(idx-1,0); st.rerun()
            if nav[1].button("Merge",use_container_width=True): st.session_state["manual_decisions"][row["pair_key"]]=build_manual_decision_record(row,"merge",column_config["entity_column"]); st.rerun()
            if nav[2].button("Keep separate",use_container_width=True): st.session_state["manual_decisions"][row["pair_key"]]=build_manual_decision_record(row,"keep_separate",column_config["entity_column"]); st.rerun()
            if nav[3].button("Unsure",use_container_width=True): st.session_state["manual_decisions"][row["pair_key"]]=build_manual_decision_record(row,"unsure",column_config["entity_column"]); st.rerun()
            if nav[4].button("Undo",use_container_width=True): st.session_state["manual_decisions"].pop(row["pair_key"],None); st.rerun()
            if nav[5].button("Next",use_container_width=True): st.session_state["pair_index"]=min(idx+1,len(visible_queue_df)-1); st.rerun()
            selected_cols=[column_config[k] for k in ["entity_column","year_column","type_column","amount_column","unit_column","notes_column_1","notes_column_2"] if column_config[k] in raw_df.columns]
            l=raw_df[raw_df[column_config["entity_column"]].fillna("").astype(str).str.strip()==row["name_a"]][selected_cols].head(sample_rows)
            r=raw_df[raw_df[column_config["entity_column"]].fillna("").astype(str).str.strip()==row["name_b"]][selected_cols].head(sample_rows)
            with st.expander("Original rows: Name A", expanded=False): st.dataframe(l,use_container_width=True,hide_index=True)
            with st.expander("Original rows: Name B", expanded=False): st.dataframe(r,use_container_width=True,hide_index=True)
            with st.expander("Technical details", expanded=False):
                show_summary_card("Side A", {"raw_name": row["name_a"]})
                show_summary_card("Side B", {"raw_name": row["name_b"]})

    with tabs[3]:
        st.caption("Cleaned workbook = reviewed Excel copy.")
        st.caption("Session JSON saves review progress. CSV exports are audit/debug logs.")
        session_payload = build_session_payload(sheet_name, column_config)
        session_bytes = json.dumps(session_payload, indent=2).encode("utf-8")
        with st.expander("Continue later", expanded=False):
            up = st.file_uploader("Upload review session JSON", type=["json"], key="session_upload")
            st.download_button("Download session JSON", session_bytes, "ship_review_session.json", "application/json", use_container_width=True)
        with st.expander("Audit exports", expanded=False):
            st.download_button("Manual review decisions CSV", make_download_bytes(pd.DataFrame(st.session_state["manual_decisions"].values())), "ship_manual_review_decisions.csv", "text/csv", use_container_width=True)
        base_name = getattr(uploaded, "name", "") or "cleaned_duplicate_review.xlsx"
        cleaned_name = f"{base_name[:-5]}_reviewed.xlsx" if base_name.lower().endswith(".xlsx") else "cleaned_duplicate_review.xlsx"
        cleaned_bytes = build_cleaned_workbook_bytes(raw_df,mapping_df,history_df,active_manual_decisions,auto_groups_df,column_config,sheet_name,file_bytes,score_filtered_queue_df)
        st.download_button("Download cleaned workbook", cleaned_bytes, cleaned_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

if __name__ == "__main__":
    app()
