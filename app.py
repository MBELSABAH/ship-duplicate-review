
import io
import json
import re
import itertools
import hashlib
from collections import Counter, defaultdict
from datetime import datetime, timezone

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz


APP_TITLE = "Duplicate Review MVP v3"
SHIP_DEFAULTS = {
    "entity_column": "Name of Vessel",
    "year_column": "Year",
    "type_column": "Type of Veseel",
    "amount_column": "Amount (primary)",
    "unit_column": "Unit (primary)",
    "notes_column_1": "Remarks from ledger",
    "notes_column_2": "Notes from transcriber",
}


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def clean_name(value: str) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().lower()
    s = s.replace("&", " and ")
    s = s.replace("?", " ")
    s = re.sub(r"[\.,;:'\"`´‘’“”()\[\]{}_/\\|-]+", " ", s)
    s = normalize_spaces(s)
    return s


def stable_hash(text: str, length: int = 12) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:length]


def make_pair_key(name_a: str, name_b: str, entity_column: str = "Name of Vessel") -> str:
    normalized_names = sorted([clean_name(name_a), clean_name(name_b)])
    raw_key = entity_column + "::" + "||".join(normalized_names)
    return "P" + stable_hash(raw_key)


def make_auto_group_key(strict_key: str, entity_column: str) -> str:
    return "A" + stable_hash(f"auto::{entity_column}::{strict_key}")


def strict_name_key(value: str) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().lower()
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def clean_vessel_type(value: str) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().lower()
    s = re.sub(r"[\.,;:'\"`´‘’“”()\[\]{}_/\\|-]+", " ", s)
    s = normalize_spaces(s)
    mapping = {
        "prop": "propeller",
        "ss": "steamship",
        "s s": "steamship",
        "stmr": "steamer",
        "stm": "steamer",
        "str": "steamer",
        "sch": "schooner",
        "schr": "schooner",
        "shnr": "schooner",
        "sb": "sailboat",
        "s b": "sailboat",
    }
    return mapping.get(s, s)


def clean_unit(value: str) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().lower()
    s = re.sub(r"[\.,;:'\"`´‘’“”()\[\]{}_/\\|-]+", " ", s)
    s = normalize_spaces(s)
    mapping = {
        "ton": "tons",
        "tons": "tons",
        "tns": "tons",
        "toise": "toise",
        "barrels": "barrels",
        "barrel": "barrels",
        "baskets": "baskets",
        "crates": "crates",
        "boxes": "boxes",
    }
    return mapping.get(s, s)


def to_float(value):
    if pd.isna(value):
        return None
    s = str(value).strip().replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", s)
    if not match:
        return None
    try:
        return float(match.group())
    except Exception:
        return None


def most_common_nonempty(values, default=""):
    vals = [v for v in values if v]
    if not vals:
        return default
    return Counter(vals).most_common(1)[0][0]


def set_to_text(values, limit=6):
    vals = sorted(v for v in set(values) if v)
    if not vals:
        return ""
    shown = vals[:limit]
    suffix = "" if len(vals) <= limit else f" (+{len(vals)-limit} more)"
    return " | ".join(shown) + suffix


def year_overlap_score(a_min, a_max, b_min, b_max):
    if pd.isna(a_min) or pd.isna(a_max) or pd.isna(b_min) or pd.isna(b_max):
        return 0.5
    if max(a_min, b_min) <= min(a_max, b_max):
        return 1.0
    gap = min(abs(a_min - b_max), abs(b_min - a_max))
    if gap <= 1:
        return 0.8
    if gap <= 3:
        return 0.5
    if gap <= 8:
        return 0.2
    return 0.0


def overlap_score(a_set, b_set, missing_default=0.5):
    if not a_set and not b_set:
        return missing_default
    if not a_set or not b_set:
        return missing_default
    return 1.0 if set(a_set) & set(b_set) else 0.0


def tons_similarity(a_value, b_value):
    if a_value is None or b_value is None:
        return 0.5
    max_val = max(abs(a_value), abs(b_value), 1.0)
    rel_diff = abs(a_value - b_value) / max_val
    if rel_diff <= 0.05:
        return 1.0
    if rel_diff <= 0.15:
        return 0.8
    if rel_diff <= 0.30:
        return 0.5
    if rel_diff <= 0.50:
        return 0.2
    return 0.0


def build_reason_list(name_score, same_clean, year_score, unit_score, type_score, cargo_amount_score):
    reasons = []
    if same_clean:
        reasons.append("same cleaned name")
    if name_score >= 0.95:
        reasons.append("very high raw-name similarity")
    elif name_score >= 0.88:
        reasons.append("high raw-name similarity")
    if unit_score == 1.0:
        reasons.append("shared cargo unit")
    if type_score == 1.0:
        reasons.append("shared vessel type")
    if cargo_amount_score >= 0.8:
        reasons.append("similar typical cargo amount")
    if year_score >= 0.8:
        reasons.append("overlapping/near years")
    if not reasons:
        reasons.append("flagged by fuzzy candidate generation")
    return reasons


def choose_canonical_name(stats_df, name_list):
    subset = stats_df[stats_df["raw_name"].isin(name_list)].copy()
    subset = subset.sort_values(["row_count", "raw_name"], ascending=[False, True])
    if subset.empty:
        return sorted(name_list, key=lambda x: (len(x), x.lower()))[0]
    max_rows = subset["row_count"].max()
    candidates = subset[subset["row_count"] == max_rows]["raw_name"].tolist()
    return sorted(candidates, key=lambda x: (len(x), x.lower()))[0]


def preprocess_rows(df: pd.DataFrame, column_config: dict) -> pd.DataFrame:
    out = df.copy()
    entity_col = column_config["entity_column"]
    out["raw_name"] = out.get(entity_col, pd.Series("", index=out.index)).fillna("").astype(str).map(lambda x: x.strip())
    out = out[out["raw_name"] != ""].copy()
    out["clean_name"] = out["raw_name"].map(clean_name)
    out["strict_name_key"] = out["raw_name"].map(strict_name_key)
    type_col = column_config.get("type_column")
    unit_col = column_config.get("unit_column")
    amount_col = column_config.get("amount_column")
    year_col = column_config.get("year_column")
    notes_col_1 = column_config.get("notes_column_1")
    notes_col_2 = column_config.get("notes_column_2")
    out["vessel_type_clean"] = out.get(type_col, pd.Series("", index=out.index)).map(clean_vessel_type) if type_col else ""
    out["unit_primary_clean"] = out.get(unit_col, pd.Series("", index=out.index)).map(clean_unit) if unit_col else ""
    out["amount_primary_num"] = out.get(amount_col, pd.Series(None, index=out.index)).map(to_float) if amount_col else None
    out["year_num"] = pd.to_numeric(out.get(year_col, pd.Series(None, index=out.index)), errors="coerce") if year_col else None
    out["notes_combined"] = (
        out.get(notes_col_1, pd.Series("", index=out.index)).fillna("").astype(str).str.strip()
        + " || "
        + out.get(notes_col_2, pd.Series("", index=out.index)).fillna("").astype(str).str.strip()
    ).str.strip(" |")
    return out


def build_name_stats(rows: pd.DataFrame) -> pd.DataFrame:
    records = []
    for raw_name, grp in rows.groupby("raw_name", sort=False):
        years = grp["year_num"].dropna()
        tons_values = grp.loc[grp["unit_primary_clean"] == "tons", "amount_primary_num"].dropna().tolist()
        records.append(
            {
                "raw_name": raw_name,
                "display_name": raw_name,
                "clean_name": most_common_nonempty(grp["clean_name"].tolist(), default=clean_name(raw_name)),
                "strict_name_key": most_common_nonempty(grp["strict_name_key"].tolist(), default=strict_name_key(raw_name)),
                "row_count": int(len(grp)),
                "min_year": int(years.min()) if len(years) else None,
                "max_year": int(years.max()) if len(years) else None,
                "vessel_types": tuple(sorted(v for v in set(grp["vessel_type_clean"]) if v)),
                "units": tuple(sorted(v for v in set(grp["unit_primary_clean"]) if v)),
                "median_tons": float(pd.Series(tons_values).median()) if tons_values else None,
                "sample_notes": " | ".join([n for n in grp["notes_combined"].dropna().astype(str).tolist() if n][:3]),
            }
        )
    stats = pd.DataFrame(records)
    stats["prefix_block"] = stats["clean_name"].map(lambda s: s[:4] if s else "")
    return stats


def build_safe_auto_groups(stats: pd.DataFrame, entity_column: str) -> pd.DataFrame:
    groups = []
    for strict_key, grp in stats.groupby("strict_name_key", sort=False):
        if not strict_key or len(grp) < 2:
            continue
        names = grp["raw_name"].tolist()
        canonical = choose_canonical_name(stats, names)
        reasons = ["same strict key after removing punctuation/case/spaces"]
        if len(set(grp["clean_name"])) == 1:
            reasons.append("same cleaned name")
        years_min = grp["min_year"].dropna()
        years_max = grp["max_year"].dropna()
        groups.append(
            {
                "auto_group_id": f"A{len(groups)+1:04d}",
                "auto_group_key": make_auto_group_key(strict_key, entity_column),
                "strict_name_key": strict_key,
                "canonical_name": canonical,
                "member_count": len(grp),
                "member_names": " | ".join(sorted(names)),
                "members_list": sorted(names),
                "reasons": " | ".join(reasons),
                "confidence": "safe_auto",
                "total_rows": int(grp["row_count"].sum()),
                "min_year": int(years_min.min()) if len(years_min) else None,
                "max_year": int(years_max.max()) if len(years_max) else None,
                "units": set_to_text(itertools.chain.from_iterable(grp["units"].tolist())),
                "vessel_types": set_to_text(itertools.chain.from_iterable(grp["vessel_types"].tolist())),
            }
        )
    out = pd.DataFrame(groups)
    if not out.empty:
        out = out.sort_values(["member_count", "total_rows", "canonical_name"], ascending=[False, False, True]).reset_index(drop=True)
    return out


def generate_candidate_pairs(stats: pd.DataFrame, entity_column: str, resolved_names=None, fuzzy_threshold: int = 88) -> pd.DataFrame:
    if resolved_names is None:
        resolved_names = set()
    candidate_stats = stats[~stats["raw_name"].isin(resolved_names)].copy()

    by_name = {row["raw_name"]: row for _, row in candidate_stats.iterrows()}
    pairs = {}
    candidate_records = []

    def add_pair(name_a, name_b, source):
        if name_a == name_b:
            return
        if name_a not in by_name or name_b not in by_name:
            return
        key = tuple(sorted((name_a, name_b)))
        if key in pairs:
            pairs[key].add(source)
            return
        pairs[key] = {source}

    for _, grp in candidate_stats.groupby("clean_name"):
        names = grp["raw_name"].tolist()
        if len(names) < 2:
            continue
        for a, b in itertools.combinations(names, 2):
            add_pair(a, b, "same_clean_name")

    for _, grp in candidate_stats.groupby("prefix_block"):
        names = grp["raw_name"].tolist()
        if len(names) < 2:
            continue
        if len(names) > 60:
            grp = grp.sort_values(["row_count", "raw_name"], ascending=[False, True]).head(60)
            names = grp["raw_name"].tolist()
        for a, b in itertools.combinations(names, 2):
            row_a = by_name[a]
            row_b = by_name[b]
            raw_score = fuzz.WRatio(row_a["raw_name"], row_b["raw_name"]) / 100.0
            clean_score = fuzz.WRatio(row_a["clean_name"], row_b["clean_name"]) / 100.0
            if max(raw_score, clean_score) >= fuzzy_threshold / 100.0:
                add_pair(a, b, "fuzzy_prefix_block")

    for (a, b), sources in pairs.items():
        row_a = by_name[a]
        row_b = by_name[b]

        raw_score = fuzz.WRatio(row_a["raw_name"], row_b["raw_name"]) / 100.0
        clean_score = fuzz.WRatio(row_a["clean_name"], row_b["clean_name"]) / 100.0
        same_clean = row_a["clean_name"] == row_b["clean_name"]

        year_score = year_overlap_score(row_a["min_year"], row_a["max_year"], row_b["min_year"], row_b["max_year"])
        unit_score = overlap_score(row_a["units"], row_b["units"])
        type_score = overlap_score(row_a["vessel_types"], row_b["vessel_types"])
        cargo_amount_score = tons_similarity(row_a["median_tons"], row_b["median_tons"])

        final_score = (
            0.45 * max(raw_score, clean_score)
            + 0.20 * (1.0 if same_clean else clean_score)
            + 0.12 * unit_score
            + 0.10 * year_score
            + 0.08 * type_score
            + 0.05 * cargo_amount_score
        )
        pair_key = make_pair_key(row_a["display_name"], row_b["display_name"], entity_column)

        candidate_records.append(
            {
                "pair_key": pair_key,
                "candidate_id": pair_key,
                "name_a": row_a["display_name"],
                "name_b": row_b["display_name"],
                "score": round(final_score, 4),
                "raw_name_score": round(raw_score, 4),
                "clean_name_score": round(clean_score, 4),
                "year_score": round(year_score, 4),
                "unit_score": round(unit_score, 4),
                "type_score": round(type_score, 4),
                "cargo_amount_score": round(cargo_amount_score, 4),
                "reasons": " | ".join(build_reason_list(max(raw_score, clean_score), same_clean, year_score, unit_score, type_score, cargo_amount_score)),
                "suggested_canonical": choose_canonical_name(candidate_stats, [row_a["display_name"], row_b["display_name"]]),
            }
        )

    queue = pd.DataFrame(candidate_records)
    if not queue.empty:
        queue = queue.sort_values(["score"], ascending=[False]).reset_index(drop=True)
    return queue


class UnionFind:
    def __init__(self):
        self.parent = {}

    def find(self, x):
        if x not in self.parent:
            self.parent[x] = x
            return x
        if self.parent[x] != x:
            self.parent[x] = self.find(self.parent[x])
        return self.parent[x]

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


def build_merge_outputs(stats_df, auto_groups_df, auto_status, manual_decisions):
    uf = UnionFind()
    history_records = []

    if not auto_groups_df.empty:
        for _, row in auto_groups_df.iterrows():
            gid = row["auto_group_key"]
            if auto_status.get(gid) == "accepted":
                members = row["members_list"]
                for m in members[1:]:
                    uf.union(members[0], m)
                history_records.append(
                    {
                        "merge_source": "auto_group",
                        "merge_id": gid,
                        "canonical_name": row["canonical_name"],
                        "members": " | ".join(members),
                        "reason": row["reasons"],
                        "status": "active",
                    }
                )

    if manual_decisions:
        for pair_key, row in manual_decisions.items():
            if row.get("decision") == "merge":
                uf.union(row["name_a"], row["name_b"])
                history_records.append(
                    {
                        "merge_source": "manual_pair",
                        "merge_id": pair_key,
                        "canonical_name": row["suggested_canonical"],
                        "members": f"{row['name_a']} | {row['name_b']}",
                        "reason": row["reasons"],
                        "status": "active",
                    }
                )

    all_names = set()
    for rec in history_records:
        parts = [p.strip() for p in rec["members"].split("|")]
        all_names.update([p for p in parts if p])

    groups = defaultdict(list)
    for name in all_names:
        groups[uf.find(name)].append(name)

    row_count_map = dict(zip(stats_df["raw_name"], stats_df["row_count"]))
    clean_map = dict(zip(stats_df["raw_name"], stats_df["clean_name"]))

    mapping_records = []
    for idx, members in enumerate(sorted(groups.values(), key=lambda g: (-len(g), sorted(g)[0])), start=1):
        canonical = sorted(members, key=lambda n: (-row_count_map.get(n, 0), len(n), n.lower()))[0]
        cluster_id = f"M{idx:04d}"
        for member in sorted(members):
            mapping_records.append(
                {
                    "cluster_id": cluster_id,
                    "canonical_name": canonical,
                    "member_name": member,
                    "row_count": row_count_map.get(member, 0),
                    "clean_name": clean_map.get(member, ""),
                }
            )

    return pd.DataFrame(history_records), pd.DataFrame(mapping_records)


def resolved_names_from_auto(auto_groups_df, auto_status):
    resolved = set()
    if auto_groups_df.empty:
        return resolved
    for _, row in auto_groups_df.iterrows():
        if auto_status.get(row["auto_group_key"]) == "accepted":
            resolved.update(row["members_list"])
    return resolved

def active_manual_decisions_for_config(manual_decisions: dict, column_config: dict) -> dict:
    active_entity = column_config["entity_column"]
    return {
        key: record
        for key, record in manual_decisions.items()
        if record.get("entity_column") == active_entity
    }


def init_state():
    defaults = {
        "auto_status": {},
        "manual_decisions": {},
        "auto_index": 0,
        "pair_index": 0,
        "column_config": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if "pair_decisions" in st.session_state and st.session_state["pair_decisions"] and not st.session_state["manual_decisions"]:
        st.session_state["manual_decisions"] = {}


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


def make_download_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def build_cleaned_workbook_bytes(
    raw_df: pd.DataFrame,
    mapping_df: pd.DataFrame,
    history_df: pd.DataFrame,
    manual_decisions: dict,
    auto_groups_df: pd.DataFrame,
    column_config: dict,
    sheet_name: str,
    workbook_bytes: bytes | None = None,
    candidate_queue_df: pd.DataFrame | None = None,
) -> bytes:
    entity_col = column_config["entity_column"]
    cleaned_df = raw_df.copy()
    original_values = cleaned_df[entity_col].fillna("").astype(str)

    canonical_map = {}
    cluster_map = {}
    merged_values = set()
    if not mapping_df.empty:
        canonical_map = dict(zip(mapping_df["member_name"], mapping_df["canonical_name"]))
        cluster_map = dict(zip(mapping_df["member_name"], mapping_df["cluster_id"]))
        merged_values = set(mapping_df["member_name"].tolist())

    status_map = {}
    source_map = {}
    score_map = {}
    reason_map = {}

    if not auto_groups_df.empty:
        accepted = set(history_df.loc[history_df["merge_source"] == "auto_group", "merge_id"].tolist()) if not history_df.empty else set()
        for _, row in auto_groups_df.iterrows():
            if row["auto_group_key"] in accepted:
                for member in row["members_list"]:
                    status_map[member] = "merged"
                    source_map[member] = "auto_group"
                    reason_map[member] = row.get("reasons", "")

    for rec in manual_decisions.values():
        names = [rec.get("name_a", ""), rec.get("name_b", "")]
        decision = rec.get("decision", "")
        for name in names:
            if name in merged_values or decision == "merge":
                status_map[name] = "merged"
                source_map[name] = "manual_pair"
            elif decision == "keep_separate":
                status_map[name] = "reviewed_keep_separate"
                source_map[name] = "manual_keep_separate"
            elif decision == "unsure":
                status_map[name] = "reviewed_unsure"
                source_map[name] = "manual_unsure"
            score_map[name] = rec.get("score", "")
            reason_map[name] = rec.get("reasons", "")

    unreviewed_names = set()
    if candidate_queue_df is not None and not candidate_queue_df.empty:
        for _, row in candidate_queue_df.iterrows():
            unreviewed_names.add(row["name_a"])
            unreviewed_names.add(row["name_b"])

    cleaned_df["dedupe_entity_column"] = entity_col
    cleaned_df["dedupe_original_value"] = original_values
    cleaned_df["dedupe_canonical_value"] = original_values.map(lambda x: canonical_map.get(x, x))
    cleaned_df["dedupe_cluster_id"] = original_values.map(lambda x: cluster_map.get(x, ""))
    cleaned_df["dedupe_review_status"] = original_values.map(lambda x: status_map.get(x, ""))
    cleaned_df["dedupe_decision_source"] = original_values.map(lambda x: source_map.get(x, ""))
    cleaned_df["dedupe_score"] = original_values.map(lambda x: score_map.get(x, ""))
    cleaned_df["dedupe_reason"] = original_values.map(lambda x: reason_map.get(x, ""))

    def default_status(v: str) -> str:
        if v in merged_values:
            return "merged"
        if v in unreviewed_names:
            return "unreviewed"
        return "not_flagged"

    missing_status = cleaned_df["dedupe_review_status"] == ""
    cleaned_df.loc[missing_status, "dedupe_review_status"] = cleaned_df.loc[missing_status, "dedupe_original_value"].map(default_status)

    if workbook_bytes:
        all_sheets = pd.read_excel(io.BytesIO(workbook_bytes), sheet_name=None)
        all_sheets[sheet_name] = cleaned_df
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
        return out.getvalue()

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
    return out.getvalue()


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def build_manual_decision_record(row: pd.Series, decision: str, entity_column: str):
    return {
        "pair_key": row["pair_key"],
        "entity_column": entity_column,
        "name_a": row["name_a"],
        "name_b": row["name_b"],
        "decision": decision,
        "score": float(row["score"]),
        "raw_name_score": float(row["raw_name_score"]),
        "clean_name_score": float(row["clean_name_score"]),
        "year_score": float(row["year_score"]),
        "unit_score": float(row["unit_score"]),
        "type_score": float(row["type_score"]),
        "cargo_amount_score": float(row["cargo_amount_score"]),
        "reasons": row["reasons"],
        "suggested_canonical": row["suggested_canonical"],
        "updated_at": now_iso(),
    }


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
    st.caption("Safe auto-merges + manual review + merge history + undo.")

    with st.sidebar:
        st.header("Setup")
        uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx", "xls"])
        fuzzy_threshold = st.slider("Manual review fuzzy threshold", 80, 98, 88, 1)
        min_manual_score = st.slider("Minimum manual-review score", 0.0, 1.0, 0.75, 0.01)
        sample_rows = st.slider("Sample original rows per side", 3, 12, 5, 1)
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

    raw_df_preview = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, nrows=50)
    available_cols = raw_df_preview.columns.tolist()
    text_candidates = [c for c in available_cols if raw_df_preview[c].dtype == "object"]
    default_entity = SHIP_DEFAULTS["entity_column"] if SHIP_DEFAULTS["entity_column"] in available_cols else (text_candidates[0] if text_candidates else available_cols[0])

    saved_config = st.session_state.get("column_config", {}) or {}
    with st.sidebar:
        st.header("Column mapping")
        entity_default = saved_config.get("entity_column") if saved_config.get("entity_column") in available_cols else default_entity
        entity_column = st.selectbox("Primary column to deduplicate/review", available_cols, index=available_cols.index(entity_default), key="entity_column_select")
        optional_options = ["(None)"] + available_cols
        def optional_select(label, cfg_key, widget_key):
            from_saved = saved_config.get(cfg_key)
            if from_saved in available_cols:
                default_val = from_saved
            else:
                default_val = SHIP_DEFAULTS[cfg_key] if SHIP_DEFAULTS[cfg_key] in available_cols else "(None)"
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

    if pd.api.types.is_numeric_dtype(raw_df[column_config["entity_column"]]):
        st.warning("This column looks numeric. The tool works best for text/entity columns such as names, places, categories, or labels.")

    resolved_names = resolved_names_from_auto(auto_groups_df, st.session_state["auto_status"])
    queue_df = generate_candidate_pairs(stats_df, entity_column=column_config["entity_column"], resolved_names=resolved_names, fuzzy_threshold=fuzzy_threshold)
    if not queue_df.empty:
        queue_df = queue_df[queue_df["score"] >= min_manual_score].reset_index(drop=True)

    active_manual_decisions = active_manual_decisions_for_config(st.session_state["manual_decisions"], column_config)
    hidden_decision_count = len(st.session_state["manual_decisions"]) - len(active_manual_decisions)
    if hidden_decision_count > 0:
        st.info("Some saved decisions belong to a different primary column and are hidden from the current mapping.")
    history_df, mapping_df = build_merge_outputs(stats_df, auto_groups_df, st.session_state["auto_status"], active_manual_decisions)

    stat_lookup = stats_df.set_index("raw_name").to_dict("index")

    metrics = st.columns(5)
    metrics[0].metric("Unique raw primary values", len(stats_df))
    metrics[1].metric("Safe auto-groups", len(auto_groups_df))
    metrics[2].metric("Accepted auto-groups", sum(1 for v in st.session_state["auto_status"].values() if v == "accepted"))
    metrics[3].metric("Manual queue", len(queue_df) if queue_df is not None else 0)
    metrics[4].metric("Merged names now", len(mapping_df))
    st.caption(f"Manual decisions in current mapping: {len(active_manual_decisions)} (total saved: {len(st.session_state['manual_decisions'])}).")

    tabs = st.tabs(["1) Safe Auto-Merges", "2) Manual Review Queue", "3) Merge History + Undo", "4) Export"])

    with tabs[0]:
        st.subheader("Safe auto-merges")
        st.write("These are the most conservative merges. They trigger only when names share the same strict key after removing punctuation, case, and spaces.")

        if auto_groups_df.empty:
            st.info("No safe auto-groups found.")
        else:
            actions = st.columns(3)
            if actions[0].button("Accept all safe auto-merges", use_container_width=True):
                for gid in auto_groups_df["auto_group_key"].tolist():
                    st.session_state["auto_status"][gid] = "accepted"
                st.rerun()
            if actions[1].button("Reject all safe auto-merges", use_container_width=True):
                for gid in auto_groups_df["auto_group_key"].tolist():
                    st.session_state["auto_status"][gid] = "rejected"
                st.rerun()
            if actions[2].button("Clear all auto decisions", use_container_width=True):
                for gid in auto_groups_df["auto_group_key"].tolist():
                    st.session_state["auto_status"].pop(gid, None)
                st.rerun()

            idx = min(st.session_state["auto_index"], len(auto_groups_df) - 1)
            st.session_state["auto_index"] = idx
            row = auto_groups_df.iloc[idx]
            status = st.session_state["auto_status"].get(row["auto_group_key"], "pending")

            info = st.columns([2, 1.5, 1.5, 2])
            info[0].markdown(f"### Group {idx + 1} / {len(auto_groups_df)}")
            info[1].markdown(f"**Status:** {status}")
            info[2].markdown(f"**Canonical:** {row['canonical_name']}")
            info[3].markdown(f"**Reason:** {row['reasons']}")

            nav = st.columns(5)
            if nav[0].button("⬅️ Previous group", use_container_width=True):
                st.session_state["auto_index"] = max(st.session_state["auto_index"] - 1, 0)
                st.rerun()
            if nav[1].button("✅ Accept auto-merge", use_container_width=True):
                st.session_state["auto_status"][row["auto_group_key"]] = "accepted"
                st.session_state["auto_index"] = min(st.session_state["auto_index"] + 1, len(auto_groups_df) - 1)
                st.rerun()
            if nav[2].button("❌ Reject auto-merge", use_container_width=True):
                st.session_state["auto_status"][row["auto_group_key"]] = "rejected"
                st.session_state["auto_index"] = min(st.session_state["auto_index"] + 1, len(auto_groups_df) - 1)
                st.rerun()
            if nav[3].button("↩️ Undo this auto decision", use_container_width=True):
                st.session_state["auto_status"].pop(row["auto_group_key"], None)
                st.rerun()
            if nav[4].button("➡️ Next group", use_container_width=True):
                st.session_state["auto_index"] = min(st.session_state["auto_index"] + 1, len(auto_groups_df) - 1)
                st.rerun()

            detail = st.columns(4)
            detail[0].metric("Member names", row["member_count"])
            detail[1].metric("Total rows", row["total_rows"])
            detail[2].metric("Min year", row["min_year"] if pd.notna(row["min_year"]) else "—")
            detail[3].metric("Max year", row["max_year"] if pd.notna(row["max_year"]) else "—")

            st.markdown(f"**Members:** {row['member_names']}")
            st.markdown(f"**Units:** {row['units'] or '—'}")
            st.markdown(f"**Vessel types:** {row['vessel_types'] or '—'}")

            preview = auto_groups_df[["auto_group_id", "auto_group_key", "canonical_name", "member_count", "total_rows", "reasons"]].copy()
            preview["status"] = preview["auto_group_key"].map(lambda gid: st.session_state["auto_status"].get(gid, "pending"))
            st.dataframe(preview, use_container_width=True, hide_index=True)

    with tabs[1]:
        st.subheader("Manual review queue")
        st.write("These are the remaining ambiguous candidates after any accepted safe auto-merges are removed from review.")

        if queue_df is None or queue_df.empty:
            st.success("No manual-review candidates remain under the current settings.")
        else:
            idx = min(st.session_state["pair_index"], len(queue_df) - 1)
            st.session_state["pair_index"] = idx
            row = queue_df.iloc[idx]
            decision = active_manual_decisions.get(row["pair_key"], {}).get("decision", "unreviewed")

            info = st.columns([2, 1.5, 1.5, 2])
            info[0].markdown(f"### Candidate {idx + 1} / {len(queue_df)}")
            info[1].markdown(f"**Score:** `{row['score']:.3f}`")
            info[2].markdown(f"**Decision:** {decision}")
            info[3].markdown(f"**Suggested canonical:** {row['suggested_canonical']}")

            nav = st.columns(6)
            if nav[0].button("⬅️ Previous pair", use_container_width=True):
                st.session_state["pair_index"] = max(st.session_state["pair_index"] - 1, 0)
                st.rerun()
            if nav[1].button("✅ Merge", use_container_width=True):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(row, "merge", column_config["entity_column"])
                st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(queue_df) - 1)
                st.rerun()
            if nav[2].button("❌ Keep separate", use_container_width=True):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(row, "keep_separate", column_config["entity_column"])
                st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(queue_df) - 1)
                st.rerun()
            if nav[3].button("🤔 Unsure", use_container_width=True):
                st.session_state["manual_decisions"][row["pair_key"]] = build_manual_decision_record(row, "unsure", column_config["entity_column"])
                st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(queue_df) - 1)
                st.rerun()
            if nav[4].button("↩️ Undo this pair decision", use_container_width=True):
                st.session_state["manual_decisions"].pop(row["pair_key"], None)
                st.rerun()
            if nav[5].button("➡️ Next pair", use_container_width=True):
                st.session_state["pair_index"] = min(st.session_state["pair_index"] + 1, len(queue_df) - 1)
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
            left_rows = raw_df[raw_df[column_config["entity_column"]].fillna("").astype(str).str.strip() == row["name_a"]][display_columns].head(sample_rows)
            right_rows = raw_df[raw_df[column_config["entity_column"]].fillna("").astype(str).str.strip() == row["name_b"]][display_columns].head(sample_rows)

            previews = st.columns(2)
            with previews[0]:
                st.markdown("#### Original rows for Side A")
                st.dataframe(left_rows, use_container_width=True, hide_index=True)
            with previews[1]:
                st.markdown("#### Original rows for Side B")
                st.dataframe(right_rows, use_container_width=True, hide_index=True)

    with tabs[2]:
        st.subheader("Merge history and undo")
        st.write("Every active merge appears here with the reason it was merged. You can undo anything from this list.")

        if history_df.empty:
            st.info("No active merges yet.")
        else:
            st.dataframe(history_df, use_container_width=True, hide_index=True)
            selected_merge = st.selectbox("Select a merge to undo", history_df["merge_id"].tolist())
            if st.button("Undo selected merge", use_container_width=True):
                selected = history_df[history_df["merge_id"] == selected_merge].iloc[0]
                if selected["merge_source"] == "auto_group":
                    st.session_state["auto_status"].pop(selected_merge, None)
                else:
                    st.session_state["manual_decisions"].pop(selected_merge, None)
                st.rerun()

            st.markdown("#### Current canonical mapping")
            st.dataframe(mapping_df, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.subheader("Export")
        st.info("This exports a reviewed copy. The original uploaded workbook is not overwritten.")
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
        if st.button("Load review session", use_container_width=False):
            if uploaded_session is None:
                st.warning("Please upload a review session JSON first.")
            else:
                data = json.loads(uploaded_session.getvalue().decode("utf-8"))
                saved_cfg = data.get("column_config", {})
                missing_cols = [v for v in saved_cfg.values() if v and v not in raw_df.columns]
                if missing_cols:
                    st.warning(f"Some saved columns are missing in the current sheet: {', '.join(sorted(set(missing_cols)))}. Using current mapping where needed.")
                else:
                    st.success("Loaded session column mapping.")
                st.session_state["column_config"] = saved_cfg
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
                    if cfg_key == "entity_column":
                        fallback = default_entity
                    else:
                        fallback = "(None)"
                    if value in available_cols:
                        st.session_state[widget_key] = value
                    else:
                        st.session_state[widget_key] = fallback
                st.session_state["auto_status"] = data.get("auto_status", {})
                st.session_state["manual_decisions"] = data.get("manual_decisions", {})
                st.session_state["auto_index"] = 0
                st.session_state["pair_index"] = 0
                st.success("Review session loaded.")
                st.rerun()

        buttons = st.columns(5)
        buttons[0].download_button("Download review session JSON", session_bytes, "ship_review_session.json", "application/json", use_container_width=True)
        buttons[1].download_button("Download auto-merge decisions CSV", make_download_bytes(auto_export if not auto_export.empty else pd.DataFrame()), "ship_auto_merge_decisions.csv", "text/csv", use_container_width=True)
        buttons[2].download_button("Download manual review decisions CSV", make_download_bytes(pair_export if not pair_export.empty else pd.DataFrame()), "ship_manual_review_decisions.csv", "text/csv", use_container_width=True)
        buttons[3].download_button("Download merge history CSV", make_download_bytes(history_df if not history_df.empty else pd.DataFrame()), "ship_merge_history.csv", "text/csv", use_container_width=True)
        buttons[4].download_button("Download canonical mapping CSV", make_download_bytes(mapping_df if not mapping_df.empty else pd.DataFrame()), "ship_canonical_mapping.csv", "text/csv", use_container_width=True)

        base_name = getattr(uploaded, "name", "") or "cleaned_duplicate_review.xlsx"
        if base_name.lower().endswith(".xlsx"):
            cleaned_name = f"{base_name[:-5]}_reviewed.xlsx"
        else:
            cleaned_name = "cleaned_duplicate_review.xlsx"
        cleaned_bytes = build_cleaned_workbook_bytes(
            raw_df=raw_df,
            mapping_df=mapping_df,
            history_df=history_df,
            manual_decisions=active_manual_decisions,
            auto_groups_df=auto_groups_df,
            column_config=column_config,
            sheet_name=sheet_name,
            workbook_bytes=file_bytes,
            candidate_queue_df=queue_df,
        )
        st.download_button(
            "Download cleaned workbook",
            cleaned_bytes,
            cleaned_name,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


if __name__ == "__main__":
    app()
