import io
import itertools
import hashlib
import re
from copy import copy
from math import ceil
from collections import Counter, defaultdict
from datetime import datetime, timezone

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import TableColumn
from rapidfuzz import fuzz



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
    return re.sub('\\s+', ' ', text).strip()

def clean_name(value: str) -> str:
    if pd.isna(value):
        return ''
    s = str(value).strip().lower()
    s = s.replace('&', ' and ')
    s = s.replace('?', ' ')
    s = re.sub('[\\.,;:\'\\"`´‘’“”()\\[\\]{}_/\\\\|-]+', ' ', s)
    s = normalize_spaces(s)
    return s

def stable_hash(text: str, length: int=12) -> str:
    return hashlib.sha1(text.encode('utf-8')).hexdigest()[:length]

def make_pair_key(name_a: str, name_b: str, entity_column: str='Name of Vessel') -> str:
    normalized_names = sorted([clean_name(name_a), clean_name(name_b)])
    raw_key = entity_column + '::' + '||'.join(normalized_names)
    return 'P' + stable_hash(raw_key)

def make_auto_group_key(strict_key: str, entity_column: str) -> str:
    return 'A' + stable_hash(f'auto::{entity_column}::{strict_key}')

def strict_name_key(value: str) -> str:
    if pd.isna(value):
        return ''
    s = str(value).strip().lower()
    s = re.sub('[^a-z0-9]+', '', s)
    return s

def clean_vessel_type(value: str) -> str:
    if pd.isna(value):
        return ''
    s = str(value).strip().lower()
    s = re.sub('[\\.,;:\'\\"`´‘’“”()\\[\\]{}_/\\\\|-]+', ' ', s)
    s = normalize_spaces(s)
    mapping = {'prop': 'propeller', 'ss': 'steamship', 's s': 'steamship', 'stmr': 'steamer', 'stm': 'steamer', 'str': 'steamer', 'sch': 'schooner', 'schr': 'schooner', 'shnr': 'schooner', 'sb': 'sailboat', 's b': 'sailboat'}
    return mapping.get(s, s)

def clean_unit(value: str) -> str:
    if pd.isna(value):
        return ''
    s = str(value).strip().lower()
    s = re.sub('[\\.,;:\'\\"`´‘’“”()\\[\\]{}_/\\\\|-]+', ' ', s)
    s = normalize_spaces(s)
    mapping = {'ton': 'tons', 'tons': 'tons', 'tns': 'tons', 'toise': 'toise', 'barrels': 'barrels', 'barrel': 'barrels', 'baskets': 'baskets', 'crates': 'crates', 'boxes': 'boxes'}
    return mapping.get(s, s)

def to_float(value):
    if pd.isna(value):
        return None
    s = str(value).strip().replace(',', '')
    match = re.search('-?\\d+(?:\\.\\d+)?', s)
    if not match:
        return None
    try:
        return float(match.group())
    except Exception:
        return None

def most_common_nonempty(values, default=''):
    vals = [v for v in values if v]
    if not vals:
        return default
    return Counter(vals).most_common(1)[0][0]

def set_to_text(values, limit=6):
    vals = sorted((v for v in set(values) if v))
    if not vals:
        return ''
    shown = vals[:limit]
    suffix = '' if len(vals) <= limit else f' (+{len(vals) - limit} more)'
    return ' | '.join(shown) + suffix

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
    if not a_set and (not b_set):
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
    if rel_diff <= 0.3:
        return 0.5
    if rel_diff <= 0.5:
        return 0.2
    return 0.0

def build_reason_list(name_score, same_clean, year_score, unit_score, type_score, cargo_amount_score):
    reasons = []
    if same_clean:
        reasons.append('same cleaned name')
    if name_score >= 0.95:
        reasons.append('very high raw-name similarity')
    elif name_score >= 0.88:
        reasons.append('high raw-name similarity')
    if unit_score == 1.0:
        reasons.append('shared cargo unit')
    if type_score == 1.0:
        reasons.append('shared vessel type')
    if cargo_amount_score >= 0.8:
        reasons.append('similar typical cargo amount')
    if year_score >= 0.8:
        reasons.append('overlapping/near years')
    if not reasons:
        reasons.append('flagged by fuzzy candidate generation')
    return reasons

def canonical_sort_key(name: str):
    whitespace_normalized = normalize_spaces(str(name))
    compact_alnum = re.sub('[^A-Za-z0-9]', '', whitespace_normalized)
    letters_only = re.sub('[^A-Za-z]', '', whitespace_normalized)
    all_caps = int(bool(letters_only) and letters_only.isupper())
    punctuation_count = len(re.findall('[^\\w\\s]', whitespace_normalized))
    cleaned_for_length = re.sub('[^A-Za-z0-9\\s]', '', whitespace_normalized)
    cleaned_for_length = normalize_spaces(cleaned_for_length)
    return (
        all_caps,
        punctuation_count,
        len(cleaned_for_length),
        whitespace_normalized.lower(),
        compact_alnum.lower(),
    )

def choose_canonical_name(stats_df, name_list):
    subset = stats_df[stats_df['raw_name'].isin(name_list)].copy()
    subset = subset.sort_values(['row_count', 'raw_name'], ascending=[False, True])
    if subset.empty:
        return sorted(name_list, key=canonical_sort_key)[0]
    max_rows = subset['row_count'].max()
    candidates = subset[subset['row_count'] == max_rows]['raw_name'].tolist()
    return sorted(candidates, key=canonical_sort_key)[0]

def preprocess_rows(df: pd.DataFrame, column_config: dict) -> pd.DataFrame:
    out = df.copy()
    entity_col = column_config['entity_column']
    out['raw_name'] = out.get(entity_col, pd.Series('', index=out.index)).fillna('').astype(str).map(lambda x: x.strip())
    out = out[out['raw_name'] != ''].copy()
    out['clean_name'] = out['raw_name'].map(clean_name)
    out['strict_name_key'] = out['raw_name'].map(strict_name_key)
    type_col = column_config.get('type_column')
    unit_col = column_config.get('unit_column')
    amount_col = column_config.get('amount_column')
    year_col = column_config.get('year_column')
    notes_col_1 = column_config.get('notes_column_1')
    notes_col_2 = column_config.get('notes_column_2')
    out['vessel_type_clean'] = out.get(type_col, pd.Series('', index=out.index)).map(clean_vessel_type) if type_col else ''
    out['unit_primary_clean'] = out.get(unit_col, pd.Series('', index=out.index)).map(clean_unit) if unit_col else ''
    out['amount_primary_num'] = out.get(amount_col, pd.Series(None, index=out.index)).map(to_float) if amount_col else None
    out['year_num'] = pd.to_numeric(out.get(year_col, pd.Series(None, index=out.index)), errors='coerce') if year_col else None
    out['notes_combined'] = (out.get(notes_col_1, pd.Series('', index=out.index)).fillna('').astype(str).str.strip() + ' || ' + out.get(notes_col_2, pd.Series('', index=out.index)).fillna('').astype(str).str.strip()).str.strip(' |')
    return out

def build_name_stats(rows: pd.DataFrame) -> pd.DataFrame:
    records = []
    for raw_name, grp in rows.groupby('raw_name', sort=False):
        years = grp['year_num'].dropna()
        tons_values = grp.loc[grp['unit_primary_clean'] == 'tons', 'amount_primary_num'].dropna().tolist()
        records.append({'raw_name': raw_name, 'display_name': raw_name, 'clean_name': most_common_nonempty(grp['clean_name'].tolist(), default=clean_name(raw_name)), 'strict_name_key': most_common_nonempty(grp['strict_name_key'].tolist(), default=strict_name_key(raw_name)), 'row_count': int(len(grp)), 'min_year': int(years.min()) if len(years) else None, 'max_year': int(years.max()) if len(years) else None, 'vessel_types': tuple(sorted((v for v in set(grp['vessel_type_clean']) if v))), 'units': tuple(sorted((v for v in set(grp['unit_primary_clean']) if v))), 'median_tons': float(pd.Series(tons_values).median()) if tons_values else None, 'sample_notes': ' | '.join([n for n in grp['notes_combined'].dropna().astype(str).tolist() if n][:3])})
    stats = pd.DataFrame(records)
    stats['prefix_block'] = stats['clean_name'].map(lambda s: s[:4] if s else '')
    return stats

def build_safe_auto_groups(stats: pd.DataFrame, entity_column: str) -> pd.DataFrame:
    groups = []
    for strict_key, grp in stats.groupby('strict_name_key', sort=False):
        if not strict_key or len(grp) < 2:
            continue
        names = grp['raw_name'].tolist()
        canonical = choose_canonical_name(stats, names)
        reasons = ['same strict key after removing punctuation/case/spaces']
        if len(set(grp['clean_name'])) == 1:
            reasons.append('same cleaned name')
        years_min = grp['min_year'].dropna()
        years_max = grp['max_year'].dropna()
        groups.append({'auto_group_id': f'A{len(groups) + 1:04d}', 'auto_group_key': make_auto_group_key(strict_key, entity_column), 'strict_name_key': strict_key, 'canonical_name': canonical, 'member_count': len(grp), 'member_names': ' | '.join(sorted(names)), 'members_list': sorted(names), 'reasons': ' | '.join(reasons), 'confidence': 'safe_auto', 'total_rows': int(grp['row_count'].sum()), 'min_year': int(years_min.min()) if len(years_min) else None, 'max_year': int(years_max.max()) if len(years_max) else None, 'units': set_to_text(itertools.chain.from_iterable(grp['units'].tolist())), 'vessel_types': set_to_text(itertools.chain.from_iterable(grp['vessel_types'].tolist()))})
    out = pd.DataFrame(groups)
    if not out.empty:
        out = out.sort_values(['member_count', 'total_rows', 'canonical_name'], ascending=[False, False, True]).reset_index(drop=True)
    return out

def generate_candidate_pairs(stats: pd.DataFrame, entity_column: str, resolved_names=None, fuzzy_threshold: int=88) -> pd.DataFrame:
    if resolved_names is None:
        resolved_names = set()
    candidate_stats = stats[~stats['raw_name'].isin(resolved_names)].copy()
    by_name = {row['raw_name']: row for _, row in candidate_stats.iterrows()}
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
    for _, grp in candidate_stats.groupby('clean_name'):
        names = grp['raw_name'].tolist()
        if len(names) < 2:
            continue
        for a, b in itertools.combinations(names, 2):
            add_pair(a, b, 'same_clean_name')
    for _, grp in candidate_stats.groupby('prefix_block'):
        names = grp['raw_name'].tolist()
        if len(names) < 2:
            continue
        if len(names) > 60:
            grp = grp.sort_values(['row_count', 'raw_name'], ascending=[False, True]).head(60)
            names = grp['raw_name'].tolist()
        for a, b in itertools.combinations(names, 2):
            row_a = by_name[a]
            row_b = by_name[b]
            raw_score = fuzz.WRatio(row_a['raw_name'], row_b['raw_name']) / 100.0
            clean_score = fuzz.WRatio(row_a['clean_name'], row_b['clean_name']) / 100.0
            if max(raw_score, clean_score) >= fuzzy_threshold / 100.0:
                add_pair(a, b, 'fuzzy_prefix_block')
    for (a, b), sources in pairs.items():
        row_a = by_name[a]
        row_b = by_name[b]
        raw_score = fuzz.WRatio(row_a['raw_name'], row_b['raw_name']) / 100.0
        clean_score = fuzz.WRatio(row_a['clean_name'], row_b['clean_name']) / 100.0
        same_clean = row_a['clean_name'] == row_b['clean_name']
        year_score = year_overlap_score(row_a['min_year'], row_a['max_year'], row_b['min_year'], row_b['max_year'])
        unit_score = overlap_score(row_a['units'], row_b['units'])
        type_score = overlap_score(row_a['vessel_types'], row_b['vessel_types'])
        cargo_amount_score = tons_similarity(row_a['median_tons'], row_b['median_tons'])
        final_score = 0.45 * max(raw_score, clean_score) + 0.2 * (1.0 if same_clean else clean_score) + 0.12 * unit_score + 0.1 * year_score + 0.08 * type_score + 0.05 * cargo_amount_score
        pair_key = make_pair_key(row_a['display_name'], row_b['display_name'], entity_column)
        candidate_records.append({'pair_key': pair_key, 'candidate_id': pair_key, 'name_a': row_a['display_name'], 'name_b': row_b['display_name'], 'score': round(final_score, 4), 'raw_name_score': round(raw_score, 4), 'clean_name_score': round(clean_score, 4), 'year_score': round(year_score, 4), 'unit_score': round(unit_score, 4), 'type_score': round(type_score, 4), 'cargo_amount_score': round(cargo_amount_score, 4), 'reasons': ' | '.join(build_reason_list(max(raw_score, clean_score), same_clean, year_score, unit_score, type_score, cargo_amount_score)), 'suggested_canonical': choose_canonical_name(candidate_stats, [row_a['display_name'], row_b['display_name']])})
    queue = pd.DataFrame(candidate_records)
    if not queue.empty:
        queue = queue.sort_values(['score'], ascending=[False]).reset_index(drop=True)
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
        ra, rb = (self.find(a), self.find(b))
        if ra != rb:
            self.parent[rb] = ra

def build_merge_outputs(stats_df, auto_groups_df, auto_status, manual_decisions):
    uf = UnionFind()
    history_records = []
    if not auto_groups_df.empty:
        for _, row in auto_groups_df.iterrows():
            gid = row['auto_group_key']
            if auto_status.get(gid) == 'accepted':
                members = row['members_list']
                for m in members[1:]:
                    uf.union(members[0], m)
                history_records.append({'merge_source': 'auto_group', 'merge_id': gid, 'canonical_name': row['canonical_name'], 'members': ' | '.join(members), 'reason': row['reasons'], 'status': 'active'})
    if manual_decisions:
        for pair_key, row in manual_decisions.items():
            if row.get('decision') == 'merge':
                uf.union(row['name_a'], row['name_b'])
                history_records.append({'merge_source': 'manual_pair', 'merge_id': pair_key, 'canonical_name': row['suggested_canonical'], 'members': f"{row['name_a']} | {row['name_b']}", 'reason': row['reasons'], 'status': 'active'})
    all_names = set()
    for rec in history_records:
        parts = [p.strip() for p in rec['members'].split('|')]
        all_names.update([p for p in parts if p])
    groups = defaultdict(list)
    for name in all_names:
        groups[uf.find(name)].append(name)
    row_count_map = dict(zip(stats_df['raw_name'], stats_df['row_count']))
    clean_map = dict(zip(stats_df['raw_name'], stats_df['clean_name']))
    mapping_records = []
    for idx, members in enumerate(sorted(groups.values(), key=lambda g: (-len(g), sorted(g)[0])), start=1):
        canonical = sorted(members, key=lambda n: (-row_count_map.get(n, 0), *canonical_sort_key(n)))[0]
        cluster_id = f'M{idx:04d}'
        for member in sorted(members):
            mapping_records.append({'cluster_id': cluster_id, 'canonical_name': canonical, 'member_name': member, 'row_count': row_count_map.get(member, 0), 'clean_name': clean_map.get(member, '')})
    return (pd.DataFrame(history_records), pd.DataFrame(mapping_records))

def resolved_names_from_auto(auto_groups_df, auto_status):
    resolved = set()
    if auto_groups_df.empty:
        return resolved
    for _, row in auto_groups_df.iterrows():
        if auto_status.get(row['auto_group_key']) == 'accepted':
            resolved.update(row['members_list'])
    return resolved

def active_manual_decisions_for_config(manual_decisions: dict, column_config: dict) -> dict:
    active_entity = column_config['entity_column']
    return {key: record for key, record in manual_decisions.items() if record.get('entity_column') == active_entity}

def make_download_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode('utf-8')

def build_cleaned_workbook_bytes(raw_df: pd.DataFrame, mapping_df: pd.DataFrame, history_df: pd.DataFrame, manual_decisions: dict, auto_groups_df: pd.DataFrame, column_config: dict, sheet_name: str, workbook_bytes: bytes | None=None, candidate_queue_df: pd.DataFrame | None=None, overwrite_entity_column: bool=False) -> bytes:
    entity_col = column_config['entity_column']
    cleaned_df = raw_df.copy()
    original_values = cleaned_df[entity_col].fillna('').astype(str)
    normalized_original_values = original_values.map(lambda x: str(x).strip())
    canonical_map = {}
    cluster_map = {}
    merged_values = set()
    if not mapping_df.empty:
        canonical_map = dict(zip(mapping_df['member_name'], mapping_df['canonical_name']))
        cluster_map = dict(zip(mapping_df['member_name'], mapping_df['cluster_id']))
        merged_values = set(mapping_df['member_name'].tolist())
    status_map = {}
    source_map = {}
    score_map = {}
    reason_map = {}
    if not auto_groups_df.empty:
        accepted = set(history_df.loc[history_df['merge_source'] == 'auto_group', 'merge_id'].tolist()) if not history_df.empty else set()
        for _, row in auto_groups_df.iterrows():
            if row['auto_group_key'] in accepted:
                for member in row['members_list']:
                    status_map[member] = 'merged'
                    source_map[member] = 'auto_group'
                    reason_map[member] = row.get('reasons', '')
    for rec in manual_decisions.values():
        names = [rec.get('name_a', ''), rec.get('name_b', '')]
        decision = rec.get('decision', '')
        for name in names:
            if name in merged_values or decision == 'merge':
                status_map[name] = 'merged'
                source_map[name] = 'manual_pair'
            elif decision == 'keep_separate':
                status_map[name] = 'reviewed_keep_separate'
                source_map[name] = 'manual_keep_separate'
            elif decision == 'unsure':
                status_map[name] = 'reviewed_unsure'
                source_map[name] = 'manual_unsure'
            score_map[name] = rec.get('score', '')
            reason_map[name] = rec.get('reasons', '')
    unreviewed_names = set()
    if candidate_queue_df is not None and (not candidate_queue_df.empty):
        for _, row in candidate_queue_df.iterrows():
            unreviewed_names.add(row['name_a'])
            unreviewed_names.add(row['name_b'])
    cleaned_df['dedupe_entity_column'] = entity_col
    cleaned_df['dedupe_original_value'] = original_values
    cleaned_df['dedupe_canonical_value'] = normalized_original_values.map(lambda x: canonical_map.get(x, x))
    cleaned_df['dedupe_cluster_id'] = normalized_original_values.map(lambda x: cluster_map.get(x, ''))
    cleaned_df['dedupe_review_status'] = normalized_original_values.map(lambda x: status_map.get(x, '')).fillna('')
    cleaned_df['dedupe_decision_source'] = normalized_original_values.map(lambda x: source_map.get(x, '')).fillna('')
    cleaned_df['dedupe_score'] = normalized_original_values.map(lambda x: score_map.get(x, '')).fillna('')
    cleaned_df['dedupe_reason'] = normalized_original_values.map(lambda x: reason_map.get(x, '')).fillna('')

    def default_status(v: str) -> str:
        if v in merged_values:
            return 'merged'
        if v in unreviewed_names:
            return 'unreviewed'
        return 'not_flagged'
    missing_status = cleaned_df['dedupe_review_status'].isna() | (cleaned_df['dedupe_review_status'] == '')
    cleaned_df.loc[missing_status, 'dedupe_review_status'] = normalized_original_values[missing_status].map(default_status)

    dedupe_columns = [
        'dedupe_entity_column',
        'dedupe_original_value',
        'dedupe_canonical_value',
        'dedupe_cluster_id',
        'dedupe_review_status',
        'dedupe_decision_source',
        'dedupe_score',
        'dedupe_reason',
    ]

    if overwrite_entity_column:
        cleaned_df[entity_col] = cleaned_df['dedupe_canonical_value']

    def to_excel_value(value):
        if pd.isna(value):
            return None
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()
        return value

    if workbook_bytes:
        workbook = load_workbook(io.BytesIO(workbook_bytes))
        if sheet_name not in workbook.sheetnames:
            raise KeyError(f'Sheet not found: {sheet_name}')
        worksheet = workbook[sheet_name]

        def normalize_header(value):
            return value.strip() if isinstance(value, str) else value

        header_map = {}
        for col_idx in range(1, worksheet.max_column + 1):
            header_value = worksheet.cell(row=1, column=col_idx).value
            key = normalize_header(header_value)
            if key is not None:
                header_map[key] = col_idx

        existing_dedupe_indexes = [header_map.get(col) for col in dedupe_columns if header_map.get(col) is not None]
        if existing_dedupe_indexes:
            template_col_idx = max(1, min(existing_dedupe_indexes) - 1)
        else:
            template_col_idx = max(1, worksheet.max_column)
        template_letter = get_column_letter(template_col_idx)
        template_dim = worksheet.column_dimensions.get(template_letter)

        dedupe_column_indexes = {}
        for dedupe_col in dedupe_columns:
            existing_idx = header_map.get(dedupe_col)
            if existing_idx is not None:
                dedupe_column_indexes[dedupe_col] = existing_idx
            else:
                new_idx = worksheet.max_column + 1
                dedupe_column_indexes[dedupe_col] = new_idx
                worksheet.cell(row=1, column=new_idx, value=dedupe_col)
                if template_dim is not None:
                    new_letter = get_column_letter(new_idx)
                    worksheet.column_dimensions[new_letter].width = template_dim.width

        dedupe_max_col_idx = max(dedupe_column_indexes.values()) if dedupe_column_indexes else worksheet.max_column

        # If the worksheet uses an Excel table, extend the right edge to include dedupe columns
        # so table styling (header + banded rows) naturally applies to the new columns.
        for table in worksheet.tables.values():
            min_col, min_row, max_col, max_row = range_boundaries(table.ref)
            if min_row != 1:
                continue
            if max_col < template_col_idx:
                continue
            if dedupe_max_col_idx <= max_col:
                continue
            new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(dedupe_max_col_idx)}{max_row}"
            table.ref = new_ref
            if table.autoFilter is not None:
                table.autoFilter.ref = new_ref

            total_cols = dedupe_max_col_idx - min_col + 1
            existing_cols = list(table.tableColumns)
            updated_cols = []
            for position in range(total_cols):
                col_idx = min_col + position
                header_value = worksheet.cell(row=min_row, column=col_idx).value
                if header_value is None or str(header_value).strip() == '':
                    column_name = f'Column{position + 1}'
                else:
                    column_name = str(header_value)
                if position < len(existing_cols):
                    col_obj = existing_cols[position]
                    col_obj.id = position + 1
                    col_obj.name = column_name
                else:
                    col_obj = TableColumn(id=position + 1, name=column_name)
                updated_cols.append(col_obj)
            table.tableColumns = updated_cols

        data_row_count = len(cleaned_df)
        for dedupe_col, col_idx in dedupe_column_indexes.items():
            template_header = worksheet.cell(row=1, column=template_col_idx)
            target_header = worksheet.cell(row=1, column=col_idx)
            if template_header.has_style:
                target_header._style = copy(template_header._style)

            for row_idx in range(data_row_count):
                excel_row = row_idx + 2
                template_cell = worksheet.cell(row=excel_row, column=template_col_idx)
                target_cell = worksheet.cell(row=excel_row, column=col_idx)
                if template_cell.has_style:
                    target_cell._style = copy(template_cell._style)

            values = cleaned_df[dedupe_col].tolist()
            for row_idx, value in enumerate(values):
                excel_row = row_idx + 2
                worksheet.cell(row=excel_row, column=col_idx, value=to_excel_value(value))

            clear_start = data_row_count + 2
            if clear_start <= worksheet.max_row:
                for excel_row in range(clear_start, worksheet.max_row + 1):
                    worksheet.cell(row=excel_row, column=col_idx, value=None)

        if overwrite_entity_column:
            raw_columns = raw_df.columns.tolist()
            if entity_col in raw_columns:
                entity_col_idx = raw_columns.index(entity_col) + 1
                values = cleaned_df['dedupe_canonical_value'].tolist()
                for row_idx, value in enumerate(values):
                    excel_row = row_idx + 2
                    worksheet.cell(row=excel_row, column=entity_col_idx, value=to_excel_value(value))

        def autofit_columns_and_rows(ws):
            min_width = 8.0
            max_width = 80.0
            col_width_map = {}
            max_row = ws.max_row
            max_col = ws.max_column

            for col_idx in range(1, max_col + 1):
                max_len = 0
                for row_idx in range(1, max_row + 1):
                    cell_value = ws.cell(row=row_idx, column=col_idx).value
                    if cell_value is None:
                        continue
                    text = str(cell_value)
                    longest_segment = max((len(segment) for segment in text.splitlines()), default=0)
                    if longest_segment > max_len:
                        max_len = longest_segment
                width = min(max_width, max(min_width, float(max_len + 2)))
                ws.column_dimensions[get_column_letter(col_idx)].width = width
                col_width_map[col_idx] = width

            base_height = ws.sheet_format.defaultRowHeight or 15.0
            max_height = 180.0
            for row_idx in range(1, max_row + 1):
                max_lines = 1
                for col_idx in range(1, max_col + 1):
                    cell_value = ws.cell(row=row_idx, column=col_idx).value
                    if cell_value is None:
                        continue
                    text = str(cell_value)
                    segments = text.splitlines() if text else ['']
                    usable_width = max(1.0, col_width_map.get(col_idx, min_width) - 2.0)
                    line_count = 0
                    for segment in segments:
                        segment_len = len(segment)
                        line_count += max(1, int(ceil(segment_len / usable_width)))
                    if line_count > max_lines:
                        max_lines = line_count
                ws.row_dimensions[row_idx].height = min(max_height, float(base_height * max_lines))

        autofit_columns_and_rows(worksheet)

        out = io.BytesIO()
        workbook.save(out)
        return out.getvalue()

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
    return out.getvalue()

def build_standardized_workbook_bytes(raw_df: pd.DataFrame, mapping_df: pd.DataFrame, column_config: dict, sheet_name: str, workbook_bytes: bytes | None) -> bytes:
    entity_col = column_config['entity_column']
    canonical_map = {}
    if not mapping_df.empty:
        canonical_map = dict(zip(mapping_df['member_name'], mapping_df['canonical_name']))

    standardized_df = raw_df.copy()
    standardized_original_values = standardized_df[entity_col].fillna('').astype(str).map(lambda x: str(x).strip())
    standardized_df[entity_col] = standardized_original_values.map(lambda x: canonical_map.get(x, x))

    def to_excel_value(value):
        if pd.isna(value):
            return None
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()
        return value

    if workbook_bytes:
        workbook = load_workbook(io.BytesIO(workbook_bytes))
        if sheet_name not in workbook.sheetnames:
            raise KeyError(f'Sheet not found: {sheet_name}')
        worksheet = workbook[sheet_name]

        raw_columns = raw_df.columns.tolist()
        if entity_col not in raw_columns:
            raise KeyError(f'Entity column not found: {entity_col}')
        entity_col_idx = raw_columns.index(entity_col) + 1

        values = standardized_df[entity_col].tolist()
        for row_idx, value in enumerate(values):
            excel_row = row_idx + 2
            worksheet.cell(row=excel_row, column=entity_col_idx, value=to_excel_value(value))

        out = io.BytesIO()
        workbook.save(out)
        return out.getvalue()

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        standardized_df.to_excel(writer, sheet_name=sheet_name, index=False)
    return out.getvalue()

def now_iso():
    return datetime.now(timezone.utc).isoformat()

def build_manual_decision_record(row: pd.Series, decision: str, entity_column: str):
    return {'pair_key': row['pair_key'], 'entity_column': entity_column, 'name_a': row['name_a'], 'name_b': row['name_b'], 'decision': decision, 'score': float(row['score']), 'raw_name_score': float(row['raw_name_score']), 'clean_name_score': float(row['clean_name_score']), 'year_score': float(row['year_score']), 'unit_score': float(row['unit_score']), 'type_score': float(row['type_score']), 'cargo_amount_score': float(row['cargo_amount_score']), 'reasons': row['reasons'], 'suggested_canonical': row['suggested_canonical'], 'updated_at': now_iso()}
