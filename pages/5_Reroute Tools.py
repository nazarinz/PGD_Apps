from utils.auth import require_login

require_login()

"""
Reroute Tools Suite - Unified Streamlit App
Menggabungkan:
  1. Reroute Merge Tool V2   – Compare Old vs New PO + size breakdown
  2. Old PO Filter & Export  – Filter baris Old PO, 2-sheet output
  3. Email Prep Tool         – Filter, reorder, isi Email Subject & Remark2
"""

# ============================================================
# IMPORTS
# ============================================================
import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from typing import List, Tuple, Dict, Optional

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Reroute Tools Suite",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

[data-testid="stSidebar"] {
    background: #0f1117;
    border-right: 1px solid #1e2130;
}
[data-testid="stSidebar"] * { color: #e0e0e0 !important; }
[data-testid="stSidebar"] .stRadio label { font-size: 0.9rem; padding: 6px 0; cursor: pointer; }
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover { color: #60a5fa !important; }

.app-header {
    background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
    border-radius: 12px; padding: 28px 36px; margin-bottom: 28px;
    border: 1px solid #2a3a4a; position: relative; overflow: hidden;
}
.app-header::before {
    content: ""; position: absolute; top: -40px; right: -40px;
    width: 180px; height: 180px;
    background: radial-gradient(circle, rgba(96,165,250,0.15) 0%, transparent 70%);
    border-radius: 50%;
}
.app-header h1 {
    font-family: 'IBM Plex Mono', monospace; font-size: 1.8rem; font-weight: 600;
    color: #e2e8f0; margin: 0 0 6px 0; letter-spacing: -0.5px;
}
.app-header p { color: #94a3b8; margin: 0; font-size: 0.9rem; }
.badge {
    display: inline-block; background: rgba(96,165,250,0.15); color: #60a5fa;
    border: 1px solid rgba(96,165,250,0.3); border-radius: 20px; padding: 2px 10px;
    font-size: 0.72rem; font-family: 'IBM Plex Mono', monospace; margin-top: 8px;
}

.tool-card {
    background: #161b27; border: 1px solid #1e2a3a; border-radius: 10px;
    padding: 20px 24px; margin-bottom: 16px; transition: border-color 0.2s;
}
.tool-card:hover { border-color: #60a5fa44; }
.tool-card h3 { font-family: 'IBM Plex Mono', monospace; font-size: 1rem; color: #60a5fa; margin: 0 0 6px 0; }
.tool-card p { color: #94a3b8; font-size: 0.85rem; margin: 0; line-height: 1.5; }

.stat-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 16px 0; }
.stat-box {
    background: #161b27; border: 1px solid #1e2a3a; border-radius: 8px;
    padding: 14px 20px; flex: 1; min-width: 130px;
}
.stat-box .num { font-family: 'IBM Plex Mono', monospace; font-size: 1.6rem; font-weight: 600; color: #60a5fa; display: block; }
.stat-box .lbl { font-size: 0.75rem; color: #64748b; display: block; margin-top: 2px; }

.step-label { display: inline-flex; align-items: center; gap: 8px; font-family: 'IBM Plex Mono', monospace; font-size: 0.8rem; color: #60a5fa; margin-bottom: 8px; }
.step-num { background: rgba(96,165,250,0.15); border: 1px solid rgba(96,165,250,0.3); border-radius: 50%; width: 22px; height: 22px; display: inline-flex; align-items: center; justify-content: center; font-size: 0.7rem; }

.sect-divider { border: none; border-top: 1px solid #1e2a3a; margin: 24px 0; }

.log-box {
    background: #0a0e17; border: 1px solid #1e2a3a; border-radius: 8px;
    padding: 14px 18px; font-family: 'IBM Plex Mono', monospace; font-size: 0.78rem;
    color: #94a3b8; max-height: 260px; overflow-y: auto; line-height: 1.7; white-space: pre-wrap;
}

.stButton > button {
    background: linear-gradient(135deg, #1d4ed8, #2563eb); color: white; border: none;
    border-radius: 8px; padding: 10px 24px; font-family: 'IBM Plex Sans', sans-serif;
    font-weight: 600; font-size: 0.9rem; transition: all 0.2s; width: 100%;
}
.stButton > button:hover { background: linear-gradient(135deg, #1e40af, #1d4ed8); box-shadow: 0 4px 20px rgba(29,78,216,0.3); }
.stDownloadButton > button {
    background: linear-gradient(135deg, #065f46, #059669) !important; color: white !important;
    border: none !important; border-radius: 8px !important; font-weight: 600 !important; width: 100% !important;
}
div[data-testid="stFileUploader"] { border: 1px dashed #1e2a3a; border-radius: 10px; padding: 8px; background: #0f1117; }
.stTextInput input, .stTextArea textarea {
    background: #0f1117 !important; border: 1px solid #1e2a3a !important;
    color: #e2e8f0 !important; border-radius: 8px !important; font-family: 'IBM Plex Mono', monospace !important;
}
.stSuccess { border-radius: 8px; }
.stWarning { border-radius: 8px; }
.stError   { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


# ============================================================
#  CORE LOGIC — V2 REROUTE MERGE ENGINE
# ============================================================

class Config:
    UNICODE_SPACES = r"[\u00A0\u1680\u2000-\u200B\u202F\u205F\u3000]"
    SIZE_PATTERN   = re.compile(r'^(?:UK[_\-\s]*)?(\d{1,2})(K|-K|-)?$', re.I)
    MAX_HEADER_SCAN_ROWS = 40

    NEW_PO_RENAMES = {
        'PO Batch Date': 'PO Date', 'PO Number': 'Sold-To PO No.',
        'Market PO Number': 'Ship-To Party PO No.', 'Customer Request Date (CRD)': 'CRD',
        'Plan Date': 'PD', 'Article Number': 'Cust Article No.',
        'Gps Customer Number': 'Ship-To Search Term', 'Class Code': 'Classification Code',
        'Grand Total': 'Order Quantity', 'MTFC Number': 'Customer Contract ID',
    }
    OLD_PO_RENAMES = {'Article Number': 'Cust Article No.'}

    HEADER_TOKENS = [
        'articlenumber','articleno','article','modelname','ponumber','marketponumber',
        'customerrequestdate','plandate','grandtotal','classcode','gpscustomernumber',
        'shipmentmode','plantcode','orderquantity','customercontractid','custarticleno',
    ]
    DATE_TOKENS = ['po date','crd','pd','fpd','posdd','lpd','podd','vas cut-off date','document date', 'prod. team atp']

    OUTPUT_COLUMNS = [
        'Remark','Select','Status','Working Status','Working Status Descr.',
        'PO Date','Requirement Segment','Order Type','Site','Work Center',
        'Sales Order','Customer Contract ID','Sold-To PO No.','Ship-To Party PO No.',
        'CRD','PD','Prod. Team ATP','FPD','FPD-DRC','POSDD','POSDD-DRC',
        'LPD','LPD-DRC','PODD','PODD-DRC','FGR','Cust Article No.','Model Name',
        'Gender','Article','Article Lead Time','Develop Type','Last Code','Season',
        'Product Hierarchy 3','Outsole Mold','Pattern Code (Upper','Ship-To No.',
        'Ship-To Search Term','Ship-To Name','Ship-To Country','Shipping Type',
        'Shipping Type 2','Packing Type','VAS Cut-Off Date','Classification Code',
        'Changed By','Document Date','Order Quantity',
    ]
    TEXT_FORMAT_COLUMNS = [
        'Customer Contract ID','Sold-To PO No.','Ship-To Party PO No.',
        'Sales Order','Ship-To No.','Ship-To Search Term',
    ]


def canonicalize(text: str) -> str:
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

def clean_headers(columns) -> List[str]:
    headers = pd.Index(columns).astype(str)
    headers = headers.str.replace(Config.UNICODE_SPACES, ' ', regex=True)
    headers = headers.str.replace(r'\s+', ' ', regex=True).str.strip()
    return list(headers)

def to_numeric_safe(series) -> pd.Series:
    return pd.to_numeric(series, errors='coerce')


class HeaderDetector:
    @staticmethod
    def detect_header_row(df: pd.DataFrame, max_scan: int = Config.MAX_HEADER_SCAN_ROWS) -> int:
        n_scan = min(max_scan, len(df))
        best_score = (-1, 0)
        size_pattern = re.compile(r"(?:uk\s*[_-]?)?\d{1,2}(?:-k|k|-)?$", re.I)
        for i in range(n_scan):
            row = df.iloc[i].astype(str).fillna("")
            row_clean = [re.sub(Config.UNICODE_SPACES, ' ', x).strip() for x in row]
            row_canon = [canonicalize(x) for x in row_clean]
            nonempty  = sum(1 for x in row_clean if x)
            token_hits = sum(any(tok in c for tok in Config.HEADER_TOKENS) for c in row_canon)
            size_hits  = sum(1 for x in row_clean if size_pattern.fullmatch(x))
            score = token_hits * 3 + size_hits + nonempty * 0.1
            if score > best_score[0]:
                best_score = (score, i)
        return best_score[1]


class ExcelReader:
    @staticmethod
    def read_with_autodetect(content_bytes: bytes) -> Tuple[pd.DataFrame, int]:
        df_raw = pd.read_excel(
            io.BytesIO(content_bytes), header=None, dtype=str,
            na_values=['', 'NA', 'N/A', 'null', 'NULL']
        )
        header_idx = HeaderDetector.detect_header_row(df_raw)
        header_row = df_raw.iloc[header_idx].astype(str).fillna("")
        header_row = [re.sub(Config.UNICODE_SPACES, ' ', x).strip() for x in header_row]
        headers = [h if h else f"Unnamed: {j}" for j, h in enumerate(header_row)]
        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = clean_headers(headers)
        df = df.dropna(how='all')
        return df, header_idx


class ColumnManager:
    @staticmethod
    def alias_column(df, target, exact_matches, loose_tokens):
        if target in df.columns: return
        for c in exact_matches:
            if c in df.columns:
                df.rename(columns={c: target}, inplace=True); return
        avoid = ["leadtime", "lead", "time"]
        for col in list(df.columns):
            cc = canonicalize(col)
            if any(a in cc for a in avoid): continue
            if all(t in cc for t in loose_tokens):
                df.rename(columns={col: target}, inplace=True); return

    @staticmethod
    def deduplicate_columns(df, name="DataFrame", log=None):
        if df.columns.duplicated().any():
            dups = list(pd.Index(df.columns)[df.columns.duplicated()])
            msg = f"⚠️ {name} duplicate columns removed: {dups}"
            if log is not None: log.append(msg)
            df = df.loc[:, ~df.columns.duplicated()].copy()
        return df


class SizeColumnHandler:
    @staticmethod
    def canonical_size_name(col: str) -> str:
        s = re.sub(r'\s+', '', str(col)).upper()
        s = s.replace('UK-', 'UK_').replace('UK', 'UK_').replace('__', '_')
        m = Config.SIZE_PATTERN.match(s.replace('_', ''))
        if m: return f'UK_{m.group(1)}{m.group(2) or ""}'
        m2 = re.match(r'^UK_(\d{1,2})(K|-K|-)?$', s)
        if m2: return f'UK_{m2.group(1)}{m2.group(2) or ""}'
        return col

    @staticmethod
    def size_sort_key(name: str) -> Tuple:
        m = re.fullmatch(r'UK_(\d{1,2})(K|-K|-)?', str(name))
        if not m: return (2, 999, 9, str(name))
        base, suffix = int(m.group(1)), m.group(2) or ''
        group  = 0 if 'K' in suffix else 1
        within = {'K': 0, '-K': 1, '': 0, '-': 1}.get(suffix, 9)
        return (group, base, within, str(name))

    @staticmethod
    def normalize_size_columns(df):
        return df.rename(columns={c: SizeColumnHandler.canonical_size_name(c) for c in df.columns})

    @staticmethod
    def get_size_columns(df):
        return [c for c in df.columns if re.match(r'^UK_\d{1,2}(K|-K|-)?$', str(c))]


class SizeSignature:
    @staticmethod
    def create(df):
        size_cols = SizeColumnHandler.get_size_columns(df)
        if not size_cols:
            return pd.Series(['NOSIZE'] * len(df), index=df.index)
        size_cols_sorted = sorted(size_cols, key=SizeColumnHandler.size_sort_key)
        sig_parts = [to_numeric_safe(df[c]).fillna(0).astype(int).astype(str) for c in size_cols_sorted]
        sig_df = pd.concat(sig_parts, axis=1)
        sig_df.columns = size_cols_sorted
        return sig_df.apply(lambda row: ','.join(row.values.astype(str)), axis=1)


class DataProcessor:
    @staticmethod
    def compute_sizes_and_qty(df):
        df = df.copy()
        size_cols = SizeColumnHandler.get_size_columns(df)
        if size_cols:
            df[size_cols] = df[size_cols].apply(to_numeric_safe).fillna(0)
            df['SizeSum'] = df[size_cols].sum(axis=1).astype('Int64')
        else:
            df['SizeSum'] = pd.NA
        if 'Order Quantity' not in df.columns and 'Grand Total' in df.columns:
            df['Order Quantity'] = df['Grand Total']
        order_qty = to_numeric_safe(df.get('Order Quantity'))
        df['OrderQty_fix'] = order_qty
        needs_fix = (df['SizeSum'].notna() & (df['OrderQty_fix'].isna() | (df['OrderQty_fix'] != df['SizeSum'])))
        df.loc[needs_fix, 'OrderQty_fix'] = df.loc[needs_fix, 'SizeSum']
        df['OrderQty_fix'] = df['OrderQty_fix'].fillna(0).astype('Int64')
        df['Qty_Diff'] = (df['OrderQty_fix'] - order_qty.fillna(0)).astype('Int64')
        return df

    @staticmethod
    def normalize_article_column(df):
        df = df.copy()
        if 'Cust Article No.' not in df.columns:
            ColumnManager.alias_column(df, 'Cust Article No.',
                ['Cust Article No.', 'Article Number', 'Article No', 'Article'], ['article', 'number'])
        if 'Cust Article No.' not in df.columns:
            raise ValueError("Column 'Cust Article No.' not found")
        df['Cust Article No.'] = df['Cust Article No.'].astype(str).str.strip().str.upper()
        return df

    @staticmethod
    def create_merge_key(df, include_sales_order=False):
        parts = [df['Cust Article No.'].astype(str), df['OrderQty_fix'].astype(str), df['size_sig'].astype(str)]
        if include_sales_order and 'Sales Order' in df.columns:
            parts.append(df['Sales Order'].astype(str).fillna('').str.strip())
        key = parts[0]
        for p in parts[1:]: key = key + '|' + p
        return key

    @staticmethod
    def aggregate_by_merge_key(df):
        size_cols = SizeColumnHandler.get_size_columns(df)
        df = df.loc[:, ~df.columns.duplicated()].copy()
        base_cols = ['merge_key'] + size_cols
        meta_cols = [c for c in df.columns if c not in base_cols]
        if size_cols:
            sum_part = df.groupby('merge_key', sort=False)[size_cols].sum(numeric_only=True).reset_index()
        else:
            sum_part = df[['merge_key']].drop_duplicates()
        rows = []
        for name, group in df.groupby('merge_key', sort=False):
            vals = {'merge_key': name}
            for col in meta_cols:
                nn = group[col].dropna()
                vals[col] = nn.iloc[0] if len(nn) else pd.NA
            rows.append(vals)
        first_part = pd.DataFrame(rows)
        result = first_part.merge(sum_part, on='merge_key', how='left')
        if size_cols:
            result['SizeSum'] = result[size_cols].sum(axis=1).astype('Int64')
        return result


class PONumberNormalizer:
    @staticmethod
    def normalize_value(value):
        if pd.isna(value): return pd.NA
        s = str(value).strip().replace(',', '')
        if s.lower() in {'', 'nan', 'none'}: return pd.NA
        if re.fullmatch(r'\d+\.0+', s): s = s.split('.')[0]
        return s

    @staticmethod
    def normalize_column(df, col_name):
        if col_name in df.columns:
            df[col_name] = df[col_name].apply(PONumberNormalizer.normalize_value)
        return df


class CustomerContractNormalizer:
    @staticmethod
    def normalize_value(value):
        if pd.isna(value): return pd.NA
        s = str(value).strip().replace(',', '')
        if s.lower() in {'', 'nan', 'none'}: return pd.NA
        try:
            f = float(s)
            if abs(f - round(f)) < 1e-9: return str(int(round(f)))
            return s
        except: pass
        if re.fullmatch(r'\d+\.0', s): s = s[:-2]
        return s

    @staticmethod
    def propagate_across_groups(df):
        if 'Customer Contract ID' not in df.columns: return df
        df = df.copy()
        df['Customer Contract ID'] = df['Customer Contract ID'].apply(CustomerContractNormalizer.normalize_value)
        group_keys = ['Cust Article No.', 'OrderQty_fix', 'size_sig']
        if 'Sold-To PO No.' in df.columns: group_keys.append('Sold-To PO No.')
        cc_series = (df.sort_values(group_keys + ['sort_order'])
                     .groupby(group_keys)['Customer Contract ID']
                     .apply(lambda s: next((x for x in s if pd.notna(x)), pd.NA)))
        cc_map = cc_series.reset_index(name='CC_fill')
        df = df.merge(cc_map, on=group_keys, how='left')
        df['Customer Contract ID'] = df['CC_fill'].apply(CustomerContractNormalizer.normalize_value)
        df.drop(columns=['CC_fill'], inplace=True)
        return df


class POPreparation:
    @staticmethod
    def prepare(df, remark, sort_order, include_sales_order=False):
        if 'Shipping Type.1' in df.columns and 'Shipping Type 2' not in df.columns:
            df = df.rename(columns={'Shipping Type.1': 'Shipping Type 2'})
        for col in ['Sold-To PO No.','Ship-To Party PO No.','Sales Order','Ship-To No.','Ship-To Search Term']:
            df = PONumberNormalizer.normalize_column(df, col)
        df = SizeColumnHandler.normalize_size_columns(df)
        if 'Order Quantity' not in df.columns: df['Order Quantity'] = pd.NA
        if 'Customer Contract ID' not in df.columns: df['Customer Contract ID'] = pd.NA
        df = DataProcessor.normalize_article_column(df)
        df = DataProcessor.compute_sizes_and_qty(df)
        df['size_sig']  = SizeSignature.create(df)
        df['merge_key'] = DataProcessor.create_merge_key(df, include_sales_order)
        df = DataProcessor.aggregate_by_merge_key(df)
        df['Remark']     = remark
        df['sort_order'] = sort_order
        return df


class SizeComparison:
    @staticmethod
    def create_comparison(merged_df, size_cols):
        if not size_cols: return pd.DataFrame()
        cols_needed = ['Remark','Cust Article No.','OrderQty_fix','size_sig']
        if 'Sold-To PO No.' in merged_df.columns: cols_needed.append('Sold-To PO No.')
        cols_needed += size_cols
        base = merged_df[cols_needed].copy()
        base[size_cols] = base[size_cols].apply(to_numeric_safe).fillna(0).astype(int)
        base['comparison_key'] = (base['Cust Article No.'].astype(str) + '|' +
                                  base['OrderQty_fix'].astype(str) + '|' + base['size_sig'].astype(str))
        wide = base.pivot_table(index=['comparison_key','Cust Article No.','OrderQty_fix'],
                                columns='Remark', values=size_cols, aggfunc='sum', fill_value=0)
        has_old = 'Old PO - Canceled' in wide.columns.get_level_values(1)
        has_new = 'New PO - Reroute'  in wide.columns.get_level_values(1)
        if not (has_old and has_new): return pd.DataFrame()
        old  = wide.xs('Old PO - Canceled', axis=1, level=1).reindex(columns=size_cols, fill_value=0)
        new  = wide.xs('New PO - Reroute',  axis=1, level=1).reindex(columns=size_cols, fill_value=0)
        diff = new - old
        po_val = 'Sold-To PO No.' if 'Sold-To PO No.' in base.columns else 'Remark'
        po_piv = base.pivot_table(index=['comparison_key'], columns='Remark', values=po_val, aggfunc='first')
        po_old = po_piv.get('Old PO - Canceled', pd.Series(index=po_piv.index, dtype=object))
        po_new = po_piv.get('New PO - Reroute',  pd.Series(index=po_piv.index, dtype=object))
        summary = pd.DataFrame({
            'Cust Article No.': [i[1] for i in new.index],
            'OrderQty_fix': [i[2] for i in new.index],
            'Old_PO_No': po_old.reindex(old.index.get_level_values(0)).values,
            'New_PO_No': po_new.reindex(new.index.get_level_values(0)).values,
            'All_Sizes_Equal': (diff == 0).all(axis=1).values,
            'Diff_Count': (diff != 0).sum(axis=1).values,
        }, index=new.index)
        def suf(d, s):
            d = d.copy(); d.columns = [f"{c}__{s}" for c in d.columns]; return d
        return pd.concat([summary.reset_index(drop=True),
                          suf(old.reset_index(drop=True), "old"),
                          suf(new.reset_index(drop=True), "new"),
                          suf(diff.reset_index(drop=True), "diff")], axis=1)

    @staticmethod
    def split_by_equality(compare_df):
        if compare_df.empty: return pd.DataFrame(), pd.DataFrame()
        return (compare_df[~compare_df['All_Sizes_Equal']].reset_index(drop=True),
                compare_df[ compare_df['All_Sizes_Equal']].reset_index(drop=True))


def get_date_columns(df):
    date_cols = []
    for col in df.columns:
        cl = col.lower()
        if 'drc' in cl or '-dr' in cl: continue
        if any(t in cl for t in Config.DATE_TOKENS): date_cols.append(col)
    return date_cols


def build_excel_bytes(main_df, compare_df, only_diff, only_equal, date_cols, size_cols):
    buf = io.BytesIO()
    def col_letter(n):
        s = ""
        while n >= 0: s = chr(n % 26 + 65) + s; n = n // 26 - 1
        return s
    with pd.ExcelWriter(buf, engine='xlsxwriter', datetime_format='m/d/yyyy') as writer:
        main_df.to_excel(writer, index=False, sheet_name='Sheet1')
        wb = writer.book; ws = writer.sheets['Sheet1']
        ws.autofilter(0, 0, len(main_df), len(main_df.columns) - 1)
        ws.freeze_panes(1, 0)
        text_fmt = wb.add_format({'num_format': '@'})
        date_fmt = wb.add_format({'num_format': 'm/d/yyyy'})
        for col in Config.TEXT_FORMAT_COLUMNS:
            if col in main_df.columns:
                idx = main_df.columns.get_loc(col); ws.set_column(idx, idx, 18, text_fmt)
        for col in date_cols:
            if col in main_df.columns:
                idx = main_df.columns.get_loc(col); ws.set_column(idx, idx, 12, date_fmt)
        fmt_set = set(Config.TEXT_FORMAT_COLUMNS + date_cols)
        for ci, cn in enumerate(main_df.columns):
            if cn in fmt_set: continue
            mv = max([len(str(cn))] + [len(str(x)) for x in main_df[cn].head(500).fillna("").astype(str)])
            ws.set_column(ci, ci, min(45, max(10, mv + 2)))
        if 'Remark' in main_df.columns:
            red = wb.add_format({'font_color': 'red'})
            lc  = col_letter(len(main_df.columns) - 1)
            ws.conditional_format(f"A2:{lc}{len(main_df)+1}", {
                'type': 'formula', 'criteria': '=$A2="Old PO - Canceled"', 'format': red})
        if not compare_df.empty: compare_df.to_excel(writer, sheet_name='Size_Compare', index=False)
        if not only_diff.empty:  only_diff.to_excel(writer,  sheet_name='Only_Different', index=False)
        if not only_equal.empty: only_equal.to_excel(writer,  sheet_name='Only_Equal', index=False)
    buf.seek(0); return buf.read()


# ============================================================
#  CORE LOGIC — OLD PO FILTER TOOL
# ============================================================

OLD_PO_FILTER_COL   = "Remark"
OLD_PO_FILTER_VALUE = "Old PO"
OLD_PO_STRING_COLS  = ["Sold-To PO No.", "Sales Order", "Cust Article No."]

SHEET1_ORDER = [
    "Working Status","Select","Working Status Descr.","Requirement Segment",
    "SO","Sold-To PO No.","CRD","CRD-DRC","PD","POSDD-DRC","POSDD",
    "FPD","FPD-DRC","PODD","PODD-DRC","Est. Inspection Date",
    "LPD","LPD-DRC","FGR","Cust Article No.","Model Name","Article",
    "Lead Time","Season","Product Hierarchy 3","Ship-To Search Term",
    "Ship-To Country","Document Date","Order Quantity","Order Type",
]
SHEET2_ORDER = [
    "Sales Order","Sold-To PO No.","Cust Article No.","Model Name",
    "Ship-To Search Term","Ship-To Country","Document Date","Order Quantity","Order Type",
]


def read_excel_preserve_strings_generic(file_bytes: bytes, string_cols: list) -> Dict[str, pd.DataFrame]:
    xls_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine="openpyxl", header=0)
    result = {}
    for sheet_name, df_raw in xls_raw.items():
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        converters = {col: str for col in string_cols if col in df_raw.columns}
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name,
                           engine="openpyxl", converters=converters)
        df.columns = [str(c).strip() for c in df.columns]
        for col in string_cols:
            if col in df.columns:
                df[col] = (df[col].astype(str).str.strip()
                           .str.replace(r"\.0$", "", regex=True)
                           .replace("nan", pd.NA))
        result[sheet_name] = df
    return result


def filter_old_po_strict(df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    if OLD_PO_FILTER_COL not in df.columns:
        return pd.DataFrame(columns=df.columns), f"Kolom '{OLD_PO_FILTER_COL}' tidak ditemukan"
    remark = df[OLD_PO_FILTER_COL].astype(str).str.strip()
    mask = remark.str.contains(OLD_PO_FILTER_VALUE, case=False, na=False) & \
           ~remark.isin(["","nan","none","NaN","None","NaT"])
    filtered = df[mask].copy()
    return filtered, f"Diambil {len(filtered)} / {len(df)} baris"


def build_sheet_cols(df: pd.DataFrame, desired: list, rename_map: dict = None):
    df = df.copy()
    if rename_map: df = df.rename(columns=rename_map)
    missing = [c for c in desired if c not in df.columns]
    for col in missing: df[col] = pd.NA
    return df[desired], missing


def build_old_po_excel(df_combined: pd.DataFrame) -> bytes:
    df_s1, _ = build_sheet_cols(df_combined, SHEET1_ORDER, rename_map={"Sales Order": "SO"})
    df_s2, _ = build_sheet_cols(df_combined, SHEET2_ORDER)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_s1.to_excel(writer, sheet_name="Old PO - Full",    index=False)
        df_s2.to_excel(writer, sheet_name="Old PO - Summary", index=False)
    buf.seek(0); return buf.read()


# ============================================================
#  CORE LOGIC — EMAIL PREP TOOL
# ============================================================

# Columns that must be read as plain text to preserve leading zeros
EMAIL_PREP_STRING_COLS = [
    "Sold-To PO No.",
    "Ship-To Party PO No.",
    "Sales Order",
    "Ship-To Search Term",
    "Cust Article No.",
]

# ── NEW column order (updated per user request) ──────────────
EMAIL_PREP_TARGET = [
    "Sales Order",
    "Remark",
    "Sold-To PO No.",
    "Ship-To Party PO No.",
    "Model Name",
    "Cust Article No.",
    "Article",
    "Order Quantity",
    "CRD",
    "PD",
    "LPD",
    "PODD",
    "Ship-To Search Term",
    "New Ship-To Search Term",   # blank – filled by user later
    "Ship-To Country",
    "Email Subject",             # auto-filled
    "Remark2",                   # auto-filled = "Full Reroute"
    "PO remark",                 # blank – filled by user later
    "Prod. Status",
]


def _clean_po_string(series: pd.Series) -> pd.Series:
    """Strip whitespace, remove trailing .0 from float-strings, blank out nan/None."""
    cleaned = (
        series.astype(str)
        .str.strip()
        .str.replace(r"(?<=\d)\.0$", "", regex=True)
    )
    cleaned = cleaned.replace({"nan": pd.NA, "None": pd.NA, "NaT": pd.NA, "": pd.NA})
    return cleaned


def email_prep_process(df: pd.DataFrame, email_subject: str) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Stash Prod. Status before reordering
    prod_status = df["Prod. Status"].values if "Prod. Status" in df.columns else [pd.NA] * len(df)

    # Add missing columns as blank
    for col in EMAIL_PREP_TARGET:
        if col not in df.columns:
            df[col] = pd.NA

    out = df[EMAIL_PREP_TARGET].copy()

    # Auto-fill
    out["Email Subject"] = email_subject
    out["Remark2"]       = "Full Reroute"
    out["Prod. Status"]  = prod_status
    # "New Ship-To Search Term" and "PO remark" remain blank (pd.NA)

    return out


def build_email_prep_excel(sheets_out: Dict[str, pd.DataFrame]) -> bytes:
    """
    Write output Excel.
    PO-number columns are cell-formatted as Text (@) so Excel preserves leading zeros.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sname, df in sheets_out.items():
            df.to_excel(writer, sheet_name=sname, index=False)
            ws = writer.sheets[sname]
            # Apply Text number format to all PO / ID string columns
            for col_idx, col_name in enumerate(df.columns, start=1):
                if col_name in EMAIL_PREP_STRING_COLS:
                    for row_idx in range(2, len(df) + 2):   # row 1 = header
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.number_format = "@"             # Text format in Excel
    buf.seek(0); return buf.read()


# ============================================================
#  UI HELPERS
# ============================================================

def stat_box_html(stats: dict) -> str:
    items = "".join(f'<div class="stat-box"><span class="num">{v}</span><span class="lbl">{k}</span></div>'
                    for k, v in stats.items())
    return f'<div class="stat-row">{items}</div>'


def step(n, label):
    st.markdown(f'<div class="step-label"><span class="step-num">{n}</span>{label}</div>',
                unsafe_allow_html=True)


# ============================================================
#  PAGES
# ============================================================

def page_home():
    st.markdown("""
    <div class="app-header">
        <h1>🔄 Reroute Tools Suite</h1>
        <p>Alat bantu pemrosesan PO Reroute — compare, filter, dan siapkan data ekspor.</p>
        <span class="badge">v2.0 · Unified Edition</span>
    </div>
    """, unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""<div class="tool-card"><h3>🔀 Reroute Merge</h3>
        <p>Gabungkan Old PO (Canceled) dan New PO (Reroute) dalam satu file dengan breakdown size, validasi, dan color-coding otomatis.</p></div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""<div class="tool-card"><h3>📋 Input Pending Cancel</h3>
        <p>Filter baris Old PO dari file merge, lalu ekspor ke 2 sheet: <em>Full</em> (lengkap) dan <em>Summary</em> (ringkasan).</p></div>""", unsafe_allow_html=True)
    with col3:
        st.markdown("""<div class="tool-card"><h3>✉️ Input Tracking Reroute Report</h3>
        <p>Reorder kolom, isi Email Subject dan Remark2 otomatis, siap kirim ke customer.</p></div>""", unsafe_allow_html=True)
    st.markdown('<hr class="sect-divider">', unsafe_allow_html=True)
    st.markdown("**Pilih tool dari sidebar kiri untuk mulai.**")


def page_reroute_merge():
    st.markdown("""
    <div class="app-header">
        <h1>🔀 Reroute Merge Tool</h1>
        <p>Upload New PO + Old PO → output Excel dengan 4 sheet: main, size compare, only different, only equal.</p>
        <span class="badge">V2 · MTFC Number support · Size Validation</span>
    </div>
    """, unsafe_allow_html=True)

    step(1, "Upload File")
    col_new, col_old = st.columns(2)
    with col_new: file_new = st.file_uploader("📄 New PO (Reroute)", type=["xlsx","xls"], key="merge_new")
    with col_old: file_old = st.file_uploader("📄 Old PO (Canceled)", type=["xlsx","xls"], key="merge_old")
    st.markdown('<hr class="sect-divider">', unsafe_allow_html=True)
    if not (file_new and file_old):
        st.info("Upload kedua file untuk melanjutkan."); return

    step(2, "Konfigurasi")
    col_a, col_b = st.columns(2)
    with col_a: include_sales = st.checkbox("Sertakan Sales Order dalam merge key", value=False)
    with col_b: output_name   = st.text_input("Nama file output", value="Merged_Old_New_V2.xlsx")
    st.markdown('<hr class="sect-divider">', unsafe_allow_html=True)

    step(3, "Proses")
    if not st.button("▶ Jalankan Merge"): return

    log = []
    with st.spinner("Memproses…"):
        try:
            bytes_new = file_new.read(); bytes_old = file_old.read()
            df_new_raw, hdr_new = ExcelReader.read_with_autodetect(bytes_new)
            df_old_raw, hdr_old = ExcelReader.read_with_autodetect(bytes_old)
            log.append(f"✅ NEW header row: {hdr_new} | cols: {len(df_new_raw.columns)}")
            log.append(f"✅ OLD header row: {hdr_old} | cols: {len(df_old_raw.columns)}")
            df_new_raw = ColumnManager.deduplicate_columns(df_new_raw, "NEW", log)
            df_old_raw = ColumnManager.deduplicate_columns(df_old_raw, "OLD", log)
            df_new = df_new_raw.copy()
            df_new.rename(columns={k: v for k, v in Config.NEW_PO_RENAMES.items() if k in df_new.columns}, inplace=True)
            df_old = df_old_raw.copy()
            df_old.rename(columns={k: v for k, v in Config.OLD_PO_RENAMES.items() if k in df_old.columns}, inplace=True)
            ColumnManager.alias_column(df_new, 'Cust Article No.',
                ['Cust Article No.','Article Number','Article No','Article'], ['article','number'])
            ColumnManager.alias_column(df_new, 'Order Quantity',
                ['Order Quantity','Grand Total','Total Qty','Order Qty','Quantity'], ['order','qty'])
            inc_sales = include_sales and ('Sales Order' in df_old.columns) and ('Sales Order' in df_new.columns)
            log.append(f"🔑 Sales Order in merge key: {inc_sales}")
            df_old_prep = POPreparation.prepare(df_old, 'Old PO - Canceled', 0, inc_sales)
            df_new_prep = POPreparation.prepare(df_new, 'New PO - Reroute',  1, inc_sales)
            log.append(f"✅ OLD prep: {len(df_old_prep)} rows | NEW prep: {len(df_new_prep)} rows")
            merged = pd.concat([df_old_prep, df_new_prep], ignore_index=True)
            merged = merged.sort_values(['Cust Article No.','OrderQty_fix','sort_order']).drop_duplicates(['merge_key','Remark'])
            log.append(f"✅ Merged: {len(merged)} rows")
            if 'Order Type' in merged.columns:
                merged.loc[merged['Remark'] == 'New PO - Reroute', 'Order Type'] = pd.NA
            merged = CustomerContractNormalizer.propagate_across_groups(merged)
            all_sz = sorted(set(SizeColumnHandler.get_size_columns(df_old_prep)) |
                            set(SizeColumnHandler.get_size_columns(df_new_prep)),
                            key=SizeColumnHandler.size_sort_key)
            for c in all_sz:
                if c not in merged.columns: merged[c] = pd.NA
            used_sz = [c for c in all_sz if merged[c].fillna(0).ne(0).any()]
            log.append(f"📏 Size columns: {len(used_sz)}")
            merged['Order Quantity'] = merged['OrderQty_fix']
            merged = merged.sort_values(['Cust Article No.','OrderQty_fix','sort_order']).reset_index(drop=True)
            avail  = [c for c in Config.OUTPUT_COLUMNS + used_sz if c in merged.columns]
            final  = merged[avail].copy()
            date_cols = get_date_columns(final)
            for c in date_cols: final[c] = pd.to_datetime(final[c], errors='coerce')
            log.append(f"📅 Date columns: {len(date_cols)}")
            export_df = final.copy()
            if used_sz:
                export_df[used_sz] = export_df[used_sz].where(export_df[used_sz] != 0, other=pd.NA)
                truly_empty = [c for c in used_sz if export_df[c].isna().all()]
                if truly_empty:
                    export_df.drop(columns=truly_empty, inplace=True)
                    used_sz = [c for c in used_sz if c not in truly_empty]
            compare_df = SizeComparison.create_comparison(merged, used_sz)
            only_diff, only_equal = SizeComparison.split_by_equality(compare_df)
            if not compare_df.empty:
                log.append(f"📊 Compare: {len(compare_df)} total | {len(only_diff)} diff | {len(only_equal)} equal")
            else:
                log.append("⚠️ No size comparison generated")
            excel_bytes = build_excel_bytes(export_df, compare_df, only_diff, only_equal, date_cols, used_sz)
            n_old = len(export_df[export_df['Remark'] == 'Old PO - Canceled'])
            n_new = len(export_df[export_df['Remark'] == 'New PO - Reroute'])
            st.markdown(stat_box_html({
                "Total Rows": len(export_df), "Old PO": n_old, "New PO": n_new,
                "Size Cols": len(used_sz), "Diff Articles": len(only_diff) if not only_diff.empty else 0,
            }), unsafe_allow_html=True)
            st.markdown('<div class="log-box">' + '\n'.join(log) + '</div>', unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(label="⬇️ Download Hasil Merge (.xlsx)", data=excel_bytes,
                file_name=output_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.markdown('<div class="log-box">' + '\n'.join(log) + '</div>', unsafe_allow_html=True)
            raise e


def page_old_po_filter():
    st.markdown("""
    <div class="app-header">
        <h1>📋 Input Pending Cancel</h1>
        <p>Filter baris Old PO dari file hasil merge → ekspor 2 sheet: <em>Full</em> dan <em>Summary</em>.</p>
        <span class="badge">Filter Strict · Leading Zero Safe</span>
    </div>
    """, unsafe_allow_html=True)

    step(1, "Upload File")
    uploaded = st.file_uploader("Upload file Excel (bisa multi-sheet)", type=["xlsx","xls"], key="opf_file")
    st.markdown('<hr class="sect-divider">', unsafe_allow_html=True)
    step(2, "Konfigurasi Output")
    col_a, _ = st.columns(2)
    with col_a: out_name = st.text_input("Nama file output", value="Old_PO_Filtered.xlsx")
    if not uploaded:
        st.info("Upload file untuk melanjutkan."); return

    st.markdown('<hr class="sect-divider">', unsafe_allow_html=True)
    step(3, "Proses")
    if not st.button("▶ Jalankan Filter Old PO"): return

    log = []
    with st.spinner("Memproses…"):
        try:
            file_bytes = uploaded.read()
            xls = read_excel_preserve_strings_generic(file_bytes, OLD_PO_STRING_COLS)
            log.append(f"✅ Sheet ditemukan: {list(xls.keys())}")
            all_filtered = []
            for sname, df in xls.items():
                filtered, msg = filter_old_po_strict(df)
                log.append(f"  [{sname}] {msg}")
                if not filtered.empty: all_filtered.append(filtered)
            if not all_filtered:
                st.warning("Tidak ada baris 'Old PO' ditemukan."); return
            combined = pd.concat(all_filtered, ignore_index=True)
            log.append(f"✅ Total Old PO: {len(combined)} baris")
            excel_bytes = build_old_po_excel(combined)
            st.markdown(stat_box_html({"Total Old PO": len(combined), "Sheet Input": len(xls), "Sheet Output": 2}), unsafe_allow_html=True)
            st.markdown('<div class="log-box">' + '\n'.join(log) + '</div>', unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(label="⬇️ Download Old PO (.xlsx)", data=excel_bytes,
                file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.markdown('<div class="log-box">' + '\n'.join(log) + '</div>', unsafe_allow_html=True)
            raise e


def page_email_prep():
    st.markdown("""
    <div class="app-header">
        <h1>✉️ Input Tracking Reroute Report</h1>
        <p>Reorder kolom, isi Email Subject dan Remark2 otomatis, siap untuk pengiriman ke customer.</p>
        <span class="badge">Auto-fill · Leading Zero Safe · Multi-sheet</span>
    </div>
    """, unsafe_allow_html=True)

    step(1, "Upload File")
    uploaded = st.file_uploader("Upload file Excel", type=["xlsx","xls"], key="ep_file")
    st.markdown('<hr class="sect-divider">', unsafe_allow_html=True)
    step(2, "Konfigurasi")
    col_a, col_b = st.columns(2)
    with col_a: email_subject = st.text_input("Email Subject (berlaku semua baris)", placeholder="Contoh: Reroute PO - May 2025")
    with col_b: out_name = st.text_input("Nama file output", value="Email_Prep.xlsx")

    st.markdown("""
    <div style="background:#0f1117;border:1px solid #1e2a3a;border-radius:8px;padding:12px 16px;margin:8px 0;">
        <span style="font-family:'IBM Plex Mono',monospace;font-size:0.78rem;color:#60a5fa;">Kolom output (urutan baru):</span><br>
        <span style="font-size:0.78rem;color:#94a3b8;">
        Sales Order · Remark · Sold-To PO No. · Ship-To Party PO No. · Model Name · Cust Article No. · Article · Order Quantity ·
        CRD · PD · LPD · PODD · Ship-To Search Term ·
        <strong style="color:#60a5fa;">New Ship-To Search Term (kosong)</strong> · Ship-To Country ·
        <strong style="color:#60a5fa;">Email Subject (auto)</strong> ·
        <strong style="color:#60a5fa;">Remark2 = "Full Reroute" (auto)</strong> ·
        <strong style="color:#60a5fa;">PO remark (kosong)</strong> · Prod. Status
        </span><br><br>
        <span style="font-family:'IBM Plex Mono',monospace;font-size:0.75rem;color:#475569;">
        ⚠️ Sold-To PO No. &amp; Ship-To Party PO No. dibaca sebagai teks — leading zero dipertahankan
        </span>
    </div>
    """, unsafe_allow_html=True)

    if not uploaded:
        st.info("Upload file untuk melanjutkan."); return
    if not email_subject.strip():
        st.warning("Isi Email Subject terlebih dahulu."); return

    st.markdown('<hr class="sect-divider">', unsafe_allow_html=True)
    step(3, "Proses")
    if not st.button("▶ Jalankan Email Prep"): return

    log = []
    with st.spinner("Memproses…"):
        try:
            file_bytes = uploaded.read()

            # Discover sheets
            xls_meta = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine="openpyxl", nrows=0)
            sheet_names = list(xls_meta.keys())
            log.append(f"✅ Sheet ditemukan: {sheet_names}")

            xls = {}
            for sname in sheet_names:
                # Peek at headers to know which string cols exist in this sheet
                df_peek = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sname, engine="openpyxl", nrows=0)
                df_peek.columns = [str(c).strip() for c in df_peek.columns]
                converters = {col: str for col in EMAIL_PREP_STRING_COLS if col in df_peek.columns}

                df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sname,
                                   engine="openpyxl", converters=converters)
                df.columns = [str(c).strip() for c in df.columns]

                # Clean PO string cols
                for col in EMAIL_PREP_STRING_COLS:
                    if col in df.columns:
                        df[col] = _clean_po_string(df[col])

                xls[sname] = df

            sheets_out = {}
            total_rows = 0
            for sname, df in xls.items():
                out_df = email_prep_process(df, email_subject.strip())
                sheets_out[sname] = out_df
                total_rows += len(out_df)
                missing = [c for c in EMAIL_PREP_TARGET
                           if c not in df.columns
                           and c not in ("Email Subject","Remark2","Prod. Status",
                                         "New Ship-To Search Term","PO remark")]
                log.append(f"  [{sname}] {len(out_df)} rows | kolom kosong ditambahkan: {missing if missing else '-'}")

            excel_bytes = build_email_prep_excel(sheets_out)

            st.markdown(stat_box_html({
                "Total Rows": total_rows, "Sheet": len(sheets_out), "Kolom Output": len(EMAIL_PREP_TARGET),
            }), unsafe_allow_html=True)
            st.markdown('<div class="log-box">' + '\n'.join(log) + '</div>', unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(label="⬇️ Download Email Prep (.xlsx)", data=excel_bytes,
                file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.markdown('<div class="log-box">' + '\n'.join(log) + '</div>', unsafe_allow_html=True)
            raise e


# ============================================================
#  SIDEBAR + ROUTING
# ============================================================

with st.sidebar:
    st.markdown("""
    <div style="padding:20px 0 12px 0;">
        <span style="font-family:'IBM Plex Mono',monospace;font-size:1.1rem;font-weight:600;color:#60a5fa;">🔄 Reroute Suite</span>
        <br><span style="font-size:0.72rem;color:#475569;">v2.0 Unified</span>
    </div>
    <hr style="border:none;border-top:1px solid #1e2a3a;margin:0 0 16px 0;">
    """, unsafe_allow_html=True)

    page = st.radio(
        "Navigasi",
        ["🏠  Home", "🔀  Reroute Merge", "📋  Input Pending Cancel", "✉️  Input Tracking Reroute Report"],
        label_visibility="collapsed",
    )

    st.markdown("""
    <hr style="border:none;border-top:1px solid #1e2a3a;margin:20px 0 12px 0;">
    <div style="font-size:0.72rem;color:#334155;line-height:1.7;padding-bottom:8px;">
        <strong style="color:#475569;">Urutan Workflow:</strong><br>
        1 → Reroute Merge<br>
        2 → Input Pending Cancel<br>
        3 → Input Tracking Reroute Report
    </div>
    """, unsafe_allow_html=True)


if   "Home"                         in page: page_home()
elif "Reroute Merge"                 in page: page_reroute_merge()
elif "Input Pending Cancel"          in page: page_old_po_filter()
elif "Input Tracking Reroute Report" in page: page_email_prep()
