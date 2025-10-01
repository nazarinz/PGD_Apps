# Without SO - Reroute Old PO - New PO Tools (FIXED)
# =======================================================
# Reroute Merge (UPLOAD FILE) + Size Comparison Old vs New
# REVISI:
#  - Order Type baris New dikosongkan
#  - Customer Contract ID disamakan per (Article, Qty) dan ditulis tanpa ".0"
#  - Perbaikan NameError (variabel df tak terdefinisi)
#  - Pembatasan scope per sub-tools (tiap blok di dalam if tool == ...)
#  - Perbaikan indentasi bagian download compare_df
#  - Konsistensi pemakaian variabel (df_baru/df_reroute)
# =======================================================

import io, re
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
from utils import set_page, header, footer, write_excel_autofit

# ================== Page & Header ==================
set_page("PGD Apps — Reroute Tools", "🔁")
header("🔁 Reroute Tools")

# ================== Select Sub-Tool ==================
tool = st.selectbox(
    "Pilih sub-tools",
    [
        "Merge Reroute: Old vs New (size compare)",
        "Cek Konsistensi Ukuran per Article",
        "Old PO Finder (Batch)",
    ],
)

# ================== Util & Helpers (shared) ==================
SIZE_PAT = re.compile(r'^(?:UK[_\-\s]*)?(\d{1,2})(K|-K|-)?$', flags=re.I)

def canonical_size_name(col: str) -> str:
    s = re.sub(r"\s+", "", str(col)).upper()
    s = s.replace("UK-", "UK_").replace("UK", "UK_").replace("__", "_")
    m = SIZE_PAT.match(s.replace("_", ""))
    if m:
        base = m.group(1); suf = m.group(2) or ''
        return f'UK_{base}{suf}'
    m2 = re.match(r'^UK_(\d{1,2})(K|-K|-)?$', s)
    if m2:
        return f"UK_{m2.group(1)}{m2.group(2) or ''}"
    return col

def normalize_size_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={c: canonical_size_name(c) for c in df.columns})

def to_num(x):
    if isinstance(x, str):
        x = x.replace(",", "")
    return pd.to_numeric(x, errors='coerce')

def compute_sizes_and_qty(df: pd.DataFrame) -> pd.DataFrame:
    size_cols = [c for c in df.columns if re.match(r'^UK_\d{1,2}(K|-K|-)?$', str(c))]
    if size_cols:
        df[size_cols] = df[size_cols].apply(to_num).fillna(0)
        df['SizeSum'] = df[size_cols].sum(axis=1).astype('Int64')
    else:
        df['SizeSum'] = pd.NA
    oq = to_num(df.get('Order Quantity'))
    df['OrderQty_fix'] = oq
    need_fix = df['SizeSum'].notna() & (df['OrderQty_fix'].isna() | (df['OrderQty_fix'] != df['SizeSum']))
    df.loc[need_fix, 'OrderQty_fix'] = df.loc[need_fix, 'SizeSum']
    df['OrderQty_fix'] = df['OrderQty_fix'].fillna(0).astype('Int64')
    return df

def norm_article_col(df: pd.DataFrame) -> pd.DataFrame:
    if 'Cust Article No.' not in df.columns:
        raise ValueError("Kolom 'Cust Article No.' tidak ditemukan.")
    df['Cust Article No.'] = df['Cust Article No.'].astype(str).str.strip().str.upper()
    return df

def add_merge_key(df: pd.DataFrame) -> pd.DataFrame:
    return df.assign(merge_key=df['Cust Article No.'].astype(str) + '|' + df['OrderQty_fix'].astype(str))

def aggregate_side(df: pd.DataFrame) -> pd.DataFrame:
    size_cols  = [c for c in df.columns if re.match(r'^UK_\d{1,2}(K|-K|-)?$', str(c))]
    meta_cols  = [c for c in df.columns if c not in size_cols and c != 'merge_key']
    agg = {c:'sum' for c in size_cols}
    agg.update({c:'first' for c in meta_cols})
    out = df.groupby('merge_key', as_index=False).agg(agg)
    if size_cols:
        out['SizeSum'] = out[size_cols].sum(axis=1).astype('Int64')
    return out

def norm_cc(v):
    if pd.isna(v):
        return pd.NA
    s = str(v).strip().replace(',', '')
    if s.lower() in {'', 'nan', 'none'}:
        return pd.NA
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    if re.fullmatch(r'\d+\.0', s):
        return s[:-2]
    return s

def prep_side(df: pd.DataFrame, remark: str, sort_order: int) -> pd.DataFrame:
    df = normalize_size_columns(df.copy())
    df = norm_article_col(df)
    df = compute_sizes_and_qty(df)
    df = add_merge_key(df)
    df['Remark'] = remark
    df['sort_order'] = sort_order
    df = aggregate_side(df)
    return df

# =======================================================
# 1) Merge Reroute: Old vs New (size compare)
# =======================================================
if tool == "Merge Reroute: Old vs New (size compare)":
    colA, colB = st.columns(2)
    with colA:
        new_po_file = st.file_uploader("Upload New PO (.xlsx)", type=["xlsx","xlsm","xls"], accept_multiple_files=False, key="rr_new")
    with colB:
        old_po_file = st.file_uploader("Upload Old PO (.xlsx)", type=["xlsx","xlsm","xls"], accept_multiple_files=False, key="rr_old")

    if not new_po_file or not old_po_file:
        st.info("Silakan upload kedua file untuk mulai.")
        st.stop()

    # === Read
    df_baru = pd.read_excel(new_po_file)
    df_reroute = pd.read_excel(old_po_file)

    # === Normalisasi header untuk file New (kalau beda nama kolom)
    df_baru.rename(columns={
        'PO Batch Date': 'PO Date',
        'PO Number': 'Sold-To PO No.',
        'Market PO Number': 'Ship-To Party PO No.',
        'Customer Request Date (CRD)': 'CRD',
        'Plan Date': 'PD',
        'Article Number': 'Cust Article No.',
        'Gps Customer Number': 'Ship-To Search Term',
        'Class Code': 'Classification Code',
        'Grand Total': 'Order Quantity'
    }, inplace=True)

    # === Build merged (1 baris per sisi)
    df_reroute_p = prep_side(df_reroute, 'Old PO - Canceled', 0)
    df_baru_p    = prep_side(df_baru,    'New PO - Reroute', 1)
    merged = pd.concat([df_reroute_p, df_baru_p], ignore_index=True).sort_values(['merge_key','sort_order'])
    merged = merged.drop_duplicates(['merge_key','Remark'])

    # === REVISI 1: Order Type baris New dikosongkan
    if 'Order Type' in merged.columns:
        merged.loc[merged['Remark'].eq('New PO - Reroute'), 'Order Type'] = pd.NA

    # === REVISI 2: Samakan Customer Contract ID per (Article, Qty) & hapus ".0"
    if 'Customer Contract ID' in merged.columns:
        merged['Customer Contract ID'] = merged['Customer Contract ID'].apply(norm_cc)
        cc_map = (
            merged.sort_values(['Cust Article No.','OrderQty_fix','sort_order'])
                  .groupby(['Cust Article No.','OrderQty_fix'])['Customer Contract ID']
                  .apply(lambda s: next((x for x in s if pd.notna(x)), pd.NA))
                  .rename('CC_fill')
        )
        merged = merged.merge(cc_map, on=['Cust Article No.','OrderQty_fix'], how='left')
        merged['Customer Contract ID'] = merged['CC_fill'].apply(norm_cc)
        merged.drop(columns=['CC_fill'], inplace=True)

    # === Lengkapi kolom size
    uk_cols_r = [c for c in df_reroute_p.columns if c.startswith("UK_")]
    uk_cols_n = [c for c in df_baru_p.columns    if c.startswith("UK_")]
    all_uk_cols = sorted(set(uk_cols_r) | set(uk_cols_n))
    for c in all_uk_cols:
        if c not in merged.columns:
            merged[c] = pd.NA

    # === Susun output
    merged['Order Quantity'] = merged['OrderQty_fix']
    kolom_output = [
        'Remark','Select','Status','Working Status','Working Status Descr.',
        'PO Date','Requirement Segment','Order Type','Site','Work Center',
        'Sales Order','Customer Contract ID','Sold-To PO No.','Ship-To Party PO No.',
        'CRD','PD','Prod. Team ATP','FPD','FPD-DRC','POSDD','POSDD-DRC',
        'LPD','LPD-DRC','PODD','PODD-DRC','FGR','Cust Article No.','Model Name',
        'Gender','Article','Article Lead Time','Develop Type','Last Code','Season',
        'Product Hierarchy 3','Outsole Mold','Pattern Code (Upper','Ship-To No.',
        'Ship-To Search Term','Ship-To Name','Ship-To Country','Shipping Type',
        'Packing Type','VAS Cut-Off Date','Classification Code','Changed By',
        'Document Date','Order Quantity'
    ] + all_uk_cols

    kolom_tersedia = [k for k in kolom_output if k in merged.columns]
    final_output = merged[kolom_tersedia].copy()

    # tanggal → datetime
    date_cols = [c for c in [
        'PO Date','CRD','PD','FPD','FPD-DRC','POSDD','POSDD-DRC',
        'LPD','LPD-DRC','PODD','PODD-DRC','VAS Cut-Off Date','Document Date'
    ] if c in final_output.columns]
    for c in date_cols:
        final_output[c] = pd.to_datetime(final_output[c], errors='coerce')

    # Kosongkan nilai size==0 pada tampilan
    size_cols_all = [c for c in final_output.columns if re.match(r'^UK_\d{1,2}(K|-K|-)?$', str(c))]
    export_df = final_output.copy()
    if size_cols_all:
        export_df[size_cols_all] = export_df[size_cols_all].where(export_df[size_cols_all] != 0, other=pd.NA)

    # ======= Perbandingan size Old vs New =======
    cmp_base = merged[['merge_key','Remark','Cust Article No.','OrderQty_fix'] + all_uk_cols].copy()
    cmp_base[all_uk_cols] = cmp_base[all_uk_cols].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    wide = cmp_base.pivot_table(index=['merge_key','Cust Article No.','OrderQty_fix'],
                                columns='Remark', values=all_uk_cols, aggfunc='first')

    compare_df = pd.DataFrame()
    only_diff = pd.DataFrame()
    only_equal = pd.DataFrame()
    if ('Old PO - Canceled' in wide.columns.get_level_values(1)) and ('New PO - Reroute' in wide.columns.get_level_values(1)):
        old = wide.xs('Old PO - Canceled', axis=1, level=1).reindex(columns=all_uk_cols)
        new = wide.xs('New PO - Reroute',  axis=1, level=1).reindex(columns=all_uk_cols)
        diff = new - old
        summary = pd.DataFrame({
            'Cust Article No.': [i[1] for i in new.index],
            'OrderQty_fix':     [i[2] for i in new.index],
            'All_Sizes_Equal':  diff.eq(0).all(axis=1).values,
            'Diff_Count':       diff.ne(0).sum(axis=1).values
        }, index=new.index)
        def add_suffix(df, suf):
            df = df.copy(); df.columns=[f"{c}__{suf}" for c in df.columns]; return df
        compare_df = pd.concat([
            summary.reset_index(drop=True),
            add_suffix(old.reset_index(drop=True),"old"),
            add_suffix(new.reset_index(drop=True),"new"),
            add_suffix(diff.reset_index(drop=True),"diff")
        ], axis=1)
        only_diff  = compare_df[~compare_df['All_Sizes_Equal']].reset_index(drop=True)
        only_equal = compare_df[ compare_df['All_Sizes_Equal']].reset_index(drop=True)

    # ======= Download (header formatting + old rows red) =======
    st.subheader("📥 Download Hasil")

    def _export_with_format(main_df: pd.DataFrame,
                            date_cols: list[str],
                            compare_df: pd.DataFrame,
                            only_diff: pd.DataFrame,
                            only_equal: pd.DataFrame) -> bytes:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yy") as writer:
            # --- write main sheet ---
            main_df.to_excel(writer, index=False, sheet_name="Sheet1")
            wb  = writer.book
            ws  = writer.sheets["Sheet1"]

            # Formats
            hdr_gray   = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#D9D9D9", "border": 1})
            hdr_yellow = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#FFF2CC", "border": 1})
            txt_fmt    = wb.add_format({"num_format": "@"})
            date_fmt   = wb.add_format({"num_format": "m/d/yy"})
            red_font   = wb.add_format({"font_color": "#FF0000"})

            # Header coloring (row 0)
            for c_idx, col_name in enumerate(main_df.columns):
                fmt = hdr_yellow if col_name == "Remark" else hdr_gray
                ws.write(0, c_idx, col_name, fmt)

            # Set text format for CC ID to avoid .0/scientific
            if 'Customer Contract ID' in main_df.columns:
                i_cc = main_df.columns.get_loc('Customer Contract ID')
                ws.set_column(i_cc, i_cc, 18, txt_fmt)

            # Date columns width + format
            for dc in date_cols:
                if dc in main_df.columns:
                    i = main_df.columns.get_loc(dc)
                    ws.set_column(i, i, 12, date_fmt)

            # Simple autofit & usability
            ws.autofilter(0, 0, len(main_df), len(main_df.columns)-1)
            ws.freeze_panes(1, 0)
            for col_idx, col_name in enumerate(main_df.columns):
                maxlen = max([len(str(col_name))] + [len(str(x)) for x in main_df[col_name].head(1000).fillna("").astype(str)])
                ws.set_column(col_idx, col_idx, min(45, max(10, maxlen + 2)))

            # Conditional font red for rows with Remark == 'Old PO - Canceled'
            if 'Remark' in main_df.columns:
                i_remark = main_df.columns.get_loc('Remark')
                last_row = len(main_df) + 1
                last_col = len(main_df.columns)
                # Build A1-style range for full table excluding header
                def col_letter(n):
                    s = ""
                    while n >= 0:
                        s = chr(n % 26 + 65) + s
                        n = n // 26 - 1
                    return s
                rng = f"A2:{col_letter(last_col-1)}{last_row}"
                # Use formula anchored to Remark column
                first_data_row = 2
                # Example formula uses the column letter for Remark
                rem_col_letter = col_letter(i_remark)
                ws.conditional_format(rng, {
                    'type': 'formula',
                    'criteria': f'=${rem_col_letter}{first_data_row}="Old PO - Canceled"',
                    'format': red_font
                })

            # --- optional compare sheets ---
            if not compare_df.empty:
                compare_df.to_excel(writer, index=False, sheet_name="Size_Compare")
                only_diff.to_excel(writer, index=False, sheet_name="Only_Different")
                only_equal.to_excel(writer, index=False, sheet_name="Only_Equal")
        return output.getvalue()

    payload = _export_with_format(export_df, date_cols, compare_df, only_diff, only_equal)
    st.success("Selesai diproses.")
    st.download_button(
        label="⬇️ Download Excel",
        data=payload,
        file_name="Hasil_Format_Dua_Baris.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# =======================================================
# 2) CEK KONSISTENSI UKURAN PER ARTICLE NUMBER
# =======================================================
if tool == "Cek Konsistensi Ukuran per Article":
    header("Cek Konsistensi Ukuran per Article Number")
    file_cons = st.file_uploader(
        "Upload file Excel (berisi Article Number, Grand Total, kolom size)",
        type=["xlsx","xlsm","xls"],
        key="cons",
    )
    if not file_cons:
        st.stop()
    df = pd.read_excel(file_cons)

    # --------- Konfigurasi kolom ---------
    ART_COL   = "Article Number"
    QTY_COL   = "Grand Total"

    # Daftar kandidat kolom ukuran sesuai contohmu:
    CAND_SIZE_COLS = [
        "3K","4K","5-K","6-K","7-K","8-K","9-K",
        "3-","4-","5-","6-","7-",
        "8","8-","9","9-","10","10-","11","11-","12-","13-"
    ]
    SIZE_COLS = [c for c in CAND_SIZE_COLS if c in df.columns]

    if not SIZE_COLS:
        st.error("Kolom ukuran tidak ditemukan. Pastikan nama kolom sesuai (3K, 4K, 5-K, ..., 13-).")
        st.stop()

    # --------- Normalisasi data ---------
    df[ART_COL] = df[ART_COL].astype(str).str.strip().str.upper()
    df[QTY_COL] = df[QTY_COL].apply(to_num)
    df[SIZE_COLS] = df[SIZE_COLS].apply(pd.to_numeric, errors="coerce").fillna(0).astype("Int64")

    # Hitung total ukuran per baris (opsional untuk validasi)
    df["SizeSum"] = df[SIZE_COLS].sum(axis=1).astype("Int64")
    df["Qty_Match"] = (df["SizeSum"] == df[QTY_COL]).fillna(False)

    # --------- Cek konsistensi pola ukuran ---------
    df["_pattern"] = df[SIZE_COLS].apply(lambda r: tuple(int(x) if pd.notna(x) else 0 for x in r), axis=1)
    pat_counts = df.groupby(ART_COL)["_pattern"].nunique().reset_index(name="Unique_Size_Patterns")
    pat_counts["Consistent"] = pat_counts["Unique_Size_Patterns"].eq(1)

    bad_arts = pat_counts[pat_counts["Unique_Size_Patterns"] > 1][ART_COL].tolist()

    def sample_patterns(g: pd.DataFrame) -> pd.DataFrame:
        out = (
            g.drop_duplicates(subset=["_pattern"])[[ART_COL, QTY_COL] + SIZE_COLS]
             .reset_index(drop=True)
        )
        out.insert(1, "Pattern_ID", range(1, len(out)+1))
        return out

    mismatch_df = pd.DataFrame()
    if bad_arts:
        mismatch_df = (
            df[df[ART_COL].isin(bad_arts)]
            .groupby(ART_COL, group_keys=True)
            .apply(sample_patterns)
            .reset_index(drop=True)
        )

    payload2 = write_excel_autofit({
        "Summary": pat_counts,
        "Mismatch_Samples": mismatch_df if not mismatch_df.empty else pd.DataFrame(),
        "Data_Normalized": df[[ART_COL, QTY_COL] + SIZE_COLS + ["SizeSum","Qty_Match"]],
    })
    st.download_button(
        "⬇️ Download Size Consistency Report",
        data=payload2,
        file_name="Size_Consistency_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# =======================================================
# 3) Old PO Finder (Batch)
# =======================================================
if tool == "Old PO Finder (Batch)":
    header("Old PO Finder - Reroute (Batch)")
    import os

    sap_file = st.file_uploader(
        "Upload file SAP (ZRSD1013 .xlsx/.csv)",
        type=["xlsx","xlsm","xls","csv"],
        key="sap",
    )
    if not sap_file:
        st.stop()

    st.caption("Masukkan pasangan (Article No, Quantity)")
    pairs_text = st.text_area("Pairs (satu per baris, format: ART, QTY)", value="GW4140, 620\nID5273, 496")

    def parse_pairs(txt: str):
        out = []
        for ln in txt.splitlines():
            ln = ln.strip()
            if not ln:
                continue
            try:
                a, q = ln.split(",")
                out.append((a.strip(), int(q)))
            except Exception:
                pass
        return out

    pairs = parse_pairs(pairs_text)
    if not pairs:
        st.warning("Tidak ada pasangan valid.")
        st.stop()

    def read_any(uploaded) -> pd.DataFrame:
        name = getattr(uploaded, 'name', '').lower()
        if name.endswith((".xlsx", ".xlsm", ".xls")):
            return pd.read_excel(uploaded, engine="openpyxl")
        if name.endswith(".csv"):
            return pd.read_csv(uploaded, sep=None, engine="python")
        raise ValueError(f"Format file tidak didukung: {name}")

    def normalize_sap(df: pd.DataFrame) -> pd.DataFrame:
        if "Quanity" in df.columns and "Quantity" not in df.columns:
            df = df.rename(columns={"Quanity": "Quantity"})
        needed = {"PO No.(Full)", "Article No", "Quantity", "FCR Date"}
        missing = needed - set(df.columns)
        if missing:
            raise ValueError(f"Kolom wajib hilang: {missing}")
        df["PO No.(Full)"] = df["PO No.(Full)"].astype(str).str.strip()
        df["Article No"]   = df["Article No"].astype(str).str.strip()
        df["Quantity"]     = pd.to_numeric(df["Quantity"], errors="coerce").astype("Int64")
        if not np.issubdtype(df["FCR Date"].dtype, np.datetime64):
            df["FCR Date"] = pd.to_datetime(df["FCR Date"], errors="coerce")
        for dcol in ["Document Date","FPD","LPD","CRD","PSDD","PODD","PD","PO Date","Actual PGI"]:
            if dcol in df.columns and not np.issubdtype(df[dcol].dtype, np.datetime64):
                df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
        return df

    def find_po(df: pd.DataFrame, article_no: str, quantity: int) -> pd.DataFrame:
        mask = (
            df["Article No"].eq(str(article_no).strip()) &
            df["Quantity"].eq(int(quantity)) &
            df["FCR Date"].isna()
        )
        cols_show = [c for c in [
            "PO No.(Full)", "Article No", "Model Name", "Quantity",
            "CRD", "PSDD", "PD", "PO Date", "Document Date", "FCR Date",
            "FPD", "LPD", "PODD", "Actual PGI"
        ] if c in df.columns]
        out = df.loc[mask, cols_show].copy()
        sort_cols = [c for c in ["CRD","PSDD","PO Date","Document Date"] if c in out.columns]
        if sort_cols:
            out = out.sort_values(by=sort_cols, ascending=True, kind="stable")
        out = out.drop_duplicates(subset=["PO No.(Full)"], keep="first").reset_index(drop=True)
        return out

    df_sap = normalize_sap(read_any(sap_file))

    all_results = []
    for i, (art, qty) in enumerate(pairs, 1):
        match = find_po(df_sap, art, qty)
        if match.empty:
            all_results.append(pd.DataFrame({
                "Seq": [i],
                "Cari_Article": [art],
                "Cari_Quantity": [qty],
                "Match_Count": [0],
                "PO No.(Full)": [pd.NA],
                "Keterangan": ["Tidak ditemukan (FCR Date kosong)"]
            }))
        else:
            match.insert(0, "Seq", i)
            match.insert(1, "Cari_Article", art)
            match.insert(2, "Cari_Quantity", qty)
            match.insert(3, "Match_Count", len(match))
            all_results.append(match)

    result = pd.concat(all_results, ignore_index=True)
    ringkas = (
        result.groupby(["Seq","Cari_Article","Cari_Quantity"], dropna=False)
              .agg(Matches=("PO No.(Full)", lambda s: s.dropna().nunique()))
              .reset_index()
              .sort_values("Seq")
    )

    payload3 = write_excel_autofit({
        "Summary": ringkas,
        "Matches": result,
    })

    st.download_button(
        "⬇️ Download PO Finder Batch",
        data=payload3,
        file_name=f"PO_Finder_Batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

footer()
