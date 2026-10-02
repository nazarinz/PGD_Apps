import io
import re
import itertools
import datetime

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# HELPER
# =========================================================
def safe_parse_date(val):
    """Parse tanggal dari Timestamp / serial Excel / string M/D/YYYY."""
    if pd.isna(val):
        return pd.NaT

    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.date)):
        return pd.Timestamp(val)

    if isinstance(val, (int, float, np.integer, np.floating)):
        try:
            return pd.to_datetime(val, unit="D", origin="1899-12-30")
        except Exception:
            return pd.NaT

    s = str(val).strip()
    try:
        return pd.to_datetime(s, format="%m/%d/%Y", errors="raise")
    except (ValueError, TypeError):
        pass
    try:
        return pd.to_datetime(s, errors="coerce", dayfirst=False)
    except Exception:
        return pd.NaT


def is_empty_val(val):
    """Nilai dianggap kosong: NaN, string kosong, atau 'nan'."""
    if pd.isna(val):
        return True
    if str(val).strip() == "" or str(val).strip().lower() == "nan":
        return True
    return False


def get_day(date_val):
    parsed = safe_parse_date(date_val)
    if pd.isna(parsed):
        return np.nan
    return parsed.day


def get_month_period(day_value):
    if pd.isna(day_value):
        return np.nan
    return "1st" if day_value <= 15 else "2nd"


def move_columns_after(df, target_col, cols_to_insert, warn):
    cols = list(df.columns)
    for c in cols_to_insert:
        if c in cols:
            cols.remove(c)
    if target_col in cols:
        idx = cols.index(target_col)
        for i, c in enumerate(cols_to_insert):
            cols.insert(idx + 1 + i, c)
    else:
        warn(f"Kolom '{target_col}' tidak ditemukan untuk penyusunan ulang.")
    return df[cols]


def get_date_diff(podd_val, compare_val):
    if is_empty_val(compare_val):
        return "Not Yet Exported"
    podd_parsed = safe_parse_date(podd_val)
    compare_parsed = safe_parse_date(compare_val)
    if pd.isna(podd_parsed) or pd.isna(compare_parsed):
        return np.nan
    return (podd_parsed - compare_parsed).days


def get_dol(podd_val, crd_val):
    podd_parsed = safe_parse_date(podd_val)
    crd_parsed = safe_parse_date(crd_val)
    if pd.isna(podd_parsed) or pd.isna(crd_parsed):
        return np.nan
    diff = (podd_parsed - crd_parsed).days
    if diff <= 0:
        return np.nan
    elif diff <= 7:
        return "DOL7"
    elif diff <= 15:
        return "DOL15"
    elif diff <= 30:
        return "DOL30"
    else:
        return "DOL45"


def combine_unique_values(series):
    """Gabungkan nilai unik yang terisi, dipisah ' | '."""
    seen = []
    for v in series:
        if is_empty_val(v) or str(v).strip().lower() == "null":
            continue
        v_str = str(v).strip()
        if v_str not in seen:
            seen.append(v_str)
    if not seen:
        return np.nan
    return " | ".join(seen)


MAX_SUBSET_SIZE = 4
TOL = 1e-9


def crd_gap_days(sap_crd, dash_crds):
    """Total selisih hari (absolut) antara CRD SAP dan CRD Dashboard terpilih."""
    if pd.isna(sap_crd):
        return 0
    total = 0
    for d in dash_crds:
        total += 9999 if pd.isna(d) else abs((pd.Timestamp(sap_crd) - d).days)
    return total


def best_matching_subset(remaining_idx, qty_map, target, sap_crd=None, crd_map=None):
    """Cari kombinasi baris Dashboard yang paling mirip dengan baris SAP.
    Prioritas: (1) selisih qty terkecil, (2) selisih CRD terkecil, (3) jumlah baris paling sedikit.
    Return (subset terbaik, selisih qty)."""
    best_subset, best_key = None, None
    max_size = min(MAX_SUBSET_SIZE, len(remaining_idx))
    for size in range(1, max_size + 1):
        for combo in itertools.combinations(remaining_idx, size):
            total = sum(qty_map[i] for i in combo)
            diff = abs(total - target)
            gap = (
                crd_gap_days(sap_crd, [crd_map.get(i, pd.NaT) for i in combo])
                if crd_map else 0
            )
            key = (diff, gap)
            if best_key is None or key < best_key:
                best_subset, best_key = combo, key
        if best_key is not None and best_key[0] <= TOL:
            break

    # Kalau belum ada yang pas, tapi gabungan SEMUA sisa baris Dashboard = qty SAP
    if (best_key is None or best_key[0] > TOL) and len(remaining_idx) > max_size:
        if abs(sum(qty_map[i] for i in remaining_idx) - target) <= TOL:
            return tuple(remaining_idx), 0

    return best_subset, (best_key[0] if best_key else None)


def true_false_to_y(val):
    if is_empty_val(val):
        return np.nan
    parts = [p.strip() for p in str(val).split(" | ")]
    mapped = []
    for p in parts:
        pu = p.upper()
        if pu == "TRUE":
            mapped.append("Y")
        elif pu == "FALSE":
            mapped.append("")
        else:
            mapped.append(p)
    return " | ".join(mapped)


def parse_model_list(text):
    """Ubah teks paste (satu Model No per baris) jadi list unik, huruf besar."""
    if not text:
        return []
    parts = re.split(r"[\n\r\t,;]+", text)
    seen, out = set(), []
    for p in parts:
        m = p.strip().upper()
        if m and m not in seen:
            seen.add(m)
            out.append(m)
    return out


def read_excel_file(uploaded):
    """Baca file upload (.xlsx / .xlsb / .xls) ke DataFrame (sheet pertama)."""
    name = uploaded.name.lower()
    uploaded.seek(0)
    if name.endswith(".xlsb"):
        return pd.read_excel(uploaded, engine="pyxlsb")
    return pd.read_excel(uploaded)


# =========================================================
# PIPELINE UTAMA
# =========================================================
def clean_data(df, df_dash, cpr_models, log):
    warn = log.append

    qty_mismatch_report = []
    dash_num_map = {}
    dash_date_map = {}
    df_dash_only = pd.DataFrame()
    available_dash_cols = []
    qty_col_dash = None

    # ---------- 2. Rapikan nama kolom ----------
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    df.columns = [c.replace("Quanity", "Quantity") for c in df.columns]

    # ---------- 3. Bersihkan spasi pada data teks ----------
    text_cols = df.select_dtypes(include=["object", "string"]).columns
    for col in text_cols:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": np.nan, "None": np.nan, "": np.nan})

    # ---------- 4. Quantity & Unit Price -> numerik ----------
    for col in ["Quantity", "Unit Price"]:
        if col in df.columns:
            df[col] = (
                df[col].astype(str).str.replace(",", "", regex=False).str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # ---------- 5. Kolom tanggal -> tipe date asli ----------
    date_cols = [
        "Document Date", "FPD", "LPD", "CRD", "PSDD",
        "FCR Date", "PODD", "PD", "PO Date", "Actual PGI", "S&P LPD",
    ]
    for col in date_cols:
        if col in df.columns:
            df[col] = df[col].apply(safe_parse_date)

    # ---------- 6. Remark (sebelum diisi) ----------
    check_cols = ["FPD", "LPD", "PSDD", "PODD"]
    existing_check_cols = [c for c in check_cols if c in df.columns]
    unconfirmed = df[existing_check_cols].isna().any(axis=1)
    df["Remark"] = np.where(unconfirmed, "Unconfirmed Order", "")
    df["Remark"] = df["Remark"].replace("", np.nan)

    if "SO" in df.columns:
        cols = list(df.columns)
        cols.remove("Remark")
        cols.insert(cols.index("SO"), "Remark")
        df = df[cols]
    else:
        warn("Kolom 'SO' tidak ditemukan, 'Remark' ditambahkan di akhir.")

    # ---------- 6b. Isi kosong FPD/LPD/PSDD/PODD dengan CRD ----------
    if "CRD" in df.columns:
        for c in ["FPD", "LPD", "PSDD", "PODD"]:
            if c in df.columns:
                df[c] = df[c].fillna(df["CRD"])
    else:
        warn("Kolom 'CRD' tidak ditemukan, FPD/LPD/PSDD/PODD tidak bisa diisi.")

    # ---------- 6c. Day - X & Month - X ----------
    for col in ["LPD", "CRD", "PODD"]:
        if col in df.columns:
            df[f"Day - {col}"] = df[col].apply(get_day)
            df[f"Month - {col}"] = df[f"Day - {col}"].apply(get_month_period)
        else:
            warn(f"Kolom '{col}' tidak ditemukan, Day/Month tidak dibuat.")

    # ---------- 6c2. CRD Monthly & 6c3. Triggering ----------
    if "CRD" in df.columns:
        df["CRD Monthly"] = df["CRD"].dt.strftime("%Y%m")
        df["Triggering"] = np.where(
            df["CRD"].dt.month <= 6,
            df["CRD"].dt.year.astype("Int64").astype(str) + "-H1",
            df["CRD"].dt.year.astype("Int64").astype(str) + "-H2",
        )
    else:
        warn("Kolom 'CRD' tidak ditemukan, CRD Monthly/Triggering tidak dibuat.")

    # ---------- 6d. Posisi Day/Month setelah kolom tanggal ----------
    if "LPD" in df.columns:
        df = move_columns_after(df, "LPD", ["Day - LPD", "Month - LPD"], warn)
    if "CRD" in df.columns:
        df = move_columns_after(df, "CRD", ["Day - CRD", "Month - CRD"], warn)
    if "PODD" in df.columns:
        df = move_columns_after(df, "PODD", ["Day - PODD", "Month - PODD"], warn)

    # ---------- 6e. PODD vs FCR / Actual PGI ----------
    if "PODD" in df.columns and "FCR Date" in df.columns:
        df["PODD vs FCR"] = df.apply(
            lambda r: get_date_diff(r["PODD"], r["FCR Date"]), axis=1
        )
    else:
        warn("Kolom 'PODD' dan/atau 'FCR Date' tidak ditemukan.")

    if "PODD" in df.columns and "Actual PGI" in df.columns:
        df["PODD vs Actual PGI"] = df.apply(
            lambda r: get_date_diff(r["PODD"], r["Actual PGI"]), axis=1
        )
    else:
        warn("Kolom 'PODD' dan/atau 'Actual PGI' tidak ditemukan.")

    # =====================================================
    # 6f. LOOKUP DASHBOARD (berdasarkan PO)
    # =====================================================
    if df_dash is None:
        warn("File Dashboard tidak diupload. Lookup Dashboard dilewati.")
    else:
        df_dash = df_dash.copy()
        df_dash.columns = (
            df_dash.columns.astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )
        df_dash_raw = df_dash.copy()

        dash_cols_to_pull = [
            "is_SPOMA", "is_Product Prio", "ReAct", "Chase", "is_Key Franchise",
            "Elevated Check", "Campaign Name", "Bran Partner Description",
            "Responsiveness", "Purchase Order Status Desc",
            "MDP Status Adjusted", "SDP Status Adjusted",
        ]
        missing_dash_cols = [c for c in dash_cols_to_pull if c not in df_dash.columns]
        if missing_dash_cols:
            warn(f"Kolom berikut tidak ditemukan di Dashboard: {missing_dash_cols}")
        available_dash_cols = [c for c in dash_cols_to_pull if c in df_dash.columns]

        if "PO" not in df_dash.columns:
            warn("Kolom 'PO' tidak ditemukan di Dashboard, lookup dilewati.")
        elif "PO No.(Full)" not in df.columns:
            warn("Kolom 'PO No.(Full)' tidak ditemukan di ZRSD1013, lookup dilewati.")
        else:
            df["PO No.(Full)"] = df["PO No.(Full)"].astype(str).str.strip()
            df_dash["PO"] = df_dash["PO"].astype(str).str.strip()

            # Anomali: PO Dashboard berawalan "M" seharusnya "0"
            anomaly_mask = df_dash["PO"].str.startswith("M")
            n_anomaly = int(anomaly_mask.sum())
            if n_anomaly > 0:
                warn(f"{n_anomaly} PO di Dashboard berawalan 'M' dikoreksi jadi '0'.")
                df_dash.loc[anomaly_mask, "PO"] = "0" + df_dash.loc[anomaly_mask, "PO"].str[1:]

            df_dash_dedup = (
                df_dash.groupby("PO", as_index=False)[available_dash_cols]
                .agg(combine_unique_values)
            )
            lookup_table = df_dash_dedup.rename(columns={"PO": "PO No.(Full)"})

            # --- Kolom numerik Dashboard (dijumlah per PO) ---
            qty_col_candidates = ["PORD Order Qty.Key Date 1", "GR Qty", "PORD Open Qty.Key Date 1"]
            qty_col_dash = next((c for c in qty_col_candidates if c in df_dash.columns), None)

            if qty_col_dash:
                dash_num_map[qty_col_dash] = "Dashboard Quantity"
            dash_num_map["MDP Delay Adjusted"] = "Dashboard MDP Delay"
            dash_num_map["SDP Delay Adjusted"] = "Dashboard SDP Delay"

            missing_num = [c for c in dash_num_map if c not in df_dash.columns]
            if missing_num:
                warn(f"Kolom numerik berikut tidak ditemukan di Dashboard: {missing_num}")
            dash_num_map = {k: v for k, v in dash_num_map.items() if k in df_dash.columns}

            for c in dash_num_map:
                df_dash[c] = pd.to_numeric(
                    df_dash[c].astype(str).str.replace(",", "", regex=False),
                    errors="coerce",
                )

            if dash_num_map:
                dash_num_sum = (
                    df_dash.groupby("PO", as_index=False)[list(dash_num_map.keys())]
                    .sum(min_count=1)
                    .rename(columns={"PO": "PO No.(Full)", **dash_num_map})
                )
                lookup_table = lookup_table.merge(dash_num_sum, on="PO No.(Full)", how="left")

            # --- Dashboard CRD (kolom CRD_) -> tanggal paling awal per PO ---
            if "CRD_" in df_dash.columns:
                dash_date_map["CRD_"] = "Dashboard CRD"
            else:
                warn("Kolom 'CRD_' tidak ditemukan di Dashboard, 'Dashboard CRD' dilewati.")

            for c in dash_date_map:
                df_dash[c] = pd.to_datetime(df_dash[c].apply(safe_parse_date), errors="coerce")

            if dash_date_map:
                dash_date_min = (
                    df_dash.groupby("PO", as_index=False)[list(dash_date_map.keys())]
                    .min()
                    .rename(columns={"PO": "PO No.(Full)", **dash_date_map})
                )
                lookup_table = lookup_table.merge(dash_date_min, on="PO No.(Full)", how="left")

            df = df.merge(lookup_table, on="PO No.(Full)", how="left")

            # --- PO ada di Dashboard tapi tidak ada di SAP ---
            not_in_sap_mask = ~df_dash["PO"].isin(set(df["PO No.(Full)"]))
            df_dash_only = df_dash_raw.loc[not_in_sap_mask].copy()
            df_dash_only["PO"] = df_dash.loc[not_in_sap_mask, "PO"]
            warn(
                f"PO di Dashboard yang tidak ada di SAP: "
                f"{df_dash_only['PO'].nunique()} PO ({len(df_dash_only)} baris)."
            )

            # =================================================
            # 6f2. PO multi-baris di SAP & Dashboard (subset-sum)
            # =================================================
            if qty_col_dash and "Quantity" in df.columns:
                dash_po_counts_map = df_dash.groupby("PO").size()
                sap_po_counts_map = df.groupby("PO No.(Full)").size()

                multi_pos = [
                    po for po, cnt in dash_po_counts_map.items()
                    if cnt > 1 and sap_po_counts_map.get(po, 0) >= 1
                ]

                for po in multi_pos:
                    sap_idx_qty = [
                        (si, df.at[si, "Quantity"])
                        for si in df.index[df["PO No.(Full)"] == po]
                    ]

                    dash_sub = df_dash[df_dash["PO"] == po].copy()
                    dash_sub[qty_col_dash] = pd.to_numeric(dash_sub[qty_col_dash], errors="coerce")
                    qty_map = dash_sub[qty_col_dash].dropna().to_dict()
                    crd_map = dash_sub["CRD_"].to_dict() if "CRD_" in dash_sub.columns else {}
                    remaining_idx = list(qty_map.keys())
                    all_dash_qtys = list(qty_map.values())

                    assigned = {}  # index SAP -> (baris Dashboard terpilih, selisih qty)

                    # Pass 1: qty SAMA PERSIS dan CRD SAMA PERSIS -> pasangkan dulu
                    for si, sap_qty in sap_idx_qty:
                        sap_crd = df.at[si, "CRD"] if "CRD" in df.columns else pd.NaT
                        if pd.isna(sap_qty) or pd.isna(sap_crd):
                            continue
                        sap_crd_day = pd.Timestamp(sap_crd).normalize()
                        for di in remaining_idx:
                            d_crd = crd_map.get(di, pd.NaT)
                            if (
                                abs(qty_map[di] - sap_qty) <= TOL
                                and pd.notna(d_crd)
                                and pd.Timestamp(d_crd).normalize() == sap_crd_day
                            ):
                                assigned[si] = ([di], 0)
                                remaining_idx.remove(di)
                                break

                    # Pass 2: sisanya (qty terbesar dulu) -> kombinasi paling mirip (qty lalu CRD)
                    pending = [(si, q) for si, q in sap_idx_qty if si not in assigned]
                    pending.sort(key=lambda x: (pd.isna(x[1]), -(x[1] if pd.notna(x[1]) else 0)))
                    for si, sap_qty in pending:
                        if pd.isna(sap_qty) or not remaining_idx:
                            assigned[si] = ([], None)
                            continue
                        sap_crd = df.at[si, "CRD"] if "CRD" in df.columns else pd.NaT
                        chosen, diff = best_matching_subset(
                            remaining_idx, qty_map, sap_qty, sap_crd, crd_map
                        )
                        chosen = list(chosen) if chosen else []
                        for di in chosen:
                            if di in remaining_idx:
                                remaining_idx.remove(di)
                        assigned[si] = (chosen, diff)

                    # Terapkan hasil ke tiap baris SAP
                    for si, sap_qty in sap_idx_qty:
                        chosen, diff = assigned.get(si, ([], None))

                        if diff is None or abs(diff) > TOL:
                            chosen_qtys = [qty_map[i] for i in chosen] if chosen else []
                            qty_mismatch_report.append({
                                "PO": po,
                                "SAP Quantity": sap_qty,
                                "Kombinasi Dashboard Qty Dicoba": (
                                    " + ".join(str(q) for q in chosen_qtys)
                                    if chosen_qtys else "(tidak ada kombinasi ditemukan)"
                                ),
                                "Total Kombinasi Terpilih": sum(chosen_qtys) if chosen_qtys else np.nan,
                                "Selisih": diff if diff is not None else np.nan,
                                "Semua Qty Dashboard di PO ini": " , ".join(str(q) for q in all_dash_qtys),
                            })

                        if not chosen:
                            for new_col in dash_num_map.values():
                                df.at[si, new_col] = np.nan
                            for new_col in dash_date_map.values():
                                df.at[si, new_col] = pd.NaT
                            continue

                        for old_col, new_col in dash_num_map.items():
                            df.at[si, new_col] = dash_sub.loc[chosen, old_col].sum(min_count=1)

                        for old_col, new_col in dash_date_map.items():
                            df.at[si, new_col] = dash_sub.loc[chosen, old_col].min()

                        for col in available_dash_cols:
                            df.at[si, col] = combine_unique_values(dash_sub.loc[chosen, col])

                # =============================================
                # 6f2b. PO 1 baris Dashboard, banyak baris SAP
                # =============================================
                single_dash_multi_sap = [
                    po for po, cnt in dash_po_counts_map.items()
                    if cnt == 1 and sap_po_counts_map.get(po, 0) > 1
                ]

                n_split_applied = 0
                for po in single_dash_multi_sap:
                    sap_mask = df["PO No.(Full)"] == po
                    sap_total = df.loc[sap_mask, "Quantity"].sum()
                    dash_total = df.loc[sap_mask, "Dashboard Quantity"].iloc[0]

                    if pd.isna(dash_total) or abs(sap_total - dash_total) > 1e-9:
                        continue

                    ratio_base = dash_total if dash_total != 0 else 1
                    for new_col in dash_num_map.values():
                        if new_col == "Dashboard Quantity":
                            continue
                        dash_val = df.loc[sap_mask, new_col].iloc[0]
                        if pd.isna(dash_val):
                            continue
                        df.loc[sap_mask, new_col] = df.loc[sap_mask, "Quantity"] * (dash_val / ratio_base)

                    df.loc[sap_mask, "Dashboard Quantity"] = df.loc[sap_mask, "Quantity"]
                    n_split_applied += 1

                if single_dash_multi_sap:
                    warn(
                        f"1 baris Dashboard -> banyak baris SAP: "
                        f"{n_split_applied} dari {len(single_dash_multi_sap)} PO diterapkan."
                    )
                if multi_pos:
                    warn(f"Qty-matching (subset-sum) selesai untuk {len(multi_pos)} PO.")
                    if qty_mismatch_report:
                        warn(
                            f"{len(qty_mismatch_report)} baris qty tidak match persis "
                            f"-> lihat tab 'Qty Mismatch Report'."
                        )
            else:
                warn("Kolom qty pembanding di Dashboard tidak ditemukan, qty-matching dilewati.")

    # ---------- 6f3. Compare Quantity ----------
    if "Dashboard Quantity" in df.columns and "Quantity" in df.columns:
        df["Qty Diff"] = df["Quantity"] - df["Dashboard Quantity"]
        df["Qty Compare"] = np.select(
            [df["Dashboard Quantity"].isna(), df["Qty Diff"].abs() < 1e-9],
            ["Not in Dashboard", "Match"],
            default="Mismatch",
        )

    # ---------- 6f4. Compare CRD ----------
    if "Dashboard CRD" in df.columns and "CRD" in df.columns:
        df["Dashboard CRD"] = pd.to_datetime(df["Dashboard CRD"], errors="coerce")
        df["CRD Diff"] = (
            df["CRD"].dt.normalize() - df["Dashboard CRD"].dt.normalize()
        ).dt.days
        df["CRD Compare"] = np.select(
            [df["Dashboard CRD"].isna(), df["CRD"].isna(), df["CRD Diff"] == 0],
            ["Not in Dashboard", "CRD Empty", "Match"],
            default="Mismatch",
        )

    # ---------- 6h. Normalisasi MDP, PDP, SDP ----------
    for col in ["MDP", "PDP", "SDP"]:
        if col in df.columns:
            status = df[col].astype(str).str.strip().str.upper()
            is_fail = status == "FAIL"
            df.loc[is_fail, col] = "FAIL"
            df.loc[~is_fail, col] = "ON TIME"
        else:
            warn(f"Kolom '{col}' tidak ditemukan, dilewati.")

    # ---------- 6h2. DOL ----------
    if "PODD" in df.columns and "CRD" in df.columns:
        df["DOL"] = df.apply(lambda r: get_dol(r["PODD"], r["CRD"]), axis=1)
    else:
        warn("Kolom 'PODD' dan/atau 'CRD' tidak ditemukan, 'DOL' dilewati.")

    # ---------- 6h3. MDP/SDP Ontime & Delay ----------
    for status_col in ["MDP", "SDP"]:
        if status_col in df.columns and "Quantity" in df.columns:
            st_up = df[status_col].astype(str).str.strip().str.upper()
            df[f"{status_col} Ontime"] = np.where(st_up == "ON TIME", df["Quantity"], 0)
            df[f"{status_col} Delay"] = np.where(st_up == "FAIL", df["Quantity"], 0)
        else:
            warn(f"Kolom '{status_col}' dan/atau 'Quantity' tidak ditemukan, Ontime/Delay dilewati.")

    # ---------- 6h4. GAP MDP & SDP ----------
    for prefix in ["MDP", "SDP"]:
        dash_delay = f"Dashboard {prefix} Delay"
        sap_delay = f"{prefix} Delay"
        if dash_delay in df.columns and sap_delay in df.columns:
            df[f"GAP {prefix}"] = df[dash_delay] - df[sap_delay]
            df[f"Has GAP {prefix}"] = df[f"GAP {prefix}"].notna() & (df[f"GAP {prefix}"].abs() > 1e-9)
        else:
            warn(f"Kolom '{dash_delay}' dan/atau '{sap_delay}' tidak ditemukan, GAP {prefix} dilewati.")

    # ---------- 6i. DRC MDP/SDP ----------
    a, b = "Delay - PO PSDD Update", "Delay/Early - Confirmation CRD"
    if a in df.columns and b in df.columns:
        df["DRC MDP/SDP"] = df[a].fillna(df[b])
    elif a in df.columns:
        df["DRC MDP/SDP"] = df[a]
    elif b in df.columns:
        df["DRC MDP/SDP"] = df[b]
    else:
        warn("Kolom DRC sumber tidak ditemukan, 'DRC MDP/SDP' dilewati.")

    # ---------- 6g. CPR (Y jika Model No ada di daftar yang di-paste) ----------
    if "Model No" not in df.columns:
        warn("Kolom 'Model No' tidak ditemukan di ZRSD1013, kolom CPR dilewati.")
    elif not cpr_models:
        warn("Daftar Model No CPR kosong. Kolom CPR dibiarkan kosong.")
        df["CPR"] = np.nan
    else:
        model_norm = df["Model No"].astype(str).str.strip().str.upper()
        is_cpr = model_norm.isin(set(cpr_models))
        df["CPR"] = np.where(is_cpr, "Y", "")
        df["CPR"] = df["CPR"].replace("", np.nan)
        not_found = [m for m in cpr_models if m not in set(model_norm)]
        warn(
            f"CPR: {len(cpr_models)} Model No di daftar, "
            f"{int(is_cpr.sum())} baris ditandai 'Y'."
        )
        if not_found:
            shown = ", ".join(not_found[:30])
            more = f" (+{len(not_found) - 30} lainnya)" if len(not_found) > 30 else ""
            warn(f"CPR: {len(not_found)} Model No di daftar tidak ada di data: {shown}{more}")

    # ---------- 6j. Shipped Qty ----------
    if "FCR Date" in df.columns and "Quantity" in df.columns:
        df["Shipped Qty"] = np.where(df["FCR Date"].notna(), df["Quantity"], 0)
    else:
        warn("Kolom 'FCR Date' dan/atau 'Quantity' tidak ditemukan, Shipped Qty tidak dibuat.")

    # ---------- 6k. Shipped Rspsv Qty ----------
    if all(c in df.columns for c in ["Responsiveness", "FCR Date", "Quantity"]):
        responsiveness_filled = (
            df["Responsiveness"].notna()
            & df["Responsiveness"].astype(str).str.strip().ne("")
        )
        df["Shipped Rspsv Qty"] = np.where(
            responsiveness_filled & df["FCR Date"].notna(), df["Quantity"], 0
        )
    else:
        warn("Kolom 'Responsiveness', 'FCR Date', dan/atau 'Quantity' tidak ditemukan, Shipped Rspsv Qty tidak dibuat.")

    # ---------- 7. Order Type (Domestic/Export) ----------
    if "Ship-to Country" in df.columns:
        df["Order Type (Domestic/Export)"] = np.where(
            df["Ship-to Country"].str.upper() == "INDONESIA", "Domestic", "Export"
        )
    else:
        warn("Kolom 'Ship-to Country' tidak ditemukan.")
        df["Order Type (Domestic/Export)"] = np.nan

    if "Order Type" in df.columns:
        cols = list(df.columns)
        cols.remove("Order Type (Domestic/Export)")
        cols.insert(cols.index("Order Type") + 1, "Order Type (Domestic/Export)")
        df = df[cols]
    else:
        warn("Kolom 'Order Type' tidak ditemukan, kolom baru ditambahkan di akhir.")

    # ---------- 7b. Rename kolom hasil lookup ----------
    rename_map = {
        "is_SPOMA": "Spoma",
        "is_Product Prio": "Priority Product",
        "is_Key Franchise": "Key Franchaise",
        "Bran Partner Description": "Brand Partner",
        "Purchase Order Status Desc": "PO Status",
        "MDP Status Adjusted": "Dashboard MDP",
        "SDP Status Adjusted": "Dashboard SDP",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # ---------- 7c. TRUE/FALSE -> Y/blank ----------
    for col in ["Spoma", "Priority Product", "ReAct", "Chase", "Key Franchaise", "Elevated Check"]:
        if col in df.columns:
            df[col] = df[col].apply(true_false_to_y)
        else:
            warn(f"Kolom '{col}' tidak ditemukan, dilewati.")

    # ---------- 7d. EFD Ontime/Delay Qty ----------
    if all(c in df.columns for c in ["Elevated Check", "MDP", "Quantity"]):
        is_elevated = df["Elevated Check"].astype(str).str.strip().str.upper() == "Y"
        mdp_up = df["MDP"].astype(str).str.strip().str.upper()
        df["EFD Ontime Qty"] = np.where(is_elevated & (mdp_up == "ON TIME"), df["Quantity"], 0)
        df["EFD Delay Qty"] = np.where(is_elevated & (mdp_up == "FAIL"), df["Quantity"], 0)
    else:
        warn("Kolom 'Elevated Check', 'MDP', dan/atau 'Quantity' tidak ditemukan, EFD Qty dilewati.")

    # ---------- 7e. Sample Qty / Sample Delay Qty ----------
    required_cols = ["Client No", "Order Type Description", "MDP", "Quantity"]
    if all(c in df.columns for c in required_cols):
        is_sample_order = (
            (df["Client No"].astype(str).str.strip().str.upper() == "ZSAS")
            & (df["Order Type Description"].astype(str).str.strip().str.upper() == "SALES SAMPLE ORDER")
        )
        is_fail_sample = is_sample_order & (df["MDP"].astype(str).str.strip().str.upper() == "FAIL")
        df["Sample Qty"] = np.where(is_sample_order, df["Quantity"], 0)
        df["Sample Delay Qty"] = np.where(is_fail_sample, df["Quantity"], 0)
    else:
        missing = [c for c in required_cols if c not in df.columns]
        warn(f"Kolom {missing} tidak ditemukan, Sample Qty/Sample Delay Qty dilewati.")

    # ---------- 8. Urutan kolom final ----------
    desired_order = [
        "Triggering", "Client No", "Order Plant", "Brand Plant Name", "Remark", "SO",
        "Order Type", "Order Type (Domestic/Export)", "Order Type Description",
        "PO No.(Full)", "Customer PO item", "PO No.(Short)",
        "Merchandise Category 2", "Shipped Qty",
        "Shipped Rspsv Qty", "Quantity", "Dashboard Quantity", "Qty Diff", "Qty Compare",
        "Sample Qty", "Sample Delay Qty", "Model Name", "Article No",
        "SAP Material", "Pattern Code(Up.No.)", "Model No", "Outsole Mold",
        "Gender", "Category 1", "Category 2", "Category 3", "Unit Price",
        "Classification Code", "DRC",
        "Delay/Early - Confirmation PD", "Delay/Early - Confirmation CRD",
        "Delay - PO PSDD Update", "Delay - PO PD Update", "DRC MDP/SDP",
        "MDP", "MDP Ontime", "MDP Delay",
        "Dashboard MDP", "Dashboard MDP Delay", "GAP MDP", "Has GAP MDP",
        "PDP",
        "SDP", "SDP Ontime", "SDP Delay",
        "Dashboard SDP", "Dashboard SDP Delay", "GAP SDP", "Has GAP SDP",
        "DOL", "Article Lead time", "Cust Ord No",
        "Ship-to-Sort1", "Ship-to Country", "Ship to Name", "Packing Type",
        "Document Date", "FPD",
        "LPD", "Day - LPD", "Month - LPD",
        "CRD", "Day - CRD", "Month - CRD", "CRD Monthly",
        "Dashboard CRD", "CRD Diff", "CRD Compare",
        "PSDD",
        "PODD", "Day - PODD", "Month - PODD",
        "FCR Date", "Actual PGI", "PODD vs FCR", "PODD vs Actual PGI",
        "PD", "PO Date", "Segment", "S&P LPD", "Currency",
        "Spoma", "Priority Product", "CPR", "ReAct", "Chase",
        "Key Franchaise", "Elevated Check", "EFD Ontime Qty", "EFD Delay Qty", "Campaign Name",
        "Brand Partner", "Responsiveness", "PO Status",
    ]
    existing_desired = [c for c in desired_order if c in df.columns]
    missing_from_source = [c for c in desired_order if c not in df.columns]
    if missing_from_source:
        warn(f"Kolom berikut tidak ditemukan di data (dilewati): {missing_from_source}")
    remaining_cols = [c for c in df.columns if c not in existing_desired]
    df = df[existing_desired + remaining_cols]

    # ---------- 9. Cek missing & duplicate ----------
    missing_summary = df.isna().sum()
    missing_summary = missing_summary[missing_summary > 0]
    dup_count = int(df.duplicated().sum())

    df_qty_report = pd.DataFrame(qty_mismatch_report) if qty_mismatch_report else pd.DataFrame(
        columns=["PO", "SAP Quantity", "Kombinasi Dashboard Qty Dicoba",
                 "Total Kombinasi Terpilih", "Selisih", "Semua Qty Dashboard di PO ini"]
    )

    return {
        "df": df,
        "qty_report": df_qty_report,
        "dash_only": df_dash_only,
        "missing_summary": missing_summary,
        "dup_count": dup_count,
    }


def to_excel_bytes(df, df_qty_report, df_dash_only):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        df_qty_report.to_excel(writer, index=False, sheet_name="Qty Mismatch Report")
        df_dash_only.to_excel(writer, index=False, sheet_name="Dashboard Not in SAP")
    return buf.getvalue()


def for_display(d, n=500):
    """Siapkan DataFrame untuk st.dataframe (hindari error tipe campuran)."""
    out = d.head(n).copy()
    for c in out.select_dtypes(include=["object", "string"]).columns:
        out[c] = out[c].map(lambda v: "" if pd.isna(v) else str(v))
    return out


# =========================================================
# UI STREAMLIT
# =========================================================
st.set_page_config(page_title="ZRSD1013 Cleaner", page_icon="📦", layout="wide")
st.title("📦 ZRSD1013 Cleaner")
st.caption("Bersihkan export ZRSD1013, lookup Dashboard & CPR, bandingkan Qty / CRD / MDP / SDP.")

with st.sidebar:
    st.header("Upload file")
    sap_file = st.file_uploader("ZRSD1013 (wajib)", type=["xlsx", "xlsb", "xls"])
    dash_file = st.file_uploader("Dashboard (opsional)", type=["xlsx", "xlsb", "xls"])
    cpr_text = st.text_area(
        "Model No CPR (paste, satu per baris)",
        height=180,
        placeholder="NJG80\nIH1234\n...",
    )
    run = st.button("Proses", type="primary", disabled=sap_file is None, use_container_width=True)
    st.caption("Sheet pertama dari tiap file yang dibaca. Kolom CPR = Y jika Model No ada di daftar.")

if run and sap_file is not None:
    log = []
    try:
        with st.spinner("Memproses data..."):
            df_raw = read_excel_file(sap_file)
            df_dash_in = read_excel_file(dash_file) if dash_file else None
            cpr_models = parse_model_list(cpr_text)

            result = clean_data(df_raw, df_dash_in, cpr_models, log)
            result["excel"] = to_excel_bytes(result["df"], result["qty_report"], result["dash_only"])
            result["log"] = log
            result["shape_awal"] = df_raw.shape
            result["stamp"] = datetime.datetime.now().strftime("%Y%m%d_%H%M")
            st.session_state["result"] = result
    except Exception as e:
        st.error("Terjadi error saat memproses data.")
        st.exception(e)
        st.session_state.pop("result", None)

if "result" in st.session_state:
    res = st.session_state["result"]
    df = res["df"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Baris awal", f"{res['shape_awal'][0]:,}")
    c2.metric("Baris final", f"{df.shape[0]:,}")
    c3.metric("Kolom final", f"{df.shape[1]:,}")
    c4.metric("Baris duplikat", f"{res['dup_count']:,}")

    st.download_button(
        "⬇️ Download hasil (Excel)",
        data=res["excel"],
        file_name=f"ZRSD1013_cleaned_{res['stamp']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    tab_data, tab_sum, tab_qty, tab_dash, tab_log = st.tabs(
        ["Hasil", "Ringkasan Compare", "Qty Mismatch Report", "Dashboard Not in SAP", "Log"]
    )

    with tab_data:
        view = df
        f1, f2, f3 = st.columns(3)
        if "Qty Compare" in df.columns:
            sel = f1.multiselect("Qty Compare", sorted(df["Qty Compare"].dropna().unique()))
            if sel:
                view = view[view["Qty Compare"].isin(sel)]
        if "CRD Compare" in df.columns:
            sel = f2.multiselect("CRD Compare", sorted(df["CRD Compare"].dropna().unique()))
            if sel:
                view = view[view["CRD Compare"].isin(sel)]
        gap_cols = [c for c in ["Has GAP MDP", "Has GAP SDP"] if c in df.columns]
        if gap_cols:
            sel = f3.multiselect("Punya GAP (MDP/SDP)", gap_cols)
            if sel:
                mask = view[sel].any(axis=1)
                view = view[mask]
        st.write(f"Menampilkan {min(len(view), 500):,} dari {len(view):,} baris (preview maks 500; file Excel berisi semua).")
        st.dataframe(for_display(view), use_container_width=True, height=500)

    with tab_sum:
        s1, s2 = st.columns(2)
        if "Qty Compare" in df.columns:
            s1.subheader("Qty Compare")
            s1.dataframe(df["Qty Compare"].value_counts().rename("Jumlah").to_frame(), use_container_width=True)
        if "CRD Compare" in df.columns:
            s2.subheader("CRD Compare")
            s2.dataframe(df["CRD Compare"].value_counts().rename("Jumlah").to_frame(), use_container_width=True)
        g1, g2 = st.columns(2)
        for col_ui, name in [(g1, "Has GAP MDP"), (g2, "Has GAP SDP")]:
            if name in df.columns:
                col_ui.subheader(name)
                col_ui.dataframe(df[name].value_counts().rename("Jumlah").to_frame(), use_container_width=True)
        st.subheader("Kolom dengan missing value")
        ms = res["missing_summary"]
        if ms.empty:
            st.write("Tidak ada missing value.")
        else:
            st.dataframe(ms.rename("Jumlah kosong").to_frame(), use_container_width=True)

    with tab_qty:
        st.write(f"{len(res['qty_report']):,} baris")
        st.dataframe(for_display(res["qty_report"]), use_container_width=True)

    with tab_dash:
        st.write(f"{len(res['dash_only']):,} baris")
        st.dataframe(for_display(res["dash_only"]), use_container_width=True)

    with tab_log:
        if res["log"]:
            for msg in res["log"]:
                st.write(f"- {msg}")
        else:
            st.write("Tidak ada peringatan.")
else:
    st.info("Upload file ZRSD1013 di sidebar, lalu klik **Proses**.")
