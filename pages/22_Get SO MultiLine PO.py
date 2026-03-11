from utils.auth import require_login

require_login()

import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Get SO from SAP", layout="wide")
st.title("📥 Get SO from SAP (Match by PO + CRD | Fallback by Qty)")

# Upload Section
email_file = st.sidebar.file_uploader("Upload Email Data", type=["xlsx", "xls", "csv"])
sap_file = st.sidebar.file_uploader("Upload SAP Data", type=["xlsx", "xls", "csv"])


# Read File

def read_file(file):
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file, dtype=str)
    return pd.read_excel(file, engine="openpyxl", dtype=str)


# Column detection

def find_col(cols, candidates):
    for cand in candidates:
        for i, c in enumerate(cols):
            if cand.lower() == str(c).lower().strip():
                return i

    for cand in candidates:
        for i, c in enumerate(cols):
            if cand.lower() in str(c).lower():
                return i

    return 0


# Auto detect date columns

def auto_detect_date_cols(cols):
    hints = ["date", "crd", "podd", "po dd", "lpd"]
    return [c for c in cols if any(h in str(c).lower() for h in hints)]


# Normalization helpers

def normalize_po(s):
    return s.astype(str).str.strip().str.lstrip("0")


def normalize_crd(s):
    dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return dt.dt.strftime("%Y-%m-%d")


def normalize_qty(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce"
    )


# Column candidates
SAP_PO_CANDIDATES = ["PO No.(Full)", "PO Number", "PO No", "Purchase Order"]
SAP_CRD_CANDIDATES = ["CRD", "Confirmation CRD", "Confirmed CRD"]
SAP_SO_CANDIDATES = ["SO", "Sales Order", "SO Number", "SO number", "Sales Order Number"]
SAP_QTY_CANDIDATES = ["Quanity", "Quantity", "Qty", "Order Qty", "PO Quantity"]

EMAIL_PO_CANDIDATES = ["PO Number", "PO No", "PO No.(Full)"]
EMAIL_CRD_CANDIDATES = ["CRD"]
EMAIL_QTY_CANDIDATES = ["PO quantity", "Quantity", "Qty"]


# Build SAP lookup

def build_sap_lookup(df, po_col, crd_col, so_col, qty_col=None):
    df = df.copy()

    df["_PO"] = normalize_po(df[po_col])
    df["_CRD"] = normalize_crd(df[crd_col])
    df["_SO"] = df[so_col].astype(str).str.strip()

    agg = {
        "_SO": lambda x: " | ".join(sorted(pd.Series(x).dropna().astype(str).unique()))
    }

    if qty_col:
        df["_QTY"] = normalize_qty(df[qty_col])
        agg["_QTY"] = "sum"

    return df.groupby(["_PO", "_CRD"], dropna=False).agg(agg).reset_index()


if email_file and sap_file:

    with st.spinner("Reading files..."):
        df_email_raw = read_file(email_file)
        df_sap_raw = read_file(sap_file)

    st.success("Files Loaded Successfully")

    with st.expander("🔍 Lihat Kolom File"):
        c1, c2 = st.columns(2)
        c1.write("**Email Columns:**")
        c1.write(list(df_email_raw.columns))
        c2.write("**SAP Columns:**")
        c2.write(list(df_sap_raw.columns))

    st.divider()

    email_cols = list(df_email_raw.columns)
    sap_cols = list(df_sap_raw.columns)

    st.subheader("🗂️ Map Columns")

    c1, c2, c3 = st.columns(3)
    email_po_col = c1.selectbox("Email → PO", email_cols, index=find_col(email_cols, EMAIL_PO_CANDIDATES))
    email_crd_col = c2.selectbox("Email → CRD", email_cols, index=find_col(email_cols, EMAIL_CRD_CANDIDATES))
    email_qty_col = c3.selectbox("Email → Qty", email_cols, index=find_col(email_cols, EMAIL_QTY_CANDIDATES))

    c4, c5, c6, c7 = st.columns(4)
    sap_po_col = c4.selectbox("SAP → PO", sap_cols, index=find_col(sap_cols, SAP_PO_CANDIDATES))
    sap_crd_col = c5.selectbox("SAP → CRD", sap_cols, index=find_col(sap_cols, SAP_CRD_CANDIDATES))
    sap_so_col = c6.selectbox("SAP → SO", sap_cols, index=find_col(sap_cols, SAP_SO_CANDIDATES))
    sap_qty_col = c7.selectbox("SAP → Qty", sap_cols, index=find_col(sap_cols, SAP_QTY_CANDIDATES))

    with st.expander("👀 Preview Mapping Result (5 baris)"):
        try:
            d1, d2 = st.columns(2)
            d1.write("**Email:**")
            d1.dataframe(df_email_raw[[email_po_col, email_crd_col, email_qty_col]].head(5), use_container_width=True)

            d2.write("**SAP:**")
            d2.dataframe(df_sap_raw[[sap_po_col, sap_crd_col, sap_so_col, sap_qty_col]].head(5), use_container_width=True)
        except Exception as e:
            st.warning(str(e))

    st.divider()

    st.subheader("📅 Pilih Kolom Date di Email")

    auto_dates = auto_detect_date_cols(email_cols)

    date_cols_to_convert = st.multiselect(
        "Kolom Date (Short Date di Excel):",
        options=email_cols,
        default=[c for c in auto_dates if c in email_cols]
    )

    st.divider()

    col_m1, col_m2 = st.columns(2)
    col_m1.metric("Email Rows", f"{len(df_email_raw):,}")
    col_m2.metric("SAP Rows", f"{len(df_sap_raw):,}")


    if st.button("🚀 Run Matching"):
        try:
            with st.spinner("Running 2-stage matching..."):

                df_email = df_email_raw.copy()

                df_email["_PO"] = normalize_po(df_email[email_po_col])
                df_email["_CRD"] = normalize_crd(df_email[email_crd_col])
                df_email["_QTY"] = normalize_qty(df_email[email_qty_col])

                for col in date_cols_to_convert:
                    df_email[col] = pd.to_datetime(df_email[col], errors="coerce").dt.normalize()

                sap_lookup = build_sap_lookup(
                    df_sap_raw,
                    sap_po_col,
                    sap_crd_col,
                    sap_so_col,
                    sap_qty_col
                )

                result = df_email.merge(
                    sap_lookup[["_PO", "_CRD", "_SO"]],
                    on=["_PO", "_CRD"],
                    how="left"
                ).rename(columns={"_SO": "SAP SO"})

                result["Match Type"] = result["SAP SO"].notna().map(
                    lambda x: "Exact" if x else None
                )

                unmatched_mask = result["Match Type"].isna()

                if unmatched_mask.sum() > 0:

                    email_qty_sum = (
                        result.loc[unmatched_mask, ["_PO", "_CRD", "_QTY"]]
                        .groupby(["_PO", "_CRD"]) ["_QTY"]
                        .sum()
                        .reset_index()
                        .rename(columns={"_QTY": "_EMAIL_QTY_SUM"})
                    )

                    sap_qty_lookup = sap_lookup[["_PO", "_CRD", "_SO", "_QTY"]].copy().rename(columns={
                        "_CRD": "_SAP_CRD",
                        "_SO": "_SAP_SO",
                        "_QTY": "_SAP_QTY"
                    })

                    fallback = email_qty_sum.merge(sap_qty_lookup, on="_PO", how="left")

                    fallback = fallback[
                        fallback["_EMAIL_QTY_SUM"].notna() &
                        fallback["_SAP_QTY"].notna() &
                        (fallback["_EMAIL_QTY_SUM"] == fallback["_SAP_QTY"])
                    ][["_PO", "_CRD", "_SAP_SO", "_SAP_CRD"]].drop_duplicates()

                    fallback_map = fallback.set_index(["_PO", "_CRD"])

                    for idx in result[unmatched_mask].index:

                        key = (result.at[idx, "_PO"], result.at[idx, "_CRD"])

                        if key in fallback_map.index:

                            row = fallback_map.loc[key]

                            if isinstance(row, pd.DataFrame):
                                so_val = row["_SAP_SO"].iloc[0]
                                crd_val = row["_SAP_CRD"].iloc[0]
                            else:
                                so_val = row["_SAP_SO"]
                                crd_val = row["_SAP_CRD"]

                            result.at[idx, "SAP SO"] = so_val
                            result.at[idx, "Match Type"] = f"Qty Match (SAP CRD: {crd_val})"

                result["Match Type"] = result["Match Type"].fillna("Unmatched")

            exact_n = (result["Match Type"] == "Exact").sum()
            qty_n = result["Match Type"].astype(str).str.startswith("Qty Match").sum()
            unmatched_n = (result["Match Type"] == "Unmatched").sum()

            c1, c2, c3 = st.columns(3)
            c1.metric("✅ Exact Match", f"{exact_n:,}")
            c2.metric("🔄 Qty Match", f"{qty_n:,}")
            c3.metric("❌ Unmatched", f"{unmatched_n:,}")

            st.subheader("Preview Result (Top 50)")

            preview_df = result.drop(columns=["_PO", "_CRD", "_QTY"], errors="ignore").copy()

            for col in date_cols_to_convert:
                if col in preview_df.columns:
                    preview_df[col] = pd.to_datetime(preview_df[col], errors="coerce").dt.strftime("%Y-%m-%d")

            st.dataframe(preview_df.head(50), use_container_width=True)

            buffer = BytesIO()
            export_df = result.drop(columns=["_PO", "_CRD", "_QTY"], errors="ignore")

            with pd.ExcelWriter(buffer, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
                export_df.to_excel(writer, index=False, sheet_name="Result")

                ws = writer.sheets["Result"]

                header = [cell.value for cell in ws[1]]

                for col_name in date_cols_to_convert:

                    if col_name in header:

                        col_idx = header.index(col_name) + 1

                        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):

                            for cell in row:

                                cell.number_format = "YYYY-MM-DD"

            st.download_button(
                label="⬇️ Download Result Excel",
                data=buffer.getvalue(),
                file_name="Email_With_SO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.exception(e)

else:
    st.info("⬆️ Upload Email & SAP file terlebih dahulu.")
