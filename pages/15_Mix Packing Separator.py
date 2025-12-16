import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Mix-Packing Separator",
    page_icon="🧱",
    layout="wide"
)

st.title("🧱 Mix-Packing Separator Tool")
st.markdown("""
Upload **final packing result Excel**.  
App ini akan **menyisipkan baris kosong sebelum setiap mix-packing group**.
""")

# ======================
# CORE SEPARATOR LOGIC
# ======================
def insert_mixpacking_separators(df):
    required_cols = ["PO No.", "Packing Rule No."]
    for c in required_cols:
        if c not in df.columns:
            raise ValueError(f"Missing required column: {c}")

    df = df.copy().reset_index(drop=True)

    # hitung jumlah baris per PO + Packing Rule
    grp_size = (
        df.groupby(["PO No.", "Packing Rule No."], dropna=False)
        .size()
        .rename("grp_size")
        .reset_index()
    )

    df = df.merge(grp_size, on=["PO No.", "Packing Rule No."], how="left")

    result = []
    seen = set()

    for _, row in df.iterrows():
        key = (row["PO No."], row["Packing Rule No."])

        # === separator sebelum mix-packing pertama ===
        if row["grp_size"] > 1 and key not in seen:
            if len(result) > 0:
                result.append({c: "" for c in df.columns if c != "grp_size"})
            seen.add(key)

        result.append(row.drop("grp_size").to_dict())

    return pd.DataFrame(result, columns=[c for c in df.columns if c != "grp_size"])


# ======================
# UI
# ======================
uploaded = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"]
)

if uploaded:
    df = pd.read_excel(uploaded, dtype=str).fillna("")

    st.subheader("📊 Preview (Before)")
    st.dataframe(df.head(30), use_container_width=True)

    if st.button("➕ Insert Mix-Packing Separators"):
        try:
            df_out = insert_mixpacking_separators(df)

            st.subheader("✅ Preview (After)")
            st.dataframe(df_out.head(40), use_container_width=True)

            # download
            bio = BytesIO()
            with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
                df_out.to_excel(writer, index=False, sheet_name="Result")
            bio.seek(0)

            st.download_button(
                "📥 Download Excel with Separators",
                data=bio,
                file_name="packing_with_separators.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(str(e))
