# pages/1_Quantity Change Tools.py
# Adapted from user's PO Tools (Extractor + Normalizer)
import io, re
from datetime import datetime, date
from typing import List, Tuple, Dict, Optional

import numpy as np
import pandas as pd
import streamlit as st
from utils import set_page, header, footer

try:
    from bs4 import BeautifulSoup  # type: ignore
    HAS_BS4 = True
except Exception:
    HAS_BS4 = False

set_page("PGD Apps — Quantity Change Tools", "🧾")
header("🧾 Quantity Change Tools")

# ----------------------------- Util (Extractor)
def split_blocks(raw: str, delimiter: str = "---") -> List[str]:
    if not raw.strip():
        return []
    parts, buf = [], []
    for line in raw.splitlines():
        if line.strip() == delimiter:
            if buf:
                parts.append("\n".join(buf).strip())
                buf = []
        else:
            buf.append(line)
    if buf:
        parts.append("\n".join(buf).strip())
    return [p for p in parts if p]

def looks_like_html(text: str) -> bool:
    t = text.lower()
    return ("<html" in t) or ("<body" in t) or ("</div>" in t) or ("<table" in t)

def normalize_lines(txt: str) -> List[str]:
    return [ln.strip() for ln in txt.splitlines() if ln.strip()]

SECTION_END_MARKERS = {
    "tracking log", "outcome", "comments", "attachments",
    "information"
}

def is_data_row(ln: str) -> bool:
    if "\t" in ln:
        return True
    if re.search(r"\s{2,}", ln):
        return True
    return bool(re.search(r"\d+\s+\d+", ln))

def parse_row_like_line(line: str) -> List[str]:
    parts = re.split(r"\t+|\s{2,}", line.strip())
    if len(parts) <= 1:
        parts = re.split(r"\s+", line.strip())
    while parts and parts[-1] == "":
        parts.pop()
    return parts

def slice_po_lines_area(lines: List[str]) -> List[str]:
    start = None
    for i, ln in enumerate(lines):
        if re.fullmatch(r"po\s*lines\s*(\(\d+\))?", ln, flags=re.IGNORECASE):
            start = i
            break
    if start is None:
        return []
    out = []
    for j in range(start + 1, len(lines)):
        if lines[j].lower() in SECTION_END_MARKERS:
            break
        out.append(lines[j])
    return out

def get_label_value(area: List[str], label: str, start_idx: int = 0, headers_set: Optional[set] = None) -> Optional[str]:
    lab = label.lower()
    lower_headers = {h.lower() for h in (headers_set or set())}
    for i in range(start_idx, len(area)):
        if area[i].lower() == lab:
            for j in range(i + 1, min(i + 6, len(area))):
                v = area[j].strip()
                if v and (v.lower() not in lower_headers):
                    return v
    return None

def extract_from_po_lines(lines: List[str]) -> Tuple[Optional[str], Optional[str]]:
    area = slice_po_lines_area(lines)
    if not area:
        return None, None

    headers: List[str] = []
    data_line: Optional[str] = None
    data_idx = None
    for idx, ln in enumerate(area):
        if is_data_row(ln):
            data_line = ln
            data_idx = idx
            break
        headers.append(ln)

    search_start = (data_idx + 1) if data_idx is not None else 0
    headers_set = set(headers)
    ts_label = get_label_value(area, "Tech_Size", start_idx=search_start, headers_set=headers_set)
    oq_label = get_label_value(area, "Original PO Qty", start_idx=search_start, headers_set=headers_set)
    if oq_label:
        m = re.search(r"\d{1,10}", oq_label.replace(",", ""))
        oq_label = m.group(0) if m else oq_label
    if ts_label or oq_label:
        return ts_label, oq_label

    if data_line:
        parts = parse_row_like_line(data_line)
        ts = oq = None
        if headers:
            try:
                idx_ts = headers.index("Tech_Size")
            except ValueError:
                idx_ts = None
            try:
                idx_oq = headers.index("Original PO Qty")
            except ValueError:
                idx_oq = None

            if idx_ts is not None and idx_ts < len(parts):
                ts = parts[idx_ts]
            if idx_oq is not None and idx_oq < len(parts):
                oq = parts[idx_oq]

        if (ts is None or oq is None) and len(parts) >= 2:
            ts = ts or parts[-2]
            oq = oq or parts[-1]

        if oq:
            m = re.search(r"\d{1,10}", oq.replace(",", ""))
            oq = m.group(0) if m else oq

        return ts, oq

    ts = get_label_value(area, "Tech_Size", headers_set=headers_set)
    oq = get_label_value(area, "Original PO Qty", headers_set=headers_set)
    if oq:
        m = re.search(r"\d{1,10}", oq.replace(",", ""))
        oq = m.group(0) if m else oq
    return ts, oq

def parse_plain_text_block(txt: str) -> Dict[str, Optional[str]]:
    lines = normalize_lines(txt)

    # 1) BTP Ticket Number
    btp_ticket = None
    for i, ln in enumerate(lines):
        if ln.lower() == "btp ticket number" and i + 1 < len(lines):
            cand = lines[i + 1]
            if re.fullmatch(r"\d{7,12}", cand):
                btp_ticket = cand
                break
    if not btp_ticket:
        for ln in lines:
            if re.fullmatch(r"\d{7,12}", ln):
                btp_ticket = ln
                break
    if not btp_ticket:
        m = re.search(r"\b(\d{7,12})\b\s*-\s*", txt)
        if m:
            btp_ticket = m.group(1)

    # 2) New PO Qty (Outcome)
    new_po_qty = None
    try:
        idx_out = [i for i,l in enumerate(lines) if l.lower()=="outcome"][0]
        for j in range(idx_out + 1, len(lines)):
            if lines[j].lower().startswith("new po qty"):
                m = re.search(r"(\d+)\s*$", lines[j]) or re.search(r"-\s*(\d+)", lines[j])
                new_po_qty = m.group(1) if m else None
                break
    except IndexError:
        pass

    # 3) Tech_Size & Original PO Qty
    tech_size, original_po_qty = extract_from_po_lines(lines)

    if original_po_qty is None:
        m = re.search(r"Original PO Qty\s*\n([^\n]+)", txt, flags=re.IGNORECASE)
        if m:
            tail = m.group(1)
            m2 = re.search(r"(\d{1,10})", tail.replace(",", ""))
            if m2:
                original_po_qty = m2.group(1)

    return {
        "BTP Ticket Number": btp_ticket,
        "Tech_Size": tech_size,
        "Original PO Qty": original_po_qty,
        "New PO Qty": new_po_qty,
    }

def parse_html_block(html: str) -> Dict[str, Optional[str]]:
    if not HAS_BS4:
        return parse_plain_text_block(html)
    soup = BeautifulSoup(html, "lxml")
    txt = soup.get_text("\n", strip=True)
    return parse_plain_text_block(txt)

def parse_block_auto(block: str) -> Dict[str, Optional[str]]:
    if looks_like_html(block):
        return parse_html_block(block)
    return parse_plain_text_block(block)

tab1, tab2 = st.tabs(["① Extractor (Text/HTML)", "② Normalizer (Excel)"])

with tab1:
    st.subheader("① Extractor (Text/HTML)")

    with st.sidebar:
        st.header("⚙️ Pengaturan (Extractor)")
        delimiter = st.text_input("Delimiter antar blok:", value="---",
                                  help="Pisahkan setiap tiket/blok dengan baris yang persis berisi ---")

    col1, col2 = st.columns(2)
    with col1:
        raw = st.text_area(
            "Paste 1..N blok teks/HTML (pisahkan dengan delimiter):",
            height=320,
            placeholder="All Tasks (1)\n\n10400439\n10400439 - Provide Feedback\n...\n---\n<blok ke-2 atau HTML>..."
        )
    with col2:
        uploads = st.file_uploader(
            "Atau upload file .txt / .html (boleh banyak):",
            type=["txt", "html", "htm"],
            accept_multiple_files=True
        )

    if st.button("🔎 Ekstrak (Text/HTML)"):
        blocks: List[str] = []
        blocks.extend(split_blocks(raw or "", delimiter=delimiter))
        if uploads:
            for f in uploads:
                content = f.read().decode("utf-8", errors="ignore")
                if content.strip():
                    blocks.append(content)

        if not blocks:
            st.warning("Tidak ada data yang diproses. Paste teks / upload file dulu ya.")
        else:
            results = []
            for i, b in enumerate(blocks, start=1):
                data = parse_block_auto(b)
                data["_Block"] = f"Block_{i:02d}"
                results.append(data)

            df = pd.DataFrame(results)[["_Block", "BTP Ticket Number", "Tech_Size", "Original PO Qty", "New PO Qty"]]
            st.success(f"Berhasil ekstrak {len(df)} blok.")
            st.dataframe(df, use_container_width=True)

            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Extract")
                ws = writer.sheets["Extract"]
                for col_idx, col in enumerate(df.columns):
                    max_len = max(len(str(col)), *(len(str(x)) for x in df[col].astype(str).tolist()))
                    ws.set_column(col_idx, col_idx, min(max(10, max_len + 2), 60))
                ws.freeze_panes(1, 0)
            st.download_button(
                "⬇️ Download Excel (Extractor)",
                data=bio.getvalue(),
                file_name=f"po_qty_extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ====================== NORMALIZER ======================
def normalize_input_columns(df: pd.DataFrame) -> pd.DataFrame:
    def canon(s: str) -> str:
        s = (s or "").strip()
        s = re.sub(r"\uFEFF", "", s)
        s_low = s.lower()
        s_low = re.sub(r"[._:/\\\-]+", " ", s_low)
        s_low = re.sub(r"#", " #", s_low)
        s_low = re.sub(r"\s+", " ", s_low).strip()
        return s_low

    mapping = {
        "prod fact": "Work Center",
        "prod fact #": "Work Center",
        "work center": "Work Center",

        "so no": "Sales Order",
        "so": "Sales Order",
        "sales order": "Sales Order",

        "customer contract no": "Customer Contract ID",
        "customer contract id": "Customer Contract ID",
        "customer contract": "Customer Contract ID",

        "po": "Sold-To PO No.",
        "po no": "Sold-To PO No.",
        "po #": "Sold-To PO No.",
        "po number": "Sold-To PO No.",
        "sold to po no": "Sold-To PO No.",
        "sold to po number": "Sold-To PO No.",

        "ship to party po no": "Ship-To Party PO No.",
        "ship to party po number": "Ship-To Party PO No.",

        "change type": "Status",
        "status": "Status",

        "cost type": "Cost Category",
        "cost category": "Cost Category",

        "crd": "CRD",
        "pd": "PD",
        "lpd": "LPD",
        "podd": "PODD",

        "art name": "Model Name",
        "model name": "Model Name",

        "art #": "Cust Article No.",
        "art": "Cust Article No.",
        "cust article no": "Cust Article No.",
        "cust article": "Cust Article No.",

        "article": "Article",
        "article lead time": "Article Lead Time",

        "cust #": "Ship-To Search Term",
        "cust": "Ship-To Search Term",
        "ship to search term": "Ship-To Search Term",

        "country": "Ship-To Country",
        "ship to country": "Ship-To Country",

        "document date": "Document Date",
        "doc date": "Document Date",

        "size": "Size",

        "ticket #": "Ticket",
        "ticket": "Ticket",

        "claim cost": "Claim Cost",

        "qty": "Old Quantity",
        "old qty": "Old Quantity",
        "new qty": "New Quantity",
        "reduce qty": "Reduce",
        "reduce": "Reduce",
    }

    rename_map = {}
    for col in df.columns:
        if str(col).startswith("UK_"):
            continue
        key = canon(str(col))
        target = mapping.get(key)
        if target:
            rename_map[col] = target

    df2 = df.copy().rename(columns=rename_map)
    df2.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df2.columns]
    return df2

def rename_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    out_map = {
        "Work Center": "Prod Fact.",
        "Sales Order": "SO NO",
        "Customer Contract ID": "Customer Contract No",
        "Sold-To PO No.": "PO#",
        "Status": "Change Type",
        "Cost Category": "Cost Type",
        "Model Name": "Art.Name",
        "Cust Article No.": "Art #",
        "Ship-To Search Term": "Cust#",
        "Ship-To Country": "Country",
        "Ticket": "Ticket#",
        "Old Quantity": "Qty",
        "New Quantity": "New Qty",
        "Reduce": "Reduce Qty",
    }
    return df.rename(columns=out_map)

def _clean_money(x: str) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    x = str(x).strip()
    return re.sub(r"[,$]", "", x)

def _to_float(x):
    x = _clean_money(x)
    if x == "":
        return np.nan
    try:
        return float(x)
    except Exception:
        return np.nan

def _fmt_shortdate_series(s: pd.Series) -> pd.Series:
    dt = pd.to_datetime(s, errors="coerce")
    out = dt.dt.strftime("%m/%d/%Y")
    return out.mask(dt.isna(), "")

def reshape_po(df: pd.DataFrame,
               fixed_cols_all: List[str] = None,
               size_prefix: str = "UK_") -> pd.DataFrame:
    if fixed_cols_all is None:
        fixed_cols_all = [
            "Work Center","Sales Order","Customer Contract ID","Sold-To PO No.",
            "Ship-To Party PO No.","Status","Cost Category","CRD","PD","LPD","PODD",
            "Model Name","Cust Article No.","Article","Article Lead Time",
            "Ship-To Search Term","Ship-To Country","Document Date","Remark","Order Quantity"
        ]

    df = df.copy().dropna(axis=1, how="all")
    for c in df.columns:
        if df[c].dtype == "O":
            df[c] = df[c].astype(str).str.strip()
    df = df.fillna("")

    fixed_cols = [c for c in fixed_cols_all if c in df.columns]
    size_cols = [c for c in df.columns if str(c).startswith(size_prefix)]
    if not size_cols:
        raise ValueError(f"Tidak ditemukan kolom size yang diawali '{size_prefix}' (mis. UK_5K, UK_1-, dst).")

    ffill_cols = [c for c in fixed_cols if c not in ("Remark","Order Quantity")]
    if ffill_cols:
        df[ffill_cols] = df[ffill_cols].replace("", pd.NA).ffill().fillna("")

    group_key_cols = [c for c in fixed_cols if c not in ("Remark","Order Quantity")]
    if not group_key_cols:
        raise ValueError("Kolom kunci tidak ditemukan. Pastikan header sesuai.")
    df["_group_key"] = df[group_key_cols].apply(lambda r: "|".join([str(v) for v in r.values]), axis=1)

    ticket_map, claim_map = {}, {}
    for _, row in df.iterrows():
        remark = row.get("Remark", "")
        if remark not in ("Ticket", "Claim Cost"):
            continue
        gkey = row["_group_key"]
        for sc in size_cols:
            val = str(row.get(sc, "")).strip()
            if val == "" or val.lower() in ("nan", "none"):
                continue
            if remark == "Ticket":
                ticket_map[(gkey, sc)] = val
            else:
                claim_map[(gkey, sc)] = val

    if "Remark" not in df.columns:
        raise ValueError("Kolom 'Remark' tidak ada di data.")
    use = df[df["Remark"].isin({"Old Quantity","New Quantity","Reduce"})].copy()

    long = use.melt(
        id_vars=group_key_cols + ["Remark"],
        value_vars=size_cols,
        var_name="Size",
        value_name="Qty"
    )
    long = long[long["Qty"].astype(str).str.strip() != ""]

    pivot = long.pivot_table(
        index=group_key_cols + ["Size"],
        columns="Remark",
        values="Qty",
        aggfunc="first"
    ).reset_index()

    def _map_ticket(row):
        gk = "|".join([str(row[c]) for c in group_key_cols])
        return ticket_map.get((gk, row["Size"]), "")

    def _map_claim(row):
        gk = "|".join([str(row[c]) for c in group_key_cols])
        return claim_map.get((gk, row["Size"]), "")

    pivot["Ticket"] = pivot.apply(_map_ticket, axis=1)
    pivot["Claim Cost"] = pivot.apply(_map_claim, axis=1)

    pivot = pivot[pivot["Ticket"].astype(str).str.strip() != ""].copy()

    for c in ["Old Quantity","New Quantity","Reduce"]:
        if c in pivot.columns:
            pivot[c] = pivot[c].apply(_to_float)

    if "Claim Cost" in pivot.columns:
        pivot["Claim Cost"] = pivot["Claim Cost"].apply(
            lambda x: (f"${_to_float(x):,.2f}" if pd.notna(_to_float(x)) else "")
        )

    for dc in ["CRD","PD","LPD","PODD","Document Date"]:
        if dc in pivot.columns:
            pivot[dc] = _fmt_shortdate_series(pivot[dc])

    if "Customer Contract ID" in pivot.columns:
        col = pivot["Customer Contract ID"].astype(str).str.strip()
        mask_empty = (col.isin({"", "nan", "NaN", "None"}) | col.str.fullmatch(r"0+"))
        pivot["Customer Contract ID"] = col.mask(mask_empty, "")

    pivot["Ticket"] = pivot["Ticket"].astype(str).str.strip().replace({"nan":"","None":""})

    std_order = [
        "Work Center","Sales Order","Customer Contract ID","Sold-To PO No.","Ship-To Party PO No.",
        "Status","Cost Category","CRD","PD","LPD","PODD","Model Name","Cust Article No.","Article",
        "Article Lead Time","Ship-To Search Term","Ship-To Country","Document Date",
        "Size","Ticket","Claim Cost","Old Quantity","New Quantity","Reduce"
    ]
    std_order = [c for c in std_order if c in pivot.columns] + [c for c in pivot.columns if c not in std_order]
    out = pivot[std_order].copy()

    sort_cols = [c for c in ["Work Center","Sales Order","Model Name","Cust Article No.","Size"] if c in out.columns]
    if sort_cols:
        out = out.sort_values(sort_cols, kind="mergesort")

    return out.reset_index(drop=True)

FINAL_ORDER = [
    "Ticket Date","Prod Fact.","Document Date","SO NO","Customer Contract No","PO#","Ticket#",
    "Factory E-mail Subject","Art.Name","Art #","Article","Cust#","Country","Size",
    "Qty","Reduce Qty","Increase Qty","New Qty","LPD","PODD","Change Type","Cost Type","Claim Cost"
]

def add_fixed_fields_and_select(df_out: pd.DataFrame, ticket_date_str: str, subject_str: str) -> pd.DataFrame:
    df_out = df_out.copy()
    dt = pd.to_datetime(ticket_date_str, errors="coerce")
    df_out["Ticket Date"] = dt.dt.strftime("%m/%d/%Y") if dt.notna().all() else ticket_date_str
    df_out["Factory E-mail Subject"] = str(subject_str or "").strip()

    qty = pd.to_numeric(df_out.get("Qty"), errors="coerce")
    newq = pd.to_numeric(df_out.get("New Qty"), errors="coerce")
    inc = np.where((~pd.isna(qty)) & (~pd.isna(newq)) & (newq > qty), newq - qty, np.nan)
    df_out["Increase Qty"] = inc

    have_cols = [c for c in FINAL_ORDER if c in df_out.columns]
    df_out = df_out[have_cols].copy()

    for col in ["Qty","Reduce Qty","Increase Qty","New Qty"]:
        if col in df_out.columns:
            as_float = pd.to_numeric(df_out[col], errors="coerce")
            df_out[col] = as_float.apply(lambda v: ("" if pd.isna(v) else (str(int(v)) if float(v).is_integer() else f"{v}")))
    return df_out

with tab2:
    st.subheader("② Normalizer (Excel) — Reshape UK_* + Ticket Date & Subject")
    st.markdown('''
**Cara pakai singkat:**
1) Upload Excel sumber (sheet pertama).  
2) Isi **Ticket Date** (sekali untuk semua baris) dan **Factory E-mail Subject**.  
3) Klik **Proses & Download**.  
**Catatan:** Aplikasi otomatis menangkap **semua kolom yang diawali `UK_`** (robust untuk variasi variasi ukuran).
''')

    file_xlsx = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)
    colA, colB, colC = st.columns([1,1,1])
    with colA:
        tdate: Optional[date] = st.date_input("Ticket Date", value=None, format="MM/DD/YYYY")
    with colB:
        subj = st.text_input("Factory E-mail Subject", value="")
    with colC:
        clicked = st.button("⚙️ Proses & Download")

    if clicked:
        if file_xlsx is None:
            st.error("Silakan upload file Excel dulu.")
        elif tdate is None:
            st.error("Silakan isi Ticket Date.")
        else:
            try:
                df_in = pd.read_excel(file_xlsx, sheet_name=0, dtype=str)
                df_in = normalize_input_columns(df_in)
                hasil_std = reshape_po(df_in)
                hasil_lbl = rename_output_columns(hasil_std)
                tdate_str = tdate.strftime("%Y-%m-%d")
                hasil_final = add_fixed_fields_and_select(hasil_lbl, tdate_str, subj)

                st.success(f"Sukses! {len(hasil_final):,} baris dihasilkan.")
                st.dataframe(hasil_final, use_container_width=True)

                bio = io.BytesIO()
                with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
                    hasil_final.to_excel(writer, index=False, sheet_name="Result")
                    ws = writer.sheets["Result"]
                    for i, c in enumerate(hasil_final.columns):
                        max_len = max(len(str(c)), *(len(str(x)) for x in hasil_final[c].astype(str).tolist()))
                        ws.set_column(i, i, min(max(10, max_len + 2), 50))
                    ws.freeze_panes(1, 0)
                st.download_button(
                    "⬇️ Download Excel (Normalizer)",
                    data=bio.getvalue(),
                    file_name=f"hasil_konversi_PO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.exception(e)

footer()