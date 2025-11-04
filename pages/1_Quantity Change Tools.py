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
    """
    1) Kalau ada baris delimiter persis '---', pakai itu.
    2) Jika tidak ada, auto-split berdasarkan baris yang berisi nomor tiket 7-12 digit (mis. 10462867).
    """
    text = raw or ""
    if not text.strip():
        return []

    # mode 1: explicit delimiter line
    if f"\n{delimiter}\n" in f"\n{text}\n":
        parts, buf = [], []
        for line in text.splitlines():
            if line.strip() == delimiter:
                if buf:
                    parts.append("\n".join(buf).strip())
                    buf = []
            else:
                buf.append(line)
        if buf:
            parts.append("\n".join(buf).strip())
        return [p for p in parts if p]

    # mode 2: auto-split by ticket number start markers
    lines = text.splitlines()
    idxs = []
    for i, ln in enumerate(lines):
        if re.fullmatch(r"\s*\d{7,12}\s*", ln or ""):
            idxs.append(i)
    if not idxs:
        # fallback: split by 3+ blank lines
        chunks = [b for b in re.split(r"\n{3,}", text.strip()) if b.strip()]
        return chunks

    idxs.append(len(lines))
    parts = []
    for a, b in zip(idxs, idxs[1:]):
        chunk = "\n".join(lines[a:b]).strip()
        if chunk:
            parts.append(chunk)
    return parts

def looks_like_html(text: str) -> bool:
    t = (text or "").lower()
    return ("<html" in t) or ("<body" in t) or ("</div>" in t) or ("<table" in t) or ("<div" in t)

def normalize_lines(txt: str) -> List[str]:
    return [ln.strip() for ln in (txt or "").splitlines() if (ln or "").strip()]

SECTION_END_MARKERS = {
    "tracking log", "outcome", "comments", "attachments", "information"
}

def is_data_row(ln: str) -> bool:
    parts = re.split(r"\t+|\s+", (ln or "").strip())
    parts = [p for p in parts if p]
    return (len(parts) >= 5) and any(re.search(r"\d", p) for p in parts)

def parse_row_like_line(line: str) -> List[str]:
    # Split robust: TAB atau whitespace umum (1+)
    parts = re.split(r"\t+|\s+", (line or "").strip())
    return [p for p in parts if p != ""]

def slice_po_lines_area(lines: List[str]) -> List[str]:
    start = None
    for i, ln in enumerate(lines):
        if re.fullmatch(r"po\s*lines\s*(\(\d+\))?", ln or "", flags=re.IGNORECASE):
            start = i
            break
    if start is None:
        return []
    out = []
    for j in range(start + 1, len(lines)):
        if (lines[j] or "").lower() in SECTION_END_MARKERS:
            break
        out.append(lines[j])
    return out

def get_label_value(area: List[str], label: str, start_idx: int = 0, headers_set: Optional[set] = None) -> Optional[str]:
    lab = (label or "").lower()
    lower_headers = { (h or "").lower() for h in (headers_set or set()) }
    for i in range(start_idx, len(area)):
        if (area[i] or "").lower() == lab:
            for j in range(i + 1, min(i + 6, len(area))):
                v = (area[j] or "").strip()
                if v and (v.lower() not in lower_headers):
                    return v
    return None

def _norm_hdr(s: str) -> str:
    return re.sub(r"[\s/_#\-]+", "", (s or "").strip().lower())

def extract_from_po_lines(lines: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """
    Cari 'Tech_Size' dan 'Original PO Qty' dari blok 'PO Lines (...)' menggunakan:
    - header berbaris (Aggregator, PO/Contract Line, ..., Tech_Size, Original PO Qty)
    - baris data pertama setelah header (tokenized)
    """
    area = slice_po_lines_area(lines)
    if not area:
        return None, None

    headers: List[str] = []
    data_line: Optional[str] = None

    for ln in area:
        if is_data_row(ln):
            data_line = ln
            break
        headers.append(ln)

    # Fallback label scan jika tidak ketemu baris data
    if not data_line:
        ts = get_label_value(area, "Tech_Size", headers_set=set(headers))
        oq = get_label_value(area, "Original PO Qty", headers_set=set(headers))
        if oq:
            m = re.search(r"\d{1,10}", str(oq).replace(",", ""))
            oq = m.group(0) if m else oq
        return ts, oq

    # Map header → index
    hnorm = [_norm_hdr(h) for h in headers if h is not None]
    def _idx(hname: str) -> Optional[int]:
        try:
            return hnorm.index(_norm_hdr(hname))
        except ValueError:
            return None

    idx_ts = _idx("Tech_Size")
    idx_oq = _idx("Original PO Qty")

    parts = parse_row_like_line(data_line)

    ts = parts[idx_ts] if (idx_ts is not None and idx_ts < len(parts)) else None
    oq = parts[idx_oq] if (idx_oq is not None and idx_oq < len(parts)) else None

    # Cleanup OQ → angka saja
    if oq:
        m = re.search(r"\d{1,10}", str(oq).replace(",", ""))
        oq = m.group(0) if m else str(oq)

    # Fallback terakhir: ambil 2 token paling akhir sebagai (Tech_Size, OQ)
    if (ts is None or ts == "") and len(parts) >= 2:
        ts = parts[-2]
    if (oq is None or oq == "") and len(parts) >= 1:
        m = re.search(r"\d{1,10}", str(parts[-1]).replace(",", ""))
        oq = m.group(0) if m else parts[-1]

    return ts, oq

def _get_new_po_qty(lines: List[str]) -> Optional[str]:
    # Cari baris yang mengandung 'New PO Qty' lalu ambil integer terakhir (mis. "00020 - 50" → 50)
    for ln in lines:
        if "new po qty" in (ln or "").lower():
            nums = re.findall(r"\d+", ln or "")
            if nums:
                return nums[-1]
    return None

def parse_plain_text_block(txt: str) -> Dict[str, Optional[str]]:
    lines = normalize_lines(txt)

    # 1) BTP Ticket Number
    btp_ticket = None
    for i, ln in enumerate(lines):
        if (ln or "").lower() == "btp ticket number" and i + 1 < len(lines):
            cand = lines[i + 1]
            if re.fullmatch(r"\d{7,12}", cand or ""):
                btp_ticket = cand
                break
    if not btp_ticket:
        for ln in lines:
            if re.fullmatch(r"\d{7,12}", ln or ""):
                btp_ticket = ln
                break
    if not btp_ticket:
        m = re.search(r"\b(\d{7,12})\b\s*-\s*", txt or "")
        if m:
            btp_ticket = m.group(1)

    # 2) New PO Qty (Outcome)
    new_po_qty = _get_new_po_qty(lines)

    # 3) Tech_Size & Original PO Qty
    tech_size, original_po_qty = extract_from_po_lines(lines)

    if original_po_qty is None:
        m = re.search(r"Original PO Qty\s*\n([^\n]+)", txt or "", flags=re.IGNORECASE)
        if m:
            tail = m.group(1)
            m2 = re.search(r"(\d{1,10})", (tail or "").replace(",", ""))
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
    if looks_like_html(block or ""):
        return parse_html_block(block)
    return parse_plain_text_block(block)

tab1, tab2 = st.tabs(["① Extractor (Text/HTML)", "② Normalizer (Excel)"])

with tab1:
    st.subheader("① Extractor (Text/HTML)")

    with st.sidebar:
        st.header("⚙️ Pengaturan (Extractor)")
        delimiter = st.text_input(
            "Delimiter antar blok:",
            value="---",
            help="Pisahkan setiap tiket/blok dengan baris yang persis berisi --- (opsional; tanpa ini pun app akan auto-split)."
        )
        show_transposed = st.toggle("🔃 Tampilkan hasil sebagai transpose", value=True)

    col1, col2 = st.columns(2)
    with col1:
        raw = st.text_area(
            "Paste 1..N blok teks/HTML (boleh tanpa delimiter):",
            height=360,
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
        # 1) dari textarea
        if raw and raw.strip():
            blocks.extend(split_blocks(raw, delimiter=delimiter))
        # 2) dari upload files
        if uploads:
            for f in uploads:
                content = f.read().decode("utf-8", errors="ignore")
                if content.strip():
                    blocks.extend(split_blocks(content, delimiter=delimiter))

        # filter kosong
        blocks = [b for b in blocks if b and b.strip()]

        if not blocks:
            st.warning("Tidak ada data yang diproses. Paste teks / upload file dulu ya.")
        else:
            results = []
            for i, b in enumerate(blocks, start=1):
                data = parse_block_auto(b)
                data["_Block"] = f"Block_{i:02d}"
                results.append(data)

            base_cols = ["_Block", "BTP Ticket Number", "Tech_Size", "Original PO Qty", "New PO Qty"]
            df = pd.DataFrame(results)[base_cols]

            st.success(f"Berhasil ekstrak {len(df)} blok.")

            if show_transposed:
                wide = df.set_index("_Block").T.reset_index().rename(columns={"index": "_Block"})
                st.dataframe(wide, use_container_width=True)
            else:
                st.dataframe(df, use_container_width=True)

            # Export Excel: 2 sheet (Extract & Transposed)
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Extract")
                wide = df.set_index("_Block").T.reset_index().rename(columns={"index": "_Block"})
                wide.to_excel(writer, index=False, sheet_name="Transposed")

                for ws_name, frame in [("Extract", df), ("Transposed", wide)]:
                    ws = writer.sheets[ws_name]
                    for col_idx, col in enumerate(frame.columns):
                        max_len = max(len(str(col)), *(len(str(x)) for x in frame[col].astype(str).tolist()))
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

        # tambahan alias umum
        "order quantity": "Order Quantity",
        "order qty": "Order Quantity",
        "old quantity": "Old Quantity",
        "new quantity": "New Quantity",
        "reduce quantity": "Reduce",
        "reduce qty": "Reduce",
        "ship-to country": "Ship-To Country",
        "ship to  country": "Ship-To Country",
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

# ====== FIX: helper & revised function ======
def _format_ticket_date_any(val) -> str:
    """Terima string/date/datetime/Timestamp -> 'MM/DD/YYYY' atau '' jika invalid."""
    if val is None or val == "":
        return ""
    try:
        d = pd.to_datetime(val, errors="coerce")
    except Exception:
        return ""
    return d.strftime("%m/%d/%Y") if pd.notna(d) else ""

def add_fixed_fields_and_select(df_out: pd.DataFrame, ticket_date_val, subject_str: str) -> pd.DataFrame:
    df_out = df_out.copy()

    # 1) Ticket Date: format sekali lalu broadcast
    df_out["Ticket Date"] = _format_ticket_date_any(ticket_date_val)

    # 2) Subject
    df_out["Factory E-mail Subject"] = (subject_str or "").strip()

    # 3) Increase Qty (jika New Qty > Qty)
    qty  = pd.to_numeric(df_out.get("Qty"), errors="coerce")
    newq = pd.to_numeric(df_out.get("New Qty"), errors="coerce")
    inc = np.where((~pd.isna(qty)) & (~pd.isna(newq)) & (newq > qty), newq - qty, np.nan)
    df_out["Increase Qty"] = inc

    # 4) Susun kolom final (hanya yang tersedia)
    have_cols = [c for c in FINAL_ORDER if c in df_out.columns]
    df_out = df_out[have_cols].copy()

    # 5) Format angka
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

                # PASS langsung objek tdate (date) — tidak perlu diubah string
                hasil_final = add_fixed_fields_and_select(hasil_lbl, tdate, subj)

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
