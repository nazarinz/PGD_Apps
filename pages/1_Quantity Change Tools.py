# pages/1_Quantity Change Extractor.py
# ① Extractor (Text/HTML) — PGD Apps
import io
import re
from datetime import datetime
from typing import List, Tuple, Dict, Optional

import pandas as pd
import streamlit as st

from utils import set_page, header, footer

try:
    from bs4 import BeautifulSoup  # type: ignore
    HAS_BS4 = True
except Exception:
    HAS_BS4 = False

set_page("PGD Apps — Quantity Change Extractor", "🧾")
header("🧾 Quantity Change Tools — ① Extractor (Text/HTML)")

# =========================================================
#                   UTIL EXTRACTOR
# =========================================================
def split_blocks(raw: str, delimiter: str = "---") -> List[str]:
    """
    Urutan prioritas:
    1) Jika ada delimiter '---' → gunakan itu.
    2) Jika tidak ada, split berdasarkan baris yang persis 'BTP Ticket Number'.
       Setiap kemunculan label ini dianggap awal blok baru.
    3) Fallback: 3+ baris kosong.
    """
    text = raw or ""
    if not text.strip():
        return []

    # Buang code fence kalau ada
    text = "\n".join([ln for ln in text.splitlines() if ln.strip() not in {"```", "``"}])

    # Mode 1: explicit delimiter
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

    # Mode 2: split by label 'BTP Ticket Number' (BUKAN angka polos)
    lines = text.splitlines()
    starts = [i for i, ln in enumerate(lines) if (ln or "").strip().lower() == "btp ticket number"]
    if starts:
        starts.append(len(lines))
        parts = []
        for a, b in zip(starts, starts[1:]):
            chunk = "\n".join(lines[a:b]).strip()
            if chunk:
                parts.append(chunk)
        return parts

    # Mode 3: fallback by 3+ blank lines
    chunks = [b for b in re.split(r"\n{3,}", text.strip()) if b.strip()]
    return chunks


def looks_like_html(text: str) -> bool:
    t = (text or "").lower()
    return ("<html" in t) or ("<body" in t) or ("</div" in t) or ("<table" in t) or ("<div" in t)


def normalize_lines(txt: str) -> List[str]:
    return [ln.strip() for ln in (txt or "").splitlines() if (ln or "").strip()]


SECTION_END_MARKERS = {
    "tracking log", "outcome", "comments", "attachments", "information"
}


def parse_row_like_line(line: str) -> List[str]:
    """
    Split robust:
    - Jika baris mengandung TAB, pakai split('\t') dan TIDAK membuang token kosong.
      (Penting agar posisi kolom tetap align saat ada PODD kosong, dll.)
    - Jika tidak ada TAB, fallback ke split whitespace (termasuk multi-spasi).
    """
    line = (line or "").rstrip("\n")
    if "\t" in line:
        return line.split("\t")  # preserve empties for alignment
    return re.split(r"\s{2,}|\s+", line.strip())


def is_data_row(ln: str) -> bool:
    parts = parse_row_like_line(ln)
    return (len(parts) >= 5) and any(re.search(r"\d", p or "") for p in parts)


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


def get_label_value(
    area: List[str],
    label: str,
    start_idx: int = 0,
    headers_set: Optional[set] = None
) -> Optional[str]:
    lab = (label or "").lower()
    lower_headers = {(h or "").lower() for h in (headers_set or set())}
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
    Ambil 'Tech_Size' dan 'Original PO Qty' dari blok 'PO Lines (...)' dengan:
    - Header multi-baris (Aggregator, ..., Tech_Size, Original PO Qty)
    - Baris data pertama (tokenized) dengan TAB-split yang preserve empty tokens
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

    # Kalau tidak ketemu baris data, fallback ke label scan
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

    parts = parse_row_like_line(data_line)  # <-- preserve empties saat TAB

    ts = parts[idx_ts] if (idx_ts is not None and idx_ts < len(parts)) else None
    oq = parts[idx_oq] if (idx_oq is not None and idx_oq < len(parts)) else None

    # Bersihkan qty → angka saja
    if oq:
        m = re.search(r"\d{1,10}", str(oq).replace(",", ""))
        oq = m.group(0) if m else str(oq)

    # Fallback kanan bila mapping gagal (ambil token terakhir sebagai OQ dan sebelumnya sebagai TS)
    if (ts is None or ts == "") or (oq is None or oq == ""):
        tokens_no_empty = [t for t in parts if t != ""]
        if len(tokens_no_empty) >= 2:
            oq2 = re.search(r"\d{1,10}", str(tokens_no_empty[-1]).replace(",", ""))
            oq = oq or (oq2.group(0) if oq2 else tokens_no_empty[-1])
            ts = ts or tokens_no_empty[-2]

    return ts, oq


def _get_new_po_qty(lines: List[str]) -> Optional[str]:
    # Cari baris yang mengandung 'New PO Qty' → ambil integer terakhir ("00020 - 50" → 50)
    for ln in lines:
        if "new po qty" in (ln or "").lower():
            nums = re.findall(r"\d+", ln or "")
            if nums:
                return nums[-1]
    return None


def parse_plain_text_block(txt: str) -> Dict[str, Optional[str]]:
    lines = normalize_lines(txt)

    # 1) BTP Ticket Number (STRICT: harus ada labelnya)
    btp_ticket = None
    for i, ln in enumerate(lines):
        if (ln or "").lower() == "btp ticket number" and i + 1 < len(lines):
            cand = lines[i + 1]
            if re.fullmatch(r"\d{7,12}", cand or ""):
                btp_ticket = cand
                break
    # Fallback ringan: pola satu baris (kalau HTML/minify dsb)
    if not btp_ticket:
        m = re.search(
            r"(?i)btp\s*ticket\s*number\s*[:\-]?\s*(\d{7,12})",
            (txt or "").replace("\n", " ")
        )
        if m:
            btp_ticket = m.group(1)

    # 2) New PO Qty (Outcome)
    new_po_qty = _get_new_po_qty(lines)

    # 3) Tech_Size & Original PO Qty
    tech_size, original_po_qty = extract_from_po_lines(lines)

    # fallback lagi untuk Original PO Qty jika masih None
    if original_po_qty is None:
        m = re.search(r"Original PO Qty\s*\n([^\n]+)", txt or "", flags=re.IGNORECASE)
        if m:
            tail = m.group(1)
            m2 = re.search(r"(\d{1,10})", (tail or "").replace(",", ""))
            if m2:
                original_po_qty = m2.group(1)

    return {
        "BTP Ticket Number": btp_ticket,
        "Tech_Size":     tech_size,
        "Original PO Qty": original_po_qty,
        "New PO Qty":      new_po_qty,
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


# =========================================================
#                   STREAMLIT UI
# =========================================================
st.subheader("① Extractor (Text/HTML)")

with st.sidebar:
    st.header("⚙️ Pengaturan (Extractor)")
    delimiter = st.text_input(
        "Delimiter antar blok:",
        value="---",
        help="Opsional. Tanpa delimiter pun aplikasi akan split berdasarkan label 'BTP Ticket Number'."
    )
    show_transposed = st.toggle("🔃 Tampilkan hasil sebagai transpose", value=True)

col1, col2 = st.columns(2)
with col1:
    raw = st.text_area(
        "Paste 1..N blok teks/HTML (boleh tanpa delimiter):",
        height=360,
        placeholder=(
            "All Tasks (1)\n\n10400439\n10400439 - Provide Feedback\n...\n"
            "---\n<blok ke-2 atau HTML>..."
        ),
    )
with col2:
    uploads = st.file_uploader(
        "Atau upload file .txt / .html (boleh banyak):",
        type=["txt", "html", "htm"],
        accept_multiple_files=True,
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
                    col_values = frame[col].astype(str).tolist()
                    if col_values:
                        max_content = max(len(v) for v in col_values)
                        max_len = max(len(str(col)), max_content)
                    else:
                        max_len = len(str(col))
                    ws.set_column(col_idx, col_idx, min(max(10, max_len + 2), 60))
                ws.freeze_panes(1, 0)

        st.download_button(
            "⬇️ Download Excel (Extractor)",
            data=bio.getvalue(),
            file_name=f"po_qty_extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

footer()
