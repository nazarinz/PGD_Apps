"""Vendor Booking KPI - Streamlit app.

Input : ZRSD013, Export Plan, FlexView_PGDVendorBooking
Output: KPI detail per baris ZRSD013 + KPI harian, dipisah sheet 3PL / Non 3PL.

Aturan:
  Matching PTS -> ZRSD013 : PO + quantity. Total qty baris ZRSD013 dari PO yang
                            tertera di PTS harus = "Total Line Quantity" FlexView.
  Ex-Plan                 : vlookup by SO ke export plan (kolom Ex-Factory); tidak ada -> PODD
  Tanggal acuan           : Estimated Cargo Ready Date (ECRD) -> Ex-Plan -> PODD
  Max Booking             : tanggal acuan - 14 hari
  KPI                     : ONTIME bila Issue Date <= Max Booking, selain itu FAIL
                            (belum ada PTS: FAIL bila hari ini > Max Booking, else PENDING)
"""
from __future__ import annotations

import io
import logging
import re
from dataclasses import dataclass
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
log = logging.getLogger("booking_kpi")

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
DEFAULT_LEAD_DAYS = 14
PLAN_DATE_COL = "Ex-Factory"
DEFAULT_3PL_KEYS = ["maersk", "dsv air & sea", "kuehne"]
DEFAULT_EXCLUDED_STATUS = ["Rejected"]
ALWAYS_DROP_STATUS = {"canceled", "cancelled"}  # duplikat -> selalu dibuang
YEAR_ROLLOVER_DAYS = 180   # koreksi tahun untuk tanggal plan tanpa tahun
MAX_SUBSET_LINES = 300     # batas kandidat baris untuk pencarian kombinasi qty

REMARK_3PL = "3PL"
REMARK_NON = "Non 3PL"
REMARK_UNKNOWN = "Unknown"

# status hasil matching PTS -> baris ZRSD013
M_EXACT = "EXACT"
M_SUBSET = "SUBSET"
M_AMBIG = "SUBSET - AMBIGUOUS"
M_MISMATCH = "QTY MISMATCH"
M_NO_LINES = "NO LINES AVAILABLE"
M_NO_TLQ = "NO TOTAL QTY"
M_NOT_BOOKED = "NOT BOOKED"
M_OK = (M_EXACT, M_SUBSET)

ZR_PO, ZR_QTY, ZR_PODD = "PO No.(Full)", "Quanity", "PODD"
ZR_SO, ZR_ITEM = "SO", "Customer PO item"
PLAN_SO, PLAN_FORWARDER = "SO", "Forwarder"
FV_RENAME = {
    "Plan To Ship ID": "PTS",
    "Status": "PTS Status",
    "Consolidator": "LSP",
    "Total Line Quantity": "TLQ",
}
FV_REQUIRED = [
    "Plan To Ship ID",
    "Po Number(s)",
    "Status",
    "Issue Date",
    "Estimated Cargo Ready Date",
    "Consolidator",
    "Total Line Quantity",
]

# kolom hasil hitung, ditaruh setelah semua kolom asli ZRSD013 (urutan sesuai permintaan)
COMPUTED_COLS = [
    "Ex-Plan",
    "PTS",
    "Issue Date",
    "Estimated Cargo Ready Date",
    "PTS Status",
    "Max Booking",
    "Vendor's Booking KPI",
    "Vendor's Booking Ontime Qty",
    "LSP Remark",
    "LSP",
    "Mode",          # tambahan dari FlexView
    "Match Status",  # tambahan: status matching PTS -> ZRSD013
]
KPI_ONTIME = "ONTIME"


@dataclass(frozen=True)
class KpiConfig:
    lead_days: int
    keys_3pl: tuple[str, ...]
    excluded_status: tuple[str, ...]
    today: pd.Timestamp


# --------------------------------------------------------------------------- #
# Processing layer (tanpa kode streamlit)
# --------------------------------------------------------------------------- #
def read_table(file, dtype: dict | None = None) -> pd.DataFrame:
    """Baca xlsx/csv dari file upload, rapikan nama kolom."""
    name = getattr(file, "name", "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file, dtype=dtype)
    else:
        df = pd.read_excel(file, dtype=dtype)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def require_columns(df: pd.DataFrame, cols: list[str], label: str) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"File {label}: kolom tidak ditemukan -> {', '.join(missing)}")
    if df.empty:
        raise ValueError(f"File {label}: tidak ada data.")


def normalize_po(s: pd.Series) -> pd.Series:
    """Samakan format PO/SO: trim, buang '.0' hasil numerik dan leading zero."""
    return (
        s.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.lstrip("0")
    )


def valid_key(s: pd.Series) -> pd.Series:
    return ~s.isin(["", "nan", "None", "NaT"])


def clean_text(s: pd.Series) -> pd.Series:
    """Teks bersih: buang '.0' hasil float dan nilai kosong (1.0 -> 1)."""
    out = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return out.where(~out.isin(["nan", "None", "NaT"]), "")


def clean_key(s: pd.Series) -> pd.Series:
    """Normalisasi + kosongkan nilai tidak valid."""
    out = normalize_po(s)
    return out.where(valid_key(out), "")


def prepare_zrsd013(zr: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (lines, raw). raw = kolom asli ZRSD013; lines = kolom kerja, line_id = indeks raw."""
    if "Quantity" in zr.columns and ZR_QTY not in zr.columns:
        zr = zr.rename(columns={"Quantity": ZR_QTY})
    require_columns(zr, [ZR_PO, ZR_QTY, ZR_PODD], "ZRSD013")

    lines = pd.DataFrame(
        {
            "PO_key": clean_key(zr[ZR_PO]),
            "SO_key": clean_key(zr[ZR_SO]) if ZR_SO in zr.columns else "",
            "Item": clean_text(zr[ZR_ITEM]) if ZR_ITEM in zr.columns else "",
            "Qty": pd.to_numeric(zr[ZR_QTY], errors="coerce").fillna(0).astype(int),
            "PODD": pd.to_datetime(zr[ZR_PODD], errors="coerce"),
        }
    )
    keep = (lines["PO_key"] != "").to_numpy()
    lines = lines[keep].reset_index(drop=True)
    lines.insert(0, "line_id", np.arange(len(lines)))

    raw = zr[keep].reset_index(drop=True).copy()
    raw[ZR_QTY] = lines["Qty"]   # qty numerik, kolom lain apa adanya
    log.info("ZRSD013: %d baris, %d PO", len(lines), lines["PO_key"].nunique())
    return lines, raw


def parse_plan_dates(values: pd.Series, podd: pd.Series) -> pd.Series:
    """Parse tanggal plan. Format 'm/d' (tanpa tahun) memakai tahun PODD."""
    text = values.astype(str).str.strip()
    md = text.str.extract(r"^(\d{1,2})/(\d{1,2})$").astype(float)

    year = podd.dt.year.fillna(pd.Timestamp.today().year).astype(int)
    parts = pd.DataFrame(
        {"year": year, "month": md[0].fillna(0), "day": md[1].fillna(0)}
    ).astype(int)
    built = pd.to_datetime(parts, errors="coerce")

    full = pd.to_datetime(values.where(md[0].isna()), errors="coerce")
    result = built.fillna(full)

    # koreksi lintas tahun (mis. plan Jan untuk PODD Des)
    delta = (result - podd).dt.days
    result = result.mask(delta > YEAR_ROLLOVER_DAYS, result - pd.DateOffset(years=1))
    result = result.mask(delta < -YEAR_ROLLOVER_DAYS, result + pd.DateOffset(years=1))
    return result


def prepare_plan(plan: pd.DataFrame, lines: pd.DataFrame) -> pd.DataFrame:
    """Ex-Plan per SO (vlookup by SO) dari kolom Ex-Factory export plan."""
    require_columns(plan, [PLAN_SO, PLAN_DATE_COL], "Export Plan")

    plan = plan.copy()
    if PLAN_FORWARDER not in plan.columns:
        plan[PLAN_FORWARDER] = pd.NA
        log.warning("Export Plan tidak punya kolom %s", PLAN_FORWARDER)

    plan["SO_key"] = clean_key(plan[PLAN_SO])
    plan = plan[plan["SO_key"] != ""]

    # PODD dari ZRSD013 (export plan sendiri punya kolom PODD, jadi jangan di-merge)
    podd_by_so = lines[lines["SO_key"] != ""].groupby("SO_key")["PODD"].min()
    plan["_podd"] = plan["SO_key"].map(podd_by_so)
    plan["Ex-Plan"] = parse_plan_dates(plan[PLAN_DATE_COL], plan["_podd"])

    # bila SO muncul di beberapa line plan, ambil Ex-Factory paling awal
    ex = (
        plan.sort_values("Ex-Plan")
        .drop_duplicates("SO_key")[["SO_key", "Ex-Plan", PLAN_FORWARDER]]
        .rename(columns={PLAN_FORWARDER: "Plan LSP"})
    )
    log.info("Export Plan: %d SO", len(ex))
    return ex


def prepare_flexview(fv: pd.DataFrame, cfg: KpiConfig) -> pd.DataFrame:
    """Return 1 baris per PTS dengan daftar PO (PO_list) dan Total Line Quantity."""
    require_columns(fv, FV_REQUIRED, "FlexView")

    fv = fv.rename(columns=FV_RENAME).copy()
    if "Mode" not in fv.columns:
        fv["Mode"] = pd.NA
        log.warning("FlexView tidak punya kolom Mode")

    # Status Canceled = duplikat -> hapus
    canceled = fv["PTS Status"].astype(str).str.strip().str.lower().isin(ALWAYS_DROP_STATUS)
    log.info("FlexView: %d baris Canceled dibuang", int(canceled.sum()))
    fv = fv[~canceled].copy()

    fv["Issue Date"] = pd.to_datetime(fv["Issue Date"], errors="coerce")
    fv["Estimated Cargo Ready Date"] = pd.to_datetime(
        fv["Estimated Cargo Ready Date"], errors="coerce"
    )
    fv["TLQ"] = pd.to_numeric(fv["TLQ"], errors="coerce")
    fv = fv[~fv["PTS Status"].isin(cfg.excluded_status)]

    sort_src = fv["Create Date"] if "Create Date" in fv.columns else fv["Issue Date"]
    fv["sort_date"] = pd.to_datetime(sort_src, errors="coerce")
    fv["PO_raw"] = fv["Po Number(s)"]

    po_series = fv["Po Number(s)"].astype(str).str.split(",").explode()
    po_series = clean_key(po_series)
    po_series = po_series[po_series != ""]
    fv["PO_list"] = po_series.groupby(level=0).agg(list)
    fv["PO_list"] = fv["PO_list"].map(lambda v: v if isinstance(v, list) else [])

    pts = fv.drop_duplicates("PTS", keep="last")[
        [
            "PTS", "PO_raw", "PO_list", "TLQ", "Issue Date",
            "Estimated Cargo Ready Date", "PTS Status", "LSP", "Mode", "sort_date",
        ]
    ]
    log.info("FlexView: %d PTS", len(pts))
    return pts.reset_index(drop=True)


def find_subset(qtys: list[int], target: int) -> tuple[list[int] | None, bool]:
    """Cari subset qty yang jumlahnya = target. Return (indeks, unik?)."""
    if target <= 0:
        return None, False
    reach: dict[int, tuple[int, tuple[int, ...]]] = {0: (1, ())}
    for i, q in enumerate(qtys):
        if q <= 0 or q > target:
            continue
        for s, (cnt, idx) in list(reach.items()):
            ns = s + q
            if ns > target:
                continue
            if ns in reach:
                old_cnt, old_idx = reach[ns]
                reach[ns] = (min(old_cnt + cnt, 2), old_idx)
            else:
                reach[ns] = (cnt, idx + (i,))
    hit = reach.get(target)
    if hit is None:
        return None, False
    return list(hit[1]), hit[0] == 1


def match_pts_to_lines(
    lines: pd.DataFrame, pts: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Alokasikan tiap PTS ke baris ZRSD013 berdasarkan PO + Total Line Quantity.

    Pass 1: PTS yang total qty seluruh baris PO-nya = Total Line Quantity.
    Pass 2: sisanya (urut Create Date) -> cari kombinasi baris yang jumlahnya
            = Total Line Quantity dari baris yang belum teralokasi.
    Return (line_map, report).
    """
    qty = lines.set_index("line_id")["Qty"]
    by_po = lines.groupby("PO_key")["line_id"].apply(list).to_dict()
    assigned: dict[int, str] = {}
    status: dict[str, str] = {}
    matched: dict[str, int] = {}
    n_lines: dict[str, int] = {}

    order = pts.sort_values(["sort_date", "PTS"], na_position="last")

    def candidates(po_list: list[str]) -> list[int]:
        ids: list[int] = []
        for po in dict.fromkeys(po_list):
            ids += [i for i in by_po.get(po, []) if i not in assigned]
        return sorted(ids)

    def assign(pts_id: str, ids: list[int], st_: str) -> None:
        for i in ids:
            assigned[i] = pts_id
        status[pts_id] = st_
        matched[pts_id] = int(qty.loc[ids].sum()) if ids else 0
        n_lines[pts_id] = len(ids)

    # Pass 1: cocok penuh
    for r in order.itertuples():
        ids = candidates(r.PO_list)
        if ids and pd.notna(r.TLQ) and int(qty.loc[ids].sum()) == int(r.TLQ):
            assign(r.PTS, ids, M_EXACT)

    # Pass 2: sisanya
    for r in order.itertuples():
        if r.PTS in status:
            continue
        ids = candidates(r.PO_list)
        if not ids:
            assign(r.PTS, [], M_NO_LINES)
            continue
        if pd.isna(r.TLQ):
            assign(r.PTS, ids, M_NO_TLQ)
            continue
        total, tlq = int(qty.loc[ids].sum()), int(r.TLQ)
        if total == tlq:
            assign(r.PTS, ids, M_EXACT)
        elif total < tlq or len(ids) > MAX_SUBSET_LINES:
            assign(r.PTS, ids, M_MISMATCH)
        else:
            picked, unique = find_subset([int(qty.loc[i]) for i in ids], tlq)
            if picked is None:
                assign(r.PTS, ids, M_MISMATCH)
            else:
                assign(r.PTS, [ids[k] for k in picked], M_SUBSET if unique else M_AMBIG)

    line_map = pd.DataFrame(
        {"line_id": list(assigned.keys()), "PTS": list(assigned.values())}
    ).merge(
        pts[["PTS", "Issue Date", "Estimated Cargo Ready Date", "PTS Status", "LSP", "Mode"]],
        on="PTS",
        how="left",
    )
    line_map["Match Status"] = line_map["PTS"].map(status)

    report = pts[["PTS", "PO_raw", "TLQ", "Issue Date", "PTS Status", "LSP", "Mode"]].copy()
    report["Matched Qty"] = report["PTS"].map(matched)
    report["Matched Lines"] = report["PTS"].map(n_lines)
    report["Match Status"] = report["PTS"].map(status)
    report["Qty Diff"] = report["Matched Qty"] - report["TLQ"]
    report = report.rename(columns={"PO_raw": "Po Number(s)", "TLQ": "Total Line Quantity"})
    log.info("Match: %s", report["Match Status"].value_counts().to_dict())
    return line_map, report


def classify_lsp(lsp: pd.Series, keys: tuple[str, ...]) -> pd.Series:
    remark = pd.Series(REMARK_NON, index=lsp.index)
    if keys:
        pattern = "|".join(re.escape(k.lower()) for k in keys)
        is_3pl = lsp.fillna("").str.lower().str.contains(pattern, regex=True)
        remark[is_3pl] = REMARK_3PL
    remark[lsp.isna() | (lsp.astype(str).str.strip() == "")] = REMARK_UNKNOWN
    return remark


def build_kpi(
    lines: pd.DataFrame,
    raw: pd.DataFrame,
    ex: pd.DataFrame,
    line_map: pd.DataFrame,
    cfg: KpiConfig,
) -> pd.DataFrame:
    df = lines.merge(line_map, on="line_id", how="left")
    df = df.merge(
        ex.rename(columns={"Ex-Plan": "Ex-Plan SO", "Plan LSP": "Plan LSP SO"}),
        on="SO_key",
        how="left",
    )

    # Ex-Plan: by SO -> PODD
    df["Ex-Plan"] = df["Ex-Plan SO"].fillna(df["PODD"])
    # LSP: FlexView dulu, jika belum booking pakai Forwarder export plan
    df["LSP"] = df["LSP"].fillna(df["Plan LSP SO"])
    df["LSP Remark"] = classify_lsp(df["LSP"], cfg.keys_3pl)
    df["Match Status"] = df["Match Status"].fillna(M_NOT_BOOKED)

    # Max Booking = (ECRD -> Ex-Plan) - lead days
    ref_date = df["Estimated Cargo Ready Date"].fillna(df["Ex-Plan"])
    df["Max Booking"] = ref_date - pd.Timedelta(days=cfg.lead_days)

    booked = df["Issue Date"].notna()
    has_deadline = df["Max Booking"].notna()
    on_time = df["Issue Date"] <= df["Max Booking"]
    overdue = df["Max Booking"] < cfg.today
    df["Vendor's Booking KPI"] = np.select(
        [~has_deadline, booked & on_time, booked & ~on_time, ~booked & overdue],
        ["NO DATE", KPI_ONTIME, "FAIL", "FAIL"],
        default="PENDING",
    )
    df["Vendor's Booking Ontime Qty"] = df["Qty"].where(
        df["Vendor's Booking KPI"] == KPI_ONTIME, 0
    )

    # semua kolom asli ZRSD013 (urutan asli) + kolom hasil hitung
    assert len(df) == len(raw) and (df["line_id"].to_numpy() == raw.index.to_numpy()).all()
    return pd.concat([raw, df[COMPUTED_COLS].reset_index(drop=True)], axis=1)


def daily_kpi(df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
    """KPI harian per tanggal Max Booking (hanya yang deadline-nya sudah lewat)."""
    due = df[df["Max Booking"] <= today]
    daily = (
        due.groupby("Max Booking")
        .agg(
            PO=(ZR_PO, "nunique"),
            Total_Qty=(ZR_QTY, "sum"),
            Ontime_Qty=("Vendor's Booking Ontime Qty", "sum"),
        )
        .reset_index()
    )
    daily["Fail_Qty"] = daily["Total_Qty"] - daily["Ontime_Qty"]
    daily["KPI_pct"] = (
        daily["Ontime_Qty"] / daily["Total_Qty"].replace(0, np.nan) * 100
    ).round(1)
    return daily


def split_by_remark(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    parts = {
        REMARK_3PL: df[df["LSP Remark"] == REMARK_3PL],
        REMARK_NON: df[df["LSP Remark"] == REMARK_NON],
    }
    unknown = df[df["LSP Remark"] == REMARK_UNKNOWN]
    if not unknown.empty:
        parts[REMARK_UNKNOWN] = unknown
    return parts


def for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.select_dtypes(include=["datetime"]).columns:
        out[c] = out[c].dt.date
    return out


@dataclass
class ZrStyle:
    """Format file ZRSD013 asli (header, format kolom, lebar, freeze, filter)."""

    headers: dict[str, dict]
    cols: dict[str, dict]
    widths: dict[str, float]
    freeze: tuple[int, int] | None
    autofilter: bool
    header_height: float | None
    date_fmt: str


DEFAULT_DATE_FMT = "m/d/yyyy"
DATE_COMPUTED_COLS = {"Ex-Plan", "Issue Date", "Estimated Cargo Ready Date", "Max Booking"}
_BORDER = {"thin": 1, "medium": 2, "dashed": 3, "dotted": 4, "thick": 5, "double": 6, "hair": 7}


def _rgb(color) -> str | None:
    try:
        if color is not None and color.type == "rgb" and isinstance(color.rgb, str):
            return "#" + color.rgb[-6:]
    except Exception:  # noqa: BLE001 - warna tema/indexed diabaikan
        pass
    return None


def _cell_fmt(cell) -> dict:
    """Ubah style sel openpyxl menjadi dict format xlsxwriter."""
    f: dict = {"num_format": cell.number_format or "General"}
    font = cell.font
    if font is not None:
        if font.name:
            f["font_name"] = font.name
        if font.size:
            f["font_size"] = font.size
        if font.bold:
            f["bold"] = True
        if font.italic:
            f["italic"] = True
        if _rgb(font.color):
            f["font_color"] = _rgb(font.color)
    if cell.fill is not None and cell.fill.fill_type == "solid" and _rgb(cell.fill.fgColor):
        f["bg_color"] = _rgb(cell.fill.fgColor)
    al = cell.alignment
    if al is not None:
        if al.horizontal in ("left", "center", "right", "fill", "justify"):
            f["align"] = al.horizontal
        if al.vertical:
            f["valign"] = {"center": "vcenter"}.get(al.vertical, al.vertical)
        if al.wrap_text:
            f["text_wrap"] = True
    for side in ("left", "right", "top", "bottom"):
        st_ = getattr(cell.border, side).style if cell.border else None
        if st_:
            f[side] = _BORDER.get(st_, 1)
    return f


def extract_zr_style(file) -> ZrStyle | None:
    """Baca format file ZRSD013 asli. None bila csv / gagal dibaca."""
    if getattr(file, "name", "").lower().endswith(".csv"):
        return None
    try:
        from openpyxl import load_workbook
        from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

        ws = load_workbook(io.BytesIO(file.getvalue())).worksheets[0]
        names = {
            c.column: str(c.value).strip() for c in ws[1] if c.value is not None
        }
        headers = {n: _cell_fmt(ws.cell(1, i)) for i, n in names.items()}
        cols = (
            {n: _cell_fmt(ws.cell(2, i)) for i, n in names.items()}
            if ws.max_row >= 2
            else {}
        )

        widths: dict[str, float] = {}
        for dim in ws.column_dimensions.values():
            if not dim.width:
                continue
            lo = dim.min or column_index_from_string(dim.index)
            hi = dim.max or lo
            for idx in range(lo, hi + 1):
                if idx in names:
                    widths[names[idx]] = dim.width

        freeze = None
        if ws.freeze_panes:
            col_letter, row = coordinate_from_string(ws.freeze_panes)
            freeze = (row - 1, column_index_from_string(col_letter) - 1)

        podd_fmt = cols.get(ZR_PODD, {}).get("num_format", "General")
        date_fmt = podd_fmt if re.search(r"[ymd]", podd_fmt.lower()) else DEFAULT_DATE_FMT
        height = ws.row_dimensions[1].height
        return ZrStyle(headers, cols, widths, freeze, bool(ws.auto_filter.ref), height, date_fmt)
    except Exception:  # noqa: BLE001 - styling opsional, jangan gagalkan proses
        log.exception("Format file ZRSD013 gagal dibaca, pakai format default")
        return None


def write_detail_sheet(
    writer: pd.ExcelWriter, name: str, df: pd.DataFrame, style: ZrStyle | None
) -> None:
    out = for_display(df)
    if style is None:
        out.to_excel(writer, sheet_name=name, index=False)
        return

    out.to_excel(writer, sheet_name=name, index=False, header=False, startrow=1)
    book, ws = writer.book, writer.sheets[name]
    cache: dict[str, object] = {}

    def fmt(d: dict):
        key = repr(sorted(d.items()))
        if key not in cache:
            cache[key] = book.add_format(d)
        return cache[key]

    orig_cols = [c for c in out.columns if c in style.headers]
    last_orig = orig_cols[-1] if orig_cols else None
    header_default = style.headers.get(last_orig, {"bold": True})
    data_default = dict(style.cols.get(ZR_PODD) or style.cols.get(last_orig) or {})

    for i, col in enumerate(out.columns):
        ws.write(0, i, col, fmt(style.headers.get(col, header_default)))
        if col in style.cols:
            col_fmt = style.cols[col]
        else:  # kolom hasil hitung: gaya kolom terakhir, format tanggal sesuai file asli
            col_fmt = {
                **data_default,
                "num_format": style.date_fmt if col in DATE_COMPUTED_COLS else "General",
            }
        width = style.widths.get(col, 18 if col not in style.headers else 12)
        ws.set_column(i, i, width, fmt(col_fmt))

    if style.header_height:
        ws.set_row(0, style.header_height)
    if style.freeze:
        ws.freeze_panes(*style.freeze)
    if style.autofilter:
        ws.autofilter(0, 0, max(len(out), 1), len(out.columns) - 1)


def to_excel_bytes(
    parts: dict[str, pd.DataFrame],
    report: pd.DataFrame,
    today: pd.Timestamp,
    style: ZrStyle | None = None,
) -> bytes:
    date_fmt = style.date_fmt if style else DEFAULT_DATE_FMT
    buf = io.BytesIO()
    with pd.ExcelWriter(
        buf, engine="xlsxwriter", datetime_format=date_fmt, date_format=date_fmt
    ) as writer:
        for remark, part in parts.items():
            write_detail_sheet(writer, f"Detail {remark}", part, style)
            for_display(daily_kpi(part, today)).to_excel(
                writer, sheet_name=f"Daily {remark}", index=False
            )
        for_display(report).to_excel(writer, sheet_name="PTS Match", index=False)
    return buf.getvalue()


def run_pipeline(
    zr_file, plan_file, fv_file, cfg: KpiConfig
) -> tuple[pd.DataFrame, pd.DataFrame]:
    zr = read_table(zr_file, dtype={ZR_PO: str, ZR_SO: str, ZR_ITEM: str})
    plan = read_table(plan_file, dtype={PLAN_SO: str})
    fv = read_table(fv_file, dtype=str)

    lines, raw = prepare_zrsd013(zr)
    ex = prepare_plan(plan, lines)
    pts = prepare_flexview(fv, cfg)
    line_map, report = match_pts_to_lines(lines, pts)
    return build_kpi(lines, raw, ex, line_map, cfg), report


# --------------------------------------------------------------------------- #
# UI layer (hanya memanggil fungsi di atas)
# --------------------------------------------------------------------------- #
def render_match_summary(report: pd.DataFrame) -> None:
    counts = report["Match Status"].value_counts()
    bad = report[~report["Match Status"].isin(M_OK)]
    st.caption(
        "Matching PTS ke ZRSD013: "
        + ", ".join(f"{k}: {v}" for k, v in counts.items())
    )
    if not bad.empty:
        st.warning(
            f"{len(bad)} PTS tidak cocok sempurna dengan qty ZRSD013 "
            "(QTY MISMATCH / AMBIGUOUS / NO LINES). Periksa sebelum dipakai."
        )
        with st.expander("Lihat PTS yang perlu dicek"):
            st.dataframe(for_display(bad), use_container_width=True, hide_index=True)


def render_part(remark: str, part: pd.DataFrame, today: pd.Timestamp) -> None:
    daily = daily_kpi(part, today)
    total_qty = int(daily["Total_Qty"].sum()) if not daily.empty else 0
    ontime_qty = int(daily["Ontime_Qty"].sum()) if not daily.empty else 0
    pct = round(ontime_qty / total_qty * 100, 1) if total_qty else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Baris jatuh tempo", f"{int(len(part[part['Max Booking'] <= today])):,}")
    c2.metric("Total Qty", f"{total_qty:,}")
    c3.metric("Ontime Qty", f"{ontime_qty:,}")
    c4.metric("KPI Booking", f"{pct}%")

    status_counts = part["Vendor's Booking KPI"].value_counts().to_dict()
    st.caption("Jumlah baris per status: " + ", ".join(f"{k}: {v}" for k, v in status_counts.items()))

    st.subheader("KPI harian (per tanggal Max Booking)")
    if daily.empty:
        st.info("Belum ada PO yang jatuh tempo.")
    else:
        st.line_chart(daily.set_index("Max Booking")["KPI_pct"])
        st.dataframe(for_display(daily), use_container_width=True, hide_index=True)

    st.subheader("Detail")
    options = sorted(part["Vendor's Booking KPI"].unique())
    chosen = st.multiselect(
        "Filter status KPI", options, default=options, key=f"flt_{remark}"
    )
    st.dataframe(
        for_display(part[part["Vendor's Booking KPI"].isin(chosen)]),
        use_container_width=True,
        hide_index=True,
    )


def main() -> None:
    st.set_page_config(page_title="Vendor Booking KPI", layout="wide")
    st.title("Vendor Booking KPI")

    with st.sidebar:
        st.header("Pengaturan")
        today = pd.Timestamp(st.date_input("Tanggal acuan (hari ini)", value=date.today()))
        lead_days = st.number_input(
            "Max Booking = ECRD minus (hari)", min_value=0, value=DEFAULT_LEAD_DAYS
        )
        keys_text = st.text_area(
            "Kata kunci LSP 3PL (satu per baris)", value="\n".join(DEFAULT_3PL_KEYS)
        )
        excluded = st.multiselect(
            "Status PTS yang dikeluarkan",
            ["Rejected", "Active", "Accepted"],
            default=DEFAULT_EXCLUDED_STATUS,
        )

    c1, c2, c3 = st.columns(3)
    zr_file = c1.file_uploader("ZRSD013", type=["xlsx", "xlsm", "csv"])
    plan_file = c2.file_uploader("Export Plan", type=["xlsx", "xlsm", "csv"])
    fv_file = c3.file_uploader("FlexView_PGDVendorBooking", type=["xlsx", "xlsm", "csv"])

    ready = all([zr_file, plan_file, fv_file])
    if st.button("Hitung KPI", type="primary", disabled=not ready):
        cfg = KpiConfig(
            lead_days=int(lead_days),
            keys_3pl=tuple(k.strip() for k in keys_text.splitlines() if k.strip()),
            excluded_status=tuple(excluded),
            today=today,
        )
        try:
            with st.spinner("Menghitung..."):
                df, report = run_pipeline(zr_file, plan_file, fv_file, cfg)
            st.session_state["result"] = (df, report, today, extract_zr_style(zr_file))
        except ValueError as exc:
            st.error(str(exc))
            st.session_state.pop("result", None)
        except Exception as exc:  # noqa: BLE001 - tampilkan ke user, log traceback
            log.exception("Pipeline gagal")
            st.error(f"Terjadi error saat memproses data: {exc}")
            st.session_state.pop("result", None)

    if "result" not in st.session_state:
        if not ready:
            st.info("Upload ketiga file untuk memulai.")
        return

    df, report, result_today, zr_style = st.session_state["result"]
    parts = split_by_remark(df)

    st.download_button(
        "Download Excel (sheet 3PL / Non 3PL / PTS Match)",
        data=to_excel_bytes(parts, report, result_today, zr_style),
        file_name=f"Booking_KPI_{result_today:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    render_match_summary(report)

    if REMARK_UNKNOWN in parts:
        st.warning(
            f"{len(parts[REMARK_UNKNOWN])} baris tidak punya LSP (tidak ada di FlexView maupun "
            "kolom Forwarder export plan). Lihat tab Unknown."
        )

    tabs = st.tabs(list(parts.keys()))
    for tab, (remark, part) in zip(tabs, parts.items()):
        with tab:
            render_part(remark, part, result_today)


if __name__ == "__main__":
    main()
