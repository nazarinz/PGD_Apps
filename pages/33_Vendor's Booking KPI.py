"""Vendor Booking KPI - Streamlit app.

Input : ZRSD013, Export Plan, FlexView_PGDVendorBooking
Output: KPI detail per PO + KPI harian, dipisah sheet 3PL / Non 3PL.

Aturan:
  Tanggal acuan = Estimated Cargo Ready Date (ECRD) -> Ex-Plan -> PODD
  Max Booking   = tanggal acuan - 14 hari
  KPI           = PASS bila Issue Date <= Max Booking, selain itu FAIL
                  (belum ada PTS: FAIL bila hari ini > Max Booking, selain itu PENDING)
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
PLAN_DATE_OPTIONS = ["PPIC Plan", "Ex-Factory"]
DEFAULT_3PL_KEYS = ["maersk", "dsv air & sea", "kuehne"]
DEFAULT_EXCLUDED_STATUS = ["Rejected"]
YEAR_ROLLOVER_DAYS = 180  # koreksi tahun untuk tanggal plan tanpa tahun

REMARK_3PL = "3PL"
REMARK_NON = "Non 3PL"
REMARK_UNKNOWN = "Unknown"

ZR_PO, ZR_QTY, ZR_PODD = "PO No.(Full)", "Quanity", "PODD"
PLAN_PO, PLAN_FORWARDER = "PO", "Forwarder"
FV_RENAME = {
    "Plan To Ship ID": "PTS",
    "Status": "PTS Status",
    "Consolidator": "LSP",
}
FV_REQUIRED = [
    "Plan To Ship ID",
    "Po Number(s)",
    "Status",
    "Issue Date",
    "Estimated Cargo Ready Date",
    "Consolidator",
]

OUTPUT_COLS = [
    "Ex-Plan",
    "PO",
    "Qty",
    "PTS",
    "Issue Date",
    "Estimated Cargo Ready Date",
    "PTS Status",
    "Max Booking",
    "Vendor's Booking KPI",
    "Vendor's Booking Ontime Qty",
    "LSP Remark",
    "LSP",
]


@dataclass(frozen=True)
class KpiConfig:
    lead_days: int
    plan_date_col: str
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
    """Samakan format PO: trim, buang '.0' hasil numerik dan leading zero."""
    return (
        s.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.lstrip("0")
    )


def valid_po(s: pd.Series) -> pd.Series:
    return ~s.isin(["", "nan", "None", "NaT"])


def prepare_zrsd013(zr: pd.DataFrame) -> pd.DataFrame:
    """Basis: 1 baris per PO (Qty dijumlah, PODD paling awal)."""
    if "Quantity" in zr.columns and ZR_QTY not in zr.columns:
        zr = zr.rename(columns={"Quantity": ZR_QTY})
    require_columns(zr, [ZR_PO, ZR_QTY, ZR_PODD], "ZRSD013")

    zr = zr.copy()
    zr["PO_key"] = normalize_po(zr[ZR_PO])
    zr = zr[valid_po(zr["PO_key"])]
    zr[ZR_PODD] = pd.to_datetime(zr[ZR_PODD], errors="coerce")
    zr[ZR_QTY] = pd.to_numeric(zr[ZR_QTY], errors="coerce").fillna(0)

    base = zr.groupby("PO_key", as_index=False).agg(
        Qty=(ZR_QTY, "sum"), PODD=(ZR_PODD, "min")
    )
    base["Qty"] = base["Qty"].astype(int)
    log.info("ZRSD013: %d PO", len(base))
    return base


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


def prepare_plan(plan: pd.DataFrame, base: pd.DataFrame, cfg: KpiConfig) -> pd.DataFrame:
    require_columns(plan, [PLAN_PO, cfg.plan_date_col], "Export Plan")

    plan = plan.copy()
    if PLAN_FORWARDER not in plan.columns:
        plan[PLAN_FORWARDER] = pd.NA
        log.warning("Export Plan tidak punya kolom %s", PLAN_FORWARDER)

    plan["PO_key"] = normalize_po(plan[PLAN_PO])
    plan = plan[valid_po(plan["PO_key"])]
    plan = plan.merge(base[["PO_key", "PODD"]], on="PO_key", how="left")
    plan["Ex-Plan"] = parse_plan_dates(plan[cfg.plan_date_col], plan["PODD"])

    ex = (
        plan.sort_values("Ex-Plan")
        .drop_duplicates("PO_key", keep="first")[["PO_key", "Ex-Plan", PLAN_FORWARDER]]
        .rename(columns={PLAN_FORWARDER: "Plan LSP"})
    )
    log.info("Export Plan: %d PO", len(ex))
    return ex


def prepare_flexview(fv: pd.DataFrame, cfg: KpiConfig) -> pd.DataFrame:
    require_columns(fv, FV_REQUIRED, "FlexView")

    fv = fv.rename(columns=FV_RENAME).copy()
    fv["Issue Date"] = pd.to_datetime(fv["Issue Date"], errors="coerce")
    fv["Estimated Cargo Ready Date"] = pd.to_datetime(
        fv["Estimated Cargo Ready Date"], errors="coerce"
    )
    fv = fv[~fv["PTS Status"].isin(cfg.excluded_status)]

    # 1 PTS bisa berisi banyak PO -> explode
    fv["PO_key"] = fv["Po Number(s)"].astype(str).str.split(",")
    fv = fv.explode("PO_key")
    fv["PO_key"] = normalize_po(fv["PO_key"])
    fv = fv[valid_po(fv["PO_key"])]

    # 1 PO dengan >1 PTS: ambil Issue Date paling awal
    fv = (
        fv.sort_values("Issue Date")
        .drop_duplicates("PO_key", keep="first")[
            ["PO_key", "PTS", "Issue Date", "Estimated Cargo Ready Date", "PTS Status", "LSP"]
        ]
    )
    log.info("FlexView: %d PO ter-booking", len(fv))
    return fv


def classify_lsp(lsp: pd.Series, keys: tuple[str, ...]) -> pd.Series:
    remark = pd.Series(REMARK_NON, index=lsp.index)
    if keys:
        pattern = "|".join(re.escape(k.lower()) for k in keys)
        is_3pl = lsp.fillna("").str.lower().str.contains(pattern, regex=True)
        remark[is_3pl] = REMARK_3PL
    remark[lsp.isna() | (lsp.astype(str).str.strip() == "")] = REMARK_UNKNOWN
    return remark


def build_kpi(
    base: pd.DataFrame, ex: pd.DataFrame, fv: pd.DataFrame, cfg: KpiConfig
) -> pd.DataFrame:
    df = base.merge(ex, on="PO_key", how="left").merge(fv, on="PO_key", how="left")

    # Ex-Plan fallback -> PODD
    df["Ex-Plan"] = df["Ex-Plan"].fillna(df["PODD"])
    # LSP: FlexView dulu, jika belum booking pakai Forwarder export plan
    df["LSP"] = df["LSP"].fillna(df["Plan LSP"])
    df["LSP Remark"] = classify_lsp(df["LSP"], cfg.keys_3pl)

    # Max Booking = (ECRD -> Ex-Plan) - lead days
    ref_date = df["Estimated Cargo Ready Date"].fillna(df["Ex-Plan"])
    df["Max Booking"] = ref_date - pd.Timedelta(days=cfg.lead_days)

    booked = df["Issue Date"].notna()
    has_deadline = df["Max Booking"].notna()
    on_time = df["Issue Date"] <= df["Max Booking"]
    overdue = df["Max Booking"] < cfg.today
    df["Vendor's Booking KPI"] = np.select(
        [~has_deadline, booked & on_time, booked & ~on_time, ~booked & overdue],
        ["NO DATE", "PASS", "FAIL", "FAIL"],
        default="PENDING",
    )
    df["Vendor's Booking Ontime Qty"] = df["Qty"].where(
        df["Vendor's Booking KPI"] == "PASS", 0
    )

    df = df.rename(columns={"PO_key": "PO"})
    return df[OUTPUT_COLS].sort_values(["Max Booking", "PO"]).reset_index(drop=True)


def daily_kpi(df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
    """KPI harian per tanggal Max Booking (hanya yang deadline-nya sudah lewat)."""
    due = df[df["Max Booking"] <= today]
    daily = (
        due.groupby("Max Booking")
        .agg(
            PO=("PO", "nunique"),
            Total_Qty=("Qty", "sum"),
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


def to_excel_bytes(parts: dict[str, pd.DataFrame], today: pd.Timestamp) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(
        buf, engine="openpyxl", datetime_format="m/d/yyyy", date_format="m/d/yyyy"
    ) as writer:
        for remark, part in parts.items():
            for_display(part).to_excel(writer, sheet_name=f"Detail {remark}", index=False)
            for_display(daily_kpi(part, today)).to_excel(
                writer, sheet_name=f"Daily {remark}", index=False
            )
    return buf.getvalue()


def run_pipeline(zr_file, plan_file, fv_file, cfg: KpiConfig) -> pd.DataFrame:
    zr = read_table(zr_file, dtype={ZR_PO: str})
    plan = read_table(plan_file, dtype={PLAN_PO: str})
    fv = read_table(fv_file, dtype=str)

    base = prepare_zrsd013(zr)
    ex = prepare_plan(plan, base, cfg)
    fvp = prepare_flexview(fv, cfg)
    return build_kpi(base, ex, fvp, cfg)


# --------------------------------------------------------------------------- #
# UI layer (hanya memanggil fungsi di atas)
# --------------------------------------------------------------------------- #
def render_part(remark: str, part: pd.DataFrame, today: pd.Timestamp) -> None:
    daily = daily_kpi(part, today)
    total_qty = int(daily["Total_Qty"].sum()) if not daily.empty else 0
    ontime_qty = int(daily["Ontime_Qty"].sum()) if not daily.empty else 0
    pct = round(ontime_qty / total_qty * 100, 1) if total_qty else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("PO jatuh tempo", f"{int(daily['PO'].sum()):,}" if not daily.empty else "0")
    c2.metric("Total Qty", f"{total_qty:,}")
    c3.metric("Ontime Qty", f"{ontime_qty:,}")
    c4.metric("KPI Booking", f"{pct}%")

    status_counts = part["Vendor's Booking KPI"].value_counts().to_dict()
    st.caption("Jumlah PO per status: " + ", ".join(f"{k}: {v}" for k, v in status_counts.items()))

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
        plan_col = st.selectbox("Kolom Ex-Plan di export plan", PLAN_DATE_OPTIONS)
        keys_text = st.text_area(
            "Kata kunci LSP 3PL (satu per baris)", value="\n".join(DEFAULT_3PL_KEYS)
        )
        excluded = st.multiselect(
            "Status PTS yang dikeluarkan",
            ["Rejected", "Cancelled", "Active", "Accepted"],
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
            plan_date_col=plan_col,
            keys_3pl=tuple(k.strip() for k in keys_text.splitlines() if k.strip()),
            excluded_status=tuple(excluded),
            today=today,
        )
        try:
            with st.spinner("Menghitung..."):
                df = run_pipeline(zr_file, plan_file, fv_file, cfg)
            st.session_state["result"] = (df, today)
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

    df, result_today = st.session_state["result"]
    parts = split_by_remark(df)

    st.download_button(
        "Download Excel (sheet 3PL / Non 3PL)",
        data=to_excel_bytes(parts, result_today),
        file_name=f"Booking_KPI_{result_today:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if REMARK_UNKNOWN in parts:
        st.warning(
            f"{len(parts[REMARK_UNKNOWN])} PO tidak punya LSP (tidak ada di FlexView maupun "
            "kolom Forwarder export plan). Lihat tab Unknown."
        )

    tabs = st.tabs(list(parts.keys()))
    for tab, (remark, part) in zip(tabs, parts.items()):
        with tab:
            render_part(remark, part, result_today)


if __name__ == "__main__":
    main()
