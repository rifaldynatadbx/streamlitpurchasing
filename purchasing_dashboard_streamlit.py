import io
import re
import time
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials

    GSPREAD_AVAILABLE = True

except ImportError:
    GSPREAD_AVAILABLE = False


# =========================================================
# PAGE CONFIGURATION
# =========================================================
st.set_page_config(
    page_title="Procurement Control Center",
    page_icon="◼",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# DEFAULT FILE
# =========================================================
BASE_DIR = Path(__file__).resolve().parent.parent

DEFAULT_FILE = (
    BASE_DIR
    / "data"
    / "purchasing_report.xlsx"
)


# =========================================================
# MONTH MAPPING
# =========================================================
MONTH_ORDER = {
    "januari": 1,
    "jan": 1,
    "january": 1,

    "februari": 2,
    "feb": 2,
    "february": 2,

    "maret": 3,
    "mar": 3,
    "march": 3,

    "april": 4,
    "apr": 4,

    "mei": 5,
    "may": 5,

    "juni": 6,
    "jun": 6,
    "june": 6,

    "juli": 7,
    "jul": 7,
    "july": 7,

    "agustus": 8,
    "aug": 8,
    "august": 8,

    "september": 9,
    "sep": 9,

    "oktober": 10,
    "oct": 10,
    "october": 10,

    "november": 11,
    "nov": 11,

    "desember": 12,
    "dec": 12,
    "december": 12,
}


# =========================================================
# MONTH ABBREVIATION (for MMM-YYYY display label)
# =========================================================
MONTH_ABBR = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}


# =========================================================
# SOURCE COLUMN MAPPING
# =========================================================
REQUIRED_COLUMNS = {
    "cost_center": "Cost Center",
    "cost_category": "COST CATEGORY",
    "site": "SITE",

    "pr_no": "PR Nos",
    "pr_date": "TANGGAL PR",

    "po_no": "PO Nos",
    "po_date": "TANGGAL PO",

    "request_date": "TANGGAL PERMINTAAN BARANG",
    "arrival_date": "TANGGAL KEDATANGAN BARANG",
    "usage_date": "TANGGAL PENGGUNAAN BARANG",

    "costing_month": "COSTING MONTH",
    "vendor": "SUPPLIER/VENDOR/PIC",

    "invoice_received_date": "INVOICE RECEIVED DATE",
    "invoice_status": "INVOICE STATUS",

    "payment_schedule": "JADWAL PEMBAYARAN",
    "payment_schedule_month": "JADWAL PEMBAYARAN (MONTH)",
    "actual_payout_date": "ACTUAL PAYOUT DATE",
    "payment_overdue_days": "PAYMENT OVERDUE (days)",
    "payment_status": "PAYMENT STATUS",

    "item_category": "ITEMS CATEGORY",
    "item_description": "ITEMS DESCRIPTION",
    "unit": "UNIT",

    "qty_po": "QTY PERMINTAAN BARANG\n(PO)",
    "qty_do": "QTY KEDATANGAN BARANG\n(DO)",

    "price_unit": "PRICE / UNIT",

    "grand_total_po": "GRAND TOTAL\n(PO)",
    "grand_total_do": "GRAND TOTAL\n(DO)",

    "total_po": "TOTAL\n(PO)",
    "total_do": "TOTAL\n(DO)",
}


# =========================================================
# DISPLAY COLUMN NAMES
# =========================================================
DISPLAY_NAMES = {
    "cost_center": "Cost Center",
    "cost_category": "Cost Category",
    "site": "Site",

    "pr_no": "PR Number",
    "pr_date": "PR Date",

    "po_no": "PO Number",
    "po_date": "PO Date",

    "request_date": "Request Date",
    "arrival_date": "Arrival Date",
    "usage_date": "Usage Date",

    "costing_month": "Costing Month",
    "vendor": "Vendor",

    "invoice_received_date": "Invoice Received",
    "invoice_status": "Invoice Status",

    "payment_schedule": "Payment Schedule",
    "payment_schedule_month": "Payment Month",
    "actual_payout_date": "Actual Payout",
    "payment_overdue_days": "Payment Overdue Days",
    "payment_status": "Payment Status",

    "item_category": "Item Category",
    "item_description": "Item Description",
    "unit": "Unit",

    "qty_po": "Qty PO",
    "qty_do": "Qty DO",
    "price_unit": "Price / Unit",

    "grand_total_po": "Grand Total PO",
    "grand_total_do": "Grand Total DO",

    "total_po": "Total PO",
    "total_do": "Total DO",

    "qty_variance": "Qty Variance",
    "fulfillment_rate": "Fulfillment Rate",
    "delivery_status": "Delivery Status",

    "value_variance_do_vs_po": "Value Variance DO vs PO",

    "spend_do": "Actual Spend",
    "spend_share": "Spend Share",
    "transaction_count": "Transactions",
    "vendor_count": "Vendor Count",
    "item_count": "Item Count",
    "po_count": "PO Count",

    "avg_price": "Average Price",
    "min_price": "Minimum Price",
    "median_price": "Median Price",
    "max_price": "Maximum Price",
    "price_gap": "Price Gap",
    "price_gap_rate": "Price Gap %",
    "potential_saving": "Potential Saving",

    "avg_fulfillment_rate": "Average Fulfillment",
    "avg_delivery_days": "Average Delivery Days",
    "avg_payment_overdue_days": "Average Overdue Days",

    "vendor_rank": "Vendor Item Rank",
    "vendor_coverage": "Vendor Coverage",
    "lowest_price": "Lowest Price",
    "highest_price": "Highest Price",
    "best_vendor": "Best Price Vendor",

    "line_count": "Line Count",
    "value_do": "DO Value",
    "value_variance": "Value Variance",

    "average_days": "Average Days",
    "median_days": "Median Days",
    "p90_days": "P90 Days",
    "observations": "Observations",

    "missing_count": "Missing Count",
    "missing_rate": "Missing Rate",
    "unique_count": "Unique Count",

    "mom_growth": "MoM Growth",
    "moving_avg": "Moving Average",
    "is_forecast": "Is Forecast",

    "risk_score": "Risk Score",
    "risk_quadrant": "Segment",

    "rank": "Rank",
    "cumulative_value": "Cumulative Spend",
    "cumulative_share": "Cumulative Share",
    "abc_class": "ABC Class",

    "bucket": "Aging Bucket",
    "exposure_value": "Exposure Value",
    "exposure_share": "Exposure Share",

    "source_file": "Source File",
}


# =========================================================
# BASE COLOR
# =========================================================
COLOR = {
    "ink": "#111827",
    "muted": "#6B7280",
    "line": "#E5E7EB",
    "soft": "#F7F8FA",
    "accent": "#334155",
    "accent_2": "#64748B",
    "danger": "#9F1239",
    "success": "#166534",
    "warning": "#92400E",
}


# =========================================================
# VISUAL STYLE
# =========================================================
def apply_app_style(dark_mode: bool) -> None:
    """
    Mengatur seluruh warna dashboard.

    Dark mode dan light mode mempunyai:
    - warna background;
    - warna card;
    - warna sidebar;
    - warna tulisan;
    - warna input;
    - warna button.
    """

    if dark_mode:
        theme = {
            "background": "#0F172A",
            "surface": "#111827",
            "surface_2": "#1E293B",
            "text": "#F8FAFC",
            "muted": "#94A3B8",
            "line": "#334155",
            "accent": "#CBD5E1",
            "button": "#1E293B",
            "sidebar": "#111827",
            "header": "rgba(15, 23, 42, 0.92)",
        }

    else:
        theme = {
            "background": "#FFFFFF",
            "surface": "#FFFFFF",
            "surface_2": "#F7F8FA",
            "text": "#111827",
            "muted": "#6B7280",
            "line": "#E5E7EB",
            "accent": "#334155",
            "button": "#FFFFFF",
            "sidebar": "#F7F8FA",
            "header": "rgba(255, 255, 255, 0.92)",
        }

    st.markdown(
        f"""
        <style>
        :root {{
            color-scheme: {"dark" if dark_mode else "light"};
        }}

        .stApp {{
            background: {theme["background"]};
            color: {theme["text"]};
        }}

        [data-testid="stHeader"] {{
            background: {theme["header"]};
        }}

        [data-testid="stSidebar"] {{
            background: {theme["sidebar"]};
            border-right: 1px solid {theme["line"]};
        }}

        [data-testid="stSidebar"] * {{
            color: {theme["text"]};
        }}

        .block-container {{
            padding-top: 1.7rem;
            padding-bottom: 3rem;
            max-width: 1500px;
        }}

        h1,
        h2,
        h3,
        p,
        label,
        .stMarkdown {{
            color: {theme["text"]};
        }}

        h1,
        h2,
        h3 {{
            letter-spacing: -0.02em;
        }}

        h1 {{
            font-size: 2rem !important;
            font-weight: 720 !important;
        }}

        h2 {{
            font-size: 1.35rem !important;
            font-weight: 680 !important;
        }}

        h3 {{
            font-size: 1.05rem !important;
            font-weight: 650 !important;
        }}

        div[data-testid="stMetric"] {{
            background: {theme["surface"]};
            border: 1px solid {theme["line"]};
            border-radius: 12px;
            padding: 14px 16px;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.12);
        }}

        div[data-testid="stMetricLabel"] {{
            color: {theme["muted"]};
        }}

        div[data-testid="stMetricValue"] {{
            color: {theme["text"]};
            font-size: 1.45rem;
        }}

        .section-label {{
            color: {theme["muted"]};
            font-size: 0.75rem;
            font-weight: 700;
            letter-spacing: 0.10em;
            text-transform: uppercase;
            margin-bottom: 0.25rem;
        }}

        .app-subtitle {{
            color: {theme["muted"]};
            font-size: 0.95rem;
            margin-top: -0.55rem;
            margin-bottom: 1.15rem;
        }}

        .insight-box {{
            background: {theme["surface_2"]};
            border: 1px solid {theme["line"]};
            border-left: 4px solid {theme["accent"]};
            border-radius: 10px;
            padding: 14px 16px;
            margin: 8px 0 16px 0;
        }}

        .insight-title {{
            color: {theme["text"]};
            font-weight: 700;
            margin-bottom: 4px;
        }}

        .insight-text {{
            color: {theme["muted"]};
            font-size: 0.9rem;
        }}

        .guide-box {{
            background: transparent;
            border: 1px dashed {theme["line"]};
            border-radius: 10px;
            padding: 12px 16px;
            margin: 4px 0 18px 0;
        }}

        .guide-title {{
            color: {theme["muted"]};
            font-weight: 700;
            font-size: 0.78rem;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            margin-bottom: 6px;
        }}

        .guide-text {{
            color: {theme["muted"]};
            font-size: 0.87rem;
            line-height: 1.55;
        }}

        .guide-text ul {{
            margin: 0;
            padding-left: 1.1rem;
        }}

        .tag-pill {{
            display: inline-block;
            border-radius: 999px;
            padding: 2px 10px;
            font-size: 0.78rem;
            font-weight: 700;
        }}

        .stTabs [data-baseweb="tab-list"] {{
            gap: 6px;
            border-bottom: 1px solid {theme["line"]};
        }}

        .stTabs [data-baseweb="tab"] {{
            height: 42px;
            background: transparent;
            border-radius: 8px 8px 0 0;
            color: {theme["muted"]};
            padding: 0 14px;
        }}

        .stTabs [aria-selected="true"] {{
            color: {theme["text"]};
            font-weight: 650;
        }}

        .stButton > button,
        .stDownloadButton > button,
        .stFormSubmitButton > button {{
            border-radius: 9px;
            border: 1px solid {theme["line"]};
            background: {theme["button"]};
            color: {theme["text"]};
            font-weight: 600;
        }}

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        .stFormSubmitButton > button:hover {{
            border-color: {theme["accent"]};
            color: {theme["text"]};
        }}

        [data-testid="stDataFrame"] {{
            border: 1px solid {theme["line"]};
            border-radius: 10px;
        }}

        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        [data-testid="stDateInput"] input,
        [data-testid="stNumberInput"] input {{
            background: {theme["surface"]};
            color: {theme["text"]};
            border-color: {theme["line"]};
        }}

        div[data-testid="stExpander"] {{
            background: {theme["surface"]};
            border-color: {theme["line"]};
        }}

        .stAlert {{
            background: {theme["surface_2"]};
            color: {theme["text"]};
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# DATA PREPARATION
# =========================================================
def normalize_col(col: object) -> str:
    """
    Membersihkan nama kolom Excel / Google Sheets.

    Kolom tanpa judul (sel header kosong) bisa muncul dalam
    beberapa bentuk tergantung sumbernya: `None` (Excel via
    openpyxl), teks "Unnamed: N" (pandas auto-naming untuk sel
    header kosong), atau string kosong "" (Google Sheets).
    Semuanya distandarkan jadi "unnamed" supaya konsisten dibuang
    oleh filter kolom di `prepare_dataframe`.
    """

    if col is None:
        return "unnamed"

    text = str(col).strip()

    if text == "" or text.lower() == "none":
        return "unnamed"

    if re.match(
        r"^unnamed(:\s*\d+)?$",
        text,
        flags=re.IGNORECASE,
    ):
        return "unnamed"

    text = text.replace("\n", " ")

    return re.sub(
        r"\s+",
        " ",
        text,
    )


def month_to_no(value: object) -> float:
    """
    Mengubah nama bulan menjadi nomor bulan.
    """

    if pd.isna(value):
        return np.nan

    text = str(value).strip().lower()

    for key, number in MONTH_ORDER.items():
        if key in text:
            return number

    return np.nan


def normalize_month_label(value: object) -> object:
    """
    Menstandarkan kolom label bulan (mis. Costing Month,
    Payment Schedule Month) menjadi teks 'MMM-YYYY'
    (contoh: 'Jan-2025').

    Sumbernya bisa berupa:
    - nama bulan sebagai teks (mis. "Januari 2025"); atau
    - tanggal Excel asli (kolom sumber ke-format sebagai Date,
      bukan Text — kasus umum kalau orang mengetik "Jan-25"
      tanpa memaksa format Text di Excel, sehingga Excel diam-
      diam mengubahnya jadi tanggal).

    Kedua kasus di atas distandarkan ke format yang sama supaya
    pencocokan bulan (`month_to_no`) dan tampilan di layar/Excel
    tetap konsisten, apa pun format sumbernya.
    """

    if pd.isna(value):
        return pd.NA

    if isinstance(value, datetime):
        return f"{MONTH_ABBR[value.month]}-{value.year}"

    text = str(value).strip()

    if not text:
        return pd.NA

    month_no = month_to_no(text)
    year_match = re.search(r"(19|20)\d{2}", text)

    if not pd.isna(month_no) and year_match:
        return f"{MONTH_ABBR[int(month_no)]}-{year_match.group(0)}"

    # Fallback: teks ini kemungkinan hasil `astype(str)` dari
    # kolom yang ke-detect sebagai Date oleh Excel/pandas
    # (contoh: "2025-01-15" atau "2025-01-15 00:00:00").
    #
    # `year_match` disyaratkan juga di sini supaya nama bulan
    # polos tanpa tahun (mis. "January") tidak salah diparse
    # oleh pandas jadi tanggal tahun 1 (0001-01-01).
    if year_match:
        parsed = pd.to_datetime(
            text,
            errors="coerce",
        )

        if pd.notna(parsed):
            return f"{MONTH_ABBR[parsed.month]}-{parsed.year}"

    return text


def read_raw_data_sheet(
    source: object,
) -> pd.DataFrame:
    """
    Membaca sheet `RAW DATA` dari workbook Excel seefisien
    mungkin.

    `python-calamine` (engine "calamine") jauh lebih cepat
    dibanding `openpyxl` untuk file besar (diukur ~5-10x lebih
    cepat), jadi dicoba lebih dulu. Kalau library-nya tidak
    terinstall atau workbook punya format yang tidak didukung
    calamine, otomatis jatuh ke `openpyxl` supaya tetap bisa
    dibaca.
    """

    try:
        return pd.read_excel(
            source,
            sheet_name="RAW DATA",
            header=1,
            engine="calamine",
        )

    except Exception:
        if hasattr(source, "seek"):
            source.seek(0)

        return pd.read_excel(
            source,
            sheet_name="RAW DATA",
            header=1,
            engine="openpyxl",
        )


@st.cache_data(
    show_spinner="Reading and preparing purchasing data..."
)
def load_data(
    file_bytes: bytes | None,
    file_path: str | None,
) -> pd.DataFrame:
    """
    Membaca sheet RAW DATA dari file Excel.
    """

    if file_bytes is not None:
        source = io.BytesIO(file_bytes)

    elif file_path:
        source = file_path

    else:
        raise FileNotFoundError(
            "Excel source is not available."
        )

    df = read_raw_data_sheet(source)

    return prepare_dataframe(df)


def prepare_dataframe(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Membersihkan & menyiapkan dataframe mentah (dari Excel
    maupun Google Sheets) menjadi dataset siap pakai dashboard.

    Dipisah dari `load_data` supaya logika pembersihan yang
    sama bisa dipakai ulang oleh sumber data mana pun (Excel
    upload, file lokal, atau Google Sheets), tanpa duplikasi.
    """

    df.columns = [
        normalize_col(col)
        for col in df.columns
    ]

    df = df.loc[
        :,
        df.columns != "unnamed",
    ]

    df = df.dropna(
        how="all",
    )

    normalized_lookup = {
        normalize_col(source_name): technical_name
        for technical_name, source_name
        in REQUIRED_COLUMNS.items()
    }

    rename_map = {
        col: normalized_lookup[col]
        for col in df.columns
        if col in normalized_lookup
    }

    df = df.rename(
        columns=rename_map,
    )

    text_cols = [
        "cost_center",
        "cost_category",
        "site",
        "vendor",
        "invoice_status",
        "payment_status",
        "item_category",
        "item_description",
        "unit",
        "pr_no",
        "po_no",
    ]

    for col in text_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype("string")
                .str.strip()
            )

            df[col] = df[col].replace({
                "<NA>": pd.NA,
                "nan": pd.NA,
                "None": pd.NA,
                "": pd.NA,
            })

    # Kolom label bulan ditangani terpisah (bukan lewat text_cols
    # di atas) supaya nilai tanggal asli Excel (kalau sumbernya
    # ke-format sebagai Date, bukan Text) tetap bisa dibaca
    # dengan benar, lalu distandarkan jadi teks "MMM-YYYY".
    month_label_cols = [
        "costing_month",
        "payment_schedule_month",
    ]

    for col in month_label_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .apply(normalize_month_label)
                .astype("string")
            )

    uppercase_cols = [
        "cost_category",
        "item_category",
        "invoice_status",
        "payment_status",
    ]

    for col in uppercase_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .str.upper()
            )

    for col in [
        "cost_category",
        "item_category",
    ]:
        if col in df.columns:
            df[col] = df[col].replace({
                "CHEMICALS": "CHEMICAL",
            })

    numeric_cols = [
        "qty_po",
        "qty_do",
        "price_unit",
        "grand_total_po",
        "grand_total_do",
        "total_po",
        "total_do",
        "payment_overdue_days",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce",
            )

    date_cols = [
        "pr_date",
        "po_date",
        "request_date",
        "arrival_date",
        "usage_date",
        "invoice_received_date",
        "payment_schedule",
        "actual_payout_date",
    ]

    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(
                df[col],
                errors="coerce",
            )

    if {
        "qty_po",
        "qty_do",
    }.issubset(df.columns):

        df["qty_variance"] = (
            df["qty_do"].fillna(0)
            - df["qty_po"].fillna(0)
        )

        df["fulfillment_rate"] = np.where(
            df["qty_po"] > 0,
            df["qty_do"] / df["qty_po"],
            np.nan,
        )

        df["delivery_status"] = np.select(
            [
                df["qty_do"] < df["qty_po"],
                df["qty_do"] > df["qty_po"],
                df["qty_do"] == df["qty_po"],
            ],
            [
                "SHORT DELIVERY",
                "OVER DELIVERY",
                "FULFILLED",
            ],
            default="UNKNOWN",
        )

    if {
        "grand_total_po",
        "grand_total_do",
    }.issubset(df.columns):

        df["value_variance_do_vs_po"] = (
            df["grand_total_do"].fillna(0)
            - df["grand_total_po"].fillna(0)
        )

    def date_diff(
        newer: str,
        older: str,
        output: str,
    ) -> None:
        if {
            newer,
            older,
        }.issubset(df.columns):

            df[output] = (
                df[newer]
                - df[older]
            ).dt.days

    date_diff(
        "po_date",
        "pr_date",
        "lt_pr_to_po",
    )

    date_diff(
        "arrival_date",
        "po_date",
        "lt_po_to_arrival",
    )

    date_diff(
        "invoice_received_date",
        "arrival_date",
        "lt_arrival_to_invoice",
    )

    date_diff(
        "payment_schedule",
        "invoice_received_date",
        "lt_invoice_to_schedule",
    )

    date_diff(
        "actual_payout_date",
        "payment_schedule",
        "lt_schedule_to_payout",
    )

    if "costing_month" in df.columns:
        df["month_no"] = (
            df["costing_month"]
            .apply(month_to_no)
        )

        costing_year = (
            df["costing_month"]
            .astype("string")
            .str.extract(
                r"((?:19|20)\d{2})",
                expand=False,
            )
            .astype("Float64")
        )

        # Dipakai untuk mengurutkan tren bulanan secara
        # kronologis (bukan cuma 1-12) supaya data yang
        # melewati pergantian tahun tidak tercampur urutannya.
        df["month_sort_key"] = (
            costing_year * 12
            + df["month_no"]
        )

    else:
        df["month_no"] = np.nan
        df["month_sort_key"] = np.nan

    return df


@st.cache_data(
    show_spinner="Combining multiple workbooks..."
)
def load_combined_data(
    named_file_bytes: tuple[tuple[str, bytes], ...],
) -> pd.DataFrame:
    """
    Menggabungkan beberapa workbook purchasing menjadi satu
    dataset, dengan kolom `source_file` supaya asal barisnya
    tetap bisa dilacak.
    """

    frames = []

    for file_name, file_bytes in named_file_bytes:
        single = load_data(
            file_bytes,
            None,
        ).copy()

        single["source_file"] = file_name

        frames.append(single)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(
        frames,
        ignore_index=True,
        sort=False,
    )

    return combined


# =========================================================
# GOOGLE SHEETS SOURCE (auto-sync, tanpa upload manual)
# =========================================================
GOOGLE_SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def gsheet_is_configured() -> bool:
    """
    Mengecek apakah kredensial & ID Google Sheet sudah
    disiapkan di Streamlit Secrets (Manage app -> Settings ->
    Secrets). Kalau belum, dashboard otomatis jatuh ke sumber
    lain (upload manual / file lokal) tanpa error.
    """

    if not GSPREAD_AVAILABLE:
        return False

    try:
        return (
            "gcp_service_account" in st.secrets
            and "gsheet_id" in st.secrets
        )

    except Exception:
        return False


def get_gsheet_client():
    """
    Membuat client Google Sheets dari service account yang
    disimpan di Streamlit Secrets. Kredensial ini tidak pernah
    terlihat oleh user dashboard, hanya dibaca dari server.
    """

    credentials_info = dict(
        st.secrets["gcp_service_account"]
    )

    credentials = Credentials.from_service_account_info(
        credentials_info,
        scopes=GOOGLE_SHEETS_SCOPES,
    )

    return gspread.authorize(credentials)


@st.cache_data(
    ttl=300,
    show_spinner="Syncing latest data from Google Sheets...",
)
def load_data_from_gsheet(
    sheet_id: str,
    worksheet_name: str = "RAW DATA",
    header_row: int = 1,
) -> pd.DataFrame:
    """
    Membaca data purchasing langsung dari Google Sheets, lalu
    memprosesnya lewat pipeline pembersihan yang sama dengan
    file Excel (`prepare_dataframe`).

    Hasilnya di-cache 5 menit (`ttl=300`) supaya dashboard tidak
    memanggil Google Sheets API di setiap interaksi user, tapi
    tetap otomatis dapat data terbaru tanpa perlu redeploy.
    Tombol "Refresh data" di sidebar bisa memaksa ambil versi
    terbaru sebelum 5 menit itu berakhir.
    """

    client = get_gsheet_client()

    spreadsheet = client.open_by_key(sheet_id)
    worksheet = spreadsheet.worksheet(worksheet_name)

    # Sengaja pakai get_all_values() (nilai mentah), bukan
    # get_all_records() — get_all_records() melempar error kalau
    # ada header kosong/duplikat (kasus umum kalau di Sheet ada
    # kolom pemisah/kosong). Header kosong/duplikat di sini
    # ditangani dengan cara yang sama seperti kolom "Unnamed" di
    # jalur Excel: dibuang lewat `prepare_dataframe`.
    # `value_render_option="UNFORMATTED_VALUE"` penting: tanpa ini,
    # Google Sheets mengirim angka sebagai TEKS TAMPILAN (mis.
    # "12,500,000" lengkap dengan pemisah ribuan), yang gagal
    # dibaca `pd.to_numeric` dan diam-diam jadi kosong — bikin
    # total spend di dashboard jatuh jauh lebih kecil dari
    # aslinya. Dengan opsi ini, angka dikirim sebagai angka asli.
    # `date_time_render_option="FORMATTED_STRING"` supaya kolom
    # tanggal tetap dikirim sebagai teks yang bisa dibaca
    # (mis. "1/15/2025"), bukan berubah jadi serial number.
    all_values = worksheet.get_all_values(
        value_render_option="UNFORMATTED_VALUE",
        date_time_render_option="FORMATTED_STRING",
    )

    if len(all_values) < header_row:
        return pd.DataFrame()

    header = all_values[header_row - 1]
    data_rows = all_values[header_row:]

    if not data_rows:
        return pd.DataFrame()

    df = pd.DataFrame(
        data_rows,
        columns=header,
    )

    # Sel kosong dari Google Sheets datang sebagai string ""
    # (bukan NaN asli seperti sel kosong di Excel) — disamakan
    # dulu supaya deteksi baris/kolom kosong di prepare_dataframe
    # bekerja sama seperti pada file Excel.
    df = df.replace("", np.nan)

    return prepare_dataframe(df)


# =========================================================
# FORMATTERS
# =========================================================
def rupiah(value: float) -> str:
    """
    Mengubah angka menjadi format Rupiah.
    """

    if pd.isna(value):
        return "-"

    value = float(value)
    magnitude = abs(value)

    if magnitude >= 1_000_000_000:
        result = (
            f"Rp {value / 1_000_000_000:,.2f} B"
        )

        return (
            result
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    if magnitude >= 1_000_000:
        result = (
            f"Rp {value / 1_000_000:,.2f} M"
        )

        return (
            result
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    return (
        f"Rp {value:,.0f}"
        .replace(",", ".")
    )


def number_id(
    value: float,
    decimals: int = 0,
) -> str:
    """
    Format angka Indonesia.
    """

    if pd.isna(value):
        return "-"

    text = f"{value:,.{decimals}f}"

    return (
        text
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def pct(
    value: float,
    decimals: int = 1,
) -> str:
    """
    Format angka desimal menjadi persen.
    """

    if pd.isna(value):
        return "-"

    return (
        f"{value * 100:.{decimals}f}%"
        .replace(".", ",")
    )


def safe_group_sum(
    df: pd.DataFrame,
    group_cols: list[str],
    value_col: str = "grand_total_do",
) -> pd.DataFrame:
    """
    Grouping data dengan validasi kolom.
    """

    required = set(
        group_cols + [value_col]
    )

    if not required.issubset(df.columns):
        return pd.DataFrame()

    result = (
        df
        .dropna(subset=group_cols)
        .groupby(
            group_cols,
            dropna=False,
        )[value_col]
        .sum()
        .reset_index()
        .sort_values(
            value_col,
            ascending=False,
        )
    )

    return result


def display_df(
    df: pd.DataFrame,
    currency_cols: list[str] | None = None,
    percent_cols: list[str] | None = None,
    highlight_col: str | None = None,
    highlight_map: dict[str, str] | None = None,
) -> None:
    """
    Menampilkan DataFrame dengan format kolom.

    `highlight_col` + `highlight_map` opsional dipakai untuk
    memberi warna latar semantik pada kolom kategori tertentu
    (mis. status risiko, kelas ABC, bucket aging) supaya lebih
    cepat dibaca tanpa harus menyisir tiap baris.
    """

    if df.empty:
        st.info(
            "No data is available for the current filter selection."
        )
        return

    currency_cols = currency_cols or []
    percent_cols = percent_cols or []

    view = df.copy()

    for col in percent_cols:
        if col in view.columns:
            view[col] = (
                view[col] * 100
            )

    view = view.rename(
        columns=DISPLAY_NAMES,
    )

    column_config = {}

    for col in currency_cols:
        display_name = DISPLAY_NAMES.get(
            col,
            col,
        )

        if display_name in view.columns:
            column_config[display_name] = (
                st.column_config.NumberColumn(
                    display_name,
                    format="Rp %,.0f",
                )
            )

    for col in percent_cols:
        display_name = DISPLAY_NAMES.get(
            col,
            col,
        )

        if display_name in view.columns:
            column_config[display_name] = (
                st.column_config.NumberColumn(
                    display_name,
                    format="%.1f%%",
                )
            )

    render_target = view

    if highlight_col and highlight_map:
        highlight_display_name = DISPLAY_NAMES.get(
            highlight_col,
            highlight_col,
        )

        if highlight_display_name in view.columns:

            def apply_highlight(value):
                background = highlight_map.get(
                    str(value)
                )

                if not background:
                    return ""

                return (
                    f"background-color: {background}; "
                    "color: #FFFFFF; "
                    "font-weight: 600;"
                )

            render_target = view.style.map(
                apply_highlight,
                subset=[highlight_display_name],
            )

    st.dataframe(
        render_target,
        use_container_width=True,
        hide_index=True,
        column_config=column_config,
    )


# =========================================================
# EXCEL EXPORT
# =========================================================
def clean_sheet_name(name: str) -> str:
    """
    Membersihkan nama worksheet.
    """

    return re.sub(
        r"[\\/*?:\[\]]",
        "-",
        name,
    )[:31]


def to_excel_bytes(
    sheets: dict[str, pd.DataFrame],
    title: str,
) -> bytes:
    """
    Membuat file Excel di memory.
    """

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        datetime_format="dd-mmm-yyyy",
    ) as writer:

        workbook = writer.book

        title_format = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "font_color": "#111827",
            "bottom": 2,
            "bottom_color": "#334155",
        })

        header_format = workbook.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#334155",
            "border": 0,
            "align": "center",
            "valign": "vcenter",
        })

        currency_format = workbook.add_format({
            "num_format": (
                "Rp #,##0;"
                "[Red]-Rp #,##0"
            ),
        })

        percent_format = workbook.add_format({
            "num_format": "0.0%",
        })

        date_format = workbook.add_format({
            "num_format": "dd-mmm-yyyy",
        })

        meta_format = workbook.add_format({
            "font_color": "#6B7280",
            "italic": True,
        })

        for sheet_name, source_df in sheets.items():
            export_df = source_df.copy()

            export_df = export_df.rename(
                columns=DISPLAY_NAMES,
            )

            safe_name = clean_sheet_name(
                sheet_name
            )

            export_df.to_excel(
                writer,
                sheet_name=safe_name,
                startrow=4,
                index=False,
            )

            worksheet = writer.sheets[
                safe_name
            ]

            worksheet.hide_gridlines(2)

            worksheet.write(
                0,
                0,
                title,
                title_format,
            )

            worksheet.write(
                1,
                0,
                (
                    "Generated: "
                    f"{datetime.now():%d %b %Y %H:%M}"
                ),
                meta_format,
            )

            worksheet.write(
                2,
                0,
                f"Rows: {len(export_df):,}",
                meta_format,
            )

            for col_index, col_name in enumerate(
                export_df.columns
            ):
                worksheet.write(
                    4,
                    col_index,
                    col_name,
                    header_format,
                )

                if len(export_df):
                    sample_length = (
                        export_df[col_name]
                        .astype(str)
                        .str.len()
                        .quantile(0.90)
                    )

                    # Kolom yang isinya kosong semua (NaN semua)
                    # membuat quantile() ikut menghasilkan NaN di
                    # sebagian versi pandas, jadi perlu fallback
                    # supaya int() di bawah tidak error.
                    if pd.isna(sample_length):
                        sample_length = len(
                            col_name
                        )

                else:
                    sample_length = len(
                        col_name
                    )

                width = min(
                    max(
                        len(col_name) + 2,
                        int(sample_length) + 2,
                    ),
                    34,
                )

                cell_format = None

                original_name = next(
                    (
                        key
                        for key, value
                        in DISPLAY_NAMES.items()
                        if value == col_name
                    ),
                    col_name,
                )

                currency_tokens = [
                    "price",
                    "total",
                    "spend",
                    "saving",
                    "value",
                ]

                if any(
                    token in original_name.lower()
                    for token in currency_tokens
                ):
                    cell_format = (
                        currency_format
                    )

                elif (
                    "rate"
                    in original_name.lower()
                    or "share"
                    in original_name.lower()
                ):
                    cell_format = (
                        percent_format
                    )

                elif pd.api.types.is_datetime64_any_dtype(
                    export_df[col_name]
                ):
                    cell_format = (
                        date_format
                    )

                worksheet.set_column(
                    col_index,
                    col_index,
                    width,
                    cell_format,
                )

            worksheet.freeze_panes(
                5,
                0,
            )

            if len(export_df):
                worksheet.autofilter(
                    4,
                    0,
                    4 + len(export_df),
                    len(export_df.columns) - 1,
                )

    return output.getvalue()


def export_button(
    label: str,
    file_name: str,
    sheets: dict[str, pd.DataFrame],
    title: str,
    key: str,
) -> None:
    """
    Tombol export Excel.
    """

    st.download_button(
        label=f"↓ {label}",
        data=to_excel_bytes(
            sheets,
            title,
        ),
        file_name=file_name,
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key=key,
        use_container_width=False,
    )


# =========================================================
# CHART HELPERS
# =========================================================
def clean_layout(
    fig: go.Figure,
    height: int = 390,
) -> go.Figure:
    """
    Menyesuaikan chart dengan light atau dark mode.
    """

    dark_mode = bool(
        st.session_state.get(
            "dark_mode",
            False,
        )
    )

    if dark_mode:
        text_color = "#F8FAFC"
        muted_color = "#94A3B8"
        line_color = "#334155"
        hover_background = "#1E293B"

    else:
        text_color = "#111827"
        muted_color = "#6B7280"
        line_color = "#E5E7EB"
        hover_background = "#FFFFFF"

    fig.update_layout(
        height=height,
        margin={
            "l": 12,
            "r": 12,
            "t": 48,
            "b": 12,
        },
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={
            "family": "Arial",
            "color": text_color,
            "size": 12,
        },
        title_font={
            "size": 15,
            "color": text_color,
        },
        legend_title_text="",
        legend={
            "font": {
                "color": text_color,
            }
        },
        hoverlabel={
            "bgcolor": hover_background,
            "font_color": text_color,
        },
    )

    fig.update_xaxes(
        showgrid=False,
        linecolor=line_color,
        tickfont={
            "color": muted_color,
        },
        title_font={
            "color": text_color,
        },
    )

    fig.update_yaxes(
        gridcolor=line_color,
        zeroline=False,
        tickfont={
            "color": muted_color,
        },
        title_font={
            "color": text_color,
        },
    )

    return fig


def bar_chart(
    data: pd.DataFrame,
    x: str,
    y: str,
    title: str,
    horizontal: bool = False,
    height: int = 390,
) -> go.Figure:
    """
    Membuat bar chart.
    """

    if horizontal:
        fig = px.bar(
            data,
            x=x,
            y=y,
            orientation="h",
            title=title,
            color_discrete_sequence=[
                COLOR["accent"],
            ],
        )

        fig.update_layout(
            yaxis={
                "categoryorder":
                "total ascending",
            }
        )

        hover_template = (
            "%{y}<br>"
            "Rp %{x:,.0f}"
            "<extra></extra>"
        )

    else:
        fig = px.bar(
            data,
            x=x,
            y=y,
            title=title,
            color_discrete_sequence=[
                COLOR["accent"],
            ],
        )

        hover_template = (
            "%{x}<br>"
            "Rp %{y:,.0f}"
            "<extra></extra>"
        )

    fig.update_traces(
        marker_line_width=0,
        hovertemplate=hover_template,
    )

    return clean_layout(
        fig,
        height,
    )


def section_heading(
    label: str,
    title: str,
    description: str = "",
) -> None:
    """
    Judul modul.
    """

    st.markdown(
        (
            '<div class="section-label">'
            f"{label}"
            "</div>"
        ),
        unsafe_allow_html=True,
    )

    st.subheader(title)

    if description:
        st.caption(description)


def insight_box(
    title: str,
    text: str,
) -> None:
    """
    Kotak executive insight.
    """

    st.markdown(
        (
            '<div class="insight-box">'
            f'<div class="insight-title">{title}</div>'
            f'<div class="insight-text">{text}</div>'
            "</div>"
        ),
        unsafe_allow_html=True,
    )


def reading_guide(
    points: list[str],
    title: str = "Cara membaca & tindak lanjut operasional",
) -> None:
    """
    Kotak panduan cara membaca sebuah section analitik,
    supaya tim operasional tahu tindakan apa yang perlu
    diambil dari angka yang ditampilkan.
    """

    if not points:
        return

    items_html = "".join(
        f"<li>{point}</li>"
        for point in points
    )

    st.markdown(
        (
            '<div class="guide-box">'
            f'<div class="guide-title">{title}</div>'
            f'<div class="guide-text"><ul>{items_html}</ul></div>'
            "</div>"
        ),
        unsafe_allow_html=True,
    )


# =========================================================
# SEMANTIC STATUS COLORS
# =========================================================
STATUS_COLORS = {
    "good": COLOR["success"],
    "warn": COLOR["warning"],
    "bad": COLOR["danger"],
    "neutral": COLOR["muted"],
}

RISK_QUADRANT_COLORS = {
    "CRITICAL - manage closely": COLOR["danger"],
    "STRATEGIC - maintain": COLOR["success"],
    "MONITOR - watch closely": COLOR["warning"],
    "ROUTINE - low touch": COLOR["accent_2"],
}

ABC_CLASS_COLORS = {
    "A": COLOR["danger"],
    "B": COLOR["warning"],
    "C": COLOR["accent_2"],
}

AGING_BUCKET_COLORS = {
    "Not yet due": COLOR["success"],
    "1-30 days": COLOR["accent_2"],
    "31-60 days": COLOR["warning"],
    "61-90 days": COLOR["danger"],
    "90+ days": COLOR["danger"],
}

DELIVERY_STATUS_COLORS = {
    "SHORT DELIVERY": COLOR["danger"],
    "OVER DELIVERY": COLOR["warning"],
    "FULFILLED": COLOR["success"],
    "UNKNOWN": COLOR["accent_2"],
}


def pill(
    text: str,
    level: str = "neutral",
) -> str:
    """
    Membuat badge kecil berwarna sesuai status
    (good / warn / bad / neutral). Label teks selalu
    ikut ditampilkan supaya makna tidak hanya
    mengandalkan warna.
    """

    background = STATUS_COLORS.get(
        level,
        STATUS_COLORS["neutral"],
    )

    return (
        '<span class="tag-pill" '
        f'style="background:{background}; color:#FFFFFF;">'
        f"{text}"
        "</span>"
    )


def status_pill_markdown(
    label: str,
    text: str,
    level: str,
) -> None:
    """
    Menampilkan satu baris "Label: [pill]" di bawah metric,
    supaya status baik/butuh perhatian/kritis langsung
    kelihatan tanpa buka tabel.
    """

    st.markdown(
        (
            f'<div style="margin: -6px 0 14px 0; font-size: 0.85rem;">'
            f"{label}: "
            f"{pill(text, level)}"
            "</div>"
        ),
        unsafe_allow_html=True,
    )


def fulfillment_level(
    rate: float,
) -> tuple[str, str]:
    """
    Menentukan level status untuk fulfillment rate.
    """

    if pd.isna(rate):
        return (
            "neutral",
            "No Data",
        )

    if rate >= 0.95:
        return (
            "good",
            "On Track",
        )

    if rate >= 0.85:
        return (
            "warn",
            "Needs Attention",
        )

    return (
        "bad",
        "Critical Shortfall",
    )


def overdue_exposure_level(
    share: float,
) -> tuple[str, str]:
    """
    Menentukan level status untuk porsi exposure overdue.
    """

    if pd.isna(share):
        return (
            "neutral",
            "No Data",
        )

    if share < 0.05:
        return (
            "good",
            "Healthy",
        )

    if share < 0.15:
        return (
            "warn",
            "Watch",
        )

    return (
        "bad",
        "High Risk",
    )


# =========================================================
# PER-TAB LOCAL FILTERS
# =========================================================
LOCAL_FILTER_DIMENSIONS = [
    (
        "source_file",
        "Source File",
    ),
    (
        "costing_month",
        "Costing Month",
    ),
    (
        "site",
        "Site",
    ),
    (
        "cost_center",
        "Cost Center",
    ),
    (
        "vendor",
        "Vendor",
    ),
    (
        "cost_category",
        "Cost Category",
    ),
    (
        "item_category",
        "Item Category",
    ),
    (
        "payment_status",
        "Payment Status",
    ),
]


def render_local_filters(
    df: pd.DataFrame,
    scope: str,
) -> pd.DataFrame:
    """
    Filter lokal khusus untuk satu tab/halaman.

    Setiap tab sudah dibungkus @st.fragment, dan widget filter
    di sini pakai key yang unik per `scope` — jadi mengubah
    filter di satu tab (mis. Price) HANYA menghitung ulang tab
    itu sendiri, tab lain tidak ikut ke-render ulang. Filter
    langsung aktif begitu dipilih, tidak perlu tombol Apply
    karena reruns di sini sudah murah/terisolasi per tab.
    """

    available_specs = [
        (col, label)
        for col, label in LOCAL_FILTER_DIMENSIONS
        if col in df.columns
    ]

    has_date = (
        "arrival_date" in df.columns
        and not df["arrival_date"].dropna().empty
    )

    if not available_specs and not has_date:
        return df

    with st.expander(
        "Filter halaman ini",
        expanded=False,
    ):
        st.caption(
            "Filter di sini hanya berlaku untuk tab ini — "
            "tab lain tidak terpengaruh dan tidak ikut "
            "dihitung ulang."
        )

        column_count = min(
            len(available_specs),
            4,
        ) or 1

        widget_columns = st.columns(
            column_count
        )

        for index, (col, label) in enumerate(
            available_specs
        ):
            options = sorted(
                df[col]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            widget_key = f"locfilt_{scope}_{col}"

            with widget_columns[
                index % column_count
            ]:
                st.multiselect(
                    label,
                    options,
                    default=st.session_state.get(
                        widget_key,
                        [],
                    ),
                    placeholder=f"All {label}",
                    key=widget_key,
                )

        date_range = None

        if has_date:
            valid_dates = (
                df["arrival_date"]
                .dropna()
            )

            min_date = (
                valid_dates
                .min()
                .date()
            )

            max_date = (
                valid_dates
                .max()
                .date()
            )

            date_key = f"locfilt_{scope}_arrival_range"

            date_range = st.date_input(
                "Arrival Date",
                value=st.session_state.get(
                    date_key,
                    (
                        min_date,
                        max_date,
                    ),
                ),
                min_value=min_date,
                max_value=max_date,
                key=date_key,
            )

    filtered = df.copy()

    active_count = 0

    for col, _ in available_specs:
        widget_key = f"locfilt_{scope}_{col}"

        selected = st.session_state.get(
            widget_key,
            [],
        )

        if selected:
            active_count += 1

            filtered = filtered[
                filtered[col]
                .astype(str)
                .isin(selected)
            ]

    if (
        has_date
        and isinstance(
            date_range,
            (tuple, list),
        )
        and len(date_range) == 2
    ):
        if (
            date_range[0] != min_date
            or date_range[1] != max_date
        ):
            active_count += 1

        start = pd.Timestamp(
            date_range[0]
        )

        end = (
            pd.Timestamp(date_range[1])
            + pd.Timedelta(days=1)
            - pd.Timedelta(seconds=1)
        )

        filtered = filtered[
            filtered["arrival_date"]
            .between(
                start,
                end,
            )
        ]

    filter_note = (
        f"{number_id(len(filtered))} baris ditampilkan"
    )

    if active_count:
        filter_note += (
            f" · {active_count} filter aktif di tab ini"
        )

    st.caption(
        filter_note
    )

    return filtered


def reset_all_local_filters() -> None:
    """
    Menghapus semua state filter lokal di semua tab —
    dipanggil saat sumber data berubah (ganti file / mode
    combine) supaya tidak ada filter lama yang diam-diam
    menutup data baru.
    """

    for key in list(st.session_state.keys()):
        if key.startswith("locfilt_"):
            del st.session_state[key]


# =========================================================
# ANALYTICAL DATASETS
# =========================================================
@st.cache_data(
    show_spinner=False,
)
def executive_datasets(
    df: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    """
    Dataset ringkasan executive.
    """

    monthly = safe_group_sum(
        df,
        [
            "month_no",
            "costing_month",
        ],
    )

    if (
        not monthly.empty
        and "month_no" in monthly.columns
        and "month_sort_key" in df.columns
    ):
        # Ambil month_sort_key (tahun x 12 + bulan) lewat left-
        # merge, bukan dijadikan kolom group langsung — supaya
        # baris yang label bulannya tidak menyebut tahun (mis.
        # cuma "Januari" tanpa tahun) tetap ikut, bukan malah
        # ke-drop dari grouping.
        sort_key_lookup = (
            df[
                [
                    "month_no",
                    "costing_month",
                    "month_sort_key",
                ]
            ]
            .dropna(
                subset=[
                    "month_no",
                    "costing_month",
                ]
            )
            .drop_duplicates(
                subset=[
                    "month_no",
                    "costing_month",
                ]
            )
        )

        monthly = monthly.merge(
            sort_key_lookup,
            on=[
                "month_no",
                "costing_month",
            ],
            how="left",
        )

        # Kalau tahunnya tidak terdeteksi, fallback ke urutan
        # per nomor bulan saja (perilaku lama) daripada dibuang.
        monthly["month_sort_key"] = (
            monthly["month_sort_key"]
            .fillna(
                monthly["month_no"]
            )
        )

        monthly = monthly.sort_values(
            "month_sort_key"
        )

    elif (
        not monthly.empty
        and "month_no" in monthly.columns
    ):
        monthly = monthly.sort_values(
            "month_no"
        )

    if (
        not monthly.empty
        and "grand_total_po" in df.columns
    ):
        monthly_po = (
            df
            .dropna(
                subset=[
                    "month_no",
                    "costing_month",
                ]
            )
            .groupby(
                [
                    "month_no",
                    "costing_month",
                ],
                dropna=False,
            )["grand_total_po"]
            .sum()
            .reset_index()
        )

        monthly = monthly.merge(
            monthly_po,
            on=[
                "month_no",
                "costing_month",
            ],
            how="left",
        )

        monthly["value_variance_do_vs_po"] = (
            monthly["grand_total_do"].fillna(0)
            - monthly["grand_total_po"].fillna(0)
        )

    vendor = safe_group_sum(
        df,
        ["vendor"],
    ).head(15)

    category = safe_group_sum(
        df,
        ["cost_category"],
    )

    site = safe_group_sum(
        df,
        ["site"],
    )

    return {
        "Monthly Spend": monthly,
        "Top Vendors": vendor,
        "Category Spend": category,
        "Site Spend": site,
    }


@st.cache_data(
    show_spinner=False,
)
def vendor_scorecard(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Membuat vendor scorecard.
    """

    required = {
        "vendor",
        "grand_total_do",
    }

    if not required.issubset(df.columns):
        return pd.DataFrame()

    aggregations = {
        "spend_do": (
            "grand_total_do",
            "sum",
        ),
    }

    if "po_no" in df.columns:
        aggregations["po_count"] = (
            "po_no",
            "nunique",
        )

    if "item_description" in df.columns:
        aggregations["item_count"] = (
            "item_description",
            "nunique",
        )

    if "fulfillment_rate" in df.columns:
        aggregations[
            "avg_fulfillment_rate"
        ] = (
            "fulfillment_rate",
            "mean",
        )

    if "lt_po_to_arrival" in df.columns:
        aggregations[
            "avg_delivery_days"
        ] = (
            "lt_po_to_arrival",
            "mean",
        )

    if "payment_overdue_days" in df.columns:
        aggregations[
            "avg_payment_overdue_days"
        ] = (
            "payment_overdue_days",
            "mean",
        )

    scorecard = (
        df
        .dropna(subset=["vendor"])
        .groupby("vendor")
        .agg(**aggregations)
        .reset_index()
    )

    total = (
        scorecard["spend_do"]
        .sum()
    )

    scorecard["spend_share"] = np.where(
        total > 0,
        scorecard["spend_do"] / total,
        0,
    )

    return scorecard.sort_values(
        "spend_do",
        ascending=False,
    )


@st.cache_data(
    show_spinner=False,
)
def price_opportunity(
    df: pd.DataFrame,
    min_spend: float,
    min_txn: int,
) -> pd.DataFrame:
    """
    Mencari variasi harga dan potential saving.
    """

    needed = {
        "item_description",
        "price_unit",
        "grand_total_do",
        "qty_do",
    }

    if not needed.issubset(df.columns):
        return pd.DataFrame()

    work = df.dropna(
        subset=[
            "item_description",
            "price_unit",
        ]
    ).copy()

    work = work[
        work["price_unit"] > 0
    ]

    aggregation = {
        "transaction_count": (
            "item_description",
            "size",
        ),
        "qty_do": (
            "qty_do",
            "sum",
        ),
        "spend_do": (
            "grand_total_do",
            "sum",
        ),
        "min_price": (
            "price_unit",
            "min",
        ),
        "avg_price": (
            "price_unit",
            "mean",
        ),
        "median_price": (
            "price_unit",
            "median",
        ),
        "max_price": (
            "price_unit",
            "max",
        ),
    }

    if "vendor" in work.columns:
        aggregation["vendor_count"] = (
            "vendor",
            "nunique",
        )

    item = (
        work
        .groupby(
            "item_description"
        )
        .agg(**aggregation)
        .reset_index()
    )

    item["price_gap"] = (
        item["max_price"]
        - item["min_price"]
    )

    item["price_gap_rate"] = np.where(
        item["min_price"] > 0,
        (
            item["price_gap"]
            / item["min_price"]
        ),
        np.nan,
    )

    item["potential_saving"] = np.maximum(
        (
            item["avg_price"]
            - item["min_price"]
        )
        * item["qty_do"],
        0,
    )

    item = item[
        (
            item["spend_do"]
            >= min_spend
        )
        & (
            item["transaction_count"]
            >= min_txn
        )
    ]

    return item.sort_values(
        [
            "potential_saving",
            "spend_do",
        ],
        ascending=False,
    )


@st.cache_data(
    show_spinner=False,
)
def vendor_top_items(
    df: pd.DataFrame,
    top_n: int = 50,
) -> pd.DataFrame:
    """
    Mengambil top 50 item berdasarkan spend
    untuk setiap vendor.
    """

    required = {
        "vendor",
        "item_description",
        "grand_total_do",
    }

    if not required.issubset(df.columns):
        return pd.DataFrame()

    aggregation = {
        "spend_do": (
            "grand_total_do",
            "sum",
        ),
        "transaction_count": (
            "item_description",
            "size",
        ),
    }

    if "qty_do" in df.columns:
        aggregation["qty_do"] = (
            "qty_do",
            "sum",
        )

    if "price_unit" in df.columns:
        aggregation["avg_price"] = (
            "price_unit",
            "mean",
        )

        aggregation["min_price"] = (
            "price_unit",
            "min",
        )

        aggregation["max_price"] = (
            "price_unit",
            "max",
        )

    if "unit" in df.columns:
        aggregation["unit"] = (
            "unit",
            lambda series: (
                series
                .dropna()
                .mode()
                .iloc[0]
                if not series
                .dropna()
                .empty
                else pd.NA
            ),
        )

    result = (
        df
        .dropna(
            subset=[
                "vendor",
                "item_description",
            ]
        )
        .groupby(
            [
                "vendor",
                "item_description",
            ],
            dropna=False,
        )
        .agg(**aggregation)
        .reset_index()
        .sort_values(
            [
                "vendor",
                "spend_do",
            ],
            ascending=[
                True,
                False,
            ],
        )
    )

    result["vendor_rank"] = (
        result
        .groupby("vendor")["spend_do"]
        .rank(
            method="first",
            ascending=False,
        )
        .astype(int)
    )

    result = result[
        result["vendor_rank"] <= top_n
    ]

    return result.sort_values(
        [
            "vendor",
            "vendor_rank",
        ]
    )


@st.cache_data(
    show_spinner=False,
)
def vendor_price_comparison(
    df: pd.DataFrame,
    vendors: list[str],
) -> pd.DataFrame:
    """
    Membandingkan rata-rata harga item
    pada maksimal tiga vendor.
    """

    required = {
        "vendor",
        "item_description",
        "price_unit",
    }

    if (
        not required.issubset(df.columns)
        or not vendors
    ):
        return pd.DataFrame()

    work = df[
        df["vendor"].isin(vendors)
    ].copy()

    work = work.dropna(
        subset=[
            "item_description",
            "price_unit",
        ]
    )

    work = work[
        work["price_unit"] > 0
    ]

    if work.empty:
        return pd.DataFrame()

    summary = (
        work
        .groupby(
            [
                "item_description",
                "vendor",
            ],
            dropna=False,
        )
        .agg(
            avg_price=(
                "price_unit",
                "mean",
            ),
            min_price=(
                "price_unit",
                "min",
            ),
            max_price=(
                "price_unit",
                "max",
            ),
            transaction_count=(
                "price_unit",
                "size",
            ),
        )
        .reset_index()
    )

    pivot = summary.pivot_table(
        index="item_description",
        columns="vendor",
        values="avg_price",
        aggfunc="mean",
    )

    pivot = pivot.reindex(
        columns=vendors,
    )

    pivot["lowest_price"] = (
        pivot[vendors]
        .min(
            axis=1,
            skipna=True,
        )
    )

    pivot["highest_price"] = (
        pivot[vendors]
        .max(
            axis=1,
            skipna=True,
        )
    )

    pivot["price_gap"] = (
        pivot["highest_price"]
        - pivot["lowest_price"]
    )

    pivot["price_gap_rate"] = np.where(
        pivot["lowest_price"] > 0,
        (
            pivot["price_gap"]
            / pivot["lowest_price"]
        ),
        np.nan,
    )

    pivot["best_vendor"] = (
        pivot[vendors]
        .idxmin(
            axis=1,
            skipna=True,
        )
    )

    pivot["vendor_coverage"] = (
        pivot[vendors]
        .notna()
        .sum(axis=1)
    )

    pivot = pivot.reset_index()

    return pivot.sort_values(
        [
            "vendor_coverage",
            "price_gap",
        ],
        ascending=[
            False,
            False,
        ],
    )


def time_intelligence(
    monthly: pd.DataFrame,
    forecast_periods: int = 2,
    trend_window: int = 3,
) -> pd.DataFrame:
    """
    Menambahkan growth rate, moving average, dan forecast
    linear sederhana ke dataset spend bulanan.

    `monthly` diharapkan sudah terurut berdasarkan month_no
    dan punya kolom costing_month & grand_total_do.
    """

    if (
        monthly.empty
        or "grand_total_do" not in monthly.columns
    ):
        return pd.DataFrame()

    work = monthly.reset_index(
        drop=True,
    ).copy()

    work["is_forecast"] = False

    work["mom_growth"] = (
        work["grand_total_do"]
        .pct_change()
    )

    work["moving_avg"] = (
        work["grand_total_do"]
        .rolling(
            window=trend_window,
            min_periods=1,
        )
        .mean()
    )

    history = work[
        "grand_total_do"
    ].to_numpy(
        dtype=float,
    )

    n_points = len(history)

    if n_points >= 2:
        x = np.arange(n_points)

        slope, intercept = np.polyfit(
            x,
            history,
            1,
        )

        future_rows = []

        for step in range(1, forecast_periods + 1):
            future_index = n_points + step - 1

            forecast_value = float(
                slope * future_index
                + intercept
            )

            future_rows.append({
                "costing_month": (
                    f"Forecast +{step}"
                ),
                "grand_total_do": max(
                    forecast_value,
                    0.0,
                ),
                "is_forecast": True,
            })

        forecast_df = pd.DataFrame(
            future_rows
        )

        work = pd.concat(
            [
                work,
                forecast_df,
            ],
            ignore_index=True,
        )

    return work


@st.cache_data(
    show_spinner=False,
)
def vendor_segmentation(
    scorecard: pd.DataFrame,
) -> pd.DataFrame:
    """
    Memberi skor risiko komposit ke vendor scorecard dan
    mengelompokkan vendor ke kuadran spend vs risiko.
    """

    if scorecard.empty:
        return pd.DataFrame()

    work = scorecard.copy()

    def normalize(
        series: pd.Series,
        invert: bool = False,
    ) -> pd.Series:
        values = pd.to_numeric(
            series,
            errors="coerce",
        )

        span = (
            values.max()
            - values.min()
        )

        if pd.isna(span) or span == 0:
            scaled = pd.Series(
                0.5,
                index=values.index,
            )

        else:
            scaled = (
                (values - values.min())
                / span
            )

        scaled = scaled.fillna(0.5)

        if invert:
            scaled = 1 - scaled

        return scaled

    risk_components = []

    if "avg_fulfillment_rate" in work.columns:
        risk_components.append(
            normalize(
                work["avg_fulfillment_rate"],
                invert=True,
            )
        )

    if "avg_delivery_days" in work.columns:
        risk_components.append(
            normalize(
                work["avg_delivery_days"],
                invert=False,
            )
        )

    if "avg_payment_overdue_days" in work.columns:
        risk_components.append(
            normalize(
                work["avg_payment_overdue_days"],
                invert=False,
            )
        )

    if "spend_share" in work.columns:
        risk_components.append(
            normalize(
                work["spend_share"],
                invert=False,
            )
        )

    if risk_components:
        risk_matrix = pd.concat(
            risk_components,
            axis=1,
        )

        work["risk_score"] = (
            risk_matrix.mean(axis=1) * 100
        )

    else:
        work["risk_score"] = 50.0

    spend_median = (
        work["spend_do"].median()
        if "spend_do" in work.columns
        else 0
    )

    risk_median = (
        work["risk_score"].median()
    )

    def label_quadrant(row) -> str:
        high_spend = (
            row.get("spend_do", 0)
            >= spend_median
        )

        high_risk = (
            row.get("risk_score", 0)
            >= risk_median
        )

        if high_spend and high_risk:
            return "CRITICAL - manage closely"

        if high_spend and not high_risk:
            return "STRATEGIC - maintain"

        if not high_spend and high_risk:
            return "MONITOR - watch closely"

        return "ROUTINE - low touch"

    work["risk_quadrant"] = work.apply(
        label_quadrant,
        axis=1,
    )

    return work.sort_values(
        [
            "risk_score",
            "spend_do",
        ],
        ascending=[
            False,
            False,
        ],
    )


@st.cache_data(
    show_spinner=False,
)
def abc_classification(
    df: pd.DataFrame,
    dimension: str,
    value_col: str = "grand_total_do",
    class_a_cutoff: float = 0.80,
    class_b_cutoff: float = 0.95,
) -> pd.DataFrame:
    """
    Klasifikasi ABC berdasarkan kontribusi kumulatif spend
    (Pareto analysis).
    """

    required = {
        dimension,
        value_col,
    }

    if not required.issubset(df.columns):
        return pd.DataFrame()

    grouped = (
        df
        .dropna(subset=[dimension])
        .groupby(dimension)[value_col]
        .sum()
        .reset_index()
        .sort_values(
            value_col,
            ascending=False,
        )
        .reset_index(drop=True)
    )

    total_value = grouped[value_col].sum()

    if total_value <= 0:
        return pd.DataFrame()

    grouped["cumulative_value"] = (
        grouped[value_col].cumsum()
    )

    grouped["cumulative_share"] = (
        grouped["cumulative_value"]
        / total_value
    )

    grouped["spend_share"] = (
        grouped[value_col]
        / total_value
    )

    grouped["rank"] = (
        grouped.index + 1
    )

    grouped["abc_class"] = np.select(
        [
            grouped["cumulative_share"]
            <= class_a_cutoff,

            grouped["cumulative_share"]
            <= class_b_cutoff,
        ],
        [
            "A",
            "B",
        ],
        default="C",
    )

    return grouped


@st.cache_data(
    show_spinner=False,
)
def payment_aging(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Membuat bucket umur tunggakan (aging) untuk exposure
    pembayaran yang masih OPEN.
    """

    required = {
        "payment_status",
        "grand_total_do",
    }

    if not required.issubset(df.columns):
        return pd.DataFrame()

    work = df.copy()

    work["payment_status"] = (
        work["payment_status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    open_only = work[
        work["payment_status"] == "OPEN"
    ].copy()

    if open_only.empty:
        return pd.DataFrame()

    if "payment_overdue_days" in open_only.columns:
        overdue_days = pd.to_numeric(
            open_only["payment_overdue_days"],
            errors="coerce",
        ).fillna(0)

    else:
        overdue_days = pd.Series(
            0,
            index=open_only.index,
            dtype=float,
        )

    bucket_labels = [
        "Not yet due",
        "1-30 days",
        "31-60 days",
        "61-90 days",
        "90+ days",
    ]

    bucket_order = pd.cut(
        overdue_days,
        bins=[
            -np.inf,
            0,
            30,
            60,
            90,
            np.inf,
        ],
        labels=bucket_labels,
        right=True,
    )

    open_only["aging_bucket"] = bucket_order

    aggregation = {
        "exposure_value": (
            "grand_total_do",
            "sum",
        ),
        "line_count": (
            "grand_total_do",
            "size",
        ),
    }

    summary = (
        open_only
        .groupby(
            "aging_bucket",
            observed=False,
        )
        .agg(**aggregation)
        .reindex(bucket_labels)
        .fillna(0)
        .reset_index()
        .rename(
            columns={
                "aging_bucket": "bucket",
            }
        )
    )

    total_exposure = (
        summary["exposure_value"].sum()
    )

    summary["exposure_share"] = np.where(
        total_exposure > 0,
        summary["exposure_value"] / total_exposure,
        0,
    )

    return summary


# =========================================================
# EXECUTIVE MODULE
# =========================================================
@st.fragment
def show_executive(
    df: pd.DataFrame,
) -> None:
    section_heading(
        "Decision View",
        "Executive Overview",
        (
            "The most important procurement, "
            "cost, and operational signals."
        ),
    )

    df = render_local_filters(df, "executive")

    total_do = df.get(
        "grand_total_do",
        pd.Series(dtype=float),
    ).sum()

    total_po = df.get(
        "grand_total_po",
        pd.Series(dtype=float),
    ).sum()

    variance = (
        total_do
        - total_po
    )

    po_count = df.get(
        "po_no",
        pd.Series(dtype="string"),
    ).nunique(
        dropna=True
    )

    vendor_count = df.get(
        "vendor",
        pd.Series(dtype="string"),
    ).nunique(
        dropna=True
    )

    fulfillment = (
        df.get(
            "fulfillment_rate",
            pd.Series(dtype=float),
        )
        .replace(
            [
                np.inf,
                -np.inf,
            ],
            np.nan,
        )
        .mean()
    )

    if (
        "payment_status" in df.columns
        and "grand_total_do" in df.columns
    ):
        open_payment = df.loc[
            df["payment_status"].eq("OPEN"),
            "grand_total_do",
        ].sum()

    else:
        open_payment = 0

    c1, c2, c3, c4, c5, c6 = (
        st.columns(6)
    )

    c1.metric(
        "Actual Spend / DO",
        rupiah(total_do),
    )

    c2.metric(
        "Committed / PO",
        rupiah(total_po),
        delta=rupiah(variance),
        delta_color="inverse",
    )

    c3.metric(
        "Purchase Orders",
        number_id(po_count),
    )

    c4.metric(
        "Active Vendors",
        number_id(vendor_count),
    )

    c5.metric(
        "Fulfillment",
        pct(fulfillment),
    )

    c6.metric(
        "Open Payment",
        rupiah(open_payment),
    )

    fulfillment_level_key, fulfillment_label = (
        fulfillment_level(
            fulfillment
        )
    )

    status_pill_markdown(
        "Fulfillment status",
        fulfillment_label,
        fulfillment_level_key,
    )

    datasets = executive_datasets(
        df
    )

    monthly = datasets[
        "Monthly Spend"
    ]

    vendors = datasets[
        "Top Vendors"
    ]

    left, right = st.columns([
        1.2,
        1,
    ])

    with left:
        if not monthly.empty:
            figure = go.Figure()

            figure.add_trace(
                go.Bar(
                    x=monthly["costing_month"],
                    y=monthly["grand_total_do"],
                    name="Actual (DO)",
                    marker_color=COLOR["accent"],
                )
            )

            if "grand_total_po" in monthly.columns:
                figure.add_trace(
                    go.Bar(
                        x=monthly["costing_month"],
                        y=monthly["grand_total_po"],
                        name="Committed (PO)",
                        marker_color=COLOR["accent_2"],
                    )
                )

            if "value_variance_do_vs_po" in monthly.columns:
                figure.add_trace(
                    go.Scatter(
                        x=monthly["costing_month"],
                        y=monthly["value_variance_do_vs_po"],
                        name="Variance (DO - PO)",
                        mode="lines+markers",
                        yaxis="y2",
                        line={
                            "color": COLOR["warning"],
                            "width": 2,
                            "dash": "dot",
                        },
                    )
                )

            figure.update_layout(
                title="Monthly Spend: Actual vs Committed",
                barmode="group",
                yaxis={
                    "title": "Spend",
                },
                yaxis2={
                    "title": "Variance",
                    "overlaying": "y",
                    "side": "right",
                    "showgrid": False,
                },
            )

            st.plotly_chart(
                clean_layout(
                    figure,
                    height=390,
                ),
                use_container_width=True,
            )

            reading_guide([
                (
                    "Bandingkan tinggi bar <b>Actual (DO)</b> vs "
                    "<b>Committed (PO)</b> tiap bulan — gap besar "
                    "berarti realisasi jauh dari komitmen awal."
                ),
                (
                    "Garis putus-putus (Variance) di atas nol "
                    "berarti realisasi melebihi PO (potensi over-"
                    "delivery/over-invoice); di bawah nol berarti "
                    "belum semua PO terealisasi bulan itu."
                ),
            ])

    with right:
        if not vendors.empty:
            figure = bar_chart(
                vendors.head(10),
                "grand_total_do",
                "vendor",
                "Top 10 Vendors",
                horizontal=True,
            )

            st.plotly_chart(
                figure,
                use_container_width=True,
            )

    if not vendors.empty:
        largest_vendor = vendors.iloc[0]

    else:
        largest_vendor = None

    if total_do:
        concentration = (
            vendors
            .head(5)["grand_total_do"]
            .sum()
            / total_do
        )

    else:
        concentration = 0

    delivery_status = df.get(
        "delivery_status",
        pd.Series(dtype="string"),
    )

    short_lines = int(
        (
            delivery_status
            == "SHORT DELIVERY"
        ).sum()
    )

    if largest_vendor is not None:
        largest_vendor_name = (
            largest_vendor["vendor"]
        )

        largest_vendor_value = rupiah(
            largest_vendor[
                "grand_total_do"
            ]
        )

    else:
        largest_vendor_name = "-"
        largest_vendor_value = "-"

    insight_box(
        "Management signal",
        (
            f"Top 5 vendors represent "
            f"{pct(concentration)} of visible spend. "
            f"Largest vendor: "
            f"{largest_vendor_name} "
            f"({largest_vendor_value}). "
            f"There are "
            f"{number_id(short_lines)} "
            f"short-delivery lines requiring review."
        ),
    )

    st.markdown(
        "#### Spend Trend & Forecast"
    )

    trend = time_intelligence(
        monthly
    )

    if not trend.empty:
        actual_trend = trend[
            ~trend["is_forecast"]
        ]

        last_growth = (
            actual_trend["mom_growth"]
            .dropna()
            .iloc[-1]
            if actual_trend["mom_growth"]
            .dropna()
            .shape[0]
            else np.nan
        )

        avg_growth = (
            actual_trend["mom_growth"]
            .dropna()
            .tail(3)
            .mean()
        )

        forecast_rows = trend[
            trend["is_forecast"]
        ]

        next_forecast = (
            forecast_rows.iloc[0][
                "grand_total_do"
            ]
            if not forecast_rows.empty
            else np.nan
        )

        t1, t2, t3 = st.columns(3)

        t1.metric(
            "Last Month Growth (MoM)",
            pct(last_growth),
        )

        t2.metric(
            "3-Month Avg Growth",
            pct(avg_growth),
        )

        t3.metric(
            "Next Month Forecast",
            rupiah(next_forecast),
        )

        figure = go.Figure()

        figure.add_trace(
            go.Bar(
                x=actual_trend["costing_month"],
                y=actual_trend["grand_total_do"],
                name="Actual Spend",
                marker_color=COLOR["accent"],
            )
        )

        figure.add_trace(
            go.Scatter(
                x=trend["costing_month"],
                y=trend["moving_avg"],
                name="Moving Average",
                mode="lines",
                line={
                    "color": COLOR["accent_2"],
                    "width": 2,
                },
            )
        )

        if not forecast_rows.empty:
            bridge_x = (
                [actual_trend["costing_month"].iloc[-1]]
                + forecast_rows["costing_month"].tolist()
            )

            bridge_y = (
                [actual_trend["grand_total_do"].iloc[-1]]
                + forecast_rows["grand_total_do"].tolist()
            )

            figure.add_trace(
                go.Scatter(
                    x=bridge_x,
                    y=bridge_y,
                    name="Forecast",
                    mode="lines+markers",
                    line={
                        "color": COLOR["warning"],
                        "width": 2,
                        "dash": "dash",
                    },
                )
            )

        figure.update_layout(
            title="Monthly Spend: Actual, Trend & Forecast",
            barmode="overlay",
        )

        st.plotly_chart(
            clean_layout(
                figure,
                height=420,
            ),
            use_container_width=True,
        )

        reading_guide([
            (
                "<b>Bar</b> = actual spend per bulan yang sudah "
                "terjadi (dari data ter-filter saat ini)."
            ),
            (
                "<b>Garis abu-abu</b> (moving average) meredam "
                "noise bulan-ke-bulan agar arah tren lebih "
                "jelas — naik terus artinya belanja sedang "
                "membesar secara struktural, bukan sekadar "
                "lonjakan satu bulan."
            ),
            (
                "<b>Garis putus-putus</b> adalah proyeksi 1-2 "
                "bulan ke depan berdasarkan tren linear historis. "
                "Gunakan sebagai early-warning untuk kebutuhan "
                "cash dan budget, bukan angka final."
            ),
            (
                "Jika <b>Last Month Growth</b> jauh di atas "
                "<b>3-Month Avg Growth</b>, cek apakah ada "
                "pembelian besar satu kali (one-off) yang perlu "
                "dikonfirmasi ke tim requester."
            ),
        ])

    else:
        st.info(
            "Monthly spend data is not sufficient for trend analysis."
        )

    c1, c2 = st.columns(2)

    with c1:
        st.markdown(
            "#### Spend by Category"
        )

        display_df(
            datasets["Category Spend"],
            currency_cols=[
                "grand_total_do",
            ],
        )

    with c2:
        st.markdown(
            "#### Spend by Site"
        )

        display_df(
            datasets["Site Spend"],
            currency_cols=[
                "grand_total_do",
            ],
        )

    export_button(
        label="Export Executive Pack",
        file_name="executive_overview.xlsx",
        sheets={
            **datasets,
            "Spend Trend Forecast": trend,
            "Filtered Raw Data": df,
        },
        title=(
            "Procurement Executive Overview"
        ),
        key="export_executive",
    )


# =========================================================
# SPEND MODULE
# =========================================================
@st.fragment
def show_spend(
    df: pd.DataFrame,
) -> None:
    section_heading(
        "Cost Control",
        "Spend & Budget Lens",
        (
            "Analyze where purchasing value is "
            "concentrated and how it changes over time."
        ),
    )

    df = render_local_filters(df, "spend")

    dimensions = [
        col
        for col in [
            "cost_category",
            "site",
            "vendor",
            "cost_center",
            "item_category",
            "item_description",
        ]
        if col in df.columns
    ]

    if not dimensions:
        st.warning(
            "No spend dimension is available."
        )
        return

    c1, c2 = st.columns([
        2,
        1,
    ])

    with c1:
        dimension = st.selectbox(
            "Analysis dimension",
            dimensions,
            format_func=lambda value: (
                DISPLAY_NAMES.get(
                    value,
                    value,
                )
            ),
            key="spend_dimension",
        )

    with c2:
        top_n = st.slider(
            "Top entities",
            min_value=5,
            max_value=50,
            value=15,
            key="spend_top_n",
        )

    grouped = safe_group_sum(
        df,
        [dimension],
    ).head(top_n)

    total_spend = df.get(
        "grand_total_do",
        pd.Series(dtype=float),
    ).sum()

    if total_spend:
        top_share = (
            grouped["grand_total_do"]
            .sum()
            / total_spend
        )

    else:
        top_share = 0

    po_count = df.get(
        "po_no",
        pd.Series(dtype="string"),
    ).nunique()

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Visible Spend",
        rupiah(total_spend),
    )

    c2.metric(
        f"Top {top_n} Share",
        pct(top_share),
    )

    c3.metric(
        "Average Spend / PO",
        rupiah(
            total_spend
            / max(
                po_count,
                1,
            )
        ),
    )

    if not grouped.empty:
        figure = bar_chart(
            grouped,
            "grand_total_do",
            dimension,
            f"Top {top_n} by Actual Spend",
            horizontal=True,
            height=480,
        )

        st.plotly_chart(
            figure,
            use_container_width=True,
        )

    display_df(
        grouped,
        currency_cols=[
            "grand_total_do",
        ],
    )

    pivot = pd.DataFrame()

    required = {
        dimension,
        "costing_month",
        "grand_total_do",
    }

    if required.issubset(df.columns):
        pivot = df.pivot_table(
            index=dimension,
            columns="costing_month",
            values="grand_total_do",
            aggfunc="sum",
            fill_value=0,
        )

        order = (
            pivot
            .sum(axis=1)
            .sort_values(
                ascending=False
            )
            .index[:top_n]
        )

        pivot = pivot.loc[order]

        month_columns = sorted(
            pivot.columns,
            key=lambda label: (
                month_to_no(label)
                if pd.notna(
                    month_to_no(label)
                )
                else 99
            ),
        )

        pivot = pivot[month_columns]

        st.markdown(
            "#### Monthly Spend Matrix"
        )

        heatmap_figure = go.Figure(
            data=go.Heatmap(
                z=pivot.values,
                x=pivot.columns,
                y=[
                    str(label)
                    for label in pivot.index
                ],
                colorscale="Blues",
                colorbar={
                    "title": "Spend",
                },
                hovertemplate=(
                    f"{DISPLAY_NAMES.get(dimension, dimension)}: "
                    "%{y}<br>Month: %{x}<br>"
                    "Spend: Rp %{z:,.0f}<extra></extra>"
                ),
            )
        )

        heatmap_figure.update_layout(
            title=(
                "Spend Heatmap: "
                f"{DISPLAY_NAMES.get(dimension, dimension)} "
                "x Month"
            ),
            yaxis={
                "autorange": "reversed",
            },
        )

        st.plotly_chart(
            clean_layout(
                heatmap_figure,
                height=max(
                    320,
                    28 * len(pivot.index),
                ),
            ),
            use_container_width=True,
        )

        reading_guide([
            (
                "Semakin gelap warnanya, semakin besar spend "
                f"{DISPLAY_NAMES.get(dimension, dimension).lower()} "
                "tersebut pada bulan itu — pindai kolom untuk "
                "melihat bulan mana yang paling berat."
            ),
            (
                "Baris dengan warna gelap merata di semua bulan "
                "menandakan biaya rutin/berulang; warna gelap "
                "hanya di satu bulan menandakan pembelian "
                "one-off yang perlu dikonfirmasi."
            ),
        ])

        pivot = pivot.reset_index()

        display_df(
            pivot,
            currency_cols=list(
                pivot.columns[1:]
            ),
        )

    export_button(
        label="Export Spend Analysis",
        file_name="spend_analysis.xlsx",
        sheets={
            "Spend Ranking": grouped,
            "Monthly Matrix": pivot,
            "Filtered Raw Data": df,
        },
        title=(
            "Spend and Cost Control Analysis"
        ),
        key="export_spend",
    )


# =========================================================
# VENDOR MODULE
# =========================================================
@st.fragment
def show_vendor(
    df: pd.DataFrame,
) -> None:
    section_heading(
        "Purchasing Management",
        "Vendor Performance",
        (
            "Review spend concentration, service "
            "performance, top items, and vendor dependency."
        ),
    )

    df = render_local_filters(df, "vendor")

    scorecard = vendor_scorecard(
        df
    )

    if scorecard.empty:
        st.warning(
            "Vendor data is not available."
        )
        return

    total_spend = (
        scorecard["spend_do"]
        .sum()
    )

    top_vendor_share = (
        scorecard.iloc[0][
            "spend_share"
        ]
    )

    top_five_share = (
        scorecard
        .head(5)["spend_share"]
        .sum()
    )

    c1, c2, c3, c4 = (
        st.columns(4)
    )

    c1.metric(
        "Vendor Count",
        number_id(
            len(scorecard)
        ),
    )

    c2.metric(
        "Top Vendor Share",
        pct(top_vendor_share),
    )

    c3.metric(
        "Top 5 Concentration",
        pct(top_five_share),
    )

    c4.metric(
        "Average Spend / Vendor",
        rupiah(
            total_spend
            / max(
                len(scorecard),
                1,
            )
        ),
    )

    left, right = st.columns([
        1.15,
        1,
    ])

    with left:
        figure = bar_chart(
            scorecard.head(15),
            "spend_do",
            "vendor",
            "Vendor Spend Ranking",
            horizontal=True,
            height=500,
        )

        st.plotly_chart(
            figure,
            use_container_width=True,
        )

    with right:
        if (
            "avg_delivery_days"
            in scorecard.columns
        ):
            scatter_x = (
                "avg_delivery_days"
            )

        else:
            scatter_x = "po_count"

        if (
            "avg_fulfillment_rate"
            in scorecard.columns
        ):
            scatter_y = (
                "avg_fulfillment_rate"
            )

        else:
            scatter_y = "spend_share"

        figure = px.scatter(
            scorecard.head(50),
            x=scatter_x,
            y=scatter_y,
            size="spend_do",
            hover_name="vendor",
            title=(
                "Vendor Risk / Performance Map"
            ),
            color_discrete_sequence=[
                COLOR["accent"],
            ],
        )

        figure = clean_layout(
            figure,
            height=500,
        )

        st.plotly_chart(
            figure,
            use_container_width=True,
        )

    display_df(
        scorecard,
        currency_cols=[
            "spend_do",
        ],
        percent_cols=[
            "spend_share",
            "avg_fulfillment_rate",
        ],
    )

    st.markdown(
        "#### Vendor Risk Segmentation"
    )

    segmentation = vendor_segmentation(
        scorecard
    )

    if not segmentation.empty:
        quadrant_colors = RISK_QUADRANT_COLORS

        quadrant_counts = (
            segmentation["risk_quadrant"]
            .value_counts()
        )

        s1, s2, s3, s4 = st.columns(4)

        s1.metric(
            "Critical Vendors",
            number_id(
                quadrant_counts.get(
                    "CRITICAL - manage closely",
                    0,
                )
            ),
        )

        s2.metric(
            "Strategic Vendors",
            number_id(
                quadrant_counts.get(
                    "STRATEGIC - maintain",
                    0,
                )
            ),
        )

        s3.metric(
            "Monitor Vendors",
            number_id(
                quadrant_counts.get(
                    "MONITOR - watch closely",
                    0,
                )
            ),
        )

        s4.metric(
            "Routine Vendors",
            number_id(
                quadrant_counts.get(
                    "ROUTINE - low touch",
                    0,
                )
            ),
        )

        figure = px.scatter(
            segmentation,
            x="spend_do",
            y="risk_score",
            size="spend_do",
            color="risk_quadrant",
            hover_name="vendor",
            log_x=True,
            title="Vendor Segmentation: Spend vs Risk Score",
            color_discrete_map=quadrant_colors,
        )

        figure.add_hline(
            y=segmentation["risk_score"].median(),
            line_dash="dot",
            line_color=COLOR["muted"],
        )

        figure.add_vline(
            x=segmentation["spend_do"].median(),
            line_dash="dot",
            line_color=COLOR["muted"],
        )

        st.plotly_chart(
            clean_layout(
                figure,
                height=460,
            ),
            use_container_width=True,
        )

        reading_guide([
            (
                "Sumbu X = total spend vendor (skala log), "
                "sumbu Y = <b>risk score</b> gabungan dari "
                "fulfillment rendah, keterlambatan pengiriman, "
                "hari overdue pembayaran, dan konsentrasi spend."
            ),
            (
                "<b>CRITICAL</b> (kanan-atas): spend besar & "
                "risiko tinggi — prioritas utama untuk review "
                "kontrak, business review rutin, dan mitigasi "
                "supply risk."
            ),
            (
                "<b>STRATEGIC</b> (kanan-bawah): spend besar "
                "tapi risiko rendah — jaga relasi baik, ini "
                "vendor andalan."
            ),
            (
                "<b>MONITOR</b> (kiri-atas): spend kecil tapi "
                "risiko tinggi — pantau agar tidak membesar "
                "jadi masalah, atau pertimbangkan mengurangi "
                "ketergantungan."
            ),
            (
                "<b>ROUTINE</b> (kiri-bawah): spend kecil & "
                "risiko rendah — cukup dipantau berkala, tidak "
                "perlu effort khusus."
            ),
        ])

        display_df(
            segmentation[
                [
                    col
                    for col in [
                        "vendor",
                        "spend_do",
                        "spend_share",
                        "avg_fulfillment_rate",
                        "avg_delivery_days",
                        "avg_payment_overdue_days",
                        "risk_score",
                        "risk_quadrant",
                    ]
                    if col in segmentation.columns
                ]
            ],
            currency_cols=[
                "spend_do",
            ],
            percent_cols=[
                "spend_share",
                "avg_fulfillment_rate",
            ],
            highlight_col="risk_quadrant",
            highlight_map=RISK_QUADRANT_COLORS,
        )

    else:
        st.info(
            "Not enough vendor metrics to compute risk segmentation."
        )
        segmentation = pd.DataFrame()

    st.markdown(
        "#### Top 50 Items Purchased per Vendor"
    )

    top_items = vendor_top_items(
        df,
        top_n=50,
    )

    vendor_options = (
        scorecard["vendor"]
        .dropna()
        .astype(str)
        .tolist()
    )

    selected_vendor = st.selectbox(
        "Choose vendor",
        vendor_options,
        key="top_item_vendor",
    )

    if selected_vendor:
        vendor_items = top_items[
            top_items["vendor"]
            == selected_vendor
        ].copy()

    else:
        vendor_items = pd.DataFrame()

    display_df(
        vendor_items,
        currency_cols=[
            "spend_do",
            "avg_price",
            "min_price",
            "max_price",
        ],
    )

    matrix = pd.DataFrame()

    required = {
        "vendor",
        "cost_category",
        "grand_total_do",
    }

    if required.issubset(df.columns):
        matrix = df.pivot_table(
            index="vendor",
            columns="cost_category",
            values="grand_total_do",
            aggfunc="sum",
            fill_value=0,
        )

        matrix["TOTAL"] = (
            matrix.sum(axis=1)
        )

        matrix = (
            matrix
            .sort_values(
                "TOTAL",
                ascending=False,
            )
            .reset_index()
        )

        with st.expander(
            "Vendor × Cost Category matrix"
        ):
            display_df(
                matrix,
                currency_cols=list(
                    matrix.columns[1:]
                ),
            )

    export_button(
        label="Export Vendor Scorecard",
        file_name="vendor_performance.xlsx",
        sheets={
            "Vendor Scorecard": scorecard,
            "Vendor Risk Segmentation": segmentation,
            "Top 50 Items All Vendors": top_items,
            "Selected Vendor Items": vendor_items,
            "Vendor Category Matrix": matrix,
            "Filtered Raw Data": df,
        },
        title=(
            "Vendor Performance Scorecard"
        ),
        key="export_vendor",
    )


# =========================================================
# ABC / PARETO SEGMENTATION MODULE
# =========================================================
@st.fragment
def show_segmentation(
    df: pd.DataFrame,
) -> None:
    """
    Modul klasifikasi ABC (Pareto) untuk item dan vendor.
    """

    section_heading(
        "Prioritization",
        "ABC / Pareto Classification",
        (
            "Fokuskan effort negosiasi, audit, dan kontrol "
            "pada item atau vendor yang paling menentukan "
            "total spend."
        ),
    )

    df = render_local_filters(df, "abc")

    dimension_options = [
        col
        for col in [
            "item_description",
            "vendor",
            "cost_category",
            "site",
        ]
        if col in df.columns
    ]

    if not dimension_options:
        st.warning(
            "No dimension is available for ABC classification."
        )
        return

    dimension = st.selectbox(
        "Classify by",
        dimension_options,
        format_func=lambda value: (
            DISPLAY_NAMES.get(
                value,
                value,
            )
        ),
        key="abc_dimension",
    )

    classified = abc_classification(
        df,
        dimension,
    )

    if classified.empty:
        st.info(
            "Not enough data to build ABC classification "
            "for this dimension."
        )
        return

    class_summary = (
        classified
        .groupby("abc_class")
        .agg(
            entity_count=(
                dimension,
                "count",
            ),
            spend_do=(
                "grand_total_do",
                "sum",
            ),
        )
        .reindex(["A", "B", "C"])
        .fillna(0)
        .reset_index()
    )

    total_entities = len(classified)
    total_spend = classified["grand_total_do"].sum()

    class_summary["entity_share"] = np.where(
        total_entities > 0,
        class_summary["entity_count"] / total_entities,
        0,
    )

    class_summary["spend_share"] = np.where(
        total_spend > 0,
        class_summary["spend_do"] / total_spend,
        0,
    )

    a_row = class_summary[
        class_summary["abc_class"] == "A"
    ]

    a_entity_share = (
        a_row["entity_share"].iloc[0]
        if not a_row.empty
        else 0
    )

    a_spend_share = (
        a_row["spend_share"].iloc[0]
        if not a_row.empty
        else 0
    )

    m1, m2, m3 = st.columns(3)

    m1.metric(
        "Total Entities",
        number_id(total_entities),
    )

    m2.metric(
        "Class A Count",
        number_id(
            int(
                a_row["entity_count"].iloc[0]
                if not a_row.empty
                else 0
            )
        ),
        delta=pct(a_entity_share),
    )

    m3.metric(
        "Class A Spend Share",
        pct(a_spend_share),
    )

    figure = go.Figure()

    bar_colors = [
        ABC_CLASS_COLORS.get(
            cls,
            COLOR["accent_2"],
        )
        for cls in classified["abc_class"]
    ]

    figure.add_trace(
        go.Bar(
            x=classified["rank"],
            y=classified["grand_total_do"],
            name="Spend",
            marker_color=bar_colors,
        )
    )

    figure.add_trace(
        go.Scatter(
            x=classified["rank"],
            y=classified["cumulative_share"] * 100,
            name="Cumulative %",
            yaxis="y2",
            mode="lines",
            line={
                "color": COLOR["accent"],
                "width": 2,
            },
        )
    )

    figure.update_layout(
        title=(
            "Pareto Chart: "
            f"{DISPLAY_NAMES.get(dimension, dimension)}"
        ),
        yaxis={
            "title": "Spend",
        },
        yaxis2={
            "title": "Cumulative %",
            "overlaying": "y",
            "side": "right",
            "range": [0, 105],
        },
        xaxis={
            "title": "Rank",
        },
    )

    st.plotly_chart(
        clean_layout(
            figure,
            height=440,
        ),
        use_container_width=True,
    )

    reading_guide([
        (
            "<b>Class A</b> (merah) = kontributor sekitar 80% "
            "spend pertama secara kumulatif — jumlahnya sedikit "
            "tapi dampaknya paling besar. Ini prioritas utama "
            "untuk negosiasi harga, kontrak jangka panjang, "
            "dan audit rutin."
        ),
        (
            "<b>Class B</b> (kuning) = kontributor 80%-95% "
            "kumulatif — layak dipantau, tapi tidak perlu "
            "seketat Class A."
        ),
        (
            "<b>Class C</b> (abu-abu) = sisa 5% terakhir, "
            "biasanya jumlahnya banyak tapi nilainya kecil per "
            "item/vendor. Cocok untuk simplifikasi proses "
            "(misalnya PO blanket/consolidated) daripada "
            "dinegosiasikan satu per satu."
        ),
        (
            "Kalau Class A count kecil tapi spend share-nya "
            "sangat besar, itu tanda konsentrasi tinggi — good "
            "untuk leverage negosiasi, tapi juga sinyal risiko "
            "ketergantungan yang perlu mitigasi."
        ),
    ])

    display_df(
        classified,
        currency_cols=[
            "grand_total_do",
            "cumulative_value",
        ],
        percent_cols=[
            "cumulative_share",
            "spend_share",
        ],
        highlight_col="abc_class",
        highlight_map=ABC_CLASS_COLORS,
    )

    st.markdown(
        "#### Class Summary"
    )

    display_df(
        class_summary,
        currency_cols=[
            "spend_do",
        ],
        percent_cols=[
            "entity_share",
            "spend_share",
        ],
        highlight_col="abc_class",
        highlight_map=ABC_CLASS_COLORS,
    )

    export_button(
        label="Export ABC Classification",
        file_name="abc_classification.xlsx",
        sheets={
            "ABC Detail": classified,
            "Class Summary": class_summary,
            "Filtered Raw Data": df,
        },
        title=(
            "ABC / Pareto Spend Classification"
        ),
        key="export_abc",
    )


# =========================================================
# PRICE MODULE
# =========================================================
@st.fragment
def show_price(
    df: pd.DataFrame,
) -> None:
    section_heading(
        "Savings",
        "Price Intelligence & Vendor Comparison",
        (
            "Identify inconsistent prices and compare "
            "the same items across up to three vendors."
        ),
    )

    df = render_local_filters(df, "price")

    c1, c2 = st.columns(2)

    with c1:
        min_spend = st.number_input(
            "Minimum item spend",
            min_value=0,
            value=10_000_000,
            step=5_000_000,
            key="min_spend",
        )

    with c2:
        min_transactions = st.number_input(
            "Minimum transaction rows",
            min_value=1,
            value=3,
            step=1,
            key="min_txn",
        )

    opportunity = price_opportunity(
        df,
        min_spend,
        min_transactions,
    )

    if not opportunity.empty:
        total_saving = (
            opportunity[
                "potential_saving"
            ].sum()
        )

        spend_scope = (
            opportunity[
                "spend_do"
            ].sum()
        )

    else:
        total_saving = 0
        spend_scope = 0

    if spend_scope:
        saving_rate = (
            total_saving
            / spend_scope
        )

    else:
        saving_rate = 0

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Potential Saving",
        rupiah(total_saving),
    )

    c2.metric(
        "Candidate Items",
        number_id(
            len(opportunity)
        ),
    )

    c3.metric(
        "Opportunity Rate",
        pct(saving_rate),
    )

    if not opportunity.empty:
        figure = px.scatter(
            opportunity.head(100),
            x="spend_do",
            y="price_gap_rate",
            size="potential_saving",
            color="price_gap_rate",
            hover_name="item_description",
            title=(
                "Saving Priority: "
                "Spend vs Price Dispersion"
            ),
            color_continuous_scale=[
                COLOR["success"],
                COLOR["warning"],
                COLOR["danger"],
            ],
        )

        figure.update_yaxes(
            tickformat=".0%"
        )

        figure.update_layout(
            coloraxis_colorbar={
                "title": "Price Gap",
                "tickformat": ".0%",
            },
        )

        figure = clean_layout(
            figure,
            height=440,
        )

        st.plotly_chart(
            figure,
            use_container_width=True,
        )

        reading_guide([
            (
                "Titik <b>hijau</b> = harga antar transaksi "
                "relatif konsisten. Titik <b>merah</b> = gap "
                "harga besar antar pembelian item yang sama — "
                "kandidat kuat untuk negosiasi atau standardisasi "
                "vendor."
            ),
            (
                "Ukuran bubble = besarnya potential saving. "
                "Prioritaskan titik merah yang juga besar "
                "(kanan-atas) karena dampak finansialnya paling "
                "signifikan."
            ),
        ])

    display_df(
        opportunity,
        currency_cols=[
            "spend_do",
            "min_price",
            "avg_price",
            "median_price",
            "max_price",
            "price_gap",
            "potential_saving",
        ],
        percent_cols=[
            "price_gap_rate",
        ],
    )

    st.markdown(
        "#### Compare Prices Across Vendors"
    )

    if "vendor" in df.columns:
        vendor_options = sorted(
            df["vendor"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    else:
        vendor_options = []

    default_vendors = vendor_options[
        :min(
            3,
            len(vendor_options),
        )
    ]

    selected_vendors = st.multiselect(
        "Select up to 3 vendors",
        vendor_options,
        default=default_vendors,
        max_selections=3,
        key="compare_vendors",
    )

    comparison = vendor_price_comparison(
        df,
        selected_vendors,
    )

    shown_comparison = pd.DataFrame()

    if len(selected_vendors) < 2:
        st.info(
            "Select at least two vendors to compare prices."
        )

    elif comparison.empty:
        st.info(
            "No comparable item-price data was found "
            "for the selected vendors."
        )

    else:
        common_only = st.toggle(
            (
                "Show only items available from at least "
                "2 selected vendors"
            ),
            value=True,
            key="common_items_only",
        )

        if common_only:
            shown_comparison = comparison[
                comparison["vendor_coverage"]
                >= 2
            ].copy()

        else:
            shown_comparison = (
                comparison.copy()
            )

        display_df(
            shown_comparison,
            currency_cols=(
                selected_vendors
                + [
                    "lowest_price",
                    "highest_price",
                    "price_gap",
                ]
            ),
            percent_cols=[
                "price_gap_rate",
            ],
        )

        chart_data = (
            shown_comparison
            .head(30)
            .melt(
                id_vars=(
                    "item_description"
                ),
                value_vars=(
                    selected_vendors
                ),
                var_name="vendor",
                value_name="avg_price",
            )
            .dropna(
                subset=["avg_price"]
            )
        )

        if not chart_data.empty:
            figure = px.bar(
                chart_data,
                x="item_description",
                y="avg_price",
                color="vendor",
                barmode="group",
                title=(
                    "Average Item Price "
                    "by Selected Vendor"
                ),
            )

            figure.update_layout(
                xaxis_tickangle=-45
            )

            figure = clean_layout(
                figure,
                height=520,
            )

            st.plotly_chart(
                figure,
                use_container_width=True,
            )

    detail = pd.DataFrame()

    if not opportunity.empty:
        item_options = (
            opportunity[
                "item_description"
            ]
            .head(150)
            .tolist()
        )

    else:
        item_options = []

    selected_item = st.selectbox(
        "Inspect vendor prices for item",
        [None] + item_options,
        key="price_item",
    )

    if selected_item:
        detail_columns = [
            col
            for col in [
                "costing_month",
                "site",
                "vendor",
                "po_no",
                "arrival_date",
                "unit",
                "qty_do",
                "price_unit",
                "grand_total_do",
            ]
            if col in df.columns
        ]

        detail = (
            df[
                df["item_description"]
                == selected_item
            ][detail_columns]
            .sort_values(
                "price_unit"
            )
        )

        display_df(
            detail,
            currency_cols=[
                "price_unit",
                "grand_total_do",
            ],
        )

    export_button(
        label="Export Price Analysis",
        file_name="price_intelligence.xlsx",
        sheets={
            "Saving Opportunities": opportunity,
            "Vendor Price Comparison": comparison,
            "Visible Comparison": shown_comparison,
            "Selected Item Detail": detail,
            "Filtered Raw Data": df,
        },
        title=(
            "Price Intelligence "
            "and Vendor Comparison"
        ),
        key="export_price",
    )


# =========================================================
# PO VS DO MODULE
# =========================================================
@st.fragment
def show_po_do(
    df: pd.DataFrame,
) -> None:
    section_heading(
        "Operational Control",
        "PO vs DO Reconciliation",
        (
            "Monitor quantity fulfillment and value "
            "variances between orders and deliveries."
        ),
    )

    df = render_local_filters(df, "podo")

    if "delivery_status" not in df.columns:
        st.warning(
            "PO and DO quantity columns are incomplete."
        )
        return

    aggregation = {
        "line_count": (
            "delivery_status",
            "size",
        ),
    }

    if "grand_total_do" in df.columns:
        aggregation["value_do"] = (
            "grand_total_do",
            "sum",
        )

    if (
        "value_variance_do_vs_po"
        in df.columns
    ):
        aggregation["value_variance"] = (
            "value_variance_do_vs_po",
            "sum",
        )

    status = (
        df
        .groupby("delivery_status")
        .agg(**aggregation)
        .reset_index()
        .sort_values(
            "line_count",
            ascending=False,
        )
    )

    if {
        "qty_po",
        "qty_do",
    }.issubset(df.columns):
        total_qty_po = (
            df["qty_po"]
            .sum()
        )

        total_qty_do = (
            df["qty_do"]
            .sum()
        )

        if total_qty_po:
            weighted_fulfillment = (
                total_qty_do
                / total_qty_po
            )

        else:
            weighted_fulfillment = np.nan

    else:
        weighted_fulfillment = np.nan

    short_count = int(
        (
            df["delivery_status"]
            == "SHORT DELIVERY"
        ).sum()
    )

    over_count = int(
        (
            df["delivery_status"]
            == "OVER DELIVERY"
        ).sum()
    )

    absolute_variance = (
        df.get(
            "value_variance_do_vs_po",
            pd.Series(dtype=float),
        )
        .abs()
        .sum()
    )

    c1, c2, c3, c4 = (
        st.columns(4)
    )

    c1.metric(
        "Weighted Fulfillment",
        pct(weighted_fulfillment),
    )

    c2.metric(
        "Short Delivery Lines",
        number_id(short_count),
    )

    c3.metric(
        "Over Delivery Lines",
        number_id(over_count),
    )

    c4.metric(
        "Absolute Value Variance",
        rupiah(absolute_variance),
    )

    fulfillment_level_key, fulfillment_label = (
        fulfillment_level(
            weighted_fulfillment
        )
    )

    status_pill_markdown(
        "Fulfillment status",
        fulfillment_label,
        fulfillment_level_key,
    )

    figure = px.bar(
        status,
        x="delivery_status",
        y="line_count",
        title="Delivery Status Distribution",
        color="delivery_status",
        color_discrete_map=DELIVERY_STATUS_COLORS,
    )

    figure.update_layout(
        showlegend=False,
    )

    figure = clean_layout(
        figure
    )

    st.plotly_chart(
        figure,
        use_container_width=True,
    )

    reading_guide([
        (
            "<b>SHORT DELIVERY</b> (merah) = barang datang "
            "lebih sedikit dari PO — cek ke vendor apakah "
            "masih ada sisa kiriman (partial) atau perlu "
            "revisi PO."
        ),
        (
            "<b>OVER DELIVERY</b> (kuning) = barang datang "
            "lebih banyak dari PO — perlu konfirmasi apakah "
            "kelebihan diterima, dikembalikan, atau memang "
            "PO-nya perlu direvisi."
        ),
        (
            "<b>FULFILLED</b> (hijau) = qty DO sama dengan "
            "qty PO, tidak perlu tindakan lanjutan."
        ),
    ])

    display_df(
        status,
        currency_cols=[
            "value_do",
            "value_variance",
        ],
        highlight_col="delivery_status",
        highlight_map=DELIVERY_STATUS_COLORS,
    )

    detail_columns = [
        col
        for col in [
            "site",
            "vendor",
            "po_no",
            "item_description",
            "unit",
            "qty_po",
            "qty_do",
            "qty_variance",
            "grand_total_po",
            "grand_total_do",
            "value_variance_do_vs_po",
            "delivery_status",
        ]
        if col in df.columns
    ]

    detail = df[
        detail_columns
    ].copy()

    if (
        "value_variance_do_vs_po"
        in detail.columns
    ):
        detail["absolute_variance"] = (
            detail[
                "value_variance_do_vs_po"
            ].abs()
        )

        detail = (
            detail
            .sort_values(
                "absolute_variance",
                ascending=False,
            )
            .drop(
                columns=[
                    "absolute_variance",
                ]
            )
        )

    st.markdown(
        "#### Largest Reconciliation Exceptions"
    )

    display_df(
        detail.head(200),
        currency_cols=[
            "grand_total_po",
            "grand_total_do",
            "value_variance_do_vs_po",
        ],
        highlight_col="delivery_status",
        highlight_map=DELIVERY_STATUS_COLORS,
    )

    export_button(
        label="Export PO-DO Control",
        file_name="po_do_control.xlsx",
        sheets={
            "Status Summary": status,
            "Variance Detail": detail,
            "Filtered Raw Data": df,
        },
        title=(
            "PO versus DO Reconciliation"
        ),
        key="export_podo",
    )


# =========================================================
# LEAD TIME MODULE
# =========================================================
@st.fragment
def show_lead_time(
    df: pd.DataFrame,
) -> None:
    section_heading(
        "Process Efficiency",
        "Lead Time Analysis",
        (
            "Measure cycle time from request initiation "
            "through payment completion."
        ),
    )

    df = render_local_filters(df, "leadtime")

    lead_time_specs = [
        (
            "lt_pr_to_po",
            "PR to PO",
        ),
        (
            "lt_po_to_arrival",
            "PO to Arrival",
        ),
        (
            "lt_arrival_to_invoice",
            "Arrival to Invoice",
        ),
        (
            "lt_invoice_to_schedule",
            "Invoice to Schedule",
        ),
        (
            "lt_schedule_to_payout",
            "Schedule to Payout",
        ),
    ]

    available = [
        (
            column,
            label,
        )
        for column, label
        in lead_time_specs
        if column in df.columns
    ]

    if not available:
        st.warning(
            "Date columns are insufficient "
            "for lead-time analysis."
        )
        return

    summary_rows = []

    for column, label in available:
        values = (
            df[column]
            .dropna()
        )

        values = values[
            (values >= -365)
            & (values <= 365)
        ]

        summary_rows.append({
            "stage": label,
            "average_days": values.mean(),
            "median_days": values.median(),
            "p90_days": values.quantile(0.90),
            "observations": values.count(),
        })

    summary = pd.DataFrame(
        summary_rows
    )

    valid_summary = summary.dropna(
        subset=["average_days"]
    )

    if not valid_summary.empty:
        slowest = (
            valid_summary
            .sort_values(
                "average_days",
                ascending=False,
            )
            .iloc[0]
        )

        slowest_name = (
            slowest["stage"]
        )

        slowest_average = (
            slowest["average_days"]
        )

        slowest_p90 = (
            slowest["p90_days"]
        )

    else:
        slowest_name = "-"
        slowest_average = np.nan
        slowest_p90 = np.nan

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Slowest Stage",
        slowest_name,
    )

    c2.metric(
        "Average Days",
        number_id(
            slowest_average,
            decimals=1,
        ),
    )

    c3.metric(
        "P90 Days",
        number_id(
            slowest_p90,
            decimals=1,
        ),
    )

    stage_bar_colors = [
        COLOR["danger"]
        if stage == slowest_name
        else COLOR["accent_2"]
        for stage in summary["stage"]
    ]

    figure = px.bar(
        summary,
        x="stage",
        y="average_days",
        title="Average Cycle Time by Stage",
    )

    figure.update_traces(
        marker_color=stage_bar_colors,
    )

    figure = clean_layout(
        figure
    )

    st.plotly_chart(
        figure,
        use_container_width=True,
    )

    reading_guide([
        (
            f"Stage <b>{slowest_name}</b> (batang merah) "
            "adalah bottleneck utama dalam siklus procurement "
            "saat ini — paling banyak menyerap waktu rata-rata."
        ),
        (
            "Bandingkan <b>Average</b> vs <b>P90</b>: kalau "
            "gap-nya besar, artinya sebagian kecil transaksi "
            "sangat lambat (outlier) dan menyeret rata-rata — "
            "cek detail per vendor di bawah untuk temukan "
            "penyebabnya."
        ),
    ])

    display_df(
        summary
    )

    selected_label = st.selectbox(
        "Detailed lead-time stage",
        [
            label
            for _, label in available
        ],
        key="lead_stage",
    )

    selected_column = {
        label: column
        for column, label in available
    }[selected_label]

    detail_columns = [
        col
        for col in [
            "vendor",
            "site",
            "po_no",
            selected_column,
        ]
        if col in df.columns
    ]

    detail = (
        df[detail_columns]
        .dropna(
            subset=[selected_column]
        )
        .sort_values(
            selected_column,
            ascending=False,
        )
    )

    if "vendor" in detail.columns:
        vendor_lead_time = (
            detail
            .groupby("vendor")[
                selected_column
            ]
            .agg([
                "mean",
                "median",
                "count",
            ])
            .reset_index()
            .sort_values(
                "mean",
                ascending=False,
            )
        )

        st.markdown(
            "#### Vendor Lead-Time Ranking"
        )

        display_df(
            vendor_lead_time.head(100)
        )

    else:
        vendor_lead_time = pd.DataFrame()

    export_button(
        label="Export Lead Time",
        file_name="lead_time_analysis.xlsx",
        sheets={
            "Stage Summary": summary,
            "Vendor Ranking": vendor_lead_time,
            "Lead Time Detail": detail,
        },
        title=(
            "Procurement Lead Time Analysis"
        ),
        key="export_lead",
    )


# =========================================================
# PAYMENT MODULE
# =========================================================
@st.fragment
def show_payment(
    df: pd.DataFrame,
) -> None:
    """
    Modul payment yang dipanggil oleh tabs[6].

    Nama fungsi ini harus tetap:
    show_payment
    """

    section_heading(
        "Cash Exposure",
        "Invoice & Payment Monitoring",
        (
            "Track outstanding exposure, payment timing, "
            "and invoice completeness."
        ),
    )

    df = render_local_filters(df, "payment")

    if "payment_status" in df.columns:
        payment_status = (
            df["payment_status"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    else:
        payment_status = pd.Series(
            "",
            index=df.index,
            dtype="string",
        )

    if "invoice_status" in df.columns:
        invoice_status = (
            df["invoice_status"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    else:
        invoice_status = pd.Series(
            "",
            index=df.index,
            dtype="string",
        )

    if "grand_total_do" in df.columns:
        payment_value = pd.to_numeric(
            df["grand_total_do"],
            errors="coerce",
        ).fillna(0)

    else:
        payment_value = pd.Series(
            0,
            index=df.index,
            dtype=float,
        )

    if (
        "payment_overdue_days"
        in df.columns
    ):
        overdue_days = pd.to_numeric(
            df["payment_overdue_days"],
            errors="coerce",
        ).fillna(0)

    else:
        overdue_days = pd.Series(
            0,
            index=df.index,
            dtype=float,
        )

    open_value = payment_value[
        payment_status.eq("OPEN")
    ].sum()

    done_value = payment_value[
        payment_status.isin([
            "DONE",
            "PAID",
            "CLOSED",
        ])
    ].sum()

    uninvoiced_value = payment_value[
        invoice_status.isin([
            "UNINVOICE",
            "UNINVOICED",
            "NO INVOICE",
        ])
    ].sum()

    overdue_value = payment_value[
        overdue_days > 0
    ].sum()

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Open Payment",
        rupiah(open_value),
    )

    c2.metric(
        "Paid / Done",
        rupiah(done_value),
    )

    c3.metric(
        "Uninvoiced Exposure",
        rupiah(uninvoiced_value),
    )

    c4.metric(
        "Overdue Exposure",
        rupiah(overdue_value),
    )

    overdue_share_of_open = (
        overdue_value / open_value
        if open_value
        else np.nan
    )

    overdue_level_key, overdue_label = (
        overdue_exposure_level(
            overdue_share_of_open
        )
    )

    status_pill_markdown(
        "Overdue exposure status",
        overdue_label,
        overdue_level_key,
    )

    forecast = safe_group_sum(
        df,
        ["payment_schedule_month"],
    )

    if {
        "vendor",
        "payment_status",
        "grand_total_do",
    }.issubset(df.columns):

        open_payment_data = df[
            payment_status.eq("OPEN")
        ].copy()

        outstanding = safe_group_sum(
            open_payment_data,
            ["vendor"],
        )

    else:
        outstanding = pd.DataFrame()

    left, right = st.columns([
        1.15,
        1,
    ])

    with left:
        if not forecast.empty:
            figure = bar_chart(
                forecast,
                "payment_schedule_month",
                "grand_total_do",
                "Forecast Cash Out",
            )

            st.plotly_chart(
                figure,
                use_container_width=True,
            )

        else:
            st.info(
                "Payment schedule data is not available."
            )

    with right:
        if not outstanding.empty:
            figure = bar_chart(
                outstanding.head(10),
                "grand_total_do",
                "vendor",
                "Outstanding by Vendor",
                horizontal=True,
            )

            st.plotly_chart(
                figure,
                use_container_width=True,
            )

        else:
            st.info(
                "Outstanding vendor data is not available."
            )

    st.markdown(
        "#### Payment Forecast"
    )

    display_df(
        forecast,
        currency_cols=[
            "grand_total_do",
        ],
    )

    st.markdown(
        "#### Outstanding Vendor Exposure"
    )

    display_df(
        outstanding,
        currency_cols=[
            "grand_total_do",
        ],
    )

    st.markdown(
        "#### AP Aging (Outstanding Exposure by Age)"
    )

    aging = payment_aging(
        df
    )

    if not aging.empty:
        total_exposure = (
            aging["exposure_value"].sum()
        )

        overdue_only = aging[
            aging["bucket"] != "Not yet due"
        ]

        overdue_exposure = (
            overdue_only["exposure_value"].sum()
        )

        critical_exposure = aging.loc[
            aging["bucket"].isin(
                [
                    "61-90 days",
                    "90+ days",
                ]
            ),
            "exposure_value",
        ].sum()

        overdue_share = (
            overdue_exposure / total_exposure
            if total_exposure
            else 0
        )

        a1, a2, a3 = st.columns(3)

        a1.metric(
            "Total Open Exposure",
            rupiah(total_exposure),
        )

        a2.metric(
            "Overdue Exposure",
            rupiah(overdue_exposure),
            delta=pct(overdue_share),
            delta_color="inverse",
        )

        a3.metric(
            "Critical (61+ days)",
            rupiah(critical_exposure),
        )

        bucket_colors = AGING_BUCKET_COLORS

        figure = px.bar(
            aging,
            x="bucket",
            y="exposure_value",
            title="Open Payment Exposure by Age Bucket",
            color="bucket",
            color_discrete_map=bucket_colors,
        )

        figure.update_traces(
            hovertemplate=(
                "%{x}<br>Rp %{y:,.0f}<extra></extra>"
            )
        )

        figure.update_layout(
            showlegend=False,
        )

        st.plotly_chart(
            clean_layout(
                figure
            ),
            use_container_width=True,
        )

        reading_guide([
            (
                "Bucket ini hanya menghitung transaksi dengan "
                "<b>Payment Status = OPEN</b> — exposure yang "
                "belum benar-benar keluar cash-nya."
            ),
            (
                "<b>61-90 hari</b> dan <b>90+ hari</b> adalah "
                "zona kritis: segera cek dengan tim AP kenapa "
                "belum dibayar (dispute invoice? approval macet? "
                "cash tidak cukup?) karena berpotensi merusak "
                "relasi vendor atau kena penalti keterlambatan."
            ),
            (
                "Kalau <b>Overdue Exposure</b> naik terus tiap "
                "periode dibanding total spend, itu sinyal ada "
                "masalah proses approval/cash planning, bukan "
                "sekadar kasus vendor per vendor."
            ),
        ])

        display_df(
            aging,
            currency_cols=[
                "exposure_value",
            ],
            percent_cols=[
                "exposure_share",
            ],
            highlight_col="bucket",
            highlight_map=AGING_BUCKET_COLORS,
        )

    else:
        st.info(
            "No open payment exposure found for aging analysis."
        )
        aging = pd.DataFrame()

    detail_columns = [
        col
        for col in [
            "site",
            "vendor",
            "po_no",
            "invoice_status",
            "invoice_received_date",
            "payment_status",
            "payment_schedule",
            "payment_schedule_month",
            "actual_payout_date",
            "payment_overdue_days",
            "grand_total_do",
        ]
        if col in df.columns
    ]

    if detail_columns:
        payment_detail = (
            df[detail_columns]
            .copy()
        )

    else:
        payment_detail = pd.DataFrame()

    if (
        not payment_detail.empty
        and "payment_overdue_days"
        in payment_detail.columns
    ):
        payment_detail = (
            payment_detail
            .sort_values(
                "payment_overdue_days",
                ascending=False,
            )
        )

    st.markdown(
        "#### Payment Detail"
    )

    display_df(
        payment_detail,
        currency_cols=[
            "grand_total_do",
        ],
    )

    export_button(
        label="Export Payment Monitoring",
        file_name="payment_monitoring.xlsx",
        sheets={
            "Payment Forecast": forecast,
            "Outstanding Vendors": outstanding,
            "AP Aging": aging,
            "Payment Detail": payment_detail,
        },
        title=(
            "Invoice and Payment Monitoring"
        ),
        key="export_payment",
    )


# =========================================================
# DATA QUALITY MODULE
# =========================================================
@st.fragment
def show_quality(
    df: pd.DataFrame,
) -> None:
    section_heading(
        "Data Governance",
        "Data Quality",
        (
            "Assess completeness, consistency, "
            "and master-data readiness."
        ),
    )

    df = render_local_filters(df, "quality")

    quality = pd.DataFrame({
        "column": df.columns,
        "missing_count": (
            df.isna()
            .sum()
            .values
        ),
        "missing_rate": (
            df.isna()
            .mean()
            .values
        ),
        "unique_count": [
            df[col].nunique(
                dropna=True
            )
            for col in df.columns
        ],
    })

    quality = quality.sort_values(
        "missing_rate",
        ascending=False,
    )

    critical_columns = [
        col
        for col in [
            "vendor",
            "po_no",
            "item_description",
            "qty_po",
            "qty_do",
            "grand_total_do",
        ]
        if col in df.columns
    ]

    if critical_columns:
        critical_missing = (
            df[critical_columns]
            .isna()
            .any(axis=1)
            .sum()
        )

    else:
        critical_missing = 0

    duplicate_rows = (
        df.duplicated()
        .sum()
    )

    average_completeness = (
        1
        - df.isna()
        .mean()
        .mean()
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Average Completeness",
        pct(average_completeness),
    )

    c2.metric(
        "Critical Incomplete Rows",
        number_id(critical_missing),
    )

    c3.metric(
        "Exact Duplicate Rows",
        number_id(duplicate_rows),
    )

    st.markdown(
        "#### Column Completeness"
    )

    display_df(
        quality,
        percent_cols=[
            "missing_rate",
        ],
    )

    item_master = pd.DataFrame()

    if "item_description" in df.columns:
        aggregation = {
            "transaction_count": (
                "item_description",
                "size",
            ),
        }

        if "grand_total_do" in df.columns:
            aggregation["spend_do"] = (
                "grand_total_do",
                "sum",
            )

        if "unit" in df.columns:
            aggregation["unit_count"] = (
                "unit",
                "nunique",
            )

        if "vendor" in df.columns:
            aggregation["vendor_count"] = (
                "vendor",
                "nunique",
            )

        item_master = (
            df
            .groupby("item_description")
            .agg(**aggregation)
            .reset_index()
        )

        if "spend_do" in item_master.columns:
            item_master = (
                item_master
                .sort_values(
                    "spend_do",
                    ascending=False,
                )
            )

        st.markdown(
            "#### Item Master Review"
        )

        display_df(
            item_master.head(300),
            currency_cols=[
                "spend_do",
            ],
        )

    export_button(
        label="Export Data Quality",
        file_name="data_quality.xlsx",
        sheets={
            "Column Quality": quality,
            "Item Master Review": item_master,
            "Filtered Raw Data": df,
        },
        title=(
            "Purchasing Data Quality Review"
        ),
        key="export_quality",
    )


# =========================================================
# DETAIL DATA (RAW ROW-LEVEL EXPLORER)
# =========================================================
DETAIL_SEARCH_COLUMNS = [
    "po_no",
    "pr_no",
    "vendor",
    "item_description",
    "item_category",
    "cost_category",
    "cost_center",
    "site",
]

DETAIL_CURRENCY_COLS = [
    "price_unit",
    "grand_total_po",
    "grand_total_do",
    "total_po",
    "total_do",
    "value_variance_do_vs_po",
]

DETAIL_PERCENT_COLS = [
    "fulfillment_rate",
]

DETAIL_ROW_DISPLAY_CAP = 3000


@st.fragment
def show_detail_data(
    df: pd.DataFrame,
) -> None:
    """
    Pembacaan detail: tabel mentah baris-per-baris yang bisa
    dicari, difilter, dan diekspor — untuk audit menyeluruh.
    """

    section_heading(
        "Data Explorer",
        "Pembacaan Detail (Raw Data)",
        (
            "Telusuri data transaksi baris demi baris — "
            "cari, urutkan, dan ekspor sesuai kebutuhan audit."
        ),
    )

    df = render_local_filters(df, "detail")

    search_columns = [
        col
        for col in DETAIL_SEARCH_COLUMNS
        if col in df.columns
    ]

    search_term = st.text_input(
        "Cari (Nomor PO, Nomor PR, Vendor, Item, Kategori, dll.)",
        value="",
        placeholder="Ketik kata kunci lalu tekan Enter...",
        key="detail_search_term",
    )

    filtered = df

    if search_term and search_columns:
        mask = pd.Series(
            False,
            index=filtered.index,
        )

        for col in search_columns:
            mask = mask | (
                filtered[col]
                .astype(str)
                .str.contains(
                    search_term,
                    case=False,
                    na=False,
                    regex=False,
                )
            )

        filtered = filtered[mask]

    all_columns = list(df.columns)

    default_columns = [
        col
        for col in all_columns
        if col not in ("month_no", "month_sort_key")
    ]

    selected_columns = st.multiselect(
        "Kolom yang ditampilkan",
        options=all_columns,
        default=default_columns,
        key="detail_selected_columns",
    )

    display_columns = (
        selected_columns
        if selected_columns
        else default_columns
    )

    row_note = f"{number_id(len(filtered))} baris ditemukan"

    if search_term:
        row_note += f" untuk pencarian \"{search_term}\""

    st.caption(row_note)

    table_view = filtered[display_columns]

    if len(table_view) > DETAIL_ROW_DISPLAY_CAP:
        st.info(
            (
                f"Menampilkan {number_id(DETAIL_ROW_DISPLAY_CAP)} "
                f"baris pertama dari {number_id(len(table_view))} baris "
                "agar tabel tetap ringan. Persempit pencarian/filter, "
                "atau gunakan tombol export di bawah untuk mengambil "
                "seluruh baris."
            )
        )

        display_df(
            table_view.head(DETAIL_ROW_DISPLAY_CAP),
            currency_cols=DETAIL_CURRENCY_COLS,
            percent_cols=DETAIL_PERCENT_COLS,
        )

    else:
        display_df(
            table_view,
            currency_cols=DETAIL_CURRENCY_COLS,
            percent_cols=DETAIL_PERCENT_COLS,
        )

    export_button(
        label="Export Detail Data (seluruh baris)",
        file_name="detail_data.xlsx",
        sheets={
            "Detail Data": table_view,
        },
        title="Purchasing Detail Data Export",
        key="export_detail_data",
    )

    reading_guide([
        (
            "Gunakan kolom pencarian untuk menemukan transaksi "
            "spesifik lewat Nomor PO, Nomor PR, nama vendor, "
            "atau nama item — tidak perlu scroll manual."
        ),
        (
            "Klik header kolom pada tabel untuk mengurutkan "
            "(ascending/descending) sesuai kebutuhan audit."
        ),
        (
            "Gunakan \"Filter halaman ini\" untuk mempersempit "
            "ke periode, site, vendor, atau kategori tertentu "
            "sebelum mencari lebih detail."
        ),
        (
            "Tombol export selalu mengambil SELURUH baris hasil "
            "filter/pencarian saat ini, bukan hanya yang tampil "
            "di layar — aman dipakai untuk dokumentasi audit."
        ),
    ])


# =========================================================
# PEMBACAAN BERDASARKAN NOMOR PO (PO LOOKUP)
# =========================================================
PO_LIFECYCLE_MILESTONES = [
    ("pr_date", "PR Diajukan", "min"),
    ("po_date", "PO Diterbitkan", "min"),
    ("arrival_date", "Barang/Jasa Diterima (DO)", "max"),
    ("invoice_received_date", "Invoice Diterima", "max"),
    ("actual_payout_date", "Pembayaran Dilakukan", "max"),
]

PO_LEAD_TIME_STEPS = [
    ("lt_pr_to_po", "PR → PO"),
    ("lt_po_to_arrival", "PO → Arrival"),
    ("lt_arrival_to_invoice", "Arrival → Invoice"),
    ("lt_invoice_to_schedule", "Invoice → Jadwal Bayar"),
    ("lt_schedule_to_payout", "Jadwal → Bayar Aktual"),
]


def build_po_timeline_figure(
    milestones: list[tuple[str, "pd.Timestamp"]],
) -> go.Figure:
    """
    Membuat visual timeline horizontal untuk satu Nomor PO,
    dari PR sampai pembayaran.
    """

    ordered = sorted(
        milestones,
        key=lambda item: item[1],
    )

    x_values = [item[1] for item in ordered]
    labels = [item[0] for item in ordered]

    figure = go.Figure()

    figure.add_trace(
        go.Scatter(
            x=x_values,
            y=[0] * len(x_values),
            mode="lines+markers",
            marker={
                "size": 16,
                "color": COLOR["accent"],
                "line": {
                    "width": 2,
                    "color": "#FFFFFF",
                },
            },
            line={
                "color": COLOR["accent_2"],
                "width": 3,
            },
            text=labels,
            hovertemplate=(
                "%{text}<br>%{x|%d %b %Y}<extra></extra>"
            ),
            showlegend=False,
        )
    )

    # Label tiap milestone dipasang lewat annotation (bukan
    # trace teks) dan diselang-seling atas/bawah, supaya kalau
    # dua tanggal berdekatan, teksnya tidak saling tumpang tindih.
    annotations = []

    for index, (label, value) in enumerate(
        zip(labels, x_values)
    ):
        goes_up = (index % 2 == 0)

        annotations.append({
            "x": value,
            "y": 0,
            "text": (
                f"<b>{label}</b><br>"
                f"{value.strftime('%d %b %Y')}"
            ),
            "showarrow": True,
            "arrowhead": 0,
            "arrowcolor": COLOR["line"],
            "ax": 0,
            "ay": -46 if goes_up else 46,
            "align": "center",
            "font": {
                "size": 11,
            },
        })

    figure.update_yaxes(
        visible=False,
        range=[-1.6, 1.6],
    )

    figure.update_layout(
        showlegend=False,
        title="Timeline Perjalanan PO",
        annotations=annotations,
    )

    return figure


@st.fragment
def show_po_lookup(
    df: pd.DataFrame,
) -> None:
    """
    Pembacaan berdasarkan Nomor PO: cari satu Nomor PO lalu
    tampilkan seluruh riwayatnya dari PR sampai pembayaran.
    """

    section_heading(
        "Data Explorer",
        "Pembacaan berdasarkan Nomor PO",
        (
            "Cari satu Nomor PO untuk melihat seluruh "
            "perjalanannya — dari pengajuan PR sampai "
            "pembayaran selesai."
        ),
    )

    if (
        "po_no" not in df.columns
        or df["po_no"].dropna().empty
    ):
        st.warning(
            "Kolom Nomor PO tidak tersedia pada data ini."
        )
        return

    df = render_local_filters(df, "polookup")

    po_options = sorted(
        df["po_no"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    if not po_options:
        st.info(
            "Tidak ada Nomor PO pada data hasil filter saat ini."
        )
        return

    selected_po = st.selectbox(
        "Pilih atau ketik Nomor PO",
        options=po_options,
        index=None,
        placeholder="Contoh: ketik sebagian nomor PO untuk mencari...",
        key="po_lookup_selected",
    )

    if not selected_po:
        st.info(
            f"{number_id(len(po_options))} Nomor PO tersedia. "
            "Pilih salah satu di atas untuk melihat riwayat "
            "lengkapnya."
        )
        return

    po_rows = df[
        df["po_no"].astype(str) == str(selected_po)
    ].copy()

    if po_rows.empty:
        st.warning(
            "Nomor PO tidak ditemukan pada data yang sedang aktif."
        )
        return

    vendor_names = (
        sorted(
            po_rows["vendor"].dropna().unique().tolist()
        )
        if "vendor" in po_rows.columns
        else []
    )

    site_names = (
        sorted(
            po_rows["site"].dropna().unique().tolist()
        )
        if "site" in po_rows.columns
        else []
    )

    total_po_value = (
        po_rows["grand_total_po"].sum()
        if "grand_total_po" in po_rows.columns
        else np.nan
    )

    total_do_value = (
        po_rows["grand_total_do"].sum()
        if "grand_total_do" in po_rows.columns
        else np.nan
    )

    variance = (
        total_do_value - total_po_value
        if pd.notna(total_po_value) and pd.notna(total_do_value)
        else np.nan
    )

    m1, m2, m3, m4, m5 = st.columns(5)

    m1.metric(
        "Vendor",
        ", ".join(vendor_names) if vendor_names else "-",
    )

    m2.metric(
        "Jumlah Line Item",
        number_id(len(po_rows)),
    )

    m3.metric(
        "Total PO",
        rupiah(total_po_value),
    )

    m4.metric(
        "Total DO",
        rupiah(total_do_value),
    )

    m5.metric(
        "Variance DO vs PO",
        rupiah(variance),
    )

    status_columns = st.columns(3)

    if "delivery_status" in po_rows.columns:
        delivery_values = (
            po_rows["delivery_status"].dropna().unique()
        )

        delivery_label = (
            delivery_values[0]
            if len(delivery_values) == 1
            else "MIXED"
        )

        delivery_level = {
            "FULFILLED": "good",
            "SHORT DELIVERY": "bad",
            "OVER DELIVERY": "warn",
        }.get(delivery_label, "neutral")

        with status_columns[0]:
            status_pill_markdown(
                "Delivery Status",
                delivery_label,
                delivery_level,
            )

    if "payment_status" in po_rows.columns:
        payment_values = (
            po_rows["payment_status"].dropna().unique()
        )

        payment_label = (
            payment_values[0]
            if len(payment_values) == 1
            else "MIXED"
        )

        payment_level = (
            "good"
            if payment_label == "PAID"
            else "warn"
            if payment_label == "MIXED"
            else "bad"
            if payment_label == "OPEN"
            else "neutral"
        )

        with status_columns[1]:
            status_pill_markdown(
                "Payment Status",
                payment_label,
                payment_level,
            )

    if site_names:
        with status_columns[2]:
            status_pill_markdown(
                "Site",
                ", ".join(site_names),
                "neutral",
            )

    milestones = []

    for column, label, aggregation in PO_LIFECYCLE_MILESTONES:
        if column not in po_rows.columns:
            continue

        valid_values = po_rows[column].dropna()

        if valid_values.empty:
            continue

        value = (
            valid_values.min()
            if aggregation == "min"
            else valid_values.max()
        )

        milestones.append((label, value))

    if len(milestones) >= 2:
        st.plotly_chart(
            clean_layout(
                build_po_timeline_figure(milestones),
                height=260,
            ),
            use_container_width=True,
        )

    else:
        st.caption(
            "Data tanggal tidak lengkap untuk membuat timeline "
            "PO ini."
        )

    lead_time_available = [
        (column, label)
        for column, label in PO_LEAD_TIME_STEPS
        if column in po_rows.columns
        and po_rows[column].notna().any()
    ]

    if lead_time_available:
        st.markdown(
            "#### Lead Time per Tahap"
        )

        lead_time_columns = st.columns(
            len(lead_time_available)
        )

        for index, (column, label) in enumerate(
            lead_time_available
        ):
            average_days = po_rows[column].mean()

            lead_time_columns[index].metric(
                label,
                f"{average_days:,.0f} hari".replace(",", "."),
            )

    st.markdown(
        "#### Detail Line Item"
    )

    line_item_columns = [
        col
        for col in [
            "item_description",
            "item_category",
            "unit",
            "qty_po",
            "qty_do",
            "price_unit",
            "grand_total_po",
            "grand_total_do",
            "delivery_status",
            "payment_status",
        ]
        if col in po_rows.columns
    ]

    display_df(
        po_rows[line_item_columns],
        currency_cols=DETAIL_CURRENCY_COLS,
        percent_cols=DETAIL_PERCENT_COLS,
    )

    export_button(
        label="Export Riwayat PO Ini",
        file_name=f"po_{selected_po}_detail.xlsx",
        sheets={
            "PO Detail": po_rows,
        },
        title=f"Riwayat Purchase Order {selected_po}",
        key="export_po_lookup",
    )

    reading_guide([
        (
            "Ketik sebagian Nomor PO pada kotak pilihan untuk "
            "mencari dengan cepat, tidak perlu scroll daftar panjang."
        ),
        (
            "Timeline menunjukkan urutan tanggal aktual PO ini — "
            "jika ada tahap yang meloncat jauh, itu tanda proses "
            "tertahan di tahap tersebut."
        ),
        (
            "Status \"MIXED\" pada Delivery/Payment berarti PO ini "
            "punya lebih dari satu line item dengan status berbeda — "
            "cek tabel detail di bawah untuk rinciannya."
        ),
        (
            "Gunakan tombol export untuk mengirim riwayat PO ini "
            "sebagai bukti/lampiran ke vendor atau tim terkait."
        ),
    ])


# =========================================================
# MAIN APPLICATION
# =========================================================
def main() -> None:
    """
    Fungsi utama aplikasi.
    """

    dark_mode = st.sidebar.toggle(
        "Dark Mode",
        value=False,
        key="dark_mode",
    )

    apply_app_style(
        dark_mode
    )

    st.markdown(
        (
            '<div class="section-label">'
            "Procurement Analytics Workspace"
            "</div>"
        ),
        unsafe_allow_html=True,
    )

    st.title(
        "Procurement Control Center"
    )

    st.markdown(
        (
            '<div class="app-subtitle">'
            "Personalized for data analysis, "
            "purchasing management, and cost control."
            "</div>"
        ),
        unsafe_allow_html=True,
    )

    gsheet_ready = gsheet_is_configured()

    uploaded_files = (
        st.sidebar.file_uploader(
            "Upload purchasing workbook(s)",
            type=["xlsx"],
            accept_multiple_files=True,
            help=(
                (
                    "Opsional — kalau tidak upload apa pun, "
                    "dashboard otomatis pakai data yang sudah "
                    "tersinkron dari Google Sheets."
                )
                if gsheet_ready
                else None
            ),
        )
    )

    active_file = None
    file_bytes = None
    file_path = None
    combine_mode = False
    named_file_bytes = None
    use_gsheet = False
    gsheet_id = None
    gsheet_worksheet = "RAW DATA"
    gsheet_header_row = 1

    if uploaded_files:
        file_lookup = {
            uploaded.name: uploaded
            for uploaded in uploaded_files
        }

        file_names = list(
            file_lookup.keys()
        )

        if len(file_names) > 1:
            mode_choice = st.sidebar.radio(
                "Multi-file mode",
                [
                    "Switch (one active file)",
                    "Combine (merge all files)",
                ],
                key="multi_file_mode",
            )

            combine_mode = mode_choice.startswith(
                "Combine"
            )

            if combine_mode:
                named_file_bytes = tuple(
                    (
                        name,
                        file_lookup[name].getvalue(),
                    )
                    for name in file_names
                )

                st.sidebar.success(
                    (
                        f"Combined source: {len(file_names)} "
                        "workbooks merged"
                    )
                )

            else:
                selected_name = st.sidebar.selectbox(
                    "Active workbook",
                    file_names,
                    key="active_workbook_name",
                )

                active_file = file_lookup[
                    selected_name
                ]

                file_bytes = (
                    active_file.getvalue()
                )

                st.sidebar.success(
                    f"Active source: {selected_name}"
                )

                st.sidebar.caption(
                    (
                        f"{len(file_names)} workbooks uploaded. "
                        "Switch to \"Combine\" above to merge "
                        "them into one analysis."
                    )
                )

        else:
            selected_name = file_names[0]

            active_file = file_lookup[
                selected_name
            ]

            file_bytes = (
                active_file.getvalue()
            )

            st.sidebar.success(
                f"Active source: {selected_name}"
            )

    elif gsheet_ready:
        use_gsheet = True
        gsheet_id = st.secrets["gsheet_id"]

        gsheet_worksheet = st.secrets.get(
            "gsheet_worksheet_name",
            "RAW DATA",
        )

        gsheet_header_row = int(
            st.secrets.get(
                "gsheet_header_row",
                1,
            )
        )

        st.sidebar.success(
            "Live source: Google Sheets (auto-sync)"
        )

        st.sidebar.caption(
            "Data disegarkan otomatis tiap 5 menit. "
            "Baru saja update Sheet-nya? Klik refresh "
            "di bawah supaya langsung kebaca."
        )

        if st.sidebar.button(
            "🔄 Refresh data sekarang",
            key="refresh_gsheet",
            use_container_width=True,
        ):
            load_data_from_gsheet.clear()
            st.rerun()

    elif DEFAULT_FILE.exists():
        file_path = str(
            DEFAULT_FILE
        )

        st.sidebar.success(
            (
                "Local source: "
                f"{DEFAULT_FILE.name}"
            )
        )

    else:
        st.info(
            (
                "Upload one or more Excel workbooks containing "
                "the `RAW DATA` sheet (header on row 2), or "
                "connect a Google Sheet as the live data source "
                "via Streamlit Secrets."
            )
        )
        return

    if combine_mode and named_file_bytes:
        data_fingerprint = (
            "combine",
            tuple(
                sorted(
                    name
                    for name, _ in named_file_bytes
                )
            ),
        )

    elif use_gsheet:
        data_fingerprint = (
            "gsheet",
            gsheet_id,
            gsheet_worksheet,
            gsheet_header_row,
        )

    elif file_path:
        data_fingerprint = (
            "default",
            DEFAULT_FILE.name,
        )

    elif active_file is not None:
        data_fingerprint = (
            "switch",
            active_file.name,
        )

    else:
        data_fingerprint = (
            "unknown",
        )

    previous_fingerprint = st.session_state.get(
        "data_fingerprint"
    )

    if (
        previous_fingerprint is not None
        and previous_fingerprint != data_fingerprint
    ):
        # Sumber data berubah (ganti file / switch ke combine).
        # Filter lama (terutama Arrival Date) disimpan di
        # session_state per-widget dan TIDAK otomatis
        # menyesuaikan rentang data yang baru — kalau dibiarkan,
        # data dari file baru bisa diam-diam ke-filter habis
        # (grafik jadi kosong/tidak muncul). Jadi semua filter
        # lokal per-tab direset supaya selalu konsisten dengan
        # data yang sedang aktif.
        reset_all_local_filters()

        st.sidebar.info(
            "Sumber data berubah — semua filter direset "
            "otomatis supaya tidak ada data baru yang "
            "tertutup filter lama."
        )

    st.session_state.data_fingerprint = data_fingerprint

    load_start_time = time.perf_counter()

    try:
        with st.spinner(
            "Loading and preparing purchasing data..."
        ):
            if combine_mode and named_file_bytes:
                df = load_combined_data(
                    named_file_bytes
                )

            elif use_gsheet:
                df = load_data_from_gsheet(
                    gsheet_id,
                    gsheet_worksheet,
                    gsheet_header_row,
                )

            else:
                df = load_data(
                    file_bytes,
                    file_path,
                )

    except FileNotFoundError as error:
        st.error(
            f"File error: {error}"
        )
        return

    except ValueError as error:
        st.error(
            (
                "Workbook structure is not valid. "
                f"Detail: {error}"
            )
        )
        return

    except Exception as error:
        message = (
            "Unable to read the Google Sheet. "
            f"Detail: {error}"
            if use_gsheet
            else (
                "Unable to read the workbook. "
                f"Detail: {error}"
            )
        )

        st.error(
            message
        )
        return

    load_duration_seconds = (
        time.perf_counter()
        - load_start_time
    )

    st.sidebar.caption(
        (
            f"Data dimuat dalam "
            f"{load_duration_seconds:.2f} detik "
            f"({number_id(len(df))} baris)."
        )
    )

    if df.empty:
        st.warning(
            (
                "No records match the current "
                "filter selection."
            )
        )
        return

    tabs = st.tabs([
        "Overview",
        "Spend",
        "Vendors",
        "ABC Classification",
        "Price",
        "PO vs DO",
        "Lead Time",
        "Payment",
        "Data Quality",
        "Detail Data",
        "Cari PO",
    ])

    with tabs[0]:
        show_executive(
            df
        )

    with tabs[1]:
        show_spend(
            df
        )

    with tabs[2]:
        show_vendor(
            df
        )

    with tabs[3]:
        show_segmentation(
            df
        )

    with tabs[4]:
        show_price(
            df
        )

    with tabs[5]:
        show_po_do(
            df
        )

    with tabs[6]:
        show_lead_time(
            df
        )

    with tabs[7]:
        show_payment(
            df
        )

    with tabs[8]:
        show_quality(
            df
        )

    with tabs[9]:
        show_detail_data(
            df
        )

    with tabs[10]:
        show_po_lookup(
            df
        )

    st.divider()

    st.caption(
        (
            "Analytical note: saving opportunities are "
            "directional. Validate specifications, units, "
            "tax treatment, vendor terms, and comparable "
            "periods before negotiation or accrual decisions."
        )
    )


# =========================================================
# RUN APPLICATION
# =========================================================
if __name__ == "__main__":
    main()
