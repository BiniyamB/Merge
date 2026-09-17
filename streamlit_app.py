"""Report Merger + Digital Transaction Value Snapshot -- Streamlit version."""

import gc
import time
import html as html_lib
import io
from datetime import date as _date

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd

st.set_page_config(
    page_title="Report Merger (POS, ATM, IPS, QR & P2P)",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

from merger import MODES, merge_reports, build_filtered_workbook, build_workbook

from snapshot_module import REPORT_DEFAULTS, SERVICE_DEFAULTS, calc_all, build_report_html, fmt_int
from acquirer_analysis import analyze_atm, analyze_pos, analyze_ips
from nbe_report import (
    generate_nbe_report,
    build_nbe_report_excel,
    generate_sett_sum_report,
    build_sett_sum_report_excel,
)
from pos_success_rate import generate_pos_success_rate_report, build_pos_success_rate_excel
from interop_report import (
    merge_bank_summary_files,
    build_success_report_excel,
    build_merged_summary_excel,
    success_report_filename,
)
from ips_report import (
    parse_ips_report,
    collect_ips_dates,
    filter_ips_by_dates,
)

# ── Global CSS (dark / light themes) ──────────────────────────────────────
_THEME = st.session_state.get("theme", "dark")
st.session_state.setdefault("mode_key", "pos_decline")

_PALETTES = {
    "dark": {
        "BG": "linear-gradient(160deg,#050a18 0%,#0a1128 40%,#0d1a30 100%)",
        "PANEL": "#0b1122",
        "CARD": "rgba(255,255,255,0.035)",
        "CARD_HOVER": "rgba(255,255,255,0.06)",
        "BORDER": "rgba(255,255,255,0.08)",
        "TEXT": "#e7edff",
        "MUTED": "#8a94b8",
        "ACCENT": "#a855f7",
        "GRAD": "linear-gradient(135deg,#ff3cac 0%,#784ba0 40%,#2b86c5 80%,#00d4ff 100%)",
        "SEP_BG": "rgba(10,17,40,0.95)",
        "RADIO_C": "#c4b5fd",
        "SHEET_BG": "rgba(168,85,247,0.06)",
        "SHEET_BORDER": "rgba(168,85,247,0.18)",
        "SHADOW": "0 18px 50px rgba(0,0,0,0.5)",
    },
    "light": {
        "BG": "linear-gradient(160deg,#eef1f9 0%,#f7f9ff 45%,#e9efff 100%)",
        "PANEL": "#ffffff",
        "CARD": "rgba(255,255,255,0.82)",
        "CARD_HOVER": "rgba(255,255,255,1)",
        "BORDER": "rgba(24,38,78,0.12)",
        "TEXT": "#16203c",
        "MUTED": "#5a6a9a",
        "ACCENT": "#7c3aed",
        "GRAD": "linear-gradient(135deg,#d81f8b 0%,#7c3aed 45%,#1d7de0 85%,#0096c7 100%)",
        "SEP_BG": "rgba(247,249,255,0.98)",
        "RADIO_C": "#6d28d9",
        "SHEET_BG": "rgba(124,58,237,0.07)",
        "SHEET_BORDER": "rgba(124,58,237,0.25)",
        "SHADOW": "0 18px 50px rgba(24,38,78,0.14)",
    },
}
_P = _PALETTES[_THEME]

st.markdown("""
<style>
    @keyframes fadeUp {
        from { opacity: 0; transform: translateY(14px); }
        to   { opacity: 1; transform: translateY(0); } }
    @keyframes glowPulse {
        0%, 100% { filter: drop-shadow(0 0 6px rgba(168,85,247,0.35)); }
        50%      { filter: drop-shadow(0 0 18px rgba(0,212,255,0.45)); } }
    @keyframes floatY {
        0%, 100% { transform: translateY(0); }
        50%      { transform: translateY(-6px); } }

    /* Override Streamlit's native theme variables so every widget follows */
    :root {
        --background-color: @@BG@@ !important;
        --secondary-background-color: @@PANEL@@ !important;
        --text-color: @@TEXT@@ !important;
        --primary-color: @@ACCENT@@ !important;
        --font: 'Plus Jakarta Sans', sans-serif !important;
    }

    /* Hide Streamlit chrome */
    #MainMenu, footer, header[data-testid="stHeader"] {
        visibility: hidden !important; height: 0 !important;
        margin: 0 !important; padding: 0 !important; }
    div[data-testid="stToolbar"] { display: none !important; }
    div[data-testid="stDecoration"] { display: none !important; }
    div[data-testid="stStatusWidget"] { display: none !important; }
    div[data-testid="stDeployButton"] { display: none !important; }
    button[title="View fullscreen"] { display: none !important; }
    div.stProfiler { display: none !important; }

    /* Base */
    .stApp {
        background: @@BG@@ !important;
        font-family: 'Plus Jakarta Sans', sans-serif; }
    .block-container { max-width: 1160px; padding-top: 1.4rem; padding-bottom: 2rem; }
    html, body, [data-testid="stAppViewContainer"] { background: transparent !important; }

    /* Gradient title */
    .gradient-title {
        font-size: 2.3rem; font-weight: 800; letter-spacing: -1px;
        background: @@GRAD@@;
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text; margin-bottom: 0;
        animation: floatY 6s ease-in-out infinite; }
    .subtitle { color: @@MUTED@@; font-size: 0.95rem; margin-top: -4px; }

    /* Theme toggle */
    .theme-toggle {
        display: flex; align-items: center; justify-content: flex-end; gap: 8px;
        min-height: 44px; }
    .theme-toggle span { color: @@MUTED@@; font-size: 0.75rem; font-weight: 600;
        text-transform: uppercase; letter-spacing: 0.6px; }

    /* Cards */
    .card {
        background: @@CARD@@;
        border: 1px solid @@BORDER@@;
        border-radius: 18px; padding: 24px 28px; margin-bottom: 20px;
        backdrop-filter: blur(12px);
        box-shadow: @@SHADOW@@;
        animation: fadeUp 0.5s ease both; }
    .card:hover { border-color: rgba(168,85,247,0.28); }
    .card-head {
        display: flex; align-items: center; gap: 12px; margin-bottom: 18px; }
    .card-icon {
        width: 40px; height: 40px; border-radius: 12px; display: flex;
        align-items: center; justify-content: center; font-size: 1.15rem;
        flex-shrink: 0; }
    .icon-blue { background: rgba(0,212,255,0.12); border: 1px solid rgba(0,212,255,0.2); }
    .icon-purple { background: rgba(168,85,247,0.12); border: 1px solid rgba(168,85,247,0.2); }
    .icon-green { background: rgba(0,232,143,0.12); border: 1px solid rgba(0,232,143,0.2); }
    .icon-pink { background: rgba(255,60,172,0.12); border: 1px solid rgba(255,60,172,0.2); }
    .icon-amber { background: rgba(251,191,36,0.14); border: 1px solid rgba(251,191,36,0.28); }
    .card-title { font-size: 1.05rem; font-weight: 800; color: @@TEXT@@; margin: 0; }
    .card-sub { font-size: 0.8rem; color: @@MUTED@@; margin: 2px 0 0; }

    /* Mode cards (report type picker) */
    .mode-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px;
        margin: 6px 0 4px; }
    .mode-card {
        border-radius: 16px; padding: 18px 16px 14px; text-align: center;
        border: 1px solid @@BORDER@@; background: @@CARD@@;
        cursor: pointer; transition: all 0.25s ease;
        animation: fadeUp 0.5s ease both; }
    .mode-card:hover { transform: translateY(-5px); box-shadow: @@SHADOW@@; }
    .mode-card.mode-active {
        border-color: @@ACCENT@@;
        box-shadow: 0 0 0 3px rgba(168,85,247,0.18), @@SHADOW@@; }
    .mode-card .mc-icon { font-size: 2rem; line-height: 1; display: block;
        margin-bottom: 8px; }
    .mode-card .mc-check {
        position: absolute; top: 10px; right: 14px; font-size: 1rem;
        display: none; }
    .mode-card.mode-active .mc-check { display: block; }
    .mode-card .mc-name { font-weight: 800; font-size: 1rem; color: @@TEXT@@; }
    .mode-card .mc-desc { font-size: 0.72rem; color: @@MUTED@@; margin-top: 5px;
        line-height: 1.35; }
    .mode-card.mode-active .mc-name { color: @@ACCENT@@; }
    .type-red     { background: linear-gradient(150deg, rgba(255,71,87,0.10), @@CARD@@ 60%); }
    .type-green   { background: linear-gradient(150deg, rgba(0,232,143,0.10), @@CARD@@ 60%); }
    .type-purple  { background: linear-gradient(150deg, rgba(168,85,247,0.10), @@CARD@@ 60%); }
    .type-blue    { background: linear-gradient(150deg, rgba(0,212,255,0.10), @@CARD@@ 60%); }
    .type-cyan    { background: linear-gradient(150deg, rgba(34,211,238,0.10), @@CARD@@ 60%); }
    .type-amber   { background: linear-gradient(150deg, rgba(251,191,36,0.12), @@CARD@@ 60%); }
    .type-violet  { background: linear-gradient(150deg, rgba(139,92,246,0.12), @@CARD@@ 60%); }
    .mode-card .stButton > button { margin-top: 10px !important; }

    /* Metrics */
    div[data-testid="stMetric"] {
        background: @@CARD@@;
        border: 1px solid @@BORDER@@;
        border-radius: 14px; padding: 16px 18px;
        transition: transform 0.2s, border-color 0.2s;
        box-shadow: 0 8px 24px rgba(0,0,0,0.12); }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-3px);
        border-color: rgba(168,85,247,0.3); }
    div[data-testid="stMetric"] label {
        color: @@MUTED@@ !important; font-size: 0.72rem !important;
        text-transform: uppercase; letter-spacing: 0.8px; font-weight: 700 !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: @@ACCENT@@ !important; font-size: 1.6rem !important; font-weight: 800 !important; }

    /* Radio */
    div[data-baseweb="radio"] > label {
        background: @@CARD@@;
        border: 1px solid @@BORDER@@;
        border-radius: 12px; padding: 12px 20px; margin: 0;
        cursor: pointer; transition: all 0.25s; color: @@TEXT@@; }
    div[data-baseweb="radio"] > label:hover {
        border-color: rgba(168,85,247,0.35); }
    div[data-baseweb="radio"] > label[data-checked="true"] {
        border-color: rgba(168,85,247,0.45);
        background: rgba(168,85,247,0.08); }
    div[data-baseweb="radio"] > label[data-checked="true"] > div {
        color: @@RADIO_C@@ !important; font-weight: 700 !important; }

    /* Buttons */
    .stButton > button {
        background: @@GRAD@@ !important;
        color: white !important; border: none !important; border-radius: 12px !important;
        font-weight: 700 !important; font-size: 0.92rem !important;
        padding: 0.6rem 2.5rem !important;
        box-shadow: 0 4px 20px rgba(168,85,247,0.25) !important;
        transition: all 0.3s !important; font-family: 'Plus Jakarta Sans', sans-serif !important; }
    .stButton > button:hover {
        box-shadow: 0 6px 28px rgba(168,85,247,0.4) !important;
        transform: translateY(-2px) !important; }
    .stButton > button:active { transform: translateY(0) !important; }

    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #00e88f 0%, #00b4d8 100%) !important;
        color: #04221a !important; border: none !important; border-radius: 12px !important;
        font-weight: 700 !important; font-size: 0.92rem !important;
        padding: 0.6rem 2.5rem !important;
        box-shadow: 0 4px 20px rgba(0,232,143,0.25) !important;
        transition: all 0.3s !important; font-family: 'Plus Jakarta Sans', sans-serif !important; }
    div[data-testid="stDownloadButton"] > button:hover {
        box-shadow: 0 6px 28px rgba(0,232,143,0.4) !important;
        transform: translateY(-2px) !important; }

    /* Wiggle a selectbox label */
    .stSelectbox label, .stMultiselect label, .stDateInput label { font-weight: 700 !important; }

    .filter-dl > button {
        background: linear-gradient(135deg, #a855f7 0%, #6366f1 50%, #3b82f6 100%) !important;
        color: white !important;
        box-shadow: 0 4px 20px rgba(168,85,247,0.2) !important; }

    /* Selectbox */
    div[data-baseweb="select"] > div {
        background: @@CARD@@ !important;
        border-color: @@BORDER@@ !important;
        border-radius: 10px !important; color: @@TEXT@@ !important; }
    div[data-baseweb="select"] > div:hover { border-color: rgba(168,85,247,0.35) !important; }
    div[data-baseweb="select"] > div:focus-within {
        border-color: @@ACCENT@@ !important;
        box-shadow: 0 0 0 2px rgba(168,85,247,0.18) !important; }
    div[data-baseweb="select"] [role="listbox"] { color: @@TEXT@@ !important; }
    div[data-baseweb="popover"] > div { background: @@PANEL@@ !important; color: @@TEXT@@ !important; }

    /* Multiselect */
    div[data-baseweb="tag"] { background: rgba(168,85,247,0.16) !important;
        border-radius: 8px !important; }
    div[data-baseweb="tag"] span { color: @@TEXT@@ !important; }
    ul[data-testid="stMultiselectDropdown"] { background: @@PANEL@@ !important; }
    ul[data-testid="stMultiselectDropdown"] img { border-radius: 6px !important; }

    /* File uploader */
    section[data-testid="stFileUploadDropzone"] {
        background: @@CARD@@ !important;
        border: 2px dashed rgba(168,85,247,0.3) !important;
        border-radius: 16px !important; }
    section[data-testid="stFileUploadDropzone"]:hover {
        border-color: rgba(168,85,247,0.55) !important;
        background: rgba(168,85,247,0.04) !important; }
    section[data-testid="stFileUploadDropzone"] button {
        background: @@GRAD@@ !important; }
    div[data-testid="stFileUploaderDropzoneInstructions"] div { color: @@MUTED@@ !important; }

    /* Tables */
    .stDataFrame { border-radius: 12px !important; overflow: hidden; }

    /* Expanders */
    details[data-testid="stExpander"] {
        background: @@CARD@@ !important;
        border: 1px solid @@BORDER@@ !important;
        border-radius: 12px !important; }
    details[data-testid="stExpander"] summary {
        font-weight: 700 !important; color: @@TEXT@@ !important; }

    /* Dividers */
    hr { border-color: @@BORDER@@ !important; opacity: 0.5; }

    /* Alerts */
    div[data-testid="stAlert"] { border-radius: 12px !important; }
    div[data-testid="stAlert"] [data-testid="stAlertContainer"] { color: @@TEXT@@ !important; }

    /* Toggle (theme + dedupe) */
    div[data-testid="stToggle"] > div[role="switch"] { background: @@ACCENT@@ !important; }
    div[data-testid="stToggle"] label p { color: @@TEXT@@ !important; font-weight: 600 !important; }

    /* Badges */
    .badge {
        display: inline-block; padding: 3px 10px; border-radius: 999px;
        font-size: 0.7rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.5px; }
    .badge-green { background: rgba(0,232,143,0.14); color: #00e88f; border: 1px solid rgba(0,232,143,0.3); }
    .badge-purple { background: rgba(168,85,247,0.14); color: #c4b5fd; border: 1px solid rgba(168,85,247,0.3); }
    .badge-red { background: rgba(255,71,87,0.14); color: #ff8a8a; border: 1px solid rgba(255,71,87,0.3); }
    .badge-blue { background: rgba(0,212,255,0.14); color: #00d4ff; border: 1px solid rgba(0,212,255,0.3); }
    .badge-amber { background: rgba(251,191,36,0.16); color: #fbbf24; border: 1px solid rgba(251,191,36,0.34); }

    /* Filter sheet items */
    .sheet-item {
        background: @@SHEET_BG@@;
        border: 1px solid @@SHEET_BORDER@@;
        border-radius: 10px; padding: 12px 16px; margin-bottom: 8px;
        display: flex; align-items: center; gap: 10px; }
    .sheet-num {
        background: @@GRAD@@; color: white;
        width: 26px; height: 26px; border-radius: 8px; display: inline-flex;
        align-items: center; justify-content: center;
        font-size: 0.7rem; font-weight: 800; flex-shrink: 0; }
    .sheet-name { font-weight: 700; color: @@RADIO_C@@; }
    .sheet-desc { color: @@MUTED@@; font-size: 0.82rem; }
    .sheet-rows { color: @@MUTED@@; font-size: 0.75rem; margin-left: auto; white-space: nowrap; }

    /* Sidebar */
    section[data-testid="stSidebar"] { background: @@PANEL@@ !important; }

    /* Separator with text */
    .section-sep { text-align: center; margin: 28px 0 10px; position: relative; }
    .section-sep::before {
        content: ''; position: absolute; top: 50%; left: 0; right: 0; height: 1px;
        background: linear-gradient(90deg, transparent, rgba(168,85,247,0.3), transparent); }
    .section-sep span {
        position: relative; background: @@SEP_BG@@; padding: 6px 20px;
        font-size: 0.75rem; font-weight: 700; color: @@ACCENT@@;
        text-transform: uppercase; letter-spacing: 1px; border-radius: 999px;
        border: 1px solid rgba(168,85,247,0.2); }

    /* Page navigation tabs */
    .page-nav { display: flex; gap: 14px; margin: 4px 0 20px; }
    .page-nav .stButton > button {
        background: @@CARD@@ !important;
        border: 1px solid @@BORDER@@ !important;
        color: @@TEXT@@ !important;
        box-shadow: none !important;
        border-radius: 12px !important;
        font-weight: 700 !important; }
    .page-nav .stButton > button:hover {
        border-color: rgba(168,85,247,0.4) !important;
        background: rgba(168,85,247,0.06) !important;
        color: @@ACCENT@@ !important; }

    /* Snapshot editor */
    div[data-testid="stDataEditor"] {
        border: 1px solid @@BORDER@@ !important;
        border-radius: 12px !important; overflow: hidden; }

    /* Captions */
    .stCaption, [data-testid="stCaptionContainer"] p { color: @@MUTED@@ !important; }
    .info-caption { color: @@MUTED@@ !important; }

    /* Date pill list (IPS mode) */
    .date-pill {
        display: inline-block; padding: 5px 12px; margin: 0 6px 6px 0;
        border-radius: 999px; font-size: 0.75rem; font-weight: 700;
        background: rgba(251,191,36,0.14); color: #fbbf24;
        border: 1px solid rgba(251,191,36,0.3); }
    .date-pill-all {
        background: rgba(0,232,143,0.14); color: #00e88f;
        border: 1px solid rgba(0,232,143,0.3); }
</style>
""".replace("@@BG@@", _P["BG"]).replace("@@PANEL@@", _P["PANEL"])
.replace("@@CARD@@", _P["CARD"]).replace("@@CARD_HOVER@@", _P["CARD_HOVER"])
.replace("@@BORDER@@", _P["BORDER"]).replace("@@TEXT@@", _P["TEXT"])
.replace("@@MUTED@@", _P["MUTED"]).replace("@@ACCENT@@", _P["ACCENT"])
.replace("@@GRAD@@", _P["GRAD"]).replace("@@SEP_BG@@", _P["SEP_BG"])
.replace("@@RADIO_C@@", _P["RADIO_C"]).replace("@@SHEET_BG@@", _P["SHEET_BG"])
.replace("@@SHEET_BORDER@@", _P["SHEET_BORDER"]).replace("@@SHADOW@@", _P["SHADOW"]),
    unsafe_allow_html=True)

# ── State ────────────────────────────────────────────────────────────────────
for key, default in [
    ("merged_meta", None), ("filter_sheets", []),
    ("records", []), ("columns", []), ("mode_key", None),
    ("pending_filters", {}),
    ("unique_values_cache", {}),
    ("snap_page", False),
    ("duplicate_records", []),
    ("interop_records", []),
    ("theme", "dark"),
    ("ips_all_records", []), ("ips_dates", []), ("ips_selected_dates", []),
    ("ips_per_file", []), ("ips_warnings", []),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── Transaction Snapshot page ────────────────────────────────────────────────
def _render_snapshot_page():
    st.markdown('<p class="subtitle">Build the executive Digital Transaction Value Snapshot report &mdash; '
                'edit the numbers, the report updates instantly.</p>', unsafe_allow_html=True)

    with st.expander("Report Information", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            r_title = st.text_input("Report title", value=REPORT_DEFAULTS["title"], key="snap_rep_title")
            r_org = st.text_input("Organization", value=REPORT_DEFAULTS["organization"], key="snap_rep_org")
            r_subtitle = st.text_input("Subtitle", value=REPORT_DEFAULTS["subtitle"], key="snap_rep_subt")
            r_tagline1 = st.text_input("Tagline", value=REPORT_DEFAULTS["tagline1"], key="snap_rep_tag1")
        with c2:
            r_brand = st.text_input("Brand", value=REPORT_DEFAULTS["brand"], key="snap_rep_brand")
            r_tagline2 = st.text_input("Tagline 2 (footer)", value=REPORT_DEFAULTS["tagline2"], key="snap_rep_tag2")
            c2a, c2b = st.columns(2)
            r_date_from = c2a.text_input("Date From", value=REPORT_DEFAULTS["dateFrom"], key="snap_rep_date_from")
            r_date_to = c2b.text_input("Date To", value=REPORT_DEFAULTS["dateTo"], key="snap_rep_date_to")

    report = {
        "title": r_title, "organization": r_org, "brand": r_brand,
        "tagline1": r_tagline1, "tagline2": r_tagline2,
        "subtitle": r_subtitle, "dateFrom": r_date_from, "dateTo": r_date_to,
    }

    st.markdown('<div class="card"><div class="card-head"><div class="card-icon icon-green">&#128202;</div>'
                '<div><p class="card-title">Transaction services</p>'
                '<p class="card-sub">Edit / add / remove rows &mdash; calculations update live</p></div></div>',
                unsafe_allow_html=True)

    default_df = pd.DataFrame([
        {"name": s["name"], "type": s["type"],
         "target": float(s.get("target", 0)),
         "transactionVolume": float(s["transactionVolume"]),
         "totalValue": float(s["totalValue"]),
         "keyMessage": s["keyMessage"], "highlighted": s["highlighted"]}
        for s in SERVICE_DEFAULTS
    ])

    if st.button("Reset to reference data", key="snap_reset"):
        st.session_state.snap_editor = default_df
        st.rerun()

    edited = st.data_editor(
        default_df,
        key="snap_editor",
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "name": st.column_config.TextColumn("Service", width="medium"),
            "type": st.column_config.SelectboxColumn(
                "Type", options=["financial", "non-financial", "success-rate"], width="small"),
            "target": st.column_config.NumberColumn(
                "Monthly Plan (Target)", min_value=0.0, step=0.01, format="%.2f", width="small"),
            "transactionVolume": st.column_config.NumberColumn(
                "Performance", min_value=0.0, step=0.01, format="%.2f", width="small"),
            "totalValue": st.column_config.NumberColumn(
                "Total Value (ETB)", min_value=0.0, step=0.01, format="%.2f", width="large"),
            "keyMessage": st.column_config.TextColumn("Key Message", width="medium"),
            "highlighted": st.column_config.CheckboxColumn("Highlight", width="small"),
        },
    )

    s1, s2, s3 = st.columns(3)
    show_bars = s1.checkbox("Show metric bars", value=True, key="snap_bars")
    auto_hl = s2.checkbox("Auto-highlight highest average", value=True, key="snap_auto_hl")
    takeaway_override = s3.text_input("Key takeaway (overrides auto)", key="snap_takeaway")

    services = []
    for idx, row in enumerate(edited.to_dict("records")):
        services.append({
            "uid": f"row_{idx}",
            "name": str(row.get("name") or "") or f"Service {idx + 1}",
            "type": row.get("type") or "financial",
            "transactionVolume": row.get("transactionVolume"),
            "totalValue": row.get("totalValue"),
            "target": row.get("target"),
            "keyMessage": row.get("keyMessage") or "",
            "highlighted": bool(row.get("highlighted")),
        })

    calc = calc_all(services)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Top Acquirer Report Uploads ──────────────────────────────────────────
    st.markdown('<div class="section-sep"><span>Top Acquirer Analysis</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="card"><div class="card-head"><div class="card-icon icon-purple">&#128200;</div>'
                '<div><p class="card-title">Upload raw reports for Top 3 Acquirer analysis</p>'
                '<p class="card-sub">ATM (daily), POS (daily) and IPS (source/destination) &mdash; '
                'the system analyses cash withdrawals, POS purchases and IPS transaction counts</p>'
                '</div></div>', unsafe_allow_html=True)

    acq_col1, acq_col2, acq_col3 = st.columns(3)
    with acq_col1:
        atm_file = st.file_uploader("ATM Daily Report", type=["xls", "xlsx"], key="snap_atm_file")
    with acq_col2:
        pos_file = st.file_uploader("POS Daily Report", type=["xls", "xlsx"], key="snap_pos_file")
    with acq_col3:
        ips_file = st.file_uploader("IPS Source & Destination", type=["xls", "xlsx"], key="snap_ips_file")

    acquirer_data = {}
    acq_warnings = []

    if atm_file:
        with st.spinner("Analysing ATM report..."):
            atm_result = analyze_atm(atm_file.getvalue(), atm_file.name)
        acquirer_data["atm"] = atm_result
        acq_warnings.extend(atm_result.warnings)
    if pos_file:
        with st.spinner("Analysing POS report..."):
            pos_result = analyze_pos(pos_file.getvalue(), pos_file.name)
        acquirer_data["pos"] = pos_result
        acq_warnings.extend(pos_result.warnings)
    if ips_file:
        with st.spinner("Analysing IPS report..."):
            ips_result = analyze_ips(ips_file.getvalue(), ips_file.name)
        acquirer_data["ips"] = ips_result
        acq_warnings.extend(ips_result.warnings)

    if acq_warnings:
        for w in acq_warnings:
            st.warning(w)

    # Show quick summary of analysis results
    if acquirer_data:
        summary_cols = st.columns(4)
        with summary_cols[0]:
            if "atm" in acquirer_data and acquirer_data["atm"].top3:
                names = ", ".join(b.name for b in acquirer_data["atm"].top3)
                st.metric("Top ATM Acquirers", names)
            else:
                st.metric("Top ATM Acquirers", "No data")
        with summary_cols[1]:
            if "pos" in acquirer_data and acquirer_data["pos"].top3:
                names = ", ".join(b.name for b in acquirer_data["pos"].top3)
                st.metric("Top POS Acquirers", names)
            else:
                st.metric("Top POS Acquirers", "No data")
        with summary_cols[2]:
            if "ips" in acquirer_data and acquirer_data["ips"].top3_senders:
                names = ", ".join(b.name for b in acquirer_data["ips"].top3_senders)
                st.metric("Top IPS Senders", names)
            else:
                st.metric("Top IPS Senders", "No data")
        with summary_cols[3]:
            if "ips" in acquirer_data and acquirer_data["ips"].top3_receivers:
                names = ", ".join(b.name for b in acquirer_data["ips"].top3_receivers)
                st.metric("Top IPS Receivers", names)
            else:
                st.metric("Top IPS Receivers", "No data")

    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-sep"><span>Live Preview</span></div>', unsafe_allow_html=True)
    st.caption("The report updates live. Use the buttons inside the preview to print (save as PDF) "
               "or download as PNG. Icons and image export load from CDNs, so an internet connection is required.")
    components.html(
        build_report_html(report, calc, show_bars=show_bars, auto_highlight=auto_hl,
                          takeaway_override=takeaway_override,
                          acquirer_data=acquirer_data if acquirer_data else None),
        height=860 + (180 if acquirer_data else 0),
        scrolling=True,
    )


# ── Header ───────────────────────────────────────────────────────────────────
h1, h2 = st.columns([5, 1])
with h1:
    st.markdown('<p class="gradient-title">Report Merger</p>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Consolidate POS, ATM, IPS, QR &amp; P2P transaction reports &mdash; in memory, nothing saved to disk.</p>', unsafe_allow_html=True)
st.session_state.setdefault("theme", "dark")

with h2:
    st.markdown('<div class="theme-toggle">', unsafe_allow_html=True)
    st.toggle("Dark mode", key="theme_toggle")
    if st.session_state.theme_toggle is not None:
        st.session_state.theme = "dark" if st.session_state.theme_toggle else "light"
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ── Page Navigation ──────────────────────────────────────────────────────────
st.markdown(
    '<div class="page-nav">', unsafe_allow_html=True)
n1, n2 = st.columns([1, 1])
with n1:
    if st.button("Merged Reports", key="nav_merger", use_container_width=True):
        st.session_state.snap_page = False
with n2:
    if st.button("Transaction Snapshot", key="nav_snapshot", use_container_width=True):
        st.session_state.snap_page = True
st.markdown("</div>", unsafe_allow_html=True)

if st.session_state.snap_page:
    _render_snapshot_page()
    st.stop()

# ── Mode Selection (report-type cards) ──────────────────────────────────────
st.markdown('<div class="card"><div class="card-head"><div class="card-icon icon-purple">1</div><div><p class="card-title">Choose report type</p><p class="card-sub">Select the type of reports you want to merge</p></div></div>', unsafe_allow_html=True)

MODE_CARDS = [
    {"key": "pos_decline", "name": "POS Decline", "icon": "📉",
     "desc": "Declined POS transactions", "type": "type-red", "badge": "badge-red"},
    {"key": "pos_success", "name": "POS Success", "icon": "✅",
     "desc": "Successful POS transactions", "type": "type-green", "badge": "badge-green"},
    {"key": "pos", "name": "POS (Daily)", "icon": "💳",
     "desc": "SmartVista POS daily report", "type": "type-purple", "badge": "badge-purple"},
    {"key": "atm", "name": "ATM (Daily)", "icon": "🏧",
     "desc": "SmartVista ATM daily report", "type": "type-blue", "badge": "badge-blue"},
    {"key": "qr", "name": "QR", "icon": "🔗",
     "desc": "QR success bank summaries", "type": "type-cyan", "badge": "badge-blue"},
    {"key": "p2p", "name": "P2P", "icon": "↔️",
     "desc": "P2P success bank summaries", "type": "type-green", "badge": "badge-green"},
    {"key": "ips", "name": "IPS", "icon": "🔁",
     "desc": "Raw IPS transactions (all sheets)", "type": "type-amber", "badge": "badge-amber"},
    {"key": "sett_sum", "name": "Sett(Sum)", "icon": "🧮",
     "desc": "Settlement workbook summarised", "type": "type-violet", "badge": "badge-purple"},
]
mode_keys_map = {c["name"]: c["key"] for c in MODE_CARDS}
mode_colors = {c["name"]: c["badge"] for c in MODE_CARDS}

_active_key = st.session_state.mode_key
card_cols = None
for i, card in enumerate(MODE_CARDS):
    if i % 4 == 0:
        card_cols = st.columns(4)
    with card_cols[i % 4]:
        active = " mode-active" if _active_key == card["key"] else ""
        st.markdown(
            f'<div class="mode-card {card["type"]}{active}" data-mode="{card["key"]}">'
            f'<span class="mc-check">✅</span>'
            f'<span class="mc-icon">{card["icon"]}</span>'
            f'<div class="mc-name">{card["name"]}</div>'
            f'<div class="mc-desc">{card["desc"]}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        if st.button(f"Open {card['name']}", key=f"mode_{card['key']}", use_container_width=True):
            st.session_state.mode_key = card["key"]
            st.session_state.merged_meta = None
            st.session_state.filter_sheets = []
            st.session_state.records = []
            st.session_state.columns = []
            st.session_state.duplicate_records = []
            st.session_state.unique_values_cache = {}
            st.session_state.ips_all_records = []
            st.session_state.ips_dates = []
            st.session_state.ips_selected_dates = []
            gc.collect()

if st.session_state.mode_key is None:
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("Select a report type above to get started.")
    st.stop()

mode_key = st.session_state.mode_key
mode_label = [k for k, v in mode_keys_map.items() if v == mode_key][0]
mode_color = mode_colors.get(mode_label, "badge-blue")

# ── Sett(Sum) Report (sett_sum mode) ───────────────────────────────────────
if mode_key == "sett_sum":
    st.markdown(f'<span class="badge {mode_color}">SETT(SUM)</span>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="card"><div class="card-head"><div class="card-icon icon-blue">2</div>'
        '<div><p class="card-title">Upload settlement workbook</p>'
        '<p class="card-sub">Drag &amp; drop your .xls or .xlsx file. Banks that appear with the '
        'same name in the two sheets (ISS_BANKS and ACQ_BANKS) are merged into one row and every '
        'column is summed.</p></div></div>',
        unsafe_allow_html=True,
    )

    sett_file = st.file_uploader("Upload files", type=["xls", "xlsx"], label_visibility="collapsed", key="sett_file")

    if sett_file is not None:
        with st.spinner("Building Sett(Sum) report..."):
            try:
                sett_reports = generate_sett_sum_report(sett_file.getvalue())
            except ValueError as e:
                sett_reports = None
                st.error(str(e))

        # generate_sett_sum_report returns {label: DataFrame}; a stale
        # deployment may still hand back a single DataFrame, so normalise it.
        sett_report_items = None
        if isinstance(sett_reports, dict):
            sett_report_items = list(sett_reports.items())
        elif sett_reports is not None:
            sett_report_items = [("Sett(Sum)", sett_reports)]

        if sett_report_items:
            tabs = st.tabs([label for label, _ in sett_report_items])
            for tab, (label, sett_df) in zip(tabs, sett_report_items):
                with tab:
                    st.dataframe(sett_df, use_container_width=True, hide_index=True, height=360)
            if st.button("Download Sett(Sum) Report", use_container_width=True, key="dl_sett_btn"):
                with st.spinner("Building Sett(Sum) workbook..."):
                    sett_excel_bytes = build_sett_sum_report_excel(dict(sett_report_items))
                sett_filename = "Sett_Sum_Report.xlsx"
                st.download_button(
                    label="Click to save Sett(Sum) Report",
                    data=sett_excel_bytes,
                    file_name=sett_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_sett_actual",
                )
                del sett_excel_bytes
                gc.collect()

    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ── QR / P2P Successful Transaction Report (interop modes) ────────────────
if mode_key in ("qr", "p2p"):
    interop_label = "QR" if mode_key == "qr" else "P2P"
    st.markdown(f'<span class="badge {mode_color}">{interop_label} (Summary)</span>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(
        f'<div class="card"><div class="card-head"><div class="card-icon icon-blue">2</div>'
        f'<div><p class="card-title">Upload "{interop_label} success for source and destination" exports</p>'
        f'<p class="card-sub">Drag &amp; drop one or more .xlsx bank summary workbooks '
        f'(NO, BANK_ID, BANK_NAME, ISSUER_TXN_COUNT, ISSUER_TOTAL_AMOUNT, '
        f'ACQUIRER_TXN_COUNT, ACQUIRER_TOTAL_AMOUNT). Rows are merged per bank '
        f'and every numeric column is summed across the uploaded files.</p></div></div>',
        unsafe_allow_html=True,
    )

    interop_files = st.file_uploader(
        "Upload files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if not interop_files:
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    report_date = st.date_input("Report date", value=_date.today(), key="interop_report_date")

    if st.session_state.merged_meta is None or st.session_state.merged_meta.get("mode_key") not in ("qr", "p2p"):
        if st.button("Merge Reports", use_container_width=True):
            with st.spinner("Merging bank summaries..."):
                try:
                    payloads = [(f.name, f.getvalue()) for f in interop_files]
                    interop_records, interop_per_file, interop_warnings = merge_bank_summary_files(payloads, mode_key)
                except ValueError as e:
                    st.error(str(e))
                    st.stop()
                except Exception as e:
                    st.error(f"Unexpected error during merge: {e}")
                    st.stop()

            st.session_state.interop_records = interop_records
            st.session_state.merged_meta = {
                "mode_key": mode_key,
                "interop_per_file": interop_per_file,
                "interop_warnings": interop_warnings,
                "report_date": report_date,
            }
            st.session_state.filter_sheets = []
            st.session_state.pending_filters = {}
            gc.collect()
            st.rerun()

        st.stop()

    meta = st.session_state.merged_meta
    records = st.session_state.interop_records
    report_date = meta.get("report_date") or report_date

    bank_rows = [r for r in records if r.get("BANK_NAME") not in ("Total", "TOTAL")]
    total_row = records[-1]
    st.markdown(
        '<div class="card"><div class="card-head"><div class="card-icon icon-green">&#10003;</div>'
        '<div><p class="card-title">Merged Report</p>'
        '<p class="card-sub">Bank summaries consolidated successfully</p></div></div>',
        unsafe_allow_html=True,
    )
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Banks", len(bank_rows))
    c2.metric("Issuer Txn Count", f"{total_row['ISSUER_TXN_COUNT']:,}")
    c3.metric("Issuer Value (ETB)", f"{total_row['ISSUER_TOTAL_AMOUNT']:,.2f}")
    c4.metric("Acquirer Txn Count", f"{total_row['ACQUIRER_TXN_COUNT']:,}")
    c5.metric("Acquirer Value (ETB)", f"{total_row['ACQUIRER_TOTAL_AMOUNT']:,.2f}")

    if meta.get("interop_warnings"):
        with st.expander(f"Warnings ({len(meta['interop_warnings'])})", expanded=False):
            for w in meta["interop_warnings"]:
                st.warning(w)

    with st.expander("Files merged", expanded=False):
        file_df = pd.DataFrame(meta["interop_per_file"])
        st.dataframe(file_df, use_container_width=True, hide_index=True)

    st.markdown(
        '<div class="card-head"><div class="card-icon icon-blue">&#128269;</div>'
        '<div><p class="card-title">Preview</p>'
        '<p class="card-sub">All banks merged across the uploaded files</p></div></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(pd.DataFrame(records), use_container_width=True, hide_index=True, height=420)

    st.markdown('<div class="section-sep"><span>Download</span></div>', unsafe_allow_html=True)
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        if st.button("Download Successful Transaction Report", use_container_width=True, key="dl_interop_styled"):
            with st.spinner("Building styled report..."):
                styled_bytes = build_success_report_excel(records, mode_key, report_date)
            styled_filename = success_report_filename(mode_key, report_date)
            st.download_button(
                label="Click to save",
                data=styled_bytes,
                file_name=styled_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_interop_styled_actual",
            )
            del styled_bytes
            gc.collect()

    with col_dl2:
        if st.button("Download Merged Summary", use_container_width=True, key="dl_interop_merged"):
            with st.spinner("Building merged workbook..."):
                merged_bytes = build_merged_summary_excel(records, mode_key)
            merged_filename = f"Successful_{interop_label}_Summary_{report_date.strftime('%Y-%m-%d')}.xlsx"
            st.download_button(
                label="Click to save",
                data=merged_bytes,
                file_name=merged_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_interop_merged_actual",
            )
            del merged_bytes
            gc.collect()

    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ── IPS Transactions Report (ips mode) ──────────────────────────────────────
def _ips_selection_meta(selected_keys):
    """Build (records, merged_meta) for the currently selected IPS dates."""
    all_records = st.session_state.ips_all_records
    keys = sorted(selected_keys)
    records = filter_ips_by_dates(all_records, keys) if keys else []
    labels = {d["key"]: d["label"] for d in st.session_state.ips_dates}
    from_date = keys[0] if keys else "-"
    to_date = keys[-1] if keys else "-"
    from_lbl = labels.get(from_date, from_date)
    to_lbl = labels.get(to_date, to_date)
    if from_lbl == to_lbl:
        filename = f"IPS_Transactions_{from_lbl}_Merged.xlsx"
    else:
        filename = f"IPS_Transactions_{from_lbl}_to_{to_lbl}_Merged.xlsx"
    resp_counts = {}
    for r in records:
        status = str(r.get("STATUS", "") or "").strip() or "(blank)"
        resp_counts[status] = resp_counts.get(status, 0) + 1
    resp_counts = dict(sorted(resp_counts.items(), key=lambda kv: -kv[1]))
    meta = {
        "filename": filename,
        "from_date": from_date,
        "to_date": to_date,
        "total_rows": len(records),
        "duplicate_count": 0,
        "per_file": st.session_state.get("ips_per_file", []),
        "resp_counts": resp_counts,
        "warnings": [],
        "mode_key": "ips",
        "mode_label": "IPS",
        "sort_by": "date_time",
        "sort_dir": "asc",
    }
    return records, meta


if mode_key == "ips":
    st.markdown(f'<span class="badge {mode_color}">IPS</span>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    ips_ready = (
        st.session_state.merged_meta is not None
        and st.session_state.merged_meta.get("mode_key") == "ips"
    )

    if not ips_ready:
        st.markdown(
            '<div class="card"><div class="card-head"><div class="card-icon icon-amber">2</div>'
            '<div><p class="card-title">Upload IPS transaction exports</p>'
            '<p class="card-sub">Drag &amp; drop one or more raw IPS .xls / .xlsx exports. Every '
            'sheet is scanned for transaction dates, then you pick which dates to include.</p>'
            '</div></div>',
            unsafe_allow_html=True,
        )
        ips_files = st.file_uploader(
            "Upload files", type=["xls", "xlsx"],
            accept_multiple_files=True, label_visibility="collapsed", key="ips_files",
        )
        if not ips_files:
            st.markdown('</div>', unsafe_allow_html=True)
            st.stop()

        for f in ips_files:
            size_mb = f.size / (1024 * 1024)
            st.markdown(
                f'<div class="sheet-item"><span class="sheet-name">{f.name}</span>'
                f'<span class="sheet-rows">{size_mb:.1f} MB</span></div>',
                unsafe_allow_html=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)

        total_ips_mb = sum(f.size for f in ips_files) / (1024 * 1024)
        if total_ips_mb > 800:
            st.error(f"The combined IPS files are {total_ips_mb:.1f} MB – exceeds the 800 MB limit.")
            st.stop()
        elif total_ips_mb > 600:
            st.warning(
                f"IPS files are {total_ips_mb:.1f} MB. The in‑memory parser may take a few seconds "
                f"and could appear to ‘crash’. Please wait for the progress bar."
            )

        if st.button("Analyse Reports & Pick Dates", use_container_width=True, key="ips_analyze"):
            with st.spinner("Scanning every sheet for transaction dates..."):
                try:
                    parsed = [parse_ips_report(f.getvalue(), f.name) for f in ips_files]
                except MemoryError:
                    st.error("Not enough memory to process these files. Try a smaller file or fewer files at once.")
                    st.stop()
                except ValueError as e:
                    st.error(str(e))
                    st.stop()
                except Exception as e:  # noqa: BLE001
                    st.error(f"Unexpected error during analysis: {e}")
                    st.stop()

            all_records = [r for p in parsed for r in p["records"]]
            dates = collect_ips_dates(parsed)
            warnings = [f"{p['filename']}: {w}" for p in parsed for w in p["warnings"]]
            per_file = [
                {
                    "filename": p["filename"], "status": "ok", "sheet": "all sheets",
                    "raw_rows": p["rows_scanned"], "data_rows": len(p["records"]),
                    "columns_kept": len(p["columns"]),
                }
                for p in parsed
            ]

            if not all_records:
                for w in warnings:
                    st.warning(w)
                st.error("No IPS transaction rows were found in the uploaded file(s).")
                st.stop()

            selected_keys = [d["key"] for d in dates]
            st.session_state.ips_all_records = all_records
            st.session_state.ips_dates = dates
            st.session_state.ips_per_file = per_file
            st.session_state.ips_warnings = warnings
            st.session_state.ips_selected_dates = selected_keys
            st.session_state["ips_date_multiselect"] = selected_keys
            st.session_state.columns = list(all_records[0].keys())
            st.session_state.filter_sheets = []
            st.session_state.pending_filters = {}
            st.session_state.duplicate_records = []
            st.session_state.unique_values_cache = {}
            st.session_state.records, st.session_state.merged_meta = _ips_selection_meta(selected_keys)
            del parsed, all_records
            gc.collect()
            st.rerun()

        st.stop()

    # Metadata ready -> show the date picker, then fall through to Results.
    st.markdown(
        '<div class="card"><div class="card-head"><div class="card-icon icon-amber">3</div>'
        '<div><p class="card-title">Choose transaction dates</p>'
        '<p class="card-sub">Every distinct date found across <b>all sheets</b> of the uploaded '
        'file(s). Pick one or more to include in the report.</p></div></div>',
        unsafe_allow_html=True,
    )

    ips_dates = st.session_state.ips_dates
    key_to_label = {d["key"]: d["label"] for d in ips_dates}
    key_to_count = {d["key"]: d["count"] for d in ips_dates}

    if "ips_date_multiselect" not in st.session_state:
        st.session_state["ips_date_multiselect"] = [d["key"] for d in ips_dates]

    selected_dates = st.multiselect(
        "Transaction dates",
        options=[d["key"] for d in ips_dates],
        format_func=lambda k: f"{key_to_label.get(k, k)}  ·  {key_to_count.get(k, 0):,} txns",
        key="ips_date_multiselect",
        help="Dates are read from the TRX_DATE column on every sheet of the workbook.",
    )

    if selected_dates:
        pills = "".join(
            f'<span class="date-pill">{key_to_label.get(k, k)} · {key_to_count.get(k, 0):,}</span>'
            for k in selected_dates
        )
    else:
        pills = ('<span class="date-pill" style="background:rgba(255,71,87,0.14);'
                 'color:#ff8a8a;border-color:rgba(255,71,87,0.3);">No dates selected</span>')
    st.markdown(f'<div style="margin:6px 0 10px;">{pills}</div>', unsafe_allow_html=True)

    if selected_dates != st.session_state.ips_selected_dates:
        st.session_state.ips_selected_dates = list(selected_dates)
        st.session_state.records, st.session_state.merged_meta = _ips_selection_meta(list(selected_dates))
        st.session_state.unique_values_cache = {}
        st.session_state.filter_sheets = []
        st.session_state.pending_filters = {}
        gc.collect()
        st.rerun()

    if st.session_state.ips_warnings:
        with st.expander(f"Warnings ({len(st.session_state.ips_warnings)})", expanded=False):
            for w in st.session_state.ips_warnings:
                st.warning(w)

    st.markdown('</div>', unsafe_allow_html=True)
    # Falls through to the shared Results / Filter & Export sections below.


if mode_key != "ips":
    mode = MODES[mode_key]

    st.markdown(f'<span class="badge {mode_color}">{mode.label}</span>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── File Upload ──────────────────────────────────────────────────────────
    st.markdown(f'<div class="card"><div class="card-head"><div class="card-icon icon-blue">2</div><div><p class="card-title">Upload {mode.label} reports</p><p class="card-sub">Drag & drop your .xls or .xlsx files</p></div></div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Upload files",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if not uploaded_files:
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    for f in uploaded_files:
        size_kb = f.size / 1024
        label = f"{size_kb:.0f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
        st.markdown(f'<div class="sheet-item"><div class="sheet-num">{len(uploaded_files)}</div><span class="sheet-name">{f.name}</span><span class="sheet-rows">{label}</span></div>', unsafe_allow_html=True)

    # ── Upload‑progress widget ─────────────────────────────────────────────
    prog_placeholder = st.empty()
    total_bytes = sum(f.size for f in uploaded_files)
    start_time = time.time()
    avg_speed_mbps = 10.0  # MB/s rough estimate
    elapsed = time.time() - start_time
    percent = min(100, (elapsed * avg_speed_mbps / (total_bytes / 1_000_000)) * 100) if total_bytes else 0
    seconds_left = (total_bytes / (avg_speed_mbps * 1_000_000) - elapsed) if total_bytes else 0
    prog_placeholder.progress(max(0, min(100, percent)))
    prog_placeholder.caption(
        f"Upload progress: {percent:5.1f}% • estimated {seconds_left:6.1f}s left"
    )

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Merge ────────────────────────────────────────────────────────────────
    if st.session_state.merged_meta is None:
        total_size_mb = sum(f.size for f in uploaded_files) / (1024 * 1024)
        if total_size_mb > 700:
            st.warning(f"Large upload ({total_size_mb:.0f} MB total). Processing may take a while and could hit memory limits on free hosting.")

        # Sort options - dates written with spelled months (Aug, Jul...) and
        # numbers (20260813) are handled automatically by the smart sort key.
        sort_by = st.selectbox(
            "Sort by",
            options=["date_time"] + list(mode.canonical_columns),
            format_func=lambda c: "Date \u00b7 Time (default)" if c == "date_time" else c,
            help="Choose which column to order the merged report by. Dates written with "
                 "spelled-out months (e.g. 15-Aug-26) and numbers (20260813) are sorted together.",
        )
        sort_dir = st.radio("Direction", ["asc", "desc"], horizontal=True,
                            format_func=lambda d: "Ascending" if d == "asc" else "Descending")

        dedupe = st.toggle(
            "Remove duplicate rows",
            value=False,
            help="Keep only the first occurrence of each identical row and remove the rest. "
                 "ACQUIRER, ISSUER, TRANS_TYPE and CURRENCY are excluded from the comparison "
                 "in every report mode (POS Decline, POS Success, POS Daily, ATM) — "
                 "so two rows are duplicates when every other column (card/account number, "
                 "date, time, amount, response code, reference numbers, terminal, address) matches, "
                 "regardless of which acquirer or issuer processed them. "
                 "The 'Download Duplicates' button always shows all rows involved, even when this is off.",
        )

        if st.button("Merge Reports", use_container_width=True):
            with st.spinner("Merging reports..."):
                try:
                    payloads = [(f.name, f.getvalue()) for f in uploaded_files]
                    result = merge_reports(
                        payloads,
                        mode_key=mode_key,
                        skip_workbook=True,
                        sort_by=sort_by,
                        sort_dir=sort_dir,
                        dedupe=dedupe,
                    )
                except MemoryError:
                    st.error("Not enough memory to process these files. Try uploading smaller files or fewer at a time.")
                    st.stop()
                except ValueError as e:
                    st.error(str(e))
                    st.stop()
                except Exception as e:
                    st.error(f"Unexpected error during merge: {e}")
                    st.stop()

            st.session_state.records = result.records
            st.session_state.columns = list(result.records[0].keys()) if result.records else []
            st.session_state.filter_sheets = []
            st.session_state.pending_filters = {}

            st.session_state.merged_meta = {
                "filename": result.filename,
                "from_date": result.from_date,
                "to_date": result.to_date,
                "total_rows": result.total_rows,
                "duplicate_count": len(result.duplicate_records),
                "per_file": result.per_file,
                "resp_counts": result.resp_counts,
                "warnings": result.warnings,
                "mode_key": result.mode_key,
                "mode_label": result.mode_label,
                "sort_by": result.sort_by,
                "sort_dir": result.sort_dir,
            }
            st.session_state.duplicate_records = result.duplicate_records
            del result, payloads
            gc.collect()
            st.rerun()

        st.stop()

# ── Results ──────────────────────────────────────────────────────────────────
meta = st.session_state.merged_meta

st.markdown('<div class="card"><div class="card-head"><div class="card-icon icon-green">&#10003;</div><div><p class="card-title">Merged Report</p><p class="card-sub">All reports consolidated successfully</p></div></div>', unsafe_allow_html=True)

ok_files = [p for p in meta["per_file"] if p["status"] == "ok"]
dup_count = meta.get("duplicate_count", 0)
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Transactions", f"{meta['total_rows']:,}")
if meta["from_date"] == meta["to_date"]:
    date_str = meta["from_date"]
else:
    date_str = f"{meta['from_date']} \u2192 {meta['to_date']}"
c2.metric("Date Range", date_str)
c3.metric("Files Merged", len(ok_files))
c4.metric("Response Codes", len(meta["resp_counts"]))
c5.metric("Duplicates Found", f"{dup_count:,}")

_sort_by = meta.get("sort_by") or "date_time"
if _sort_by == "date_time":
    sort_desc = "Date \u00b7 Time (default)"
else:
    sort_desc = _sort_by
_sort_dir = meta.get("sort_dir") or "asc"
st.caption(
    f"Sorted by: **{sort_desc}** ({'descending' if _sort_dir == 'desc' else 'ascending'})"
)

if meta["warnings"]:
    with st.expander(f"Warnings ({len(meta['warnings'])})", expanded=False):
        for w in meta["warnings"]:
            st.warning(w)

with st.expander("Files merged", expanded=False):
    file_df = pd.DataFrame(meta["per_file"])
    st.dataframe(file_df, use_container_width=True, hide_index=True)

if meta["resp_counts"]:
    with st.expander("Response code distribution", expanded=False):
        resp_df = pd.DataFrame(
            list(meta["resp_counts"].items()),
            columns=["Response Code", "Count"]
        ).sort_values("Count", ascending=False).head(15)
        st.bar_chart(resp_df.set_index("Response Code"))

st.markdown(f'<div class="card-head"><div class="card-icon icon-blue">&#128269;</div><div><p class="card-title">Preview</p><p class="card-sub">First {min(len(st.session_state.records), 50)} of {meta["total_rows"]:,} rows</p></div></div>', unsafe_allow_html=True)
preview_rows = []
for row in st.session_state.records[:50]:
    preview_rows.append({k: str(v) if v is not None and v != "" else "\u2014" for k, v in row.items()})
preview_df = pd.DataFrame(preview_rows)
st.dataframe(preview_df, use_container_width=True, hide_index=True, height=380)

# Download merged — generate workbook on demand (NOT stored in session state)
st.markdown('<div class="section-sep"><span>Download</span></div>', unsafe_allow_html=True)
col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    if st.button("Download Merged Report", use_container_width=True, key="dl_merged"):
        with st.spinner("Building workbook..."):
            wb_bytes = build_workbook(
                st.session_state.records,
                meta["from_date"],
                meta["to_date"],
                MODES[meta["mode_key"]],
            )
        st.download_button(
            label="Click to save",
            data=wb_bytes,
            file_name=meta["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_merged_actual",
        )
        del wb_bytes
        gc.collect()

with col_dl2:
    dup_count = meta.get("duplicate_count", 0)
    if dup_count > 0:
        if st.button(f"Download Duplicates Only ({dup_count:,} rows)", use_container_width=True, key="dl_dupes"):
            with st.spinner("Building duplicates workbook..."):
                dup_wb = build_workbook(
                    st.session_state.duplicate_records,
                    "",
                    "",
                    MODES[meta["mode_key"]],
                )
            dup_filename = meta["filename"].replace("_Merged.xlsx", "_Duplicates.xlsx")
            st.download_button(
                label="Click to save duplicates",
                data=dup_wb,
                file_name=dup_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_dupes_actual",
            )
            del dup_wb
            gc.collect()
    else:
        st.button("No Duplicates to Download", use_container_width=True, disabled=True, key="dl_dupes_disabled")

st.markdown('</div>', unsafe_allow_html=True)

# ── NBE Institution Summary Report (POS, ATM & POS Decline) ───────────────
if meta["mode_key"] in ("pos", "atm", "pos_decline"):
    st.markdown('<div class="section-sep"><span>NBE Institution Report</span></div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="card"><div class="card-head"><div class="card-icon icon-purple">&#127974;</div>'
        '<div><p class="card-title">NBE Institution Breakdown Report</p>'
        '<p class="card-sub">Transaction counts and total value (ETB) per institution as Issuer &amp; Acquirer '
        f'(Filtered for {meta["mode_key"].upper()} '
        f'{"declined" if meta["mode_key"] == "pos_decline" else "successful"} transactions'
        f'{"" if meta["mode_key"] == "pos_decline" else " with Response Code -1 / -1.0"})</p></div></div>',
        unsafe_allow_html=True,
    )

    nbe_mode = "pos_decline" if meta["mode_key"] == "pos_decline" else meta["mode_key"]
    nbe_df = generate_nbe_report(st.session_state.records, nbe_mode)

    st.dataframe(
        nbe_df,
        use_container_width=True,
        hide_index=True,
        height=360,
    )

    col_nbe_dl1, col_nbe_dl2 = st.columns(2)
    with col_nbe_dl1:
        if st.button(f"Download {nbe_mode.upper()} NBE Institution Report", use_container_width=True, key="dl_nbe_btn"):
            with st.spinner("Building NBE Report workbook..."):
                nbe_excel_bytes = build_nbe_report_excel(nbe_df, nbe_mode)
            nbe_filename = f"NBE_{nbe_mode.upper()}_Report_{meta['from_date']}_to_{meta['to_date']}.xlsx"
            st.download_button(
                label="Click to save NBE Report",
                data=nbe_excel_bytes,
                file_name=nbe_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_nbe_actual",
            )
            del nbe_excel_bytes
            gc.collect()

    st.markdown('</div>', unsafe_allow_html=True)

# ── Balance Inquiry NBE Reports (POS / ATM daily modes) ────────────────────
if meta["mode_key"] in ("pos", "atm"):
    st.markdown('<div class="section-sep"><span>Balance Inquiry Report</span></div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="card"><div class="card-head"><div class="card-icon icon-purple">&#129534;</div>'
        '<div><p class="card-title">NBE Balance Inquiry Breakdown Report</p>'
        '<p class="card-sub">Balance inquiry transactions split by result '
        '(balance inquiries have no monetary amount)</p></div></div>',
        unsafe_allow_html=True,
    )

    bi_meta = [
        ("balance_inquiry_success", "Success Balance Inquiry",
         "Successful balance inquiries with Response Code -1 / -1.0",
         "BALANCE_INQUIRY_SUCCESS_Report"),
        ("balance_inquiry_decline", "Decline Balance Inquiry",
         "Declined balance inquiries excluding response codes -1, 503, 821, 862, 901, 904, 911, 912, 915",
         "BALANCE_INQUIRY_DECLINE_Report"),
    ]
    for mode_key, title, subtitle, filename_prefix in bi_meta:
        st.markdown(
            f'<div class="card"><div class="card-head"><div class="card-icon icon-purple">{title[:1]}</div>'
            f'<div><p class="card-title">{title}</p>'
            f'<p class="card-sub">{subtitle}</p></div></div>',
            unsafe_allow_html=True,
        )

        bi_df = generate_nbe_report(st.session_state.records, mode_key)

        st.dataframe(
            bi_df,
            use_container_width=True,
            hide_index=True,
            height=360,
        )

        dl_key1 = f"dl_{mode_key}_btn"
        dl_key2 = f"dl_{mode_key}_actual"
        if st.button(f"Download {title}", use_container_width=True, key=dl_key1):
            with st.spinner(f"Building {title} workbook..."):
                bi_excel_bytes = build_nbe_report_excel(bi_df, mode_key)
            bi_filename = f"{filename_prefix}_{meta['from_date']}_to_{meta['to_date']}.xlsx"
            st.download_button(
                label=f"Click to save {title} Report",
                data=bi_excel_bytes,
                file_name=bi_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=dl_key2,
            )
            del bi_excel_bytes
            gc.collect()

    st.markdown('</div>', unsafe_allow_html=True)

# ── ATM Decline Response Codes Report (ATM daily mode only) ─────────────────
if meta["mode_key"] == "atm":
    st.markdown('<div class="section-sep"><span>ATM Decline Response Codes</span></div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="card"><div class="card-head"><div class="card-icon icon-pink">&#9888;</div>'
        '<div><p class="card-title">ATM Decline Response Codes Breakdown Report</p>'
        '<p class="card-sub">Count and value (ETB) of ATM transactions per institution as Issuer &amp; Acquirer, '
        'excluding response codes -1, 503, 821, 862, 901, 904, 911, 912, 915</p></div></div>',
        unsafe_allow_html=True,
    )

    adc_df = generate_nbe_report(st.session_state.records, "atm_decline")

    st.dataframe(
        adc_df,
        use_container_width=True,
        hide_index=True,
        height=360,
    )

    col_adc_dl1, col_adc_dl2 = st.columns(2)
    with col_adc_dl1:
        if st.button("Download ATM Decline Response Codes Report", use_container_width=True, key="dl_adc_btn"):
            with st.spinner("Building ATM Decline Response Codes workbook..."):
                adc_excel_bytes = build_nbe_report_excel(adc_df, "atm_decline")
            adc_filename = f"ATM_DECLINE_RESPONSE_CODES_Report_{meta['from_date']}_to_{meta['to_date']}.xlsx"
            st.download_button(
                label="Click to save ATM Decline Response Codes Report",
                data=adc_excel_bytes,
                file_name=adc_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_adc_actual",
            )
            del adc_excel_bytes
            gc.collect()

    st.markdown('</div>', unsafe_allow_html=True)

# ── POS Success Rate Summary Report (POS modes only) ──────────────────────────
if meta["mode_key"] in ("pos", "pos_decline", "pos_success"):
    st.markdown('<div class="section-sep"><span>POS Success Rate Report</span></div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="card"><div class="card-head"><div class="card-icon icon-green">&#128200;</div>'
        '<div><p class="card-title">POS Transaction Decline &amp; Success Rate Summary (Issuer View)</p>'
        '<p class="card-sub">Response code matrix grouped strictly by ISSUER institutions, including Cardholder-Related Declines '
        '(RC 821, 901, 904, 906, 911, 912, 914, 915), Total Success, Total Attempted Transactions &amp; Adjusted Success Rate %</p></div></div>',
        unsafe_allow_html=True,
    )

    matrix_df, desc_df = generate_pos_success_rate_report(st.session_state.records)

    # Format success rate row values as percentages for UI preview
    display_matrix = matrix_df.copy()
    rate_idx = display_matrix[display_matrix["RC/BANK NAME"] == "success rate"].index
    if not rate_idx.empty:
        r_i = rate_idx[0]
        for col_name in display_matrix.columns:
            if col_name != "RC/BANK NAME":
                val = display_matrix.at[r_i, col_name]
                try:
                    fval = float(val)
                    display_matrix.at[r_i, col_name] = f"{fval * 100:.2f}%"
                except (ValueError, TypeError):
                    pass

    st.dataframe(
        display_matrix,
        use_container_width=True,
        hide_index=True,
        height=380,
    )

    st.markdown(
        '<div style="font-size:0.78rem;color:#a0aec0;margin-bottom:12px;">'
        '<b>Success Rate Color Fills:</b> '
        '<span style="background:#C6EFCE;color:#006100;padding:2px 8px;border-radius:4px;font-weight:700;">🟩 97%–100% Green</span> &nbsp;'
        '<span style="background:#FFEB9C;color:#9C6500;padding:2px 8px;border-radius:4px;font-weight:700;">🟨 86%–96% Yellow</span> &nbsp;'
        '<span style="background:#FFF2CC;color:#7F6000;padding:2px 8px;border-radius:4px;font-weight:700;">🟧 79%–85% L. Yellow</span> &nbsp;'
        '<span style="background:#FFC7CE;color:#9C0006;padding:2px 8px;border-radius:4px;font-weight:700;">🟥 &le;78% Red</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    with st.expander("Response Code Descriptions & Remarks Lookup Table", expanded=False):
        st.dataframe(desc_df, use_container_width=True, hide_index=True)

    col_sr_dl1, col_sr_dl2 = st.columns(2)
    with col_sr_dl1:
        if st.button("Download POS Success Rate Report", use_container_width=True, key="dl_sr_btn"):
            with st.spinner("Building POS Success Rate workbook..."):
                sr_excel_bytes = build_pos_success_rate_excel(matrix_df, desc_df, report_date=meta['from_date'])
            sr_filename = f"POS_Success_Rate_Report_Issuer_{meta['from_date']}_to_{meta['to_date']}.xlsx"
            st.download_button(
                label="Click to save POS Success Rate Report",
                data=sr_excel_bytes,
                file_name=sr_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_sr_actual",
            )
            del sr_excel_bytes
            gc.collect()

    st.markdown('</div>', unsafe_allow_html=True)

# ── Filter Panel ─────────────────────────────────────────────────────────────
st.markdown('<div class="card"><div class="card-head"><div class="card-icon icon-pink">&#9660;</div><div><p class="card-title">Filter & Export to Sheets</p><p class="card-sub">Build filters column by column, then create sheets. Each sheet becomes a separate tab in the downloaded Excel file.</p></div></div>', unsafe_allow_html=True)

if "pending_filters" not in st.session_state:
    st.session_state.pending_filters = {}

MAX_FILTER_VALUES = 500

def _get_col_values(col):
    """Compute unique values for a column lazily and cache."""
    if col in st.session_state.unique_values_cache:
        return st.session_state.unique_values_cache[col]
    raw = {
        str(r.get(col, "")).strip()
        for r in st.session_state.records
        if r.get(col, "") not in ("", None)
    }
    vals = sorted(raw)[:MAX_FILTER_VALUES]
    st.session_state.unique_values_cache[col] = vals
    return vals

def _count_filtered_records(records, filters):
    """Count records matching all filters (AND logic)."""
    result = records
    for col, fvs in filters.items():
        allowed = {str(v).strip().upper() for v in fvs}
        result = [r for r in result if str(r.get(col, "")).strip().upper() in allowed]
    return len(result)

col_a, col_b, col_c = st.columns([2, 4, 1])

with col_a:
    filter_col = st.selectbox(
        "Column",
        options=[""] + st.session_state.columns,
        key="filter_col_select",
        index=0,
        placeholder="Select column...",
    )

with col_b:
    if filter_col:
        all_vals = _get_col_values(filter_col)
        filter_vals = st.multiselect(
            "Values (pick one or more)",
            options=all_vals,
            key="filter_vals_multi",
            placeholder=f"Choose {filter_col} values...",
        )
    else:
        filter_vals = []
        st.multiselect(
            "Values",
            options=["(Select a column first)"],
            key="filter_vals_multi",
            disabled=True,
            default=[],
        )

with col_c:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    add_clicked = st.button(
        "+ Add Filter",
        key="add_filter_btn",
        disabled=(not filter_col or not filter_vals),
        use_container_width=True,
    )

if add_clicked and filter_col and filter_vals:
    st.session_state.pending_filters[filter_col] = list(filter_vals)
    st.rerun()

if st.session_state.pending_filters:
    st.markdown('<p style="color:#5a6a9a;font-size:0.75rem;margin:8px 0 4px;font-weight:600;">Current sheet filters (AND between columns, OR within column):</p>', unsafe_allow_html=True)
    for pcol, pvals in list(st.session_state.pending_filters.items()):
        vals_display = ", ".join(str(v) for v in pvals[:5])
        if len(pvals) > 5:
            vals_display += f" +{len(pvals)-5} more"
        pcol1, pcol2 = st.columns([5, 1])
        with pcol1:
            st.markdown(
                f'<div class="sheet-item">'
                f'<span class="sheet-name">{pcol}</span>'
                f'<span class="sheet-desc" style="margin-left:8px;">{vals_display}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        with pcol2:
            if st.button("X", key=f"rm_pending_{pcol}", use_container_width=True):
                del st.session_state.pending_filters[pcol]
                st.rerun()

    cc1, cc2 = st.columns([1, 1])
    with cc1:
        if st.button("Create Sheet", key="create_sheet_btn", use_container_width=True):
            import copy
            filters = copy.deepcopy(st.session_state.pending_filters)
            count = _count_filtered_records(st.session_state.records, filters)
            first_col = list(filters.keys())[0]
            first_vals = list(filters[first_col])
            val_label = "+".join(str(v)[:10] for v in first_vals[:3])
            if len(first_vals) > 3:
                val_label += "+..."
            sheet_name = f"{first_col}_{val_label}"
            if len(filters) > 1:
                sheet_name += f"+{len(filters)-1}col"
            st.session_state.filter_sheets.append({
                "name": sheet_name, "filters": filters, "count": count,
            })
            st.session_state.pending_filters = {}
            st.rerun()
    with cc2:
        if st.button("Clear", key="clear_pending_btn", use_container_width=True):
            st.session_state.pending_filters = {}
            st.rerun()
else:
    st.markdown('<p style="color:#3a4570;font-size:0.8rem;font-style:italic;margin:8px 0 0;">Select a column + values, click "+ Add Filter". Add more columns for AND logic. Then "Create Sheet".</p>', unsafe_allow_html=True)

if st.session_state.filter_sheets:
    st.markdown('<div class="section-sep"><span>Sheets</span></div>', unsafe_allow_html=True)
    for i, sheet in enumerate(st.session_state.filter_sheets):
        parts = []
        for k, v in sheet["filters"].items():
            if isinstance(v, (list, tuple, set)):
                vs = list(v)
                if len(vs) <= 3:
                    parts.append(f"<code>{k}</code> IN ({', '.join(f'<code>{vv}</code>' for vv in vs)})")
                else:
                    parts.append(f"<code>{k}</code> IN ({', '.join(f'<code>{vv}</code>' for vv in vs[:3])}, +{len(vs)-3})")
            else:
                parts.append(f"<code>{k}</code> = <code>{v}</code>")
        desc = " AND ".join(parts) if parts else "All rows (unfiltered)"
        st.markdown(
            f'<div class="sheet-item">'
            f'<div class="sheet-num">{i+1}</div>'
            f'<span class="sheet-name">{sheet["name"]}</span>'
            f'<span class="sheet-desc">{desc}</span>'
            f'<span class="sheet-rows">~{sheet["count"]:,} rows</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

    rm_cols = st.columns(len(st.session_state.filter_sheets))
    for i, col in enumerate(rm_cols):
        with col:
            if st.button("Remove", key=f"rm_{i}", use_container_width=True):
                st.session_state.filter_sheets.pop(i)
                st.rerun()

    st.markdown('<div class="section-sep"><span>Download Filtered</span></div>', unsafe_allow_html=True)

    total_records = sum(s["count"] for s in st.session_state.filter_sheets)
    sheet_count = len(st.session_state.filter_sheets)

    if st.button(
        f"Download Filtered Report ({sheet_count} sheets, {total_records:,} rows)",
        use_container_width=True,
        key=f"dl_filtered_{sheet_count}_{total_records}",
    ):
        with st.spinner("Building filtered workbook..."):
            mode_obj = MODES[st.session_state.mode_key]
            filtered_wb = build_filtered_workbook(
                st.session_state.records,
                st.session_state.columns,
                st.session_state.filter_sheets,
                mode_obj,
            )
        if sheet_count == 1:
            dl_name = f"{mode_obj.output_prefix}_{st.session_state.filter_sheets[0]['name']}_Filtered.xlsx"
        else:
            dl_name = f"{mode_obj.output_prefix}_Filtered_{sheet_count}_Sheets.xlsx"

        st.download_button(
            label="Click to save",
            data=filtered_wb,
            file_name=dl_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_filtered_actual_{sheet_count}",
        )
        del filtered_wb
        gc.collect()

    col_clr1, col_clr2 = st.columns([4, 1])
    with col_clr2:
        if st.button("Clear All", key="clear_sheets"):
            st.session_state.filter_sheets = []
            st.session_state.pending_filters = {}
            st.rerun()
else:
    st.info("Build filters above, then create sheets. Each sheet becomes a separate tab in the downloaded Excel file.")

st.markdown('</div>', unsafe_allow_html=True)

# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown('<div style="text-align:center;color:#2a3560;font-size:0.75rem;padding:16px 0 0;">Reports are processed entirely in memory &mdash; nothing is saved to disk.</div>', unsafe_allow_html=True)
