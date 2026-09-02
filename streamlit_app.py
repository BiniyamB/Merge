"""Report Merger + Digital Transaction Value Snapshot -- Streamlit version."""

import gc
import html as html_lib
import io

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd

st.set_page_config(
    page_title="Report Merger (POS, ATM & QR)",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

from merger import MODES, merge_reports, build_filtered_workbook, build_workbook

from snapshot_module import REPORT_DEFAULTS, SERVICE_DEFAULTS, calc_all, build_report_html, fmt_int

# ── Global CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

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
        background: linear-gradient(160deg, #050a18 0%, #0a1128 40%, #0d1a30 100%);
        font-family: 'Plus Jakarta Sans', sans-serif; }
    .block-container { max-width: 1100px; padding-top: 1.5rem; padding-bottom: 2rem; }

    /* Gradient text */
    .gradient-title {
        font-size: 2.2rem; font-weight: 800; letter-spacing: -1px;
        background: linear-gradient(135deg, #ff3cac 0%, #784ba0 40%, #2b86c5 80%, #00d4ff 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        margin-bottom: 0; }
    .subtitle { color: #5a6a9a; font-size: 0.95rem; margin-top: -4px; }

    /* Cards */
    .card {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.06);
        border-radius: 16px; padding: 24px 28px; margin-bottom: 20px;
        backdrop-filter: blur(12px); }
    .card-head {
        display: flex; align-items: center; gap: 12px; margin-bottom: 18px; }
    .card-icon {
        width: 38px; height: 38px; border-radius: 10px; display: flex;
        align-items: center; justify-content: center; font-size: 1.1rem;
        flex-shrink: 0; }
    .icon-blue { background: rgba(0,212,255,0.1); border: 1px solid rgba(0,212,255,0.15); }
    .icon-purple { background: rgba(168,85,247,0.1); border: 1px solid rgba(168,85,247,0.15); }
    .icon-green { background: rgba(0,232,143,0.1); border: 1px solid rgba(0,232,143,0.15); }
    .icon-pink { background: rgba(255,60,172,0.1); border: 1px solid rgba(255,60,172,0.15); }
    .card-title { font-size: 1.05rem; font-weight: 700; color: #f0f4ff; margin: 0; }
    .card-sub { font-size: 0.8rem; color: #5a6a9a; margin: 2px 0 0; }

    /* Metrics */
    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.06);
        border-radius: 14px; padding: 16px 18px;
        transition: transform 0.2s, border-color 0.2s; }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        border-color: rgba(168,85,247,0.2); }
    div[data-testid="stMetric"] label {
        color: #5a6a9a !important; font-size: 0.72rem !important;
        text-transform: uppercase; letter-spacing: 0.8px; font-weight: 700 !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #00d4ff !important; font-size: 1.6rem !important; font-weight: 800 !important; }

    /* Radio mode selector */
    div[data-baseweb="radio"] > label {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.06);
        border-radius: 12px; padding: 14px 22px; margin: 0;
        cursor: pointer; transition: all 0.25s;
        display: flex; align-items: center; gap: 8px; }
    div[data-baseweb="radio"] > label:hover {
        border-color: rgba(168,85,247,0.3);
        background: rgba(168,85,247,0.04); }
    div[data-baseweb="radio"] > label[data-checked="true"] {
        border-color: rgba(168,85,247,0.4);
        background: rgba(168,85,247,0.08); }
    div[data-baseweb="radio"] > label[data-checked="true"] > div {
        color: #c4b5fd !important; font-weight: 700 !important; }

    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #ff3cac 0%, #784ba0 50%, #2b86c5 100%) !important;
        color: white !important; border: none !important; border-radius: 12px !important;
        font-weight: 700 !important; font-size: 0.92rem !important;
        padding: 0.6rem 2.5rem !important;
        box-shadow: 0 4px 20px rgba(255,60,172,0.2) !important;
        transition: all 0.3s !important; font-family: 'Plus Jakarta Sans', sans-serif !important; }
    .stButton > button:hover {
        box-shadow: 0 6px 28px rgba(255,60,172,0.35) !important;
        transform: translateY(-2px) !important; }
    .stButton > button:active { transform: translateY(0) !important; }

    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #00e88f 0%, #00b4d8 100%) !important;
        color: #052e16 !important; border: none !important; border-radius: 12px !important;
        font-weight: 700 !important; font-size: 0.92rem !important;
        padding: 0.6rem 2.5rem !important;
        box-shadow: 0 4px 20px rgba(0,232,143,0.2) !important;
        transition: all 0.3s !important; font-family: 'Plus Jakarta Sans', sans-serif !important; }
    div[data-testid="stDownloadButton"] > button:hover {
        box-shadow: 0 6px 28px rgba(0,232,143,0.35) !important;
        transform: translateY(-2px) !important; }

    /* Filter download button */
    .filter-dl > button {
        background: linear-gradient(135deg, #a855f7 0%, #6366f1 50%, #3b82f6 100%) !important;
        color: white !important;
        box-shadow: 0 4px 20px rgba(168,85,247,0.2) !important; }

    /* Selectbox */
    div[data-baseweb="select"] > div {
        background: rgba(255,255,255,0.04) !important;
        border-color: rgba(255,255,255,0.08) !important;
        border-radius: 10px !important; }
    div[data-baseweb="select"]:hover > div {
        border-color: rgba(168,85,247,0.3) !important; }
    div[data-baseweb="select"] > div:focus-within {
        border-color: #a855f7 !important;
        box-shadow: 0 0 0 2px rgba(168,85,247,0.15) !important; }

    /* File uploader */
    section[data-testid="stFileUploadDropzone"] {
        background: rgba(255,255,255,0.02) !important;
        border: 2px dashed rgba(0,212,255,0.2) !important;
        border-radius: 16px !important; }
    section[data-testid="stFileUploadDropzone"]:hover {
        border-color: rgba(0,212,255,0.4) !important;
        background: rgba(0,212,255,0.02) !important; }

    /* Tables */
    .stDataFrame { border-radius: 12px !important; overflow: hidden; }

    /* Expanders */
    details[data-testid="stExpander"] {
        background: rgba(255,255,255,0.02) !important;
        border: 1px solid rgba(255,255,255,0.06) !important;
        border-radius: 12px !important; }
    details[data-testid="stExpander"] summary { font-weight: 700 !important; }

    /* Dividers */
    hr { border-color: rgba(255,255,255,0.05) !important; opacity: 0.5; }

    /* Success / Warning / Error boxes */
    div[data-testid="stAlert"] { border-radius: 12px !important; }

    /* Badges */
    .badge {
        display: inline-block; padding: 3px 10px; border-radius: 999px;
        font-size: 0.7rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.5px; }
    .badge-green { background: rgba(0,232,143,0.12); color: #00e88f; border: 1px solid rgba(0,232,143,0.25); }
    .badge-purple { background: rgba(168,85,247,0.12); color: #c4b5fd; border: 1px solid rgba(168,85,247,0.25); }
    .badge-red { background: rgba(255,71,87,0.12); color: #ff8a8a; border: 1px solid rgba(255,71,87,0.25); }
    .badge-blue { background: rgba(0,212,255,0.12); color: #00d4ff; border: 1px solid rgba(0,212,255,0.25); }

    /* Filter sheet items */
    .sheet-item {
        background: rgba(168,85,247,0.04);
        border: 1px solid rgba(168,85,247,0.12);
        border-radius: 10px; padding: 12px 16px; margin-bottom: 8px;
        display: flex; align-items: center; gap: 10px; }
    .sheet-num {
        background: linear-gradient(135deg, #a855f7, #6366f1); color: white;
        width: 26px; height: 26px; border-radius: 8px; display: inline-flex;
        align-items: center; justify-content: center;
        font-size: 0.7rem; font-weight: 800; flex-shrink: 0; }
    .sheet-name { font-weight: 700; color: #c4b5fd; }
    .sheet-desc { color: #5a6a9a; font-size: 0.82rem; }
    .sheet-rows { color: #5a6a9a; font-size: 0.75rem; margin-left: auto; white-space: nowrap; }

    /* Sidebar */
    section[data-testid="stSidebar"] { background: #050a18 !important; }

    /* Separator with text */
    .section-sep {
        text-align: center; margin: 28px 0 10px; position: relative; }
    .section-sep::before {
        content: ''; position: absolute; top: 50%; left: 0; right: 0; height: 1px;
        background: linear-gradient(90deg, transparent, rgba(168,85,247,0.2), transparent); }
    .section-sep span {
        position: relative; background: rgba(10,17,40,0.95); padding: 6px 20px;
        font-size: 0.75rem; font-weight: 700; color: #a855f7;
        text-transform: uppercase; letter-spacing: 1px; border-radius: 999px;
        border: 1px solid rgba(168,85,247,0.15); }

    /* Page navigation tabs */
    .page-nav { display: flex; gap: 14px; margin: 4px 0 20px; }
    .page-nav .stButton > button {
        background: rgba(255,255,255,0.05) !important;
        border: 1px solid rgba(168,85,247,0.25) !important;
        color: #c4b5fd !important;
        box-shadow: none !important;
        border-radius: 12px !important;
        font-weight: 700 !important; }
    .page-nav .stButton > button:hover {
        border-color: rgba(0,212,255,0.5) !important;
        background: rgba(0,212,255,0.06) !important;
        color: #00d4ff !important; }

    /* Snapshot editor styling */
    div[data-testid="stDataEditor"] {
        border: 1px solid rgba(255,255,255,0.08) !important;
        border-radius: 12px !important; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ── State ────────────────────────────────────────────────────────────────────
for key, default in [
    ("merged_meta", None), ("filter_sheets", []),
    ("records", []), ("columns", []), ("mode_key", None),
    ("pending_filters", {}),
    ("unique_values_cache", {}),
    ("snap_page", False),
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
            r_date = st.text_input("Date", value=REPORT_DEFAULTS["date"], key="snap_rep_date")
            r_tagline2 = st.text_input("Tagline 2 (footer)", value=REPORT_DEFAULTS["tagline2"], key="snap_rep_tag2")

    report = {
        "title": r_title, "organization": r_org, "brand": r_brand,
        "tagline1": r_tagline1, "tagline2": r_tagline2,
        "subtitle": r_subtitle, "date": r_date,
    }

    st.markdown('<div class="card"><div class="card-head"><div class="card-icon icon-green">&#128202;</div>'
                '<div><p class="card-title">Transaction services</p>'
                '<p class="card-sub">Edit / add / remove rows &mdash; calculations update live</p></div></div>',
                unsafe_allow_html=True)

    default_df = pd.DataFrame([
        {"name": s["name"], "type": s["type"], "transactionVolume": int(s["transactionVolume"]),
         "totalValue": float(s["totalValue"]), "keyMessage": s["keyMessage"],
         "highlighted": s["highlighted"]}
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
                "Type", options=["financial", "non-financial"], width="small"),
            "transactionVolume": st.column_config.NumberColumn(
                "Transaction Volume", min_value=0, step=100, format="%.0f", width="small"),
            "totalValue": st.column_config.NumberColumn(
                "Total Value (ETB)", min_value=0.0, step=1000.0, format="%.2f", width="small"),
            "keyMessage": st.column_config.TextColumn("Key Message", width="large"),
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
            "keyMessage": row.get("keyMessage") or "",
            "highlighted": bool(row.get("highlighted")),
        })

    calc = calc_all(services)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-sep"><span>Live Preview</span></div>', unsafe_allow_html=True)
    st.caption("The report updates live. Use the buttons inside the preview to print (save as PDF) "
               "or download as PNG. Icons and image export load from CDNs, so an internet connection is required.")
    components.html(
        build_report_html(report, calc, show_bars=show_bars, auto_highlight=auto_hl,
                          takeaway_override=takeaway_override),
        height=860,
        scrolling=True,
    )


# ── Header ───────────────────────────────────────────────────────────────────
st.markdown('<p class="gradient-title">Report Merger</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Consolidate POS, ATM &amp; QR transaction reports &mdash; in memory, nothing saved to disk.</p>', unsafe_allow_html=True)

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

# ── Mode Selection ───────────────────────────────────────────────────────────
st.markdown('<div class="card"><div class="card-head"><div class="card-icon icon-purple">1</div><div><p class="card-title">Choose report type</p><p class="card-sub">Select the type of reports you want to merge</p></div></div>', unsafe_allow_html=True)

mode_options = ["POS Decline", "POS Success", "POS (Daily)", "ATM (Daily)", "QR"]
mode_keys_map = {
    "POS Decline": "pos_decline", "POS Success": "pos_success",
    "POS (Daily)": "pos", "ATM (Daily)": "atm", "QR": "qr",
}
mode_colors = {
    "POS Decline": "badge-red", "POS Success": "badge-green",
    "POS (Daily)": "badge-purple", "ATM (Daily)": "badge-blue",
    "QR": "badge-blue",
}

cols = st.columns(len(mode_options))
selected = None
for i, opt in enumerate(mode_options):
    with cols[i]:
        if st.button(opt, key=f"mode_{i}", use_container_width=True):
            st.session_state.mode_key = mode_keys_map[opt]
            st.session_state.merged_meta = None
            st.session_state.filter_sheets = []
            st.session_state.records = []
            st.session_state.columns = []
            st.session_state.unique_values_cache = {}
            gc.collect()

if st.session_state.mode_key is None:
    st.markdown('</div>', unsafe_allow_html=True)
    st.info("Select a report type above to get started.")
    st.stop()

mode_key = st.session_state.mode_key
mode = MODES[mode_key]
mode_label = [k for k, v in mode_keys_map.items() if v == mode_key][0]

st.markdown(f'<span class="badge {mode_colors.get(mode_label, "badge-blue")}">{mode.label}</span>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── File Upload ──────────────────────────────────────────────────────────────
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

st.markdown('</div>', unsafe_allow_html=True)

# ── Merge ────────────────────────────────────────────────────────────────────
if st.session_state.merged_meta is None:
    total_size_mb = sum(f.size for f in uploaded_files) / (1024 * 1024)
    if total_size_mb > 100:
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
            "per_file": result.per_file,
            "resp_counts": result.resp_counts,
            "warnings": result.warnings,
            "mode_key": result.mode_key,
            "mode_label": result.mode_label,
            "sort_by": result.sort_by,
            "sort_dir": result.sort_dir,
        }
        del result, payloads
        gc.collect()
        st.rerun()

    st.stop()

# ── Results ──────────────────────────────────────────────────────────────────
meta = st.session_state.merged_meta

st.markdown('<div class="card"><div class="card-head"><div class="card-icon icon-green">&#10003;</div><div><p class="card-title">Merged Report</p><p class="card-sub">All reports consolidated successfully</p></div></div>', unsafe_allow_html=True)

ok_files = [p for p in meta["per_file"] if p["status"] == "ok"]
c1, c2, c3, c4 = st.columns(4)
c1.metric("Transactions", f"{meta['total_rows']:,}")
if meta["from_date"] == meta["to_date"]:
    date_str = meta["from_date"]
else:
    date_str = f"{meta['from_date']} \u2192 {meta['to_date']}"
c2.metric("Date Range", date_str)
c3.metric("Files Merged", len(ok_files))
c4.metric("Response Codes", len(meta["resp_counts"]))

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
