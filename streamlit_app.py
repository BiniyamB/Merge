"""POS & ATM Report Merger -- Streamlit version."""

import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="POS & ATM Report Merger",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

from merger import MODES, merge_reports, build_filtered_workbook

st.markdown("""
<style>
    #MainMenu, footer { display: none !important; }
    .stApp { background: #0a1128; }
    .block-container { max-width: 1100px; padding-top: 2rem; }
    h1 span { background: linear-gradient(135deg, #ff3cac, #784ba0, #2b86c5);
              -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08);
        border-radius: 12px; padding: 12px 16px; }
    div[data-testid="stMetric"] label { color: #5a6a9a !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: #00d4ff !important; }
    section[data-testid="stSidebar"] { background: #050a18; }
    .stButton > button {
        background: linear-gradient(135deg, #ff3cac, #784ba0, #2b86c5) !important;
        color: white !important; border: none !important; border-radius: 10px !important;
        font-weight: 700 !important; padding: 0.5rem 2rem !important; }
    .stButton > button:hover { opacity: 0.9; transform: translateY(-1px); }
    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #00e88f, #00b4d8) !important;
        color: #052e16 !important; border: none !important; border-radius: 10px !important;
        font-weight: 700 !important; }
    hr { border-color: rgba(255,255,255,0.06) !important; }
    .stRadio > div { flex-direction: row !important; gap: 1rem; }
    div[data-baseweb="radio"] > label { background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08); border-radius: 10px;
        padding: 8px 18px; margin: 0; cursor: pointer; }
    div[data-baseweb="radio"] > label:hover { border-color: rgba(168,85,247,0.4); }
    div[data-baseweb="radio"] input:checked + div {
        background: linear-gradient(135deg, #a855f7, #6366f1) !important; }
</style>
""", unsafe_allow_html=True)

# ── State ────────────────────────────────────────────────────────────────────
if "merge_result" not in st.session_state:
    st.session_state.merge_result = None
if "filter_sheets" not in st.session_state:
    st.session_state.filter_sheets = []
if "unique_values" not in st.session_state:
    st.session_state.unique_values = {}
if "records" not in st.session_state:
    st.session_state.records = []
if "columns" not in st.session_state:
    st.session_state.columns = []
if "mode_key" not in st.session_state:
    st.session_state.mode_key = None

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("# POS & ATM <span>Report Merger</span>", unsafe_allow_html=True)
st.caption("Consolidate POS & ATM transaction reports — in memory, nothing saved to disk.")

st.divider()

# ── Mode selection ───────────────────────────────────────────────────────────
st.subheader("1. Choose report type")

mode_options = ["POS Decline", "POS Success", "POS (Daily)", "ATM (Daily)"]
mode_keys_map = {
    "POS Decline": "pos_decline",
    "POS Success": "pos_success",
    "POS (Daily)": "pos",
    "ATM (Daily)": "atm",
}
mode_key = mode_keys_map[st.radio("Report type", mode_options, horizontal=True, label_visibility="collapsed")]
mode = MODES[mode_key]
st.session_state.mode_key = mode_key

# ── File upload ──────────────────────────────────────────────────────────────
st.subheader(f"2. Upload {mode.label} reports")

uploaded_files = st.file_uploader(
    "Drag & drop your Excel files here",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
    help="One or more .xls / .xlsx files",
)

if not uploaded_files:
    st.stop()

st.success(f"**{len(uploaded_files)}** file(s) ready:")
for f in uploaded_files:
    size_kb = f.size / 1024
    st.markdown(f"  - `{f.name}` ({size_kb:.1f} KB)")

# ── Merge ────────────────────────────────────────────────────────────────────
st.divider()

if st.button("Merge Reports", key="merge_btn", use_container_width=True):
    with st.spinner("Merging reports..."):
        payloads = [(f.name, f.getvalue()) for f in uploaded_files]
        try:
            result = merge_reports(payloads, mode_key=mode_key)
        except ValueError as e:
            st.error(str(e))
            st.stop()

    st.session_state.merge_result = result
    st.session_state.records = result.records
    st.session_state.columns = list(result.records[0].keys()) if result.records else []
    st.session_state.filter_sheets = []

    # Compute unique values per column
    unique_vals = {}
    for col in st.session_state.columns:
        vals = sorted({str(r.get(col, "")).strip() for r in result.records if r.get(col, "") not in ("", None)})
        unique_vals[col] = vals
    st.session_state.unique_values = unique_vals

    st.success("Merge complete!")

# ── Results ──────────────────────────────────────────────────────────────────
result = st.session_state.merge_result
if result is None:
    st.stop()

st.divider()
st.subheader(f"Merged Report — {result.mode_label}")

# Stats
ok_files = [p for p in result.per_file if p["status"] == "ok"]
c1, c2, c3, c4 = st.columns(4)
c1.metric("Transactions", f"{result.total_rows:,}")
c2.metric("Date Range", f"{result.from_date} → {result.to_date}" if result.from_date != result.to_date else result.from_date)
c3.metric("Files Merged", len(ok_files))
c4.metric("Response Codes", len(result.resp_counts))

# Warnings
if result.warnings:
    with st.expander("Warnings", expanded=False):
        for w in result.warnings:
            st.warning(w)

# Per-file details
with st.expander("Files merged", expanded=False):
    file_df = pd.DataFrame(result.per_file)
    st.dataframe(file_df, use_container_width=True, hide_index=True)

# Response code chart
if result.resp_counts:
    with st.expander("Response code distribution", expanded=False):
        resp_df = pd.DataFrame(
            list(result.resp_counts.items()),
            columns=["Response Code", "Count"]
        ).sort_values("Count", ascending=False).head(15)
        st.bar_chart(resp_df.set_index("Response Code"))

# Preview
st.subheader("Preview")
preview_rows = []
for row in result.records[:50]:
    preview_rows.append({k: str(v) if v is not None and v != "" else "—" for k, v in row.items()})
preview_df = pd.DataFrame(preview_rows)
st.dataframe(preview_df, use_container_width=True, hide_index=True, height=400)

# Download merged
st.divider()
st.download_button(
    label="Download Merged Report",
    data=result.workbook_bytes,
    file_name=result.filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

# ── Filter Panel ─────────────────────────────────────────────────────────────
st.divider()
st.subheader("3. Filter & Export to Sheets")
st.caption("Each filter combination becomes a separate sheet in the downloaded Excel file.")

col_a, col_b, col_c = st.columns([2, 3, 1])

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
        vals = ["(All — no filter)"] + st.session_state.unique_values.get(filter_col, [])
        filter_val = st.selectbox("Value", options=vals, key="filter_val_select", index=0)
    else:
        filter_val = st.selectbox("Value", options=["Select a column first..."], key="filter_val_select", disabled=True)

with col_c:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("+ Add Sheet", key="add_sheet_btn", disabled=not filter_col, use_container_width=True):
        filters = {}
        if filter_col and filter_val and filter_val != "(All — no filter)":
            filters[filter_col] = filter_val

        # Count rows
        if not filters:
            count = len(st.session_state.records)
            name = "All"
        else:
            val_upper = str(list(filters.values())[0]).upper()
            col_name = list(filters.keys())[0]
            count = sum(1 for r in st.session_state.records if str(r.get(col_name, "")).upper() == val_upper)
            val_short = filter_val[:15] + "..." if len(filter_val) > 15 else filter_val
            name = f"{filter_col}_{val_short}"

        # Check duplicate
        key = str(filters)
        if not any(str(s["filters"]) == key for s in st.session_state.filter_sheets):
            st.session_state.filter_sheets.append({"name": name, "filters": filters, "count": count})
            st.rerun()
        else:
            st.warning("This filter combination already exists.")

# Display filter sheets
if st.session_state.filter_sheets:
    st.markdown("**Sheets to export:**")
    for i, sheet in enumerate(st.session_state.filter_sheets):
        col1, col2, col3 = st.columns([0.5, 5, 0.5])
        with col1:
            st.markdown(f"**{i+1}.**")
        with col2:
            if sheet["filters"]:
                desc = ", ".join(f"`{k}` = `{v}`" for k, v in sheet["filters"].items())
            else:
                desc = "All rows (unfiltered)"
            st.markdown(f"**{sheet['name']}** — {desc} — ~{sheet['count']:,} rows")
        with col3:
            if st.button("X", key=f"rm_sheet_{i}"):
                st.session_state.filter_sheets.pop(i)
                st.rerun()

    # Download filtered
    st.divider()
    col_dl, col_clr = st.columns([3, 1])
    with col_dl:
        with st.spinner("Generating filtered workbook..."):
            try:
                mode_obj = MODES[st.session_state.mode_key]
                filtered_wb = build_filtered_workbook(
                    st.session_state.records,
                    st.session_state.columns,
                    st.session_state.filter_sheets,
                    mode_obj,
                )
                if len(st.session_state.filter_sheets) == 1:
                    dl_name = f"{mode_obj.output_prefix}_{st.session_state.filter_sheets[0]['name']}_Filtered.xlsx"
                else:
                    dl_name = f"{mode_obj.output_prefix}_Filtered_{len(st.session_state.filter_sheets)}_Sheets.xlsx"

                st.download_button(
                    label=f"Download Filtered Report ({len(st.session_state.filter_sheets)} sheets)",
                    data=filtered_wb,
                    file_name=dl_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Error building workbook: {e}")

    with col_clr:
        if st.button("Clear All", key="clear_sheets"):
            st.session_state.filter_sheets = []
            st.rerun()
else:
    st.info("Use the dropdowns above to add filter sheets, then download a multi-sheet Excel file.")
