"""
MMU Hub Report — Streamlit webapp
Karad Postal Division, Dept. of Posts — Pune Region, Maharashtra Circle

Implements the Meta_Prompt.md end goal:
- Permanently hosted Master file, replaceable from the UI
- India Post branded header
- 9 named file placeholders, each with its own calendar picker(s) — dates are
  NEVER inferred from filenames, only from the pickers
- "Generate" -> if files missing, warn and ask "still generate the report?"
- Live Divisional Summary widget (Range + Single Date, all 4 KPIs)

Run locally:   streamlit run app.py
Deploy:        Streamlit Community Cloud (no system packages needed).
"""
import datetime
import pandas as pd
import streamlit as st

from engine import (read_csv_any, read_master, build_full, build_divisional_summary,
                     SUB_DIVISIONS, save_master_file, load_persisted_master,
                     save_logo_file, load_persisted_logo)
from workbook import build_workbook

st.set_page_config(page_title="Karad Postal Division MMU Hub", page_icon="📮", layout="wide")

INDIA_POST_LOGO = "https://upload.wikimedia.org/wikipedia/en/3/32/India_Post.svg"
NATIONAL_EMBLEM = "https://upload.wikimedia.org/wikipedia/commons/5/55/Emblem_of_India.svg"
MAROON = "#8B0000"
MAROON_DARK = "#5C0000"
GOLD = "#F5B700"

# ------------------------------------------------------------------- CSS ---
st.markdown(f"""
<style>
    .gov-topbar {{
        background:#f2f2f2; padding:4px 16px; font-size:12px; color:#444;
        border-bottom:1px solid #ddd; margin:-1rem -1rem 0 -1rem;
    }}
    .gov-header {{
        background: linear-gradient(135deg, {MAROON} 0%, {MAROON_DARK} 100%);
        padding: 18px 28px; margin: 0 -1rem 1.2rem -1rem;
        border-bottom: 4px solid {GOLD};
        display:flex; align-items:center; gap:18px;
    }}
    .gov-header img {{ height:54px; background:white; border-radius:6px; padding:4px; }}
    .gov-header .titles h1 {{
        color:white; font-size:26px; font-weight:800; margin:0; line-height:1.25;
        font-family: 'Segoe UI', Arial, sans-serif;
    }}
    .gov-header .titles p {{
        color:#f0e0e0; font-size:13px; margin:2px 0 0 0; font-family: 'Nirmala UI', Arial, sans-serif;
    }}
    div[data-testid="stMetricValue"] {{ color:{MAROON_DARK}; }}
    .stButton>button[kind="primary"] {{ background-color:{MAROON}; border-color:{MAROON}; }}
</style>
<div class="gov-topbar">🇮🇳 Government of India &nbsp;|&nbsp; Ministry of Communications &nbsp;|&nbsp; Department of Posts</div>
""", unsafe_allow_html=True)

_logo_path = load_persisted_logo()
_logo_src = _logo_path if _logo_path else INDIA_POST_LOGO
st.markdown(f"""
<div class="gov-header">
    <img src="{NATIONAL_EMBLEM}" />
    <img src="{_logo_src}" />
    <div class="titles">
        <h1>Karad Postal Division MMU Hub</h1>
        <p>भारत सरकार | Government of India · संचार मंत्रालय · डाक विभाग · Department of Posts</p>
        <p>Pune Region, Maharashtra Circle — Mail Motor / Mobile Unit Performance Monitoring</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------ state ---
if 'confirm_missing' not in st.session_state:
    st.session_state.confirm_missing = False
if 'generated_buf' not in st.session_state:
    st.session_state.generated_buf = None

# ----------------------------------------------------------- Master file ---
st.subheader("📋 Office Master File")
persisted = load_persisted_master()
mcol1, mcol2 = st.columns([3, 2])
with mcol1:
    if persisted:
        st.success("A Master File is permanently hosted for this webapp and will be used automatically.")
    else:
        st.warning("No Master File is hosted yet — upload one below to set it up (this only needs doing once).")
with mcol2:
    with st.expander("Upload / Replace Master File"):
        new_master = st.file_uploader("Office Master File (.xlsx)", type=["xlsx"], key="master_upload")
        if new_master is not None and st.button("Save as the hosted Master File"):
            save_master_file(new_master)
            st.success("Master File saved and will now be used for every future run.")
            st.rerun()
        st.caption("⚠️ On Streamlit Community Cloud, hosted files persist while the app stays "
                    "running but reset on redeploy — for guaranteed permanence, self-host or "
                    "mount persistent storage.")
        st.markdown("---")
        st.caption("Optional: replace the India Post logo shown in the header above.")
        new_logo = st.file_uploader("Header logo (.png/.jpg/.svg)", type=["png", "jpg", "jpeg", "svg"], key="logo_upload")
        if new_logo is not None and st.button("Save as header logo"):
            save_logo_file(new_logo)
            st.success("Logo updated.")
            st.rerun()

if not persisted:
    st.stop()

master_raw = read_master(persisted)

st.divider()

# ------------------------------------------------------- input file grid ---
st.subheader("📁 Input Files — dates are always taken from the calendar pickers below, never from filenames")

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("**1️⃣ Delivery Productivity — Range 1**")
    dp_r1_file = st.file_uploader("Upload CSV", type=["csv"], key="dp_r1_f", label_visibility="collapsed")
    dp_r1_from = st.date_input("From", key="dp_r1_from", value=None)
    dp_r1_to = st.date_input("To", key="dp_r1_to", value=None)

    st.markdown("**2️⃣ Delivery Productivity — Range 2** *(only if your source system caps exports at ~15 days)*")
    dp_r2_file = st.file_uploader("Upload CSV", type=["csv"], key="dp_r2_f", label_visibility="collapsed")
    dp_r2_from = st.date_input("From", key="dp_r2_from", value=None)
    dp_r2_to = st.date_input("To", key="dp_r2_to", value=None)

    st.markdown("**3️⃣ Delivery Productivity — Single Date**")
    dp_sd_file = st.file_uploader("Upload CSV", type=["csv"], key="dp_sd_f", label_visibility="collapsed")
    dp_sd_date = st.date_input("Single Date", key="dp_sd_date", value=None)

with c2:
    st.markdown("**4️⃣ DSS Usage — Range**")
    dss_r_file = st.file_uploader("Upload CSV", type=["csv"], key="dss_r_f", label_visibility="collapsed")
    dss_r_from = st.date_input("From", key="dss_r_from", value=None)
    dss_r_to = st.date_input("To", key="dss_r_to", value=None)

    st.markdown("**5️⃣ DSS Usage — Single Date**")
    dss_sd_file = st.file_uploader("Upload CSV", type=["csv"], key="dss_sd_f", label_visibility="collapsed")
    dss_sd_date = st.date_input("Single Date", key="dss_sd_date", value=None)

    st.markdown("**6️⃣ Letter Box Clearance — Range**")
    lb_r_file = st.file_uploader("Upload CSV", type=["csv"], key="lb_r_f", label_visibility="collapsed")
    lb_r_from = st.date_input("From", key="lb_r_from", value=None)
    lb_r_to = st.date_input("To", key="lb_r_to", value=None)

with c3:
    st.markdown("**7️⃣ Letter Box Clearance — Single Date**")
    lb_sd_file = st.file_uploader("Upload CSV", type=["csv"], key="lb_sd_f", label_visibility="collapsed")
    lb_sd_date = st.date_input("Single Date", key="lb_sd_date", value=None)

    st.markdown("**8️⃣ COD Digital Transaction — Range**")
    cod_r_file = st.file_uploader("Upload CSV", type=["csv"], key="cod_r_f", label_visibility="collapsed")
    cod_r_from = st.date_input("From", key="cod_r_from", value=None)
    cod_r_to = st.date_input("To", key="cod_r_to", value=None)

    st.markdown("**9️⃣ COD Digital Transaction — Single Date**")
    cod_sd_file = st.file_uploader("Upload CSV", type=["csv"], key="cod_sd_f", label_visibility="collapsed")
    cod_sd_date = st.date_input("Single Date", key="cod_sd_date", value=None)

st.divider()

# --------------------------------------------------------- derive dates ---
def first_not_none(*vals):
    for v in vals:
        if v is not None:
            return v
    return None

canonical_single_date = first_not_none(dp_sd_date, dss_sd_date, lb_sd_date, cod_sd_date)
range_froms = [d for d in (dp_r1_from, dss_r_from, lb_r_from, cod_r_from) if d is not None]
range_tos = [d for d in (dp_r2_to or dp_r1_to, dss_r_to, lb_r_to, cod_r_to) if d is not None]
canonical_range_start = min(range_froms) if range_froms else None
canonical_range_end = max(range_tos) if range_tos else None

sd_values = [d for d in (dp_sd_date, dss_sd_date, lb_sd_date, cod_sd_date) if d is not None]
if len(set(sd_values)) > 1:
    st.warning(f"⚠️ The 4 Single Date pickers don't all agree ({sorted(set(sd_values))}). "
                "Using the earliest-listed one available (Delivery Productivity's, if provided) "
                "for the report title and office record number. Double-check this is intended.")

st.subheader("📅 Report dates (derived from your calendar picks above)")
d1, d2, d3 = st.columns(3)
d1.metric("Single Date (for title/record no.)", canonical_single_date.strftime('%d.%m.%Y') if canonical_single_date else "not set")
d2.metric("Range Start", canonical_range_start.strftime('%d.%m.%Y') if canonical_range_start else "not set")
d3.metric("Range End", canonical_range_end.strftime('%d.%m.%Y') if canonical_range_end else "not set")

if canonical_single_date is None:
    st.error("At least one Single Date picker must be set (Delivery Productivity's is preferred) "
              "before a report can be generated — the office record number depends on it.")
    st.stop()

record_date = canonical_single_date + datetime.timedelta(days=1)
st.caption(f"Office record number will read **No. KRDDN/G/Mails/MMU/Performance/"
            f"{record_date.strftime('%d.%m.%y')}** (day after Single Date).")

st.divider()

# ----------------------------------------------------------- file status ---
def status(label, present):
    return f"✅ {label}" if present else f"⬜ {label} — *not provided*"

file_checklist = [
    ("Delivery Productivity — Range 1", dp_r1_file is not None),
    ("Delivery Productivity — Range 2", dp_r2_file is not None),
    ("Delivery Productivity — Single Date", dp_sd_file is not None),
    ("DSS Usage — Range", dss_r_file is not None),
    ("DSS Usage — Single Date", dss_sd_file is not None),
    ("Letter Box Clearance — Range", lb_r_file is not None),
    ("Letter Box Clearance — Single Date", lb_sd_file is not None),
    ("COD Digital Transaction — Range", cod_r_file is not None),
    ("COD Digital Transaction — Single Date", cod_sd_file is not None),
]
missing = [label for label, present in file_checklist if not present]

st.subheader("✅ Input file status")
cols = st.columns(3)
for i, (label, present) in enumerate(file_checklist):
    cols[i % 3].markdown(status(label, present))

# ------------------------------------------------------------ read files ---
dp_r1_df = read_csv_any(dp_r1_file) if dp_r1_file else None
dp_r2_df = read_csv_any(dp_r2_file) if dp_r2_file else None
dp_sd_df = read_csv_any(dp_sd_file) if dp_sd_file else None
dss_r_df = read_csv_any(dss_r_file) if dss_r_file else None
dss_sd_df = read_csv_any(dss_sd_file) if dss_sd_file else None
lb_r_df = read_csv_any(lb_r_file) if lb_r_file else None
lb_sd_df = read_csv_any(lb_sd_file) if lb_sd_file else None
cod_r_df = read_csv_any(cod_r_file) if cod_r_file else None
cod_sd_df = read_csv_any(cod_sd_file) if cod_sd_file else None

try:
    df, excl_r, excl_sd, avail = build_full(
        master_raw, dp_r1_df, dp_r2_df, dp_sd_df, dss_r_df, dss_sd_df,
        cod_r_df, cod_sd_df, lb_r_df, lb_sd_df)
except Exception as e:
    st.error(f"Couldn't process the uploaded files — check they match the expected column layout "
              f"for their category. Error: {e}")
    st.stop()

if len(excl_r) or len(excl_sd):
    not_in_master = pd.concat([excl_r, excl_sd]).drop_duplicates('office_id')
    not_in_master = not_in_master[not_in_master['office_id'] != 0]
    if len(not_in_master):
        st.caption("⚠️ DSS file(s) contained office IDs not present in Master — excluded and "
                    "logged (not silently dropped): " +
                    ", ".join(f"{r.office_name} ({r.office_id})" for r in not_in_master.itertuples()))

st.divider()

# ------------------------------------------------------- live preview ---
st.subheader("📊 Divisional Summary — live preview (Including B.Os, Range + Single Date)")
with st.container(border=True):
    ds = build_divisional_summary(df)
    preview_rows = ds['incl_bo']
    preview_df = pd.DataFrame(preview_rows)[
        ['sub_division', 'delivery_pct_r', 'delivery_pct_sd', 'dss_pct_r', 'dss_pct_sd',
         'cod_pct_r', 'cod_pct_sd', 'lb_pct_r', 'lb_pct_sd']
    ].rename(columns={
        'sub_division': 'Sub Division', 'delivery_pct_r': 'Delivery % (R)', 'delivery_pct_sd': 'Delivery % (SD)',
        'dss_pct_r': 'DSS % (R)', 'dss_pct_sd': 'DSS % (SD)', 'cod_pct_r': 'COD % (R)', 'cod_pct_sd': 'COD % (SD)',
        'lb_pct_r': 'LB % (R)', 'lb_pct_sd': 'LB % (SD)',
    })
    pct_cols = [c for c in preview_df.columns if c != 'Sub Division']
    st.dataframe(
        preview_df.style.format({c: lambda v: '' if pd.isna(v) else f'{v*100:.2f}%' for c in pct_cols}),
        use_container_width=True, hide_index=True,
    )
    total_row = ds['incl_bo'][-1]

    def fmt(v):
        return '' if pd.isna(v) else f"{v*100:.2f}%"

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Delivery % (R) — Division", fmt(total_row['delivery_pct_r']),
               help=f"Single Date: {fmt(total_row['delivery_pct_sd'])}")
    m2.metric("DSS % (R) — Division", fmt(total_row['dss_pct_r']),
               help=f"Single Date: {fmt(total_row['dss_pct_sd'])}")
    m3.metric("COD % (R) — Division", fmt(total_row['cod_pct_r']),
               help=f"Single Date: {fmt(total_row['cod_pct_sd'])}")
    m4.metric("LB % (R) — Division", fmt(total_row['lb_pct_r']),
               help=f"Single Date: {fmt(total_row['lb_pct_sd'])}")

st.divider()

# ------------------------------------------------------------ generate ---
st.subheader("🛠️ Generate Report")


def do_generate():
    date_meta = {
        'DP R1': f"{dp_r1_from} to {dp_r1_to}" if dp_r1_from else None,
        'DP R2': f"{dp_r2_from} to {dp_r2_to}" if dp_r2_from else None,
        'DP SD': dp_sd_date, 'DSS R': f"{dss_r_from} to {dss_r_to}" if dss_r_from else None,
        'DSS SD': dss_sd_date, 'LB R': f"{lb_r_from} to {lb_r_to}" if lb_r_from else None,
        'LB SD': lb_sd_date, 'COD R': f"{cod_r_from} to {cod_r_to}" if cod_r_from else None,
        'COD SD': cod_sd_date,
    }
    with st.spinner("Building all 19 sheets..."):
        buf = build_workbook(
            master_raw, df, excl_r, excl_sd, avail,
            canonical_single_date, canonical_range_start, canonical_range_end,
            dp_r1_df, dp_r2_df, dp_sd_df, dss_r_df, dss_sd_df, lb_r_df, lb_sd_df,
            cod_r_df, cod_sd_df, date_meta=date_meta,
        )
    st.session_state.generated_buf = buf
    st.session_state.confirm_missing = False


if missing and not st.session_state.confirm_missing:
    if st.button("🛠️ Generate MMU Hub Report (.xlsx)", type="primary"):
        st.session_state.confirm_missing = True
        st.rerun()

if st.session_state.confirm_missing:
    st.warning(
        f"**{len(missing)} of 9 files not uploaded:** " + "; ".join(missing) + ".\n\n"
        "Each missing parameter will show as **blank** everywhere in the workbook "
        "(never 0%, N/A, or a dash), and defaulter/exceptional tables for that "
        "parameter will note data wasn't provided rather than falsely implying zero "
        "defaulters.\n\n**Still generate the report?**"
    )
    gc1, gc2 = st.columns([1, 4])
    if gc1.button("✅ Yes, generate anyway", type="primary"):
        do_generate()
        st.rerun()
    if gc2.button("Cancel"):
        st.session_state.confirm_missing = False
        st.rerun()

if not missing:
    if st.button("🛠️ Generate MMU Hub Report (.xlsx)", type="primary"):
        do_generate()
        st.rerun()

if st.session_state.generated_buf is not None:
    st.success("Workbook built — 8 visible sheets + 11 hidden sheets, per Build Spec v8.")
    fname = f"MMU_Hub_Report_{canonical_single_date.strftime('%d.%m.%Y')}.xlsx"
    st.download_button("⬇️ Download " + fname, data=st.session_state.generated_buf, file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
with st.expander("About this app"):
    st.markdown(
        "Implements the **MMU Hub** end goal for Karad Postal Division: a permanently-hosted "
        "Master file, 9 named input placeholders each with their own calendar picker (dates are "
        "never inferred from filenames), graceful per-parameter degradation with an explicit "
        "'Generate anyway?' confirmation, and a live Divisional Summary preview. The generated "
        "workbook follows **Build Spec v8** — Raw Data carries live cross-sheet formulas; every "
        "other visible sheet is hardcoded pre-computed values from this same engine, so nothing "
        "can silently drift between sheets."
    )
