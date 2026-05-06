"""
app.py — Thrust Test Rig Engineering Dashboard
Main entry point. Handles: page config, CSS, sidebar, data loading, routing.
All logic lives in python_functions/. Edit this file only for layout/routing changes.

Folder structure:
    app.py                          ← this file
    pdf_engine.py                   ← PDF generation
    python_functions/
        db.py                       ← SQLite functions
        data_pipeline.py            ← file loading, cleaning, normalization
        charts.py                   ← Plotly + Matplotlib helpers
        view_explorer.py            ← Log Library full-screen view
        view_dashboard.py           ← Test summary, measurable params, initial params
        view_plots.py               ← Custom plot, saved gallery, downloads
"""

import sys
import json
import importlib
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# ── Path setup so python_functions/ imports work ──
BASE_DIR  = Path(__file__).parent
LOGS_DIR  = BASE_DIR / "logs"
DB_PATH   = BASE_DIR / "thrust_logs.db"
LOGS_DIR.mkdir(exist_ok=True)

sys.path.insert(0, str(BASE_DIR))

# ── Dependency check ──
_missing = [p for p in ["openpyxl", "matplotlib", "reportlab", "plotly"]
            if importlib.util.find_spec(p) is None]
if _missing:
    st.error(
        f"Missing package(s): **{', '.join(_missing)}**\n\n"
        f"Run:  `pip install {' '.join(_missing)}`  then restart Streamlit."
    )
    st.stop()

# ── Module imports ──
from python_functions.db import (
    init_db, save_run, fetch_run, fetch_all_runs,
    delete_run, fetch_folders, move_run_to_folder
)
from python_functions.data_pipeline import (
    load_file_from_path, load_file_from_upload,
    normalize_columns, parse_time, clean_and_drop,
    extract_test_date, compute_stats, default_init_params
)
from python_functions.charts import (
    pl_single, pl_overlay, pl_multi,
    make_single_plot, make_overlay_plot, fig_to_png, DARK
)
from python_functions.view_explorer import render_explorer
from python_functions.view_dashboard import (
    render_test_summary, render_measurable_parameters, render_initial_parameters
)
from python_functions.view_plots import (
    render_custom_plot, render_saved_plots_gallery,
    render_update_parameters, render_downloads
)

# ── Init DB ──
init_db(DB_PATH)

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Thrust Test Rig Dashboard",
    page_icon="🚀",
    layout="wide",
)

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;500;700&display=swap');
  html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
  h1, h2, h3 { font-family: 'Space Mono', monospace; }
  .stMetric label {
      font-size: 0.72rem; letter-spacing: 0.08em; text-transform: uppercase;
  }
  section[data-testid="stSidebar"] { background: #0d0d0d; }
  section[data-testid="stSidebar"] * { color: #e0e0e0 !important; }
  .debug-box {
      background: #1a1a2e; color: #a0d8ef;
      font-family: 'Space Mono', monospace; font-size: 0.75rem;
      padding: 1rem; border-radius: 6px;
      border-left: 3px solid #4fc3f7; margin: 0.5rem 0; white-space: pre-wrap;
  }
  .run-card {
      background: #13161e; border: 1px solid #2a2d3a; border-radius: 8px;
      padding: 10px 14px; margin-bottom: 8px;
  }
  .run-card-active { border-color: #f97316 !important; background: #1a1a2e !important; }
  .run-date  { font-size: 0.7rem;  color: #6b7280; font-family: 'Space Mono', monospace; }
  .run-name  { font-size: 0.85rem; font-weight: 500; color: #e0e0e0; }
  .run-stats { font-size: 0.72rem; color: #9ca3af; margin-top: 3px; }
  .summary-card {
      background: #13161e; border: 1px solid #2a2d3a; border-radius: 10px;
      padding: 16px 20px; margin-bottom: 16px;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🚀 Thrust Test Rig")
    st.divider()

    sidebar_mode = st.radio(
        "Mode", ["📂 Log Library", "⬆️ Upload New Log"],
        label_visibility="collapsed"
    )
    st.divider()

    if sidebar_mode == "📂 Log Library":
        _folders = fetch_folders(DB_PATH)
        if _folders:
            st.markdown("**Folders**")
            if st.button("📋 All runs", use_container_width=True, key="folder_all"):
                st.session_state["active_folder"] = None
                st.session_state["mode"] = "explorer"
                st.rerun()
            for _f in _folders:
                _n = len(fetch_all_runs(folder=_f, db_path=DB_PATH))
                if st.button(f"📁 {_f}  ({_n})", use_container_width=True,
                             key=f"folder_{_f}"):
                    st.session_state["active_folder"] = _f
                    st.session_state["mode"] = "explorer"
                    st.rerun()
        else:
            st.caption("No saved runs yet. Import a log to get started.")

        if (st.session_state.get("mode") == "library"
                and st.session_state.get("selected_run")):
            st.divider()
            if st.button("← Back to Library", use_container_width=True):
                st.session_state["mode"] = "explorer"
                st.rerun()

        if st.session_state.get("mode") not in ("library", "explorer"):
            st.session_state["mode"] = "explorer"

    else:
        uploaded = st.file_uploader(
            "Upload log file",
            type=["xlsx", "xlsm", "xls", "csv", "txt"],
            label_visibility="collapsed"
        )
        st.session_state["mode"]          = "upload"
        st.session_state["uploaded_file"] = uploaded

    st.divider()
    st.markdown("**Compare mode**")
    _compare_on = st.toggle("Compare two runs", value=False, key="compare_on")
    if _compare_on:
        _all_runs = fetch_all_runs(db_path=DB_PATH)
        _run_opts = {r["display_name"]: r["filename"] for r in _all_runs}
        if len(_run_opts) < 2:
            st.caption("Need at least 2 saved runs.")
        else:
            _cmp_name = st.selectbox("Compare against:",
                                     list(_run_opts.keys()), key="cmp_run_select")
            st.session_state["cmp_filename"] = _run_opts[_cmp_name]
    else:
        st.session_state.pop("cmp_filename", None)

    st.divider()
    show_debug = st.toggle("Debug log",      value=False)
    show_raw   = st.toggle("Raw data table", value=False)
    st.divider()
    rpm_filter = st.slider("Min RPM filter", 0, 500, 0, step=50,
                           help="Exclude rows below this RPM")

# ─────────────────────────────────────────────
# MAIN AREA
# ─────────────────────────────────────────────
st.title("Thrust Test Rig — Engineering Dashboard")

mode = st.session_state.get("mode", "upload")

# ══════════════════════════════════════════════
# EXPLORER MODE
# ══════════════════════════════════════════════
if mode == "explorer":
    render_explorer(LOGS_DIR)
    st.stop()

# ══════════════════════════════════════════════
# DASHBOARD MODE (library or upload)
# ══════════════════════════════════════════════
df        = None
logs      = []
run_meta  = None
filename  = None

if mode == "library":
    sel = st.session_state.get("selected_run")
    if sel:
        run_meta = fetch_run(sel, DB_PATH)
        if run_meta and run_meta.get("file_path") and Path(run_meta["file_path"]).exists():
            df = load_file_from_path(Path(run_meta["file_path"]), logs)
            filename = run_meta["filename"]
        else:
            st.error("File was moved or deleted. Re-import it from the library.")
            st.stop()
    else:
        st.info("👈 Select a run from the Log Library.")
        st.stop()

elif mode == "upload":
    uploaded_file = st.session_state.get("uploaded_file")
    if not uploaded_file:
        st.info("👈 Upload a `.xlsx` or `.csv` log file, or switch to **Log Library**.")
        st.stop()
    df       = load_file_from_upload(uploaded_file, logs)
    filename = uploaded_file.name

if df is None:
    st.error("Could not parse the file.")
    st.code("\n".join(logs))
    st.stop()

# ── Clean pipeline ──
logs.append(f"📋 Raw: {df.shape[0]} rows × {df.shape[1]} cols")
df = normalize_columns(df, logs)
df = parse_time(df, logs)
df = clean_and_drop(df, logs)

if rpm_filter > 0 and "RPM" in df.columns:
    _before = len(df)
    df = df[df["RPM"] >= rpm_filter].copy()
    logs.append(f"🔧 RPM filter ≥{rpm_filter}: {len(df)}/{_before} rows kept")

try:
    _time_max_str = f"{float(df['Time'].max()):.1f}s"
except Exception:
    _time_max_str = str(df["Time"].max())
logs.append(f"✅ Final: {df.shape[0]} rows | Span: 0–{_time_max_str}")

# ── Load compare run ──
df2        = None
run_meta2  = None
label_run1 = filename or "Run A"
label_run2 = "Compare run"
_cmp_fn    = st.session_state.get("cmp_filename")
if _cmp_fn and _cmp_fn != filename:
    _cmp_meta = fetch_run(_cmp_fn, DB_PATH)
    if _cmp_meta and _cmp_meta.get("file_path") and Path(_cmp_meta["file_path"]).exists():
        _cmp_logs = []
        df2 = load_file_from_path(Path(_cmp_meta["file_path"]), _cmp_logs)
        if df2 is not None:
            df2 = normalize_columns(df2, _cmp_logs)
            df2 = parse_time(df2, _cmp_logs)
            df2 = clean_and_drop(df2, _cmp_logs)
            if rpm_filter > 0 and "RPM" in df2.columns:
                df2 = df2[df2["RPM"] >= rpm_filter].copy()
            run_meta2  = _cmp_meta
            label_run1 = run_meta["display_name"]  if run_meta  else (filename or "Run A")
            label_run2 = _cmp_meta["display_name"] if _cmp_meta else _cmp_fn

# ── Debug log ──
if show_debug:
    with st.expander("🔍 Debug log", expanded=True):
        st.markdown(
            "<div class='debug-box'>" + "\n".join(logs) + "</div>",
            unsafe_allow_html=True
        )

if df.empty:
    st.error("DataFrame is empty after cleaning.")
    st.stop()

# ── Summary card (library mode) ──
_compare_active = st.session_state.get("compare_on", False) and df2 is not None

if not _compare_active:
    if run_meta:
        saved = run_meta.get("saved_at", "")[:10]
        st.markdown(f"""
        <div class="summary-card">
          <div style="font-size:0.75rem;color:#6b7280;font-family:'Space Mono',monospace">
            {run_meta['test_date']} &nbsp;·&nbsp; saved {saved}
          </div>
          <div style="font-size:1.2rem;font-weight:600;color:#e0e0e0;margin:4px 0">
            {run_meta['display_name']}
          </div>
          <div style="font-size:0.8rem;color:#9ca3af">
            {run_meta['num_rows']:,} rows &nbsp;|&nbsp;
            {run_meta['duration_s']:.1f}s &nbsp;|&nbsp;
            Max Thrust {run_meta['max_thrust_n']:.1f} N &nbsp;|&nbsp;
            Max RPM {int(run_meta['max_rpm']):,}
          </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Dashboard sections (hidden in compare mode) ──
    render_test_summary(df)
    st.divider()
    render_measurable_parameters(df)
    st.divider()
    render_initial_parameters(df, filename)
    st.divider()

else:
    st.info(f"🔀 Compare mode: **{label_run1}** (solid) vs **{label_run2}** (dashed/dotted)")

# ── Always visible ──
render_custom_plot(df, df2, label_run1, label_run2)
render_saved_plots_gallery()

# ── Raw data table ──
if show_raw:
    st.divider()
    st.subheader("Cleaned data table")
    st.dataframe(df.reset_index(drop=True), use_container_width=True)

# ── Update parameters + downloads ──
render_update_parameters(df, filename, LOGS_DIR, mode)
render_downloads(df, filename)
