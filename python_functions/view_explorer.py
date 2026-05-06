"""
view_explorer.py — Full-screen Log Library explorer view
Renders when mode == "explorer". Handles folder nav, import, search, move.
Calls db.py and data_pipeline.py — no chart or PDF logic here.
"""

import streamlit as st
from pathlib import Path

from python_functions.db import (
    fetch_all_runs, fetch_folders, save_run, delete_run, move_run_to_folder
)
from python_functions.data_pipeline import (
    load_file_from_upload, normalize_columns, parse_time,
    clean_and_drop, compute_stats, default_init_params, extract_test_date
)


def render_explorer(logs_dir: Path):
    """
    Render the full-screen log library explorer.
    Call st.stop() after this to prevent dashboard rendering.

    Args:
        logs_dir: Path to the logs/ folder where files are stored.
    """
    active_folder = st.session_state.get("active_folder")

    # ── Top bar ──
    _ex1, _ex2, _ex3 = st.columns([3, 2, 2])
    _ex1.subheader(
        "📂 Log Library" + (f" — {active_folder}" if active_folder else " — All Runs")
    )

    # Import file uploader (right column)
    with _ex3:
        _imp_file = st.file_uploader(
            "Import log",
            type=["xlsx", "xlsm", "xls", "csv", "txt"],
            label_visibility="collapsed",
            key="explorer_import"
        )
        if _imp_file:
            st.session_state["pending_import"] = _imp_file

    st.divider()

    # ── Pending import form ──
    if st.session_state.get("pending_import"):
        _pf = st.session_state["pending_import"]
        st.markdown("#### Import new log")
        _if1, _if2, _if3 = st.columns([3, 2, 1])

        _imp_name = _if1.text_input(
            "Run name",
            value=_pf.name.rsplit(".", 1)[0],
            key="imp_run_name"
        )

        _all_folders = fetch_folders() or ["Uncategorised"]
        _imp_folder  = _if2.selectbox(
            "Folder",
            _all_folders + ["+ New folder…"],
            index=_all_folders.index(active_folder)
                  if active_folder and active_folder in _all_folders else 0,
            key="imp_folder_sel"
        )
        if _imp_folder == "+ New folder…":
            _imp_folder = _if2.text_input("New folder name", key="imp_new_folder_txt")

        if _if3.button("✅ Confirm Import", type="primary", use_container_width=True):
            _imp_logs = []
            _imp_df   = load_file_from_upload(_pf, _imp_logs)
            if _imp_df is None:
                st.error("Could not parse file. Check the file format.")
            else:
                _imp_df = normalize_columns(_imp_df, _imp_logs)
                _imp_df = parse_time(_imp_df, _imp_logs)
                _imp_df = clean_and_drop(_imp_df, _imp_logs)

                # Save file to logs/
                _dest = logs_dir / _pf.name
                _pf.seek(0)
                _dest.write_bytes(_pf.read())

                # Save to library immediately with auto-filled params
                _imp_stats  = compute_stats(_imp_df)
                _imp_params = default_init_params(_imp_df)
                _imp_date   = extract_test_date(_pf.name)
                save_run(
                    _pf.name,
                    _imp_name or _pf.name,
                    _imp_date,
                    _dest,
                    _imp_stats,
                    _imp_params,
                    folder=_imp_folder or "Uncategorised"
                )
                st.session_state.pop("pending_import", None)
                st.success(f"✅ '{_imp_name}' imported to folder '{_imp_folder}'")
                st.rerun()

        if st.button("✕ Cancel", key="imp_cancel"):
            st.session_state.pop("pending_import", None)
            st.rerun()

        st.divider()

    # ── Search bar ──
    _sc1, _sc2, _sc3 = st.columns([3, 2, 2])
    _search  = _sc1.text_input(
        "🔍 Search", placeholder="Run name, date…",
        label_visibility="collapsed", key="exp_search"
    )
    _df_from = _sc2.date_input("From", value=None,
                                label_visibility="collapsed", key="exp_from")
    _df_to   = _sc3.date_input("To",   value=None,
                                label_visibility="collapsed", key="exp_to")

    _runs = fetch_all_runs(_search, _df_from, _df_to, folder=active_folder)

    if not _runs:
        st.info("No runs found. Import a log file above to get started.")
        return

    # ── Runs table grouped by folder ──
    _by_folder = {}
    for _r in _runs:
        _by_folder.setdefault(_r["folder"] or "Uncategorised", []).append(_r)

    for _fold, _fold_runs in _by_folder.items():
        st.markdown(f"### 📁 {_fold}  `{len(_fold_runs)} run(s)`")

        # Table header
        _h = st.columns([3, 1, 1, 1, 1, 1, 1])
        for _col, _label in zip(
            _h, ["Run name", "Date", "Max RPM", "Max Thrust", "Duration", "", ""]
        ):
            _col.markdown(f"**{_label}**")

        for _r in _fold_runs:
            _rc = st.columns([3, 1, 1, 1, 1, 1, 1])
            _rc[0].write(_r["display_name"])
            _rc[1].write(_r["test_date"] or "—")
            _rc[2].write(f"{int(_r['max_rpm'])} RPM"      if _r["max_rpm"]      else "—")
            _rc[3].write(f"{_r['max_thrust_n']:.0f} N"    if _r["max_thrust_n"] else "—")
            _rc[4].write(f"{_r['duration_s']:.0f}s"       if _r["duration_s"]   else "—")

            if _rc[5].button("Open", key=f"exp_open_{_r['filename']}",
                             use_container_width=True, type="primary"):
                st.session_state["selected_run"] = _r["filename"]
                st.session_state["mode"] = "library"
                st.rerun()

            if _rc[6].button("🗑", key=f"exp_del_{_r['filename']}",
                             use_container_width=True):
                delete_run(_r["filename"])
                st.rerun()

        # Move between folders
        with st.expander(f"Move runs in '{_fold}' to another folder"):
            _mv1, _mv2, _mv3 = st.columns([3, 2, 1])
            _mv_run = _mv1.selectbox(
                "Run",
                [r["display_name"] for r in _fold_runs],
                key=f"mv_run_{_fold}"
            )
            _all_f = fetch_folders()
            _mv_to = _mv2.selectbox(
                "Move to folder",
                [f for f in _all_f if f != _fold] + ["+ New folder…"],
                key=f"mv_to_{_fold}"
            )
            if _mv_to == "+ New folder…":
                _mv_to = _mv2.text_input("New folder name", key=f"mv_new_{_fold}")
            if _mv3.button("Move", key=f"mv_btn_{_fold}", use_container_width=True):
                _mv_fn = next(
                    (r["filename"] for r in _fold_runs if r["display_name"] == _mv_run),
                    None
                )
                if _mv_fn and _mv_to:
                    move_run_to_folder(_mv_fn, _mv_to)
                    st.rerun()

        st.divider()
