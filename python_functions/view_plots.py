"""
view_plots.py — Custom plot, saved plots gallery, update parameters, downloads.
Called from app.py after the dashboard sections render.
"""

import json
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from pathlib import Path

from python_functions.db import fetch_run, save_run
from python_functions.data_pipeline import compute_stats, extract_test_date
from python_functions.charts import (
    pl_single, pl_overlay, pl_multi,
    make_single_plot, fig_to_png, DARK
)
from pdf_engine import build_pdf_report


# ─────────────────────────────────────────────
# CUSTOM PLOT
# ─────────────────────────────────────────────
def render_custom_plot(df: pd.DataFrame, df2, label_run1: str, label_run2: str):
    """Render the interactive custom Plotly plot with save-to-PDF capability."""
    try:
        import plotly.graph_objects as go
        _plotly_ok = True
    except ImportError:
        _plotly_ok = False

    st.subheader("📈 Custom Plot")

    if not _plotly_ok:
        st.warning("Plotly not installed. Run:  `pip install plotly`  then restart Streamlit.")
        return

    _EXCLUDE   = {"Thrust_0deg_kg", "Thrust_90deg_kg", "Thrust_180deg_kg",
                  "Thrust_270deg_kg", "Total_Weight"}
    _plot_cols = [c for c in df.columns
                  if c not in _EXCLUDE and pd.api.types.is_numeric_dtype(df[c])]
    _none_opts = ["None"] + _plot_cols

    # ── Fixed colours per Y axis (cycles for unlimited axes) ──
    _Y_COLOURS = ["#f97316", "#38bdf8", "#a78bfa", "#4ade80",
                  "#fb923c", "#22d3ee", "#c084fc", "#86efac"]

    # ── Axis selectors ──
    _x_default = "Time"   if "Time"   in _plot_cols else _plot_cols[0]
    _y_default = "Thrust" if "Thrust" in _plot_cols else (
                 "RPM"    if "RPM"    in _plot_cols else _plot_cols[1])

    # Session state tracks list of extra Y columns (unlimited)
    if "plt_extra_y" not in st.session_state:
        st.session_state["plt_extra_y"] = []

    # X + Y1 always shown
    _ax1, _ax2, _ax_btn_add, _ax_btn_rem = st.columns([2, 2, 1, 1])
    x_col = _ax1.selectbox("X axis",  _plot_cols,
                            index=_plot_cols.index(_x_default), key="plt_x")
    y_col = _ax2.selectbox("Y1 axis", _plot_cols,
                            index=_plot_cols.index(_y_default), key="plt_y")

    if _ax_btn_add.button("＋ Y axis", use_container_width=True, key="plt_add_y"):
        st.session_state["plt_extra_y"].append("None")
        st.rerun()
    if _ax_btn_rem.button("－ Y axis", use_container_width=True,
                          disabled=len(st.session_state["plt_extra_y"]) == 0,
                          key="plt_rem_y"):
        st.session_state["plt_extra_y"].pop()
        st.rerun()

    # Show extra Y axis selectors
    _extra_y_cols = []
    _none_opts    = ["None"] + _plot_cols
    for _yi, _yval in enumerate(st.session_state["plt_extra_y"]):
        _yn = st.selectbox(
            f"Y{_yi + 2} axis", _none_opts,
            index=_none_opts.index(_yval) if _yval in _none_opts else 0,
            key=f"plt_y_extra_{_yi}"
        )
        st.session_state["plt_extra_y"][_yi] = _yn
        _extra_y_cols.append(_yn)

    # ── Controls row — Type + Time window only ──
    _ctrl1, _ctrl2, _ctrl3, _ctrl4 = st.columns([1, 1, 2, 2])
    _plt_type   = _ctrl1.radio("Type", ["Line", "Scatter"], key="plt_type")
    _use_window = _ctrl2.checkbox("Time\nwindow", value=False, key="plt_win")

    _plt_tmin = 0.0
    _plt_tmax = float(df["Time"].max()) if "Time" in df.columns else 0.0
    if _use_window and "Time" in df.columns:
        _plt_tmin = _ctrl3.number_input("From (s)", value=0.0,
                                         max_value=float(df["Time"].max()),
                                         step=1.0, key="plt_tmin")
        _plt_tmax = _ctrl4.number_input("To (s)", value=float(df["Time"].max()),
                                         max_value=float(df["Time"].max()),
                                         step=1.0, key="plt_tmax")
    else:
        _df_plot  = df.copy()
        _df_plot2 = df2.copy() if df2 is not None else None

    if _use_window and "Time" in df.columns:
        _df_plot  = df[(df["Time"] >= _plt_tmin) & (df["Time"] <= _plt_tmax)].copy()
        _df_plot2 = df2[(df2["Time"] >= _plt_tmin) & (df2["Time"] <= _plt_tmax)].copy() \
                    if df2 is not None and "Time" in df2.columns else df2
    else:
        _df_plot  = df.copy()
        _df_plot2 = df2.copy() if df2 is not None else None

    _step    = max(1, len(_df_plot) // 5000)
    _df_plot = _df_plot.iloc[::_step]
    _ds_note = f"({len(_df_plot):,} pts)"

    _BG = "#0d0f14"; _PP = "#13161e"; _GR = "#1e2130"; _TX = "#c8ccd8"
    _ms = "lines" if _plt_type == "Line" else "markers"
    _mz = 3 if _plt_type == "Scatter" else 4

    # Active extra Y columns (ignore "None")
    _active_extras = [(i, col) for i, col in enumerate(_extra_y_cols)
                      if col != "None" and col in _df_plot.columns]

    fig = go.Figure()

    # Y1 trace
    _c1 = _Y_COLOURS[0]
    fig.add_trace(go.Scatter(
        x=_df_plot[x_col], y=_df_plot[y_col], mode=_ms,
        name=f"{y_col} ({label_run1})",
        line=dict(color=_c1, width=1.6) if _plt_type == "Line" else None,
        marker=dict(size=_mz, color=_c1), yaxis="y",
        hovertemplate=f"<b>{x_col}</b>: %{{x:.3f}}<br>"
                      f"<b>{y_col}</b>: %{{y:.3f}}<extra>{label_run1}</extra>",
    ))

    # Extra Y traces
    for _ei, _ecol in _active_extras:
        _ec = _Y_COLOURS[(_ei + 1) % len(_Y_COLOURS)]
        _yaxis_id = f"y{_ei + 2}"
        fig.add_trace(go.Scatter(
            x=_df_plot[x_col], y=_df_plot[_ecol], mode=_ms,
            name=f"{_ecol} ({label_run1})",
            line=dict(color=_ec, width=1.4) if _plt_type == "Line" else None,
            marker=dict(size=_mz, color=_ec), yaxis=_yaxis_id,
            hovertemplate=f"<b>{x_col}</b>: %{{x:.3f}}<br>"
                          f"<b>{_ecol}</b>: %{{y:.3f}}<extra>{label_run1}</extra>",
        ))

    # Compare run traces
    if _df_plot2 is not None:
        _step2    = max(1, len(_df_plot2) // 5000)
        _df_plot2 = _df_plot2.iloc[::_step2]
        if x_col in _df_plot2.columns and y_col in _df_plot2.columns:
            fig.add_trace(go.Scatter(
                x=_df_plot2[x_col], y=_df_plot2[y_col], mode=_ms,
                name=f"{y_col} ({label_run2})",
                line=dict(color=_c1, width=1.2, dash="dot") if _plt_type == "Line" else None,
                marker=dict(size=_mz, color=_c1, symbol="diamond"), yaxis="y",
                hovertemplate=f"<b>{x_col}</b>: %{{x:.3f}}<br>"
                              f"<b>{y_col}</b>: %{{y:.3f}}<extra>{label_run2}</extra>",
            ))
        for _ei, _ecol in _active_extras:
            _ec = _Y_COLOURS[(_ei + 1) % len(_Y_COLOURS)]
            _yaxis_id = f"y{_ei + 2}"
            if _ecol in _df_plot2.columns:
                fig.add_trace(go.Scatter(
                    x=_df_plot2[x_col], y=_df_plot2[_ecol], mode=_ms,
                    name=f"{_ecol} ({label_run2})",
                    line=dict(color=_ec, width=1.0, dash="dot") if _plt_type == "Line" else None,
                    marker=dict(size=_mz, color=_ec, symbol="diamond"), yaxis=_yaxis_id,
                    hovertemplate=f"<b>{x_col}</b>: %{{x:.3f}}<br>"
                                  f"<b>{_ecol}</b>: %{{y:.3f}}<extra>{label_run2}</extra>",
                ))

    # ── Layout ──
    _n_extra  = len(_active_extras)
    _r_margin = max(60, 60 + _n_extra * 60)
    _x_domain = [0, max(0.60, 1.0 - _n_extra * 0.10)]

    _layout = dict(
        plot_bgcolor=_BG, paper_bgcolor=_PP,
        font=dict(color=_TX, family="monospace", size=11),
        xaxis=dict(
            title=dict(text=x_col, font=dict(color=_TX)),
            gridcolor=_GR, gridwidth=0.5,
            showline=True, linecolor="#2a2d3a",
            tickfont=dict(color="#6b7280"),
            domain=_x_domain,
            rangeslider=dict(visible=True, thickness=0.06),
        ),
        yaxis=dict(
            title=dict(text=y_col, font=dict(color=_c1)),
            tickfont=dict(color=_c1),
            gridcolor=_GR, gridwidth=0.5,
            showline=True, linecolor="#2a2d3a",
        ),
        legend=dict(bgcolor="#13161e", bordercolor="#2a2d3a", borderwidth=1,
                    font=dict(color=_TX)),
        margin=dict(l=60, r=_r_margin, t=40, b=50),
        hovermode="x unified",
        height=530,
    )

    # Add yaxis2, yaxis3... for each extra axis
    for _ei, _ecol in _active_extras:
        _ec       = _Y_COLOURS[(_ei + 1) % len(_Y_COLOURS)]
        _axis_key = f"yaxis{_ei + 2}"
        _pos      = round(1.0 - _ei * 0.10, 2)
        _layout[_axis_key] = dict(
            title=dict(text=_ecol, font=dict(color=_ec)),
            tickfont=dict(color=_ec),
            overlaying="y", side="right",
            anchor="free", position=_pos,
            showgrid=False, showline=True, linecolor="#2a2d3a",
        )

    fig.update_layout(**_layout)
    st.plotly_chart(fig, use_container_width=True)

    _extra_labels = "".join([f"  |  Y{i+2}: **{c}**" for i, c in _active_extras])
    st.caption(
        f"X: **{x_col}**  |  Y1: **{y_col}**" + _extra_labels
        + (f"  |  vs **{label_run2}**" if _df_plot2 is not None else "")
        + f"  |  {_ds_note}"
    )

    # ── Save Plot button ──
    _sp1, _sp2 = st.columns([3, 1])
    _extra_names  = " & ".join([c for _, c in _active_extras])
    _auto_title   = f"{y_col} vs {x_col}" + (f" & {_extra_names}" if _extra_names else "")
    _title_key    = f"plot_title_{x_col}_{y_col}_{'_'.join([c for _,c in _active_extras])}"
    _plot_title_input = _sp1.text_input(
        "Plot title for report", value=_auto_title,
        key=_title_key, label_visibility="collapsed",
        placeholder="Enter plot title…"
    )
    if _sp2.button("💾 Save Plot", type="primary", use_container_width=True, key="save_plot_btn"):
        if "saved_plots" not in st.session_state:
            st.session_state["saved_plots"] = []

        # Render as matplotlib PNG for PDF
        with plt.style.context(DARK):
            _pdf_fig, _pdf_ax1 = plt.subplots(figsize=(11, 3.4))
            _dfp = _df_plot
            _pdf_ax1.plot(_dfp[x_col], _dfp[y_col],
                          color=_c1, linewidth=1.4, label=y_col)
            _pdf_ax1.set_xlabel(x_col, fontsize=8)
            _pdf_ax1.set_ylabel(y_col, color=_c1, fontsize=8)
            _pdf_ax1.tick_params(axis="y", colors=_c1)
            _prev_ax = _pdf_ax1
            for _ei, _ecol in _active_extras:
                _ec   = _Y_COLOURS[(_ei + 1) % len(_Y_COLOURS)]
                _axi  = _prev_ax.twinx()
                if _ei > 0:
                    _axi.spines["right"].set_position(("axes", 1.0 + _ei * 0.12))
                _axi.plot(_dfp[x_col], _dfp[_ecol],
                          color=_ec, linewidth=1.2, label=_ecol)
                _axi.set_ylabel(_ecol, color=_ec, fontsize=8)
                _axi.tick_params(axis="y", colors=_ec)
                _prev_ax = _axi
            _pdf_fig.legend(loc="upper left", fontsize=8, bbox_to_anchor=(0.08, 0.95))
            _pdf_fig.tight_layout()
        _png = fig_to_png(_pdf_fig)

        _save_title = _plot_title_input or _auto_title
        _existing   = [p["title"] for p in st.session_state["saved_plots"]]
        if _save_title in _existing:
            for _p in st.session_state["saved_plots"]:
                if _p["title"] == _save_title:
                    _p["png"] = _png
            st.toast(f"Updated: {_save_title}", icon="✅")
        else:
            st.session_state["saved_plots"].append({
                "title": _save_title, "png": _png,
                "x": x_col, "y": y_col,
            })
            st.toast(f"Saved: {_save_title}", icon="✅")

    # ── Compare diff table ──
    if df2 is not None:
        st.divider()
        st.markdown("#### 📊 Run comparison summary")
        _diff_cols = ["RPM", "Thrust", "Voltage", "Current",
                      "Motor_Temp", "ESC_Temp", "Power", "Torque"]
        _diff_rows = []
        for _dc in _diff_cols:
            v1 = df[_dc].mean()  if _dc in df.columns  and df[_dc].notna().any()  else None
            v2 = df2[_dc].mean() if _dc in df2.columns and df2[_dc].notna().any() else None
            if v1 is None and v2 is None:
                continue
            _d  = (v2 - v1)     if (v1 is not None and v2 is not None) else None
            _dp = (_d/v1*100)   if (_d is not None and v1 != 0)        else None
            _diff_rows.append({
                "Metric":   _dc,
                label_run1: f"{v1:.2f}" if v1 is not None else "—",
                label_run2: f"{v2:.2f}" if v2 is not None else "—",
                "Delta":    f"{_d:+.2f}"    if _d  is not None else "—",
                "Delta %":  f"{_dp:+.1f}%" if _dp is not None else "—",
            })
        if _diff_rows:
            st.dataframe(pd.DataFrame(_diff_rows).set_index("Metric"),
                         use_container_width=True)


# ─────────────────────────────────────────────
# SAVED PLOTS GALLERY
# ─────────────────────────────────────────────
def render_saved_plots_gallery():
    """Render the saved plots gallery below the custom plot."""
    _saved = st.session_state.get("saved_plots", [])
    if not _saved:
        return

    st.divider()
    st.subheader(f"📌 Saved Plots ({len(_saved)}) — these will appear in the PDF report")

    for _idx, _sp in enumerate(_saved):
        _ga, _gb = st.columns([5, 1])
        _ga.markdown(f"**{_idx + 1}. {_sp['title']}**  "
                     + (f"`{_sp['x']} × {_sp['y']}`" if _sp.get("x") else ""))
        if _gb.button("🗑 Remove", key=f"rm_plot_{_idx}", use_container_width=True):
            st.session_state["saved_plots"].pop(_idx)
            st.rerun()
        st.image(_sp["png"], width=700)
        st.divider()


# ─────────────────────────────────────────────
# UPDATE PARAMETERS
# ─────────────────────────────────────────────
def render_update_parameters(df: pd.DataFrame, filename: str,
                              logs_dir: Path, mode: str):
    """Render the Update Parameters section at the bottom."""
    st.divider()
    st.subheader("💾 Save to Library")

    _db_row_save = fetch_run(filename) if filename else None
    sav1, sav2   = st.columns([3, 1])
    display_name = sav1.text_input(
        "Run name (for library)",
        value=_db_row_save["display_name"]
              if _db_row_save else filename.rsplit(".", 1)[0] if filename else "",
        key="display_name",
        placeholder="e.g. Motor 4 – RUN 1 – Step Test"
    )

    if sav2.button("💾 Update Parameters", type="primary", use_container_width=True):
        _ip_save = st.session_state.get(f"ip_{filename}", {})
        if not _ip_save and _db_row_save and _db_row_save.get("init_params"):
            try:
                _ip_save = json.loads(_db_row_save["init_params"])
            except Exception:
                _ip_save = {}
        if not _ip_save:
            from python_functions.data_pipeline import default_init_params
            _ip_save = default_init_params(df)

        dest_path   = logs_dir / filename
        stats       = compute_stats(df)
        test_date   = extract_test_date(filename)
        _cur_folder = _db_row_save["folder"] \
                      if _db_row_save and _db_row_save.get("folder") \
                      else "Uncategorised"
        save_run(filename, display_name or filename, test_date,
                 dest_path, stats, _ip_save, folder=_cur_folder)
        st.success("✅ Parameters updated in library.")


# ─────────────────────────────────────────────
# DOWNLOADS
# ─────────────────────────────────────────────
def render_downloads(df: pd.DataFrame, filename: str, show_meas: bool = False):
    """Render the CSV and PDF download section."""
    st.divider()
    st.subheader("Downloads")

    base_name = filename.rsplit(".", 1)[0] if filename else "log"
    dl1, dl2  = st.columns(2)

    # ── CSV ──
    dl1.download_button(
        "⬇️  Download cleaned CSV",
        df.to_csv(index=False).encode("utf-8"),
        f"{base_name}_cleaned.csv",
        "text/csv",
        use_container_width=True,
    )

    # ── PDF ──
    if dl2.button("📄 Generate & Download PDF Report",
                  type="primary", use_container_width=True):
        with st.spinner("Building PDF report…"):

            _db_row_dl = fetch_run(filename) if filename else None
            _ip_dl     = {}
            if _db_row_dl and _db_row_dl.get("init_params"):
                try:
                    _ip_dl = json.loads(_db_row_dl["init_params"])
                except Exception:
                    pass

            # Smart start/end detection
            _rpm_ok = "RPM" in df.columns and df["RPM"].notna().any()
            if _rpm_ok:
                _rpm_max  = df["RPM"].max()
                _start_df = df[df["RPM"] >= 50]
                _start_row = _start_df.iloc[0] if len(_start_df) else df.iloc[0]
                _end_df   = df[df["RPM"] >= 0.90 * _rpm_max]
                _end_row  = _end_df.iloc[-1] if len(_end_df) else df.iloc[-1]
            else:
                _start_row = df.iloc[0]
                _end_row   = df.iloc[-1]

            def _col(row, col, fmt=".2f"):
                try:
                    v = float(row[col])
                    return f"{v:{fmt}}" if pd.notna(v) else ""
                except Exception:
                    return ""

            def _maxcol(col, fmt=".2f"):
                try:
                    return f"{float(df[col].max()):{fmt}}" if col in df.columns else ""
                except Exception:
                    return ""

            _time_at_rpm = ""
            if _rpm_ok:
                _at_rpm = df[df["RPM"] >= 0.90 * df["RPM"].max()]
                if len(_at_rpm) > 1:
                    _time_at_rpm = f"{float(_at_rpm['Time'].max() - _at_rpm['Time'].min()):.1f}"

            pdf_data = {
                "filename":             filename or "",
                "test_date":            extract_test_date(filename) if filename else "",
                "test_time":            "",
                "saved_at":             _db_row_dl["saved_at"] if _db_row_dl else "",
                "duration_s":           f"{float(df['Time'].max()):.1f}" if "Time" in df.columns else "",
                "num_rows":             str(len(df)),
                "run_name":             _db_row_dl["display_name"]
                                        if _db_row_dl and _db_row_dl.get("display_name")
                                        else base_name,
                "operator":             "",
                "max_rpm":              _maxcol("RPM", ".0f"),
                "max_thrust":           _maxcol("Thrust", ".2f"),
                "max_torque":           f"{float(df['Torque'].abs().max()):.2f}"
                                        if "Torque" in df.columns else "",
                "max_esc_temp":         _maxcol("ESC_Temp", ".1f"),
                "max_motor_temp":       _maxcol("Motor_Temp", ".1f"),
                "max_esc_inlet_temp":   _maxcol("ESC_Inlet_Temp_C", ".1f"),
                "max_motor_inlet_temp": _maxcol("Motor_Inlet_Temp_C", ".1f"),
                "max_esc_pressure":     _maxcol("ESC_Pressure", ".3f"),
                "max_fin_inlet_temp":   _maxcol("Fin_Inlet_Temp_C", ".1f"),
                "max_fin_outlet_temp":  _maxcol("Fin_Outlet_Temp_C", ".1f"),
                "battery_voltage_post": _col(_end_row, "Voltage", ".2f"),
                "time_at_target_rpm":   _time_at_rpm,
                "mechanical_power":     "",
                "electrical_power":     "",
                "mechanical_efficiency": "",
                "overall_efficiency":    "",
                "init_esc_temp":         _col(_start_row, "ESC_Temp", ".1f"),
                "init_motor_temp":       _col(_start_row, "Motor_Temp", ".1f"),
                "esc_inlet_coolant":     _col(_start_row, "ESC_Inlet_Temp_C", ".1f"),
                "motor_inlet_coolant":   _col(_start_row, "Motor_Inlet_Temp_C", ".1f"),
                "esc_inlet_flow":        _col(_start_row, "ESC_Flow", ".2f"),
                "esc_inlet_pressure":    _col(_start_row, "ESC_Pressure", ".3f"),
                "battery_voltage":       _col(_start_row, "Voltage", ".2f"),
                **{k: str(v) for k, v in _ip_dl.items()},
            }
            _ip_sess = st.session_state.get(f"ip_{filename}", {})
            pdf_data.update({k: str(v) for k, v in _ip_sess.items()})

            # Test parameter check
            _tpc = st.session_state.get(f"tpc_{filename}")
            if not _tpc and _db_row_dl and _db_row_dl.get("test_param_check"):
                try:
                    _tpc = json.loads(_db_row_dl["test_param_check"])
                except Exception:
                    _tpc = None
            if _tpc:
                pdf_data["test_param_check"] = _tpc

            # Efficiency results — only if Measurable Parameters was toggled on
            if show_meas:
                # Base electrical power from raw df
                if "Voltage" in df.columns and "Current" in df.columns:
                    pdf_data["electrical_power"] = f"{float(df['Voltage'].mean() * df['Current'].mean()):.0f}"
                # Override with more precise eff_results if user ran the calculator
                _eff = st.session_state.get("eff_results", {})
                if _eff:
                    _pm = _eff.get("P_mech", (None, None))[0]
                    _pd = _eff.get("P_DC",   (None, None))[0]
                    _eo = _eff.get("eta_overall", (None, None))[0]
                    _em = _eff.get("eta_mech",    (None, None))[0]
                    if _pm: pdf_data["mechanical_power"]      = f"{_pm:.0f}"
                    if _pd: pdf_data["electrical_power"]      = f"{_pd:.0f}"
                    if _eo: pdf_data["overall_efficiency"]    = f"{_eo:.4f}"
                    if _em: pdf_data["mechanical_efficiency"] = f"{_em:.2f}"

            # Saved plots
            _saved    = st.session_state.get("saved_plots", [])
            chart_imgs = [(p["title"], p["png"]) for p in _saved]
            _rname    = pdf_data["run_name"]
            pdf_bytes = build_pdf_report(pdf_data, chart_imgs, _rname)

        st.download_button(
            "⬇️  Download PDF Report",
            pdf_bytes,
            f"{base_name}_report.pdf",
            "application/pdf",
            use_container_width=True,
        )
        st.success(f"✅ PDF ready — {len(chart_imgs)} saved plot(s) included.")