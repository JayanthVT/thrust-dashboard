"""
view_dashboard.py — Dashboard view sections
Contains: Test Summary, Measurable Parameters & Efficiency, Initial Parameters.
Called from app.py after data is loaded and cleaned.
"""

import json
import numpy as np
import pandas as pd
import streamlit as st

from python_functions.db import fetch_run, save_run
from python_functions.data_pipeline import (
    compute_stats, default_init_params, extract_test_date
)


def _fmt_time(secs):
    return f"{int(secs // 60)}m {secs % 60:.1f}s"


def _safe(row, col):
    try:
        v = row[col]
        return float(v) if pd.notna(v) else None
    except Exception:
        return None


# ─────────────────────────────────────────────
# TEST SUMMARY
# ─────────────────────────────────────────────
def render_test_summary(df: pd.DataFrame):
    """Render test summary — 4 key metrics then RPM lookup."""
    st.subheader("Test summary")

    if "RPM" not in df.columns or df.empty:
        _mcols = st.columns(4)
        _mcols[0].metric("Peak Thrust",   f"{df['Thrust'].max():.1f} N"    if "Thrust"   in df.columns else "—")
        _mcols[1].metric("Max ESC Temp",  f"{df['ESC_Temp'].max():.1f} °C" if "ESC_Temp"  in df.columns else "—")
        _mcols[2].metric("Max Motor Temp",f"{df['Motor_Temp'].max():.1f} °C" if "Motor_Temp" in df.columns else "—")
        _mcols[3].metric("Duration",      f"{df['Time'].max():.1f} s")
        return

    # ── Row 1: 5 key metrics ──
    _rpm_idx      = df["RPM"].idxmax()
    _pr           = df.loc[_rpm_idx]
    _pr_rpm       = _pr["RPM"]
    _peak_thrust  = df["Thrust"].max()        if "Thrust"     in df.columns else None
    _max_esc      = df["ESC_Temp"].max()      if "ESC_Temp"   in df.columns else None
    _max_motor    = df["Motor_Temp"].max()    if "Motor_Temp" in df.columns else None
    _peak_torque  = df["Torque"].abs().max()  if "Torque"     in df.columns else None

    _r1 = st.columns(5)
    _r1[0].metric("Peak RPM",       f"{int(_pr_rpm):,}")
    _r1[1].metric("Peak Thrust",    f"{_peak_thrust:.1f} N"  if _peak_thrust is not None else "—",
                  help="Maximum thrust over full run")
    _r1[2].metric("Peak Torque",    f"{_peak_torque:.2f} Nm" if _peak_torque is not None else "—",
                  help="Maximum absolute torque over full run")
    _r1[3].metric("Max ESC Temp",   f"{_max_esc:.1f} °C"    if _max_esc    is not None else "—")
    _r1[4].metric("Max Motor Temp", f"{_max_motor:.1f} °C"  if _max_motor  is not None else "—")

    st.divider()

    # ── RPM lookup ──
    st.caption("🔍 **RPM lookup** — type any RPM to see values at that operating point")
    _lc1, _lc2 = st.columns([1, 1])
    _lookup_rpm = _lc1.number_input(
        "Target RPM", min_value=0, max_value=int(_pr_rpm),
        value=int(_pr_rpm), step=50, key="rpm_lookup"
    )
    _tol = _lc2.number_input(
        "Tolerance ±", min_value=1, max_value=200,
        value=25, step=5, key="rpm_tol",
        help="Rows within ±tolerance RPM are averaged"
    )
    _band = df[(df["RPM"] >= _lookup_rpm - _tol) & (df["RPM"] <= _lookup_rpm + _tol)]

    if len(_band) == 0:
        st.warning(f"No data within ±{_tol} RPM of {_lookup_rpm}.")
    else:
        _lv   = _band["Voltage"].mean()         if "Voltage"    in _band.columns else None
        _li   = _band["Current"].mean()         if "Current"    in _band.columns else None
        _lpe  = (_lv * _li)                    if (_lv and _li)                  else None
        _lt   = _band["Thrust"].mean()          if "Thrust"     in _band.columns else None
        _ltor = _band["Torque"].abs().mean()    if "Torque"     in _band.columns else None
        _lmt  = _band["Motor_Temp"].mean()      if "Motor_Temp" in _band.columns else None
        _let  = _band["ESC_Temp"].mean()        if "ESC_Temp"   in _band.columns else None
        _lrpm = _band["RPM"].mean()

        st.caption(f"Mean over {len(_band)} rows  |  actual RPM: {_lrpm:.1f}")
        _lr = st.columns(7)
        _lr[0].metric("Actual RPM",       f"{_lrpm:.1f}")
        _lr[1].metric("DC Voltage",       f"{_lv:.1f} V"   if _lv   is not None else "—")
        _lr[2].metric("Current",          f"{_li:.1f} A"   if _li   is not None else "—")
        _lr[3].metric("Electrical Power", f"{_lpe:.0f} W"  if _lpe  is not None else "—",
                      help="V × I")
        _lr[4].metric("Thrust",           f"{_lt:.1f} N"   if _lt   is not None else "—")
        _lr[5].metric("Torque",           f"{_ltor:.2f} Nm" if _ltor is not None else "—",
                      help="|Torque| — absolute value of mean torque in band")
        _lr[6].metric("Motor / ESC Temp",
                      f"{_lmt:.1f} / {_let:.1f} °C"
                      if (_lmt is not None and _let is not None) else "—")


# ─────────────────────────────────────────────
# MEASURABLE PARAMETERS & EFFICIENCY
# ─────────────────────────────────────────────
def render_measurable_parameters(df: pd.DataFrame):
    """Render the Measurable Parameters & Efficiency expander."""
    with st.expander("📐 Measurable Parameters & Efficiency", expanded=False):

        has_thrust  = "Thrust"  in df.columns and (df["Thrust"].abs() > 0).any()
        has_torque  = "Torque"  in df.columns and (df["Torque"].abs() > 0).any()
        has_rpm     = "RPM"     in df.columns and df["RPM"].notna().any()
        has_dc_volt = "Voltage" in df.columns
        has_curr    = "Current" in df.columns

        # ── Auto-detect steady-state window ──
        st.markdown("**Steady-state window**")
        st.caption("Auto-detected from lowest RPM variance — verify and edit if needed.")

        _auto_s, _auto_e = 0.0, float(df["Time"].max())
        if has_rpm and len(df) > 50:
            _gm  = df[df["RPM"] > 0]["RPM"].mean()
            _thr = 0.01 * _gm
            _dt  = df["Time"].diff().median()
            _wr  = max(int(5.0 / _dt), 10) if _dt > 0 else 50
            _rs  = df["RPM"].rolling(_wr, center=True).std().fillna(9999)
            _sm  = _rs < _thr
            _bl = _bs = _cl = _cs = 0
            for _i, _st in enumerate(_sm):
                if _st:
                    if _cl == 0: _cs = _i
                    _cl += 1
                    if _cl > _bl: _bl, _bs = _cl, _cs
                else:
                    _cl = 0
            if _bl > 10:
                _auto_s = float(df["Time"].iloc[_bs])
                _auto_e = float(df["Time"].iloc[min(_bs + _bl - 1, len(df) - 1)])

        _sw1, _sw2 = st.columns(2)
        win_start = _sw1.number_input("Window start (s)", value=round(_auto_s, 1),
                                       min_value=0.0,
                                       max_value=float(df["Time"].max()),
                                       step=0.5, key="eff_ws")
        win_end   = _sw2.number_input("Window end (s)",   value=round(_auto_e, 1),
                                       min_value=0.0,
                                       max_value=float(df["Time"].max()),
                                       step=0.5, key="eff_we")

        # ── Manual inputs ──
        _dfw_preview = df[(df["Time"] >= win_start) & (df["Time"] <= win_end)]
        st.divider()
        st.markdown("**Manual inputs**")
        _inp1, _inp2 = st.columns(2)

        _torque_default = float(_dfw_preview["Torque"].abs().mean()) \
                          if has_torque and len(_dfw_preview) else 0.0
        _thrust_default = float(_dfw_preview["Thrust"].abs().mean()) \
                          if has_thrust and len(_dfw_preview) else 0.0

        inp_torque = _inp1.number_input(
            "Torque (Nm) — load cell", value=_torque_default,
            step=0.5, key="inp_torque",
            help="Auto-filled from log mean over window — edit to override")
        inp_thrust = _inp2.number_input(
            "Thrust (N) — load cell", value=_thrust_default,
            step=1.0, key="inp_thrust",
            help="Auto-filled from log mean over window — edit to override")

        _inp1.caption(f"Log mean: {_torque_default:.2f} Nm" if has_torque else "Not in log")
        _inp2.caption(f"Log mean: {_thrust_default:.2f} N"  if has_thrust else "Not in log")

        st.divider()

        if st.button("⚙️  Calculate", type="primary", key="eff_calc"):
            _dfw = df[(df["Time"] >= win_start) & (df["Time"] <= win_end)].copy()
            if len(_dfw) < 5:
                st.error("Window too short — fewer than 5 rows.")
            else:
                _dfw["omega"]  = _dfw["RPM"] * (2 * np.pi / 60)
                _dfw["P_mech"] = inp_torque * _dfw["omega"]
                _dfw["V_DC"]   = _dfw["Voltage"] if has_dc_volt else np.nan
                _dfw["I_DC"]   = _dfw["Current"]  if has_curr    else np.nan
                _dfw["P_DC"]   = _dfw["V_DC"] * _dfw["I_DC"]
                _dfw["T_g"]    = inp_thrust * 101.972
                _dfw["eta_overall"] = np.where(
                    _dfw["P_DC"] > 0, _dfw["T_g"] / _dfw["P_DC"], np.nan)
                _dfw["eta_mech"] = np.where(
                    _dfw["P_DC"] > 0,
                    (_dfw["P_mech"] / _dfw["P_DC"]) * 100, np.nan)

                def _s(series):
                    s = pd.to_numeric(series, errors="coerce").dropna()
                    return (float(s.mean()), float(s.std())) if len(s) else (None, None)

                st.session_state["eff_results"] = {
                    "torque":      inp_torque,
                    "thrust":      inp_thrust,
                    "win_start":   win_start,
                    "win_end":     win_end,
                    "omega":       _s(_dfw["omega"]),
                    "P_mech":      _s(_dfw["P_mech"]),
                    "V_DC":        _s(_dfw["V_DC"]),
                    "I_DC":        _s(_dfw["I_DC"]),
                    "P_DC":        _s(_dfw["P_DC"]),
                    "eta_overall": _s(pd.Series(_dfw["eta_overall"])),
                    "eta_mech":    _s(pd.Series(_dfw["eta_mech"])),
                    "n_rows":      len(_dfw),
                }

        # ── Display results from session state ──
        _res = st.session_state.get("eff_results")
        if _res:
            def _met(label, key, unit, fmt=".2f", help_txt=""):
                mv, sv = _res.get(key, (None, None))
                if mv is None:
                    st.metric(label, "—", help=help_txt)
                else:
                    st.metric(label, f"{mv:{fmt}} {unit}",
                              delta=f"±{sv:{fmt}} σ",
                              delta_color="off", help=help_txt)

            st.caption(
                f"Results for window {_res['win_start']:.1f}s → {_res['win_end']:.1f}s  "
                f"({_res['n_rows']:,} rows)  |  "
                f"Torque: {_res['torque']:.2f} Nm  |  "
                f"Thrust: {_res['thrust']:.2f} N"
            )

            st.markdown("#### 🔧 Mechanical")
            _m1, _m2, _m3 = st.columns(3)
            with _m1:
                st.metric("Shaft Torque (input)", f"{_res['torque']:.2f} Nm",
                          help="Entered above — auto-filled from log mean")
            with _m2:
                _met("Angular Velocity ω", "omega", "rad/s", ".2f",
                     help_txt="ω = RPM × 2π / 60")
            with _m3:
                _met("Mechanical Power", "P_mech", "W", ".0f",
                     help_txt="P_mech = Torque × ω")

            st.divider()
            st.markdown("#### ⚡ Electrical")
            _e1, _e2, _e3 = st.columns(3)
            with _e1:
                _met("DC Bus Voltage", "V_DC", "V", ".2f",
                     help_txt="Mean DC bus voltage over window")
            with _e2:
                _met("DC Bus Current", "I_DC", "A", ".2f",
                     help_txt="Mean DC bus current over window")
            with _e3:
                _met("Electrical Power (DC)", "P_DC", "W", ".0f",
                     help_txt="P_elec = V_DC × I_DC")

            st.divider()
            st.markdown("#### 📊 Efficiency")
            _f1, _f2 = st.columns(2)
            with _f1:
                _met("Overall Efficiency", "eta_overall", "g/W", ".4f",
                     help_txt="= (Thrust_N × 101.972) / (V_DC × I_DC)")
            with _f2:
                _met("Mechanical Efficiency", "eta_mech", "%", ".2f",
                     help_txt="= (P_mech / P_elec) × 100")

        st.divider()


# ─────────────────────────────────────────────
# INITIAL PARAMETERS
# ─────────────────────────────────────────────
def render_initial_parameters(df: pd.DataFrame, filename: str):
    """
    Render the Initial Parameters expander.
    Returns the current ip dict so app.py can pass it to the update button.
    """
    _db_row = fetch_run(filename) if filename else None
    _saved_ip = {}
    if _db_row and _db_row.get("init_params"):
        try:
            _saved_ip = json.loads(_db_row["init_params"])
        except Exception:
            pass

    _defaults = default_init_params(df)
    ip = {**_defaults, **_saved_ip}

    with st.expander("📋 Initial Parameters", expanded=False):
        if _db_row:
            st.caption("✅ Loaded from library — edit and click **Update Parameters** to save.")
        else:
            st.caption("Auto-filled from first log row — edit as needed.")

        st.markdown("**Reservoir**")
        rc1, rc2, rc3 = st.columns(3)
        ip["res_capacity"]    = rc1.text_input("Capacity (L)",     value=ip["res_capacity"],    key="rc1")
        ip["res_composition"] = rc2.text_input("Composition",      value=ip["res_composition"], key="rc2")
        ip["res_temperature"] = rc3.text_input("Temperature (°C)", value=ip["res_temperature"], key="rc3")

        st.markdown("**Duty Cycle & Flowrate**")
        ip["duty_cycle"] = st.text_input("Duty cycle & flowrate", value=ip["duty_cycle"], key="dc")

        st.markdown("**Temperature (°C)**")
        tc1, tc2, tc3, tc4, tc5 = st.columns(5)
        ip["init_esc_temp"]       = tc1.text_input("Initial ESC",        value=ip["init_esc_temp"],       key="t1")
        ip["init_motor_temp"]     = tc2.text_input("Initial Motor",       value=ip["init_motor_temp"],     key="t2")
        ip["ambient_temp"]        = tc3.text_input("Ambient",             value=ip["ambient_temp"],        placeholder="e.g. 31", key="t3")
        ip["esc_inlet_coolant"]   = tc4.text_input("ESC Inlet Coolant",   value=ip["esc_inlet_coolant"],   key="t4")
        ip["motor_inlet_coolant"] = tc5.text_input("Motor Inlet Coolant", value=ip["motor_inlet_coolant"], key="t5")

        st.markdown("**Flowrate (LPM)**")
        fc1, fc2 = st.columns(2)
        ip["esc_inlet_flow"]   = fc1.text_input("ESC Inlet",   value=ip["esc_inlet_flow"],   key="f1")
        ip["motor_inlet_flow"] = fc2.text_input("Motor Inlet", value=ip["motor_inlet_flow"],  placeholder="-", key="f2")

        st.markdown("**Pressure (Bar)**")
        pc1, pc2 = st.columns(2)
        ip["esc_inlet_pressure"]   = pc1.text_input("ESC Inlet",   value=ip["esc_inlet_pressure"],   key="p1")
        ip["motor_inlet_pressure"] = pc2.text_input("Motor Inlet", value=ip["motor_inlet_pressure"],  placeholder="-", key="p2")

        st.markdown("**Battery**")
        bc1, bc2, bc3 = st.columns(3)
        ip["battery_voltage"] = bc1.text_input("Voltage (V)", value=ip["battery_voltage"], key="b1")
        ip["battery_soc"]     = bc2.text_input("SOC",         value=ip["battery_soc"],     key="b2")
        ip["battery_soh"]     = bc3.text_input("SOH",         value=ip["battery_soh"],     key="b3")

        st.markdown("**Fintube**")
        fi1, fi2 = st.columns(2)
        ip["fin_inlet_temp"]  = fi1.text_input("Inlet Temp (°C)",  value=ip["fin_inlet_temp"],  key="fi1")
        ip["fin_outlet_temp"] = fi2.text_input("Outlet Temp (°C)", value=ip["fin_outlet_temp"], key="fi2")

        st.markdown("**Notes**")
        ip["notes"] = st.text_area(
            "Test notes", value=ip["notes"], height=80,
            placeholder="Any observations, anomalies, or setup details…",
            key="notes"
        )

    # Store in session state so it's accessible outside the expander
    st.session_state[f"ip_{filename}"] = ip
    return ip