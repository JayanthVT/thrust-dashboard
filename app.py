"""
Thrust Test Rig — Engineering Dashboard  v3
- SQLite log library with persistent memory
- File explorer sidebar with date search
- Editable initial parameters that survive restarts
- All charts from v2 preserved
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import io, os, sqlite3, shutil, json, re
import importlib.util
from datetime import datetime, date
from pathlib import Path

# ── PDF generation ──
from pdf_engine import build_pdf_report, fig_to_png, TEMPLATE_PATH, LOGO_PATH

# ─────────────────────────────────────────────
# DEPENDENCY CHECK
# ─────────────────────────────────────────────
_missing = [p for p in ["openpyxl", "matplotlib"] if importlib.util.find_spec(p) is None]
if _missing:
    st.error(
        f"Missing required package(s): **{', '.join(_missing)}**\n\n"
        f"Run:  `pip install {' '.join(_missing)}`  then restart Streamlit."
    )
    st.stop()

# ─────────────────────────────────────────────
# PATHS  (everything lives next to this script)
# ─────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent
LOGS_DIR  = BASE_DIR / "logs"
DB_PATH   = BASE_DIR / "thrust_logs.db"
LOGS_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────
# DATABASE  SETUP
# ─────────────────────────────────────────────
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS runs (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                filename      TEXT UNIQUE,
                display_name  TEXT,
                test_date     TEXT,
                saved_at      TEXT,
                file_path     TEXT,
                -- summary stats
                max_thrust_n  REAL,
                max_rpm       REAL,
                max_power_w   REAL,
                max_current_a REAL,
                max_voltage_v REAL,
                max_esc_temp  REAL,
                max_motor_temp REAL,
                duration_s    REAL,
                num_rows      INTEGER,
                -- initial parameters (JSON blob)
                init_params   TEXT
            )
        """)
        conn.commit()

init_db()

def save_run(filename, display_name, test_date, file_path, stats, init_params):
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO runs
              (filename, display_name, test_date, saved_at, file_path,
               max_thrust_n, max_rpm, max_power_w, max_current_a,
               max_voltage_v, max_esc_temp, max_motor_temp, duration_s,
               num_rows, init_params)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(filename) DO UPDATE SET
              display_name   = excluded.display_name,
              test_date      = excluded.test_date,
              saved_at       = excluded.saved_at,
              file_path      = excluded.file_path,
              max_thrust_n   = excluded.max_thrust_n,
              max_rpm        = excluded.max_rpm,
              max_power_w    = excluded.max_power_w,
              max_current_a  = excluded.max_current_a,
              max_voltage_v  = excluded.max_voltage_v,
              max_esc_temp   = excluded.max_esc_temp,
              max_motor_temp = excluded.max_motor_temp,
              duration_s     = excluded.duration_s,
              num_rows       = excluded.num_rows,
              init_params    = excluded.init_params
        """, (
            filename, display_name, test_date,
            datetime.now().isoformat(timespec="seconds"),
            str(file_path),
            stats.get("max_thrust_n"), stats.get("max_rpm"),
            stats.get("max_power_w"), stats.get("max_current_a"),
            stats.get("max_voltage_v"), stats.get("max_esc_temp"),
            stats.get("max_motor_temp"), stats.get("duration_s"),
            stats.get("num_rows"),
            json.dumps(init_params),
        ))
        conn.commit()

def update_init_params(filename, init_params):
    with get_conn() as conn:
        conn.execute(
            "UPDATE runs SET init_params=?, saved_at=? WHERE filename=?",
            (json.dumps(init_params),
             datetime.now().isoformat(timespec="seconds"),
             filename)
        )
        conn.commit()

def fetch_all_runs(search_text="", date_from=None, date_to=None):
    q = "SELECT * FROM runs WHERE 1=1"
    params = []
    if search_text:
        q += " AND (display_name LIKE ? OR filename LIKE ?)"
        params += [f"%{search_text}%", f"%{search_text}%"]
    if date_from:
        q += " AND test_date >= ?"
        params.append(str(date_from))
    if date_to:
        q += " AND test_date <= ?"
        params.append(str(date_to))
    q += " ORDER BY test_date DESC, saved_at DESC"
    with get_conn() as conn:
        return [dict(r) for r in conn.execute(q, params).fetchall()]

def fetch_run(filename):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM runs WHERE filename=?", (filename,)).fetchone()
        return dict(row) if row else None

def delete_run(filename):
    row = fetch_run(filename)
    if row and row["file_path"] and Path(row["file_path"]).exists():
        Path(row["file_path"]).unlink(missing_ok=True)
    with get_conn() as conn:
        conn.execute("DELETE FROM runs WHERE filename=?", (filename,))
        conn.commit()




# ─────────────────────────────────────────────
# PLOTLY CHART HELPERS  (interactive, used on screen)
# Matplotlib kept separately for PDF only
# ─────────────────────────────────────────────
_PL_BG     = "#0d0f14"
_PL_PAPER  = "#13161e"
_PL_GRID   = "#1e2130"
_PL_TEXT   = "#c8ccd8"
_PL_TICK   = "#6b7280"
_PL_BORDER = "#2a2d3a"

def _pl_base_layout(title="", height=380):
    return dict(
        title=dict(text=title, font=dict(color=_PL_TEXT, size=12), x=0.01),
        plot_bgcolor=_PL_BG, paper_bgcolor=_PL_PAPER,
        font=dict(color=_PL_TEXT, family="monospace", size=11),
        legend=dict(bgcolor=_PL_PAPER, bordercolor=_PL_BORDER, borderwidth=1,
                    font=dict(color=_PL_TEXT)),
        margin=dict(l=60, r=60, t=45, b=45),
        hovermode="x unified",
        height=height,
    )

def _pl_xaxis(label="Time (s)"):
    return dict(title=dict(text=label, font=dict(color=_PL_TICK)),
                gridcolor=_PL_GRID, gridwidth=0.5,
                showline=True, linecolor=_PL_BORDER,
                tickfont=dict(color=_PL_TICK))

def _pl_yaxis(label, color):
    return dict(title=dict(text=label, font=dict(color=color)),
                tickfont=dict(color=color),
                gridcolor=_PL_GRID, gridwidth=0.5,
                showline=True, linecolor=_PL_BORDER)

def pl_single(df, y_col, color, ylabel, unit, title,
              df2=None, label1="Run A", label2="Run B"):
    """Single Y axis plotly chart. Optionally overlay df2 for comparison."""
    import plotly.graph_objects as _go
    fig = _go.Figure()
    # Downsample
    _step = max(1, len(df) // 5000)
    _df = df.iloc[::_step]
    fig.add_trace(_go.Scatter(
        x=_df["Time"], y=_df[y_col], mode="lines",
        name=f"{y_col} ({label1})",
        line=dict(color=color, width=1.6),
        hovertemplate=f"<b>Time</b>: %{{x:.2f}}s<br><b>{y_col}</b>: %{{y:.3f}} {unit}<extra>{label1}</extra>",
    ))
    if df2 is not None and y_col in df2.columns and "Time" in df2.columns:
        _step2 = max(1, len(df2) // 5000)
        _df2 = df2.iloc[::_step2]
        fig.add_trace(_go.Scatter(
            x=_df2["Time"], y=_df2[y_col], mode="lines",
            name=f"{y_col} ({label2})",
            line=dict(color="#38bdf8", width=1.4, dash="dash"),
            hovertemplate=f"<b>Time</b>: %{{x:.2f}}s<br><b>{y_col}</b>: %{{y:.3f}} {unit}<extra>{label2}</extra>",
        ))
    layout = _pl_base_layout(title)
    layout["xaxis"] = _pl_xaxis()
    layout["yaxis"] = _pl_yaxis(f"{ylabel} ({unit})", color)
    fig.update_layout(**layout)
    return fig

def pl_overlay(df, y1, y2, c1, c2, l1, l2, u1, u2, title,
               df2=None, label1="Run A", label2="Run B"):
    """Dual Y axis plotly chart. Optionally overlay df2 for comparison."""
    import plotly.graph_objects as _go
    fig = _go.Figure()
    _step = max(1, len(df) // 5000)
    _df = df.iloc[::_step]
    fig.add_trace(_go.Scatter(
        x=_df["Time"], y=_df[y1], mode="lines", name=f"{y1} ({label1})",
        line=dict(color=c1, width=1.6), yaxis="y1",
        hovertemplate=f"<b>{y1}</b>: %{{y:.2f}} {u1}<extra>{label1}</extra>",
    ))
    fig.add_trace(_go.Scatter(
        x=_df["Time"], y=_df[y2], mode="lines", name=f"{y2} ({label1})",
        line=dict(color=c2, width=1.4, dash="dash"), yaxis="y2",
        hovertemplate=f"<b>{y2}</b>: %{{y:.2f}} {u2}<extra>{label1}</extra>",
    ))
    if df2 is not None:
        _step2 = max(1, len(df2) // 5000)
        _df2 = df2.iloc[::_step2]
        if y1 in df2.columns:
            fig.add_trace(_go.Scatter(
                x=_df2["Time"], y=_df2[y1], mode="lines",
                name=f"{y1} ({label2})",
                line=dict(color=c1, width=1.2, dash="dot"), yaxis="y1",
                hovertemplate=f"<b>{y1}</b>: %{{y:.2f}} {u1}<extra>{label2}</extra>",
            ))
        if y2 in df2.columns:
            fig.add_trace(_go.Scatter(
                x=_df2["Time"], y=_df2[y2], mode="lines",
                name=f"{y2} ({label2})",
                line=dict(color=c2, width=1.0, dash="dot"), yaxis="y2",
                hovertemplate=f"<b>{y2}</b>: %{{y:.2f}} {u2}<extra>{label2}</extra>",
            ))
    layout = _pl_base_layout(title)
    layout["xaxis"]  = _pl_xaxis()
    layout["yaxis"]  = _pl_yaxis(f"{l1} ({u1})", c1)
    layout["yaxis2"] = dict(
        title=dict(text=f"{l2} ({u2})", font=dict(color=c2)),
        tickfont=dict(color=c2), overlaying="y", side="right",
        showgrid=False, showline=True, linecolor=_PL_BORDER)
    fig.update_layout(**layout)
    return fig

def pl_multi(df, cols_colors, title, unit="", df2=None, label1="Run A", label2="Run B"):
    """Multiple traces on one Y axis. Optionally overlay df2."""
    import plotly.graph_objects as _go
    fig = _go.Figure()
    _step = max(1, len(df) // 5000)
    _df = df.iloc[::_step]
    for col, color in cols_colors:
        if col in _df.columns:
            fig.add_trace(_go.Scatter(
                x=_df["Time"], y=_df[col], mode="lines",
                name=f"{col} ({label1})", line=dict(color=color, width=1.4),
                hovertemplate=f"<b>{col}</b>: %{{y:.2f}} {unit}<extra>{label1}</extra>",
            ))
    if df2 is not None:
        _step2 = max(1, len(df2) // 5000)
        _df2 = df2.iloc[::_step2]
        for col, color in cols_colors:
            if col in _df2.columns and "Time" in _df2.columns:
                fig.add_trace(_go.Scatter(
                    x=_df2["Time"], y=_df2[col], mode="lines",
                    name=f"{col} ({label2})",
                    line=dict(color=color, width=1.0, dash="dash"),
                    hovertemplate=f"<b>{col}</b>: %{{y:.2f}} {unit}<extra>{label2}</extra>",
                ))
    layout = _pl_base_layout(title)
    layout["xaxis"] = _pl_xaxis()
    layout["yaxis"] = _pl_yaxis(unit, _PL_TEXT)
    fig.update_layout(**layout)
    return fig

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
  .stMetric label { font-size: 0.72rem; letter-spacing: 0.08em; text-transform: uppercase; }
  section[data-testid="stSidebar"] { background: #0d0d0d; }
  section[data-testdata="stSidebar"] * { color: #e0e0e0 !important; }
  .debug-box {
    background: #1a1a2e; color: #a0d8ef; font-family: 'Space Mono', monospace;
    font-size: 0.75rem; padding: 1rem; border-radius: 6px;
    border-left: 3px solid #4fc3f7; margin: 0.5rem 0; white-space: pre-wrap;
  }
  .run-card {
    background: #13161e; border: 1px solid #2a2d3a; border-radius: 8px;
    padding: 10px 14px; margin-bottom: 8px; cursor: pointer;
  }
  .run-card:hover { border-color: #f97316; }
  .run-card-active { border-color: #f97316 !important; background: #1a1a2e !important; }
  .run-date { font-size: 0.7rem; color: #6b7280; font-family: 'Space Mono', monospace; }
  .run-name { font-size: 0.85rem; font-weight: 500; color: #e0e0e0; }
  .run-stats { font-size: 0.72rem; color: #9ca3af; margin-top: 3px; }
  .summary-card {
    background: #13161e; border: 1px solid #2a2d3a; border-radius: 10px;
    padding: 16px 20px; margin-bottom: 16px;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# COLUMN ALIAS MAP
# ─────────────────────────────────────────────
COLUMN_ALIASES = {
    "Time":         ["Timestamp", "Time", "time", "timestamp", "T", "elapsed"],
    "Thrust":       ["Net_Thrust_N", "Thrust", "thrust", "Force_N", "ThrustN"],
    "RPM":          ["Actual_RPM", "RPM", "rpm", "Motor_RPM", "actual_rpm"],
    "Cmd_RPM":      ["Commanded_RPM", "Cmd_RPM", "commanded_rpm", "Target_RPM"],
    "Motor_Temp":   ["Motor_Temp_C", "Motor_Temp", "motor_temp_c"],
    "ESC_Temp":     ["ESC_Temp_C", "ESC_Temp", "esc_temp_c", "esc_temp"],
    "Power":        ["Power_W", "Power", "power_w"],
    "Current":      ["Current_A", "Current", "current_a"],
    "Voltage":      ["DC_Voltage_V", "DC_Voltage", "Voltage", "dc_voltage_v"],
    "Torque":       ["ESC_Torque_Nm", "Torque", "esc_torque_nm"],
    "Total_Weight": ["Total_Weight_kg", "total_weight_kg"],
    "Accel_X":      ["Accel_X_g", "accel_x_g"],
    "Accel_Y":      ["Accel_Y_g", "accel_y_g"],
    "Accel_Z":      ["Accel_Z_g", "accel_z_g"],
    "ESC_Pressure": ["ESC_Inlet_Pressure_Bar", "esc_inlet_pressure_bar"],
    "ESC_Flow":     ["ESC_Inlet_Flow_Lpm", "esc_inlet_flow_lpm"],
    "Motor_Flow":   ["Motor_Flow_Lpm", "motor_flow_lpm"],
}

# ─────────────────────────────────────────────
# INGESTION HELPERS
# ─────────────────────────────────────────────
def load_file_from_path(path: Path, logs):
    name = path.name.lower()
    raw = path.read_bytes()
    return _parse_raw(raw, name, logs)

def load_file_from_upload(uploaded_file, logs):
    raw = uploaded_file.read()
    name = uploaded_file.name.lower()
    return _parse_raw(raw, name, logs)

def _parse_raw(raw, name, logs):
    if name.endswith((".xlsx", ".xlsm")):
        for engine in ["openpyxl", "calamine"]:
            try:
                xl = pd.ExcelFile(io.BytesIO(raw), engine=engine)
                df = pd.read_excel(io.BytesIO(raw), sheet_name=0, engine=engine)
                logs.append(f"✅ Excel loaded (engine={engine}) — {df.shape[0]} rows × {df.shape[1]} cols")
                return df
            except ImportError as e:
                logs.append(f"⚠️  engine={engine} not installed: {e}")
            except Exception as e:
                logs.append(f"⚠️  engine={engine} failed: {e}")
        logs.append("❌ Run:  pip install openpyxl")
        return None
    if name.endswith(".xls"):
        try:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=0, engine="xlrd")
            logs.append(f"✅ .xls loaded — {df.shape[0]} rows")
            return df
        except Exception as e:
            logs.append(f"❌ xlrd: {e}")
            return None
    for enc in ["utf-8", "latin1", "cp1252"]:
        for eng in ["c", "python"]:
            try:
                df = pd.read_csv(io.BytesIO(raw), encoding=enc, engine=eng, on_bad_lines="skip")
                logs.append(f"✅ CSV parsed enc={enc} eng={eng}")
                return df
            except Exception:
                pass
    return None

def normalize_columns(df, logs):
    rename_map = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in df.columns and canonical not in df.columns:
                rename_map[alias] = canonical
                logs.append(f"🔀 {alias} → {canonical}")
                break
    return df.rename(columns=rename_map)

def parse_time(df, logs):
    if "Time" not in df.columns:
        logs.append("❌ No Time column.")
        return df
    col = df["Time"]

    # Already datetime64
    if pd.api.types.is_datetime64_any_dtype(col):
        df["Time"] = (col - col.iloc[0]).dt.total_seconds()
        logs.append(f"✅ Datetime → elapsed s ({float(df['Time'].max()):.1f}s)")
        return df

    # Already numeric
    if pd.api.types.is_numeric_dtype(col):
        logs.append("✅ Time already numeric.")
        return df

    # Try datetime string parse
    parsed = pd.to_datetime(col, errors="coerce")
    if parsed.notna().sum() / max(len(df), 1) > 0.5:
        df["Time"] = (parsed - parsed.iloc[0]).dt.total_seconds()
        logs.append(f"✅ String → elapsed s ({float(df['Time'].max()):.1f}s)")
        return df

    # Try numeric coerce
    num = pd.to_numeric(col, errors="coerce")
    if num.notna().sum() / max(len(df), 1) > 0.5:
        df["Time"] = num
        logs.append("✅ Time coerced to numeric.")
        return df

    # Last resort — use row index as time proxy (assumes uniform sampling)
    logs.append("⚠️  Could not parse Time column — using row index as time proxy.")
    df["Time"] = np.arange(len(df), dtype=float)
    return df

def clean_and_drop(df, logs):
    for col in [c for c in df.columns if c != "Time"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    present = [c for c in ["Time", "Thrust", "RPM"] if c in df.columns]
    before = len(df)
    df = df.dropna(subset=present)
    if before - len(df):
        logs.append(f"🗑️  Dropped {before-len(df)} rows with NaN in {present}")
    return df

def extract_test_date(filename: str) -> str:
    """Pull YYYYMMDD from filename like Motor4_RUN1_20260416_114449, else today."""
    m = re.search(r"(\d{8})", filename)
    if m:
        raw = m.group(1)
        try:
            return datetime.strptime(raw, "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            pass
    return date.today().isoformat()

def compute_stats(df) -> dict:
    def safe_max(col):
        return float(df[col].max()) if col in df.columns else None
    return {
        "max_thrust_n":   safe_max("Thrust"),
        "max_rpm":        safe_max("RPM"),
        "max_power_w":    safe_max("Power"),
        "max_current_a":  safe_max("Current"),
        "max_voltage_v":  safe_max("Voltage"),
        "max_esc_temp":   safe_max("ESC_Temp"),
        "max_motor_temp": safe_max("Motor_Temp"),
        "duration_s":     float(df["Time"].max()) if "Time" in df.columns else None,
        "num_rows":       len(df),
    }

def default_init_params(df) -> dict:
    t0 = df.iloc[0]
    def g(col, fmt="{:.2f}"):
        try:
            v = t0[col]
            return fmt.format(float(v)) if pd.notna(v) else ""
        except Exception:
            return ""
    return {
        "res_capacity":          "10",
        "res_composition":       "1:1 Glycol:Distilled water",
        "res_temperature":       g("Motor_Inlet_Temp_C", "{:.1f}"),
        "duty_cycle":            "70% ~6lpm",
        "init_esc_temp":         g("ESC_Temp",   "{:.1f}"),
        "init_motor_temp":       g("Motor_Temp",  "{:.1f}"),
        "ambient_temp":          "",
        "esc_inlet_coolant":     g("ESC_Inlet_Temp_C",   "{:.1f}"),
        "motor_inlet_coolant":   g("Motor_Inlet_Temp_C", "{:.1f}"),
        "esc_inlet_flow":        g("ESC_Flow",   "{:.2f}"),
        "motor_inlet_flow":      "-",
        "esc_inlet_pressure":    g("ESC_Pressure", "{:.3f}"),
        "motor_inlet_pressure":  "-",
        "battery_voltage":       g("Voltage", "{:.2f}"),
        "battery_soc":           "-%",
        "battery_soh":           "100%",
        "fin_inlet_temp":        g("Fin_Inlet_Temp_C",  "{:.1f}"),
        "fin_outlet_temp":       g("Fin_Outlet_Temp_C", "{:.1f}"),
        "notes":                 "",
    }

# ─────────────────────────────────────────────
# PLOT STYLE
# ─────────────────────────────────────────────
DARK = {
    "figure.facecolor": "#0d0f14", "axes.facecolor": "#13161e",
    "axes.edgecolor": "#2a2d3a", "axes.labelcolor": "#c8ccd8",
    "axes.grid": True, "grid.color": "#1e2130",
    "grid.linestyle": "--", "grid.linewidth": 0.5,
    "xtick.color": "#6b7280", "ytick.color": "#6b7280",
    "text.color": "#c8ccd8", "legend.facecolor": "#13161e",
    "legend.edgecolor": "#2a2d3a",
}

def make_single_plot(df, y_col, color, ylabel, unit, title):
    with plt.style.context(DARK):
        fig, ax = plt.subplots(figsize=(9, 3.2))
        ax.plot(df["Time"], df[y_col], color=color, linewidth=1.4, alpha=0.9)
        ax.fill_between(df["Time"], df[y_col], alpha=0.10, color=color)
        idx = df[y_col].idxmax()
        px, py = df.loc[idx, "Time"], df.loc[idx, y_col]
        ax.annotate(f"Peak {py:.2f} {unit}", xy=(px, py), xytext=(px, py*0.78),
                    arrowprops=dict(arrowstyle="->", color=color, lw=1.0),
                    color=color, fontsize=8, fontfamily="monospace")
        ax.set_title(title, fontsize=10, fontweight="bold", pad=8)
        ax.set_xlabel("Time (s)", fontsize=8)
        ax.set_ylabel(f"{ylabel} ({unit})", fontsize=8)
        fig.tight_layout()
        return fig

def make_overlay_plot(df, y1, y2, c1, c2, l1, l2, u1, u2, title):
    with plt.style.context(DARK):
        fig, ax1 = plt.subplots(figsize=(12, 3.6))
        ax2 = ax1.twinx()
        ax1.plot(df["Time"], df[y1], color=c1, linewidth=1.4, label=l1)
        ax1.fill_between(df["Time"], df[y1], alpha=0.08, color=c1)
        ax2.plot(df["Time"], df[y2], color=c2, linewidth=1.2, linestyle="--", label=l2, alpha=0.85)
        ax1.set_xlabel("Time (s)", fontsize=8)
        ax1.set_ylabel(f"{l1} ({u1})", color=c1, fontsize=8)
        ax2.set_ylabel(f"{l2} ({u2})", color=c2, fontsize=8)
        ax1.tick_params(axis="y", colors=c1)
        ax2.tick_params(axis="y", colors=c2)
        lines = ax1.get_lines() + ax2.get_lines()
        ax1.legend(lines, [l.get_label() for l in lines], loc="upper left", fontsize=8)
        ax1.set_title(title, fontsize=10, fontweight="bold", pad=8)
        fig.tight_layout()
        return fig

# ─────────────────────────────────────────────
# SIDEBAR — FILE EXPLORER
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🚀 Thrust Test Rig")
    st.divider()

    sidebar_mode = st.radio("Mode", ["📂 Log Library", "⬆️ Upload New Log"],
                            label_visibility="collapsed")
    st.divider()

    if sidebar_mode == "📂 Log Library":
        st.markdown("**Search logs**")
        search_text = st.text_input("Name / keyword", placeholder="e.g. Motor4, RUN2…",
                                    label_visibility="collapsed")
        col_df, col_dt = st.columns(2)
        date_from = col_df.date_input("From", value=None, label_visibility="collapsed")
        date_to   = col_dt.date_input("To",   value=None, label_visibility="collapsed")

        runs = fetch_all_runs(search_text, date_from, date_to)

        if not runs:
            st.caption("No saved logs yet. Upload a log and hit **Save to Library**.")
            selected_filename = None
        else:
            st.caption(f"{len(runs)} run(s) found")
            selected_filename = st.session_state.get("selected_run")

            for run in runs:
                is_active = run["filename"] == selected_filename
                card_class = "run-card run-card-active" if is_active else "run-card"
                thrust_str = f"{run['max_thrust_n']:.0f}N" if run["max_thrust_n"] else "—"
                rpm_str    = f"{int(run['max_rpm'])}RPM"  if run["max_rpm"]        else "—"
                dur_str    = f"{run['duration_s']:.0f}s"  if run["duration_s"]     else "—"
                st.markdown(f"""
                <div class="{card_class}">
                  <div class="run-date">{run['test_date']} · saved {run['saved_at'][:10]}</div>
                  <div class="run-name">{run['display_name']}</div>
                  <div class="run-stats">↑{thrust_str} &nbsp;|&nbsp; ↻{rpm_str} &nbsp;|&nbsp; ⏱{dur_str}</div>
                </div>""", unsafe_allow_html=True)

                btn_col, del_col = st.columns([4, 1])
                if btn_col.button("Open", key=f"open_{run['filename']}",
                                  type="primary" if is_active else "secondary",
                                  use_container_width=True):
                    st.session_state["selected_run"] = run["filename"]
                    st.session_state["mode"] = "library"
                    st.rerun()
                if del_col.button("🗑", key=f"del_{run['filename']}",
                                  use_container_width=True, help="Delete this run"):
                    delete_run(run["filename"])
                    if st.session_state.get("selected_run") == run["filename"]:
                        st.session_state.pop("selected_run", None)
                    st.rerun()

    else:  # Upload mode
        uploaded = st.file_uploader(
            "Upload log file", type=["xlsx", "xlsm", "xls", "csv", "txt"],
            label_visibility="collapsed"
        )
        st.session_state["mode"] = "upload"
        st.session_state["uploaded_file"] = uploaded

    st.divider()
    # ── Compare mode ──
    st.divider()
    st.markdown("**Compare mode**")
    _compare_on = st.toggle("Compare two runs", value=False, key="compare_on")
    if _compare_on:
        _all_runs = fetch_all_runs()
        _run_opts = {r["display_name"]: r["filename"] for r in _all_runs}
        if len(_run_opts) < 2:
            st.caption("Need at least 2 saved runs to compare.")
        else:
            _cmp_name = st.selectbox("Compare against:", list(_run_opts.keys()),
                                     key="cmp_run_select")
            st.session_state["cmp_filename"] = _run_opts[_cmp_name]
    else:
        st.session_state.pop("cmp_filename", None)

    st.divider()
    show_debug = st.toggle("Debug log",        value=False)
    show_raw   = st.toggle("Raw data table",   value=False)
    st.divider()
    rpm_filter = st.slider("Min RPM filter", 0, 500, 0, step=50,
                           help="Exclude rows below this RPM (trims idle/spin-up)")

# ─────────────────────────────────────────────
# MAIN AREA
# ─────────────────────────────────────────────
st.title("Thrust Test Rig — Engineering Dashboard")

# ── Determine what to show ──
mode = st.session_state.get("mode", "upload")
df        = None
logs      = []
run_meta  = None   # db row if loaded from library
filename  = None

if mode == "library":
    sel = st.session_state.get("selected_run")
    if sel:
        run_meta = fetch_run(sel)
        if run_meta and run_meta["file_path"] and Path(run_meta["file_path"]).exists():
            df = load_file_from_path(Path(run_meta["file_path"]), logs)
            filename = run_meta["filename"]
        else:
            st.error("⚠️  The file for this log was moved or deleted. "
                     "Re-upload it to restore the charts.")
            st.stop()
    else:
        st.info("👈  Select a log from the library, or switch to **Upload New Log**.")
        st.stop()

elif mode == "upload":
    uploaded = st.session_state.get("uploaded_file")
    if not uploaded:
        st.info("👈  Upload a `.xlsx` or `.csv` log file to get started.")
        with st.expander("Expected column schema"):
            st.markdown("""
| Dashboard name | Source column |
|---|---|
| Time | `Timestamp` |
| Thrust | `Net_Thrust_N` |
| RPM | `Actual_RPM` |
| Cmd_RPM | `Commanded_RPM` |
| Motor/ESC Temp | `Motor_Temp_C / ESC_Temp_C` |
| Power / Current / Voltage | `Power_W / Current_A / DC_Voltage_V` |
| Torque | `ESC_Torque_Nm` |
| Accel X/Y/Z | `Accel_X_g / Accel_Y_g / Accel_Z_g` |
            """)
        st.stop()
    df = load_file_from_upload(uploaded, logs)
    filename = uploaded.name

if df is None:
    st.error("Could not parse the file.")
    st.code("\n".join(logs))
    st.stop()

# ── CLEAN ──
logs.append(f"📋 Raw: {df.shape[0]} rows × {df.shape[1]} cols")
df = normalize_columns(df, logs)
df = parse_time(df, logs)
df = clean_and_drop(df, logs)

if rpm_filter > 0 and "RPM" in df.columns:
    before = len(df)
    df = df[df["RPM"] >= rpm_filter].copy()
    logs.append(f"🔧 RPM filter ≥{rpm_filter}: {len(df)}/{before} rows kept")

_time_max = df['Time'].max()
try:
    _time_max_str = f"{float(_time_max):.1f}s"
except (ValueError, TypeError):
    _time_max_str = str(_time_max)
logs.append(f"✅ Final: {df.shape[0]} rows | Span: 0–{_time_max_str}")

# ── Load compare run if selected ──
df2        = None
run_meta2  = None
label_run1 = filename or "Run A"
label_run2 = "Compare run"
_cmp_fn = st.session_state.get("cmp_filename")
if _cmp_fn and _cmp_fn != filename:
    _cmp_meta = fetch_run(_cmp_fn)
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

if show_debug:
    with st.expander("🔍 Debug log", expanded=True):
        st.markdown("<div class='debug-box'>" + "\n".join(logs) + "</div>", unsafe_allow_html=True)

if df.empty:
    st.error("DataFrame is empty after cleaning.")
    st.stop()

# ─────────────────────────────────────────────
# SUMMARY CARD  (library mode shows this at top)
# ─────────────────────────────────────────────
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
        {run_meta['duration_s']:.1f}s duration &nbsp;|&nbsp;
        Max Thrust {run_meta['max_thrust_n']:.1f} N &nbsp;|&nbsp;
        Max RPM {int(run_meta['max_rpm']):,}
      </div>
    </div>
    """, unsafe_allow_html=True)

if not _compare_active:
    # ─────────────────────────────────────────────
    # METRICS
    # ─────────────────────────────────────────────
    st.subheader("Test summary")

    def _safe(row, col):
        try:
            v = row[col]
            return float(v) if pd.notna(v) else None
        except Exception:
            return None

    def _fmt_time(secs):
        return f"{int(secs//60)}m {secs%60:.1f}s"

    if "RPM" in df.columns and not df.empty:

        # ── ROW 1: Peak RPM — timestamp, RPM, electrical power ──
        _rpm_idx  = df["RPM"].idxmax()
        _pr       = df.loc[_rpm_idx]
        _pr_time  = _pr["Time"]
        _pr_rpm   = _pr["RPM"]
        _pr_volt  = _safe(_pr, "Voltage")
        _pr_curr  = _safe(_pr, "Current")
        _pr_pelec = (_pr_volt * _pr_curr) if (_pr_volt and _pr_curr) else None

        st.caption("🔴 **Peak RPM** — all values at the moment of max RPM")
        _r1 = st.columns(6)
        _r1[0].metric("Timestamp",        _fmt_time(_pr_time))
        _r1[1].metric("Max RPM",          f"{int(_pr_rpm):,}")
        _r1[2].metric("DC Voltage",       f"{_pr_volt:.1f} V"  if _pr_volt  else "—")
        _r1[3].metric("Current",          f"{_pr_curr:.1f} A"  if _pr_curr  else "—")
        _r1[4].metric("Electrical Power", f"{_pr_pelec:.0f} W" if _pr_pelec else "—",
                      help="Electrical Power = DC Voltage × Current")
        _r1[5].metric("Thrust" if "Thrust" in df.columns else "—",
                      f"{_safe(_pr,'Thrust'):.1f} N" if _safe(_pr,'Thrust') is not None else "—")

        st.divider()

        # ── ROW 2: Temperatures (max over full run) + Duration ──
        st.caption("🌡️ **Temperatures** — max over full run")
        _r2 = st.columns(4)
        _r2[0].metric("Max Motor Temp", f"{df['Motor_Temp'].max():.1f} °C" if "Motor_Temp" in df.columns else "—")
        _r2[1].metric("Max ESC Temp",   f"{df['ESC_Temp'].max():.1f} °C"   if "ESC_Temp"   in df.columns else "—")
        _r2[2].metric("Duration",       f"{df['Time'].max():.1f} s")
        _r2[3].metric("Data points",    f"{len(df):,}")

        st.divider()

        # ── ROW 3: RPM lookup — user types any RPM, get snapshot ──
        st.caption("🔍 **RPM lookup** — type any RPM to see values at that operating point")
        _lc1, _lc2, _lc3 = st.columns([1, 1, 4])
        _lookup_rpm = _lc1.number_input("Target RPM", min_value=0, max_value=int(df["RPM"].max()),
                                        value=int(_pr_rpm), step=50, key="rpm_lookup")
        _tol        = _lc2.number_input("Tolerance ±", min_value=1, max_value=200,
                                        value=25, step=5, key="rpm_tol",
                                        help="RPM band around target — rows within ±tolerance are averaged")

        _band = df[(df["RPM"] >= _lookup_rpm - _tol) & (df["RPM"] <= _lookup_rpm + _tol)]

        if len(_band) == 0:
            st.warning(f"No data found within ±{_tol} RPM of {_lookup_rpm}. Try a wider tolerance.")
        else:
            _lv   = _band["Voltage"].mean() if "Voltage" in _band.columns else None
            _li   = _band["Current"].mean() if "Current" in _band.columns else None
            _lpe  = (_lv * _li) if (_lv is not None and _li is not None) else None
            _lt   = _band["Thrust"].mean()    if "Thrust"    in _band.columns else None
            _lmt  = _band["Motor_Temp"].mean()if "Motor_Temp"in _band.columns else None
            _let  = _band["ESC_Temp"].mean()  if "ESC_Temp"  in _band.columns else None
            _lrpm = _band["RPM"].mean()

            st.caption(f"Showing mean over {len(_band)} rows  |  actual RPM mean: {_lrpm:.1f}")
            _lr = st.columns(6)
            _lr[0].metric("Actual RPM",        f"{_lrpm:.1f}")
            _lr[1].metric("DC Voltage",        f"{_lv:.1f} V"   if _lv   is not None else "—")
            _lr[2].metric("Current",           f"{_li:.1f} A"   if _li   is not None else "—")
            _lr[3].metric("Electrical Power",  f"{_lpe:.0f} W"  if _lpe  is not None else "—",
                          help="Electrical Power = DC Voltage × Current")
            _lr[4].metric("Thrust",            f"{_lt:.1f} N"   if _lt   is not None else "—")
            _lr[5].metric("Motor / ESC Temp",
                          f"{_lmt:.1f} / {_let:.1f} °C" if (_lmt is not None and _let is not None) else "—")

    else:
        _mcols = st.columns(4)
        _mcols[0].metric("Max Thrust",   f"{df['Thrust'].max():.1f} N"   if "Thrust"  in df.columns else "—")
        _mcols[1].metric("Max ESC Temp", f"{df['ESC_Temp'].max():.1f} °C"if "ESC_Temp"in df.columns else "—")
        _mcols[2].metric("Duration",     f"{df['Time'].max():.1f} s")
        _mcols[3].metric("Data points",  f"{len(df):,}")

    st.divider()

    # ─────────────────────────────────────────────
    # MEASURABLE PARAMETERS & EFFICIENCY
    # ─────────────────────────────────────────────
    with st.expander("📐 Measurable Parameters & Efficiency", expanded=False):

        has_thrust  = "Thrust"  in df.columns and (df["Thrust"].abs() > 0).any()
        has_torque  = "Torque"  in df.columns and (df["Torque"].abs() > 0).any()
        has_rpm     = "RPM"     in df.columns and df["RPM"].notna().any()
        has_dc_volt = "Voltage" in df.columns
        has_curr    = "Current" in df.columns

        # ── Steady-state window ──
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
                _auto_e = float(df["Time"].iloc[min(_bs + _bl - 1, len(df)-1)])

        _sw1, _sw2 = st.columns(2)
        win_start = _sw1.number_input("Window start (s)", value=round(_auto_s, 1),
                      min_value=0.0, max_value=float(df["Time"].max()), step=0.5, key="eff_ws")
        win_end   = _sw2.number_input("Window end (s)",   value=round(_auto_e, 1),
                      min_value=0.0, max_value=float(df["Time"].max()), step=0.5, key="eff_we")

        # ── Torque and Thrust inputs — above Calculate so they survive reruns ──
        _dfw_preview = df[(df["Time"] >= win_start) & (df["Time"] <= win_end)]
        st.divider()
        st.markdown("**Manual inputs**")
        _inp1, _inp2 = st.columns(2)

        _torque_default = float(_dfw_preview["Torque"].abs().mean()) if has_torque and len(_dfw_preview) else 0.0
        _thrust_default = float(_dfw_preview["Thrust"].abs().mean()) if has_thrust and len(_dfw_preview) else 0.0

        inp_torque = _inp1.number_input(
            "Torque (Nm) — load cell",
            value=_torque_default,
            step=0.5, key="inp_torque",
            help="Auto-filled from log mean over window — edit to override before hitting Calculate")
        inp_thrust = _inp2.number_input(
            "Thrust (N) — load cell",
            value=_thrust_default,
            step=1.0, key="inp_thrust",
            help="Auto-filled from log mean over window — edit to override before hitting Calculate")

        _inp1.caption(f"Log mean: {_torque_default:.2f} Nm" if has_torque else "Not in log")
        _inp2.caption(f"Log mean: {_thrust_default:.2f} N"  if has_thrust else "Not in log")

        st.divider()

        if st.button("⚙️  Calculate", type="primary", key="eff_calc"):

            _dfw = df[(df["Time"] >= win_start) & (df["Time"] <= win_end)].copy()

            if len(_dfw) < 5:
                st.error("Window too short — fewer than 5 rows. Adjust start/end.")
            else:
                # ── compute everything and store in session state ──
                _dfw["omega"]  = _dfw["RPM"] * (2 * np.pi / 60)
                _dfw["P_mech"] = inp_torque * _dfw["omega"]
                _dfw["V_DC"]   = _dfw["Voltage"] if has_dc_volt else np.nan
                _dfw["I_DC"]   = _dfw["Current"] if has_curr    else np.nan
                _dfw["P_DC"]   = _dfw["V_DC"] * _dfw["I_DC"]
                _dfw["T_g"]    = inp_thrust * 101.972
                _dfw["eta_overall"] = np.where(_dfw["P_DC"] > 0, _dfw["T_g"] / _dfw["P_DC"], np.nan)
                _dfw["eta_mech"]    = np.where(_dfw["P_DC"] > 0, (_dfw["P_mech"] / _dfw["P_DC"]) * 100, np.nan)

                def _s(series):
                    s = pd.to_numeric(series, errors="coerce").dropna()
                    return (float(s.mean()), float(s.std())) if len(s) else (None, None)

                st.session_state["eff_results"] = {
                    "torque":        inp_torque,
                    "thrust":        inp_thrust,
                    "win_start":     win_start,
                    "win_end":       win_end,
                    "omega":         _s(_dfw["omega"]),
                    "P_mech":        _s(_dfw["P_mech"]),
                    "V_DC":          _s(_dfw["V_DC"]),
                    "I_DC":          _s(_dfw["I_DC"]),
                    "P_DC":          _s(_dfw["P_DC"]),
                    "eta_overall":   _s(pd.Series(_dfw["eta_overall"])),
                    "eta_mech":      _s(pd.Series(_dfw["eta_mech"])),
                    "n_rows":        len(_dfw),
                }

        # ── Display results from session state (persists across reruns) ──
        _res = st.session_state.get("eff_results")
        if _res:
            def _met(label, key, unit, fmt=".2f", help_txt=""):
                mv, sv = _res.get(key, (None, None))
                if mv is None:
                    st.metric(label, "—", help=help_txt)
                else:
                    st.metric(label, f"{mv:{fmt}} {unit}",
                              delta=f"±{sv:{fmt}} σ", delta_color="off", help=help_txt)

            st.caption(f"Results for window {_res['win_start']:.1f}s → {_res['win_end']:.1f}s  "
                       f"({_res['n_rows']:,} rows)  |  "
                       f"Torque input: {_res['torque']:.2f} Nm  |  "
                       f"Thrust input: {_res['thrust']:.2f} N")

            st.markdown("#### 🔧 Mechanical")
            _m1, _m2, _m3 = st.columns(3)
            with _m1:
                st.metric("Shaft Torque (input)", f"{_res['torque']:.2f} Nm",
                          help="Torque entered above — used for P_mech calculation")
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
                     help_txt="Overall Efficiency = Thrust(g) / P_elec\n"
                              "= (Thrust_N × 101.972) / (V_DC × I_DC)\n"
                              "Uses manually entered Thrust value above")
            with _f2:
                _met("Mechanical Efficiency", "eta_mech", "%", ".2f",
                     help_txt="Mechanical Efficiency = P_mech / P_elec × 100\n"
                              "= (Torque × ω) / (V_DC × I_DC)\n"
                              "Drivetrain efficiency — ESC + motor losses combined")

    st.divider()

    # ─────────────────────────────────────────────
    # INITIAL PARAMETERS  (persistent via SQLite)
    # ─────────────────────────────────────────────

    # Load from DB if this run is already saved, else use defaults
    _db_row = fetch_run(filename) if filename else None
    if _db_row and _db_row.get("init_params"):
        try:
            _saved_ip = json.loads(_db_row["init_params"])
        except Exception:
            _saved_ip = {}
    else:
        _saved_ip = {}

    # Merge: DB values win over defaults (so previously saved fields are always restored)
    _defaults = default_init_params(df)
    ip = {**_defaults, **_saved_ip}

    with st.expander("📋 Initial Parameters", expanded=True):
        _is_saved = bool(_db_row)
        if _is_saved:
            st.caption("✅ Loaded from library — edit any field and click **Save to Library** to update.")
        else:
            st.caption("Auto-filled from first log row — edit as needed, then **Save to Library** to make permanent.")

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
        ip["notes"] = st.text_area("Test notes", value=ip["notes"], height=80,
                                   placeholder="Any observations, anomalies, or setup details…",
                                   key="notes")



# ─────────────────────────────────────────────
# CUSTOM PLOT
# ─────────────────────────────────────────────
try:
    import plotly.graph_objects as go
    _plotly_ok = True
except ImportError:
    _plotly_ok = False

st.divider()
st.subheader("📈 Custom Plot")

if not _plotly_ok:
    st.warning("Plotly not installed. Run:  `pip install plotly`  then restart Streamlit.")
else:
    _EXCLUDE = {"Thrust_0deg_kg","Thrust_90deg_kg","Thrust_180deg_kg",
                "Thrust_270deg_kg","Total_Weight"}
    _plot_cols = [c for c in df.columns if c not in _EXCLUDE
                  and pd.api.types.is_numeric_dtype(df[c])]

    _cp1, _cp2, _cp3, _cp4 = st.columns([2, 2, 2, 1])
    _x_default = "Time"   if "Time"   in _plot_cols else _plot_cols[0]
    _y_default = "Thrust" if "Thrust" in _plot_cols else (
                 "RPM"    if "RPM"    in _plot_cols else _plot_cols[1])
    _y2_opts   = ["None"] + _plot_cols

    x_col  = _cp1.selectbox("X axis",  _plot_cols,
                             index=_plot_cols.index(_x_default), key="plt_x")
    y_col  = _cp2.selectbox("Y axis",  _plot_cols,
                             index=_plot_cols.index(_y_default), key="plt_y")
    y2_col = _cp3.selectbox("Y2 axis (optional)",
                             _y2_opts, index=0, key="plt_y2")
    _plot_type = _cp4.radio("Type", ["Line", "Scatter"], key="plt_type")

    _pw1, _pw2, _pw3 = st.columns([2, 2, 3])
    _use_window = _pw3.checkbox("Filter time window", value=False, key="plt_win")
    if _use_window and "Time" in df.columns:
        _plt_tmin = _pw1.number_input("From (s)", value=0.0,
                                       max_value=float(df["Time"].max()),
                                       step=1.0, key="plt_tmin")
        _plt_tmax = _pw2.number_input("To (s)", value=float(df["Time"].max()),
                                       max_value=float(df["Time"].max()),
                                       step=1.0, key="plt_tmax")
        _df_plot  = df[(df["Time"] >= _plt_tmin) & (df["Time"] <= _plt_tmax)].copy()
        _df_plot2 = df2[(df2["Time"] >= _plt_tmin) & (df2["Time"] <= _plt_tmax)].copy() if df2 is not None and "Time" in df2.columns else df2
    else:
        _df_plot  = df.copy()
        _df_plot2 = df2.copy() if df2 is not None else None

    _step    = max(1, len(_df_plot) // 5000)
    _df_plot = _df_plot.iloc[::_step]
    _ds_note = f"({len(_df_plot):,} pts)"

    _DARK_BG = "#0d0f14"; _DARK_PAPER = "#13161e"
    _GRID_COL = "#1e2130"; _TEXT_COL = "#c8ccd8"
    _C1="#f97316"; _C2="#38bdf8"; _C3="#a78bfa"; _C4="#4ade80"
    _mode_str = "lines" if _plot_type == "Line" else "markers"
    _msz = 3 if _plot_type == "Scatter" else 4

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=_df_plot[x_col], y=_df_plot[y_col], mode=_mode_str,
        name=f"{y_col} ({label_run1})",
        line=dict(color=_C1, width=1.6) if _plot_type=="Line" else None,
        marker=dict(size=_msz, color=_C1), yaxis="y1",
        hovertemplate=f"<b>{x_col}</b>: %{{x:.3f}}<br><b>{y_col}</b>: %{{y:.3f}}<extra>{label_run1}</extra>",
    ))
    _has_y2 = y2_col != "None" and y2_col in _df_plot.columns
    if _has_y2:
        fig.add_trace(go.Scatter(
            x=_df_plot[x_col], y=_df_plot[y2_col], mode=_mode_str,
            name=f"{y2_col} ({label_run1})",
            line=dict(color=_C2, width=1.4, dash="dash") if _plot_type=="Line" else None,
            marker=dict(size=_msz, color=_C2), yaxis="y2",
            hovertemplate=f"<b>{x_col}</b>: %{{x:.3f}}<br><b>{y2_col}</b>: %{{y:.3f}}<extra>{label_run1}</extra>",
        ))
    if _df_plot2 is not None:
        _step2 = max(1, len(_df_plot2) // 5000)
        _df_plot2 = _df_plot2.iloc[::_step2]
        if x_col in _df_plot2.columns and y_col in _df_plot2.columns:
            fig.add_trace(go.Scatter(
                x=_df_plot2[x_col], y=_df_plot2[y_col], mode=_mode_str,
                name=f"{y_col} ({label_run2})",
                line=dict(color=_C3, width=1.4, dash="dot") if _plot_type=="Line" else None,
                marker=dict(size=_msz, color=_C3, symbol="diamond"), yaxis="y1",
                hovertemplate=f"<b>{x_col}</b>: %{{x:.3f}}<br><b>{y_col}</b>: %{{y:.3f}}<extra>{label_run2}</extra>",
            ))
        if _has_y2 and y2_col in _df_plot2.columns:
            fig.add_trace(go.Scatter(
                x=_df_plot2[x_col], y=_df_plot2[y2_col], mode=_mode_str,
                name=f"{y2_col} ({label_run2})",
                line=dict(color=_C4, width=1.2, dash="dot") if _plot_type=="Line" else None,
                marker=dict(size=_msz, color=_C4, symbol="diamond"), yaxis="y2",
                hovertemplate=f"<b>{x_col}</b>: %{{x:.3f}}<br><b>{y2_col}</b>: %{{y:.3f}}<extra>{label_run2}</extra>",
            ))

    _layout = dict(
        plot_bgcolor=_DARK_BG, paper_bgcolor=_DARK_PAPER,
        font=dict(color=_TEXT_COL, family="monospace", size=11),
        xaxis=dict(title=dict(text=x_col, font=dict(color=_TEXT_COL)),
                   gridcolor=_GRID_COL, gridwidth=0.5,
                   showline=True, linecolor="#2a2d3a",
                   tickfont=dict(color="#6b7280")),
        yaxis=dict(title=dict(text=y_col, font=dict(color=_C1)),
                   tickfont=dict(color=_C1),
                   gridcolor=_GRID_COL, gridwidth=0.5,
                   showline=True, linecolor="#2a2d3a"),
        legend=dict(bgcolor="#13161e", bordercolor="#2a2d3a", borderwidth=1,
                    font=dict(color=_TEXT_COL)),
        margin=dict(l=60, r=60, t=40, b=50),
        hovermode="x unified", height=450,
    )
    if _has_y2:
        _layout["yaxis2"] = dict(
            title=dict(text=y2_col, font=dict(color=_C2)),
            tickfont=dict(color=_C2), overlaying="y", side="right",
            showgrid=False, showline=True, linecolor="#2a2d3a")
    fig.update_layout(**_layout)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(f"X: **{x_col}**  |  Y1: **{y_col}**"
               + (f"  |  Y2: **{y2_col}**" if _has_y2 else "")
               + (f"  |  vs **{label_run2}**" if _df_plot2 is not None else "")
               + f"  |  {_ds_note}")

    # ── Save Plot button ──
    _sp1, _sp2 = st.columns([3, 1])
    _plot_title_input = _sp1.text_input(
        "Plot title for report",
        value=f"{y_col} vs {x_col}" + (f" & {y2_col}" if _has_y2 else ""),
        key="plot_title_input",
        label_visibility="collapsed",
        placeholder="Enter plot title…"
    )
    if _sp2.button("💾 Save Plot", type="primary", use_container_width=True, key="save_plot_btn"):
        # Render current selection as matplotlib PNG for PDF
        if "saved_plots" not in st.session_state:
            st.session_state["saved_plots"] = []

        # Build matplotlib version of current plot for PDF embedding
        with plt.style.context(DARK):
            _pdf_fig, _pdf_ax1 = plt.subplots(figsize=(11, 3.4))
            _step_pdf = max(1, len(_df_plot) // 5000)
            _dfp = _df_plot.iloc[::_step_pdf]
            _pdf_ax1.plot(_dfp[x_col], _dfp[y_col],
                          color=_C1, linewidth=1.4, label=y_col)
            _pdf_ax1.set_xlabel(x_col, fontsize=8)
            _pdf_ax1.set_ylabel(y_col, color=_C1, fontsize=8)
            _pdf_ax1.tick_params(axis="y", colors=_C1)
            if _has_y2:
                _pdf_ax2 = _pdf_ax1.twinx()
                _pdf_ax2.plot(_dfp[x_col], _dfp[y2_col],
                              color=_C2, linewidth=1.2,
                              linestyle="--", label=y2_col)
                _pdf_ax2.set_ylabel(y2_col, color=_C2, fontsize=8)
                _pdf_ax2.tick_params(axis="y", colors=_C2)
                _pdf_fig.legend(loc="upper left", fontsize=8,
                                bbox_to_anchor=(0.08, 0.95))
            _pdf_fig.tight_layout()
        _png = fig_to_png(_pdf_fig)

        # Check for duplicate title
        _existing_titles = [p["title"] for p in st.session_state["saved_plots"]]
        _save_title = _plot_title_input or f"{y_col} vs {x_col}"
        if _save_title in _existing_titles:
            # Update existing
            for _p in st.session_state["saved_plots"]:
                if _p["title"] == _save_title:
                    _p["png"] = _png
            st.toast(f"Updated: {_save_title}", icon="✅")
        else:
            st.session_state["saved_plots"].append({
                "title": _save_title,
                "png":   _png,
                "x":     x_col,
                "y":     y_col,
                "y2":    y2_col if _has_y2 else None,
            })
            st.toast(f"Saved: {_save_title}", icon="✅")

    # ── Compare diff table ──
    if df2 is not None:
        st.divider()
        st.markdown("#### 📊 Run comparison summary")
        _diff_cols = ["RPM","Thrust","Voltage","Current","Motor_Temp",
                      "ESC_Temp","Power","Torque"]
        _diff_rows = []
        for _dc in _diff_cols:
            v1 = df[_dc].mean()  if _dc in df.columns  and df[_dc].notna().any() else None
            v2 = df2[_dc].mean() if _dc in df2.columns and df2[_dc].notna().any() else None
            if v1 is None and v2 is None:
                continue
            _d  = (v2 - v1)       if (v1 is not None and v2 is not None) else None
            _dp = (_d/v1*100)     if (_d is not None and v1 != 0)        else None
            _diff_rows.append({
                "Metric":   _dc,
                label_run1: f"{v1:.2f}" if v1 is not None else "—",
                label_run2: f"{v2:.2f}" if v2 is not None else "—",
                "Delta":    f"{_d:+.2f}"   if _d  is not None else "—",
                "Delta %":  f"{_dp:+.1f}%" if _dp is not None else "—",
            })
        if _diff_rows:
            st.dataframe(pd.DataFrame(_diff_rows).set_index("Metric"),
                         use_container_width=True)

# ── Saved Plots Gallery ──
_saved_plots = st.session_state.get("saved_plots", [])
if _saved_plots:
    st.divider()
    st.subheader(f"📌 Saved Plots ({len(_saved_plots)}) — these will appear in PDF report")
    for _idx, _sp in enumerate(_saved_plots):
        _ga, _gb = st.columns([5, 1])
        _ga.markdown(f"**{_idx+1}. {_sp['title']}**  "
                     + (f"`{_sp['x']} × {_sp['y']}`" if _sp.get('x') else ""))
        if _gb.button("🗑 Remove", key=f"rm_plot_{_idx}", use_container_width=True):
            st.session_state["saved_plots"].pop(_idx)
            st.rerun()
        # Show thumbnail
        st.image(_sp["png"], width=700)
        st.divider()

# RAW TABLE
# ─────────────────────────────────────────────
if show_raw:
    st.divider()
    st.subheader("Cleaned data table")
    st.dataframe(df.reset_index(drop=True), use_container_width=True)


# ─────────────────────────────────────────────
# SAVE TO LIBRARY
# ─────────────────────────────────────────────
st.divider()
st.subheader("💾 Save to Library")
_db_row_save = fetch_run(filename) if filename else None
sav1, sav2 = st.columns([3, 1])
display_name = sav1.text_input(
    "Run name (for library)",
    value=_db_row_save["display_name"] if _db_row_save else filename.rsplit(".", 1)[0] if filename else "",
    key="display_name",
    placeholder="e.g. Motor 4 – RUN 1 – Step Test"
)

if sav2.button("💾 Save to Library", type="primary", use_container_width=True):
    # Load ip from session state or db (safe whether Initial Params expander is visible or not)
    _ip_save = st.session_state.get(f"ip_{filename}", {})
    if not _ip_save and _db_row_save and _db_row_save.get("init_params"):
        try:
            _ip_save = json.loads(_db_row_save["init_params"])
        except Exception:
            _ip_save = {}
    if not _ip_save:
        _ip_save = default_init_params(df)
    ip = _ip_save

    # Copy file into logs/ folder if uploading for the first time
    dest_path = LOGS_DIR / filename
    if mode == "upload":
        uploaded_file_obj = st.session_state.get("uploaded_file")
        if uploaded_file_obj and not dest_path.exists():
            uploaded_file_obj.seek(0)
            dest_path.write_bytes(uploaded_file_obj.read())

    stats = compute_stats(df)
    test_date = extract_test_date(filename)
    save_run(filename, display_name or filename, test_date, dest_path, stats, ip)

    # Also offer CSV download
    param_rows = [
        ("Reservoir",   "Capacity (L)",             ip["res_capacity"]),
        ("Reservoir",   "Composition",              ip["res_composition"]),
        ("Reservoir",   "Temperature (°C)",         ip["res_temperature"]),
        ("Duty Cycle",  "Duty Cycle & Flowrate",    ip["duty_cycle"]),
        ("Temperature", "Initial ESC Temp (°C)",    ip["init_esc_temp"]),
        ("Temperature", "Initial Motor Temp (°C)",  ip["init_motor_temp"]),
        ("Temperature", "Ambient (°C)",             ip["ambient_temp"]),
        ("Temperature", "ESC Inlet Coolant (°C)",   ip["esc_inlet_coolant"]),
        ("Temperature", "Motor Inlet Coolant (°C)", ip["motor_inlet_coolant"]),
        ("Flowrate",    "ESC Inlet (LPM)",          ip["esc_inlet_flow"]),
        ("Flowrate",    "Motor Inlet (LPM)",        ip["motor_inlet_flow"]),
        ("Pressure",    "ESC Inlet (Bar)",          ip["esc_inlet_pressure"]),
        ("Pressure",    "Motor Inlet (Bar)",        ip["motor_inlet_pressure"]),
        ("Battery",     "Voltage (V)",              ip["battery_voltage"]),
        ("Battery",     "SOC",                      ip["battery_soc"]),
        ("Battery",     "SOH",                      ip["battery_soh"]),
        ("Fintube",     "Inlet Temp (°C)",          ip["fin_inlet_temp"]),
        ("Fintube",     "Outlet Temp (°C)",         ip["fin_outlet_temp"]),
        ("Notes",       "Test Notes",               ip["notes"]),
    ]
    param_csv = pd.DataFrame(param_rows, columns=["Group","Parameter","Value"])\
                  .to_csv(index=False).encode("utf-8")
    base = filename.rsplit(".", 1)[0]
    st.download_button("⬇️  Download parameters CSV", param_csv,
                       f"{base}_initial_parameters.csv", "text/csv", key="dl_params")
    st.success("✅ Saved to library. Switch to **Log Library** in the sidebar to find this run anytime.")

# ─────────────────────────────────────────────
# DOWNLOAD SECTION — CSV + PDF REPORT
# ─────────────────────────────────────────────
st.divider()
st.subheader("Downloads")

base_name = filename.rsplit(".", 1)[0] if filename else "log"
dl1, dl2 = st.columns(2)

# ── CSV ──
dl1.download_button(
    "⬇️  Download cleaned CSV",
    df.to_csv(index=False).encode("utf-8"),
    f"{base_name}_cleaned.csv",
    "text/csv",
    use_container_width=True,
)

# ── PDF REPORT ──
if not TEMPLATE_PATH.exists():
    dl2.warning("⚠️  `Format.xlsx` not found next to dashboard. "
                "Place the template file in the same folder as app.py.")
else:
    if dl2.button("📄 Generate & Download PDF Report",
                  type="primary", use_container_width=True):

        with st.spinner("Building PDF report…"):

            # ── Collect data dict ──
            _db_row_dl = fetch_run(filename) if filename else None
            _ip_dl = {}
            if _db_row_dl and _db_row_dl.get("init_params"):
                try:
                    _ip_dl = json.loads(_db_row_dl["init_params"])
                except Exception:
                    pass

            # ── Smart start/end detection ──
            # Start: first row where RPM >= 50 (motor spinning, not zero)
            # End:   last row before RPM drops more than 10% from max (ramp-down begins)
            _rpm_col_ok = "RPM" in df.columns and df["RPM"].notna().any()
            if _rpm_col_ok:
                _rpm_max  = df["RPM"].max()
                _start_df = df[df["RPM"] >= 50]
                _start_row = _start_df.iloc[0] if len(_start_df) else df.iloc[0]
                # End = last row where RPM is still above 90% of peak
                _end_df   = df[df["RPM"] >= 0.90 * _rpm_max]
                _end_row  = _end_df.iloc[-1] if len(_end_df) else df.iloc[-1]
            else:
                _start_row = df.iloc[0]
                _end_row   = df.iloc[-1]

            def _col(row, col, fmt=".2f"):
                try:
                    v = float(row[col])
                    return f"{v:{fmt}}" if pd.notna(v) else ""
                except: return ""

            def _maxcol(col, fmt=".2f"):
                try:
                    return f"{float(df[col].max()):{fmt}}" if col in df.columns else ""
                except: return ""

            # Time at target RPM = duration where RPM >= 90% of peak
            _time_at_rpm = ""
            if _rpm_col_ok:
                _at_rpm = df[df["RPM"] >= 0.90 * df["RPM"].max()]
                if len(_at_rpm) > 1:
                    _time_at_rpm = f"{float(_at_rpm['Time'].max() - _at_rpm['Time'].min()):.1f}"

            pdf_data = {
                # run meta
                "filename":   filename or "",
                "test_date":  extract_test_date(filename) if filename else "",
                "test_time":  extract_test_date(filename)[-5:] if filename else "",
                "saved_at":   _db_row_dl["saved_at"] if _db_row_dl else "",
                "duration_s": f"{float(df['Time'].max()):.1f}" if "Time" in df.columns else "",
                "num_rows":   str(len(df)),
                "run_name":   _db_row_dl["display_name"] if _db_row_dl and _db_row_dl.get("display_name") else base_name,
                "operator":   "",
                # results — all from log
                "max_rpm":              _maxcol("RPM", ".0f"),
                "max_thrust":           _maxcol("Thrust", ".2f"),
                "max_torque":           f"{float(df['Torque'].abs().max()):.2f}" if "Torque" in df.columns else "",
                "max_esc_temp":         _maxcol("ESC_Temp", ".1f"),
                "max_motor_temp":       _maxcol("Motor_Temp", ".1f"),
                "max_esc_inlet_temp":   _maxcol("ESC_Inlet_Temp_C", ".1f"),
                "max_motor_inlet_temp": _maxcol("Motor_Inlet_Temp_C", ".1f"),
                "max_esc_pressure":     _maxcol("ESC_Pressure", ".3f"),
                "max_fin_inlet_temp":   _maxcol("Fin_Inlet_Temp_C", ".1f"),
                "max_fin_outlet_temp":  _maxcol("Fin_Outlet_Temp_C", ".1f"),
                "battery_voltage_post": _col(_end_row, "Voltage", ".2f"),
                "time_at_target_rpm":   _time_at_rpm,
                # efficiency placeholders — filled if user has calculated
                "mechanical_power":     "",
                "electrical_power":     f"{float(df['Voltage'].mean() * df['Current'].mean()):.0f}" if "Voltage" in df.columns and "Current" in df.columns else "",
                "mechanical_efficiency":"",
                "overall_efficiency":   "",
                # initial params from db / session (start conditions)
                "init_esc_temp":        _col(_start_row, "ESC_Temp", ".1f"),
                "init_motor_temp":      _col(_start_row, "Motor_Temp", ".1f"),
                "esc_inlet_coolant":    _col(_start_row, "ESC_Inlet_Temp_C", ".1f"),
                "motor_inlet_coolant":  _col(_start_row, "Motor_Inlet_Temp_C", ".1f"),
                "esc_inlet_flow":       _col(_start_row, "ESC_Flow", ".2f"),
                "esc_inlet_pressure":   _col(_start_row, "ESC_Pressure", ".3f"),
                "battery_voltage":      _col(_start_row, "Voltage", ".2f"),
                **{k: str(v) for k, v in _ip_dl.items()},
            }
            # session state overrides db for manually entered fields
            _ip_sess = st.session_state.get(f"ip_{filename}", {})
            pdf_data.update({k: str(v) for k, v in _ip_sess.items()})

            # ── Charts come from saved plots in session state ──
            _saved = st.session_state.get("saved_plots", [])
            chart_images = [(p["title"], p["png"]) for p in _saved]

            # ── Run name for PDF title ──
            _rname = (_db_row_dl["display_name"] if _db_row_dl and _db_row_dl.get("display_name")
                      else base_name)

            pdf_bytes = build_pdf_report(pdf_data, chart_images, _rname)

        st.download_button(
            "⬇️  Download PDF Report",
            pdf_bytes,
            f"{base_name}_report.pdf",
            "application/pdf",
            use_container_width=True,
        )
        _n = len(st.session_state.get("saved_plots", []))
        st.success(f"✅ PDF ready — {_n} saved plot(s) included.")