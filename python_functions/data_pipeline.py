"""
data_pipeline.py — Data ingestion and processing for Thrust Test Rig Dashboard
Handles file loading, column normalisation, time parsing, cleaning, and stats.
No Streamlit imports. Pure Python + Pandas + NumPy.
"""

import io
import re
import numpy as np
import pandas as pd
from datetime import datetime, date
from pathlib import Path

# ─────────────────────────────────────────────
# COLUMN ALIAS MAP
# Add new aliases here when new log formats arrive.
# ─────────────────────────────────────────────
COLUMN_ALIASES = {
    "Time":         ["Timestamp", "Time", "T", "elapsed"],
    "Thrust":       ["Net_Thrust_N", "Thrust", "Force_N", "ThrustN", "Net_Thrust"],
    "RPM":          ["Actual_RPM", "RPM", "Motor_RPM"],
    "Cmd_RPM":      ["Commanded_RPM", "Cmd_RPM", "Target_RPM"],
    "Motor_Temp":   ["Motor_Temp_C", "Motor_Temp"],
    "ESC_Temp":     ["ESC_Temp_C", "ESC_Temp"],
    "Power":        ["Power_W", "Power"],
    "Current":      ["Current_A", "Current"],
    "Voltage":      ["DC_Voltage_V", "DC_Voltage", "Voltage"],
    "Torque":       ["Torque(0.23m)", "Net_Torque"],
    "Total_Weight": ["Total_Weight_kg", "Total_Weight"],
    "Accel_X":      ["Accel_X_g", "Accel_X"],
    "Accel_Y":      ["Accel_Y_g", "Accel_Y"],
    "Accel_Z":      ["Accel_Z_g", "Accel_Z"],
    "ESC_Pressure": ["ESC_Inlet_Pressure_Bar", "ESC_Inlet_Pressure"],
    "ESC_Flow":     ["ESC_Inlet_Flow_Lpm", "ESC_Inlet_Flow"],
    "ESC_Inlet_Temp_C":   ["ESC_Inlet_Temp_C", "esc_inlet_temp", "ESC_Inlet_Temp"],
    "Motor_Inlet_Temp_C": ["Motor_Inlet_Temp_C", "motor_inlet_temp", "Motor_Inlet_Temp"],
    "Fin_Inlet_Temp_C":   ["Fin_Inlet_Temp_C", "fin_inlet_temp", "Fin_Inlet_Temp"],
    "Fin_Outlet_Temp_C":  ["Fin_Outlet_Temp_C", "fin_outlet_temp", "Fin_Outlet_Temp"],
    "Motor_Flow":   ["Motor_Flow_Lpm", "Motor_Flow"],
}


def load_file_from_path(path: Path, logs: list):
    """Load a log file from a local path."""
    name = path.name.lower()
    raw  = path.read_bytes()
    return _parse_raw(raw, name, logs)


def load_file_from_upload(uploaded_file, logs: list):
    """Load a log file from a Streamlit UploadedFile object."""
    raw  = uploaded_file.read()
    name = uploaded_file.name.lower()
    return _parse_raw(raw, name, logs)


def _parse_raw(raw: bytes, name: str, logs: list):
    """Try all known parsers and return a DataFrame or None."""
    if name.endswith((".xlsx", ".xlsm")):
        for engine in ["openpyxl", "calamine"]:
            try:
                pd.ExcelFile(io.BytesIO(raw), engine=engine)
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

    # CSV fallback — try multiple encodings and engines
    for enc in ["utf-8", "latin1", "cp1252"]:
        for eng in ["c", "python"]:
            try:
                df = pd.read_csv(io.BytesIO(raw), encoding=enc,
                                 engine=eng, on_bad_lines="skip")
                logs.append(f"✅ CSV parsed enc={enc} eng={eng}")
                return df
            except Exception:
                pass
    logs.append("❌ Could not parse file with any known format.")
    return None


def normalize_columns(df: pd.DataFrame, logs: list) -> pd.DataFrame:
    """
    Rename columns to canonical names using COLUMN_ALIASES map.
    Matching is case-insensitive so any capitalisation of the source
    column name will be found automatically.
    """
    # Build a lowercase lookup: lowercase_alias -> canonical
    _lower_map = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            _lower_map[alias.lower()] = canonical

    # Build a lowercase -> original column name map from the dataframe
    _df_lower = {col.lower(): col for col in df.columns}

    rename_map = {}
    for lower_alias, canonical in _lower_map.items():
        if lower_alias in _df_lower and canonical not in rename_map.values():
            original_col = _df_lower[lower_alias]
            if original_col not in rename_map and canonical not in df.columns:
                rename_map[original_col] = canonical
                logs.append(f"🔀 {original_col} → {canonical}")

    return df.rename(columns=rename_map)


def parse_time(df: pd.DataFrame, logs: list) -> pd.DataFrame:
    """Convert Time column to numeric elapsed seconds. Four fallback layers."""
    if "Time" not in df.columns:
        logs.append("❌ No Time column.")
        return df

    col = df["Time"]

    # Layer 1: already datetime64
    if pd.api.types.is_datetime64_any_dtype(col):
        df["Time"] = (col - col.iloc[0]).dt.total_seconds()
        logs.append(f"✅ Datetime → elapsed s ({float(df['Time'].max()):.1f}s)")
        return df

    # Layer 2: already numeric
    if pd.api.types.is_numeric_dtype(col):
        logs.append("✅ Time already numeric.")
        return df

    # Layer 3: datetime string parse
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        parsed = pd.to_datetime(col, errors="coerce")
    if parsed.notna().sum() / max(len(df), 1) > 0.5:
        df["Time"] = (parsed - parsed.iloc[0]).dt.total_seconds()
        logs.append(f"✅ String → elapsed s ({float(df['Time'].max()):.1f}s)")
        return df

    # Layer 4: numeric coerce
    num = pd.to_numeric(col, errors="coerce")
    if num.notna().sum() / max(len(df), 1) > 0.5:
        df["Time"] = num
        logs.append("✅ Time coerced to numeric.")
        return df

    # Layer 5: MM:SS.f format (e.g. "34:15.6" → 2055.6 seconds)
    def _parse_mmss(val):
        try:
            parts = str(val).strip().split(":")
            if len(parts) == 2:
                return int(parts[0]) * 60 + float(parts[1])
            if len(parts) == 3:   # HH:MM:SS.f
                return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
        except Exception:
            pass
        return None

    mmss = col.apply(_parse_mmss)
    if mmss.notna().sum() / max(len(df), 1) > 0.5:
        # Convert to elapsed seconds from first value
        first = mmss.dropna().iloc[0]
        df["Time"] = (mmss - first).round(3)
        logs.append(f"✅ MM:SS format → elapsed s ({float(df['Time'].max()):.1f}s)")
        return df

    # Last resort: row index as proxy
    logs.append("⚠️  Could not parse Time — using row index as time proxy.")
    df["Time"] = np.arange(len(df), dtype=float)
    return df


def clean_and_drop(df: pd.DataFrame, logs: list) -> pd.DataFrame:
    """Coerce all non-Time columns to numeric and drop rows missing key columns."""
    for col in [c for c in df.columns if c != "Time"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    present = [c for c in ["Time", "Thrust", "RPM"] if c in df.columns]
    before  = len(df)
    df      = df.dropna(subset=present)
    if before - len(df):
        logs.append(f"🗑️  Dropped {before - len(df)} rows with NaN in {present}")
    return df


def extract_test_date(filename: str) -> str:
    """Pull YYYYMMDD from filename like Motor4_RUN1_20260416_114449, else today."""
    m = re.search(r"(\d{8})", filename)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            pass
    return date.today().isoformat()


def compute_stats(df: pd.DataFrame) -> dict:
    """Extract peak statistics from a cleaned dataframe."""
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


def default_init_params(df: pd.DataFrame) -> dict:
    """
    Build default initial parameters dict from the first row of the dataframe.
    Values are auto-filled where column exists, empty string otherwise.
    """
    t0 = df.iloc[0]

    def g(col, fmt="{:.2f}"):
        try:
            v = t0[col]
            return fmt.format(float(v)) if pd.notna(v) else ""
        except Exception:
            return ""

    return {
        "res_capacity":        "10",
        "res_composition":     "1:1 Glycol:Distilled water",
        "res_temperature":     g("Motor_Inlet_Temp_C", "{:.1f}"),
        "duty_cycle":          "70% ~6lpm",
        "init_esc_temp":       g("ESC_Temp",            "{:.1f}"),
        "init_motor_temp":     g("Motor_Temp",           "{:.1f}"),
        "ambient_temp":        "",
        "esc_inlet_coolant":   g("ESC_Inlet_Temp_C",    "{:.1f}"),
        "motor_inlet_coolant": g("Motor_Inlet_Temp_C",  "{:.1f}"),
        "esc_inlet_flow":      g("ESC_Flow",             "{:.2f}"),
        "motor_inlet_flow":    "-",
        "esc_inlet_pressure":  g("ESC_Pressure",         "{:.3f}"),
        "motor_inlet_pressure":"-",
        "battery_voltage":     g("Voltage",              "{:.2f}"),
        "battery_soc":         "-%",
        "battery_soh":         "100%",
        "fin_inlet_temp":      g("Fin_Inlet_Temp_C",    "{:.1f}"),
        "fin_outlet_temp":     g("Fin_Outlet_Temp_C",   "{:.1f}"),
        "notes":               "",
    }