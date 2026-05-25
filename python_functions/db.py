"""
db.py — SQLite database functions for Thrust Test Rig Dashboard
All run storage, retrieval, folder management lives here.
No Streamlit imports. Pure Python + SQLite.
"""

import sqlite3
import json
from datetime import datetime
from pathlib import Path

# DB_PATH is set by app.py and passed in — default here for standalone use
_DEFAULT_DB = Path(__file__).parent.parent / "thrust_logs.db"


def get_conn(db_path=None):
    path = db_path or _DEFAULT_DB
    conn = sqlite3.connect(str(path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path=None):
    with get_conn(db_path) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS runs (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                filename       TEXT UNIQUE,
                display_name   TEXT,
                folder         TEXT DEFAULT 'Uncategorised',
                test_date      TEXT,
                saved_at       TEXT,
                file_path      TEXT,
                max_thrust_n   REAL,
                max_rpm        REAL,
                max_power_w    REAL,
                max_current_a  REAL,
                max_voltage_v  REAL,
                max_esc_temp   REAL,
                max_motor_temp REAL,
                duration_s     REAL,
                num_rows       INTEGER,
                init_params    TEXT
            )
        """)
        # Migrate existing databases — add folder column if missing
        try:
            conn.execute("ALTER TABLE runs ADD COLUMN folder TEXT DEFAULT 'Uncategorised'")
        except Exception:
            pass  # column already exists
        conn.commit()


def save_run(filename, display_name, test_date, file_path, stats,
             init_params, folder="Uncategorised", db_path=None):
    with get_conn(db_path) as conn:
        conn.execute("""
            INSERT INTO runs
              (filename, display_name, folder, test_date, saved_at, file_path,
               max_thrust_n, max_rpm, max_power_w, max_current_a,
               max_voltage_v, max_esc_temp, max_motor_temp, duration_s,
               num_rows, init_params)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(filename) DO UPDATE SET
              display_name   = excluded.display_name,
              folder         = excluded.folder,
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
            filename, display_name, folder, test_date,
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


def update_init_params(filename, init_params, db_path=None):
    with get_conn(db_path) as conn:
        conn.execute(
            "UPDATE runs SET init_params=?, saved_at=? WHERE filename=?",
            (json.dumps(init_params),
             datetime.now().isoformat(timespec="seconds"),
             filename)
        )
        conn.commit()


def fetch_all_runs(search_text="", date_from=None, date_to=None,
                   folder=None, db_path=None):
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
    if folder:
        q += " AND folder = ?"
        params.append(folder)
    q += " ORDER BY folder ASC, test_date DESC, saved_at DESC"
    with get_conn(db_path) as conn:
        return [dict(r) for r in conn.execute(q, params).fetchall()]


def fetch_run(filename, db_path=None):
    with get_conn(db_path) as conn:
        row = conn.execute(
            "SELECT * FROM runs WHERE filename=?", (filename,)
        ).fetchone()
        return dict(row) if row else None


def delete_run(filename, db_path=None):
    row = fetch_run(filename, db_path)
    if row and row["file_path"] and Path(row["file_path"]).exists():
        Path(row["file_path"]).unlink(missing_ok=True)
    with get_conn(db_path) as conn:
        conn.execute("DELETE FROM runs WHERE filename=?", (filename,))
        conn.commit()


def fetch_folders(db_path=None):
    """Return sorted list of distinct folder names."""
    with get_conn(db_path) as conn:
        rows = conn.execute(
            "SELECT DISTINCT folder FROM runs ORDER BY folder ASC"
        ).fetchall()
        return [r["folder"] for r in rows] if rows else []


def move_run_to_folder(filename, folder, db_path=None):
    """Move a run to a different folder."""
    with get_conn(db_path) as conn:
        conn.execute(
            "UPDATE runs SET folder=? WHERE filename=?",
            (folder, filename)
        )
        conn.commit()
