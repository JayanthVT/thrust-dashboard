"""
charts.py — Chart helpers for Thrust Test Rig Dashboard
Plotly helpers for interactive screen charts.
Matplotlib helpers for PDF-quality static charts.
No Streamlit imports.
"""

import io
import numpy as np
import matplotlib.pyplot as plt

# ─────────────────────────────────────────────
# MATPLOTLIB DARK STYLE
# ─────────────────────────────────────────────
DARK = {
    "figure.facecolor": "#0d0f14",
    "axes.facecolor":   "#13161e",
    "axes.edgecolor":   "#2a2d3a",
    "axes.labelcolor":  "#c8ccd8",
    "axes.grid":        True,
    "grid.color":       "#1e2130",
    "grid.linestyle":   "--",
    "grid.linewidth":   0.5,
    "xtick.color":      "#6b7280",
    "ytick.color":      "#6b7280",
    "text.color":       "#c8ccd8",
    "legend.facecolor": "#13161e",
    "legend.edgecolor": "#2a2d3a",
}

# ─────────────────────────────────────────────
# PLOTLY STYLE CONSTANTS
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
    return dict(
        title=dict(text=label, font=dict(color=_PL_TICK)),
        gridcolor=_PL_GRID, gridwidth=0.5,
        showline=True, linecolor=_PL_BORDER,
        tickfont=dict(color=_PL_TICK)
    )


def _pl_yaxis(label, color):
    return dict(
        title=dict(text=label, font=dict(color=color)),
        tickfont=dict(color=color),
        gridcolor=_PL_GRID, gridwidth=0.5,
        showline=True, linecolor=_PL_BORDER
    )


# ─────────────────────────────────────────────
# PLOTLY CHART FUNCTIONS
# All accept optional df2 for compare mode overlay.
# ─────────────────────────────────────────────

def pl_single(df, y_col, color, ylabel, unit, title,
              df2=None, label1="Run A", label2="Run B"):
    """Single Y axis Plotly chart. Optionally overlay df2 for comparison."""
    import plotly.graph_objects as _go
    fig   = _go.Figure()
    _step = max(1, len(df) // 5000)
    _df   = df.iloc[::_step]

    fig.add_trace(_go.Scatter(
        x=_df["Time"], y=_df[y_col], mode="lines",
        name=f"{y_col} ({label1})",
        line=dict(color=color, width=1.6),
        hovertemplate=f"<b>Time</b>: %{{x:.2f}}s<br>"
                      f"<b>{y_col}</b>: %{{y:.3f}} {unit}<extra>{label1}</extra>",
    ))
    if df2 is not None and y_col in df2.columns and "Time" in df2.columns:
        _step2 = max(1, len(df2) // 5000)
        _df2   = df2.iloc[::_step2]
        fig.add_trace(_go.Scatter(
            x=_df2["Time"], y=_df2[y_col], mode="lines",
            name=f"{y_col} ({label2})",
            line=dict(color="#38bdf8", width=1.4, dash="dash"),
            hovertemplate=f"<b>Time</b>: %{{x:.2f}}s<br>"
                          f"<b>{y_col}</b>: %{{y:.3f}} {unit}<extra>{label2}</extra>",
        ))
    layout = _pl_base_layout(title)
    layout["xaxis"] = _pl_xaxis()
    layout["yaxis"] = _pl_yaxis(f"{ylabel} ({unit})", color)
    fig.update_layout(**layout)
    return fig


def pl_overlay(df, y1, y2, c1, c2, l1, l2, u1, u2, title,
               df2=None, label1="Run A", label2="Run B"):
    """Dual Y axis Plotly chart. Optionally overlay df2 for comparison."""
    import plotly.graph_objects as _go
    fig   = _go.Figure()
    _step = max(1, len(df) // 5000)
    _df   = df.iloc[::_step]

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
        _df2   = df2.iloc[::_step2]
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
    layout           = _pl_base_layout(title)
    layout["xaxis"]  = _pl_xaxis()
    layout["yaxis"]  = _pl_yaxis(f"{l1} ({u1})", c1)
    layout["yaxis2"] = dict(
        title=dict(text=f"{l2} ({u2})", font=dict(color=c2)),
        tickfont=dict(color=c2), overlaying="y", side="right",
        showgrid=False, showline=True, linecolor=_PL_BORDER)
    fig.update_layout(**layout)
    return fig


def pl_multi(df, cols_colors, title, unit="",
             df2=None, label1="Run A", label2="Run B"):
    """Multiple traces on one Y axis. Optionally overlay df2."""
    import plotly.graph_objects as _go
    fig   = _go.Figure()
    _step = max(1, len(df) // 5000)
    _df   = df.iloc[::_step]

    for col, color in cols_colors:
        if col in _df.columns:
            fig.add_trace(_go.Scatter(
                x=_df["Time"], y=_df[col], mode="lines",
                name=f"{col} ({label1})", line=dict(color=color, width=1.4),
                hovertemplate=f"<b>{col}</b>: %{{y:.2f}} {unit}<extra>{label1}</extra>",
            ))
    if df2 is not None:
        _step2 = max(1, len(df2) // 5000)
        _df2   = df2.iloc[::_step2]
        for col, color in cols_colors:
            if col in _df2.columns and "Time" in _df2.columns:
                fig.add_trace(_go.Scatter(
                    x=_df2["Time"], y=_df2[col], mode="lines",
                    name=f"{col} ({label2})",
                    line=dict(color=color, width=1.0, dash="dash"),
                    hovertemplate=f"<b>{col}</b>: %{{y:.2f}} {unit}<extra>{label2}</extra>",
                ))
    layout           = _pl_base_layout(title)
    layout["xaxis"]  = _pl_xaxis()
    layout["yaxis"]  = _pl_yaxis(unit, _PL_TEXT)
    fig.update_layout(**layout)
    return fig


# ─────────────────────────────────────────────
# MATPLOTLIB CHART FUNCTIONS (PDF only)
# ─────────────────────────────────────────────

def make_single_plot(df, y_col, color, ylabel, unit, title):
    """Matplotlib single-axis chart for PDF embedding."""
    with plt.style.context(DARK):
        fig, ax = plt.subplots(figsize=(9, 3.2))
        ax.plot(df["Time"], df[y_col], color=color, linewidth=1.4, alpha=0.9)
        ax.fill_between(df["Time"], df[y_col], alpha=0.10, color=color)
        idx  = df[y_col].idxmax()
        px   = df.loc[idx, "Time"]
        py   = df.loc[idx, y_col]
        ax.annotate(f"Peak {py:.2f} {unit}", xy=(px, py), xytext=(px, py * 0.78),
                    arrowprops=dict(arrowstyle="->", color=color, lw=1.0),
                    color=color, fontsize=8, fontfamily="monospace")
        ax.set_title(title, fontsize=10, fontweight="bold", pad=8)
        ax.set_xlabel("Time (s)", fontsize=8)
        ax.set_ylabel(f"{ylabel} ({unit})", fontsize=8)
        fig.tight_layout()
        return fig


def make_overlay_plot(df, y1, y2, c1, c2, l1, l2, u1, u2, title):
    """Matplotlib dual-axis chart for PDF embedding."""
    with plt.style.context(DARK):
        fig, ax1 = plt.subplots(figsize=(12, 3.6))
        ax2 = ax1.twinx()
        ax1.plot(df["Time"], df[y1], color=c1, linewidth=1.4, label=l1)
        ax1.fill_between(df["Time"], df[y1], alpha=0.08, color=c1)
        ax2.plot(df["Time"], df[y2], color=c2, linewidth=1.2,
                 linestyle="--", label=l2, alpha=0.85)
        ax1.set_xlabel("Time (s)", fontsize=8)
        ax1.set_ylabel(f"{l1} ({u1})", color=c1, fontsize=8)
        ax2.set_ylabel(f"{l2} ({u2})", color=c2, fontsize=8)
        ax1.tick_params(axis="y", colors=c1)
        ax2.tick_params(axis="y", colors=c2)
        lines = ax1.get_lines() + ax2.get_lines()
        ax1.legend(lines, [l.get_label() for l in lines],
                   loc="upper left", fontsize=8)
        ax1.set_title(title, fontsize=10, fontweight="bold", pad=8)
        fig.tight_layout()
        return fig


def fig_to_png(fig, dpi=150) -> bytes:
    """Convert a matplotlib figure to PNG bytes and close it."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf.read()
