"""
ValuAI - Display Helper Functions
==================================
Small utility functions for consistent formatting of numbers, percentages,
currency, and dates in the Streamlit UI.
"""

from datetime import datetime, date


def fmt_pct(value, decimals=2):
    """Format a number as a percentage string. Input: 7.21 -> Output: '7.21%'.
    Returns 'N/A' for None or non-numeric inputs."""
    if value is None:
        return "N/A"
    try:
        return f"{float(value):.{decimals}f}%"
    except (TypeError, ValueError):
        return "N/A"


def fmt_money(value, decimals=2):
    """Format a number with thousand separators. Input: 1375.58 -> '1,375.58'.
    Returns '—' for None or non-numeric (cleaner table display than 'N/A')."""
    if value is None:
        return "—"
    try:
        v = float(value)
        return f"{v:,.{decimals}f}"
    except (TypeError, ValueError):
        return "—"


def fmt_number(value, decimals=4):
    """Format plain number with given decimals. Used for ratios like Beta, multiples."""
    if value is None:
        return "—"
    try:
        return f"{float(value):.{decimals}f}"
    except (TypeError, ValueError):
        return "—"


def fmt_date(value):
    """Display a date in 'DD-MMM-YYYY' format. Input can be ISO string, date,
    datetime, or None."""
    if value is None or value == "":
        return "—"
    if isinstance(value, datetime):
        return value.strftime("%d-%b-%Y")
    if isinstance(value, date):
        return value.strftime("%d-%b-%Y")
    try:
        if isinstance(value, str):
            d = datetime.fromisoformat(value).date()
            return d.strftime("%d-%b-%Y")
    except (TypeError, ValueError):
        pass
    return str(value)


def list_to_columns_dict(year_labels, values):
    """Convert two parallel lists into a dict suitable for st.dataframe display."""
    out = {}
    for label, v in zip(year_labels, values):
        out[label] = fmt_money(v) if isinstance(v, (int, float)) else (v if v is not None else "—")
    return out