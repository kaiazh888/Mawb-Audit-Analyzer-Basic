from __future__ import annotations

import re
import pandas as pd


def safe_numeric(s: pd.Series) -> pd.Series:
    """
    Convert a Series to numeric safely.
    Handles commas, blanks, and invalid values.
    """
    if s is None:
        return pd.Series(dtype="float64")

    s2 = (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(s2, errors="coerce").fillna(0.0)


def find_first_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """
    Find the first matching column name from a candidate list.
    Match is case-insensitive and ignores surrounding spaces.
    """
    if df is None or df.empty:
        return None

    norm_map = {str(c).strip().lower(): c for c in df.columns}
    for c in candidates:
        key = str(c).strip().lower()
        if key in norm_map:
            return norm_map[key]
    return None


def find_sheet_with_required_cols(xls: pd.ExcelFile, required: dict[str, list[str]]) -> str | None:
    """
    Find the first sheet that contains all required logical columns.
    `required` is a dict like:
        {
            "MAWB": ["MAWB", "Master AWB"],
            "Cost Amount": ["Cost Amount", "Cost"],
        }
    """
    for sheet in xls.sheet_names:
        try:
            preview = pd.read_excel(xls, sheet_name=sheet, nrows=5)
        except Exception:
            continue

        ok = True
        for _, candidates in required.items():
            if find_first_col(preview, candidates) is None:
                ok = False
                break
        if ok:
            return sheet

    return None


def normalize_mawb(x) -> str:
    """
    Normalize MAWB:
    - keep digits and hyphen
    - trim spaces
    - if 11 digits like 12345678901 => 123-45678901
    - otherwise keep cleaned text
    """
    if pd.isna(x):
        return ""

    s = str(x).strip()
    if not s:
        return ""

    s = s.replace(" ", "")
    s = re.sub(r"[^0-9\-]", "", s)

    digits = re.sub(r"[^0-9]", "", s)
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:]}"
    return s


def parse_mawb_list(mawb_text: str) -> list[str]:
    """
    Parse user-entered MAWB list from text area.
    Supports comma/newline/space separated values.
    """
    if not mawb_text or not str(mawb_text).strip():
        return []

    raw = re.split(r"[\s,;]+", str(mawb_text).strip())
    vals = [normalize_mawb(x) for x in raw if str(x).strip()]
    vals = [x for x in vals if x]
    return list(dict.fromkeys(vals))


def clean_eta_series(s: pd.Series) -> pd.Series:
    """
    Clean ETA column into pandas datetime.
    Compatible with newer pandas versions.
    """
    if s is None:
        return pd.Series(dtype="datetime64[ns]")

    s2 = s.astype(str).str.strip()

    # Common placeholders -> NaT
    s2 = s2.replace(
        {
            "": None,
            "nan": None,
            "NaN": None,
            "None": None,
            "NULL": None,
            "N/A": None,
            "NA": None,
        }
    )

    # New pandas: no infer_datetime_format argument
    dt1 = pd.to_datetime(s2, errors="coerce")
    return dt1


def pct(num, den):
    """
    Safe percentage calculation.
    Returns 0 when denominator is 0 or missing.
    """
    num_s = pd.Series(num) if not isinstance(num, pd.Series) else num
    den_s = pd.Series(den) if not isinstance(den, pd.Series) else den

    out = num_s / den_s.replace(0, pd.NA)
    out = out.fillna(0.0)
    return out


def format_pct_str(x) -> str:
    """
    Format a ratio like 0.1234 -> '12.34%'.
    """
    try:
        if pd.isna(x):
            return ""
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return ""


def display_df(df: pd.DataFrame, date_cols: list[str] | None = None) -> pd.DataFrame:
    """
    Return a display-friendly DataFrame:
    - dates formatted as YYYY-MM-DD
    - percentage columns formatted as xx.xx%
    - numeric columns kept as-is otherwise
    """
    if df is None:
        return pd.DataFrame()

    out = df.copy()

    date_cols = date_cols or []
    for c in date_cols:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce").dt.strftime("%Y-%m-%d")
            out[c] = out[c].fillna("")

    for c in out.columns:
        if "margin" in str(c).lower() and pd.api.types.is_numeric_dtype(out[c]):
            out[c] = out[c].apply(lambda x: format_pct_str(x) if pd.notna(x) else "")

    return out


def to_date_only(df: pd.DataFrame, date_cols: list[str] | None = None) -> pd.DataFrame:
    """
    Convert datetime columns to date-only values for Excel export.
    Keeps them as datetime/date-compatible values instead of strings.
    """
    if df is None:
        return pd.DataFrame()

    out = df.copy()
    date_cols = date_cols or []

    for c in date_cols:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce").dt.normalize()

    return out
