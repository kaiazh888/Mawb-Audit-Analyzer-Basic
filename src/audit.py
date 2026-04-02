from __future__ import annotations

import re
import pandas as pd


def safe_numeric(s: pd.Series) -> pd.Series:
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
    if df is None:
        return None

    norm_map = {str(c).strip().lower(): c for c in df.columns}
    for c in candidates:
        key = str(c).strip().lower()
        if key in norm_map:
            return norm_map[key]
    return None


def find_sheet_with_required_cols(xls: pd.ExcelFile, required: dict[str, list[str]]) -> str | None:
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
    if not mawb_text or not str(mawb_text).strip():
        return []

    raw = re.split(r"[\s,;]+", str(mawb_text).strip())
    vals = [normalize_mawb(x) for x in raw if str(x).strip()]
    vals = [x for x in vals if x]
    return list(dict.fromkeys(vals))


def clean_eta_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="datetime64[ns]")

    s2 = s.astype(str).str.strip()
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
    return pd.to_datetime(s2, errors="coerce")


def pct(num, den):
    num_s = pd.Series(num) if not isinstance(num, pd.Series) else num
    den_s = pd.Series(den) if not isinstance(den, pd.Series) else den

    out = num_s / den_s.replace(0, pd.NA)
    return out.fillna(0.0)


def format_pct_str(x) -> str:
    try:
        if pd.isna(x):
            return ""
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return ""


def display_df(df: pd.DataFrame, date_cols: list[str] | None = None) -> pd.DataFrame:
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
    if df is None:
        return pd.DataFrame()

    out = df.copy()
    date_cols = date_cols or []

    for c in date_cols:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce").dt.normalize()

    return out
