import re
import pandas as pd


def safe_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


def norm_colname(s: str) -> str:
    return re.sub(r"[\s_\-]+", "", str(s).strip().lower())


def find_first_col(df: pd.DataFrame, candidates: list[str]) -> str:
    mapping = {norm_colname(c): c for c in df.columns.astype(str)}
    for cand in candidates:
        key = norm_colname(cand)
        if key in mapping:
            return mapping[key]
    return ""


def find_sheet_with_required_cols(xls: pd.ExcelFile, required_candidates: dict[str, list[str]]) -> str:
    for sh in xls.sheet_names:
        try:
            tmp = pd.read_excel(xls, sheet_name=sh, nrows=60)
        except Exception:
            continue

        ok = True
        for _, cand_list in required_candidates.items():
            if not find_first_col(tmp, cand_list):
                ok = False
                break
        if ok:
            return sh
    return ""


def pct(numer: pd.Series, denom: pd.Series) -> pd.Series:
    return (numer / denom).where(denom != 0, 0)


def normalize_mawb(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip().upper()
    if not s or s in {"NAN", "NONE"}:
        return ""

    s_alnum = re.sub(r"[^0-9A-Z]", "", s)

    # digits 11 => 3+8
    if s_alnum.isdigit() and len(s_alnum) == 11:
        return f"{s_alnum[:3]}-{s_alnum[3:]}"
    # digits 12 => take last 11
    if s_alnum.isdigit() and len(s_alnum) == 12:
        s11 = s_alnum[-11:]
        return f"{s11[:3]}-{s11[3:]}"
    return s_alnum


def parse_mawb_list(text: str) -> list[str]:
    if not text or not str(text).strip():
        return []
    tokens = re.split(r"[,\s]+", str(text).strip())
    tokens = [normalize_mawb(t) for t in tokens if str(t).strip()]
    tokens = [t for t in tokens if t]
    return sorted(set(tokens))


def clean_eta_series(s: pd.Series) -> pd.Series:
    """
    Robust ETA parser (text) + normalize to DATE (no time).
    """
    s = s.astype(str).fillna("").str.strip()
    s = s.str.replace(r"(?i)^\s*eta\s*[:\-]\s*", "", regex=True)
    s = s.str.replace(r"\s+", " ", regex=True)

    # YYYYMMDD
    yyyymmdd = s.str.match(r"^\d{8}$")
    s2 = s.copy()
    if yyyymmdd.any():
        parsed = pd.to_datetime(s.loc[yyyymmdd], format="%Y%m%d", errors="coerce")
        s2.loc[yyyymmdd] = parsed.astype("datetime64[ns]").astype(str)

    dt1 = pd.to_datetime(s2, errors="coerce", infer_datetime_format=True)

    mask = dt1.isna() & s2.ne("")
    if mask.any():
        dt2 = pd.to_datetime(s2[mask], errors="coerce", dayfirst=True, infer_datetime_format=True)
        dt1.loc[mask] = dt2

    return dt1.dt.normalize()


def to_date_only(df_in: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    df_out = df_in.copy()
    for c in cols:
        if c in df_out.columns:
            df_out[c] = pd.to_datetime(df_out[c], errors="coerce").dt.date
    return df_out


def format_pct_str(x) -> str:
    try:
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return ""


def display_df(df_in: pd.DataFrame, date_cols: list[str] | None = None) -> pd.DataFrame:
    """
    For Streamlit display only: convert date columns to date-only, and % columns to formatted strings.
    Does NOT mutate original input df.
    """
    out = df_in.copy()
    if date_cols:
        out = to_date_only(out, date_cols)

    # common % columns
    for col in ["Profit Margin %", "Closed %", "ETA Filled %", "Overall Profit Margin %"]:
        if col in out.columns:
            out[col] = out[col].apply(format_pct_str)

    return out
