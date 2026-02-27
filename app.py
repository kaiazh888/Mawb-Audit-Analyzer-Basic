import io
import re
import pandas as pd
import streamlit as st


# ---------------- Page Config ----------------
st.set_page_config(page_title="MAWB Audit Analyzer", layout="wide")

st.title("MAWB Audit Analyzer (Billing-only)")
st.caption(
    "Upload Billing charges export + optional MAWB→ETA mapping file. "
    "Supports MAWB filter box, profit margin analysis, zero buckets, "
    "outliers, negative profit, and Charge Code / Vendor summaries."
)

# ---------------- Helpers ----------------

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


def pct(numer: pd.Series, denom: pd.Series) -> pd.Series:
    return (numer / denom).where(denom != 0, 0)


def normalize_mawb(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip().upper()
    if not s or s in {"NAN", "NONE"}:
        return ""

    s_alnum = re.sub(r"[^0-9A-Z]", "", s)

    if s_alnum.isdigit() and len(s_alnum) == 11:
        return f"{s_alnum[:3]}-{s_alnum[3:]}"
    if s_alnum.isdigit() and len(s_alnum) == 12:
        s11 = s_alnum[-11:]
        return f"{s11[:3]}-{s11[3:]}"

    if "-" in s and len(s.split("-")[0]) == 3:
        return s

    return s_alnum or s


def parse_mawb_list(text: str) -> list[str]:
    if not text or not str(text).strip():
        return []
    tokens = re.split(r"[,\s]+", str(text).strip())
    tokens = [normalize_mawb(t) for t in tokens if str(t).strip()]
    return sorted(set([t for t in tokens if t]))


# ---------------- Uploaders ----------------

billing_file = st.file_uploader(
    "Upload Billing Charges Excel (.xlsx)",
    type=["xlsx"],
)

eta_file = st.file_uploader(
    "Optional: Upload MAWB→ETA mapping Excel (.xlsx)",
    type=["xlsx"],
)

st.divider()

st.subheader("Optional Filter: Keep only specified MAWBs")

mawb_text = st.text_area(
    "Paste MAWBs here (comma / space / newline separated).",
    height=140,
)

# ---------------- Main ----------------

if not billing_file:
    st.info("Please upload a Billing Charges Excel file to start.")
    st.stop()

try:

    df = pd.read_excel(billing_file)

    df["MAWB"] = df["MAWB"].apply(normalize_mawb)
    df["Cost Amount"] = safe_numeric(df["Cost Amount"])
    df["Sell Amount"] = safe_numeric(df["Sell Amount"])

    # Optional MAWB filter
    mawb_keep = parse_mawb_list(mawb_text)
    if mawb_keep:
        df = df[df["MAWB"].isin(mawb_keep)]

    # MAWB Summary
    summary = (
        df.groupby("MAWB", as_index=False)
        .agg(
            Total_Cost=("Cost Amount", "sum"),
            Total_Sell=("Sell Amount", "sum"),
        )
    )

    summary["Profit"] = summary["Total_Sell"] - summary["Total_Cost"]
    summary["Profit Margin %"] = pct(
        summary["Profit"],
        summary["Total_Sell"],
    )

    st.subheader("MAWB Summary")
    st.dataframe(summary, use_container_width=True)

    # Export
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary.to_excel(writer, index=False)

    st.download_button(
        "Download Report Excel",
        data=output.getvalue(),
        file_name="MAWB_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.exception(e)
