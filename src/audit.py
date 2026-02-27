from __future__ import annotations

import pandas as pd
import streamlit as st
from dataclasses import dataclass

from .helpers import (
    safe_numeric,
    find_first_col,
    find_sheet_with_required_cols,
    normalize_mawb,
    parse_mawb_list,
    clean_eta_series,
    pct,
    display_df,
)


BILLING_REQUIRED = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "Cost Amount": ["Cost Amount", "Cost", "AP Amount", "Total Cost", "CostAmount"],
    "Sell Amount": ["Sell Amount", "Sell", "AR Amount", "Total Sell", "SellAmount"],
}
BILLING_OPTIONAL = {
    "Client": ["Client", "Customer", "Account", "Shipper", "Bill To", "Billed To"],
    "Charge Code": ["Charge Code", "ChargeCode", "Charge", "Code"],
    "Vendor": ["Vendor", "Carrier", "Supplier"],
}
ETA_REQUIRED = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "ETA": ["ETA", "Eta", "Estimated Time of Arrival", "Arrival", "Arrival Date", "ETA Date"],
}


@dataclass
class AuditResult:
    # Inputs
    mawb_keep: list[str]
    mawb_not_found_df: pd.DataFrame
    eta_parse_note: str | None

    # KPI
    kpi_vertical: pd.DataFrame
    neg_summary: pd.DataFrame

    # Core dataframes (raw for export)
    df: pd.DataFrame
    summary: pd.DataFrame
    exceptions: pd.DataFrame
    client_summary: pd.DataFrame
    margin_outliers: pd.DataFrame
    negative_profit: pd.DataFrame
    zero_margin: pd.DataFrame
    zero_profit: pd.DataFrame
    both_zero: pd.DataFrame
    sell_zero_only: pd.DataFrame
    cost_zero_only: pd.DataFrame
    chargecode_summary: pd.DataFrame
    vendor_summary: pd.DataFrame
    chargecode_profit_le0_mawb: pd.DataFrame
    mawb_not_found: list[str]
    margin_label: str

    # Display dfs (formatted strings)
    display_summary: pd.DataFrame
    display_exceptions: pd.DataFrame
    display_client_summary: pd.DataFrame
    display_margin_outliers: pd.DataFrame
    display_negative_profit: pd.DataFrame
    display_zero_margin: pd.DataFrame
    display_zero_profit: pd.DataFrame
    display_both_zero: pd.DataFrame
    display_sell_zero_only: pd.DataFrame
    display_cost_zero_only: pd.DataFrame
    display_chargecode_summary: pd.DataFrame
    display_vendor_summary: pd.DataFrame
    display_chargecode_profit_le0_mawb: pd.DataFrame


def _make_kpi_vertical(kpi_dict: dict, pct_keys: set[str]) -> pd.DataFrame:
    from .helpers import format_pct_str
    rows = []
    for k, v in kpi_dict.items():
        if k in pct_keys:
            rows.append({"Metric": k, "Value": format_pct_str(v)})
        else:
            rows.append({"Metric": k, "Value": v})
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def run_audit(
    billing_file,
    eta_file,
    mawb_text: str,
    low_thr: float = 0.30,
    high_thr: float = 0.80,
) -> AuditResult:
    MARGIN_LABEL = f"Margin<{int(low_thr*100)}% or >{int(high_thr*100)}%"

    # ---- Read billing charges ----
    xls = pd.ExcelFile(billing_file)
    billing_sheet = find_sheet_with_required_cols(xls, BILLING_REQUIRED)
    if not billing_sheet:
        raise ValueError(
            "Could not find a sheet containing required fields: MAWB, Cost Amount, Sell Amount."
        )

    raw_df = pd.read_excel(xls, sheet_name=billing_sheet)

    mawb_col = find_first_col(raw_df, BILLING_REQUIRED["MAWB"])
    cost_col = find_first_col(raw_df, BILLING_REQUIRED["Cost Amount"])
    sell_col = find_first_col(raw_df, BILLING_REQUIRED["Sell Amount"])
    client_col = find_first_col(raw_df, BILLING_OPTIONAL["Client"])
    charge_code_col = find_first_col(raw_df, BILLING_OPTIONAL["Charge Code"])
    vendor_col = find_first_col(raw_df, BILLING_OPTIONAL["Vendor"])

    if not (mawb_col and cost_col and sell_col):
        raise ValueError("Billing sheet found but required columns could not be detected after scanning.")

    # Normalize billing
    df = raw_df.copy()
    df["MAWB"] = df[mawb_col].apply(normalize_mawb)
    df["Cost Amount"] = safe_numeric(df[cost_col])
    df["Sell Amount"] = safe_numeric(df[sell_col])

    if client_col:
        df["Client"] = df[client_col].astype(str).str.strip()
        df.loc[df["Client"].isin(["", "nan", "None"]), "Client"] = "UNKNOWN"
    else:
        df["Client"] = "UNKNOWN"

    if charge_code_col:
        df["Charge Code"] = df[charge_code_col].astype(str).str.strip()
        df.loc[df["Charge Code"].isin(["", "nan", "None"]), "Charge Code"] = "UNKNOWN"
    else:
        df["Charge Code"] = "UNKNOWN"

    if vendor_col:
        df["Vendor"] = df[vendor_col].astype(str).str.strip()
        df.loc[df["Vendor"].isin(["", "nan", "None"]), "Vendor"] = "UNKNOWN"
    else:
        df["Vendor"] = "UNKNOWN"

    df = df[df["MAWB"].ne("")].copy()

    # ---- Optional MAWB filter ----
    mawb_keep = parse_mawb_list(mawb_text)
    if mawb_keep:
        before_mawb = df["MAWB"].nunique()
        df = df[df["MAWB"].isin(mawb_keep)].copy()
        after_mawb = df["MAWB"].nunique()
        found_set = set(df["MAWB"].unique())
        mawb_not_found = sorted(set(mawb_keep) - found_set)
        mawb_not_found_df = pd.DataFrame({"MAWB": mawb_not_found})

        # Keep a small info note in dataframe? UI shows separately anyway.
        _ = (before_mawb, after_mawb)
    else:
        mawb_not_found = []
        mawb_not_found_df = pd.DataFrame({"MAWB": []})

    # ---- Read ETA mapping (optional) ----
    eta_map = None
    eta_parse_note = None

    if eta_file:
        xls2 = pd.ExcelFile(eta_file)
        map_sheet = find_sheet_with_required_cols(xls2, ETA_REQUIRED)
        if map_sheet:
            mdf0 = pd.read_excel(xls2, sheet_name=map_sheet)
            m_mawb = find_first_col(mdf0, ETA_REQUIRED["MAWB"])
            m_eta = find_first_col(mdf0, ETA_REQUIRED["ETA"])
            if m_mawb and m_eta:
                mdf = mdf0[[m_mawb, m_eta]].copy()
                mdf.columns = ["MAWB", "ETA"]
                mdf["MAWB"] = mdf["MAWB"].apply(normalize_mawb)
                mdf["ETA"] = clean_eta_series(mdf["ETA"])

                bad_eta_rows = int(mdf["ETA"].isna().sum())
                total_rows = int(len(mdf))
                if total_rows > 0 and bad_eta_rows > 0:
                    eta_parse_note = (
                        f"ETA parsing note: {bad_eta_rows} / {total_rows} ETA values could not be parsed and were left blank."
                    )

                eta_map = (
                    mdf.dropna(subset=["MAWB"])
                    .groupby("MAWB", as_index=False)["ETA"]
                    .max()
                )

    # ---- Merge ETA into billing ----
    if eta_map is not None and not eta_map.empty:
        df = df.merge(eta_map, on="MAWB", how="left")
    else:
        df["ETA"] = pd.NaT

    df["ETA"] = pd.to_datetime(df["ETA"], errors="coerce").dt.normalize()

    # ---- MAWB summary ----
    summary = (
        df.groupby("MAWB", as_index=False)
        .agg(
            Client=("Client", "first"),
            Total_Cost=("Cost Amount", "sum"),
            Total_Sell=("Sell Amount", "sum"),
            Line_Count=("MAWB", "size"),
            ETA=("ETA", "max"),
        )
    )

    summary["ETA Month"] = summary["ETA"].dt.to_period("M").astype(str).replace("NaT", "")
    summary["Profit"] = summary["Total_Sell"] - summary["Total_Cost"]
    summary["Profit Margin %"] = pct(summary["Profit"], summary["Total_Sell"])

    def is_closed(r) -> str:
        if not (r["Total_Cost"] > 0 and r["Total_Sell"] > 0):
            return "Open"
        pm = r["Profit Margin %"]
        if (pm < low_thr) or (pm > high_thr):
            return "Open"
        return "Closed"

    summary["Classification"] = summary.apply(is_closed, axis=1)

    def exception_type(r) -> str:
        if r["Total_Cost"] == 0 and r["Total_Sell"] == 0:
            return "Cost=Sell=0"
        if r["Total_Sell"] == 0:
            return "Revenue=0"
        if r["Total_Cost"] == 0:
            return "Cost=0"
        pm = r["Profit Margin %"]
        if (pm != 0) and ((pm < low_thr) or (pm > high_thr)):
            return MARGIN_LABEL
        return ""

    summary["Exception_Type"] = summary.apply(exception_type, axis=1)

    exceptions = summary[summary["Classification"].eq("Open")].copy()

    # ---- Client Summary ----
    client_summary = (
        df.groupby("Client", as_index=False)
        .agg(
            Total_Cost=("Cost Amount", "sum"),
            Total_Sell=("Sell Amount", "sum"),
            Line_Count=("Client", "size"),
            MAWB_Count=("MAWB", pd.Series.nunique),
            Latest_ETA=("ETA", "max"),
        )
    )
    client_summary["Profit"] = client_summary["Total_Sell"] - client_summary["Total_Cost"]
    client_summary["Profit Margin %"] = pct(client_summary["Profit"], client_summary["Total_Sell"])
    client_summary = client_summary.sort_values("Profit", ascending=False)

    # ---- Margin Outliers / Negative Profit ----
    margin_outliers = summary[
        ((summary["Profit Margin %"] < low_thr) | (summary["Profit Margin %"] > high_thr))
        & (summary["Profit Margin %"] != 0)
    ].copy().sort_values("Profit Margin %")

    negative_profit = summary[summary["Profit"] < 0].copy().sort_values("Profit")

    # ---- Zero buckets ----
    zero_margin = summary[summary["Profit Margin %"] == 0].copy().sort_values(
        ["Total_Sell", "Total_Cost"], ascending=False
    )
    zero_profit = summary[summary["Profit"] == 0].copy().sort_values(
        ["Total_Sell", "Total_Cost"], ascending=False
    )

    both_zero = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] == 0)].copy().sort_values("MAWB")
    sell_zero_only = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] > 0)].copy().sort_values(
        "Total_Cost", ascending=False
    )
    cost_zero_only = summary[(summary["Total_Cost"] == 0) & (summary["Total_Sell"] > 0)].copy().sort_values(
        "Total_Sell", ascending=False
    )

    # ---- Charge Code Summary ----
    chargecode_summary = (
        df.groupby("Charge Code", as_index=False)
        .agg(
            Total_Cost=("Cost Amount", "sum"),
            Total_Sell=("Sell Amount", "sum"),
            Line_Count=("Charge Code", "size"),
            MAWB_Count=("MAWB", pd.Series.nunique),
        )
    )
    chargecode_summary["Profit"] = chargecode_summary["Total_Sell"] - chargecode_summary["Total_Cost"]
    chargecode_summary["Profit Margin %"] = pct(chargecode_summary["Profit"], chargecode_summary["Total_Sell"])
    chargecode_summary = chargecode_summary.sort_values("Profit", ascending=False)

    # Charge code exception counts (MAWB-level flags)
    mawb_flags = summary[["MAWB", "Exception_Type"]].copy()
    mawb_charge = df[["MAWB", "Charge Code"]].drop_duplicates()
    cc_exc = mawb_charge.merge(mawb_flags, on="MAWB", how="left")

    chargecode_exceptions = (
        cc_exc.pivot_table(
            index="Charge Code",
            columns="Exception_Type",
            values="MAWB",
            aggfunc=pd.Series.nunique,
            fill_value=0,
        )
        .reset_index()
    )

    chargecode_summary = (
        chargecode_summary.merge(chargecode_exceptions, on="Charge Code", how="left")
        .fillna(0)
    )

    # ---- Vendor Summary ----
    vendor_summary = (
        df.groupby("Vendor", as_index=False)
        .agg(
            Total_Cost=("Cost Amount", "sum"),
            Total_Sell=("Sell Amount", "sum"),
            Line_Count=("Vendor", "size"),
            MAWB_Count=("MAWB", pd.Series.nunique),
        )
    )
    vendor_summary["Profit"] = vendor_summary["Total_Sell"] - vendor_summary["Total_Cost"]
    vendor_summary["Profit Margin %"] = pct(vendor_summary["Profit"], vendor_summary["Total_Sell"])
    vendor_summary = vendor_summary.sort_values("Profit", ascending=False)

    mawb_vendor = df[["MAWB", "Vendor"]].drop_duplicates()
    v_exc = mawb_vendor.merge(mawb_flags, on="MAWB", how="left")

    vendor_exceptions = (
        v_exc.pivot_table(
            index="Vendor",
            columns="Exception_Type",
            values="MAWB",
            aggfunc=pd.Series.nunique,
            fill_value=0,
        )
        .reset_index()
    )

    vendor_summary = vendor_summary.merge(vendor_exceptions, on="Vendor", how="left").fillna(0)

    # ---- Charge Code Profit <= 0 by MAWB ----
    cc_mawb = (
        df.groupby(["MAWB", "Charge Code"], as_index=False)
        .agg(
            Client=("Client", "first"),
            Vendor=("Vendor", "first"),
            Total_Cost=("Cost Amount", "sum"),
            Total_Sell=("Sell Amount", "sum"),
            ETA=("ETA", "max"),
        )
    )
    cc_mawb["Profit"] = cc_mawb["Total_Sell"] - cc_mawb["Total_Cost"]
    cc_mawb["Profit Margin %"] = pct(cc_mawb["Profit"], cc_mawb["Total_Sell"])
    cc_mawb["ETA Month"] = pd.to_datetime(cc_mawb["ETA"], errors="coerce").dt.to_period("M").astype(str).replace("NaT", "")

    chargecode_profit_le0_mawb = (
        cc_mawb[cc_mawb["Profit"] <= 0]
        .copy()
        .sort_values(["Profit", "Total_Sell"], ascending=[True, False])
    )

    # ---- KPI / Summary numbers ----
    total_mawb = int(len(summary))
    closed_cnt = int((summary["Classification"] == "Closed").sum())
    open_cnt = total_mawb - closed_cnt

    total_sell_sum = float(summary["Total_Sell"].sum())
    total_profit_sum = float(summary["Profit"].sum())
    overall_pm = (total_profit_sum / total_sell_sum) if total_sell_sum else 0

    neg_profit_cnt = int((summary["Profit"] < 0).sum())
    neg_profit_amt = float(summary.loc[summary["Profit"] < 0, "Profit"].sum())
    neg_profit_ratio = (neg_profit_cnt / total_mawb) if total_mawb else 0

    eta_filled_ratio = float((summary["ETA"].notna().sum() / total_mawb)) if total_mawb else 0

    kpi_dict = {
        "Total MAWB": total_mawb,
        "Closed Count": closed_cnt,
        "Closed %": (closed_cnt / total_mawb) if total_mawb else 0,
        "Open Count": open_cnt,
        "Revenue=0 Count": int((summary["Exception_Type"] == "Revenue=0").sum()),
        "Cost=0 Count": int((summary["Exception_Type"] == "Cost=0").sum()),
        "Cost=Sell=0 Count": int((summary["Exception_Type"] == "Cost=Sell=0").sum()),
        f"{MARGIN_LABEL} Count": int((summary["Exception_Type"] == MARGIN_LABEL).sum()),
        "Total Cost": float(summary["Total_Cost"].sum()),
        "Total Sell": total_sell_sum,
        "Total Profit": total_profit_sum,
        "Overall Profit Margin %": overall_pm,
        "ETA Filled %": eta_filled_ratio,
    }
    KPI_PCT_KEYS = {"Closed %", "Overall Profit Margin %", "ETA Filled %"}
    kpi_vertical = _make_kpi_vertical(kpi_dict, KPI_PCT_KEYS)

    neg_summary = pd.DataFrame(
        [
            {"Metric": "Profit < 0 Count", "Value": neg_profit_cnt},
            {"Metric": "Profit < 0 Total Amount", "Value": neg_profit_amt},
            {"Metric": "Profit < 0 % of MAWBs", "Value": f"{neg_profit_ratio*100:.2f}%"},
        ]
    )

    # ---- Display versions for Streamlit ----
    display_summary = display_df(summary, date_cols=["ETA"])
    display_exceptions = display_df(exceptions, date_cols=["ETA"])
    display_client_summary = display_df(client_summary, date_cols=["Latest_ETA"])
    display_margin_outliers = display_df(margin_outliers, date_cols=["ETA"])
    display_negative_profit = display_df(negative_profit, date_cols=["ETA"])
    display_zero_margin = display_df(zero_margin, date_cols=["ETA"])
    display_zero_profit = display_df(zero_profit, date_cols=["ETA"])
    display_both_zero = display_df(both_zero, date_cols=["ETA"])
    display_sell_zero_only = display_df(sell_zero_only, date_cols=["ETA"])
    display_cost_zero_only = display_df(cost_zero_only, date_cols=["ETA"])
    display_chargecode_summary = display_df(chargecode_summary)
    display_vendor_summary = display_df(vendor_summary)
    display_chargecode_profit_le0_mawb = display_df(chargecode_profit_le0_mawb, date_cols=["ETA"])

    return AuditResult(
        mawb_keep=mawb_keep,
        mawb_not_found_df=mawb_not_found_df,
        eta_parse_note=eta_parse_note,
        kpi_vertical=kpi_vertical,
        neg_summary=neg_summary,
        df=df,
        summary=summary,
        exceptions=exceptions,
        client_summary=client_summary,
        margin_outliers=margin_outliers,
        negative_profit=negative_profit,
        zero_margin=zero_margin,
        zero_profit=zero_profit,
        both_zero=both_zero,
        sell_zero_only=sell_zero_only,
        cost_zero_only=cost_zero_only,
        chargecode_summary=chargecode_summary,
        vendor_summary=vendor_summary,
        chargecode_profit_le0_mawb=chargecode_profit_le0_mawb,
        mawb_not_found=mawb_not_found,
        margin_label=MARGIN_LABEL,
        display_summary=display_summary,
        display_exceptions=display_exceptions,
        display_client_summary=display_client_summary,
        display_margin_outliers=display_margin_outliers,
        display_negative_profit=display_negative_profit,
        display_zero_margin=display_zero_margin,
        display_zero_profit=display_zero_profit,
        display_both_zero=display_both_zero,
        display_sell_zero_only=display_sell_zero_only,
        display_cost_zero_only=display_cost_zero_only,
        display_chargecode_summary=display_chargecode_summary,
        display_vendor_summary=display_vendor_summary,
        display_chargecode_profit_le0_mawb=display_chargecode_profit_le0_mawb,
    )
