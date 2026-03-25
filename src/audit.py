from __future__ import annotations

import re
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
    mawb_keep: list[str]
    mawb_not_found_df: pd.DataFrame
    eta_parse_note: str | None

    kpi_vertical: pd.DataFrame
    neg_summary: pd.DataFrame

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


def _clean_text_value(x) -> str:
    s = str(x).strip()
    if s.lower() in {"", "nan", "none"}:
        return "UNKNOWN"
    return s


def _extract_last8(mawb: str) -> str:
    """
    例如:
    125-22479096 -> 22479096
    777-22479096 -> 22479096
    """
    s = str(mawb).strip()
    m = re.search(r"(\d{8})$", s)
    return m.group(1) if m else ""


def _apply_procaresx_match(df: pd.DataFrame) -> pd.DataFrame:
    """
    对 PROCARESX:
    若存在 125-XXXXXXXX 和 777-XXXXXXXX 且最后8位相同，
    则将 777-XXXXXXXX 全部并到 125-XXXXXXXX 下做后续分析。
    """
    out = df.copy()

    if "Client" not in out.columns or "MAWB" not in out.columns:
        return out

    mask = out["Client"].astype(str).str.upper().eq("PROCARESX")
    if not mask.any():
        return out

    p = out.loc[mask, ["MAWB"]].copy()
    p["MAWB_STR"] = p["MAWB"].astype(str).str.strip()
    p["LAST8"] = p["MAWB_STR"].apply(_extract_last8)
    p["PREFIX3"] = p["MAWB_STR"].str[:3]

    have_125 = set(p.loc[p["PREFIX3"] == "125", "LAST8"])
    have_777 = set(p.loc[p["PREFIX3"] == "777", "LAST8"])
    matched_last8 = have_125 & have_777

    if not matched_last8:
        return out

    def remap_mawb(row):
        if str(row["Client"]).strip().upper() != "PROCARESX":
            return row["MAWB"]

        mawb = str(row["MAWB"]).strip()
        last8 = _extract_last8(mawb)

        if mawb.startswith("777") and last8 in matched_last8:
            return f"125-{last8}"
        return row["MAWB"]

    out["MAWB"] = out.apply(remap_mawb, axis=1)
    return out


def _build_mawb_active_codes(df: pd.DataFrame) -> dict[str, set[str]]:
    """
    MAWB 下，只要 cost 或 sell 非0 的 code，就认为 active。
    """
    d = df.loc[
        (df["Cost Amount"].fillna(0) != 0) | (df["Sell Amount"].fillna(0) != 0),
        ["MAWB", "Charge Code"],
    ].copy()

    d["Charge Code"] = d["Charge Code"].astype(str).str.upper().str.strip()
    out = (
        d.groupby("MAWB")["Charge Code"]
        .apply(lambda s: set(x for x in s if x and x != "UNKNOWN"))
        .to_dict()
    )
    return out


def _classify_margin_exception(
    client: str,
    active_codes: set[str],
    profit_margin: float,
    low_thr: float,
    high_thr: float,
) -> str:
    """
    新规则说明：
    1) 删除旧的 HANCAIWUX / 4PXDIGHKG 高毛利豁免逻辑
    2) 新增：
       - HANCAIWUX 和 4PXDIGHKG：
         当 active code 仅有 THAWB / DDOC 时，profit margin < 85% 记异常
       - SHELIUSZX 和 LIBEXPLHR：
         当 active code 中没有 THAWB 时，profit margin > 35% 记异常
    3) 其他情况保留默认规则：
       - margin < 30% 或 > 80%
    """
    client_u = str(client).strip().upper()
    codes = {str(x).strip().upper() for x in (active_codes or set()) if str(x).strip()}
    special_codes = {"THAWB", "DDOC"}

    if client_u in {"HANCAIWUX", "4PXDIGHKG"}:
        if codes and codes.issubset(special_codes):
            if pd.notna(profit_margin) and profit_margin < 0.85:
                return "Margin<85%"
            return ""

    if client_u in {"SHELIUSZX", "LIBEXPLHR"}:
        if "THAWB" not in codes:
            if pd.notna(profit_margin):
                if profit_margin > 0.35:
                    return "Margin>35%"
                if profit_margin < low_thr:
                    return f"Margin<{int(low_thr*100)}%"
            return ""

    if pd.notna(profit_margin):
        if profit_margin > high_thr:
            return f"Margin>{int(high_thr*100)}%"
        if profit_margin < low_thr:
            return f"Margin<{int(low_thr*100)}%"

    return ""


def _keep_chargecode_negative_exception(row) -> bool:
    """
    code级负利润异常规则：
    1) WHALECBOS 的 TISC/TABD/DSTOR：profit < -10 才算异常
    2) 所有客户的 TISC：只有 profit < -10 才算异常
    3) 其他 code：只有 profit < 0 才算异常
    """
    client = str(row.get("Client", "")).strip().upper()
    code = str(row.get("Charge Code", "")).strip().upper()
    profit_val = row.get("Profit", 0)
    profit = float(profit_val) if pd.notna(profit_val) else 0.0

    if code == "TISC":
        return profit < -10

    if client == "WHALECBOS" and code in {"TABD", "DSTOR"}:
        return profit < -10

    return profit < 0


@st.cache_data(show_spinner=False)
def run_audit(
    billing_file,
    eta_file,
    mawb_text: str,
    low_thr: float = 0.30,
    high_thr: float = 0.80,
) -> AuditResult:
    MARGIN_LABEL = f"Margin<{int(low_thr*100)}% or >{int(high_thr*100)}%"

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

    df = raw_df.copy()
    df["MAWB"] = df[mawb_col].apply(normalize_mawb)
    df["Cost Amount"] = safe_numeric(df[cost_col])
    df["Sell Amount"] = safe_numeric(df[sell_col])

    if client_col:
        df["Client"] = df[client_col].apply(_clean_text_value)
    else:
        df["Client"] = "UNKNOWN"

    if charge_code_col:
        df["Charge Code"] = df[charge_code_col].apply(_clean_text_value)
    else:
        df["Charge Code"] = "UNKNOWN"

    if vendor_col:
        df["Vendor"] = df[vendor_col].apply(_clean_text_value)
    else:
        df["Vendor"] = "UNKNOWN"

    df = df[df["MAWB"].ne("")].copy()

    mawb_keep = parse_mawb_list(mawb_text)
    if mawb_keep:
        df = df[df["MAWB"].isin(mawb_keep)].copy()
        found_set = set(df["MAWB"].unique())
        mawb_not_found = sorted(set(mawb_keep) - found_set)
        mawb_not_found_df = pd.DataFrame({"MAWB": mawb_not_found})
    else:
        mawb_not_found = []
        mawb_not_found_df = pd.DataFrame({"MAWB": []})

    df = _apply_procaresx_match(df)

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
                mdf = pd.DataFrame()
                mdf["MAWB"] = mdf0[m_mawb].apply(normalize_mawb)
                mdf["ETA"] = clean_eta_series(mdf0[m_eta])

                branch_col = find_first_col(mdf0, ["Destination", "Dest", "Branch", "Station", "To"])
                if branch_col:
                    mdf["Branch"] = mdf0[branch_col].apply(_clean_text_value).str.upper()
                elif len(mdf0.columns) >= 6:
                    mdf["Branch"] = mdf0.iloc[:, 5].apply(_clean_text_value).str.upper()
                else:
                    mdf["Branch"] = "UNKNOWN"

                bad_eta_rows = int(mdf["ETA"].isna().sum())
                total_rows = int(len(mdf))
                if total_rows > 0 and bad_eta_rows > 0:
                    eta_parse_note = (
                        f"ETA parsing note: {bad_eta_rows} / {total_rows} ETA values could not be parsed and were left blank."
                    )

                eta_map = (
                    mdf.dropna(subset=["MAWB"])
                    .groupby("MAWB", as_index=False)
                    .agg(
                        ETA=("ETA", "max"),
                        Branch=("Branch", "first"),
                    )
                )

    if eta_map is not None and not eta_map.empty:
        df = df.merge(eta_map, on="MAWB", how="left")
    else:
        df["ETA"] = pd.NaT
        df["Branch"] = "UNKNOWN"

    df["ETA"] = pd.to_datetime(df["ETA"], errors="coerce").dt.normalize()
    df["Branch"] = df["Branch"].fillna("UNKNOWN").astype(str).str.upper()

    mawb_active_codes = _build_mawb_active_codes(df)

    summary = (
        df.groupby("MAWB", as_index=False)
        .agg(
            Client=("Client", "first"),
            Branch=("Branch", "first"),
            Total_Cost=("Cost Amount", "sum"),
            Total_Sell=("Sell Amount", "sum"),
            Line_Count=("MAWB", "size"),
            ETA=("ETA", "max"),
        )
    )

    summary["ETA Month"] = summary["ETA"].dt.to_period("M").astype(str).replace("NaT", "")
    summary["Profit"] = summary["Total_Sell"] - summary["Total_Cost"]
    summary["Profit Margin %"] = pct(summary["Profit"], summary["Total_Sell"])

    def exception_type(r) -> str:
        if r["Total_Cost"] == 0 and r["Total_Sell"] == 0:
            return "Cost=Sell=0"
        if r["Total_Sell"] == 0:
            return "Revenue=0"
        if r["Total_Cost"] == 0:
            return "Cost=0"

        active_codes = mawb_active_codes.get(r["MAWB"], set())
        return _classify_margin_exception(
            client=r["Client"],
            active_codes=active_codes,
            profit_margin=r["Profit Margin %"],
            low_thr=low_thr,
            high_thr=high_thr,
        )

    summary["Exception_Type"] = summary.apply(exception_type, axis=1)
    summary["Classification"] = summary["Exception_Type"].apply(lambda x: "Closed" if x == "" else "Open")

    exceptions = summary[summary["Classification"].eq("Open")].copy()

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

    margin_outliers = summary[
        summary["Exception_Type"].astype(str).str.startswith("Margin")
    ].copy().sort_values("Profit Margin %")

    negative_profit = summary[summary["Profit"] < 0].copy().sort_values("Profit")

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

    cc_mawb = (
        df.groupby(["MAWB", "Charge Code"], as_index=False)
        .agg(
            Client=("Client", "first"),
            Branch=("Branch", "first"),
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
        cc_mawb[
            ~((cc_mawb["Total_Cost"] == 0) & (cc_mawb["Total_Sell"] == 0))
        ]
        .copy()
    )

    chargecode_profit_le0_mawb = chargecode_profit_le0_mawb[
        chargecode_profit_le0_mawb.apply(_keep_chargecode_negative_exception, axis=1)
    ].copy()

    chargecode_profit_le0_mawb = chargecode_profit_le0_mawb.sort_values(
        ["Profit", "Total_Sell"], ascending=[True, False]
    )

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

    cc_profit_neg = (
        chargecode_profit_le0_mawb[chargecode_profit_le0_mawb["Profit"] < 0]
        .groupby("Charge Code")["MAWB"]
        .nunique()
        .reset_index(name="Profit<0")
    )

    chargecode_summary = (
        chargecode_summary
        .merge(chargecode_exceptions, on="Charge Code", how="left")
        .merge(cc_profit_neg, on="Charge Code", how="left")
        .fillna(0)
    )

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
    v_exc = mawb_vendor.merge(summary[["MAWB", "Exception_Type"]], on="MAWB", how="left")

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
    }

    margin_labels = sorted(
        x for x in summary["Exception_Type"].dropna().astype(str).unique()
        if x.startswith("Margin")
    )
    for label in margin_labels:
        kpi_dict[f"{label} Count"] = int((summary["Exception_Type"] == label).sum())

    kpi_dict.update({
        "Total Cost": float(summary["Total_Cost"].sum()),
        "Total Sell": total_sell_sum,
        "Total Profit": total_profit_sum,
        "Overall Profit Margin %": overall_pm,
        "ETA Filled %": eta_filled_ratio,
    })

    KPI_PCT_KEYS = {"Closed %", "Overall Profit Margin %", "ETA Filled %"}
    kpi_vertical = _make_kpi_vertical(kpi_dict, KPI_PCT_KEYS)

    neg_summary = pd.DataFrame(
        [
            {"Metric": "Profit < 0 Count", "Value": neg_profit_cnt},
            {"Metric": "Profit < 0 Total Amount", "Value": neg_profit_amt},
            {"Metric": "Profit < 0 % of MAWBs", "Value": f"{neg_profit_ratio*100:.2f}%"},
        ]
    )

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
