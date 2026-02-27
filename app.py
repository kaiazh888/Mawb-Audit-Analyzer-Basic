import streamlit as st
from src.audit import run_audit
from src.export_excel import build_excel_report

st.set_page_config(page_title="MAWB Audit Analyzer", layout="wide")

st.title("MAWB Audit Analyzer (Billing-only)")
st.caption(
    "Upload Billing charges export + optional MAWB→ETA mapping file. "
    "Supports MAWB filter, profit margin analysis, zero buckets, outliers, negative profit, "
    "and Charge Code / Vendor summaries."
)

# ---------------- Sidebar controls ----------------
with st.sidebar:
    st.header("Inputs")

    billing_file = st.file_uploader(
        "Billing Charges Excel (.xlsx)",
        type=["xlsx"],
        key="billing",
    )

    eta_file = st.file_uploader(
        "Optional: MAWB→ETA mapping Excel (.xlsx)",
        type=["xlsx"],
        key="eta_mapping",
    )

    st.divider()

    st.subheader("Optional MAWB Filter")
    mawb_text = st.text_area(
        "Paste MAWBs (comma/space/newline). Supports 99934022122 → 999-34022122.",
        height=140,
        placeholder="Example:\n999-34022122\n99934022133\n999 34022144",
        key="mawb_filter",
    )

    st.divider()

    st.subheader("Rules")
    low_thr = st.number_input(
        "Margin low threshold",
        min_value=0.0,
        max_value=1.0,
        value=0.30,
        step=0.01,
    )
    high_thr = st.number_input(
        "Margin high threshold",
        min_value=0.0,
        max_value=1.0,
        value=0.80,
        step=0.01,
    )
    st.caption("Rule: if (Cost>0 and Sell>0) AND (PM < low OR PM > high) => Open; else Closed.")

# ---------------- Main ----------------
if not billing_file:
    st.info("Please upload a Billing Charges Excel file to start.")
    st.stop()

try:
    result = run_audit(
        billing_file=billing_file,
        eta_file=eta_file,
        mawb_text=mawb_text,
        low_thr=float(low_thr),
        high_thr=float(high_thr),
    )

    # Notes
    if result.eta_parse_note:
        st.info(result.eta_parse_note)

    if result.mawb_keep:
        st.subheader("MAWB Not Found (in uploaded Billing file)")
        st.dataframe(result.mawb_not_found_df, use_container_width=True)

    # KPI
    st.subheader("Analysis Summary (KPI)")
    st.dataframe(result.kpi_vertical, use_container_width=True)

    st.subheader("Summary: Profit < 0 (Count / Amount / Ratio)")
    st.dataframe(result.neg_summary, use_container_width=True)

    # Main tables
    st.subheader("Exceptions (Open items)")
    st.dataframe(result.display_exceptions, use_container_width=True)

    st.subheader("MAWB Summary (All)")
    st.dataframe(result.display_summary, use_container_width=True)

    st.subheader("Client Profit Summary")
    st.dataframe(result.display_client_summary, use_container_width=True)

    st.subheader("Profit Margin Outliers (PM!=0)")
    st.dataframe(result.display_margin_outliers, use_container_width=True)

    st.subheader("Negative Profit (Profit < 0)")
    st.dataframe(result.display_negative_profit, use_container_width=True)

    st.subheader("Zero Margin (Profit Margin % = 0)")
    st.dataframe(result.display_zero_margin, use_container_width=True)

    st.subheader("Zero Profit (Profit = 0)")
    st.dataframe(result.display_zero_profit, use_container_width=True)

    st.subheader("Cost=Sell=0 (Both Zero)")
    st.dataframe(result.display_both_zero, use_container_width=True)

    st.subheader("Sell=0 ONLY (Total_Sell=0 and Total_Cost>0)")
    st.dataframe(result.display_sell_zero_only, use_container_width=True)

    st.subheader("Cost=0 ONLY (Total_Cost=0 and Total_Sell>0)")
    st.dataframe(result.display_cost_zero_only, use_container_width=True)

    st.subheader("Charge Code Summary")
    st.dataframe(result.display_chargecode_summary, use_container_width=True)

    st.subheader("Vendor Summary")
    st.dataframe(result.display_vendor_summary, use_container_width=True)

    st.subheader("Charge Code Profit <= 0 (by MAWB)")
    st.dataframe(result.display_chargecode_profit_le0_mawb, use_container_width=True)

    # Export
    st.divider()
    st.subheader("Export")
    excel_bytes = build_excel_report(result)

    st.download_button(
        "Download Report Excel",
        data=excel_bytes,
        file_name="MAWB_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.exception(e)
