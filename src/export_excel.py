import io
import pandas as pd
from .helpers import to_date_only

PERCENT_FMT = "0.00%"
NUMBER_FMT = "#,##0.00"


def _set_col_format(ws, workbook, col_idx: int, width: int, num_format: str, startcol: int = 0):
    fmt = workbook.add_format({"num_format": num_format})
    ws.set_column(startcol + col_idx, startcol + col_idx, width, fmt)


def build_excel_report(result) -> bytes:
    """
    Build a multi-tab Excel report, with:
    - Analysis Summary (hyperlinks + KPI table + embedded ChargeCode/Vendor)
    - Detail sheets
    - Profit Margin % columns formatted as percent where numeric
    """
    output = io.BytesIO()

    # Prepare export copies (date-only)
    summary_x = to_date_only(result.summary, ["ETA"])
    exceptions_x = to_date_only(result.exceptions, ["ETA"])
    client_summary_x = to_date_only(result.client_summary, ["Latest_ETA"])
    margin_outliers_x = to_date_only(result.margin_outliers, ["ETA"])
    negative_profit_x = to_date_only(result.negative_profit, ["ETA"])
    zero_margin_x = to_date_only(result.zero_margin, ["ETA"])
    zero_profit_x = to_date_only(result.zero_profit, ["ETA"])
    both_zero_x = to_date_only(result.both_zero, ["ETA"])
    sell_zero_only_x = to_date_only(result.sell_zero_only, ["ETA"])
    cost_zero_only_x = to_date_only(result.cost_zero_only, ["ETA"])
    df_x = to_date_only(result.df, ["ETA"])
    chargecode_profit_le0_mawb_x = to_date_only(result.chargecode_profit_le0_mawb, ["ETA"])

    # For formatting (which sheets have Profit Margin %)
    percent_sheets = {
        "Exceptions": exceptions_x,
        "MAWB_Summary": summary_x,
        "Client_Summary": client_summary_x,
        "Margin_Outliers": margin_outliers_x,
        "Negative_Profit": negative_profit_x,
        "Zero_Margin": zero_margin_x,
        "Zero_Profit": zero_profit_x,
        "Both_Zero": both_zero_x,
        "Sell_Zero_Only": sell_zero_only_x,
        "Cost_Zero_Only": cost_zero_only_x,
        "ChargeCode_Summary": result.chargecode_summary,
        "Vendor_Summary": result.vendor_summary,
        "ChargeCode_ProfitLE0_MAWB": chargecode_profit_le0_mawb_x,
    }

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({"bold": True, "font_size": 14})
        subheader_fmt = workbook.add_format({"bold": True, "font_size": 12})
        bold_fmt = workbook.add_format({"bold": True})
        percent_fmt = workbook.add_format({"num_format": PERCENT_FMT})
        number_fmt = workbook.add_format({"num_format": NUMBER_FMT})

        # ---------------- Analysis Summary sheet ----------------
        ws = workbook.add_worksheet("Analysis Summary")
        writer.sheets["Analysis Summary"] = ws

        ws.write(0, 0, "Analysis Summary", header_fmt)

        link_start_row = 2
        ws.write(link_start_row, 0, "This page provides an overview. Click detail links below:", bold_fmt)

        tab_links = [
            ("Open exceptions overview + detail", "Exceptions"),
            ("MAWB level summary + detail", "MAWB_Summary"),
            ("Client margin summary + detail", "Client_Summary"),
            (f"Margin anomalies ({result.margin_label}) + detail", "Margin_Outliers"),
            ("Negative profit MAWBs + detail", "Negative_Profit"),
            ("Zero margin tickets + detail", "Zero_Margin"),
            ("Zero profit tickets + detail", "Zero_Profit"),
            ("Cost=Sell=0 tickets + detail", "Both_Zero"),
            ("Sell=0 only tickets + detail", "Sell_Zero_Only"),
            ("Cost=0 only tickets + detail", "Cost_Zero_Only"),
            ("Charge code summary + detail", "ChargeCode_Summary"),
            ("Vendor summary + detail", "Vendor_Summary"),
            ("ChargeCode Profit<=0 by MAWB + detail", "ChargeCode_ProfitLE0_MAWB"),
            ("Raw enriched billing + detail", "Raw_Billing_Enriched"),
        ]
        if result.mawb_keep:
            tab_links.insert(0, ("MAWB not found from filter + detail", "MAWB_Not_Found"))

        r = link_start_row + 1
        for text, sheet_name in tab_links:
            ws.write_url(r, 0, f"internal:'{sheet_name}'!A1", string=text)
            r += 1

        # KPI block
        kpi_row = r + 1
        ws.write(kpi_row, 0, "KPI (two-column)", subheader_fmt)
        ws.write(kpi_row + 1, 0, "Metric", bold_fmt)
        ws.write(kpi_row + 1, 1, "Value", bold_fmt)

        # We rebuild KPI dict from the vertical table (Metric/Value)
        kpi_write_row = kpi_row + 2
        for i in range(len(result.kpi_vertical)):
            metric = result.kpi_vertical.loc[i, "Metric"]
            value = result.kpi_vertical.loc[i, "Value"]
            ws.write(kpi_write_row + i, 0, metric)

            # If looks like percent string "xx.xx%"
            if isinstance(value, str) and value.endswith("%"):
                try:
                    v = float(value.replace("%", "")) / 100.0
                    ws.write_number(kpi_write_row + i, 1, v, percent_fmt)
                except Exception:
                    ws.write(kpi_write_row + i, 1, value)
            else:
                # try numeric
                try:
                    ws.write_number(kpi_write_row + i, 1, float(value), number_fmt)
                except Exception:
                    ws.write(kpi_write_row + i, 1, str(value))

        # Negative profit summary
        neg_row = kpi_write_row + len(result.kpi_vertical) + 2
        ws.write(neg_row, 0, "Summary: Profit < 0", subheader_fmt)
        ws.write(neg_row + 1, 0, "Metric", bold_fmt)
        ws.write(neg_row + 1, 1, "Value", bold_fmt)

        for i in range(len(result.neg_summary)):
            ws.write(neg_row + 2 + i, 0, result.neg_summary.loc[i, "Metric"])
            v = result.neg_summary.loc[i, "Value"]
            if isinstance(v, str) and v.endswith("%"):
                try:
                    vv = float(v.replace("%", "")) / 100.0
                    ws.write_number(neg_row + 2 + i, 1, vv, percent_fmt)
                except Exception:
                    ws.write(neg_row + 2 + i, 1, v)
            else:
                try:
                    ws.write_number(neg_row + 2 + i, 1, float(v), number_fmt)
                except Exception:
                    ws.write(neg_row + 2 + i, 1, str(v))

        # Embed ChargeCode_Summary + Vendor_Summary (preview)
        cc_row = neg_row + 6
        ws.write(cc_row, 0, "ChargeCode_Summary (embedded)", subheader_fmt)
        result.chargecode_summary.to_excel(writer, index=False, sheet_name="Analysis Summary", startrow=cc_row + 1, startcol=0)

        # Format Profit Margin % column in embedded CC table
        try:
            pm_idx = list(result.chargecode_summary.columns).index("Profit Margin %")
            _set_col_format(ws, workbook, pm_idx, width=16, num_format=PERCENT_FMT, startcol=0)
        except Exception:
            pass

        v_row = cc_row + 2 + len(result.chargecode_summary) + 3
        ws.write(v_row, 0, "Vendor_Summary (embedded)", subheader_fmt)
        result.vendor_summary.to_excel(writer, index=False, sheet_name="Analysis Summary", startrow=v_row + 1, startcol=0)

        # Format Profit Margin % in embedded vendor table too
        try:
            pm_idx_v = list(result.vendor_summary.columns).index("Profit Margin %")
            _set_col_format(ws, workbook, pm_idx_v, width=16, num_format=PERCENT_FMT, startcol=0)
        except Exception:
            pass

        # ---------------- Detail sheets ----------------
        exceptions_x.to_excel(writer, index=False, sheet_name="Exceptions")
        summary_x.to_excel(writer, index=False, sheet_name="MAWB_Summary")
        client_summary_x.to_excel(writer, index=False, sheet_name="Client_Summary")
        margin_outliers_x.to_excel(writer, index=False, sheet_name="Margin_Outliers")
        negative_profit_x.to_excel(writer, index=False, sheet_name="Negative_Profit")
        zero_margin_x.to_excel(writer, index=False, sheet_name="Zero_Margin")
        zero_profit_x.to_excel(writer, index=False, sheet_name="Zero_Profit")
        both_zero_x.to_excel(writer, index=False, sheet_name="Both_Zero")
        sell_zero_only_x.to_excel(writer, index=False, sheet_name="Sell_Zero_Only")
        cost_zero_only_x.to_excel(writer, index=False, sheet_name="Cost_Zero_Only")
        result.chargecode_summary.to_excel(writer, index=False, sheet_name="ChargeCode_Summary")
        result.vendor_summary.to_excel(writer, index=False, sheet_name="Vendor_Summary")
        chargecode_profit_le0_mawb_x.to_excel(writer, index=False, sheet_name="ChargeCode_ProfitLE0_MAWB")

        if result.mawb_keep:
            result.mawb_not_found_df.to_excel(writer, index=False, sheet_name="MAWB_Not_Found")

        df_x.to_excel(writer, index=False, sheet_name="Raw_Billing_Enriched")

        # Apply percent format to Profit Margin % column in all relevant sheets
        for sh, dfx in percent_sheets.items():
            if sh in writer.sheets and "Profit Margin %" in dfx.columns:
                ws2 = writer.sheets[sh]
                pm_col = list(dfx.columns).index("Profit Margin %")
                _set_col_format(ws2, workbook, pm_col, width=16, num_format=PERCENT_FMT, startcol=0)

        # Make columns a bit wider for readability (optional quick wins)
        for sh in writer.sheets:
            wsx = writer.sheets[sh]
            wsx.set_column(0, 0, 18)  # first col
            wsx.set_column(1, 5, 16)  # middle cols

    return output.getvalue()
