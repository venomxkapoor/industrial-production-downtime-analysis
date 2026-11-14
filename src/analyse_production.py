"""Analyse real beverage bottling production and downtime records."""

from pathlib import Path
import textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.image as mpimg
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
import pandas as pd

REQUIRED_COLS = [
    "date", "product_type_l", "production_units", "liters_produced",
    "production_start_time", "production_end_time", "efficiency",
    "gallons_per_hour", "monitored_time_dec", "operation_time_dec", "pause_time_dec",
]

def safe_pct(numerator, denominator):
    """Return numerator / denominator * 100, or 0.0 when denominator <= 0."""
    return 0.0 if denominator <= 0 else (numerator / denominator) * 100.0

def safe_rate(numerator, denominator):
    """Return numerator / denominator, or 0.0 when denominator <= 0."""
    return 0.0 if denominator <= 0 else numerator / denominator

def load_and_check_data(xlsx_path):
    """Load daily_operation_summary and event count from production_raw.xlsx."""
    df = pd.read_excel(xlsx_path, sheet_name="daily_operation_summary")
    events_df = pd.read_excel(xlsx_path, sheet_name="downtime_event_log")
    missing_cols = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing expected columns: {missing_cols}")
    if int(df[REQUIRED_COLS].isna().sum().sum()) > 0:
        raise ValueError("Source sheet contains missing values.")
    df["date"] = pd.to_datetime(df["date"])
    if (df["monitored_time_dec"] <= 0).any() or (df["operation_time_dec"] <= 0).any():
        raise ValueError("Monitored and operating hours must be positive.")
    df = df.sort_values(["date", "production_start_time"]).reset_index(drop=True)
    return df, len(events_df), int(events_df["date"].nunique())

def calculate_run_metrics(df):
    """Add independently calculated efficiency, downtime, and production rates."""
    out = df.copy()
    out["date_str"] = out["date"].dt.strftime("%Y-%m-%d")
    out["year_month"] = out["date"].dt.to_period("M").astype(str)
    out["source_efficiency_pct"] = out["efficiency"] * 100.0
    calc_eff, dt_pct, u_rate, l_rate = [], [], [], []
    for _, row in out.iterrows():
        calc_eff.append(safe_pct(row["operation_time_dec"], row["monitored_time_dec"]))
        dt_pct.append(safe_pct(row["pause_time_dec"], row["monitored_time_dec"]))
        u_rate.append(safe_rate(row["production_units"], row["operation_time_dec"]))
        l_rate.append(safe_rate(row["liters_produced"], row["operation_time_dec"]))
    out["calc_efficiency_pct"] = calc_eff
    out["downtime_pct"] = dt_pct
    out["calc_units_per_op_hour"] = u_rate
    out["calc_liters_per_op_hour"] = l_rate
    return out

def summarize_by_group(df, group_col):
    """Aggregate production, time, and weighted/mean KPIs by group_col."""
    rows = []
    for key, g in df.groupby(group_col, sort=True):
        mon, op, pause = g["monitored_time_dec"].sum(), g["operation_time_dec"].sum(), g["pause_time_dec"].sum()
        units, liters = int(g["production_units"].sum()), int(g["liters_produced"].sum())
        rows.append({
            group_col: key, "runs": len(g), "unique_dates": int(g["date_str"].nunique()),
            "production_units": units, "liters_produced": liters,
            "monitored_hours": round(mon, 2), "operating_hours": round(op, 2), "pause_hours": round(pause, 2),
            "weighted_eff_pct": round(safe_pct(op, mon), 2), "mean_eff_pct": round(g["calc_efficiency_pct"].mean(), 2),
            "weighted_downtime_pct": round(safe_pct(pause, mon), 2), "mean_downtime_pct": round(g["downtime_pct"].mean(), 2),
            "overall_units_per_op_h": round(safe_rate(units, op), 2),
            "mean_units_per_op_h": round(g["calc_units_per_op_hour"].mean(), 2),
            "mean_liters_per_op_h": round(g["calc_liters_per_op_hour"].mean(), 2),
        })
    return pd.DataFrame(rows)

def build_summaries(run_df, event_count, event_days):
    """Compute 56-day calendar totals, overall KPIs, product summary, and monthly summary."""
    daily_df = summarize_by_group(run_df, "date_str").drop(columns=["unique_dates"])
    product_df = summarize_by_group(run_df, "product_type_l")
    monthly_df = summarize_by_group(run_df, "year_month")
    mon = run_df["monitored_time_dec"].sum()
    op = run_df["operation_time_dec"].sum()
    pause = run_df["pause_time_dec"].sum()
    units, liters = int(run_df["production_units"].sum()), int(run_df["liters_produced"].sum())
    kpi_rows = [
        ("Production Runs Recorded", len(run_df), "runs", "Rows in daily_operation_summary"),
        ("Unique Operating Dates", int(run_df["date_str"].nunique()), "days", "3 dates had 2 non-overlapping runs"),
        ("Logged Stop Events", event_count, "events", f"Rows in downtime_event_log ({event_days} dates)"),
        ("Total Production Units", units, "bottles", "3L and 5L bottles combined"),
        ("Total Volume Produced", liters, "liters", "Sum of liters_produced (units x size)"),
        ("Total Monitored Time", round(mon, 2), "hours", "Sum of monitored_time_dec"),
        ("Total Operating Time", round(op, 2), "hours", "Sum of operation_time_dec"),
        ("Total Pause / Downtime", round(pause, 2), "hours", "Sum of pause_time_dec"),
        ("Time-Weighted Efficiency", round(safe_pct(op, mon), 2), "%", "Total operating / total monitored x 100"),
        ("Mean Run Efficiency", round(run_df["calc_efficiency_pct"].mean(), 2), "%", "Simple mean across 59 runs"),
        ("Time-Weighted Downtime", round(safe_pct(pause, mon), 2), "%", "Total pause / total monitored x 100"),
        ("Mean Run Downtime", round(run_df["downtime_pct"].mean(), 2), "%", "Simple mean across 59 runs"),
        ("Overall Operating Rate", round(safe_rate(units, op), 2), "units/op h", "Total units / total operating hours"),
        ("Mean Run Operating Rate", round(run_df["calc_units_per_op_hour"].mean(), 2), "units/op h", "Simple mean across 59 runs"),
    ]
    kpi_df = pd.DataFrame(kpi_rows, columns=["Metric", "Value", "Unit", "Notes"])
    return daily_df, kpi_df, product_df, monthly_df

def save_charts(daily_df, monthly_df, charts_dir):
    """Create two clear charts for daily output/time and monthly efficiency/downtime."""
    charts_dir.mkdir(parents=True, exist_ok=True)
    x = range(len(daily_df))
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6.2), sharex=True)
    ax1.bar(x, daily_df["production_units"], color="#2b5c8f")
    ax1.set_title("Daily Bottle Output Across 56 Operating Dates (2022-07 to 2023-02)")
    ax1.set_ylabel("Units Produced")
    ax2.bar(x, daily_df["operating_hours"], color="#2e8b57", label="Operating Hours")
    ax2.bar(x, daily_df["pause_hours"], bottom=daily_df["operating_hours"], color="#d95f02", label="Pause Hours")
    ax2.set_title("Daily Operating vs. Pause Hours (Non-Overlapping Multi-Run Dates Summed)")
    ax2.set_ylabel("Hours")
    ax2.legend(loc="upper left")
    step = 4
    ax2.set_xticks(list(x)[::step])
    ax2.set_xticklabels(daily_df["date_str"].iloc[::step], rotation=45, ha="right", fontsize=8)
    fig.tight_layout()
    fig.savefig(charts_dir / "daily_production_and_downtime.png", dpi=150)
    plt.close(fig)

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10.5, 4.2))
    m_x = list(range(len(monthly_df)))
    w = 0.36
    ax1.bar([i - w / 2 for i in m_x], monthly_df["operating_hours"], width=w, color="#2e8b57", label="Operating h")
    ax1.bar([i + w / 2 for i in m_x], monthly_df["pause_hours"], width=w, color="#d95f02", label="Pause h")
    ax1.set_title("Monthly Operating vs. Pause Hours")
    ax1.set_ylabel("Hours")
    ax1.set_xticks(m_x)
    ax1.set_xticklabels(monthly_df["year_month"], rotation=30, ha="right", fontsize=8)
    ax1.legend()
    ax2.bar([i - w / 2 for i in m_x], monthly_df["weighted_eff_pct"], width=w, color="#2b5c8f", label="Efficiency %")
    ax2.bar([i + w / 2 for i in m_x], monthly_df["weighted_downtime_pct"], width=w, color="#c0392b", label="Downtime %")
    ax2.set_title("Monthly Time-Weighted Efficiency vs. Downtime (%)")
    ax2.set_ylabel("Percentage (%)")
    ax2.set_ylim(0, 80)
    ax2.set_xticks(m_x)
    ax2.set_xticklabels(monthly_df["year_month"], rotation=30, ha="right", fontsize=8)
    ax2.legend()
    fig.tight_layout()
    fig.savefig(charts_dir / "monthly_efficiency_and_downtime.png", dpi=150)
    plt.close(fig)

def write_df_to_sheet(ws, df, start_row=1, title=None):
    """Write a DataFrame block to an openpyxl worksheet with simple formatting."""
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(bold=True, color="FFFFFF")
    if title:
        ws.cell(start_row, 1, title).font = Font(bold=True, size=11)
        start_row += 1
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(start_row, c_idx, str(col_name))
        cell.fill, cell.font = hdr_fill, hdr_font
    for r_idx, row in enumerate(df.itertuples(index=False), start_row + 1):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(r_idx, c_idx, round(val, 4) if isinstance(val, float) else val)
            if isinstance(val, float):
                cell.number_format = "0.00"
    return start_row + len(df) + 1

def save_excel(run_df, daily_df, kpi_df, product_df, monthly_df, top_dt_df, xlsx_out):
    """Save clean 3-sheet Excel workbook with live formulas on Selected Data."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Selected Data"
    cols = ["date_str", "product_type_l", "production_units", "liters_produced",
            "production_start_time", "production_end_time", "monitored_time_dec",
            "operation_time_dec", "pause_time_dec", "source_efficiency_pct",
            "gallons_per_hour", "calc_efficiency_pct", "downtime_pct", "calc_units_per_op_hour"]
    sel = run_df[cols].copy()
    end_r = write_df_to_sheet(ws1, sel, start_row=1)
    ws1.cell(end_r, 1, "TOTAL / OVERALL").font = Font(bold=True)
    for col_let in ["C", "D", "G", "H", "I"]:
        c = ws1.cell(end_r, ord(col_let) - 64, f"=SUM({col_let}2:{col_let}{end_r - 1})")
        c.font = Font(bold=True)
        c.number_format = "#,##0.00" if col_let in ["G", "H", "I"] else "#,##0"
    ws1.cell(end_r, 12, f"=H{end_r}/G{end_r}*100").number_format = "0.00"
    ws1.cell(end_r, 13, f"=I{end_r}/G{end_r}*100").number_format = "0.00"
    ws1.cell(end_r, 14, f"=C{end_r}/H{end_r}").number_format = "0.00"

    ws2 = wb.create_sheet("KPI Summary")
    r = write_df_to_sheet(ws2, kpi_df, start_row=1, title="Overall Production & Downtime KPIs (59 Runs)")
    r = write_df_to_sheet(ws2, product_df, start_row=r + 1, title="Comparison by Bottle Size (3L vs 5L)")
    write_df_to_sheet(ws2, top_dt_df, start_row=r + 1, title="Top 5 Runs by Pause Hours")

    ws3 = wb.create_sheet("Period Comparison")
    r = write_df_to_sheet(ws3, monthly_df, start_row=1, title="Monthly Production & Downtime Breakdown")
    write_df_to_sheet(ws3, daily_df, start_row=r + 1, title="Calendar-Date Totals (56 Unique Operating Dates)")

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 38)
    wb.save(xlsx_out)

def build_report_text(run_df, kpi_df, product_df, monthly_df):
    """Derive all PDF summary text and time-balance discrepancy metrics from calculated DataFrames."""
    kpi = dict(zip(kpi_df["Metric"], kpi_df["Value"]))
    runs = int(kpi["Production Runs Recorded"])
    unique_dates = int(kpi["Unique Operating Dates"])
    stop_events = int(kpi["Logged Stop Events"])
    total_units = int(kpi["Total Production Units"])
    total_liters = int(kpi["Total Volume Produced"])
    mon_h = float(kpi["Total Monitored Time"])
    op_h = float(kpi["Total Operating Time"])
    pause_h = float(kpi["Total Pause / Downtime"])
    w_eff = float(kpi["Time-Weighted Efficiency"])
    w_dt = float(kpi["Time-Weighted Downtime"])

    multi_run_dates = int((run_df.groupby("date_str").size() > 1).sum())
    min_date, max_date = run_df["date_str"].min(), run_df["date_str"].max()
    min_month, max_month = monthly_df["year_month"].min(), monthly_df["year_month"].max()

    best_prod = product_df.loc[product_df["weighted_eff_pct"].idxmax()]
    other_prods = product_df[product_df["product_type_l"] != best_prod["product_type_l"]]
    if not other_prods.empty:
        other_row = other_prods.iloc[0]
        prod_cmp = (
            f"{int(best_prod['product_type_l'])}L runs achieved higher weighted efficiency "
            f"({best_prod['weighted_eff_pct']:.2f}%) than {int(other_row['product_type_l'])}L runs "
            f"({other_row['weighted_eff_pct']:.2f}%)."
        )
    else:
        prod_cmp = (
            f"{int(best_prod['product_type_l'])}L runs achieved "
            f"{best_prod['weighted_eff_pct']:.2f}% weighted efficiency."
        )

    peak_out = monthly_df.loc[monthly_df["production_units"].idxmax()]
    peak_pause = monthly_df.loc[monthly_df["pause_hours"].idxmax()]
    best_eff_m = monthly_df.loc[monthly_df["weighted_eff_pct"].idxmax()]
    if peak_out["year_month"] == peak_pause["year_month"]:
        month_obs = (
            f"{peak_out['year_month']} had both the highest output ({int(peak_out['production_units']):,} units) "
            f"and pause time ({peak_pause['pause_hours']:.2f} h; {peak_pause['weighted_downtime_pct']:.2f}% downtime), "
            f"while {best_eff_m['year_month']} had the highest weighted efficiency ({best_eff_m['weighted_eff_pct']:.2f}%)."
        )
    else:
        month_obs = (
            f"{peak_out['year_month']} had the highest output ({int(peak_out['production_units']):,} units) and "
            f"{peak_pause['year_month']} had the highest pause time ({peak_pause['pause_hours']:.2f} h; "
            f"{peak_pause['weighted_downtime_pct']:.2f}% downtime), while {best_eff_m['year_month']} had the highest "
            f"weighted efficiency ({best_eff_m['weighted_eff_pct']:.2f}%)."
        )

    diff = run_df["operation_time_dec"] + run_df["pause_time_dec"] - run_df["monitored_time_dec"]
    disc_runs = int((diff.abs() > 0.01).sum())
    net_diff_h = float(diff.sum())
    max_diff_h = float(diff.abs().max()) if len(run_df) else 0.0
    max_diff_min = max_diff_h * 60.0

    subtitle_dates = f"Beverage Bottling Line Operations ({min_date} to {max_date}) | Zenodo DOI: 10.5281/zenodo.18146866 (CC BY 4.0)"
    kpi_title = f"Overall KPI Summary ({runs} Runs / {unique_dates} Dates)"
    month_title = f"Monthly Summary ({min_month} to {max_month})"
    prov = (
        textwrap.fill(
            f"1. Provenance & Scope: Real historical records from daily_operation_summary ({runs} production runs across "
            f"{unique_dates} dates; {stop_events:,} supplementary stop records in downtime_event_log). On {multi_run_dates} dates "
            "with multiple runs (3L and 5L), start/end timestamps do not overlap.",
            92,
        )
        + "\n"
        + textwrap.fill(
            f"2. Key Findings: Total output was {total_units:,} units ({total_liters:,} L) across {mon_h:.2f} monitored hours "
            f"({op_h:.2f} operating h, {pause_h:.2f} pause h). Time-weighted efficiency was {w_eff:.2f}% and downtime was "
            f"{w_dt:.2f}%. {prod_cmp}",
            92,
        )
    )
    notes = textwrap.fill(
        f"Observations & Limitations: (1) {month_obs} "
        f"(2) Source time-balance discrepancy: {disc_runs} of {runs} runs have |operation + pause - monitored| > 0.01 h "
        f"({net_diff_h:+.2f} h cumulative net diff; max single-run diff {max_diff_h:.3f} h / {max_diff_min:.1f} min). "
        f"Source data is unaltered and discrepancies are not mere rounding errors; efficiency ({w_eff:.2f}%) and "
        f"downtime ({w_dt:.2f}%) are computed independently from monitored time, so their sum ({w_eff + w_dt:.2f}%) differs from 100%. "
        "(3) Source gallons_per_hour uses units/operating_time in rows 1-18 and units/monitored_time in most later rows, "
        "so consistent operating-hour rates are reported separately. (4) No fault codes exist; pauses are not assumed breakdowns.",
        102,
    )
    return subtitle_dates, prov, notes, kpi_title, month_title

def save_pdf_report(run_df, kpi_df, product_df, monthly_df, charts_dir, pdf_path):
    """Build a concise 2-page PDF report matching saved CSV and Excel values."""
    subtitle_dates, prov, notes, kpi_title, month_title = build_report_text(run_df, kpi_df, product_df, monthly_df)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    with PdfPages(pdf_path) as pdf:
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.text(0.07, 0.94, "Industrial Production & Downtime Analysis", fontsize=17, fontweight="bold", color="#1f4e78")
        fig.text(0.07, 0.915, subtitle_dates, fontsize=8.5, color="#444444")
        fig.text(0.07, 0.88, prov, fontsize=8.8, va="top", bbox=dict(boxstyle="round,pad=0.5", fc="#f4f7fa", ec="#c5d3e0"))
        ax1 = fig.add_axes([0.07, 0.35, 0.86, 0.39])
        ax1.axis("off")
        ax1.set_title(kpi_title, loc="left", fontsize=10.5, fontweight="bold")
        t1 = ax1.table(cellText=kpi_df.values, colLabels=kpi_df.columns, loc="center", cellLoc="left", colWidths=[0.28, 0.13, 0.13, 0.46])
        t1.auto_set_font_size(False)
        t1.set_fontsize(7.8)
        t1.scale(1, 1.32)
        ax2 = fig.add_axes([0.07, 0.13, 0.86, 0.17])
        ax2.axis("off")
        ax2.set_title("Product Type Comparison (3L vs. 5L Bottles)", loc="left", fontsize=10.5, fontweight="bold")
        p_cols = ["product_type_l", "runs", "production_units", "liters_produced", "monitored_hours", "operating_hours", "pause_hours", "weighted_eff_pct", "weighted_downtime_pct"]
        t2 = ax2.table(cellText=product_df[p_cols].values, colLabels=["Bottle (L)", "Runs", "Units", "Liters", "Mon h", "Op h", "Pause h", "W. Eff %", "W. DT %"], loc="center", cellLoc="center")
        t2.auto_set_font_size(False)
        t2.set_fontsize(8)
        t2.scale(1, 1.4)
        pdf.savefig(fig)
        plt.close(fig)

        fig = plt.figure(figsize=(8.27, 11.69))
        fig.text(0.07, 0.945, "Period Trends, Charts & Limitations", fontsize=14, fontweight="bold", color="#1f4e78")
        ax_m = fig.add_axes([0.07, 0.68, 0.86, 0.23])
        ax_m.axis("off")
        ax_m.set_title(month_title, loc="left", fontsize=10, fontweight="bold")
        m_cols = ["year_month", "runs", "unique_dates", "production_units", "monitored_hours", "operating_hours", "pause_hours", "weighted_eff_pct", "weighted_downtime_pct"]
        tm = ax_m.table(cellText=monthly_df[m_cols].values, colLabels=["Month", "Runs", "Days", "Units", "Mon h", "Op h", "Pause h", "W. Eff %", "W. DT %"], loc="center", cellLoc="center")
        tm.auto_set_font_size(False)
        tm.set_fontsize(8)
        tm.scale(1, 1.25)
        for ax_box, img_name in [([0.06, 0.415, 0.88, 0.245], "daily_production_and_downtime.png"), ([0.06, 0.175, 0.88, 0.22], "monthly_efficiency_and_downtime.png")]:
            ax = fig.add_axes(ax_box)
            ax.imshow(mpimg.imread(charts_dir / img_name))
            ax.axis("off")
        fig.text(0.07, 0.155, notes, fontsize=7.9, va="top", bbox=dict(boxstyle="round,pad=0.45", fc="#fff9f0", ec="#e0c9a6"))
        pdf.savefig(fig)
        plt.close(fig)

def run_pipeline(base_dir):
    """Run full data load, metric calculations, CSV/Excel/chart exports, and PDF report."""
    base = Path(base_dir)
    raw_df, ev_count, ev_days = load_and_check_data(base / "data" / "production_raw.xlsx")
    run_df = calculate_run_metrics(raw_df)
    daily_df, kpi_df, product_df, monthly_df = build_summaries(run_df, ev_count, ev_days)
    top_dt_df = run_df.nlargest(5, "pause_time_dec")[["date_str", "product_type_l", "production_units", "monitored_time_dec", "operation_time_dec", "pause_time_dec", "calc_efficiency_pct", "downtime_pct"]]
    out_dir = base / "outputs"
    out_dir.mkdir(parents=True, exist_ok=True)
    run_df.to_csv(out_dir / "daily_summary.csv", index=False)
    save_charts(daily_df, monthly_df, out_dir / "charts")
    save_excel(run_df, daily_df, kpi_df, product_df, monthly_df, top_dt_df, out_dir / "production_summary.xlsx")
    save_pdf_report(run_df, kpi_df, product_df, monthly_df, out_dir / "charts", base / "reports" / "production_report.pdf")
    return run_df, daily_df, kpi_df, product_df, monthly_df

if __name__ == "__main__":
    project_root = Path(__file__).resolve().parent.parent
    _, _, kpi, prod, month = run_pipeline(project_root)
    print(kpi.to_string(index=False))
    print("\n--- By Product ---\n", prod.to_string(index=False))
    print("\n--- By Month ---\n", month.to_string(index=False))
