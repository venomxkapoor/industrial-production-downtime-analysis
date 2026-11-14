"""
Build the two-page PDF report from the saved analysis outputs.

All numbers (and the words "higher"/"lower") are read from the CSV files
in outputs/, so the report always agrees with the latest analysis run.
Only matplotlib is used - no extra PDF library is needed.

Run from the project folder (after analyse_schedules.py):
    python src/build_report.py
"""

import textwrap
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
OUTPUT_FOLDER = PROJECT_FOLDER / "outputs"
CHART_FOLDER = OUTPUT_FOLDER / "charts"
REPORT_FILE = PROJECT_FOLDER / "reports" / "production_planning_report.pdf"

A4_SIZE_INCHES = (8.27, 11.69)
LEFT_MARGIN = 0.07
TEXT_WIDTH_CHARS = 112
LINE_HEIGHT = 0.0145
HEADING_COLOUR = "#1f4e79"

COUNT_METRICS = ["Number of orders", "Late orders"]


# ---------------------------------------------------------------------------
# Reading the results
# ---------------------------------------------------------------------------

def load_results():
    """Read the saved outputs into simple Python dictionaries."""
    comparison = pd.read_csv(OUTPUT_FOLDER / "kpi_comparison.csv")
    workload = pd.read_csv(OUTPUT_FOLDER / "work_centre_load.csv")
    kpis_wc = pd.read_csv(OUTPUT_FOLDER / "kpi_by_work_centre.csv")

    fifo = {}
    edd = {}
    for index, row in comparison.iterrows():
        fifo[row["metric"]] = row["FIFO"]
        edd[row["metric"]] = row["EDD"]

    return {"comparison": comparison, "workload": workload, "kpis_wc": kpis_wc,
            "fifo": fifo, "edd": edd}


def format_value(metric, value):
    """Counts without decimals, everything else with one or two decimals."""
    if metric in COUNT_METRICS:
        return f"{int(value)}"
    if metric == "Late orders (%)":
        return f"{value:.1f}"
    return f"{value:.2f}".rstrip("0").rstrip(".")


def change_word(fifo_value, edd_value):
    """Describe how EDD compares with FIFO for a metric."""
    if edd_value < fifo_value:
        return "lower under EDD"
    if edd_value > fifo_value:
        return "higher under EDD"
    return "identical for both rules"


def better_policy(fifo_value, edd_value):
    """For metrics where smaller is better: which policy is better?"""
    if edd_value < fifo_value:
        return "EDD"
    if fifo_value < edd_value:
        return "FIFO"
    return "neither"


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def write_heading(fig, y, text):
    fig.text(LEFT_MARGIN, y, text, fontsize=11, fontweight="bold", color=HEADING_COLOUR, va="top")
    return y - 0.022


def write_paragraph(fig, y, text, size=8.5):
    for line in textwrap.wrap(text, TEXT_WIDTH_CHARS):
        fig.text(LEFT_MARGIN, y, line, fontsize=size, va="top")
        y = y - LINE_HEIGHT
    return y - 0.006


def write_bullets(fig, y, items, size=8.5):
    for item in items:
        lines = textwrap.wrap(item, TEXT_WIDTH_CHARS - 4)
        for line_number in range(len(lines)):
            prefix = "\u2022  " if line_number == 0 else "    "
            fig.text(LEFT_MARGIN + 0.01, y, prefix + lines[line_number], fontsize=size, va="top")
            y = y - LINE_HEIGHT
    return y - 0.006


def draw_table(fig, rect, header, rows, column_widths):
    """Draw a simple table inside the rectangle [left, bottom, width, height]."""
    ax = fig.add_axes(rect)
    ax.axis("off")
    table = ax.table(cellText=rows, colLabels=header, colWidths=column_widths,
                     loc="upper center", cellLoc="center")
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.25)
    for (row_number, column_number), cell in table.get_celld().items():
        cell.set_edgecolor("#bfbfbf")
        if row_number == 0:
            cell.set_text_props(fontweight="bold", color="white")
            cell.set_facecolor(HEADING_COLOUR)
        elif column_number == 0:
            cell.set_text_props(ha="left")
    return ax


def add_image(fig, rect, image_file):
    ax = fig.add_axes(rect)
    ax.imshow(plt.imread(image_file))
    ax.axis("off")


# ---------------------------------------------------------------------------
# Findings written from the numbers
# ---------------------------------------------------------------------------

def make_findings(results):
    fifo = results["fifo"]
    edd = results["edd"]
    workload = results["workload"]
    kpis_wc = results["kpis_wc"]
    n = int(fifo["Number of orders"])

    findings = []

    findings.append(
        f"Late orders: {int(fifo['Late orders'])} of {n} ({fifo['Late orders (%)']:.1f} %) under FIFO and "
        f"{int(edd['Late orders'])} ({edd['Late orders (%)']:.1f} %) under EDD - "
        f"{change_word(fifo['Late orders'], edd['Late orders'])}. EDD does not guarantee fewer late orders."
    )
    tardiness_text = (
        f"Average tardiness: {fifo['Average tardiness (h)']:.2f} h (FIFO) vs {edd['Average tardiness (h)']:.2f} h (EDD) - "
        f"{change_word(fifo['Average tardiness (h)'], edd['Average tardiness (h)'])}. Maximum lateness: "
        f"{fifo['Maximum lateness (h)']:g} h vs {edd['Maximum lateness (h)']:g} h - "
        f"{change_word(fifo['Maximum lateness (h)'], edd['Maximum lateness (h)'])}."
    )
    if edd["Maximum lateness (h)"] < fifo["Maximum lateness (h)"] and edd["Late orders"] >= fifo["Late orders"]:
        tardiness_text = tardiness_text + (" Under EDD, at least as many orders are late, but the delays are "
                                           "spread more evenly and the worst delay is smaller.")
    findings.append(tardiness_text)
    findings.append(
        f"Average waiting time: {fifo['Average waiting time (h)']:.2f} h (FIFO) vs {edd['Average waiting time (h)']:.2f} h "
        f"(EDD) - {change_word(fifo['Average waiting time (h)'], edd['Average waiting time (h)'])}. "
        f"Average flow time (completion - release): {fifo['Average flow time (h)']:.2f} h vs {edd['Average flow time (h)']:.2f} h."
    )
    findings.append(
        f"Total processing time ({fifo['Total processing time (h)']:g} h) and makespan ({fifo['Makespan (h)']:g} h) are "
        f"{change_word(fifo['Makespan (h)'], edd['Makespan (h)'])}, as expected: re-ordering jobs without setup or idle "
        f"time does not change the amount of work or when the last job finishes."
    )

    load_parts = []
    for index, row in workload.iterrows():
        text = (f"{row['work_centre']} has {row['assigned_hours']:g} h assigned against a {row['capacity_budget_h']:g} h "
                f"budget (ratio {row['capacity_pressure_ratio']:.2f}")
        if row["overload_h"] > 0:
            text = text + f", {row['overload_h']:g} h over budget)"
        else:
            text = text + ", within budget)"
        load_parts.append(text)
    late_parts = []
    for policy in ["FIFO", "EDD"]:
        policy_rows = kpis_wc[kpis_wc["policy"] == policy]
        part = policy + ": "
        pieces = []
        for index, row in policy_rows.iterrows():
            pieces.append(f"{row['work_centre']} {int(row['Late orders'])}")
        late_parts.append(part + ", ".join(pieces))
    findings.append(
        "; ".join(load_parts) + ". Late orders by work centre - " + "; ".join(late_parts) +
        ". No sequencing rule can remove an overload; it needs capacity or due-date changes."
    )

    count_winner = better_policy(fifo["Late orders"], edd["Late orders"])
    max_winner = better_policy(fifo["Maximum lateness (h)"], edd["Maximum lateness (h)"])
    tardiness_winner = better_policy(fifo["Average tardiness (h)"], edd["Average tardiness (h)"])
    if count_winner != max_winner:
        findings.append(
            f"Trade-off: if the main goal is the number of late orders, {count_winner} performed better in this "
            f"dataset; if the goal is to limit how late orders become (maximum lateness, average tardiness), "
            f"{max_winner} performed better. The right rule depends on the business objective."
        )
    else:
        findings.append(
            f"In this dataset {max_winner} was better for both the number of late orders and maximum lateness; "
            f"average tardiness was better under {tardiness_winner}. This may not hold for other order sets."
        )
    return findings


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

def make_page_one(results):
    fig = plt.figure(figsize=A4_SIZE_INCHES)
    fifo = results["fifo"]

    fig.text(LEFT_MARGIN, 0.965, "Production Scheduling & Workload Analysis", fontsize=17,
             fontweight="bold", color=HEADING_COLOUR, va="top")
    fig.text(LEFT_MARGIN, 0.935, "FIFO vs Earliest Due Date (EDD) on two work centres  |  Python and Excel",
             fontsize=10, va="top")
    fig.text(LEFT_MARGIN, 0.915,
             "Personal educational portfolio project. All orders are SIMULATED (fixed random seed) - not real factory data.",
             fontsize=8.5, style="italic", color="#c0392b", va="top")

    y = 0.885
    y = write_heading(fig, y, "1. Business question")
    y = write_paragraph(fig, y,
        f"A fictional manufacturer has {int(fifo['Number of orders'])} open production orders, each with a processing time "
        "and a delivery due time. Management wants to know how much work is assigned, which work centre is under more "
        "capacity pressure, how many orders finish late, how long orders wait, and how two simple sequencing rules "
        "compare - including the trade-offs between them.")

    y = write_heading(fig, y, "2. Data and assumptions")
    y = write_bullets(fig, y, [
        "80 simulated orders (seed 42): order ID, original sequence, work centre (WC1 / WC2, about 60 / 40 % split), "
        "product category, processing time 1-8 h (0.5 h steps), due time 8-160 h after the start of the planning period.",
        "All orders available at time 0. Each order runs entirely on its assigned work centre; the two work centres "
        "work in parallel and independently; one order at a time per work centre; no interruption (no preemption).",
        "Deterministic processing times; no setup times, breakdowns, operator shortages, shift breaks or transfers. "
        "Time is measured in continuous processing hours - a simplification, not a claim about real shift patterns.",
        "Capacity budget: 160 h per work centre (illustrative: 10 days x 2 shifts x 8 h). Capacity-pressure ratio = "
        "assigned hours / budget. This is a planning load ratio, not measured machine utilisation.",
    ])

    y = write_heading(fig, y, "3. Scheduling rules and KPIs")
    y = write_bullets(fig, y, [
        "FIFO (First In, First Out): orders processed in their original sequence within each work centre.",
        "EDD (Earliest Due Date): orders processed by earliest due time within each work centre; "
        "ties broken by original sequence.",
        "Per order: start, completion = start + processing time, waiting = start - release, flow time = completion - "
        "release, lateness = completion - due (negative = early), tardiness = max(0, lateness), late if tardiness > 0.",
        "Per rule: late orders (count, %), average waiting, flow time and tardiness, maximum lateness, makespan "
        "(latest completion time) and total processing time.",
    ])

    y = write_heading(fig, y, "4. KPI comparison (all 80 orders)")
    header = ["Metric", "FIFO", "EDD", "EDD - FIFO"]
    rows = []
    for index, row in results["comparison"].iterrows():
        metric = row["metric"]
        rows.append([metric, format_value(metric, row["FIFO"]), format_value(metric, row["EDD"]),
                     format_value(metric, row["difference_EDD_minus_FIFO"])])
    table_height = 0.205
    draw_table(fig, [LEFT_MARGIN, y - table_height, 0.86, table_height], header, rows, [0.46, 0.18, 0.18, 0.18])
    y = y - table_height - 0.012

    y = write_heading(fig, y, "5. Work-centre load")
    header = ["Work centre", "Orders", "Assigned (h)", "Budget (h)", "Pressure ratio", "Overload (h)"]
    rows = []
    for index, row in results["workload"].iterrows():
        rows.append([row["work_centre"], f"{int(row['number_of_orders'])}", f"{row['assigned_hours']:g}",
                     f"{row['capacity_budget_h']:g}", f"{row['capacity_pressure_ratio']:.2f}",
                     f"{row['overload_h']:g}"])
    draw_table(fig, [LEFT_MARGIN, y - 0.07, 0.86, 0.07], header, rows, [0.2, 0.12, 0.17, 0.15, 0.18, 0.18])

    fig.text(0.5, 0.02, "Page 1 of 2", fontsize=7.5, ha="center", color="grey")
    return fig


def make_page_two(results):
    fig = plt.figure(figsize=A4_SIZE_INCHES)

    fig.text(LEFT_MARGIN, 0.965, "6. Results in charts", fontsize=11, fontweight="bold",
             color=HEADING_COLOUR, va="top")
    add_image(fig, [0.03, 0.72, 0.47, 0.235], CHART_FOLDER / "late_orders_by_policy.png")
    add_image(fig, [0.50, 0.72, 0.47, 0.235], CHART_FOLDER / "time_kpis_by_policy.png")
    add_image(fig, [0.265, 0.485, 0.47, 0.235], CHART_FOLDER / "workload_vs_capacity.png")

    y = 0.475
    y = write_heading(fig, y, "7. Principal conclusions (supported by the simulated model)")
    y = write_bullets(fig, y, make_findings(results))

    y = write_heading(fig, y, "8. Practical limitations")
    y = write_bullets(fig, y, [
        "Simulated data: due times and processing times are random, so results show how the rules behave, "
        "not how a real plant performs. Real order data would be needed for business decisions.",
        "Single-stage model: each order uses one work centre only. A real multi-stage system can have bottlenecks "
        "that this model cannot detect; WC1 is only a capacity-pressure point under these assumptions.",
        "No setups, breakdowns, shift calendars, new order arrivals or priorities/penalties per customer. "
        "Any of these could change which rule performs better.",
        "Only two simple rules were compared; no optimisation was attempted. One dataset was analysed, so the "
        "differences are not tested statistically across many scenarios.",
    ])

    fig.text(0.5, 0.02, "Page 2 of 2  |  Source code, data and Excel workbook: production-scheduling-analysis repository",
             fontsize=7.5, ha="center", color="grey")
    return fig


def main():
    results = load_results()
    REPORT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with PdfPages(REPORT_FILE) as pdf:
        for fig in [make_page_one(results), make_page_two(results)]:
            pdf.savefig(fig)
            plt.close(fig)
    print(f"Saved report to {REPORT_FILE}")


if __name__ == "__main__":
    main()
