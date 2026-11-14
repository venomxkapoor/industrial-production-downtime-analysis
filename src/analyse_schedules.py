"""
Schedule simulated production orders with two simple rules,
calculate performance indicators (KPIs) and save the results.

Policy A - FIFO: orders are processed in their original sequence.
Policy B - EDD:  orders are processed by earliest due time
                 (ties are broken by the original sequence).

Each work centre processes its own orders one after another.
The two work centres work in parallel and independently.

Run from the project folder:
    python src/analyse_schedules.py
"""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # save charts to files without opening a window
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl.styles import Font, PatternFill

PROJECT_FOLDER = Path(__file__).resolve().parent.parent
INPUT_FILE = PROJECT_FOLDER / "data" / "production_orders.csv"
OUTPUT_FOLDER = PROJECT_FOLDER / "outputs"
CHART_FOLDER = OUTPUT_FOLDER / "charts"

# Illustrative capacity budget per work centre for the planning period:
# 10 working days x 2 shifts x 8 hours = 160 hours.
CAPACITY_BUDGET_H = 160

REQUIRED_COLUMNS = [
    "order_id",
    "original_sequence",
    "work_centre",
    "product_category",
    "processing_time_h",
    "due_time_h",
    "release_time_h",
]

POLICY_COLOURS = {"FIFO": "#8c96a3", "EDD": "#2e75b6"}

# Small example with 5 orders on one work centre. The expected results
# were calculated by hand (see the README and the Excel "Worked Example").
WORKED_EXAMPLE_ORDERS = [
    {"order_id": "EX-1", "original_sequence": 1, "processing_time_h": 3.0, "due_time_h": 5},
    {"order_id": "EX-2", "original_sequence": 2, "processing_time_h": 2.0, "due_time_h": 4},
    {"order_id": "EX-3", "original_sequence": 3, "processing_time_h": 4.0, "due_time_h": 12},
    {"order_id": "EX-4", "original_sequence": 4, "processing_time_h": 1.0, "due_time_h": 3},
    {"order_id": "EX-5", "original_sequence": 5, "processing_time_h": 5.0, "due_time_h": 15},
]


# ---------------------------------------------------------------------------
# Stage A - load and check the data
# ---------------------------------------------------------------------------

def load_orders(file_path):
    """Read the order CSV file into a DataFrame."""
    return pd.read_csv(file_path)


def check_orders(orders):
    """Return a list of data problems. An empty list means the data is OK."""
    problems = []

    for column in REQUIRED_COLUMNS:
        if column not in orders.columns:
            problems.append(f"Missing column: {column}")
    if len(problems) > 0:
        return problems

    if orders[REQUIRED_COLUMNS].isna().any().any():
        problems.append("Some required values are missing (empty cells).")
    if orders["order_id"].duplicated().any():
        problems.append("Some order IDs appear more than once.")
    if orders["original_sequence"].duplicated().any():
        problems.append("Some original sequence numbers appear more than once.")
    if (orders["processing_time_h"] <= 0).any():
        problems.append("Some processing times are zero or negative.")
    if (orders["due_time_h"] < 0).any():
        problems.append("Some due times are negative.")
    if (orders["release_time_h"] < 0).any():
        problems.append("Some release times are negative.")

    return problems


def make_worked_example_orders():
    """Return the 5-order worked example as a DataFrame (one work centre)."""
    example = pd.DataFrame(WORKED_EXAMPLE_ORDERS)
    example["work_centre"] = "WC1"
    example["product_category"] = "Example"
    example["release_time_h"] = 0
    return example


# ---------------------------------------------------------------------------
# Stages B and C - put the orders in sequence and build the schedule
# ---------------------------------------------------------------------------

def sort_fifo(orders):
    """FIFO: within each work centre, keep the original order sequence."""
    sorted_orders = orders.sort_values(["work_centre", "original_sequence"])
    return sorted_orders.reset_index(drop=True)


def sort_edd(orders):
    """EDD: within each work centre, earliest due time first.
    Ties on due time are broken by the original sequence (smaller first)."""
    sorted_orders = orders.sort_values(["work_centre", "due_time_h", "original_sequence"])
    return sorted_orders.reset_index(drop=True)


def build_schedule(sorted_orders, policy_name):
    """Go through the orders in the given sequence and calculate the
    start time, completion time and delivery measures of every order."""
    machine_free_at = {}   # work centre -> time (h) when it becomes free
    position_counter = {}  # work centre -> how many orders were scheduled
    rows = []

    for index, order in sorted_orders.iterrows():
        work_centre = order["work_centre"]

        # Every work centre starts empty at time 0.
        if work_centre not in machine_free_at:
            machine_free_at[work_centre] = 0.0
            position_counter[work_centre] = 0

        # An order starts when the machine is free (and the order is released).
        start_time = max(machine_free_at[work_centre], order["release_time_h"])
        completion_time = start_time + order["processing_time_h"]
        machine_free_at[work_centre] = completion_time
        position_counter[work_centre] = position_counter[work_centre] + 1

        waiting_time = start_time - order["release_time_h"]
        flow_time = completion_time - order["release_time_h"]
        lateness = completion_time - order["due_time_h"]
        tardiness = max(0.0, lateness)
        is_late = tardiness > 0

        rows.append({
            "policy": policy_name,
            "work_centre": work_centre,
            "position_on_work_centre": position_counter[work_centre],
            "order_id": order["order_id"],
            "original_sequence": order["original_sequence"],
            "product_category": order["product_category"],
            "processing_time_h": order["processing_time_h"],
            "release_time_h": order["release_time_h"],
            "due_time_h": order["due_time_h"],
            "start_time_h": start_time,
            "completion_time_h": completion_time,
            "waiting_time_h": waiting_time,
            "flow_time_h": flow_time,
            "lateness_h": lateness,
            "tardiness_h": tardiness,
            "is_late": is_late,
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Stage D - KPIs, comparison and workload
# ---------------------------------------------------------------------------

def calculate_kpis(schedule):
    """Summarise one schedule with simple performance indicators."""
    number_of_orders = len(schedule)
    late_orders = int(schedule["is_late"].sum())

    kpis = {
        "Number of orders": number_of_orders,
        "Late orders": late_orders,
        "Late orders (%)": round(100 * late_orders / number_of_orders, 1),
        "Average waiting time (h)": round(schedule["waiting_time_h"].mean(), 2),
        "Average flow time (h)": round(schedule["flow_time_h"].mean(), 2),
        "Average tardiness (h)": round(schedule["tardiness_h"].mean(), 2),
        "Total tardiness (h)": round(schedule["tardiness_h"].sum(), 2),
        "Maximum lateness (h)": round(schedule["lateness_h"].max(), 2),
        "Total processing time (h)": round(schedule["processing_time_h"].sum(), 2),
        "Makespan (h)": round(schedule["completion_time_h"].max(), 2),
    }
    return kpis


def compare_policies(fifo_kpis, edd_kpis):
    """Put the KPIs of both policies side by side in one table."""
    rows = []
    for metric in fifo_kpis:
        fifo_value = fifo_kpis[metric]
        edd_value = edd_kpis[metric]
        rows.append({
            "metric": metric,
            "FIFO": fifo_value,
            "EDD": edd_value,
            "difference_EDD_minus_FIFO": round(edd_value - fifo_value, 2),
        })
    return pd.DataFrame(rows)


def kpis_by_work_centre(schedule):
    """Calculate the same KPIs separately for every work centre."""
    rows = []
    for work_centre in sorted(schedule["work_centre"].unique()):
        subset = schedule[schedule["work_centre"] == work_centre]
        row = {"policy": schedule["policy"].iloc[0], "work_centre": work_centre}
        row.update(calculate_kpis(subset))
        rows.append(row)
    return pd.DataFrame(rows)


def calculate_workload(orders, capacity_budget_h):
    """Assigned processing hours per work centre compared with the budget.
    The ratio is a planning load ratio, not measured machine utilisation."""
    workload = orders.groupby("work_centre").agg(
        number_of_orders=("order_id", "count"),
        assigned_hours=("processing_time_h", "sum"),
    ).reset_index()

    workload["capacity_budget_h"] = capacity_budget_h
    workload["capacity_pressure_ratio"] = (workload["assigned_hours"] / capacity_budget_h).round(3)
    workload["overload_h"] = (workload["assigned_hours"] - capacity_budget_h).clip(lower=0)
    workload["is_overloaded"] = workload["assigned_hours"] > capacity_budget_h
    return workload


# ---------------------------------------------------------------------------
# Stage E - charts
# ---------------------------------------------------------------------------

def draw_grouped_bars(ax, labels, fifo_values, edd_values):
    """Draw FIFO and EDD bars next to each other for every label."""
    bar_width = 0.38
    positions = range(len(labels))
    fifo_positions = []
    edd_positions = []
    for position in positions:
        fifo_positions.append(position - bar_width / 2)
        edd_positions.append(position + bar_width / 2)

    fifo_bars = ax.bar(fifo_positions, fifo_values, bar_width, label="FIFO", color=POLICY_COLOURS["FIFO"])
    edd_bars = ax.bar(edd_positions, edd_values, bar_width, label="EDD", color=POLICY_COLOURS["EDD"])
    ax.bar_label(fifo_bars, fmt="%g", fontsize=8)
    ax.bar_label(edd_bars, fmt="%g", fontsize=8)
    ax.set_xticks(list(positions))
    ax.set_xticklabels(labels)
    ax.legend()
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def chart_late_orders(fifo_kpis, edd_kpis, kpis_wc, file_path):
    """Number of late orders per work centre and in total."""
    labels = []
    fifo_values = []
    edd_values = []
    for work_centre in sorted(kpis_wc["work_centre"].unique()):
        fifo_row = kpis_wc[(kpis_wc["policy"] == "FIFO") & (kpis_wc["work_centre"] == work_centre)]
        edd_row = kpis_wc[(kpis_wc["policy"] == "EDD") & (kpis_wc["work_centre"] == work_centre)]
        labels.append(work_centre)
        fifo_values.append(int(fifo_row["Late orders"].iloc[0]))
        edd_values.append(int(edd_row["Late orders"].iloc[0]))
    labels.append("All orders")
    fifo_values.append(fifo_kpis["Late orders"])
    edd_values.append(edd_kpis["Late orders"])

    fig, ax = plt.subplots(figsize=(6, 4))
    draw_grouped_bars(ax, labels, fifo_values, edd_values)
    ax.set_ylabel("Number of late orders")
    ax.set_title("Late orders by policy")
    fig.tight_layout()
    fig.savefig(file_path, dpi=150)
    plt.close(fig)


def chart_time_metrics(fifo_kpis, edd_kpis, file_path):
    """Average waiting time, average tardiness and maximum lateness."""
    metrics = ["Average waiting time (h)", "Average tardiness (h)", "Maximum lateness (h)"]
    labels = ["Avg. waiting", "Avg. tardiness", "Max. lateness"]
    fifo_values = []
    edd_values = []
    for metric in metrics:
        fifo_values.append(fifo_kpis[metric])
        edd_values.append(edd_kpis[metric])

    fig, ax = plt.subplots(figsize=(6, 4))
    draw_grouped_bars(ax, labels, fifo_values, edd_values)
    ax.set_ylabel("Hours")
    ax.set_title("Time-based KPIs by policy (all orders)")
    fig.tight_layout()
    fig.savefig(file_path, dpi=150)
    plt.close(fig)


def chart_workload(workload, file_path):
    """Assigned processing hours per work centre versus the capacity budget."""
    fig, ax = plt.subplots(figsize=(6, 4))
    bars = ax.bar(workload["work_centre"], workload["assigned_hours"], color="#5b9bd5", width=0.5)
    ax.bar_label(bars, fmt="%g h", fontsize=8)
    ax.axhline(CAPACITY_BUDGET_H, color="#c0392b", linestyle="--",
               label=f"Capacity budget ({CAPACITY_BUDGET_H} h)")

    for index, row in workload.iterrows():
        ax.text(index, row["assigned_hours"] / 2, f"ratio {row['capacity_pressure_ratio']:.2f}",
                ha="center", color="white", fontsize=9, fontweight="bold")

    ax.set_ylabel("Assigned processing hours")
    ax.set_title("Workload vs capacity budget per work centre")
    ax.legend(loc="lower right")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(file_path, dpi=150)
    plt.close(fig)


# ---------------------------------------------------------------------------
# Excel workbook
# ---------------------------------------------------------------------------

def make_assumptions_table():
    """Text table with the model assumptions for the Excel workbook."""
    rows = [
        ["Data", "Simulated (fictional) orders created by src/generate_data.py with random seed 42. Not real factory data."],
        ["Time unit", "All times are in hours from the start of the planning period (continuous processing hours)."],
        ["Release", "All orders are available at time 0."],
        ["Machines", "Two independent work centres (WC1, WC2) working in parallel. Each processes one order at a time."],
        ["Assignment", "Each order is processed entirely on its assigned work centre. No transfers, no multi-step routing."],
        ["Interruptions", "No preemption: an order is never interrupted once started."],
        ["Simplifications", "Deterministic processing times. No setup times, breakdowns, operator shortages or shift breaks."],
        ["FIFO", "Orders processed in original sequence within each work centre."],
        ["EDD", "Orders processed by earliest due time within each work centre; ties broken by original sequence."],
        ["Waiting time", "Start time - release time"],
        ["Flow time", "Completion time - release time (equals completion time because release = 0)"],
        ["Lateness", "Completion time - due time (negative = early)"],
        ["Tardiness", "max(0, lateness)"],
        ["Late order", "Tardiness > 0 (an order finishing exactly at its due time is on time)"],
        ["Makespan", "Latest completion time over all orders"],
        ["Capacity budget", f"{CAPACITY_BUDGET_H} h per work centre (illustrative: 10 days x 2 shifts x 8 h)."],
        ["Capacity pressure", "Assigned processing hours / capacity budget. A planning load ratio, not measured utilisation."],
        ["Python vs Excel", "All sheets except 'Worked Example' contain values calculated by Python. 'Worked Example' uses live Excel formulas."],
    ]
    return pd.DataFrame(rows, columns=["Topic", "Description"])


def write_example_block(sheet, first_row, title, example_orders, python_kpis):
    """Write one worked-example table with live Excel formulas.
    Returns the next free row."""
    sheet.cell(row=first_row, column=1, value=title).font = Font(bold=True, size=12)

    headers = ["Order", "Processing time (h)", "Due time (h)", "Release time (h)",
               "Start (h)", "Completion (h)", "Waiting (h)", "Flow time (h)",
               "Lateness (h)", "Tardiness (h)", "Late?"]
    header_row = first_row + 1
    for column_number, header in enumerate(headers, start=1):
        cell = sheet.cell(row=header_row, column=column_number, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")

    first_data_row = header_row + 1
    row = first_data_row
    for index, order in example_orders.iterrows():
        sheet.cell(row=row, column=1, value=order["order_id"])
        sheet.cell(row=row, column=2, value=float(order["processing_time_h"]))
        sheet.cell(row=row, column=3, value=float(order["due_time_h"]))
        sheet.cell(row=row, column=4, value=float(order["release_time_h"]))
        if row == first_data_row:
            sheet.cell(row=row, column=5, value=f"=D{row}")
        else:
            sheet.cell(row=row, column=5, value=f"=MAX(F{row - 1},D{row})")
        sheet.cell(row=row, column=6, value=f"=E{row}+B{row}")
        sheet.cell(row=row, column=7, value=f"=E{row}-D{row}")
        sheet.cell(row=row, column=8, value=f"=F{row}-D{row}")
        sheet.cell(row=row, column=9, value=f"=F{row}-C{row}")
        sheet.cell(row=row, column=10, value=f"=MAX(0,I{row})")
        sheet.cell(row=row, column=11, value=f'=IF(J{row}>0,"Yes","No")')
        row = row + 1
    last_data_row = row - 1

    row = row + 1
    sheet.cell(row=row, column=1, value="Summary").font = Font(bold=True)
    sheet.cell(row=row, column=2, value="Excel formula").font = Font(bold=True)
    sheet.cell(row=row, column=3, value="Python result").font = Font(bold=True)
    row = row + 1

    summary_lines = [
        ["Late orders", f'=COUNTIF(K{first_data_row}:K{last_data_row},"Yes")', "Late orders"],
        ["Average waiting time (h)", f"=AVERAGE(G{first_data_row}:G{last_data_row})", "Average waiting time (h)"],
        ["Average flow time (h)", f"=AVERAGE(H{first_data_row}:H{last_data_row})", "Average flow time (h)"],
        ["Average tardiness (h)", f"=AVERAGE(J{first_data_row}:J{last_data_row})", "Average tardiness (h)"],
        ["Maximum lateness (h)", f"=MAX(I{first_data_row}:I{last_data_row})", "Maximum lateness (h)"],
        ["Makespan (h)", f"=MAX(F{first_data_row}:F{last_data_row})", "Makespan (h)"],
    ]
    for label, formula, kpi_name in summary_lines:
        sheet.cell(row=row, column=1, value=label)
        sheet.cell(row=row, column=2, value=formula)
        sheet.cell(row=row, column=3, value=python_kpis[kpi_name])
        row = row + 1

    return row + 1


def add_worked_example_sheet(workbook):
    """Add a sheet with a 5-order example that uses real Excel formulas."""
    sheet = workbook.create_sheet("Worked Example")
    sheet["A1"] = "Worked example: 5 orders on one work centre (can be checked by hand)"
    sheet["A1"].font = Font(bold=True, size=13)
    sheet["A2"] = ("Columns E-K and the 'Excel formula' column are live Excel formulas. "
                   "The 'Python result' column was calculated by src/analyse_schedules.py "
                   "for the same 5 orders. Both should agree.")

    example = make_worked_example_orders()
    fifo_example = sort_fifo(example)
    edd_example = sort_edd(example)
    fifo_kpis = calculate_kpis(build_schedule(fifo_example, "FIFO"))
    edd_kpis = calculate_kpis(build_schedule(edd_example, "EDD"))

    next_row = write_example_block(sheet, 4, "FIFO (original sequence)", fifo_example, fifo_kpis)
    write_example_block(sheet, next_row, "EDD (earliest due time first)", edd_example, edd_kpis)

    sheet.column_dimensions["A"].width = 26
    for letter in "BCDEFGHIJK":
        sheet.column_dimensions[letter].width = 16


def format_sheet(sheet):
    """Bold header row, freeze it and set readable column widths."""
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
    sheet.freeze_panes = "A2"

    for column_cells in sheet.columns:
        longest = 0
        for cell in column_cells:
            if cell.value is not None:
                longest = max(longest, len(str(cell.value)))
        column_letter = column_cells[0].column_letter
        sheet.column_dimensions[column_letter].width = min(longest + 2, 110)


def write_excel(orders, fifo_schedule, edd_schedule, comparison, kpis_wc, workload, file_path):
    """Save all results into one Excel workbook."""
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        make_assumptions_table().to_excel(writer, sheet_name="Assumptions", index=False)
        orders.to_excel(writer, sheet_name="Original Orders", index=False)
        fifo_schedule.to_excel(writer, sheet_name="FIFO Schedule", index=False)
        edd_schedule.to_excel(writer, sheet_name="EDD Schedule", index=False)
        comparison.to_excel(writer, sheet_name="KPI Comparison", index=False)
        kpis_wc.to_excel(writer, sheet_name="KPI by Work Centre", index=False)
        workload.to_excel(writer, sheet_name="Work Centre Load", index=False)

        for sheet in writer.book.worksheets:
            format_sheet(sheet)
        add_worked_example_sheet(writer.book)


# ---------------------------------------------------------------------------
# Main program
# ---------------------------------------------------------------------------

def main():
    orders = load_orders(INPUT_FILE)

    problems = check_orders(orders)
    if len(problems) > 0:
        for problem in problems:
            print("Data problem:", problem)
        raise SystemExit("Stopping because the input data has problems.")
    print(f"Loaded {len(orders)} orders - data checks passed.")

    fifo_schedule = build_schedule(sort_fifo(orders), "FIFO")
    edd_schedule = build_schedule(sort_edd(orders), "EDD")

    fifo_kpis = calculate_kpis(fifo_schedule)
    edd_kpis = calculate_kpis(edd_schedule)
    comparison = compare_policies(fifo_kpis, edd_kpis)
    kpis_wc = pd.concat([kpis_by_work_centre(fifo_schedule), kpis_by_work_centre(edd_schedule)],
                        ignore_index=True)
    workload = calculate_workload(orders, CAPACITY_BUDGET_H)

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    CHART_FOLDER.mkdir(parents=True, exist_ok=True)

    fifo_schedule.to_csv(OUTPUT_FOLDER / "fifo_schedule.csv", index=False)
    edd_schedule.to_csv(OUTPUT_FOLDER / "edd_schedule.csv", index=False)
    comparison.to_csv(OUTPUT_FOLDER / "kpi_comparison.csv", index=False)
    kpis_wc.to_csv(OUTPUT_FOLDER / "kpi_by_work_centre.csv", index=False)
    workload.to_csv(OUTPUT_FOLDER / "work_centre_load.csv", index=False)

    chart_late_orders(fifo_kpis, edd_kpis, kpis_wc, CHART_FOLDER / "late_orders_by_policy.png")
    chart_time_metrics(fifo_kpis, edd_kpis, CHART_FOLDER / "time_kpis_by_policy.png")
    chart_workload(workload, CHART_FOLDER / "workload_vs_capacity.png")

    write_excel(orders, fifo_schedule, edd_schedule, comparison, kpis_wc, workload,
                OUTPUT_FOLDER / "production_analysis.xlsx")

    print()
    print("KPI comparison (all orders):")
    print(comparison.to_string(index=False))
    print()
    print("Work centre load:")
    print(workload.to_string(index=False))
    print()
    print(f"Results saved in {OUTPUT_FOLDER}")


if __name__ == "__main__":
    main()
