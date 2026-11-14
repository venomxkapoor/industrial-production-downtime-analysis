"""
Automated checks for the scheduling logic and the saved outputs.

Run from the project folder (after running the three scripts in src/):
    python -m pytest
"""

import re

import pandas as pd
import pytest
from openpyxl import load_workbook

import analyse_schedules as an
import build_report as br
import generate_data as gen

TOLERANCE = 1e-9


# ---------------------------------------------------------------------------
# Helper data
# ---------------------------------------------------------------------------

@pytest.fixture
def orders():
    """The simulated dataset, created in memory with the fixed seed."""
    return gen.generate_orders(gen.NUMBER_OF_ORDERS, gen.SEED)


@pytest.fixture
def fifo_schedule(orders):
    return an.build_schedule(an.sort_fifo(orders), "FIFO")


@pytest.fixture
def edd_schedule(orders):
    return an.build_schedule(an.sort_edd(orders), "EDD")


# ---------------------------------------------------------------------------
# 1. Data generation and data checks
# ---------------------------------------------------------------------------

def test_data_generation_is_reproducible():
    first = gen.generate_orders(80, 42)
    second = gen.generate_orders(80, 42)
    pd.testing.assert_frame_equal(first, second)


def test_different_seed_gives_different_data():
    first = gen.generate_orders(80, 42)
    other = gen.generate_orders(80, 7)
    assert not first.equals(other)


def test_saved_csv_matches_generator(orders):
    saved = pd.read_csv(an.INPUT_FILE)
    pd.testing.assert_frame_equal(saved, orders, check_dtype=False)


def test_generated_data_passes_checks(orders):
    assert an.check_orders(orders) == []
    assert orders["order_id"].is_unique
    assert (orders["processing_time_h"] > 0).all()
    assert orders["due_time_h"].notna().all()
    assert (orders["release_time_h"] == 0).all()


def test_check_orders_finds_problems(orders):
    broken = orders.copy()
    broken.loc[0, "processing_time_h"] = -2
    broken.loc[1, "order_id"] = broken.loc[2, "order_id"]
    broken.loc[3, "due_time_h"] = None
    problems = an.check_orders(broken)
    assert len(problems) == 3


# ---------------------------------------------------------------------------
# 2. Hand-calculated 5-order example (also shown in the Excel workbook)
# ---------------------------------------------------------------------------
#   Order  p  due   FIFO start-end   EDD position / start-end
#   EX-1   3   5    0-3   (L=-2)     3rd: 3-6   (L=+1)
#   EX-2   2   4    3-5   (L=+1)     2nd: 1-3   (L=-1)
#   EX-3   4  12    5-9   (L=-3)     4th: 6-10  (L=-2)
#   EX-4   1   3    9-10  (L=+7)     1st: 0-1   (L=-2)
#   EX-5   5  15   10-15  (L= 0)     5th: 10-15 (L= 0)

def test_worked_example_fifo_matches_hand_calculation():
    schedule = an.build_schedule(an.sort_fifo(an.make_worked_example_orders()), "FIFO")
    assert list(schedule["order_id"]) == ["EX-1", "EX-2", "EX-3", "EX-4", "EX-5"]
    assert list(schedule["start_time_h"]) == [0, 3, 5, 9, 10]
    assert list(schedule["completion_time_h"]) == [3, 5, 9, 10, 15]
    assert list(schedule["lateness_h"]) == [-2, 1, -3, 7, 0]
    assert list(schedule["tardiness_h"]) == [0, 1, 0, 7, 0]
    assert list(schedule["is_late"]) == [False, True, False, True, False]

    kpis = an.calculate_kpis(schedule)
    assert kpis["Late orders"] == 2
    assert kpis["Late orders (%)"] == 40.0
    assert kpis["Average waiting time (h)"] == 5.4    # (0+3+5+9+10)/5
    assert kpis["Average flow time (h)"] == 8.4       # (3+5+9+10+15)/5
    assert kpis["Average tardiness (h)"] == 1.6       # (0+1+0+7+0)/5
    assert kpis["Maximum lateness (h)"] == 7
    assert kpis["Makespan (h)"] == 15


def test_worked_example_edd_matches_hand_calculation():
    schedule = an.build_schedule(an.sort_edd(an.make_worked_example_orders()), "EDD")
    assert list(schedule["order_id"]) == ["EX-4", "EX-2", "EX-1", "EX-3", "EX-5"]
    assert list(schedule["start_time_h"]) == [0, 1, 3, 6, 10]
    assert list(schedule["completion_time_h"]) == [1, 3, 6, 10, 15]
    assert list(schedule["lateness_h"]) == [-2, -1, 1, -2, 0]
    assert list(schedule["tardiness_h"]) == [0, 0, 1, 0, 0]

    kpis = an.calculate_kpis(schedule)
    assert kpis["Late orders"] == 1
    assert kpis["Average waiting time (h)"] == 4.0    # (0+1+3+6+10)/5
    assert kpis["Average flow time (h)"] == 7.0       # (1+3+6+10+15)/5
    assert kpis["Average tardiness (h)"] == 0.2       # 1/5
    assert kpis["Maximum lateness (h)"] == 1
    assert kpis["Makespan (h)"] == 15


def test_order_finishing_exactly_on_due_time_is_not_late():
    schedule = an.build_schedule(an.sort_fifo(an.make_worked_example_orders()), "FIFO")
    ex5 = schedule[schedule["order_id"] == "EX-5"].iloc[0]
    assert ex5["lateness_h"] == 0
    assert not ex5["is_late"]


def test_edd_tie_is_broken_by_original_sequence():
    tied = pd.DataFrame({
        "order_id": ["A", "B", "C"],
        "original_sequence": [1, 2, 3],
        "work_centre": ["WC1", "WC1", "WC1"],
        "product_category": ["X", "X", "X"],
        "processing_time_h": [2.0, 1.0, 1.0],
        "due_time_h": [10, 5, 5],
        "release_time_h": [0, 0, 0],
    })
    schedule = an.build_schedule(an.sort_edd(tied), "EDD")
    assert list(schedule["order_id"]) == ["B", "C", "A"]


def test_two_work_centres_run_in_parallel():
    two_machines = pd.DataFrame({
        "order_id": ["A", "B", "C"],
        "original_sequence": [1, 2, 3],
        "work_centre": ["WC1", "WC2", "WC1"],
        "product_category": ["X", "X", "X"],
        "processing_time_h": [2.0, 3.0, 4.0],
        "due_time_h": [10, 10, 10],
        "release_time_h": [0, 0, 0],
    })
    schedule = an.build_schedule(an.sort_fifo(two_machines), "FIFO")
    starts = dict(zip(schedule["order_id"], schedule["start_time_h"]))
    assert starts == {"A": 0, "B": 0, "C": 2}   # B starts at 0 on WC2


# ---------------------------------------------------------------------------
# 3. Properties of the full schedules
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("policy", ["FIFO", "EDD"])
def test_every_order_scheduled_exactly_once(orders, fifo_schedule, edd_schedule, policy):
    schedule = fifo_schedule if policy == "FIFO" else edd_schedule
    assert len(schedule) == len(orders)
    assert schedule["order_id"].is_unique
    assert set(schedule["order_id"]) == set(orders["order_id"])


@pytest.mark.parametrize("policy", ["FIFO", "EDD"])
def test_completion_equals_start_plus_processing(fifo_schedule, edd_schedule, policy):
    schedule = fifo_schedule if policy == "FIFO" else edd_schedule
    difference = schedule["completion_time_h"] - (schedule["start_time_h"] + schedule["processing_time_h"])
    assert (difference.abs() < TOLERANCE).all()


@pytest.mark.parametrize("policy", ["FIFO", "EDD"])
def test_no_overlap_and_no_idle_time_on_a_work_centre(fifo_schedule, edd_schedule, policy):
    schedule = fifo_schedule if policy == "FIFO" else edd_schedule
    for work_centre in schedule["work_centre"].unique():
        jobs = schedule[schedule["work_centre"] == work_centre].sort_values("position_on_work_centre")
        starts = list(jobs["start_time_h"])
        ends = list(jobs["completion_time_h"])
        assert starts[0] == 0
        for i in range(1, len(starts)):
            assert starts[i] >= ends[i - 1] - TOLERANCE          # no overlap
            assert abs(starts[i] - ends[i - 1]) < TOLERANCE      # no idle gap


@pytest.mark.parametrize("policy", ["FIFO", "EDD"])
def test_delivery_measures_are_correct(fifo_schedule, edd_schedule, policy):
    s = fifo_schedule if policy == "FIFO" else edd_schedule
    assert ((s["waiting_time_h"] - (s["start_time_h"] - s["release_time_h"])).abs() < TOLERANCE).all()
    assert ((s["flow_time_h"] - (s["completion_time_h"] - s["release_time_h"])).abs() < TOLERANCE).all()
    assert ((s["lateness_h"] - (s["completion_time_h"] - s["due_time_h"])).abs() < TOLERANCE).all()
    for index, row in s.iterrows():
        assert row["tardiness_h"] == max(0, row["lateness_h"])
        assert row["is_late"] == (row["tardiness_h"] > 0)


def test_fifo_follows_original_sequence(fifo_schedule):
    for work_centre in fifo_schedule["work_centre"].unique():
        jobs = fifo_schedule[fifo_schedule["work_centre"] == work_centre].sort_values("position_on_work_centre")
        assert jobs["original_sequence"].is_monotonic_increasing


def test_edd_follows_due_time_order(edd_schedule):
    for work_centre in edd_schedule["work_centre"].unique():
        jobs = edd_schedule[edd_schedule["work_centre"] == work_centre].sort_values("position_on_work_centre")
        assert jobs["due_time_h"].is_monotonic_increasing
        # within equal due times, the original sequence must increase
        due = list(jobs["due_time_h"])
        seq = list(jobs["original_sequence"])
        for i in range(1, len(due)):
            if due[i] == due[i - 1]:
                assert seq[i] > seq[i - 1]


def test_both_policies_use_identical_input(fifo_schedule, edd_schedule):
    columns = ["order_id", "work_centre", "product_category", "processing_time_h", "due_time_h", "release_time_h"]
    fifo_input = fifo_schedule[columns].sort_values("order_id").reset_index(drop=True)
    edd_input = edd_schedule[columns].sort_values("order_id").reset_index(drop=True)
    pd.testing.assert_frame_equal(fifo_input, edd_input)


def test_total_processing_and_makespan_unchanged(fifo_schedule, edd_schedule):
    assert fifo_schedule["processing_time_h"].sum() == edd_schedule["processing_time_h"].sum()
    for work_centre in fifo_schedule["work_centre"].unique():
        fifo_end = fifo_schedule[fifo_schedule["work_centre"] == work_centre]["completion_time_h"].max()
        edd_end = edd_schedule[edd_schedule["work_centre"] == work_centre]["completion_time_h"].max()
        workload = fifo_schedule[fifo_schedule["work_centre"] == work_centre]["processing_time_h"].sum()
        assert fifo_end == edd_end == workload
    assert fifo_schedule["completion_time_h"].max() == edd_schedule["completion_time_h"].max()


def test_edd_max_lateness_not_worse_than_fifo_per_work_centre(fifo_schedule, edd_schedule):
    # Classical single-machine result: EDD minimises maximum lateness.
    for work_centre in fifo_schedule["work_centre"].unique():
        fifo_max = fifo_schedule[fifo_schedule["work_centre"] == work_centre]["lateness_h"].max()
        edd_max = edd_schedule[edd_schedule["work_centre"] == work_centre]["lateness_h"].max()
        assert edd_max <= fifo_max


# ---------------------------------------------------------------------------
# 4. Saved outputs agree with the schedules (run the scripts first)
# ---------------------------------------------------------------------------

def test_saved_schedules_match_calculation(fifo_schedule, edd_schedule):
    saved_fifo = pd.read_csv(an.OUTPUT_FOLDER / "fifo_schedule.csv")
    saved_edd = pd.read_csv(an.OUTPUT_FOLDER / "edd_schedule.csv")
    pd.testing.assert_frame_equal(saved_fifo, fifo_schedule, check_dtype=False)
    pd.testing.assert_frame_equal(saved_edd, edd_schedule, check_dtype=False)


def test_kpi_file_matches_detailed_schedules(fifo_schedule, edd_schedule):
    kpis = pd.read_csv(an.OUTPUT_FOLDER / "kpi_comparison.csv").set_index("metric")
    for policy, schedule in [("FIFO", fifo_schedule), ("EDD", edd_schedule)]:
        assert kpis.loc["Number of orders", policy] == len(schedule)
        assert kpis.loc["Late orders", policy] == (schedule["tardiness_h"] > 0).sum()
        assert abs(kpis.loc["Average waiting time (h)", policy] - schedule["waiting_time_h"].mean()) < 0.005
        assert abs(kpis.loc["Average tardiness (h)", policy] - schedule["tardiness_h"].mean()) < 0.005
        assert abs(kpis.loc["Maximum lateness (h)", policy] - schedule["lateness_h"].max()) < 0.005
        assert abs(kpis.loc["Makespan (h)", policy] - schedule["completion_time_h"].max()) < 0.005


def test_workload_file_matches_orders(orders):
    workload = pd.read_csv(an.OUTPUT_FOLDER / "work_centre_load.csv").set_index("work_centre")
    for work_centre in orders["work_centre"].unique():
        assigned = orders[orders["work_centre"] == work_centre]["processing_time_h"].sum()
        assert workload.loc[work_centre, "assigned_hours"] == assigned
        expected_ratio = assigned / an.CAPACITY_BUDGET_H
        assert abs(workload.loc[work_centre, "capacity_pressure_ratio"] - expected_ratio) < 0.0005
        assert workload.loc[work_centre, "overload_h"] == max(0, assigned - an.CAPACITY_BUDGET_H)


def test_excel_values_match_csv_outputs():
    excel_file = an.OUTPUT_FOLDER / "production_analysis.xlsx"
    sheet_to_csv = {
        "FIFO Schedule": "fifo_schedule.csv",
        "EDD Schedule": "edd_schedule.csv",
        "KPI Comparison": "kpi_comparison.csv",
        "KPI by Work Centre": "kpi_by_work_centre.csv",
        "Work Centre Load": "work_centre_load.csv",
    }
    for sheet_name, csv_name in sheet_to_csv.items():
        from_excel = pd.read_excel(excel_file, sheet_name=sheet_name)
        from_csv = pd.read_csv(an.OUTPUT_FOLDER / csv_name)
        pd.testing.assert_frame_equal(from_excel, from_csv, check_dtype=False)

    orders_excel = pd.read_excel(excel_file, sheet_name="Original Orders")
    pd.testing.assert_frame_equal(orders_excel, pd.read_csv(an.INPUT_FILE), check_dtype=False)


def test_excel_worked_example_has_formulas_and_hand_values():
    workbook = load_workbook(an.OUTPUT_FOLDER / "production_analysis.xlsx")
    sheet = workbook["Worked Example"]
    formulas = []
    max_lateness_results = []
    late_order_results = []
    for row in sheet.iter_rows(values_only=True):
        for value in row:
            if isinstance(value, str) and value.startswith("="):
                formulas.append(value)
        if row[0] == "Maximum lateness (h)":
            max_lateness_results.append(row[2])
        if row[0] == "Late orders":
            late_order_results.append(row[2])
    assert "=MAX(0,I6)" in formulas          # tardiness formula of the first FIFO order
    assert max_lateness_results == [7, 1]    # FIFO, EDD (hand calculation)
    assert late_order_results == [2, 1]


# ---------------------------------------------------------------------------
# 5. The PDF report agrees with the results
# ---------------------------------------------------------------------------

def test_report_findings_use_actual_numbers():
    results = br.load_results()
    text = " ".join(br.make_findings(results))
    fifo = results["fifo"]
    edd = results["edd"]
    assert f"{int(fifo['Late orders'])} of {int(fifo['Number of orders'])}" in text
    assert f"{int(edd['Late orders'])} ({edd['Late orders (%)']:.1f} %) under EDD" in text
    assert f"{fifo['Maximum lateness (h)']:g} h vs {edd['Maximum lateness (h)']:g} h" in text
    assert f"makespan ({fifo['Makespan (h)']:g} h)" in text
    # direction words must match the numbers
    if edd["Late orders"] > fifo["Late orders"]:
        assert "higher under EDD. EDD does not guarantee" in text
    if edd["Maximum lateness (h)"] < fifo["Maximum lateness (h)"]:
        assert "Maximum lateness: " in text and "- lower under EDD." in text


def test_report_pdf_exists_and_has_two_pages():
    pdf_bytes = br.REPORT_FILE.read_bytes()
    assert pdf_bytes.startswith(b"%PDF")
    page_count = len(re.findall(rb"/Type\s*/Page[^s]", pdf_bytes))
    assert page_count == 2
