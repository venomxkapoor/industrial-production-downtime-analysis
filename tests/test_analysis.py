"""Tests for industrial production and downtime analysis."""

from pathlib import Path
import sys
import openpyxl
import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from analyse_production import (
    build_summaries,
    calculate_run_metrics,
    load_and_check_data,
    run_pipeline,
    safe_pct,
    safe_rate,
)


def test_source_dataset_provenance_and_shape():
    xlsx = ROOT / "data" / "production_raw.xlsx"
    assert xlsx.exists()
    df, ev_count, ev_days = load_and_check_data(xlsx)
    assert len(df) == 59
    assert df["date"].nunique() == 56
    assert ev_count == 1388
    assert ev_days == 55


def test_data_validity_and_liters_formula():
    df, _, _ = load_and_check_data(ROOT / "data" / "production_raw.xlsx")
    assert int(df.isna().sum().sum()) == 0
    assert (df["monitored_time_dec"] > 0).all()
    assert (df["operation_time_dec"] > 0).all()
    assert (df["liters_produced"] == df["production_units"] * df["product_type_l"]).all()


def test_hand_calculated_first_run():
    df, _, _ = load_and_check_data(ROOT / "data" / "production_raw.xlsx")
    run_df = calculate_run_metrics(df)
    r0 = run_df.iloc[0]
    assert r0["date_str"] == "2022-07-22"
    assert int(r0["production_units"]) == 2849
    assert int(r0["liters_produced"]) == 8547
    assert r0["calc_efficiency_pct"] == pytest.approx(57.4513, abs=1e-3)
    assert r0["downtime_pct"] == pytest.approx(42.5487, abs=1e-3)
    assert r0["calc_units_per_op_hour"] == pytest.approx(562.6159, abs=1e-3)


def test_duplicate_dates_aggregation_no_double_counting():
    df, ev_count, ev_days = load_and_check_data(ROOT / "data" / "production_raw.xlsx")
    run_df = calculate_run_metrics(df)
    daily_df, _, _, _ = build_summaries(run_df, ev_count, ev_days)
    assert len(daily_df) == 56
    assert daily_df["runs"].sum() == 59
    assert int(daily_df["production_units"].sum()) == 60888
    assert int(daily_df["liters_produced"].sum()) == 212358
    assert daily_df["monitored_hours"].sum() == pytest.approx(238.65, abs=0.1)


def test_overall_and_product_kpis():
    df, ev_count, ev_days = load_and_check_data(ROOT / "data" / "production_raw.xlsx")
    run_df = calculate_run_metrics(df)
    _, kpi_df, product_df, monthly_df = build_summaries(run_df, ev_count, ev_days)
    kpis = dict(zip(kpi_df["Metric"], kpi_df["Value"]))
    assert kpis["Time-Weighted Efficiency"] == 42.66
    assert kpis["Time-Weighted Downtime"] == 57.64
    assert kpis["Overall Operating Rate"] == 598.06
    p3 = product_df[product_df["product_type_l"] == 3].iloc[0]
    p5 = product_df[product_df["product_type_l"] == 5].iloc[0]
    assert int(p3["runs"]) == 39 and int(p3["production_units"]) == 46041
    assert int(p5["runs"]) == 20 and int(p5["production_units"]) == 14847
    assert len(monthly_df) == 8


def test_zero_division_and_invalid_input_guards():
    assert safe_pct(25.0, 0.0) == 0.0
    assert safe_rate(100.0, -2.0) == 0.0
    sample = pd.DataFrame([{
        "date": pd.Timestamp("2023-01-01"), "product_type_l": 3, "production_units": 600,
        "liters_produced": 1800, "production_start_time": "10:00", "production_end_time": "12:00",
        "efficiency": 0.75, "gallons_per_hour": 400.0,
        "monitored_time_dec": 2.0, "operation_time_dec": 1.5, "pause_time_dec": 0.5,
    }])
    out = calculate_run_metrics(sample)
    assert out.loc[0, "calc_efficiency_pct"] == pytest.approx(75.0)
    assert out.loc[0, "downtime_pct"] == pytest.approx(25.0)
    assert out.loc[0, "calc_units_per_op_hour"] == pytest.approx(400.0)


def test_generated_files_and_excel_sheets():
    run_pipeline(ROOT)
    csv_path = ROOT / "outputs" / "daily_summary.csv"
    xlsx_path = ROOT / "outputs" / "production_summary.xlsx"
    pdf_path = ROOT / "reports" / "production_report.pdf"
    assert csv_path.exists()
    assert xlsx_path.exists()
    assert pdf_path.exists()
    assert pdf_path.read_bytes().startswith(b"%PDF")
    assert (ROOT / "outputs" / "charts" / "daily_production_and_downtime.png").exists()
    assert (ROOT / "outputs" / "charts" / "monthly_efficiency_and_downtime.png").exists()

    saved = pd.read_csv(csv_path)
    assert len(saved) == 59
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    assert wb.sheetnames == ["Selected Data", "KPI Summary", "Period Comparison"]
    assert wb["Selected Data"]["C61"].value == "=SUM(C2:C60)"
    assert wb["KPI Summary"]["B6"].value == 60888
