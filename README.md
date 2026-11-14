# Industrial Production & Downtime Analysis — Python and Excel

A transparent, beginner-friendly industrial operations analysis using genuine 2022–2023 beverage bottling line records. The project quantifies production volume, monitored operating versus pause time, downtime percentages, bottle-size differences (3 L vs. 5 L), and monthly trends using Python (`pandas`, `matplotlib`, `openpyxl`) and Excel.

---

## Business Questions

1. **Production Output**: How many bottle units and litres were recorded across the observation period?
2. **Time Breakdown**: How much monitored time was spent actively operating versus paused?
3. **Downtime Peaks**: Which production runs and calendar periods experienced the highest downtime hours and downtime share?
4. **Category & Period Variation**: How do efficiency, downtime share, and operating output rate vary across bottle sizes (3 L vs. 5 L) and months?
5. **Operational Observations**: Which data-supported patterns and reporting limitations should an analyst note before investigating root causes?

---

## Official Dataset Provenance

- **Title**: *Industrial Production Time-Series Dataset from a Beverage Bottling Line*
- **Source Repository**: Zenodo (`https://zenodo.org/records/18146866`)
- **DOI**: `10.5281/zenodo.18146866`
- **Licence**: Creative Commons Attribution 4.0 International (**CC BY 4.0**) — see [LICENSE_DATA.md](LICENSE_DATA.md)
- **File**: `data/production_raw.xlsx` (preserved unmodified)

These are **real historical operational records** released through a public research dataset. They represent preprocessed daily and hourly operational summaries rather than raw PLC sensor streams.

### Worksheet Selection & Duplicate-Date Handling

- **Principal source**: `daily_operation_summary` (**59 production runs** across **56 unique calendar dates**, from `2022-07-22` to `2023-02-10`, with zero missing values).
- **Supplementary log**: `downtime_event_log` (**1,388 logged stop events** across 55 dates, counted for reference without merging all sheets).
- **Multi-run dates**: Three dates (`2022-09-28`, `2022-10-10`, `2022-10-24`) have two rows—one 3 L run and one 5 L run. Inspecting `production_start_time` and `production_end_time` confirms these runs occurred sequentially without time overlap, so summing hours and output by calendar date does not double-count.

---

## Column Definitions & Formulas

### Source Columns (`daily_operation_summary`)
| Column | Unit / Scale | Description |
| :--- | :--- | :--- |
| `date` | Date | Operating calendar date (`2022-07-22` to `2023-02-10`) |
| `product_type_l` | Litres (`3` or `5`) | Bottle format produced during the run |
| `production_units` | Bottles | Total bottle count recorded |
| `liters_produced` | Litres | Exactly `production_units × product_type_l` on all 59 runs |
| `monitored_time_dec` | Decimal hours | Total monitored span for the run |
| `operation_time_dec` | Decimal hours | Active operating time during the run |
| `pause_time_dec` | Decimal hours | Recorded pause / downtime during the run |
| `efficiency` | Ratio (`0–1`) | **Source-reported** ratio (`operation_time_dec / monitored_time_dec`) |
| `gallons_per_hour` | Rate | **Source-reported** rate (see Data Quality Note below) |

### Independently Calculated Metrics
- **Downtime Percentage (%)**: `(pause_time_dec / monitored_time_dec) × 100` (guarded against division by zero)
- **Calculated Efficiency (%)**: `(operation_time_dec / monitored_time_dec) × 100`
- **Calculated Operating Production Rate**: `production_units / operation_time_dec` (consistent denominator across all 59 runs)
- **Time-Weighted Summary Rates**: Calculated as `SUM(hours) / SUM(monitored_hours) × 100` across grouped runs, reported alongside unweighted run averages.

---

## Key Findings

### Overall Summary (59 Runs / 56 Unique Dates)
| Metric | Value | Unit |
| :--- | :---: | :--- |
| Production Runs Recorded | 59 | runs |
| Unique Operating Dates | 56 | days |
| Logged Stop Events (`downtime_event_log`) | 1,388 | events |
| Total Production Output | 60,888 | bottles |
| Total Volume Produced | 212,358 | litres |
| Total Monitored Time | 238.65 | hours |
| Total Operating Time | 101.81 | hours |
| Total Pause / Downtime | 137.56 | hours |
| Time-Weighted Efficiency | 42.66 | % |
| Average Run Efficiency (unweighted) | 46.52 | % |
| Time-Weighted Downtime | 57.64 | % |
| Average Run Downtime (unweighted) | 54.28 | % |
| Overall Operating Production Rate | 598.06 | units / operating h |

### Product Category Comparison (3 L vs. 5 L)
| Bottle Size | Runs | Units | Litres | Monitored h | Operating h | Pause h | Weighted Eff. (%) | Weighted Downtime (%) | Mean Rate (units/op h) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **3 L** | 39 | 46,041 | 138,123 | 163.56 | 73.11 | 90.14 | 44.70 | 55.11 | 654.57 |
| **5 L** | 20 | 14,847 | 74,235 | 75.10 | 28.70 | 47.43 | 38.22 | 63.15 | 583.74 |

### Period Highlights
- **Highest Volume & Downtime Month (`2022-12`)**: 15 runs producing 16,051 units across 79.30 monitored hours, with **50.29 pause hours (63.42% weighted downtime)** and 37.67% weighted efficiency.
- **Highest Efficiency Month (`2023-01`)**: 9 runs producing 10,042 units across 26.34 monitored hours, achieving **53.98% weighted efficiency** (44.56% weighted downtime).
- **Highest Single-Run Downtime**:
  - By hours: `2022-12-13` (5 L run, **7.00 pause hours** out of 9.88 monitored hours; 70.85% downtime) and `2022-12-06` (3 L run, **6.86 pause hours** out of 8.36 monitored hours; 82.10% downtime).
  - By percentage: `2022-12-19` (5 L run, **84.62% downtime**, 4.86 pause hours out of 5.74 monitored hours).

---

## Data Quality Notes & Limitations

1. **Time-Balance Discrepancy in Source Records**: On 44 of 59 runs, `operation_time_dec + pause_time_dec` matches `monitored_time_dec` within 0.01 h. Across the remaining 15 runs, small timing offsets lead to a cumulative `+0.72 h` difference (`101.81 + 137.56 = 239.37 h` vs. `238.65 h` monitored), so weighted efficiency (42.66%) plus weighted downtime (57.64%) sums to 100.30%.
2. **Inconsistent Denominator in Source `gallons_per_hour`**: Checking the source workbook shows `gallons_per_hour` matches `production_units / operation_time_dec` in the first 18 rows and `production_units / monitored_time_dec` in 40 of the remaining 41 rows. Therefore, this analysis separates source-reported `gallons_per_hour` from the consistently defined `calc_units_per_op_hour` (`production_units / operation_time_dec`).
3. **No Root-Cause Attribution**: The public dataset does not include failure reason codes, changeover logs, operator shifts, or maintenance records. High pause percentages indicate periods worth investigating locally, not proven equipment faults or staffing issues.
4. **Scope**: This is a descriptive operational analysis built as an educational portfolio project using Python, pandas, openpyxl, and matplotlib.

---

## Project Structure

```text
industrial-production-downtime-analysis/
├── README.md
├── LICENSE_DATA.md
├── requirements.txt
├── data/
│   └── production_raw.xlsx
├── src/
│   └── analyse_production.py
├── outputs/
│   ├── daily_summary.csv
│   ├── production_summary.xlsx
│   └── charts/
│       ├── daily_production_and_downtime.png
│       └── monthly_efficiency_and_downtime.png
├── reports/
│   └── production_report.pdf
└── tests/
    └── test_analysis.py
```

---

## How to Reproduce (Windows PowerShell)

```powershell
# 1. Install the 4 dependencies
pip install -r requirements.txt

# 2. Run the analysis script (creates CSV, charts, 3-sheet Excel workbook, and 2-page PDF)
python src/analyse_production.py

# 3. Run the automated verification suite
python -m pytest -v
```
