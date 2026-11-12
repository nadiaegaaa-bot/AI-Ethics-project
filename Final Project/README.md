# Bulk – Energy Intelligence dashboard

Final project for AI Ethics (5th semester, ESADE). A dark-themed energy dashboard for Bulk's data centre sites: N01 (Vennesla/Kristiansand), OS-IX (Oslo) and DK01 (Esbjerg).

## How to open
Double-click `index.html`. You don't need to install anything or run a server, because Chart.js is bundled in `js/vendor/`.

## Data
Every figure comes from `data/Bulk_Energy_Intelligence_Data_and_Competitor_Benchmarks.xlsx`. All sheets are used:

| Workbook sheet | Used in |
|---|---|
| Bulk Monthly Synthetic (63 site-months, Jan 2025 – Sep 2026) | Overview, Sites, Energy & Cost, AI Insights, Capacity, Reports |
| Competitor Benchmarks | Benchmarking → Reported PUE (verified), comparability notes |
| Peer Synthetic Scenarios | Benchmarking → Monthly scenarios (synthetic) |
| Real Company Facts | Capacity (N01 59 MW, 530 MW planned), PUE target 1.20, Data & Sources |
| Dashboard | Data & Sources → Import check (2025 totals recalculated and matched) |
| Sources & Methods, Data Dictionary, START HERE | Data & Sources |

The workbook is turned into `js/data.js` by a small script that uses only the Python standard library. After editing the Excel file, run:

```
python3 tools/excel_to_data.py
```

## Pages
- **Overview:** PUE, total energy, cooling energy and estimated cost, with change vs. the previous period. Energy by site. AI insights built from the workbook's REVIEW flags.
- **Sites:** a site selector and four tabs. Performance shows the PUE trend vs. the 1.20 design target and an energy donut. Energy breakdown shows monthly bars. Cost & prices shows cost and the price index. AI analysis shows the monthly cooling residual vs. the 12% threshold.
- **Energy & Cost:** monthly energy and cost, average price, cooling share, and a cost-by-site table.
- **AI Insights:** actual vs. expected cooling with the flagged months highlighted, a details pop-up, suggested actions (ticks are saved in the browser) and estimated impact.
- **Capacity:** N01 load vs. its reported 59 MW, and IT load per site.
- **Benchmarking:** verified reported PUE, synthetic monthly peer scenarios, and electricity price by zone.
- **Reports:** PDF reports open in a print view. The data export downloads as CSV. You can create new reports.
- **Data & Sources:** verified facts, the source register, the data dictionary, the workbook guide and the import check.

## Ethics and limitations (from the workbook)
- All monthly operating values are **synthetic** and are labelled "Synthetic data" in the UI.
- Cost = MWh × a synthetic price index. It is not Bulk's invoice, hedge or tariff.
- The "AI model" is a transparent heuristic: expected cooling = IT MWh × (0.115 + 0.0055 × max(temp − 8, 0) + site adjustment). It is not trained ML, and flags need engineer validation.
- Peer PUE figures differ in year, scope and method, so they are shown as context, not as a ranking. STACK is listed without a number.
- N01's 59 MW is reported capacity, not usage. OS-IX and DK01 capacity is unverified, so no utilisation is shown for them.
