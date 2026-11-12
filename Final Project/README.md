# Bulk – Energy Intelligence dashboard

Final project for AI Ethics (5th semester, ESADE). A dark-themed energy dashboard for Bulk's data centre sites (OS-IX Oslo, N01 Kristiansand, DK01 Esbjerg).

## How to open
Double-click `index.html`. No install or server is needed; Chart.js is bundled in `js/vendor/`.

## Pages
- **Overview** – PUE, total energy, cooling energy, cost; energy by site; AI insights.
- **Sites** – per-site view with tabs: Performance (PUE trend vs. target, energy breakdown donut), Energy breakdown, Cost & prices, AI analysis.
- **Energy & Cost** – monthly energy and cost, average price, renewable share, cost by site.
- **AI Insights** – N01 cooling anomaly (actual vs. expected), suggested actions (ticks are saved in the browser), estimated impact, details pop-up.
- **Capacity** – installed, utilised and headroom per site, with status.
- **Benchmarking** – PUE, energy cost and renewable share vs. competitors; N01 electricity price.
- **Reports** – download reports (PDF via the print dialog, capacity forecast as CSV) and create new ones.

The period selector (last 30 days / 6 months / 12 months) and site selector recalculate every figure.

## Data
All data lives in `js/data.js` (`window.BULK_DATA`):
- `sites` – location, installed/utilised MW, renewable share
- `monthly` – per site and month: IT, cooling and other energy (MWh) and cost (NOK)
- `coolingDailyN01` – daily actual vs. expected cooling for June 2024
- `benchmark` – competitor PUE, cost, renewable share and N01 price
- `insights`, `reports`

The current values reproduce the figures in the design screenshots. To use the Excel workbook, replace these tables with its sheets; the pages calculate everything from them.
