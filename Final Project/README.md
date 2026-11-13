# Bulk – Energy Intelligence dashboard

Final project for AI Ethics (5th semester, ESADE). A dark-themed energy dashboard for Bulk's data centre sites: N01 (Vennesla/Kristiansand), OS-IX (Oslo) and DK01 (Esbjerg). It is also a progressive web app (PWA), so you can install it and use it offline.

**Live app:** <https://nadiaegaaa-bot.github.io/AI-Ethics-project/>

## How to open
- **As an app:** open the live link in Chrome or Edge and click **Install app** in the sidebar. On iPhone/iPad, use Safari's **Share → Add to Home Screen**. On a Mac with Safari, use **File → Add to Dock**. The installed app updates itself and works offline.
- **From disk:** double-click `index.html`. Everything works, but installing and offline mode need the web address.

## Data
Every figure comes from `data/Bulk_Energy_Intelligence_Data_and_Competitor_Benchmarks.xlsx`. All sheets are used:

| Workbook sheet | Used in |
|---|---|
| Bulk Monthly Synthetic (63 site-months, Jan 2025 – Sep 2026) | Overview, Sites, Energy & Cost, AI Insights, Capacity, Reports |
| Competitor Benchmarks | Benchmarking → Reported PUE (verified), comparability notes |
| Peer Synthetic Scenarios | Benchmarking → Monthly scenarios (synthetic) |
| Real Company Facts | Capacity (N01 59 MW, 530 MW planned), PUE target 1.20, Data & Methods |
| Dashboard | Data & Methods → Import check (2025 totals recalculated and matched) |
| Sources & Methods, Data Dictionary, START HERE | Data & Methods |

`js/data.js` is generated from the workbook by a script that uses only the Python standard library. After editing the Excel file, run this in the `Final Project` folder:

```
python3 tools/excel_to_data.py
```

Then bump `VERSION` in `sw.js` so installed copies pick up the change.

## Uploading new data (Reports page)
On **Reports → Upload a report to update the dashboard**, drop or choose a file:
- **Excel (.xlsx)** in the Bulk data workbook format. The monthly sheet updates the numbers. Any of the other sheets present (benchmarks, peer scenarios, company facts, sources, dictionary, guide, Dashboard summary) replace those tables.
- **CSV** with one row per site and month. Required columns are `Month`, `Site`, `IT energy (MWh)` and `Cooling (MWh)`. Optional columns are other overhead, total, price, temperature, expected cooling, IT load, alert and provenance. Comma or semicolon separators, comma decimals and months like `2026-10`, `Oct 2026` or `01.10.2026` all work. **Download template** gives you the current data in the right format to edit.

A preview shows how many site-months are new or updated, any new sites, the cooling flags and warnings, all before anything changes. You can **merge** (update matching months, add new ones) or **replace** the monthly data. Missing values are calculated the same way as in the workbook: total = IT + cooling + other, PUE = total / IT, cost = total × price, and expected cooling uses formula M05. Rows above the 12% residual are flagged REVIEW.

Uploads are stored only in your browser. Each one can be removed, or you can **Restore original data**. PDF files can't be read automatically, so upload the numbers as Excel or CSV.

## Help, labels and navigation
- **Info buttons (ⓘ):** every card and every menu item has an ⓘ. Hover over it or click it to see what the number means, what it is based on and how to read the chart. The ⓘ next to each page title explains the page.
- **Executive summary:** the Overview starts with "About this platform", a short introduction and the key numbers for the selected period.
- **Confidential:** every page title has a Confidential badge, and the printed reports are marked confidential.
- **Data centre picker:** on Sites and AI Insights, choose the data centre with the buttons under the title.
- **Verne:** Verne Iceland and Verne Finland are combined into one competitor (simple average 1.295, because per-site energy is not published).
- **Settings → Contact & help:** a placeholder; this will be developed later.

## AI Copilot
The **Ask Copilot** button (bottom right, or **Ctrl+J** / **⌘J**) opens a chat panel on both the computer and the phone. It answers questions about cooling anomalies, optimisation potential, capacity risks, peer comparison, cost drivers, PUE and energy use. Answers are calculated in the browser from the loaded data with the dashboard's own formulas, with **no external AI service**, and are split into *Observations (from the data)*, *Possible explanations (not confirmed)* and *Suggested next steps*, plus the scope and basis. Name a site (N01, Oslo…) or a period (2025, last month) in the question to change the focus; otherwise the current page's context is used.

## Phone app vs. computer dashboard
On screens up to 640 px wide (phones, including the installed app on a phone), the app shows only the core pages: **Overview, Sites, Energy & Cost and AI Insights**, in a bottom tab bar. The "About this platform" text is shortened to one line. Other pages redirect to the Overview on a phone. On a computer or tablet, the full dashboard with all pages and the full description is unchanged.

## Pages
- **Overview:** PUE, total energy, cooling energy and estimated cost, with change vs. the previous period. Energy by site. AI insights from the workbook's REVIEW flags.
- **Sites:** four tabs. Performance shows the PUE trend vs. the 1.20 design target and an energy donut. Energy breakdown shows monthly bars. Cost & prices shows cost and the price index. AI analysis shows the cooling residual vs. the 12% threshold.
- **Energy & Cost:** monthly energy and cost, average price, cooling share, and a cost-by-site table.
- **AI Insights:** actual vs. expected cooling with the flagged months highlighted, a details pop-up, suggested actions and estimated impact.
- **Capacity:** N01 load vs. its reported 59 MW, and IT load per site.
- **Benchmarking:** verified reported PUE, synthetic monthly peer scenarios, and electricity price by zone. **Click any competitor** (a bar, a row in the list, or a name in the chart legend) to add it to a side-by-side comparison with Bulk, up to 4 at a time. The comparison shows PUE and the difference vs. Bulk, the gap to the 1.20 target, what the gap would mean in GWh and NOK at Bulk's scale, the monthly scenario range, energy scope, status, cautions and source. The selection is remembered.
- **Reports:** PDF reports open in a print view, and the data export downloads as CSV. You can create new reports.
- **Data & Methods:** verified facts, the source register, the data dictionary, the workbook guide and the import check.

## Ethics and limitations (from the workbook)
- All monthly operating values are **synthetic** and are labelled "Synthetic data" in the UI.
- Cost = MWh × a synthetic price index. It is not Bulk's invoice, hedge or tariff.
- The "AI model" is a transparent heuristic: expected cooling = IT MWh × (0.115 + 0.0055 × max(temp − 8, 0) + site adjustment). It is not trained ML, and flags need engineer validation.
- Peer PUE figures differ in year, scope and method, so they are shown as context, not as a ranking. STACK is listed without a number.
- N01's 59 MW is reported capacity, not usage. OS-IX and DK01 capacity is unverified, so no utilisation is shown for them.

## PWA files
- `manifest.webmanifest`: app name, colours, icons and shortcuts.
- `sw.js`: the service worker that caches the app for offline use.
- `js/pwa.js`: handles the install button, the update prompt and the offline badge.
- `icons/`: the app icons.
