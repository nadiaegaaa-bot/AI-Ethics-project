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
- **Computer:** *Intelligence → Copilot* in the menu (or **Ctrl+J** / **⌘J**) opens a full Copilot page: a chat on the left, "What you can ask" grouped by topic on the right, and a site and period selector. The round **Ask Copilot** button still opens a quick side panel on the other pages. The conversation is kept while you move between pages; **New chat** clears it.
- **Phone:** the round button opens the Copilot as a bottom sheet.
- **For everyone:** **Simple** gives a short plain-language answer, with key terms explained and the numbers one click away. **Detailed** shows all observations, possible explanations, next steps and the basis. It can explain terms ("What is PUE?", "MW vs. MWh"), answer specific questions ("What was cooling at N01 in June 2026?", "When will N01 reach 80% of capacity?", "Which site has the best PUE?", "How is expected cooling calculated?", "Tell me about DK01", "How do I reduce cooling?"), and analyse cooling, savings, capacity, peers and cost.
- **Navigation help:** ask "Where can I compare competitors?", "How do I upload new data?" or "Where is Capacity?". On the computer it points to the left menu; on the phone it points to the bottom bar and says when a feature is only on the computer.
- Answers are calculated in the browser from the loaded data with the dashboard's formulas. **No external AI service is used.** The chosen site or period applies unless the question names another one.

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
