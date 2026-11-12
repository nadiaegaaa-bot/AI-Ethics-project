# Bulk – Energy Intelligence dashboard

Final project for AI Ethics (5th semester, ESADE). A dark-themed energy dashboard for Bulk's data centre sites: OS-IX Oslo, N01 Kristiansand and DK01 Esbjerg. It is also a progressive web app (PWA), so you can install it and use it offline.

> **Data status:** `js/data.js` contains **placeholder data** that reproduces the numbers in the design screenshots. The Excel workbook data has not been added yet.

## How to open
- **Quick look:** double-click `index.html`. The dashboard works straight from disk, but installing it and using it offline need a web address (see below).
- **As an app:** open the hosted address in Chrome or Edge and click **Install app** in the sidebar, or the install icon in the address bar. On iPhone/iPad, use Safari's **Share → Add to Home Screen**. On a Mac with Safari, use **File → Add to Dock**.
- **Locally over http:** in this folder, run `python3 -m http.server 8000` and open <http://localhost:8000>.

## Pages
- **Overview:** PUE, total energy, cooling energy and cost; energy by site; AI insights.
- **Sites:** a per-site view with four tabs. Performance shows the PUE trend vs. target and an energy breakdown donut. The other tabs are Energy breakdown, Cost & prices, and AI analysis.
- **Energy & Cost:** monthly energy and cost, average price, renewable share, and cost by site.
- **AI Insights:** the N01 cooling anomaly (actual vs. expected), suggested actions (ticks are saved), estimated impact, and a details pop-up.
- **Capacity:** installed, utilised and headroom per site, with a status.
- **Benchmarking:** PUE, energy cost and renewable share vs. competitors, plus the N01 electricity price.
- **Reports:** download reports (PDF via the print dialog; the capacity forecast as CSV) and create new ones.

The period selector and the site selector recalculate every figure.

## PWA files
- `manifest.webmanifest`: app name, colours, icons, and shortcuts to AI Insights, Energy & Cost and Reports.
- `sw.js`: the service worker. It caches the app for offline use. Bump `VERSION` when you change files so that installed copies update.
- `js/pwa.js`: registers the service worker and handles the install button, the "new version" prompt and the offline badge.
- `icons/`: app icons (192, 512, maskable 512, Apple touch, favicon).

On phones, the sidebar turns into a bottom tab bar.
