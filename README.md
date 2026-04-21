# Copart Competitive Dashboard

An interactive, browser-based benchmark dashboard comparing **Copart** against IAA, Manheim, OpenLane, ACV Auctions, CarMax, and Carvana across 13 metrics — built for partner presentations and internal use.

**Live:** https://ma-ai-mo.github.io/copart-competitive-dashboard/

---

## Metrics Covered

| # | Metric | Source |
|---|--------|--------|
| 1 | Annual Revenue | SEC filings / investor reports |
| 2 | Market Capitalization | Public market data, year-end 2024 |
| 3 | Physical Locations | Company websites & press releases |
| 4 | Land Footprint (acres) | Copart annual report; peer estimates |
| 5 | Countries of Operation | Public disclosures |
| 6 | Vehicles Processed Annually | FY2024 estimates |
| 7 | Total Employees | LinkedIn & public disclosures |
| 8 | Mobile App Rating (iOS) | App Store, 2024 |
| 9 | Glassdoor Employee Rating | Glassdoor.com, 2024 |
| 10 | Years in Operation | Company founding dates |
| 11 | Registered Global Buyers | Copart investor disclosures; peer estimates |
| 12 | Buyer Nations | Copart public disclosures; peer estimates |
| 13 | Average Daily Inventory | Estimated from annual volume data |

---

## How to Run Locally

No build step needed — it's a single HTML file.

```bash
# Clone the repo
git clone https://github.com/MA-AI-MO/copart-competitive-dashboard.git
cd copart-competitive-dashboard

# Open in browser (any of these)
open index.html                         # macOS
start index.html                        # Windows
xdg-open index.html                     # Linux
```

---

## How to Update Data

All data lives in the `<script>` block at the bottom of `index.html`.

Each chart is created with a `mkBar(...)` call. The array order is always:

```
['Copart', 'IAA', 'Manheim', 'OpenLane', 'ACV', 'CarMax', 'Carvana']
```

**Example — updating Market Cap:**
```js
mkBar('mktCapChart',
  [48, 7, null, 0.8, 1.2, 12, 35],   // ← edit these values
  '$B', 'Market Cap (USD Billions)',
  v => v == null ? '' : `$${v}B`
);
```

Use `null` for companies where a metric is not applicable (e.g. private companies, retail-only platforms).

The **radar chart** at the bottom auto-normalizes all values, so it updates automatically when you change the bar chart data.

---

## How to Contribute

1. Fork this repo
2. Make your changes in `index.html`
3. Open a Pull Request with a brief description of what changed and why

For data updates, please include a source link in the PR description.

---

## Tech Stack

- [Chart.js 4.4](https://www.chartjs.org/) — charting
- [chartjs-plugin-datalabels 2.2](https://chartjs-plugin-datalabels.netlify.app/) — bar labels
- [Inter](https://fonts.google.com/specimen/Inter) — typography
- Pure HTML/CSS/JS — no build tools, no dependencies to install

---

## Deployment

The dashboard is hosted on **GitHub Pages** from the `main` branch root.  
Any push to `main` automatically updates the live page within ~1 minute.

---

*Data is directional and approximate — intended for presentation use, not investment decisions.*
