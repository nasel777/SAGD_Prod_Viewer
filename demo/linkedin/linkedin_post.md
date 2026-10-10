Monitoring SAGD well performance means juggling a lot of data at once: oil, water and steam rates, ESP parameters, pressures, operating indicators and downhole thermocouple temperatures. The data is usually there. What's hard is seeing all of it side by side quickly enough to trust what it's telling you.

So I built two browser-based tools for that: the SAGD Production Viewer and the SAGD Temp Viewer.

🔹 Production Viewer
• Time series with any number of variables, either stacked tracks or overlaid axes, with daily, weekly or monthly resolution and moving averages
• Crossplots of any two parameters, colored by date, e.g. emulsion rate vs. ESP frequency to spot pump degradation over time
• Statistics as bar or box charts that compare wells over a date range, or before vs. after an event such as a redrill
• Derived KPIs computed on the fly: SOR / CSOR, water cut, cumulative volumes, ESP flow normalized to 60 Hz, motor current per unit flow
• Shading for shut-in or flag periods, a "days since first oil" or "days since redrill" x-axis, and layouts you can save and reload

🔹 Temp Viewer
• Temperature or subcool profiles along the lateral, comparing dates or wells
• A time × position heatmap that shows steam breakthrough or cold spots developing along the wellbore
• Automatic cleaning of dead thermocouples and physically impossible subcool values

Being able to look at the same problem in several ways (trend, crossplot, profile, before/after statistics) is what turns a hunch into an analysis you can defend.

A few design choices I'm happy with:
✅ Each tool is a single offline HTML file. No installation and no server: open it, drop in the Excel workbook, and the data never leaves your machine.
✅ It's easy to adapt. It was built for SAGD, but variables, derived KPIs and screening views can be tailored to any field's characteristics and to what each team needs to watch.
✅ I developed it with AI coding tools (Claude Code), which let me go from a monitoring idea to a working feature in hours instead of weeks. This is a big shift for engineers who build their own tools.

The screenshots use a fully synthetic demo dataset (14 fictional wells), so no real field data is shown.

👉 Try the live demo: [demo link]

I'd love to hear how you monitor well performance, and which views you'd add.

#SAGD #OilAndGas #PetroleumEngineering #ProductionEngineering #ReservoirEngineering #DataVisualization #ArtificialLift #ESP #WellPerformance #AI

---
Suggested image order (upload in this order; LinkedIn shows the first four in the preview grid):
01_prod_viewer_overview.png: Production Viewer, multi-plot dashboard
02_temp_subcool_profile_heatmap.png: subcool profile and heatmap, steam breakthrough at point 3 before a redrill
03_esp_crossplot_by_date.png: ESP curve shifting over time (pump degradation)
04_redrill_before_after_bars.png: oil rate last normal / just before / after redrill
05_production_performance.png: oil, water cut, SOR stacked tracks with shut-in shading
06_steam_breakthrough_watch.png: subcool min, WHT, max temperature
07_multiwell_esp_screening.png: ESP flow at 60 Hz across producers
08_oil_vs_days_since_redrill.png: wells aligned on redrill day
09_field_oil_rate_bars.png: field-wide average oil rate comparison
10_temp_compare_wells.png: temperature profiles across wells, dead sensor cleaned in heatmap
