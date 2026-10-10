As a SAGD production engineer, I built two tools to make well performance data easier to explore: SAGD Production Viewer and SAGD Temp Viewer.

Time-series plots are my starting point. But looking at the same data from several angles (trend, crossplot, profile, before/after statistics) is what lets me challenge an interpretation and reach conclusions I can trust.

🔹 Production Viewer brings together oil, water and steam rates, ESP parameters and well operating indicators:
• Time-series plots for single wells or multi-well comparisons, as stacked tracks or overlaid axes
• Crossplots between any two parameters, colored by well or date. Plotting emulsion rate against ESP frequency over time, for example, shows pump degradation at a glance
• Statistical bar and box charts, including before vs. after redrill comparisons and trends aligned by days since redrill
• Derived KPIs calculated on the fly: SOR / CSOR, water cut, cumulative volumes, ESP flow normalized to 60 Hz

🔹 Temp Viewer focuses on downhole thermocouple temperatures and subcool:
• Profiles along the lateral, compared across dates or wells
• A time × position heatmap that shows steam breakthrough or cold spots developing along the wellbore
• Automatic cleanup of dead sensors and physically impossible subcool values

Each viewer is a single offline HTML file. Open it in a browser and load an Excel workbook. There's no installation or server, and the data stays on your machine.

I built both tools through vibe coding: I described what I wanted to see as an engineer, then iterated on the result in plain language until each view answered a real monitoring question. That approach also makes the tools easy to adapt. They're built around SAGD workflows, but they can be reshaped at any time to fit each field's characteristics and each team's needs.

All screenshots use a fully synthetic demo dataset of 14 fictional wells, not real field data.

👉 Explore the live demo and share your feedback: https://sagd-prod-monitoring.vercel.app

How do you monitor well performance today, and which views would you add?

#SAGD #ProductionEngineering #OilAndGas #PetroleumEngineering #DataVisualization #ArtificialLift #WellPerformance #VibeCoding

---
Recommended image upload order (LinkedIn shows the first four in the preview grid):
1. 01_prod_viewer_overview.png: Production Viewer, multiple views of well performance in one workspace
2. 02_temp_subcool_profile_heatmap.png: subcool profiles across dates and a time × position heatmap (steam breakthrough at point 3 before a redrill)
3. 03_esp_crossplot_by_date.png: emulsion rate vs. ESP frequency colored by date, showing pump degradation
4. 04_redrill_before_after_bars.png: mean oil rate for the last-normal, just-before and after-redrill windows
5. 05_production_performance.png: oil rate, water cut and SOR in stacked tracks with shut-in shading
6. 06_steam_breakthrough_watch.png: minimum subcool, wellhead temperature and maximum downhole temperature together
7. 07_multiwell_esp_screening.png: ESP flow at 60 Hz compared across producers
8. 08_oil_vs_days_since_redrill.png: oil-rate histories aligned on redrill day
9. 09_field_oil_rate_bars.png: mean oil rate by well, January–June 2026
10. 10_temp_compare_wells.png: temperature profiles across six wells on the same date
