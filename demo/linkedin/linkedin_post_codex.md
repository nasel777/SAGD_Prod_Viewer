### 1. Post text

As a SAGD production engineer, I built two tools to make well performance data easier to explore: **SAGD_Prod_Viewer** and **SAGD_Temp_Viewer**.

Time-series plots are my starting point. But checking the same data from several angles helps me challenge an interpretation and reach more reliable conclusions.

**SAGD_Prod_Viewer** brings together oil, water and steam rates, ESP parameters, and well operating indicators. It supports:
• Time-series plots for individual wells and multiwell comparisons
• Crossplots between selected parameters, colored by well or date
• Statistical bar charts and box plots
• Comparisons before and after redrill, with trends aligned by days since redrill

**SAGD_Temp_Viewer** focuses on downhole thermocouple temperatures and subcool. It lets users compare profiles across dates or wells, step through time, and explore heatmaps showing how conditions change along the well.

Each viewer is a **single offline HTML file**. Open it in a browser and load an Excel workbook—no installation or server required. Data stays on the user’s machine.

I developed these tools with **Claude Code**, combining my production engineering experience with AI-assisted coding. They are built around SAGD workflows, and the code can be modified to suit each field’s characteristics and users’ needs.

All screenshots use a **fully synthetic demo dataset of 14 fictional wells**, not real field data.

#SAGD #ProductionEngineering #OilAndGas #DataVisualization #AIAssistedDevelopment

Explore the live demo and share your feedback: [demo link]

### 2. Recommended image upload order

| Order | Image | Caption |
|---|---|---|
| 1 | `01_prod_viewer_overview.png` | Production Viewer overview: multiple views of well performance in one workspace. |
| 2 | `02_temp_subcool_profile_heatmap.png` | Temp Viewer: compare subcool profiles across dates alongside a time–position heatmap. |
| 3 | `05_production_performance.png` | Track oil rate, water cut and SOR together using stacked time-series plots. |
| 4 | `03_esp_crossplot_by_date.png` | Explore emulsion rate versus ESP frequency, with points colored by date. |
| 5 | `07_multiwell_esp_screening.png` | Compare ESP flow at 60 Hz across selected producers. |
| 6 | `06_steam_breakthrough_watch.png` | Review minimum subcool, wellhead temperature and maximum downhole temperature together. |
| 7 | `08_oil_vs_days_since_redrill.png` | Align oil-rate histories around redrill day to compare well trajectories. |
| 8 | `04_redrill_before_after_bars.png` | Compare mean oil rates for last-normal, just-before and after-redrill windows. |
| 9 | `09_field_oil_rate_bars.png` | Compare mean oil rates by well for January–June 2026. |
| 10 | `10_temp_compare_wells.png` | Compare thermocouple temperature profiles across six wells on the same date. |