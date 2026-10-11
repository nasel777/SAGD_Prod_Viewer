"""Assemble the public demo site in site/ (deployed to Vercel). Only synthetic demo data goes in.

    python3 build.py && python3 tools/build_site.py && (cd site && vercel deploy --prod --yes)

The site root serves the Production Viewer. The viewer copies get a `sagd-demo` meta tag, which turns on
the "Open demo data" button and the demo download links; the offline files in the repo root stay as they are.
"""
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / "site"
URL = "https://sagd-prod-monitoring.vercel.app"
DATA, LAYOUT = "SAGD_Demo_Data.xlsx", "layout_demo_showcase.json"

COPY = {
    DATA: ROOT / "demo" / DATA,
    LAYOUT: ROOT / "demo" / LAYOUT,
    "img/prod_overview.png": ROOT / "demo" / "linkedin" / "01_prod_viewer_overview.png",
}
DEMO_META = f'<meta name="sagd-demo" data-xlsx="{DATA}" data-layout="{LAYOUT}">'
SHARE_META = f"""<meta name="description" content="Monitor SAGD well performance with time series, crossplots, statistics and wellbore temperature profiles. Live demo with synthetic data.">
<meta property="og:title" content="SAGD Production &amp; Temp Viewers">
<meta property="og:description" content="Monitor SAGD well performance with time series, crossplots, statistics and wellbore temperature profiles. Live demo with synthetic data.">
<meta property="og:image" content="{URL}/img/prod_overview.png">
<meta property="og:url" content="{URL}/">"""
VIEWERS = {"SAGD_Prod_Viewer.html": DEMO_META + "\n" + SHARE_META, "SAGD_Temp_Viewer.html": DEMO_META}

for p in list(SITE.glob("*")) if SITE.exists() else []:
    if p.name != ".vercel":  # keep the Vercel project link
        shutil.rmtree(p) if p.is_dir() else p.unlink()
SITE.mkdir(exist_ok=True)

for name, meta in VIEWERS.items():
    html = (ROOT / name).read_text(encoding="utf-8")
    anchor = '<meta name="viewport"'
    if anchor not in html:
        raise SystemExit(f"{name}: viewport meta not found")
    (SITE / name).write_text(html.replace(anchor, meta + "\n" + anchor, 1), encoding="utf-8")
for dest, src in COPY.items():
    out = SITE / dest
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, out)
(SITE / "vercel.json").write_text(json.dumps(
    {"rewrites": [{"source": "/", "destination": "/SAGD_Prod_Viewer.html"}]}, indent=2) + "\n")

for p in sorted(SITE.rglob("*")):
    if p.is_file() and ".vercel" not in p.parts:
        print(f"{p.relative_to(SITE)} ({p.stat().st_size / 1048576:.1f} MB)")
