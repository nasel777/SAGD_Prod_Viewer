"""Assemble the public demo site in site/ (deployed to Vercel). Only synthetic demo data goes in.

    python3 build.py && python3 tools/build_site.py && vercel deploy site --prod
"""
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / "site"
FILES = {
    "index.html": ROOT / "tools" / "site_index.html",
    "SAGD_Prod_Viewer.html": ROOT / "SAGD_Prod_Viewer.html",
    "SAGD_Temp_Viewer.html": ROOT / "SAGD_Temp_Viewer.html",
    "SAGD_Demo_Data.xlsx": ROOT / "demo" / "SAGD_Demo_Data.xlsx",
    "layout_demo_showcase.json": ROOT / "demo" / "layout_demo_showcase.json",
    "img/prod_overview.png": ROOT / "demo" / "linkedin" / "01_prod_viewer_overview.png",
    "img/temp_profile.png": ROOT / "demo" / "linkedin" / "02_temp_subcool_profile_heatmap.png",
}

for name in list(SITE.glob("*")) if SITE.exists() else []:
    if name.name != ".vercel":  # keep the Vercel project link
        shutil.rmtree(name) if name.is_dir() else name.unlink()
for dest, src in FILES.items():
    out = SITE / dest
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, out)
    print(f"{dest} ({out.stat().st_size / 1048576:.1f} MB)")
