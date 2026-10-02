"""Bundle src/ into a single offline HTML file: SAGD_Prod_Viewer.html"""
from pathlib import Path

ROOT = Path(__file__).parent
SRC = ROOT / "src"


def read(p):
    return (SRC / p).read_text(encoding="utf-8")


html = read("app.html")
parts = {
    "/*@@CSS@@*/": read("app.css"),
    "/*@@APP@@*/": read("app.js"),
    "/*@@XLSX@@*/": read("lib/xlsx.full.min.js"),
    "/*@@PLOTLY@@*/": read("lib/plotly.min.js"),
}
for key, text in parts.items():
    if "</script" in text.lower():
        raise SystemExit(f"{key}: contains </script, cannot inline")
    html = html.replace(key, text, 1)

out = ROOT / "SAGD_Prod_Viewer.html"
out.write_text(html, encoding="utf-8")
print(f"wrote {out} ({out.stat().st_size / 1048576:.1f} MB)")
