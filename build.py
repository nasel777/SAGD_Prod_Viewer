"""Bundle src/ into single offline HTML files: SAGD_Prod_Viewer.html and SAGD_Profile_Viewer.html"""
from pathlib import Path

ROOT = Path(__file__).parent
SRC = ROOT / "src"


def read(p):
    return (SRC / p).read_text(encoding="utf-8")


def bundle(page, parts, out_name):
    html = read(page)
    for key, text in parts.items():
        if "</script" in text.lower():
            raise SystemExit(f"{key}: contains </script, cannot inline")
        if key not in html:
            raise SystemExit(f"{page}: placeholder {key} not found")
        html = html.replace(key, text, 1)
    out = ROOT / out_name
    out.write_text(html, encoding="utf-8")
    print(f"wrote {out} ({out.stat().st_size / 1048576:.1f} MB)")


common = {
    "/*@@CSS@@*/": read("app.css"),
    "/*@@READER@@*/": read("reader.js"),
    "/*@@XLSX@@*/": read("lib/xlsx.full.min.js"),
    "/*@@PLOTLY@@*/": read("lib/plotly.min.js"),
}
bundle("app.html", {**common, "/*@@APP@@*/": read("app.js")}, "SAGD_Prod_Viewer.html")
bundle("profile.html", {**common, "/*@@PROFILE_CSS@@*/": read("profile.css"), "/*@@PROFILE@@*/": read("profile.js")},
       "SAGD_Profile_Viewer.html")
