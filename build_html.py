"""
build_html.py — Injects dashboard_data.json into template.html to produce
dashboard.html, the file to publish (or republish) as the Artifact.

Usage:  python build_data.py && python build_html.py
"""

import datetime
import json
import os

HERE = os.path.dirname(__file__)
DATA_PATH = os.path.join(HERE, "dashboard_data.json")
TEMPLATE_PATH = os.path.join(HERE, "template.html")
OUT_PATH = os.path.join(HERE, "index.html")


def main():
    with open(DATA_PATH, encoding="utf-8") as f:
        data_raw = f.read()
    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        html = f.read()

    stamp = datetime.datetime.now().strftime("%d %b %Y, %H:%M")
    # Defensive: a literal "</script" anywhere inside the embedded JSON (e.g.
    # in a scraped headline) would prematurely close the <script> block when
    # the browser's HTML tokenizer scans it, regardless of JS string context.
    data_raw = data_raw.replace("</script", "<\\/script")
    html = html.replace("__DASHBOARD_DATA__", data_raw)
    html = html.replace("__BUILD_STAMP__", stamp)

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Wrote {OUT_PATH} ({len(html):,} bytes), stamped {stamp}")


if __name__ == "__main__":
    main()
