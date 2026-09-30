#!/usr/bin/env python3
"""Đóng gói Tool tổng hợp Dealreg thành MỘT file HTML chạy offline.

Nhúng trực tiếp: CSS, JS ứng dụng, thư viện (xlsx-js-style, Chart.js,
chartjs-plugin-datalabels) và logo (base64). Không cần internet khi chạy.

    python3 dealreg/build.py
"""
import base64
import pathlib

VERSION = "1.0.0"
ROOT = pathlib.Path(__file__).resolve().parent
OUT = ROOT.parent / "Tool tổng hợp Dealreg.html"
# Bản trùng nội dung cho GitHub Pages (URL gọn: https://<user>.github.io/<repo>/)
PAGES = ROOT.parent / "index.html"


def read(p):
    return (ROOT / p).read_text(encoding="utf-8")


def safe_js(s):
    # tránh đóng thẻ <script> sớm khi nhúng inline
    return s.replace("</script", "<\\/script")


def main():
    logo = "data:image/png;base64," + base64.b64encode((ROOT / "assets/logo-nts.png").read_bytes()).decode()
    html = read("src/index.html")
    parts = {
        "/*__CSS__*/": read("src/styles.css"),
        "/*__XLSX__*/": safe_js(read("vendor/xlsx-js-style.min.js")),
        "/*__CHARTJS__*/": safe_js(read("vendor/chart.umd.min.js")),
        "/*__DATALABELS__*/": safe_js(read("vendor/chartjs-plugin-datalabels.min.js")),
        "/*__APP__*/": safe_js(read("src/app.js").replace("__VERSION__", VERSION)),
    }
    for k, v in parts.items():
        assert k in html, k
        html = html.replace(k, v, 1)
    html = html.replace("__LOGO__", logo).replace("__VERSION__", VERSION)
    OUT.write_text(html, encoding="utf-8")
    PAGES.write_text(html, encoding="utf-8")
    print(f"OK -> {OUT} ({OUT.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
