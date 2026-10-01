#!/usr/bin/env python3
"""Đóng gói Tool tổng hợp Dealreg thành MỘT file HTML chạy offline.

Nhúng CSS, JS ứng dụng, thư viện (xlsx-js-style, Chart.js, chartjs-plugin-datalabels)
và logo (base64) vào một file duy nhất.

Ví dụ:
    python3 scripts/build.py --out-dir /path/to/output
    python3 scripts/build.py --template /path/to/edited-template --out-dir . --pages --version 1.1.0

--template : thư mục có src/, vendor/, assets/logo-nts.png (mặc định: ../assets/tool-template)
--out-dir  : nơi ghi "Tool tổng hợp Dealreg.html"
--pages    : ghi thêm index.html (cho GitHub Pages / web server)
"""
import argparse
import base64
import pathlib

HERE = pathlib.Path(__file__).resolve().parent
DEFAULT_TEMPLATE = HERE.parent / "assets" / "tool-template"
OUT_NAME = "Tool tổng hợp Dealreg.html"


def safe_js(s):
    # tránh đóng thẻ <script> sớm khi nhúng inline
    return s.replace("</script", "<\\/script")


def build(template: pathlib.Path, out_dir: pathlib.Path, version: str, pages: bool, logo: pathlib.Path | None):
    rd = lambda p: (template / p).read_text(encoding="utf-8")
    logo_path = logo or (template / "assets" / "logo-nts.png")
    mime = "image/svg+xml" if logo_path.suffix.lower() == ".svg" else "image/png"
    logo_uri = f"data:{mime};base64," + base64.b64encode(logo_path.read_bytes()).decode()
    html = rd("src/index.html")
    parts = {
        "/*__CSS__*/": rd("src/styles.css"),
        "/*__XLSX__*/": safe_js(rd("vendor/xlsx-js-style.min.js")),
        "/*__CHARTJS__*/": safe_js(rd("vendor/chart.umd.min.js")),
        "/*__DATALABELS__*/": safe_js(rd("vendor/chartjs-plugin-datalabels.min.js")),
        "/*__APP__*/": safe_js(rd("src/app.js").replace("__VERSION__", version)),
    }
    for k, v in parts.items():
        assert k in html, f"Thiếu placeholder {k} trong src/index.html"
        html = html.replace(k, v, 1)
    html = html.replace("__LOGO__", logo_uri).replace("__VERSION__", version)
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / OUT_NAME
    out.write_text(html, encoding="utf-8")
    print(f"OK -> {out} ({out.stat().st_size / 1024:.0f} KB)")
    if pages:
        (out_dir / "index.html").write_text(html, encoding="utf-8")
        print(f"OK -> {out_dir / 'index.html'}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", type=pathlib.Path, default=DEFAULT_TEMPLATE)
    ap.add_argument("--out-dir", type=pathlib.Path, default=pathlib.Path.cwd())
    ap.add_argument("--version", default="1.0.0")
    ap.add_argument("--pages", action="store_true")
    ap.add_argument("--logo", type=pathlib.Path, help="logo thay thế (png/svg)")
    a = ap.parse_args()
    build(a.template, a.out_dir, a.version, a.pages, a.logo)


if __name__ == "__main__":
    main()
