#!/usr/bin/env python3
"""Soi nhanh file Excel raw Dealreg (báo cáo service-tickets của Base) trước khi build/sửa tool.

In ra: sheet + dòng tiêu đề được nhận diện, cột nào khớp/thiếu so với COLMAP của tool,
phân bố giá trị các cột quan trọng (Người tạo, Sale, Người thực hiện, Tên hãng, Tên khối,
Trạng thái...), định dạng ngày, và các dấu hiệu cần chuẩn hoá (Sale lẫn @username/họ tên,
hãng viết khác nhau).

    python3 scripts/profile_excel.py "service-tickets.report.xlsx"
Cần: pip install openpyxl
"""
import collections
import re
import sys
import unicodedata

import openpyxl

COLMAP = {  # đồng bộ với COLMAP trong assets/tool-template/src/app.js
    "no": ["no.", "no", "stt"], "id": ["id", "ma phieu"], "creator": ["nguoi tao"], "service": ["dich vu"],
    "name": ["ten", "ten phieu"], "category": ["danh muc phieu"], "desc": ["mo ta phieu", "mo ta"],
    "attach": ["tep dinh kem phieu"], "pm": ["nguoi thuc hien"], "followers": ["nguoi theo doi"],
    "created": ["tao luc"], "updated": ["cap nhat lan cuoi luc"], "sale": ["sale"],
    "eu": ["ten end user (bang tieng anh)", "ten end user", "end user"], "euEmail": ["email cua end user"],
    "custInfo": ["thong tin khach hang"], "project": ["ten du an"], "reseller": ["reseller"], "vendor": ["ten hang"],
    "bom": ["thong tin bom (model, so luong, nam support...)", "thong tin bom"], "timeline": ["timeline du an"],
    "website": ["website khach hang"], "note": ["note", "ghi chu"], "crm": ["link deal crm"],
    "origin": ["link to origin object"], "blockId": ["id khoi cua phieu"], "block": ["ten khoi"],
    "blockType": ["loai khoi"], "status": ["trang thai"], "blockCreated": ["tao khoi luc"],
    "blockUpdated": ["cap nhat khoi lan cuoi luc"], "sla": ["sla/thoi han"], "overdue": ["qua han"],
}


def fold(s):
    s = str(s or "").replace("đ", "d").replace("Đ", "d")
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower().strip()


def main(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    best = None
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for i, r in enumerate(rows[:25]):
            heads = [fold(h) for h in r]
            hits = {k: heads.index(a) for k, alts in COLMAP.items() for a in alts if a in heads}
            if "id" in hits and (not best or len(hits) > len(best[2])):
                best = (ws.title, i, hits, [h for h in r], rows[i + 1:])
    if not best:
        print("Không tìm thấy dòng tiêu đề có cột ID."); return
    title, hi, hits, heads, data = best
    data = [r for r in data if r and r[hits["id"]] not in (None, "")]
    print(f"Sheet: {title} | dòng tiêu đề: {hi + 1} | {len(heads)} cột | {len(data)} phiếu")
    important = ["id", "name", "status", "block", "sale", "vendor", "pm", "creator", "updated", "timeline", "eu"]
    print("Thiếu cột quan trọng:", [k for k in important if k not in hits] or "không")
    print("Thiếu cột phụ:", [k for k in COLMAP if k not in hits and k not in important] or "không")
    print("Cột tool không đọc:", [h for h in heads if h and not any(fold(h) in alts for alts in COLMAP.values())])
    for k in ["creator", "pm", "sale", "vendor", "block", "status", "overdue", "reseller"]:
        if k in hits:
            c = collections.Counter(str(r[hits[k]]).strip() if r[hits[k]] is not None else "(trống)" for r in data)
            print(f"\n== {k} ({len(c)} giá trị):", dict(c.most_common(30)))
    if "sale" in hits:
        vals = [str(r[hits["sale"]]) for r in data if r[hits["sale"]]]
        at = sum(v.startswith("@") for v in vals)
        print(f"\nSale: {at} giá trị dạng @username, {len(vals) - at} dạng họ tên → cần gộp")
    if "vendor" in hits:
        groups = collections.defaultdict(set)
        for r in data:
            for v in re.split(r"\s*[,;/]\s*", str(r[hits["vendor"]] or "")):
                if v: groups[re.sub(r"[^a-z0-9]", "", fold(v))].add(v)
        print("Hãng có nhiều cách viết:", {k: v for k, v in groups.items() if len(v) > 1})
    for k in ["created", "updated", "timeline", "sla"]:
        if k in hits:
            sample = [r[hits[k]] for r in data if r[hits[k]]][:3]
            print(f"Định dạng {k}:", sample)


if __name__ == "__main__":
    main(sys.argv[1])
