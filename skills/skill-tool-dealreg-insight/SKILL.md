---
name: skill-tool-dealreg-insight
description: "skill tool Dealreg insight — quy trình & bộ mã nguồn dựng \"Tool tổng hợp Dealreg\": một file HTML offline duy nhất (logo + màu nhận diện NTS Hà Nội) nạp file Excel raw đăng ký dự án (Deal Registration / DR) xuất từ Base Service (service-tickets.report*.xlsx), lọc đa điều kiện theo Sale, người tạo, PM phụ trách (Người thực hiện), hãng, tiến độ (Tên khối), trạng thái, timeline; tìm nhanh có gợi ý; cờ nhắc theo thời gian chưa cập nhật; biểu đồ động; chi tiết phiếu; soạn nhắc; xuất Excel \"Tổng hợp Dealreg Tháng [tháng].[năm].xlsx\" nhiều sheet. Dùng skill này mỗi khi người dùng (PM/presales/sales admin) muốn tổng hợp, phân tích, báo cáo, dashboard dữ liệu dealreg / đăng ký dự án / DR, build lại hoặc sửa / nâng cấp / đổi màu tool Dealreg, thêm bộ lọc / biểu đồ / cột Excel, áp dụng tool cho file raw mới hoặc cấu trúc cột mới, hay public tool thành trang web — kể cả khi họ chỉ nói \"làm tool tổng hợp phiếu đăng ký dự án\", \"dashboard DR theo sale/hãng\", \"tổng hợp file service tickets\" mà không nhắc tên skill."
---

# skill tool Dealreg insight

Tái tạo / bảo trì **Tool tổng hợp Dealreg** của Phòng Quản lý sản phẩm – NTS Hanoi Corp.: một file HTML duy nhất, mở bằng trình duyệt, **không cần internet, không gửi dữ liệu đi đâu**, nạp file Excel raw từ Base Service và cho ra dashboard + file Excel tổng hợp.

Skill đã chứa sẵn bản tool hoàn chỉnh đã được người dùng duyệt (`assets/tool-template/`). Vì vậy việc mặc định là **dùng lại template và chỉ sửa phần cần thiết**, không viết lại từ đầu — viết lại sẽ mất các chi tiết đã được tinh chỉnh (chuẩn hoá Sale, freeze panes Excel, đếm facet, highlight không dấu…).

## Nội dung skill
```
assets/tool-template/
  src/index.html   khung trang + placeholder /*__CSS__*/ /*__XLSX__*/ /*__CHARTJS__*/ /*__DATALABELS__*/ /*__APP__*/ __LOGO__ __VERSION__
  src/styles.css   giao diện, token màu thương hiệu (--brand-*, --orange-*, --cyan-*)
  src/app.js       toàn bộ logic (đọc Excel, chuẩn hoá, lọc, biểu đồ, chi tiết, nhắc, xuất Excel)
  vendor/          xlsx-js-style 1.2.0 (Apache-2.0), Chart.js 4 (MIT), chartjs-plugin-datalabels 2 (MIT)
  assets/logo-nts.png
scripts/build.py          gộp tất cả thành "Tool tổng hợp Dealreg.html" (+ index.html với --pages)
scripts/profile_excel.py  soi file raw: cột khớp/thiếu, phân bố giá trị, dấu hiệu cần chuẩn hoá
scripts/smoke_test.js     kiểm thử Playwright: nạp → tìm → lọc → chi tiết → xuất Excel, bắt lỗi & request mạng
references/data-model.md  cột raw, luồng Tên khối, quy tắc chuẩn hoá, cờ nhắc, cảnh báo dữ liệu, số liệu đối chiếu
references/ui-spec.md     bố cục, bộ lọc, biểu đồ, drawer chi tiết, bảng màu NTS
references/excel-output.md 8 sheet đầu ra, kỹ thuật style/freeze panes
references/publishing.md  GitHub Pages / nội bộ / gửi file
```

## Quy trình

### 1. Xác định yêu cầu
Phân loại yêu cầu để biết cần đọc gì:
- **Chỉ cần tool** (người dùng mới, máy mới) → bước 4 (build) rồi giao file.
- **Có file raw mới / cột thay đổi** → bước 2.
- **Sửa tính năng / thêm bộ lọc, biểu đồ, cột Excel** → đọc `references/ui-spec.md` hoặc `references/excel-output.md`, sửa trong bản sao template (bước 3).
- **Đổi màu / logo / nhận diện** → `references/ui-spec.md` (mục Nhận diện), sửa token `:root` trong `styles.css`, màu cứng trong `app.js` (biểu đồ `grad('#…')`, `ST`/`FILL` của Excel).
- **Public lên web** → `references/publishing.md`.

Nếu yêu cầu có chỗ mơ hồ (ảnh tham chiếu bị thiếu, quy tắc mâu thuẫn như hai ngưỡng cùng "mức 3", câu bị cụt), hỏi lại trước khi làm — người dùng này đã nói rõ muốn được hỏi khi chưa rõ. Khi thiếu ảnh logo/màu, đề xuất làm trước với mặc định hoặc chờ, để người dùng chọn.

### 2. Soi dữ liệu raw (khi có file mới)
```bash
pip install openpyxl -q
python3 scripts/profile_excel.py "<file raw>.xlsx"
```
So với `references/data-model.md`: cột nào thiếu/đổi tên, giá trị Tên khối mới, Sale có lẫn họ tên/@username không, hãng có cách viết mới không. Tên cột mới → thêm vào `COLMAP` trong `app.js`; lỗi chính tả hãng mới → `VENDOR_ALIAS`/`VENDOR_CANON`; khối mới tự xử lý (chèn trước Approved) nhưng nên thêm vào `FLOW` đúng vị trí.

### 3. Sửa mã (nếu cần)
Sao chép template ra thư mục làm việc rồi sửa bản sao (thư mục skill có thể chỉ đọc):
```bash
cp -r <skill>/assets/tool-template ./dealreg-src
```
Nguyên tắc khi sửa:
- Giữ **một file HTML, không request mạng**: thư viện mới phải tải về và nhúng qua placeholder trong `index.html` + `build.py`, không dùng CDN/Google Fonts.
- Mọi thứ hiển thị theo `S.view` (đã lọc) để biểu đồ/KPI/insight/bảng/Excel luôn khớp bộ lọc. Thêm bộ lọc mới = thêm một mục vào `FACETS` (`label`, `vals(r)` trả mảng giá trị, tuỳ chọn `lab`, `order`, `dot`) + một `<div class="field" data-facet="…">` trong `index.html`; combobox, chip, đếm facet, chip đang áp dụng tự có.
- Trường phụ thuộc ngày mốc tính trong `derive()`; trường tĩnh trong `rebuild()`.
- Thêm cột bảng → `COLS`; thêm sheet/cột Excel → `exportExcel()` (nhớ cập nhật mảng `patchViews` theo thứ tự sheet).
- Văn bản giao diện bằng tiếng Việt có dấu; chân trang giữ nguyên câu "Dữ liệu được tổng hợp nội bộ từ phòng Quản lý sản phẩm - Công ty Nam Trường Sơn Hà Nội".

### 4. Build
```bash
python3 <skill>/scripts/build.py --out-dir <nơi giao file>                       # dùng template gốc
python3 <skill>/scripts/build.py --template ./dealreg-src --out-dir . --version 1.1.0 [--pages] [--logo logo.png]
```
Kết quả: `Tool tổng hợp Dealreg.html` (~850 KB). `--pages` ghi thêm `index.html` để host web.

### 5. Kiểm thử trước khi giao
Kiểm thử là bắt buộc vì tool chạy ở máy người khác — lỗi JS nhỏ là trang trắng.
```bash
node --check ./dealreg-src/src/app.js          # nếu đã sửa app.js
LANG=C.UTF-8 node <skill>/scripts/smoke_test.js "Tool tổng hợp Dealreg.html" "<raw>.xlsx" <out_dir> "$(npm root -g)/playwright"
```
(Trong môi trường có Chromium cài sẵn ở chỗ khác, đặt `CHROMIUM_PATH`.) Script phải in "Không lỗi, không request mạng" và xuất được file `Tổng hợp Dealreg Tháng …xlsx`. Sau đó:
- Xem ảnh `full.png` / `detail.png` (đọc bằng công cụ xem ảnh) để soát bố cục, màu, chữ bị tràn.
- Mở file Excel bằng openpyxl kiểm tra số sheet, freeze panes, định dạng ngày, số liệu khớp KPI trên giao diện.
- Với file mẫu gốc, đối chiếu số liệu ở cuối `references/data-model.md`.

### 6. Giao & báo cáo
- Gửi file HTML (và file Excel mẫu nếu hữu ích); nếu làm trong repo thì commit + push.
- Tóm tắt cho người dùng bằng tiếng Việt, ngắn gọn: đã làm gì, đã kiểm thử ra sao, các giả định cần họ xác nhận (VD Mức 4 cho >1 năm, phạm vi áp dụng cờ mặc định "Tất cả phiếu").

## Những điểm đã được người dùng chốt (giữ nguyên trừ khi họ yêu cầu khác)
- Tên file đầu ra: `Tổng hợp Dealreg Tháng <m>.<yyyy>.xlsx`, tháng theo ngày mốc.
- PM phụ trách = cột "Người thực hiện"; Tiến độ xử lý = "Tên khối"; Trạng thái = "Trạng thái".
- Sale tự động gộp họ tên ↔ @username; hãng gộp hoa/thường & lỗi chính tả.
- Cờ nhắc: >1 tháng M1, >3 tháng M2, >6 tháng M3, >1 năm M4 (tách từ yêu cầu gốc ghi hai lần "mức 3").
- Bộ lọc xếp theo chuẩn dashboard: tìm nhanh + thời gian → 6 bộ lọc chính → lọc nhanh → nâng cao thu gọn → chip đang áp dụng.
- Màu: xanh dương NTS #0B5CB5 chủ đạo + cam #F26522 nhấn (theo ntshanoi.com.vn), có chiều sâu; logo NTS góc trên trái.
