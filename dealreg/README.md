# Tool tổng hợp Dealreg — NTS Hanoi Corp.

Sản phẩm: **`Tool tổng hợp Dealreg.html`** (ở thư mục gốc repo). Đây là một file HTML duy nhất, mở bằng Chrome/Edge là dùng được. Tool chạy offline 100%, không gửi dữ liệu đi đâu.

## Cách dùng
1. Mở file HTML, sau đó kéo-thả hoặc chọn file Excel xuất từ Base Service (báo cáo "service-tickets"). Có thể nạp nhiều file cùng lúc; phiếu trùng ID sẽ được gộp, giữ bản cập nhật mới nhất.
2. Lọc / tìm kiếm. Kết quả, biểu đồ, KPI và insight cập nhật ngay theo bộ lọc.
3. Bấm **Xuất Excel** để tải `Tổng hợp Dealreg Tháng <m>.<yyyy>.xlsx` (tháng lấy theo *ngày mốc*).

## Sửa mã nguồn
- `src/index.html`, `src/styles.css`, `src/app.js`: giao diện và logic.
- `vendor/`: thư viện nhúng (xlsx-js-style 1.2.0 – Apache-2.0, Chart.js 4 – MIT, chartjs-plugin-datalabels 2 – MIT).
- `assets/logo-nts.png`: logo.
- Build lại: `python3 dealreg/build.py`.

Chuẩn hoá tên hãng (alias lỗi chính tả) nằm ở `VENDOR_ALIAS` / `VENDOR_CANON` trong `src/app.js`. Ngưỡng cờ nhắc nằm ở hàm `derive()`.
