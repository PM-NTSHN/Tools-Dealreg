# File Excel đầu ra

Hàm `exportExcel()` trong `app.js`. Thư viện **xlsx-js-style** (SheetJS có style). Tên file: `Tổng hợp Dealreg Tháng <m>.<yyyy>.xlsx` theo tháng của **ngày mốc** (VD "Tổng hợp Dealreg Tháng 9.2026.xlsx"). Mặc định xuất **theo bộ lọc**; menu ▾ cho "Xuất toàn bộ dữ liệu".

| # | Sheet | Nội dung | Cố định |
|---|---|---|---|
| 1 | Tổng quan | Tiêu đề tháng, công ty, ngày mốc, thời điểm xuất, nguồn file, **phạm vi bộ lọc**; bảng Chỉ tiêu chính (tổng, đang xử lý, approved, rejected kèm %, tỷ lệ duyệt, số sale/hãng/EU, cần nhắc, quá hạn SLA, qua timeline, trùng EU+Hãng, TG duyệt trung vị); Theo tiến độ; Theo cờ nhắc (tô màu); Theo timeline; Top 10 Hãng; Top 10 Sale; dòng chân trang | ẩn lưới |
| 2 | Danh sách Dealreg | 33 cột: STT, ID, Cờ nhắc (tô màu), Số ngày chưa cập nhật, Tên phiếu, Hãng, End user, Tên dự án, Reseller, BOM, Timeline, Quý timeline, Tình trạng timeline, Sale, Người tạo, PM, Tiến độ, Trạng thái, Kết quả DR (tô màu), Ngày tạo, Cập nhật cuối, Tuổi phiếu, Hạn SLA, Quá hạn SLA, Cảnh báo dữ liệu, Note, Link CRM (hyperlink), Email EU, Website, Thông tin KH, Người theo dõi, Mô tả, Tệp đính kèm | dòng 1 + 5 cột |
| 3 | Cần nhắc update | Phiếu cờ ≥ Mức 1, sắp theo mức rồi số ngày | dòng 1 + 6 cột |
| 4–6 | Theo Sale / Theo Hãng / Theo PM | Tổng, Đang xử lý, Approved, Rejected, Tỷ lệ duyệt, Có cờ nhắc, M1–M4, Qua timeline chưa xong, Cập nhật gần nhất + dòng TỔNG | dòng 4 + cột 1 |
| 7 | Theo tháng & quý | Xu hướng theo tháng tạo; Pipeline theo quý timeline (không Rejected) | |
| 8 | Kiểm tra dữ liệu | Phiếu có cảnh báo (conflict lên đầu), hãng ở tên phiếu, danh sách phiếu trùng | dòng 1 + 4 cột |

## Kỹ thuật cần nhớ
- Ngày ghi thành số serial Excel tự tính theo giờ địa phương (`serial()`), đặt cả `cell.z` và `s.numFmt` (`dd/mm/yyyy hh:mm`, `dd/mm/yyyy`, `0.0%`).
- Style: header nền #0B5CB5 chữ trắng; thanh mục ở Tổng quan nền cam #F26522; viền mảnh #CFD9E0; `!autofilter`, `!cols`, `!merges`, `!rows`.
- **Freeze panes / ẩn lưới**: SheetJS bản cộng đồng không ghi được → hậu xử lý `patchViews()`: đọc zip bằng `XLSX.CFB`, thay `<sheetView workbookViewId="0"/>` trong `xl/worksheets/sheetN.xml` bằng thẻ có `<pane … state="frozen"/>`, ghi lại bằng `XLSX.CFB.write(…, {fileType:'zip'})`. Thứ tự mảng views = thứ tự sheet.
- Tải file: Blob + `<a download>`. (Trong container headless thiếu locale, tên file có dấu bị đổi thành "download" — chỉ là lỗi môi trường test, chạy với `LANG=C.UTF-8`.)
- Kiểm tra lại bằng openpyxl: số sheet, `freeze_panes`, `auto_filter.ref`, `number_format` của cột ngày, màu fill cột Cờ nhắc.
