# Đặc tả giao diện & nhận diện thương hiệu

Đọc khi sửa bố cục, bộ lọc, biểu đồ, màu sắc. File: `assets/tool-template/src/index.html`, `styles.css`, phần render trong `app.js`.

## Bố cục (từ trên xuống)
1. **Thanh dính** (hiện khi cuộn qua bộ lọc): logo, số phiếu đang hiển thị, nút "Bộ lọc ↑", "Xuất Excel".
2. **Banner**: logo NTS trong khung trắng bo góc (trên cùng bên trái) · eyebrow "NTS HANOI CORP. · PHÒNG QUẢN LÝ SẢN PHẨM" · tiêu đề "Tổng hợp Đăng ký dự án (Dealreg)" · dòng phụ (tên file, số phiếu, khoảng ngày tạo, cập nhật mới nhất, ngày mốc) · nút "Nạp file Excel", "Xuất Excel ▾" (theo bộ lọc / toàn bộ).
3. **Vùng kéo-thả** (khi chưa có dữ liệu; kéo-thả được cả trang) → sau khi nạp thu thành thanh "Dữ liệu nguồn" (chip file, nạp thêm, xoá dữ liệu).
4. **Card Bộ lọc & tìm kiếm** — sắp xếp theo mức dùng thường xuyên:
   - Hàng 1: Ô tìm nhanh (rộng) | Khoảng thời gian [trường: Ngày tạo / Cập nhật cuối / Timeline DA] [từ] [đến] [chọn nhanh: 7/30 ngày, tháng này/trước, quý này/trước/sau, năm nay/sau — tính theo ngày mốc].
   - Hàng 2 (6 ô combobox nhiều lựa chọn): Sale · Người tạo · PM phụ trách · Hãng · Tiến độ xử lý · Trạng thái.
   - Hàng chip "Lọc nhanh" có số đếm: Đang xử lý, Chờ hãng check, Có cờ nhắc, Cờ ≥ Mức 2, Timeline ≤30 ngày, Đã qua timeline–chưa xong, Quá hạn SLA, Trùng EU+Hãng, Approved, Rejected.
   - "Lọc nâng cao & thiết lập cờ nhắc" (thu gọn): Cờ nhắc · Kết quả DR · Tình trạng Timeline · Reseller · Quá hạn SLA · Cảnh báo dữ liệu · Người theo dõi · Tháng tạo; Ngày mốc; Phạm vi cờ; chú giải mức cờ.
   - Thanh chip bộ lọc đang áp dụng (✕ từng cái, "Xoá tất cả") + "Đang hiển thị X / Y phiếu".
5. **Thanh cảnh báo nhắc update**: số phiếu >1 tháng, phiếu lâu nhất, chip theo mức (bấm để lọc), "Xem danh sách", "✉ Soạn nội dung nhắc" (modal: nhóm theo PM/Sale/Hãng, từ mức, chỉ đang xử lý; sao chép / tải .txt).
6. **2 panel donut**: (a) Kết quả DR + "Phiếu đang xử lý nằm ở bước"; (b) Cờ nhắc + Timeline dự án. Legend bấm để lọc.
7. **8 KPI**: Tổng phiếu (số sale/hãng) · Đang xử lý · Approved · Rejected · Tỷ lệ duyệt · Cần nhắc update · Quá hạn SLA · TG duyệt trung vị. KPI có bộ lọc thì bấm được.
8. **Insight tự động** (2 cột): tổng quan & tỷ lệ duyệt, top hãng/sale/PM, PM tồn đọng, phiếu lâu chưa cập nhật nhất, hãng Rejected cao nhất (≥5 phiếu có kết quả), timeline, trùng/xung đột, tháng cao điểm, thời gian xử lý, lệch tên hãng.
9. **Biểu đồ (Chart.js, cập nhật theo bộ lọc, bấm để lọc)**: Sale (cột chồng theo kết quả, top 15) · Hãng (thanh ngang) · PM (cột chồng theo cờ) · Tiến độ (thanh ngang theo thứ tự luồng) · Xu hướng tạo mới/Approved/Rejected/Đang xử lý (tuần/tháng/quý) · Người tạo · Pipeline theo quý Timeline (không tính Rejected; bấm → lọc khoảng timeline) · Top Reseller.
10. **Bảng kết quả**: cột Cờ nhắc, ID, Tên phiếu (+ dự án), End user, Hãng (⚠ nếu lệch), Sale, PM, Tiến độ, Trạng thái, Timeline DA, Cập nhật cuối; ẩn/hiện thêm Người tạo, Reseller, Dự án, Ngày tạo, Hạn SLA, Người theo dõi, Note. Bấm tiêu đề để sắp xếp, phân trang 25/50/100/200, highlight từ khoá.
11. **Chi tiết phiếu** (drawer bên phải khi bấm dòng): header (#ID, hãng, tên, badge kết quả/tiến độ/cờ/số ngày) · stepper tiến độ · 4 mốc thời gian (tạo, cập nhật cuối, timeline, SLA) · cảnh báo · Khách hàng & dự án (EU, email, dự án, website, reseller, hãng, thông tin KH, BOM) · Phụ trách (Sale, Người tạo, PM, Người theo dõi) · Note & CRM · Mô tả · Tệp đính kèm (link) · Phiếu khác cùng End user · Thông tin hệ thống. Phím ←/→ chuyển phiếu, Esc đóng, nút sao chép tóm tắt.
12. **Chân trang**: logo + "Dữ liệu được tổng hợp nội bộ từ phòng Quản lý sản phẩm - Công ty Nam Trường Sơn Hà Nội" + phiên bản.

## Tìm kiếm & combobox
- Tìm nhanh: không dấu, tách token (mọi token phải khớp) trên ID, ID khối, tên, EU, email, dự án, thông tin KH, note, mô tả, BOM, website, reseller, hãng, sale, người tạo, PM, tiến độ, trạng thái, link CRM. Gợi ý thả xuống: nhóm "Lọc theo người/hãng/đối tác" (bấm → thêm bộ lọc) + "Phiếu khớp" xếp hạng (ID trùng > ID bắt đầu > tên > EU > dự án); ↑/↓/Enter.
- Combobox: gõ trực tiếp để lọc danh sách (không dấu, khớp cả họ tên & @username), checkbox nhiều lựa chọn, số đếm theo **facet** (đếm với mọi bộ lọc khác trừ chính nó), mục 0 phiếu làm mờ, "Chọn các mục đang hiện"/"Bỏ chọn", Backspace xoá tag cuối.

## Nhận diện NTS Hà Nội (tham chiếu ntshanoi.com.vn)
| Token | Màu | Dùng cho |
|---|---|---|
| --brand-600 | #0B5CB5 | xanh dương chủ đạo (thanh menu website): banner, nút chính, header bảng |
| --brand-900/800/700 | #062F63 / #073F80 / #0A4F9E | chiều sâu gradient, chữ tiêu đề |
| --brand-500/400 | #1470D0 / #2A8AE3 | điểm sáng gradient |
| --orange-500 | #F26522 | cam nhấn (mục đang chọn / gạch chân tiêu đề): viền dưới banner, gạch chân tiêu đề card, bước hiện tại, nút "Soạn nội dung nhắc" |
| --cyan-500 / --mint-500 | #10B7DE / #3FD3B6 | màu phụ (icon dịch vụ website): panel donut, biểu đồ Người tạo |
- Màu ngữ nghĩa giữ cố định: Approved xanh lá #1FA55B, Rejected đỏ #E5484D, Đang xử lý vàng cam #F59E0B; cờ M1 #F2B705, M2 #F07C1B, M3 #E03A3E, M4 #8B1538.
- Chiều sâu: gradient nhiều điểm + radial highlight, hoạ tiết sọc chéo mờ trên banner, bóng nhiều lớp (`--sh-1/2/3`), card bo 16px.
- Font hệ thống hỗ trợ tiếng Việt (Segoe UI…) — không dùng Google Fonts vì phải chạy offline.
- Responsive: KPI 8→4→2 cột, bộ lọc 6→3→1 cột, biểu đồ 1 cột dưới 1100px.
