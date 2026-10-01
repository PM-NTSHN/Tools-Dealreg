# Logic dữ liệu Dealreg (Base Service – "Official_Yêu cầu Đăng ký dự án")

Đọc file này khi cần sửa cách đọc / chuẩn hoá / tính toán dữ liệu, hoặc khi file raw đổi cấu trúc.
Code tương ứng nằm trong `assets/tool-template/src/app.js` (các khối `COLMAP`, `rebuild()`, `derive()`).

## 1. File raw đầu vào
- Xuất từ Base Service: `service-tickets.report.<hh.mm.dd.mm.yy>.xlsx`, sheet "Service Tickets" (có thể có thêm sheet rỗng "Worksheet 1" → bỏ qua).
- Mỗi dòng = 1 phiếu đăng ký dự án (ID duy nhất). 36 cột. Ngày dạng chuỗi `hh:mm dd/mm/yyyy`; Timeline dạng `dd/mm/yyyy`.
- Tool chọn sheet có nhiều tiêu đề khớp nhất, dò dòng tiêu đề trong 25 dòng đầu, map cột **theo tên** (không dấu, chữ thường) → đổi thứ tự cột không ảnh hưởng.

| Cột raw | Trường | Ghi chú |
|---|---|---|
| No. | no | |
| ID | id | khoá chính, dùng loại trùng khi nạp nhiều file (giữ bản cập nhật mới nhất) |
| Người tạo | creator | username không @ |
| Dịch vụ | service | luôn "Official_Yêu cầu Đăng ký dự án" |
| Tên | name | quy ước `DR_ <End user>_ <Hãng>` |
| Mô tả phiếu | desc | text tự do (EU/Project/BOM/Partner/Timeline) |
| Tệp đính kèm phiếu | attach | `01. file.xlsx - (https://api.base.vn/...)`, có thể nhiều |
| Người thực hiện | pm | `@username` — **PM phụ trách** |
| Người theo dõi | followers | nhiều `@username` cách nhau dấu cách |
| Tạo lúc / Cập nhật lần cuối lúc | created / updated | |
| Sale | sale | **lẫn họ tên ("Vy Công Quý") và "@quyvy"** → phải gộp |
| Tên End user (Bằng tiếng Anh), Email của End user, Thông tin khách hàng, Tên dự án, Reseller, Website Khách hàng, Note, Link Deal CRM | eu, euEmail, custInfo, project, reseller, website, note, crm | |
| Tên hãng | vendor | có thể nhiều hãng "Netscout, Opentext"; viết hoa/thường lẫn lộn, lỗi chính tả |
| Thông tin BOM (...) | bom | |
| Timeline dự án | timeline | ngày dự kiến |
| ID khối của phiếu, Tên khối, Loại khối | blockId, block, blockType | **Tên khối = tiến độ xử lý** |
| Trạng thái | status | "Đang xử lý" / "Hoàn thành" |
| Tạo khối lúc / Cập nhật khối lần cuối lúc | blockCreated / blockUpdated | |
| SLA/Thời hạn, Quá hạn | sla, overdue | Quá hạn = Có/Không/trống |

Cột thiếu → tool cảnh báo nhưng vẫn chạy (bộ lọc liên quan để trống). Bắt buộc: ID, Tên, Trạng thái, Tên khối.

## 2. Luồng tiến độ (Tên khối)
`Tạo yêu cầu đăng ký dự án → Tự động tìm PM hãng → PM tiếp nhận → PM gửi DR đến hãng → Hãng check → Approved | Rejected`
- Approved/Rejected luôn đi kèm Trạng thái "Hoàn thành". Khối lạ (xuất hiện sau này) tự chèn trước Approved.
- **Kết quả DR** (trường suy ra `result`): tên khối chứa "approv"/"duyet" → Approved; "reject"/"tu choi" → Rejected; Trạng thái Hoàn thành khác → "Hoàn thành khác"; còn lại → Đang xử lý.

## 3. Chuẩn hoá
- **Không dấu khi so khớp**: `fold()` bỏ dấu từng ký tự, giữ nguyên độ dài chuỗi để highlight đúng vị trí; đ→d. Chuẩn hoá NFC khi nạp.
- **Người (Sale/Người tạo/PM/Người theo dõi)**: khoá = username chữ thường. Họ tên → khoá theo quy ước Base: *tên cuối + họ đầu, không dấu* ("Vy Công Quý" → `quyvy`, "Dương Văn Thế" → `theduong`). Nếu 2 họ tên khác nhau ra cùng khoá → **không gộp**, giữ riêng theo họ tên. Hiển thị "Họ Tên" + phụ "@username".
- **Hãng**: khoá = slug chữ thường; `VENDOR_ALIAS` sửa lỗi chính tả (optswat→opswat…); `VENDOR_CANON` cho tên hiển thị chuẩn (OPSWAT, NETSCOUT, InfoExpress…). Tách nhiều hãng theo `, ; / &`.
- **Reseller**: khoá slug, nhãn = cách viết phổ biến nhất; tách theo dấu phẩy.
- **Cập nhật cuối** `lastUpdate` = max(Cập nhật lần cuối lúc, Cập nhật khối lần cuối lúc), fallback Tạo lúc.

## 4. Trường phụ thuộc ngày mốc (`derive()`)
Ngày mốc mặc định = hôm nay, người dùng đổi được. Tên file xuất lấy tháng của ngày mốc.
- `days` = số ngày từ lastUpdate đến hết ngày mốc.
- **Cờ nhắc** (theo tháng lịch, không phải 30 ngày cố định): ≤1 tháng = 0 (Bình thường); >1 tháng = Mức 1; >3 tháng = Mức 2; >6 tháng = Mức 3; >1 năm = Mức 4. (Yêu cầu gốc ghi ">6 tháng" và ">1 năm" đều "mức 3" — đã tách Mức 4, đổi trong `derive()` nếu cần.)
- Phạm vi cờ (`S.scope`): `all` (mặc định, đúng yêu cầu) / `notRejected` / `open`; ngoài phạm vi → 'na'.
- **Tình trạng timeline** `tl`: past (đã qua), d30 (≤30 ngày), d90 (31–90), later (>90), none.
- `procDays` = lastUpdate − created cho phiếu đã hoàn thành (dùng cho "TG duyệt trung vị").

## 5. Cảnh báo dữ liệu (`issues`)
| Mã | Ý nghĩa |
|---|---|
| conflict | Cùng End user + Hãng (không tính Rejected) nhưng **khác Sale** → nguy cơ xung đột deal |
| dup | Cùng End user + Hãng đăng ký >1 lần |
| vendorMismatch | Đoạn cuối tên phiếu (sau "_") là một hãng khác cột Tên hãng |
| tlPastOpen | Đang xử lý nhưng đã qua timeline |
| slaOver | Đang xử lý và Quá hạn = Có |
| noSale / noVendor / noTimeline / noEU | Thiếu thông tin |

End user so khớp bằng slug (bỏ dấu, bỏ ký tự đặc biệt). "Phiếu khác cùng End user" hiện trong chi tiết phiếu.

## 6. Kết quả trên file mẫu (30/09/2026, 820 phiếu) — dùng để đối chiếu khi sửa
596 Approved / 86 Rejected / 138 Đang xử lý; tỷ lệ duyệt 87,4%; 16 người tạo; Sale 32 giá trị raw → 22 người sau gộp; 26 hãng sau chuẩn hoá; cờ: 193 OK / 203 M1 / 400 M2 / 24 M3; 113 trùng EU+Hãng, 8 conflict, 8 vendorMismatch, 48 quá hạn SLA, 27 thiếu timeline.
