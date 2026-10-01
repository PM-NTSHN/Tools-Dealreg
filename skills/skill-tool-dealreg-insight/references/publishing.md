# Chia sẻ / public tool thành trang web

Tool là file tĩnh; dữ liệu Excel chỉ xử lý trong trình duyệt người dùng → host ở đâu cũng không lộ dữ liệu (chỉ lộ giao diện & mã nguồn).

## GitHub Pages (miễn phí, repo public)
1. Build kèm `--pages` để có `index.html` ở gốc repo (URL gọn `https://<owner>.github.io/<repo>/`); thêm file rỗng `.nojekyll`.
2. Commit + push.
3. Người dùng mở `https://github.com/<owner>/<repo>/settings/pages` → Source: *Deploy from a branch* → chọn nhánh chứa `index.html`, thư mục `/ (root)` → Save → đợi 1–2 phút.
4. Mỗi lần push bản mới, trang tự cập nhật sau ~1 phút.
- Repo private: Pages cần gói GitHub trả phí.
- Repo hiện tại của NTS PM: `PM-NTSHN/Tools` → https://pm-ntshn.github.io/Tools/

## Chỉ cho nội bộ truy cập
- **Cloudflare Pages + Cloudflare Access** (miễn phí ≤50 user): đăng nhập bằng email công ty.
- **Web server nội bộ** (IIS/nginx/Apache): chép `index.html` vào thư mục web, VD `http://intranet/dealreg/`; chỉ LAN/VPN.
- SharePoint/OneDrive/Google Drive **không** hiển thị HTML như trang web (chỉ tải về).

## Gửi file trực tiếp
Gửi `Tool tổng hợp Dealreg.html` qua email/chat; người nhận mở bằng Chrome/Edge.
