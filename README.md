# Kho tài liệu và VBA

Trang web tĩnh chạy trên GitHub Pages để lưu, tìm, xem và tải tài liệu cùng code VBA.

## Cài đặt

1. Tạo repo **Public** mới trên GitHub, ví dụ `kho-tai-lieu`.
2. Trong repo, chọn **Add file > Upload files**, kéo toàn bộ nội dung thư mục này vào, kể cả thư mục `tai-lieu` và `vba`, rồi bấm **Commit changes**.
3. Vào **Settings > Pages**. Ở **Source** chọn **Deploy from a branch**, chọn nhánh `main` và thư mục `/ (root)`, rồi bấm **Save**.
4. Chờ 1 đến 2 phút, trang sẽ có tại `https://TEN-GITHUB.github.io/kho-tai-lieu/`.

## Tạo token để upload từ trang web

1. Mở GitHub, vào **Settings > Developer settings > Personal access tokens > Fine-grained tokens > Generate new token**.
2. Ở **Repository access**, chọn **Only select repositories** rồi chọn đúng repo này.
3. Ở **Permissions**, đặt **Contents** thành **Read and write**.
4. Tạo token, sao chép, rồi dán vào nút **Kết nối để upload** trên trang.

Không gửi token cho người khác và không dán token vào file trong repo. Nếu lộ token, hãy xóa nó trên GitHub ngay.

## Sử dụng

- Bấm **Thêm file** hoặc kéo file vào bảng để upload. Mỗi file tối đa 25 MB.
- File `.bas`, `.cls`, `.frm`, `.xlsm` và các định dạng macro khác tự vào nhóm VBA. Có thể đổi nhóm trước khi upload.
- Bấm **Xem** để mở code, PDF, ảnh và file Word, Excel, PowerPoint ngay trên trang, hoặc bấm **Tải** để tải xuống.
- Trang tự nhận code tiếng Nhật (Shift-JIS) và tiếng Việt (Windows-1258). Nếu chữ vẫn lỗi, đổi ô **Bảng mã** trong cửa sổ xem. File tải về luôn giữ nguyên bản gốc.
- File Office được hiển thị qua trình xem của Microsoft, nên file lớn có thể không mở được.
- UserForm có hai file `.frm` và `.frx`. Hãy upload cả hai.

## Cài thành app

- iPhone: mở trang bằng Safari, bấm Chia sẻ, chọn **Thêm vào MH chính**.
- Android: mở trang bằng Chrome, bấm menu **⋮**, chọn **Cài đặt ứng dụng** hoặc **Thêm vào Màn hình chính**.
- Máy tính: mở trang bằng Chrome hoặc Edge, chọn biểu tượng cài đặt ở cuối thanh địa chỉ.

App vẫn cần mạng vì file nằm trên GitHub.

## Lưu ý

- Repo và trang đều công khai. Không upload dữ liệu cá nhân, mật khẩu hay tài liệu nội bộ.
- Mỗi người xem có khoảng 60 lượt gọi GitHub API mỗi giờ, mỗi lần mở trang dùng 2 lượt.
- Khi thay file trùng tên, trang có thể hiển thị nội dung cũ trong vài phút.
- Nếu dùng tên miền riêng, điền `owner` và `repo` trong `CAU_HINH` ở đầu `app.js`.
