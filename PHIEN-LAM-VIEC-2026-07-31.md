# Lịch sử trò chuyện & chỉnh sửa — 2026-07-31

Dự án: `G:\01. Github\RP---GG-Studio-` (Electrical Panel Reviewer / VoltReview AI)

---

## 1. Preview web app

Yêu cầu: "preview web app này".

- Phát hiện port 3000 đang bị chiếm bởi dev server của phiên chat khác.
- Dùng cấu hình dự phòng `dev-3100` trong [.claude/launch.json](.claude/launch.json) (port 3100).
- Lần khởi động đầu tiên gặp lỗi tạm thời từ Vite (`Cannot find module './ResultPlugin'` trong `enhanced-resolve`) — chạy thử trực tiếp qua Git Bash thì hết lỗi, khởi động lại qua tool là chạy sạch. Kết luận: lỗi thoáng qua của node_modules/vite cache, không phải bug thật.
- App chạy tốt tại `http://localhost:3100`, tự nạp file demo `BANG_TINH_TAI_DIEN_DEMO.xlsx`.

---

## 2. Tìm & đọc lịch sử trò chuyện cũ của dự án ReviewTuDien

- User cung cấp đường dẫn `.claude` project session không tồn tại trên máy này (`00c091a4-...jsonl`).
- User gửi file backup `H:\My Drive\xx. Zalo Download\.claude.zip` — kiểm tra bên trong, xác nhận đây là bản backup toàn bộ thư mục `.claude` từ máy khác, **chứa đúng file jsonl bị thiếu** (`projects/D--09--Githup-ReviewTuDien/00c091a4-b05e-456a-a33b-80c41ea09b1c.jsonl`, 5MB, 1537 dòng).
- Đã kiểm tra: không có API key/token dạng plaintext trong các bản backup `.claude.json`, chỉ có metadata tài khoản.
- Trích xuất và tóm tắt toàn bộ nội dung phiên cũ (28/07/2026) — các mốc chính:
  1. Đọc hiểu kiến trúc dự án (React 19 + Vite + TS, patch ZIP OOXML trực tiếp thay vì rebuild bằng SheetJS).
  2. Sửa rule nhận diện tủ MSB — chỉ dựa vào **tên sheet** chứa "MSB", bỏ cách quét toàn ô A1:F9.
  3. Đổi bố cục: đưa khối chọn tủ lên trên, thêm nút Reload, theo dõi file Excel thay đổi real-time.
  4. Đổi hiển thị lỗi mạch sang icon (lỗi/warning/info) thay vì text.
  5. Cho nhập tỉ lệ fill ống tuỳ chỉnh.
  6. **Đổi form tủ điện** — viết module [panelLayout.ts](src/utils/panelLayout.ts) tự dò cột theo tiêu đề thay vì gán cứng toạ độ, hỗ trợ cả form cũ và form mới; phát hiện & sửa thêm 4 lỗi thật (cú pháp cáp `x` không nhận, mất hệ số nhân số sợi cáp, lưu `.xls` ra toàn số 0, ngưỡng chống hỏng file quá chặt).
  7. Đổi danh sách chọn tủ từ hàng tab sang dropdown (khắc phục vỡ layout khi nhiều tủ lỗi).
  8. Tin nhắn cuối: yêu cầu load bảng tra `Spec. Cable.xlsx` từ folder `data/` — **bị cắt ngang do hết session limit**, chưa hoàn thành.
- User xác nhận việc "đổi form tủ điện" ở trên đã hoàn tất trong code hiện tại (đã kiểm tra `panelLayout.ts` tồn tại trong project đang mở).

---

## 3. Các chỉnh sửa UI trong phiên này (đều đã kiểm chứng qua trình duyệt bằng HMR, không cần reload)

Tất cả chỉnh sửa nằm trong [src/components/PanelViewer.tsx](src/components/PanelViewer.tsx), trừ mục 3.7.

### 3.1. Thu hẹp khoảng cách 2 dòng tổng kết (footer)
- `footCellCls`: `py-2.5` → `py-1`.
- Kết quả: khối footer từ 113.25px → 89.25px (−24px), không cắt nội dung.

### 3.2. Bỏ khung viền badge "Lộ vào"
- Bỏ `px-3 py-1 rounded-lg bg-white border-2 border-[#1B7A45]` — chỉ giữ icon + chữ trên nền xanh của dòng footer.

### 3.3. Thu hẹp cột "Loại CB" + căn giữa tiêu đề
- Giảm `min-w` của `<select>` (`cbTypeSelectCls`) và của `<td>` bọc ngoài.
- Phát hiện: `<td>` có `min-w-[150px]` riêng mới là ràng buộc thật (không phải class của `<select>`) — sửa đúng chỗ đó thì mới có tác dụng.
- Kết quả cuối: `<td>` min-w 150px → 110px, select rộng thực tế 138px → 118px, 0/6 select bị cắt chữ.
- Header "Loại CB": thêm `text-center`.

### 3.4. Đồng bộ cỡ chữ toàn bộ ô dữ liệu bảng = 12px (bằng tiêu đề "MSB-01")
- Đổi `inputCls`, `cbTypeSelectCls`, class `<tbody>`, và các `<td>` có size riêng (16px/15px/14px) → đồng nhất `text-[12px]`.

### 3.5. Cỡ chữ + căn giữa toàn bộ tiêu đề cột (`<thead>`)
- `<thead>`: `text-[14px]` → `text-[12px]`.
- Toàn bộ 16 cột: thêm/đổi thành `text-center` (trước đó nhiều cột `text-right` hoặc mặc định trái).

### 3.6. Đồng bộ cỡ chữ khối tổng kết (tfoot) = 12px
- `<tfoot>` base, `footNumCls` override (Full), dòng "Dòng điện tính toán", 2 span "Lộ vào:" — tất cả về `text-[12px]`.

### 3.7. Sửa lỗi hiển thị sai "Lộ vào" (nguồn cấp) — [src/utils/excelParser.ts](src/utils/excelParser.ts)
- **Lỗi**: với form LAV3 (`LAV3 - PANEL.xlsx`, 109 tủ), cột mô tả/tên mạch ở dòng chứa CB tổng lại trùng nhãn của chính dòng đó (vd. "DÒNG ĐIỆN TÍNH TOÁN/ CURRENT (A)"), bị nhầm thành nguồn cấp.
- **Sửa**:
  1. Thêm chặn: nếu chuỗi lấy được trùng nhãn dòng tổng kết (`isSummaryRow`) → bỏ, không hiển thị.
  2. Thêm hàm `findHeaderSupplySource()` — dự phòng đọc "Nguồn cấp từ/ Supply from:" ở khối tiêu đề đầu sheet (dòng 1–9), xử lý cả 2 kiểu nhãn+giá trị chung ô hoặc tách ô.
- **Kiểm chứng trên file thật** (script trực tiếp qua `parseExcelWorkbook`, 109 sheet):
  - `LV-MSB-01`: sai *"DÒNG ĐIỆN TÍNH TOÁN/ CURRENT (A)"* → đúng **"MÁY BIẾN ÁP KHÔ TR-01 / DRY TRANSFORMER TR-01"**.
  - `Busway THÁP 1 (TRỤC 1)`: ô nguồn cấp trong file gốc thực sự để trống → nay hiện đúng trống ("—") thay vì nhãn sai.
  - 0/109 tủ còn source dạng nhãn rác.
  - Demo (`MSB-01`, `DB-L01`) không bị ảnh hưởng, vẫn đúng `FROM LV-TRANSFORMER` / `FROM MSB-01`.

---

## Ghi chú
- Dev server preview đang chạy ở `http://localhost:3100` (cấu hình `dev-3100` trong `.claude/launch.json`), vì port 3000 bị chiếm bởi phiên chat khác.
- File này do assistant xuất theo yêu cầu, tóm tắt lại nội dung đã trao đổi — không thay thế lịch sử chat gốc.
