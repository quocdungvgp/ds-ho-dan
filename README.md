# Danh sách hộ dân · Bản đồ dự án — hướng dẫn dùng cho một dự án mới

**Trang web:** https://quocdungvgp.github.io/ds-ho-dan/ · Làm tiếp code: đọc [PHAT-TRIEN.md](PHAT-TRIEN.md)

## Các file
| File | Dùng để |
|---|---|
| `index.html` | Trang web (2 thẻ: Danh sách hộ dân, Bản đồ dự án). Đưa lên host bất kỳ (Cloudflare Pages, GitHub Pages, hosting của Ban). |
| `Code.gs` | Máy chủ Google Apps Script: lưu danh sách hộ vào Google Sheet, file thông báo/bản vẽ/mẫu vào Google Drive. |
| `dxf2bando.py` | Đổi bản vẽ CAD (DXF, VN2000) ra `bando.json` cho bản đồ. |
| `mau-thong-bao-vi-du.docx` | Ví dụ mẫu Word có ô điền «HOTEN» «DT»… |

## 1. Danh sách hộ
**Cách A – Excel (chạy ngay, lưu trên máy đang mở trang):** ⚙ Cấu hình → *Tải file Excel mẫu* → điền → *Nhập danh sách từ Excel*.
Trang tự tìm hàng tiêu đề có “Họ và tên”, tự nhận cột: STT, Mã số hồ sơ, Cán bộ thụ lý, Địa chỉ thu hồi, Phường, Khu phố, Tổ NVBT (ghi 1/2/3 hay “Tổ 1” đều được), Số tờ, Số thửa (nhiều thửa ghi `5, 6`), Diện tích thu hồi (`12,5` hay `12.5`), Giải tỏa (toàn phần/một phần), Số/Ngày thông báo, Số ảnh, Link hồ sơ hiện trạng, Ghi chú.

**Cách B – Google Sheet (cả Ban dùng chung):**
1. Tạo Google Sheet mới cho dự án → Tiện ích mở rộng → Apps Script → dán `Code.gs`.
2. Điền `THU_MUC_ID` = ID thư mục Drive chứa thông báo (đoạn sau `/folders/` trong link thư mục), đổi `MK_ADMIN_DAU`.
3. Chạy hàm `caiDat()` một lần (cấp quyền) → tạo tab “Tổng”, “NguoiDung”, “DeNghi”, “NhatKy” và tài khoản `admin`.
4. Triển khai → Tùy chọn triển khai mới → Ứng dụng web; Thực thi: *Tôi*; Truy cập: *Bất kỳ ai* → copy URL `/exec`.
5. Trên trang: ⚙ Cấu hình → dán URL → Lưu → đăng nhập `admin` → bấm tên mình góc trên → **đổi mật khẩu**.
6. *Nhập danh sách từ Excel* một lần để đổ dữ liệu lên Sheet (hoặc dán thẳng vào Sheet rồi chạy lại `caiDat()` để cấp id).
Sửa trên web ghi luôn vào Sheet; sửa trên Sheet thì bấm “Đọc lại” trên web.

### Người dùng & phân quyền (chỉ khi nối Google Sheet)
Quản trị bấm **👥 Người dùng** → thêm tài khoản: tên đăng nhập, họ tên, mật khẩu, quyền, phạm vi tổ.
| Quyền | Được làm |
|---|---|
| Quản trị | Mọi việc; duyệt / từ chối đề nghị; quản lý người dùng; nạp bản vẽ, mẫu TB; xem nhật ký; nhập Excel thay toàn bộ |
| Sửa | Thêm / sửa / xóa hộ, đính kèm / gỡ / tạo thông báo — có hiệu lực ngay; nhập Excel (thêm vào) |
| Đề nghị | Như Sửa nhưng thành **đề nghị chờ duyệt** (hiện ⏳ trên dòng hộ, quản trị bấm Duyệt / Từ chối; người gửi Rút được) |
| Xem | Chỉ xem, lọc, tải, in, bản đồ |
**Phạm vi tổ**: chọn Tổ 1…4 thì chỉ thao tác hộ của tổ đó (vẫn xem toàn bộ). Muốn ngưng ai: tích **Khóa** (giữ lịch sử) thay vì xóa.
Máy chủ kiểm quyền ở mọi yêu cầu, sai mật khẩu 5 lần khóa 15 phút, mật khẩu lưu dạng mã hóa có muối, phiên đăng nhập 6 giờ.
Mọi thao tác ghi vào tab **NhatKy** (nút 📜 Nhật ký; nút “Lịch sử” trên từng hộ).
Lưu ý: người mở thẳng file Google Sheet vẫn sửa được mọi thứ — chỉ chia sẻ Sheet cho quản trị.

## Link dự án (gửi anh em)
Quản trị mở ⚙ Cấu hình → ô **Link dự án** → *Chép* → gửi qua Zalo. Ai bấm link là trang tự nối đúng Google Sheet của dự án đó, chỉ còn đăng nhập. Link không chứa mật khẩu.
Tên dự án, tên đơn vị, đuôi mã hồ sơ, ngưỡng DT còn lại: quản trị sửa trong ⚙ Cấu hình là áp dụng cho mọi người của dự án.

## Tiện ích khác
- **Đính kèm hàng loạt** (quyền Sửa trở lên): chọn nhiều file .docx hoặc 1 file .zip; trang ghép hộ theo STT trong tên file (`TBTHD_STT008_…`, `STT8…`, `008_…`), báo trước file nào không ghép được.
- **Tự cập nhật**: máy khác sửa thì trong vòng 1 phút trang đang mở tự đọc lại.
- **Sắp xếp**: bấm tiêu đề cột STT, Mã, Cán bộ, Họ tên, Địa chỉ, Tổ, Tờ·Thửa, DT ở bảng Theo hộ.
- **Nhập Excel thay toàn bộ**: hộ trùng STT + họ tên giữ nguyên mã hồ sơ, thông báo đã đính kèm, cán bộ thụ lý.
- **Mã hồ sơ** cấp ở máy chủ, không trùng khi nhiều người mở cùng lúc.
- Bản đồ vẽ bằng canvas, nhãn chỉ hiện thửa trong khung nhìn — bản vẽ vài nghìn thửa vẫn mượt.

## Tiến độ giải phóng mặt bằng
Khung **Tiến độ giải phóng mặt bằng** (dưới các thẻ số): 7 bước *Kiểm đếm → Nhận định pháp lý → Lập phương án → Công khai PA → Phê duyệt PA → Chi trả → Bàn giao MB*.
- Mỗi bước: số hộ đã xong / tổng, % — bấm **còn … hộ chưa xong** để lọc ra đúng các hộ đó. Thanh màu trên cùng: số hộ đang dừng ở từng bước, bấm để lọc.
- **Hạn bàn giao mặt bằng**: quản trị đặt ở ⚙ Cấu hình (cấu hình chung, mọi người thấy) → hiện đếm ngày còn lại / quá hạn và số hộ chưa bàn giao.
- Cập nhật từng hộ: **✎ Sửa** → ghi **ngày xong** của bước nào là tính hộ đã xong bước đó (lấy bước cao nhất); chưa có ngày thì chọn ô *Tiến độ GPMB — bước đã xong*.
- Ô lọc **Mọi tiến độ GPMB**: đang ở bước / chưa xong bước / đã xong bước. Bảng Tổ có cột **Đã bàn giao MB**. Cột *Giải tỏa · Tiến độ* bấm tiêu đề để sắp xếp theo tiến độ.
- Bản đồ: *Tô màu thửa theo* → **Tiến độ GPMB** (thửa nhiều hộ lấy hộ chậm nhất).
- Excel: cột *Tiến độ GPMB* (ghi chữ hoặc số 1–7) và *Ngày Kiểm đếm*, *Ngày Chi trả*… (dd/mm/yyyy). Xuất Excel có đủ các cột này.
- Sheet cũ: dán Code.gs mới (Phiên bản mới) rồi chạy lại `caiDat()` để thêm 2 cột *Tiến độ GPMB*, *Ngày các bước GPMB* — chưa chạy thì sửa tiến độ không lưu được.

## Kinh phí bồi thường, hỗ trợ, tái định cư & chi trả
Khung **Kinh phí…**: tổng kinh phí phương án, đã chi, còn phải chi, số hộ chưa có số tiền; bảng theo tổ (bồi thường, hỗ trợ, TĐC, tổng, đã chi, còn phải chi, tỷ lệ). Bấm số để lọc hộ.
- Nhập ở **✎ Sửa**: *Tiền bồi thường, hỗ trợ, tái định cư, Đã chi trả* (gõ 1.500.000.000 hay 1500000000 đều được). Ngày chi trả ghi ở ô *6. Chi trả* phần tiến độ.
- Ô lọc **Mọi tình trạng chi trả**: chưa có số tiền PA / có PA / chưa chi đủ / đã chi đủ.
- **Xuất Excel** có dòng TỔNG CỘNG và trang *Theo tổ*. Nhập Excel nhận các cột *Tiền bồi thường, Tiền hỗ trợ, Tiền tái định cư, Đã chi trả*.
- Sheet cũ: dán Code.gs mới rồi chạy `caiDat()` (thêm 4 cột tiền).

## Cập nhật Code.gs khi có bản mới
Dán đè Code.gs → Triển khai → Quản lý triển khai → ✎ → Phiên bản: *Phiên bản mới* → Triển khai (URL /exec giữ nguyên).

## Thông báo Word tự sửa theo web
Sửa họ tên, diện tích, toàn phần/một phần, thửa, tờ trên web → mỗi lần **Tải / Xem trước / Tải zip** thông báo, trang tự thay đúng chỗ lệch trong file (giữ định dạng chữ). File gốc trên Drive chỉ đổi khi bấm **⟳ Đồng bộ file Word**: trang so mọi thông báo đang lọc, liệt kê chỗ lệch, đồng ý thì ghi hẳn.
Trang tìm theo câu chữ mẫu thông báo thu hồi đất chuẩn: “Thu hồi đất của Ông/Bà …”, “Diện tích đất dự kiến thu hồi: … m²”, “thuộc toàn bộ/một phần thửa đất số …, tờ bản đồ số …”. Mẫu dự án khác câu chữ thì sửa mảng `TB_FLD` trong index.html.

## Ảnh hiện trạng (PDF)
Nút **📷 Ảnh HT (PDF)** trên dòng hộ: đính kèm file PDF hồ sơ hiện trạng (Timemark: mỗi trang 1 ảnh). Trang tự đếm số ảnh; bấm 📷 để xem lưới ảnh, bấm ảnh xem lớn, ← → lật. Lưu trên Drive cùng thư mục thông báo. Sheet cũ chạy lại `caiDat()` để thêm cột *File ảnh hiện trạng*.

## Ảnh vệ tinh theo năm
Thẻ Bản đồ → ô **Ảnh vệ tinh theo năm** (Esri Wayback): mỗi năm một mốc, ghi ngày chụp thật tại dự án, bỏ năm trùng ảnh. Kéo **Độ rõ** để so với ảnh hiện tại.

## Sửa bản vẽ (quản trị)
Thẻ Bản đồ → **Sửa bản vẽ**: ✂ *Tách thửa* (bấm thửa → bấm 2 điểm đường cắt → đặt nhãn phần mới), ⊕ *Gộp thửa* (bấm thửa giữ nhãn → bấm thửa gộp vào: 2 mảnh tính chung 1 thửa, cộng diện tích), ✎ *Sửa nhãn*, ↶ *Hoàn tác*, 💾 *Lưu bản vẽ* (lên Drive cho cả dự án).

## 2. Bản đồ thửa
1. AutoCAD: bản vẽ ở tọa độ VN2000 thật (TP.HCM: X ≈ 600 000, Y ≈ 1 190 000). SAVEAS → DXF 2013.
2. `pip install ezdxf` rồi:
   `python dxf2bando.py "Ban do du an.dxf" --thua THUA_DAT --nhan SO_THUA --ranh RANH_THU_HOI`
   - `--thua`: lớp polyline KÍN ranh thửa · `--nhan`: lớp text “tờ/thửa” nằm trong thửa (chỉ ghi số thửa thì thêm `--to 134`) · `--ranh`: lớp ranh dự án (lặp được).
   - Sai tên lớp thì script in ra danh sách lớp có trong bản vẽ.
3. Bỏ `bando.json` cạnh `index.html` trên host (cả Ban thấy), hoặc ⚙ Cấu hình → *Nạp bản vẽ* (lưu trên máy đó; có Apps Script thì lưu lên Drive cho máy khác).
Thửa ghép với hộ theo “Số tờ/Số thửa”. Hệ VN2000 TP.HCM kinh tuyến trục 105°45′ — dự án tỉnh khác thì sửa dòng `proj4.defs('VN2000HCM', …)` trong `index.html` (+lon_0=…).

## 3. Thông báo thu hồi đất
- Đính kèm từng hộ: nút *Đính kèm* trên dòng hộ (file .docx).
- Tạo theo mẫu: soạn mẫu Word của dự án, chỗ thay đổi gõ đúng: `«HOTEN»` (tự thêm “Ông/Bà”, tổ chức thì giữ tên), `«TEN»`, `«DT»`, `«MUC»` (toàn bộ/một phần), `«THUA»`, `«TO»`, `«STT»`, `«MA»`, `«DIACHI»`, `«PHUONG»`, `«KP»`, `«SOTB»`, `«NGAY»`. Gõ liền một lần, đừng sửa lẻ từng chữ trong ô (Word tách đoạn thì không nhận).
  ⚙ Cấu hình → *Nạp mẫu TB* → hiện nút *Tạo TB* từng hộ và *Tạo TB hàng loạt* (chỉ hộ đủ họ tên, tờ, thửa, diện tích, toàn phần/một phần).
- *Tải TB (zip)* tải mọi thông báo của các hộ đang lọc.

## 4. Ảnh hiện trạng
Ghi số ảnh và link thư mục/tệp Drive hồ sơ hiện trạng vào cột “Số ảnh hiện trạng”, “Link hồ sơ hiện trạng”. Nút 📷 trên dòng hộ mở link.

## Chưa có so với trang Ban đang dùng
Tô màu tay từng thửa; gộp thửa là gộp theo nhãn (không hợp nhất đường ranh thành 1 đa giác).
