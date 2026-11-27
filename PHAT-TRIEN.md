# PHÁT TRIỂN TIẾP — ds-ho-dan

> **Đọc file này trước khi sửa code.** Đủ để làm tiếp mà không phải đọc lại cả `index.html` (~1.500 dòng).
> Cách nhờ Claude làm tiếp, đỡ tốn token: *"Đọc PHAT-TRIEN.md trong repo quocdungvgp/ds-ho-dan, làm mục X ở phần Việc tiếp theo, chạy tests rồi đẩy lên."*

- Trang: https://quocdungvgp.github.io/ds-ho-dan/ (GitHub Pages, nhánh `main`, thư mục gốc — đẩy lên ~1 phút là cập nhật)
- Repo: https://github.com/quocdungvgp/ds-ho-dan (công khai — **không đưa dữ liệu hộ dân, bando.json thật, mật khẩu vào repo**)
- Bản gốc tham khảo: banqldabinhquoi.com `#dsmetro2`, `#bandometro2` (trang nội bộ của Ban, cần đăng nhập)
- Hướng dẫn người dùng: `HUONG-DAN.md` (= `README.md`)

---

## 1. Kiến trúc

```
Trình duyệt (index.html, 1 file, không build)          Google (mỗi dự án 1 bộ)
  ├─ Danh sách hộ #ds  ─┐                               ┌─ Google Sheet: tab Tổng, NguoiDung, DeNghi, NhatKy
  ├─ Bản đồ #bando     ─┼── POST JSON {action,token} ──►│  Apps Script Code.gs (Ứng dụng web /exec)
  └─ IndexedDB/localStorage (chế độ không nối Sheet)    └─ Drive: TB .docx, ảnh HT .pdf, bando.json, mau-thong-bao.docx
dxf2bando.py: bản vẽ CAD (DXF, VN2000) → bando.json · sw.js + manifest.json + icon.svg: cài như ứng dụng, mở được khi mất sóng
```

- **2 chế độ**: `CFG.api` trống → chạy một mình trên máy (dữ liệu mẫu / Excel, lưu `localStorage` + IndexedDB `dsho`), `ME` = quản trị.
  Có `CFG.api` → đăng nhập, mọi thứ qua Apps Script; máy chủ kiểm quyền lại.
- **Link dự án**: `…/ds-ho-dan/?api=<URL /exec>` → tự lưu vào cấu hình, xóa tham số khỏi thanh địa chỉ.
- Thư viện tải từ jsDelivr, ghim phiên bản (bảng `LIB` trong mục 1 và `<head>`): leaflet 1.9.4, proj4 2.12.1 (nạp sẵn);
  xlsx 0.18.5, jszip 3.10.1, pizzip 3.1.7, docxtemplater 3.47.4, docx-preview 0.3.2, pdfjs-dist 3.11.174, qrcode-generator 1.4.4, jsqr 1.4.0, jspdf 2.5.1, polygon-clipping 0.15.7 (nạp khi cần qua `need('xlsx',…)`).
  Thêm thư viện: `LIB` + `tests/package.json` + bảng `CDN` trong `tests/common.py`.
- Hệ tọa độ: `proj4.defs('VN2000HCM', …)` — kinh tuyến trục 105°45′, múi 3°, k=0,9999, 7 tham số towgs84. Tỉnh khác sửa `+lon_0`.

## 2. Bản đồ code `index.html` (tìm bằng tiêu đề mục)

`grep -n "^   [0-9]" index.html` ra danh sách mục:

| Mục | Nội dung | Hàm chính |
|---|---|---|
| 1. TIỆN ÍCH | `$ esc num f1 fold toast download b64 need() idb` | `need(...ks)` nạp thư viện CDN |
| 2. CẤU HÌNH | `CFG0`, `CFG` (localStorage `dsho_cfg`), link `?api=` | `saveCfg linkDA NG()` |
| 3. NGƯỜI DÙNG & QUYỀN | `ME TOKEN RANK has() inScope() canRow() isAdmin()`; nhóm dữ liệu `NHOM NHOM_COT nhomOf() canG() canRowG(r,g)` | khớp `NHOM/NHOM_COT` trong Code.gs |
| 4. DỮ LIỆU | `api() load(nhe) saveRow delRow capMa getTB putTB delTB`, `SHARED` cấu hình chung | `xong(j,viec)` báo + đọc lại |
| 4b. DỮ LIỆU MẪU | `sample() sampleMap()` | |
| 5. LỌC & SỐ LIỆU | `st` (bộ lọc), `PARCEL`, `conLai()` (dùng `THU` tổng DT theo tờ\|thửa, xóa ở `render()`), **kiểm tra dữ liệu `tinhLoi()` → `LOI`**, `filtered()`, đề nghị `CHO choOf virtRows choHtml`; **tiến độ GPMB** `GD GD_COL gdIdx gdOf gdNgay gdMatch gdCell renderGd()` | `gdIdx(r)` = bước đã xong (−1…6) |
| 6. VẼ DANH SÁCH | `render()` (thẻ số, bảng tổ, chip, ô lọc) `tbCell() list() sortRows()`; phân trang `TRANG nHien xemThem inDu` (in thì vẽ đủ) | ô ít quan trọng có class `ph` (điện thoại ẩn khi thẻ gọn) |
| 7. THÔNG BÁO WORD | `attach xem taoTB taoLoat`, **`tbFld() docFix()`** tự sửa Word theo web (mốc câu chữ `tbm_*` trong cấu hình chung), `dongBo() zipTB hoi(msg,nhãn,mặcĐịnh)` | |
| 7b. NHIỀU MẪU VĂN BẢN | `truong(r)` bộ trường chung (cả mẫu TB), `TRUONG_DS`, `docSo()` số→chữ, `dien(mau,r)`, `MAUVB loadMauVB getMauVB putMauVB delMauVB moVB taoVB` | thêm trường: sửa `truong()` + `TRUONG_DS` |
| 8. EXCEL | `COLS` (khóa, tiêu đề, regex nhận cột; `gdn_<bước>` = cột ngày ảo), `colVal()` giá trị xuất, `importXls exportXls` | |
| 9. BẢN ĐỒ | `initMap CLS(tô màu: tb to gt cl anh gd) drawMap labels pickParcel suggest meas* exportKML`, `HO_IDX hoOf` | |
| 10. SỰ KIỆN | `route()` (#ds/#bando), click `data-a=…`, hộp sửa hộ (gửi `chg`+`goc` qua `saveRow(r,viec,chg,goc)`), cấu hình; điện thoại: `#bMore`, `#bLoc demLoc()`, khung gập `.panel.gap` (`dsho_dong`), `[data-mo]` | |
| 11. ĐĂNG NHẬP… | `topBar showLogin`, người dùng `openUsers userForm`, `duyet drawDuyet openLog` | |
| 12. BỔ SUNG | sắp xếp cột, chép link, tự đọc lại (`ping` 60 s), đính kèm hàng loạt | |
| 13. ẢNH HIỆN TRẠNG | `pdfLib getHT attachHT openAnh drawThumbs bigAnh` | |
| 14. ẢNH THEO NĂM | `loadWayback` (Esri waybackconfig.json + metadata ngày chụp) | |
| 15. SỬA BẢN VẼ | `ED edSet edClick clipHalf edPush edApply`, lưu `putMap` | |
| 17. BẢN ĐỒ NÂNG CAO | `khungTen()` (in A3/A4, `#khung`, `@page` qua `<style id=pageKho>`), `exportDXF() aci() diemNhan()`, `docBanVe() tuVung() nhanTu()` (GeoJSON/KML), `gopHinh()` | DXF R12: lớp THUA_DAT/SO_THUA/RANH_THU_HOI — đọc lại bằng dxf2bando.py |
| 16. QR & CHỤP ẢNH HT | `hoLink moHo PEND_HO inNhan quetQR qrTuAnh docQR`, `pickFiles viTri dongDau chupHT` → `luuHT(r,buf,name)` (mục 13) | link `#ho=<id>` xử lý trong `route()` + cuối `render()` |

Quy ước: nút thao tác trên dòng dùng `data-a="<việc>" data-id="<id hộ>"`, bắt chung ở mục 10. Lọc nhanh dùng `data-f='{json bộ lọc}'`.
Giá trị từ Sheet luôn qua `esc()` trước khi ghép HTML. Hộp xác nhận dùng `hoi(msg, nhãnÔNhập?)` (không dùng confirm/prompt).

## 3. Dữ liệu (Sheet tab "Tổng") — mảng `COT` trong Code.gs

`stt ma canbo ten diachi phuong kp to(Tổ NVBT) to_bd(số tờ) thua(“5, 6” nhiều thửa) dt giaitoa(Toàn phần|Một phần) tb_so tb_ngay tb_file tb_ten anh ht(link) ht_file ht_ten tien_bt tien_ht tien_tdc da_chi gd gd_ngay duyet ghichu by at id`
- Thêm cột: thêm vào `COT` (Code.gs) + `COLS` (index.html, nếu nhập Excel) + form `#fm` + chỗ hiển thị; chạy lại `caiDat()` trên Sheet.
- Khóa thửa: `"tờ/thửa"` (`keyThua`) — ghép hộ ↔ bando.json (`thua[].l`).
- Mã hồ sơ: `###/P{phường}/{đuôi}.{số tổ}`, máy chủ cấp (`capMa_`), cấp rồi không đổi.
- **Tiến độ GPMB**: `gd` = nhãn bước đã xong gần nhất (“Chi trả”; đọc được cả số 1–7 hoặc chữ gần đúng qua regex trong `GD`), `gd_ngay` = JSON `{kd,pl,pa,ck,pd,ct,bg: "yyyy-mm-dd"}`.
  Bước thật = max(`gd`, bước cao nhất có ngày) — `gdIdx()`. Lưu lại thì trang ghi `gd` theo bước thật. Thêm/bớt/đổi tên bước: sửa mảng `GD` (+ `GD_COL` đủ `GD.length+1` màu).
- **Kinh phí**: `tien_bt tien_ht tien_tdc da_chi` lưu số đồng nguyên (chuỗi). `tien()` đọc được “1.500.000.000”, “1,5e9” không; `tongPA() ctOf()`; ngày chi trả = `gd_ngay.ct` (không có cột riêng). Lọc `st.ct`: chua/co/thieu/du.
  Bộ lọc `st.gd`: `e<i>` đang ở bước i (`e-1` chưa bắt đầu), `c<i>` chưa xong bước i, `x<i>` đã xong. Hạn bàn giao: `CFG.hanBG` (trong `SHARED`).

**bando.json**: `{v, o:[x0,y0], thua:[{id, l:"tờ/thửa", v:[[x,y]…] (m, đã trừ o), dt?}], lop:{ranh:[[[x,y]…]]}}`

**Tab khác**: `NguoiDung` (tên, họ tên, quyền, phạm vi tổ, hash, muối, khóa, đăng nhập cuối) · `DeNghi` (id, loại, hộ id, nội dung JSON, người, lúc, trạng thái cho/duyet/tuchoi/rut/loi, …) · `NhatKy`.

## 4. Máy chủ Code.gs — các `action`

| action | quyền | ghi chú |
|---|---|---|
| login | — | trả `token`: Cache 6 giờ + bản chính trong ScriptProperties `tk_<token>`={ten,het} (`PHIEN_GIO`=72), sai 5 lần khóa 15 phút |
| me, logout, changePw, list, ping, getFile, getMap, getMau | đăng nhập | `list` trả `rows, cho, me, cfg, sheetUrl`; `getFile` chỉ file đang gắn hộ/đề nghị |
| upsert, delete, uploadTB(kind tb/ht), delTB, delHT | denghi+ (trong phạm vi tổ, đúng **nhóm dữ liệu**) | quyền `denghi` → thành đề nghị (`propose_`). `upsert` nhận `goc` (giá trị lúc mở form) → ô người khác vừa đổi thì trả `{xungdot:[{k,goc,moi,minh}],by,at}`, `ep:true` = ghi đè |
| bulk, capMa | sua+ | bulk cần mọi nhóm; capMa cần nhóm hoso |
| mauDS, getMauVB | đăng nhập | mẫu văn bản khác: ScriptProperties `mau_vb`=[{ten,id}] |
| replaceAll, approve, reject, users, userSave, userDel, putMap, putMau, putMauVB, delMauVB, putCfg, log | quantri | |
| withdraw | người gửi | |

**Nhóm dữ liệu** (cột NguoiDung “Được sửa (nhóm)”, trống/* = mọi nhóm): `hoso tb ht tien gd` — `NHOM_COT` gán cột vào nhóm, cột còn lại là `hoso`; thêm/xóa hộ cần `hoso`; đính kèm/gỡ TB cần `tb`, ảnh HT cần `ht`. Kiểm ở `change_`, `uploadTB_`, `bulk`, `capMa` (`nhom_()`).
Hàm chạy tay/trigger: `caiDat()` (tạo tab, cột), `caiTuDong()` (cài trigger `saoLuu` 1 giờ sáng + `guiBaoDeNghi` 15 phút). `list_()` nhớ trong `MEMO_` suốt 1 yêu cầu, mọi chỗ ghi Sheet đặt `MEMO_ = null`.
Mọi thay đổi gọi `log_()` → ghi NhatKy + đổi `ver` (trang đang mở tự đọc lại). `onEdit()` làm việc này khi sửa tay trên Sheet.
Giới hạn Apps Script cần nhớ: 1 lần gọi tối đa ~6 phút; khối dữ liệu gửi lên ~50 MB (trang chặn PDF > 35 MB); CacheService tối đa 6 giờ.

## 5. Chạy thử (không cần mạng, không cần Google)

```bash
cd tests && npm i && pip install playwright python-docx openpyxl pillow --break-system-packages
python3 fixtures.py                 # tạo file mẫu vào tests/fx
python3 ui_local.py                 # chế độ một máy: 15 bước
node mock-gas.js &                  # Apps Script giả lập (chạy lại mỗi lần cho sạch dữ liệu), cổng 8787
python3 ui_server.py                # chế độ Sheet + phân quyền: 14 bước
```
`common.py` chặn CDN → trả từ `tests/node_modules`, chuyển `script.google.com` → máy giả lập. Lần cuối: **73/73 OK** (27/11/2026). Test DXF cần `pip install ezdxf`. mock-gas có lối test `/__het` (xóa Cache) và `/__mail` (chạy `guiBaoDeNghi`, trả thư). Chạy cả bộ: `sh tests/chay.sh`.
Chrome: `/usr/bin/google-chrome` (đổi bằng biến `CHROME`). File tải về trong test không có đuôi → `save_as()` ra file có đuôi rồi mới mở bằng openpyxl.
Chưa thử: Google Apps Script thật, ảnh nền bản đồ (chỉ thử tay trên trang thật — ảnh vệ tinh và Wayback chạy), camera quét QR thật + GPS thật (test dùng nhánh “chọn ảnh”).

## 6. Đẩy lên

Sửa → `sh tests/chay.sh` → `git commit` → `git push` (GitHub Pages tự cập nhật). Sửa `Code.gs` thì nhắc người dùng: dán đè trong Apps Script → Triển khai → Quản lý triển khai → ✎ → **Phiên bản mới**; thêm cột thì chạy `caiDat()`; trigger mới thì chạy `caiTuDong()`.

---

## 7. Việc tiếp theo

### Đã xong (27/11/2026) — chi tiết ở mục 8
Tiến độ GPMB 7 bước + hạn bàn giao · Kinh phí BT–HT–TĐC & chi trả · Nhiều mẫu văn bản · QR hồ sơ · Chụp ảnh HT bằng điện thoại ·
Sao lưu tự động · Email báo đề nghị · Xuất DXF · In A3/A4 khung tên · Gộp hình học thật · Nhập GeoJSON/KML · Dùng khi mất sóng (PWA + hàng đợi) ·
Câu chữ mẫu TB vào cấu hình chung · Phiên 3 ngày · Đọc Sheet 1 lần/yêu cầu. **Đã bỏ: nối Base.vn** (theo yêu cầu).

### Còn lại
1. **Đọc hồ sơ quét bằng AI → điền thông tin hộ** (chưa làm: cần máy AI offline của Ban / API có khóa — trang công khai không được chứa khóa).
   Hướng làm: nút “Đọc hồ sơ” trên dòng hộ gửi PDF tới 1 dịch vụ nội bộ (máy AI offline, Ollama + model đọc ảnh) qua địa chỉ khai trong cấu hình chung → trả JSON {ten, diachi, thua, to_bd, dt, …} → mở form sửa hộ điền sẵn, người dùng kiểm rồi mới lưu.
   Nối skill nhận định nhà đất Metro 2 để ra luôn Giấy xác nhận pháp lý.
2. Hạn riêng từng bước GPMB (vd hạn chi trả) + tô đỏ hộ trễ hạn; biểu đồ số hộ xong theo tuần (từ `gd_ngay`).
3. Mất sóng: lưu sẵn ảnh nền bản đồ khu dự án (sw.js hiện không lưu tile — mất sóng thì thửa hiện trên nền xám). Cẩn thận dung lượng.
4. Bản vẽ lớn (> 5 MB): chia `bando.json` theo tờ, nạp tờ đang xem.
5. Bỏ cột “Chờ duyệt” (`duyet`) — cờ tay, trùng nghĩa với đề nghị chờ duyệt.
6. Xuất báo cáo tiến độ ra Word theo mẫu Ban (số liệu tiến độ + kinh phí theo tổ) — nối skill báo cáo tiến độ Metro 2.
7. Điện thoại: thẻ hộ còn ~250 px/hộ — có thể gộp Tổ, Tờ·Thửa, DT lên 1 dòng.

### Kỹ thuật / giới hạn đã biết
- `bando.json` lưu Drive 1 file, đọc qua Apps Script (base64 JSON) — bản vẽ > 5 MB sẽ chậm.
- Chế độ một máy: dữ liệu chỉ trong 1 trình duyệt (xóa dữ liệu trình duyệt là mất) → nhắc xuất Excel thường xuyên.
- Ai mở thẳng file Google Sheet sửa được hết, không qua phân quyền → chỉ chia sẻ Sheet cho quản trị.
- Hàng đợi mất sóng gửi theo thứ tự; máy chủ vẫn kiểm quyền/phạm vi khi gửi — việc bị từ chối thì báo rồi bỏ khỏi hàng đợi.
  Sửa trùng hộ với người khác lúc mất sóng: bản gửi sau ghi đè (không trộn).
- Service worker chỉ chạy trên https (GitHub Pages); test chạy http nên không thử được SW, chỉ thử hàng đợi.
- `sw.js`: đổi hằng `BAN` khi muốn ép mọi máy bỏ bộ nhớ cũ (index.html luôn lấy mạng trước nên thường không cần).

---

## 8. Nhật ký làm

### 27/11/2026 — Tiến độ GPMB theo giai đoạn (mục 7.1)
- `index.html`: khung **Tiến độ giải phóng mặt bằng** (`#pGd`: hạn bàn giao đếm ngày, thanh màu số hộ theo bước, 7 ô bước có % và nút “còn … hộ chưa xong”);
  ô lọc `#fGd`; cột *Đã bàn giao MB* ở bảng tổ; cột *Giải tỏa · Tiến độ* (vạch 7 bước + bước/ngày, sắp xếp được); form sửa hộ có chọn bước + 7 ô ngày (`gdn_<mã>` → gộp thành `gd_ngay` khi lưu);
  ô *Hạn bàn giao mặt bằng* (`hanBG`, cấu hình chung); bản đồ chế độ tô màu `gd` (hộ chậm nhất trên thửa) + dòng tiến độ ở khung thông tin thửa + mô tả KML;
  Excel nhập/xuất cột *Tiến độ GPMB* + *Ngày <bước>* (dd/mm/yyyy hoặc ngày Excel); dữ liệu mẫu có tiến độ rải đều; đề nghị chờ duyệt hiện `gd_ngay` dạng chữ.
  Regex cột Phường đổi `/phuong/` → `/^phuong|^p$/` (tránh ăn nhầm cột “Ngày lập phương án”).
- `Code.gs`: thêm `gd`, `gd_ngay` vào `COT` → **người dùng phải dán Code.gs mới + Phiên bản mới + chạy `caiDat()`** (Sheet cũ chưa có 2 cột thì sửa tiến độ bị bỏ qua im lặng).
- Tests: thêm `fx/ds_gd.xlsx`, 6 bước ở ui_local, 1 bước ở ui_server → 36/36 OK.

### 27/11/2026 — Kinh phí BT–HT–TĐC & chi trả (mục 7.2)
- Khung `#pTien`: 4 thẻ (tổng PA, đã chi, còn phải chi, hộ chưa có số tiền) + bảng theo tổ (`renderTien`). Dòng hộ: “PA 1,7 tỷ · đã chi 50%” (`tienCell`), khung thông tin thửa trên bản đồ cũng có.
- Form sửa hộ 4 ô tiền (hiện có dấu chấm nghìn, lưu số nguyên). Ô lọc `#fCt`.
- Xuất Excel: thêm cột Tổng PA, Còn phải chi; **dòng TỔNG CỘNG** cuối (ô Họ tên trống nên nhập lại không thành hộ); trang 2 **“Theo tổ”**; định dạng `#,##0`.
- `Code.gs` `COT` thêm 4 cột → dán Code.gs mới + Phiên bản mới + `caiDat()`.
- `tests/chay.sh` chạy cả 2 bộ.

### 27/11/2026 — Máy chủ: sao lưu, email đề nghị, phiên dài, đọc Sheet 1 lần (mục 7.6, 7.7 + kỹ thuật)
- `Code.gs`: hằng `PHIEN_GIO=72 EMAIL_QT SAO_LUU_GIU`; `caiTuDong() saoLuu() guiBaoDeNghi() emailQT_()`; token lưu ScriptProperties; `MEMO_`.
- **Người dùng phải**: dán Code.gs mới → Phiên bản mới → chạy `caiTuDong()` 1 lần (cho phép quyền Drive + Gmail). `saoLuu` chưa thử được trên máy giả lập (cần Drive thật).

### 27/11/2026 — Nhiều mẫu văn bản (mục 7.3)
- Section 7b mới. Nút hero **📄 Tạo văn bản (theo mẫu)** (các hộ đang lọc → zip) + nút **📄 Văn bản** trên dòng hộ; ⚙ Cấu hình: **Nạp mẫu văn bản khác**, danh sách mẫu có ✕ xóa (quản trị).
- Trường mới: «TO_NVBT CANBO GIAITOA GHICHU TIEN_BT TIEN_HT TIEN_TDC TONG TONG_CHU DA_CHI DA_CHI_CHU CON_LAI GIAIDOAN NGAY_KD…NGAY_BG DUAN DONVI HOM_NAY HAN_BG».
  `docSo()` đọc số kiểu “Một nghìn năm trăm tỷ”, “Một triệu không trăm linh năm nghìn”. Nạp mẫu có thử `dien()` trước → báo lỗi trường hỏng.
- `Code.gs`: action `mauDS getMauVB putMauVB delMauVB` → dán Code.gs mới + Phiên bản mới (không cần `caiDat`).

### 27/11/2026 — QR hồ sơ + chụp ảnh hiện trạng bằng điện thoại (mục 7.4, 7.5)
- Hero: **🏷 In nhãn QR** (các hộ đang lọc), **⌗ Quét QR hồ sơ**. Dòng hộ: **🏷 QR**, **📷 Chụp ảnh HT** (nút đính kèm PDF đổi thành 📎).
- `attachHT` tách thành `attachHT` (chọn file) + `luuHT` (đếm trang, lưu) để chụp ảnh dùng lại.
- Không đổi Code.gs.

### 27/11/2026 — Bản đồ: xuất DXF, in khung tên, gộp hình học, nhập GeoJSON/KML (mục 7.8–7.11)
- Thanh bên bản đồ: ô khổ giấy (A3/A4) + **⇩ Xuất DXF (AutoCAD)**; ⚙ Cấu hình: Nạp bản vẽ nhận `.geojson`, `.kml`.
- Test: DXF xuất ra được `dxf2bando.py` đọc lại đủ thửa + ranh; tách rồi gộp 2 nửa thành 1 đa giác; GeoJSON/KML đổi VN2000 đúng diện tích (~98 m²).
- Không đổi Code.gs.

### 27/11/2026 — Dùng khi mất sóng + câu chữ mẫu TB vào cấu hình (mục 7.13 + kỹ thuật)
- `sw.js` (index.html mạng trước/đã lưu khi mất mạng; thư viện jsDelivr lưu 1 lần), `manifest.json`, `icon.svg` → “Thêm vào màn hình chính”.
- `api()`: lỗi mạng với lệnh sửa (`QUEUE_A`) → cất IndexedDB `hang`, trả `{queued}`; `apHang()` áp lên danh sách để xem (dòng có nhãn “⇡ chờ gửi”);
  `guiHang()` chạy khi sự kiện `online`, khi ping 60 s thành công, hoặc bấm nút **⇡ N chờ gửi** trên thanh đầu. Danh sách đọc lần cuối lưu IndexedDB `dsCache`.
- `TB_FLD` thay bằng `tbFld()` dựng từ cấu hình chung `tbm_<trường>_a/_b` (mốc chữ trước/sau giá trị; `TBM`, `TBM0`, `reMoc`). ⚙ Cấu hình có mục gập “Câu chữ mẫu thông báo”.
- Test: `tests/common.py` có `MANG['tat']` giả mất mạng tới Apps Script. 60/60 OK. Không đổi Code.gs.

### 27/11/2026 — Chống ghi đè, phân quyền theo nhóm dữ liệu, tốc độ, điện thoại, kiểm tra dữ liệu
- **Chống ghi đè**: trước đây lưu gửi cả dòng → người lưu sau ghi đè ô người khác vừa sửa. Nay form chỉ gửi ô đổi (`chg`) + giá trị lúc mở (`goc`); máy chủ báo xung đột, trang hỏi “ghi đè / lấy bản mới”. Hàng đợi mất sóng cũng mang `goc` → xung đột thì báo bị từ chối.
- **Nhóm dữ liệu**: 👥 Người dùng có ô “Được sửa nhóm dữ liệu”. Form khóa ô ngoài nhóm, ẩn nút Xóa/Đính kèm/Thêm hộ tương ứng; máy chủ chặn cả khi gọi thẳng API (đã test). `bulk` nay kiểm cả tổ cũ của hộ.
- **Tốc độ** (3.000 hộ, máy tính): mở trang 2,5 s → 0,3 s; bỏ lọc 2 s → 0,24 s; phần tử trang 118.000 → ~4.800. Gõ tìm trễ 180 ms mới lọc.
- **Điện thoại** (≤ 700 px): thanh trên 1 dòng; hero chỉ Quét QR / Thêm hộ / Đọc lại + ⋯; thẻ số 2 cột; 3 khung gập sẵn; ô tìm + “Bộ lọc ▾” ghim khi cuộn; thẻ hộ gọn, bấm “chi tiết · thao tác” mới hiện. Trang dài 14.700 px → ~8.900 px.
- **Kiểm tra dữ liệu**: thẻ “hộ cần kiểm tra dữ liệu” + dòng ⚠ trên hộ + cột “Cần kiểm tra” khi xuất Excel. Luật: trùng tên + tờ/thửa, thiếu DT, tổng DT thu hồi vượt DT thửa, tờ/thửa không có trên bản vẽ, đã chi > tổng PA, chi tiền khi chưa phê duyệt, bàn giao chưa có ngày chi trả, ngày các bước lùi, có file TB chưa ghi số.
- `Code.gs` đổi → dán + Phiên bản mới + **`caiDat()`** (thêm cột “Được sửa (nhóm)” tab NguoiDung; người cũ để trống = mọi nhóm như trước).

### Việc người dùng phải làm sau đợt 27/11/2026 (1 lần)
1. Apps Script: dán đè **Code.gs** mới → Triển khai → Quản lý triển khai → ✎ → **Phiên bản mới**.
2. Chạy **`caiDat()`** (thêm cột tiến độ + tiền + “Được sửa (nhóm)”) rồi **`caiTuDong()`** (sao lưu đêm + email đề nghị; cho phép quyền Drive, Gmail).
3. Quản trị: 👥 Người dùng → giao **nhóm dữ liệu** cho từng người (vd kế toán chỉ “Kinh phí, chi trả”). ⚙ Cấu hình → đặt **Hạn bàn giao mặt bằng**, nạp **mẫu văn bản khác** nếu cần, kiểm **Câu chữ mẫu thông báo** nếu mẫu TB dự án khác chuẩn.
