# PHÁT TRIỂN TIẾP — ds-ho-dan

> **Đọc file này trước khi sửa code.** Đủ để làm tiếp mà không phải đọc lại cả `index.html` (~1.180 dòng).
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
dxf2bando.py: bản vẽ CAD (DXF, VN2000) → bando.json
```

- **2 chế độ**: `CFG.api` trống → chạy một mình trên máy (dữ liệu mẫu / Excel, lưu `localStorage` + IndexedDB `dsho`), `ME` = quản trị.
  Có `CFG.api` → đăng nhập, mọi thứ qua Apps Script; máy chủ kiểm quyền lại.
- **Link dự án**: `…/ds-ho-dan/?api=<URL /exec>` → tự lưu vào cấu hình, xóa tham số khỏi thanh địa chỉ.
- Thư viện tải từ jsDelivr, ghim phiên bản (bảng `LIB` trong mục 1 và `<head>`): leaflet 1.9.4, proj4 2.12.1 (nạp sẵn);
  xlsx 0.18.5, jszip 3.10.1, pizzip 3.1.7, docxtemplater 3.47.4, docx-preview 0.3.2, pdfjs-dist 3.11.174 (nạp khi cần qua `need('xlsx',…)`).
- Hệ tọa độ: `proj4.defs('VN2000HCM', …)` — kinh tuyến trục 105°45′, múi 3°, k=0,9999, 7 tham số towgs84. Tỉnh khác sửa `+lon_0`.

## 2. Bản đồ code `index.html` (tìm bằng tiêu đề mục)

`grep -n "^   [0-9]" index.html` ra danh sách mục:

| Mục | Nội dung | Hàm chính |
|---|---|---|
| 1. TIỆN ÍCH | `$ esc num f1 fold toast download b64 need() idb` | `need(...ks)` nạp thư viện CDN |
| 2. CẤU HÌNH | `CFG0`, `CFG` (localStorage `dsho_cfg`), link `?api=` | `saveCfg linkDA NG()` |
| 3. NGƯỜI DÙNG & QUYỀN | `ME TOKEN RANK has() inScope() canRow() isAdmin()` | |
| 4. DỮ LIỆU | `api() load(nhe) saveRow delRow capMa getTB putTB delTB`, `SHARED` cấu hình chung | `xong(j,viec)` báo + đọc lại |
| 4b. DỮ LIỆU MẪU | `sample() sampleMap()` | |
| 5. LỌC & SỐ LIỆU | `st` (bộ lọc), `PARCEL`, `conLai()`, `filtered()`, đề nghị `CHO choOf virtRows choHtml`; **tiến độ GPMB** `GD GD_COL gdIdx gdOf gdNgay gdMatch gdCell renderGd()` | `gdIdx(r)` = bước đã xong (−1…6) |
| 6. VẼ DANH SÁCH | `render()` (thẻ số, bảng tổ, chip, ô lọc) `tbCell() list() sortRows()` | |
| 7. THÔNG BÁO WORD | `attach xem taoTB taoLoat`, **`TB_FLD docFix()`** tự sửa Word theo web, `dongBo() zipTB hoi()` | |
| 8. EXCEL | `COLS` (khóa, tiêu đề, regex nhận cột; `gdn_<bước>` = cột ngày ảo), `colVal()` giá trị xuất, `importXls exportXls` | |
| 9. BẢN ĐỒ | `initMap CLS(tô màu: tb to gt cl anh gd) drawMap labels pickParcel suggest meas* exportKML`, `HO_IDX hoOf` | |
| 10. SỰ KIỆN | `route()` (#ds/#bando), click `data-a=…`, hộp sửa hộ, cấu hình | |
| 11. ĐĂNG NHẬP… | `topBar showLogin`, người dùng `openUsers userForm`, `duyet drawDuyet openLog` | |
| 12. BỔ SUNG | sắp xếp cột, chép link, tự đọc lại (`ping` 60 s), đính kèm hàng loạt | |
| 13. ẢNH HIỆN TRẠNG | `pdfLib getHT attachHT openAnh drawThumbs bigAnh` | |
| 14. ẢNH THEO NĂM | `loadWayback` (Esri waybackconfig.json + metadata ngày chụp) | |
| 15. SỬA BẢN VẼ | `ED edSet edClick clipHalf edPush edApply`, lưu `putMap` | |

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
| upsert, delete, uploadTB(kind tb/ht), delTB, delHT | denghi+ (trong phạm vi tổ) | quyền `denghi` → thành đề nghị (`propose_`) |
| bulk, capMa | sua+ | |
| replaceAll, approve, reject, users, userSave, userDel, putMap, putMau, putCfg, log | quantri | |
| withdraw | người gửi | |

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
`common.py` chặn CDN → trả từ `tests/node_modules`, chuyển `script.google.com` → máy giả lập. Lần cuối: **42/42 OK** (27/11/2026). mock-gas có lối test `/__het` (xóa Cache) và `/__mail` (chạy `guiBaoDeNghi`, trả thư). Chạy cả bộ: `sh tests/chay.sh`.
Chrome: `/usr/bin/google-chrome` (đổi bằng biến `CHROME`). File tải về trong test không có đuôi → `save_as()` ra file có đuôi rồi mới mở bằng openpyxl.
Chưa thử: Google Apps Script thật, ảnh nền bản đồ (chỉ thử tay trên trang thật — ảnh vệ tinh và Wayback chạy).

## 6. Đẩy lên

Sửa → chạy tests → `git commit` → `git push` (GitHub Pages tự cập nhật). Sửa `Code.gs` thì nhắc người dùng: dán đè trong Apps Script → Triển khai → Quản lý triển khai → ✎ → **Phiên bản mới**; thêm cột thì chạy `caiDat()`.

---

## 7. Việc tiếp theo (đã nghiên cứu, xếp theo giá trị cho công tác bồi thường GPMB)

### Ưu tiên 1
1. ~~**Tiến độ GPMB theo giai đoạn**~~ — **XONG 27/11/2026** (xem mục 3 + Nhật ký làm). Còn có thể làm thêm: hạn riêng từng bước (vd hạn chi trả) + tô đỏ hộ trễ hạn; biểu đồ số hộ xong theo tuần (từ `gd_ngay`); mở rộng tới 23 giai đoạn như thẻ của Ban nếu cần (chỉ sửa mảng `GD`).
2. ~~**Tiền bồi thường – hỗ trợ – TĐC & chi trả**~~ — **XONG 27/11/2026**. Ngày chi dùng chung ngày bước “Chi trả”, không thêm cột `ngay_chi`.
3. **Nhiều mẫu văn bản** (giấy mời, TB kiểm đếm, biên bản, phiếu chi…) thay vì 1 mẫu TB: lưu `mau/<tên>.docx` trên Drive (`putMau` thêm tên), nút “Tạo văn bản” chọn mẫu, dùng chung trường «HOTEN» «DT» «THUA»… + thêm trường tiền, ngày. Tạo hàng loạt → zip.
4. **QR mã hồ sơ**: in nhãn bìa hồ sơ (QR = link `#ho=<id>` mở đúng hộ), quét QR bằng camera điện thoại (thư viện qrcode-generator + jsQR). Ban đang dùng “QR hồ sơ / Quét QR”.
5. **Chụp ảnh hiện trạng thẳng bằng điện thoại**: `<input type=file accept=image/* capture=environment multiple>`, gắn giờ + GPS lên ảnh, gộp thành PDF (jsPDF) rồi dùng luồng `attachHT` sẵn có.
6. ~~**Sao lưu tự động**~~ — **XONG**: `caiTuDong()` → `saoLuu()` mỗi đêm, thư mục “Sao lưu”, giữ `SAO_LUU_GIU`=30.

### Ưu tiên 2
7. ~~**Báo có đề nghị chờ duyệt** qua email~~ — **XONG**: `propose_` đặt cờ `dn_moi`, trigger `guiBaoDeNghi()` 15 phút gửi 1 thư tới `EMAIL_QT` (trống = chủ script).
8. **Xuất ngược DXF** bản vẽ đã tách/gộp/sửa nhãn (viết DXF ASCII: LWPOLYLINE lớp THUA_DAT + TEXT lớp SO_THUA, cộng lại gốc `o`) để về AutoCAD.
9. **In bản đồ A3 có khung tên** (tên dự án, chú giải, mũi tên Bắc, thước tỉ lệ, ngày in) — hiện chỉ in khung bản đồ.
10. **Gộp hình học thật** khi gộp thửa kề (turf.js `union`), hiện chỉ gộp theo nhãn.
11. **Nhập bản vẽ khác DXF**: GeoJSON/KML (đã có VN2GE ra KML), DWG cần ODA File Converter chuyển DXF trước.
12. **Đọc hồ sơ quét bằng AI** → điền thông tin hộ (nối skill nhận định nhà đất Metro 2 / máy AI offline của Ban).
13. **Dùng ngoài thực địa khi mất sóng** (PWA: service worker + hàng đợi sửa, gửi khi có mạng).

### Kỹ thuật / giới hạn đã biết
- `TB_FLD` (mục 7) cứng theo câu chữ mẫu TB thu hồi đất chuẩn → nên đưa vào cấu hình chung để mỗi dự án tự khai.
- ~~Phiên 6 giờ~~ — đã lưu token + hạn vào ScriptProperties (dọn token hết hạn mỗi lần đăng nhập).
- ~~Sheet lớn đọc nhiều lần~~ — `list_()` nhớ `MEMO_` trong 1 yêu cầu (`rowIndex_` vẫn đọc 1 cột id, nhẹ).
- `bando.json` lưu Drive 1 file, đọc qua Apps Script (base64 JSON) — bản vẽ > 5 MB sẽ chậm → nén hoặc chia theo tờ.
- Chế độ một máy: dữ liệu chỉ trong 1 trình duyệt (xóa dữ liệu trình duyệt là mất) → nhắc xuất Excel thường xuyên.
- Ai mở thẳng file Google Sheet sửa được hết, không qua phân quyền → chỉ chia sẻ Sheet cho quản trị.
- Cột “Chờ duyệt” (`duyet`) là cờ tay, khác với đề nghị chờ duyệt (tab DeNghi) — có thể bỏ cột này.

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
