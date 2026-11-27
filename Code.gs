/**
 * MÁY CHỦ cho trang "Danh sách hộ dân · Bản đồ dự án" — Google Apps Script.
 * Dữ liệu: Google Sheet (tab "Tổng" = hộ, "NguoiDung" = tài khoản, "DeNghi" = đề nghị chờ duyệt, "NhatKy" = lịch sử).
 * File thông báo, bản vẽ, mẫu Word: thư mục Google Drive.
 *
 * CÀI ĐẶT (1 lần):
 *  1. Mở Google Sheet của dự án → Tiện ích mở rộng → Apps Script → xóa hết, dán file này.
 *  2. Sửa THU_MUC_ID (ID thư mục Drive) và MK_ADMIN_DAU (mật khẩu tạm của tài khoản admin).
 *  3. Chạy hàm caiDat() một lần (cho phép quyền) → tạo các tab + tài khoản "admin".
 *  4. Triển khai → Tùy chọn triển khai mới → Ứng dụng web · Thực thi: Tôi · Truy cập: Bất kỳ ai → copy URL /exec.
 *  5. Trên trang: ⚙ Cấu hình → dán URL → đăng nhập admin → ĐỔI MẬT KHẨU → thêm người dùng ở nút "👥 Người dùng".
 *  Sửa code: Triển khai → Quản lý triển khai → ✎ → Phiên bản mới (URL giữ nguyên).
 *
 * QUYỀN (cột "Quyền" trong tab NguoiDung):
 *   quantri  — mọi việc, duyệt đề nghị, quản lý người dùng, nạp bản vẽ / mẫu thông báo
 *   sua      — thêm / sửa / xóa hộ, đính kèm / gỡ / tạo thông báo: có hiệu lực ngay
 *   denghi   — như "sua" nhưng thành ĐỀ NGHỊ, quản trị duyệt rồi mới có hiệu lực
 *   xem      — chỉ xem, tải, in
 * PHẠM VI (cột "Phạm vi tổ"): * = mọi tổ; hoặc "Tổ 1, Tổ 3" = chỉ thao tác hộ của các tổ đó (vẫn xem được toàn bộ).
 * Máy chủ kiểm quyền ở mọi yêu cầu — sửa trang web cũng không vượt quyền được.
 */
const THU_MUC_ID = '';                 // ID thư mục Drive lưu thông báo (trống = thư mục gốc Drive)
const MK_ADMIN_DAU = 'Doi-mat-khau-ngay-1';   // mật khẩu tạm cho tài khoản admin khi chạy caiDat()
const TEN_TAB = 'Tổng';
const PHIEN_GIO = 12;                  // phiên đăng nhập kéo dài bao nhiêu giờ

const COT = [
  ['stt', 'STT'], ['ma', 'Mã số hồ sơ'], ['canbo', 'Cán bộ thụ lý'], ['ten', 'Họ và tên'],
  ['diachi', 'Địa chỉ thu hồi'], ['phuong', 'Phường (cũ)'], ['kp', 'Khu phố'], ['to', 'Tổ NVBT'],
  ['to_bd', 'Số tờ'], ['thua', 'Số thửa'], ['dt', 'Diện tích thu hồi (m²)'], ['giaitoa', 'Giải tỏa'],
  ['tb_so', 'Số thông báo'], ['tb_ngay', 'Ngày thông báo'], ['tb_file', 'File thông báo (Drive ID)'],
  ['tb_ten', 'Tên file thông báo'], ['anh', 'Số ảnh hiện trạng'], ['ht', 'Link hồ sơ hiện trạng'],
  ['ht_file', 'File ảnh hiện trạng (Drive ID)'], ['ht_ten', 'Tên file ảnh hiện trạng'],
  ['tien_bt', 'Tiền bồi thường (đ)'], ['tien_ht', 'Tiền hỗ trợ (đ)'], ['tien_tdc', 'Tiền tái định cư (đ)'], ['da_chi', 'Đã chi trả (đ)'],
  ['gd', 'Tiến độ GPMB (bước đã xong)'], ['gd_ngay', 'Ngày các bước GPMB (JSON)'],
  ['duyet', 'Chờ duyệt'], ['ghichu', 'Ghi chú'], ['by', 'Người sửa cuối'], ['at', 'Lúc sửa cuối'], ['id', 'id']
];
const DATA_KEYS = COT.map(c => c[0]).filter(k => ['id', 'by', 'at', 'tb_file', 'tb_ten', 'ht_file', 'ht_ten'].indexOf(k) < 0);
const U_COT = ['Tên đăng nhập', 'Họ tên', 'Quyền', 'Phạm vi tổ', 'Mật khẩu (mã hóa)', 'Muối', 'Khóa', 'Đăng nhập lần cuối'];
const DN_COT = ['id', 'Loại', 'Hộ id', 'Nội dung', 'Người đề nghị', 'Lúc', 'Trạng thái', 'Người duyệt', 'Lúc duyệt', 'Lý do'];
const NK_COT = ['Lúc', 'Người', 'Việc', 'Hộ id', 'Chi tiết'];
const QUYEN = ['quantri', 'sua', 'denghi', 'xem'];
const LOAI_TEN = { sua: 'sửa hộ', them: 'thêm hộ', xoa: 'xóa hộ', kem: 'đính kèm thông báo', go: 'gỡ thông báo', kemht: 'đính kèm ảnh hiện trạng', goht: 'gỡ ảnh hiện trạng' };

/* ================================================================== */
function caiDat() {
  const ss = SpreadsheetApp.getActive();
  const sh = tab_(TEN_TAB);
  const head = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
  COT.forEach(([, t]) => { if (head.indexOf(t) < 0) head.push(t); });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1);
  const c = cols_(sh), n = sh.getLastRow() - 1;
  if (n > 0) { const r = sh.getRange(2, c.id + 1, n, 1), v = r.getValues(); v.forEach(x => { if (!x[0]) x[0] = uid_(); }); r.setValues(v); }
  [['NguoiDung', U_COT], ['DeNghi', DN_COT], ['NhatKy', NK_COT]].forEach(([t, h]) => {
    const s = tab_(t); if (!s.getLastRow()) { s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold'); s.setFrozenRows(1); }
  });
  if (!users_().some(u => u.ten === 'admin')) {
    const salt = uid_();
    tab_('NguoiDung').appendRow(['admin', 'Quản trị hệ thống', 'quantri', '*', hash_(MK_ADMIN_DAU, salt), salt, '', '']);
  }
  Logger.log('Xong. Đăng nhập: admin / ' + MK_ADMIN_DAU + ' — đổi mật khẩu ngay sau lần đầu.');
}

/* sửa tay trên Sheet → báo các máy đang mở trang tự đọc lại (trigger đơn giản, không cần cài) */
function onEdit() { CacheService.getScriptCache().put('ver', String(Date.now()), 21600); }

function doGet() { return out_({ ok: true, msg: 'Máy chủ danh sách hộ đang chạy. Trang dùng POST.' }); }

function doPost(e) {
  let q;
  try { q = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Dữ liệu gửi lên lỗi' }); }
  try {
    if (q.action === 'login') return out_(login_(q.user, q.pass));
    const me = auth_(q.token);
    if (!me) return out_({ ok: false, error: 'Chưa đăng nhập hoặc hết phiên', auth: true });
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { return out_(route_(me, q)); } finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function route_(me, q) {
  const A = q.action;
  switch (A) {
    case 'me': return { ok: true, me: pub_(me) };
    case 'logout': CacheService.getScriptCache().remove('tk_' + q.token); return { ok: true };
    case 'changePw': return changePw_(me, q.old, q.pass);
    case 'list': { const rows = list_(); return { ok: true, rows, cho: pending_(me, rows), me: pub_(me), cfg: cfg_(), sheetUrl: me.q === 'quantri' ? SpreadsheetApp.getActive().getUrl() : '' }; }
    case 'ping': return { ok: true, v: CacheService.getScriptCache().get('ver') || '0' };
    case 'putCfg': need_(me, 'quantri'); PropertiesService.getScriptProperties().setProperty('cfg', JSON.stringify(q.cfg || {})); log_(me, 'sửa cấu hình dự án', '', JSON.stringify(q.cfg)); return { ok: true };
    case 'capMa': need_(me, 'sua'); return { ok: true, n: capMa_(me, q.duoi) };
    case 'getFile': return { ok: true, b64: Utilities.base64Encode(DriveApp.getFileById(fileOk_(q.id)).getBlob().getBytes()) };
    case 'getMap': { const f = file_('bando.json'); return { ok: true, map: f ? JSON.parse(f.getBlob().getDataAsString()) : null }; }
    case 'getMau': { const f = file_('mau-thong-bao.docx'); return { ok: true, b64: f ? Utilities.base64Encode(f.getBlob().getBytes()) : '' }; }
    case 'log': need_(me, 'quantri'); return { ok: true, log: nhatKy_(q.hoId) };
    /* ---- thay đổi dữ liệu hộ ---- */
    case 'upsert': return change_(me, q.row.id && rowIndex_(q.row.id) > 0 ? 'sua' : 'them', q.row);
    case 'bulk': need_(me, 'sua'); (q.rows || []).forEach(r => { scope_(me, r.to); upsert_(r, me); }); log_(me, 'nhập nhiều', '', (q.rows || []).length + ' hộ'); return { ok: true };
    case 'replaceAll': need_(me, 'quantri'); replaceAll_(q.rows || [], me); log_(me, 'thay toàn bộ danh sách', '', (q.rows || []).length + ' hộ'); return { ok: true };
    case 'delete': return change_(me, 'xoa', { id: q.id, lydo: q.lydo || '' });
    case 'uploadTB': return uploadTB_(me, q);
    case 'delTB': return change_(me, 'go', { id: q.id });
    case 'delHT': return change_(me, 'goht', { id: q.id });
    /* ---- đề nghị ---- */
    case 'approve': need_(me, 'quantri'); return { ok: true, n: decide_(me, q.ids || [q.id], true) };
    case 'reject': need_(me, 'quantri'); return { ok: true, n: decide_(me, q.ids || [q.id], false, q.lydo) };
    case 'withdraw': return withdraw_(me, q.id);
    /* ---- quản trị ---- */
    case 'putMap': need_(me, 'quantri'); saveFile_('bando.json', Utilities.newBlob(JSON.stringify(q.map), 'application/json', 'bando.json')); log_(me, 'nạp bản vẽ', '', (q.map.thua || []).length + ' thửa'); return { ok: true };
    case 'putMau': need_(me, 'quantri'); saveFile_('mau-thong-bao.docx', Utilities.newBlob(Utilities.base64Decode(q.b64), null, 'mau-thong-bao.docx')); log_(me, 'nạp mẫu thông báo', '', ''); return { ok: true };
    case 'users': need_(me, 'quantri'); return { ok: true, users: users_().map(pub_) };
    case 'userSave': need_(me, 'quantri'); return userSave_(me, q.u);
    case 'userDel': need_(me, 'quantri'); return userDel_(me, q.ten);
    default: return { ok: false, error: 'Không rõ yêu cầu ' + A };
  }
}

/* ================= TÀI KHOẢN ================= */
function hash_(pw, salt) {
  let b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  for (let i = 0; i < 500; i++) b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, b.concat(Utilities.newBlob(salt).getBytes()));
  return Utilities.base64Encode(b);
}
function users_() {
  const s = tab_('NguoiDung'); const n = s.getLastRow() - 1; if (n < 1) return [];
  return s.getRange(2, 1, n, U_COT.length).getValues().map((r, i) => ({
    row: i + 2, ten: String(r[0]).trim(), hoten: String(r[1]), q: String(r[2]).trim(), to: String(r[3] || '*').trim(),
    hash: String(r[4]), salt: String(r[5]), khoa: !!String(r[6]).trim(), last: r[7]
  })).filter(u => u.ten);
}
function pub_(u) { return { ten: u.ten, hoten: u.hoten, q: u.q, to: u.to, khoa: u.khoa, last: u.last ? String(u.last) : '' }; }
function login_(ten, pass) {
  const cache = CacheService.getScriptCache(), k = 'sai_' + String(ten).toLowerCase(), sai = +(cache.get(k) || 0);
  if (sai >= 5) return { ok: false, error: 'Sai mật khẩu quá 5 lần — thử lại sau 15 phút' };
  const u = users_().find(x => x.ten.toLowerCase() === String(ten || '').trim().toLowerCase());
  if (!u || u.khoa || hash_(String(pass || ''), u.salt) !== u.hash) {
    cache.put(k, String(sai + 1), 900);
    return { ok: false, error: u && u.khoa ? 'Tài khoản đang bị khóa' : 'Sai tên đăng nhập hoặc mật khẩu' };
  }
  cache.remove(k);
  const token = uid_() + uid_() + uid_();
  cache.put('tk_' + token, u.ten, PHIEN_GIO * 3600 > 21600 ? 21600 : PHIEN_GIO * 3600);   // CacheService tối đa 6 giờ
  tab_('NguoiDung').getRange(u.row, 8).setValue(new Date());
  log_(u, 'đăng nhập', '', '');
  return { ok: true, token, me: pub_(u) };
}
function auth_(token) {
  if (!token) return null;
  const ten = CacheService.getScriptCache().get('tk_' + token); if (!ten) return null;
  const u = users_().find(x => x.ten === ten);
  return u && !u.khoa ? u : null;
}
function changePw_(me, old, pw) {
  if (hash_(String(old || ''), me.salt) !== me.hash) return { ok: false, error: 'Mật khẩu cũ không đúng' };
  if (String(pw || '').length < 8) return { ok: false, error: 'Mật khẩu mới tối thiểu 8 ký tự' };
  const salt = uid_(); tab_('NguoiDung').getRange(me.row, 5, 1, 2).setValues([[hash_(pw, salt), salt]]);
  log_(me, 'đổi mật khẩu', '', ''); return { ok: true };
}
function userSave_(me, u) {
  const ten = String(u.ten || '').trim();
  if (!/^[a-z0-9._-]{3,30}$/i.test(ten)) return { ok: false, error: 'Tên đăng nhập 3–30 ký tự: chữ không dấu, số, . _ -' };
  if (QUYEN.indexOf(u.q) < 0) return { ok: false, error: 'Quyền không hợp lệ' };
  const all = users_(), cu = all.find(x => x.ten.toLowerCase() === ten.toLowerCase());
  if (cu && cu.ten === me.ten && (u.q !== 'quantri' || u.khoa)) return { ok: false, error: 'Không tự hạ quyền hoặc tự khóa chính mình' };
  const s = tab_('NguoiDung');
  if (!cu && String(u.pass || '').length < 8) return { ok: false, error: 'Tài khoản mới cần mật khẩu tối thiểu 8 ký tự' };
  const row = cu ? cu.row : s.getLastRow() + 1;
  s.getRange(row, 1, 1, 4).setValues([[ten, u.hoten || '', u.q, u.to || '*']]);
  s.getRange(row, 7).setValue(u.khoa ? 'x' : '');
  if (u.pass) { const salt = uid_(); s.getRange(row, 5, 1, 2).setValues([[hash_(u.pass, salt), salt]]); }
  log_(me, cu ? 'sửa người dùng' : 'thêm người dùng', '', ten + ' · ' + u.q + ' · ' + (u.to || '*') + (u.pass ? ' · đặt mật khẩu' : '') + (u.khoa ? ' · KHÓA' : ''));
  return { ok: true };
}
function userDel_(me, ten) {
  if (ten === me.ten) return { ok: false, error: 'Không tự xóa chính mình' };
  const u = users_().find(x => x.ten === ten); if (!u) return { ok: false, error: 'Không thấy tài khoản' };
  tab_('NguoiDung').deleteRow(u.row); log_(me, 'xóa người dùng', '', ten); return { ok: true };
}
const RANK = { quantri: 3, sua: 2, denghi: 1, xem: 0 };
function need_(me, q) { if ((RANK[me.q] || 0) < RANK[q]) throw new Error('Tài khoản ' + me.ten + ' không có quyền này'); }
function inScope_(me, to) {
  if (me.q === 'quantri' || !me.to || me.to === '*') return true;
  return me.to.split(/[,;]/).map(s => s.trim()).indexOf(String(to || '').trim()) >= 0;
}
function scope_(me, to) { if (!inScope_(me, to)) throw new Error('Hộ thuộc ' + (to || 'chưa phân tổ') + ' — ngoài phạm vi của ' + me.ten + ' (' + me.to + ')'); }

/* ================= HỘ DÂN ================= */
function tab_(t) { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(t) || ss.insertSheet(t); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function cols_(sh) {
  sh = sh || tab_(TEN_TAB);
  const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(s => String(s).trim());
  const c = {}; COT.forEach(([k, t]) => { const i = head.indexOf(t); if (i >= 0) c[k] = i; });
  if (c.id === undefined) throw new Error('Sheet chưa có cột "id" — chạy hàm caiDat() trước');
  c._n = head.length; return c;
}
const val_ = v => v instanceof Date ? Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd') : (v === null || v === undefined ? '' : String(v));
function list_() {
  const sh = tab_(TEN_TAB), c = cols_(sh), n = sh.getLastRow() - 1; if (n < 1) return [];
  return sh.getRange(2, 1, n, c._n).getValues().filter(r => r.join('').trim()).map(r => {
    const o = {}; COT.forEach(([k]) => { o[k] = c[k] === undefined ? '' : val_(r[c[k]]); }); return o;
  }).filter(o => o.id);
}
function rowIndex_(id) {
  const sh = tab_(TEN_TAB), c = cols_(sh), n = sh.getLastRow() - 1; if (n < 1 || !id) return -1;
  const ids = sh.getRange(2, c.id + 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function getRow_(id) { return list_().find(r => String(r.id) === String(id)) || null; }
function upsert_(row, me) {
  const sh = tab_(TEN_TAB), c = cols_(sh);
  if (!row.id) row.id = uid_();
  row.by = me ? (me.hoten || me.ten) : ''; row.at = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm');
  const at = rowIndex_(row.id);
  const line = at > 0 ? sh.getRange(at, 1, 1, c._n).getValues()[0] : new Array(c._n).fill('');
  COT.forEach(([k]) => { if (c[k] !== undefined && row[k] !== undefined) line[c[k]] = row[k]; });
  if (at > 0) sh.getRange(at, 1, 1, c._n).setValues([line]); else sh.appendRow(line);
  return row.id;
}
function replaceAll_(rows, me) {
  const sh = tab_(TEN_TAB), c = cols_(sh);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, c._n).clearContent();
  if (!rows.length) return;
  rows.forEach(r => { if (!r.id) r.id = uid_(); });
  sh.getRange(2, 1, rows.length, c._n).setValues(rows.map(r => { const l = new Array(c._n).fill(''); COT.forEach(([k]) => { if (c[k] !== undefined) l[c[k]] = r[k] === undefined ? '' : r[k]; }); return l; }));
}
/* sửa / thêm / xóa / gỡ TB: quyền "sua"+"quantri" có hiệu lực ngay, "denghi" thành đề nghị */
function change_(me, loai, data) {
  need_(me, 'denghi');
  const cu = data.id ? getRow_(data.id) : null;
  if (loai === 'sua' || loai === 'xoa' || loai === 'go' || loai === 'goht') { if (!cu) throw new Error('Không thấy hộ'); scope_(me, cu.to); }
  if (loai === 'them') scope_(me, data.to);
  if (loai === 'sua' && data.to !== undefined) scope_(me, data.to);   // đổi sang tổ khác cũng phải trong phạm vi
  if (loai === 'sua') {   // chỉ giữ trường có thay đổi
    const d = { id: data.id }; let co = false;
    DATA_KEYS.forEach(k => { if (data[k] !== undefined && String(data[k]) !== String(cu[k])) { d[k] = data[k]; co = true; } });
    if (!co) return { ok: true, id: data.id, nochange: true };
    data = d;
  }
  if (me.q === 'denghi') {
    const id = propose_(me, loai, loai === 'them' ? (data.id = data.id || uid_()) : data.id, data);
    return { ok: true, pending: id, id: data.id };
  }
  apply_(me, loai, data.id, data);
  return { ok: true, id: data.id };
}
function apply_(me, loai, hoId, d) {
  if (loai === 'them' || loai === 'sua') {
    const ct = loai === 'sua' ? JSON.stringify(diffOnly_(d)) : d.ten;
    if (loai === 'sua') { const cu = getRow_(hoId); if (!cu) throw new Error('Hộ đã bị xóa'); d = Object.assign(cu, d); }
    upsert_(d, me); log_(me, loai === 'them' ? 'thêm hộ' : 'sửa hộ', d.id, ct);
  } else if (loai === 'xoa') {
    const at = rowIndex_(hoId); if (at > 0) tab_(TEN_TAB).deleteRow(at);
    log_(me, 'xóa hộ', hoId, d.lydo || '');
  } else if (loai === 'kem' || loai === 'kemht') {
    const [fF, fT] = loai === 'kem' ? ['tb_file', 'tb_ten'] : ['ht_file', 'ht_ten'];
    const cu = getRow_(hoId); if (!cu) throw new Error('Hộ đã bị xóa');
    if (cu[fF] && cu[fF] !== d.fileId) { try { DriveApp.getFileById(cu[fF]).setTrashed(true); } catch (e) {} }
    const o = { id: hoId }; o[fF] = d.fileId; o[fT] = d.name; if (loai === 'kemht' && d.anh) o.anh = d.anh;
    upsert_(o, me); log_(me, LOAI_TEN[loai], hoId, d.name);
  } else if (loai === 'go' || loai === 'goht') {
    const [fF, fT] = loai === 'go' ? ['tb_file', 'tb_ten'] : ['ht_file', 'ht_ten'];
    const cu = getRow_(hoId); if (cu && cu[fF]) { try { DriveApp.getFileById(cu[fF]).setTrashed(true); } catch (e) {} }
    const o = { id: hoId }; o[fF] = ''; o[fT] = ''; upsert_(o, me); log_(me, LOAI_TEN[loai], hoId, cu ? cu[fT] : '');
  }
}
function diffOnly_(d) { const o = {}; Object.keys(d).forEach(k => { if (DATA_KEYS.indexOf(k) >= 0) o[k] = d[k]; }); return o; }
function uploadTB_(me, q) {
  need_(me, 'denghi');
  const cu = getRow_(q.id); if (!cu) throw new Error('Không thấy hộ'); scope_(me, cu.to);
  const ht = q.kind === 'ht', loai = ht ? 'kemht' : 'kem';
  const f = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(q.b64), ht ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', q.name || (ht ? 'hien-trang.pdf' : 'thong-bao.docx')));
  const d = { fileId: f.getId(), name: q.name, anh: q.anh || '' };
  if (me.q === 'denghi') return { ok: true, pending: propose_(me, loai, q.id, d), fileId: f.getId() };
  apply_(me, loai, q.id, d);
  return { ok: true, fileId: f.getId() };
}

/* ================= ĐỀ NGHỊ CHỜ DUYỆT ================= */
function dnRows_() {
  const s = tab_('DeNghi'), n = s.getLastRow() - 1; if (n < 1) return [];
  return s.getRange(2, 1, n, DN_COT.length).getValues().map((r, i) => ({
    row: i + 2, id: String(r[0]), loai: String(r[1]), hoId: String(r[2]), data: r[3] ? JSON.parse(r[3]) : {},
    by: String(r[4]), at: val2_(r[5]), st: String(r[6]), dby: String(r[7]), dat: val2_(r[8]), lydo: String(r[9])
  }));
}
const val2_ = v => v instanceof Date ? Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm') : String(v || '');
function propose_(me, loai, hoId, data) {
  // đề nghị sửa trùng hộ của cùng người → gộp vào đề nghị cũ
  if (loai === 'sua') {
    const old = dnRows_().find(d => d.st === 'cho' && d.loai === 'sua' && d.hoId === hoId && d.by === me.ten);
    if (old) { const m = Object.assign(old.data, data); tab_('DeNghi').getRange(old.row, 4, 1, 3).setValues([[JSON.stringify(m), me.ten, new Date()]]); return old.id; }
  }
  const id = uid_();
  tab_('DeNghi').appendRow([id, loai, hoId, JSON.stringify(data), me.ten, new Date(), 'cho', '', '', data.lydo || '']);
  log_(me, 'gửi đề nghị ' + (LOAI_TEN[loai] || loai), hoId, JSON.stringify(data).slice(0, 300));
  return id;
}
function pending_(me, rows) {
  const by = {}; (rows || list_()).forEach(r => { by[r.id] = r; });
  return dnRows_().filter(d => d.st === 'cho' && (me.q === 'quantri' || d.by === me.ten || inScope_(me, (by[d.hoId] || d.data).to)))
    .map(d => ({ id: d.id, loai: d.loai, hoId: d.hoId, data: d.data, by: d.by, at: d.at, lydo: d.lydo }));
}
function decide_(me, ids, ok, lydo) {
  const s = tab_('DeNghi'); let n = 0;
  dnRows_().filter(d => ids.indexOf(d.id) >= 0 && d.st === 'cho').forEach(d => {
    if (ok) { try { apply_({ ten: d.by, hoten: d.by + ' (duyệt: ' + me.ten + ')' }, d.loai, d.hoId, d.data); } catch (e) { s.getRange(d.row, 7, 1, 4).setValues([['loi', me.ten, new Date(), String(e.message)]]); return; } }
    else if ((d.loai === 'kem' || d.loai === 'kemht') && d.data.fileId) { try { DriveApp.getFileById(d.data.fileId).setTrashed(true); } catch (e) {} }
    s.getRange(d.row, 7, 1, 4).setValues([[ok ? 'duyet' : 'tuchoi', me.ten, new Date(), ok ? d.lydo : (lydo || '')]]);
    log_(me, (ok ? 'duyệt ' : 'từ chối ') + (LOAI_TEN[d.loai] || d.loai) + ' của ' + d.by, d.hoId, lydo || '');
    n++;
  });
  return n;
}
function withdraw_(me, id) {
  const d = dnRows_().find(x => x.id === id && x.st === 'cho'); if (!d) return { ok: false, error: 'Không thấy đề nghị' };
  if (d.by !== me.ten && me.q !== 'quantri') return { ok: false, error: 'Chỉ người gửi mới rút được' };
  if ((d.loai === 'kem' || d.loai === 'kemht') && d.data.fileId) { try { DriveApp.getFileById(d.data.fileId).setTrashed(true); } catch (e) {} }
  tab_('DeNghi').getRange(d.row, 7, 1, 3).setValues([['rut', me.ten, new Date()]]);
  return { ok: true };
}

/* ================= NHẬT KÝ & DRIVE ================= */
function log_(me, viec, hoId, ct) {
  tab_('NhatKy').appendRow([new Date(), me.hoten || me.ten, viec, hoId || '', ct || '']);
  if (viec !== 'đăng nhập') CacheService.getScriptCache().put('ver', String(Date.now()), 21600);   // báo các máy đang mở: có thay đổi
}
/* cấu hình chung của dự án (tên dự án, đơn vị, đuôi mã, ngưỡng DT còn lại) — quản trị đặt, mọi người dùng chung */
function cfg_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('cfg') || '{}'); } catch (e) { return {}; } }
/* MÃ SỐ HỒ SƠ cấp ở máy chủ (trong khóa) — không trùng khi nhiều người mở cùng lúc: ###/P{phường}/{đuôi}.{số tổ} */
function capMa_(me, duoi) {
  const sh = tab_(TEN_TAB), c = cols_(sh), n = sh.getLastRow() - 1; if (n < 1 || c.ma === undefined) return 0;
  const v = sh.getRange(2, 1, n, c._n).getValues();
  let mx = v.reduce((m, r) => Math.max(m, parseInt(r[c.ma], 10) || 0), 0), k = 0;
  duoi = duoi || cfg_().maDuoi || 'DA';
  v.forEach((r, i) => {
    const to = String(c.to === undefined ? '' : r[c.to]).trim();
    if (!String(r[c.ma]).trim() && to && String(r[c.id]).trim()) {
      mx++; k++;
      sh.getRange(i + 2, c.ma + 1).setValue(String(mx).padStart(3, '0') + '/P' + String(c.phuong === undefined ? '' : r[c.phuong]).trim() + '/' + duoi + '.' + ((to.match(/\d+/) || [''])[0]));
    }
  });
  if (k) log_(me, 'cấp mã hồ sơ', '', k + ' hộ');
  return k;
}
function nhatKy_(hoId) {
  const s = tab_('NhatKy'), n = s.getLastRow() - 1; if (n < 1) return [];
  const from = Math.max(2, s.getLastRow() - 2000 + 1);
  return s.getRange(from, 1, s.getLastRow() - from + 1, 5).getValues()
    .filter(r => !hoId || String(r[3]) === String(hoId)).slice(-300).reverse()
    .map(r => ({ at: val2_(r[0]), by: String(r[1]), viec: String(r[2]), hoId: String(r[3]), ct: String(r[4]) }));
}
function folder_() { return THU_MUC_ID ? DriveApp.getFolderById(THU_MUC_ID) : DriveApp.getRootFolder(); }
function file_(name) { const it = folder_().getFilesByName(name); return it.hasNext() ? it.next() : null; }
function saveFile_(name, blob) { const f = file_(name); if (f) f.setTrashed(true); folder_().createFile(blob.setName(name)); }
/* chỉ cho tải file thông báo đang gắn với hộ hoặc đề nghị — không cho đọc file Drive bất kỳ */
function fileOk_(id) {
  if (list_().some(r => r.tb_file === id || r.ht_file === id) || dnRows_().some(d => d.st === 'cho' && d.data.fileId === id)) return id;
  throw new Error('File không thuộc danh sách này');
}
