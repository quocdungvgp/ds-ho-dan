"""Tạo file mẫu để chạy thử: mau.docx (mẫu TB có «»), tb_that.docx (TB lệch web, chữ tách nhiều run),
ds.xlsx / ds2.xlsx (danh sách hộ), ht.pdf (3 trang ảnh), tb.zip (đính kèm hàng loạt). Cần: python-docx openpyxl pillow"""
import os, zipfile, shutil
from docx import Document
import openpyxl
from PIL import Image, ImageDraw
D = os.path.join(os.path.dirname(__file__), 'fx'); os.makedirs(D, exist_ok=True)
d = Document(); d.add_paragraph("THÔNG BÁO THU HỒI ĐẤT"); d.add_paragraph("Thu hồi đất của «HOTEN».")
d.add_paragraph("Diện tích đất dự kiến thu hồi: «DT» m², thuộc «MUC» thửa đất số «THUA», tờ bản đồ số «TO». STT «STT»"); d.save(f'{D}/mau.docx')
d = Document(); d.add_paragraph("PHIẾU CHI TIỀN BỒI THƯỜNG"); d.add_paragraph("Chi cho «HOTEN», thửa «THUA» tờ «TO». Kiểm đếm ngày «NGAY_KD»."); d.add_paragraph("Số tiền: «TONG» đồng (Bằng chữ: «TONG_CHU»)."); d.save(f'{D}/mau_vb.docx')
d = Document(); p = d.add_paragraph(); p.add_run("Thu hồi đất của "); r = p.add_run("Ông/Bà Nguyễn Văn "); r.bold = True; r = p.add_run("Ann."); r.bold = True
p = d.add_paragraph(); p.add_run("Diện tích đất dự kiến thu hồi: "); p.add_run("20"); p.add_run(",5 m², thuộc một phần thửa đất số 1, tờ bản đồ số 134."); d.save(f'{D}/tb_that.docx')
H = ['STT', 'Họ và tên', 'Địa chỉ thu hồi', 'Phường', 'Khu phố', 'Tổ', 'Số tờ', 'Số thửa', 'Diện tích thu hồi', 'Giải tỏa']
for name, rows in [('ds.xlsx', [[1, 'Nguyễn Test', '1 Đường A', '19', '4', 2, '134', '3', '12,5', 'một phần'], [2, 'Trần Test', '2 Đường A', '19', '4', 2, '134', '4', 20, 'Toàn phần']]),
                   ('ds2.xlsx', [[1, 'Nguyễn Văn An', '10 Đường A', 19, 4, 1, 134, 1, '21,5', 'một phần'], [2, 'Trần Thị Bích', '12 Đường A', 19, 4, 2, 134, 2, 30, 'toàn phần'], [3, 'Lê Văn Mới', '16 Đường A', 19, 4, 2, 134, 5, 12, 'một phần']])]:
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Tổng'; ws.append(H); [ws.append(r) for r in rows]; wb.save(f'{D}/{name}')
# có cột tiến độ GPMB: bước ghi bằng số (3 = Lập phương án) + ngày chi trả dạng chữ dd/mm/yyyy → tự lên "Chi trả"
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Tổng'; ws.append(H + ['Tiến độ GPMB', 'Ngày kiểm đếm', 'Ngày chi trả'])
ws.append([1, 'Phạm Tiến Độ', '3 Đường A', '19', '4', 3, '135', '2', 15, 'một phần', 3, '01/09/2026', '15/10/2026']); wb.save(f'{D}/ds_gd.xlsx')
ims = []
for i in range(3):
    im = Image.new('RGB', (800, 600), (40 + i * 60, 120, 160)); ImageDraw.Draw(im).text((40, 40), f"Anh {i+1}", fill=(255, 255, 255)); ims.append(im)
ims[0].save(f'{D}/ht.pdf', save_all=True, append_images=ims[1:])
with zipfile.ZipFile(f'{D}/tb.zip', 'w') as z:
    z.write(f'{D}/mau.docx', 'TB/TBTHD_STT002_Tran Thi Bich.docx'); z.write(f'{D}/mau.docx', 'TB/TBTHD_STT099_X.docx')
shutil.copy(f'{D}/mau.docx', f'{D}/TBTHD_STT001_Nguyen Van An.docx')
print('Đã tạo', D)
