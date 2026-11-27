"""Chạy thử chế độ KHÔNG nối Sheet (dữ liệu mẫu, lưu trong trình duyệt). Chạy: python3 tests/ui_local.py"""
import asyncio, zipfile, re, openpyxl
from common import Run, FX, ok


async def main():
    async with Run() as R:
        pg = await R.page()
        ok(await pg.locator('#tMain tbody tr').count() == 24, 'dữ liệu mẫu 24 hộ')
        # lọc + xem theo thửa
        await pg.click('#tTo tbody tr:nth-child(1) td:nth-child(4) button'); ok('hộ' in await pg.inner_text('#cnt'), 'bấm số trong bảng tổ → lọc')
        await pg.click('#bClear'); await pg.click('#segView button[data-v=ho]')
        # sửa hộ
        await pg.click('#tMain [data-a=sua] >> nth=0'); await pg.fill('#fm [name=dt]', '33,5'); await pg.click('#fm button[value=ok]'); await pg.wait_for_timeout(400)
        ok('33,5' in await pg.inner_text('#tMain tbody tr:nth-child(1)'), 'sửa diện tích')
        # tiến độ GPMB
        ok(await pg.locator('#gdSteps .gds').count() == 7 and await pg.locator('#gdBar button').count() >= 5, 'tiến độ GPMB: 7 bước + thanh tiến độ')
        await pg.click('#bCfg'); await pg.fill('#fCfg [name=hanBG]', '2027-03-31'); await pg.click('#fCfg button[value=ok]'); await pg.wait_for_timeout(400)
        t = await pg.inner_text('#gdHan'); ok('31/03/2027' in t and 'hộ chưa bàn giao' in t, 'hạn bàn giao MB: đếm ngày + số hộ chưa bàn giao')
        await pg.click('#tMain tbody tr:nth-child(1) [data-a=sua]'); await pg.fill('#fm [name=gdn_kd]', '2026-09-01'); await pg.fill('#fm [name=gdn_pd]', '2026-11-20')
        await pg.click('#fm button[value=ok]'); await pg.wait_for_timeout(400)
        ok('Phê duyệt PA · 20/11/2026' in await pg.inner_text('#tMain tbody tr:nth-child(1)'), 'ghi ngày bước → tự tính đã xong bước cao nhất')
        b = pg.locator('#gdSteps .gds >> nth=5').locator('button'); n = re.search(r'\d+', await b.inner_text()).group()
        await b.click(); await pg.wait_for_timeout(300)
        ok(await pg.inner_text('#cnt') == n + ' hộ' and await pg.input_value('#fGd') == 'c5', 'bấm "còn … hộ chưa xong" → lọc đúng số hộ')
        await pg.click('#bClear')
        # kinh phí BT-HT-TĐC
        ok(await pg.locator('#tienCards .card').count() == 4 and 'Tổng cộng' in await pg.inner_text('#tTien'), 'kinh phí: thẻ số + bảng theo tổ')
        await pg.click('#tMain tbody tr:nth-child(1) [data-a=sua]'); await pg.fill('#fm [name=tien_bt]', '1.500.000.000'); await pg.fill('#fm [name=tien_ht]', '200.000.000'); await pg.fill('#fm [name=tien_tdc]', ''); await pg.fill('#fm [name=da_chi]', '850.000.000')
        await pg.click('#fm button[value=ok]'); await pg.wait_for_timeout(400)
        ok('PA 1,7 tỷ · đã chi 50%' in await pg.inner_text('#tMain tbody tr:nth-child(1)'), 'nhập tiền dạng 1.500.000.000 → PA 1,7 tỷ, đã chi 50%')
        await pg.select_option('#fCt', 'thieu'); await pg.wait_for_timeout(200)
        t = await pg.inner_text('#tMain tbody'); ok('Nguyễn Văn An' in t and 'đã chi đủ' not in t, 'lọc "có PA, chưa chi đủ"'); await pg.click('#bClear')
        # mẫu TB + tạo TB + xem trước
        async with pg.expect_file_chooser() as fc:
            await pg.click('#bCfg'); await pg.click('#iMau')
        await (await fc.value).set_files(f'{FX}/mau.docx'); await pg.wait_for_timeout(800)
        await pg.click('#tMain [data-a=tao] >> nth=0'); await pg.wait_for_timeout(2500)
        await pg.click('#tMain [data-a=xem] >> nth=0'); await pg.wait_for_timeout(2500)
        ok('Ông/Bà' in await pg.inner_text('#xemBody'), 'tạo TB theo mẫu + xem trước'); await pg.click('#xemDong')
        # nhiều mẫu văn bản
        async with pg.expect_file_chooser() as fc:
            await pg.click('#bCfg'); await pg.click('#iMauVB')
        await (await fc.value).set_files(f'{FX}/mau_vb.docx'); await pg.wait_for_timeout(500)
        ok(await pg.input_value('dialog[open] input') == 'mau_vb', 'nạp mẫu văn bản: hỏi tên, gợi ý theo tên file')
        await pg.fill('dialog[open] input', 'Phiếu chi'); await pg.click('dialog[open] button[value=ok]'); await pg.wait_for_timeout(600)
        await pg.click('#tMain tbody tr:nth-child(1) [data-a=vb]'); await pg.wait_for_timeout(200)
        async with pg.expect_download() as d: await pg.click('#vbTao')
        dl = await d.value; xp2 = '/tmp/_vb.docx'; await dl.save_as(xp2); t = re.sub(r'<[^>]+>', '', zipfile.ZipFile(xp2).read('word/document.xml').decode())
        ok(dl.suggested_filename.startswith('phieu_chi_STT001') and 'Ông/Bà Nguyễn Văn An' in t and '1.700.000.000' in t and 'Một tỷ bảy trăm triệu đồng' in t and '01/09/2026' in t, 'tạo văn bản 1 hộ: họ tên, tiền, tiền bằng chữ, ngày kiểm đếm')
        await pg.click('#bVB'); await pg.wait_for_timeout(200)
        async with pg.expect_download() as d: await pg.click('#vbTao')
        dl = await d.value; xp3 = '/tmp/_vb.zip'; await dl.save_as(xp3); ok(len(zipfile.ZipFile(xp3).namelist()) == 24, 'tạo văn bản hàng loạt → zip 24 hộ')
        # Word tự sửa theo web
        await pg.click('#tMain tbody tr:nth-child(1) [data-a=sua]'); await pg.click('#fm button[value=cancel]')
        async with pg.expect_file_chooser() as fc: await pg.click('#tMain tbody tr:nth-child(1) [data-a=kem]')
        await (await fc.value).set_files(f'{FX}/tb_that.docx'); await pg.wait_for_timeout(800)
        async with pg.expect_download() as d: await pg.click('#tMain tbody tr:nth-child(1) [data-a=tai]')
        x = zipfile.ZipFile(await (await d.value).path()).read('word/document.xml').decode()
        t = re.sub(r'<[^>]+>', '', x)
        ok('Nguyễn Văn An.' in t and '33,5' in t and x.count('<w:b/>') == 2, 'tải TB: tự sửa họ tên + diện tích, giữ chữ đậm')
        # ảnh hiện trạng PDF
        async with pg.expect_file_chooser() as fc: await pg.click('#tMain tbody tr:nth-child(2) [data-a=kemht]')
        await (await fc.value).set_files(f'{FX}/ht.pdf'); await pg.wait_for_timeout(2500)
        await pg.click('#tMain tbody tr:nth-child(2) [data-a=ht]'); await pg.wait_for_timeout(3000)
        ok(await pg.locator('#anhGrid canvas').count() == 3, 'ảnh hiện trạng PDF: 3 ảnh'); await pg.click('#anhDong')
        # chụp ảnh HT bằng điện thoại → PDF có dấu giờ/GPS; hộ đã có PDF thì gộp thêm
        async with pg.expect_file_chooser() as fc: await pg.click('#tMain tbody tr:nth-child(3) [data-a=chup]')
        await (await fc.value).set_files([f'{FX}/chup1.jpg', f'{FX}/chup2.jpg']); await pg.wait_for_timeout(9000)
        ok('📷 2' in await pg.inner_text('#tMain tbody tr:nth-child(3)'), 'chụp ảnh HT: 2 ảnh → PDF 2 trang gắn vào hộ')
        async with pg.expect_file_chooser() as fc: await pg.click('#tMain tbody tr:nth-child(2) [data-a=chup]')
        await (await fc.value).set_files([f'{FX}/chup1.jpg']); await pg.wait_for_timeout(500); await pg.click('dialog[open] button[value=ok]'); await pg.wait_for_timeout(9000)
        ok('📷 4' in await pg.inner_text('#tMain tbody tr:nth-child(2)'), 'chụp thêm vào hồ sơ đã có 3 ảnh → 4 ảnh')
        # QR: in nhãn → quét từ ảnh → mở đúng hộ; link #ho=
        await pg.click('#tMain tbody tr:nth-child(1) [data-a=nhan]'); await pg.wait_for_timeout(800)
        src = await pg.get_attribute('#inNhan img', 'src'); import base64; open('/tmp/_qr.gif', 'wb').write(base64.b64decode(src.split(',')[1]))
        ok(await pg.locator('#inNhan .nhan').count() == 1 and 'Nguyễn Văn An' in await pg.inner_text('#inNhan'), 'in nhãn QR hồ sơ')
        await pg.click('#bQuet'); await pg.wait_for_timeout(800)
        async with pg.expect_file_chooser() as fc: await pg.click('#qrAnh')
        await (await fc.value).set_files('/tmp/_qr.gif'); await pg.wait_for_timeout(800)
        ok(await pg.inner_text('#cnt') == '1 hộ' and 'Nguyễn Văn An' in await pg.inner_text('#tMain tbody'), 'quét QR (ảnh) → mở đúng hộ'); await pg.click('#bClear')
        p2 = await R.page('http://local/index.html#ho=m5'); await p2.wait_for_timeout(500)
        ok(await p2.inner_text('#cnt') == '1 hộ' and 'Huỳnh Văn Sang' in await p2.inner_text('#tMain tbody'), 'mở link #ho=<id> → đúng hộ'); await p2.close()
        # zip, excel
        async with pg.expect_download() as d: await pg.click('#bZip')
        ok((await d.value).suggested_filename.endswith('.zip'), 'tải TB zip')
        async with pg.expect_download() as d: await pg.click('#bXls')
        xp = '/tmp/_xuat.xlsx'; await (await d.value).save_as(xp); ws = openpyxl.load_workbook(xp).active; hd = [c.value for c in ws[1]]; r1 = dict(zip(hd, [c.value for c in ws[2]]))
        ok(r1.get('Tiến độ GPMB (bước đã xong)') == 'Phê duyệt PA' and r1.get('Ngày Kiểm đếm') == '01/09/2026', 'xuất Excel (có cột tiến độ + ngày từng bước)')
        wb = openpyxl.load_workbook(xp); last = [c.value for c in wb['Tổng'][wb['Tổng'].max_row]]; tt = list(wb['Theo tổ'].values)
        ok(str(last[0]).startswith('TỔNG CỘNG') and r1.get('Tổng PA (đ)') == 1700000000 and tt[-1][0] == 'TỔNG CỘNG' and tt[-1][9] == sum(x[9] for x in tt[1:-1]), 'xuất Excel: dòng tổng + trang "Theo tổ" cộng đúng')
        async with pg.expect_file_chooser() as fc:
            await pg.click('#bCfg'); await pg.click('#iXls')
        await (await fc.value).set_files(f'{FX}/ds.xlsx'); await pg.wait_for_timeout(800)
        await pg.click('dialog[open] button[value=no]'); await pg.wait_for_timeout(500)
        ok('26 hộ' in await pg.inner_text('#cnt'), 'nhập Excel thêm 2 hộ')
        async with pg.expect_file_chooser() as fc:
            await pg.click('#bCfg'); await pg.click('#iXls')
        await (await fc.value).set_files(f'{FX}/ds_gd.xlsx'); await pg.wait_for_timeout(800)
        await pg.click('dialog[open] button[value=no]'); await pg.wait_for_timeout(500)
        await pg.fill('#q', 'Tiến Độ'); await pg.wait_for_timeout(200)
        ok('Chi trả · 15/10/2026' in await pg.inner_text('#tMain tbody'), 'nhập Excel: cột tiến độ (số) + ngày dd/mm/yyyy'); await pg.click('#bClear')
        # bản đồ
        await pg.click('#tabMap'); await pg.wait_for_timeout(1200)
        await pg.fill('#mq', '134/3'); await pg.press('#mq', 'Enter'); await pg.wait_for_timeout(800)
        ok('Thửa 3' in await pg.inner_text('#mInfo'), 'tìm thửa trên bản đồ')
        box = await pg.locator('#map').bounding_box(); cx, cy = box['x'] + box['width'] / 2, box['y'] + box['height'] / 2
        await pg.click('[data-ed=tach]'); await pg.mouse.click(cx, cy); await pg.mouse.click(cx, cy - 250); await pg.mouse.click(cx, cy + 250); await pg.wait_for_timeout(500)
        await pg.fill('dialog[open] input', '134/3A'); await pg.click('dialog[open] button[value=ok]'); await pg.wait_for_timeout(600)
        ok('Đã tách' in await pg.inner_text('#edMsg'), 'tách thửa'); await pg.click('[data-ed=tach]')
        await pg.click('[data-ed=gop]'); await pg.mouse.click(cx - 12, cy); await pg.wait_for_timeout(300); await pg.mouse.click(cx + 12, cy); await pg.wait_for_timeout(1200)
        ok('thành 1 thửa' in await pg.inner_text('#edMsg'), 'gộp 2 thửa liền nhau → 1 đa giác thật (' + (await pg.inner_text('#edMsg'))[:60] + ')'); await pg.click('[data-ed=gop]')
        async with pg.expect_download() as d: await pg.click('#mDxf')
        await (await d.value).save_as('/tmp/_bd.dxf'); import subprocess, json, os
        rc = subprocess.run(['python3', os.path.join(os.path.dirname(FX), '..', 'dxf2bando.py'), '/tmp/_bd.dxf', '--ranh', 'RANH_THU_HOI', '-o', '/tmp/_bd.json'], capture_output=True, text=True)
        bj = json.load(open('/tmp/_bd.json')) if rc.returncode == 0 else {'thua': []}; ls = [t['l'] for t in bj['thua']]
        ok(len(ls) == 24 and '134/7' in ls and not any(l.startswith('?') for l in ls) and bj['lop']['ranh'], 'xuất DXF → dxf2bando.py đọc lại đủ 24 thửa, đúng nhãn, có ranh ' + rc.stderr[-200:])
        await pg.select_option('#mKho', 'A3'); await pg.click('#mIn'); await pg.wait_for_timeout(900)
        t = await pg.inner_text('#khung'); ok('Chú giải' in t and 'Tỷ lệ xấp xỉ 1:' in t and 'A3' in t, 'in bản đồ: khung tên, chú giải, tỷ lệ, khổ A3')
        await pg.click('#mDt')
        for dx, dy in [(-150, -100), (150, -100), (150, 100)]: await pg.mouse.click(cx + dx, cy + dy)
        ok('Diện tích' in await pg.inner_text('#measBar'), 'đo diện tích'); await pg.click('#mDt')
        async with pg.expect_download() as d: await pg.click('#mKml')
        ok((await d.value).suggested_filename.endswith('.kml'), 'xuất KML')
        await pg.select_option('#mMode', 'gd'); await pg.wait_for_timeout(300)
        ok('Chưa bắt đầu' in await pg.inner_text('#mLeg') and 'Xong: Bàn giao MB' in await pg.inner_text('#mLeg'), 'bản đồ tô màu theo tiến độ GPMB')
        for fx, n, k in [('thua.geojson', 2, '50/1'), ('thua.kml', 3, '51/3')]:
            async with pg.expect_file_chooser() as fc:
                await pg.click('#bCfg'); await pg.click('#iMap')
            await (await fc.value).set_files(f'{FX}/{fx}'); await pg.wait_for_timeout(700)
            await pg.fill('#mq', k); await pg.press('#mq', 'Enter'); await pg.wait_for_timeout(500)
            m = re.search(r'DT thửa trên bản vẽ: ([\d.,]+) m²', await pg.inner_text('#mInfo')); dt = float(m.group(1).replace('.', '').replace(',', '.')) if m else 0
            ok(f'{n} thửa' in await pg.inner_text('#mSrc') and 90 < dt < 105, f'nạp bản vẽ {fx}: {n} thửa, đổi WGS84 → VN2000, DT {dt} m²')
        await pg.set_viewport_size({'width': 390, 'height': 800}); await pg.click('#tabDs'); await pg.wait_for_timeout(400)
        ok(await pg.evaluate('document.documentElement.scrollWidth') <= 390, 'điện thoại không tràn ngang')
        ok(not R.errs, 'không có lỗi JS ' + str(R.errs))

asyncio.run(main())
