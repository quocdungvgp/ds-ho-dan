"""Chạy thử chế độ KHÔNG nối Sheet (dữ liệu mẫu, lưu trong trình duyệt). Chạy: python3 tests/ui_local.py"""
import asyncio, zipfile, re
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
        # mẫu TB + tạo TB + xem trước
        async with pg.expect_file_chooser() as fc:
            await pg.click('#bCfg'); await pg.click('#iMau')
        await (await fc.value).set_files(f'{FX}/mau.docx'); await pg.wait_for_timeout(800)
        await pg.click('#tMain [data-a=tao] >> nth=0'); await pg.wait_for_timeout(2500)
        await pg.click('#tMain [data-a=xem] >> nth=0'); await pg.wait_for_timeout(2500)
        ok('Ông/Bà' in await pg.inner_text('#xemBody'), 'tạo TB theo mẫu + xem trước'); await pg.click('#xemDong')
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
        # zip, excel
        async with pg.expect_download() as d: await pg.click('#bZip')
        ok((await d.value).suggested_filename.endswith('.zip'), 'tải TB zip')
        async with pg.expect_download() as d: await pg.click('#bXls')
        ok((await d.value).suggested_filename.endswith('.xlsx'), 'xuất Excel')
        async with pg.expect_file_chooser() as fc:
            await pg.click('#bCfg'); await pg.click('#iXls')
        await (await fc.value).set_files(f'{FX}/ds.xlsx'); await pg.wait_for_timeout(800)
        await pg.click('dialog[open] button[value=no]'); await pg.wait_for_timeout(500)
        ok('26 hộ' in await pg.inner_text('#cnt'), 'nhập Excel thêm 2 hộ')
        # bản đồ
        await pg.click('#tabMap'); await pg.wait_for_timeout(1200)
        await pg.fill('#mq', '134/3'); await pg.press('#mq', 'Enter'); await pg.wait_for_timeout(800)
        ok('Thửa 3' in await pg.inner_text('#mInfo'), 'tìm thửa trên bản đồ')
        box = await pg.locator('#map').bounding_box(); cx, cy = box['x'] + box['width'] / 2, box['y'] + box['height'] / 2
        await pg.click('[data-ed=tach]'); await pg.mouse.click(cx, cy); await pg.mouse.click(cx, cy - 250); await pg.mouse.click(cx, cy + 250); await pg.wait_for_timeout(500)
        await pg.fill('dialog[open] input', '134/3A'); await pg.click('dialog[open] button[value=ok]'); await pg.wait_for_timeout(600)
        ok('Đã tách' in await pg.inner_text('#edMsg'), 'tách thửa'); await pg.click('[data-ed=tach]')
        await pg.click('#mDt')
        for dx, dy in [(-150, -100), (150, -100), (150, 100)]: await pg.mouse.click(cx + dx, cy + dy)
        ok('Diện tích' in await pg.inner_text('#measBar'), 'đo diện tích'); await pg.click('#mDt')
        async with pg.expect_download() as d: await pg.click('#mKml')
        ok((await d.value).suggested_filename.endswith('.kml'), 'xuất KML')
        await pg.set_viewport_size({'width': 390, 'height': 800}); await pg.click('#tabDs'); await pg.wait_for_timeout(400)
        ok(await pg.evaluate('document.documentElement.scrollWidth') <= 390, 'điện thoại không tràn ngang')
        ok(not R.errs, 'không có lỗi JS ' + str(R.errs))

asyncio.run(main())
