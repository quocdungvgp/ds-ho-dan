"""Chạy thử chế độ nối Google Sheet: Code.gs chạy trên máy chủ giả lập (node tests/mock-gas.js — chạy lại mỗi lần test
để dữ liệu sạch). Tài khoản admin / Doi-mat-khau-ngay-1 (MK_ADMIN_DAU). Chạy: python3 tests/ui_server.py"""
import asyncio, urllib.parse
from common import Run, FX, API, login, ok

LINK = 'http://local/index.html?api=' + urllib.parse.quote(API)


async def user(pg, ten, q, to):
    await pg.click('#bUsers'); await pg.wait_for_timeout(500)
    await pg.fill('#fUser [name=ten]', ten); await pg.fill('#fUser [name=hoten]', ten); await pg.select_option('#fUser [name=q]', q)
    await pg.fill('#fUser [name=pass]', '12345678')
    if to != '*': await pg.uncheck('#uTo [value="*"]'); await pg.check(f'#uTo [value="{to}"]')
    await pg.click('#uLuu'); await pg.wait_for_timeout(600); await pg.click('#uDong')


async def main():
    async with Run() as R:
        ad = await R.page(LINK)
        ok(await ad.evaluate("document.querySelector('#dLogin').open"), 'link dự án → hỏi đăng nhập')
        await login(ad, 'admin', 'sai'); ok('Sai' in await ad.inner_text('#loginMsg'), 'sai mật khẩu bị chặn')
        await login(ad, 'admin', 'Doi-mat-khau-ngay-1')
        ok('001/' in await ad.inner_text('#tMain tbody tr:nth-child(1)'), 'đăng nhập admin, máy chủ cấp mã hồ sơ')
        await user(ad, 'to2', 'denghi', 'Tổ 2'); await user(ad, 'xem1', 'xem', '*')
        # cấu hình chung
        await ad.click('#bCfg'); await ad.fill('#fCfg [name=duAn]', 'Dự án THỬ'); await ad.click('#fCfg button[value=ok]'); await ad.wait_for_timeout(800)
        xm = await R.page(LINK); await login(xm, 'xem1', '12345678')
        ok(await xm.locator('#tMain [data-a=sua]').count() == 0 and await xm.locator('#bAdd').is_hidden(), 'quyền Xem: không có nút sửa')
        ok('THỬ' in await xm.inner_text('#duAn'), 'cấu hình chung áp cho người khác')
        # đề nghị → duyệt
        t2 = await R.page(LINK); await login(t2, 'to2', '12345678')
        ok(await t2.locator('#tMain [data-a=sua]').count() == 2, 'phạm vi Tổ 2: chỉ sửa được 2 hộ Tổ 2')
        await t2.click('#tMain [data-a=sua] >> nth=0'); await t2.fill('#fm [name=dt]', '45,5'); await t2.click('#fm button[value=ok]'); await t2.wait_for_timeout(1200)
        ok('đề nghị' in await t2.inner_text('#toast'), 'sửa thành đề nghị')
        async with t2.expect_file_chooser() as fc: await t2.click('#tMain [data-a=kem] >> nth=0')
        await (await fc.value).set_files(f'{FX}/mau.docx'); await t2.wait_for_timeout(1500)
        await ad.click('#bReload'); await ad.wait_for_timeout(1000)
        ok(await ad.inner_text('#nDuyet') == '2', 'admin thấy 2 đề nghị')
        await ad.click('#bDuyet'); await ad.click('#dAll'); await ad.click('dialog[open] button[value=ok]'); await ad.wait_for_timeout(1500); await ad.click('#dDong')
        row = await ad.inner_text('#tMain tbody tr:nth-child(2)')
        ok('45,5' in row and 'Xem trước' in row, 'duyệt toàn bộ: sửa + đính kèm có hiệu lực')
        # đính kèm hàng loạt, nhập Excel thay toàn bộ giữ TB
        async with ad.expect_file_chooser() as fc: await ad.click('#bKemLoat')
        await (await fc.value).set_files([f'{FX}/tb.zip', f'{FX}/TBTHD_STT001_Nguyen Van An.docx']); await ad.wait_for_timeout(1500)
        ok('bỏ qua 1' in await ad.inner_text('dialog[open] h3'), 'hàng loạt: báo file không ghép được')
        await ad.click('dialog[open] button[value=ok]'); await ad.wait_for_timeout(2500)
        async with ad.expect_file_chooser() as fc:
            await ad.click('#bCfg'); await ad.click('#iXls')
        await (await fc.value).set_files(f'{FX}/ds2.xlsx'); await ad.wait_for_timeout(800)
        await ad.click('dialog[open] button[value=ok]'); await ad.wait_for_timeout(2500)
        ok(await ad.locator('#tMain [data-a=tai]').count() == 2, 'nhập Excel thay toàn bộ vẫn giữ thông báo')
        # nhật ký, ảnh HT qua máy chủ
        async with ad.expect_file_chooser() as fc: await ad.click('#tMain tbody tr:nth-child(1) [data-a=kemht]')
        await (await fc.value).set_files(f'{FX}/ht.pdf'); await ad.wait_for_timeout(2500)
        await ad.click('#tMain tbody tr:nth-child(1) [data-a=ht]'); await ad.wait_for_timeout(3000)
        ok(await ad.locator('#anhGrid canvas').count() == 3, 'ảnh hiện trạng lưu/đọc qua máy chủ'); await ad.click('#anhDong')
        # tiến độ GPMB lưu qua máy chủ, hạn bàn giao là cấu hình chung
        await ad.click('#bCfg'); await ad.fill('#fCfg [name=hanBG]', '2027-03-31'); await ad.click('#fCfg button[value=ok]'); await ad.wait_for_timeout(800)
        await ad.click('#tMain tbody tr:nth-child(1) [data-a=sua]'); await ad.fill('#fm [name=gdn_ct]', '2026-11-25'); await ad.click('#fm button[value=ok]'); await ad.wait_for_timeout(1200)
        await xm.click('#bReload'); await xm.wait_for_timeout(1000)
        ok('Chi trả · 25/11/2026' in await xm.inner_text('#tMain tbody tr:nth-child(1)') and '31/03/2027' in await xm.inner_text('#gdHan'), 'tiến độ GPMB lưu Sheet, hạn bàn giao áp cho mọi người')
        await ad.click('#bLog'); await ad.wait_for_timeout(800); ok('duyệt' in await ad.inner_text('#tLog'), 'nhật ký ghi thao tác')
        ok(not R.errs, 'không có lỗi JS ' + str(R.errs))

asyncio.run(main())
