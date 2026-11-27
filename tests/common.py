"""Khung chạy thử chung: mở trang bằng Chrome thật (playwright), KHÔNG cần mạng.
- Thư viện CDN (cdn.jsdelivr.net/npm/...) được trả từ tests/node_modules (chạy `npm i` trong tests/).
- Trang (index.html, bando.json nếu có) trả từ thư mục gốc repo qua địa chỉ giả http://local/.
- Apps Script (https://script.google.com/...) được chuyển sang máy chủ giả lập: node tests/mock-gas.js (cổng 8787).
Biến môi trường CHROME: đường dẫn Chrome (mặc định /usr/bin/google-chrome)."""
import os, urllib.request
from playwright.async_api import async_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
NM = os.path.join(HERE, 'node_modules')
FX = os.path.join(HERE, 'fx')
API = 'https://script.google.com/macros/s/TEST/exec'
CDN = {  # đường dẫn CDN trong index.html -> file trong node_modules
    'leaflet@1.9.4/dist/leaflet.js': 'leaflet/dist/leaflet.js', 'leaflet@1.9.4/dist/leaflet.css': 'leaflet/dist/leaflet.css',
    'proj4@2.12.1/dist/proj4.js': 'proj4/dist/proj4.js', 'xlsx@0.18.5/dist/xlsx.full.min.js': 'xlsx/dist/xlsx.full.min.js',
    'jszip@3.10.1/dist/jszip.min.js': 'jszip/dist/jszip.min.js', 'pizzip@3.1.7/dist/pizzip.min.js': 'pizzip/dist/pizzip.min.js',
    'docxtemplater@3.47.4/build/docxtemplater.js': 'docxtemplater/build/docxtemplater.js',
    'docx-preview@0.3.2/dist/docx-preview.min.js': 'docx-preview/dist/docx-preview.min.js',
    'qrcode-generator@1.4.4/qrcode.js': 'qrcode-generator/qrcode.js', 'jsqr@1.4.0/dist/jsQR.js': 'jsqr/dist/jsQR.js', 'jspdf@2.5.1/dist/jspdf.umd.min.js': 'jspdf/dist/jspdf.umd.min.js', 'polygon-clipping@0.15.7/dist/polygon-clipping.umd.min.js': 'polygon-clipping/dist/polygon-clipping.umd.min.js',
    'pdfjs-dist@3.11.174/build/pdf.min.js': 'pdfjs-dist/build/pdf.min.js', 'pdfjs-dist@3.11.174/build/pdf.worker.min.js': 'pdfjs-dist/build/pdf.worker.min.js',
}


async def route(r):
    u = r.request.url
    try:
        if 'cdn.jsdelivr.net/npm/' in u:
            k = u.split('/npm/')[1]
            if k not in CDN: return await r.fulfill(status=404, body='')
            return await r.fulfill(path=os.path.join(NM, CDN[k]), content_type='text/css' if k.endswith('.css') else 'application/javascript')
        if u.startswith('http://local/'):
            f = os.path.join(ROOT, u.split('http://local/')[1].split('#')[0].split('?')[0] or 'index.html')
            if os.path.exists(f): return await r.fulfill(path=f)
            return await r.fulfill(status=404, body='')
        if u.startswith('https://script.google.com/'):
            req = urllib.request.Request('http://localhost:8787', data=r.request.post_data_buffer, method='POST')
            return await r.fulfill(body=urllib.request.urlopen(req).read(), content_type='application/json', headers={'access-control-allow-origin': '*'})
        await r.abort()   # ảnh nền bản đồ, font… không cần
    except Exception:
        pass


class Run:
    def __init__(self): self.errs = []

    async def __aenter__(self):
        self.p = await async_playwright().start()
        self.b = await self.p.chromium.launch(executable_path=os.environ.get('CHROME', '/usr/bin/google-chrome'), args=['--no-sandbox'])
        return self

    async def __aexit__(self, *a):
        await self.b.close(); await self.p.stop()

    async def page(self, url='http://local/index.html', w=1300, h=900):
        ctx = await self.b.new_context(viewport={'width': w, 'height': h}, accept_downloads=True)
        pg = await ctx.new_page()
        pg.on('pageerror', lambda e: self.errs.append(str(e)))
        await pg.route('**/*', route)
        await pg.goto(url); await pg.wait_for_timeout(1200)
        return pg


async def login(pg, u, pw):
    await pg.fill('#fLogin [name=user]', u); await pg.fill('#fLogin [name=pass]', pw)
    await pg.click('#fLogin button[value=ok]'); await pg.wait_for_timeout(1200)


def ok(cond, msg):
    print(('  OK  ' if cond else '  LỖI ') + msg)
    return cond
