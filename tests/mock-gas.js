// Giả lập Google Apps Script tối thiểu để chạy Code.gs bằng Node
const crypto = require('crypto'), fs = require('fs'), vm = require('vm'), http = require('http');
const sheets = {};
function Sheet(name) {
  const d = []; // rows of arrays
  const s = {
    name, d,
    getLastRow: () => { let n = d.length; while (n && !d[n - 1].some(v => v !== '' && v !== null && v !== undefined)) n--; return n; },
    getLastColumn: () => d.reduce((m, r) => Math.max(m, r.length), 0),
    getRange(r, c, nr = 1, nc = 1) {
      const rg = {
        getValues() { const o = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push((d[r - 1 + i] || [])[c - 1 + j] ?? ''); o.push(row); } return o; },
        setValues(v) { for (let i = 0; i < nr; i++) { d[r - 1 + i] = d[r - 1 + i] || []; for (let j = 0; j < nc; j++) d[r - 1 + i][c - 1 + j] = v[i][j]; } return rg; },
        setValue(v) { d[r - 1] = d[r - 1] || []; d[r - 1][c - 1] = v; return rg; },
        clearContent() { for (let i = 0; i < nr; i++) if (d[r - 1 + i]) for (let j = 0; j < nc; j++) d[r - 1 + i][c - 1 + j] = ''; return rg; },
        setFontWeight() { return rg; }
      }; return rg;
    },
    appendRow(a) { d[s.getLastRow()] = a.slice(); },
    deleteRow(i) { d.splice(i - 1, 1); },
    setFrozenRows() {}
  }; return s;
}
const files = {};
const mkFile = (blob) => { const id = 'f' + crypto.randomBytes(6).toString('hex'); const f = { id, blob, trashed: false, getId: () => id, getBlob: () => blob, setTrashed(v) { f.trashed = v; } }; files[id] = f; return f; };
const Blob = (bytes, mime, name) => { const b = { bytes: Buffer.from(bytes), name, getBytes: () => [...b.bytes].map(x => x > 127 ? x - 256 : x), getDataAsString: () => b.bytes.toString('utf8'), setName(n) { b.name = n; return b; } }; return b; };
const folder = { createFile(a, content, mime) { const blob = typeof a === 'string' ? Blob(Buffer.from(content), mime, a) : a; return mkFile(blob); },
  getFilesByName(n) { const l = Object.values(files).filter(f => !f.trashed && f.blob.name === n); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; } };
const cache = {}, props = {};
const toBuf = x => typeof x === 'string' ? Buffer.from(x, 'utf8') : Buffer.from(x.map(v => v & 255));
const G = {
  SpreadsheetApp: { getActive: () => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = Sheet(n)), getUrl: () => 'https://docs.google.com/spreadsheets/d/TEST' }) },
  DriveApp: { getRootFolder: () => folder, getFolderById: () => folder, getFileById: id => { if (!files[id]) throw new Error('no file ' + id); return files[id]; } },
  Utilities: {
    DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
    computeDigest: (a, v) => [...crypto.createHash('sha256').update(toBuf(v)).digest()].map(x => x > 127 ? x - 256 : x),
    base64Encode: v => toBuf(v).toString('base64'), base64Decode: s => [...Buffer.from(s, 'base64')].map(x => x > 127 ? x - 256 : x),
    newBlob: (v, mime, name) => Blob(typeof v === 'string' ? Buffer.from(v) : toBuf(v), mime, name),
    getUuid: () => crypto.randomUUID(),
    formatDate: (d, tz, f) => d.toISOString().slice(0, 16).replace('T', ' ')
  },
  CacheService: { getScriptCache: () => ({ get: k => cache[k] ?? null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ t, setMimeType() { return this; } }) },
  Logger: { log: console.log }, console
};
vm.createContext(G);
vm.runInContext(fs.readFileSync(process.argv[2] || __dirname + '/../Code.gs', 'utf8') + '\n;this.doPost=doPost;this.caiDat=caiDat;', G);
G.caiDat();
// nạp vài hộ mẫu
const rows = [['1', '', '', 'Nguyễn Văn An', '10 Đường A', '19', '4', 'Tổ 1', '134', '1', '20,5', 'Một phần'],
              ['2', '', '', 'Trần Thị Bích', '12 Đường A', '19', '4', 'Tổ 2', '134', '2', '30', 'Toàn phần'],
              ['3', '', '', 'Lê Văn Cường', '14 Đường A', '19', '4', 'Tổ 2', '134', '3', '', '']];
rows.forEach(r => sheets['Tổng'].appendRow(r));
G.caiDat();
http.createServer((req, res) => {
  let b = ''; req.on('data', c => b += c); req.on('end', () => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    if (req.method !== 'POST') return res.end('{}');
    const o = G.doPost({ postData: { contents: b } }); res.setHeader('content-type', 'application/json'); res.end(o.t);
  });
}).listen(8787, () => console.log('mock on 8787'));
