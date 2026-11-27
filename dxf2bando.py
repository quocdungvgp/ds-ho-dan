"""
Đổi bản vẽ CAD (DXF, tọa độ VN2000) ra bando.json cho trang "Bản đồ dự án".

Cách dùng (AutoCAD: SAVEAS → DXF 2007 hoặc 2013):
    pip install ezdxf
    python dxf2bando.py "Ban do du an.dxf"
    python dxf2bando.py "Ban do du an.dxf" --thua THUA_DAT --nhan SO_THUA --ranh RANH_THU_HOI --to 134

Quy ước bản vẽ (sửa bằng tham số nếu lớp tên khác):
  --thua   lớp ranh thửa: polyline KÍN (LWPOLYLINE / POLYLINE / HATCH ranh ngoài)
  --nhan   lớp nhãn thửa: TEXT/MTEXT nằm TRONG thửa, dạng "tờ/thửa" (vd 134/7) hoặc chỉ "thửa" (khi đó dùng --to)
  --ranh   lớp ranh dự án / ranh thu hồi (vẽ đường vàng), có thể nhiều lớp: --ranh RANH --ranh RANH_KP
  --to     số tờ mặc định cho nhãn chỉ ghi số thửa
Thửa không có nhãn vẫn được vẽ (nhãn "?/id") để dễ thấy chỗ thiếu.
Kết quả: bando.json — đặt cạnh index.html trên máy chủ, hoặc nạp ở ⚙ Cấu hình → "Nạp bản vẽ".
"""
import argparse, json, math, re, sys

try:
    import ezdxf
except ImportError:
    sys.exit("Chưa có ezdxf — chạy: pip install ezdxf")


def area(v):
    s = 0.0
    for i in range(len(v)):
        x1, y1 = v[i]; x2, y2 = v[(i + 1) % len(v)]
        s += x1 * y2 - x2 * y1
    return abs(s) / 2


def inside(x, y, v):
    c = False; j = len(v) - 1
    for i in range(len(v)):
        xi, yi = v[i]; xj, yj = v[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
            c = not c
        j = i
    return c


def ring_of(e):
    t = e.dxftype()
    if t == "LWPOLYLINE":
        pts = [(p[0], p[1]) for p in e.get_points("xy")]
        closed = e.closed
    elif t == "POLYLINE":
        pts = [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices]
        closed = e.is_closed
    else:
        return None
    if len(pts) >= 3 and not closed and math.dist(pts[0], pts[-1]) > 0.05:
        return None                       # hở → không phải thửa
    if math.dist(pts[0], pts[-1]) < 1e-6:
        pts = pts[:-1]
    return pts if len(pts) >= 3 else None


def text_of(e):
    if e.dxftype() == "TEXT":
        return e.dxf.text, e.dxf.insert
    if e.dxftype() == "MTEXT":
        return e.plain_text(), e.dxf.insert
    return None, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf")
    ap.add_argument("--thua", default="THUA_DAT")
    ap.add_argument("--nhan", default="SO_THUA")
    ap.add_argument("--ranh", action="append", default=None)
    ap.add_argument("--to", default="")
    ap.add_argument("-o", "--out", default="bando.json")
    a = ap.parse_args()
    ranh_layers = [s.upper() for s in (a.ranh or ["RANH", "RANH_THU_HOI", "RANH_DU_AN"])]

    msp = ezdxf.readfile(a.dxf).modelspace()
    layers = {}
    for e in msp:
        layers.setdefault(e.dxf.layer.upper(), 0)
        layers[e.dxf.layer.upper()] += 1

    rings = []
    for e in msp.query(f'LWPOLYLINE POLYLINE[layer=="{a.thua}"]i'):
        r = ring_of(e)
        if r:
            rings.append(r)
    for h in msp.query(f'HATCH[layer=="{a.thua}"]i'):
        for p in h.paths:
            if hasattr(p, "vertices") and len(p.vertices) >= 3:
                rings.append([(v[0], v[1]) for v in p.vertices])
    if not rings:
        print("Không thấy thửa kín nào trên lớp", a.thua)
        print("Các lớp có trong bản vẽ:", ", ".join(f"{k}({n})" for k, n in sorted(layers.items())))
        sys.exit(1)

    texts = []
    for e in msp.query(f'TEXT MTEXT[layer=="{a.nhan}"]i'):
        s, p = text_of(e)
        if s:
            texts.append((s.strip(), p.x, p.y))

    xs = [x for r in rings for x, _ in r]; ys = [y for r in rings for _, y in r]
    o = [math.floor(min(xs) / 1000) * 1000, math.floor(min(ys) / 1000) * 1000]

    # nhãn → thửa nhỏ nhất chứa điểm chèn
    ring_area = [area(r) for r in rings]
    label = [None] * len(rings)
    for s, x, y in texts:
        cand = [i for i, r in enumerate(rings) if inside(x, y, r)]
        if not cand:
            continue
        i = min(cand, key=lambda k: ring_area[k])
        m = re.match(r"^\s*(\d+)\s*[/\-]\s*(\d+\w*)", s)
        if m:
            label[i] = f"{m.group(1)}/{m.group(2)}"
        elif re.match(r"^\s*\d+\w*\s*$", s) and a.to:
            label[i] = f"{a.to}/{s.strip()}"

    thua = []
    for i, r in enumerate(rings):
        thua.append({
            "id": f"t{i}",
            "l": label[i] or f"?/{i}",
            "v": [[round(x - o[0], 2), round(y - o[1], 2)] for x, y in r],
            "dt": round(ring_area[i], 1),
        })

    ranh = []
    for e in msp.query("LWPOLYLINE POLYLINE LINE"):
        if e.dxf.layer.upper() not in ranh_layers:
            continue
        if e.dxftype() == "LINE":
            pts = [(e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)]
        elif e.dxftype() == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points("xy")]
            if e.closed: pts.append(pts[0])
        else:
            pts = [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices]
            if e.is_closed: pts.append(pts[0])
        ranh.append([[round(x - o[0], 2), round(y - o[1], 2)] for x, y in pts])

    out = {"v": 1, "crs": f"VN2000 TP.HCM (m), gốc {o[0]},{o[1]}", "o": o, "src": a.dxf,
           "thua": thua, "lop": {"ranh": ranh}}
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))

    nolab = sum(1 for l in label if not l)
    print(f"Đã ghi {a.out}: {len(thua)} thửa ({len(thua) - nolab} có nhãn tờ/thửa, {nolab} chưa có), {len(ranh)} đoạn ranh.")
    if nolab:
        print("  → thửa chưa nhãn hiện trên bản đồ là '?/số'. Kiểm tra lớp nhãn --nhan hoặc thêm --to <số tờ>.")
    if not (300000 < o[0] < 800000 and 1000000 < o[1] < 2600000):
        print("  ⚠ Tọa độ không giống VN2000 (X ~ 600000, Y ~ 1190000 ở TP.HCM) — bản vẽ có đang ở tọa độ giả định không?")


if __name__ == "__main__":
    main()
