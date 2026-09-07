# '통합 문서1.xlsx' 의 김천 시트에서 농지소재지가 개령면인 신청농지만 뽑아
# lee.html 의 PARCELS / GEO / MAPC / MAPG / MAPL 블록을 새로 만들어 끼운다.
# 대장이 갱신되면 이 스크립트만 다시 돌리면 된다 (앱의 나머지 코드는 건드리지 않는다).
#
# 좌표와 지적도 경계는 새로 만들지 않는다 — 2025년 농지전수조사 앱
# (github.com/Hyeokgi/gaeryeong-survey 의 kim.html·moon.html) 에 이미 들어 있는 것을
# (리, 지번) 으로 맞춰 가져온다. 그 앱의 조사 대상이 아니었던 지번은 좌표가 없다.
#
#   사용법:  python tools/build_nongji.py [gaeryeong-survey 저장소 경로]
#   기본값:  이 파일 기준 ../gaeryeong-survey
#
#   필요:    openpyxl

import openpyxl, re, io, json, math, os, sys, collections

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

XL = os.path.join(ROOT, "통합 문서1.xlsx")
APP = os.path.join(ROOT, "lee.html")
REPO = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "gaeryeong-survey")

EMD = "개령면"
BUF = 700.0          # 신청농지에서 이 거리(m) 안의 지적도 폴리곤만 남긴다 (약도 최대 배율이 1km)
PREC = 1e6


# ── 전수조사 앱에서 좌표·지적도 읽기 ──────────────────────────
def grab(f):
    t = io.open(os.path.join(REPO, f), encoding="utf-8").read()
    return (json.loads(re.search(r"const PARCELS = (\[.*?\]);\n", t, re.S).group(1)),
            json.loads(re.search(r"const GEO = (\{.*?\});\n", t, re.S).group(1)),
            re.search(r'const MAPC="(.*?)"\n', t, re.S).group(1),
            json.loads(re.search(r"const MAPG=(\[.*?\]);\n", t, re.S).group(1)),
            json.loads(re.search(r"const MAPL=(\{.*?\});\n", t, re.S).group(1)))


def decode(s):
    """폴리라인 풀기 — 구글 방식 그대로다"""
    out, i, x, y = [], 0, 0, 0
    while i < len(s):
        for which in (0, 1):
            r = sh = 0
            while True:
                b = ord(s[i]) - 63; i += 1
                r |= (b & 0x1f) << sh; sh += 5
                if b < 0x20: break
            d = ~(r >> 1) if (r & 1) else (r >> 1)
            if which == 0: x += d
            else: y += d
        out.append((x / PREC, y / PREC))
    return out


def enc_num(v):
    v = ~(v << 1) if v < 0 else (v << 1)
    o = []
    while v >= 0x20:
        o.append(chr((0x20 | (v & 0x1f)) + 63)); v >>= 5
    o.append(chr(v + 63))
    return "".join(o)


def encode(pts):
    px = py = 0; o = []
    for x, y in pts:
        ix, iy = int(round(x * PREC)), int(round(y * PREC))
        o.append(enc_num(ix - px)); o.append(enc_num(iy - py))
        px, py = ix, iy
    return "".join(o)


def load_source():
    geo, lot, polys = {}, {}, []
    for f in ("kim.html", "moon.html"):
        path = os.path.join(REPO, f)
        if not os.path.exists(path):
            sys.exit("!! %s 이 없다. 전수조사 저장소 경로를 인자로 넘겨라.\n"
                     "   git clone https://github.com/Hyeokgi/gaeryeong-survey.git" % path)
        P, G, C, Gg, L = grab(f)
        base = len(polys)
        for i, s in enumerate(Gg):
            polys.append((int(C[i]), decode(s)))
        for p in P:
            k = (p["village"], p["jibun"])
            if p["id"] in G: geo.setdefault(k, tuple(G[p["id"]]))
            if p["id"] in L: lot.setdefault(k, base + L[p["id"]])
    return geo, lot, polys


# ── 대장에서 신청농지 뽑기 ────────────────────────────────────
ADDR = re.compile(EMD + r"\s+(\S+리)\s+(\S+)")


def num(v):
    if v is None: return None
    try: return round(float(v), 2)
    except (TypeError, ValueError): return None


def fmt(v):
    return "" if v is None else ("%.2f" % v).rstrip("0").rstrip(".")


def com(v):
    if v is None: return ""
    return format(int(v), ",") if float(v) == int(v) else format(v, ",.2f")


def load_rows():
    ws = openpyxl.load_workbook(XL, data_only=True).worksheets[0]
    recs = []
    for r in ws.iter_rows(min_row=8, values_only=True):
        if not r[5] or EMD not in str(r[5]):
            continue
        site = str(r[5]).strip()
        m = ADDR.search(site)
        if not m:
            print("   주소를 못 읽었다:", site); continue
        recs.append(dict(vil=m.group(1), jb=re.sub(r"번지$", "", m.group(2)).strip(), site=site,
                         who=str(r[2]).strip(), addr=str(r[3]).strip(), tel=str(r[4]).strip(),
                         ra=num(r[6]), aa=num(r[7]),
                         c1=str(r[8] or "").strip(), c2=str(r[9] or "").strip(),
                         item=str(r[10] or "").strip(), grp=str(r[11] or "").strip(),
                         sub=str(r[12] or "").strip(), prev=str(r[13] or "").strip(),
                         note=str(r[18] or "").strip()))
    return recs


def jkey(j):
    a = re.match(r"(\d+)(?:-(\d+))?", j)
    return (int(a.group(1)), int(a.group(2) or 0)) if a else (10 ** 9, 0)


def main():
    geo, lot, polys = load_source()
    print("전수조사 앱: 좌표 %d필지, 지적도 폴리곤 %d개" % (len(geo), len(polys)))

    recs = load_rows()
    # 리는 필지가 많은 곳부터, 그 안은 신청인 → 지번 순 (현장을 도는 순서와 맞춘다)
    vord = [v for v, _ in collections.Counter(x["vil"] for x in recs).most_common()]
    recs.sort(key=lambda x: (vord.index(x["vil"]), x["who"], jkey(x["jb"])))

    PARCELS, GEO, MAPL = [], {}, {}
    seqn = collections.Counter()
    for i, x in enumerate(recs, 1):
        pid = "n%d" % i
        seqn[(x["vil"], x["who"])] += 1
        PARCELS.append({
            "id": pid, "no": i, "village": x["vil"], "jibun": x["jb"],
            "label": "%s %s" % (x["vil"], x["jb"]), "addr": x["site"],
            "who": x["who"], "tel": x["tel"], "home": x["addr"],
            "seq": seqn[(x["vil"], x["who"])],
            "ra": fmt(x["ra"]), "aa": fmt(x["aa"]), "raT": com(x["ra"]), "aaT": com(x["aa"]),
            "c1": x["c1"], "c2": x["c2"], "item": x["item"],
            "grp": x["grp"], "sub": x["sub"], "prev": x["prev"], "note": x["note"],
            "fname": "%s %s_%s" % (x["who"], x["vil"], x["jb"]),
        })
        k = (x["vil"], x["jb"])
        if k in geo: GEO[pid] = geo[k]
        if k in lot: MAPL[pid] = lot[k]

    print("신청농지 %d필지 · 신청인 %d명 · 좌표 %d · 지적도 연결 %d"
          % (len(PARCELS), len({p["who"] for p in PARCELS}), len(GEO), len(MAPL)))
    for p in PARCELS:
        if p["id"] not in GEO:
            print("       좌표 없음: %s %s (%s)" % (p["village"], p["jibun"], p["who"]))

    # ── 지적도는 신청농지 주변만 남긴다 ──
    pts = list(GEO.values())
    if not pts:
        sys.exit("!! 좌표가 하나도 안 맞았다. 지번 형식을 확인해라.")
    lat0 = sum(p[1] for p in pts) / len(pts)
    dl = BUF / (111320 * math.cos(lat0 * math.pi / 180))
    dt = BUF / 110540
    boxes = [(p[0] - dl, p[1] - dt, p[0] + dl, p[1] + dt) for p in pts]

    def near(pl):
        xs = [p[0] for p in pl]; ys = [p[1] for p in pl]
        a, b, c, d = min(xs), min(ys), max(xs), max(ys)
        return any(a <= x1 and c >= x0 and b <= y1 and d >= y0 for x0, y0, x1, y1 in boxes)

    keep, remap, must = [], {}, set(MAPL.values())
    for i, (cat, pl) in enumerate(polys):
        if i in must or near(pl):
            remap[i] = len(keep); keep.append((cat, pl))
    MAPL = {k: remap[v] for k, v in MAPL.items() if v in remap}
    MAPC = "".join(str(c) for c, _ in keep)
    MAPG = [encode(pl) for _, pl in keep]
    print("지적도 %d → %d 폴리곤 (%.0f KB)"
          % (len(polys), len(keep), sum(len(s) for s in MAPG) / 1024))

    J = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))
    block = (
        "const PARCELS = %s;\n" % J(PARCELS) +
        "/* 필지 중심 좌표 [경도, 위도] · 2025년 농지전수조사 앱(kim/moon)의 GEO 에서 지번으로 맞춰 가져왔다.\n"
        "   그 조사의 대상이 아니었던 지번은 좌표가 없어 키 자체가 없다. tools/build_nongji.py 로 다시 만든다. */\n" +
        "const GEO = {%s};\n" % ",".join('"%s":[%.6f,%.6f]' % (k, v[0], v[1]) for k, v in GEO.items()) +
        "/* 연속지적도 필지 경계 · 신청농지 반경 %dm 안만 남겼다.\n"
        "   MAPC 는 지목 갈래, MAPG 는 폴리라인으로 감은 경계선, MAPL 은 우리 필지와의 연결이다. */\n" % BUF +
        'const MAPC="%s";\n' % MAPC +
        "const MAPG=%s;\n" % J(MAPG) +
        "const MAPL=%s;\n" % J(MAPL))

    t = io.open(APP, encoding="utf-8").read()
    old = re.search(r"const PARCELS = \[.*?\n(?:.*?\n)*?const MAPL=\{.*?\};\n", t, re.S)
    if not old:
        sys.exit("!! lee.html 에서 기존 데이터 블록을 못 찾았다.")
    io.open(APP, "w", encoding="utf-8", newline="\n").write(t[:old.start()] + block + t[old.end():])
    print("[OK] %s  %.0f KB" % (APP, os.path.getsize(APP) / 1024))


if __name__ == "__main__":
    main()
