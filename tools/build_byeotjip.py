# '2026년 볏짚환원사업 확정내역(최종)-이행점검.xlsx' 에서 신청농지를 뽑아
# lim.html 의 PARCELS / GEO / MAPC / MAPG / MAPL 블록을 새로 만들어 끼운다.
# 확정내역이 갱신되면 이 스크립트만 다시 돌리면 된다 (앱의 나머지 코드는 건드리지 않는다).
#
#   사용법:  python -u tools/build_byeotjip.py [gaeryeong-survey 저장소 경로]
#   필요:    openpyxl
#
# ── 개인정보 ──────────────────────────────────────────────
# 원본에는 생년월일 · 연락처 · 자택주소 · 은행명 · 계좌번호가 들어 있다.
# 이 저장소는 공개이고 앱은 GitHub Pages 로 나간다. 그래서 **아래 다섯 칸은 읽지도 않는다.**
# 앱에 들어가는 신원 정보는 신청인 성명뿐이다 (lee.html 과 같은 기준).
#
# ── 좌표 ──────────────────────────────────────────────────
# 좌표와 지적도는 2025년 농지전수조사 앱(kim.html · moon.html)에서 (리, 지번) 으로 맞춰 가져온다.
# 그 앱이 개령면만 다루므로 **545필지 중 8필지에만 좌표가 붙는다.**
# 나머지는 약도를 그릴 수 없어 카드에서 네이버지도로 넘긴다.

import openpyxl, re, io, json, math, os, sys, collections

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

XLNAME = "2026년 볏짚환원사업 확정내역(최종)-이행점검.xlsx"
APP = os.path.join(ROOT, "lim.html")
REPO = sys.argv[1] if len(sys.argv) > 1 else ROOT


def find_xl():
    """원본은 개인정보가 있어 저장소에 두지 않는다. 경로를 주지 않으면 가까운 곳부터 찾는다.
       직접 줄 수도 있다:  python tools/build_byeotjip.py <저장소 경로> <확정내역.xlsx>"""
    if len(sys.argv) > 2:
        if not os.path.exists(sys.argv[2]):
            sys.exit("!! 준 경로에 파일이 없다: %s" % sys.argv[2])
        return sys.argv[2]
    drive = os.path.splitdrive(ROOT)[0] + os.sep
    here = [ROOT, os.path.dirname(ROOT),
            os.path.join(os.path.expanduser("~"), "Downloads"),
            os.path.join(drive, "user", "Downloads")]
    for d in here:
        p = os.path.join(d, XLNAME)
        if os.path.exists(p):
            return p
    sys.exit("!! '%s' 을 찾지 못했다.\n   찾아본 곳:\n     %s\n   두 번째 인자로 경로를 직접 줘도 된다."
             % (XLNAME, "\n     ".join(here)))


XL = find_xl()

BUF = 700.0          # 좌표가 붙은 필지에서 이 거리(m) 안의 지적도만 남긴다 (약도 최대 배율이 1km)
PREC = 1e6

# 원본 열 번호 (0-based) — 읽지 않는 칸은 아예 적지 않았다
C_NO, C_JURI, C_NAME = 0, 1, 2
C_EMD, C_RI, C_JIBUN = 6, 7, 8
C_RA, C_AA = 9, 10
C_ECO, C_BIZ, C_HARV = 11, 12, 13
C_UNIT, C_FEE = 14, 15
C_NOTE = 20


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
            sys.exit("!! %s 이 없다. 전수조사 저장소 경로를 인자로 넘겨라." % path)
        P, G, C, Gg, L = grab(f)
        base = len(polys)
        for i, s in enumerate(Gg):
            polys.append((int(C[i]), decode(s)))
        for p in P:
            k = (p["village"], p["jibun"])
            if p["id"] in G: geo.setdefault(k, tuple(G[p["id"]]))
            if p["id"] in L: lot.setdefault(k, base + L[p["id"]])
    return geo, lot, polys


# ── 확정내역 읽기 ─────────────────────────────────────────────
def num(v):
    if v is None: return None
    try: return round(float(v), 2)
    except (TypeError, ValueError): return None


def fmt(v):
    return "" if v is None else ("%.2f" % v).rstrip("0").rstrip(".")


def com(v):
    if v is None: return ""
    return format(int(v), ",") if float(v) == int(v) else format(v, ",.2f")


def txt(v):
    """엑셀이 날짜·숫자로 바꿔 놓은 칸까지 사람이 쓴 모양으로 되돌린다"""
    if v is None: return ""
    if hasattr(v, "strftime"): return v.strftime("%m.%d.")
    if isinstance(v, float) and v == int(v): v = int(v)
    return str(v).strip()


def load_rows():
    """연번은 농가의 첫 행에만 있으므로 아래로 끌어 채운다. '계'·'합계' 행은 버린다."""
    ws = openpyxl.load_workbook(XL, data_only=True).worksheets[0]
    recs, fno = [], None
    for r in ws.iter_rows(min_row=7, values_only=True):
        tag = txt(r[C_NO])
        if tag in ("계", "합계"):
            continue
        if isinstance(r[C_NO], int):
            fno = r[C_NO]
        if not (r[C_NAME] or r[C_EMD]):
            continue
        emd, ri, jb = txt(r[C_EMD]), txt(r[C_RI]), txt(r[C_JIBUN])
        if not jb:
            print("   지번이 비어 있다 — 건너뛴다: 연번 %s %s %s %s" % (fno, txt(r[C_NAME]), emd, ri))
            continue
        recs.append(dict(fno=fno, juri=txt(r[C_JURI]), who=txt(r[C_NAME]),
                         emd=emd, ri=ri, jb=re.sub(r"번지$", "", jb).strip(),
                         ra=num(r[C_RA]), aa=num(r[C_AA]),
                         eco=txt(r[C_ECO]), biz=txt(r[C_BIZ]), harv=txt(r[C_HARV]),
                         unit=num(r[C_UNIT]), fee=num(r[C_FEE]),
                         note=txt(r[C_NOTE]).replace("\n", " ")))
    return recs


def jkey(j):
    a = re.match(r"(\d+)(?:-(\d+))?", j)
    return (int(a.group(1)), int(a.group(2) or 0)) if a else (10 ** 9, 0)


def metres(a, b):
    dx = (b[0] - a[0]) * 111320 * math.cos(a[1] * math.pi / 180)
    dy = (b[1] - a[1]) * 110540
    return math.hypot(dx, dy)


def main():
    geo, lot, polys = load_source()
    print("전수조사 앱: 좌표 %d필지, 지적도 폴리곤 %d개 (개령면뿐이다)" % (len(geo), len(polys)))

    recs = load_rows()
    farms = len({r["fno"] for r in recs})
    print("확정내역: %d필지 · %d농가" % (len(recs), farms))

    # 읍면동은 필지가 많은 곳부터. 그 안은 리동 → 농가(연번) → 지번 순.
    # 좌표가 8필지뿐이라 lee.html 처럼 이동경로로 순번을 매길 수가 없다. 지번 순이 차선이다.
    eord = [v for v, _ in collections.Counter(r["emd"] for r in recs).most_common()]
    rord = {}
    for v in eord:
        rord[v] = [x for x, _ in collections.Counter(
            r["ri"] for r in recs if r["emd"] == v).most_common()]
    recs.sort(key=lambda r: (eord.index(r["emd"]), rord[r["emd"]].index(r["ri"]),
                             r["fno"], jkey(r["jb"])))

    PARCELS, GEO, MAPL = [], {}, {}
    seqn = collections.Counter()
    for i, x in enumerate(recs, 1):
        pid = "b%d" % i
        seqn[x["emd"]] += 1
        PARCELS.append({
            "id": pid, "no": i, "fno": x["fno"], "seq": seqn[x["emd"]],
            "village": x["emd"], "ri": x["ri"], "jibun": x["jb"],
            "juri": x["juri"], "who": x["who"],
            "label": "%s %s %s" % (x["emd"], x["ri"], x["jb"]),
            "addr": "김천시 %s %s %s" % (x["emd"], x["ri"], x["jb"]),
            "ra": fmt(x["ra"]), "aa": fmt(x["aa"]),
            "raT": com(x["ra"]), "aaT": com(x["aa"]),
            "fee": fmt(x["fee"]), "feeT": com(x["fee"]), "unit": fmt(x["unit"]),
            "eco": x["eco"], "biz": x["biz"], "harv": x["harv"], "note": x["note"],
            "fname": "%s %s %s_%s" % (x["who"], x["emd"], x["ri"], x["jb"]),
        })
        k = (x["ri"], x["jb"])
        if k in geo: GEO[pid] = geo[k]
        if k in lot: MAPL[pid] = lot[k]

    print("좌표 %d / %d필지 · 지적도 연결 %d" % (len(GEO), len(PARCELS), len(MAPL)))
    for v in eord:
        vs = [p for p in PARCELS if p["village"] == v]
        print("   %-5s %3d필지 · %2d리동 · %2d농가 · 좌표 %d"
              % (v, len(vs), len({p["ri"] for p in vs}), len({p["fno"] for p in vs}),
                 sum(1 for p in vs if p["id"] in GEO)))

    # ── 지적도는 좌표가 붙은 필지 주변만 남긴다 ──
    pts = list(GEO.values())
    keep, remap = [], {}
    if pts:
        lat0 = sum(p[1] for p in pts) / len(pts)
        dl = BUF / (111320 * math.cos(lat0 * math.pi / 180))
        dt = BUF / 110540
        boxes = [(p[0] - dl, p[1] - dt, p[0] + dl, p[1] + dt) for p in pts]

        def near(pl):
            xs = [p[0] for p in pl]; ys = [p[1] for p in pl]
            a, b, c, d = min(xs), min(ys), max(xs), max(ys)
            return any(a <= x1 and c >= x0 and b <= y1 and d >= y0
                       for x0, y0, x1, y1 in boxes)

        must = set(MAPL.values())
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
        "/* 필지 중심 좌표 [경도, 위도] · 2025년 농지전수조사 앱(kim/moon)의 GEO 에서 지번으로 맞춰 왔다.\n"
        "   그 조사가 개령면만 다뤄 545필지 중 8필지에만 붙는다. tools/build_byeotjip.py 로 다시 만든다. */\n" +
        "const GEO = {%s};\n" % ",".join('"%s":[%.6f,%.6f]' % (k, v[0], v[1])
                                         for k, v in GEO.items()) +
        "/* 연속지적도 필지 경계 · 좌표가 붙은 필지 반경 %dm 안만 남겼다.\n"
        "   MAPC 는 지목 갈래, MAPG 는 폴리라인으로 감은 경계선, MAPL 은 우리 필지와의 연결이다. */\n" % BUF +
        'const MAPC="%s";\n' % MAPC +
        "const MAPG=%s;\n" % J(MAPG) +
        "const MAPL=%s;\n" % J(MAPL))

    t = io.open(APP, encoding="utf-8").read()
    old = re.search(r"const PARCELS = \[.*?\n(?:.*?\n)*?const MAPL=\{.*?\};\n", t, re.S)
    if not old:
        sys.exit("!! lim.html 에서 기존 데이터 블록을 못 찾았다.")
    io.open(APP, "w", encoding="utf-8", newline="\n").write(t[:old.start()] + block + t[old.end():])
    print("[OK] %s  %.0f KB" % (APP, os.path.getsize(APP) / 1024))


if __name__ == "__main__":
    main()
