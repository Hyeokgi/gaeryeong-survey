# 연속지적도 shapefile 에서 각 조사원 담당 리의 필지 경계를 뽑아 kim.html / moon.html 에 심는다
# 조사원마다 담당 리가 다르므로 파일별로 필요한 리만 넣어 용량을 절반으로 줄인다.
import shapefile, re, io, os, json, math, sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHP = os.path.join(REPO, "_source", "7223_202608210201_LP_PA_CBND_BUBUN_연속지적도")

TOL = 4e-6          # 약 0.4m — 화면에서 구분되지 않는 꼭짓점은 버린다
PREC = 1e6          # 좌표 정밀도 약 0.11m

# 지목을 화면에서 구분할 몇 갈래로 묶는다
CAT = {"답": 0, "전": 1, "과": 1, "목": 1,
       "도": 2, "철": 2,
       "구": 3, "천": 3, "유": 3, "제": 3,
       "대": 4, "창": 4, "공": 4, "학": 4, "종": 4, "묘": 4, "주": 4,
       "산": 5, "임": 5}
CATN = ["답", "전·과수원", "도로", "물길", "대지·건물", "임야", "그 밖"]


def dp(pts, tol):
    """더글러스-포이커 단순화. 재귀 대신 스택을 쓴다."""
    n = len(pts)
    if n < 3:
        return pts
    keep = [False] * n
    keep[0] = keep[n - 1] = True
    stack = [(0, n - 1)]
    while stack:
        a, b = stack.pop()
        if b <= a + 1:
            continue
        ax, ay = pts[a]; bx, by = pts[b]
        dx, dy = bx - ax, by - ay
        dd = dx * dx + dy * dy
        far, fd = -1, tol
        for i in range(a + 1, b):
            px, py = pts[i]
            if dd == 0:
                d = math.hypot(px - ax, py - ay)
            else:
                t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / dd))
                d = math.hypot(px - (ax + t * dx), py - (ay + t * dy))
            if d > fd:
                far, fd = i, d
        if far > 0:
            keep[far] = True
            stack.append((a, far)); stack.append((far, b))
    return [pts[i] for i in range(n) if keep[i]]


def enc_num(v):
    """구글 폴리라인 방식 — 부호를 섞고 5비트씩 끊어 아스키로 만든다"""
    v = ~(v << 1) if v < 0 else (v << 1)
    out = []
    while v >= 0x20:
        out.append(chr((0x20 | (v & 0x1f)) + 63))
        v >>= 5
    out.append(chr(v + 63))
    return "".join(out)


def encode(ring):
    out, lx, ly = [], 0, 0
    for x, y in ring:
        ix, iy = int(round(x * PREC)), int(round(y * PREC))
        out.append(enc_num(ix - lx)); out.append(enc_num(iy - ly))
        lx, ly = ix, iy
    return "".join(out)


def load():
    sf = shapefile.Reader(SHP, encoding="utf-8", encodingErrors="replace")
    I = {f[0]: i for i, f in enumerate(sf.fields[1:])}
    digits = lambda s: re.sub(r"\D", "", str(s) or "")
    out = []
    for sh, r in zip(sf.iterShapes(), sf.records()):
        li = str(r[I["li_nm"]]).strip()
        if not li:
            continue
        jm = (re.sub(r"[\d\s\-]", "", str(r[I["jibun"]])) or "?")[:1]
        bon, bu = digits(r[I["bonbun"]]), digits(r[I["bubun"]])
        parts = list(sh.parts) + [len(sh.points)]
        rings = [sh.points[parts[i]:parts[i + 1]] for i in range(len(parts) - 1)]
        rings = [g for g in rings if len(g) >= 4]
        if not rings:
            continue
        out.append(dict(li=li, cat=CAT.get(jm, 6), jm=jm,
                        key=(li, int(bon or 0), int(bu or 0)),
                        rings=max(rings, key=len)))          # 구멍은 무시하고 가장 큰 테두리만
    return out


def build(fname, cad):
    path = os.path.join(REPO, fname)
    t = io.open(path, encoding="utf-8").read()
    parcels = json.loads(re.search(r"const PARCELS = (\[.*?\]);\n", t, re.S).group(1))
    villages = sorted({p["village"] for p in parcels})

    mine = [c for c in cad if c["li"] in villages]
    idx = {}
    enc, cats = [], []
    for c in mine:
        ring = dp(c["rings"], TOL)
        if len(ring) < 4:
            ring = c["rings"]
        enc.append(encode(ring)); cats.append(c["cat"])
        idx.setdefault(c["key"], len(enc) - 1)

    linked, miss = {}, []
    for p in parcels:
        m = re.match(r"^(\d+)(?:-(\d+))?$", p["jibun"].strip())
        k = (p["village"], int(m.group(1)), int(m.group(2) or 0)) if m else None
        if k in idx:
            linked[p["id"]] = idx[k]
        else:
            miss.append(p)

    block = ("/* 연속지적도 필지 경계 · %s 담당 %s\n"
             "   MAPC 는 지목 갈래, MAPG 는 폴리라인으로 감은 경계선, MAPL 은 우리 필지와의 연결이다.\n"
             "   tools/build_cadastral.py 로 다시 만든다. */\n"
             "const MAPC=\"%s\";\nconst MAPG=%s;\nconst MAPL=%s;\n"
             % (fname, "·".join(villages),
                "".join(str(c) for c in cats),
                json.dumps(enc, ensure_ascii=False, separators=(",", ":")),
                json.dumps(linked, separators=(",", ":"))))

    old = re.search(r"/\* 연속지적도 필지 경계.*?\nconst MAPL=\{.*?\};\n", t, re.S)
    if old:
        t = t[:old.start()] + block + t[old.end():]
    else:
        a = re.search(r"const GEO = \{.*?\};\n", t, re.S)
        t = t[:a.end()] + block + t[a.end():]

    io.open(path, "w", encoding="utf-8", newline="").write(t)
    pts = sum(len(e) for e in enc)
    print("[OK] %-10s %s" % (fname, "·".join(villages)))
    print("       필지 %d개, 부호화 %d글자 (%.2fMB), 우리 필지 연결 %d/%d"
          % (len(enc), pts, len(block) / 1048576, len(linked), len(parcels)))
    for p in miss:
        print("       경계 없음: %s %s (%d일차 %02d번)" % (p["village"], p["jibun"], p["day"], p["seq"]))


if __name__ == "__main__":
    cad = load()
    print("지적도 %d필지 읽음\n" % len(cad))
    for f in ("kim.html", "moon.html"):
        build(f, cad)
