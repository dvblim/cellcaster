#!/usr/bin/env python3
"""Reference port of the CellCaster renderer.

A line-by-line Python port of src/Raycaster.bas, src/Textures.bas and the
level/movement code in src/World.bas. It exists for two reasons:

  * to check the VBA maths outside Excel (frames are compared number by number);
  * to draw docs/preview.png without needing a Windows machine.

Constants, the palette and the level are read from the .bas files, so this
script cannot drift away from them silently.

    python tools/reference_renderer.py            # writes docs/preview.png
    python tools/reference_renderer.py --check    # validates the level only
"""
import math
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"


# --- read constants, palette and level straight from the VBA sources ---------
def read(name):
    return (SRC / name).read_text(encoding="ascii").replace("\r\n", "\n")


def constants(text):
    out = {}
    for m in re.finditer(r"Public Const (\w+) As (\w+) = ([^'\n]+)", text):
        name, kind, raw = m.group(1), m.group(2), m.group(3).strip()
        if kind == "String":
            out[name] = raw.strip('"')
        else:
            out[name] = float(raw.rstrip("#")) if kind == "Double" else int(raw)
    return out


CFG = constants(read("Config.bas"))
WORLD = constants(read("World.bas"))
SCREEN_W, SCREEN_H = CFG["SCREEN_W"], CFG["SCREEN_H"]
FOV_PLANE, FOG = CFG["FOV_PLANE"], CFG["FOG_DISTANCE"]
NUM_SHADES, TEX = CFG["NUM_SHADES"], CFG["TEX_SIZE"]
MAP_W, MAP_H = WORLD["MAP_W"], WORLD["MAP_H"]

COLORS = {}  # name -> id
for m in re.finditer(r"^\s+(col\w+) = (\d+)", read("Config.bas"), re.M):
    COLORS[m.group(1)] = int(m.group(2))
BASE = {}    # id -> (r, g, b)
for m in re.finditer(r"Case (col\w+):\s+BaseColor = RGB\((\d+), (\d+), (\d+)\)", read("Config.bas")):
    BASE[COLORS[m.group(1)]] = tuple(int(m.group(i)) for i in (2, 3, 4))
C = COLORS

LEVEL = re.findall(r'AddRow "([^"]+)"', read("World.bas"))


def palette_rgb(index):
    if index == 0:
        return (0, 0, 0)
    base = BASE[(index - 1) // NUM_SHADES]
    level = (index - 1) % NUM_SHADES + 1
    f = 0.15 + 0.85 * level / NUM_SHADES
    return tuple(int(c * f) for c in base)


# --- World --------------------------------------------------------------------
class World:
    def __init__(self):
        self.map = [[0] * MAP_H for _ in range(MAP_W)]
        self.gems = []
        for y, row in enumerate(LEVEL):
            for x, ch in enumerate(row):
                if ch.isdigit() and ch != "0":
                    self.map[x][y] = int(ch)
                elif ch == "P":
                    self.px, self.py = x + 0.5, y + 0.5
                elif ch == "*":
                    self.gems.append([x + 0.5, y + 0.5, True])
        self.dx, self.dy = 1.0, 0.0
        self.plx, self.ply = 0.0, FOV_PLANE

    def is_wall(self, x, y):
        mx, my = math.floor(x), math.floor(y)
        if mx < 0 or my < 0 or mx >= MAP_W or my >= MAP_H:
            return True
        return self.map[mx][my] != 0

    def try_move(self, dx, dy):
        r = CFG["PLAYER_RADIUS"]
        sgn = lambda v: (v > 0) - (v < 0)
        nx = self.px + dx
        if not self.is_wall(nx + sgn(dx) * r, self.py - r) and not self.is_wall(nx + sgn(dx) * r, self.py + r):
            self.px = nx
        ny = self.py + dy
        if not self.is_wall(self.px - r, ny + sgn(dy) * r) and not self.is_wall(self.px + r, ny + sgn(dy) * r):
            self.py = ny

    def rotate(self, a):
        c, s = math.cos(a), math.sin(a)
        self.dx, self.dy = self.dx * c - self.dy * s, self.dx * s + self.dy * c
        self.plx, self.ply = self.plx * c - self.ply * s, self.plx * s + self.ply * c

    def collect(self):
        n = 0
        for g in self.gems:
            if g[2] and (g[0] - self.px) ** 2 + (g[1] - self.py) ** 2 < CFG["PICKUP_RADIUS"] ** 2:
                g[2] = False
                n += 1
        return n


# --- Textures -----------------------------------------------------------------
def _hash(x, y, seed):
    return (x * 73 + y * 151 + seed * 199 + x * y * 31) % 11


def _brick(x, y):
    course = y // 4
    if (y & 3) == 3 or ((x + (course & 1) * 4) & 7) == 7:
        return C["colMortar"]
    return C["colBrick"]


def _stone(x, y):
    if (x & 7) == 0 or (y & 7) == 0 or _hash(x, y, 1) == 0:
        return C["colStoneDark"]
    return C["colStone"]


def _wood(x, y):
    if (x & 3) == 3 or _hash(x // 4, y, 2) < 2:
        return C["colWoodDark"]
    return C["colWood"]


def _metal(x, y):
    if x == 0 or y == 0 or x == TEX - 1 or y == TEX - 1:
        return C["colStoneDark"]
    if x in (2, 13) and y in (2, 13):
        return C["colWhite"]
    if y in (7, 8):
        return C["colStoneDark"]
    return C["colMetal"]


def _moss(x, y):
    return C["colMoss"] if _hash(x, y, 3) < y // 3 else _stone(x, y)


def _gem(x, y):
    d = abs(2 * x - (TEX - 1)) + abs(2 * y - (TEX - 1))
    if d > TEX - 2:
        return -1
    if x < TEX // 2 and y < TEX // 2 and d <= TEX // 2:
        return C["colWhite"]
    return C["colGold"]


TEXFN = {1: _brick, 2: _stone, 3: _wood, 4: _metal, 5: _moss}
TEXTURES = {k: [[fn(x, y) for y in range(TEX)] for x in range(TEX)] for k, fn in TEXFN.items()}
GEM = [[_gem(x, y) for y in range(TEX)] for x in range(TEX)]


# --- Raycaster ----------------------------------------------------------------
_B = [0, 8, 2, 10, 12, 4, 14, 6, 3, 11, 1, 9, 15, 7, 13, 5]
BAYER = [[(_B[by * 4 + bx] + 0.5) / 16.0 for by in range(4)] for bx in range(4)]
SCALE = SCREEN_W / (2.0 * FOV_PLANE)


def brightness(dist):
    return max(0.0, 1.0 - dist / FOG)


def shade(bright, bx, by):
    return min(NUM_SHADES, math.floor(bright * NUM_SHADES + BAYER[bx][by]))


def pal(color, level):
    return color * NUM_SHADES + level if level > 0 else 0


def render(w, t, minimap=False):
    """Returns frame[y][x] of palette indices, like the VBA frame (0-based)."""
    frame = [[0] * SCREEN_W for _ in range(SCREEN_H)]
    half = SCREEN_H // 2
    zbuf = [0.0] * SCREEN_W

    for r in range(half):
        y, yc = half + r, half - 1 - r
        dist = (SCALE * 0.5) / (r + 0.5)
        sx = dist * 2.0 * w.plx / SCREEN_W
        sy = dist * 2.0 * w.ply / SCREEN_W
        fx = w.px + dist * (w.dx - w.plx) + sx * 0.5
        fy = w.py + dist * (w.dy - w.ply) + sy * 0.5
        for x in range(SCREEN_W):
            bx = x & 3
            level = shade(brightness(dist), bx, y & 3)
            if level == 0:
                frame[y][x] = 0
            elif ((math.floor(fx) + math.floor(fy)) & 1) == 0:
                frame[y][x] = C["colFloor"] * NUM_SHADES + level
            else:
                frame[y][x] = C["colStoneDark"] * NUM_SHADES + level
            frame[yc][x] = pal(C["colCeiling"], shade(brightness(dist) * 0.9, bx, yc & 3))
            fx += sx
            fy += sy

    for x in range(SCREEN_W):
        cam = 2.0 * (x + 0.5) / SCREEN_W - 1.0
        rx, ry = w.dx + w.plx * cam, w.dy + w.ply * cam
        mx, my = math.floor(w.px), math.floor(w.py)
        ddx = 1e30 if rx == 0 else abs(1.0 / rx)
        ddy = 1e30 if ry == 0 else abs(1.0 / ry)
        if rx < 0:
            stx, sdx = -1, (w.px - mx) * ddx
        else:
            stx, sdx = 1, (mx + 1.0 - w.px) * ddx
        if ry < 0:
            sty, sdy = -1, (w.py - my) * ddy
        else:
            sty, sdy = 1, (my + 1.0 - w.py) * ddy
        while True:
            if sdx < sdy:
                sdx += ddx
                mx += stx
                side = 0
            else:
                sdy += ddy
                my += sty
                side = 1
            if mx < 0 or my < 0 or mx >= MAP_W or my >= MAP_H or w.map[mx][my] != 0:
                break
        dist = max(0.05, sdx - ddx if side == 0 else sdy - ddy)
        zbuf[x] = dist
        inside = 0 <= mx < MAP_W and 0 <= my < MAP_H
        tex = TEXTURES[w.map[mx][my] if inside else 1]

        wx = w.py + dist * ry if side == 0 else w.px + dist * rx
        wx -= math.floor(wx)
        tx = min(TEX - 1, math.floor(wx * TEX))
        if (side == 0 and rx < 0) or (side == 1 and ry > 0):
            tx = TEX - 1 - tx

        line_h = SCALE / dist
        y_top = (SCREEN_H - line_h) * 0.5
        y0 = max(0, math.floor(y_top + 0.5))
        y1 = min(SCREEN_H - 1, math.floor(y_top + line_h - 0.5))
        bright = brightness(dist) * (CFG["SIDE_SHADE"] if side == 1 else 1.0)
        levels = [shade(bright, x & 3, k) for k in range(4)]
        tstep = TEX / line_h
        tpos = (y0 + 0.5 - y_top) * tstep
        for y in range(y0, y1 + 1):
            level = levels[y & 3]
            if level == 0:
                frame[y][x] = 0
            else:
                frame[y][x] = tex[tx][min(TEX - 1, math.floor(tpos))] * NUM_SHADES + level
            tpos += tstep

    inv = 1.0 / (w.plx * w.dy - w.dx * w.ply)
    vis = []
    for i, g in enumerate(w.gems, start=1):
        if g[2]:
            cz = inv * (-w.ply * (g[0] - w.px) + w.plx * (g[1] - w.py))
            if cz > 0.2:
                vis.append((cz, i))
    vis.sort(key=lambda p: -p[0])  # stable, like the insertion sort
    for cz, i in vis:
        g = w.gems[i - 1]
        cx = inv * (w.dy * (g[0] - w.px) - w.dx * (g[1] - w.py))
        size = SCALE / cz * CFG["GEM_SCALE"]
        if size < 1.0:
            continue
        ctx = (SCREEN_W * 0.5) * (1.0 + cx / cz)
        cty = SCREEN_H * 0.5 + math.sin(t * 3.0 + i) * size * 0.15
        x0, x1 = math.floor(ctx - size * 0.5 + 0.5), math.floor(ctx + size * 0.5 - 0.5)
        y0, y1 = math.floor(cty - size * 0.5 + 0.5), math.floor(cty + size * 0.5 - 0.5)
        bright = max(0.5, brightness(cz))
        for sx in range(x0, x1 + 1):
            if 0 <= sx < SCREEN_W and cz < zbuf[sx]:
                u = min(TEX - 1, max(0, math.floor((sx + 0.5 - (ctx - size * 0.5)) * TEX / size)))
                for sy in range(y0, y1 + 1):
                    if 0 <= sy < SCREEN_H:
                        v = min(TEX - 1, max(0, math.floor((sy + 0.5 - (cty - size * 0.5)) * TEX / size)))
                        texel = GEM[u][v]
                        if texel >= 0:
                            frame[sy][sx] = texel * NUM_SHADES + shade(bright, sx & 3, sy & 3)

    if minimap and MAP_W + 2 <= SCREEN_W and MAP_H + 2 <= SCREEN_H:
        for y in range(MAP_H):
            for x in range(MAP_W):
                frame[y + 1][x + 1] = C["colStone"] * NUM_SHADES + 2 if w.map[x][y] else 0
        for g in w.gems:
            if g[2]:
                frame[math.floor(g[1]) + 1][math.floor(g[0]) + 1] = C["colGold"] * NUM_SHADES + NUM_SHADES
        hx, hy = math.floor(w.px + w.dx * 1.5), math.floor(w.py + w.dy * 1.5)
        if 0 <= hx < MAP_W and 0 <= hy < MAP_H and w.map[hx][hy] == 0:
            frame[hy + 1][hx + 1] = C["colWhite"] * NUM_SHADES + 2
        frame[math.floor(w.py) + 1][math.floor(w.px) + 1] = C["colWhite"] * NUM_SHADES + NUM_SHADES
    return frame


# --- Level validation ---------------------------------------------------------
def check_level():
    problems = []
    if len(LEVEL) != MAP_H:
        problems.append(f"level has {len(LEVEL)} rows, expected {MAP_H}")
    for i, row in enumerate(LEVEL):
        if len(row) != MAP_W:
            problems.append(f"row {i} has {len(row)} columns, expected {MAP_W}")
        bad = set(row) - set(".P*") - {str(d) for d in range(1, CFG["NUM_TEXTURES"] + 1)}
        if bad:
            problems.append(f"row {i} has unknown symbols {sorted(bad)}")
    if problems:
        return problems
    if sum(row.count("P") for row in LEVEL) != 1:
        problems.append("level needs exactly one player start (P)")
        return problems
    edge = LEVEL[0] + LEVEL[-1] + "".join(r[0] + r[-1] for r in LEVEL)
    if not all(ch.isdigit() for ch in edge):
        problems.append("the outer ring must be solid wall")
    w = World()
    seen, todo = set(), [(math.floor(w.px), math.floor(w.py))]
    while todo:
        x, y = todo.pop()
        if (x, y) in seen or w.map[x][y]:
            continue
        seen.add((x, y))
        todo += [(x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)]
    for gx, gy, _ in w.gems:
        if (math.floor(gx), math.floor(gy)) not in seen:
            problems.append(f"gem at {gx},{gy} cannot be reached")
    return problems


def save_png(frames, path, scale=5):
    from PIL import Image
    pad = 2
    sheet = Image.new("RGB", (SCREEN_W * scale, (SCREEN_H * scale + pad) * len(frames) - pad), (30, 30, 36))
    for n, frame in enumerate(frames):
        img = Image.new("RGB", (SCREEN_W, SCREEN_H))
        img.putdata([palette_rgb(v) for row in frame for v in row])
        sheet.paste(img.resize((SCREEN_W * scale, SCREEN_H * scale), Image.NEAREST), (0, n * (SCREEN_H * scale + pad)))
    sheet.save(path)


if __name__ == "__main__":
    issues = check_level()
    for issue in issues:
        print("LEVEL ERROR:", issue)
    if issues:
        sys.exit(1)
    print(f"level ok: {MAP_W}x{MAP_H}, {len(World().gems)} gems, all reachable")
    if "--check" in sys.argv:
        sys.exit(0)

    shots = []
    w = World(); w.px, w.py = 3.5, 4.5; w.rotate(0.35)
    shots.append(render(w, 0.4, minimap=True))
    w = World(); w.px, w.py = 16.6, 17.5; w.rotate(-0.25)
    shots.append(render(w, 1.1))
    out = ROOT / "docs" / "preview.png"
    save_png(shots, out)
    print("wrote", out.relative_to(ROOT))
