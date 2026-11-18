"""簡易軌跡シミュレーター（NTX 向け、N ブロック単位）。

対応:
  - G0 / G1 / G2 / G3（R 指定・I/J/K 指定）、G90 / G91、U / V / W（X / Y / Z のインクリメンタル）
  - X は直径値（旋削・ミーリング・極座標補間すべて）、I は半径値
  - 極座標補間 G112〜G113（X は直径値、C は mm の直線座標）
  - 穴あけ固定サイクル G81〜G89（R は初期点からのインクリメンタル、戻りは初期点）
  - G28 / G30 / G53 / G361 は線を切るだけ（機械原点の座標は不明なため）
未対応（notices に記録）:
  - 旋削固定サイクル G70〜G76、変数（#）・マクロ（G65/G66）・IF/GOTO/WHILE
  - 同時5軸 G43.4、円筒補間、径補正、B 軸、工具形状
C 軸割り出し（G112 の外の C 指令）は、ミーリングの点を C 角度ぶん回してワーク基準の端面図に描く。
"""
from __future__ import annotations

import math
import re
import tkinter as tk
from bisect import bisect_right
from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# 解析
# ---------------------------------------------------------------------------

SPACE_TURN = "turn"     # 点 = (Z, X半径, Y)
SPACE_MILL = "mill"     # 点 = (X半径, Y, Z)
SPACE_POLAR = "polar"   # 点 = (X半径, C, Z)
SPACE_MILL_C = "mill_c"  # 点 = (ワークX, ワークY, Z, 機械X半径)  C 割り出しを反映

# C のプラス方向。+1: C+ でワークが反時計回り（端面を正面から見て）。逆なら -1 にする（仮決め）
C_SIGN = 1.0

# 画面ごとの表示定義: (タイトル, 横軸の添字, 縦軸の添字, 横軸名, 縦軸名)
SPACE_VIEWS = {
    SPACE_TURN: [("旋削  Z–X", 0, 1, "Z", "X(半径)")],
    SPACE_MILL: [("ミーリング  X–Y", 0, 1, "X(半径)", "Y"), ("ミーリング  X–Z", 0, 2, "X(半径)", "Z")],
    SPACE_POLAR: [("極座標補間  X–C（端面）", 0, 1, "X(半径)", "C"), ("極座標補間  X–Z", 0, 2, "X(半径)", "Z")],
    SPACE_MILL_C: [("端面（ワーク基準・C割り出し反映）", 0, 1, "X", "Y"), ("機械座標  X–Z", 3, 2, "X(半径)", "Z")],
}

RAPID_MM_MIN = 20000.0      # 早送りの見かけ速度（実機値は不明なので見た目用）
DEFAULT_FEED_MM_MIN = 1000.0
DEFAULT_RPM_LIMIT = 10000.0
ARC_STEP_RAD = math.pi / 90  # 円弧の分割（2°）

_WORD_RE = re.compile(r"([A-Z])([+-]?(?:\d+\.?\d*|\.\d+))")
_MACRO_RE = re.compile(r"#|\[|\b(?:IF|GOTO|WHILE|DO|END)\b")


@dataclass
class Segment:
    line_no: int
    kind: str                       # "rapid" | "feed"
    pts: list                       # 空間座標の 3 要素タプルの列
    weight: float = 0.0             # アニメーション用の相対時間
    c0: float | None = None         # 始点・終点の C 角度（割り出し、度）
    c1: float | None = None


@dataclass
class SimResult:
    space: str
    segments: list = field(default_factory=list)
    notices: list = field(default_factory=list)   # [(行番号, メッセージ)]
    dropped: int = 0                # 表示空間が違うため省略した移動数
    c_index: bool = False


def _strip_comments(line: str) -> str:
    return re.sub(r"\([^)]*\)", "", line)


def _parse_words(code: str) -> list[tuple[str, float, str]]:
    words = []
    for letter, raw in _WORD_RE.findall(code):
        try:
            words.append((letter, float(raw), raw))
        except ValueError:
            continue
    return words


def _point(space: str, pos: dict) -> tuple | None:
    x, y, z, c = pos.get("X"), pos.get("Y"), pos.get("Z"), pos.get("C")
    if space == SPACE_TURN:
        if x is None or z is None:
            return None
        return (z, x / 2.0, y if y is not None else 0.0)
    if space == SPACE_POLAR:
        if x is None or z is None or c is None:
            return None
        return (x / 2.0, c, z)
    if x is None or z is None:
        return None
    return (x / 2.0, y if y is not None else 0.0, z)


def _arc_plane(space: str, plane: int) -> tuple[int, int, int, str, str]:
    """円弧平面: (横の添字, 縦の添字, 残りの添字, 横の中心ワード, 縦の中心ワード)"""
    if space == SPACE_TURN:
        return 0, 1, 2, "K", "I"
    if space == SPACE_POLAR:
        return 0, 1, 2, "I", "J"
    if plane == 18:
        return 2, 0, 1, "K", "I"
    if plane == 19:
        return 1, 2, 0, "J", "K"
    return 0, 1, 2, "I", "J"


def arc_points(p0, p1, ia, ib, iw, center, clockwise, full=False):
    a0, b0, a1, b1 = p0[ia], p0[ib], p1[ia], p1[ib]
    ca, cb = center
    r0 = math.hypot(a0 - ca, b0 - cb)
    r1 = math.hypot(a1 - ca, b1 - cb)
    t0 = math.atan2(b0 - cb, a0 - ca)
    t1 = math.atan2(b1 - cb, a1 - ca)
    if clockwise:
        sweep = (t0 - t1) % (2 * math.pi)
    else:
        sweep = (t1 - t0) % (2 * math.pi)
    if sweep < 1e-9 and full:
        sweep = 2 * math.pi
    n = min(720, max(2, math.ceil(sweep / ARC_STEP_RAD)))
    pts = []
    for k in range(n + 1):
        f = k / n
        t = t0 - sweep * f if clockwise else t0 + sweep * f
        r = r0 + (r1 - r0) * f
        p = [0.0, 0.0, 0.0]
        p[ia] = ca + r * math.cos(t)
        p[ib] = cb + r * math.sin(t)
        p[iw] = p0[iw] + (p1[iw] - p0[iw]) * f
        pts.append(tuple(p))
    pts[0] = tuple(p0)
    pts[-1] = tuple(p1)
    return pts


def _r_center(p0, p1, ia, ib, radius, clockwise):
    a0, b0, a1, b1 = p0[ia], p0[ib], p1[ia], p1[ib]
    da, db = a1 - a0, b1 - b0
    chord = math.hypot(da, db)
    if chord < 1e-12:
        return None
    h = math.sqrt(max(radius * radius - (chord / 2) ** 2, 0.0))
    na, nb = -db / chord, da / chord          # 進行方向の左側
    sign = 1.0 if not clockwise else -1.0
    if radius < 0:
        sign = -sign
    return (a0 + da / 2 + sign * h * na, b0 + db / 2 + sign * h * nb)


def _polyline_length(pts) -> float:
    return sum(math.dist(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def simulate_block(rows: list[tuple[int, str]], mode: str = "") -> SimResult:
    """N ブロックの行 [(行番号, 行)] を解釈して軌跡を作る。mode は "TURN" / "MILL" / ""。"""
    base_space = SPACE_TURN if mode == "TURN" else SPACE_MILL
    if not mode:
        joined = " ".join(_strip_comments(line).upper() for _, line in rows)
        base_space = SPACE_TURN if re.search(r"G18(?!\d)", joined) and not re.search(r"G1[79](?!\d)", joined) else SPACE_MILL

    pos: dict[str, float | None] = {"X": None, "Y": 0.0, "Z": None, "C": None}
    motion = 0
    absolute = True
    plane = 18 if base_space == SPACE_TURN else 17
    polar = False
    feed = None
    per_rev = base_space == SPACE_TURN
    spindle = 0.0
    css = False
    rpm_limit = DEFAULT_RPM_LIMIT
    cycle = None            # {"code", "axis", "bottom", "r_inc"}
    segments: list[tuple[str, Segment]] = []
    notices: list[tuple[int, str]] = []
    seen_notice: set[tuple[int, str]] = set()
    c_index = False

    def note(line_no: int, msg: str) -> None:
        key = (line_no, msg)
        if key not in seen_notice:
            seen_notice.add(key)
            notices.append(key)

    def current_space() -> str:
        return SPACE_POLAR if polar else base_space

    def feed_rate() -> float:
        f = feed if feed and feed > 0 else None
        if f is None:
            return DEFAULT_FEED_MM_MIN
        if not per_rev:
            return f
        if css:
            x = pos.get("X")
            dia = max(abs(x) if x is not None else 50.0, 1.0)
            rpm = min(spindle * 1000.0 / (math.pi * dia), rpm_limit) if spindle else 1000.0
        else:
            rpm = spindle or 1000.0
        return max(f * max(rpm, 1.0), 1.0)

    def emit(line_no: int, kind: str, pts, c0=None, c1=None) -> None:
        length = _polyline_length(pts)
        if not polar and c0 is not None and c1 is not None and abs(c1 - c0) > 1e-9:
            # C 軸が回る移動: 回転ぶんの周長も長さに含める
            radius = max(math.hypot(pts[0][0], pts[0][1]), math.hypot(pts[-1][0], pts[-1][1]))
            length = math.hypot(length, radius * math.radians(abs(c1 - c0)))
        if length < 1e-9:
            return
        rate = RAPID_MM_MIN if kind == "rapid" else feed_rate()
        segments.append((current_space(), Segment(line_no, kind, list(pts), length / rate, c0, c1)))

    def move_linear(line_no: int, kind: str, target: dict) -> None:
        space = current_space()
        p0 = _point(space, pos)
        c0 = pos.get("C")
        pos.update(target)
        p1 = _point(space, pos)
        if p0 is not None and p1 is not None:
            emit(line_no, kind, [p0, p1], c0, pos.get("C"))

    for line_no, raw in rows:
        code = _strip_comments(raw).strip().upper()
        if code.startswith("/"):
            code = code[1:]
        code = re.sub(r"\s+", "", code)
        if not code or code.startswith("%"):
            continue
        if _MACRO_RE.search(code):
            note(line_no, "変数・マクロ文は描画しません")
            continue
        words = _parse_words(code)
        gcodes = [round(v, 1) for k, v, _ in words if k == "G"]
        values: dict[str, float] = {}
        for k, v, _ in words:
            if k not in ("G", "M", "N", "O"):
                values.setdefault(k, v)

        if 65 in gcodes or 66 in gcodes:
            note(line_no, "G65/G66 マクロ呼び出しは描画しません")
            continue
        if any(70 <= g <= 76 for g in gcodes):
            note(line_no, "旋削固定サイクルは未対応です")
            continue
        if 43.4 in gcodes:
            note(line_no, "G43.4（同時5軸）は未対応です")
        if 4 in gcodes or 10 in gcodes or 11 in gcodes:
            continue
        if 28 in gcodes or 30 in gcodes or 53 in gcodes:
            for axis, inc in (("X", "U"), ("Y", "V"), ("Z", "W"), ("C", "H")):
                if axis in values or inc in values:
                    pos[axis] = None
            continue
        if 361 in gcodes or 362 in gcodes:
            pos.update({"X": None, "Z": None, "C": None})
            continue

        for g in gcodes:
            if g in (0, 1, 2, 3):
                motion = int(g)
                cycle = None
            elif g == 90:
                absolute = True
            elif g == 91:
                absolute = False
            elif g in (17, 18, 19):
                plane = int(g)
            elif g == 112:
                polar = True
                pos["C"] = None
            elif g == 113:
                polar = False
                pos["C"] = None
            elif g == 96:
                css = True
            elif g == 97:
                css = False
            elif g == 98:
                per_rev = False
            elif g == 99:
                per_rev = True
            elif g == 80:
                cycle = None

        if 50 in gcodes:
            if "S" in values:
                rpm_limit = values["S"]
            if "X" in values or "Z" in values:
                note(line_no, "G50 の座標系設定は描画に反映しません")
                pos.update({"X": None, "Z": None})
            continue
        if "S" in values:
            spindle = values["S"]
        if "F" in values:
            feed = values["F"]

        # 目標位置（X は直径値のまま保持する）
        target: dict[str, float] = {}
        for axis, inc in (("X", "U"), ("Y", "V"), ("Z", "W")):
            if axis in values:
                if absolute:
                    target[axis] = values[axis]
                elif pos[axis] is not None:
                    target[axis] = pos[axis] + values[axis]
                else:
                    target[axis] = None
            elif inc in values:
                target[axis] = pos[axis] + values[inc] if pos[axis] is not None else None
        has_h_as_c = "H" in values and 43 not in gcodes and 43.4 not in gcodes and 49 not in gcodes
        if "C" in values or has_h_as_c:
            if "C" in values:
                c_val = values["C"] if absolute or pos["C"] is None else pos["C"] + values["C"]
                if not absolute and pos["C"] is None:
                    c_val = None
            else:
                c_val = pos["C"] + values["H"] if pos["C"] is not None else None
            target["C"] = c_val
            if not polar and c_val not in (None, 0.0):
                if not c_index:
                    note(line_no, "C軸割り出しをワーク基準で描画（C+ の回転方向は仮決め）")
                c_index = True

        new_cycle = next((int(g) for g in gcodes if 81 <= g <= 89), None)
        if new_cycle is not None:
            axis = "X" if new_cycle >= 87 else "Z"
            cycle = {"code": new_cycle, "axis": axis, "bottom": None, "r_inc": 0.0}

        if cycle is not None:
            if new_cycle is None and not target and "R" not in values:
                continue
            axis = cycle["axis"]
            if axis in target:
                cycle["bottom"] = target.pop(axis)
            if "R" in values:
                cycle["r_inc"] = values["R"] * (2.0 if axis == "X" else 1.0)
            initial = pos[axis]
            if initial is None or cycle["bottom"] is None:
                note(line_no, "穴あけサイクルの位置が不明なため描画しません")
                pos.update(target)
                continue
            move_linear(line_no, "rapid", target)
            r_point = initial + cycle["r_inc"]
            move_linear(line_no, "rapid", {axis: r_point})
            move_linear(line_no, "feed", {axis: cycle["bottom"]})
            move_linear(line_no, "rapid", {axis: initial})
            continue

        if not target:
            continue
        if motion in (0, 1):
            move_linear(line_no, "rapid" if motion == 0 else "feed", target)
            continue

        # 円弧
        space = current_space()
        p0 = _point(space, pos)
        c0 = pos.get("C")
        pos.update(target)
        p1 = _point(space, pos)
        if p0 is None or p1 is None:
            continue
        c1 = pos.get("C")
        ia, ib, iw, wa, wb = _arc_plane(space, plane)
        clockwise = motion == 2
        if "R" in values:
            center = _r_center(p0, p1, ia, ib, values["R"], clockwise)
            full = False
        else:
            center = (p0[ia] + values.get(wa, 0.0), p0[ib] + values.get(wb, 0.0))
            full = True
        if center is None:
            emit(line_no, "feed", [p0, p1], c0, c1)
            continue
        emit(line_no, "feed", arc_points(p0, p1, ia, ib, iw, center, clockwise, full), c0, c1)

    polar_count = sum(1 for space, _ in segments if space == SPACE_POLAR)
    shown = SPACE_POLAR if polar_count else base_space
    kept = [seg for space, seg in segments if space == shown]
    dropped = len(segments) - len(kept)
    if shown == SPACE_MILL and c_index:
        shown = SPACE_MILL_C
        for seg in kept:
            seg.pts = _rotate_by_c(seg)
    return SimResult(space=shown, segments=kept, notices=notices, dropped=dropped, c_index=c_index)


def _rotate_by_c(seg: Segment) -> list:
    """ミーリングの点 (X半径, Y, Z) を、C 角度ぶん戻してワーク基準 (X, Y, Z, 機械X半径) にする。
    移動中に C が変わる場合は、移動距離に比例して C を補間する。"""
    c0 = seg.c0 or 0.0
    c1 = seg.c1 if seg.c1 is not None else c0
    pts = seg.pts
    if abs(c1 - c0) > 1e-9 and len(pts) == 2:
        # 直線移動中に C が回るときは 2° ごとに分けて、回転の軌跡を描く
        n = min(720, max(2, math.ceil(abs(c1 - c0) / 2.0)))
        a, b = pts
        pts = [tuple(x + (y - x) * k / n for x, y in zip(a, b)) for k in range(n + 1)]
    total = _polyline_length(pts)
    out = []
    acc = 0.0
    for i, p in enumerate(pts):
        if i:
            acc += math.dist(pts[i - 1], p)
        f = acc / total if total > 0 else i / max(len(pts) - 1, 1)
        theta = -math.radians(C_SIGN * (c0 + (c1 - c0) * f))
        x, y, z = p[0], p[1], p[2]
        out.append((x * math.cos(theta) - y * math.sin(theta),
                    x * math.sin(theta) + y * math.cos(theta), z, x))
    return out


# ---------------------------------------------------------------------------
# 画面
# ---------------------------------------------------------------------------

BG = "#0B0F14"
PANEL = "#131A22"
TEXT = "#DCEAF4"
MUTED = "#8296A6"
ACCENT = "#22D3C5"
GRID = "#1E2833"
AXIS = "#33475A"
HILITE = "#FBBF24"
SLIDER = "#4A6680"
COLORS = {
    ("rapid", False): "#2C3846", ("rapid", True): "#7D90A2",
    ("feed", False): "#174A50", ("feed", True): "#22D3C5",
}


class PathView(tk.Canvas):
    """1 画面分の描画。ワールド座標 → 画面: sx = ox + s*wx, sy = oy - s*wy"""

    def __init__(self, master, title, ia, ib, label_a, label_b, on_pick):
        super().__init__(master, bg=BG, highlightthickness=1, highlightbackground=GRID, bd=0)
        self.title, self.ia, self.ib = title, ia, ib
        self.label_a, self.label_b = label_a, label_b
        self.on_pick = on_pick
        self.s, self.ox, self.oy = 1.0, 0.0, 0.0
        self.items: list[int] = []
        self._drag = None
        self._fitted_size = None
        self.bind("<MouseWheel>", self._wheel)
        self.bind("<Button-4>", lambda e: self._zoom(e.x, e.y, 1.2))
        self.bind("<Button-5>", lambda e: self._zoom(e.x, e.y, 1 / 1.2))
        self.bind("<ButtonPress-1>", self._press)
        self.bind("<B1-Motion>", self._motion)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Configure>", self._on_resize)

    def to_screen(self, p):
        return self.ox + self.s * p[self.ia], self.oy - self.s * p[self.ib]

    def build(self, segments, center_line=False):
        self.delete("all")
        self.items = []
        w, h = max(self.winfo_width(), 50), max(self.winfo_height(), 50)
        pts = [p for seg in segments for p in seg.pts]
        if pts:
            a = [p[self.ia] for p in pts]
            b = [p[self.ib] for p in pts]
            amin, amax, bmin, bmax = min(a), max(a), min(b), max(b)
        else:
            amin, amax, bmin, bmax = -10.0, 10.0, -10.0, 10.0
        span_a, span_b = max(amax - amin, 1e-3), max(bmax - bmin, 1e-3)
        self.s = 0.85 * min((w - 40) / span_a, (h - 50) / span_b)
        self.s = max(self.s, 1e-6)
        self.ox = w / 2 - self.s * (amin + amax) / 2
        self.oy = h / 2 + 10 + self.s * (bmin + bmax) / 2
        self._fitted_size = (w, h)

        big = 1e6
        x0, y0 = self.to_screen(self._world(0.0, 0.0))
        self.create_line(x0 - big, y0, x0 + big, y0, fill=AXIS, dash=(6, 4) if center_line else None, tags=("world", "axis"))
        self.create_line(x0, y0 - big, x0, y0 + big, fill=AXIS, tags=("world", "axis"))
        for seg in segments:
            if seg.kind == "feed" and self.is_point(seg):
                # この図では動きが点になる切削（正面から見た穴あけなど）は丸印で示す
                x, y = self.to_screen(seg.pts[0])
                color = COLORS[(seg.kind, True)]
                item = self.create_oval(x - 4, y - 4, x + 4, y + 4, outline=color, fill=color, width=1,
                                        tags=("world", "seg"))
                self.items.append(item)
                continue
            coords = []
            for p in seg.pts:
                coords.extend(self.to_screen(p))
            item = self.create_line(*coords, fill=COLORS[(seg.kind, True)], width=1.6 if seg.kind == "feed" else 1,
                                    dash=(3, 3) if seg.kind == "rapid" else None, tags=("world", "seg"))
            self.items.append(item)
        self._draw_overlay_text()

    def paint(self, item, color) -> None:
        if self.type(item) == "oval":
            self.itemconfigure(item, fill=color, outline=color)
        else:
            self.itemconfigure(item, fill=color)

    def is_point(self, seg) -> bool:
        a0, b0 = seg.pts[0][self.ia], seg.pts[0][self.ib]
        return all(abs(p[self.ia] - a0) < 1e-6 and abs(p[self.ib] - b0) < 1e-6 for p in seg.pts)

    def _world(self, a, b):
        p = [0.0, 0.0, 0.0, 0.0]
        p[self.ia], p[self.ib] = a, b
        return p

    def _draw_overlay_text(self):
        self.delete("hud")
        self.create_text(10, 8, text=self.title, anchor="nw", fill=TEXT, font=("Yu Gothic UI", 10, "bold"), tags="hud")
        self.create_text(10, 28, text=f"横: {self.label_a}   縦: {self.label_b}", anchor="nw", fill=MUTED,
                         font=("Yu Gothic UI", 9), tags="hud")

    def _apply(self, f, x, y):
        self.scale("world", x, y, f, f)
        self.ox = x + (self.ox - x) * f
        self.oy = y + (self.oy - y) * f
        self.s *= f

    def _zoom(self, x, y, f):
        self._apply(f, x, y)
        self.event_generate("<<ViewChanged>>")

    def _wheel(self, event):
        self._zoom(event.x, event.y, 1.2 if event.delta > 0 else 1 / 1.2)

    def _press(self, event):
        self._drag = (event.x, event.y, event.x, event.y)

    def _motion(self, event):
        if self._drag is None:
            return
        sx, sy, lx, ly = self._drag
        dx, dy = event.x - lx, event.y - ly
        self.move("world", dx, dy)
        self.ox += dx
        self.oy += dy
        self._drag = (sx, sy, event.x, event.y)
        self.event_generate("<<ViewChanged>>")

    def _release(self, event):
        if self._drag is None:
            return
        sx, sy, _, _ = self._drag
        self._drag = None
        if abs(event.x - sx) + abs(event.y - sy) > 3:
            return
        hits = [i for i in self.find_overlapping(event.x - 4, event.y - 4, event.x + 4, event.y + 4) if i in self.items]
        if hits:
            self.on_pick(self.items.index(hits[-1]))

    def _on_resize(self, event):
        if self._fitted_size and self._fitted_size != (event.width, event.height):
            self.event_generate("<<NeedFit>>")


class ToolpathPanel(tk.Frame):
    """メイン画面に埋め込む軌跡パネル。on_close があれば右上に「閉じる」ボタンを出す。"""
    PLAY_SECONDS = 20.0     # 1倍速でブロック全体を再生する時間
    TICK_MS = 30

    def __init__(self, master, on_jump, on_follow, on_close=None, button_factory=None):
        super().__init__(master, bg=PANEL)
        self.on_jump, self.on_follow, self.on_close = on_jump, on_follow, on_close
        self._button = button_factory or (lambda parent, text, cmd: tk.Button(parent, text=text, command=cmd))
        self._key = None
        self.result: SimResult | None = None
        self.views: list[PathView] = []
        self._cum: list[float] = []
        self._total = 0.0
        self.t = 0.0
        self._done = 0
        self._playing = False
        self._after = None
        self._follow_line = None
        self._hl_line = None
        self._fit_pending = None

        top = tk.Frame(self, bg=PANEL)
        top.pack(fill="x", padx=12, pady=(10, 4))
        self.title_var = tk.StringVar(value="エディタでN番号の中にカーソルを置いてください")
        tk.Label(top, textvariable=self.title_var, bg=PANEL, fg=TEXT, font=("Yu Gothic UI", 12, "bold")).pack(side="left")
        self.follow_var = tk.BooleanVar(value=True)
        tk.Checkbutton(top, text="再生中の行をエディタで追う", variable=self.follow_var, bg=PANEL, fg=MUTED,
                       selectcolor=BG, activebackground=PANEL, activeforeground=TEXT,
                       font=("Yu Gothic UI", 9)).pack(side="right")
        self._button(top, "全体表示", self.fit).pack(side="right", padx=8)
        if on_close is not None:
            self._button(top, "✕ 閉じる", on_close).pack(side="right", padx=(0, 8))

        self.notice_var = tk.StringVar()
        notice = tk.Label(self, textvariable=self.notice_var, bg=PANEL, fg=HILITE, anchor="w", justify="left",
                          font=("Yu Gothic UI", 9), wraplength=600)
        notice.pack(fill="x", padx=14)
        notice.bind("<Configure>", lambda e: notice.configure(wraplength=max(200, e.width - 8)))

        self.view_frame = tk.Frame(self, bg=PANEL)
        self.view_frame.pack(fill="both", expand=True, padx=12, pady=6)

        ctrl = tk.Frame(self, bg=PANEL)
        ctrl.pack(fill="x", padx=12, pady=(0, 10))
        self._button(ctrl, "⏮ 最初", self.rewind).pack(side="left")
        self._button(ctrl, "◀ 1行", lambda: self.step_line(-1)).pack(side="left", padx=(6, 0))
        self.play_button = self._button(ctrl, "▶ 再生", self.toggle_play)
        self.play_button.pack(side="left", padx=6)
        self._button(ctrl, "1行 ▶", lambda: self.step_line(1)).pack(side="left")
        tk.Label(ctrl, text="速度", bg=PANEL, fg=MUTED, font=("Yu Gothic UI", 9)).pack(side="left", padx=(14, 4))
        self.speed_var = tk.DoubleVar(value=33.0)
        self.speed_label = tk.Label(ctrl, text="", width=6, bg=PANEL, fg=TEXT, font=("Consolas", 10))
        tk.Scale(ctrl, from_=0, to=100, orient="horizontal", variable=self.speed_var, showvalue=False, length=120,
                 bg=SLIDER, activebackground=ACCENT, troughcolor=BG, highlightthickness=0, bd=0,
                 sliderrelief="flat", sliderlength=16).pack(side="left")
        self.speed_label.pack(side="left")
        self.progress_var = tk.DoubleVar(value=1000.0)
        self.progress = tk.Scale(ctrl, from_=0, to=1000, orient="horizontal", variable=self.progress_var,
                                 showvalue=False, bg=SLIDER, activebackground=ACCENT, troughcolor=BG,
                                 highlightthickness=0, bd=0, sliderrelief="flat", sliderlength=16,
                                 command=self._on_progress_drag)
        self.progress.pack(side="left", fill="x", expand=True, padx=(14, 0))
        self.status_var = tk.StringVar()
        tk.Label(self, textvariable=self.status_var, bg=PANEL, fg=MUTED, anchor="w",
                 font=("Consolas", 9)).pack(fill="x", padx=14, pady=(0, 8))
        self.speed_var.trace_add("write", lambda *_: self._update_speed_label())
        self._update_speed_label()

    # ---- 外部から ----
    def set_block(self, label: str, rows, mode: str) -> None:
        key = (label, tuple(rows), mode)
        if key == self._key:
            return
        self._key = key
        self.stop()
        self.result = simulate_block(list(rows), mode)
        res = self.result
        names = {SPACE_TURN: "旋削", SPACE_MILL: "ミーリング", SPACE_POLAR: "極座標補間（G112）",
                 SPACE_MILL_C: "ミーリング（C割り出し）"}
        self.title_var.set(f"{label}   {names[res.space]}   移動 {len(res.segments)} 本")
        msgs = [f"{ln}行: {msg}" for ln, msg in res.notices]
        if res.dropped:
            msgs.insert(0, f"G112 の外の移動 {res.dropped} 本は省略")
        if len(msgs) > 4:
            msgs = msgs[:4] + [f"ほか {len(msgs) - 4} 件"]
        self.notice_var.set("未対応・省略: " + " / ".join(msgs) if msgs else "")

        for v in self.views:
            v.destroy()
        self.views = []
        defs = SPACE_VIEWS[res.space]
        for col, (title, ia, ib, la, lb) in enumerate(defs):
            view = PathView(self.view_frame, title, ia, ib, la, lb, self._pick)
            view.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 6, 0))
            self.views.append(view)
            view.bind("<<NeedFit>>", lambda e: self._schedule_fit())
            view.bind("<<ViewChanged>>", lambda e: self._redraw_dynamic())
            self.view_frame.columnconfigure(col, weight=1, uniform="v")
        for col in range(len(defs), 3):
            self.view_frame.columnconfigure(col, weight=0, uniform="")
        self.view_frame.rowconfigure(0, weight=1)

        self._cum = []
        acc = 0.0
        for seg in res.segments:
            acc += max(seg.weight, 1e-9)
            self._cum.append(acc)
        self._total = acc
        self.t = self._total
        self._done = len(res.segments)
        self._hl_line = None
        self.update_idletasks()
        self.fit()

    def highlight_line(self, line_no: int | None) -> None:
        if line_no == self._hl_line:
            return
        self._hl_line = line_no
        self._draw_highlight()

    # ---- 表示 ----
    def fit(self) -> None:
        if self.result is None:
            return
        for v in self.views:
            v.build(self.result.segments, center_line=self.result.space == SPACE_TURN)
        self._recolor_all()
        self._redraw_dynamic()

    def _schedule_fit(self):
        if self._fit_pending is not None:
            self.after_cancel(self._fit_pending)
        self._fit_pending = self.after(120, self._do_fit)

    def _do_fit(self):
        self._fit_pending = None
        self.fit()

    def _recolor_all(self):
        if self.result is None:
            return
        for v in self.views:
            for i, item in enumerate(v.items):
                seg = self.result.segments[i]
                v.paint(item, COLORS[(seg.kind, i < self._done)])

    def _set_done(self, new_done):
        if self.result is None or new_done == self._done:
            return
        lo, hi = sorted((self._done, new_done))
        for v in self.views:
            for i in range(lo, hi):
                seg = self.result.segments[i]
                v.paint(v.items[i], COLORS[(seg.kind, i < new_done)])
        self._done = new_done

    def _draw_highlight(self):
        for v in self.views:
            v.delete("hl")
        if self.result is None or self._hl_line is None:
            return
        for v in self.views:
            for seg in self.result.segments:
                if seg.line_no == self._hl_line:
                    if v.is_point(seg):
                        x, y = v.to_screen(seg.pts[0])
                        v.create_oval(x - 6, y - 6, x + 6, y + 6, outline=HILITE, width=3, tags=("world", "hl"))
                        continue
                    coords = []
                    for p in seg.pts:
                        coords.extend(v.to_screen(p))
                    v.create_line(*coords, fill=HILITE, width=3, tags=("world", "hl"))

    def _current(self):
        """現在時刻の (セグメント番号, 工具位置)"""
        res = self.result
        if res is None or not res.segments:
            return None, None
        if self.t >= self._total:
            return len(res.segments) - 1, res.segments[-1].pts[-1]
        i = bisect_right(self._cum, self.t)
        i = min(i, len(res.segments) - 1)
        start = self._cum[i - 1] if i > 0 else 0.0
        seg = res.segments[i]
        frac = (self.t - start) / max(self._cum[i] - start, 1e-12)
        return i, _point_along(seg.pts, frac)

    def _redraw_dynamic(self):
        for v in self.views:
            v.delete("dyn")
        idx, tool = self._current()
        self._draw_highlight()
        if idx is None:
            self.status_var.set("描画できる移動がありません")
            return
        res = self.result
        done = idx if self.t < self._total else len(res.segments)
        self._set_done(done)
        seg = res.segments[idx]
        if self.t < self._total:
            frac_pts = _partial(seg.pts, tool)
            for v in self.views:
                coords = []
                for p in frac_pts:
                    coords.extend(v.to_screen(p))
                if len(coords) >= 4:
                    v.create_line(*coords, fill=COLORS[(seg.kind, True)], width=2, tags=("world", "dyn"))
        for v in self.views:
            x, y = v.to_screen(tool)
            v.create_oval(x - 5, y - 5, x + 5, y + 5, outline="#FFFFFF", width=2, fill=HILITE, tags=("world", "dyn"))
        self.status_var.set(
            f"{seg.line_no}行  {'早送り' if seg.kind == 'rapid' else '切削'}  "
            + "  ".join(f"{name}={tool[i]:.3f}" for name, i in self._axis_names()))
        if self._playing and self.follow_var.get() and seg.line_no != self._follow_line:
            self._follow_line = seg.line_no
            self.on_follow(seg.line_no)

    def _axis_names(self):
        if self.result.space == SPACE_TURN:
            return [("Z", 0), ("X(半径)", 1)]
        if self.result.space == SPACE_POLAR:
            return [("X(半径)", 0), ("C", 1), ("Z", 2)]
        if self.result.space == SPACE_MILL_C:
            return [("ワークX", 0), ("ワークY", 1), ("Z", 2)]
        return [("X(半径)", 0), ("Y", 1), ("Z", 2)]

    # ---- 再生 ----
    def _speed(self) -> float:
        return 0.5 * (40.0 ** (self.speed_var.get() / 100.0))   # 0.5〜20倍（対数）

    def _update_speed_label(self):
        self.speed_label.configure(text=f"x{self._speed():.1f}")

    def toggle_play(self):
        if self._playing:
            self.stop()
            return
        if self.result is None or not self.result.segments:
            return
        if self.t >= self._total:
            self.t = 0.0
        self._playing = True
        self.play_button.configure(text="■ 停止")
        self._tick()

    def stop(self):
        self._playing = False
        if self._after is not None:
            self.after_cancel(self._after)
            self._after = None
        try:
            self.play_button.configure(text="▶ 再生")
        except tk.TclError:
            pass

    def _tick(self):
        self._after = None
        if not self._playing:
            return
        self.t = min(self._total, self.t + self._total * self._speed() * self.TICK_MS / 1000.0 / self.PLAY_SECONDS)
        self._sync_progress()
        self._redraw_dynamic()
        if self.t >= self._total:
            self.stop()
            return
        self._after = self.after(self.TICK_MS, self._tick)

    def rewind(self):
        self.stop()
        self.t = 0.0
        self._sync_progress()
        self._redraw_dynamic()

    def step_line(self, direction: int):
        """前後の行の終わりまで進める / 戻す"""
        self.stop()
        res = self.result
        if res is None or not res.segments:
            return
        idx, _ = self._current()
        line = res.segments[idx].line_no
        ends = []
        for i, seg in enumerate(res.segments):
            if i + 1 == len(res.segments) or res.segments[i + 1].line_no != seg.line_no:
                ends.append((self._cum[i], seg.line_no))
        if direction > 0:
            nxt = next((t for t, ln in ends if t > self.t + 1e-12), self._total)
        else:
            prev = [t for t, ln in ends if t < self.t - 1e-12 and ln != line]
            nxt = prev[-1] if prev else 0.0
        self.t = nxt
        self._sync_progress()
        self._redraw_dynamic()
        idx, _ = self._current()
        if self.follow_var.get():
            self._follow_line = res.segments[idx].line_no
            self.on_follow(self._follow_line)

    def _sync_progress(self):
        self.progress_var.set(1000.0 * self.t / self._total if self._total else 1000.0)

    def _on_progress_drag(self, value):
        if self._playing or self._total <= 0:
            return
        t = float(value) / 1000.0 * self._total
        if abs(t - self.t) > 1e-12:
            self.t = t
            self._redraw_dynamic()

    def _pick(self, index: int):
        if self.result is not None and 0 <= index < len(self.result.segments):
            self.on_jump(self.result.segments[index].line_no)

    def destroy(self):
        self.stop()
        super().destroy()


def _point_along(pts, frac):
    total = _polyline_length(pts)
    if total <= 0:
        return pts[-1]
    target = total * max(0.0, min(1.0, frac))
    for i in range(len(pts) - 1):
        d = math.dist(pts[i], pts[i + 1])
        if target <= d or i == len(pts) - 2:
            f = target / d if d > 0 else 1.0
            return tuple(a + (b - a) * min(f, 1.0) for a, b in zip(pts[i], pts[i + 1]))
        target -= d
    return pts[-1]


def _partial(pts, tool):
    """セグメントの始点から工具位置までの折れ線"""
    out = [pts[0]]
    best = 0
    best_d = math.inf
    for i in range(len(pts) - 1):
        a, b = pts[i], pts[i + 1]
        ab = [y - x for x, y in zip(a, b)]
        denom = sum(c * c for c in ab) or 1e-12
        f = max(0.0, min(1.0, sum((t - x) * c for t, x, c in zip(tool, a, ab)) / denom))
        proj = [x + c * f for x, c in zip(a, ab)]
        d = math.dist(proj, tool)
        if d < best_d - 1e-12:
            best, best_d = i, d
    out.extend(pts[1:best + 1])
    out.append(tuple(tool))
    return out
