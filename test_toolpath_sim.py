"""簡易軌跡シミュレーターのテスト：解析（旋削・極座標補間・穴あけ・円弧・未対応）と画面の操作。"""
import math
import tkinter as tk
import toolpath_sim as t


def rows(text):
    return list(enumerate(text.strip("\n").split("\n"), start=1))


def end(seg):
    return tuple(round(v, 3) for v in seg.pts[-1])


POLAR = """\
N1
G361B-90.D0
G43H18
G97S7579M13
G98G17
G0C30.3
X64.16Y0
Z2.
G112
G1X64.16C0F1000.
Z1.
Z-1.F606.305
X61.885C0.006
G3X61.016C-0.242I-0.003J-0.5
G1X31.65C-25.369
G2X-2.667C-29.87I-15.825J25.369
G1X-50.738C-15.825
Z2.F1000.
G113
G0X62.569Y0Z22.25M9
G28W0
M1
"""

# ---- 極座標補間: X は直径値、I は半径値、C は mm ----
r = t.simulate_block(rows(POLAR), "MILL")
assert r.space == t.SPACE_POLAR
assert r.c_index and any("C軸割り出し" in msg for _, msg in r.notices)
assert r.dropped == 1                                  # G113 の後の G0 は省略
assert end(r.segments[0]) == (32.08, 0.0, 1.0)        # G112 直後の C 未確定区間は描かない（Z1. から）
arc = next(s for s in r.segments if s.line_no == 16)
assert {round(math.hypot(p[0], p[1]), 3) for p in arc.pts} == {29.9}   # 中心 (0,0)・半径 29.9
assert arc.pts[0][:2] == (15.825, -25.369) and math.dist(arc.pts[-1][:2], (-1.3335, -29.87)) < 1e-9
small = next(s for s in r.segments if s.line_no == 14)
assert {round(math.dist(p[:2], (30.9395, -0.494)), 2) for p in small.pts} == {0.5}

# ---- 旋削: Z–X（X は半径に直す）、中心越え、G28 で線が切れる ----
TURN = """\
N1
G97S279M3
G99G18
G0Y0
X274.Z100.
Z15.
G50S1500
G96S240
G1Z1.775F1.
G1X-3.6F0.3
G0Z4.414
G28U0
G0Z10.
X100.
G1W-5.
U10.
M1
"""
r = t.simulate_block(rows(TURN), "TURN")
assert r.space == t.SPACE_TURN and r.notices == []
assert r.segments[0].pts[0] == (100.0, 137.0, 0.0) and end(r.segments[0]) == (15.0, 137.0, 0.0)
assert end(r.segments[2]) == (1.775, -1.8, 0.0)
lines = [s.line_no for s in r.segments]
assert 13 not in lines                                 # G28U0 の後は X 不明なので Z10. は描かない
assert end(r.segments[-2]) == (5.0, 50.0, 0.0)         # W-5. はインクリメンタル
assert end(r.segments[-1]) == (5.0, 55.0, 0.0)         # U10. は直径で 10 → 半径 5
assert all(s.weight > 0 for s in r.segments)

# ---- R 指定の円弧（旋削 G3、Z 右・X 上で反時計回り） ----
r = t.simulate_block(rows("N1\nG18\nG0X0Z0\nG3X20.Z-10.R10.F0.1\n"), "TURN")
arc = r.segments[-1]
assert {round(math.dist(p[:2], (-10.0, 0.0)), 6) for p in arc.pts} == {10.0}
assert end(arc) == (-10.0, 10.0, 0.0)
r = t.simulate_block(rows("N1\nG17\nG0X0Y0Z0\nG2X20.Y0R-10.\n"), "MILL")   # R<0 は大円弧（ここは半円）
assert len(r.segments[-1].pts) > 10

# ---- 穴あけ: R は初期点からのインクリメンタル、戻りは初期点 ----
DRILL = """\
N2
G97S1061M13
G98G17
G0C0
X0Y0
Z10.
Z2.1
G83Z-46.2R0F99.M68
X20.
G80
G0Z10.
G65P7007A3.B2.C0D500.E2.
M1
"""
r = t.simulate_block(rows(DRILL), "MILL")
assert not r.c_index                                    # C0 は割り出しとして扱わない
assert [msg for _, msg in r.notices] == ["G65/G66 マクロ呼び出しは描画しません"]
cyc = [s for s in r.segments if s.line_no == 8]
assert [s.kind for s in cyc] == ["feed", "rapid"]       # R0 なので R 点までの早送りは長さ 0
assert end(cyc[0]) == (0.0, 0.0, -46.2) and end(cyc[1]) == (0.0, 0.0, 2.1)
second = [s for s in r.segments if s.line_no == 9]
assert end(second[0]) == (10.0, 0.0, 2.1) and end(second[1]) == (10.0, 0.0, -46.2)
r = t.simulate_block(rows("N1\nG98G17\nG0X0Y0Z10.\nG83Z-5.R-5.F100.\n"), "MILL")
assert end(r.segments[0]) == (0.0, 0.0, 5.0)            # R-5. → 初期点 10 から 5 へ

# ---- 未対応 ----
r = t.simulate_block(rows("N1\nG0X100.Z5.\nG71U2.R1.\nG71P10Q20U0.4W0.1F0.3\n#100=5\nG0X#100\nG43.4H5\n"), "TURN")
msgs = [msg for _, msg in r.notices]
assert "旋削固定サイクルは未対応です" in msgs and "変数・マクロ文は描画しません" in msgs
assert "G43.4（同時5軸）は未対応です" in msgs

# ---- 画面 ----
# ---- C 軸割り出し: ワーク基準の端面図に回して描く ----
CIDX = """\
N3
G98G17
G0C0
X100.Y0Z5.
G83Z-10.R0F100.
G80
G0C90.
G83Z-10.R0F100.
G80
G0C180.
G87X60.R0F100.
G80
M1
"""
r = t.simulate_block(rows(CIDX), "MILL")
assert r.space == t.SPACE_MILL_C and r.c_index
assert [msg for _, msg in r.notices] == ["C軸割り出しをワーク基準で描画（C+ の回転方向は仮決め）"]
hole0 = [s for s in r.segments if s.line_no == 5]
hole90 = [s for s in r.segments if s.line_no == 8]
side180 = [s for s in r.segments if s.line_no == 11]
def near(p, q):
    return all(abs(a - b) < 1e-9 for a, b in zip(p, q))
assert near(end(hole0[0]), (50.0, 0.0, -10.0, 50.0))          # C0: そのまま
assert near(hole90[0].pts[-1], (0.0, -50.0, -10.0, 50.0))      # C90: ワーク基準では -90° 側（C_SIGN=+1）
assert near(side180[0].pts[-1], (-30.0, 0.0, 5.0, 30.0))       # C180 の側面穴あけ（X 方向）
assert near(side180[0].pts[0], (-50.0, 0.0, 5.0, 50.0))
# C を動かしながらの切削は円弧状になる（補間）
r = t.simulate_block(rows("N1\nG98G17\nG0X100.Y0Z0C0\nG1C90.F100.\n"), "MILL")
sweep = r.segments[-1]
assert {round(math.hypot(p[0], p[1]), 6) for p in sweep.pts} == {50.0}
assert near(sweep.pts[-1][:2], (0.0, -50.0))
# C0 だけなら従来どおりのミーリング表示
assert t.simulate_block(rows(DRILL), "MILL").space == t.SPACE_MILL

root = tk.Tk()
root.withdraw()
jumps, follows = [], []
closed = []
top = tk.Toplevel(root)
top.geometry("900x600")
win = t.ToolpathPanel(top, on_jump=jumps.append, on_follow=follows.append, on_close=lambda: closed.append(1))
win.pack(fill="both", expand=True)
win.set_block("N1", rows(POLAR), "MILL")
root.update()
assert len(win.views) == 2 and all(len(v.items) == len(win.result.segments) for v in win.views)
first_items = win.views[0].items
win.set_block("N1", rows(POLAR), "MILL")                 # 同じ内容なら作り直さない
assert win.views[0].items is first_items

win.highlight_line(16)
assert win.views[0].find_withtag("hl")

win.rewind()
assert win.t == 0 and win._done == 0
win.speed_var.set(100)
assert win.speed_label.cget("text") == "x20.0"
win.toggle_play()
for _ in range(5):
    root.update()
    root.after(40)
win._tick()
assert 0 < win.t < win._total and follows, (win.t, follows)
win.stop()
win.step_line(1)
line_after_step = win.result.segments[win._current()[0]].line_no
win.step_line(-1)
assert win.result.segments[win._current()[0]].line_no <= line_after_step
win._on_progress_drag(1000)
assert win.t == win._total and win._done == len(win.result.segments)

v = win.views[0]
s_before = v.s
v._zoom(100, 100, 2.0)
assert abs(v.s - 2 * s_before) < 1e-9
win.fit()
assert abs(win.views[0].s - s_before) < 1e-6

v = win.views[0]
pick = win.result.segments[5]
mx, my = v.to_screen(pick.pts[len(pick.pts) // 2])     # 円弧上の点を押す
v._press(type("E", (), {"x": int(mx), "y": int(my)}))
v._release(type("E", (), {"x": int(mx), "y": int(my)}))
assert jumps and jumps[-1] == win.result.segments[5].line_no, jumps

win.set_block("N1", rows(TURN), "TURN")
root.update()
assert len(win.views) == 1 and win.result.space == t.SPACE_TURN
win.set_block("N3", rows(CIDX), "MILL")
root.update()
assert len(win.views) == 2 and win.views[1].ia == 3      # 2画面目は機械座標の X–Z
win.highlight_line(8)
win.fit()
win.destroy()
top.destroy()
root.destroy()
print("PASS: C-index (face/side, C sweep), polar (diameter X, radius I), turning U/W/G28, R arcs, drilling R/return, notices, window play/step/drag/zoom/fit/pick")
