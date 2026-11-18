"""軌跡モードの画面切替：4ペイン → NCプログラム30%｜軌跡70% → 閉じると元の並びと幅に戻る。"""
import tkinter as tk
import nc_program_check_tool_v40 as m

PROGRAM = """\
N1
G97S279M3
G99G18
G0X274.Z100.
G1Z1.F0.2
X-3.6
M1

N2
G98G17
G0C0
X100.Y0Z5.
G83Z-10.R0F100.
G80
G0C90.
G83Z-10.R0F100.
G80
M1
"""

import sys
m.NcCheckApp._install_mascot = lambda self: None   # マスコットはこのテストの対象外
root = tk.Tk()
if sys.platform != "win32":                          # Linux の Tk には "zoomed" が無い
    _state = root.state
    root.state = lambda new=None: None if new == "zoomed" else _state(new)
root.geometry("1400x800+0+0")
errors = []
root.report_callback_exception = lambda exc, val, tb: errors.append(val)
app = m.NcCheckApp(root)
root.update()
app.input_text.insert("1.0", PROGRAM)
app.input_text.mark_set(tk.INSERT, "3.0")
root.update()

body = app._main_panes
before = [str(p) for p in body.panes()]
before_sashes = [body.sash_coord(i)[0] for i in range(len(before) - 1)]
assert len(before) == 4

# 入る
app._toggle_toolpath_mode()
root.update()
panes = [str(p) for p in body.panes()]
assert panes == [str(app._input_panel), str(app._toolpath_panel)], panes
ratio = body.sash_coord(0)[0] / body.winfo_width()
assert 0.27 < ratio < 0.33, ratio
assert app._toolpath_panel.result is not None and app._toolpath_panel._key[0] == "N1"

# カーソルを N2 に動かすと自動で描き直す（C 割り出しブロック）
app.jump_to_input_line(13)
root.after(300, root.quit)
root.mainloop()
assert app._toolpath_panel._key[0] == "N2"
assert app._toolpath_panel.result.space == "mill_c"

# 線をクリックしたらその行へ（jump_to_input_line）
seg = app._toolpath_panel.result.segments[0]
app._toolpath_panel._pick(0)
assert int(app.input_text.index(tk.INSERT).split(".")[0]) == seg.line_no

# 閉じる（パネルの「閉じる」と同じ処理）
app._toolpath_panel.on_close()
root.update()
assert [str(p) for p in body.panes()] == before
after_sashes = [body.sash_coord(i)[0] for i in range(len(before) - 1)]
assert all(abs(a - b) <= 2 for a, b in zip(before_sashes, after_sashes)), (before_sashes, after_sashes)

# 閉じている間はカーソル移動で描き直さない
key = app._toolpath_panel._key
app.jump_to_input_line(3)
root.after(300, root.quit)
root.mainloop()
assert app._toolpath_panel._key == key

# もう一度入る → ボタンでもう一度押すと閉じる
app._toggle_toolpath_mode()
root.update()
assert len(body.panes()) == 2 and app._toolpath_panel._key[0] == "N1"
app._toggle_toolpath_mode()
root.update()
assert [str(p) for p in body.panes()] == before

assert not errors, errors
root.destroy()
print("PASS: enter 30/70, auto redraw on cursor, pick jumps, close restores panes and widths, toggle")
