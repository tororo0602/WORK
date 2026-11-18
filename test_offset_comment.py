"""補正コメント (H30) 自動挿入のテスト：挿入位置・書き換え・再実行・Undo・工具名の読み取り。"""
from types import SimpleNamespace
import tkinter as tk
from unittest.mock import patch
import nc_program_check_tool_v40 as m
import replacement_preview

SAMPLE = """\
N1
(OD H39 S50-DCLNN R0.8)
G54
T2001
G361B-45.D1.
T2030
G43H39
M1

N2
(U DRILL D30)
G54
T2030
G361B-90.D0
G28H0
G43H30
G65P7007A3.B2.C0D500.E2.
M1

N3
G49
T2005
G361B0.D0
G43H05
G0X0.
G43H39
M1

N4(FACE MILL)
T2004
G361B0.D0
G43H04
M1

N5
(NO TOOL CHANGE)
G43H07
M1

N6
(NO H)
T2006
G361B0.D0
M1
"""

lines = SAMPLE.split("\n")
plans = {p["n_label"]: p for p in m.plan_offset_comments(lines)}

# 対象: G361 があり G43 H があるブロックだけ（N5: G361 なし、N6: H なし は対象外）
assert set(plans) == {"N1", "N2", "N3", "N4"}, plans

# 工具名コメントの下（＝工具名コメント行の後ろに挿入）
assert plans["N1"]["action"] == "insert" and plans["N1"]["line_no"] == 2 and plans["N1"]["comment"] == "(H39)"
# G28H0 と G65 の D は拾わない
assert plans["N2"]["comment"] == "(H30)" and plans["N2"]["line_no"] == 11
# 工具名コメントが無ければ N 行の下、H は複数並べる
assert plans["N3"]["line_no"] == 20 and plans["N3"]["comment"] == "(H05 H39)"
# 工具名が N 行にあるときは N 行の下
assert plans["N4"]["line_no"] == 29 and plans["N4"]["comment"] == "(H04)"

# ---- エディタへの適用（Undo 1回で戻る） ----
root = tk.Tk()
root.withdraw()
editor = tk.Text(root, undo=True)
editor.insert("1.0", SAMPLE)
editor.edit_reset()
calls = []
app = SimpleNamespace(root=root, input_text=editor, status_var=tk.StringVar(),
                      _update_linenumbers=lambda: calls.append("ln"),
                      _highlight_all=lambda: calls.append("hl"),
                      _auto_rerun_check=lambda: calls.append("check"))
captured = {}


class FakePreview:
    def __init__(self, parent, title, rows, indices, on_confirm, grouped=False):
        captured.update(title=title, rows=rows, indices=indices, confirm=on_confirm, grouped=grouped)


with patch.object(replacement_preview, "ReplacementPreview", FakePreview):
    m.NcCheckApp._open_offset_comment_preview(app)
assert captured["grouped"] and len(captured["rows"]) == 4 and captured["title"] == "補正コメント（4箇所）"

# N3 だけ外して適用
order = [p["n_label"] for p in m.plan_offset_comments(lines)]
selected = {i for i, label in enumerate(order) if label != "N3"}
captured["confirm"](selected, False)
after = editor.get("1.0", "end-1c")
assert "(OD H39 S50-DCLNN R0.8)\n(H39)\nG54" in after
assert "(U DRILL D30)\n(H30)\nG54" in after
assert "N4(FACE MILL)\n(H04)\nT2004" in after
assert "N3\nG49" in after  # 外した分は入らない
assert calls == ["ln", "hl", "check"]

# 工具名の読み取りは (H..) 行を飛ばす
blocks = {b.n_label: b for b in m.split_n_blocks(after.split("\n"))}
assert m.extract_tool_name(blocks["N1"].rows)[0] == "(OD H39 S50-DCLNN R0.8)"
assert m.extract_tool_name(blocks["N2"].rows)[0] == "(U DRILL D30)"

# 2回目: 書いたブロックは変更なし、N3 だけ残る
again = {p["n_label"] for p in m.plan_offset_comments(after.split("\n"))}
assert again == {"N3"}, again

# Undo 1回で元通り
editor.edit_undo()
assert editor.get("1.0", "end-1c") == SAMPLE

# 既存の (H..) が古ければ書き換え
changed = SAMPLE.replace("(U DRILL D30)\n", "(U DRILL D30)\n(H31)\n")
plan = {p["n_label"]: p for p in m.plan_offset_comments(changed.split("\n"))}["N2"]
assert plan["action"] == "replace" and plan["old"] == "(H31)" and plan["new"] == "(H30)", plan

# 編集後に確定しようとしたら反映しない
editor.delete("1.0", "end")
editor.insert("1.0", SAMPLE)
with patch.object(replacement_preview, "ReplacementPreview", FakePreview):
    m.NcCheckApp._open_offset_comment_preview(app)
editor.insert("end", "\nM30")
with patch.object(m.messagebox, "showwarning") as warn:
    captured["confirm"]({0}, False)
assert warn.called and "(H39)" not in editor.get("1.0", "end-1c")

root.destroy()
print("PASS: targets, position under tool name, multi H, N-line tool name, apply/partial, undo, rerun, replace stale, stale snapshot")
