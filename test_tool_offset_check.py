"""T番号とH補正の一致チェック（tool_offset_mismatch）のテスト。"""
import nc_program_check_tool_v39 as m

SAMPLE = """\
N1
(OD H39 S50-DCLNN R0.8)
G54
G49
T2001
G361B-45.D1.
T2030
M301
G43H39
G97S279M3
G0Y0
X274.Z100.
M1

N2
(U DRILL D30 & C6-TSL40-85)
G54
G49
T2030
G361B-90.D0
T2003
M302
G28H0
G43H30
G97S1061M13
G83Z-46.2R0F99.M68
G80
G65P7007A3.B2.C0D500.E2.
M1
"""


def kinds(text):
    return [f for f in m.detect_tool_offset_mismatch_findings(text.splitlines())]


def block(n, t, h_lines, g361="G361B0.D0"):
    body = [n, "G49", t, g361, "T9999"] + h_lines + ["M1", ""]
    return "\n".join(body)


# 実プログラム: N1 は H39（パーマネント）でスルー、N2 は T2030/H30 で一致、G28H0・G65 の H/D は無視
assert kinds(SAMPLE) == [], kinds(SAMPLE)

# 下2桁と不一致 → 警告。行は G43 行、先読み T9999 は比較に使わない
f = kinds(block("N5", "T2005", ["G43H12"]))
assert len(f) == 1 and f[0].text.startswith("N5 T2005↔H12") and f[0].line_no == 6, f
assert f[0].reason == "T番号の下2桁はH05"

# H38 は対象、H39 以上はスルー
assert len(kinds(block("N5", "T2005", ["G43H38"]))) == 1
assert kinds(block("N5", "T2005", ["G43H39"])) == []
assert kinds(block("N5", "T2005", ["G43H50"])) == []

# H05 / H5 どちらの書き方でも一致扱い
assert kinds(block("N5", "T2005", ["G43H05"])) == []
assert kinds(block("N5", "T2005", ["G43 H5"])) == []

# G361 があるのに G43 H なし → G361 行に警告
f = kinds(block("N7", "T2007", ["G0X100."]))
assert len(f) == 1 and "G43のH補正なし" in f[0].text and f[0].line_no == 4, f

# G43 以外の H（G28H0 など）だけでは「H補正あり」にならない
assert "G43のH補正なし" in kinds(block("N7", "T2007", ["G28H0"]))[0].text

# 同一ブロックに H が複数
f = kinds(block("N8", "T2008", ["G43H08", "G43H09"]))
msgs = [x.text for x in f]
assert any("H08/H09" in x and "複数" in x for x in msgs), msgs
assert any("H09" in x and "不一致" in x for x in msgs), msgs

# G361 の無いブロックは対象外
assert kinds("N9\nT2001\nG43H12\nM1\n") == []

# コメント内の H は無視
assert kinds(block("N5", "T2005", ["G43H05(H12)"])) == []

# G43.4（TCP）の H も対象
assert len(kinds(block("N5", "T2005", ["G43.4H12"]))) == 1

# レポート・目次への組み込み
report = m.build_report(block("N5", "T2005", ["G43H12"]), m.ThresholdSettings())
assert report.summary.tool_offset_mismatch_count == 1
assert any(line.text == "T/H補正不一致: 1件" for line in report.lines)
assert any(g["kind"] == "tool_offset_mismatch" and g["sev"] == "wrn" for g in m.TOC_GROUPS)

print("PASS: sample program, mismatch, permanent H39+, missing G43 H, multiple H, comments, G43.4, report")
