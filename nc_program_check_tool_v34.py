#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import json
import os
import re
import sys
import tkinter as tk
from dataclasses import dataclass, field
from datetime import datetime
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font as XlFont, PatternFill, Alignment as XlAlign, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

APP_TITLE = "NCプログラム確認ツール v34"
WINDOW_SIZE = "1420x860"
READ_ENCODINGS = ["utf-8-sig", "utf-8", "cp932", "shift_jis", "latin-1"]
MODE_TURNING = "ターニング"
MODE_MILLING = "ミーリング"

BG_APP = "#0B0F14"
BG_PANEL = "#131A22"
BG_HEADER = "#1B4B73"
FG_HEADER = "#FFFFFF"
FG_SUB = "#A8CCEA"
ACCENT = "#22D3C5"
ACCENT_DARK = "#1BA89D"
TEXT_MAIN = "#FFFFFF"
TEXT_MUTED = "#8296A6"
INPUT_BG = "#080C11"
RESULT_BG = "#080C11"
BORDER = "#1E2833"
ERROR_LINE_BG = "#3A1620"
JUMP_LINE_BG = "#0C2628"

# 警告系の配色（双方向入替の警告枠など、注意喚起UI用）
WARNING_BG = "#241A08"        # 濃い琥珀系の背景
WARNING_BORDER = "#FBBF24"    # 警告アンバー（枠線・タイトル）
WARNING_TITLE = "#FDE68A"     # タイトル文字（少し明るめ）
WARNING_TEXT = "#E6D4A8"      # 本文（読みやすい薄い色）

# v35 判定セマンティックカラー（TC Suite調：OK/WARN/ALERT/INFO/VIOLET）
OK_COLOR = "#34D399"
WARN_COLOR = "#FBBF24"
ALERT_COLOR = "#FB5E7E"
INFO_COLOR = "#6CB6FF"
VIOLET_COLOR = "#C08CFF"

# v35 secbar（各ペイン見出し帯）用の配色。BG_HEADERと同系のネイビーバー。
BAR_LO = "#153A5A"
BAR_INK = "#CFE6F7"
BAR_EN = "#7FB2D9"

# シンタックスハイライト色（TC Suite調 — ネイビー地に映えるHUD配色）
SYNTAX_COLORS: list[tuple[str, str, str]] = [
    # (タグ名, 正規表現, 文字色)
    ("hl_comment", r"\([^)]*\)", "#5C6E62"),                          # コメント: オリーブグレー（視認性UP）
    ("hl_n_number", r"\bN\d+\b", "#F0D060"),                          # N番号: HUDアンバー（明るく）
    ("hl_o_number", r"\bO\d+", "#F0D060"),                             # O番号: HUDアンバー（明るく）
    ("hl_g_code", r"G-?(?:\d+(?:\.\d*)?|\.\d+)", "#FF6B6B"),           # Gコード: 警告レッド
    ("hl_m_code", r"M-?(?:\d+(?:\.\d*)?|\.\d+)", "#FF9557"),           # Mコード: 警告オレンジ
    ("hl_s_val", r"S-?(?:\d+(?:\.\d*)?|\.\d+)", "#3DDDFF"),            # S値: HUDシアン
    ("hl_f_val", r"F-?(?:\d+(?:\.\d*)?|\.\d+)", "#34D399"),            # F値: HUDグリーン
    ("hl_t_val", r"\bT\d+", "#66CCFF"),                                # T番号: HUDライトブルー
    ("hl_h_val", r"\bH\d+", "#88AAFF"),                                # H番号: HUDブルー
    ("hl_x_axis", r"X-?(?:\d+(?:\.\d*)?|\.\d+)", "#5BE584"),          # X軸: ブライトグリーン
    ("hl_u_axis", r"U-?(?:\d+(?:\.\d*)?|\.\d+)", "#5BE584"),          # U軸: ブライトグリーン
    ("hl_y_axis", r"Y-?(?:\d+(?:\.\d*)?|\.\d+)", "#2DDBC0"),          # Y軸: グリーン→ティール
    ("hl_v_axis", r"V-?(?:\d+(?:\.\d*)?|\.\d+)", "#2DDBC0"),          # V軸: グリーン→ティール
    ("hl_z_axis", r"Z-?(?:\d+(?:\.\d*)?|\.\d+)", "#3DC7EA"),          # Z軸: ティール→ブルー
    ("hl_w_axis", r"W-?(?:\d+(?:\.\d*)?|\.\d+)", "#3DC7EA"),          # W軸: ティール→ブルー
    ("hl_c_axis", r"C-?(?:\d+(?:\.\d*)?|\.\d+)", "#5FA8F5"),          # C軸: ブルー寄り
    ("hl_b_axis", r"B-?(?:\d+(?:\.\d*)?|\.\d+)", "#7C93FF"),          # B軸: ディープブルー
]

RESULT_STYLE_MAP = {
    "section_heading": {
        "background": "#15324A",
        "foreground": "#CFE6F7",
        "font": ("Yu Gothic UI", 11, "bold"),
        "spacing1": 10,
        "spacing3": 4,
        "lmargin1": 6,
        "lmargin2": 6,
    },
    "program_number": {
        "background": "#10161F",
        "foreground": "#FFFFFF",
        "font": ("Consolas", 12, "bold"),
        "spacing3": 6,
        "lmargin1": 6,
        "lmargin2": 6,
    },
    "summary_overspeed": {"background": "#301019", "foreground": "#FFC2CE", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "summary_overfeed": {"background": "#2A2008", "foreground": "#FDE9A8", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "summary_tcp": {"background": "#0F1D2E", "foreground": "#BBDDFF", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "summary_decimal_error": {"background": "#0E241C", "foreground": "#8FF0C7", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "summary_radius": {"background": "#101E16", "foreground": "#8FE8B8", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "summary_tailstock_macro_missing": {"background": "#221530", "foreground": "#D9C2FF", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "summary_duplicate_n": {"background": "#16202E", "foreground": "#B7D4F0", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "finding_overspeed": {"background": "#3A121C", "foreground": "#FFD0DA", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "finding_overfeed": {"background": "#362A0C", "foreground": "#FDE9A8", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "finding_tcp": {"background": "#12233A", "foreground": "#CFE6F7", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "finding_decimal_error": {"background": "#123328", "foreground": "#7FF0CB", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "finding_tailstock_macro_missing": {"background": "#241A34", "foreground": "#DCC4FF", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "finding_default": {"background": "#141B24", "foreground": "#C7D4DE", "lmargin1": 10, "lmargin2": 10, "spacing1": 2, "spacing3": 2},
    "block_label": {"background": "#142433", "foreground": "#A9E4DB", "font": ("Consolas", 11, "bold"), "spacing1": 8, "spacing3": 2, "lmargin1": 6, "lmargin2": 6},
    "tool_t": {"background": "#101B26", "foreground": "#88DDFF", "font": ("Consolas", 11, "bold"), "lmargin1": 14, "lmargin2": 14},
    "tool_name": {"background": "#10161F", "foreground": "#DEE7EE", "lmargin1": 14, "lmargin2": 14},
    "offset_h": {"background": "#101E28", "foreground": "#77DDEE", "lmargin1": 14, "lmargin2": 14},
    "radius_notice": {"background": "#0E2620", "foreground": "#5BE584", "font": ("Yu Gothic UI", 10, "bold"), "lmargin1": 14, "lmargin2": 14},
    "duplicate_n_notice": {"background": "#14273A", "foreground": "#9CCBEF", "font": ("Yu Gothic UI", 10, "bold"), "lmargin1": 14, "lmargin2": 14},
    "empty_notice": {"background": "#141B24", "foreground": "#C7D4DE", "lmargin1": 10, "lmargin2": 10},
}

# v35 右ペイン「目次」用の検査項目グループ定義。
# kind は Finding.kind（overspeed/overfeed/decimal_error/tailstock_macro_missing/tcp）
# のほか、SummaryData から集計する duplicate_n / radius_comp_n を含む7種。
TOC_GROUPS: list[dict] = [
    {"kind": "overspeed", "label": "回転数超え", "color": ALERT_COLOR, "sev": "err"},
    {"kind": "overfeed", "label": "送り超え", "color": "#FF9557", "sev": "err"},
    {"kind": "decimal_error", "label": "小数点間違い", "color": "#F0D060", "sev": "err"},
    {"kind": "tailstock_macro_missing", "label": "芯押しマクロ忘れ", "color": VIOLET_COLOR, "sev": "err"},
    {"kind": "tcp", "label": "G43.4 (TCP)", "color": INFO_COLOR, "sev": "wrn"},
    {"kind": "duplicate_n", "label": "N番号重複", "color": ACCENT, "sev": "wrn"},
    {"kind": "radius_comp_n", "label": "径補正使用N", "color": TEXT_MUTED, "sev": "info"},
]

# v35 M/G記号凡例の初期シードデータ。gm_glossary.json が存在しない初回起動時のみ使われ、
# 以降はJSONファイルが実データ（ユーザーが自由に追加・編集・削除できる）。
GM_GLOSSARY_SEED: dict[str, str] = {
    "G0": "送り",
    "G1": "直線切削",
    "G2": "円弧切削（時計回り）",
    "G3": "円弧切削（半時計回り）",
    "G4": "ドウェル",
    "G17": "平面選択（XY）",
    "G18": "平面選択（ZX)",
    "G19": "平面選択（YZ）",
    "G28": "機械原点復帰",
    "G30": "第2、第３、第４、原点復帰",
    "G38": "ワーク押し付け確認",
    "G40": "工具径補正（キャンセル）",
    "G41": "工具径補正（左）",
    "G42": "工具径補正（右）",
    "G43": "工具補正有効",
    "G49": "工具補正無効",
    "G50": "最高回転速度設定",
    "G54": "ワーク座標系１選択",
    "G55": "ワーク座標系２選択",
    "G56": "ワーク座標系３選択",
    "G57": "ワーク座標系４選択",
    "G58": "ワーク座標系５選択",
    "G59": "ワーク座標系６選択",
    "G68.1": "３次元座標変換",
    "G69.1": "３次元座標変換キャンセル",
    "G80": "穴あけ固定サイクルキャンセル",
    "G83": "端面ドリルサイクル",
    "G83.5": "端面ドリルサイクル（ステップ）",
    "G83.6": "端面ドリルサイクル（イニシャル点戻り）",
    "G84": "端面同期タップサイクル",
    "G85": "端面ボーリングサイクル",
    "G87": "側面ドリルサイクル",
    "G87.5": "側面ドリルサイクル（ステップ）",
    "G87.6": "側面ドリルサイクル（イニシャル点戻り）",
    "G88": "側面同期タップサイクル",
    "G89": "側面ボーリングサイクル",
    "G96": "切削速度指令",
    "G97": "回転速度指令",
    "G98": "毎分送り",
    "G99": "毎回転送り",
    "G330": "芯押し台/第二主軸原点復帰",
    "G361": "工具交換指令（機械原点経由）",
    "G362": "工具交換指令（第4原点経由）",
    "M0": "プログラム一時停止",
    "M1": "オプショナルストップ",
    "M2": "プログラム終了",
    "M3": "主軸正転（第一主軸）",
    "M4": "主軸逆転（第一主軸）",
    "M5": "主軸停止",
    "M8": "切削油ON",
    "M9": "切削油OFF",
    "M10": "チャッククランプ（第一主軸）",
    "M11": "チャックアンクランプ（第一主軸）",
    "M13": "工具正転",
    "M14": "工具逆転",
    "M19": "主軸定位置停止（第一主軸）",
    "M23": "チャンファリングON",
    "M24": "チャンファリングOFF",
    "M25": "芯押し台前進",
    "M26": "芯押し台後退",
    "M30": "リセット＆リワインド",
    "M33": "工具収納",
    "M34": "位相同期運転有効",
    "M35": "速度同期運転有効",
    "M36": "同期運転無効",
    "M45": "C軸接続",
    "M46": "C軸接続解除",
    "M68": "主軸ブレーキクランプ（第一主軸）",
    "M69": "主軸ブレーキアンクランプ（第一主軸）",
    "M80": "突っ切り確認",
    "M81": "ワーク押し付け確認有効",
    "M82": "ワーク押し付け確認無効",
    "M90": "第一主軸/工具主軸同時運転モードON",
    "M91": "第一主軸/工具主軸同時運転モードOFF",
    "M98": "サブプログラム呼出し",
    "M99": "サブプログラム終了",
    "M203": "第二主軸正転",
    "M204": "第二主軸逆転",
    "M210": "チャッククランプ（第二主軸）",
    "M211": "チャックアンクランプ（第二主軸）",
    "M219": "主軸定位置停止（第二主軸）",
    "M245": "C軸接続（第二主軸）",
    "M246": "C軸接続解除（第二主軸）",
    "M268": "主軸ブレーキクランプ（第二主軸）",
    "M269": "主軸ブレーキアンクランプ（第二主軸）",
    "M290": "第二主軸/工具主軸同時運転モードON",
    "M291": "第二主軸/工具主軸同時運転モードOFF",
    "M301": "ツールヘッドクランプ",
    "M302": "ツールヘッドアンクランプ",
    "M303": "第一主軸選択信号ON",
    "M304": "第二主軸選択信号ON",
    "M319": "工具主軸/回転工具主軸定位置停止",
    "M329": "主軸同期タップモードON",
    "M480": "C軸同期モードON",
    "M481": "C軸同期モードOFF",
    "M560": "工具主軸逆転モードON",
    "M561": "工具主軸逆転モードOFF",
    "M594": "B軸コンタリングモードON",
    "M595": "B軸コンタリングモードOFF",
    "M200": "チップコンベヤ正転",
    "M201": "チップコンベヤ停止",
    "M484": "スルースピンドルクーラントON",
    "M485": "スルースピンドルクーラントOFF",
    "G65P7001": "引き抜き後突っ切りマクロ（C2→C1）",
    "G65P7002": "突っ切り後網かごで回収マクロ",
    "G65P7003": "XY平面溝荒加工マクロ",
    "G65P7004": "受け渡しマクロ（C1→C2）",
    "G65P7005": "センサー、C軸中心割り出しマクロ",
    "G65P7006": "切りくず掃除マクロ",
    "G65P7007": "切りくず掃除マクロ",
    "G65P7008": "C軸割り出しマクロ",
    "G65P7009": "ひっかき棒製品回収マクロ",
    "G65P7011": "受け渡しマクロ（引き抜き）",
    "G65P7012": "C軸割り出しマクロ（外径）",
    "G65P7013": "C軸割り出しマクロ（C軸同期）",
    "G65P7014": "掴みかえマクロ（C1→C2引き抜き）",
}

# G65Pxxxx（マクロ呼出し）→ G小数点コード → 通常のG/Mコードの順で長いトークンを優先マッチ
GM_TOKEN_RE = re.compile(r"G65P\d+|G\d+\.\d+|G\d+|M\d+", re.IGNORECASE)


def extract_gm_tokens(line: str) -> list[str]:
    """1行からG/Mコードのトークンを（コメント除く）出現順に抽出する。表示専用の軽量パーサ。"""
    if is_comment_line(line.strip()):
        return []
    target = strip_inline_comments(line)
    seen: set[str] = set()
    tokens: list[str] = []
    for m in GM_TOKEN_RE.finditer(target):
        token = m.group(0).upper()
        if token not in seen:
            seen.add(token)
            tokens.append(token)
    return tokens


_MACRO_ARG_INT_RE = re.compile(r"^[+-]?\d+$")


def ensure_decimal_value(value: str) -> str:
    """マクロ呼出し支援：引数値が整数だけの場合、小数点を必須で付与する（例: "100" -> "100."）。
    既に小数点がある・数値以外（式やブラケット等）はそのまま返す。
    """
    value = value.strip()
    if value and _MACRO_ARG_INT_RE.fullmatch(value):
        return value + "."
    return value


@dataclass
class ThresholdSettings:
    speed_limit: float = 12000.0
    feed_limit: float = 40000.0


@dataclass
class ReportLine:
    text: str
    target_line_no: int | None = None
    style_key: str | None = None


@dataclass
class Finding:
    line_no: int
    text: str
    kind: str
    mode: str | None = None


@dataclass
class NBlock:
    n_label: str
    start_line_no: int
    rows: list[tuple[int, str]]


@dataclass
class SummaryData:
    overspeed_count: int = 0
    overfeed_count: int = 0
    tcp_count: int = 0
    decimal_error_count: int = 0
    radius_comp_n_count: int = 0
    tailstock_macro_missing_count: int = 0
    duplicate_n_count: int = 0
    duplicate_n_details: list = field(default_factory=list)  # [(n_label, [start_line_no, ...]), ...]


@dataclass
class ReportData:
    lines: list[ReportLine]
    summary: SummaryData
    error_line_nos: set[int]
    findings: list = field(default_factory=list)

    def to_text(self) -> str:
        return "\n".join(line.text for line in self.lines)


# =============================================================================
# マクロ機能用データモデル（v28〜）
# =============================================================================

@dataclass
class MacroStep:
    """マクロの1ステップ。既存テンプレ実行情報を1つ保持。
    params はテンプレ依存の自由形式 dict（factor, value, axis 等）。
    全ステップは "全体" スコープで実行される。
    """
    template_id: str
    template_label: str
    params: dict
    enabled: bool = True

    def to_dict(self) -> dict:
        return {
            "template_id": self.template_id,
            "template_label": self.template_label,
            "params": dict(self.params),
            "enabled": self.enabled,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "MacroStep":
        return cls(
            template_id=d.get("template_id", ""),
            template_label=d.get("template_label", ""),
            params=dict(d.get("params", {})),
            enabled=bool(d.get("enabled", True)),
        )


@dataclass
class Macro:
    """マクロ本体。ステップのリスト＋メタ情報。"""
    name: str = ""
    description: str = ""
    steps: list = field(default_factory=list)
    created_at: str = ""
    updated_at: str = ""

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "description": self.description,
            "steps": [s.to_dict() for s in self.steps],
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "format_version": 1,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Macro":
        return cls(
            name=d.get("name", ""),
            description=d.get("description", ""),
            steps=[MacroStep.from_dict(s) for s in d.get("steps", [])],
            created_at=d.get("created_at", ""),
            updated_at=d.get("updated_at", ""),
        )


def get_macros_dir() -> str:
    """マクロ保存ディレクトリ。実行ファイルと同じ階層に macros/ を作る。"""
    if getattr(sys, "frozen", False):
        # PyInstaller exe 化時
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    macros_dir = os.path.join(base, "macros")
    try:
        os.makedirs(macros_dir, exist_ok=True)
    except OSError:
        pass
    return macros_dir


def _app_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_gm_glossary_path() -> str:
    """G/M記号一覧（全件）を保存するJSONファイルのパス"""
    return os.path.join(_app_base_dir(), "gm_glossary.json")


def load_gm_glossary() -> dict[str, str]:
    """G/M記号一覧（全件）をJSONから読み込む。
    ファイルが無ければ組み込みシード(GM_GLOSSARY_SEED)で新規作成する。
    以降はこのJSONファイルが唯一の実データで、ユーザーが自由に追加・編集・削除できる。
    """
    path = get_gm_glossary_path()
    if not os.path.isfile(path):
        save_gm_glossary(GM_GLOSSARY_SEED)
        return dict(GM_GLOSSARY_SEED)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {str(k).upper(): str(v) for k, v in data.items()}
    except (OSError, json.JSONDecodeError, ValueError):
        pass
    return dict(GM_GLOSSARY_SEED)


def save_gm_glossary(entries: dict[str, str]) -> None:
    """G/M記号一覧（全件）をJSONへ保存する"""
    path = get_gm_glossary_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(entries, f, ensure_ascii=False, indent=2, sort_keys=True)
    except OSError:
        pass


def get_nc_macros_path() -> str:
    """NCカスタムマクロ（O番号サブプログラム）定義一覧を保存するJSONファイルのパス"""
    return os.path.join(_app_base_dir(), "nc_macros.json")


# マクロ入力支援機能の初期サンプル（ユーザー提供のO7021をそのまま同梱）
DEFAULT_NC_MACRO_TEXT = """%
O7021(ORIGIN AUTO SET MACRO)

(==ARGUMENT==)
(A=#1  PART LENGTH)
(B=#2  FINISH ALLOWANCE)
(C=#3  TOUCH C AXIS ANGLE)
(E=#8  MODE 1=JAW 2=WORK 3=JIG 4=CUTOFF)
(H=#11 CHUCK TO MEASURE FACE)
(Q=#17 SPINDLE 1 OR 2)
(X=#24 TOUCH X POSITION)
%
"""

MACRO_ARG_LINE_RE = re.compile(r"^\(([A-Z])=#\d+\s+(.+)\)\s*$")
MACRO_HEADER_RE = re.compile(r"^O(\d+)\s*\(([^)]*)\)")


def parse_macro_definition(text: str) -> dict | None:
    """NCカスタムマクロ本文からO番号・タイトル・引数一覧(A〜Z=#N 説明)を抽出する（自動解析）。
    マクロ呼出し支援ダイアログの入力フォーム生成専用の軽量パーサ（検査ロジックとは無関係）。
    戻り値はそのままマクロ定義編集ダイアログへプリフィルされ、手動で修正できる。
    """
    o_number = None
    title = ""
    args: list[dict] = []
    seen_letters: set[str] = set()
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if o_number is None:
            m = MACRO_HEADER_RE.match(line)
            if m:
                o_number = f"O{m.group(1)}"
                title = m.group(2).strip()
                continue
        m = MACRO_ARG_LINE_RE.match(line)
        if m:
            letter, desc = m.group(1), m.group(2).strip()
            if letter not in seen_letters:
                seen_letters.add(letter)
                args.append({"letter": letter, "desc": desc})
    if o_number is None:
        return None
    return {"o_number": o_number, "title": title, "args": args, "memo": "", "raw_text": text}


def load_nc_macro_defs() -> list[dict]:
    """マクロ定義一覧をJSONから読み込む。
    ファイルが無ければユーザー提供のO7021を自動解析してシードとして書き込む。
    """
    path = get_nc_macros_path()
    if not os.path.isfile(path):
        seed = parse_macro_definition(DEFAULT_NC_MACRO_TEXT)
        defs = [seed] if seed else []
        save_nc_macro_defs(defs)
        return defs
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            defs = []
            for item in data:
                if isinstance(item, dict) and item.get("o_number"):
                    item.setdefault("title", "")
                    item.setdefault("args", [])
                    item.setdefault("memo", "")
                    item.setdefault("raw_text", "")
                    defs.append(item)
            defs.sort(key=lambda d: d["o_number"])
            return defs
    except (OSError, json.JSONDecodeError, ValueError):
        pass
    return []


def save_nc_macro_defs(defs: list[dict]) -> None:
    """マクロ定義一覧をJSONへ保存する"""
    path = get_nc_macros_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(defs, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def sanitize_macro_filename(name: str) -> str:
    """マクロ名をファイル名として安全な形に変換"""
    # OS で使えない文字を除去
    safe = re.sub(r'[\\/:*?"<>|]', "_", name)
    safe = safe.strip().strip(".")
    return safe if safe else "macro"


def normalize_fragment(line: str) -> str:
    return re.sub(r"\s+", "", line.strip())


def is_comment_line(line: str) -> bool:
    s = line.strip()
    return s.startswith("(") and s.endswith(")")


def strip_inline_comments(line: str) -> str:
    """行中の (...)コメント部分を除去して返す"""
    return re.sub(r"\([^)]*\)", "", line)


def is_g65_line(line: str) -> bool:
    """G65マクロ呼び出し行かどうか"""
    return bool(re.search(r"G65(?!\d)", normalize_fragment(line), flags=re.IGNORECASE))


def parse_axis_decimal_error(line: str) -> str | None:
    """軸指令の小数点間違いを検出する。
    - 小数点2個以上: X100..5, Z50.2.3, F1.5.0 等（XYZUVW + F が対象）
    - 小数点なし（非ゼロ整数値）: X100, Z-50 等（XYZUVW のみ。F は対象外）
    """
    stripped = line.strip()
    # 小数点2個以上: X100..5, Z50.2.3, U-3..0, F1.5.0 等
    multi_dot = re.search(r"([XYZUVWF]-?\d*\.\d*\.[\d.]*)", stripped, flags=re.IGNORECASE)
    if multi_dot:
        return multi_dot.group(1).upper()
    # 小数点なし（XYZUVWのみ。Fは小数点なしでもOKなので除外）
    missing = re.search(r"([XYZUVW]-?(?!0(?!\d))\d+)(?=\s|$|[^0-9.])", stripped, flags=re.IGNORECASE)
    if missing:
        return missing.group(1).upper()
    return None


TURNING_TOOL_HINTS = (
    "DCLNN", "DTJNR", "STLPR", "DDJNR", "SCLCR", "SVJBR", "PCLNR", "PDJNR",
    "MTJNR", "MVJNR", "MVLNR", "DNMG", "CNMG", "TNMG", "VNMG", "WNMG",
)
MILLING_TOOL_HINTS = (
    "ENDMILL", "EMILL", "MILL", "DRILL", "TAP", "REAM", "BSL", "WEX", "WEZ",
    "ALPEN", "CARBIDE DRILL", "CENTER DRILL",
)


class ModeHeuristic:
    def __init__(self) -> None:
        self.current_n = ""
        self.turning_score = 0
        self.milling_score = 0

    def reset_for_new_block(self, n_label: str) -> None:
        self.current_n = n_label
        self.turning_score = 0
        self.milling_score = 0

    def update(self, line: str) -> None:
        compact = normalize_fragment(line).upper()
        if not compact:
            return

        if any(keyword in compact for keyword in TURNING_TOOL_HINTS):
            self.turning_score += 3
        if any(keyword in compact for keyword in MILLING_TOOL_HINTS):
            self.milling_score += 3

        if re.search(r"G96(?!\d)", compact):
            self.turning_score += 3
        if re.search(r"G97(?!\d)", compact):
            self.turning_score += 1
        if re.search(r"G50S[+-]?\d", compact):
            self.turning_score += 2
        if re.search(r"G43\.4(?!\d)", compact):
            self.milling_score += 4
        if re.search(r"G17(?!\d)|G19(?!\d)", compact):
            self.milling_score += 2
        if re.search(r"G18(?!\d)", compact):
            self.turning_score += 1
        if re.search(r"(?:^|[^A-Z])(Y|A|B|C)[+-]?\d", compact):
            self.milling_score += 1
        if re.search(r"M13[3-6](?!\d)", compact):
            self.milling_score += 2

    def classify(self) -> str:
        if self.milling_score > self.turning_score:
            return MODE_MILLING
        return MODE_TURNING


def classify_mode_for_line(line: str, mode_heuristic: ModeHeuristic) -> str:
    preview = ModeHeuristic()
    preview.current_n = mode_heuristic.current_n
    preview.turning_score = mode_heuristic.turning_score
    preview.milling_score = mode_heuristic.milling_score
    preview.update(line)
    return preview.classify()


def detect_speed_findings(line: str, line_no: int, settings: ThresholdSettings, mode_heuristic: ModeHeuristic) -> list[Finding]:
    compact = normalize_fragment(line)
    findings: list[Finding] = []

    for match in re.finditer(r"S([+-]?(?:\d+(?:\.\d+)?|\.\d+))", compact, flags=re.IGNORECASE):
        try:
            value = float(match.group(1))
        except ValueError:
            continue
        if value >= settings.speed_limit:
            findings.append(Finding(line_no=line_no, text=compact, kind="overspeed"))
            break
    return findings


def detect_feed_findings(line: str, line_no: int, settings: ThresholdSettings, mode_heuristic: ModeHeuristic) -> list[Finding]:
    compact = normalize_fragment(line)
    findings: list[Finding] = []

    for match in re.finditer(r"F([+-]?(?:\d+(?:\.\d+)?|\.\d+))", compact, flags=re.IGNORECASE):
        try:
            value = float(match.group(1))
        except ValueError:
            continue
        if value >= settings.feed_limit:
            findings.append(Finding(line_no=line_no, text=compact, kind="overfeed"))
            break
    return findings


def detect_tcp_finding(line: str, line_no: int) -> Finding | None:
    compact = normalize_fragment(line)
    if re.search(r"G43\.4(?!\d)", compact, flags=re.IGNORECASE):
        return Finding(line_no=line_no, text="G43.4", kind="tcp")
    return None


def detect_decimal_error_finding(line: str, line_no: int) -> Finding | None:
    token = parse_axis_decimal_error(line)
    if token:
        return Finding(line_no=line_no, text=token, kind="decimal_error")
    return None


def detect_tailstock_macro_missing_findings(lines: list[str]) -> list[Finding]:
    findings: list[Finding] = []
    m25_seen = False

    for line_no, line in enumerate(lines, start=1):
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue

        check_target = strip_inline_comments(stripped)
        compact = normalize_fragment(check_target).upper()
        if not compact:
            continue

        if re.search(r"M25(?!\d)", compact):
            m25_seen = True

        if m25_seen and re.search(r"G361(?!\d)", compact):
            findings.append(Finding(line_no=line_no, text="芯押しマクロ忘れ", kind="tailstock_macro_missing"))
            m25_seen = False

    return findings


def detect_radius_comp_used(line: str) -> bool:
    compact = normalize_fragment(line).upper()
    return bool(re.search(r"G41(?!\d)|G42(?!\d)", compact))


def extract_program_number(lines: list[str]) -> tuple[str, int | None]:
    for line_no, line in enumerate(lines, start=1):
        match = re.match(r"^\s*(O\d+)(?:\s*\([^)]*\))?\s*$", line, flags=re.IGNORECASE)
        if match:
            return match.group(1).upper(), line_no
    return "O番号なし", None


def extract_program_comment(lines: list[str]) -> str:
    for line in lines:
        match = re.match(r"^\s*O\d+\s*\(([^)]*)\)\s*$", line, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""


def split_n_blocks(lines: list[str]) -> list[NBlock]:
    blocks: list[NBlock] = []
    current_label = ""
    current_line_no = 0
    current_rows: list[tuple[int, str]] = []

    for line_no, line in enumerate(lines, start=1):
        match = re.match(r"^\s*(N\d+)\b", line, flags=re.IGNORECASE)
        if match:
            if current_rows:
                blocks.append(NBlock(n_label=current_label, start_line_no=current_line_no, rows=current_rows))
            current_label = match.group(1).upper()
            current_line_no = line_no
            current_rows = [(line_no, line)]
        elif current_rows:
            current_rows.append((line_no, line))

    if current_rows:
        blocks.append(NBlock(n_label=current_label, start_line_no=current_line_no, rows=current_rows))

    return blocks


def extract_tool_name(block_rows: list[tuple[int, str]]) -> tuple[str, int | None]:
    # 最初の行（N番号を含む行）に (...) コメントがあれば優先で拾う
    if block_rows:
        first_line_no, first_line = block_rows[0]
        m = re.search(r"\(([^)]*)\)", first_line)
        if m:
            return f"({m.group(1)})", first_line_no
    # 次行以降のコメント行を探す（従来動作）
    for line_no, line in block_rows[1:]:
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("("):
            return stripped, line_no
    return "", None


def extract_t_before_g36x(block_rows: list[tuple[int, str]]) -> tuple[str, int | None]:
    # T番号パターン: 直前が英字でない & 直後が小数点でない
    # 例: G0T2001, M6T2001, N100T2001, T2001 → ヒット
    # 例: XT2001, CT2001 → ヒットしない（軸文字直後は除外）
    # 表示は T + 先頭4桁固定
    t_pat = re.compile(r"(?<![A-Z])T(\d+)(?!\.)", flags=re.IGNORECASE)

    def _format_t(digits: str) -> str:
        return f"T{digits[:4]}"

    g36x_index: int | None = None
    for index, (line_no, line) in enumerate(block_rows):
        if re.search(r"G36[12](?!\d)", line, flags=re.IGNORECASE):
            same_line_match = t_pat.search(line)
            if same_line_match:
                return _format_t(same_line_match.group(1)), line_no
            g36x_index = index
            break

    if g36x_index is None:
        return "", None

    for line_no, line in reversed(block_rows[:g36x_index]):
        match = t_pat.search(line)
        if match:
            return _format_t(match.group(1)), line_no
    return "", None


def extract_h_list(block_rows: list[tuple[int, str]]) -> list[tuple[str, int]]:
    h_list: list[tuple[str, int]] = []
    seen: set[str] = set()

    for line_no, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        for match in re.finditer(r"H(\d+)", stripped, flags=re.IGNORECASE):
            num = match.group(1)
            if int(num) == 0:
                continue
            token = f"H{num}"
            if token not in seen:
                seen.add(token)
                h_list.append((token, line_no))

    return h_list


def block_radius_comp_line(block_rows: list[tuple[int, str]]) -> int | None:
    for line_no, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        if detect_radius_comp_used(stripped):
            return line_no
    return None


def extract_spindle_info(block_rows: list[tuple[int, str]]) -> dict[str, list[str]]:
    """Nブロック内のG50/G96/G97ごとのS値をモーダル管理で抽出する。
    戻り値: {"G50": ["3000"], "G96": ["180", "220"], "G97": ["2000"]} のような辞書
    """
    result: dict[str, list[str]] = {}
    current_mode: str = ""  # "", "G50", "G96", "G97"

    for _, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        code = strip_inline_comments(stripped)
        compact = normalize_fragment(code).upper()
        if not compact:
            continue

        # G50は同一行でS値を拾う（クランプ指定は必ず同一行）
        g50_match = re.search(r"G50(?!\d).*?S([+-]?(?:\d+(?:\.\d+)?|\.\d+))", compact)
        if g50_match:
            val = g50_match.group(1)
            result.setdefault("G50", [])
            if val not in result["G50"]:
                result["G50"].append(val)

        # G96/G97のモード切り替え
        if re.search(r"G96(?!\d)", compact):
            current_mode = "G96"
        elif re.search(r"G97(?!\d)", compact):
            current_mode = "G97"

        # 現在のモードに応じてS値を拾う（G50のS値は除外済み）
        if current_mode in ("G96", "G97", ""):
            for m in re.finditer(r"S([+-]?(?:\d+(?:\.\d+)?|\.\d+))", compact):
                val = m.group(1)
                if g50_match and val == g50_match.group(1):
                    continue
                result.setdefault(current_mode, [])
                if val not in result[current_mode]:
                    result[current_mode].append(val)

    return result


def format_spindle_info(info: dict[str, list[str]]) -> str:
    """extract_spindle_infoの結果を表示用文字列に変換"""
    parts = []
    for key in ("G50", "G96", "G97", ""):
        if key in info:
            vals = ", ".join(f"S{v}" for v in info[key])
            if key == "":
                parts.append(vals)  # モード不明はS値だけ
            else:
                parts.append(f"{key} {vals}")
    return "\n".join(parts) if parts else ""


def extract_feed_list(block_rows: list[tuple[int, str]]) -> str:
    """Nブロック内のF値を重複除外で全部列挙する"""
    seen: list[str] = []
    for _, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        if is_g65_line(stripped):
            continue
        code = strip_inline_comments(stripped)
        compact = normalize_fragment(code)
        for m in re.finditer(r"F([+-]?(?:\d+(?:\.\d+)?|\.\d+))", compact, flags=re.IGNORECASE):
            val = m.group(1)
            if val not in seen:
                seen.append(val)
    return ", ".join(f"F{v}" for v in seen)


# クーラントMコード → 表示名 のマップ
COOLANT_M_MAP = {
    "M8":   "ジェット",
    "M484": "スルースピンドル",
}


def extract_coolant_info(block_rows: list[tuple[int, str]]) -> list[tuple[str, str]]:
    """Nブロック内のクーラントMコード(M8/M484)を出現順・重複除外で抽出する。
    Returns: [(Mコード, 表示名), ...]
    """
    found: list[tuple[str, str]] = []
    seen_codes: set[str] = set()
    for _, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        if is_g65_line(stripped):
            continue
        code = strip_inline_comments(stripped)
        compact = normalize_fragment(code)
        # M番号を全部拾って、クーラント対象だけ通す
        # M8 と M484 は数字続き禁止で区別（M80, M84 とは別物）
        for m in re.finditer(r"M(\d+)", compact, flags=re.IGNORECASE):
            num = m.group(1)
            mcode = f"M{int(num)}"  # ゼロ埋め除去
            if mcode in COOLANT_M_MAP and mcode not in seen_codes:
                seen_codes.add(mcode)
                found.append((mcode, COOLANT_M_MAP[mcode]))
    return found


def format_coolant_info(coolant: list[tuple[str, str]]) -> str:
    """クーラント情報を表示用文字列に整形。
    例: [("M8","ジェット"), ("M484","スルースピンドル")]
        → "M8 ジェット\nM484 スルースピンドル"
    """
    if not coolant:
        return ""
    return "\n".join(f"{mc} {name}" for mc, name in coolant)


# 回転方向 Mコード → 表示名 のマップ
ROTATION_M_MAP = {
    "M3":   "主軸 正転",
    "M4":   "主軸 逆転",
    "M13":  "ミーリング 正転",
    "M14":  "ミーリング 逆転",
    "M203": "主軸 正転",
    "M204": "主軸 逆転",
}


def extract_rotation_info(block_rows: list[tuple[int, str]]) -> list[tuple[str, str]]:
    """Nブロック内の回転方向Mコード(M3/M4/M13/M14)を出現順・重複除外で抽出する。
    Returns: [(Mコード, 表示名), ...]
    """
    found: list[tuple[str, str]] = []
    seen_codes: set[str] = set()
    for _, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        if is_g65_line(stripped):
            continue
        code = strip_inline_comments(stripped)
        compact = normalize_fragment(code)
        # M3/M4/M13/M14 は数字続き禁止で M30/M40/M130/M140 と区別
        for m in re.finditer(r"M(\d+)(?!\d)", compact, flags=re.IGNORECASE):
            num = m.group(1)
            mcode = f"M{int(num)}"  # ゼロ埋め除去
            if mcode in ROTATION_M_MAP and mcode not in seen_codes:
                seen_codes.add(mcode)
                found.append((mcode, ROTATION_M_MAP[mcode]))
    return found


def format_rotation_info(rot: list[tuple[str, str]]) -> str:
    """回転方向情報を表示用文字列に整形"""
    if not rot:
        return ""
    return "\n".join(f"{mc} {name}" for mc, name in rot)


# ワーク座標系 Gコード → 表示名 のマップ
WORK_COORD_G_MAP = {
    "G54": "G54",  "G55": "G55",  "G56": "G56",
    "G57": "G57",  "G58": "G58",  "G59": "G59",
}


def extract_work_coordinate(block_rows: list[tuple[int, str]]) -> list[str]:
    """Nブロック内のワーク座標系指令(G54〜G59)を出現順・重複除外で抽出する。
    Returns: ["G54", "G55", ...]
    """
    found: list[str] = []
    seen: set[str] = set()
    for _, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        if is_g65_line(stripped):
            continue
        code = strip_inline_comments(stripped)
        compact = normalize_fragment(code)
        for m in re.finditer(r"G(\d+)(?!\d)", compact, flags=re.IGNORECASE):
            num = m.group(1)
            gcode = f"G{int(num)}"
            if gcode in WORK_COORD_G_MAP and gcode not in seen:
                seen.add(gcode)
                found.append(gcode)
    return found


def format_work_coordinate(wc: list[str]) -> str:
    """ワーク座標系情報を表示用文字列に整形"""
    if not wc:
        return ""
    return ", ".join(wc)


def extract_b_axis_angles(block_rows: list[tuple[int, str]]) -> list[str]:
    """Nブロック内のB軸角度値を抽出する。
    ※ G361 / G362 と同じ行にあるBのみ拾う（加工系の工具姿勢指令のみ対象）
    ※ G0 B-90. のような ATC 準備動作などは除外される
    Returns: ["0", "90", "-30"] のような重複除外リスト
    """
    found: list[str] = []
    seen: set[str] = set()
    for _, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        if is_g65_line(stripped):
            continue
        code = strip_inline_comments(stripped)
        compact = normalize_fragment(code)
        # この行に G361 または G362 が含まれているかチェック
        if not re.search(r"G36[12](?!\d)", compact, flags=re.IGNORECASE):
            continue
        # G361/G362 と同じ行にあるBを拾う
        for m in re.finditer(r"(?<![A-Za-z])B([+-]?(?:\d+(?:\.\d+)?|\.\d+))",
                              compact, flags=re.IGNORECASE):
            val = m.group(1)
            if val not in seen:
                seen.add(val)
                found.append(val)
    return found


def format_b_axis_angles(angles: list[str]) -> str:
    """B軸角度情報を表示用文字列に整形
    例: ["0", "90", "-30"] → "B: 0, 90, -30"
    """
    if not angles:
        return ""
    return "B: " + ", ".join(angles)


def classify_block_mode(block_rows: list[tuple[int, str]]) -> str:
    """Nブロック単体のモード（ターニング/ミーリング）を判定する"""
    heuristic = ModeHeuristic()
    for _, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        if is_g65_line(stripped):
            heuristic.update(stripped)
            continue
        heuristic.update(strip_inline_comments(stripped))
    return heuristic.classify()


def extract_z_range(block_rows: list[tuple[int, str]]) -> tuple[tuple[float, int] | None, tuple[float, int] | None]:
    """Nブロック内のZ値から最小値＋行番号、最大値＋行番号を返す。
    方針: モーダル追跡はしない。Zが明示された行のみ拾う。
    除外: コメント行、コメント内Z、G65マクロ呼び出し行、
          機械座標系退避指令(G28/G30/G53)、ワーク座標系設定(G50でZが含まれる形)。
    """
    z_values: list[tuple[float, int]] = []

    for line_no, line in block_rows:
        stripped = line.strip()
        if not stripped or is_comment_line(stripped):
            continue
        if is_g65_line(stripped):
            continue
        code = strip_inline_comments(stripped)
        compact = normalize_fragment(code).upper()
        if not compact:
            continue

        # 退避指令・ワーク座標設定は除外（実加工位置ではない）
        if re.search(r"G(?:28|30|53)(?!\d)", compact):
            continue
        if re.search(r"G50(?!\d)", compact):
            continue

        # Z値を拾う
        z_match = re.search(r"Z([+-]?(?:\d+(?:\.\d+)?|\.\d+))", compact)
        if z_match:
            try:
                val = float(z_match.group(1))
                z_values.append((val, line_no))
            except ValueError:
                pass

    if not z_values:
        return None, None

    z_min = min(z_values, key=lambda x: x[0])
    z_max = max(z_values, key=lambda x: x[0])
    return z_min, z_max


def format_z_with_line(z_info: tuple[float, int] | None) -> str:
    """(値, 行番号)タプルを表示用文字列に整形"""
    if z_info is None:
        return ""
    val, line_no = z_info
    if val == int(val):
        val_str = f"Z{int(val)}."
    else:
        val_str = f"Z{val:g}"
    return f"{val_str}\n({line_no}行)"


def collect_program_findings(lines: list[str], settings: ThresholdSettings) -> list[Finding]:
    findings: list[Finding] = []
    heuristic = ModeHeuristic()

    findings.extend(detect_tailstock_macro_missing_findings(lines))

    for line_no, line in enumerate(lines, start=1):
        stripped = line.strip()
        if not stripped:
            continue

        n_match = re.match(r"^\s*(N\d+)\b", line, flags=re.IGNORECASE)
        if n_match:
            heuristic.reset_for_new_block(n_match.group(1).upper())

        if not is_comment_line(stripped):
            if is_g65_line(stripped):
                heuristic.update(stripped)
                continue
            check_target = strip_inline_comments(stripped)
            findings.extend(detect_speed_findings(check_target, line_no, settings, heuristic))
            findings.extend(detect_feed_findings(check_target, line_no, settings, heuristic))
            tcp = detect_tcp_finding(check_target, line_no)
            if tcp:
                findings.append(tcp)
            decimal_error = detect_decimal_error_finding(check_target, line_no)
            if decimal_error:
                findings.append(decimal_error)

        heuristic.update(stripped)

    deduped: list[Finding] = []
    seen: set[tuple[int, str, str]] = set()
    for finding in findings:
        key = (finding.line_no, finding.text, finding.kind)
        if key not in seen:
            seen.add(key)
            deduped.append(finding)
    return deduped


def build_report(program_text: str, settings: ThresholdSettings) -> ReportData:
    lines = program_text.splitlines()
    if not lines:
        return ReportData(lines=[ReportLine("入力が空です。", style_key="empty_notice")], summary=SummaryData(), error_line_nos=set(), findings=[])

    report_lines: list[ReportLine] = []
    summary = SummaryData()

    program_number, program_line_no = extract_program_number(lines)
    report_lines.append(ReportLine(program_number, program_line_no, "program_number"))
    report_lines.append(ReportLine(""))

    findings = collect_program_findings(lines, settings)
    error_line_nos = {finding.line_no for finding in findings}
    for finding in findings:
        if finding.kind == "overspeed":
            summary.overspeed_count += 1
        elif finding.kind == "overfeed":
            summary.overfeed_count += 1
        elif finding.kind == "tcp":
            summary.tcp_count += 1
        elif finding.kind == "decimal_error":
            summary.decimal_error_count += 1
        elif finding.kind == "tailstock_macro_missing":
            summary.tailstock_macro_missing_count += 1

    radius_comp_n_count = 0
    blocks = split_n_blocks(lines)
    for block in blocks:
        if block_radius_comp_line(block.rows) is not None:
            radius_comp_n_count += 1
    summary.radius_comp_n_count = radius_comp_n_count

    # N番号重複検出: 同じN番号が複数回出現するブロックを集計
    n_label_lines: dict[str, list[int]] = {}
    for block in blocks:
        n_label_lines.setdefault(block.n_label, []).append(block.start_line_no)
    duplicates = [(label, line_list) for label, line_list in n_label_lines.items()
                  if len(line_list) >= 2]
    summary.duplicate_n_count = len(duplicates)
    summary.duplicate_n_details = duplicates

    report_lines.append(ReportLine("═══ SUMMARY / サマリ ═══════════════════════════", style_key="section_heading"))
    report_lines.append(ReportLine(f"回転数超え: {summary.overspeed_count}件", style_key="summary_overspeed"))
    report_lines.append(ReportLine(f"送り超え: {summary.overfeed_count}件", style_key="summary_overfeed"))
    report_lines.append(ReportLine(f"G43.4: {summary.tcp_count}件", style_key="summary_tcp"))
    report_lines.append(ReportLine(f"小数点間違い: {summary.decimal_error_count}件", style_key="summary_decimal_error"))
    report_lines.append(ReportLine(f"径補正使用N: {summary.radius_comp_n_count}件", style_key="summary_radius"))
    report_lines.append(ReportLine(f"芯押しマクロ忘れ: {summary.tailstock_macro_missing_count}件", style_key="summary_tailstock_macro_missing"))
    report_lines.append(ReportLine(f"N番号重複: {summary.duplicate_n_count}件", style_key="summary_duplicate_n"))
    report_lines.append(ReportLine(""))

    report_lines.append(ReportLine("═══ SCAN RESULT / プログラムチェック ═══════════", style_key="section_heading"))
    finding_style_map = {
        "overspeed": "finding_overspeed",
        "overfeed": "finding_overfeed",
        "tcp": "finding_tcp",
        "decimal_error": "finding_decimal_error",
        "tailstock_macro_missing": "finding_tailstock_macro_missing",
    }
    for finding in findings:
        report_lines.append(ReportLine(f"{finding.line_no}行{finding.text}", finding.line_no, finding_style_map.get(finding.kind, "finding_default")))

    # 重複N番号 → 出現行リスト の辞書（ブロック内マーク用）
    duplicate_n_map = {label: line_list for label, line_list in duplicates}

    for block in blocks:
        report_lines.append(ReportLine(""))
        # 「─── BLOCK N1 (L245-L398) ──────────」形式
        # 行範囲を block.rows から計算
        row_lines = [ln for ln, _ in block.rows]
        if row_lines:
            line_min = min(row_lines)
            line_max = max(row_lines)
            range_str = f" (L{line_min}-L{line_max})"
        else:
            range_str = ""
        prefix = f"─── BLOCK {block.n_label}{range_str} "
        block_header = prefix + "─" * max(3, 50 - len(prefix))
        report_lines.append(ReportLine(block_header, block.start_line_no, "block_label"))

        t_number, t_line_no = extract_t_before_g36x(block.rows)
        if t_number:
            report_lines.append(ReportLine(t_number, t_line_no, "tool_t"))

        tool_name, tool_line_no = extract_tool_name(block.rows)
        if tool_name:
            report_lines.append(ReportLine(tool_name, tool_line_no, "tool_name"))

        for h_token, h_line_no in extract_h_list(block.rows):
            report_lines.append(ReportLine(h_token, h_line_no, "offset_h"))

        radius_comp_line = block_radius_comp_line(block.rows)
        if radius_comp_line is not None:
            report_lines.append(ReportLine("径補正使用", radius_comp_line, "radius_notice"))

        # N番号重複マーク（このブロックのN番号が他の場所でも使われている場合）
        if block.n_label in duplicate_n_map:
            other_lines = [ln for ln in duplicate_n_map[block.n_label]
                            if ln != block.start_line_no]
            if other_lines:
                others_str = ", ".join(f"L{ln}" for ln in other_lines)
                report_lines.append(ReportLine(
                    f"N番号重複（他: {others_str}）",
                    block.start_line_no, "duplicate_n_notice"
                ))

    return ReportData(lines=report_lines, summary=summary, error_line_nos=error_line_nos, findings=findings)


class ThresholdEntry(tk.Frame):
    """回転数・送りの判定上限を入力するエントリー"""

    def __init__(self, master: tk.Misc, title: str, default_value: float, on_change) -> None:
        super().__init__(master, bg="#101720", bd=1, relief="solid")
        self.configure(padx=12, pady=10)
        self._on_change = on_change
        tk.Label(self, text=title, bg="#101720", fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(anchor="w")
        self.var = tk.StringVar(value=str(int(default_value)))
        self.entry = tk.Entry(
            self, textvariable=self.var,
            font=("Consolas", 14, "bold"),
            bg=INPUT_BG, fg=ACCENT,
            insertbackground=ACCENT, relief="solid", bd=1,
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT, justify="right", width=10,
        )
        self.entry.pack(anchor="w", pady=(4, 0), fill="x")
        self.var.trace_add("write", self._handle_change)

    def _handle_change(self, *_args) -> None:
        if callable(self._on_change):
            self._on_change()

    def get_value(self) -> float | None:
        raw = self.var.get().strip()
        if not raw:
            return None
        try:
            return float(raw)
        except ValueError:
            return None


class VerdictBand(tk.Frame):
    """右ペイン最上部の細い判定帯。ランプ＋判定文言＋ERR/WRN/LINES件数のみを表示する。
    v34のエヴァ風大型バナーに代わる、TC Suite調の控えめな帯。
    """

    # state -> (lamp/vt色, vt文言, vs文言)
    _STATES = {
        "idle":     (TEXT_MUTED, "――", "チェック未実行"),
        "ok":       (OK_COLOR, "OK", "問題なし"),
        "wn":       (WARN_COLOR, "WRN", "軽度の警告"),
        "ng":       (ALERT_COLOR, "NG", "危険箇所あり"),
    }

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, bg=INPUT_BG, bd=0, highlightthickness=0)
        self.configure(padx=13, pady=10)

        row = tk.Frame(self, bg=INPUT_BG)
        row.pack(fill="x")

        self._lamp = tk.Canvas(row, width=10, height=10, bg=INPUT_BG,
                               highlightthickness=0, bd=0)
        self._lamp_oval = self._lamp.create_oval(1, 1, 9, 9, fill=TEXT_MUTED, outline="")
        self._lamp.pack(side="left", padx=(0, 11))

        self._vt = tk.Label(row, text="――", bg=INPUT_BG, fg=TEXT_MUTED,
                            font=("Consolas", 15, "bold"))
        self._vt.pack(side="left")
        self._vs = tk.Label(row, text="チェック未実行", bg=INPUT_BG, fg=TEXT_MUTED,
                            font=("Yu Gothic UI", 10))
        self._vs.pack(side="left", padx=(11, 0))

        tally = tk.Frame(row, bg=INPUT_BG)
        tally.pack(side="right")
        self._err_lbl = self._make_chip(tally, "ERR 0", ALERT_COLOR, "#5A2530")
        self._err_lbl.pack(side="left", padx=(0, 6))
        self._wrn_lbl = self._make_chip(tally, "WRN 0", WARN_COLOR, "#6E5520")
        self._wrn_lbl.pack(side="left", padx=(0, 6))
        self._line_lbl = self._make_chip(tally, "0 L", TEXT_MUTED, BORDER)
        self._line_lbl.pack(side="left")

    @staticmethod
    def _make_chip(master: tk.Misc, text: str, fg: str, border: str) -> tk.Label:
        return tk.Label(
            master, text=text, bg="#0C1219", fg=fg,
            font=("Consolas", 9, "bold"), padx=8, pady=3,
            highlightthickness=1, highlightbackground=border, highlightcolor=border,
        )

    def set_state(self, state: str, err: int = 0, wrn: int = 0, lines: int = 0) -> None:
        """state: 'idle' | 'ok' | 'wn' | 'ng'"""
        if state not in self._STATES:
            state = "idle"
        color, vt_text, vs_text = self._STATES[state]
        self._lamp.itemconfigure(self._lamp_oval, fill=color)
        self._vt.configure(text=vt_text, fg=color)
        self._vs.configure(text=vs_text)
        self._err_lbl.configure(text=f"ERR {err}")
        self._wrn_lbl.configure(text=f"WRN {wrn}")
        self._line_lbl.configure(text=f"{lines} L")


class MiniMap(tk.Canvas):
    """中央ペイン右端の細いミニマップ。全行の危険(赤)/警告(黄)分布とNブロック先頭を一望し、
    クリックで一番近い検出行へジャンプする。
    """

    def __init__(self, master: tk.Misc, on_jump, width: int = 18) -> None:
        super().__init__(master, width=width, bg="#06090D", highlightthickness=0, bd=0,
                         cursor="hand2")
        self._on_jump = on_jump
        self._total_lines = 1
        self._markers: list[tuple[int, str]] = []  # (line_no, 'err'|'wrn')
        self._n_starts: list[int] = []
        self._viewport: tuple[float, float] = (0.0, 1.0)
        self.bind("<Configure>", lambda _e: self._redraw())
        self.bind("<Button-1>", self._on_click)

    def set_data(self, total_lines: int, markers: list[tuple[int, str]], n_starts: list[int]) -> None:
        self._total_lines = max(1, total_lines)
        self._markers = markers
        self._n_starts = n_starts
        self._redraw()

    def set_viewport(self, first: float, last: float) -> None:
        try:
            self._viewport = (float(first), float(last))
        except (TypeError, ValueError):
            return
        self._redraw()

    def _redraw(self) -> None:
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1 or h <= 1:
            return
        total = self._total_lines
        top = self._viewport[0] * h
        bottom = max(self._viewport[1] * h, top + 2)
        self.create_rectangle(0, top, w, bottom, fill="#1B2836", outline="")
        for n in self._n_starts:
            y = (n / total) * h
            self.create_line(2, y, w - 2, y, fill="#3A6478", width=1)
        for line_no, sev in self._markers:
            y = (line_no / total) * h
            if sev == "err":
                self.create_rectangle(2, y - 1.5, w - 2, y + 1.5, fill=ALERT_COLOR, outline="")
            else:
                self.create_rectangle(2, y - 1, w - 2, y + 1, fill=WARN_COLOR, outline="")

    def _on_click(self, event) -> None:
        h = self.winfo_height()
        if h <= 1 or not self._markers:
            return
        ratio = max(0.0, min(1.0, event.y / h))
        target = ratio * self._total_lines
        threshold = max(10, self._total_lines / 15)
        best_line = None
        best_dist = None
        for line_no, _sev in self._markers:
            dist = abs(line_no - target)
            if best_dist is None or dist < best_dist:
                best_dist = dist
                best_line = line_no
        if best_line is not None and best_dist is not None and best_dist <= threshold:
            self._on_jump(best_line)


class NcCheckApp:
    def __init__(self, root: tk.Tk, splash_mode: bool = False, initial_file: str | None = None) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.configure(bg=BG_APP)

        self.loaded_path: str | None = None
        self.status_var = tk.StringVar(value="NCプログラムを貼り付けるか、ファイルを開いてください。")
        self.path_var = tk.StringVar(value="未読込")
        self._search_bar: tk.Frame | None = None
        self._search_var = tk.StringVar()
        self._search_matches: list[str] = []
        self._search_idx: int = -1

        # G/M記号一覧（全件）。JSON管理で自由に追加・編集・削除できる。
        self._gm_codes: dict[str, str] = load_gm_glossary()
        # マクロ入力支援：JSON管理。初回起動時はサンプル(O7021)を1つ用意しておく
        load_nc_macro_defs()

        # コマンドライン引数経由の初期読み込みファイル（Windows「プログラムから開く」用）
        self._pending_initial_file = initial_file

        if splash_mode:
            self.root.withdraw()
            self._splash = SplashScreen(self.root)
            self._splash.set_progress(0.05, "起動中...")
            self.root.update()
            self.root.after(100, self._build_with_splash)
        else:
            self._build_ui()
            self.root.state("zoomed")
            self.root.bind("<Control-f>", self._focus_search)
            self.root.bind("<Control-F>", self._focus_search)
            self.root.bind("<Control-h>", self._open_replace_dialog)
            self.root.bind("<Control-H>", self._open_replace_dialog)
            # Undoボタンの初期状態
            self.root.after(50, self._update_undo_button_state)
            # コマンドライン引数のファイルがあれば読み込む
            if self._pending_initial_file:
                self.root.after(100, self._load_pending_initial_file)

    def _build_with_splash(self) -> None:
        sp = self._splash
        stages = [
            (0.10, "システム初期化..."),
            (0.20, "カラーパレット読込..."),
            (0.30, "検出エンジン初期化..."),
            (0.40, "正規表現パターン構築..."),
            (0.50, "シンタックスハイライト準備..."),
        ]
        self._splash_stages_post = [
            (0.65, "UIフレーム構築..."),
            (0.75, "チェックエンジン接続..."),
            (0.85, "Excel出力モジュール読込..."),
            (0.92, "検索エンジン初期化..."),
            (1.00, "起動完了"),
        ]
        self._splash_pre_idx = 0
        self._splash_pre_stages = stages
        self._run_pre_stage()

    def _run_pre_stage(self) -> None:
        sp = self._splash
        stages = self._splash_pre_stages
        i = self._splash_pre_idx
        if i < len(stages):
            val, txt = stages[i]
            sp.set_progress(val, txt)
            self.root.update()
            self._splash_pre_idx += 1
            self.root.after(250, self._run_pre_stage)
        else:
            sp.set_progress(0.55, "UIコンポーネント生成...")
            self.root.update()
            self._build_ui()
            self._splash_post_idx = 0
            self.root.after(150, self._run_post_stage)

    def _run_post_stage(self) -> None:
        sp = self._splash
        stages = self._splash_stages_post
        i = self._splash_post_idx
        if i < len(stages):
            val, txt = stages[i]
            sp.set_progress(val, txt)
            self.root.update()
            self._splash_post_idx += 1
            self.root.after(200, self._run_post_stage)
        else:
            self.root.after(500, self._finish_splash)

    def _finish_splash(self) -> None:
        self._splash.destroy()
        self.root.deiconify()
        self.root.state("zoomed")
        self.root.bind("<Control-f>", self._focus_search)
        self.root.bind("<Control-F>", self._focus_search)
        self.root.bind("<Control-h>", self._open_replace_dialog)
        self.root.bind("<Control-H>", self._open_replace_dialog)
        # Undoボタンの初期状態
        self.root.after(50, self._update_undo_button_state)
        # コマンドライン引数のファイルがあれば読み込む（メイン画面表示完了後）
        if self._pending_initial_file:
            self.root.after(200, self._load_pending_initial_file)

    def _load_pending_initial_file(self) -> None:
        """コマンドライン引数で渡されたファイルを読み込む（1回限り）"""
        path = self._pending_initial_file
        self._pending_initial_file = None
        if path:
            self._load_file_from_path(path)

    def _make_action_button(self, master: tk.Misc, text: str, command, primary: bool = False) -> tk.Button:
        bg = ACCENT if primary else INPUT_BG
        fg = "#04231F" if primary else TEXT_MAIN
        active_bg = ACCENT_DARK if primary else "#1F2B3A"
        return tk.Button(
            master,
            text=text,
            command=command,
            font=("Yu Gothic UI", 9, "bold" if primary else "normal"),
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=13,
            pady=8,
            cursor="hand2",
        )

    def _make_secbar(self, parent: tk.Frame, en: str, ja: str) -> tk.Label:
        """各ペイン共通の見出し帯（secbar）。モックの .secbar に合わせたネイビーバー。
        右端の可変テキスト（pill/ヒント）用ラベルを返す。
        """
        bar = tk.Frame(parent, bg=BG_HEADER, padx=12, pady=8)
        bar.pack(fill="x")
        tk.Label(bar, text=en, bg=BG_HEADER, fg=BAR_EN,
                 font=("Consolas", 7)).pack(side="left")
        tk.Label(bar, text=ja, bg=BG_HEADER, fg=BAR_INK,
                 font=("Yu Gothic UI", 8, "bold")).pack(side="left", padx=(6, 0))
        sp_label = tk.Label(bar, text="", bg=BG_HEADER, fg=BAR_EN, font=("Consolas", 7))
        sp_label.pack(side="right")
        return sp_label

    def _build_ui(self) -> None:
        self.root.configure(bg=BG_APP)

        # ===== タイトルバー =====
        titlebar = tk.Frame(self.root, bg=BG_HEADER, padx=14, pady=8)
        titlebar.pack(fill="x")

        mark = tk.Frame(titlebar, bg=BG_HEADER, width=20, height=20,
                        highlightthickness=1, highlightbackground=ACCENT, highlightcolor=ACCENT)
        mark.pack_propagate(False)
        mark.pack(side="left", padx=(0, 12))
        tk.Label(mark, text="NC", bg=BG_HEADER, fg=ACCENT, font=("Consolas", 7, "bold")).pack(expand=True)

        tk.Label(titlebar, text="NCプログラム確認ツール", bg=BG_HEADER, fg=FG_HEADER,
                 font=("Yu Gothic UI", 10, "bold")).pack(side="left")
        tk.Label(titlebar, text="v34", bg=BG_HEADER, fg=FG_SUB,
                 font=("Yu Gothic UI", 8)).pack(side="left", padx=(8, 0))

        tk.Label(titlebar, textvariable=self.path_var, bg="#143856", fg=FG_SUB,
                 font=("Consolas", 8), anchor="w", padx=10, pady=4,
                 highlightthickness=1, highlightbackground="#2A5578", highlightcolor="#2A5578",
                 ).pack(side="right")

        # ===== ツールバー =====
        toolbar = tk.Frame(self.root, bg=BG_PANEL, padx=14, pady=8)
        toolbar.pack(fill="x")
        self._make_action_button(toolbar, "ファイルを開く", self.open_file).pack(side="left", padx=(0, 7))
        self._make_action_button(toolbar, "チェック実行", self.run_check, primary=True).pack(side="left", padx=(0, 7))
        self._make_action_button(toolbar, "保存", self.save_nc_file).pack(side="left", padx=(0, 7))
        self._make_action_button(toolbar, "初品資料出力", self.export_excel).pack(side="left", padx=(0, 7))
        tk.Frame(toolbar, bg="#2A3846", width=1, height=22).pack(side="left", padx=5)
        self._make_action_button(toolbar, "置換 (Ctrl+H)", self._open_replace_dialog).pack(side="left", padx=(0, 7))
        self._make_action_button(toolbar, "マクロ呼出し支援", self._open_macro_assist_dialog).pack(side="left", padx=(0, 7))
        self._undo_button = self._make_action_button(toolbar, "↶ 元に戻す", self._do_undo)
        self._undo_button.pack(side="left", padx=(0, 7))
        self._make_action_button(toolbar, "入力をクリア", self.clear_input).pack(side="left", padx=(0, 7))

        # フォントサイズ変更
        self._font_size = 11
        font_frame = tk.Frame(toolbar, bg=BG_PANEL)
        font_frame.pack(side="right")
        btn_kw = {"font": ("Consolas", 9), "bg": INPUT_BG, "fg": TEXT_MUTED,
                  "activebackground": "#1F2B3A", "activeforeground": TEXT_MAIN,
                  "relief": "solid", "bd": 1, "highlightthickness": 0,
                  "width": 2, "height": 1, "cursor": "hand2"}
        tk.Button(font_frame, text="A-", command=self._font_decrease, **btn_kw).pack(side="left", padx=(0, 5))
        self._font_size_label = tk.Label(font_frame, text="11pt", bg=BG_PANEL, fg=TEXT_MUTED,
                                         font=("Consolas", 8), width=4, anchor="center")
        self._font_size_label.pack(side="left", padx=(0, 5))
        tk.Button(font_frame, text="A+", command=self._font_increase, **btn_kw).pack(side="left")

        # ===== 3ペイン =====
        # ペイン同士が独立したウィンドウに見えるよう、sashを最小限の「黒い隙間」として使う
        body = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=6, sashrelief=tk.FLAT,
                              bg=BG_APP, bd=0)
        body.pack(fill="both", expand=True, padx=6, pady=6)

        block_info_panel = tk.Frame(body, bg=BG_PANEL, bd=0, highlightthickness=0)
        n_index_panel = tk.Frame(body, bg=BG_PANEL, bd=0, highlightthickness=0)
        input_panel = tk.Frame(body, bg=BG_PANEL, bd=0, highlightthickness=0)
        result_panel = tk.Frame(body, bg=BG_PANEL, bd=0, highlightthickness=0)

        # 画面幅から初期幅を計算（N目次・BLOCK INFOはやや狭く、原文は広め、結果はやや狭く）
        try:
            sw = self.root.winfo_screenwidth()
        except tk.TclError:
            sw = 1920
        unit = max(200, int((sw - 40) / 5))  # 5等分の1ユニット
        # 表示順: N番号目次 → BLOCK INFO → NCプログラム → チェック結果
        body.add(n_index_panel, stretch="always", width=int(unit * 0.8), minsize=150)
        body.add(block_info_panel, stretch="always", width=int(unit * 0.85), minsize=210)
        body.add(input_panel, stretch="always", width=int(unit * 2.0), minsize=300)
        body.add(result_panel, stretch="always", width=int(unit * 1.4), minsize=280)

        self._build_n_index_panel(n_index_panel)
        self._build_block_info_panel(block_info_panel)
        self._build_input_panel(input_panel)
        self._build_result_panel(result_panel)

        # ===== 検索バー =====
        self._search_bar = tk.Frame(self.root, bg=BG_PANEL, padx=14, pady=8,
                                    highlightthickness=1, highlightbackground=BORDER, bd=0)
        self._search_bar.pack(fill="x")

        tk.Label(self._search_bar, text="検索", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).pack(side="left", padx=(0, 8))

        self._search_entry = tk.Entry(
            self._search_bar, textvariable=self._search_var,
            font=("Consolas", 9), bg=INPUT_BG, fg=TEXT_MAIN,
            insertbackground=TEXT_MAIN, relief="solid", bd=1,
            highlightthickness=1, highlightbackground="#2A3846", highlightcolor=ACCENT,
            width=34,
        )
        self._search_entry.pack(side="left", padx=(0, 6))
        self._search_entry.bind("<Return>", self._search_next)
        self._search_entry.bind("<Shift-Return>", self._search_prev)
        self._search_entry.bind("<KeyRelease>", self._on_search_changed)

        self._search_count_label = tk.Label(self._search_bar, text="", bg=BG_PANEL, fg=TEXT_MUTED,
                                            font=("Yu Gothic UI", 8))

        btn_style = {"font": ("Yu Gothic UI", 8), "bg": INPUT_BG, "fg": TEXT_MAIN,
                     "activebackground": "#1F2B3A", "activeforeground": TEXT_MAIN,
                     "relief": "solid", "bd": 1, "highlightthickness": 0,
                     "padx": 10, "pady": 4, "cursor": "hand2"}

        tk.Button(self._search_bar, text="次へ", command=self._search_next, **btn_style).pack(side="left", padx=(0, 4))
        tk.Button(self._search_bar, text="前へ", command=self._search_prev, **btn_style).pack(side="left", padx=(0, 4))
        self._search_count_label.pack(side="left", padx=(6, 0))

        self.input_text.tag_configure("search_highlight", background="#1B3A36", foreground="#A7F3D0")
        self.input_text.tag_configure("search_current", background="#1F8F86", foreground="#101720")

        # ===== ステータスバー =====
        status_bar = tk.Frame(self.root, bg=INPUT_BG, padx=14, pady=6,
                              highlightthickness=1, highlightbackground=BORDER, bd=0)
        status_bar.pack(fill="x")
        tk.Label(status_bar, textvariable=self.status_var, bg=INPUT_BG, fg=TEXT_MUTED,
                 anchor="w", justify="left", font=("Consolas", 8)).pack(fill="x")

    def _build_input_panel(self, parent: tk.Frame) -> None:
        self._source_secbar_sp = self._make_secbar(parent, "SOURCE", "NCプログラム")

        editor_frame = tk.Frame(parent, bg=BG_PANEL, padx=0, pady=0)
        editor_frame.pack(fill="both", expand=True)

        # 行番号ガター + 本文 + スクロールバー を横並び配置（プログラム入力タブは黒背景）
        editor_inner = tk.Frame(editor_frame, bg=INPUT_BG, relief="solid", bd=1,
                                highlightthickness=1, highlightbackground=BORDER,
                                highlightcolor=ACCENT)
        editor_inner.pack(fill="both", expand=True)

        # 行番号エリアのコンテナ（ガター本体 + 右端アクセントライン）
        gutter_frame = tk.Frame(editor_inner, bg=INPUT_BG)
        gutter_frame.pack(side="left", fill="y")

        # 行番号Text（表示専用・編集不可・選択不可）
        self.linenumber_text = tk.Text(
            gutter_frame,
            width=5,
            padx=8,
            pady=12,
            takefocus=0,
            font=("Consolas", 11),
            bg=INPUT_BG,
            fg="#3A6478",  # マトリクス残像グリーン
            relief="flat",
            bd=0,
            highlightthickness=0,
            wrap=tk.NONE,
            cursor="arrow",
            state="disabled",
            selectbackground=INPUT_BG,  # 選択時も色変えない（選択させない演出）
            selectforeground="#3A6478",
            inactiveselectbackground=INPUT_BG,
        )
        self.linenumber_text.pack(side="left", fill="y")

        # 行番号タグ：右寄せ、現在行ハイライト、エラー行ハイライト
        self.linenumber_text.tag_configure("ln_right", justify="right")
        self.linenumber_text.tag_configure(
            "ln_n_block",
            foreground="#F0D060",     # HUDアンバー（N番号のシンタックス色と合わせる）
            background="#2A2008",     # アンバー背景（黒上で浮くように）
            font=("Consolas", 11, "bold"),
            justify="right",
        )
        self.linenumber_text.tag_configure(
            "ln_current",
            foreground=ACCENT,        # 毒々しい蛍光グリーン
            background="#123A38",     # 黒上でも識別できる緑系
            justify="right",
        )
        self.linenumber_text.tag_configure(
            "ln_warn",
            foreground=WARN_COLOR,
            background="#2A2008",
            justify="right",
        )
        self.linenumber_text.tag_configure(
            "ln_error",
            foreground="#FB5E7E",     # 警告レッド
            background="#3A1620",     # 黒上で浮く血だまり
            justify="right",
        )
        # 優先度：error > warn > current > n_block > right
        self.linenumber_text.tag_raise("ln_n_block", "ln_right")
        self.linenumber_text.tag_raise("ln_current", "ln_n_block")
        self.linenumber_text.tag_raise("ln_warn", "ln_current")
        self.linenumber_text.tag_raise("ln_error", "ln_warn")

        # 右端の縦アクセントライン（ガターと本文の境界）
        gutter_divider = tk.Frame(editor_inner, bg=ACCENT_DARK, width=1)
        gutter_divider.pack(side="left", fill="y")

        # ミニマップ（995行規模でも危険/警告分布を一望できる縦帯。クリックで近い検出行へ）
        self._minimap = MiniMap(editor_inner, on_jump=self.jump_to_input_line)
        self._minimap.pack(side="right", fill="y")

        # 縦スクロールバー（本文用）
        self._input_vbar = tk.Scrollbar(editor_inner, orient="vertical",
                                        bg=BG_PANEL, troughcolor="#101720",
                                        activebackground=ACCENT)
        self._input_vbar.pack(side="right", fill="y")

        # 本文Text
        self.input_text = tk.Text(
            editor_inner,
            wrap=tk.NONE,
            undo=True,
            font=("Consolas", 11),
            bg=INPUT_BG,
            fg=TEXT_MAIN,
            insertbackground=ACCENT,   # カーソルも蛍光グリーンに
            relief="flat",
            bd=0,
            highlightthickness=0,
            padx=12,
            pady=12,
            selectbackground="#1F5A54",        # 選択色（明るい緑、現在行ハイライトと差をつける）
            selectforeground="#FFFFFF",         # 選択文字を白で強調
            inactiveselectbackground="#1F5A54",
        )
        self.input_text.pack(side="left", fill="both", expand=True)

        # 縦スクロール同期：scrollbarのsetは本文→scrollbarの通知、commandはユーザ操作→両方へ
        def _on_textscroll(*args):
            # input_text の yview 変化 → scrollbar 更新 & 行番号側追従 & ミニマップの表示窓
            self._input_vbar.set(*args)
            self.linenumber_text.yview_moveto(args[0])
            self._minimap.set_viewport(*args)

        def _on_scrollbar(*args):
            # scrollbar操作 → 本文と行番号の両方を動かす
            self.input_text.yview(*args)
            self.linenumber_text.yview(*args)

        self.input_text.configure(yscrollcommand=_on_textscroll)
        self._input_vbar.configure(command=_on_scrollbar)

        # マウスホイールで行番号側スクロール時も本文連動
        def _on_wheel_linenumber(event):
            # Windows: delta / 120 単位。他プラットフォームでも同じ扱いで概ねOK
            if event.delta:
                units = int(-event.delta / 120) * 3
            else:
                units = -3 if getattr(event, "num", 0) == 4 else 3
            self.input_text.yview_scroll(units, "units")
            return "break"
        self.linenumber_text.bind("<MouseWheel>", _on_wheel_linenumber)
        self.linenumber_text.bind("<Button-4>", _on_wheel_linenumber)
        self.linenumber_text.bind("<Button-5>", _on_wheel_linenumber)

        self.input_text.tag_configure("current_line_bg", background="#141B24")
        self.input_text.tag_configure("warn_line_highlight", background="#2A2008")
        self.input_text.tag_configure("error_line_highlight", background=ERROR_LINE_BG)
        self.input_text.tag_configure("jump_highlight", background=JUMP_LINE_BG)

        # シンタックスハイライトタグ登録
        for tag_name, _, color in SYNTAX_COLORS:
            self.input_text.tag_configure(tag_name, foreground=color)
        # レイヤ優先度：current_line_bg（最下） < syntax < error < jump < sel
        self.input_text.tag_raise("error_line_highlight")
        self.input_text.tag_raise("jump_highlight")
        # selタグ（マウス選択）を最上位に持ち上げて、行ハイライトに埋もれないように
        try:
            self.input_text.tag_raise("sel")
        except tk.TclError:
            pass

        # キー入力でカーソル行をリアルタイムハイライト + 行番号更新
        self.input_text.bind("<KeyRelease>", self._on_input_key_release)
        self.input_text.bind("<<Modified>>", self._on_input_modified)
        # カーソル移動（クリック・矢印キー・Homeなど）で行番号の現在行ハイライトだけ軽量更新
        self.input_text.bind("<ButtonRelease-1>", self._on_input_cursor_move, add="+")
        self.input_text.bind("<<Selection>>", self._on_input_cursor_move, add="+")
        self._hl_last_line: int = -1
        # 初回行番号描画
        self._update_linenumbers()

    def _update_linenumbers(self) -> None:
        """input_textの行数に合わせて行番号ガターを再生成"""
        if not hasattr(self, "linenumber_text"):
            return
        current_yview = self.input_text.yview()
        total = self.input_text.index("end-1c")
        try:
            last_line = int(str(total).split(".")[0])
        except (ValueError, IndexError):
            last_line = 1
        if last_line < 1:
            last_line = 1
        width = max(4, len(str(last_line)) + 2)
        text = "\n".join(str(i) for i in range(1, last_line + 1))
        self.linenumber_text.configure(state="normal")
        self.linenumber_text.configure(width=width)
        self.linenumber_text.delete("1.0", tk.END)
        self.linenumber_text.insert("1.0", text)
        self.linenumber_text.tag_add("ln_right", "1.0", tk.END)
        self.linenumber_text.configure(state="disabled")
        self.linenumber_text.yview_moveto(current_yview[0])
        # エラー行と現在行の強調を再適用
        self._apply_linenumber_highlights()
        # Nブロック強調も再適用（行番号再生成でタグが消えるため）
        self._apply_n_block_dividers()
        self._update_source_secbar(last_line)

    def _update_source_secbar(self, total_lines: int) -> None:
        if not hasattr(self, "_source_secbar_sp"):
            return
        try:
            content = self.input_text.get("1.0", "end-1c")
        except tk.TclError:
            content = ""
        if not content.strip():
            self._source_secbar_sp.configure(text="")
            return
        program_number, _ = extract_program_number(content.splitlines())
        self._source_secbar_sp.configure(text=f"{program_number} / {total_lines} 行")

    def _apply_linenumber_highlights(self) -> None:
        """行番号ガター & 本文カーソル行の強調を再適用"""
        if not hasattr(self, "linenumber_text"):
            return
        ln = self.linenumber_text
        ln.configure(state="normal")
        ln.tag_remove("ln_current", "1.0", tk.END)
        ln.tag_remove("ln_error", "1.0", tk.END)
        ln.tag_remove("ln_warn", "1.0", tk.END)

        def _lines_from_tag(tag_name: str) -> set[int]:
            try:
                tag_ranges = self.input_text.tag_ranges(tag_name)
            except tk.TclError:
                tag_ranges = []
            result: set[int] = set()
            for i in range(0, len(tag_ranges), 2):
                try:
                    start_line = int(str(tag_ranges[i]).split(".")[0])
                    end_line = int(str(tag_ranges[i + 1]).split(".")[0])
                    for ln_no in range(start_line, end_line + 1):
                        result.add(ln_no)
                except (ValueError, IndexError):
                    pass
            return result

        # エラー行・警告行の集合を取得（エラー優先）
        error_lines = _lines_from_tag("error_line_highlight")
        warn_lines = _lines_from_tag("warn_line_highlight") - error_lines
        for ln_no in error_lines:
            try:
                ln.tag_add("ln_error", f"{ln_no}.0", f"{ln_no}.end")
            except tk.TclError:
                pass
        for ln_no in warn_lines:
            try:
                ln.tag_add("ln_warn", f"{ln_no}.0", f"{ln_no}.end")
            except tk.TclError:
                pass

        # 現在カーソル行
        cur_line = None
        try:
            cur_idx = self.input_text.index(tk.INSERT)
            cur_line = int(str(cur_idx).split(".")[0])
        except (tk.TclError, ValueError, IndexError):
            pass

        if cur_line is not None and cur_line not in error_lines and cur_line not in warn_lines:
            try:
                ln.tag_add("ln_current", f"{cur_line}.0", f"{cur_line}.end")
            except tk.TclError:
                pass

        ln.configure(state="disabled")

        # 本文側：カーソル行の背景も光らせる（IDE風）
        try:
            self.input_text.tag_remove("current_line_bg", "1.0", tk.END)
            if cur_line is not None:
                self.input_text.tag_add("current_line_bg", f"{cur_line}.0", f"{cur_line}.end + 1c")
        except tk.TclError:
            pass

    def _on_input_key_release(self, event) -> None:
        self._on_key_release(event)
        self._update_linenumbers()
        self._update_block_info_for_current_line()
        self._update_gm_legend_for_current_line()
        # 編集debounce：タイピングが止まってから400ms後にチェック全体再実行
        # （BLOCK INFOキャッシュ・警告類すべて更新される）
        self._schedule_edit_recheck()

    def _schedule_edit_recheck(self) -> None:
        """編集後にdebounceでチェック全体を自動再実行"""
        # 自動更新OFFなら何もしない
        if not getattr(self, "_auto_refresh_enabled", True):
            return
        if not hasattr(self, "_edit_debounce_id"):
            self._edit_debounce_id = None
        if self._edit_debounce_id is not None:
            try:
                self.root.after_cancel(self._edit_debounce_id)
            except Exception:
                pass
        self._edit_debounce_id = self.root.after(1000, self._auto_rerun_check_after_edit)

    def _auto_rerun_check_after_edit(self) -> None:
        """編集debounceからの呼び出し用。
        プログラムが空 or レポートが未表示の状態では何もしない（初回はユーザー操作待ち）"""
        self._edit_debounce_id = None
        program_text = self.input_text.get("1.0", tk.END).rstrip("\n")
        if not program_text.strip():
            return
        # まだ一度もチェック実行されてない場合はスキップ
        # （ユーザーが「チェック実行」を最初に押すまでは自動更新しない方針）
        try:
            existing_report = self.output_text.get("1.0", "end-1c").strip()
        except tk.TclError:
            return
        if not existing_report:
            return
        # 既存の_auto_rerun_checkを使う（スキャンアニメなし）
        self._auto_rerun_check()

    def _on_input_modified(self, event) -> None:
        try:
            if self.input_text.edit_modified():
                self._update_linenumbers()
                self._update_undo_button_state()
                self.input_text.edit_modified(False)
        except tk.TclError:
            pass

    def _on_input_cursor_move(self, event=None) -> None:
        """カーソル移動のみ → 現在行ハイライト + N情報パネル + M/G凡例を軽量更新"""
        self._apply_linenumber_highlights()
        self._update_block_info_for_current_line()
        self._update_gm_legend_for_current_line()

    def _build_result_panel(self, parent: tk.Frame) -> None:
        sp = self._make_secbar(parent, "RESULT", "チェック結果")
        sp.configure(text="クリックで原文へジャンプ")

        # ===== 上下分割：上＝チェック結果、下＝カーソル行のM/G記号凡例 =====
        vsplit = tk.PanedWindow(parent, orient=tk.VERTICAL, sashwidth=6, sashrelief=tk.FLAT,
                                bg=BG_APP, bd=0)
        vsplit.pack(fill="both", expand=True)

        top = tk.Frame(vsplit, bg=BG_PANEL, bd=0, highlightthickness=0)
        bottom = tk.Frame(vsplit, bg=BG_PANEL, bd=0, highlightthickness=0)
        vsplit.add(top, stretch="always", height=430, minsize=160)
        vsplit.add(bottom, stretch="always", height=170, minsize=110)

        self._build_gm_legend_panel(bottom)

        parent = top  # 以降の要素はすべて上段（チェック結果）に配置

        # ===== 判定は細い帯1本（デカい枠は置かない） =====
        self._verdict_band = VerdictBand(parent)
        self._verdict_band.pack(fill="x")

        # ===== ゼロ件は畳んで1行に =====
        self._zeros_text = tk.Text(
            parent, height=2, wrap="word", bd=0, highlightthickness=0,
            bg="#0F151C", fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
            padx=13, pady=8, cursor="arrow", state="disabled",
        )
        self._zeros_text.tag_configure("chk", foreground=OK_COLOR, font=("Consolas", 9, "bold"))

        # ===== 判定上限：折りたたんで最下部に格納 =====
        default = ThresholdSettings()
        self._threshold_debounce_id: str | None = None

        limits_wrap = tk.Frame(parent, bg=INPUT_BG, highlightthickness=1, highlightbackground=BORDER)
        limits_wrap.pack(fill="x", side="bottom")

        self._limits_open = False
        limits_header = tk.Frame(limits_wrap, bg=INPUT_BG, cursor="hand2")
        limits_header.pack(fill="x")
        self._limits_arrow = tk.Label(limits_header, text="▶", bg=INPUT_BG, fg=TEXT_MUTED,
                                      font=("Consolas", 9), cursor="hand2")
        self._limits_arrow.pack(side="left", padx=(13, 6), pady=8)
        tk.Label(limits_header, text="判定上限の設定", bg=INPUT_BG, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9, "bold"), cursor="hand2").pack(side="left")
        self._limits_summary = tk.Label(
            limits_header,
            text=f"S{int(default.speed_limit)} / F{int(default.feed_limit)}",
            bg=INPUT_BG, fg=TEXT_MUTED, font=("Consolas", 9), cursor="hand2",
        )
        self._limits_summary.pack(side="right", padx=(0, 13))
        for w in (limits_header, *limits_header.winfo_children()):
            w.bind("<Button-1>", lambda _e: self._toggle_limits())

        self._limits_body = tk.Frame(limits_wrap, bg=INPUT_BG, padx=13, pady=12)
        threshold_grid = tk.Frame(self._limits_body, bg=INPUT_BG)
        threshold_grid.pack(fill="x")
        threshold_grid.grid_columnconfigure(0, weight=1)
        threshold_grid.grid_columnconfigure(1, weight=1)
        self.speed_entry = ThresholdEntry(threshold_grid, "回転数 判定上限 (S)",
                                          default.speed_limit, self._on_threshold_changed)
        self.feed_entry = ThresholdEntry(threshold_grid, "送り 判定上限 (F)",
                                         default.feed_limit, self._on_threshold_changed)
        self.speed_entry.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.feed_entry.grid(row=0, column=1, sticky="nsew")
        # 初期状態は閉じる（bodyは意図的にpackしない）

        # ===== 目次本体（スクロール可能・検出ありのグループだけ開いて表示） =====
        self._toc_wrap = tk.Frame(parent, bg=BG_PANEL)
        toc_wrap = self._toc_wrap
        toc_wrap.pack(fill="both", expand=True)
        toc_canvas = tk.Canvas(toc_wrap, bg=BG_PANEL, highlightthickness=0, bd=0)
        toc_canvas.pack(side="left", fill="both", expand=True)
        toc_vbar = tk.Scrollbar(toc_wrap, orient="vertical", bg=BG_PANEL,
                                troughcolor="#101720", activebackground=ACCENT,
                                command=toc_canvas.yview)
        toc_vbar.pack(side="right", fill="y")
        toc_canvas.configure(yscrollcommand=toc_vbar.set)

        self._toc_inner = tk.Frame(toc_canvas, bg=BG_PANEL)
        toc_window = toc_canvas.create_window((0, 0), window=self._toc_inner, anchor="nw")

        def _on_toc_inner_configure(_event=None):
            toc_canvas.configure(scrollregion=toc_canvas.bbox("all"))
        self._toc_inner.bind("<Configure>", _on_toc_inner_configure)

        def _on_toc_canvas_configure(event):
            toc_canvas.itemconfigure(toc_window, width=event.width)
        toc_canvas.bind("<Configure>", _on_toc_canvas_configure)

        def _on_toc_wheel(event):
            if event.delta:
                toc_canvas.yview_scroll(int(-event.delta / 120), "units")
            else:
                toc_canvas.yview_scroll(-1 if getattr(event, "num", 0) == 4 else 1, "units")
        toc_canvas.bind("<MouseWheel>", _on_toc_wheel)
        toc_canvas.bind("<Button-4>", _on_toc_wheel)
        toc_canvas.bind("<Button-5>", _on_toc_wheel)

        self._toc_group_open: dict[str, bool] = {}
        self._toc_selected_line: int | None = None
        self._toc_item_widgets: dict[int, list[tk.Widget]] = {}
        self._last_toc_groups: list[dict] | None = None
        self._render_toc_placeholder("チェック実行後に表示されます")

        # ===== 非表示の内部レポートバッファ =====
        # 旧テキストレポート生成・行エラーハイライト判定・「編集後の初回チェック済み」
        # 判定など既存ロジックが参照するため保持するが、画面には表示しない。
        self.output_text = ScrolledText(
            parent, wrap=tk.NONE, undo=False, font=("Consolas", 11),
            bg=RESULT_BG, fg=TEXT_MAIN,
        )
        self.output_text.tag_configure("heading", font=("Consolas", 11, "bold"), foreground="#8FE8B8")
        self.output_text.tag_configure("jump_link", foreground="#8FF0C7", underline=True)
        for style_name, style_kwargs in RESULT_STYLE_MAP.items():
            self.output_text.tag_configure(style_name, **style_kwargs)

    def _build_gm_legend_panel(self, parent: tk.Frame) -> None:
        """チェック結果ペイン下段：カーソル行にあるM/Gコードの意味をリアルタイム表示する"""
        self._gm_legend_sp = self._make_secbar(parent, "LEGEND", "M/G記号")
        # 記号一覧はユーザー側でも追加できるよう、secbarに追加ボタンを差し込む
        add_btn = tk.Label(self._gm_legend_sp.master, text="＋ 追加", bg=BG_HEADER, fg=BAR_INK,
                           font=("Yu Gothic UI", 8, "bold"), cursor="hand2")
        add_btn.pack(side="right", padx=(0, 10))
        add_btn.bind("<Button-1>", lambda _e: self._open_gm_glossary_dialog())

        legend_wrap = tk.Frame(parent, bg=BG_PANEL)
        legend_wrap.pack(fill="both", expand=True)

        canvas = tk.Canvas(legend_wrap, bg=BG_PANEL, highlightthickness=0, bd=0)
        canvas.pack(side="left", fill="both", expand=True)
        vbar = tk.Scrollbar(legend_wrap, orient="vertical", bg=BG_PANEL,
                            troughcolor="#101720", activebackground=ACCENT,
                            command=canvas.yview)
        vbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=vbar.set)
        self._gm_legend_canvas = canvas

        self._gm_legend_inner = tk.Frame(canvas, bg=BG_PANEL, padx=13, pady=10)
        inner_window = canvas.create_window((0, 0), window=self._gm_legend_inner, anchor="nw")

        def _on_inner_configure(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        self._gm_legend_inner.bind("<Configure>", _on_inner_configure)

        def _on_canvas_configure(event):
            canvas.itemconfigure(inner_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.bind("<MouseWheel>", self._on_gm_legend_wheel)
        canvas.bind("<Button-4>", self._on_gm_legend_wheel)
        canvas.bind("<Button-5>", self._on_gm_legend_wheel)

        self._render_gm_legend([], None)

    def _on_gm_legend_wheel(self, event) -> None:
        if not hasattr(self, "_gm_legend_canvas"):
            return
        if event.delta:
            self._gm_legend_canvas.yview_scroll(int(-event.delta / 120), "units")
        else:
            self._gm_legend_canvas.yview_scroll(-1 if getattr(event, "num", 0) == 4 else 1, "units")

    def _render_gm_legend(self, tokens: list[str], line_no: int | None) -> None:
        if not hasattr(self, "_gm_legend_inner"):
            return
        for w in self._gm_legend_inner.winfo_children():
            w.destroy()
        if hasattr(self, "_gm_legend_sp"):
            self._gm_legend_sp.configure(text=(f"{line_no}行目" if line_no else ""))

        if not tokens:
            tk.Label(self._gm_legend_inner, text="カーソル行にG/Mコードはありません",
                     bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                     wraplength=260, justify="left").pack(anchor="w", pady=6)
            return

        for token in tokens:
            meaning = self._gm_lookup(token)
            row = tk.Frame(self._gm_legend_inner, bg=BG_PANEL)
            row.pack(fill="x", pady=4)
            code_color = ACCENT if meaning else TEXT_MUTED
            tk.Label(row, text=token, bg=BG_PANEL, fg=code_color,
                     font=("Consolas", 13, "bold"), anchor="w", width=9).pack(side="left")
            tk.Label(row, text=meaning or "凡例未登録のコード", bg=BG_PANEL,
                     fg=(TEXT_MAIN if meaning else TEXT_MUTED),
                     font=("Yu Gothic UI", 10), anchor="w", justify="left",
                     wraplength=190).pack(side="left", fill="x", expand=True)

    def _update_gm_legend_for_current_line(self) -> None:
        if not hasattr(self, "_gm_legend_inner"):
            return
        try:
            cur_idx = self.input_text.index(tk.INSERT)
            cur_line = int(str(cur_idx).split(".")[0])
            line_text = self.input_text.get(f"{cur_line}.0", f"{cur_line}.end")
        except (tk.TclError, ValueError, IndexError):
            return
        tokens = extract_gm_tokens(line_text)
        self._render_gm_legend(tokens, cur_line)

    def _gm_lookup(self, token: str) -> str | None:
        """G/M記号一覧(self._gm_codes、JSON管理)から意味を引く"""
        return getattr(self, "_gm_codes", {}).get(token)

    def _open_gm_glossary_dialog(self) -> None:
        """G/M記号一覧の追加・編集ダイアログを開く（既に開いていれば前面に）"""
        existing = getattr(self, "_gm_glossary_dialog", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.lift()
                    existing.focus_set()
                    return
            except tk.TclError:
                pass
        self._gm_glossary_dialog = GmGlossaryDialog(self)

    def _on_gm_glossary_saved(self, entries: dict[str, str]) -> None:
        """GmGlossaryDialogからの保存コールバック：一覧を差し替えて即反映"""
        self._gm_codes = entries
        save_gm_glossary(entries)
        self._update_gm_legend_for_current_line()

    def _open_macro_assist_dialog(self) -> None:
        """マクロ呼出し支援ダイアログを開く（既に開いていれば前面に）"""
        existing = getattr(self, "_macro_assist_dialog", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.lift()
                    existing.focus_set()
                    return
            except tk.TclError:
                pass
        self._macro_assist_dialog = NcMacroAssistDialog(self)

    def insert_text_at_cursor(self, text: str) -> None:
        """カーソル位置（選択範囲があれば置き換え）にテキストを挿入する"""
        try:
            if self.input_text.tag_ranges("sel"):
                self.input_text.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        self.input_text.insert(tk.INSERT, text)
        self.input_text.see(tk.INSERT)
        self.input_text.focus_set()
        self._update_linenumbers()
        self._highlight_all()
        self._update_gm_legend_for_current_line()
        self._schedule_edit_recheck()

    def _toggle_limits(self) -> None:
        self._limits_open = not self._limits_open
        if self._limits_open:
            self._limits_arrow.configure(text="▼")
            self._limits_body.pack(fill="x")
        else:
            self._limits_arrow.configure(text="▶")
            self._limits_body.pack_forget()

    def _render_toc_placeholder(self, message: str) -> None:
        if not hasattr(self, "_toc_inner"):
            return
        for w in self._toc_inner.winfo_children():
            w.destroy()
        self._toc_item_widgets = {}
        tk.Label(self._toc_inner, text=message, bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 10), wraplength=340, justify="left").pack(
                     anchor="w", pady=20, padx=13)
        self._zeros_text.pack_forget()

    def _build_block_info_panel(self, parent: tk.Frame) -> None:
        """カーソル位置のNブロック情報をリアルタイム表示するパネル"""
        self._block_secbar_sp = self._make_secbar(parent, "BLOCK", "ブロック情報")

        # フッター（自動更新トグル）— content_wrapより先にpackして下端に固定
        footer = tk.Frame(parent, bg=BG_PANEL, padx=12, pady=9,
                          highlightthickness=1, highlightbackground=BORDER, bd=0)
        footer.pack(side="bottom", fill="x")

        self._auto_refresh_enabled = True  # デフォルトON
        self._auto_refresh_btn = tk.Frame(footer, bg=BG_PANEL, cursor="hand2")
        self._auto_refresh_btn.pack(anchor="w")
        self._auto_refresh_dot = tk.Canvas(self._auto_refresh_btn, width=7, height=7,
                                           bg=BG_PANEL, highlightthickness=0, bd=0)
        self._auto_refresh_dot_oval = self._auto_refresh_dot.create_oval(0, 0, 7, 7, fill=OK_COLOR, outline="")
        self._auto_refresh_dot.pack(side="left", padx=(0, 7))
        self._auto_refresh_label = tk.Label(self._auto_refresh_btn, text="自動更新 ON",
                                            bg=BG_PANEL, fg=OK_COLOR, font=("Yu Gothic UI", 8), cursor="hand2")
        self._auto_refresh_label.pack(side="left")
        for w in (self._auto_refresh_btn, self._auto_refresh_dot, self._auto_refresh_label):
            w.bind("<Button-1>", lambda _e: self._toggle_auto_refresh())
        self._update_auto_refresh_ui()  # 初期表示

        # 内容コンテナ（スクロール可能）— footerより後にpackすることで間を埋める
        content_wrap = tk.Frame(parent, bg=BG_PANEL, padx=14, pady=0)
        content_wrap.pack(fill="both", expand=True)

        # キャンバスベースのスクロールエリア
        canvas = tk.Canvas(content_wrap, bg=BG_PANEL, highlightthickness=0, bd=0)
        canvas.pack(side="left", fill="both", expand=True)
        vbar = tk.Scrollbar(content_wrap, orient="vertical",
                            bg=BG_PANEL, troughcolor="#101720",
                            activebackground=ACCENT, command=canvas.yview)
        vbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=vbar.set)

        self._block_info_inner = tk.Frame(canvas, bg=BG_PANEL)
        inner_window = canvas.create_window((0, 0), window=self._block_info_inner, anchor="nw")

        def _on_inner_configure(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        self._block_info_inner.bind("<Configure>", _on_inner_configure)

        def _on_canvas_configure(event):
            canvas.itemconfigure(inner_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        # マウスホイール対応
        def _on_wheel(event):
            if event.delta:
                canvas.yview_scroll(int(-event.delta / 120), "units")
            else:
                canvas.yview_scroll(-1 if getattr(event, "num", 0) == 4 else 1, "units")
        canvas.bind("<MouseWheel>", _on_wheel)
        canvas.bind("<Button-4>", _on_wheel)
        canvas.bind("<Button-5>", _on_wheel)

        # キャッシュとカレント追跡
        self._block_info_cache: dict = {}  # {n_label: dict of fields}
        self._block_info_ranges: list = []  # [(start_line, end_line, n_label), ...]
        self._current_displayed_label: str | None = None

        # 初期表示（プレースホルダー）
        self._show_block_info_placeholder()

    def _build_n_index_panel(self, parent: tk.Frame) -> None:
        """N番号ごとの目次。クリックでそのNブロックの先頭行へジャンプする。"""
        self._make_secbar(parent, "INDEX", "N番号目次")

        content_wrap = tk.Frame(parent, bg=BG_PANEL, padx=0, pady=0)
        content_wrap.pack(fill="both", expand=True)

        canvas = tk.Canvas(content_wrap, bg=BG_PANEL, highlightthickness=0, bd=0)
        canvas.pack(side="left", fill="both", expand=True)
        vbar = tk.Scrollbar(content_wrap, orient="vertical",
                            bg=BG_PANEL, troughcolor="#101720",
                            activebackground=ACCENT, command=canvas.yview)
        vbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=vbar.set)
        self._n_index_canvas = canvas

        self._n_index_inner = tk.Frame(canvas, bg=BG_PANEL)
        inner_window = canvas.create_window((0, 0), window=self._n_index_inner, anchor="nw")

        def _on_inner_configure(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        self._n_index_inner.bind("<Configure>", _on_inner_configure)

        def _on_canvas_configure(event):
            canvas.itemconfigure(inner_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.bind("<MouseWheel>", self._on_n_index_wheel)
        canvas.bind("<Button-4>", self._on_n_index_wheel)
        canvas.bind("<Button-5>", self._on_n_index_wheel)
        self._n_index_inner.bind("<MouseWheel>", self._on_n_index_wheel)
        self._n_index_inner.bind("<Button-4>", self._on_n_index_wheel)
        self._n_index_inner.bind("<Button-5>", self._on_n_index_wheel)

        self._n_index_item_widgets: dict[int, tuple] = {}
        self._render_n_index()

    def _on_n_index_wheel(self, event) -> None:
        if not hasattr(self, "_n_index_canvas"):
            return
        if event.delta:
            self._n_index_canvas.yview_scroll(int(-event.delta / 120), "units")
        else:
            self._n_index_canvas.yview_scroll(-1 if getattr(event, "num", 0) == 4 else 1, "units")

    def _render_n_index(self) -> None:
        """BLOCK INFOと同じキャッシュから、Nブロック一覧をクリック可能な行として並べる"""
        if not hasattr(self, "_n_index_inner"):
            return
        for w in self._n_index_inner.winfo_children():
            w.destroy()
        self._n_index_item_widgets = {}

        if not getattr(self, "_block_info_cache", None):
            tk.Label(self._n_index_inner, text="チェック実行後に表示されます",
                     bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 8),
                     wraplength=100, justify="left").pack(anchor="w", padx=10, pady=14)
            return

        for block_id in sorted(self._block_info_cache.keys()):
            info = self._block_info_cache[block_id]
            selected = (self._current_displayed_label == block_id)
            row_bg = "#171F29" if selected else BG_PANEL
            row = tk.Frame(self._n_index_inner, bg=row_bg, cursor="hand2")
            row.pack(fill="x")
            bar = tk.Frame(row, bg=(ACCENT if selected else BG_PANEL), width=4)
            bar.pack(side="left", fill="y")
            label_col = tk.Frame(row, bg=row_bg)
            label_col.pack(side="left", fill="x", expand=True, padx=(10, 6), pady=9)
            lbl = tk.Label(label_col, text=info["n_label"], bg=row_bg,
                           fg=(ACCENT if selected else TEXT_MAIN),
                           font=("Consolas", 15, "bold"), anchor="w")
            lbl.pack(anchor="w", fill="x")
            ln_lbl = tk.Label(label_col, text=info.get("tool_name", "") or "―", bg=row_bg,
                              fg=TEXT_MUTED, font=("Consolas", 10), anchor="w",
                              wraplength=190, justify="left")
            ln_lbl.pack(anchor="w", fill="x", pady=(2, 0))
            self._n_index_item_widgets[block_id] = (row, bar, label_col, lbl, ln_lbl)
            for w in (row, bar, label_col, lbl, ln_lbl):
                w.bind("<Button-1>", lambda _e, ln=info["line_min"]: self.jump_to_input_line(ln))
                w.bind("<MouseWheel>", self._on_n_index_wheel)
                w.bind("<Button-4>", self._on_n_index_wheel)
                w.bind("<Button-5>", self._on_n_index_wheel)

    def _refresh_n_index_selection(self) -> None:
        """カーソル位置のNブロックが変わった時、目次側のハイライトだけ軽量に更新"""
        if not hasattr(self, "_n_index_item_widgets"):
            return
        for block_id, widgets in self._n_index_item_widgets.items():
            row, bar, label_col, lbl, ln_lbl = widgets
            selected = (self._current_displayed_label == block_id)
            row_bg = "#171F29" if selected else BG_PANEL
            row.configure(bg=row_bg)
            bar.configure(bg=(ACCENT if selected else BG_PANEL))
            label_col.configure(bg=row_bg)
            lbl.configure(bg=row_bg, fg=(ACCENT if selected else TEXT_MAIN))
            ln_lbl.configure(bg=row_bg)

    def _toggle_auto_refresh(self) -> None:
        """自動更新トグル切替"""
        self._auto_refresh_enabled = not self._auto_refresh_enabled
        self._update_auto_refresh_ui()
        # OFFに切り替えた瞬間にdebounceタイマーが残ってたら止める
        if not self._auto_refresh_enabled:
            if hasattr(self, "_edit_debounce_id") and self._edit_debounce_id is not None:
                try:
                    self.root.after_cancel(self._edit_debounce_id)
                except Exception:
                    pass
                self._edit_debounce_id = None

    def _update_auto_refresh_ui(self) -> None:
        """自動更新ライブインジケータの見た目をON/OFF状態に合わせて更新"""
        color = OK_COLOR if self._auto_refresh_enabled else TEXT_MUTED
        text = "自動更新 ON" if self._auto_refresh_enabled else "自動更新 OFF"
        self._auto_refresh_dot.itemconfigure(self._auto_refresh_dot_oval, fill=color)
        self._auto_refresh_label.configure(text=text, fg=color)

    def _show_block_info_placeholder(self, message: str = "チェック実行後に表示されます") -> None:
        """プレースホルダー or 「N番号外」表示"""
        if not hasattr(self, "_block_info_inner"):
            return
        # 既存ウィジェットをクリア
        for w in self._block_info_inner.winfo_children():
            w.destroy()
        tk.Label(self._block_info_inner, text=message,
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 10), wraplength=250, justify="left").pack(
                     anchor="w", pady=14, padx=4)
        self._current_displayed_label = None
        if hasattr(self, "_block_secbar_sp"):
            self._block_secbar_sp.configure(text="")

    def _render_block_info(self, info: dict) -> None:
        """1ブロック分の情報をカード表示。モックの nblock + gauges 構成に合わせ、
        N番号を大表示 + T番号/H補正/送りをゲージ化し、残りはkey-value行で並べる。
        """
        if not hasattr(self, "_block_info_inner"):
            return
        for w in self._block_info_inner.winfo_children():
            w.destroy()

        # ===== Nブロック大表示（N番号 + 工具名 + 行範囲） =====
        nblock = tk.Frame(self._block_info_inner, bg=BG_PANEL)
        nblock.pack(fill="x", pady=(10, 0))
        tk.Label(nblock, text=info.get("n_label", ""),
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 30, "bold")).pack(anchor="center")
        tool_name = info.get("tool_name", "") or "(名称なし)"
        tk.Label(nblock, text=tool_name,
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Consolas", 11, "bold"),
                 wraplength=240, justify="center").pack(anchor="center", pady=(6, 0))
        line_min = info.get("line_min")
        line_max = info.get("line_max")
        if line_min and line_max:
            span = line_max - line_min + 1
            range_text = f"{line_min}行 – {line_max}行 / {span} lines"
            pill_text = f"L{line_min}–L{line_max}"
        else:
            range_text = "―"
            pill_text = ""
        tk.Label(nblock, text=range_text,
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Consolas", 9)).pack(anchor="center", pady=(3, 10))
        if hasattr(self, "_block_secbar_sp"):
            self._block_secbar_sp.configure(text=pill_text)

        # 危険件数バッジ（あれば）
        danger_count = int(info.get("danger_count", 0) or 0)
        if danger_count > 0:
            badge_color_bg = "#301A10"
            badge_color_fg = "#FF9557"
            if danger_count >= 3:
                badge_color_bg = "#402030"
                badge_color_fg = ALERT_COLOR
            badge_frame = tk.Frame(self._block_info_inner, bg=badge_color_bg,
                                    highlightthickness=1, highlightbackground=badge_color_fg)
            badge_frame.pack(pady=(0, 10), padx=0, fill="x")
            tk.Label(badge_frame,
                     text=f"⚠ 危険検出 {danger_count} 件",
                     bg=badge_color_bg, fg=badge_color_fg,
                     font=("Yu Gothic UI", 9, "bold"),
                     pady=4).pack()

        # ===== ゲージ：T番号/H補正は少し大きめに強調表示 =====
        gauges = tk.Frame(self._block_info_inner, bg=BG_PANEL,
                          highlightthickness=1, highlightbackground=BORDER)
        gauges.pack(fill="x", pady=(0, 4))
        gauge_defs = [
            ("T番号", info.get("t_number", "") or "―"),
            ("H補正", info.get("h_offsets", "") or "―"),
        ]
        for idx, (label, value) in enumerate(gauge_defs):
            cell = tk.Frame(gauges, bg=BG_PANEL)
            cell.grid(row=0, column=idx, sticky="nsew", padx=(1 if idx else 0, 0))
            gauges.grid_columnconfigure(idx, weight=1)
            tk.Label(cell, text=label, bg=BG_PANEL, fg=TEXT_MUTED,
                     font=("Yu Gothic UI", 8)).pack(pady=(9, 3))
            tk.Label(cell, text=value, bg=BG_PANEL, fg=TEXT_MAIN,
                     font=("Consolas", 15, "bold"), wraplength=90,
                     justify="center").pack(pady=(0, 9))

        # ===== 残り項目をkey-value形式で（送りは回転数の上に通常行として表示） =====
        rows = [
            ("径補正", info.get("radius_comp", "") or "―"),
            ("ワーク座標", info.get("work_coord", "") or "―"),
            ("送り", info.get("feed_list", "") or "―"),
            ("回転数", info.get("spindle", "") or "―"),
            ("回転方向", info.get("rotation", "") or "―"),
            ("クーラント", info.get("coolant", "") or "―"),
            ("B軸", info.get("b_axis", "") or "―"),
            ("Z最小",  info.get("z_min_display", "") or "―"),
            ("Z最大",  info.get("z_max_display", "") or "―"),
        ]
        for label, value in rows:
            self._make_info_row(self._block_info_inner, label, value)

    def _make_info_row(self, parent: tk.Frame, label: str, value: str) -> None:
        """ラベル：値 の1行をカード風に表示"""
        row = tk.Frame(parent, bg="#101720", bd=1, relief="solid",
                       highlightthickness=0)
        row.pack(fill="x", pady=2)
        tk.Label(row, text=label, bg="#101720", fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9), width=8, anchor="w", padx=8, pady=4).pack(side="left")
        tk.Label(row, text=value, bg="#101720", fg=TEXT_MAIN,
                 font=("Consolas", 10), anchor="w",
                 wraplength=160, justify="left", padx=4, pady=4).pack(side="left", fill="x", expand=True)

    def _refresh_block_info_cache(self, findings: list = None) -> None:
        """チェック実行時、現在のプログラムから全Nブロックの情報を集計してキャッシュする
        findings: build_report の結果を渡すと、各ブロックの危険検出件数バッジを表示
        """
        program_text = self.input_text.get("1.0", tk.END).rstrip("\n")
        if not program_text.strip():
            self._block_info_cache = {}
            self._block_info_ranges = []
            self._render_n_index()
            return

        lines = program_text.splitlines()
        blocks = split_n_blocks(lines)
        cache: dict[int, dict] = {}
        ranges: list[tuple[int, int, int]] = []  # (line_min, line_max, block_id)

        for block_id, block in enumerate(blocks):
            t_num, _ = extract_t_before_g36x(block.rows)
            tool_name_raw, _ = extract_tool_name(block.rows)
            tool_name = tool_name_raw.strip("() ") if tool_name_raw else ""
            h_list = extract_h_list(block.rows)
            h_str = ", ".join(h for h, _ in h_list)
            has_radius = block_radius_comp_line(block.rows) is not None
            is_turning = classify_block_mode(block.rows) == MODE_TURNING
            if is_turning:
                z_min, z_max = extract_z_range(block.rows)
                z_min_str = format_z_with_line(z_min).replace("\n", " ")
                z_max_str = format_z_with_line(z_max).replace("\n", " ")
            else:
                z_min_str = "(ミーリング)"
                z_max_str = "(ミーリング)"

            # 行番号範囲を先に計算（危険件数集計に使う）
            row_lines = [ln for ln, _ in block.rows]
            line_min = min(row_lines) if row_lines else 0
            line_max = max(row_lines) if row_lines else 0

            # 危険検出件数を集計（findingsがあれば）
            danger_count = 0
            if findings and row_lines:
                for f in findings:
                    line_no = getattr(f, "line_no", None)
                    if line_no is not None and line_min <= line_no <= line_max:
                        danger_count += 1

            cache[block_id] = {
                "n_label": block.n_label,
                "line_min": line_min,
                "line_max": line_max,
                "t_number": t_num,
                "tool_name": tool_name,
                "h_offsets": h_str,
                "radius_comp": "G41/G42" if has_radius else "",
                "spindle": format_spindle_info(extract_spindle_info(block.rows)),
                "feed_list": extract_feed_list(block.rows),
                "coolant": format_coolant_info(extract_coolant_info(block.rows)),
                "rotation": format_rotation_info(extract_rotation_info(block.rows)),
                "work_coord": format_work_coordinate(extract_work_coordinate(block.rows)),
                "b_axis": format_b_axis_angles(extract_b_axis_angles(block.rows)),
                "z_min_display": z_min_str,
                "z_max_display": z_max_str,
                "danger_count": danger_count,
            }
            # 行番号範囲を保存（カーソル位置からブロック特定用）
            if row_lines:
                ranges.append((line_min, line_max, block_id))

        self._block_info_cache = cache
        self._block_info_ranges = ranges
        self._render_n_index()

        # 現在カーソル位置のブロックを表示
        self._update_block_info_for_current_line(force=True)

    def _find_block_for_line(self, line_no: int) -> int | None:
        """行番号からブロックID（連番）を特定。N番号外ならNone
        同じN番号が複数あっても、行番号で正しく区別できる
        """
        for start, end, block_id in self._block_info_ranges:
            if start <= line_no <= end:
                return block_id
        return None

    def _update_block_info_for_current_line(self, force: bool = False) -> None:
        """カーソル位置から該当Nブロックを判定し、別ブロックなら表示更新"""
        if not hasattr(self, "_block_info_cache"):
            return
        if not self._block_info_cache:
            # チェック未実行
            if force or self._current_displayed_label is not None:
                self._show_block_info_placeholder("チェック実行後に表示されます")
            return

        try:
            cur_idx = self.input_text.index(tk.INSERT)
            cur_line = int(str(cur_idx).split(".")[0])
        except (tk.TclError, ValueError, IndexError):
            return

        block_id = self._find_block_for_line(cur_line)

        # ブロックIDが変わった時 or 強制更新時のみ再描画
        if not force and block_id == self._current_displayed_label:
            return

        if block_id is None:
            self._show_block_info_placeholder("[N番号外]")
            self._current_displayed_label = None
        else:
            info = self._block_info_cache.get(block_id)
            if info:
                self._render_block_info(info)
                self._current_displayed_label = block_id
            else:
                self._show_block_info_placeholder("[N番号外]")
                self._current_displayed_label = None
        self._refresh_n_index_selection()

    def _get_settings(self) -> ThresholdSettings:
        default = ThresholdSettings()
        speed = self.speed_entry.get_value() if hasattr(self, "speed_entry") else None
        feed = self.feed_entry.get_value() if hasattr(self, "feed_entry") else None
        if speed is None or speed <= 0:
            speed = default.speed_limit
        if feed is None or feed <= 0:
            feed = default.feed_limit
        return ThresholdSettings(speed_limit=speed, feed_limit=feed)

    def _on_threshold_changed(self) -> None:
        """閾値入力欄の値が変わったらdebounceをかけて再チェック実行"""
        if self._threshold_debounce_id is not None:
            try:
                self.root.after_cancel(self._threshold_debounce_id)
            except Exception:
                pass
        self._threshold_debounce_id = self.root.after(350, self._auto_rerun_check)

    def _auto_rerun_check(self) -> None:
        self._threshold_debounce_id = None
        program_text = self.input_text.get("1.0", tk.END).rstrip("\n")
        if not program_text.strip():
            return
        settings = self._get_settings()
        self._update_limits_summary(settings)
        report = build_report(program_text, settings)
        self._on_report_ready(report, program_text)
        self._highlight_all()
        # ブロック情報キャッシュも更新（findings渡して危険件数集計）
        self._refresh_block_info_cache(report.findings)

    def _update_limits_summary(self, settings: ThresholdSettings) -> None:
        if hasattr(self, "_limits_summary"):
            self._limits_summary.configure(
                text=f"S{int(settings.speed_limit)} / F{int(settings.feed_limit)}"
            )

    def _clear_output_jump_tags(self) -> None:
        jump_tags = [tag for tag in self.output_text.tag_names() if tag.startswith("jump_")]
        if jump_tags:
            self.output_text.tag_delete(*jump_tags)

    def _apply_error_highlights(self, error_line_nos: set[int], warn_line_nos: set[int] | None = None) -> None:
        warn_line_nos = warn_line_nos or set()
        self.input_text.tag_remove("error_line_highlight", "1.0", tk.END)
        self.input_text.tag_remove("warn_line_highlight", "1.0", tk.END)
        for line_no in sorted(warn_line_nos - error_line_nos):
            start = f"{line_no}.0"
            end = f"{line_no}.end"
            self.input_text.tag_add("warn_line_highlight", start, end)
        for line_no in sorted(error_line_nos):
            start = f"{line_no}.0"
            end = f"{line_no}.end"
            self.input_text.tag_add("error_line_highlight", start, end)
        self.input_text.tag_raise("jump_highlight")
        # 行番号ガターのエラー/警告行ハイライトも追従
        self._apply_linenumber_highlights()

    def _render_report(self, report: ReportData, err_lines: set[int] | None = None,
                       warn_lines: set[int] | None = None) -> None:
        """チェック結果を内部レポートバッファ(output_text)へ書き出す。
        画面には表示しないが、旧テキストレポート・保存機能・エラー行判定など
        既存ロジックが参照するため維持する。
        err_lines/warn_linesを渡すと、エディタの赤/黄ハイライトはそちらを優先する
        （目次の7分類ベースの重大度で塗り分けるため）。
        """
        self.output_text.delete("1.0", tk.END)
        self.output_text.config(cursor="arrow")
        self._clear_output_jump_tags()
        self._apply_error_highlights(
            err_lines if err_lines is not None else report.error_line_nos,
            warn_lines,
        )

        for idx, report_line in enumerate(report.lines):
            start_index = self.output_text.index("end-1c")
            self.output_text.insert(tk.END, report_line.text)
            end_index = self.output_text.index("end-1c")
            self.output_text.insert(tk.END, "\n")

            if start_index != end_index and report_line.style_key:
                self.output_text.tag_add(report_line.style_key, start_index, end_index)

            # section_heading スタイルが付いた行は heading タグも付与
            if start_index != end_index and report_line.style_key == "section_heading":
                self.output_text.tag_add("heading", start_index, end_index)

            if report_line.target_line_no is None or start_index == end_index:
                continue

            tag_name = f"jump_{idx}"
            self.output_text.tag_add("jump_link", start_index, end_index)
            self.output_text.tag_add(tag_name, start_index, end_index)
            self.output_text.tag_bind(
                tag_name,
                "<Button-1>",
                lambda event, line_no=report_line.target_line_no: self.jump_to_input_line(line_no),
            )
            self.output_text.tag_bind(tag_name, "<Enter>", lambda event: self.output_text.config(cursor="hand2"))
            self.output_text.tag_bind(tag_name, "<Leave>", lambda event: self.output_text.config(cursor="arrow"))

    # =====================================================================
    # v35 目次（TOC）まわり：findings/summary を7分類のヒット一覧に組み立て、
    # 判定帯・ゼロ件畳み表示・グループ目次・ミニマップ・エディタ塗り分けへ配線する。
    # 検査ロジック（collect_program_findings等）自体には一切手を入れない。
    # =====================================================================

    def _toc_val_msg(self, finding: Finding) -> tuple[str, str]:
        """Finding 1件を目次の (値, メッセージ) 表示用に整形する（表示専用の整形）"""
        if finding.kind == "overspeed":
            m = re.search(r"S[+-]?\d+(?:\.\d+)?", finding.text, flags=re.IGNORECASE)
            return (m.group(0).upper() if m else finding.text[:10]), "回転数上限を超過"
        if finding.kind == "overfeed":
            m = re.search(r"F[+-]?\d+(?:\.\d+)?", finding.text, flags=re.IGNORECASE)
            return (m.group(0).upper() if m else finding.text[:10]), "送り上限を超過"
        if finding.kind == "tcp":
            return finding.text, "TCP指令を検出"
        if finding.kind == "decimal_error":
            return finding.text, "小数点の誤りを検出"
        if finding.kind == "tailstock_macro_missing":
            return "M25→G361", finding.text
        return finding.text, ""

    def _collect_duplicate_n_hits(self, report: ReportData) -> list[dict]:
        hits: list[dict] = []
        for label, line_list in report.summary.duplicate_n_details:
            for line_no in sorted(line_list):
                others = [str(x) for x in sorted(line_list) if x != line_no]
                msg = f"N番号が重複（{', '.join(others)}行と）" if others else "N番号が重複"
                hits.append({"line": line_no, "val": label, "msg": msg})
        return hits

    def _collect_radius_hits(self, program_lines: list[str]) -> list[dict]:
        hits: list[dict] = []
        for block in split_n_blocks(program_lines):
            line_no = block_radius_comp_line(block.rows)
            if line_no is not None:
                hits.append({"line": line_no, "val": block.n_label, "msg": "径補正(G41/G42)を使用"})
        return hits

    def _build_toc_groups(self, report: ReportData, program_lines: list[str]) -> list[dict]:
        finding_kinds = {"overspeed", "overfeed", "decimal_error", "tailstock_macro_missing", "tcp"}
        hits_by_kind: dict[str, list[dict]] = {k: [] for k in finding_kinds}
        for finding in report.findings:
            if finding.kind in hits_by_kind:
                val, msg = self._toc_val_msg(finding)
                hits_by_kind[finding.kind].append({"line": finding.line_no, "val": val, "msg": msg})
        hits_by_kind["duplicate_n"] = self._collect_duplicate_n_hits(report)
        hits_by_kind["radius_comp_n"] = self._collect_radius_hits(program_lines)

        groups: list[dict] = []
        for g in TOC_GROUPS:
            hits = sorted(hits_by_kind.get(g["kind"], []), key=lambda h: h["line"])
            groups.append({**g, "hits": hits})
        return groups

    def _render_verdict_and_toc(self, groups: list[dict], total_lines: int) -> None:
        err_count = sum(len(g["hits"]) for g in groups if g["sev"] == "err")
        wrn_count = sum(len(g["hits"]) for g in groups if g["sev"] == "wrn")

        if err_count > 0:
            state = "ng"
        elif wrn_count > 0:
            state = "wn"
        else:
            state = "ok"
        if hasattr(self, "_verdict_band"):
            self._verdict_band.set_state(state, err=err_count, wrn=wrn_count, lines=total_lines)

        self._render_toc(groups)

    def _render_toc(self, groups: list[dict]) -> None:
        if not hasattr(self, "_toc_inner"):
            return
        for w in self._toc_inner.winfo_children():
            w.destroy()
        self._toc_item_widgets = {}

        zero_groups = [g for g in groups if not g["hits"]]
        active_groups = [g for g in groups if g["hits"]]

        # ゼロ件は畳んで1行に
        self._zeros_text.configure(state="normal")
        self._zeros_text.delete("1.0", tk.END)
        for g in zero_groups:
            self._zeros_text.insert(tk.END, "✔ ", "chk")
            self._zeros_text.insert(tk.END, g["label"] + "    ")
        self._zeros_text.configure(state="disabled")
        if zero_groups:
            self._zeros_text.pack(fill="x", before=self._toc_wrap)
        else:
            self._zeros_text.pack_forget()

        if not active_groups:
            tk.Label(self._toc_inner, text="✔ ALL CLEAR", bg=BG_PANEL, fg=OK_COLOR,
                     font=("Consolas", 16, "bold")).pack(anchor="w", padx=13, pady=(24, 4))
            tk.Label(self._toc_inner, text="検出された問題はありません", bg=BG_PANEL, fg=TEXT_MUTED,
                     font=("Yu Gothic UI", 10)).pack(anchor="w", padx=13)
            return

        for g in active_groups:
            kind = g["kind"]
            if kind not in self._toc_group_open:
                self._toc_group_open[kind] = True  # 検出ありのグループはデフォルト展開
            self._render_toc_group(g)

    def _render_toc_group(self, g: dict) -> None:
        kind = g["kind"]
        is_open = self._toc_group_open.get(kind, True)
        glyph = "✕" if g["sev"] == "err" else ("▲" if g["sev"] == "wrn" else "・")

        group_frame = tk.Frame(self._toc_inner, bg=BG_PANEL)
        group_frame.pack(fill="x")

        header = tk.Frame(group_frame, bg=BG_PANEL, cursor="hand2")
        header.pack(fill="x")
        arrow = tk.Label(header, text=("▼" if is_open else "▶"), bg=BG_PANEL, fg=TEXT_MUTED,
                         font=("Consolas", 8), cursor="hand2")
        arrow.pack(side="left", padx=(13, 8), pady=9)
        sw = tk.Frame(header, bg=g["color"], width=3, height=15)
        sw.pack(side="left", padx=(0, 8))
        tk.Label(header, text=glyph, bg=BG_PANEL, fg=g["color"], font=("Consolas", 10),
                 width=1, cursor="hand2").pack(side="left")
        tk.Label(header, text=g["label"], bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 11), cursor="hand2").pack(side="left", padx=(6, 0))
        tk.Label(header, text=str(len(g["hits"])), bg=BG_PANEL, fg=g["color"],
                 font=("Consolas", 11, "bold"), cursor="hand2").pack(side="right", padx=13)

        for w in (header, arrow, sw, *header.winfo_children()):
            w.bind("<Button-1>", lambda _e, k=kind: self._toggle_toc_group(k))

        sep = tk.Frame(group_frame, bg="#131C25", height=1)
        sep.pack(fill="x")

        if not is_open:
            return

        for hit in g["hits"]:
            line_no = hit["line"]
            selected = (self._toc_selected_line == line_no)
            item = tk.Frame(group_frame, bg=("#171F29" if selected else BG_PANEL), cursor="hand2")
            item.pack(fill="x")
            bar = tk.Frame(item, bg=(g["color"] if selected else BG_PANEL), width=3)
            bar.pack(side="left", fill="y")
            ln_lbl = tk.Label(item, text=str(line_no), bg=item["bg"],
                              fg=(ACCENT if selected else TEXT_MUTED),
                              font=("Consolas", 9), width=5, anchor="e", cursor="hand2")
            ln_lbl.pack(side="left", padx=(8, 9), pady=6)
            val_lbl = tk.Label(item, text=hit["val"], bg=item["bg"], fg=g["color"],
                               font=("Consolas", 10, "bold"), cursor="hand2")
            val_lbl.pack(side="left")
            msg_lbl = tk.Label(item, text=hit["msg"], bg=item["bg"],
                               fg=(TEXT_MAIN if selected else TEXT_MUTED),
                               font=("Yu Gothic UI", 9), anchor="w", cursor="hand2")
            msg_lbl.pack(side="left", padx=(10, 8), fill="x", expand=True)

            self._toc_item_widgets.setdefault(line_no, []).append(item)
            for w in (item, bar, ln_lbl, val_lbl, msg_lbl):
                w.bind("<Button-1>", lambda _e, ln=line_no: self._select_toc_item(ln))

    def _toggle_toc_group(self, kind: str) -> None:
        self._toc_group_open[kind] = not self._toc_group_open.get(kind, True)
        if self._last_toc_groups is not None:
            self._render_toc(self._last_toc_groups)

    def _select_toc_item(self, line_no: int) -> None:
        self._toc_selected_line = line_no
        if self._last_toc_groups is not None:
            self._render_toc(self._last_toc_groups)
        self.jump_to_input_line(line_no)

    def _update_minimap(self, groups: list[dict], program_lines: list[str]) -> None:
        if not hasattr(self, "_minimap"):
            return
        markers: list[tuple[int, str]] = []
        for g in groups:
            sev = "err" if g["sev"] == "err" else ("wrn" if g["sev"] == "wrn" else None)
            if sev is None:
                continue
            for hit in g["hits"]:
                markers.append((hit["line"], sev))
        n_starts = []
        for block in split_n_blocks(program_lines):
            block_line_nos = [ln for ln, _ in block.rows]
            if block_line_nos:
                n_starts.append(min(block_line_nos))
        self._minimap.set_data(max(1, len(program_lines)), markers, n_starts)

    def _severity_line_sets(self, groups: list[dict]) -> tuple[set[int], set[int]]:
        err_lines: set[int] = set()
        warn_lines: set[int] = set()
        for g in groups:
            target = err_lines if g["sev"] == "err" else (warn_lines if g["sev"] == "wrn" else None)
            if target is None:
                continue
            for hit in g["hits"]:
                target.add(hit["line"])
        return err_lines, warn_lines

    def _on_report_ready(self, report: ReportData, program_text: str) -> None:
        """チェック実行/自動再チェック後の共通後処理。
        判定帯・目次・ミニマップ・エディタの赤/黄ハイライトをまとめて更新する。
        """
        program_lines = program_text.splitlines()
        groups = self._build_toc_groups(report, program_lines)
        self._last_toc_groups = groups
        err_lines, warn_lines = self._severity_line_sets(groups)

        self._render_report(report, err_lines=err_lines, warn_lines=warn_lines)
        self._render_verdict_and_toc(groups, len(program_lines))
        self._update_minimap(groups, program_lines)

    def jump_to_input_line(self, line_no: int) -> None:
        start = f"{line_no}.0"
        end = f"{line_no}.end"
        self.input_text.tag_remove("jump_highlight", "1.0", tk.END)
        self.input_text.tag_add("jump_highlight", start, end)
        self.input_text.tag_raise("jump_highlight")
        self.input_text.mark_set(tk.INSERT, start)
        self.input_text.see(start)
        self.input_text.focus_set()
        self._update_block_info_for_current_line()
        self._update_gm_legend_for_current_line()
        self.status_var.set(f"{line_no}行目へ移動しました。")

    def open_file(self) -> None:
        path = filedialog.askopenfilename(
            title="NCプログラムを選択",
            filetypes=[
                ("すべてのファイル", "*.*"),
            ],
        )
        if not path:
            return
        self._load_file_from_path(path)

    def _load_file_from_path(self, path: str) -> bool:
        """指定パスのNCファイルを読み込んでエディタに展開する。
        Returns: 成功時True、失敗時False
        コマンドライン引数経由・Windowsの「プログラムから開く」経由でも呼ばれる。
        """
        if not os.path.isfile(path):
            messagebox.showerror("ファイルが見つかりません", f"指定されたファイルが存在しません。\n{path}")
            return False

        text = None
        last_error = None
        for enc in READ_ENCODINGS:
            try:
                with open(path, "r", encoding=enc) as f:
                    text = f.read()
                break
            except Exception as e:
                last_error = e

        if text is None:
            messagebox.showerror("読込エラー", f"ファイルを開けませんでした。\n{path}\n\n{last_error}")
            return False

        self.loaded_path = path
        self.path_var.set(path)
        self.input_text.delete("1.0", tk.END)
        self.input_text.insert("1.0", text)
        self.output_text.delete("1.0", tk.END)
        self._clear_output_jump_tags()
        self.input_text.tag_remove("error_line_highlight", "1.0", tk.END)
        self.input_text.tag_remove("warn_line_highlight", "1.0", tk.END)
        self.input_text.tag_remove("jump_highlight", "1.0", tk.END)
        self.status_var.set(f"読込済み: {os.path.basename(path)}")
        self._highlight_all()
        self._update_linenumbers()
        # 新ファイル読込→旧キャッシュは無効、プレースホルダーへ
        self._block_info_cache = {}
        self._block_info_ranges = []
        self._show_block_info_placeholder("チェック実行後に表示されます")
        self._render_n_index()
        self._last_toc_groups = None
        self._toc_selected_line = None
        self._render_toc_placeholder("チェック実行後に表示されます")
        if hasattr(self, "_verdict_band"):
            self._verdict_band.set_state("idle")
        if hasattr(self, "_minimap"):
            self._minimap.set_data(1, [], [])
        self._update_gm_legend_for_current_line()
        return True

    def run_check(self) -> None:
        program_text = self.input_text.get("1.0", tk.END).rstrip("\n")
        if not program_text.strip():
            messagebox.showwarning("入力なし", "NCプログラムが空です。")
            return

        # スキャンライン演出 → 完了後に本処理
        self._play_scan_animation(on_complete=lambda: self._do_check_after_scan(program_text))

    def _do_check_after_scan(self, program_text: str) -> None:
        settings = self._get_settings()
        self._update_limits_summary(settings)
        report = build_report(program_text, settings)
        self._on_report_ready(report, program_text)
        self._highlight_all()
        # Nブロック情報キャッシュを更新（findings渡して危険件数集計）
        self._refresh_block_info_cache(report.findings)
        self.status_var.set("チェック完了。目次の項目をクリックすると元の行へ移動します。")

    def _play_scan_animation(self, on_complete=None) -> None:
        """入力エリアに緑のスキャンラインを上から下に走らせる演出。
        合計約400ms、全行を横断する。完了後 on_complete を呼ぶ。
        """
        try:
            content = self.input_text.get("1.0", "end-1c")
            total_lines = max(1, content.count("\n") + 1)
        except tk.TclError:
            if on_complete:
                on_complete()
            return

        # スキャンタグ準備（2段階：先頭の明るい帯 + その上の残像）
        self.input_text.tag_configure("scan_head", background="#178F86")
        self.input_text.tag_configure("scan_trail", background="#0E2A28")
        # エラータグより下に（演出が結果表示を邪魔しない）
        self.input_text.tag_lower("scan_head")
        self.input_text.tag_lower("scan_trail")

        # 画面に見えてる範囲だけを対象にする（長いファイルでも瞬間的に）
        try:
            top_idx = self.input_text.index("@0,0")
            bottom_idx = self.input_text.index(f"@0,{self.input_text.winfo_height()}")
            top_line = int(str(top_idx).split(".")[0])
            bottom_line = int(str(bottom_idx).split(".")[0])
        except (tk.TclError, ValueError):
            top_line, bottom_line = 1, total_lines

        scan_start = top_line
        scan_end = min(bottom_line, total_lines)
        span = max(1, scan_end - scan_start + 1)
        # 総フレーム数（長いプログラムでも固定8フレーム、短いと行数そのまま）
        frame_count = min(8, span)
        frame_interval = 40  # ms

        def step(i: int) -> None:
            # すべてのscanタグを一旦剥がす
            self.input_text.tag_remove("scan_head", "1.0", tk.END)
            self.input_text.tag_remove("scan_trail", "1.0", tk.END)
            if i >= frame_count:
                # 終了処理
                if on_complete:
                    on_complete()
                return
            # 現在位置（比例計算）
            progress = i / frame_count
            head_line = scan_start + int(progress * span)
            trail_line = head_line - 1 if head_line > scan_start else None
            trail_line2 = head_line - 2 if head_line - 1 > scan_start else None

            try:
                if trail_line2:
                    self.input_text.tag_add("scan_trail", f"{trail_line2}.0", f"{trail_line2}.end + 1c")
                if trail_line:
                    self.input_text.tag_add("scan_trail", f"{trail_line}.0", f"{trail_line}.end + 1c")
                if head_line <= scan_end:
                    self.input_text.tag_add("scan_head", f"{head_line}.0", f"{head_line}.end + 1c")
            except tk.TclError:
                pass

            self.root.after(frame_interval, lambda: step(i + 1))

        step(0)

    def save_result(self) -> None:
        result = self.output_text.get("1.0", tk.END).rstrip("\n")
        if not result.strip():
            messagebox.showwarning("保存不可", "先にチェック結果を作成してください。")
            return

        initial_name = "nc_check_result.txt"
        if self.loaded_path:
            base = os.path.splitext(os.path.basename(self.loaded_path))[0]
            initial_name = f"{base}_check_result.txt"

        save_path = filedialog.asksaveasfilename(
            title="チェック結果を保存",
            defaultextension=".txt",
            initialfile=initial_name,
            filetypes=[("テキスト", "*.txt"), ("すべてのファイル", "*.*")],
        )
        if not save_path:
            return

        try:
            with open(save_path, "w", encoding="cp932", newline="\r\n") as f:
                f.write(result)
            self.status_var.set(f"保存済み(SJIS/CRLF): {save_path}")
        except UnicodeEncodeError as e:
            messagebox.showerror(
                "保存エラー",
                "SJIS(CP932)で表現できない文字が含まれています。\n"
                f"詳細: {e}",
            )
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))

    def save_nc_file(self) -> None:
        program_text = self.input_text.get("1.0", tk.END).rstrip("\n")
        if not program_text.strip():
            messagebox.showwarning("保存不可", "NCプログラムが空です。")
            return

        initial_name = "program.NC"
        if self.loaded_path:
            initial_name = os.path.basename(self.loaded_path)

        save_path = filedialog.asksaveasfilename(
            title="NCプログラムを保存",
            defaultextension=".NC",
            initialfile=initial_name,
            filetypes=[("すべてのファイル", "*.*")],
        )
        if not save_path:
            return

        try:
            with open(save_path, "w", encoding="cp932", newline="\r\n") as f:
                f.write(program_text)
            self.loaded_path = save_path
            self.path_var.set(save_path)
            self.status_var.set(f"保存済み(SJIS/CRLF): {save_path}")
        except UnicodeEncodeError as e:
            messagebox.showerror(
                "保存エラー",
                "SJIS(CP932)で表現できない文字が含まれています。\n"
                "NCプログラム内の特殊文字を確認してください。\n"
                f"詳細: {e}",
            )
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))

    def clear_input(self) -> None:
        # 入力欄が空でなければ確認
        current = self.input_text.get("1.0", "end-1c")
        if current.strip():
            ans = messagebox.askyesno(
                "確認",
                "入力欄をクリアしますか？\nこの操作は元に戻せません。",
                parent=self.root,
                default="no",
            )
            if not ans:
                return
        self.input_text.delete("1.0", tk.END)
        self.output_text.delete("1.0", tk.END)
        self._clear_output_jump_tags()
        self.input_text.tag_remove("jump_highlight", "1.0", tk.END)
        self.input_text.tag_remove("error_line_highlight", "1.0", tk.END)
        self.input_text.tag_remove("warn_line_highlight", "1.0", tk.END)
        self.loaded_path = None
        self.path_var.set("未読込")
        self.status_var.set("入力をクリアしました。")
        self._update_linenumbers()
        # キャッシュクリア＋プレースホルダー
        self._block_info_cache = {}
        self._block_info_ranges = []
        self._show_block_info_placeholder("チェック実行後に表示されます")
        self._render_n_index()
        self._last_toc_groups = None
        self._toc_selected_line = None
        self._render_toc_placeholder("チェック実行後に表示されます")
        if hasattr(self, "_verdict_band"):
            self._verdict_band.set_state("idle")
        if hasattr(self, "_minimap"):
            self._minimap.set_data(1, [], [])
        self._update_gm_legend_for_current_line()

    # --- 元に戻す（Ctrl+Z相当） ---
    def _do_undo(self) -> None:
        """入力欄のUndoを実行（Ctrl+Z相当）。"""
        try:
            self.input_text.edit_undo()
            self.status_var.set("元に戻しました")
            self._update_linenumbers()
        except tk.TclError:
            self.status_var.set("これ以上戻せません")
        # 状態更新
        self._update_undo_button_state()

    def _update_undo_button_state(self) -> None:
        """Undoスタックの状態を見てボタンを有効/無効化。"""
        if not hasattr(self, "_undo_button"):
            return
        try:
            # edit_canundo の戻り値（0/1 や True/False）
            can_undo = bool(self.input_text.tk.call(
                self.input_text._w, "edit", "canundo"
            ))
        except tk.TclError:
            can_undo = False
        self._undo_button.configure(state=("normal" if can_undo else "disabled"))

    # --- フォントサイズ変更 ---

    def _font_decrease(self) -> None:
        if self._font_size > 8:
            self._font_size -= 1
            self._apply_font_size()

    def _font_increase(self) -> None:
        if self._font_size < 24:
            self._font_size += 1
            self._apply_font_size()

    def _apply_font_size(self) -> None:
        self._font_size_label.config(text=f"{self._font_size}pt")
        self.input_text.config(font=("Consolas", self._font_size))
        self.output_text.config(font=("Consolas", self._font_size))
        # 行番号ガターも同サイズで追従
        if hasattr(self, "linenumber_text"):
            self.linenumber_text.configure(font=("Consolas", self._font_size))

        # RESULT_STYLE_MAP 内の個別タグのフォントサイズも追従させる
        # ベースサイズ11ptを基準に、現在のフォントサイズとの差分を各タグの元サイズに加算
        delta = self._font_size - 11
        style_font_bases = {
            "section_heading": ("Yu Gothic UI", 11, "bold"),
            "program_number": ("Consolas", 12, "bold"),
            "block_label": ("Consolas", 11, "bold"),
            "tool_t": ("Consolas", 11, "bold"),
            "radius_notice": ("Yu Gothic UI", 10, "bold"),
            "duplicate_n_notice": ("Yu Gothic UI", 10, "bold"),
        }
        for tag_name, (family, base_size, weight) in style_font_bases.items():
            new_size = max(7, base_size + delta)
            self.output_text.tag_configure(tag_name, font=(family, new_size, weight))

        # heading タグ（セクション見出しの追加装飾）も追従
        heading_size = max(7, 11 + delta)
        self.output_text.tag_configure("heading", font=("Consolas", heading_size, "bold"))

        self._highlight_all()
        self.status_var.set(f"フォントサイズ: {self._font_size}pt")

    # --- 検索機能 ---

    def _focus_search(self, event=None) -> None:
        self._search_entry.focus_set()
        self._search_entry.select_range(0, tk.END)
        return "break"

    def _open_replace_dialog(self, event=None) -> None:
        """Ctrl+H: 置換ダイアログを開く（既に開いてれば前面に）"""
        existing = getattr(self, "_find_dialog", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.lift()
                    existing.focus_set()
                    return "break"
            except tk.TclError:
                pass
        self._find_dialog = FindReplaceDialog(self, mode="replace")
        return "break"

    def _on_search_changed(self, event=None) -> None:
        if event and event.keysym in ("Return", "Shift_L", "Shift_R", "Escape"):
            return
        self._do_search()

    def _do_search(self) -> None:
        self.input_text.tag_remove("search_highlight", "1.0", tk.END)
        self.input_text.tag_remove("search_current", "1.0", tk.END)
        self._search_matches = []
        self._search_idx = -1

        query = self._search_var.get()
        if not query:
            if hasattr(self, "_search_count_label") and self._search_count_label:
                self._search_count_label.config(text="")
            return

        start = "1.0"
        while True:
            pos = self.input_text.search(query, start, stopindex=tk.END, nocase=True)
            if not pos:
                break
            end = f"{pos}+{len(query)}c"
            self.input_text.tag_add("search_highlight", pos, end)
            self._search_matches.append(pos)
            start = end

        count = len(self._search_matches)
        if hasattr(self, "_search_count_label") and self._search_count_label:
            self._search_count_label.config(text=f"{count}件" if count > 0 else "0件")

        if self._search_matches:
            self._search_idx = 0
            self._highlight_current_match()

    def _highlight_current_match(self) -> None:
        self.input_text.tag_remove("search_current", "1.0", tk.END)
        if not self._search_matches or self._search_idx < 0:
            return
        pos = self._search_matches[self._search_idx]
        query = self._search_var.get()
        end = f"{pos}+{len(query)}c"
        self.input_text.tag_add("search_current", pos, end)
        self.input_text.see(pos)
        total = len(self._search_matches)
        current = self._search_idx + 1
        if hasattr(self, "_search_count_label") and self._search_count_label:
            self._search_count_label.config(text=f"{current}/{total}件")

    def _search_next(self, event=None) -> None:
        if not self._search_matches:
            self._do_search()
            return
        self._search_idx = (self._search_idx + 1) % len(self._search_matches)
        self._highlight_current_match()
        return "break"

    def _search_prev(self, event=None) -> None:
        if not self._search_matches:
            self._do_search()
            return
        self._search_idx = (self._search_idx - 1) % len(self._search_matches)
        self._highlight_current_match()
        return "break"

    # --- シンタックスハイライト ---

    def _highlight_line(self, line_no: int) -> None:
        """指定行のシンタックスハイライトを更新する"""
        start = f"{line_no}.0"
        end = f"{line_no}.end"
        line_text = self.input_text.get(start, end)

        # この行のハイライトタグを全て除去
        for tag_name, _, _ in SYNTAX_COLORS:
            self.input_text.tag_remove(tag_name, start, end)

        # コメントを先にマッチ（コメント内は他のハイライトをしない）
        comment_ranges: list[tuple[int, int]] = []
        for m in re.finditer(r"\([^)]*\)", line_text):
            comment_ranges.append((m.start(), m.end()))
            tag_start = f"{line_no}.{m.start()}"
            tag_end = f"{line_no}.{m.end()}"
            self.input_text.tag_add("hl_comment", tag_start, tag_end)

        def in_comment(pos: int) -> bool:
            return any(s <= pos < e for s, e in comment_ranges)

        # コメント以外のハイライト
        for tag_name, pattern, _ in SYNTAX_COLORS:
            if tag_name == "hl_comment":
                continue
            for m in re.finditer(pattern, line_text, flags=re.IGNORECASE):
                if in_comment(m.start()):
                    continue
                tag_start = f"{line_no}.{m.start()}"
                tag_end = f"{line_no}.{m.end()}"
                self.input_text.tag_add(tag_name, tag_start, tag_end)

        # エラー・ジャンプタグを最前面に維持
        self.input_text.tag_raise("error_line_highlight")
        self.input_text.tag_raise("jump_highlight")

    def _highlight_all(self) -> None:
        """全行のシンタックスハイライトを一括適用する"""
        content = self.input_text.get("1.0", tk.END)
        total_lines = content.count("\n")
        for line_no in range(1, total_lines + 1):
            self._highlight_line(line_no)
        # Nブロック区切りの縦ラインも付け直す
        self._apply_n_block_dividers()

    def _apply_n_block_dividers(self) -> None:
        """Nで始まる行の番号を行番号ガター側で強調する（工具交換境目の視認性）
        本文側には一切タグを付けない。
        """
        if not hasattr(self, "linenumber_text"):
            return
        ln = self.linenumber_text
        try:
            ln.configure(state="normal")
            ln.tag_remove("ln_n_block", "1.0", tk.END)
        except tk.TclError:
            return

        content = self.input_text.get("1.0", tk.END)
        for idx, line in enumerate(content.splitlines(), start=1):
            stripped = line.lstrip()
            if not stripped:
                continue
            if re.match(r"^N\d+\b", stripped):
                try:
                    ln.tag_add("ln_n_block", f"{idx}.0", f"{idx}.end")
                except tk.TclError:
                    pass
        ln.configure(state="disabled")

    def _on_key_release(self, event=None) -> None:
        """キー入力時にカーソル行のハイライトを更新する"""
        try:
            cursor_pos = self.input_text.index(tk.INSERT)
            current_line = int(cursor_pos.split(".")[0])
        except (ValueError, tk.TclError):
            return
        self._highlight_line(current_line)
        # 改行・削除時は前後の行も更新
        if event and event.keysym in ("Return", "BackSpace", "Delete"):
            if current_line > 1:
                self._highlight_line(current_line - 1)
            self._highlight_line(current_line + 1)

    def export_excel(self) -> None:
        if not HAS_OPENPYXL:
            messagebox.showerror("未インストール", "openpyxl がインストールされていません。\npip install openpyxl を実行してください。")
            return
        program_text = self.input_text.get("1.0", tk.END).rstrip("\n")
        if not program_text.strip():
            messagebox.showwarning("入力なし", "NCプログラムが空です。先にチェックを実行してください。")
            return
        initial_name = "初品加工確認資料.xlsx"
        if self.loaded_path:
            base = os.path.splitext(os.path.basename(self.loaded_path))[0]
            initial_name = f"{base}_初品加工確認資料.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="初品加工確認資料を保存",
            defaultextension=".xlsx",
            initialfile=initial_name,
            filetypes=[("Excel", "*.xlsx"), ("すべてのファイル", "*.*")],
        )
        if not save_path:
            return
        try:
            file_name = os.path.basename(self.loaded_path) if self.loaded_path else ""
            generate_first_article_excel(program_text, save_path, file_name=file_name)
            self.status_var.set(f"初品資料保存済み: {save_path}")
        except Exception as e:
            messagebox.showerror("Excel出力エラー", str(e))


def generate_first_article_excel(nc_text: str, output_path: str, file_name: str = "", operator_name: str = "") -> str:
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl が必要です")

    lines = nc_text.splitlines()
    settings = ThresholdSettings()
    prog_num, _ = extract_program_number(lines)
    prog_comment = extract_program_comment(lines)
    blocks = split_n_blocks(lines)
    findings = collect_program_findings(lines, settings)

    summary = SummaryData()
    for f in findings:
        if f.kind == "overspeed": summary.overspeed_count += 1
        elif f.kind == "overfeed": summary.overfeed_count += 1
        elif f.kind == "tcp": summary.tcp_count += 1
        elif f.kind == "decimal_error": summary.decimal_error_count += 1
        elif f.kind == "tailstock_macro_missing": summary.tailstock_macro_missing_count += 1
    for block in blocks:
        if block_radius_comp_line(block.rows) is not None:
            summary.radius_comp_n_count += 1

    block_info_list = []
    for block in blocks:
        t_num, _ = extract_t_before_g36x(block.rows)
        tool_name_raw, _ = extract_tool_name(block.rows)
        tool_name = tool_name_raw.strip("() ") if tool_name_raw else ""
        h_list = extract_h_list(block.rows)
        h_str = ", ".join(h for h, _ in h_list)
        has_radius = block_radius_comp_line(block.rows) is not None
        block_lines_set = set(ln for ln, _ in block.rows)
        block_findings = [f for f in findings if f.line_no in block_lines_set]
        # ターニングブロックのみZ最小/最大を拾う
        is_turning = classify_block_mode(block.rows) == MODE_TURNING
        if is_turning:
            z_min, z_max = extract_z_range(block.rows)
            z_min_str = format_z_with_line(z_min)
            z_max_str = format_z_with_line(z_max)
        else:
            z_min_str = ""
            z_max_str = ""
        block_info_list.append({
            "n_label": block.n_label, "t_number": t_num, "tool_name": tool_name,
            "h_offsets": h_str, "radius_comp": "G41/G42" if has_radius else "",
            "findings": block_findings,
            "spindle": format_spindle_info(extract_spindle_info(block.rows)),
            "feed_list": extract_feed_list(block.rows),
            "z_min": z_min_str,
            "z_max": z_max_str,
        })

    # Excel色定義
    XN = "1B2A4A"; XDH = "2D3748"; XW = "FFFFFF"; XLG = "F7F8FA"
    XRB = "FEF2F2"; XRT = "DC2626"; XAB = "FFFBEB"; XAT = "D97706"
    XPB = "F5F3FF"; XPT = "7C3AED"; XBB = "EFF6FF"; XBT = "2563EB"
    XGB = "F0FDF4"; XGT = "16A34A"

    tb = Border(left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
                top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"))
    hf = XlFont(name="Yu Gothic UI", size=8, bold=True, color=XW)
    hfl = PatternFill("solid", fgColor=XDH)
    ha = XlAlign(horizontal="center", vertical="center", wrap_text=True)
    nf = XlFont(name="Yu Gothic UI", size=10)
    na = XlAlign(vertical="center", wrap_text=True)
    ca = XlAlign(horizontal="center", vertical="center")

    def sc(ws, row, col, value, font=None, fill=None, alignment=None, border=None):
        cell = ws.cell(row=row, column=col, value=value)
        if font: cell.font = font
        if fill: cell.fill = fill
        if alignment: cell.alignment = alignment
        if border: cell.border = border
        return cell

    wb = Workbook()
    ws = wb.active
    ws.title = "初品加工確認資料"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.4; ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.5
    for cl, w in {"A": 4, "B": 6, "C": 8, "D": 28, "E": 10, "F": 8, "G": 22, "H": 20, "I": 11, "J": 11}.items():
        ws.column_dimensions[cl].width = w

    # タイトル
    ws.merge_cells("A1:J1")
    sc(ws, 1, 1, "初品加工確認資料", font=XlFont(name="Yu Gothic UI", size=16, bold=True, color=XN),
       alignment=XlAlign(horizontal="center", vertical="center"))
    ws.row_dimensions[1].height = 36

    # 基本情報
    ws.merge_cells("A2:J2"); ws.row_dimensions[2].height = 6
    ilf = XlFont(name="Yu Gothic UI", size=8, color="64748B")
    ivf = XlFont(name="Yu Gothic UI", size=11, bold=True)
    ifl = PatternFill("solid", fgColor=XLG)
    for sc_, ec_, label, value in [("A","B","プログラム番号",prog_num),("C","D","プログラム名",prog_comment),
                                    ("E","F","ファイル名",file_name if file_name else "―"),
                                    ("G","H","出力日",datetime.now().strftime("%Y/%m/%d")),
                                    ("I","J","担当者",operator_name if operator_name else "")]:
        ws.merge_cells(f"{sc_}3:{ec_}3"); ws.merge_cells(f"{sc_}4:{ec_}4")
        ci = ord(sc_) - ord("A") + 1; ei = ord(ec_) - ord("A") + 1
        sc(ws, 3, ci, label, font=ilf, fill=ifl, alignment=XlAlign(horizontal="left", vertical="center"), border=tb)
        sc(ws, 4, ci, value, font=ivf, fill=PatternFill("solid", fgColor=XW),
           alignment=XlAlign(horizontal="left", vertical="center"), border=tb)
        for c in range(ci, ei + 1):
            ws.cell(row=3, column=c).border = tb; ws.cell(row=3, column=c).fill = ifl
            ws.cell(row=4, column=c).border = tb
    ws.row_dimensions[3].height = 22; ws.row_dimensions[4].height = 28

    # 工具一覧
    ws.merge_cells("A5:J5"); ws.row_dimensions[5].height = 6
    ws.merge_cells("A6:J6")
    sc(ws, 6, 1, "■ 工具・補正一覧", font=XlFont(name="Yu Gothic UI", size=11, bold=True, color=XN),
       alignment=XlAlign(vertical="center"))
    ws.row_dimensions[6].height = 26
    for ci, ht in enumerate(["No.","N番号","T番号","工具名称","H補正","径補正","回転数","送り","Z最小(行)","Z最大(行)"], start=1):
        sc(ws, 7, ci, ht, font=hf, fill=hfl, alignment=ha, border=tb)
    ws.row_dimensions[7].height = 28

    kl_map = {"overspeed":"回転数超え","overfeed":"送り超え","tcp":"G43.4","decimal_error":"小数点間違い","tailstock_macro_missing":"芯押しマクロ忘れ"}
    kc_map = {"overspeed":XRT,"overfeed":XAT,"tcp":XPT,"decimal_error":XBT,"tailstock_macro_missing":XPT}

    row = 8
    for idx, info in enumerate(block_info_list, start=1):
        cr = 7 + idx
        rf = PatternFill("solid", fgColor=XW)
        af = PatternFill("solid", fgColor=XLG) if idx % 2 == 0 else rf
        sf = XlFont(name="Yu Gothic UI", size=8)  # データ行は8pt
        for ci, val in enumerate([idx, info["n_label"], info["t_number"], info["tool_name"],
                                   info["h_offsets"], info["radius_comp"],
                                   info["spindle"], info["feed_list"],
                                   info["z_min"], info["z_max"]], start=1):
            ft = sf; al = XlAlign(vertical="center", wrap_text=True); fl = af
            if ci == 1: al = XlAlign(horizontal="center", vertical="center", wrap_text=True)
            elif ci in (2, 3, 5, 6, 9, 10): al = XlAlign(horizontal="center", vertical="center", wrap_text=True)
            sc(ws, cr, ci, val, font=ft, fill=fl, alignment=al, border=tb)
        # 回転数・送り・Z値の行数から行高を計算
        spindle_lines = info["spindle"].count("\n") + 1 if info["spindle"] else 1
        feed_commas = info["feed_list"].count(",") + 1 if info["feed_list"] else 1
        feed_wrap_lines = max(1, (feed_commas + 3) // 4) if feed_commas > 4 else 1
        z_min_lines = info["z_min"].count("\n") + 1 if info["z_min"] else 1
        z_max_lines = info["z_max"].count("\n") + 1 if info["z_max"] else 1
        max_lines = max(spindle_lines, feed_wrap_lines, z_min_lines, z_max_lines)
        ws.row_dimensions[cr].height = max(22, 14 * max_lines)
        row = cr + 1

    # 危険検出一覧
    if findings:
        row += 1
        ws.merge_cells(f"A{row}:J{row}"); ws.row_dimensions[row].height = 6; row += 1
        ws.merge_cells(f"A{row}:J{row}")
        sc(ws, row, 1, "■ 危険検出一覧", font=XlFont(name="Yu Gothic UI", size=11, bold=True, color=XN),
           alignment=XlAlign(vertical="center"))
        ws.row_dimensions[row].height = 26; row += 1
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=10)
        for ci, h in enumerate(["No.","行番号","種別","プログラム原文"], start=1):
            sc(ws, row, ci, h, font=hf, fill=hfl, alignment=ha, border=tb)
        for c in range(1, 11):
            ws.cell(row=row, column=c).border = tb; ws.cell(row=row, column=c).fill = hfl
            ws.cell(row=row, column=c).font = hf; ws.cell(row=row, column=c).alignment = ha
        ws.row_dimensions[row].height = 28; row += 1
        for idx, f in enumerate(findings, start=1):
            kl = kl_map.get(f.kind, f.kind); fc = kc_map.get(f.kind, XRT)
            bg = {"overspeed":XRB,"overfeed":XAB,"tcp":XPB,"decimal_error":XBB,"tailstock_macro_missing":XPB}.get(f.kind, XRB)
            raw = lines[f.line_no - 1].rstrip() if 1 <= f.line_no <= len(lines) else ""
            ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=10)
            sc(ws, row, 1, idx, font=nf, fill=PatternFill("solid", fgColor=XW), alignment=ca, border=tb)
            sc(ws, row, 2, f.line_no, font=nf, fill=PatternFill("solid", fgColor=XW), alignment=ca, border=tb)
            sc(ws, row, 3, kl, font=XlFont(name="Yu Gothic UI", size=8, bold=True, color=fc),
               fill=PatternFill("solid", fgColor=bg), alignment=ca, border=tb)
            sc(ws, row, 4, raw, font=XlFont(name="Consolas", size=10),
               fill=PatternFill("solid", fgColor=XW), alignment=XlAlign(vertical="center"), border=tb)
            for c in range(5, 11): ws.cell(row=row, column=c).border = tb
            ws.row_dimensions[row].height = 24; row += 1

    ws.print_area = f"A1:J{row}"
    wb.save(output_path)
    return output_path


class GmGlossaryDialog(tk.Toplevel):
    """G/M記号一覧（全件）をユーザーが自由に追加・編集・削除できるダイアログ。
    保存内容は gm_glossary.json に書き出され、app._gm_codes として使われる。
    """

    def __init__(self, app: "NcCheckApp") -> None:
        super().__init__(app.root)
        self.app = app
        self._entries: dict[str, str] = dict(getattr(app, "_gm_codes", {}) or {})
        self._editing_code: str | None = None

        self.title("G/M記号の追加・編集")
        self.configure(bg=BG_PANEL)
        self.transient(app.root)
        self.resizable(True, True)
        self.geometry("520x620")
        self.minsize(440, 380)

        self._build_ui()
        self._render_list()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        try:
            px = app.root.winfo_rootx() + (app.root.winfo_width() - self.winfo_width()) // 2
            py = app.root.winfo_rooty() + (app.root.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        outer = tk.Frame(self, bg=BG_PANEL, padx=14, pady=12)
        outer.pack(fill="both", expand=True)

        tk.Label(outer, text="G/M記号の追加・編集", bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w")
        tk.Label(outer, text="コードと意味は自由に追加・編集・削除できます（gm_glossary.jsonで管理）。"
                             "一覧の行をクリックすると編集できます。",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                 wraplength=460, justify="left").pack(anchor="w", pady=(2, 10))

        # ===== 追加・編集フォーム =====
        add_row = tk.Frame(outer, bg=BG_PANEL)
        add_row.pack(fill="x", pady=(0, 10))
        self._code_var = tk.StringVar()
        self._meaning_var = tk.StringVar()
        tk.Label(add_row, text="コード", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).grid(row=0, column=0, sticky="w")
        tk.Label(add_row, text="意味", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).grid(row=0, column=1, sticky="w", padx=(8, 0))
        self._code_entry = tk.Entry(add_row, textvariable=self._code_var, font=("Consolas", 11),
                                    bg=INPUT_BG, fg=TEXT_MAIN, insertbackground=ACCENT,
                                    relief="solid", bd=1, highlightthickness=1,
                                    highlightbackground=BORDER, highlightcolor=ACCENT, width=10)
        self._code_entry.grid(row=1, column=0, sticky="ew")
        meaning_entry = tk.Entry(add_row, textvariable=self._meaning_var, font=("Yu Gothic UI", 10),
                                 bg=INPUT_BG, fg=TEXT_MAIN, insertbackground=ACCENT,
                                 relief="solid", bd=1, highlightthickness=1,
                                 highlightbackground=BORDER, highlightcolor=ACCENT)
        meaning_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8))
        add_row.grid_columnconfigure(1, weight=1)
        self._add_btn = tk.Button(
            add_row, text="追加", command=self._add_entry,
            font=("Yu Gothic UI", 9, "bold"), bg=ACCENT, fg="#04231F",
            activebackground=ACCENT_DARK, activeforeground="#04231F",
            relief="solid", bd=1, padx=12, pady=3, cursor="hand2",
        )
        self._add_btn.grid(row=1, column=2, sticky="e")
        self._clear_btn = tk.Button(
            add_row, text="編集をやめる", command=self._clear_form,
            font=("Yu Gothic UI", 8), bg=INPUT_BG, fg=TEXT_MUTED,
            activebackground="#1F2B3A", activeforeground=TEXT_MAIN,
            relief="solid", bd=1, padx=8, pady=3, cursor="hand2",
        )
        meaning_entry.bind("<Return>", lambda _e: self._add_entry())
        self._code_entry.bind("<Return>", lambda _e: meaning_entry.focus_set())

        # ===== 一覧（スクロール可能） =====
        list_wrap = tk.Frame(outer, bg=INPUT_BG, relief="solid", bd=1,
                             highlightthickness=1, highlightbackground=BORDER)
        list_wrap.pack(fill="both", expand=True)
        canvas = tk.Canvas(list_wrap, bg=INPUT_BG, highlightthickness=0, bd=0)
        canvas.pack(side="left", fill="both", expand=True)
        vbar = tk.Scrollbar(list_wrap, orient="vertical", bg=BG_PANEL,
                            troughcolor="#101720", activebackground=ACCENT,
                            command=canvas.yview)
        vbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=vbar.set)
        self._list_canvas = canvas
        self._list_inner = tk.Frame(canvas, bg=INPUT_BG)
        inner_window = canvas.create_window((0, 0), window=self._list_inner, anchor="nw")

        def _on_inner_configure(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        self._list_inner.bind("<Configure>", _on_inner_configure)

        def _on_canvas_configure(event):
            canvas.itemconfigure(inner_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_wheel(event):
            if event.delta:
                canvas.yview_scroll(int(-event.delta / 120), "units")
            else:
                canvas.yview_scroll(-1 if getattr(event, "num", 0) == 4 else 1, "units")
        canvas.bind("<MouseWheel>", _on_wheel)
        canvas.bind("<Button-4>", _on_wheel)
        canvas.bind("<Button-5>", _on_wheel)

        # ===== フッター =====
        footer = tk.Frame(outer, bg=BG_PANEL)
        footer.pack(fill="x", pady=(10, 0))
        tk.Button(
            footer, text="保存して閉じる", command=self._save_and_close,
            font=("Yu Gothic UI", 10, "bold"), bg=ACCENT, fg="#04231F",
            activebackground=ACCENT_DARK, activeforeground="#04231F",
            relief="solid", bd=1, padx=16, pady=6, cursor="hand2",
        ).pack(side="right")
        tk.Button(
            footer, text="キャンセル", command=self.destroy,
            font=("Yu Gothic UI", 9), bg=INPUT_BG, fg=TEXT_MAIN,
            activebackground="#1F2B3A", activeforeground=TEXT_MAIN,
            relief="solid", bd=1, padx=12, pady=6, cursor="hand2",
        ).pack(side="right", padx=(0, 8))

    def _add_entry(self) -> None:
        code = self._code_var.get().strip().upper()
        meaning = self._meaning_var.get().strip()
        if not code or not meaning:
            messagebox.showwarning("入力不足", "コードと意味の両方を入力してください。", parent=self)
            return
        # 編集中に別コードへ書き換えた場合は元コードを消してから登録し直す
        if self._editing_code and self._editing_code != code:
            self._entries.pop(self._editing_code, None)
        self._entries[code] = meaning
        self._clear_form()
        self._render_list()

    def _delete_entry(self, code: str) -> None:
        self._entries.pop(code, None)
        if self._editing_code == code:
            self._clear_form()
        self._render_list()

    def _load_for_edit(self, code: str) -> None:
        self._editing_code = code
        self._code_var.set(code)
        self._meaning_var.set(self._entries.get(code, ""))
        self._add_btn.configure(text="更新")
        self._clear_btn.grid(row=1, column=3, sticky="e", padx=(6, 0))
        self._code_entry.focus_set()

    def _clear_form(self) -> None:
        self._editing_code = None
        self._code_var.set("")
        self._meaning_var.set("")
        self._add_btn.configure(text="追加")
        self._clear_btn.grid_remove()

    def _render_list(self) -> None:
        for w in self._list_inner.winfo_children():
            w.destroy()
        if not self._entries:
            tk.Label(self._list_inner, text="コードがまだありません",
                     bg=INPUT_BG, fg=TEXT_MUTED, font=("Yu Gothic UI", 9)).pack(
                         anchor="w", padx=10, pady=12)
            return
        for code in sorted(self._entries.keys()):
            meaning = self._entries[code]
            selected = code == self._editing_code
            row_bg = "#171F29" if selected else INPUT_BG
            row = tk.Frame(self._list_inner, bg=row_bg, cursor="hand2")
            row.pack(fill="x", padx=8, pady=3)
            code_lbl = tk.Label(row, text=code, bg=row_bg, fg=ACCENT,
                                font=("Consolas", 11, "bold"), width=9, anchor="w")
            code_lbl.pack(side="left")
            meaning_lbl = tk.Label(row, text=meaning, bg=row_bg, fg=TEXT_MAIN,
                                   font=("Yu Gothic UI", 10), anchor="w", justify="left",
                                   wraplength=230)
            meaning_lbl.pack(side="left", fill="x", expand=True)
            tk.Button(
                row, text="削除", command=lambda c=code: self._delete_entry(c),
                font=("Yu Gothic UI", 8), bg=row_bg, fg="#C97C8A",
                activebackground="#452030", activeforeground="#E0A0AE",
                relief="solid", bd=1, padx=8, pady=1, cursor="hand2",
            ).pack(side="right")
            for w in (row, code_lbl, meaning_lbl):
                w.bind("<Button-1>", lambda _e, c=code: self._load_for_edit(c))

    def _save_and_close(self) -> None:
        self.app._on_gm_glossary_saved(dict(self._entries))
        self.destroy()


class NcMacroAssistDialog(tk.Toplevel):
    """カスタムマクロ（O番号サブプログラム）の呼出し行を組み立てて挿入する支援ダイアログ。
    マクロ定義（O番号・タイトル・引数一覧・メモ）は nc_macros.json に一括保存され、
    自動解析（ファイル取込）と手動作成・編集の両方に対応する。複数のマクロを登録・切替できる。
    """

    def __init__(self, app: "NcCheckApp") -> None:
        super().__init__(app.root)
        self.app = app
        self._defs: list[dict] = []
        self._selected: dict | None = None
        self._arg_vars: dict[str, tk.StringVar] = {}

        self.title("マクロ呼出し支援")
        self.configure(bg=BG_PANEL)
        self.transient(app.root)
        self.resizable(True, True)
        self.geometry("900x640")
        self.minsize(780, 500)

        self._build_ui()
        self._reload_defs()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        try:
            px = app.root.winfo_rootx() + (app.root.winfo_width() - self.winfo_width()) // 2
            py = app.root.winfo_rooty() + (app.root.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        outer = tk.Frame(self, bg=BG_PANEL, padx=14, pady=12)
        outer.pack(fill="both", expand=True)

        tk.Label(outer, text="マクロ呼出し支援", bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w")
        tk.Label(outer, text="マクロを選び、引数を入力してカーソル位置へ呼出し行を挿入します。",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9)).pack(anchor="w", pady=(2, 10))

        body = tk.Frame(outer, bg=BG_PANEL)
        body.pack(fill="both", expand=True)

        # ===== 左：マクロ一覧 =====
        left = tk.Frame(body, bg=INPUT_BG, relief="solid", bd=1,
                        highlightthickness=1, highlightbackground=BORDER, width=270)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        left_head = tk.Frame(left, bg=INPUT_BG, padx=10, pady=8)
        left_head.pack(fill="x")
        tk.Label(left_head, text="マクロ一覧", bg=INPUT_BG, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9, "bold")).pack(anchor="w")
        btn_row = tk.Frame(left_head, bg=INPUT_BG)
        btn_row.pack(fill="x", pady=(6, 0))
        tk.Button(
            btn_row, text="＋ 自動解析", command=self._import_macro,
            font=("Yu Gothic UI", 8, "bold"), bg=ACCENT, fg="#04231F",
            activebackground=ACCENT_DARK, activeforeground="#04231F",
            relief="solid", bd=1, padx=6, pady=2, cursor="hand2",
        ).pack(side="left")
        tk.Button(
            btn_row, text="＋ 手動作成", command=self._create_macro_manual,
            font=("Yu Gothic UI", 8), bg=BG_PANEL, fg=TEXT_MAIN,
            activebackground="#1F2B3A", activeforeground=TEXT_MAIN,
            relief="solid", bd=1, padx=6, pady=2, cursor="hand2",
        ).pack(side="left", padx=(6, 0))

        list_canvas = tk.Canvas(left, bg=INPUT_BG, highlightthickness=0, bd=0)
        list_canvas.pack(side="left", fill="both", expand=True)
        list_vbar = tk.Scrollbar(left, orient="vertical", bg=BG_PANEL,
                                 troughcolor="#101720", activebackground=ACCENT,
                                 command=list_canvas.yview)
        list_vbar.pack(side="right", fill="y")
        list_canvas.configure(yscrollcommand=list_vbar.set)
        self._list_inner = tk.Frame(list_canvas, bg=INPUT_BG)
        inner_window = list_canvas.create_window((0, 0), window=self._list_inner, anchor="nw")

        def _on_inner_configure(_e=None):
            list_canvas.configure(scrollregion=list_canvas.bbox("all"))
        self._list_inner.bind("<Configure>", _on_inner_configure)

        def _on_canvas_configure(event):
            list_canvas.itemconfigure(inner_window, width=event.width)
        list_canvas.bind("<Configure>", _on_canvas_configure)

        # ===== 右：選択したマクロの引数フォーム =====
        right = tk.Frame(body, bg=BG_PANEL, padx=12)
        right.pack(side="left", fill="both", expand=True)

        detail_header = tk.Frame(right, bg=BG_PANEL)
        detail_header.pack(fill="x")
        self._detail_title = tk.Label(detail_header, text="左からマクロを選んでください",
                                      bg=BG_PANEL, fg=TEXT_MAIN, font=("Consolas", 13, "bold"),
                                      anchor="w")
        self._detail_title.pack(side="left", anchor="w")
        self._edit_btn = tk.Button(
            detail_header, text="編集", command=self._edit_selected,
            font=("Yu Gothic UI", 8), bg=INPUT_BG, fg=TEXT_MAIN,
            activebackground="#1F2B3A", activeforeground=TEXT_MAIN,
            relief="solid", bd=1, padx=8, pady=2, cursor="hand2",
        )
        self._detail_sub = tk.Label(right, text="", bg=BG_PANEL, fg=TEXT_MUTED,
                                    font=("Yu Gothic UI", 9), anchor="w")
        self._detail_sub.pack(anchor="w", pady=(2, 6))

        # ===== メモ（あれば表示） =====
        self._memo_frame = tk.Frame(right, bg=BG_PANEL)
        tk.Label(self._memo_frame, text="メモ", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).pack(anchor="w")
        self._memo_text = tk.Text(
            self._memo_frame, height=3, wrap="word", bd=1, relief="solid",
            bg=INPUT_BG, fg=TEXT_MUTED, font=("Yu Gothic UI", 9), padx=6, pady=4,
            highlightthickness=1, highlightbackground=BORDER, state="disabled", cursor="arrow",
        )
        self._memo_text.pack(fill="x", pady=(2, 8))

        form_wrap = tk.Frame(right, bg=INPUT_BG, relief="solid", bd=1,
                             highlightthickness=1, highlightbackground=BORDER)
        form_wrap.pack(fill="both", expand=True)
        form_canvas = tk.Canvas(form_wrap, bg=INPUT_BG, highlightthickness=0, bd=0)
        form_canvas.pack(side="left", fill="both", expand=True)
        form_vbar = tk.Scrollbar(form_wrap, orient="vertical", bg=BG_PANEL,
                                 troughcolor="#101720", activebackground=ACCENT,
                                 command=form_canvas.yview)
        form_vbar.pack(side="right", fill="y")
        form_canvas.configure(yscrollcommand=form_vbar.set)
        self._form_inner = tk.Frame(form_canvas, bg=INPUT_BG, padx=14, pady=12)
        form_window = form_canvas.create_window((0, 0), window=self._form_inner, anchor="nw")

        def _on_form_configure(_e=None):
            form_canvas.configure(scrollregion=form_canvas.bbox("all"))
        self._form_inner.bind("<Configure>", _on_form_configure)

        def _on_form_canvas_configure(event):
            form_canvas.itemconfigure(form_window, width=event.width)
        form_canvas.bind("<Configure>", _on_form_canvas_configure)

        # ===== プレビュー & 挿入 =====
        preview_wrap = tk.Frame(right, bg=BG_PANEL)
        preview_wrap.pack(fill="x", pady=(10, 0))
        note = tk.Label(preview_wrap, text="呼出し行プレビュー（数値のみの引数には自動で小数点を付与）",
                        bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 8))
        note.pack(anchor="w")
        self._preview_var = tk.StringVar(value="")
        preview_entry = tk.Entry(
            preview_wrap, textvariable=self._preview_var, font=("Consolas", 12, "bold"),
            bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT, relief="solid", bd=1,
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
            state="readonly", readonlybackground=INPUT_BG,
        )
        preview_entry.pack(fill="x", pady=(3, 8))

        footer = tk.Frame(right, bg=BG_PANEL)
        footer.pack(fill="x")
        tk.Button(
            footer, text="カーソル位置へ挿入", command=self._insert_call,
            font=("Yu Gothic UI", 10, "bold"), bg=ACCENT, fg="#04231F",
            activebackground=ACCENT_DARK, activeforeground="#04231F",
            relief="solid", bd=1, padx=16, pady=6, cursor="hand2",
        ).pack(side="right")
        tk.Button(
            footer, text="閉じる", command=self.destroy,
            font=("Yu Gothic UI", 9), bg=INPUT_BG, fg=TEXT_MAIN,
            activebackground="#1F2B3A", activeforeground=TEXT_MAIN,
            relief="solid", bd=1, padx=12, pady=6, cursor="hand2",
        ).pack(side="right", padx=(0, 8))

    # ---- マクロ一覧 ----

    def _reload_defs(self) -> None:
        self._defs = load_nc_macro_defs()
        for w in self._list_inner.winfo_children():
            w.destroy()
        if not self._defs:
            tk.Label(self._list_inner, text="マクロがありません。\n「＋ 自動解析」または「＋ 手動作成」で追加できます。",
                     bg=INPUT_BG, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                     wraplength=230, justify="left").pack(anchor="w", padx=10, pady=12)
            return
        for d in self._defs:
            selected = self._selected is not None and self._selected.get("o_number") == d["o_number"]
            row_bg = "#171F29" if selected else INPUT_BG
            row = tk.Frame(self._list_inner, bg=row_bg, cursor="hand2")
            row.pack(fill="x")
            bar = tk.Frame(row, bg=(ACCENT if selected else INPUT_BG), width=3)
            bar.pack(side="left", fill="y")
            col = tk.Frame(row, bg=row_bg)
            col.pack(side="left", fill="x", expand=True, padx=(8, 4), pady=7)
            tk.Label(col, text=d["o_number"], bg=row_bg, fg=(ACCENT if selected else TEXT_MAIN),
                     font=("Consolas", 11, "bold"), anchor="w").pack(anchor="w")
            tk.Label(col, text=d.get("title", ""), bg=row_bg, fg=TEXT_MUTED,
                     font=("Yu Gothic UI", 8), anchor="w", wraplength=170,
                     justify="left").pack(anchor="w")
            edit_icon = tk.Label(row, text="編集", bg=row_bg, fg=TEXT_MUTED,
                                 font=("Yu Gothic UI", 8), cursor="hand2")
            edit_icon.pack(side="right", padx=(0, 8))
            edit_icon.bind("<Button-1>", lambda _e, dd=d: self._edit_macro(dd))
            for w in (row, bar, col, *col.winfo_children()):
                w.bind("<Button-1>", lambda _e, dd=d: self._select_macro(dd))

    def _import_macro(self) -> None:
        path = filedialog.askopenfilename(
            title="マクロ定義ファイルを選択",
            filetypes=[("テキスト/NCファイル", "*.txt *.nc *.NC"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        text = None
        for enc in READ_ENCODINGS:
            try:
                with open(path, "r", encoding=enc) as f:
                    text = f.read()
                break
            except (OSError, UnicodeDecodeError):
                continue
        if text is None:
            messagebox.showerror("読込エラー", f"ファイルを読み込めませんでした。\n{path}", parent=self)
            return
        parsed = parse_macro_definition(text)
        if not parsed:
            messagebox.showerror(
                "解析エラー",
                "O番号の見出し（例: O7021(タイトル)）が見つかりませんでした。\n"
                "「＋ 手動作成」から直接入力することもできます。",
                parent=self,
            )
            return
        # 自動解析はあくまで下書き。手動編集ダイアログを開いて内容を確認・修正してから保存する。
        MacroDefEditDialog(self.app, on_save=self._on_macro_def_saved,
                           macro_def=parsed, original_o_number=parsed["o_number"])

    def _create_macro_manual(self) -> None:
        MacroDefEditDialog(self.app, on_save=self._on_macro_def_saved)

    def _edit_selected(self) -> None:
        if self._selected:
            self._edit_macro(self._selected)

    def _edit_macro(self, macro_def: dict) -> None:
        MacroDefEditDialog(self.app, on_save=self._on_macro_def_saved,
                           macro_def=macro_def, original_o_number=macro_def.get("o_number"))

    def _on_macro_def_saved(self, macro_def: dict, original_o_number: str | None) -> None:
        """MacroDefEditDialogからの保存コールバック：O番号で一意になるよう置き換えて保存"""
        defs = load_nc_macro_defs()
        if original_o_number:
            defs = [d for d in defs if d.get("o_number") != original_o_number]
        defs = [d for d in defs if d.get("o_number") != macro_def["o_number"]]
        defs.append(macro_def)
        defs.sort(key=lambda d: d["o_number"])
        save_nc_macro_defs(defs)
        self._reload_defs()
        self._select_macro(macro_def)

    # ---- 引数フォーム ----

    def _select_macro(self, macro_def: dict) -> None:
        self._selected = macro_def
        self._arg_vars = {}
        self._reload_defs()

        self._detail_title.configure(text=f"{macro_def['o_number']}  {macro_def.get('title', '')}")
        self._edit_btn.pack(side="right")
        p_number = re.sub(r"\D", "", macro_def["o_number"])
        self._detail_sub.configure(text=f"呼出しコード: G65 P{p_number}")

        memo = (macro_def.get("memo") or "").strip()
        if memo:
            self._memo_text.configure(state="normal")
            self._memo_text.delete("1.0", tk.END)
            self._memo_text.insert("1.0", memo)
            self._memo_text.configure(state="disabled")
            self._memo_frame.pack(fill="x", after=self._detail_sub)
        else:
            self._memo_frame.pack_forget()

        for w in self._form_inner.winfo_children():
            w.destroy()

        args = macro_def.get("args", [])
        if not args:
            tk.Label(self._form_inner, text="このマクロには引数の定義がありません。",
                     bg=INPUT_BG, fg=TEXT_MUTED, font=("Yu Gothic UI", 9)).pack(anchor="w")
            self._update_preview()
            return

        for arg in args:
            letter = arg["letter"]
            desc = arg.get("desc", "")
            row = tk.Frame(self._form_inner, bg=INPUT_BG)
            row.pack(fill="x", pady=4)
            tk.Label(row, text=letter, bg=INPUT_BG, fg=ACCENT,
                     font=("Consolas", 13, "bold"), width=2, anchor="w").pack(side="left")
            desc_label = tk.Label(row, text=desc, bg=INPUT_BG, fg=TEXT_MUTED,
                                  font=("Yu Gothic UI", 9), anchor="w", wraplength=280,
                                  justify="left")
            desc_label.pack(side="left", fill="x", expand=True, padx=(6, 8))
            var = tk.StringVar()
            var.trace_add("write", lambda *_a: self._update_preview())
            entry = tk.Entry(row, textvariable=var, font=("Consolas", 11),
                             bg=BG_PANEL, fg=TEXT_MAIN, insertbackground=ACCENT,
                             relief="solid", bd=1, highlightthickness=1,
                             highlightbackground=BORDER, highlightcolor=ACCENT, width=10)
            entry.pack(side="right")
            # 引数値は小数点必須：フォーカスが外れた時点で整数だけなら "." を自動付与
            entry.bind("<FocusOut>", lambda _e, v=var: v.set(ensure_decimal_value(v.get())))
            self._arg_vars[letter] = var

        self._update_preview()

    def _update_preview(self) -> None:
        if not self._selected:
            self._preview_var.set("")
            return
        p_number = re.sub(r"\D", "", self._selected["o_number"])
        parts = [f"G65P{p_number}"]
        for arg in self._selected.get("args", []):
            letter = arg["letter"]
            value = self._arg_vars.get(letter)
            text = ensure_decimal_value(value.get()) if value else ""
            if text:
                parts.append(f"{letter}{text}")
        self._preview_var.set("".join(parts))

    def _insert_call(self) -> None:
        line = self._preview_var.get().strip()
        if not self._selected:
            messagebox.showwarning("未選択", "先にマクロを選んでください。", parent=self)
            return
        if not line:
            messagebox.showwarning("入力なし", "引数を1つ以上入力してください。", parent=self)
            return
        self.app.insert_text_at_cursor(line + "\n")
        self.app.status_var.set(f"マクロ呼出し行を挿入しました: {line}")


class MacroDefEditDialog(tk.Toplevel):
    """マクロ定義（O番号・タイトル・引数一覧・複数行メモ）の作成・編集ダイアログ。
    自動解析（ファイル取込）結果のプリフィルにも、ゼロからの手動作成にも使う共通フォーム。
    保存は on_save(macro_def, original_o_number) コールバック経由でNcMacroAssistDialogへ渡す。
    """

    def __init__(self, app: "NcCheckApp", on_save, macro_def: dict | None = None,
                 original_o_number: str | None = None) -> None:
        super().__init__(app.root)
        self.app = app
        self._on_save = on_save
        self._original_o_number = original_o_number
        self._arg_rows: list[dict] = []

        self.title("マクロ定義の作成・編集")
        self.configure(bg=BG_PANEL)
        self.transient(app.root)
        self.resizable(True, True)
        self.geometry("620x680")
        self.minsize(520, 500)

        self._build_ui()
        if macro_def:
            self._load(macro_def)
        else:
            self._add_arg_row()

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        try:
            px = app.root.winfo_rootx() + (app.root.winfo_width() - self.winfo_width()) // 2
            py = app.root.winfo_rooty() + (app.root.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except tk.TclError:
            pass
        self.grab_set()

    def _build_ui(self) -> None:
        outer = tk.Frame(self, bg=BG_PANEL, padx=14, pady=12)
        outer.pack(fill="both", expand=True)

        tk.Label(outer, text="マクロ定義の作成・編集", bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w")
        tk.Label(outer, text="自動解析の結果もここで手動修正できます。",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9)).pack(anchor="w", pady=(2, 10))

        head = tk.Frame(outer, bg=BG_PANEL)
        head.pack(fill="x", pady=(0, 8))
        self._o_number_var = tk.StringVar()
        self._title_var = tk.StringVar()
        tk.Label(head, text="O番号", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).grid(row=0, column=0, sticky="w")
        tk.Label(head, text="タイトル", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).grid(row=0, column=1, sticky="w", padx=(8, 0))
        tk.Entry(head, textvariable=self._o_number_var, font=("Consolas", 11),
                 bg=INPUT_BG, fg=TEXT_MAIN, insertbackground=ACCENT, relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
                 width=10).grid(row=1, column=0, sticky="ew")
        tk.Entry(head, textvariable=self._title_var, font=("Yu Gothic UI", 10),
                 bg=INPUT_BG, fg=TEXT_MAIN, insertbackground=ACCENT, relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
                 ).grid(row=1, column=1, sticky="ew", padx=(8, 0))
        head.grid_columnconfigure(1, weight=1)

        # ===== メモ（複数行） =====
        tk.Label(outer, text="メモ（複数行可）", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).pack(anchor="w", pady=(4, 2))
        memo_wrap = tk.Frame(outer, bg=INPUT_BG, relief="solid", bd=1,
                             highlightthickness=1, highlightbackground=BORDER)
        memo_wrap.pack(fill="x", pady=(0, 10))
        self._memo_text = tk.Text(
            memo_wrap, height=4, wrap="word", bd=0, highlightthickness=0,
            bg=INPUT_BG, fg=TEXT_MAIN, insertbackground=ACCENT,
            font=("Yu Gothic UI", 10), padx=6, pady=6,
        )
        self._memo_text.pack(fill="x")

        # ===== 引数一覧 =====
        args_head = tk.Frame(outer, bg=BG_PANEL)
        args_head.pack(fill="x")
        tk.Label(args_head, text="引数一覧（記号＋説明）", bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8, "bold")).pack(side="left")
        tk.Button(
            args_head, text="＋ 引数を追加", command=self._add_arg_row,
            font=("Yu Gothic UI", 8), bg=INPUT_BG, fg=TEXT_MAIN,
            activebackground="#1F2B3A", activeforeground=TEXT_MAIN,
            relief="solid", bd=1, padx=8, pady=2, cursor="hand2",
        ).pack(side="right")

        args_wrap = tk.Frame(outer, bg=INPUT_BG, relief="solid", bd=1,
                             highlightthickness=1, highlightbackground=BORDER)
        args_wrap.pack(fill="both", expand=True, pady=(4, 10))
        args_canvas = tk.Canvas(args_wrap, bg=INPUT_BG, highlightthickness=0, bd=0)
        args_canvas.pack(side="left", fill="both", expand=True)
        args_vbar = tk.Scrollbar(args_wrap, orient="vertical", bg=BG_PANEL,
                                 troughcolor="#101720", activebackground=ACCENT,
                                 command=args_canvas.yview)
        args_vbar.pack(side="right", fill="y")
        args_canvas.configure(yscrollcommand=args_vbar.set)
        self._args_inner = tk.Frame(args_canvas, bg=INPUT_BG, padx=6, pady=6)
        args_window = args_canvas.create_window((0, 0), window=self._args_inner, anchor="nw")

        def _on_args_configure(_e=None):
            args_canvas.configure(scrollregion=args_canvas.bbox("all"))
        self._args_inner.bind("<Configure>", _on_args_configure)

        def _on_args_canvas_configure(event):
            args_canvas.itemconfigure(args_window, width=event.width)
        args_canvas.bind("<Configure>", _on_args_canvas_configure)

        footer = tk.Frame(outer, bg=BG_PANEL)
        footer.pack(fill="x")
        tk.Button(
            footer, text="保存", command=self._save,
            font=("Yu Gothic UI", 10, "bold"), bg=ACCENT, fg="#04231F",
            activebackground=ACCENT_DARK, activeforeground="#04231F",
            relief="solid", bd=1, padx=16, pady=6, cursor="hand2",
        ).pack(side="right")
        tk.Button(
            footer, text="キャンセル", command=self.destroy,
            font=("Yu Gothic UI", 9), bg=INPUT_BG, fg=TEXT_MAIN,
            activebackground="#1F2B3A", activeforeground=TEXT_MAIN,
            relief="solid", bd=1, padx=12, pady=6, cursor="hand2",
        ).pack(side="right", padx=(0, 8))

    def _load(self, macro_def: dict) -> None:
        self._o_number_var.set(macro_def.get("o_number", ""))
        self._title_var.set(macro_def.get("title", ""))
        self._memo_text.delete("1.0", tk.END)
        self._memo_text.insert("1.0", macro_def.get("memo", ""))
        args = macro_def.get("args", [])
        for arg in args:
            self._add_arg_row(arg.get("letter", ""), arg.get("desc", ""))
        if not args:
            self._add_arg_row()

    def _add_arg_row(self, letter: str = "", desc: str = "") -> None:
        row = tk.Frame(self._args_inner, bg=INPUT_BG)
        row.pack(fill="x", pady=2)
        letter_var = tk.StringVar(value=letter)
        desc_var = tk.StringVar(value=desc)
        tk.Entry(row, textvariable=letter_var, font=("Consolas", 11, "bold"),
                 bg=BG_PANEL, fg=ACCENT, insertbackground=ACCENT, relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
                 width=3, justify="center").pack(side="left")
        tk.Entry(row, textvariable=desc_var, font=("Yu Gothic UI", 10),
                 bg=BG_PANEL, fg=TEXT_MAIN, insertbackground=ACCENT, relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
                 ).pack(side="left", fill="x", expand=True, padx=(6, 6))
        row_info = {"frame": row, "letter_var": letter_var, "desc_var": desc_var}
        tk.Button(
            row, text="✕", command=lambda: self._remove_arg_row(row_info),
            font=("Yu Gothic UI", 8), bg=INPUT_BG, fg="#C97C8A",
            activebackground="#452030", activeforeground="#E0A0AE",
            relief="solid", bd=1, padx=6, pady=1, cursor="hand2",
        ).pack(side="right")
        self._arg_rows.append(row_info)

    def _remove_arg_row(self, row_info: dict) -> None:
        row_info["frame"].destroy()
        self._arg_rows.remove(row_info)

    def _save(self) -> None:
        o_number = self._o_number_var.get().strip().upper()
        if not o_number:
            messagebox.showwarning("入力不足", "O番号を入力してください。", parent=self)
            return
        if not re.match(r"^O?\d+$", o_number):
            messagebox.showwarning("形式エラー", "O番号は「O7021」のように入力してください。", parent=self)
            return
        if not o_number.startswith("O"):
            o_number = "O" + o_number

        args = []
        for row_info in self._arg_rows:
            letter = row_info["letter_var"].get().strip().upper()
            desc = row_info["desc_var"].get().strip()
            if letter:
                args.append({"letter": letter, "desc": desc})

        macro_def = {
            "o_number": o_number,
            "title": self._title_var.get().strip(),
            "args": args,
            "memo": self._memo_text.get("1.0", "end-1c"),
            "raw_text": "",
        }
        self._on_save(macro_def, self._original_o_number)
        self.destroy()


def _re_match_t(s: str) -> bool:
    """T番号形式の検証（T+数字）"""
    return bool(re.fullmatch(r"T\d+", s))


class TplTargetSelector:
    """テンプレート用の対象範囲セレクタ。
    全体／選択範囲／指定したNブロック の3択ラジオと、
    Nブロック指定時のドロップダウンを管理する。
    """

    OPTIONS = ["全体", "選択範囲", "指定したNブロック"]

    def __init__(self, dialog, parent):
        self.dialog = dialog  # FindReplaceDialog
        self._var = tk.StringVar(value="全体")
        self._n_block_label = ""  # 選択中のNラベル（"N100"等）
        self._n_block_options: list[str] = []  # ドロップダウン表示用文字列リスト

        # 行1：ラベル＋ラジオ
        row1 = tk.Frame(parent, bg=BG_PANEL)
        row1.pack(fill="x", pady=4)
        tk.Label(row1, text="対象:", bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="e").pack(side="left", padx=(0, 8))
        rb_kw = {"bg": BG_PANEL, "fg": TEXT_MAIN, "selectcolor": INPUT_BG,
                 "activebackground": BG_PANEL, "activeforeground": ACCENT,
                 "font": ("Yu Gothic UI", 10), "bd": 0, "highlightthickness": 0}
        for opt in self.OPTIONS:
            tk.Radiobutton(row1, text=opt, variable=self._var, value=opt,
                           command=self._on_radio_changed, **rb_kw).pack(side="left", padx=(0, 8))

        # 行2：Nブロックドロップダウン（初期非表示）
        self._n_row = tk.Frame(parent, bg=BG_PANEL)
        # まだpackしない（_on_radio_changedで制御）
        tk.Label(self._n_row, text="Nブロック:", bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="e").pack(side="left", padx=(0, 8))

        # tk.OptionMenuベースのドロップダウン
        self._n_var = tk.StringVar(value="（チェック実行後に選択可能）")
        self._n_optionmenu = tk.OptionMenu(self._n_row, self._n_var, self._n_var.get())
        self._n_optionmenu.configure(
            bg=INPUT_BG, fg=ACCENT, font=("Consolas", 10),
            activebackground="#123A38", activeforeground=ACCENT,
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
            relief="solid", bd=1, anchor="w", width=40,
        )
        self._n_optionmenu["menu"].configure(
            bg=INPUT_BG, fg=TEXT_MAIN, font=("Consolas", 10),
            activebackground="#123A38", activeforeground=ACCENT,
        )
        self._n_optionmenu.pack(side="left", fill="x", expand=True)

    def _on_radio_changed(self):
        """ラジオ切り替え時：Nブロック選択なら再スキャン＆表示、それ以外は隠す"""
        if self._var.get() == "指定したNブロック":
            self._refresh_n_block_options()
            if not self._n_row.winfo_ismapped():
                self._n_row.pack(fill="x", pady=4)
        else:
            if self._n_row.winfo_ismapped():
                self._n_row.pack_forget()

    def _refresh_n_block_options(self):
        """プログラム本文をスキャンしてNブロック一覧を再構築。
        表示形式: "N100  T2001  (DCLNN R0.4)"
        T番号やコメントが無い場合は "N100" のみ。
        """
        target = self.dialog.target  # tk.Text widget
        text = target.get("1.0", "end-1c")
        lines = text.splitlines()

        # NBlockに分割（既存ロジックを簡易再現）
        n_blocks = self._scan_n_blocks(lines)

        # 各ブロックのT番号と工具コメントを抽出
        options = []
        for n_label, block_rows in n_blocks:
            t_num, _ = extract_t_before_g36x(block_rows)
            tool_name, _ = extract_tool_name(block_rows)
            parts = [n_label]
            if t_num:
                parts.append(t_num)
            if tool_name:
                # 長すぎるコメントは省略
                tn = tool_name if len(tool_name) <= 30 else tool_name[:27] + "..."
                parts.append(tn)
            options.append("  ".join(parts))

        self._n_block_options = options

        # OptionMenuの中身を差し替え
        menu = self._n_optionmenu["menu"]
        menu.delete(0, "end")
        if not options:
            self._n_var.set("（Nブロックが見つかりません）")
            menu.add_command(
                label="（Nブロックが見つかりません）",
                command=lambda: self._n_var.set("（Nブロックが見つかりません）"),
            )
        else:
            self._n_var.set(options[0])
            for opt in options:
                menu.add_command(
                    label=opt,
                    command=lambda v=opt: self._n_var.set(v),
                )

    def _scan_n_blocks(self, lines: list[str]) -> list[tuple[str, list[tuple[int, str]]]]:
        """行リストからNブロックを抽出。
        Returns: [(n_label, [(line_no, line), ...]), ...]
        """
        blocks = []
        current_n = None
        current_rows = []
        for i, line in enumerate(lines, start=1):
            m = re.match(r"^\s*(N\d+)\b", line)
            if m:
                # 新しいNブロック開始
                if current_n is not None:
                    blocks.append((current_n, current_rows))
                current_n = m.group(1)
                current_rows = [(i, line)]
            else:
                if current_n is not None:
                    current_rows.append((i, line))
        if current_n is not None:
            blocks.append((current_n, current_rows))
        return blocks

    def get(self) -> str:
        """選択中の対象範囲名を返す（既存互換）"""
        return self._var.get()

    def get_n_block_label(self) -> str:
        """選択中のNラベル（"N100"等）を返す。Nブロック選択モードで有効。"""
        display = self._n_var.get()
        m = re.match(r"^(N\d+)", display)
        return m.group(1) if m else ""


class FindReplaceDialog(tk.Toplevel):
    """検索/置換ダイアログ。Ghost Protocolテーマ。
    Ctrl+F: 検索モード（置換欄非表示）
    Ctrl+H: 置換モード（置換欄表示）
    """

    def __init__(self, app, mode: str = "find") -> None:
        super().__init__(app.root)
        self.app = app
        self.target = app.input_text
        self.mode = mode  # "find" | "replace"

        self.title("検索 / 置換  //  FIND & REPLACE")
        self.configure(bg=BG_PANEL)
        self.transient(app.root)
        # リサイズ可能にして、初期サイズと最小サイズを指定
        self.resizable(True, True)
        self.geometry("760x880")
        self.minsize(680, 520)

        # 検索状態
        self._last_search_end: str = "1.0"
        self._search_var = tk.StringVar()
        self._replace_var = tk.StringVar()
        self._regex_var = tk.StringVar(value="0")
        self._case_var = tk.StringVar(value="0")

        self._build_ui()
        self._setup_tags()
        self._bind_keys()

        # 選択範囲があればそれを検索文字に初期セット
        try:
            sel = self.target.get("sel.first", "sel.last")
            if sel and "\n" not in sel:
                self._search_var.set(sel)
        except tk.TclError:
            pass

        # 中央寄せ
        self.update_idletasks()
        self._center_on_parent()
        try:
            self._search_text.focus_set()
            self._search_text.tag_add("sel", "1.0", "end-1c")
        except (tk.TclError, AttributeError):
            pass

        # 閉じる時にハイライト消す
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        outer = tk.Frame(self, bg=BG_PANEL, padx=14, pady=12)
        outer.pack(fill="both", expand=True)

        # ヘッダ
        title_text = "▶ FIND" if self.mode == "find" else "▶ FIND & REPLACE"
        tk.Label(outer, text=title_text, bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w")
        tk.Label(outer, text="検索／置換／テンプレート",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(anchor="w", pady=(0, 8))

        # タブ切替ボタン群
        tab_bar = tk.Frame(outer, bg=BG_PANEL)
        tab_bar.pack(fill="x", pady=(0, 8))
        self._current_tab = tk.StringVar(value="simple")
        self._tab_buttons = {}
        for key, label in [("simple", "シンプル置換"), ("template", "テンプレート"), ("macro", "マクロ")]:
            btn = tk.Button(
                tab_bar, text=label,
                command=lambda k=key: self._switch_tab(k),
                font=("Yu Gothic UI", 10, "bold"),
                bg="#141B24", fg=TEXT_MUTED,
                activebackground="#1F2B3A", activeforeground=ACCENT,
                relief="solid", bd=1, padx=14, pady=4, cursor="hand2",
            )
            btn.pack(side="left", padx=(0, 4))
            self._tab_buttons[key] = btn

        # タブ内容を入れるコンテナ
        self._tab_container = tk.Frame(outer, bg=BG_PANEL)
        self._tab_container.pack(fill="both", expand=True)

        # 各タブのフレーム作成
        self._tab_frames = {
            "simple": tk.Frame(self._tab_container, bg=BG_PANEL),
            "template": tk.Frame(self._tab_container, bg=BG_PANEL),
            "macro": tk.Frame(self._tab_container, bg=BG_PANEL),
        }
        self._build_simple_tab(self._tab_frames["simple"])
        self._build_template_tab(self._tab_frames["template"])
        self._build_macro_tab(self._tab_frames["macro"])

        # ステータス表示（共通、最下部）
        self._status_var = tk.StringVar(value="")
        tk.Label(outer, textvariable=self._status_var,
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Consolas", 9), anchor="w").pack(fill="x", pady=(8, 0))

        # 閉じるボタン（共通）
        close_frame = tk.Frame(outer, bg=BG_PANEL)
        close_frame.pack(fill="x", pady=(8, 0))
        btn_kw = {"font": ("Yu Gothic UI", 10, "bold"),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "solid", "bd": 1, "padx": 14, "pady": 4,
                  "cursor": "hand2"}
        tk.Button(close_frame, text="閉じる", command=self._on_close,
                  **btn_kw).pack(side="right")

        # 初期タブ表示
        self._switch_tab("simple")

    def _switch_tab(self, tab_key: str) -> None:
        """タブ切替"""
        for key, frame in self._tab_frames.items():
            if key == tab_key:
                frame.pack(fill="both", expand=True)
            else:
                frame.pack_forget()
        # ボタンの見た目更新（アクティブタブを明るく）
        for key, btn in self._tab_buttons.items():
            if key == tab_key:
                btn.configure(bg="#123A38", fg=ACCENT)
            else:
                btn.configure(bg="#141B24", fg=TEXT_MUTED)
        self._current_tab.set(tab_key)

    # =========================================================================
    # Undo統合用ヘルパー：Tk Text の Undo スタックに「1操作」として記録
    # =========================================================================
    def _atomic_replace(self, text_widget, start: str, end: str, new_text: str) -> None:
        """delete + insert を1つのUndoとして記録する。
        Ctrl+Z 1回で確実に戻せるようにするため。
        """
        try:
            text_widget.config(autoseparators=False)
            text_widget.edit_separator()
            text_widget.delete(start, end)
            text_widget.insert(start, new_text)
            text_widget.edit_separator()
        finally:
            try:
                text_widget.config(autoseparators=True)
            except tk.TclError:
                pass

    def _show_regex_builder(self, initial_pattern: str, ignore_case: bool,
                              on_result, status_msg: str = "正規表現ビルダー: パターンを設定しました") -> None:
        """正規表現ビルダーを起動する共通処理。
        - 選択範囲があればテスト文字列の初期値に設定
        - ビルダー起動（モーダル）
        - 結果あれば on_result(pattern) を呼ぶ
        - ステータスメッセージを表示

        Args:
            initial_pattern: ビルダーに初期表示するパターン文字列
            ignore_case: 大文字小文字を無視するか
            on_result: callable(result_pattern: str) -> None。結果適用コールバック
            status_msg: 成功時にステータスバーへ表示するメッセージ
        """
        # メイン画面の選択範囲があればテスト文字列の初期値に
        initial_test = ""
        try:
            sel = self.target.get("sel.first", "sel.last")
            if sel:
                initial_test = sel
        except tk.TclError:
            pass

        builder = RegexBuilderDialog(self,
                                      initial_pattern=initial_pattern,
                                      initial_test_text=initial_test,
                                      initial_ignore_case=ignore_case)
        builder.transient(self)
        builder.grab_set()
        self.wait_window(builder)

        if builder.result_pattern is not None:
            on_result(builder.result_pattern)
            self._status_var.set(status_msg)

    def _open_regex_builder_for_simple(self) -> None:
        """シンプル置換タブの検索欄に対してビルダーを開く"""
        def apply_result(pattern: str):
            self._search_text.delete("1.0", "end")
            self._search_text.insert("1.0", pattern)
            self._search_var.set(pattern)
            self._regex_var.set("1")

        self._show_regex_builder(
            initial_pattern=self._search_var.get(),
            ignore_case=(self._case_var.get() == "0"),
            on_result=apply_result,
        )

    def _build_simple_tab(self, parent: tk.Frame) -> None:
        """シンプル置換タブの中身（v30〜複数行Text対応：改行込みの検索・置換が可能）"""
        # 検索文字エリア（Text、複数行）
        sf = tk.Frame(parent, bg=BG_PANEL)
        sf.pack(fill="x", pady=(8, 6))
        tk.Label(sf, text="検索:", bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=8, anchor="ne").pack(side="left", padx=(0, 8), pady=(2, 0))

        search_wrap = tk.Frame(sf, bg=BG_PANEL,
                                highlightbackground=BORDER, highlightthickness=1)
        search_wrap.pack(side="left", fill="x", expand=True)
        self._search_text = tk.Text(
            search_wrap, height=3, font=("Consolas", 11),
            bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT,
            relief="flat", bd=0, wrap="none",
            padx=4, pady=2, undo=True,
        )
        self._search_text.pack(side="left", fill="both", expand=True)
        # 既存 _search_var との同期：Text 内容を変えたら StringVar も更新
        def _on_search_text_changed(event=None):
            try:
                val = self._search_text.get("1.0", "end-1c")
                self._search_var.set(val)
            except tk.TclError:
                pass
        self._search_text.bind("<KeyRelease>", _on_search_text_changed)
        self._search_text.bind("<<Modified>>", lambda e: (
            _on_search_text_changed(),
            self._search_text.edit_modified(False)
        ))

        # 既存の _search_var 初期値を Text に流し込む（前回開いた状態の復元）
        try:
            initial = self._search_var.get()
            if initial:
                self._search_text.insert("1.0", initial)
        except tk.TclError:
            pass

        # 置換文字エリア（置換モード時のみ）
        if self.mode == "replace":
            rf = tk.Frame(parent, bg=BG_PANEL)
            rf.pack(fill="x", pady=(0, 6))
            tk.Label(rf, text="置換:", bg=BG_PANEL, fg=TEXT_MAIN,
                     font=("Yu Gothic UI", 10), width=8, anchor="ne").pack(side="left", padx=(0, 8), pady=(2, 0))

            replace_wrap = tk.Frame(rf, bg=BG_PANEL,
                                     highlightbackground=BORDER, highlightthickness=1)
            replace_wrap.pack(side="left", fill="x", expand=True)
            self._replace_text = tk.Text(
                replace_wrap, height=3, font=("Consolas", 11),
                bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT,
                relief="flat", bd=0, wrap="none",
                padx=4, pady=2, undo=True,
            )
            self._replace_text.pack(side="left", fill="both", expand=True)
            def _on_replace_text_changed(event=None):
                try:
                    val = self._replace_text.get("1.0", "end-1c")
                    self._replace_var.set(val)
                except tk.TclError:
                    pass
            self._replace_text.bind("<KeyRelease>", _on_replace_text_changed)
            self._replace_text.bind("<<Modified>>", lambda e: (
                _on_replace_text_changed(),
                self._replace_text.edit_modified(False)
            ))
            try:
                initial = self._replace_var.get()
                if initial:
                    self._replace_text.insert("1.0", initial)
            except tk.TclError:
                pass

        # 改行入力ヒント
        tk.Label(parent,
                 text="※ 改行を含む検索/置換が可能。Enterキーで改行入力（実行は下のボタンから）",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 8),
                 anchor="w").pack(fill="x", padx=(60, 0), pady=(0, 4))

        # オプション（チェックボックス）
        opt_frame = tk.Frame(parent, bg=BG_PANEL)
        opt_frame.pack(fill="x", pady=(4, 10))
        tk.Label(opt_frame, text="", bg=BG_PANEL, width=8).pack(side="left", padx=(0, 8))
        cb_kw = {"bg": BG_PANEL, "fg": TEXT_MAIN, "selectcolor": INPUT_BG,
                 "activebackground": BG_PANEL, "activeforeground": ACCENT,
                 "font": ("Yu Gothic UI", 9), "bd": 0, "highlightthickness": 0}
        tk.Checkbutton(opt_frame, text="正規表現", variable=self._regex_var,
                       onvalue="1", offvalue="0", **cb_kw).pack(side="left", padx=(0, 6))
        tk.Button(opt_frame, text="🔧 ビルダー",
                  command=self._open_regex_builder_for_simple,
                  font=("Yu Gothic UI", 9),
                  bg="#141B24", fg=ACCENT,
                  activebackground="#1F2B3A", activeforeground=ACCENT,
                  relief="solid", bd=1, padx=8, pady=1,
                  cursor="hand2").pack(side="left", padx=(0, 16))
        tk.Checkbutton(opt_frame, text="大文字小文字を区別", variable=self._case_var,
                       onvalue="1", offvalue="0", **cb_kw).pack(side="left")

        # 対象範囲セレクタ（「すべて置換」専用）
        if self.mode == "replace":
            scope_wrap = tk.Frame(parent, bg=BG_PANEL)
            scope_wrap.pack(fill="x", pady=(0, 6))
            self._simple_scope_selector = TplTargetSelector(self, scope_wrap)

            tk.Label(parent,
                     text="※ 対象範囲は「すべて置換」のみに適用されます\n"
                          "※ 「次を検索」「置換して次へ」は常に全体を対象にします",
                     bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 8),
                     wraplength=500, justify="left").pack(anchor="w", padx=(60, 0), pady=(2, 6))

            # 範囲指定エリア（折りたたみ式）
            self._range_expanded_var = tk.StringVar(value="0")
            range_toggle_row = tk.Frame(parent, bg=BG_PANEL)
            range_toggle_row.pack(fill="x", pady=(4, 0))
            tk.Label(range_toggle_row, text="", bg=BG_PANEL, width=8).pack(side="left", padx=(0, 8))
            self._range_toggle_btn = tk.Button(
                range_toggle_row,
                text="▶ 範囲指定（任意）",
                command=self._toggle_range_panel,
                font=("Yu Gothic UI", 9, "bold"),
                bg="#141B24", fg=ACCENT,
                activebackground="#1F2B3A", activeforeground=ACCENT,
                relief="solid", bd=1, padx=10, pady=2, cursor="hand2",
                anchor="w",
            )
            self._range_toggle_btn.pack(side="left")

            # 範囲指定の本体（初期は非表示）
            self._range_panel = tk.Frame(parent, bg=BG_PANEL)
            # pack はトグルで制御するのでここではしない

            # 境界マッチオプション
            self._range_regex_var = tk.StringVar(value="0")
            range_opt_row = tk.Frame(self._range_panel, bg=BG_PANEL)
            range_opt_row.pack(fill="x", pady=(6, 2))
            tk.Label(range_opt_row, text="", bg=BG_PANEL, width=8).pack(side="left", padx=(0, 8))
            tk.Checkbutton(
                range_opt_row, text="境界も正規表現で指定する",
                variable=self._range_regex_var, onvalue="1", offvalue="0",
                bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
                activebackground=BG_PANEL, activeforeground=ACCENT,
                font=("Yu Gothic UI", 9), bd=0, highlightthickness=0,
            ).pack(side="left")

            # 4つの境界入力欄（ヘルパー使う）
            self._range_entries = {}
            for key, label in [
                ("outer_start", "外側 開始:"),
                ("outer_end",   "外側 終端:"),
                ("inner_start", "内側 開始:"),
                ("inner_end",   "内側 終端:"),
            ]:
                self._range_entries[key] = self._make_range_entry_row(
                    self._range_panel, label, key
                )

            # 説明
            tk.Label(self._range_panel,
                     text="※ 範囲指定なし: 対象範囲全体で置換\n"
                          "※ 外側のみ指定: 外側範囲内のみ置換（境界行は除外）\n"
                          "※ 外側＋内側指定: 外側内にある内側範囲のみで置換\n"
                          "※ 終端が見つからない範囲は無視されます\n"
                          "※ 「境界も正規表現」ONで、境界マーカーも正規表現で書けます",
                     bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 8),
                     wraplength=500, justify="left").pack(anchor="w", padx=(60, 0), pady=(4, 4))

        # ボタン群
        btn_frame = tk.Frame(parent, bg=BG_PANEL)
        btn_frame.pack(fill="x", pady=(4, 0))

        btn_kw = {"font": ("Yu Gothic UI", 10, "bold"),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "solid", "bd": 1, "padx": 14, "pady": 4,
                  "cursor": "hand2"}
        btn_kw_primary = dict(btn_kw)
        btn_kw_primary["bg"] = "#123A38"
        btn_kw_primary["fg"] = ACCENT

        tk.Button(btn_frame, text="次を検索", command=self.find_next,
                  **btn_kw_primary).pack(side="left", padx=(0, 6))

        if self.mode == "replace":
            tk.Button(btn_frame, text="置換して次へ", command=self.replace_one,
                      **btn_kw).pack(side="left", padx=(0, 6))
            tk.Button(btn_frame, text="すべて置換", command=self.replace_all,
                      **btn_kw).pack(side="left", padx=(0, 6))

    def _toggle_range_panel(self) -> None:
        """範囲指定パネルの折りたたみ切り替え"""
        if not hasattr(self, "_range_panel"):
            return
        expanded = self._range_expanded_var.get() == "1"
        if expanded:
            # 閉じる
            self._range_panel.pack_forget()
            self._range_expanded_var.set("0")
            self._range_toggle_btn.configure(text="▶ 範囲指定（任意）")
        else:
            # 開く
            self._range_panel.pack(fill="x", pady=(2, 6))
            self._range_expanded_var.set("1")
            self._range_toggle_btn.configure(text="▼ 範囲指定（任意）")

    def _make_range_entry_row(self, parent: tk.Widget, label: str, key: str) -> tk.Text:
        """範囲指定の入力欄1行を作成（ラベル + Entry + ビルダーボタン）"""
        row = tk.Frame(parent, bg=BG_PANEL)
        row.pack(fill="x", pady=2)
        tk.Label(row, text=label, bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 9), width=10, anchor="e").pack(side="left", padx=(0, 6))

        entry_wrap = tk.Frame(row, bg=BG_PANEL,
                               highlightbackground=BORDER, highlightthickness=1)
        entry_wrap.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ent = tk.Entry(entry_wrap, font=("Consolas", 10),
                       bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT,
                       relief="flat", bd=0)
        ent.pack(fill="both", expand=True, padx=3, pady=1)

        # ビルダーボタン（境界が正規表現モード時のみ意味あるが常に表示）
        tk.Button(row, text="🔧 ビルダー",
                  command=lambda k=key: self._open_regex_builder_for_range(k),
                  font=("Yu Gothic UI", 9),
                  bg="#141B24", fg=ACCENT,
                  activebackground="#1F2B3A", activeforeground=ACCENT,
                  relief="solid", bd=1, padx=8, pady=1,
                  cursor="hand2").pack(side="left")
        return ent

    def _open_regex_builder_for_range(self, key: str) -> None:
        """範囲指定の境界欄に対する正規表現ビルダー起動"""
        if not hasattr(self, "_range_entries"):
            return
        entry = self._range_entries.get(key)
        if entry is None:
            return

        label_map = {
            "outer_start": "外側 開始",
            "outer_end":   "外側 終端",
            "inner_start": "内側 開始",
            "inner_end":   "内側 終端",
        }

        def apply_result(pattern: str):
            entry.delete(0, "end")
            entry.insert(0, pattern)
            # 結果を流し込んだら境界正規表現モードを自動ON
            self._range_regex_var.set("1")

        self._show_regex_builder(
            initial_pattern=entry.get(),
            ignore_case=(self._case_var.get() == "0"),
            on_result=apply_result,
            status_msg=f"正規表現ビルダー: {label_map.get(key, key)} にパターンを設定（境界正規表現ON）",
        )

    def _get_range_spec(self) -> dict | None:
        """シンプル置換タブの範囲指定値を取得し、整合性チェックして返す。
        Returns:
            None: 整合性エラー（ステータス設定済み）
            {}: 範囲指定なし（全行対象）
            {"outer_s":..., "outer_e":..., "inner_s":..., "inner_e":..., "regex":bool}: 通常
        """
        if not hasattr(self, "_range_entries"):
            return {}
        # 折りたたみが閉じてる場合は範囲指定無効として扱う
        if self._range_expanded_var.get() != "1":
            return {}

        outer_s = self._range_entries["outer_start"].get().strip()
        outer_e = self._range_entries["outer_end"].get().strip()
        inner_s = self._range_entries["inner_start"].get().strip()
        inner_e = self._range_entries["inner_end"].get().strip()

        # 整合性チェック
        if (outer_s and not outer_e) or (outer_e and not outer_s):
            self._status_var.set("⚠ 外側範囲は開始と終端を両方指定してください")
            return None
        if (inner_s and not inner_e) or (inner_e and not inner_s):
            self._status_var.set("⚠ 内側範囲は開始と終端を両方指定してください")
            return None
        if (inner_s or inner_e) and not (outer_s and outer_e):
            self._status_var.set("⚠ 内側範囲を使うには外側範囲の指定が必要です")
            return None

        if not outer_s:
            return {}  # 範囲指定なし

        return {
            "outer_s": outer_s, "outer_e": outer_e,
            "inner_s": inner_s, "inner_e": inner_e,
            "regex": self._range_regex_var.get() == "1",
        }

    def _build_template_tab(self, parent: tk.Frame) -> None:
        """テンプレートタブ：左にカテゴリ＆テンプレ一覧、右に詳細パネル"""
        # 17個のテンプレ定義
        # (id, label, category, description, builder_method_name)
        self._templates = [
            # 数値変更
            ("f_scale",      "F値倍率変更",       "数値変更",   "全F値に倍率を掛ける（例：0.8倍で安全側）"),
            ("f_uniform",    "F値統一",           "数値変更",   "全F値を指定値に置換"),
            ("s_scale",      "S値倍率変更",       "数値変更",   "全S値に倍率を掛ける"),
            ("s_uniform",    "S値統一",           "数値変更",   "全S値を指定値に置換"),
            # 座標操作
            ("axis_offset",  "座標オフセット加算", "座標操作",   "指定軸（X/Y/Z）の値に加算"),
            ("axis_invert",  "座標符号反転",       "座標操作",   "指定軸の値の符号を反転（裏面加工転用）"),
            # 形式整形
            ("decimal_add",  "小数点付加",         "形式整形",   "小数点なしの整数指令に「.」を追加"),
            ("upper_case",   "小文字→大文字統一", "形式整形",   "NCコマンド全体を大文字に統一"),
            ("trim_spaces",  "余分な空白除去",     "形式整形",   "行頭末尾の空白を除去"),
            # コメント
            ("comment_add",  "コメント化",         "コメント",   "対象範囲の行を ( ) で囲む"),
            ("comment_del",  "コメント解除",       "コメント",   "( ) を外して通常行に戻す"),
            ("comment_remove", "コメント削除",     "コメント",   "行内の ( ) コメントをすべて削除"),
            # 入替
            ("char_swap",    "文字入替",           "入替",       "A↔B の入替（範囲指定2段ネスト対応）"),
            # 行操作
            ("line_delete",  "行削除",             "行操作",     "空行・パターン一致行を削除"),
        ]

        # 左ペイン：カテゴリ別テンプレ一覧（スクロール対応）
        left = tk.Frame(parent, bg=BG_PANEL, width=240)
        left.pack(side="left", fill="y", padx=(0, 10), pady=(8, 0))
        left.pack_propagate(False)

        tk.Label(left, text="テンプレート一覧",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        # Canvas + Scrollbar でスクロール可能なフレームを構築
        canvas_wrap = tk.Frame(left, bg=BG_PANEL)
        canvas_wrap.pack(fill="both", expand=True)

        list_canvas = tk.Canvas(canvas_wrap, bg=BG_PANEL, bd=0,
                                highlightthickness=0, takefocus=False)
        self._tpl_list_canvas = list_canvas
        list_scrollbar = tk.Scrollbar(canvas_wrap, orient="vertical",
                                       command=list_canvas.yview,
                                       bg=BG_PANEL, troughcolor=BG_HEADER,
                                       activebackground=ACCENT_DARK,
                                       bd=0, highlightthickness=0,
                                       width=10)
        list_canvas.configure(yscrollcommand=list_scrollbar.set)

        list_scrollbar.pack(side="right", fill="y")
        list_canvas.pack(side="left", fill="both", expand=True)

        list_inner = tk.Frame(list_canvas, bg=BG_PANEL)
        list_inner_id = list_canvas.create_window((0, 0), window=list_inner, anchor="nw")

        # 内側フレームのサイズ変化に追従してscrollregion更新
        def _on_inner_configure(event):
            list_canvas.configure(scrollregion=list_canvas.bbox("all"))
        list_inner.bind("<Configure>", _on_inner_configure)

        # Canvas幅変化時に内側フレーム幅を合わせる
        def _on_canvas_configure(event):
            list_canvas.itemconfigure(list_inner_id, width=event.width)
        list_canvas.bind("<Configure>", _on_canvas_configure)

        # マウスホイール対応（カーソルがリスト上にある時のみ）
        def _on_mousewheel(event):
            list_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _bind_wheel(_e):
            list_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_wheel(_e):
            list_canvas.unbind_all("<MouseWheel>")
        list_canvas.bind("<Enter>", _bind_wheel)
        list_canvas.bind("<Leave>", _unbind_wheel)

        # カテゴリごとにグループ化して表示
        self._template_buttons: dict = {}
        last_category = None
        for tpl_id, label, category, _desc in self._templates:
            if category != last_category:
                tk.Label(list_inner, text=f"▼ {category}",
                         bg=BG_PANEL, fg=ACCENT_DARK,
                         font=("Yu Gothic UI", 9, "bold")).pack(anchor="w", pady=(6, 2), padx=4)
                last_category = category
            btn = tk.Button(
                list_inner, text=f"  {label}",
                command=lambda tid=tpl_id: self._select_template(tid),
                font=("Yu Gothic UI", 9), bg=BG_PANEL, fg=TEXT_MAIN,
                activebackground="#123A38", activeforeground=ACCENT,
                relief="flat", bd=0, anchor="w", padx=8, pady=2, cursor="hand2",
            )
            btn.pack(fill="x", pady=1)
            self._template_buttons[tpl_id] = btn

        # 右ペイン：選択中テンプレの詳細＆実行
        self._template_detail_frame = tk.Frame(parent, bg=BG_PANEL)
        self._template_detail_frame.pack(side="left", fill="both", expand=True, pady=(8, 0))
        self._show_template_placeholder()

    # =========================================================================
    # マクロタブ（v28〜）
    # =========================================================================

    def _build_macro_tab(self, parent: tk.Frame) -> None:
        """マクロタブ：左にマクロ一覧、右に編集＆実行パネル"""
        # マクロ機能の状態保持
        self._current_macro: Macro = Macro()  # 編集中のマクロ
        self._macro_step_widgets: list = []   # ステップ行ウィジェットのリスト
        self._macro_undo_snapshot: str | None = None  # Undo用スナップショット
        self._macro_dirty: bool = False       # 未保存変更フラグ

        # 左ペイン：マクロ一覧
        left = tk.Frame(parent, bg=BG_PANEL, width=240)
        left.pack(side="left", fill="y", padx=(0, 10), pady=(8, 0))
        left.pack_propagate(False)

        tk.Label(left, text="マクロ一覧",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        # 操作ボタン群（新規／更新）
        btn_kw = {"font": ("Yu Gothic UI", 9, "bold"),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "solid", "bd": 1, "padx": 8, "pady": 3,
                  "cursor": "hand2"}

        new_btn = tk.Button(left, text="＋ 新規作成",
                            command=self._macro_new,
                            **btn_kw)
        new_btn.pack(fill="x", pady=(0, 4))

        refresh_btn = tk.Button(left, text="↻ 一覧を更新",
                                command=self._macro_refresh_list,
                                **btn_kw)
        refresh_btn.pack(fill="x", pady=(0, 8))

        # 区切り
        tk.Frame(left, bg=BORDER, height=1).pack(fill="x", pady=(0, 4))
        tk.Label(left, text="▼ 保存済みマクロ",
                 bg=BG_PANEL, fg=ACCENT_DARK,
                 font=("Yu Gothic UI", 9, "bold")).pack(anchor="w", pady=(2, 4), padx=2)

        # スクロール対応リスト
        list_wrap = tk.Frame(left, bg=BG_PANEL)
        list_wrap.pack(fill="both", expand=True)

        self._macro_list_canvas = tk.Canvas(list_wrap, bg=BG_PANEL, bd=0,
                                             highlightthickness=0, takefocus=False)
        macro_scroll = tk.Scrollbar(list_wrap, orient="vertical",
                                     command=self._macro_list_canvas.yview,
                                     bg=BG_PANEL, troughcolor=BG_HEADER,
                                     activebackground=ACCENT_DARK,
                                     bd=0, highlightthickness=0, width=10)
        self._macro_list_canvas.configure(yscrollcommand=macro_scroll.set)
        macro_scroll.pack(side="right", fill="y")
        self._macro_list_canvas.pack(side="left", fill="both", expand=True)

        self._macro_list_inner = tk.Frame(self._macro_list_canvas, bg=BG_PANEL)
        self._macro_list_inner_id = self._macro_list_canvas.create_window(
            (0, 0), window=self._macro_list_inner, anchor="nw"
        )

        def _on_macro_inner_configure(event):
            self._macro_list_canvas.configure(
                scrollregion=self._macro_list_canvas.bbox("all")
            )
        self._macro_list_inner.bind("<Configure>", _on_macro_inner_configure)

        def _on_macro_canvas_configure(event):
            self._macro_list_canvas.itemconfigure(
                self._macro_list_inner_id, width=event.width
            )
        self._macro_list_canvas.bind("<Configure>", _on_macro_canvas_configure)

        # フォルダを開くボタン（最下部）
        folder_btn = tk.Button(left, text="📁 フォルダを開く",
                               command=self._macro_open_folder,
                               **btn_kw)
        folder_btn.pack(fill="x", pady=(8, 0))

        # 右ペイン：編集＆実行
        self._macro_detail_frame = tk.Frame(parent, bg=BG_PANEL)
        self._macro_detail_frame.pack(side="left", fill="both", expand=True, pady=(8, 0))

        self._build_macro_detail_panel(self._macro_detail_frame)

        # 初回読込
        self._macro_refresh_list()
        self._macro_load_into_panel(Macro())

    def _build_macro_detail_panel(self, parent: tk.Frame) -> None:
        """右ペイン：マクロ編集UI"""
        # ヘッダ
        tk.Label(parent, text="▶ マクロ編集",
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w", pady=(0, 8))

        # 名前
        name_row = tk.Frame(parent, bg=BG_PANEL)
        name_row.pack(fill="x", pady=2)
        tk.Label(name_row, text="マクロ名:",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=10, anchor="e").pack(side="left", padx=(0, 6))
        self._macro_name_var = tk.StringVar()
        name_entry = tk.Entry(name_row, textvariable=self._macro_name_var,
                              font=("Yu Gothic UI", 10),
                              bg=INPUT_BG, fg=TEXT_MAIN,
                              insertbackground=ACCENT, relief="solid", bd=1,
                              highlightcolor=ACCENT, highlightbackground=BORDER,
                              highlightthickness=1)
        name_entry.pack(side="left", fill="x", expand=True)
        self._macro_name_var.trace_add("write", lambda *a: self._macro_mark_dirty())

        # 説明（UI非表示。データは保持してJSON互換性のため保存はする）
        self._macro_desc_var = tk.StringVar()

        # 区切り
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=(8, 6))

        # ステップリスト ヘッダ
        step_header = tk.Frame(parent, bg=BG_PANEL)
        step_header.pack(fill="x", pady=(0, 4))
        tk.Label(step_header, text="ステップ一覧",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10, "bold")).pack(side="left")
        tk.Label(step_header, text="（上から順番に実行）",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(side="left", padx=(6, 0))

        # ステップリスト本体（スクロール対応）
        list_wrap = tk.Frame(parent, bg=BG_PANEL,
                             highlightbackground=BORDER, highlightthickness=1)
        list_wrap.pack(fill="both", expand=True, pady=(0, 6))

        self._macro_step_canvas = tk.Canvas(list_wrap, bg=BG_PANEL, bd=0,
                                             highlightthickness=0, takefocus=False,
                                             height=180)
        step_scroll = tk.Scrollbar(list_wrap, orient="vertical",
                                    command=self._macro_step_canvas.yview,
                                    bg=BG_PANEL, troughcolor=BG_HEADER,
                                    activebackground=ACCENT_DARK,
                                    bd=0, highlightthickness=0, width=10)
        self._macro_step_canvas.configure(yscrollcommand=step_scroll.set)
        step_scroll.pack(side="right", fill="y")
        self._macro_step_canvas.pack(side="left", fill="both", expand=True)

        self._macro_step_inner = tk.Frame(self._macro_step_canvas, bg=BG_PANEL)
        self._macro_step_inner_id = self._macro_step_canvas.create_window(
            (0, 0), window=self._macro_step_inner, anchor="nw"
        )

        def _on_step_inner_configure(event):
            self._macro_step_canvas.configure(
                scrollregion=self._macro_step_canvas.bbox("all")
            )
        self._macro_step_inner.bind("<Configure>", _on_step_inner_configure)

        def _on_step_canvas_configure(event):
            self._macro_step_canvas.itemconfigure(
                self._macro_step_inner_id, width=event.width
            )
        self._macro_step_canvas.bind("<Configure>", _on_step_canvas_configure)

        # ＋ステップ追加 ボタン
        add_btn = tk.Button(parent, text="＋ ステップ追加",
                            command=self._macro_open_step_add_dialog,
                            font=("Yu Gothic UI", 10, "bold"),
                            bg="#141B24", fg=ACCENT,
                            activebackground="#1F2B3A", activeforeground=ACCENT,
                            relief="solid", bd=1, padx=14, pady=4, cursor="hand2")
        add_btn.pack(fill="x", pady=(0, 8))

        # 区切り
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=(0, 6))

        # 実行ログ
        tk.Label(parent, text="実行ログ",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        log_wrap = tk.Frame(parent, bg=BG_PANEL,
                            highlightbackground=BORDER, highlightthickness=1)
        log_wrap.pack(fill="both", expand=False, pady=(0, 6))

        self._macro_log_text = tk.Text(log_wrap, height=6, wrap="word",
                                        font=("Consolas", 9),
                                        bg=INPUT_BG, fg=TEXT_MAIN,
                                        insertbackground=ACCENT,
                                        relief="flat", bd=0, padx=6, pady=4)
        self._macro_log_text.pack(side="left", fill="both", expand=True)
        self._macro_log_text.configure(state="disabled")

        log_scroll = tk.Scrollbar(log_wrap, orient="vertical",
                                   command=self._macro_log_text.yview,
                                   bg=BG_PANEL, troughcolor=BG_HEADER,
                                   activebackground=ACCENT_DARK,
                                   bd=0, highlightthickness=0, width=10)
        self._macro_log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")

        # ログ用タグ設定
        self._macro_log_text.tag_configure("ok", foreground=ACCENT)
        self._macro_log_text.tag_configure("warn", foreground="#F0D060")
        self._macro_log_text.tag_configure("err", foreground="#FB5E7E")
        self._macro_log_text.tag_configure("info", foreground=TEXT_MUTED)
        self._macro_log_text.tag_configure("step", foreground="#3DDDFF")

        # 実行ボタン群
        btn_row = tk.Frame(parent, bg=BG_PANEL)
        btn_row.pack(fill="x", pady=(4, 0))

        btn_kw = {"font": ("Yu Gothic UI", 10, "bold"),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "solid", "bd": 1, "padx": 12, "pady": 4,
                  "cursor": "hand2"}

        tk.Button(btn_row, text="プレビュー",
                  command=self._macro_run_preview,
                  **btn_kw).pack(side="left", padx=(0, 4))

        tk.Button(btn_row, text="▶ 実行",
                  command=self._macro_run_execute,
                  font=("Yu Gothic UI", 10, "bold"),
                  bg="#123A38", fg=ACCENT,
                  activebackground="#175E56", activeforeground=ACCENT,
                  relief="solid", bd=1, padx=12, pady=4,
                  cursor="hand2").pack(side="left", padx=(0, 4))

        tk.Label(btn_row,
                 text="※ 実行後はメイン画面の [↶ 元に戻す] か Ctrl+Z で取り消せます",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).pack(side="left", padx=(8, 0))

        # 保存系
        save_row = tk.Frame(parent, bg=BG_PANEL)
        save_row.pack(fill="x", pady=(6, 0))

        tk.Button(save_row, text="💾 保存",
                  command=self._macro_save,
                  **btn_kw).pack(side="left", padx=(0, 4))
        tk.Button(save_row, text="🗑 削除",
                  command=self._macro_delete,
                  **btn_kw).pack(side="left", padx=(0, 4))

    # ---------- マクロ一覧の管理 ----------

    def _macro_refresh_list(self) -> None:
        """macros/ ディレクトリをスキャンしてマクロ一覧を再構築"""
        # 既存のボタンクリア
        for w in self._macro_list_inner.winfo_children():
            w.destroy()

        macros_dir = get_macros_dir()
        try:
            files = sorted([f for f in os.listdir(macros_dir) if f.lower().endswith(".json")])
        except OSError:
            files = []

        if not files:
            tk.Label(self._macro_list_inner,
                     text="（保存済みマクロなし）",
                     bg=BG_PANEL, fg=TEXT_MUTED,
                     font=("Yu Gothic UI", 9),
                     anchor="w").pack(fill="x", padx=4, pady=4)
            return

        self._macro_list_buttons = {}
        for fname in files:
            fpath = os.path.join(macros_dir, fname)
            display_name = os.path.splitext(fname)[0]
            btn = tk.Button(
                self._macro_list_inner, text=f"  {display_name}",
                command=lambda p=fpath: self._macro_load_from_file(p),
                font=("Yu Gothic UI", 9), bg=BG_PANEL, fg=TEXT_MAIN,
                activebackground="#123A38", activeforeground=ACCENT,
                relief="flat", bd=0, anchor="w", padx=8, pady=2, cursor="hand2",
            )
            btn.pack(fill="x", pady=1)
            self._macro_list_buttons[fpath] = btn

    def _macro_open_folder(self) -> None:
        """OS のファイラでマクロフォルダを開く"""
        macros_dir = get_macros_dir()
        try:
            if sys.platform == "win32":
                os.startfile(macros_dir)
            elif sys.platform == "darwin":
                os.system(f'open "{macros_dir}"')
            else:
                os.system(f'xdg-open "{macros_dir}"')
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダを開けませんでした:\n{e}", parent=self)

    def _macro_load_from_file(self, path: str) -> None:
        """JSONファイルからマクロを読み込み、編集パネルにセット"""
        if self._macro_dirty:
            ans = messagebox.askyesno(
                "未保存の変更",
                "現在のマクロに未保存の変更があります。\n破棄して読み込みますか？",
                parent=self,
            )
            if not ans:
                return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            macro = Macro.from_dict(data)
            self._macro_load_into_panel(macro)
            self._macro_log("info", f"📂 読込: {os.path.basename(path)}\n", clear=True)
        except (OSError, json.JSONDecodeError, ValueError) as e:
            messagebox.showerror("読込エラー",
                                 f"マクロを読み込めませんでした:\n{e}", parent=self)

    def _macro_load_into_panel(self, macro: Macro) -> None:
        """マクロを編集パネルに反映"""
        self._current_macro = macro
        self._macro_name_var.set(macro.name)
        self._macro_desc_var.set(macro.description)
        self._macro_render_steps()
        self._macro_dirty = False

    def _macro_new(self) -> None:
        """新規マクロ作成"""
        if self._macro_dirty:
            ans = messagebox.askyesno(
                "未保存の変更",
                "現在のマクロに未保存の変更があります。\n破棄して新規作成しますか？",
                parent=self,
            )
            if not ans:
                return
        self._macro_load_into_panel(Macro())
        self._macro_log("info", "📝 新規マクロ\n", clear=True)

    def _macro_mark_dirty(self) -> None:
        """変更フラグを立てる（StringVar.trace 用）"""
        # 名前/説明を Macro に反映
        if hasattr(self, "_current_macro") and self._current_macro is not None:
            self._current_macro.name = self._macro_name_var.get()
            self._current_macro.description = self._macro_desc_var.get()
        self._macro_dirty = True

    # ---------- ステップリストの描画 ----------

    def _macro_render_steps(self) -> None:
        """ステップリストを再描画"""
        for w in self._macro_step_inner.winfo_children():
            w.destroy()
        self._macro_step_widgets = []

        if not self._current_macro.steps:
            tk.Label(self._macro_step_inner,
                     text="（ステップなし。「＋ ステップ追加」から追加してください）",
                     bg=BG_PANEL, fg=TEXT_MUTED,
                     font=("Yu Gothic UI", 9),
                     anchor="w").pack(fill="x", padx=8, pady=8)
            return

        for idx, step in enumerate(self._current_macro.steps):
            row = self._make_step_row(self._macro_step_inner, idx, step)
            self._macro_step_widgets.append(row)

    def _make_step_row(self, parent: tk.Frame, idx: int, step: "MacroStep") -> tk.Frame:
        """ステップ1行のウィジェットを作成"""
        # 全体の枠
        row = tk.Frame(parent, bg=BG_PANEL,
                       highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", pady=2, padx=2)

        inner = tk.Frame(row, bg=BG_PANEL)
        inner.pack(fill="x", padx=4, pady=3)

        # 有効化チェックボックス
        enabled_var = tk.StringVar(value="1" if step.enabled else "0")
        def _on_toggle():
            step.enabled = (enabled_var.get() == "1")
            self._macro_dirty = True
            # 有効/無効でラベル色を変える
            label_widget._target_step = step
            self._update_step_row_appearance(label_widget, step)

        cb = tk.Checkbutton(
            inner, variable=enabled_var, onvalue="1", offvalue="0",
            command=_on_toggle,
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            bd=0, highlightthickness=0,
        )
        cb.pack(side="left", padx=(0, 4))

        # 番号
        num_label = tk.Label(inner, text=f"{idx+1}.",
                             bg=BG_PANEL, fg=ACCENT,
                             font=("Consolas", 10, "bold"),
                             width=3, anchor="w")
        num_label.pack(side="left", padx=(0, 4))

        # ラベル（テンプレ名 + パラメータ要約）
        label_text = self._format_step_summary(step)
        label_widget = tk.Label(inner, text=label_text,
                                bg=BG_PANEL,
                                fg=TEXT_MAIN if step.enabled else TEXT_MUTED,
                                font=("Yu Gothic UI", 10),
                                anchor="w", justify="left")
        label_widget.pack(side="left", fill="x", expand=True)

        # 操作ボタン群（右寄せ）
        btn_kw = {"font": ("Yu Gothic UI", 9),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "flat", "bd": 0, "padx": 6, "pady": 0,
                  "cursor": "hand2"}

        tk.Button(inner, text="↓", command=lambda i=idx: self._macro_step_move(i, 1),
                  **btn_kw).pack(side="right", padx=1)
        tk.Button(inner, text="↑", command=lambda i=idx: self._macro_step_move(i, -1),
                  **btn_kw).pack(side="right", padx=1)
        tk.Button(inner, text="✕", command=lambda i=idx: self._macro_step_delete(i),
                  font=("Yu Gothic UI", 9),
                  bg="#141B24", fg="#FB5E7E",
                  activebackground="#3A1620", activeforeground="#FF93A8",
                  relief="flat", bd=0, padx=6, pady=0,
                  cursor="hand2").pack(side="right", padx=1)
        tk.Button(inner, text="✎", command=lambda i=idx: self._macro_step_edit(i),
                  **btn_kw).pack(side="right", padx=1)

        return row

    def _update_step_row_appearance(self, label_widget: tk.Label, step: "MacroStep") -> None:
        """ステップ行の見た目を有効/無効に応じて更新"""
        label_widget.configure(fg=TEXT_MAIN if step.enabled else TEXT_MUTED)

    def _format_step_summary(self, step: "MacroStep") -> str:
        """ステップの内容を1行サマリ文字列にする"""
        label = step.template_label or step.template_id
        params = step.params or {}
        tid = step.template_id

        if tid == "simple_replace":
            old = params.get("old", "?")
            new = params.get("new", "")
            # 改行を ↵ で可視化
            old_disp = old.replace("\n", "↵").replace("\r", "")
            new_disp = new.replace("\n", "↵").replace("\r", "")
            # 表示が長すぎる場合は省略
            if len(old_disp) > 24:
                old_disp = old_disp[:22] + "…"
            if len(new_disp) > 24:
                new_disp = new_disp[:22] + "…"
            flags = []
            if params.get("use_regex", False):
                flags.append("正規表現")
            if not params.get("case_sensitive", True):
                flags.append("大小無視")
            flag_str = f" [{', '.join(flags)}]" if flags else ""
            return f"{label}  ('{old_disp}' → '{new_disp}'){flag_str}"
        if tid in ("f_scale", "s_scale"):
            return f"{label}  (係数: {params.get('factor', '?')})"
        if tid in ("f_uniform", "s_uniform"):
            return f"{label}  (値: {params.get('value', '?')})"
        if tid == "axis_offset":
            return f"{label}  ({params.get('axis', '?')}軸 + {params.get('offset', '?')})"
        if tid == "axis_invert":
            return f"{label}  ({params.get('axis', '?')}軸)"
        if tid == "decimal_add":
            return f"{label}  (対象: {params.get('axis', '?')} / スケール: {params.get('scale', '?')})"
        if tid in ("upper_case", "trim_spaces", "comment_add", "comment_del", "comment_remove"):
            return f"{label}"
        if tid == "char_swap":
            a = params.get("a", "?")
            b = params.get("b", "?")
            outer_s = params.get("outer_start", "")
            outer_e = params.get("outer_end", "")
            scope_str = ""
            if outer_s and outer_e:
                inner_s = params.get("inner_start", "")
                inner_e = params.get("inner_end", "")
                if inner_s and inner_e:
                    scope_str = f" / 範囲: {outer_s}〜{outer_e} ＞ {inner_s}〜{inner_e}"
                else:
                    scope_str = f" / 範囲: {outer_s}〜{outer_e}"
            flex_str = " [柔軟]" if params.get("flex", False) else ""
            cmt_str = " [コメント保護]" if params.get("skip_comment", True) else ""
            return f"{label}  ({a} ↔ {b}){scope_str}{flex_str}{cmt_str}"
        if tid == "line_delete":
            cond_parts = []
            if params.get("delete_empty", False):
                cond_parts.append("空行")
            if params.get("delete_pattern", False):
                pat = params.get("pattern", "")
                if len(pat) > 16:
                    pat = pat[:14] + "…"
                regex_mark = "regex" if params.get("use_regex", False) else ""
                case_mark = "" if params.get("case_sensitive", True) else "ci"
                marks = "/".join(m for m in [regex_mark, case_mark] if m)
                if marks:
                    cond_parts.append(f"'{pat}' [{marks}]")
                else:
                    cond_parts.append(f"'{pat}'")
            cond_str = " + ".join(cond_parts) if cond_parts else "?"
            return f"{label}  ({cond_str})"
        return label

    # ---------- ステップ操作 ----------

    def _macro_step_move(self, idx: int, direction: int) -> None:
        """ステップを上下に移動（direction: -1=上, +1=下）"""
        steps = self._current_macro.steps
        new_idx = idx + direction
        if 0 <= new_idx < len(steps):
            steps[idx], steps[new_idx] = steps[new_idx], steps[idx]
            self._macro_render_steps()
            self._macro_dirty = True

    def _macro_step_delete(self, idx: int) -> None:
        """ステップを削除"""
        steps = self._current_macro.steps
        if 0 <= idx < len(steps):
            label = steps[idx].template_label
            if not messagebox.askyesno("確認",
                                       f"ステップ {idx+1}「{label}」を削除しますか？",
                                       parent=self):
                return
            del steps[idx]
            self._macro_render_steps()
            self._macro_dirty = True

    def _macro_step_edit(self, idx: int) -> None:
        """ステップを編集（追加ダイアログを編集モードで開く）"""
        steps = self._current_macro.steps
        if 0 <= idx < len(steps):
            self._macro_open_step_add_dialog(edit_index=idx)

    # ---------- ステップ追加ダイアログ ----------

    def _macro_open_step_add_dialog(self, edit_index: int | None = None) -> None:
        """ステップ追加（または編集）ダイアログを開く"""
        dialog = MacroStepDialog(self, edit_index=edit_index)
        # モーダル風（grab）で待機
        dialog.transient(self)
        dialog.grab_set()
        self.wait_window(dialog)

        if dialog.result_step is not None:
            if edit_index is not None:
                # 編集モード：既存ステップを置き換え
                self._current_macro.steps[edit_index] = dialog.result_step
            else:
                # 追加モード
                self._current_macro.steps.append(dialog.result_step)
            self._macro_render_steps()
            self._macro_dirty = True

    # ---------- 実行・プレビュー ----------

    def _macro_log(self, tag: str, msg: str, clear: bool = False) -> None:
        """実行ログに書き込み"""
        self._macro_log_text.configure(state="normal")
        if clear:
            self._macro_log_text.delete("1.0", "end")
        self._macro_log_text.insert("end", msg, tag)
        self._macro_log_text.see("end")
        self._macro_log_text.configure(state="disabled")

    def _macro_collect_text(self) -> str:
        """現在のエディタ全文を取得"""
        return self.target.get("1.0", "end-1c")

    def _macro_run_preview(self) -> None:
        """プレビュー：差分一覧ダイアログを表示し、OKされたら実行へ進む"""
        if not self._current_macro.steps:
            messagebox.showinfo("プレビュー", "ステップが空です。", parent=self)
            return

        original = self._macro_collect_text()

        # 全ステップを順次適用してログ収集
        current = original
        total_changes = 0
        executed = 0
        skipped = 0
        step_logs = []  # [(tag, msg), ...]
        error_occurred = False

        for i, step in enumerate(self._current_macro.steps, 1):
            label = step.template_label or step.template_id
            if not step.enabled:
                step_logs.append(("info", f"Step{i}: {label} → スキップ(無効)\n"))
                skipped += 1
                continue
            try:
                new_text, count = self.macro_apply_step(current, step)
                if count > 0:
                    step_logs.append(("step", f"Step{i}: {label} → {count}行変更\n"))
                else:
                    step_logs.append(("warn", f"Step{i}: {label} → 変更箇所なし\n"))
                current = new_text
                total_changes += count
                executed += 1
            except Exception as e:
                step_logs.append(("err", f"Step{i}: {label} → エラー: {e}\n"))
                step_logs.append(("err", "→ ここで中断\n"))
                error_occurred = True
                break

        # エラーがあった場合はダイアログ出さずログだけ表示
        if error_occurred:
            self._macro_log("info", "─── プレビュー（エラーで中断）───\n", clear=True)
            for tag, msg in step_logs:
                self._macro_log(tag, msg)
            return

        # サマリログ
        step_logs.append(("ok",
                          f"\n✅ プレビュー完了 (実行: {executed}, スキップ: {skipped}, 計 {total_changes}行変更)\n"))

        if current == original:
            step_logs.append(("warn", "※ 変更なし。実行しても元と同じです。\n"))

        # 親パネルのログにも反映
        self._macro_log("info", "─── プレビュー実行 ───\n", clear=True)
        for tag, msg in step_logs:
            self._macro_log(tag, msg)

        # プレビューダイアログを開く
        preview = MacroPreviewDialog(
            self,
            original_text=original,
            final_text=current,
            step_logs=step_logs,
            total_changes=total_changes,
            executed=executed,
            skipped=skipped,
        )
        preview.transient(self)
        preview.grab_set()
        self.wait_window(preview)

        # OK 押されたら実行へ
        if preview.confirmed:
            self._macro_apply_text(original, current, executed, skipped, total_changes)

    def _macro_apply_text(self, original: str, final: str,
                          executed: int, skipped: int, total_changes: int) -> None:
        """事前計算済みの最終テキストをエディタへ反映する。
        プレビューダイアログから直接実行される場合に使用。
        """
        if final == original:
            self._macro_log("warn", "※ 変更なし。テキストは書き換えませんでした。\n")
            return

        self._macro_undo_snapshot = original
        self._atomic_replace(self.target, "1.0", "end", final)

        self._macro_log("ok",
                        f"\n✅ 実行完了 (実行: {executed}, スキップ: {skipped}, 計 {total_changes}行変更)\n")
        self._macro_log("info", "メイン画面の [↶ 元に戻す] か Ctrl+Z で戻せます。\n")
        self._status_var.set(f"マクロ実行: {total_changes}行変更")

        try:
            self._trigger_post_replace_check_debounced()
        except Exception:
            pass

    def _macro_run_execute(self) -> None:
        """本番実行：エディタ全文を書き換える"""
        if not self._current_macro.steps:
            messagebox.showinfo("実行", "ステップが空です。", parent=self)
            return

        if not messagebox.askyesno(
            "マクロ実行",
            f"マクロ「{self._current_macro.name or '(無名)'}」を実行します。\n"
            f"ステップ数: {len(self._current_macro.steps)}\n\n"
            f"※ 実行前のテキストはバックアップされ、[元に戻す] で復元できます。\n"
            "実行してよろしいですか？",
            parent=self,
        ):
            return

        original = self._macro_collect_text()
        self._macro_log("info", "─── 本番実行 ───\n", clear=True)

        current = original
        total_changes = 0
        executed = 0
        skipped = 0

        for i, step in enumerate(self._current_macro.steps, 1):
            label = step.template_label or step.template_id
            if not step.enabled:
                self._macro_log("info", f"Step{i}: {label} → スキップ(無効)\n")
                skipped += 1
                continue
            try:
                new_text, count = self.macro_apply_step(current, step)
                if count > 0:
                    self._macro_log("step",
                                    f"Step{i}: {label} → {count}行変更\n")
                else:
                    self._macro_log("warn",
                                    f"Step{i}: {label} → 変更箇所なし\n")
                current = new_text
                total_changes += count
                executed += 1
            except Exception as e:
                self._macro_log("err",
                                f"Step{i}: {label} → エラー: {e}\n")
                self._macro_log("err", "→ ここで中断（テキスト未変更）\n")
                return

        if current == original:
            self._macro_log("warn", "※ 変更なし。テキストは書き換えませんでした。\n")
            return

        # スナップショット保存 → エディタに反映
        self._macro_undo_snapshot = original
        self._atomic_replace(self.target, "1.0", "end", current)

        # ステータス・後処理
        self._macro_log("ok",
                        f"\n✅ 実行完了 (実行: {executed}, スキップ: {skipped}, 計 {total_changes}行変更)\n")
        self._macro_log("info", "メイン画面の [↶ 元に戻す] か Ctrl+Z で戻せます。\n")
        self._status_var.set(f"マクロ実行: {total_changes}行変更")

        # 既存の置換後フック呼ぶ（行番号更新・自動再チェック）
        try:
            self._trigger_post_replace_check_debounced()
        except Exception:
            pass

    def _macro_undo(self) -> None:
        """マクロ実行を取り消し（実行前のテキストに戻す）"""
        if self._macro_undo_snapshot is None:
            messagebox.showinfo("Undo", "戻せる実行履歴がありません。", parent=self)
            return
        if not messagebox.askyesno(
            "Undo",
            "実行前の状態に戻しますか？\n（マクロ実行で行った変更がすべて取り消されます）",
            parent=self,
        ):
            return
        self.target.delete("1.0", "end")
        self.target.insert("1.0", self._macro_undo_snapshot)
        self._macro_undo_snapshot = None
        self._macro_log("info", "↶ 元に戻しました\n")
        self._status_var.set("マクロをUndoしました")
        try:
            self._trigger_post_replace_check_debounced()
        except Exception:
            pass

    # ---------- 保存・削除 ----------

    def _macro_save(self) -> None:
        """現在のマクロをJSONとして保存"""
        # 名前を Macro に同期
        self._current_macro.name = self._macro_name_var.get().strip()
        self._current_macro.description = self._macro_desc_var.get().strip()

        if not self._current_macro.name:
            messagebox.showwarning("保存", "マクロ名を入力してください。", parent=self)
            return
        if not self._current_macro.steps:
            messagebox.showwarning("保存", "ステップが空のマクロは保存できません。", parent=self)
            return

        macros_dir = get_macros_dir()
        safe_name = sanitize_macro_filename(self._current_macro.name)
        path = os.path.join(macros_dir, f"{safe_name}.json")

        # 既存上書き確認
        if os.path.exists(path) and not self._current_macro.created_at:
            # 新規作成時のみ上書き確認（既存読込→保存は黙って上書き）
            if not messagebox.askyesno(
                "上書き確認",
                f"同名のマクロが既に存在します:\n{safe_name}.json\n\n上書きしますか？",
                parent=self,
            ):
                return

        # タイムスタンプ
        now = datetime.now().isoformat(timespec="seconds")
        if not self._current_macro.created_at:
            self._current_macro.created_at = now
        self._current_macro.updated_at = now

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self._current_macro.to_dict(), f,
                          ensure_ascii=False, indent=2)
            self._macro_log("ok", f"💾 保存: {os.path.basename(path)}\n")
            self._macro_dirty = False
            self._macro_refresh_list()
        except OSError as e:
            messagebox.showerror("保存エラー", f"保存できませんでした:\n{e}", parent=self)

    def _macro_delete(self) -> None:
        """現在開いているマクロのファイルを削除"""
        name = self._macro_name_var.get().strip()
        if not name:
            messagebox.showinfo("削除", "削除対象のマクロが特定できません。", parent=self)
            return

        macros_dir = get_macros_dir()
        safe_name = sanitize_macro_filename(name)
        path = os.path.join(macros_dir, f"{safe_name}.json")

        if not os.path.exists(path):
            messagebox.showinfo("削除",
                                f"ファイルが存在しません:\n{safe_name}.json\n"
                                "（まだ保存されていないマクロです）", parent=self)
            return

        if not messagebox.askyesno(
            "削除確認",
            f"マクロ「{name}」を削除します。\nこの操作は取り消せません。\n\n本当に削除しますか？",
            parent=self,
        ):
            return

        try:
            os.remove(path)
            self._macro_log("ok", f"🗑 削除: {os.path.basename(path)}\n")
            self._macro_load_into_panel(Macro())
            self._macro_refresh_list()
        except OSError as e:
            messagebox.showerror("削除エラー", f"削除できませんでした:\n{e}", parent=self)

    def _show_template_placeholder(self) -> None:

        """テンプレ未選択時のプレースホルダー"""
        for w in self._template_detail_frame.winfo_children():
            w.destroy()
        tk.Label(self._template_detail_frame,
                 text="◀ 左から使いたいテンプレートを選択してください",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 10),
                 wraplength=300, justify="left").pack(anchor="w", pady=20, padx=4)

    def _select_template(self, tpl_id: str) -> None:
        """テンプレ選択：右ペインに該当テンプレの設定UIを表示"""
        # 全ボタンの色を戻す
        for tid, btn in self._template_buttons.items():
            if tid == tpl_id:
                btn.configure(bg="#123A38", fg=ACCENT)
            else:
                btn.configure(bg=BG_PANEL, fg=TEXT_MAIN)

        # 詳細ペインクリア
        for w in self._template_detail_frame.winfo_children():
            w.destroy()

        # 該当テンプレの情報取得
        tpl_info = next((t for t in self._templates if t[0] == tpl_id), None)
        if tpl_info is None:
            return
        _, label, category, desc = tpl_info

        # ヘッダー
        tk.Label(self._template_detail_frame, text=f"▶ {label}",
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(self._template_detail_frame, text=f"[{category}]　{desc}",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9), wraplength=380, justify="left").pack(anchor="w", pady=(0, 12))

        # 各テンプレに対応するUIビルダーを呼び出し
        builder = getattr(self, f"_tpl_ui_{tpl_id}", None)
        if builder:
            builder(self._template_detail_frame)
        else:
            tk.Label(self._template_detail_frame,
                     text="（このテンプレートは未実装です）",
                     bg=BG_PANEL, fg=TEXT_MUTED,
                     font=("Yu Gothic UI", 10)).pack(anchor="w")

    # =========================================================================
    # 共通UIヘルパー
    # =========================================================================

    def _make_tpl_entry(self, parent, label_text: str, default: str = "", width: int = 14) -> tk.Entry:
        """テンプレ用の入力欄1行（ラベル＋エントリ）"""
        row = tk.Frame(parent, bg=BG_PANEL)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label_text, bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="e").pack(side="left", padx=(0, 8))
        var = tk.StringVar(value=default)
        entry = tk.Entry(
            row, textvariable=var, font=("Consolas", 11),
            bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT,
            relief="solid", bd=1, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT, width=width,
        )
        entry.pack(side="left")
        entry._var = var  # type: ignore
        return entry

    def _make_tpl_combobox(self, parent, label_text: str, options: list, default: str = "") -> tk.StringVar:
        """テンプレ用のコンボボックス代替（ラジオボタン横並び）"""
        row = tk.Frame(parent, bg=BG_PANEL)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label_text, bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="e").pack(side="left", padx=(0, 8))
        var = tk.StringVar(value=default or options[0])
        rb_kw = {"bg": BG_PANEL, "fg": TEXT_MAIN, "selectcolor": INPUT_BG,
                 "activebackground": BG_PANEL, "activeforeground": ACCENT,
                 "font": ("Yu Gothic UI", 10), "bd": 0, "highlightthickness": 0}
        for opt in options:
            tk.Radiobutton(row, text=opt, variable=var, value=opt, **rb_kw).pack(side="left", padx=(0, 8))
        return var

    def _make_tpl_target_radio(self, parent):
        """対象範囲選択（全体/選択範囲/指定したNブロック）。
        「指定したNブロック」選択時はNブロックドロップダウンを動的表示。
        Returns: TplTargetSelector（.get() で範囲名、.get_n_block() でNラベルを返す）
        """
        return TplTargetSelector(self, parent)

    def _make_tpl_buttons(self, parent, on_run) -> None:
        """実行ボタン（押すとプレビュー→確認→実行）"""
        row = tk.Frame(parent, bg=BG_PANEL)
        row.pack(fill="x", pady=(16, 0))
        btn_kw_primary = {
            "font": ("Yu Gothic UI", 10, "bold"),
            "bg": "#123A38", "fg": ACCENT,
            "activebackground": "#1F2B3A", "activeforeground": ACCENT,
            "relief": "solid", "bd": 1, "padx": 14, "pady": 4,
            "cursor": "hand2",
        }
        tk.Button(row, text="実行", command=on_run, **btn_kw_primary).pack(side="left", padx=(0, 6))
        tk.Label(row, text="※ 押すとプレビューが表示されます",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(side="left", padx=(8, 0))

    def _get_target_text(self, scope):
        """対象範囲のテキストと範囲インデックスを返す。
        scope: '全体' | '選択範囲' | '指定したNブロック' | TplTargetSelector インスタンス
        Returns: (text, start_idx, end_idx)
        """
        # TplTargetSelectorが渡された場合の処理
        selector = None
        if isinstance(scope, TplTargetSelector):
            selector = scope
            scope_name = scope.get()
        else:
            scope_name = scope

        if scope_name == "選択範囲":
            try:
                start = self.target.index("sel.first")
                end = self.target.index("sel.last")
                text = self.target.get(start, end)
                if not text:
                    return None, None, None
                return text, start, end
            except tk.TclError:
                self._status_var.set("⚠ テキストが選択されていません")
                return None, None, None
        elif scope_name == "指定したNブロック":
            if selector is None:
                self._status_var.set("⚠ Nブロック指定が取得できません")
                return None, None, None
            n_label = selector.get_n_block_label()
            if not n_label:
                self._status_var.set("⚠ Nブロックを選択してください")
                return None, None, None
            return self._get_n_block_text_range(n_label)
        else:
            text = self.target.get("1.0", "end-1c")
            return text, "1.0", "end-1c"

    def _get_n_block_text_range(self, n_label: str):
        """指定Nラベルから次のNブロック直前までのテキストと範囲を返す。
        Returns: (text, start_idx, end_idx)
        """
        full_text = self.target.get("1.0", "end-1c")
        lines = full_text.splitlines(keepends=True)

        # 指定Nブロックの開始行を探す
        start_line_no = None
        for i, line in enumerate(lines, start=1):
            m = re.match(r"^\s*(N\d+)\b", line)
            if m and m.group(1) == n_label:
                start_line_no = i
                break

        if start_line_no is None:
            self._status_var.set(f"⚠ {n_label} が見つかりません")
            return None, None, None

        # 次のNブロック開始行（または末尾）を探す
        end_line_no = len(lines) + 1  # デフォルトは末尾の次
        for i in range(start_line_no, len(lines)):
            line = lines[i]  # 0-indexed
            m = re.match(r"^\s*(N\d+)\b", line)
            if m:
                end_line_no = i + 1  # 1-indexedの行番号
                break

        # tkインデックス
        start_idx = f"{start_line_no}.0"
        end_idx = f"{end_line_no}.0"
        text = self.target.get(start_idx, end_idx)
        return text, start_idx, end_idx

    def _strip_existing_nc_comments(self, s: str) -> str:
        """NC行から (...) コメント部分を除いた素のNCコードを返す（連続空白も圧縮）"""
        no_comment = re.sub(r"\([^)]*\)", "", s)
        # 連続空白を1個にまとめて両端トリム
        return re.sub(r"\s+", " ", no_comment).strip()

    def _build_commented_line(self, old_content: str, new_content: str, ending: str = "") -> str:
        """変更後の行に (元: <旧素コード>) を追記。ending は改行文字（"\\n" 等）。
        旧行が空 or コメントのみの場合はそのまま返す。
        """
        old_clean = self._strip_existing_nc_comments(old_content)
        if not old_clean:
            return new_content + ending
        return f"{new_content} (元: {old_clean}){ending}"

    def _show_preview_dialog(self, title: str, line_changes: list, change_indices: list, on_confirm) -> None:
        """プレビューダイアログ表示。各変更行にチェックボックス [✓]/[ ] を表示、
        クリックでトグル。「実行する」押下時、ON行のインデックス集合と
        コメント追記フラグを on_confirm に渡す。

        line_changes:   [(line_no, old_content, new_content), ...]
        change_indices: [original_lines における変更行のインデックス, ...]
        on_confirm:     callable(selected_set: set[int], with_comment: bool) → 採用処理
        """
        if not line_changes:
            messagebox.showinfo("プレビュー", "対象が見つかりませんでした。", parent=self)
            self._status_var.set("一致なし")
            return

        # コメント追記用：旧行が空 or コメントのみなら追記しない
        def _make_commented_display(old_line: str, new_line: str) -> str:
            old_clean = self._strip_existing_nc_comments(old_line)
            if not old_clean:
                return new_line
            return f"{new_line} (元: {old_clean})"

        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.configure(bg=BG_PANEL)
        dlg.transient(self)
        dlg.geometry("700x680")
        dlg.minsize(580, 460)

        outer = tk.Frame(dlg, bg=BG_PANEL, padx=14, pady=12)
        outer.pack(fill="both", expand=True)

        # ヘッダ
        tk.Label(outer, text=f"▶ {title}", bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w")

        # 件数表示（ON件数を動的更新）
        info_var = tk.StringVar(value=f"変更対象: {len(line_changes)} 行")
        tk.Label(outer, textvariable=info_var,
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10)).pack(anchor="w", pady=(2, 4))

        tk.Label(outer, text="※ 行頭の [✓] をクリックでON/OFF切替",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(anchor="w", pady=(0, 6))

        # コメント追記オプション
        with_comment_var = tk.StringVar(value="0")
        cb_row = tk.Frame(outer, bg=BG_PANEL)
        cb_row.pack(fill="x", pady=(0, 6))
        tk.Checkbutton(
            cb_row, text="元の値を行末に (元: ...) で追記する",
            variable=with_comment_var, onvalue="1", offvalue="0",
            command=lambda: render_all(),
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            font=("Yu Gothic UI", 9), bd=0, highlightthickness=0,
        ).pack(side="left")
        tk.Label(cb_row, text="※ 既存の ( ) コメントは除外して追記",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 8)).pack(side="left", padx=(8, 0))

        # 全選択/全解除ボタン群
        sel_btn_row = tk.Frame(outer, bg=BG_PANEL)
        sel_btn_row.pack(fill="x", pady=(0, 6))
        sel_btn_kw = {"font": ("Yu Gothic UI", 9, "bold"),
                      "bg": "#141B24", "fg": TEXT_MAIN,
                      "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                      "relief": "solid", "bd": 1, "padx": 10, "pady": 2,
                      "cursor": "hand2"}

        # 状態：採用する change_indices のセット（初期は全選択）
        selected: set[int] = set(change_indices)
        # change_indices と line_changes の対応付け（同順序前提）
        # i 番目のプレビュー行 ↔ change_indices[i] が original_lines のインデックス

        # スクロール可能なテキスト表示
        text_frame = tk.Frame(outer, bg=BG_PANEL)
        text_frame.pack(fill="both", expand=True)
        preview_text = tk.Text(
            text_frame, wrap=tk.NONE,
            font=("Consolas", 10), bg=INPUT_BG, fg=TEXT_MAIN,
            relief="solid", bd=1, highlightthickness=1,
            highlightbackground=BORDER, padx=8, pady=8,
            cursor="arrow",
        )
        preview_text.pack(side="left", fill="both", expand=True)
        vbar = tk.Scrollbar(text_frame, orient="vertical",
                            bg=BG_PANEL, troughcolor="#101720",
                            activebackground=ACCENT, command=preview_text.yview)
        vbar.pack(side="right", fill="y")
        preview_text.configure(yscrollcommand=vbar.set)

        # タグ設定
        preview_text.tag_configure("preview_idx",     foreground="#6CB6FF", font=("Consolas", 10, "bold"))
        preview_text.tag_configure("preview_label",   foreground=TEXT_MUTED, font=("Consolas", 9))
        preview_text.tag_configure("preview_before",  foreground="#FF9557", background="#301019")
        preview_text.tag_configure("preview_after",   foreground="#7FF0CB", background="#0E241C")
        preview_text.tag_configure("preview_off_before", foreground="#5C6B78")
        preview_text.tag_configure("preview_off_after",  foreground="#5C6B78")
        preview_text.tag_configure("checkbox_on",  foreground=ACCENT, font=("Consolas", 11, "bold"))
        preview_text.tag_configure("checkbox_off", foreground=TEXT_MUTED, font=("Consolas", 11, "bold"))

        # 各プレビュー項目の Text 内範囲を覚えておく
        # i 番目の項目: { "cb_start", "cb_end", "before_start", "before_end", "after_start", "after_end" }
        item_ranges: list[dict] = []

        def update_info_label():
            on_count = len(selected)
            total = len(change_indices)
            info_var.set(f"変更対象: {total} 行 / 採用: {on_count} 行")

        def render_all():
            """テキストを最初から書き直す。状態（ON/OFF）を反映"""
            # 再描画前に「現在画面の先頭に見えてる行」を記録
            # @0,0 = ウィジェット左上の座標 → そこにある行のインデックス
            try:
                top_index = preview_text.index("@0,0")
                top_line = int(top_index.split(".")[0])
            except Exception:
                top_line = 1

            preview_text.configure(state="normal")
            preview_text.delete("1.0", tk.END)
            item_ranges.clear()

            for i, (line_no, old_line, new_line) in enumerate(line_changes):
                orig_idx = change_indices[i]
                is_on = orig_idx in selected

                item = {}
                # チェックボックス [✓] or [ ]
                cb_start = preview_text.index("end-1c")
                if is_on:
                    preview_text.insert(tk.END, "[✓]", "checkbox_on")
                else:
                    preview_text.insert(tk.END, "[ ]", "checkbox_off")
                cb_end = preview_text.index("end-1c")
                item["cb_start"] = cb_start
                item["cb_end"] = cb_end

                # 連番＋行番号
                preview_text.insert(tk.END, f"  #{i+1:<3d} ", "preview_idx")
                preview_text.insert(tk.END, f"行 {line_no}\n", "preview_label")

                # 変更前
                preview_text.insert(tk.END, "      - ", "preview_label")
                bef_start = preview_text.index("end-1c")
                tag = "preview_before" if is_on else "preview_off_before"
                preview_text.insert(tk.END, f"{old_line}\n", tag)
                bef_end = preview_text.index("end-1c")
                item["before_start"] = bef_start
                item["before_end"] = bef_end

                # 変更後（コメント追記オプション反映）
                preview_text.insert(tk.END, "      + ", "preview_label")
                aft_start = preview_text.index("end-1c")
                tag = "preview_after" if is_on else "preview_off_after"
                if with_comment_var.get() == "1":
                    display_new = _make_commented_display(old_line, new_line)
                else:
                    display_new = new_line
                preview_text.insert(tk.END, f"{display_new}\n", tag)
                aft_end = preview_text.index("end-1c")
                item["after_start"] = aft_start
                item["after_end"] = aft_end

                preview_text.insert(tk.END, "\n")
                item_ranges.append(item)

            preview_text.configure(state="disabled")
            update_info_label()

            # 記録した先頭行を画面の先頭に戻す
            try:
                preview_text.see(f"{top_line}.0")
                # see は中央寄せになりがちなので、もう一段画面上端に寄せる
                preview_text.yview(f"{top_line}.0")
            except Exception:
                pass

        def toggle_item(i: int):
            """i番目の項目をトグル"""
            orig_idx = change_indices[i]
            if orig_idx in selected:
                selected.discard(orig_idx)
            else:
                selected.add(orig_idx)
            render_all()

        def select_all():
            selected.clear()
            selected.update(change_indices)
            render_all()

        def deselect_all():
            selected.clear()
            render_all()

        # クリックハンドラ：クリック位置がどの項目のチェックボックス範囲かを判定
        def on_text_click(event):
            try:
                idx = preview_text.index(f"@{event.x},{event.y}")
            except tk.TclError:
                return
            for i, item in enumerate(item_ranges):
                if preview_text.compare(idx, ">=", item["cb_start"]) and \
                   preview_text.compare(idx, "<", item["cb_end"]):
                    toggle_item(i)
                    return "break"

        preview_text.bind("<Button-1>", on_text_click)

        # 全選択/解除ボタン
        tk.Button(sel_btn_row, text="すべて選択", command=select_all, **sel_btn_kw).pack(side="left", padx=(0, 4))
        tk.Button(sel_btn_row, text="すべて解除", command=deselect_all, **sel_btn_kw).pack(side="left", padx=(0, 4))

        # 実行/キャンセルボタン
        btn_row = tk.Frame(outer, bg=BG_PANEL)
        btn_row.pack(fill="x", pady=(12, 0))
        btn_kw = {"font": ("Yu Gothic UI", 10, "bold"),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "solid", "bd": 1, "padx": 14, "pady": 4,
                  "cursor": "hand2"}
        btn_kw_primary = dict(btn_kw)
        btn_kw_primary["bg"] = "#123A38"
        btn_kw_primary["fg"] = ACCENT

        def _do_execute():
            if not selected:
                messagebox.showwarning("実行", "1つも選択されていません。", parent=dlg)
                return
            with_comment = with_comment_var.get() == "1"
            dlg.destroy()
            on_confirm(set(selected), with_comment)

        tk.Button(btn_row, text="実行する", command=_do_execute, **btn_kw_primary).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="キャンセル", command=dlg.destroy, **btn_kw).pack(side="left", padx=(0, 6))
        tk.Label(btn_row, text="※ Ctrl+Z で元に戻せます",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(side="right")

        # 初回描画
        render_all()

        # ダイアログ中央寄せ
        dlg.update_idletasks()
        try:
            px, py = self.winfo_rootx(), self.winfo_rooty()
            pw, ph = self.winfo_width(), self.winfo_height()
            dw, dh = dlg.winfo_width(), dlg.winfo_height()
            dlg.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 3}")
        except tk.TclError:
            pass
        dlg.grab_set()
        dlg.focus_set()

    def _apply_replacement(self, new_text: str, start_idx: str, end_idx: str, count: int, label: str) -> None:
        """テキスト置換を実行（Undo履歴1回分にまとめる）"""
        # delete と insert を1つのundo単位にまとめるため、autoseparatorsを一時的にOFF
        prev_autosep = self.target.cget("autoseparators")
        self.target.configure(autoseparators=False)
        try:
            self.target.edit_separator()  # 操作前に区切り
            self.target.delete(start_idx, end_idx)
            self.target.insert(start_idx, new_text)
            self.target.edit_separator()  # 操作後に区切り
        finally:
            self.target.configure(autoseparators=prev_autosep)
        self.app._update_linenumbers()
        self.app._highlight_all()
        self._status_var.set(f"{label}: {count}件 適用しました")
        messagebox.showinfo("完了", f"{count}件 適用しました。\n（Ctrl+Z で元に戻せます）", parent=self)
        # 置換後に自動でプログラムチェック実行（自動更新OFFでも走らせる）
        self._trigger_post_replace_check()

    def _trigger_post_replace_check(self) -> None:
        """置換完了後にプログラムチェックを再実行（軽量版）。
        自動更新トグルの ON/OFF に関わらず実行する（置換は明示アクションのため）。
        """
        try:
            program_text = self.app.input_text.get("1.0", "end-1c").rstrip("\n")
            if not program_text.strip():
                return
            # _auto_rerun_check はスキャンアニメ無しで軽量
            self.app._auto_rerun_check()
        except Exception:
            pass

    def _trigger_post_replace_check_debounced(self) -> None:
        """連打される可能性のある置換（replace_one 等）用のdebounce版"""
        if not hasattr(self, "_post_replace_debounce_id"):
            self._post_replace_debounce_id = None
        if self._post_replace_debounce_id is not None:
            try:
                self.app.root.after_cancel(self._post_replace_debounce_id)
            except Exception:
                pass
        self._post_replace_debounce_id = self.app.root.after(
            600, self._fire_debounced_replace_check
        )

    def _fire_debounced_replace_check(self) -> None:
        self._post_replace_debounce_id = None
        self._trigger_post_replace_check()

    def _compute_line_changes(self, scope: str, transform):
        """対象範囲のテキストに対して、行ごとに transform(line) を適用し、
        変更情報を集めて返す。
        Returns: dict {
            "original_lines": list[str],     # 元の行（改行込み）
            "new_lines":      list[str],     # 変換後の行（改行込み）
            "change_indices": list[int],     # original_lines のうち変更があったインデックス
            "line_changes":   list[(global_line_no, old_content, new_content)],
            "start_idx":      str,           # tk index 範囲開始
            "end_idx":        str,           # tk index 範囲終了
        }
        対象が無い場合は None を返す。
        """
        text, start_idx, end_idx = self._get_target_text(scope)
        if text is None:
            return None

        try:
            base_line = int(str(start_idx).split(".")[0])
        except (ValueError, AttributeError):
            base_line = 1

        original_lines = text.splitlines(keepends=True)
        new_lines = []
        change_indices = []
        line_changes = []
        for idx, line in enumerate(original_lines):
            content = line.rstrip("\r\n")
            ending = line[len(content):]
            new_content = transform(content)
            new_line = new_content + ending
            new_lines.append(new_line)
            if new_content != content:
                change_indices.append(idx)
                line_changes.append((base_line + idx, content, new_content))

        return {
            "original_lines": original_lines,
            "new_lines": new_lines,
            "change_indices": change_indices,
            "line_changes": line_changes,
            "start_idx": start_idx,
            "end_idx": end_idx,
        }

    def _run_template(self, title: str, label: str, scope: str, transform) -> None:
        """テンプレ実行の共通処理。
        変換関数 transform を行ごとに適用 → プレビュー（行ごとON/OFF選択） → 確認後に置換実行。
        """
        result = self._compute_line_changes(scope, transform)
        if result is None:
            return
        if not result["line_changes"]:
            messagebox.showinfo(label, "変更対象がありませんでした。", parent=self)
            self._status_var.set("一致なし")
            return

        def on_confirm(selected_set, with_comment=False):
            """selected_set: 採用する change_indices のサブセット
            with_comment: True なら採用行に (元: ...) コメントを追記
            """
            # 採用された行だけ new_lines を、それ以外は original_lines を使って再構築
            # original_lines / new_lines / line_changes の対応を取るため、change_indices→old_content マップを作る
            old_map = {}  # original_idx -> old_content
            for (line_no, old_c, new_c), orig_idx in zip(result["line_changes"], result["change_indices"]):
                old_map[orig_idx] = old_c

            assembled = []
            for idx, orig in enumerate(result["original_lines"]):
                if idx in result["change_indices"] and idx in selected_set:
                    new_line = result["new_lines"][idx]
                    if with_comment and idx in old_map:
                        # new_line は改行込み → 改行を一旦剥がしてコメント追記
                        content = new_line.rstrip("\r\n")
                        ending = new_line[len(content):]
                        new_line = self._build_commented_line(old_map[idx], content, ending)
                    assembled.append(new_line)
                else:
                    assembled.append(orig)
            new_text = "".join(assembled)
            applied = len(selected_set)
            self._apply_replacement(
                new_text, result["start_idx"], result["end_idx"], applied, label,
            )

        self._show_preview_dialog(title, result["line_changes"], result["change_indices"], on_confirm)

    # =========================================================================
    # マクロ用：params dict ベースの純粋変換ロジック（v28〜）
    # =========================================================================
    # 各テンプレートに対応する「params dict → transform 関数」のビルダー群。
    # 既存の _tpl_run_* GUI と並列に存在する。
    # マクロ実行時はこれらを使ってテキスト全体を順次変換する。
    # =========================================================================

    @staticmethod
    def _build_transform_f_scale(params: dict):
        ratio = float(params.get("factor", 1.0))
        pat = re.compile(r"(?<![A-Za-z])F([+-]?(?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def transform(line):
            return pat.sub(lambda m: f"F{float(m.group(1)) * ratio:g}", line)
        return transform

    @staticmethod
    def _build_transform_f_uniform(params: dict):
        new_val = str(params.get("value", "")).strip()
        pat = re.compile(r"(?<![A-Za-z])F([+-]?(?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def transform(line):
            return pat.sub(lambda m: f"F{new_val}", line)
        return transform

    @staticmethod
    def _build_transform_s_scale(params: dict):
        ratio = float(params.get("factor", 1.0))
        pat = re.compile(r"(?<![A-Za-z])S((?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def _sub(m):
            new_val = float(m.group(1)) * ratio
            if new_val == int(new_val):
                return f"S{int(new_val)}"
            return f"S{new_val:g}"
        def transform(line):
            return pat.sub(_sub, line)
        return transform

    @staticmethod
    def _build_transform_s_uniform(params: dict):
        new_val = str(params.get("value", "")).strip()
        pat = re.compile(r"(?<![A-Za-z])S((?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def transform(line):
            return pat.sub(lambda m: f"S{new_val}", line)
        return transform

    @staticmethod
    def _build_transform_axis_offset(params: dict):
        axis = str(params.get("axis", "X")).upper()
        offset = float(params.get("offset", 0.0))
        pat = re.compile(rf"(?<![A-Za-z]){axis}([+-]?(?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def _sub(m):
            try:
                new_val = float(m.group(1)) + offset
            except ValueError:
                return m.group(0)
            had_dot = "." in m.group(1)
            if had_dot:
                if new_val == int(new_val):
                    return f"{axis}{int(new_val)}."
                return f"{axis}{new_val:g}"
            else:
                if new_val == int(new_val):
                    return f"{axis}{int(new_val)}"
                return f"{axis}{new_val:g}"
        def transform(line):
            return pat.sub(_sub, line)
        return transform

    @staticmethod
    def _build_transform_axis_invert(params: dict):
        axis = str(params.get("axis", "Z")).upper()
        pat = re.compile(rf"(?<![A-Za-z]){axis}([+-]?(?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def _sub(m):
            try:
                v = float(m.group(1))
            except ValueError:
                return m.group(0)
            new_val = -v
            had_dot = "." in m.group(1)
            if had_dot:
                if new_val == int(new_val):
                    return f"{axis}{int(new_val)}."
                return f"{axis}{new_val:g}"
            else:
                return f"{axis}{int(new_val)}"
        def transform(line):
            return pat.sub(_sub, line)
        return transform

    @staticmethod
    def _build_transform_decimal_add(params: dict):
        axis_sel = str(params.get("axis", "全軸"))
        scale_sel = str(params.get("scale", "1."))

        if axis_sel == "全軸":
            axis_chars = "XYZUVWABCIJKFS"
        else:
            axis_chars = axis_sel

        if scale_sel == "1.":
            decimal_digits = 0
        else:
            parts = scale_sel.split(".")
            decimal_digits = len(parts[1]) if len(parts) > 1 else 0

        pat = re.compile(rf"(?<![A-Za-z])([{axis_chars}])([+-]?\d+)(?!\.)(?!\d)", re.IGNORECASE)

        def transform(line):
            def _replace(m):
                axis = m.group(1)
                num_str = m.group(2)
                if decimal_digits == 0:
                    return f"{axis}{num_str}."
                sign = ""
                digits = num_str
                if digits.startswith(("+", "-")):
                    sign = digits[0]
                    digits = digits[1:]
                    if sign == "+":
                        sign = ""
                if len(digits) <= decimal_digits:
                    padded = digits.rjust(decimal_digits + 1, "0")
                else:
                    padded = digits
                int_part = padded[:-decimal_digits] if decimal_digits > 0 else padded
                dec_part = padded[-decimal_digits:] if decimal_digits > 0 else ""
                int_part = int_part.lstrip("0") or "0"
                return f"{axis}{sign}{int_part}.{dec_part}"
            return pat.sub(_replace, line)
        return transform

    @staticmethod
    def _build_transform_upper_case(params: dict):
        def transform(line):
            placeholders = []
            def _hide(m):
                placeholders.append(m.group(0))
                return f"\x00{len(placeholders)-1}\x00"
            tmp = re.sub(r"\([^)]*\)", _hide, line)
            new_tmp = tmp.upper()
            def _restore(m):
                idx = int(m.group(1))
                return placeholders[idx]
            return re.sub(r"\x00(\d+)\x00", _restore, new_tmp)
        return transform

    @staticmethod
    def _build_transform_trim_spaces(params: dict):
        def transform(line):
            return line.strip()
        return transform

    @staticmethod
    def _build_transform_comment_add(params: dict):
        def transform(line):
            stripped = line.strip()
            if not stripped:
                return line
            if stripped.startswith("(") and stripped.endswith(")"):
                return line
            indent_len = len(line) - len(line.lstrip())
            indent = line[:indent_len]
            return f"{indent}({line[indent_len:]})"
        return transform

    @staticmethod
    def _build_transform_comment_del(params: dict):
        def transform(line):
            stripped = line.strip()
            if not (stripped.startswith("(") and stripped.endswith(")") and len(stripped) >= 2):
                return line
            indent_len = len(line) - len(line.lstrip())
            indent = line[:indent_len]
            inner = stripped[1:-1]
            return f"{indent}{inner}"
        return transform

    @staticmethod
    def _build_transform_comment_remove(params: dict):
        comment_pat = re.compile(r"\([^)]*\)")
        def transform(line):
            if not comment_pat.search(line):
                return line
            cleaned = comment_pat.sub("", line)
            cleaned = re.sub(r"[ \t]+", " ", cleaned)
            return cleaned.strip()
        return transform

    def _macro_apply_simple_replace(self, text: str, params: dict) -> tuple[str, int]:
        """シンプル置換（テキスト全体に対する置換）。
        改行をまたぐ検索・置換に対応。
        Returns: (new_text, match_count)
        """
        old = str(params.get("old", ""))
        new = str(params.get("new", ""))
        use_regex = bool(params.get("use_regex", False))
        case_sensitive = bool(params.get("case_sensitive", True))

        if not old:
            return text, 0

        flags = 0 if case_sensitive else re.IGNORECASE

        if use_regex:
            try:
                pat = re.compile(old, flags)
            except re.error:
                # 不正な正規表現 → 何もしない（マクロ実行時に止めない）
                return text, 0
            try:
                new_text, count = pat.subn(new, text)
            except re.error:
                return text, 0
            return new_text, count

        # 通常置換（テキスト全体に対して）
        if case_sensitive:
            count = text.count(old)
            return text.replace(old, new), count

        # 大小無視
        pat = re.compile(re.escape(old), re.IGNORECASE)
        new_text, count = pat.subn(new, text)
        return new_text, count

    # 行単位 transform で表現できないテンプレ（char_swap, line_delete, simple_replace）は
    # _macro_apply_step 側で個別ハンドラを呼ぶ
    _MACRO_LINE_TRANSFORM_BUILDERS = {
        "f_scale":        "_build_transform_f_scale",
        "f_uniform":      "_build_transform_f_uniform",
        "s_scale":        "_build_transform_s_scale",
        "s_uniform":      "_build_transform_s_uniform",
        "axis_offset":    "_build_transform_axis_offset",
        "axis_invert":    "_build_transform_axis_invert",
        "decimal_add":    "_build_transform_decimal_add",
        "upper_case":     "_build_transform_upper_case",
        "trim_spaces":    "_build_transform_trim_spaces",
        "comment_add":    "_build_transform_comment_add",
        "comment_del":    "_build_transform_comment_del",
        "comment_remove": "_build_transform_comment_remove",
    }

    def _macro_apply_char_swap(self, text: str, params: dict) -> tuple[str, int]:
        """char_swap をテキスト全体に適用。
        Returns: (new_text, change_count)
        """
        a = str(params.get("a", "")).strip()
        b = str(params.get("b", "")).strip()
        if not a or not b or a == b:
            return text, 0

        outer_s = str(params.get("outer_start", "")).strip()
        outer_e = str(params.get("outer_end", "")).strip()
        inner_s = str(params.get("inner_start", "")).strip()
        inner_e = str(params.get("inner_end", "")).strip()
        flex = bool(params.get("flex", False))
        skip_comment = bool(params.get("skip_comment", True))

        # 範囲指定の整合性チェック（不整合なら何もしない）
        if (outer_s and not outer_e) or (outer_e and not outer_s):
            return text, 0
        if (inner_s and not inner_e) or (inner_e and not inner_s):
            return text, 0
        if (inner_s or inner_e) and not (outer_s and outer_e):
            return text, 0

        original_lines = text.splitlines(keepends=True)

        # アクティブ行集合を計算
        if not outer_s:
            # 範囲指定なし → 全行対象
            active_indices = set(range(len(original_lines)))
        else:
            active_indices = self._compute_swap_active_indices(
                original_lines, outer_s, outer_e, inner_s, inner_e
            )

        if not active_indices:
            return text, 0

        # 共通ヘルパーで swap_one を構築（_run_template_char_swap と同一ロジック）
        swap_one = self._make_char_swap_fn(a, b, flex, skip_comment)

        new_lines = []
        change_count = 0
        for idx, line in enumerate(original_lines):
            content = line.rstrip("\r\n")
            ending = line[len(content):]
            if idx in active_indices:
                new_content = swap_one(content)
            else:
                new_content = content
            if new_content != content:
                change_count += 1
            new_lines.append(new_content + ending)

        return "".join(new_lines), change_count

    def _macro_apply_line_delete(self, text: str, params: dict) -> tuple[str, int]:
        """行削除：空行・パターン一致行を削除する。
        Returns: (new_text, deleted_count)
        params:
            delete_empty: bool          # 空行を削除
            delete_pattern: bool        # パターンに一致する行を削除
            pattern: str                # 検索パターン
            use_regex: bool             # 正規表現として扱う
            case_sensitive: bool        # 大文字小文字区別
        """
        delete_empty = bool(params.get("delete_empty", False))
        delete_pattern = bool(params.get("delete_pattern", False))
        pattern = str(params.get("pattern", ""))
        use_regex = bool(params.get("use_regex", False))
        case_sensitive = bool(params.get("case_sensitive", True))

        # 何も削除条件が無効なら何もしない
        if not delete_empty and not (delete_pattern and pattern):
            return text, 0

        # パターン一致判定関数を構築
        match_func = None
        if delete_pattern and pattern:
            flags = 0 if case_sensitive else re.IGNORECASE
            if use_regex:
                try:
                    pat = re.compile(pattern, flags)
                except re.error:
                    # 不正な正規表現 → パターン削除はスキップ
                    pat = None
                if pat is not None:
                    match_func = lambda s: bool(pat.search(s))
            else:
                if case_sensitive:
                    needle = pattern
                    match_func = lambda s: needle in s
                else:
                    needle_lower = pattern.lower()
                    match_func = lambda s: needle_lower in s.lower()

        # 行ごとに判定
        lines = text.splitlines(keepends=True)
        new_lines = []
        deleted = 0

        for line in lines:
            content = line.rstrip("\r\n")
            should_delete = False

            if delete_empty and not content.strip():
                should_delete = True
            elif match_func is not None and match_func(content):
                should_delete = True

            if should_delete:
                deleted += 1
            else:
                new_lines.append(line)

        return "".join(new_lines), deleted

    def macro_apply_step(self, text: str, step: "MacroStep") -> tuple[str, int]:
        """マクロの1ステップを text 全体に適用。
        Returns: (new_text, change_count)
        無効化されたステップ・未知のテンプレIDは (text, 0) を返す。
        """
        if not step.enabled:
            return text, 0

        # simple_replace（テキスト全体に対する置換、改行をまたぐ検索OK）
        if step.template_id == "simple_replace":
            return self._macro_apply_simple_replace(text, step.params)

        # char_swap は特殊（行単位 transform にできない）
        if step.template_id == "char_swap":
            return self._macro_apply_char_swap(text, step.params)

        # line_delete も特殊（行を消すので「行数が変わる」）
        if step.template_id == "line_delete":
            return self._macro_apply_line_delete(text, step.params)

        # 通常テンプレ：行単位 transform を適用
        builder_name = self._MACRO_LINE_TRANSFORM_BUILDERS.get(step.template_id)
        if not builder_name:
            return text, 0

        try:
            builder = getattr(self, builder_name)
            transform = builder(step.params)
        except (ValueError, KeyError, TypeError):
            return text, 0

        original_lines = text.splitlines(keepends=True)
        new_lines = []
        change_count = 0
        for line in original_lines:
            content = line.rstrip("\r\n")
            ending = line[len(content):]
            try:
                new_content = transform(content)
            except (ValueError, TypeError):
                new_content = content
            if new_content != content:
                change_count += 1
            new_lines.append(new_content + ending)

        return "".join(new_lines), change_count

    # =========================================================================
    # テンプレ実装（17個）
    # =========================================================================


    # --- 1. F値倍率変更 ---
    def _tpl_ui_f_scale(self, parent):
        e = self._make_tpl_entry(parent, "倍率（例: 0.8）:", "0.8")
        scope = self._make_tpl_target_radio(parent)
        self._tpl_state = {"entry": e, "scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_f_scale())

    def _tpl_run_f_scale(self):
        try:
            ratio = float(self._tpl_state["entry"]._var.get())
        except ValueError:
            self._status_var.set("⚠ 倍率は数値で入力してください")
            return
        scope = self._tpl_state["scope"]
        # 数値パターン: F1, F1., F.5, F1.5, F-2.5 全部対応
        # 直前が英字でない（OFF1 みたいな単語の中の F は弾く）
        pat = re.compile(r"(?<![A-Za-z])F([+-]?(?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def transform(line):
            return pat.sub(lambda m: f"F{float(m.group(1)) * ratio:g}", line)
        self._run_template("F値倍率変更 プレビュー", "F値倍率変更", scope, transform)

    # --- 2. F値統一 ---
    def _tpl_ui_f_uniform(self, parent):
        e = self._make_tpl_entry(parent, "新しいF値:", "0.15")
        scope = self._make_tpl_target_radio(parent)
        self._tpl_state = {"entry": e, "scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_f_uniform())

    def _tpl_run_f_uniform(self):
        new_val = self._tpl_state["entry"]._var.get().strip()
        if not new_val:
            self._status_var.set("⚠ F値を入力してください")
            return
        scope = self._tpl_state["scope"]
        pat = re.compile(r"(?<![A-Za-z])F([+-]?(?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def transform(line):
            return pat.sub(lambda m: f"F{new_val}", line)
        self._run_template("F値統一 プレビュー", "F値統一", scope, transform)

    # --- 3. S値倍率変更 ---
    def _tpl_ui_s_scale(self, parent):
        e = self._make_tpl_entry(parent, "倍率（例: 0.8）:", "0.8")
        scope = self._make_tpl_target_radio(parent)
        self._tpl_state = {"entry": e, "scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_s_scale())

    def _tpl_run_s_scale(self):
        try:
            ratio = float(self._tpl_state["entry"]._var.get())
        except ValueError:
            self._status_var.set("⚠ 倍率は数値で入力してください")
            return
        scope = self._tpl_state["scope"]
        # S100, S100., S2000.5 など対応
        # 直前が英字でない（G96S100, M3S2000 みたいな詰め書きも拾う）
        pat = re.compile(r"(?<![A-Za-z])S((?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def _sub(m):
            new_val = float(m.group(1)) * ratio
            if new_val == int(new_val):
                return f"S{int(new_val)}"
            return f"S{new_val:g}"
        def transform(line):
            return pat.sub(_sub, line)
        self._run_template("S値倍率変更 プレビュー", "S値倍率変更", scope, transform)

    # --- 4. S値統一 ---
    def _tpl_ui_s_uniform(self, parent):
        e = self._make_tpl_entry(parent, "新しいS値:", "2000")
        scope = self._make_tpl_target_radio(parent)
        self._tpl_state = {"entry": e, "scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_s_uniform())

    def _tpl_run_s_uniform(self):
        new_val = self._tpl_state["entry"]._var.get().strip()
        if not new_val:
            self._status_var.set("⚠ S値を入力してください")
            return
        scope = self._tpl_state["scope"]
        pat = re.compile(r"(?<![A-Za-z])S((?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def transform(line):
            return pat.sub(lambda m: f"S{new_val}", line)
        self._run_template("S値統一 プレビュー", "S値統一", scope, transform)

    # --- 5. 座標オフセット加算 ---
    def _tpl_ui_axis_offset(self, parent):
        ax = self._make_tpl_combobox(parent, "対象軸:", ["X", "Y", "Z"], "X")
        e = self._make_tpl_entry(parent, "加算値（例: 5.0、-3.0）:", "5.0")
        scope = self._make_tpl_target_radio(parent)
        self._tpl_state = {"axis": ax, "entry": e, "scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_axis_offset())

    def _tpl_run_axis_offset(self):
        axis = self._tpl_state["axis"].get()
        try:
            offset = float(self._tpl_state["entry"]._var.get())
        except ValueError:
            self._status_var.set("⚠ 加算値は数値で入力してください")
            return
        scope = self._tpl_state["scope"]
        # X1, X1., X.5, X-2.5 全部対応
        # 直前が英字でない（G0X10, G91X-5 みたいな詰め書きも拾う）
        pat = re.compile(rf"(?<![A-Za-z]){axis}([+-]?(?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def _sub(m):
            try:
                new_val = float(m.group(1)) + offset
            except ValueError:
                return m.group(0)
            had_dot = "." in m.group(1)
            if had_dot:
                if new_val == int(new_val):
                    return f"{axis}{int(new_val)}."
                return f"{axis}{new_val:g}"
            else:
                if new_val == int(new_val):
                    return f"{axis}{int(new_val)}"
                return f"{axis}{new_val:g}"
        def transform(line):
            return pat.sub(_sub, line)
        self._run_template(f"{axis}軸オフセット加算 プレビュー", f"{axis}軸オフセット加算", scope, transform)

    # --- 6. 座標符号反転 ---
    def _tpl_ui_axis_invert(self, parent):
        ax = self._make_tpl_combobox(parent, "対象軸:", ["X", "Y", "Z"], "Z")
        scope = self._make_tpl_target_radio(parent)
        self._tpl_state = {"axis": ax, "scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_axis_invert())

    def _tpl_run_axis_invert(self):
        axis = self._tpl_state["axis"].get()
        scope = self._tpl_state["scope"]
        pat = re.compile(rf"(?<![A-Za-z]){axis}([+-]?(?:\d+\.\d*|\.\d+|\d+))", re.IGNORECASE)
        def _sub(m):
            try:
                v = float(m.group(1))
            except ValueError:
                return m.group(0)
            new_val = -v
            had_dot = "." in m.group(1)
            if had_dot:
                if new_val == int(new_val):
                    return f"{axis}{int(new_val)}."
                return f"{axis}{new_val:g}"
            else:
                return f"{axis}{int(new_val)}"
        def transform(line):
            return pat.sub(_sub, line)
        self._run_template(f"{axis}軸符号反転 プレビュー", f"{axis}軸符号反転", scope, transform)

    # --- 8. 小数点付加 ---
    DECIMAL_ADD_AXES = ["全軸", "X", "Y", "Z", "U", "V", "W", "A", "B", "C", "I", "J", "K", "F", "S"]
    DECIMAL_ADD_SCALES = ["1.", "0.1", "0.01", "0.001", "0.0001"]

    def _tpl_ui_decimal_add(self, parent):
        scope = self._make_tpl_target_radio(parent)
        axis_var = self._make_tpl_combobox(parent, "対象軸:", self.DECIMAL_ADD_AXES, "全軸")
        scale_var = self._make_tpl_combobox(parent, "スケール:", self.DECIMAL_ADD_SCALES, "1.")
        tk.Label(parent,
                 text="※ スケール 1. = 末尾に「.」付加（例: F100 → F100.）\n"
                      "※ スケール 0.01 = 100で割って小数2桁表示（例: F100 → F1.00）\n"
                      "※ 既に小数点を含む値はスキップします",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                 wraplength=420, justify="left").pack(anchor="w", pady=(8, 0))
        self._tpl_state = {"scope": scope, "axis": axis_var, "scale": scale_var}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_decimal_add())

    def _tpl_run_decimal_add(self):
        scope = self._tpl_state["scope"]
        axis_sel = self._tpl_state["axis"].get()
        scale_sel = self._tpl_state["scale"].get()

        # 対象軸記号セットを決定
        if axis_sel == "全軸":
            axis_chars = "XYZUVWABCIJKFS"
        else:
            axis_chars = axis_sel

        # スケールから「小数点以下桁数」を算出
        # "1." → 0桁（整数として末尾に . を付加）
        # "0.1" → 1桁、"0.01" → 2桁、"0.001" → 3桁、"0.0001" → 4桁
        if scale_sel == "1.":
            decimal_digits = 0
        else:
            # "0.001" → 小数部 "001" → 3桁
            parts = scale_sel.split(".")
            decimal_digits = len(parts[1]) if len(parts) > 1 else 0

        pat = re.compile(rf"(?<![A-Za-z])([{axis_chars}])([+-]?\d+)(?!\.)(?!\d)", re.IGNORECASE)

        def transform(line):
            def _replace(m):
                axis = m.group(1)
                num_str = m.group(2)  # 符号付き整数文字列
                if decimal_digits == 0:
                    # 末尾に . 付加するだけ
                    return f"{axis}{num_str}."
                # スケール処理：右辺の小数点位置に合わせて元の数値を割る
                # 例: 0.01（2桁）の場合、num_str を 100 で割った値を 2 桁固定で表示
                sign = ""
                digits = num_str
                if digits.startswith(("+", "-")):
                    sign = digits[0]
                    digits = digits[1:]
                    if sign == "+":
                        sign = ""  # +は表示しない慣例
                # ゼロパディングして小数点を挿入
                if len(digits) <= decimal_digits:
                    # 桁数不足 → 先頭に0を補う（例: F1, 0.01 → 0.01）
                    padded = digits.rjust(decimal_digits + 1, "0")
                else:
                    padded = digits
                int_part = padded[:-decimal_digits] if decimal_digits > 0 else padded
                dec_part = padded[-decimal_digits:] if decimal_digits > 0 else ""
                # 整数部の先頭ゼロ削除（ただし1桁は残す）
                int_part = int_part.lstrip("0") or "0"
                return f"{axis}{sign}{int_part}.{dec_part}"

            return pat.sub(_replace, line)

        self._run_template("小数点付加 プレビュー", "小数点付加", scope, transform)

    # --- 9. 小文字→大文字統一 ---
    def _tpl_ui_upper_case(self, parent):
        scope = self._make_tpl_target_radio(parent)
        tk.Label(parent, text="※ コメント (...) 内は変更しません",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                 wraplength=380, justify="left").pack(anchor="w", pady=(8, 0))
        self._tpl_state = {"scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_upper_case())

    def _tpl_run_upper_case(self):
        scope = self._tpl_state["scope"]
        def transform(line):
            # コメント部分を退避
            placeholders = []
            def _hide(m):
                placeholders.append(m.group(0))
                return f"\x00{len(placeholders)-1}\x00"
            tmp = re.sub(r"\([^)]*\)", _hide, line)
            new_tmp = tmp.upper()
            def _restore(m):
                idx = int(m.group(1))
                return placeholders[idx]
            return re.sub(r"\x00(\d+)\x00", _restore, new_tmp)
        self._run_template("大文字統一 プレビュー", "大文字統一", scope, transform)

    # --- 10. 余分な空白除去 ---
    def _tpl_ui_trim_spaces(self, parent):
        scope = self._make_tpl_target_radio(parent)
        tk.Label(parent, text="※ 各行の行頭・行末の空白を除去します",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                 wraplength=380, justify="left").pack(anchor="w", pady=(8, 0))
        self._tpl_state = {"scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_trim_spaces())

    def _tpl_run_trim_spaces(self):
        scope = self._tpl_state["scope"]
        def transform(line):
            return line.strip()
        self._run_template("空白除去 プレビュー", "空白除去", scope, transform)

    # --- 14. コメント化 ---
    def _tpl_ui_comment_add(self, parent):
        scope = self._make_tpl_target_radio(parent)
        tk.Label(parent, text="※ 対象範囲の各行を ( ) で囲みます。\n"
                              "　 「選択範囲」モードは範囲選択してから実行してください。",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                 wraplength=380, justify="left").pack(anchor="w", pady=(8, 0))
        self._tpl_state = {"scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_comment_add())

    def _tpl_run_comment_add(self):
        scope = self._tpl_state["scope"]
        def transform(line):
            stripped = line.strip()
            if not stripped:
                return line
            if stripped.startswith("(") and stripped.endswith(")"):
                return line
            indent_len = len(line) - len(line.lstrip())
            indent = line[:indent_len]
            return f"{indent}({line[indent_len:]})"
        self._run_template("コメント化 プレビュー", "コメント化", scope, transform)

    # --- 15. コメント解除 ---
    def _tpl_ui_comment_del(self, parent):
        scope = self._make_tpl_target_radio(parent)
        tk.Label(parent, text="※ 行全体が ( ) で囲まれている行のみカッコを外します",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                 wraplength=380, justify="left").pack(anchor="w", pady=(8, 0))
        self._tpl_state = {"scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_comment_del())

    def _tpl_run_comment_del(self):
        scope = self._tpl_state["scope"]
        def transform(line):
            stripped = line.strip()
            if not (stripped.startswith("(") and stripped.endswith(")") and len(stripped) >= 2):
                return line
            indent_len = len(line) - len(line.lstrip())
            indent = line[:indent_len]
            inner = stripped[1:-1]
            return f"{indent}{inner}"
        self._run_template("コメント解除 プレビュー", "コメント解除", scope, transform)

    # --- 15-2. コメント削除（行内の ( ) すべて削除） ---
    def _tpl_ui_comment_remove(self, parent):
        scope = self._make_tpl_target_radio(parent)
        tk.Label(parent, text="※ 行内の ( ) コメントをすべて削除します。\n"
                              "　 行全体がコメントの場合は空行として残ります。\n"
                              "　 連続空白は1つにまとめ、行末空白も除去します。",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                 wraplength=420, justify="left").pack(anchor="w", pady=(8, 0))
        self._tpl_state = {"scope": scope}
        self._make_tpl_buttons(parent, lambda: self._tpl_run_comment_remove())

    def _tpl_run_comment_remove(self):
        scope = self._tpl_state["scope"]
        comment_pat = re.compile(r"\([^)]*\)")
        def transform(line):
            # コメントが含まれない行は触らない（インデント保持のため）
            if not comment_pat.search(line):
                return line
            # コメントを全削除
            cleaned = comment_pat.sub("", line)
            # 連続空白を1つに圧縮
            cleaned = re.sub(r"[ \t]+", " ", cleaned)
            # 両端の空白除去（コメント削除で生じた余分な空白を一掃）
            return cleaned.strip()
        self._run_template("コメント削除 プレビュー", "コメント削除", scope, transform)

    # --- 18. 文字入替（範囲指定2段ネスト対応） ---
    def _tpl_ui_char_swap(self, parent):
        # ⚠ 警告枠：双方向入替の説明
        warn_frame = tk.Frame(parent, bg=WARNING_BG,
                               highlightbackground=WARNING_BORDER,
                               highlightthickness=2)
        warn_frame.pack(fill="x", pady=(4, 10), padx=2)
        warn_inner = tk.Frame(warn_frame, bg=WARNING_BG)
        warn_inner.pack(fill="x", padx=10, pady=8)
        tk.Label(warn_inner,
                 text="⚠ これは「双方向入替」です（一方向の置換ではありません）",
                 bg=WARNING_BG, fg=WARNING_TITLE,
                 font=("Yu Gothic UI", 10, "bold"),
                 anchor="w").pack(fill="x")
        tk.Label(warn_inner,
                 text="A↔B が同時に入れ替わります（A→B かつ B→A を一度に実行）\n"
                      "例: A=G2, B=G3 で実行 →  元の G2行は G3 に変換／元の G3行は G2 に変換\n"
                      "※ 一方向の置換が目的なら「シンプル置換」タブを使ってください\n"
                      "　（シンプル置換にも範囲指定機能があります）",
                 bg=WARNING_BG, fg=WARNING_TEXT,
                 font=("Yu Gothic UI", 9),
                 wraplength=540, justify="left",
                 anchor="w").pack(fill="x", pady=(4, 0))

        # 入替ペア：A ⇄ B を視覚的に示す
        pair_frame = tk.Frame(parent, bg=BG_PANEL)
        pair_frame.pack(fill="x", pady=(2, 4))

        # 行1: 入替対象A
        row_a = tk.Frame(pair_frame, bg=BG_PANEL)
        row_a.pack(fill="x", pady=2)
        tk.Label(row_a, text="入替対象A:", bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="e").pack(side="left", padx=(0, 8))
        var_a = tk.StringVar(value="")
        e_a = tk.Entry(row_a, textvariable=var_a, font=("Consolas", 11),
                        bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT,
                        relief="solid", bd=1, highlightthickness=1,
                        highlightbackground=BORDER, highlightcolor=ACCENT, width=14)
        e_a.pack(side="left")
        e_a._var = var_a  # type: ignore

        # 行2: ⇄ 矢印（中央寄せ、目立つ色）
        arrow_row = tk.Frame(pair_frame, bg=BG_PANEL)
        arrow_row.pack(fill="x")
        tk.Label(arrow_row, text="", bg=BG_PANEL, width=18).pack(side="left", padx=(0, 8))
        tk.Label(arrow_row, text="⇅  双方向入替  ⇅",
                 bg=BG_PANEL, fg=WARNING_TITLE,
                 font=("Yu Gothic UI", 10, "bold"),
                 anchor="w").pack(side="left")

        # 行3: 入替対象B
        row_b = tk.Frame(pair_frame, bg=BG_PANEL)
        row_b.pack(fill="x", pady=2)
        tk.Label(row_b, text="入替対象B:", bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="e").pack(side="left", padx=(0, 8))
        var_b = tk.StringVar(value="")
        e_b = tk.Entry(row_b, textvariable=var_b, font=("Consolas", 11),
                        bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT,
                        relief="solid", bd=1, highlightthickness=1,
                        highlightbackground=BORDER, highlightcolor=ACCENT, width=14)
        e_b.pack(side="left")
        e_b._var = var_b  # type: ignore

        # マッチモード（チェックボックス）
        flex_var = tk.StringVar(value="0")  # "0"=厳密、"1"=柔軟
        flex_row = tk.Frame(parent, bg=BG_PANEL)
        flex_row.pack(fill="x", pady=4)
        tk.Label(flex_row, text="", bg=BG_PANEL, width=18).pack(side="left", padx=(0, 8))
        tk.Checkbutton(
            flex_row, text="柔軟マッチ（境界を無視して全部ヒット／プレビューで取捨選択）",
            variable=flex_var, onvalue="1", offvalue="0",
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            font=("Yu Gothic UI", 9), bd=0, highlightthickness=0,
        ).pack(side="left")

        # コメント内除外（デフォルトON、NC用途向け）
        skip_comment_var = tk.StringVar(value="1")
        sc_row = tk.Frame(parent, bg=BG_PANEL)
        sc_row.pack(fill="x", pady=4)
        tk.Label(sc_row, text="", bg=BG_PANEL, width=18).pack(side="left", padx=(0, 8))
        tk.Checkbutton(
            sc_row, text="コメント内 ( ... ) は置換しない（推奨ON）",
            variable=skip_comment_var, onvalue="1", offvalue="0",
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            font=("Yu Gothic UI", 9), bd=0, highlightthickness=0,
        ).pack(side="left")

        # 区切り線
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
        tk.Label(parent, text="範囲指定（任意）",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", padx=4)

        e_outer_start = self._make_tpl_entry(parent, "外側 開始:", "")
        e_outer_end = self._make_tpl_entry(parent, "外側 終端:", "")
        e_inner_start = self._make_tpl_entry(parent, "内側 開始:", "")
        e_inner_end = self._make_tpl_entry(parent, "内側 終端:", "")

        scope = self._make_tpl_target_radio(parent)

        tk.Label(parent,
                 text="※ 範囲指定なし: 全体で A↔B を入替\n"
                      "※ 外側のみ指定: 外側範囲内のみ入替（境界行は除外）\n"
                      "※ 外側＋内側指定: 外側内に内側がある場合のみ、内側範囲で入替\n"
                      "※ 終端が見つからない範囲は無視されます\n"
                      "※ 厳密マッチ（既定）: 後ろが数字で続く場合のみ除外\n"
                      "   例: G2X10 ✓ / G03Z-5. ✓ / G20 ✗ / G32 ✗\n"
                      "※ 柔軟マッチ: 境界を無視して部分一致もヒット（G1 が G10 の一部にもヒット）",
                 bg=BG_PANEL, fg=TEXT_MUTED, font=("Yu Gothic UI", 9),
                 wraplength=460, justify="left").pack(anchor="w", pady=(8, 0))

        self._tpl_state = {
            "a": e_a, "b": e_b,
            "flex": flex_var,
            "skip_comment": skip_comment_var,
            "outer_start": e_outer_start, "outer_end": e_outer_end,
            "inner_start": e_inner_start, "inner_end": e_inner_end,
            "scope": scope,
        }
        self._make_tpl_buttons(parent, lambda: self._tpl_run_char_swap())

    def _tpl_run_char_swap(self):
        a = self._tpl_state["a"]._var.get().strip()
        b = self._tpl_state["b"]._var.get().strip()
        if not a or not b:
            self._status_var.set("⚠ 入替対象 文字A / 文字B を入力してください")
            return
        if a == b:
            self._status_var.set("⚠ 文字Aと文字Bが同じです")
            return

        outer_s = self._tpl_state["outer_start"]._var.get().strip()
        outer_e = self._tpl_state["outer_end"]._var.get().strip()
        inner_s = self._tpl_state["inner_start"]._var.get().strip()
        inner_e = self._tpl_state["inner_end"]._var.get().strip()

        # 範囲指定の整合性チェック
        if (outer_s and not outer_e) or (outer_e and not outer_s):
            self._status_var.set("⚠ 外側範囲は開始と終端を両方指定してください")
            return
        if (inner_s and not inner_e) or (inner_e and not inner_s):
            self._status_var.set("⚠ 内側範囲は開始と終端を両方指定してください")
            return
        if (inner_s or inner_e) and not (outer_s and outer_e):
            self._status_var.set("⚠ 内側範囲を使うには外側範囲の指定が必要です")
            return

        scope = self._tpl_state["scope"]
        flex = self._tpl_state["flex"].get() == "1"
        skip_comment = self._tpl_state["skip_comment"].get() == "1"
        self._run_template_char_swap(scope, a, b, outer_s, outer_e, inner_s, inner_e, flex, skip_comment)

    # --- 行削除 ---
    def _tpl_ui_line_delete(self, parent):
        del_empty_var = tk.StringVar(value="1")
        cb1 = tk.Checkbutton(
            parent, text="空行を削除する", variable=del_empty_var,
            onvalue="1", offvalue="0",
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            font=("Yu Gothic UI", 10), bd=0, highlightthickness=0,
        )
        cb1.pack(anchor="w", pady=(2, 6), padx=4)

        # 区切り
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=(2, 6))

        del_pat_var = tk.StringVar(value="0")
        cb2 = tk.Checkbutton(
            parent, text="パターンに一致する行を削除する", variable=del_pat_var,
            onvalue="1", offvalue="0",
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            font=("Yu Gothic UI", 10), bd=0, highlightthickness=0,
        )
        cb2.pack(anchor="w", pady=(2, 4), padx=4)

        pat_entry = self._make_tpl_entry(parent, "パターン:", "")

        regex_var = tk.StringVar(value="0")
        cb3 = tk.Checkbutton(
            parent, text="正規表現として扱う", variable=regex_var,
            onvalue="1", offvalue="0",
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            font=("Yu Gothic UI", 9), bd=0, highlightthickness=0,
        )
        cb3.pack(anchor="w", pady=(2, 2), padx=20)

        case_var = tk.StringVar(value="1")
        cb4 = tk.Checkbutton(
            parent, text="大文字小文字を区別する", variable=case_var,
            onvalue="1", offvalue="0",
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            font=("Yu Gothic UI", 9), bd=0, highlightthickness=0,
        )
        cb4.pack(anchor="w", pady=(2, 6), padx=20)

        tk.Label(parent,
                 text="※ 両方ON可（空行＋パターン一致を一括削除）\n"
                      "※ パターン一致は「行に部分一致」した行を全削除\n"
                      "  例: 'G04 P' → G04 P を含む行を全削除\n"
                      "※ プレビューで実行前に確認できます",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9), wraplength=420,
                 justify="left").pack(anchor="w", padx=4, pady=(4, 8))

        # ビルダーボタン
        builder_row = tk.Frame(parent, bg=BG_PANEL)
        builder_row.pack(fill="x", pady=(0, 6), padx=4)
        tk.Button(
            builder_row, text="🔧 正規表現ビルダーを開く",
            command=self._open_regex_builder_for_line_delete,
            font=("Yu Gothic UI", 9),
            bg="#141B24", fg=ACCENT,
            activebackground="#1F2B3A", activeforeground=ACCENT,
            relief="solid", bd=1, padx=10, pady=2, cursor="hand2",
        ).pack(side="left")

        self._tpl_state = {
            "delete_empty": del_empty_var,
            "delete_pattern": del_pat_var,
            "pattern": pat_entry,
            "use_regex": regex_var,
            "case_sensitive": case_var,
        }
        self._make_tpl_buttons(parent, lambda: self._tpl_run_line_delete())

    def _open_regex_builder_for_line_delete(self) -> None:
        """テンプレタブの行削除パターン用ビルダー"""
        try:
            initial_pattern = self._tpl_state["pattern"]._var.get()
        except (KeyError, AttributeError):
            initial_pattern = ""

        ignore_case = False
        try:
            ignore_case = (self._tpl_state["case_sensitive"].get() == "0")
        except KeyError:
            pass

        def apply_result(pattern: str):
            self._tpl_state["pattern"]._var.set(pattern)
            self._tpl_state["delete_pattern"].set("1")
            self._tpl_state["use_regex"].set("1")

        self._show_regex_builder(
            initial_pattern=initial_pattern,
            ignore_case=ignore_case,
            on_result=apply_result,
        )

    def _tpl_run_line_delete(self):
        delete_empty = self._tpl_state["delete_empty"].get() == "1"
        delete_pattern = self._tpl_state["delete_pattern"].get() == "1"
        pattern = self._tpl_state["pattern"]._var.get().strip()
        use_regex = self._tpl_state["use_regex"].get() == "1"
        case_sensitive = self._tpl_state["case_sensitive"].get() == "1"

        if not delete_empty and not delete_pattern:
            self._status_var.set("⚠ 「空行削除」または「パターン削除」のどちらかをONに")
            return

        if delete_pattern and not pattern:
            self._status_var.set("⚠ パターン削除ONですが、パターンが未入力")
            return

        # 正規表現の妥当性チェック
        if delete_pattern and use_regex:
            try:
                flags = 0 if case_sensitive else re.IGNORECASE
                re.compile(pattern, flags)
            except re.error as e:
                self._status_var.set(f"⚠ 正規表現エラー: {e}")
                return

        # 全文取得 → 行削除実行
        original = self.target.get("1.0", "end-1c")
        new_text, deleted = self._macro_apply_line_delete(original, {
            "delete_empty": delete_empty,
            "delete_pattern": delete_pattern,
            "pattern": pattern,
            "use_regex": use_regex,
            "case_sensitive": case_sensitive,
        })

        if deleted == 0:
            messagebox.showinfo("行削除", "削除対象の行が見つかりませんでした。", parent=self)
            self._status_var.set("削除対象なし")
            return

        # MacroPreviewDialog を再利用してプレビュー表示
        step_logs = [("step", f"行削除: {deleted}行を削除します\n")]
        preview = MacroPreviewDialog(
            self,
            original_text=original,
            final_text=new_text,
            step_logs=step_logs,
            total_changes=deleted,
            executed=1,
            skipped=0,
        )
        preview.transient(self)
        preview.grab_set()
        self.wait_window(preview)

        if not preview.confirmed:
            self._status_var.set("行削除をキャンセルしました")
            return

        # 適用
        self._atomic_replace(self.target, "1.0", "end", new_text)
        self.app._update_linenumbers()
        self.app._highlight_all()
        self._status_var.set(f"行削除: {deleted}行を削除しました")
        try:
            self._trigger_post_replace_check_debounced()
        except Exception:
            pass
        messagebox.showinfo("完了",
                            f"{deleted}行を削除しました。\n（Ctrl+Z で元に戻せます）",
                            parent=self)

    @staticmethod
    def _make_char_swap_fn(a: str, b: str, flex: bool, skip_comment: bool):
        """A↔B 双方向入替を行う1行変換関数を構築して返す。
        _run_template_char_swap と _macro_apply_char_swap で共通利用。

        Args:
            a, b: 入替対象の文字列
            flex: True=柔軟マッチ（部分一致もヒット） / False=厳密マッチ（後が数字・小数点でない）
            skip_comment: True=行内の (...) コメント部分は保護する

        Returns:
            1行を受け取って入替後の1行を返す関数
        """
        if flex:
            pat_a = re.compile(re.escape(a), re.IGNORECASE)
            pat_b = re.compile(re.escape(b), re.IGNORECASE)
        else:
            pat_a = re.compile(rf"(?<![A-Za-z0-9_]){re.escape(a)}(?!\d)(?!\.)", re.IGNORECASE)
            pat_b = re.compile(rf"(?<![A-Za-z0-9_]){re.escape(b)}(?!\d)(?!\.)", re.IGNORECASE)
        placeholder = "\x00SWAP_PLACEHOLDER\x00"
        cmt_prefix = "\x00CMT_"
        cmt_suffix = "\x00"
        comment_re = re.compile(r"\([^)]*\)")
        cmt_restore_re = re.compile(
            rf"{re.escape(cmt_prefix)}(\d+){re.escape(cmt_suffix)}"
        )

        def swap_one(content: str) -> str:
            if skip_comment:
                # コメント区間をプレースホルダに退避
                stash: list[str] = []
                def _store(m):
                    idx = len(stash)
                    stash.append(m.group(0))
                    return f"{cmt_prefix}{idx}{cmt_suffix}"
                masked = comment_re.sub(_store, content)
                tmp = pat_a.sub(placeholder, masked)
                tmp = pat_b.sub(a, tmp)
                replaced = tmp.replace(placeholder, b)
                return cmt_restore_re.sub(lambda m: stash[int(m.group(1))], replaced)
            else:
                tmp = pat_a.sub(placeholder, content)
                tmp = pat_b.sub(a, tmp)
                return tmp.replace(placeholder, b)

        return swap_one

    def _run_template_char_swap(self, scope, a, b, outer_s, outer_e, inner_s, inner_e, flex=False, skip_comment=True):
        """文字入替専用のrun_template。
        対象テキストを取得 → 範囲スキャンでアクティブ行集合作成 →
        該当行のみ入替 → プレビュー → 確認後に置換実行
        flex=True の場合は境界条件を無視して部分一致もヒットさせる
        skip_comment=True の場合、行中の (...) コメント部分は置換対象から外す
        """
        text, start_idx, end_idx = self._get_target_text(scope)
        if text is None:
            return

        try:
            base_line = int(str(start_idx).split(".")[0])
        except (ValueError, AttributeError):
            base_line = 1

        original_lines = text.splitlines(keepends=True)

        # 入替アクティブな行インデックス集合を計算
        active_indices = self._compute_swap_active_indices(
            original_lines, outer_s, outer_e, inner_s, inner_e
        )

        if not active_indices:
            messagebox.showinfo("文字入替", "入替対象の範囲が見つかりませんでした。", parent=self)
            self._status_var.set("対象範囲なし")
            return

        # 入替実行（プレースホルダ経由で安全に）
        # 共通ヘルパー _make_char_swap_fn でロジック一元化
        # 厳密マッチ（flex=False）: 前=英数字_でない／後=数字でない＋小数点禁止
        #   → G2 は G20/G2.5 にはマッチせず、G2X10/G2 行末/G2 (コメント) はマッチ
        # 柔軟マッチ（flex=True） : 境界条件なし、部分一致もヒット
        swap_one_line = self._make_char_swap_fn(a, b, flex, skip_comment)

        new_lines = []
        change_indices = []
        line_changes = []

        for idx, line in enumerate(original_lines):
            content = line.rstrip("\r\n")
            ending = line[len(content):]

            if idx in active_indices:
                new_content = swap_one_line(content)
            else:
                new_content = content

            new_line = new_content + ending
            new_lines.append(new_line)
            if new_content != content:
                change_indices.append(idx)
                line_changes.append((base_line + idx, content, new_content))

        if not line_changes:
            messagebox.showinfo("文字入替", "変更対象がありませんでした。", parent=self)
            self._status_var.set("一致なし")
            return

        def on_confirm(selected_set, with_comment=False):
            # change_indices と line_changes の対応マップ
            old_map = {}
            for (line_no, old_c, new_c), orig_idx in zip(line_changes, change_indices):
                old_map[orig_idx] = old_c

            assembled = []
            for idx, orig in enumerate(original_lines):
                if idx in change_indices and idx in selected_set:
                    new_line = new_lines[idx]
                    if with_comment and idx in old_map:
                        content = new_line.rstrip("\r\n")
                        ending = new_line[len(content):]
                        new_line = self._build_commented_line(old_map[idx], content, ending)
                    assembled.append(new_line)
                else:
                    assembled.append(orig)
            new_text = "".join(assembled)
            applied = len(selected_set)
            self._apply_replacement(new_text, start_idx, end_idx, applied, "文字入替")

        self._show_preview_dialog("文字入替 プレビュー", line_changes, change_indices, on_confirm)

    def _compute_swap_active_indices(self, lines, outer_s, outer_e, inner_s, inner_e,
                                       regex_boundary: bool = False,
                                       ignore_case: bool = False):
        """入替対象となる行インデックスの集合を返す。
        - 範囲指定なし: 全行
        - 外側のみ: 外側内（境界除く）の全行
        - 外側＋内側: 外側内にある内側範囲のみ（境界除く）
        終端が見つからない範囲は無視。
        regex_boundary=True で境界マーカーを正規表現として解釈する。
        """
        n = len(lines)

        # 範囲指定なし → 全行アクティブ
        if not outer_s:
            return set(range(n))

        # 外側範囲をスキャン（複数回出現対応）
        # 各行に outer_s / outer_e が含まれるかで判定
        outer_ranges = self._scan_paired_markers(
            lines, outer_s, outer_e, regex_boundary, ignore_case
        )
        if not outer_ranges:
            return set()

        # 内側指定なし → 外側範囲内（境界除く）すべて
        if not inner_s:
            active = set()
            for s, e in outer_ranges:
                # s, e は境界行のインデックス。間の行 (s+1 .. e-1) を採用
                active.update(range(s + 1, e))
            return active

        # 内側指定あり → 外側範囲内で内側範囲をスキャン
        active = set()
        for o_s, o_e in outer_ranges:
            # 外側内部の行のみで内側マーカーを探索
            inner_inside_lines = lines[o_s + 1:o_e]  # 境界除いた中身
            inner_ranges = self._scan_paired_markers(
                inner_inside_lines, inner_s, inner_e, regex_boundary, ignore_case
            )
            for i_s, i_e in inner_ranges:
                # i_s, i_e は inner_inside_lines 内のインデックス
                # 元のlinesでのインデックスに変換: o_s + 1 + i_s
                global_inner_s = o_s + 1 + i_s
                global_inner_e = o_s + 1 + i_e
                active.update(range(global_inner_s + 1, global_inner_e))
        return active

    def _scan_paired_markers(self, lines, start_marker, end_marker,
                              regex_mode: bool = False,
                              ignore_case: bool = False):
        """行のリストから (start_marker を含む行, end_marker を含む行) のペアを抽出。
        - 開始マーカーを見つけたら次の終端マーカーとペアにする
        - 終端が見つからなければそのペアは無視
        - 同じ行に start_marker と end_marker の両方があっても、開始扱いとする
        - regex_mode=True で start_marker/end_marker を正規表現として解釈
        Returns: list of (start_idx, end_idx) tuples
        """
        # マッチ判定関数を構築
        if regex_mode:
            flags = re.IGNORECASE if ignore_case else 0
            try:
                start_re = re.compile(start_marker, flags)
                end_re = re.compile(end_marker, flags)
            except re.error:
                # 正規表現エラーの場合は空リストを返す
                return []
            def match_start(line: str) -> bool:
                return bool(start_re.search(line))
            def match_end(line: str) -> bool:
                return bool(end_re.search(line))
        else:
            if ignore_case:
                start_lower = start_marker.lower()
                end_lower = end_marker.lower()
                def match_start(line: str) -> bool:
                    return start_lower in line.lower()
                def match_end(line: str) -> bool:
                    return end_lower in line.lower()
            else:
                def match_start(line: str) -> bool:
                    return start_marker in line
                def match_end(line: str) -> bool:
                    return end_marker in line

        ranges = []
        i = 0
        n = len(lines)
        while i < n:
            if match_start(lines[i]):
                start_i = i
                # 次の行から終端を探す
                j = i + 1
                found_end = False
                while j < n:
                    if match_end(lines[j]):
                        ranges.append((start_i, j))
                        found_end = True
                        i = j + 1
                        break
                    j += 1
                if not found_end:
                    # 終端が見つからない → このペアは無視、次へ
                    break
            else:
                i += 1
        return ranges

    def _setup_tags(self) -> None:
        """検索ヒット用ハイライトタグを本文側に設定"""
        self.target.tag_configure("find_match", background="#3A2E08", foreground="#FDE9A8")
        # エラー・ジャンプより下、シンタックスより上
        try:
            self.target.tag_raise("find_match")
            self.target.tag_raise("error_line_highlight")
            self.target.tag_raise("jump_highlight")
        except tk.TclError:
            pass

    def _bind_keys(self) -> None:
        self.bind("<Return>", lambda e: self.find_next())
        self.bind("<Escape>", lambda e: self._on_close())

    def _center_on_parent(self) -> None:
        try:
            self.app.root.update_idletasks()
            px = self.app.root.winfo_rootx()
            py = self.app.root.winfo_rooty()
            pw = self.app.root.winfo_width()
            ph = self.app.root.winfo_height()
            w = self.winfo_width()
            h = self.winfo_height()
            x = px + (pw - w) // 2
            y = py + (ph - h) // 3
            self.geometry(f"+{x}+{y}")
        except tk.TclError:
            pass

    def _build_pattern(self):
        """検索文字列を正規表現Patternに変換。失敗時はNoneとエラー文字列を返す"""
        text = self._search_var.get()
        if not text:
            return None, "検索文字列が空です"
        flags = 0 if self._case_var.get() == "1" else re.IGNORECASE
        try:
            if self._regex_var.get() == "1":
                pat = re.compile(text, flags)
            else:
                pat = re.compile(re.escape(text), flags)
            return pat, None
        except re.error as e:
            return None, f"正規表現エラー: {e}"

    def _clear_match_highlights(self) -> None:
        try:
            self.target.tag_remove("find_match", "1.0", tk.END)
        except tk.TclError:
            pass

    def find_next(self) -> bool:
        pat, err = self._build_pattern()
        if pat is None:
            self._status_var.set(f"⚠ {err}")
            return False

        self._clear_match_highlights()
        content = self.target.get("1.0", "end-1c")
        if not content:
            self._status_var.set("本文が空です")
            return False

        # カーソル位置（または前回検索終了位置）から検索開始
        try:
            cursor_idx = self.target.index(tk.INSERT)
        except tk.TclError:
            cursor_idx = "1.0"
        start_offset = self._index_to_offset(cursor_idx)

        # 一旦カーソル以降を探す
        match = pat.search(content, start_offset)
        looped = False
        if not match:
            # 末尾まで来てたら先頭から再検索
            match = pat.search(content)
            looped = True
            if not match:
                self._status_var.set("一致なし")
                return False

        # マッチ範囲をハイライト＋スクロール
        start_idx = self._offset_to_index(match.start())
        end_idx = self._offset_to_index(match.end())
        self.target.tag_add("find_match", start_idx, end_idx)
        self.target.mark_set(tk.INSERT, end_idx)
        self.target.see(start_idx)

        if looped:
            self._status_var.set("先頭に戻って一致")
        else:
            self._status_var.set(f"一致：{match.group()}")
        return True

    def replace_one(self) -> None:
        """次の一致を置換して次の一致まで進む"""
        pat, err = self._build_pattern()
        if pat is None:
            self._status_var.set(f"⚠ {err}")
            return

        # 既にハイライト中の一致があればそれを置換、なければ次を検索してから置換
        ranges = self.target.tag_ranges("find_match")
        if ranges:
            start_idx = str(ranges[0])
            end_idx = str(ranges[1])
            replacement = self._replace_var.get()

            # 正規表現の場合は \1, \2 などの後方参照を反映
            if self._regex_var.get() == "1":
                matched_text = self.target.get(start_idx, end_idx)
                m = pat.fullmatch(matched_text)
                if m:
                    try:
                        replacement = m.expand(replacement)
                    except Exception:
                        pass

            self.target.delete(start_idx, end_idx)
            self.target.insert(start_idx, replacement)
            self._clear_match_highlights()
            # 次へ進める
            self.app._update_linenumbers()
            self.find_next()
            # 連打される可能性が高いのでdebounceで吸収
            self._trigger_post_replace_check_debounced()
        else:
            # 一致が無ければ次を検索するだけ
            self.find_next()

    def _replace_all_whole_text(self) -> None:
        """改行を含む検索/置換用ルート：テキスト全体を一気に処理し、
        MacroPreviewDialog で差分を見せて確認後に適用。
        """
        original = self.target.get("1.0", "end-1c")
        params = {
            "old": self._search_var.get(),
            "new": self._replace_var.get(),
            "use_regex": self._regex_var.get() == "1",
            "case_sensitive": self._case_var.get() == "1",
        }
        if not params["old"]:
            self._status_var.set("⚠ 検索文字列が空")
            return

        new_text, count = self._macro_apply_simple_replace(original, params)
        if count == 0:
            self._status_var.set("一致なし")
            messagebox.showinfo("置換", "一致する文字列が見つかりませんでした。", parent=self)
            return
        if new_text == original:
            self._status_var.set("変更なし")
            messagebox.showinfo("置換", "マッチはあったものの、変更内容がありませんでした。", parent=self)
            return

        # MacroPreviewDialog で差分確認
        step_logs = [("step", f"シンプル置換: {count}箇所マッチ\n")]
        preview = MacroPreviewDialog(
            self,
            original_text=original,
            final_text=new_text,
            step_logs=step_logs,
            total_changes=count,
            executed=1,
            skipped=0,
        )
        preview.transient(self)
        preview.grab_set()
        self.wait_window(preview)

        if not preview.confirmed:
            self._status_var.set("置換キャンセル")
            return

        # 適用（atomic Undo）
        self._atomic_replace(self.target, "1.0", "end", new_text)
        self._clear_match_highlights()
        self.app._update_linenumbers()
        self.app._highlight_all()
        self._status_var.set(f"{count}箇所を置換しました")
        try:
            self._trigger_post_replace_check_debounced()
        except Exception:
            pass
        messagebox.showinfo("置換完了",
                            f"{count}箇所を置換しました。\n（Ctrl+Z で元に戻せます）",
                            parent=self)

    def replace_all(self) -> None:
        pat, err = self._build_pattern()
        if pat is None:
            self._status_var.set(f"⚠ {err}")
            return

        # 範囲指定の取得＆整合性チェック
        range_spec = self._get_range_spec()
        if range_spec is None:
            # 整合性エラー（ステータス設定済み）
            return

        # 検索文字列に改行が含まれている、または正規表現で改行系メタ文字を使ってる場合は
        # 「行単位プレビュー＋チェックボックス」UIが使えないので、テキスト全体置換ルートに
        search_str = self._search_var.get()
        is_regex = self._regex_var.get() == "1"
        has_newline = (
            "\n" in search_str
            or "\n" in self._replace_var.get()
            or (is_regex and any(s in search_str for s in (r"\n", r"\r", r"\s", r".*\n", "(?s)")))
        )
        if has_newline:
            # 範囲指定モードでは改行検索は非対応（行単位の境界判定が前提のため）
            if range_spec:
                self._status_var.set("⚠ 範囲指定モードでは改行を含む検索は使えません")
                messagebox.showwarning(
                    "範囲指定の制約",
                    "範囲指定モードでは改行を含む検索／置換はサポートされていません。\n"
                    "範囲指定を外すか、検索/置換から改行を取り除いてください。",
                    parent=self,
                )
                return
            self._replace_all_whole_text()
            return

        # 対象範囲を取得（セレクタ経由）
        if hasattr(self, "_simple_scope_selector"):
            scope_text, range_start, range_end = self._get_target_text(self._simple_scope_selector)
        else:
            scope_text = self.target.get("1.0", "end-1c")
            range_start, range_end = "1.0", "end-1c"

        if scope_text is None:
            # ステータスはセレクタ側で設定済み
            return

        # 範囲指定があれば、対象テキストを行に分割してアクティブ行集合を計算
        if range_spec:
            scope_lines = scope_text.splitlines(keepends=True)
            ignore_case = self._case_var.get() == "0"
            try:
                active_indices = self._compute_swap_active_indices(
                    scope_lines,
                    range_spec["outer_s"], range_spec["outer_e"],
                    range_spec["inner_s"], range_spec["inner_e"],
                    regex_boundary=range_spec["regex"],
                    ignore_case=ignore_case,
                )
            except re.error as e:
                self._status_var.set(f"⚠ 境界の正規表現エラー: {e}")
                messagebox.showerror("正規表現エラー",
                                      f"境界マーカーの正規表現が不正です。\n{e}",
                                      parent=self)
                return
            if not active_indices:
                self._status_var.set("一致なし（範囲指定の境界が見つかりません）")
                messagebox.showinfo("置換",
                                     "範囲指定の境界が見つかりませんでした。",
                                     parent=self)
                return
        else:
            active_indices = None  # 全行対象を意味する

        matches = list(pat.finditer(scope_text))
        if not matches:
            self._status_var.set("一致なし")
            messagebox.showinfo("置換", "対象範囲に一致する文字列が見つかりませんでした。", parent=self)
            return

        # 行単位の変更情報を組み立てる
        replacement = self._replace_var.get()
        is_regex = self._regex_var.get() == "1"

        # 範囲開始行のオフセットを計算（line_no表示用）
        try:
            base_line = int(str(range_start).split(".")[0])
        except (ValueError, AttributeError):
            base_line = 1

        original_lines = scope_text.splitlines(keepends=True)
        new_lines = []
        change_indices = []
        line_changes = []  # [(line_no, old_content, new_content), ...]

        for idx, line in enumerate(original_lines):
            line_content = line.rstrip("\r\n")
            ending = line[len(line_content):]
            # 範囲指定モードでアクティブ外の行はスキップ
            if active_indices is not None and idx not in active_indices:
                new_lines.append(line)
                continue
            if is_regex:
                new_content = pat.sub(replacement, line_content)
            else:
                new_content = pat.sub(lambda m: replacement, line_content)
            new_line = new_content + ending
            new_lines.append(new_line)
            if new_content != line_content:
                change_indices.append(idx)
                line_changes.append((base_line + idx, line_content, new_content))

        if not line_changes:
            # マッチはあったが行差分なし or 範囲外しか一致しない
            if active_indices is not None:
                self._status_var.set("範囲内に該当なし")
                messagebox.showinfo("置換",
                                     "指定範囲内に一致する文字列が見つかりませんでした。",
                                     parent=self)
            else:
                self._status_var.set("変更対象なし")
                messagebox.showinfo("置換", "変更対象がありませんでした。", parent=self)
            return

        match_count = len(matches)
        affected_line_count = len(line_changes)

        def on_confirm(selected_set, with_comment=False):
            # change_indices と line_changes の対応マップ
            old_map = {}
            for (line_no, old_c, new_c), orig_idx in zip(line_changes, change_indices):
                old_map[orig_idx] = old_c

            assembled = []
            applied_lines = 0
            for idx, orig in enumerate(original_lines):
                if idx in change_indices and idx in selected_set:
                    new_line = new_lines[idx]
                    if with_comment and idx in old_map:
                        content = new_line.rstrip("\r\n")
                        ending = new_line[len(content):]
                        new_line = self._build_commented_line(old_map[idx], content, ending)
                    assembled.append(new_line)
                    applied_lines += 1
                else:
                    assembled.append(orig)
            new_text = "".join(assembled)

            # 範囲だけ差し替え（undo履歴1回分にする）
            prev_autosep = self.target.cget("autoseparators")
            self.target.configure(autoseparators=False)
            try:
                self.target.edit_separator()
                self.target.delete(range_start, range_end)
                self.target.insert(range_start, new_text)
                self.target.edit_separator()
            finally:
                self.target.configure(autoseparators=prev_autosep)

            self._clear_match_highlights()
            self.app._update_linenumbers()
            self.app._highlight_all()
            self._status_var.set(f"{applied_lines}行を置換しました")
            messagebox.showinfo(
                "置換完了",
                f"{applied_lines}行を置換しました。\n（Ctrl+Z で元に戻せます）",
                parent=self,
            )
            # 置換後に自動でプログラムチェック実行
            self._trigger_post_replace_check()

        # スコープ名をタイトルに含める
        scope_label = "全体"
        if hasattr(self, "_simple_scope_selector"):
            sel = self._simple_scope_selector
            scope_label = sel.get()
            if scope_label == "指定したNブロック":
                n = sel.get_n_block_label()
                if n:
                    scope_label = f"{n}ブロック"

        title = f"全置換 プレビュー [{scope_label}]（一致 {match_count}件 / 影響 {affected_line_count}行）"
        self._show_preview_dialog(title, line_changes, change_indices, on_confirm)

    def _index_to_offset(self, index: str) -> int:
        """tk index ('row.col') を 文字列オフセット に変換"""
        try:
            content_to_idx = self.target.get("1.0", index)
            return len(content_to_idx)
        except tk.TclError:
            return 0

    def _offset_to_index(self, offset: int) -> str:
        return self.target.index(f"1.0 + {offset} chars")

    def _on_close(self) -> None:
        self._clear_match_highlights()
        # マウスホイールのbind_allを明示的に解除（残留防止）
        try:
            if hasattr(self, "_tpl_list_canvas") and self._tpl_list_canvas:
                self._tpl_list_canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        # ダイアログ参照をクリア（再オープン用）
        if hasattr(self.app, "_find_dialog"):
            self.app._find_dialog = None
        self.destroy()


class MacroStepDialog(tk.Toplevel):
    """マクロにステップを追加（または編集）するダイアログ。
    左ペイン：テンプレ一覧、右ペイン：選択中テンプレのパラメータ入力UI
    """

    # 既存テンプレ定義（FindReplaceDialog._build_template_tab と一致させる）
    TEMPLATES = [
        ("simple_replace", "シンプル置換",      "基本",       "文字列・正規表現で検索→置換"),
        ("f_scale",        "F値倍率変更",       "数値変更",   "全F値に倍率を掛ける"),
        ("f_uniform",      "F値統一",           "数値変更",   "全F値を指定値に置換"),
        ("s_scale",        "S値倍率変更",       "数値変更",   "全S値に倍率を掛ける"),
        ("s_uniform",      "S値統一",           "数値変更",   "全S値を指定値に置換"),
        ("axis_offset",    "座標オフセット加算", "座標操作",   "指定軸の値に加算"),
        ("axis_invert",    "座標符号反転",      "座標操作",   "指定軸の値の符号を反転"),
        ("decimal_add",    "小数点付加",        "形式整形",   "整数指令に「.」を追加"),
        ("upper_case",     "小文字→大文字統一", "形式整形",   "NCコマンドを大文字に統一"),
        ("trim_spaces",    "余分な空白除去",    "形式整形",   "行頭末尾の空白を除去"),
        ("comment_add",    "コメント化",        "コメント",   "対象行を ( ) で囲む"),
        ("comment_del",    "コメント解除",      "コメント",   "( ) を外して通常行に戻す"),
        ("comment_remove", "コメント削除",      "コメント",   "行内の ( ) コメントを削除"),
        ("char_swap",      "文字入替",          "入替",       "A↔B の入替（範囲指定対応）"),
        ("line_delete",    "行削除",            "行操作",     "空行・パターン一致行を削除"),
    ]

    def __init__(self, parent_dialog, edit_index: int | None = None):
        """parent_dialog: FindReplaceDialog インスタンス
        edit_index: 編集モードの場合、対象ステップのインデックス
        """
        super().__init__(parent_dialog)
        self.parent_dialog = parent_dialog
        self.edit_index = edit_index
        self.result_step: "MacroStep | None" = None  # OK 押下時に設定

        # 編集モードの場合は元データ取得
        self._initial_step: "MacroStep | None" = None
        if edit_index is not None:
            steps = parent_dialog._current_macro.steps
            if 0 <= edit_index < len(steps):
                self._initial_step = steps[edit_index]

        title = "ステップ編集" if edit_index is not None else "ステップ追加"
        self.title(f"マクロ — {title}")
        self.configure(bg=BG_PANEL)
        self.geometry("680x560")
        self.minsize(640, 480)
        self.resizable(True, True)

        # 選択中テンプレ
        self._selected_template_id: str | None = None
        # 入力欄保持用
        self._entries: dict = {}

        self._build_ui()

        # 編集モードの場合：初期テンプレを選択して値をセット
        if self._initial_step is not None:
            self._select_template(self._initial_step.template_id)
            self._populate_from_step(self._initial_step)

        # 中央寄せ
        self.update_idletasks()
        self._center_on_parent()

        # ESC で閉じる
        self.bind("<Escape>", lambda e: self._on_cancel())
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_on_parent(self) -> None:
        try:
            self.update_idletasks()
            pw = self.parent_dialog.winfo_width()
            ph = self.parent_dialog.winfo_height()
            px = self.parent_dialog.winfo_rootx()
            py = self.parent_dialog.winfo_rooty()
            w = self.winfo_width()
            h = self.winfo_height()
            x = px + (pw - w) // 2
            y = py + (ph - h) // 2
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        outer = tk.Frame(self, bg=BG_PANEL, padx=12, pady=10)
        outer.pack(fill="both", expand=True)

        # ヘッダ
        title_text = "▶ ステップ編集" if self.edit_index is not None else "▶ ステップ追加"
        tk.Label(outer, text=title_text,
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w")
        tk.Label(outer, text="左からテンプレートを選び、パラメータを入力",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(anchor="w", pady=(0, 8))

        # 本体（左右分割）
        body = tk.Frame(outer, bg=BG_PANEL)
        body.pack(fill="both", expand=True)

        # 左：テンプレ一覧
        left = tk.Frame(body, bg=BG_PANEL, width=200)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)

        tk.Label(left, text="テンプレート",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        # スクロール対応
        list_wrap = tk.Frame(left, bg=BG_PANEL)
        list_wrap.pack(fill="both", expand=True)

        list_canvas = tk.Canvas(list_wrap, bg=BG_PANEL, bd=0,
                                highlightthickness=0, takefocus=False)
        list_scroll = tk.Scrollbar(list_wrap, orient="vertical",
                                    command=list_canvas.yview,
                                    bg=BG_PANEL, troughcolor=BG_HEADER,
                                    activebackground=ACCENT_DARK,
                                    bd=0, highlightthickness=0, width=10)
        list_canvas.configure(yscrollcommand=list_scroll.set)
        list_scroll.pack(side="right", fill="y")
        list_canvas.pack(side="left", fill="both", expand=True)

        list_inner = tk.Frame(list_canvas, bg=BG_PANEL)
        list_inner_id = list_canvas.create_window((0, 0), window=list_inner, anchor="nw")

        def _on_inner_configure(event):
            list_canvas.configure(scrollregion=list_canvas.bbox("all"))
        list_inner.bind("<Configure>", _on_inner_configure)

        def _on_canvas_configure(event):
            list_canvas.itemconfigure(list_inner_id, width=event.width)
        list_canvas.bind("<Configure>", _on_canvas_configure)

        # マウスホイール
        def _on_mousewheel(event):
            list_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _bind_wheel(_e):
            list_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_wheel(_e):
            list_canvas.unbind_all("<MouseWheel>")
        list_canvas.bind("<Enter>", _bind_wheel)
        list_canvas.bind("<Leave>", _unbind_wheel)

        # カテゴリ別表示
        self._tpl_buttons: dict = {}
        last_category = None
        for tpl_id, label, category, _desc in self.TEMPLATES:
            if category != last_category:
                tk.Label(list_inner, text=f"▼ {category}",
                         bg=BG_PANEL, fg=ACCENT_DARK,
                         font=("Yu Gothic UI", 9, "bold")).pack(anchor="w", pady=(6, 2), padx=4)
                last_category = category
            btn = tk.Button(
                list_inner, text=f"  {label}",
                command=lambda tid=tpl_id: self._select_template(tid),
                font=("Yu Gothic UI", 9), bg=BG_PANEL, fg=TEXT_MAIN,
                activebackground="#123A38", activeforeground=ACCENT,
                relief="flat", bd=0, anchor="w", padx=8, pady=2, cursor="hand2",
            )
            btn.pack(fill="x", pady=1)
            self._tpl_buttons[tpl_id] = btn

        # 右：パラメータ入力
        self._right_frame = tk.Frame(body, bg=BG_PANEL)
        self._right_frame.pack(side="left", fill="both", expand=True)
        self._show_placeholder()

        # 下部ボタン群
        btn_row = tk.Frame(outer, bg=BG_PANEL)
        btn_row.pack(fill="x", pady=(10, 0))

        btn_kw = {"font": ("Yu Gothic UI", 10, "bold"),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "solid", "bd": 1, "padx": 14, "pady": 4,
                  "cursor": "hand2"}

        ok_text = "更新" if self.edit_index is not None else "追加"
        tk.Button(btn_row, text=ok_text, command=self._on_ok,
                  font=("Yu Gothic UI", 10, "bold"),
                  bg="#123A38", fg=ACCENT,
                  activebackground="#175E56", activeforeground=ACCENT,
                  relief="solid", bd=1, padx=14, pady=4,
                  cursor="hand2").pack(side="right", padx=(4, 0))
        tk.Button(btn_row, text="キャンセル", command=self._on_cancel,
                  **btn_kw).pack(side="right")

        # ステータス
        self._status_var = tk.StringVar(value="")
        tk.Label(outer, textvariable=self._status_var,
                 bg=BG_PANEL, fg="#FB5E7E",
                 font=("Consolas", 9), anchor="w").pack(fill="x", pady=(6, 0))

    def _show_placeholder(self) -> None:
        for w in self._right_frame.winfo_children():
            w.destroy()
        tk.Label(self._right_frame,
                 text="◀ 左からテンプレートを選択してください",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 10)).pack(anchor="w", pady=20, padx=4)

    def _select_template(self, tpl_id: str) -> None:
        """テンプレ選択：右ペインに該当テンプレの設定UIを表示"""
        # ボタン色更新
        for tid, btn in self._tpl_buttons.items():
            if tid == tpl_id:
                btn.configure(bg="#123A38", fg=ACCENT)
            else:
                btn.configure(bg=BG_PANEL, fg=TEXT_MAIN)

        # 詳細クリア
        for w in self._right_frame.winfo_children():
            w.destroy()
        self._entries = {}
        self._selected_template_id = tpl_id

        # テンプレ情報
        tpl_info = next((t for t in self.TEMPLATES if t[0] == tpl_id), None)
        if tpl_info is None:
            return
        _, label, category, desc = tpl_info

        # ヘッダ
        tk.Label(self._right_frame, text=f"▶ {label}",
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(self._right_frame, text=f"[{category}]　{desc}",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9), wraplength=380, justify="left").pack(anchor="w", pady=(0, 12))

        # テンプレ別UIビルダー呼び出し
        builder = getattr(self, f"_ui_{tpl_id}", None)
        if builder:
            builder(self._right_frame)
        else:
            tk.Label(self._right_frame,
                     text="（このテンプレートはマクロでは使用できません）",
                     bg=BG_PANEL, fg=TEXT_MUTED,
                     font=("Yu Gothic UI", 10)).pack(anchor="w")

    # ---------- 入力UIヘルパー ----------

    def _make_entry(self, parent, label_text: str, default: str = "", key: str = "") -> tk.Entry:
        row = tk.Frame(parent, bg=BG_PANEL)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label_text, bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="e").pack(side="left", padx=(0, 8))
        var = tk.StringVar(value=default)
        e = tk.Entry(row, textvariable=var, font=("Consolas", 11),
                     bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT,
                     relief="solid", bd=1,
                     highlightcolor=ACCENT, highlightbackground=BORDER,
                     highlightthickness=1)
        e.pack(side="left", fill="x", expand=True)
        e._var = var
        if key:
            self._entries[key] = e
        return e

    def _make_combobox(self, parent, label_text: str, options: list, default: str = "", key: str = ""):
        row = tk.Frame(parent, bg=BG_PANEL)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label_text, bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="e").pack(side="left", padx=(0, 8))
        var = tk.StringVar(value=default)
        from tkinter import ttk
        cb = ttk.Combobox(row, textvariable=var, values=options,
                          state="readonly", font=("Consolas", 10), width=14)
        cb.pack(side="left")
        if key:
            self._entries[key] = var
        return var

    def _make_checkbox(self, parent, label_text: str, default: bool = False, key: str = ""):
        row = tk.Frame(parent, bg=BG_PANEL)
        row.pack(fill="x", pady=4)
        tk.Label(row, text="", bg=BG_PANEL, width=18).pack(side="left", padx=(0, 8))
        var = tk.StringVar(value="1" if default else "0")
        cb = tk.Checkbutton(row, text=label_text, variable=var,
                            onvalue="1", offvalue="0",
                            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
                            activebackground=BG_PANEL, activeforeground=ACCENT,
                            font=("Yu Gothic UI", 9), bd=0, highlightthickness=0)
        cb.pack(side="left")
        if key:
            self._entries[key] = var
        return var

    def _make_note(self, parent, text: str) -> None:
        tk.Label(parent, text=text, bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9), wraplength=420,
                 justify="left").pack(anchor="w", pady=(8, 0))

    def _make_textarea(self, parent, label_text: str, default: str = "",
                        key: str = "", height: int = 3) -> tk.Text:
        """複数行入力用 Text ウィジェット（改行入力可）。"""
        row = tk.Frame(parent, bg=BG_PANEL)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label_text, bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10), width=18, anchor="ne").pack(side="left", padx=(0, 8), pady=(2, 0))
        wrap = tk.Frame(row, bg=BG_PANEL,
                         highlightbackground=BORDER, highlightthickness=1)
        wrap.pack(side="left", fill="x", expand=True)
        t = tk.Text(wrap, height=height, font=("Consolas", 11),
                    bg=INPUT_BG, fg=ACCENT, insertbackground=ACCENT,
                    relief="flat", bd=0, wrap="none",
                    padx=4, pady=2, undo=True)
        t.pack(side="left", fill="both", expand=True)
        if default:
            t.insert("1.0", default)
        if key:
            self._entries[key] = t
        # 値取得用にマーカー（_make_entry の _var と区別）
        t._is_textarea = True
        return t

    def _open_regex_builder_for_field(self, key: str, use_regex: bool = True) -> None:
        """指定フィールドに対して正規表現ビルダーを開く。
        key: self._entries の key
        use_regex: True なら use_regex チェックも自動でONにする
        FindReplaceDialog 側の _show_regex_builder と似た処理だが、
        本ダイアログは Text/Entry/StringVar 3パターンに対応する必要があるため独自実装。
        """
        entry = self._entries.get(key)
        if entry is None:
            return

        # 現在値を取得して初期パターンに
        if getattr(entry, "_is_textarea", False):
            try:
                initial_pattern = entry.get("1.0", "end-1c")
            except tk.TclError:
                initial_pattern = ""
        elif hasattr(entry, "_var"):
            initial_pattern = entry._var.get()
        else:
            initial_pattern = ""

        # メイン画面の選択範囲を取得（FindReplaceDialog → app → input_text 経由）
        initial_test = ""
        try:
            target = self.parent_dialog.target
            sel = target.get("sel.first", "sel.last")
            if sel:
                initial_test = sel
        except (tk.TclError, AttributeError):
            pass

        # ignore_case の初期値
        ignore_case = False
        cs_var = self._entries.get("case_sensitive")
        if isinstance(cs_var, tk.StringVar):
            ignore_case = (cs_var.get() == "0")

        builder = RegexBuilderDialog(self,
                                      initial_pattern=initial_pattern,
                                      initial_test_text=initial_test,
                                      initial_ignore_case=ignore_case)
        builder.transient(self)
        builder.grab_set()
        self.wait_window(builder)

        if builder.result_pattern is not None:
            # 結果をフィールドに流し込む
            if getattr(entry, "_is_textarea", False):
                try:
                    entry.delete("1.0", "end")
                    entry.insert("1.0", builder.result_pattern)
                except tk.TclError:
                    pass
            elif hasattr(entry, "_var"):
                entry._var.set(builder.result_pattern)
            # use_regex チェックを自動ON
            if use_regex:
                ur_var = self._entries.get("use_regex")
                if isinstance(ur_var, tk.StringVar):
                    ur_var.set("1")
            self._status_var.set("正規表現ビルダー: パターンを設定しました")

    # ---------- 各テンプレ用UI ----------

    def _ui_simple_replace(self, parent):
        self._make_textarea(parent, "検索文字:", "", key="old", height=3)
        self._make_textarea(parent, "置換後:", "", key="new", height=3)
        self._make_checkbox(parent, "正規表現として扱う", False, key="use_regex")
        self._make_checkbox(parent, "大文字小文字を区別する", True, key="case_sensitive")

        # 正規表現ビルダー呼出ボタン
        builder_row = tk.Frame(parent, bg=BG_PANEL)
        builder_row.pack(fill="x", pady=(4, 0))
        tk.Label(builder_row, text="", bg=BG_PANEL, width=18).pack(side="left", padx=(0, 8))
        tk.Button(
            builder_row, text="🔧 正規表現ビルダーを開く",
            command=lambda: self._open_regex_builder_for_field("old", use_regex=True),
            font=("Yu Gothic UI", 9),
            bg="#141B24", fg=ACCENT,
            activebackground="#1F2B3A", activeforeground=ACCENT,
            relief="solid", bd=1, padx=10, pady=2, cursor="hand2",
        ).pack(side="left")

        self._make_note(parent,
                        "※ 改行を含む検索/置換が可能（Enterキーで改行入力）\n"
                        "※ 検索文字に一致する箇所をすべて置換します\n"
                        "※ 正規表現ON: Pythonの re モジュール構文（\\d, \\s, [], (), 等）\n"
                        "※ サクラエディタの S_ReplaceAll マクロを移植する場合は、\n"
                        "  '\\r\\n' は実際の改行（Enter）に置き換えてください")

    def _ui_f_scale(self, parent):
        self._make_entry(parent, "倍率（例: 0.8）:", "0.8", key="factor")

    def _ui_f_uniform(self, parent):
        self._make_entry(parent, "新しいF値:", "0.15", key="value")

    def _ui_s_scale(self, parent):
        self._make_entry(parent, "倍率（例: 0.8）:", "0.8", key="factor")

    def _ui_s_uniform(self, parent):
        self._make_entry(parent, "新しいS値:", "2000", key="value")

    def _ui_axis_offset(self, parent):
        self._make_combobox(parent, "対象軸:", ["X", "Y", "Z"], "X", key="axis")
        self._make_entry(parent, "加算値:", "5.0", key="offset")

    def _ui_axis_invert(self, parent):
        self._make_combobox(parent, "対象軸:", ["X", "Y", "Z"], "Z", key="axis")

    def _ui_decimal_add(self, parent):
        DECIMAL_ADD_AXES = ["全軸", "X", "Y", "Z", "U", "V", "W", "A", "B", "C", "I", "J", "K", "F", "S"]
        DECIMAL_ADD_SCALES = ["1.", "0.1", "0.01", "0.001", "0.0001"]
        self._make_combobox(parent, "対象軸:", DECIMAL_ADD_AXES, "全軸", key="axis")
        self._make_combobox(parent, "スケール:", DECIMAL_ADD_SCALES, "1.", key="scale")
        self._make_note(parent,
                        "※ スケール 1. = 末尾に「.」付加（例: F100 → F100.）\n"
                        "※ スケール 0.01 = 100で割って小数2桁表示\n"
                        "※ 既に小数点を含む値はスキップします")

    def _ui_upper_case(self, parent):
        self._make_note(parent, "※ コメント (...) 内は変更しません\n※ パラメータなし")

    def _ui_trim_spaces(self, parent):
        self._make_note(parent, "※ 各行の行頭・行末の空白を除去\n※ パラメータなし")

    def _ui_comment_add(self, parent):
        self._make_note(parent, "※ 各行を ( ) で囲みます\n※ パラメータなし")

    def _ui_comment_del(self, parent):
        self._make_note(parent, "※ 行全体が ( ) で囲まれた行のカッコを外します\n※ パラメータなし")

    def _ui_comment_remove(self, parent):
        self._make_note(parent, "※ 行内の ( ) コメントをすべて削除\n※ パラメータなし")

    def _ui_char_swap(self, parent):
        self._make_entry(parent, "文字A:", "", key="a")
        self._make_entry(parent, "文字B:", "", key="b")
        self._make_checkbox(parent, "柔軟マッチ（境界を無視して部分一致もヒット）",
                            False, key="flex")
        self._make_checkbox(parent, "コメント内 ( ... ) は置換しない（推奨ON）",
                            True, key="skip_comment")

        # 区切り
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=(8, 4))
        tk.Label(parent, text="範囲指定（任意）",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", padx=4)

        self._make_entry(parent, "外側 開始:", "", key="outer_start")
        self._make_entry(parent, "外側 終端:", "", key="outer_end")
        self._make_entry(parent, "内側 開始:", "", key="inner_start")
        self._make_entry(parent, "内側 終端:", "", key="inner_end")

        self._make_note(parent,
                        "※ 範囲指定なし: 全体で A↔B を入替\n"
                        "※ 外側のみ指定: 外側範囲内のみ入替（境界行は除外）\n"
                        "※ 外側＋内側: 外側内に内側がある場合のみ、内側範囲で入替")

    def _ui_line_delete(self, parent):
        self._make_checkbox(parent, "空行を削除する", True, key="delete_empty")

        # 区切り
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=(8, 4))

        self._make_checkbox(parent, "パターンに一致する行を削除する", False, key="delete_pattern")
        self._make_entry(parent, "パターン:", "", key="pattern")
        self._make_checkbox(parent, "正規表現として扱う", False, key="use_regex")
        self._make_checkbox(parent, "大文字小文字を区別する", True, key="case_sensitive")

        # 正規表現ビルダー呼出ボタン
        builder_row = tk.Frame(parent, bg=BG_PANEL)
        builder_row.pack(fill="x", pady=(4, 0))
        tk.Label(builder_row, text="", bg=BG_PANEL, width=18).pack(side="left", padx=(0, 8))
        tk.Button(
            builder_row, text="🔧 正規表現ビルダーを開く",
            command=lambda: self._open_regex_builder_for_field("pattern", use_regex=True),
            font=("Yu Gothic UI", 9),
            bg="#141B24", fg=ACCENT,
            activebackground="#1F2B3A", activeforeground=ACCENT,
            relief="solid", bd=1, padx=10, pady=2, cursor="hand2",
        ).pack(side="left")

        self._make_note(parent,
                        "※ 両方ON可（空行＋パターン一致を一回で削除）\n"
                        "※ パターン一致は「行に部分一致」した行を削除\n"
                        "  例: 'G04 P' → G04 P を含む行を全削除\n"
                        "※ 行を消すので、後続ステップは行番号がズレます\n"
                        "  （char_swap の範囲指定マーカーには注意）")

    # ---------- 編集モード時の値復元 ----------

    def _populate_from_step(self, step: "MacroStep") -> None:
        """編集モードで初期値を入力欄に流し込む"""
        for key, value in step.params.items():
            entry = self._entries.get(key)
            if entry is None:
                continue
            if isinstance(entry, tk.StringVar):
                # combobox/checkbox
                if isinstance(value, bool):
                    entry.set("1" if value else "0")
                else:
                    entry.set(str(value))
            elif getattr(entry, "_is_textarea", False):
                # Text ウィジェット（複数行）
                try:
                    entry.delete("1.0", "end")
                    entry.insert("1.0", str(value))
                except tk.TclError:
                    pass
            else:
                # tk.Entry（_var 経由）
                if hasattr(entry, "_var"):
                    entry._var.set(str(value))

    # ---------- 入力収集＆検証 ----------

    def _collect_params(self) -> dict | None:
        """入力欄からparams dict を作成。バリデーション失敗時は None を返す"""
        tid = self._selected_template_id
        if tid is None:
            self._status_var.set("⚠ テンプレートを選択してください")
            return None

        params = {}

        def get_str(key: str, strip: bool = True) -> str:
            """key の入力値を取得。strip=False なら改行や前後空白を保持。"""
            e = self._entries.get(key)
            if e is None:
                return ""
            if isinstance(e, tk.StringVar):
                v = e.get()
                return v.strip() if strip else v
            # Text ウィジェット（_make_textarea 由来）
            if getattr(e, "_is_textarea", False):
                try:
                    v = e.get("1.0", "end-1c")
                except tk.TclError:
                    return ""
                return v.strip() if strip else v
            # Entry (_make_entry 由来は _var 持ち)
            if hasattr(e, "_var"):
                v = e._var.get()
                return v.strip() if strip else v
            return ""

        def get_bool(key: str) -> bool:
            e = self._entries.get(key)
            if isinstance(e, tk.StringVar):
                return e.get() == "1"
            return False

        # テンプレ別パラメータ収集＆検証
        if tid == "simple_replace":
            # 改行を保持したいので strip=False
            old = get_str("old", strip=False)
            if not old:
                self._status_var.set("⚠ 検索文字を入力してください")
                return None
            new = get_str("new", strip=False)  # 空でもOK（削除動作）
            use_regex = get_bool("use_regex")
            case_sensitive = get_bool("case_sensitive")

            # 正規表現の妥当性チェック
            if use_regex:
                try:
                    flags = 0 if case_sensitive else re.IGNORECASE
                    re.compile(old, flags)
                except re.error as e:
                    self._status_var.set(f"⚠ 正規表現エラー: {e}")
                    return None

            params["old"] = old
            params["new"] = new
            params["use_regex"] = use_regex
            params["case_sensitive"] = case_sensitive

        elif tid in ("f_scale", "s_scale"):
            v = get_str("factor")
            try:
                params["factor"] = float(v)
            except ValueError:
                self._status_var.set("⚠ 倍率は数値で入力してください")
                return None
            if params["factor"] <= 0:
                self._status_var.set("⚠ 倍率は正の数で入力してください")
                return None

        elif tid in ("f_uniform", "s_uniform"):
            v = get_str("value")
            if not v:
                self._status_var.set("⚠ 値を入力してください")
                return None
            # 数値として妥当か軽くチェック
            try:
                float(v)
            except ValueError:
                self._status_var.set("⚠ 値は数値で入力してください")
                return None
            params["value"] = v

        elif tid == "axis_offset":
            params["axis"] = get_str("axis") or "X"
            v = get_str("offset")
            try:
                params["offset"] = float(v)
            except ValueError:
                self._status_var.set("⚠ 加算値は数値で入力してください")
                return None

        elif tid == "axis_invert":
            params["axis"] = get_str("axis") or "Z"

        elif tid == "decimal_add":
            params["axis"] = get_str("axis") or "全軸"
            params["scale"] = get_str("scale") or "1."

        elif tid in ("upper_case", "trim_spaces", "comment_add", "comment_del", "comment_remove"):
            pass  # パラメータなし

        elif tid == "char_swap":
            a = get_str("a")
            b = get_str("b")
            if not a or not b:
                self._status_var.set("⚠ 文字A・文字Bは両方必須")
                return None
            if a == b:
                self._status_var.set("⚠ 文字Aと文字Bが同じです")
                return None
            params["a"] = a
            params["b"] = b
            params["flex"] = get_bool("flex")
            params["skip_comment"] = get_bool("skip_comment")

            outer_s = get_str("outer_start")
            outer_e = get_str("outer_end")
            inner_s = get_str("inner_start")
            inner_e = get_str("inner_end")

            # 整合性チェック
            if (outer_s and not outer_e) or (outer_e and not outer_s):
                self._status_var.set("⚠ 外側範囲は開始と終端を両方指定してください")
                return None
            if (inner_s and not inner_e) or (inner_e and not inner_s):
                self._status_var.set("⚠ 内側範囲は開始と終端を両方指定してください")
                return None
            if (inner_s or inner_e) and not (outer_s and outer_e):
                self._status_var.set("⚠ 内側範囲は外側範囲とセットで指定")
                return None

            params["outer_start"] = outer_s
            params["outer_end"] = outer_e
            params["inner_start"] = inner_s
            params["inner_end"] = inner_e

        elif tid == "line_delete":
            delete_empty = get_bool("delete_empty")
            delete_pattern = get_bool("delete_pattern")
            pattern = get_str("pattern")
            use_regex = get_bool("use_regex")
            case_sensitive = get_bool("case_sensitive")

            # 何も削除条件が指定されていない
            if not delete_empty and not delete_pattern:
                self._status_var.set("⚠ 「空行削除」または「パターン削除」のどちらかをONにしてください")
                return None

            # パターン削除ONなのにパターン未入力
            if delete_pattern and not pattern:
                self._status_var.set("⚠ パターン削除ONですが、パターンが未入力です")
                return None

            # 正規表現の妥当性チェック
            if delete_pattern and use_regex:
                try:
                    flags = 0 if case_sensitive else re.IGNORECASE
                    re.compile(pattern, flags)
                except re.error as e:
                    self._status_var.set(f"⚠ 正規表現エラー: {e}")
                    return None

            params["delete_empty"] = delete_empty
            params["delete_pattern"] = delete_pattern
            params["pattern"] = pattern
            params["use_regex"] = use_regex
            params["case_sensitive"] = case_sensitive

        else:
            self._status_var.set("⚠ 未対応のテンプレート")
            return None

        return params

    def _on_ok(self) -> None:
        params = self._collect_params()
        if params is None:
            return

        tid = self._selected_template_id
        tpl_info = next((t for t in self.TEMPLATES if t[0] == tid), None)
        label = tpl_info[1] if tpl_info else tid

        self.result_step = MacroStep(
            template_id=tid,
            template_label=label,
            params=params,
            enabled=(self._initial_step.enabled if self._initial_step else True),
        )
        self.destroy()

    def _on_cancel(self) -> None:
        self.result_step = None
        self.destroy()


class MacroPreviewDialog(tk.Toplevel):
    """マクロのプレビュー（差分一覧）ダイアログ。
    実行前/実行後のテキストと、変更行のみの対比一覧を表示。
    OKで実行確定、キャンセルで破棄。
    """

    def __init__(self, parent_dialog, original_text: str, final_text: str,
                 step_logs: list, total_changes: int, executed: int, skipped: int):
        """parent_dialog: FindReplaceDialog インスタンス
        original_text: 実行前テキスト
        final_text: 全ステップ実行後のテキスト
        step_logs: [(tag, message), ...] のリスト
        total_changes: 総変更行数
        executed: 実行ステップ数
        skipped: スキップステップ数
        """
        super().__init__(parent_dialog)
        self.parent_dialog = parent_dialog
        self.confirmed: bool = False  # OK 押下時 True

        self.original_text = original_text
        self.final_text = final_text
        self.step_logs = step_logs
        self.total_changes = total_changes
        self.executed = executed
        self.skipped = skipped

        self.title("マクロ プレビュー — 変更内容を確認")
        self.configure(bg=BG_PANEL)
        self.geometry("1100x680")
        self.minsize(900, 520)
        self.resizable(True, True)

        self._build_ui()

        # 中央寄せ
        self.update_idletasks()
        self._center_on_parent()

        # ESC でキャンセル
        self.bind("<Escape>", lambda e: self._on_cancel())
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_on_parent(self) -> None:
        try:
            self.update_idletasks()
            pw = self.parent_dialog.winfo_width()
            ph = self.parent_dialog.winfo_height()
            px = self.parent_dialog.winfo_rootx()
            py = self.parent_dialog.winfo_rooty()
            w = self.winfo_width()
            h = self.winfo_height()
            x = px + (pw - w) // 2
            y = py + (ph - h) // 2
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        outer = tk.Frame(self, bg=BG_PANEL, padx=12, pady=10)
        outer.pack(fill="both", expand=True)

        # ヘッダ
        tk.Label(outer, text="▶ マクロ実行プレビュー",
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 13, "bold")).pack(anchor="w")

        # サマリ
        summary_text = (
            f"実行ステップ: {self.executed}　"
            f"スキップ: {self.skipped}　"
            f"変更行数: {self.total_changes}"
        )
        tk.Label(outer, text=summary_text,
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 10)).pack(anchor="w", pady=(2, 8))

        # 実行ログ（コンパクト表示）
        log_wrap = tk.Frame(outer, bg=BG_PANEL,
                             highlightbackground=BORDER, highlightthickness=1)
        log_wrap.pack(fill="x", pady=(0, 8))

        log_text = tk.Text(log_wrap, height=4, wrap="word",
                            font=("Consolas", 9),
                            bg=INPUT_BG, fg=TEXT_MAIN,
                            insertbackground=ACCENT,
                            relief="flat", bd=0, padx=6, pady=4)
        log_text.pack(side="left", fill="both", expand=True)

        log_scroll = tk.Scrollbar(log_wrap, orient="vertical",
                                   command=log_text.yview,
                                   bg=BG_PANEL, troughcolor=BG_HEADER,
                                   activebackground=ACCENT_DARK,
                                   bd=0, highlightthickness=0, width=10)
        log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")

        # ログタグ
        log_text.tag_configure("ok", foreground=ACCENT)
        log_text.tag_configure("warn", foreground="#F0D060")
        log_text.tag_configure("err", foreground="#FB5E7E")
        log_text.tag_configure("info", foreground=TEXT_MUTED)
        log_text.tag_configure("step", foreground="#3DDDFF")

        # ログ書き込み
        for tag, msg in self.step_logs:
            log_text.insert("end", msg, tag)
        log_text.configure(state="disabled")

        # 区切り
        tk.Label(outer, text="─── 変更一覧 ─────────────────",
                 bg=BG_PANEL, fg=ACCENT_DARK,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", pady=(2, 4))

        # 変更行を集計
        diff_rows = self._compute_diff_rows()

        if not diff_rows:
            tk.Label(outer, text="（変更行はありません）",
                     bg=BG_PANEL, fg=TEXT_MUTED,
                     font=("Yu Gothic UI", 10)).pack(anchor="w", pady=10)
        else:
            self._build_diff_view(outer, diff_rows)

        # 下部ボタン群
        btn_row = tk.Frame(outer, bg=BG_PANEL)
        btn_row.pack(fill="x", pady=(10, 0))

        btn_kw = {"font": ("Yu Gothic UI", 10, "bold"),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "solid", "bd": 1, "padx": 14, "pady": 4,
                  "cursor": "hand2"}

        # 変更がない場合は実行ボタン無効化
        ok_state = "normal" if diff_rows else "disabled"
        tk.Button(btn_row, text="▶ このまま実行",
                  command=self._on_ok, state=ok_state,
                  font=("Yu Gothic UI", 10, "bold"),
                  bg="#123A38", fg=ACCENT,
                  activebackground="#175E56", activeforeground=ACCENT,
                  relief="solid", bd=1, padx=14, pady=4,
                  cursor="hand2").pack(side="right", padx=(4, 0))
        tk.Button(btn_row, text="キャンセル", command=self._on_cancel,
                  **btn_kw).pack(side="right")

        tk.Label(btn_row,
                 text="※ [このまま実行] でエディタへ反映します（[元に戻す] で取り消し可能）",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(side="left")

    def _compute_diff_rows(self) -> list:
        """元テキストと最終テキストを行単位で比較し、差分のみリストアップ。
        difflib.SequenceMatcher を使い、行削除・追加・変更すべてを正しく扱う。
        Returns: list of (kind, old_line_no, new_line_no, old_content, new_content)
            kind: "modify" | "delete" | "insert"
            old_line_no, new_line_no: 該当しない側は None
        """
        import difflib

        old_lines = self.original_text.splitlines()
        new_lines = self.final_text.splitlines()

        sm = difflib.SequenceMatcher(a=old_lines, b=new_lines, autojunk=False)
        diff_rows = []

        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                continue
            elif tag == "replace":
                # i1..i2 を j1..j2 に置換 → ペアにできる範囲は modify、余りは delete/insert
                pair_count = min(i2 - i1, j2 - j1)
                for k in range(pair_count):
                    diff_rows.append((
                        "modify",
                        i1 + k + 1,
                        j1 + k + 1,
                        old_lines[i1 + k],
                        new_lines[j1 + k],
                    ))
                # 余りの古い側 → delete
                for k in range(pair_count, i2 - i1):
                    diff_rows.append((
                        "delete",
                        i1 + k + 1,
                        None,
                        old_lines[i1 + k],
                        "",
                    ))
                # 余りの新しい側 → insert
                for k in range(pair_count, j2 - j1):
                    diff_rows.append((
                        "insert",
                        None,
                        j1 + k + 1,
                        "",
                        new_lines[j1 + k],
                    ))
            elif tag == "delete":
                for k in range(i2 - i1):
                    diff_rows.append((
                        "delete",
                        i1 + k + 1,
                        None,
                        old_lines[i1 + k],
                        "",
                    ))
            elif tag == "insert":
                for k in range(j2 - j1):
                    diff_rows.append((
                        "insert",
                        None,
                        j1 + k + 1,
                        "",
                        new_lines[j1 + k],
                    ))

        return diff_rows

    def _build_diff_view(self, parent: tk.Frame, diff_rows: list) -> None:
        """差分ビュー（左右2列のテキストウィジェット、同期スクロール）"""
        # 全体枠
        wrap = tk.Frame(parent, bg=BG_PANEL,
                        highlightbackground=BORDER, highlightthickness=1)
        wrap.pack(fill="both", expand=True)

        # ヘッダ行（行番号 / 変更前 / 変更後）
        header = tk.Frame(wrap, bg=BG_HEADER, height=24)
        header.pack(fill="x")
        header.pack_propagate(False)

        # 行ヘッダ: Text(width=12) と同じ幅にするため、空のText で配置
        line_header_text = tk.Text(header, width=12, height=1,
                                    bg=BG_HEADER, fg=ACCENT,
                                    font=("Consolas", 10, "bold"),
                                    relief="flat", bd=0,
                                    highlightthickness=0, padx=6, pady=0,
                                    cursor="arrow")
        line_header_text.pack(side="left", fill="y")
        line_header_text.tag_configure("center", justify="center")
        line_header_text.insert("1.0", "行", "center")
        line_header_text.configure(state="disabled")
        tk.Frame(header, bg=BORDER, width=1).pack(side="left", fill="y")
        # 変更前/変更後 ヘッダ部: gridで完全均等化
        diff_header = tk.Frame(header, bg=BG_HEADER)
        diff_header.pack(side="left", fill="both", expand=True)
        diff_header.grid_columnconfigure(0, weight=1, uniform="diff_cols")
        diff_header.grid_columnconfigure(1, minsize=1, weight=0)
        diff_header.grid_columnconfigure(2, weight=1, uniform="diff_cols")
        diff_header.grid_rowconfigure(0, weight=1)
        tk.Label(diff_header, text="変更前",
                 bg=BG_HEADER, fg="#FF93A8",
                 font=("Consolas", 10, "bold"),
                 anchor="w", padx=8).grid(row=0, column=0, sticky="nsew")
        tk.Frame(diff_header, bg=BORDER).grid(row=0, column=1, sticky="nsew")
        tk.Label(diff_header, text="変更後",
                 bg=BG_HEADER, fg=ACCENT,
                 font=("Consolas", 10, "bold"),
                 anchor="w", padx=8).grid(row=0, column=2, sticky="nsew")
        # 縦スクロールバー幅分のダミー
        tk.Frame(header, bg=BG_HEADER, width=10).pack(side="right")

        # 本体（3列のText、同期スクロール）
        body = tk.Frame(wrap, bg=BG_PANEL)
        body.pack(fill="both", expand=True)

        text_kw = {"font": ("Consolas", 10),
                   "bg": INPUT_BG,
                   "insertbackground": ACCENT,
                   "relief": "flat", "bd": 0,
                   "wrap": "none", "cursor": "arrow",
                   "padx": 6, "pady": 4}

        # 行番号列
        line_text = tk.Text(body, width=12, fg=TEXT_MUTED, **text_kw)
        line_text.pack(side="left", fill="y")
        tk.Frame(body, bg=BORDER, width=1).pack(side="left", fill="y")

        # 縦スクロールバー（先に right に確保）
        scroll = tk.Scrollbar(body, orient="vertical",
                               bg=BG_PANEL, troughcolor=BG_HEADER,
                               activebackground=ACCENT_DARK,
                               bd=0, highlightthickness=0, width=10)
        scroll.pack(side="right", fill="y")

        # 変更前/変更後 本体部: gridで完全均等化
        diff_body = tk.Frame(body, bg=BG_PANEL)
        diff_body.pack(side="left", fill="both", expand=True)
        diff_body.grid_columnconfigure(0, weight=1, uniform="diff_cols")
        diff_body.grid_columnconfigure(1, minsize=1, weight=0)
        diff_body.grid_columnconfigure(2, weight=1, uniform="diff_cols")
        diff_body.grid_rowconfigure(0, weight=1)

        # 変更前列
        old_text = tk.Text(diff_body, fg="#FF93A8", **text_kw)
        old_text.grid(row=0, column=0, sticky="nsew")
        tk.Frame(diff_body, bg=BORDER).grid(row=0, column=1, sticky="nsew")

        # 変更後列
        new_text = tk.Text(diff_body, fg=ACCENT, **text_kw)
        new_text.grid(row=0, column=2, sticky="nsew")

        # 同期スクロール
        def _sync_yview(*args):
            line_text.yview(*args)
            old_text.yview(*args)
            new_text.yview(*args)

        def _on_scroll(*args):
            scroll.set(*args)
            # 他のText に位置を伝播（再帰防止のため delta方式は使わず set で同期）
            top, bot = args[0], args[1]
            line_text.yview_moveto(float(top))
            old_text.yview_moveto(float(top))
            new_text.yview_moveto(float(top))

        scroll.config(command=_sync_yview)
        line_text.config(yscrollcommand=_on_scroll)
        old_text.config(yscrollcommand=_on_scroll)
        new_text.config(yscrollcommand=_on_scroll)

        # 横スクロール（共通：bottom）
        hbar_frame = tk.Frame(wrap, bg=BG_PANEL)
        hbar_frame.pack(fill="x")

        # 行番号列分のスペーサー（line_text の width=12 と同じ Text を空で配置して幅完全一致）
        line_spacer = tk.Text(hbar_frame, width=12, height=1,
                              bg=BG_PANEL, relief="flat", bd=0,
                              highlightthickness=0, padx=6, pady=0)
        line_spacer.pack(side="left", fill="y")
        line_spacer.configure(state="disabled")
        tk.Frame(hbar_frame, bg=BG_PANEL, width=1).pack(side="left", fill="y")
        # 縦バー幅分のスペーサー（先にrightで確保）
        tk.Frame(hbar_frame, bg=BG_PANEL, width=10).pack(side="right")

        # 変更前/変更後 横バー部: gridで均等化
        hbar_inner = tk.Frame(hbar_frame, bg=BG_PANEL)
        hbar_inner.pack(side="left", fill="both", expand=True)
        hbar_inner.grid_columnconfigure(0, weight=1, uniform="diff_cols")
        hbar_inner.grid_columnconfigure(1, minsize=1, weight=0)
        hbar_inner.grid_columnconfigure(2, weight=1, uniform="diff_cols")
        hbar_inner.grid_rowconfigure(0, weight=1)
        old_hbar = tk.Scrollbar(hbar_inner, orient="horizontal",
                                 command=old_text.xview,
                                 bg=BG_PANEL, troughcolor=BG_HEADER,
                                 activebackground=ACCENT_DARK,
                                 bd=0, highlightthickness=0, width=10)
        new_hbar = tk.Scrollbar(hbar_inner, orient="horizontal",
                                 command=new_text.xview,
                                 bg=BG_PANEL, troughcolor=BG_HEADER,
                                 activebackground=ACCENT_DARK,
                                 bd=0, highlightthickness=0, width=10)
        old_hbar.grid(row=0, column=0, sticky="nsew")
        tk.Frame(hbar_inner, bg=BG_PANEL).grid(row=0, column=1, sticky="nsew")
        new_hbar.grid(row=0, column=2, sticky="nsew")
        old_text.configure(xscrollcommand=old_hbar.set)
        new_text.configure(xscrollcommand=new_hbar.set)

        # マウスホイール：3列同時スクロール
        def _on_mousewheel(event):
            delta = int(-1 * (event.delta / 120))
            line_text.yview_scroll(delta, "units")
            old_text.yview_scroll(delta, "units")
            new_text.yview_scroll(delta, "units")
            return "break"
        for w in (line_text, old_text, new_text):
            w.bind("<MouseWheel>", _on_mousewheel)

        # スタイルタグ設定
        # delete: 元側を赤背景＋取り消し線、新側は空表示で薄い背景
        # insert: 元側を空表示、新側を緑背景
        # modify: 通常の左右並び
        for w in (line_text, old_text, new_text):
            w.tag_configure("delete_old", background="#3A1620", foreground="#FF93A8",
                            overstrike=True)
            w.tag_configure("delete_new", background="#301019", foreground=TEXT_MUTED)
            w.tag_configure("insert_old", background="#141B24", foreground=TEXT_MUTED)
            w.tag_configure("insert_new", background="#123A38", foreground=ACCENT)
            w.tag_configure("modify_line", foreground=TEXT_MUTED)

        # 行ごとに内容を流し込む（kind ごとに表示を変える）
        for diff in diff_rows:
            kind = diff[0]
            old_no = diff[1]
            new_no = diff[2]
            old_content = diff[3]
            new_content = diff[4]

            if kind == "modify":
                line_label = f"{old_no}→{new_no}" if old_no != new_no else f"{old_no}"
                line_text.insert("end", f"{line_label}\n", "modify_line")
                old_text.insert("end", f"{old_content}\n")
                new_text.insert("end", f"{new_content}\n")
            elif kind == "delete":
                line_text.insert("end", f"{old_no}\n", "delete_old")
                old_text.insert("end", f"{old_content}\n", "delete_old")
                new_text.insert("end", "(削除)\n", "delete_new")
            elif kind == "insert":
                line_text.insert("end", f"{new_no}\n", "insert_new")
                old_text.insert("end", "(追加)\n", "insert_old")
                new_text.insert("end", f"{new_content}\n", "insert_new")

        # 編集不可
        line_text.configure(state="disabled")
        old_text.configure(state="disabled")
        new_text.configure(state="disabled")

    def _on_ok(self) -> None:
        self.confirmed = True
        self.destroy()

    def _on_cancel(self) -> None:
        self.confirmed = False
        self.destroy()


class RegexBuilderDialog(tk.Toplevel):
    """正規表現ビルダー：早見表＋ライブテストつきの汎用ダイアログ。
    呼び出し元から open_regex_builder() で呼び出され、
    OK 押下時に self.result_pattern に正規表現文字列がセットされる。
    キャンセル時は None。
    """

    # 早見表の項目定義
    CHEATSHEET = [
        ("基本", [
            (r"\d", "数字1文字（0-9）", "G01 → '0','1'"),
            (r"\s", "空白文字（スペース・タブ）", "'G01 X10' → スペース"),
            (r".", "何でも1文字", "abc → 'a','b','c'"),
            (r"\w", "英数字＋アンダースコア", "G01 → 'G','0','1'"),
        ]),
        ("繰り返し", [
            (r"+", "直前の1個以上", r"\d+ → '01','12'"),
            (r"*", "直前の0個以上", r"\d* → ''も含む"),
            (r"?", "直前の0個か1個", r"\d? → 0個か1個"),
            (r"{n}", "直前のちょうどn個", r"\d{3} → '012','345'"),
            (r"{n,m}", "直前のn〜m個", r"\d{1,3} → '0','01','012'"),
        ]),
        ("位置", [
            (r"^", "行頭にマッチ", r"^G → 行頭のG"),
            (r"$", "行末にマッチ", r"\.$ → 行末のドット"),
            (r"\b", "単語の境界", r"\bG\d → 'G01'は○"),
        ]),
        ("まとめ・選択", [
            (r"[ABC]", "AかBかCのいずれか", "[XYZ] → X,Y,Zのどれか"),
            (r"[A-Z]", "A〜Zのいずれか", "[A-Z] → 大文字英字"),
            (r"[^ABC]", "AでもBでもCでもない", "[^XYZ] → X,Y,Z以外"),
            (r"(A|B)", "AかB（どちらか）", "(G01|G02) → G01かG02"),
        ]),
        ("エスケープ", [
            (r"\.", "ドット文字そのもの", r"X10\. → 'X10.'"),
            (r"\(", "開きカッコ そのもの", r"\(test\) → '(test)'"),
            (r"\\", "バックスラッシュ そのもの", ""),
        ]),
        ("NCコードでよく使う", [
            (r"[XYZ][+-]?\d+\.?\d*", "X/Y/Z軸の値", "X10. Y-5. Z3.5 等"),
            (r"[GM]\d+", "G/Mコード", "G01, G96, M30 等"),
            (r"F[0-9.]+", "F値", "F0.2, F100 等"),
            (r"S\d+", "S値", "S2000, S100 等"),
            (r"N\d+", "Nブロック番号", "N100, N200 等"),
            (r"\([^)]*\)", "コメント", "(TEST) (T01 ...) 等"),
            (r"^\s*$", "空行", "全部空白の行"),
        ]),
    ]

    def __init__(self, parent, initial_pattern: str = "",
                 initial_test_text: str = "",
                 initial_ignore_case: bool = False):
        """parent: 呼び出し元 Toplevel または NcCheckApp
        initial_pattern: ビルダーを開いた時の初期パターン
        initial_test_text: テスト用文字列の初期値（呼び出し元から渡せる）
        initial_ignore_case: 大文字小文字無視チェックの初期状態
        """
        super().__init__(parent)
        self.parent_widget = parent
        self.result_pattern: str | None = None  # OK時にセットされる

        self.title("正規表現ビルダー  //  REGEX BUILDER")
        self.configure(bg=BG_PANEL)
        self.geometry("1020x700")
        self.minsize(880, 540)
        self.resizable(True, True)

        self._build_ui(initial_pattern, initial_test_text, initial_ignore_case)

        # 中央寄せ
        self.update_idletasks()
        self._center_on_parent()

        self.bind("<Escape>", lambda e: self._on_cancel())
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # 初回マッチ実行
        self._update_matches()

    def _center_on_parent(self) -> None:
        try:
            self.update_idletasks()
            pw = self.parent_widget.winfo_width()
            ph = self.parent_widget.winfo_height()
            px = self.parent_widget.winfo_rootx()
            py = self.parent_widget.winfo_rooty()
            w = self.winfo_width()
            h = self.winfo_height()
            x = px + (pw - w) // 2
            y = py + (ph - h) // 2
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass

    def _build_ui(self, initial_pattern: str, initial_test_text: str,
                  initial_ignore_case: bool) -> None:
        outer = tk.Frame(self, bg=BG_PANEL, padx=14, pady=12)
        outer.pack(fill="both", expand=True)

        # ヘッダ
        tk.Label(outer, text="▶ 正規表現ビルダー",
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 12, "bold")).pack(anchor="w")
        tk.Label(outer,
                 text="パターンを入力して、下のテスト用文字列でマッチを確認",
                 bg=BG_PANEL, fg=TEXT_MUTED,
                 font=("Yu Gothic UI", 9)).pack(anchor="w", pady=(0, 8))

        # ── 左右2列レイアウト ──
        body = tk.Frame(outer, bg=BG_PANEL)
        body.pack(fill="both", expand=True)

        # 左ペイン：パターン入力・テスト文字列・マッチ結果・確定ボタン
        left = tk.Frame(body, bg=BG_PANEL)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # 右ペイン：早見表（固定幅 260px）
        right = tk.Frame(body, bg=BG_PANEL, width=260)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)  # 幅固定

        # ── パターン入力欄 ──
        pat_frame = tk.Frame(left, bg=BG_PANEL)
        pat_frame.pack(fill="x", pady=(0, 4))
        tk.Label(pat_frame, text="パターン:",
                 bg=BG_PANEL, fg=TEXT_MAIN,
                 font=("Yu Gothic UI", 10, "bold"),
                 width=10, anchor="e").pack(side="left", padx=(0, 6))

        self._pattern_var = tk.StringVar(value=initial_pattern)
        self._pattern_entry = tk.Entry(
            pat_frame, textvariable=self._pattern_var,
            font=("Consolas", 12), bg=INPUT_BG, fg=ACCENT,
            insertbackground=ACCENT, relief="solid", bd=1,
            highlightcolor=ACCENT, highlightbackground=BORDER,
            highlightthickness=1,
        )
        self._pattern_entry.pack(side="left", fill="x", expand=True)
        self._pattern_var.trace_add("write", lambda *a: self._update_matches())

        # オプション
        opt_frame = tk.Frame(left, bg=BG_PANEL)
        opt_frame.pack(fill="x", pady=(2, 8))
        tk.Label(opt_frame, text="", bg=BG_PANEL, width=10).pack(side="left", padx=(0, 6))
        self._ignore_case_var = tk.StringVar(
            value="1" if initial_ignore_case else "0"
        )
        tk.Checkbutton(
            opt_frame, text="大文字小文字を区別しない",
            variable=self._ignore_case_var, onvalue="1", offvalue="0",
            command=self._update_matches,
            bg=BG_PANEL, fg=TEXT_MAIN, selectcolor=INPUT_BG,
            activebackground=BG_PANEL, activeforeground=ACCENT,
            font=("Yu Gothic UI", 9), bd=0, highlightthickness=0,
        ).pack(side="left")

        # エラー表示
        self._error_var = tk.StringVar(value="")
        tk.Label(left, textvariable=self._error_var,
                 bg=BG_PANEL, fg="#FB5E7E",
                 font=("Consolas", 9), anchor="w").pack(fill="x", padx=(80, 0))

        # ── テスト用文字列 ──
        tk.Label(left, text="── テスト用文字列 ──",
                 bg=BG_PANEL, fg=ACCENT_DARK,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", pady=(8, 4))

        test_btn_row = tk.Frame(left, bg=BG_PANEL)
        test_btn_row.pack(fill="x", pady=(0, 4))
        btn_kw = {"font": ("Yu Gothic UI", 9),
                  "bg": "#141B24", "fg": TEXT_MAIN,
                  "activebackground": "#1F2B3A", "activeforeground": ACCENT,
                  "relief": "solid", "bd": 1, "padx": 8, "pady": 2,
                  "cursor": "hand2"}
        tk.Button(test_btn_row, text="選択範囲を読込",
                  command=self._load_selection,
                  **btn_kw).pack(side="left", padx=(0, 4))
        tk.Button(test_btn_row, text="全文を読込",
                  command=self._load_all_text,
                  **btn_kw).pack(side="left", padx=(0, 4))
        tk.Button(test_btn_row, text="クリア",
                  command=self._clear_test_text,
                  **btn_kw).pack(side="left", padx=(0, 4))

        test_wrap = tk.Frame(left, bg=BG_PANEL,
                              highlightbackground=BORDER, highlightthickness=1)
        test_wrap.pack(fill="both", expand=True)

        self._test_text = tk.Text(
            test_wrap, height=8, font=("Consolas", 11),
            bg=INPUT_BG, fg=TEXT_MAIN, insertbackground=ACCENT,
            relief="flat", bd=0, wrap="none", padx=4, pady=4, undo=True,
        )
        test_vbar = tk.Scrollbar(test_wrap, orient="vertical",
                                  command=self._test_text.yview,
                                  bg=BG_PANEL, troughcolor=BG_HEADER,
                                  activebackground=ACCENT_DARK,
                                  bd=0, highlightthickness=0, width=10)
        self._test_text.configure(yscrollcommand=test_vbar.set)
        test_vbar.pack(side="right", fill="y")
        self._test_text.pack(side="left", fill="both", expand=True)

        if initial_test_text:
            self._test_text.insert("1.0", initial_test_text)

        # マッチ箇所のハイライトタグ
        self._test_text.tag_configure("regex_match",
                                       background="#123A38",
                                       foreground=ACCENT)

        # テスト文字列が変わったら再マッチ
        self._test_text.bind("<<Modified>>", self._on_test_text_modified)

        # ── マッチ結果 ──
        result_header = tk.Frame(left, bg=BG_PANEL)
        result_header.pack(fill="x", pady=(8, 2))
        tk.Label(result_header, text="── マッチ結果 ──",
                 bg=BG_PANEL, fg=ACCENT_DARK,
                 font=("Yu Gothic UI", 10, "bold")).pack(side="left")
        self._match_count_var = tk.StringVar(value="0 件")
        tk.Label(result_header, textvariable=self._match_count_var,
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Consolas", 10, "bold")).pack(side="right")

        result_wrap = tk.Frame(left, bg=BG_PANEL,
                                highlightbackground=BORDER, highlightthickness=1)
        result_wrap.pack(fill="both", expand=False, pady=(0, 4))

        self._result_text = tk.Text(
            result_wrap, height=5, font=("Consolas", 9),
            bg=INPUT_BG, fg=TEXT_MAIN, insertbackground=ACCENT,
            relief="flat", bd=0, wrap="none", padx=4, pady=2,
        )
        result_vbar = tk.Scrollbar(result_wrap, orient="vertical",
                                    command=self._result_text.yview,
                                    bg=BG_PANEL, troughcolor=BG_HEADER,
                                    activebackground=ACCENT_DARK,
                                    bd=0, highlightthickness=0, width=10)
        self._result_text.configure(yscrollcommand=result_vbar.set)
        result_vbar.pack(side="right", fill="y")
        self._result_text.pack(side="left", fill="both", expand=True)
        self._result_text.configure(state="disabled")

        # ── 右ペイン：早見表（常時表示）──
        tk.Label(right, text="── 記号一覧 ──",
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", pady=(0, 4))

        self._cheatsheet_frame = tk.Frame(right, bg=BG_PANEL,
                                           highlightbackground=BORDER,
                                           highlightthickness=1)
        self._cheatsheet_frame.pack(fill="both", expand=True)
        self._build_cheatsheet_content()

        # ── 確定ボタン（左ペイン下部）──
        ok_row = tk.Frame(left, bg=BG_PANEL)
        ok_row.pack(fill="x", pady=(8, 0))

        tk.Button(ok_row, text="この正規表現を使う",
                  command=self._on_ok,
                  font=("Yu Gothic UI", 10, "bold"),
                  bg="#123A38", fg=ACCENT,
                  activebackground="#175E56", activeforeground=ACCENT,
                  relief="solid", bd=1, padx=14, pady=4,
                  cursor="hand2").pack(side="right", padx=(4, 0))
        tk.Button(ok_row, text="キャンセル", command=self._on_cancel,
                  font=("Yu Gothic UI", 10),
                  bg="#141B24", fg=TEXT_MAIN,
                  activebackground="#1F2B3A", activeforeground=ACCENT,
                  relief="solid", bd=1, padx=14, pady=4,
                  cursor="hand2").pack(side="right")

    def _build_cheatsheet_content(self) -> None:
        """早見表の中身を構築（右ペインに常時表示）"""
        # スクロール対応
        cs_canvas = tk.Canvas(self._cheatsheet_frame,
                               bg=BG_PANEL, bd=0,
                               highlightthickness=0, takefocus=False,
                               height=240)
        cs_scroll = tk.Scrollbar(self._cheatsheet_frame, orient="vertical",
                                  command=cs_canvas.yview,
                                  bg=BG_PANEL, troughcolor=BG_HEADER,
                                  activebackground=ACCENT_DARK,
                                  bd=0, highlightthickness=0, width=10)
        cs_canvas.configure(yscrollcommand=cs_scroll.set)
        cs_scroll.pack(side="right", fill="y")
        cs_canvas.pack(side="left", fill="both", expand=True)

        cs_inner = tk.Frame(cs_canvas, bg=BG_PANEL)
        cs_inner_id = cs_canvas.create_window(
            (0, 0), window=cs_inner, anchor="nw"
        )

        def _on_cs_inner_configure(event):
            cs_canvas.configure(scrollregion=cs_canvas.bbox("all"))
        cs_inner.bind("<Configure>", _on_cs_inner_configure)

        def _on_cs_canvas_configure(event):
            cs_canvas.itemconfigure(cs_inner_id, width=event.width)
        cs_canvas.bind("<Configure>", _on_cs_canvas_configure)

        # マウスホイール
        def _on_mousewheel(event):
            cs_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _bind_wheel(e):
            cs_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_wheel(e):
            cs_canvas.unbind_all("<MouseWheel>")
        cs_canvas.bind("<Enter>", _bind_wheel)
        cs_canvas.bind("<Leave>", _unbind_wheel)

        # カテゴリごとに表示
        for category, items in self.CHEATSHEET:
            tk.Label(cs_inner, text=f"▼ {category}",
                     bg=BG_PANEL, fg=ACCENT_DARK,
                     font=("Yu Gothic UI", 10, "bold"),
                     anchor="w").pack(fill="x", padx=4, pady=(6, 2))

            for symbol, desc, example in items:
                # アイテム枠（区切り線つき）
                item_frame = tk.Frame(cs_inner, bg=BG_PANEL)
                item_frame.pack(fill="x", padx=4, pady=(2, 2))

                # 1行目：[挿入] ボタン + 記号
                row1 = tk.Frame(item_frame, bg=BG_PANEL)
                row1.pack(fill="x")

                # 挿入ボタン（左に配置）
                tk.Button(
                    row1, text="挿入",
                    command=lambda s=symbol: self._insert_symbol(s),
                    font=("Yu Gothic UI", 9),
                    bg="#141B24", fg=ACCENT,
                    activebackground="#1F2B3A", activeforeground=ACCENT,
                    relief="solid", bd=1, padx=6, pady=0, cursor="hand2",
                ).pack(side="left", padx=(0, 6))

                # 記号
                sym_label = tk.Label(
                    row1, text=symbol, bg=BG_PANEL, fg=ACCENT,
                    font=("Consolas", 11, "bold"),
                    anchor="w",
                )
                sym_label.pack(side="left")

                # 2行目：説明（小さく）
                row2 = tk.Frame(item_frame, bg=BG_PANEL)
                row2.pack(fill="x", padx=(40, 0))

                desc_label = tk.Label(
                    row2, text=desc, bg=BG_PANEL, fg=TEXT_MAIN,
                    font=("Yu Gothic UI", 9), anchor="w",
                    wraplength=200, justify="left",
                )
                desc_label.pack(side="top", anchor="w")

                # 3行目：例（あれば、さらに薄色）
                if example:
                    ex_label = tk.Label(
                        item_frame, text=f"例: {example}",
                        bg=BG_PANEL, fg=TEXT_MUTED,
                        font=("Consolas", 8), anchor="w",
                        wraplength=210, justify="left",
                    )
                    ex_label.pack(fill="x", padx=(40, 0))

    def _insert_symbol(self, symbol: str) -> None:
        """早見表からの記号挿入：パターン入力欄のカーソル位置に挿入"""
        try:
            self._pattern_entry.insert("insert", symbol)
            self._pattern_entry.focus_set()
        except tk.TclError:
            pass

    # ──────────────────────────
    # テスト文字列の取得（選択範囲・全文）
    # ──────────────────────────

    def _get_main_text_widget(self):
        """呼び出し元から辿ってメインの入力 Text（NcCheckApp.input_text）を取得"""
        # parent_widget が NcCheckApp の場合
        if hasattr(self.parent_widget, "input_text"):
            return self.parent_widget.input_text
        # parent_widget が NcCheckApp.root の場合
        if hasattr(self.parent_widget, "winfo_toplevel"):
            top = self.parent_widget.winfo_toplevel()
            # FindReplaceDialog から呼ばれた場合：app 経由
            if hasattr(self.parent_widget, "app"):
                app = self.parent_widget.app
                if hasattr(app, "input_text"):
                    return app.input_text
            # MacroStepDialog から呼ばれた場合：parent_dialog → app
            if hasattr(self.parent_widget, "parent_dialog"):
                pd = self.parent_widget.parent_dialog
                if hasattr(pd, "app") and hasattr(pd.app, "input_text"):
                    return pd.app.input_text
        return None

    def _load_selection(self) -> None:
        """メイン画面の選択範囲をテスト用文字列に取り込む"""
        widget = self._get_main_text_widget()
        if widget is None:
            self._error_var.set("⚠ メインのエディタを取得できません")
            return
        try:
            sel = widget.get("sel.first", "sel.last")
            if not sel:
                self._error_var.set("⚠ メイン画面で範囲を選択してください")
                return
            self._test_text.delete("1.0", "end")
            self._test_text.insert("1.0", sel)
            self._error_var.set("")
            self._update_matches()
        except tk.TclError:
            self._error_var.set("⚠ 選択範囲がありません")

    def _load_all_text(self) -> None:
        """メイン画面の全文をテスト用文字列に取り込む"""
        widget = self._get_main_text_widget()
        if widget is None:
            self._error_var.set("⚠ メインのエディタを取得できません")
            return
        try:
            content = widget.get("1.0", "end-1c")
            if not content:
                self._error_var.set("⚠ メイン画面が空です")
                return
            self._test_text.delete("1.0", "end")
            self._test_text.insert("1.0", content)
            self._error_var.set("")
            self._update_matches()
        except tk.TclError:
            self._error_var.set("⚠ 取得できませんでした")

    def _clear_test_text(self) -> None:
        self._test_text.delete("1.0", "end")
        self._update_matches()

    def _on_test_text_modified(self, event=None) -> None:
        try:
            if self._test_text.edit_modified():
                self._update_matches()
                self._test_text.edit_modified(False)
        except tk.TclError:
            pass

    # ──────────────────────────
    # ライブマッチ実行
    # ──────────────────────────

    def _update_matches(self) -> None:
        """パターン or テスト文字列が変わったらマッチ再計算＆表示更新"""
        # パターン取得
        pattern = self._pattern_var.get()
        ignore_case = self._ignore_case_var.get() == "1"
        flags = re.IGNORECASE if ignore_case else 0

        # ハイライトクリア
        try:
            self._test_text.tag_remove("regex_match", "1.0", "end")
        except tk.TclError:
            pass

        # 結果欄クリア
        self._result_text.configure(state="normal")
        self._result_text.delete("1.0", "end")

        if not pattern:
            self._error_var.set("")
            self._match_count_var.set("0 件")
            self._result_text.configure(state="disabled")
            return

        # 正規表現コンパイル
        try:
            pat = re.compile(pattern, flags)
        except re.error as e:
            self._error_var.set(f"⚠ 正規表現エラー: {e}")
            self._match_count_var.set("- 件")
            self._result_text.configure(state="disabled")
            return

        self._error_var.set("")

        # テスト文字列取得
        test_text = self._test_text.get("1.0", "end-1c")
        if not test_text:
            self._match_count_var.set("0 件 (テスト文字列なし)")
            self._result_text.configure(state="disabled")
            return

        # マッチ実行
        matches = list(pat.finditer(test_text))
        if not matches:
            self._match_count_var.set("0 件 (マッチなし)")
            self._result_text.insert("end", "(マッチなし)\n")
            self._result_text.configure(state="disabled")
            return

        # ハイライト＆結果リスト
        self._match_count_var.set(f"{len(matches)} 件")

        for i, m in enumerate(matches, 1):
            start_idx = self._offset_to_index(test_text, m.start())
            end_idx = self._offset_to_index(test_text, m.end())
            try:
                self._test_text.tag_add("regex_match", start_idx, end_idx)
            except tk.TclError:
                pass

            # 結果行：「N: 'マッチテキスト' (行L, col C-C)」
            line_no, col = self._offset_to_line_col(test_text, m.start())
            match_text = m.group(0)
            # 改行入りの場合は↵で表示
            disp = match_text.replace("\n", "↵")
            if len(disp) > 40:
                disp = disp[:38] + "…"
            self._result_text.insert("end",
                                      f"{i:3d}: '{disp}'  (行{line_no}, col{col}-{col + len(match_text)})\n")

            # キャプチャグループも表示
            for gi, g in enumerate(m.groups(), 1):
                if g is None:
                    continue
                gd = g.replace("\n", "↵")
                if len(gd) > 40:
                    gd = gd[:38] + "…"
                self._result_text.insert("end",
                                          f"      └ グループ{gi}: '{gd}'\n")

        self._result_text.configure(state="disabled")

    @staticmethod
    def _offset_to_index(text: str, offset: int) -> str:
        """文字列内オフセット → tkinter Text の行.列インデックス文字列"""
        # 高速化のため offset までの改行数で行を計算
        line_no = text.count("\n", 0, offset) + 1
        last_nl = text.rfind("\n", 0, offset)
        col = offset - (last_nl + 1) if last_nl >= 0 else offset
        return f"{line_no}.{col}"

    @staticmethod
    def _offset_to_line_col(text: str, offset: int) -> tuple:
        """文字列内オフセット → (行番号, 列) のタプル（1始まり）"""
        line_no = text.count("\n", 0, offset) + 1
        last_nl = text.rfind("\n", 0, offset)
        col = offset - (last_nl + 1) if last_nl >= 0 else offset
        return (line_no, col + 1)

    # ──────────────────────────
    # 確定/キャンセル
    # ──────────────────────────

    def _on_ok(self) -> None:
        # 最終バリデーション
        pat = self._pattern_var.get()
        if not pat:
            self._error_var.set("⚠ パターンを入力してください")
            return
        try:
            flags = re.IGNORECASE if self._ignore_case_var.get() == "1" else 0
            re.compile(pat, flags)
        except re.error as e:
            self._error_var.set(f"⚠ 正規表現エラー: {e}")
            return
        self.result_pattern = pat
        self.destroy()

    def _on_cancel(self) -> None:
        self.result_pattern = None
        self.destroy()


class SplashScreen(tk.Toplevel):
    """Ghost Protocolテーマのスプラッシュスクリーン"""

    def __init__(self, root: tk.Tk) -> None:
        super().__init__(root)
        self.overrideredirect(True)
        self.configure(bg="#0B0F14")
        self.attributes("-topmost", True)

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = 420, 260
        x, y = (sw - w) // 2, (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.canvas = tk.Canvas(self, bg="#0B0F14", width=w, height=h,
                                bd=0, highlightthickness=0)
        self.canvas.pack()

        # 外枠（Ghost Protocolグリーン）
        self.canvas.create_rectangle(1, 1, w - 1, h - 1, outline="#22D3C5", width=1)

        # タイトル
        self.canvas.create_text(w // 2, 55, text="NC PROGRAM CHECKER",
                                font=("Courier", 16, "bold"), fill="#22D3C5")
        self.canvas.create_text(w // 2, 90,
                                text="NCプログラム確認ツール",
                                font=("Helvetica", 13), fill="#8296A6")
        self.canvas.create_text(w // 2, 120, text=APP_TITLE.split()[-1] if " " in APP_TITLE else "",
                                font=("Courier", 11), fill="#2A3846")

        # プログレスバー
        bar_y, bar_w, bar_h = 160, 280, 4
        bx1 = (w - bar_w) // 2
        bx2 = bx1 + bar_w
        self.canvas.create_rectangle(bx1, bar_y, bx2, bar_y + bar_h,
                                     fill="#141B24", outline="")
        self._bar = self.canvas.create_rectangle(bx1, bar_y, bx1, bar_y + bar_h,
                                                 fill="#22D3C5", outline="")
        self._bar_x1 = bx1
        self._bar_x2 = bx2
        self._bar_y = bar_y
        self._bar_h = bar_h

        # ステータス
        self._status = self.canvas.create_text(w // 2, 190, text="初期化中...",
                                               font=("Courier", 9), fill="#8296A6")

        # GHOST PROTOCOL
        self.canvas.create_line(120, 220, 300, 220, fill="#1E2833", width=1)
        self.canvas.create_text(w // 2, 238, text="GHOST PROTOCOL",
                                font=("Courier", 8), fill="#2A3846")

        # パーティクル
        import math as _math
        self._math = _math
        self._particles: list[tuple[int, float, float]] = []
        for i in range(15):
            px = 20 + (w - 40) * ((hash(str(i)) % 100) / 100)
            py = 20 + (h - 40) * ((hash(str(i + 50)) % 100) / 100)
            size = 1 + hash(str(i + 99)) % 2
            pid = self.canvas.create_oval(px, py, px + size, py + size,
                                          fill="#22D3C5", outline="")
            self._particles.append((pid, px, py))

        self._anim_step = 0
        self._animate()

    def _animate(self) -> None:
        self._anim_step += 1
        math = self._math
        for pid, ox, oy in self._particles:
            dx = math.sin(self._anim_step * 0.05 + ox) * 0.5
            dy = math.cos(self._anim_step * 0.04 + oy) * 0.3
            self.canvas.move(pid, dx, dy)
            g_base = int(150 + 40 * math.sin(self._anim_step * 0.08 + ox))
            g_base = max(120, min(211, g_base))
            b_val = max(110, min(197, g_base - 14))
            color = f"#22{g_base:02x}{b_val:02x}"
            try:
                self.canvas.itemconfig(pid, fill=color)
            except Exception:
                pass
        if self.winfo_exists():
            self.after(50, self._animate)

    def set_progress(self, value: float, text: str = "") -> None:
        self._progress = max(0.0, min(1.0, value))
        x = self._bar_x1 + (self._bar_x2 - self._bar_x1) * self._progress
        self.canvas.coords(self._bar, self._bar_x1, self._bar_y,
                           x, self._bar_y + self._bar_h)
        if text:
            self.canvas.itemconfig(self._status, text=text)
        self.update_idletasks()


def main() -> None:
    root = tk.Tk()

    # PyInstaller スプラッシュを閉じる（exe化時のみ動作、自前 SplashScreen に引き継ぐ）
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except ImportError:
        pass

    # コマンドライン引数からファイルパスを取得
    # Windows「プログラムから開く」「右クリック→このアプリで開く」経由で渡される
    initial_file = None
    if len(sys.argv) >= 2:
        candidate = sys.argv[1]
        # ファイルが実在する場合のみ採用
        if os.path.isfile(candidate):
            initial_file = candidate

    app = NcCheckApp(root, splash_mode=True, initial_file=initial_file)
    root.mainloop()


if __name__ == "__main__":
    main()
