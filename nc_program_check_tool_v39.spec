# -*- mode: python ; coding: utf-8 -*-
#
# NCプログラム確認ツール v39 ビルド設定
# 等価な従来コマンド:
#   python -m PyInstaller --onedir --noupx --noconsole nc_program_check_tool_v39.py
# + スプラッシュ画像対応
#
# ビルド方法:
#   python -m PyInstaller nc_program_check_tool_v39.spec
#
# 必要ファイル（同じフォルダに置くこと）:
#   - nc_program_check_tool_v39.py
#   - splash_v37.png

from pathlib import Path
from shutil import copy2, copytree

a = Analysis(
    ['nc_program_check_tool_v39.py'],
    pathex=[],
    binaries=[],
    datas=[('wooln-wireframe-splash.png', '.'), ('mascot.webp', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'pandas', 'matplotlib', 'scipy', 'IPython', 'pytest'],
    noarchive=False,
)

pyz = PYZ(a.pure)

# スプラッシュ画像
splash = Splash(
    'splash_start.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,           # text_pos=None で動的テキスト無効（画像のまま表示）
    max_img_size=(836, 471),  # シンプルな起動準備中画像をコンパクトに表示
    text_size=10,
    text_color='#22D3C5',
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,                  # スプラッシュをexeに含める
    [],
    exclude_binaries=True,   # --onedir 用（バイナリは COLLECT で同梱）
    name='nc_program_check_tool_v39',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,               # --noupx（UPX圧縮しない）
    console=False,           # --noconsole（コンソール出さない）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',       # アイコンファイルがあればコメント解除
)

coll = COLLECT(
    exe,
    splash.binaries,         # スプラッシュ用バイナリも同梱
    a.binaries,
    a.datas,
    strip=False,
    upx=False,               # --noupx
    upx_exclude=[],
    name='nc_program_check_tool_v39',
)

# 編集可能な設定はアプリが読み書きするexeと同じ階層へ配置する。
source_dir = Path(SPECPATH)
app_dir = Path(DISTPATH) / 'nc_program_check_tool_v39'
for filename in ('gm_glossary.json', 'nc_macros.json'):
    copy2(source_dir / filename, app_dir / filename)
if (source_dir / 'macros').is_dir():
    copytree(source_dir / 'macros', app_dir / 'macros', dirs_exist_ok=True)
