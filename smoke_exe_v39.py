"""配布exeの実起動とウィンドウ表示を検証する。"""
import ctypes
from ctypes import wintypes
import subprocess
import time
from pathlib import Path
from PIL import ImageGrab

base = Path(__file__).resolve().parent
exe = base.parent.parent / '最新版/nc_program_check_tool_v39/nc_program_check_tool_v39.exe'
user = ctypes.windll.user32
user.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user.IsWindowVisible.argtypes = [wintypes.HWND]
user.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
callback = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
windows = []
def enum(hwnd, _):
    pid = wintypes.DWORD()
    user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if pid.value == process.pid and user.IsWindowVisible(hwnd):
        title = ctypes.create_unicode_buffer(512)
        user.GetWindowTextW(hwnd, title, 512)
        windows.append((hwnd, title.value))
    return True

startup = subprocess.STARTUPINFO()
startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
startup.wShowWindow = 0
process = subprocess.Popen([str(exe)], cwd=exe.parent, startupinfo=startup)
main = None
deadline = time.monotonic() + 30
while time.monotonic() < deadline and process.poll() is None:
    windows.clear()
    user.EnumWindows(callback(enum), 0)
    main = None
    for hwnd, title in windows:
        rect = wintypes.RECT()
        user.GetWindowRect(hwnd, ctypes.byref(rect))
        if title == 'NCプログラム確認ツール v39' and rect.right-rect.left >= 1000:
            main = hwnd
            break
    if main:
        time.sleep(1)
        rect = wintypes.RECT()
        user.GetWindowRect(main, ctypes.byref(rect))
        ImageGrab.grab(bbox=(max(0,rect.left),max(0,rect.top),rect.right,rect.bottom)).save(base/'v39_exe_preview.png')
        break
    time.sleep(.1)
if not main:
    ImageGrab.grab().save(base/'v39_exe_failure.png')
    raise RuntimeError(f'exe did not reach main window: code={process.poll()}, windows={windows}')
mascot_found = False
deadline = time.monotonic() + 12
while time.monotonic() < deadline:
    windows.clear()
    user.EnumWindows(callback(enum), 0)
    for hwnd, title in windows:
        rect = wintypes.RECT()
        user.GetWindowRect(hwnd, ctypes.byref(rect))
        if 150 <= rect.right-rect.left <= 165 and 140 <= rect.bottom-rect.top <= 160:
            # EnumWindows lists top-level windows in front-to-back order.
            mascot_found = [item[0] for item in windows].index(hwnd) < [item[0] for item in windows].index(main)
            break
    if mascot_found:
        break
    time.sleep(.1)
user.PostMessageW(main, 0x0010, 0, 0)
process.wait(timeout=15)
assert process.returncode == 0, process.returncode
assert mascot_found, 'Packaged mascot did not appear'
print('Packaged exe: startup, main window, mascot visible, clean exit OK')

