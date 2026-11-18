"""
mascot.py — ツール画面の上をドラッグで動かせるマスコット（Windows / tkinter）

使い方:
    from mascot import Mascot
    Mascot(root, "mascot.webp", ui_bg="#0B1622")

仕組み:
    - 枠なしの子ウィンドウ（Toplevel）を親ウィンドウの上に重ねる
    - 1色抜き（-transparentcolor）で羊の外側を透過させる
    - 羊の内側の暗い部分は、UI背景色とのスクリーン合成で背景に馴染ませる
    - 親ウィンドウの移動・最小化に追従し、ドラッグ範囲は親ウィンドウ内に制限する

操作:
    左ドラッグ … 移動
    右クリック … メニュー（アニメ停止/再開、隠す）
"""
import os
import sys
import tkinter as tk

from PIL import Image, ImageChops, ImageSequence, ImageTk

KEY = "#FF00FF"  # 透過に使う色（羊の画像内には出てこない色）


def resource_path(name):
    """PyInstaller（onedir/onefile）でもソース実行でも同梱ファイルを見つける"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def load_frames(path, ui_bg, scale=1.0):
    """RGBAのアニメWebPを、1色抜き用に加工したPhotoImageのリストにする"""
    src = Image.open(path)
    delay = src.info.get("duration", 83)
    frames = []
    for f in ImageSequence.Iterator(src):
        f = f.convert("RGBA")
        if scale != 1.0:
            f = f.resize((round(f.width * scale), round(f.height * scale)), Image.LANCZOS)
        rgb = ImageChops.screen(f.convert("RGB"), Image.new("RGB", f.size, ui_bg))
        mask = f.getchannel("A").point(lambda v: 255 if v >= 128 else 0)
        out = Image.composite(rgb, Image.new("RGB", f.size, KEY), mask)
        frames.append(ImageTk.PhotoImage(out))  # 参照をリストで保持（消えるの防止）
    return frames, delay


class Mascot(tk.Toplevel):
    def __init__(self, master, path, ui_bg="#0B1622", scale=1.0, start=("right", "bottom")):
        super().__init__(master)
        self.host = master.winfo_toplevel()
        self.withdraw()
        self.overrideredirect(True)
        self.transient(self.host)
        self.configure(bg=KEY)
        self.attributes("-transparentcolor", KEY)  # Windows専用

        self.frames, self.delay = load_frames(path, ui_bg, scale)
        self.i = 0
        self.playing = True
        w, h = self.frames[0].width(), self.frames[0].height()
        self.size = (w, h)

        self.lbl = tk.Label(self, bg=KEY, bd=0, highlightthickness=0, cursor="fleur")
        self.lbl.pack()

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="アニメ停止/再開", command=self.toggle)
        self.menu.add_command(label="隠す", command=self.withdraw)

        self.lbl.bind("<ButtonPress-1>", self._press)
        self.lbl.bind("<B1-Motion>", self._drag)
        self.lbl.bind("<Button-3>", lambda e: self.menu.tk_popup(e.x_root, e.y_root))
        self.host.bind("<Configure>", self._follow, add="+")
        self.host.bind("<Unmap>", lambda e: e.widget is self.host and self.withdraw(), add="+")
        self.host.bind("<Map>", lambda e: e.widget is self.host and self._show(), add="+")
        self.host.bind("<FocusIn>", lambda e: self.lift(), add="+")

        # 初期位置（親ウィンドウ内の相対座標）
        self.host.update_idletasks()
        hw, hh = self.host.winfo_width(), self.host.winfo_height()
        self.rx = hw - w - 16 if start[0] == "right" else 16
        self.ry = hh - h - 16 if start[1] == "bottom" else 16
        self._show()
        self._tick()

    # ---- 表示・追従 ----
    def _show(self):
        self.deiconify()
        self._place()
        self.lift()

    def _clamp(self):
        hw, hh = self.host.winfo_width(), self.host.winfo_height()
        w, h = self.size
        self.rx = max(0, min(self.rx, hw - w))
        self.ry = max(0, min(self.ry, hh - h))

    def _place(self):
        self._clamp()
        x = self.host.winfo_rootx() + self.rx
        y = self.host.winfo_rooty() + self.ry
        self.geometry(f"+{x}+{y}")

    def _follow(self, e):
        if e.widget is self.host and self.winfo_viewable():
            self._place()

    # ---- ドラッグ ----
    def _press(self, e):
        self._off = (e.x, e.y)

    def _drag(self, e):
        self.rx = e.x_root - self._off[0] - self.host.winfo_rootx()
        self.ry = e.y_root - self._off[1] - self.host.winfo_rooty()
        self._place()

    # ---- アニメ ----
    def toggle(self):
        self.playing = not self.playing

    def _tick(self):
        if self.playing:
            self.lbl.config(image=self.frames[self.i])
            self.i = (self.i + 1) % len(self.frames)
        self.after(self.delay, self._tick)


if __name__ == "__main__":
    BG = "#0B1622"  # ツールの背景色に合わせる
    root = tk.Tk()
    root.title("mascot demo")
    root.geometry("900x600")
    root.configure(bg=BG)
    tk.Label(root, text="羊をドラッグして動かしてみて", fg="#7FE3FF", bg=BG,
             font=("Meiryo", 14)).pack(pady=20)
    Mascot(root, resource_path("mascot.webp"), ui_bg=BG, scale=1.0)
    root.mainloop()
