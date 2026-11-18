"""Small, static decorations; native widgets keep their original behaviour."""
import sys
import tkinter as tk
from tkinter import font as tkfont


def round_native(widget, radius=8):
    """Clip only the widget's outside corners, preserving native event handling."""
    if sys.platform != 'win32' or getattr(widget, '_nc_rounded', False):
        return
    widget._nc_rounded = True
    import ctypes
    from ctypes import wintypes
    gdi = ctypes.WinDLL('gdi32', use_last_error=True)
    user = ctypes.WinDLL('user32', use_last_error=True)
    gdi.CreateRoundRectRgn.argtypes = [ctypes.c_int] * 6
    gdi.CreateRoundRectRgn.restype = wintypes.HANDLE
    gdi.DeleteObject.argtypes = [wintypes.HANDLE]
    user.SetWindowRgn.argtypes = [wintypes.HWND, wintypes.HANDLE, wintypes.BOOL]
    last = [None]

    def resize(event):
        if event.widget is not widget:
            return
        size = (event.width, event.height)
        if size == last[0] or min(size) < 2:
            return
        last[0] = size
        r = min(radius, min(size) // 2)
        region = gdi.CreateRoundRectRgn(0, 0, size[0] + 1, size[1] + 1, r * 2, r * 2)
        if region and not user.SetWindowRgn(widget.winfo_id(), region, True):
            gdi.DeleteObject(region)
        # Windows owns the region after a successful SetWindowRgn call.

    widget.bind('<Configure>', resize, add='+')
    from types import SimpleNamespace
    resize(SimpleNamespace(widget=widget, width=widget.winfo_width(), height=widget.winfo_height()))


class HeaderText(tk.Label):
    """Unmapped label retains the existing configure(text=...) interface."""
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

    def configure(self, cnf=None, **kwargs):
        result = super().configure(cnf, **kwargs)
        if cnf or kwargs:
            self.master.draw_text()
        return result

    config = configure


class GradientHeader(tk.Canvas):
    def __init__(self, parent, en, ja, background, ink, muted):
        self.title_font = tkfont.Font(family='Yu Gothic UI', size=10, weight='bold')
        self.small_font = tkfont.Font(family='Consolas', size=8)
        self.suffix_font = tkfont.Font(family='Consolas', size=11, weight='bold')
        height = max(self.title_font.metrics('linespace'), self.small_font.metrics('linespace')) + 16
        super().__init__(parent, bg=background, height=height, bd=0, highlightthickness=0)
        self.en, self.ja, self.ink, self.muted = en, ja, ink, muted
        self.base = tuple(int(background[i:i+2], 16) for i in (1, 3, 5))
        self.hint = HeaderText(self, text='', fg=muted)
        self.extra = None
        self.controls = []
        self.suffix = ""
        self.collapsed = False
        self._size = None
        self.bind('<Configure>', self._draw)

    def _draw(self, event):
        if (event.width, event.height) != self._size:
            self._size = (event.width, event.height)
            self.delete('gradient')
            for y in range(event.height):
                factor = 1.24 - .48 * y / max(1, event.height - 1)
                color = '#' + ''.join(f'{max(0, min(255, round(c * factor))):02x}' for c in self.base)
                self.create_line(0, y, event.width, y, fill=color, tags='gradient')
            self.tag_lower('gradient')
        self.draw_text()

    def draw_text(self):
        self.delete('caption')
        width, height = self.winfo_width(), self.winfo_height()
        if self.collapsed:
            self.create_text(width/2, 42, text="\n".join(self.ja), font=self.title_font,
                             fill=self.ink, anchor='n', tags='caption')
            for widget, item in self.controls:
                self.coords(item, width/2, 18)
            return
        self.create_text(14, height/2, text=self.ja, font=self.title_font, fill=self.ink, anchor='w', tags='caption')
        x = 14 + self.title_font.measure(self.ja) + 10
        if self.suffix:
            self.create_text(x, height/2, text=self.suffix, font=self.suffix_font,
                             fill=self.ink, anchor='w', tags='caption')
            x += self.suffix_font.measure(self.suffix) + 10
        for widget, item in self.controls:
            w = widget.winfo_reqwidth()
            self.coords(item, x+w/2, height/2)
            x += w+8
        hint = self.hint.cget('text')
        right = width-12
        if self.extra:
            text, callback = self.extra
            extra_width = self.title_font.measure(text)
            self.create_text(right, height/2, text=text, font=self.title_font, fill=self.ink,
                             anchor='e', tags=('caption','extra'))
            self.tag_bind('extra','<Button-1>', lambda e: callback())
            self.tag_bind('extra','<Enter>', lambda e: self.configure(cursor='hand2'))
            self.tag_bind('extra','<Leave>', lambda e: self.configure(cursor=''))
            right -= extra_width+10
        # Prefer titles and their controls over optional trailing metadata.
        if self.small_font.measure(hint) <= max(0,right-x):
            self.create_text(right, height/2, text=hint, font=self.small_font, fill=self.muted,
                             anchor='e', tags='caption')

    def add_control(self, widget):
        item = self.create_window(0, 0, window=widget)
        self.controls.append((widget, item))
        widget.bind('<Configure>', lambda e: self.draw_text(), add='+')
        self.draw_text()

    def add_check(self, variable, command):
        check = tk.Checkbutton(self, variable=variable, command=command, bg=self.cget('bg'),
                               fg=self.ink, activebackground=self.cget('bg'), activeforeground=self.ink,
                               selectcolor='#080C11', bd=0, highlightthickness=0, padx=0, pady=0,
                               cursor='hand2', takefocus=True)
        self.add_control(check)
        return check

    def set_suffix(self, text):
        self.suffix = text
        self.draw_text()

    def add_action(self, text, callback):
        self.extra = (text, callback)
        self.draw_text()


class SmoothButton(tk.Button):
    """Antialiased flat button face with native keyboard/command/state semantics."""
    def __init__(self, master, **kw):
        from PIL import Image, ImageDraw, ImageTk
        self._nc_rounded = True
        compact = kw.pop('compact', False)
        face = kw.pop('bg', '#080C11')
        hover = kw.pop('activebackground', face)
        parent_bg = master.cget('bg')
        f = tkfont.Font(font=kw.get('font', ('Yu Gothic UI', 10)))
        width = f.measure(kw.get('text', '')) + (16 if compact else 28)
        height = f.metrics('linespace') + (8 if compact else 20)
        for key in ('width', 'height', 'padx', 'pady', 'bd', 'relief', 'highlightthickness'):
            kw.pop(key, None)
        super().__init__(master, **kw, bg=parent_bg, activebackground=parent_bg,
                         bd=0, relief='flat', highlightthickness=0, padx=0, pady=0,
                         compound='center', takefocus=True)
        self._faces = []
        for color, border in ((face, '#2A3846'), (hover, '#22D3C5'), (face, '#22D3C5')):
            image = Image.new('RGB', (width*3, height*3), parent_bg)
            draw = ImageDraw.Draw(image)
            draw.rounded_rectangle((1,1,width*3-2,height*3-2), radius=24, fill=color, outline=border, width=3)
            self._faces.append(ImageTk.PhotoImage(image.resize((width,height), Image.Resampling.LANCZOS), master=self))
        self.configure(image=self._faces[0])
        self.bind('<Enter>', lambda e: self._face(1), add='+')
        self.bind('<Leave>', lambda e: self._face(0), add='+')
        self.bind('<FocusIn>', lambda e: self._face(2), add='+')
        self.bind('<FocusOut>', lambda e: self._face(0), add='+')

    def _face(self, index):
        if self.cget('state') == 'disabled':
            index = 0
        elif self.focus_get() is self:
            index = 2
        self.configure(image=self._faces[index])
