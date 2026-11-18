"""Shared replacement preview, with a virtualized comparison list."""
import bisect
import difflib
import math
import re
import tkinter as tk
from tkinter import font as tkfont, ttk
from ui_finish_v39 import GradientHeader, SmoothButton, round_native


class ReplacementPreview(tk.Toplevel):
    def __init__(self, parent, title, rows, indices, on_confirm, grouped=False):
        super().__init__(parent)
        self.title(title)
        self.configure(bg='#131A22')
        self.transient(parent)
        self.geometry('1000x720')
        self.minsize(700, 480)
        self.rows, self.indices = rows, indices
        self.selected = set(indices)
        self.on_confirm = on_confirm
        self.grouped = grouped
        self.unit = '箇所' if grouped else '行'
        self.comment = tk.BooleanVar(value=False)
        self._layouts = []
        self._offsets = []
        self._check_widgets = []
        self._pending = None
        self._width = 0
        self._cache = {}
        self.code_font = tkfont.Font(self, family='Consolas', size=11)
        self.line_height = self.code_font.metrics('linespace') + 5
        self.char_width = max(1, self.code_font.measure('0'))
        self._char_sizes = {}
        head = GradientHeader(self, '', '変更内容の確認', '#1B4B73', '#FFFFFF', '#A8CCEA')
        head.pack(fill='x', padx=16, pady=(16, 0))
        tk.Label(self, text=title, bg='#131A22', fg='#B8CDDD',
                 font=('Yu Gothic UI', 10), anchor='w', wraplength=920).pack(fill='x', padx=20, pady=(10, 0))
        toolbar = tk.Frame(self, bg='#131A22')
        toolbar.pack(fill='x', padx=20, pady=10)
        self.count = tk.StringVar()
        tk.Label(toolbar, textvariable=self.count, bg='#131A22', fg='#DCEAF4',
                 font=('Yu Gothic UI', 11)).pack(side='left')
        self._button(toolbar, 'すべて解除', lambda: self.select_all(False)).pack(side='right')
        self._button(toolbar, 'すべて選択', lambda: self.select_all(True)).pack(side='right', padx=8)
        if not grouped:
            tk.Checkbutton(self, text='元の値をコメントに残す', variable=self.comment,
                           command=self._comment_changed, bg='#131A22', fg='#B8CDDD',
                           selectcolor='#080C11', activebackground='#131A22',
                           activeforeground='#FFFFFF', font=('Yu Gothic UI', 10)).pack(anchor='w', padx=20, pady=(0, 8))
        footer = tk.Frame(self, bg='#17212C', padx=20, pady=14)
        footer.pack(side='bottom', fill='x')
        tk.Label(footer, text='Ctrl＋Zで元に戻せます', bg='#17212C', fg='#A8BAC9',
                 font=('Yu Gothic UI', 9)).pack(side='left')
        self.execute = self._button(footer, '選択した項目を適用', self._execute)
        self.execute.pack(side='right')
        self._button(footer, 'キャンセル', self.destroy).pack(side='right', padx=8)
        body = tk.Frame(self, bg='#131A22')
        body.pack(fill='both', expand=True, padx=16, pady=(0, 14))
        self.canvas = tk.Canvas(body, bg='#131A22', bd=0, highlightthickness=0, yscrollincrement=24)
        self.scrollbar = ttk.Scrollbar(body, orient='vertical', command=self._scroll,
                                        style='Nc.Vertical.TScrollbar')
        self.scrollbar.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)
        self.canvas.configure(yscrollcommand=self._scrolled)
        self.canvas.bind('<Configure>', self._resize)
        self.bind('<MouseWheel>', self._wheel)
        self.bind('<Escape>', lambda e: self.destroy())
        self.protocol('WM_DELETE_WINDOW', self.destroy)
        self._update_count()
        self.update_idletasks()
        x = max(0, min(parent.winfo_rootx() + (parent.winfo_width()-self.winfo_width())//2,
                       self.winfo_screenwidth()-self.winfo_width()-16))
        y = max(0, min(parent.winfo_rooty() + (parent.winfo_height()-self.winfo_height())//2,
                       self.winfo_screenheight()-self.winfo_height()-60))
        self.geometry(f'+{x}+{y}')
        self.grab_set()
        self.focus_set()

    @staticmethod
    def _button(parent, text, command):
        return SmoothButton(parent, text=text, command=command, bg='#1C2C3B', fg='#E1EBF4',
                            activebackground='#29465B', activeforeground='#FFFFFF',
                            font=('Yu Gothic UI', 10), relief='flat', bd=0, padx=12, pady=7, cursor='hand2')

    def _texts(self, index):
        _, old, new = self.rows[index]
        if self.comment.get():
            clean = re.sub(r'\s+', ' ', re.sub(r'\([^)]*\)', '', old)).strip()
            if clean:
                new += f' (元: {clean})'
        return old, new

    def _wrap(self, text, columns):
        # Measure non-ASCII comments too so they do not cross the column boundary.
        lines = []
        for line in text.expandtabs(4).split('\n'):
            current = ''
            width = 0
            for char in line:
                if char not in self._char_sizes:
                    self._char_sizes[char] = self.code_font.measure(char)
                size = self._char_sizes[char]
                if current and width+size > columns*self.char_width:
                    lines.append(current)
                    current, width = '', 0
                current += char
                width += size
            lines.append(current)
        return lines

    def _resize(self, event):
        if event.width != self._width:
            self._width = event.width
            self._layout()
        self._schedule()

    def _layout(self):
        self._columns = max(12, int((self._width/2-36)/self.char_width))
        offset = 0
        self._offsets = []
        self._layouts = []
        for i in range(len(self.rows)):
            old, new = self._texts(i)
            left, right = self._wrap(old, self._columns), self._wrap(new, self._columns)
            height = 78 + max(len(left), len(right))*self.line_height
            self._offsets.append(offset)
            self._layouts.append((left, right, height))
            offset += height + 12
        self.canvas.configure(scrollregion=(0, 0, self._width, max(0, offset)))

    def _comment_changed(self):
        self._cache.clear()
        self._layout()
        self._schedule()

    def _scroll(self, *args):
        self.canvas.yview(*args)
        self._schedule()

    def _scrolled(self, first, last):
        self.scrollbar.set(first, last)
        self._schedule()

    def _wheel(self, event):
        if event.delta:
            self.canvas.yview_scroll(-int(math.copysign(max(1, abs(event.delta)//120), event.delta)), 'units')
            self._schedule()
        return 'break'

    def _schedule(self):
        if self._pending is None:
            self._pending = self.after_idle(self._paint)

    def _update_count(self):
        n = len(self.selected)
        self.count.set(f'変更対象 {len(self.rows)}{self.unit}　／　選択 {n}{self.unit}')
        self.execute.configure(text=f'選択した{n}{self.unit}を適用', state='normal' if n else 'disabled')

    def select_all(self, enabled):
        self.selected = set(self.indices) if enabled else set()
        self._update_count()
        self._schedule()

    def toggle(self, index):
        key = self.indices[index]
        if key in self.selected:
            self.selected.remove(key)
        else:
            self.selected.add(key)
        self._update_count()
        self._schedule()

    def _draw_code(self, x, y, lines, original, changed, enabled, before=False):
        expanded = original.expandtabs(4)
        # Render highlighted character runs, rather than an entire coloured line.
        pos = 0
        for line in lines:
            start = 0
            while start < len(line):
                highlighted = pos+start in changed
                end = start+1
                while end < len(line) and ((pos+end in changed) == highlighted):
                    end += 1
                chunk = line[start:end]
                xx = x + self.code_font.measure(line[:start])
                if highlighted:
                    self.canvas.create_rectangle(xx, y, xx+self.code_font.measure(chunk), y+self.line_height-2,
                                                 fill=('#38282E' if before else '#173F3D') if enabled else '#26323E', outline='')
                self.canvas.create_text(xx, y, text=chunk, anchor='nw', font=self.code_font,
                                        fill=(('#EDB7BC' if before else '#71EDCF') if highlighted else '#DFE8F0') if enabled else '#93A3B3')
                start = end
            pos += len(line)
            if pos < len(expanded) and expanded[pos] == '\n':
                pos += 1
            y += self.line_height

    def _paint(self):
        self._pending = None
        if not self.winfo_exists() or not self._layouts:
            return
        focused = self.focus_get()
        focus_index = getattr(focused, '_preview_index', None)
        for widget in self._check_widgets:
            widget.destroy()
        self._check_widgets.clear()
        self.canvas.delete('all')
        top = self.canvas.canvasy(0)
        bottom = top+self.canvas.winfo_height()
        first = max(0, bisect.bisect_right(self._offsets, top)-1)
        width = self._width-5
        for i in range(first, len(self.rows)):
            y = self._offsets[i]
            if y > bottom:
                break
            left, right, height = self._layouts[i]
            enabled = self.indices[i] in self.selected
            radius = 8
            self.canvas.create_polygon(1+radius,y,width-radius,y,width,y,width,y+radius,
                                       width,y+height-radius,width,y+height,width-radius,y+height,
                                       1+radius,y+height,1,y+height,1,y+height-radius,1,y+radius,1,y,
                                       smooth=True,fill='#101820',outline='#344759')
            self.canvas.create_rectangle(2,y+1,width-1,y+36,fill='#1B2937',outline='')
            check = tk.Checkbutton(self.canvas, text=f'{self.rows[i][0]}行',
                                   bg='#1B2937', fg='#D5E3EE', activebackground='#1B2937',
                                   activeforeground='#FFFFFF', selectcolor='#080C11',
                                   font=('Yu Gothic UI',10), command=lambda index=i:self.toggle(index))
            check.var = tk.BooleanVar(value=enabled)
            check._preview_index = i
            check.configure(variable=check.var)
            self._check_widgets.append(check)
            self.canvas.create_window(12,y+18,window=check,anchor='w')
            if focus_index == i:
                check.focus_set()
            self.canvas.create_text(width-14,y+18,text='変更する' if enabled else '変更しない',anchor='e',
                                    fill='#68DACE' if enabled else '#A6B5C3',font=('Yu Gothic UI',9))
            self.canvas.create_line(width/2,y+36,width/2,y+height,fill='#2B3846')
            for x,label in ((15,'変更前'),(width/2+15,'変更後' if enabled else '変更後の候補')):
                self.canvas.create_text(x,y+47,text=label,anchor='nw',fill='#A7B7C5',font=('Yu Gothic UI',9))
            old,new = self._texts(i)
            if i not in self._cache:
                old_pos,new_pos=set(),set()
                for op,a,b,c,d in difflib.SequenceMatcher(None,old.expandtabs(4),new.expandtabs(4)).get_opcodes():
                    if op != 'equal':
                        old_pos.update(range(a,b));new_pos.update(range(c,d))
                self._cache[i]=(old_pos,new_pos)
            self._draw_code(15,y+69,left,old,self._cache[i][0],enabled,before=True)
            self._draw_code(width/2+15,y+69,right,new,self._cache[i][1],enabled)

    def _execute(self):
        if self.selected:
            selected, comment = set(self.selected), self.comment.get()
            self.destroy()
            self.on_confirm(selected, comment)

    def destroy(self):
        if self._pending is not None:
            self.after_cancel(self._pending)
            self._pending = None
        super().destroy()
