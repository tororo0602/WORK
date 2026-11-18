"""Draggable in-tool mascot with the approved two-item context menu."""
from pathlib import Path
import tkinter as tk
from mascot_base import load_frames, KEY


class ToolMascot(tk.Toplevel):
    def __init__(self, root, dock, on_visibility_changed=None):
        super().__init__(root)
        self.withdraw()
        self.host, self.dock = root, dock
        self.on_visibility_changed = on_visibility_changed
        self.overrideredirect(True)
        self.transient(root)
        self.configure(bg=KEY)
        self.attributes('-transparentcolor', KEY)
        self.frames, self.delay = load_frames(str(Path(__file__).with_name('mascot.webp')), '#131A22', .65)
        self.size = self.frames[0].width(), self.frames[0].height()
        self.i = 0
        self.playing = True
        self.hidden = False
        self.docked = True
        self.rx = self.ry = 0
        self.timer = None
        self.label = tk.Label(self, image=self.frames[0], bg=KEY, bd=0, cursor='fleur')
        self.label.pack()
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label='アニメ停止/再開', command=self.toggle)
        self.menu.add_command(label='隠す', command=self.hide)
        self.label.bind('<Button-3>', self.popup)
        self.label.bind('<ButtonPress-1>', self.press)
        self.label.bind('<B1-Motion>', self.drag)
        self.bindings=[]
        for widget, event, callback in [(root,'<Configure>',self.follow),(root,'<Map>',self.follow),
                                        (root,'<Unmap>',self.follow),(root,'<FocusIn>',self.focus),
                                        (dock,'<Configure>',self.follow),(dock,'<Map>',self.follow)]:
            self.bindings.append((widget,event,widget.bind(event,callback,add='+')))
        self.follow()

    def follow(self, event=None):
        if event is not None and event.widget not in (self.host,self.dock):
            return
        w,h=self.size
        if self.hidden or not self.host.winfo_viewable() or self.host.winfo_width()<w or self.host.winfo_height()<h:
            self.suspend()
            return
        if self.docked:
            if not self.dock.winfo_viewable() or self.dock.winfo_width()<w:
                self.suspend()
                return
            self.rx=self.dock.winfo_rootx()-self.host.winfo_rootx()+(self.dock.winfo_width()-w)//2
            self.ry=self.dock.winfo_rooty()-self.host.winfo_rooty()+max(0,self.dock.winfo_height()-h-10)
        self.rx=max(0,min(self.rx,self.host.winfo_width()-w))
        self.ry=max(0,min(self.ry,self.host.winfo_height()-h))
        self.geometry(f'+{self.host.winfo_rootx()+self.rx}+{self.host.winfo_rooty()+self.ry}')
        if not self.winfo_viewable():
            self.deiconify()
            self.lift(self.host)
        if self.playing and self.timer is None:
            self.timer=self.after(self.delay,self.tick)

    def suspend(self):
        if self.timer is not None:
            self.after_cancel(self.timer)
            self.timer=None
        self.withdraw()

    def focus(self,event):
        if event.widget is self.host:
            self.restore_layer()

    def restore_layer(self):
        self.follow()
        if not self.hidden and self.winfo_viewable():
            self.lift(self.host)

    def press(self,event):
        self.offset=event.x,event.y

    def drag(self,event):
        self.docked=False
        self.rx=event.x_root-self.offset[0]-self.host.winfo_rootx()
        self.ry=event.y_root-self.offset[1]-self.host.winfo_rooty()
        self.follow()

    def popup(self,event):
        try:
            self.menu.tk_popup(event.x_root,event.y_root)
        finally:
            self.menu.grab_release()

    def hide(self):
        self.hidden=True
        self.suspend()
        if self.on_visibility_changed:
            self.on_visibility_changed(False)

    def show(self):
        self.hidden=False
        self.follow()
        if self.on_visibility_changed:
            self.on_visibility_changed(True)

    def toggle(self):
        self.playing=not self.playing
        if self.timer is not None:
            self.after_cancel(self.timer)
            self.timer=None
        self.follow()

    def tick(self):
        self.timer=None
        if self.playing and self.winfo_viewable() and not self.hidden:
            self.i=(self.i+1)%len(self.frames)
            self.label.configure(image=self.frames[self.i])
            self.timer=self.after(self.delay,self.tick)

    def destroy(self):
        if self.timer is not None:
            self.after_cancel(self.timer)
            self.timer=None
        for widget,event,binding in self.bindings:
            if widget.winfo_exists():
                widget.unbind(event,binding)
        super().destroy()
