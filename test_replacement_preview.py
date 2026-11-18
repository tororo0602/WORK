import tkinter as tk
import time
from pathlib import Path
from PIL import ImageGrab
from replacement_preview import ReplacementPreview

root=tk.Tk()
root.geometry('1200x850+0+0')
root.update()
errors=[]
root.report_callback_exception=lambda *args:errors.append(args)
rows=[(42+i*3,'G00 X50. Z5.\nG01 Z-10. F100.','G00 X50. Z5.\nG01 Z-12. F80.') for i in range(1500)]
confirmed=[]
try:
    start=time.monotonic()
    dlg=ReplacementPreview(root,'複数行置換 [N20ブロック]（1500箇所）',rows,list(range(1500)),lambda chosen,comment:confirmed.append((chosen,comment)),True)
    root.update()
    assert len(dlg._check_widgets)<15
    dlg.toggle(1);root.update()
    assert 1 not in dlg.selected
    dlg.canvas.yview_moveto(1);root.update()
    assert dlg.canvas.yview()[1]>.99
    assert len(dlg._check_widgets)<15
    dlg.canvas.yview_moveto(0);root.update()
    assert 1 not in dlg.selected
    dlg.select_all(False);root.update()
    assert str(dlg.execute.cget('state'))=='disabled'
    dlg.toggle(0);root.update()
    assert dlg.selected=={0}
    ImageGrab.grab(bbox=(dlg.winfo_rootx(),dlg.winfo_rooty(),dlg.winfo_rootx()+dlg.winfo_width(),dlg.winfo_rooty()+dlg.winfo_height())).save(Path(__file__).with_name('replacement_preview.png'))
    dlg._execute();assert confirmed==[({0},False)]
    dlg=ReplacementPreview(root,'テンプレート変換',[(8,'G01 X10. Y20.','G01 X20. Y10.')],[7],lambda chosen,comment:confirmed.append((chosen,comment)))
    root.update();dlg.comment.set(True);dlg._comment_changed();root.update()
    assert '(元: G01 X10. Y20.)' in dlg._texts(0)[1]
    dlg._execute();assert confirmed[-1]==({7},True)
    dlg=ReplacementPreview(root,'通常の置換',[(1,'F100.','F80.')],[0],lambda *args:confirmed.append(args))
    root.update();dlg.destroy();assert len(confirmed)==2
    assert not errors,errors
    print(f'PASS: 1500 rows, virtual scroll, selection, disabled empty apply, callback keys, comments, cancel ({time.monotonic()-start:.2f}s)')
finally:
    root.destroy()
