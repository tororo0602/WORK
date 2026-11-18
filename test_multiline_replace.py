import tkinter as tk
from types import SimpleNamespace, MethodType
from unittest.mock import patch
import sys
sys.path.insert(0,'開発用/v39')
import nc_program_check_tool_v39 as m
root=tk.Tk();root.withdraw()
t=tk.Text(root,undo=True)
d=SimpleNamespace(target=t,app=SimpleNamespace(_update_linenumbers=lambda:None,_highlight_all=lambda:None),
 _status_var=tk.StringVar(),_search_var=tk.StringVar(value='X1\nY2'),_replace_var=tk.StringVar(value='Z3'),
 _regex_var=tk.StringVar(value='0'),_case_var=tk.StringVar(value='1'))
for name in ['_get_target_text','_get_n_block_text_range','_build_pattern','_atomic_replace','_replace_all_whole_text','replace_all']:
 setattr(d,name,MethodType(getattr(m.FindReplaceDialog,name),d))
d._clear_match_highlights=lambda:None
d._trigger_post_replace_check=lambda:None
d._get_range_spec=lambda:{}
selector=object.__new__(m.TplTargetSelector)
selector._var=tk.StringVar(value='指定したNブロック')
selector._n_var=tk.StringVar(value='N20')
d._simple_scope_selector=selector
captured={}
def preview(title,rows,ids,confirm,grouped=False):
 captured.update(title=title,rows=rows,ids=ids,confirm=confirm,grouped=grouped)
d._show_preview_dialog=preview
source='N10\nX1\nY2\nN20\nX1\nY2\nX1\nY2\nN30\nX1\nY2'
def reset():
 t.delete('1.0','end');t.insert('1.0',source);t.edit_reset();captured.clear()
try:
 with patch.object(m.messagebox,'showinfo'),patch.object(m.messagebox,'showwarning'),patch.object(m.messagebox,'showerror'):
  reset();d.replace_all()
  assert captured['grouped'] and len(captured['rows'])==2
  assert t.get('1.0','end-1c')==source
  captured['confirm']({1})
  assert t.get('1.0','end-1c')=='N10\nX1\nY2\nN20\nX1\nY2\nZ3\nN30\nX1\nY2'
  t.edit_undo();assert t.get('1.0','end-1c')==source
  reset();selector._var.set('選択範囲');t.tag_add('sel','5.0','7.0');d.replace_all()
  assert len(captured['rows'])==1;captured['confirm']({0})
  assert t.get('1.0','end-1c')==source.replace('N20\nX1\nY2','N20\nZ3',1)
  reset();selector._var.set('指定したNブロック');selector._n_var.set('N30');d.replace_all();captured['confirm']({0})
  assert t.get('1.0','end-1c')==source[:-5]+'Z3'
  reset();selector._var.set('全体');d._regex_var.set('1');d._search_var.set(r'X(1)\nY(2)');d._replace_var.set(r'Z\2\nA\1')
  d.replace_all();assert len(captured['rows'])==4;captured['confirm']({0})
  assert t.get('1.0','end-1c')==source.replace('X1\nY2','Z2\nA1',1)
  reset();d.replace_all();t.insert('end-1c','\nM30');current=t.get('1.0','end-1c');captured['confirm']({0});assert t.get('1.0','end-1c')==current
 print('PASS: N scope, selective match, cancel, undo, selection, EOF, regex/backrefs, stale snapshot')
finally:root.destroy()

