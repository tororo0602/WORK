"""NCコード支援の保存互換性・改行・削除・編集を一時データで確認。"""
import json
import tempfile
from pathlib import Path
from types import SimpleNamespace
import tkinter as tk
from unittest.mock import patch
import nc_program_check_tool_v39 as m

root = tk.Tk()
root.geometry('1420x860+0+0')
editor = tk.Text(root, undo=True)
editor.pack()
root.update()
app = SimpleNamespace(root=root, input_text=editor, status_var=tk.StringVar())
def insert(text):
    if editor.tag_ranges('sel'):
        editor.delete('sel.first', 'sel.last')
    editor.insert('insert', text)
app.insert_text_at_cursor = insert
errors = []
root.report_callback_exception = lambda *args: errors.append(args)
try:
    with tempfile.TemporaryDirectory() as temp, patch.object(m, 'get_nc_macros_path', return_value=str(Path(temp)/'codes.json')):
        old = {'o_number':'O7021','title':'既存マクロ','args':[{'letter':'A','desc':'距離'}], 'raw_text':'original'}
        m.save_nc_macro_defs([old])
        dialog = m.NcMacroAssistDialog(app)
        dialog._select_macro(m.load_nc_macro_defs()[0])
        dialog._arg_vars['A'].set('10')
        assert dialog._preview_var.get() == 'G65P7021A10.'
        call = dialog._preview_var.get()
        for before, after, expected in [('', '', call), ('N10\n', 'M30', 'N10\n'+call+'\nM30'),
                ('N10', '\nM30', 'N10\n'+call+'\nM30'), ('N10\n', '\nM30', 'N10\n'+call+'\nM30'),
                ('N10', 'M30', 'N10\n'+call+'\nM30'), ('N10\n', '', 'N10\n'+call)]:
            editor.delete('1.0', 'end'); editor.insert('1.0', before+after)
            editor.mark_set('insert', f'1.0+{len(before)}c')
            dialog._insert_call()
            assert editor.get('1.0','end-1c') == expected
        editor.delete('1.0','end'); editor.insert('1.0','N10\nOLD\nM30')
        editor.tag_add('sel','2.0','2.end'); editor.mark_set('insert','1.0')
        dialog._insert_call()
        assert editor.get('1.0','end-1c') == 'N10\n'+call+'\nM30'
        edit = m.MacroDefEditDialog(app, dialog._on_macro_def_saved)
        edit._category_var.set('固定サイクル'); edit._category_changed()
        edit._title_var.set('サンプルサイクル')
        code='G81 X10. Z-5. R2. F100.\n\nG80'
        edit._code_text.insert('1.0',code); edit._save()
        assert len(m.load_nc_macro_defs()) == 2
        saved = dialog._selected
        assert dialog._preview_var.get() == code
        edit = m.MacroDefEditDialog(app, dialog._on_macro_def_saved, saved, saved['o_number'])
        assert edit._category_var.get() == '固定サイクル'
        assert edit._code_text.get('1.0','end-1c') == code
        root.update()
        from PIL import ImageGrab
        edit.lift(); root.update()
        ImageGrab.grab(bbox=(edit.winfo_rootx(),edit.winfo_rooty(),edit.winfo_rootx()+edit.winfo_width(),edit.winfo_rooty()+edit.winfo_height())).save(Path(__file__).with_name('code_edit_preview.png'))
        edit.destroy()
        editor.delete('1.0','end'); dialog._insert_call()
        assert editor.get('1.0','end-1c') == code
        dialog.lift(); root.update()
        ImageGrab.grab(bbox=(dialog.winfo_rootx(),dialog.winfo_rooty(),dialog.winfo_rootx()+dialog.winfo_width(),dialog.winfo_rooty()+dialog.winfo_height())).save(Path(__file__).with_name('code_assist_preview.png'))
        with patch.object(m.messagebox,'askyesno',return_value=False):
            dialog._delete_selected()
        assert len(m.load_nc_macro_defs()) == 2
        with patch.object(m.messagebox,'askyesno',return_value=True):
            dialog._delete_selected()
        assert len(m.load_nc_macro_defs()) == 1 and dialog._selected is None
        assert dialog._preview_var.get() == ''
        dialog._select_macro(m.load_nc_macro_defs()[0])
        with patch.object(m.messagebox,'askyesno',return_value=True):
            dialog._delete_selected()
        assert m.load_nc_macro_defs() == []
        assert not errors, errors
        dialog.destroy()
        print('PASS: legacy data, cycle save/edit, exact multiline content, insertion boundaries/selection, deletion/cancel/empty library')
finally:
    root.destroy()
