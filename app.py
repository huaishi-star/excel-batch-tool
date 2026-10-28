"""Lsy的Excel-Tool — 图形界面入口。

界面分四个标签页：
1. 批量重命名：按单元格组合规则复制重命名
2. 模板填充：按映射关系把源文件数据填入模板
3. 行/列删除：按条件批量删除行/列
4. 端口拆分：按端口号拆分、按流量排序、标红指定行并汇总

线程安全约定：工作线程只往 queue 里放消息，主线程用 after() 轮询并更新界面，
工作线程绝不直接触碰任何 tkinter 控件。
"""

from __future__ import annotations

import queue
import sys
import threading
import time
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, font as tkfont
from tkinter.scrolledtext import ScrolledText

import excel_core
import config_store

EXCEL_TYPES = [('Excel 文件', '*.xlsx *.xlsm *.xls')]

IS_MAC = sys.platform == 'darwin'
IS_WIN = sys.platform == 'win32'
# macOS 没有微软雅黑，使用系统中文字体；Windows 保持原字体
UI_FONT = 'PingFang SC' if IS_MAC else 'Microsoft YaHei UI'


# ---------------------------------------------------------------------------
# 全局异常兜底（--windowed 打包后无控制台，把未捕获异常写入 exe 旁日志）
# ---------------------------------------------------------------------------

def _log_dir() -> Path:
    """错误日志目录：macOS 写用户日志目录，Windows 仍写在 exe/脚本旁。"""
    if IS_MAC:
        return Path.home() / 'Library' / 'Logs' / 'Lsy的Excel-Tool'
    base = Path(sys.executable if getattr(sys, 'frozen', False) else __file__)
    return base.parent


def _write_exception_log():
    try:
        log_dir = _log_dir()
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / 'Lsy的Excel-Tool_错误日志.txt'
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(''.join(traceback.format_exception(*sys.exc_info())) + '\n')
    except Exception:
        pass


def _excepthook(exc_type, exc_value, exc_tb):
    _write_exception_log()
    sys.__excepthook__(exc_type, exc_value, exc_tb)


# ---------------------------------------------------------------------------
# 文件多选列表（两个标签页共用）
# ---------------------------------------------------------------------------

class FilePicker(ttk.LabelFrame):
    """文件 / 文件夹多选列表，支持选择 .xlsx/.xlsm/.xls、含子文件夹扫描、移除与清空。"""

    def __init__(self, master, text: str, recursive_var: tk.BooleanVar | None = None):
        super().__init__(master, text=text)
        self.files: list[str] = []
        self.recursive_var = recursive_var if recursive_var is not None else tk.BooleanVar(value=False)
        self._build()

    def _build(self):
        btn_row = ttk.Frame(self)
        btn_row.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Button(btn_row, text='选择文件…', command=self.pick_files).pack(side='left', padx=(0, 4))
        ttk.Button(btn_row, text='选择文件夹…', command=self.pick_folder).pack(side='left', padx=4)
        ttk.Checkbutton(btn_row, text='包含子文件夹', variable=self.recursive_var).pack(side='left', padx=4)
        ttk.Button(btn_row, text='移除选中', command=self.remove_selected).pack(side='left', padx=4)
        ttk.Button(btn_row, text='清空', command=self.clear).pack(side='left', padx=4)

        list_frame = ttk.Frame(self)
        list_frame.pack(fill='both', expand=True, padx=6, pady=6)
        self.listbox = tk.Listbox(list_frame, selectmode='extended', height=4, exportselection=False)
        sb = ttk.Scrollbar(list_frame, orient='vertical', command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.listbox.pack(side='left', fill='both', expand=True)

    # ---- 文件收集 ----
    def pick_files(self):
        paths = filedialog.askopenfilenames(parent=self, title='选择 Excel 文件', filetypes=EXCEL_TYPES)
        if paths:
            self.add_files(paths)

    def pick_folder(self):
        folder = filedialog.askdirectory(parent=self, title='选择文件夹')
        if not folder:
            return
        found = self._scan_folder(folder)
        if found:
            self.add_files(found)
        else:
            messagebox.showinfo('提示', '该文件夹下没有找到 .xlsx / .xlsm / .xls 文件', parent=self)

    def _scan_folder(self, folder: str) -> list[str]:
        folder = Path(folder)
        recursive = self.recursive_var.get()
        it = folder.rglob('*') if recursive else folder.glob('*')
        return [str(p) for p in it if p.is_file() and p.suffix.lower() in ('.xlsx', '.xlsm', '.xls')]

    def add_files(self, paths):
        seen = {p.lower() for p in self.files}
        for p in paths:
            key = str(p).lower()
            if key not in seen:
                self.files.append(str(p))
                seen.add(key)
        self._refresh()

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            del self.files[i]
        self._refresh()

    def clear(self):
        self.files.clear()
        self._refresh()

    def _refresh(self):
        self.listbox.delete(0, 'end')
        for f in self.files:
            self.listbox.insert('end', f)


# ---------------------------------------------------------------------------
# 可滚动容器：内容超高时出现垂直滚动条，不靠扩大窗口
# ---------------------------------------------------------------------------

class ScrollableFrame(ttk.Frame):
    """把内容放进 .inner，内容超出可视高度时自动出现滚动条。"""

    def __init__(self, master):
        super().__init__(master)
        self._canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self._vsb = ttk.Scrollbar(self, orient='vertical', command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vsb.set)
        self._vsb.pack(side='right', fill='y')
        self._canvas.pack(side='left', fill='both', expand=True)
        self.inner = ttk.Frame(self._canvas)
        self._window_id = self._canvas.create_window((0, 0), window=self.inner, anchor='nw')
        self.inner.bind('<Configure>', self._on_inner_configure)
        self._canvas.bind('<Configure>', self._on_canvas_configure)
        self._canvas.bind('<Enter>', self._bind_mousewheel)
        self._canvas.bind('<Leave>', self._unbind_mousewheel)

    def _on_inner_configure(self, _event):
        self._canvas.configure(scrollregion=self._canvas.bbox('all'))

    def _on_canvas_configure(self, event):
        self._canvas.itemconfigure(self._window_id, width=event.width)

    def _bind_mousewheel(self, _event):
        self._canvas.bind_all('<MouseWheel>', self._on_mousewheel)

    def _unbind_mousewheel(self, _event):
        self._canvas.unbind_all('<MouseWheel>')

    def _on_mousewheel(self, event):
        w = event.widget
        # 自身带滚动条的控件（文本框/列表）交给它们自己处理，避免双重滚动
        if w.winfo_class() in ('Text', 'Listbox', 'TCombobox', 'Treeview'):
            return
        if IS_MAC:
            # macOS 滚轮/触控板 delta 通常只有 ±1（Windows 是 ±120）
            delta = -event.delta
            if abs(delta) >= 1:
                step = int(delta)
            else:
                step = 1 if delta > 0 else (-1 if delta < 0 else 0)
        else:
            step = int(-event.delta / 120)
        if step:
            self._canvas.yview_scroll(step, 'units')


# ---------------------------------------------------------------------------
# 标签页基类：后台线程 + 队列轮询管道
# ---------------------------------------------------------------------------

class BaseTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self._queue = queue.Queue()
        self._cancel = threading.Event()
        self._worker: threading.Thread | None = None
        self.status_var = tk.StringVar(value='就绪')
        self.run_button: ttk.Button | None = None
        self.preview_button: ttk.Button | None = None
        self.cancel_button: ttk.Button | None = None
        self.log_text: ScrolledText | None = None

    # ---- 后台线程管道 ----
    def start_worker(self, target, args=()):
        self._queue = queue.Queue()
        self._cancel.clear()
        self.set_running(True)

        def _run():
            try:
                target(self._queue, self._cancel, *args)
            except Exception:
                err = traceback.format_exc()
                try:
                    self._queue.put(('log', '[失败] 内部错误：\n' + err))
                except Exception:
                    pass
                self._queue.put(('done', None))

        self._worker = threading.Thread(target=_run, daemon=True)
        self._worker.start()
        self.after(80, self._poll)

    def _poll(self):
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == 'log':
                    self._append_log(payload)
                elif kind == 'preview':
                    self._show_preview(payload)
                elif kind == 'progress':
                    self._show_progress(*payload)
                elif kind == 'error':
                    self._show_error(payload)
                elif kind == 'done':
                    self._on_done(payload)
        except queue.Empty:
            pass
        if self._worker is not None and self._worker.is_alive():
            self.after(80, self._poll)
        else:
            self._worker = None

    def _show_error(self, msg: str):
        messagebox.showerror('提示', msg, parent=self)

    def _on_done(self, summary):
        self.set_running(False)
        self.status_var.set('就绪')
        if summary:
            self._append_log(f'[完成] 成功 {summary["ok"]} 个，跳过 {summary["skip"]} 个，失败 {summary["fail"]} 个')

    def _show_progress(self, done: int, total: int):
        self.status_var.set(f'处理中：{done} / {total}')

    def _show_preview(self, lines):
        """由子类实现：把预览行填入预览控件。"""
        raise NotImplementedError

    # ---- 日志 ----
    def _clear_log(self):
        if self.log_text is not None:
            self.log_text.configure(state='normal')
            self.log_text.delete('1.0', 'end')
            self.log_text.configure(state='disabled')

    def _append_log(self, text: str):
        if self.log_text is None:
            return
        tag = None
        if text.startswith('[OK]'):
            tag = 'ok'
        elif text.startswith('[跳过]'):
            tag = 'skip'
        elif text.startswith('[失败]'):
            tag = 'fail'
        elif text.startswith('[取消]') or text.startswith('[完成]'):
            tag = 'cancel'
        self.log_text.configure(state='normal')
        self.log_text.insert('end', text + '\n', tag)
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def _build_log_area(self, parent):
        frame = ttk.LabelFrame(parent, text='运行日志')
        frame.pack(fill='x', padx=8, pady=(4, 8))
        self.log_text = ScrolledText(frame, height=7, state='disabled', wrap='word',
                                     font=(UI_FONT, 9))
        self.log_text.pack(fill='both', expand=True, padx=4, pady=4)
        self.log_text.tag_configure('ok', foreground='#1a7f37')
        self.log_text.tag_configure('skip', foreground='#9a6700')
        self.log_text.tag_configure('fail', foreground='#cf222e')
        self.log_text.tag_configure('cancel', foreground='#8250df')

    # ---- 按钮状态 ----
    def set_running(self, running: bool):
        state = 'disabled' if running else 'normal'
        for b in (self.run_button, self.preview_button):
            if b is not None:
                b.configure(state=state)
        if self.cancel_button is not None:
            self.cancel_button.configure(state='normal' if running else 'disabled')

    def _cancel(self):
        self._cancel.set()
        self.status_var.set('正在取消…')


# ---------------------------------------------------------------------------
# Tab 1：批量重命名
# ---------------------------------------------------------------------------

class RenameTab(BaseTab):
    def __init__(self, master):
        super().__init__(master)
        self._rule_items: list = []  # 规则元素：['cell',坐标] / ['text',文字] / ['filename','']
        self._build()

    def _build(self):
        body = ScrollableFrame(self)
        body.pack(fill='both', expand=True)
        # ① 选择文件
        picker = FilePicker(body.inner, text='① 选择要处理的文件（.xlsx / .xlsm / .xls）')
        picker.pack(fill='x', padx=8, pady=(8, 4))
        self.picker = picker

        # ② 重命名规则
        rule = ttk.LabelFrame(body.inner, text='② 重命名规则（单元格 / 手打文字 / 原文件名，按顺序拼接为文件名）')
        rule.pack(fill='x', padx=8, pady=4)

        mid = ttk.Frame(rule)
        mid.pack(fill='x', padx=6, pady=(6, 2))

        list_frame = ttk.Frame(mid)
        list_frame.pack(side='left', fill='x', expand=True)
        self.rule_list = tk.Listbox(list_frame, height=4, exportselection=False)
        sb = ttk.Scrollbar(list_frame, orient='vertical', command=self.rule_list.yview)
        self.rule_list.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.rule_list.pack(side='left', fill='both', expand=True)

        btn_col = ttk.Frame(mid)
        btn_col.pack(side='left', padx=(10, 0))

        row_cell = ttk.Frame(btn_col)
        row_cell.pack(fill='x')
        ttk.Label(row_cell, text='单元格：').pack(side='left')
        self.coord_entry = ttk.Entry(row_cell, width=10)
        self.coord_entry.pack(side='left', fill='x', expand=True)
        self.coord_entry.bind('<Return>', lambda e: self._add_coord())
        ttk.Button(row_cell, text='添加', width=5, command=self._add_coord).pack(side='left', padx=(4, 0))

        row_text = ttk.Frame(btn_col)
        row_text.pack(fill='x', pady=(2, 0))
        ttk.Label(row_text, text='文字：').pack(side='left')
        self.text_entry = ttk.Entry(row_text, width=10)
        self.text_entry.pack(side='left', fill='x', expand=True)
        self.text_entry.bind('<Return>', lambda e: self._add_text())
        ttk.Button(row_text, text='添加', width=5, command=self._add_text).pack(side='left', padx=(4, 0))

        row_fn = ttk.Frame(btn_col)
        row_fn.pack(fill='x', pady=(2, 0))
        ttk.Button(row_fn, text='添加原文件名', width=14, command=self._add_filename).pack(side='left')
        ttk.Button(row_fn, text='删除', width=5, command=self._del_rule).pack(side='left', padx=(4, 0))

        row_move = ttk.Frame(btn_col)
        row_move.pack(fill='x', pady=(2, 0))
        ttk.Button(row_move, text='上移', width=9, command=lambda: self._move_rule(-1)).pack(side='left')
        ttk.Button(row_move, text='下移', width=9, command=lambda: self._move_rule(1)).pack(side='left', padx=(4, 0))

        hint_row = ttk.Frame(rule)
        hint_row.pack(fill='x', padx=6, pady=(0, 2))
        ttk.Label(hint_row, text='提示：顺序即拼接顺序；分隔符只插在相邻的单元格/原文件名之间；文字原样保留。',
                  foreground='#57606a').pack(anchor='w')

        sep_row = ttk.Frame(rule)
        sep_row.pack(fill='x', padx=6, pady=(2, 2))
        ttk.Label(sep_row, text='分隔符：').pack(side='left')
        self.sep_entry = ttk.Entry(sep_row, width=8)
        self.sep_entry.pack(side='left')
        ttk.Label(sep_row, text='（留空表示无分隔符，可填 - _ 空格 等）',
                  foreground='#57606a').pack(side='left', padx=8)

        # 规则保存/载入
        profile_row = ttk.Frame(rule)
        profile_row.pack(fill='x', padx=6, pady=(2, 6))
        ttk.Button(profile_row, text='保存当前规则…', command=self._save_rule).pack(side='left')
        ttk.Label(profile_row, text='  已保存规则：').pack(side='left', padx=(8, 2))
        self.rule_cb = ttk.Combobox(profile_row, state='readonly', width=18)
        self.rule_cb.pack(side='left')
        ttk.Button(profile_row, text='载入', command=self._load_rule).pack(side='left', padx=2)
        ttk.Button(profile_row, text='删除', command=self._delete_rule).pack(side='left', padx=2)
        self._refresh_rule_combobox()

        # ③ 预览
        prev = ttk.LabelFrame(body.inner, text='③ 预览（先预览确认无误，再执行）')
        prev.pack(fill='x', padx=8, pady=4)
        pframe = ttk.Frame(prev)
        pframe.pack(fill='x', padx=6, pady=(6, 2))
        self.preview_button = ttk.Button(pframe, text='预览', command=self._do_preview)
        self.preview_button.pack(side='left')
        ttk.Label(pframe, textvariable=self.status_var, foreground='#57606a').pack(side='left', padx=8)

        pbox = ttk.Frame(prev)
        pbox.pack(fill='both', expand=True, padx=6, pady=(2, 6))
        self.preview_text = tk.Text(pbox, height=4, state='disabled', wrap='none',
                                    font=(UI_FONT, 9))
        hsb = ttk.Scrollbar(pbox, orient='horizontal', command=self.preview_text.xview)
        self.preview_text.configure(xscrollcommand=hsb.set)
        hsb.pack(side='bottom', fill='x')
        self.preview_text.pack(side='left', fill='both', expand=True)

        # ④ 输出与执行
        out = ttk.LabelFrame(body.inner, text='④ 输出文件夹与执行')
        out.pack(fill='x', padx=8, pady=4)
        orow = ttk.Frame(out)
        orow.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(orow, text='输出文件夹：').pack(side='left')
        self.out_var = tk.StringVar()
        ttk.Entry(orow, textvariable=self.out_var).pack(side='left', fill='x', expand=True)
        ttk.Button(orow, text='浏览…', command=self._pick_outdir).pack(side='left', padx=4)

        brow = ttk.Frame(out)
        brow.pack(fill='x', padx=6, pady=(2, 6))
        self.run_button = ttk.Button(brow, text='开始重命名', command=self._run)
        self.run_button.pack(side='left')
        self.cancel_button = ttk.Button(brow, text='取消', command=self._cancel, state='disabled')
        self.cancel_button.pack(side='left', padx=4)

        self._build_log_area(self)

    # ---- 规则编辑 ----
    def _add_coord(self):
        c = self.coord_entry.get().strip().upper()
        if not c:
            return
        if not excel_core.CELL_RE.match(c):
            messagebox.showerror('格式错误', f'单元格应为 A1、B3 这样的坐标：{c!r}', parent=self)
            return
        if any(k == 'cell' and v == c for k, v in self._rule_items):
            messagebox.showinfo('提示', f'单元格 {c} 已在列表中', parent=self)
            return
        self._rule_items.append(['cell', c])
        self.coord_entry.delete(0, 'end')
        self._refresh_rule_list()

    def _add_text(self):
        t = self.text_entry.get()
        if not t.strip():
            messagebox.showinfo('提示', '请先输入要添加的文字', parent=self)
            return
        self._rule_items.append(['text', t])
        self.text_entry.delete(0, 'end')
        self._refresh_rule_list()

    def _add_filename(self):
        if any(k == 'filename' for k, _ in self._rule_items):
            messagebox.showinfo('提示', '「原文件名」已在列表中', parent=self)
            return
        self._rule_items.append(['filename', ''])
        self._refresh_rule_list()

    def _del_rule(self):
        for i in reversed(self.rule_list.curselection()):
            del self._rule_items[i]
        self._refresh_rule_list()

    def _move_rule(self, delta: int):
        sel = self.rule_list.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if j < 0 or j >= len(self._rule_items):
            return
        self._rule_items[i], self._rule_items[j] = self._rule_items[j], self._rule_items[i]
        self._refresh_rule_list()
        self.rule_list.selection_clear(0, 'end')
        self.rule_list.selection_set(j)
        self.rule_list.activate(j)

    def _refresh_rule_list(self):
        self.rule_list.delete(0, 'end')
        for kind, value in self._rule_items:
            if kind == 'cell':
                self.rule_list.insert('end', value)
            elif kind == 'text':
                self.rule_list.insert('end', f'文字：{value}')
            else:
                self.rule_list.insert('end', '原文件名')
        n = self.rule_list.size()
        if n:
            self.rule_list.selection_clear(0, 'end')
            self.rule_list.selection_set(n - 1)

    # ---- 规则保存/载入 ----
    def _save_rule(self):
        rule = [tuple(item) for item in self._rule_items]
        if not rule:
            messagebox.showwarning('提示', '请先添加规则（单元格 / 文字 / 原文件名）再保存', parent=self)
            return
        name = simpledialog.askstring('保存规则', '请输入规则名称（留空则用时间命名）：', parent=self)
        if name is None:
            return
        name = name.strip() or time.strftime('%m-%d %H:%M')
        if name in config_store.list_profiles('rename'):
            if not messagebox.askyesno('提示', f'规则「{name}」已存在，是否覆盖？', parent=self):
                return
        config_store.save_profile('rename', name, {
            'rule': [list(item) for item in rule],
            'separator': self.sep_entry.get(),
        })
        self._refresh_rule_combobox(name)
        self._append_log(f'[提示] 规则「{name}」已保存')

    def _load_rule(self):
        name = self.rule_cb.get()
        if not name:
            messagebox.showinfo('提示', '请先在列表中选择要载入的规则', parent=self)
            return
        prof = config_store.get_profile('rename', name)
        if not prof:
            messagebox.showwarning('提示', f'规则「{name}」不存在', parent=self)
            return
        self._rule_items = []
        for kind, value in prof.get('rule', []):
            if kind == 'cell' and excel_core.CELL_RE.match(str(value)):
                self._rule_items.append(['cell', str(value).upper()])
            elif kind == 'text':
                self._rule_items.append(['text', str(value)])
            elif kind == 'filename':
                self._rule_items.append(['filename', ''])
        self._refresh_rule_list()
        self.sep_entry.delete(0, 'end')
        self.sep_entry.insert(0, prof.get('separator', ''))
        self._append_log(f'[提示] 已载入规则「{name}」')

    def _delete_rule(self):
        name = self.rule_cb.get()
        if not name:
            return
        if not messagebox.askyesno('删除规则', f'确定删除规则「{name}」吗？', parent=self):
            return
        config_store.delete_profile('rename', name)
        self._refresh_rule_combobox()
        self._append_log(f'[提示] 已删除规则「{name}」')

    def _refresh_rule_combobox(self, select=None):
        names = config_store.list_profiles('rename')
        self.rule_cb['values'] = names
        if select and select in names:
            self.rule_cb.set(select)
        elif self.rule_cb.get() not in names and names:
            self.rule_cb.set(names[0])
        elif not names:
            self.rule_cb.set('')

    # ---- 交互 ----
    def _pick_outdir(self):
        d = filedialog.askdirectory(parent=self, title='选择输出文件夹')
        if d:
            self.out_var.set(d)

    def _collect_rule(self):
        files = list(self.picker.files)
        rule = [tuple(item) for item in self._rule_items]
        sep = self.sep_entry.get()
        out_dir = self.out_var.get().strip()
        return files, rule, sep, out_dir

    def _validate(self, files, rule, out_dir) -> bool:
        if not files:
            messagebox.showwarning('提示', '请先选择要处理的文件', parent=self)
            return False
        if not rule:
            messagebox.showwarning('提示', '请至少添加一项规则（单元格 / 文字 / 原文件名）', parent=self)
            return False
        if not out_dir:
            messagebox.showwarning('提示', '请选择输出文件夹', parent=self)
            return False
        return True

    def _do_preview(self):
        files, rule, sep, out_dir = self._collect_rule()
        if not self._validate(files, rule, out_dir):
            return
        self._clear_log()
        self.start_worker(self._worker_preview, (files, rule, sep, out_dir))

    def _run(self):
        files, rule, sep, out_dir = self._collect_rule()
        if not self._validate(files, rule, out_dir):
            return
        self._clear_log()
        self.start_worker(self._worker_run, (files, rule, sep, out_dir))

    # ---- 工作线程 ----
    def _worker_preview(self, queue, cancel, files, rule, sep, out_dir):
        plan, errors = excel_core.plan_rename(files, rule, sep, out_dir)
        lines = [f'{item.src.name}  →  {item.name}' for item in plan]
        for f, msg in errors:
            lines.append(f'[错误] 无法读取 {f}：{msg}')
        queue.put(('preview', lines))
        queue.put(('done', None))

    def _worker_run(self, queue, cancel, files, rule, sep, out_dir):
        log = lambda m: queue.put(('log', m))
        plan, errors = excel_core.plan_rename(files, rule, sep, out_dir)
        for f, msg in errors:
            log(f'[失败] 无法读取 {f}：{msg}')
        if not plan:
            queue.put(('error', '没有可处理的文件，请检查单元格与文件内容'))
            queue.put(('done', None))
            return
        summary = excel_core.execute_rename(plan, log,
                                            lambda d, t: queue.put(('progress', (d, t))),
                                            cancel)
        queue.put(('done', summary))

    def _show_preview(self, lines):
        self.preview_text.configure(state='normal')
        self.preview_text.delete('1.0', 'end')
        self.preview_text.insert('end', '\n'.join(lines) + ('\n' if lines else ''))
        self.preview_text.configure(state='disabled')
        self.status_var.set(f'预览：共 {len(lines)} 项')


# ---------------------------------------------------------------------------
# Tab 2：模板填充
# ---------------------------------------------------------------------------

class FillTab(BaseTab):
    def __init__(self, master):
        super().__init__(master)
        self._mappings: list = []  # ['cell',src,tpl] 或 ['column',src_col,src_row,tpl_col,tpl_row]
        self._build()

    def _build(self):
        body = ScrollableFrame(self)
        body.pack(fill='both', expand=True)
        # ① 选择源文件 + 模板
        picker = FilePicker(body.inner, text='① 选择源文件（每个文件按第一个工作表读取数据）')
        picker.pack(fill='x', padx=8, pady=(8, 4))
        self.picker = picker

        tpl = ttk.LabelFrame(body.inner, text='② 选择模板文件')
        tpl.pack(fill='x', padx=8, pady=4)
        trow = ttk.Frame(tpl)
        trow.pack(fill='x', padx=6, pady=6)
        ttk.Label(trow, text='模板文件：').pack(side='left')
        self.template_var = tk.StringVar()
        ttk.Entry(trow, textvariable=self.template_var, state='readonly').pack(side='left', fill='x', expand=True)
        ttk.Button(trow, text='选择模板…', command=self._pick_template).pack(side='left', padx=4)

        # ③ 映射表
        map_frame = ttk.LabelFrame(body.inner, text='③ 映射表（单元格 → 单元格 或 整列 → 整列，均指第一个工作表）')
        map_frame.pack(fill='x', padx=8, pady=4)

        # 映射表格行高加高，减少拥挤
        _style = ttk.Style()
        _style.configure('Map.Treeview', rowheight=26)

        tree_frame = ttk.Frame(map_frame)
        tree_frame.pack(fill='x', padx=8, pady=(8, 4))
        cols = ('type', 'src', 'tpl')
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=4, style='Map.Treeview')
        self.tree.heading('type', text='类型')
        self.tree.heading('src', text='源')
        self.tree.heading('tpl', text='模板')
        self.tree.column('type', width=80, anchor='center')
        self.tree.column('src', width=170, anchor='center')
        self.tree.column('tpl', width=170, anchor='center')
        tsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=tsb.set)
        tsb.pack(side='right', fill='y')
        self.tree.pack(side='left', fill='both', expand=True)

        edit = ttk.Frame(map_frame)
        edit.pack(fill='x', padx=8, pady=(2, 2))

        # 第一行：类型 + 输入（间距宽松）
        row1 = ttk.Frame(edit)
        row1.pack(fill='x', pady=(6, 6))
        ttk.Label(row1, text='添加方式：').pack(side='left')
        self.map_type = ttk.Combobox(row1, state='readonly', width=15,
                                     values=['单元格 → 单元格', '整列 → 整列', '列求和 → 单元格'])
        self.map_type.set('单元格 → 单元格')
        self.map_type.bind('<<ComboboxSelected>>', lambda e: self._toggle_map_type())
        self.map_type.pack(side='left', padx=(6, 0))

        self.input_frame = ttk.Frame(row1)
        self.input_frame.pack(side='left', padx=(18, 0))

        # 单元格模式
        self.cell_frame = ttk.Frame(self.input_frame)
        self.cell_frame.pack(side='left')
        ttk.Label(self.cell_frame, text='源单元格：').pack(side='left')
        self.src_entry = ttk.Entry(self.cell_frame, width=8)
        self.src_entry.pack(side='left', padx=(2, 0))
        ttk.Label(self.cell_frame, text=' 模板单元格：').pack(side='left', padx=(8, 0))
        self.tpl_entry = ttk.Entry(self.cell_frame, width=8)
        self.tpl_entry.pack(side='left', padx=(2, 0))

        # 整列模式
        self.col_frame = ttk.Frame(self.input_frame)
        ttk.Label(self.col_frame, text='源列：').pack(side='left')
        self.src_col_entry = ttk.Entry(self.col_frame, width=5)
        self.src_col_entry.pack(side='left', padx=(2, 0))
        ttk.Label(self.col_frame, text='从第').pack(side='left', padx=(4, 0))
        self.src_row_entry = ttk.Entry(self.col_frame, width=4)
        self.src_row_entry.pack(side='left', padx=(2, 0))
        ttk.Label(self.col_frame, text='行  → 模板列：').pack(side='left', padx=(6, 0))
        self.tpl_col_entry = ttk.Entry(self.col_frame, width=5)
        self.tpl_col_entry.pack(side='left', padx=(2, 0))
        ttk.Label(self.col_frame, text='从第').pack(side='left', padx=(4, 0))
        self.tpl_row_entry = ttk.Entry(self.col_frame, width=4)
        self.tpl_row_entry.pack(side='left', padx=(2, 0))
        ttk.Label(self.col_frame, text='行').pack(side='left', padx=(2, 0))

        # 列求和模式
        self.sum_frame = ttk.Frame(self.input_frame)
        ttk.Label(self.sum_frame, text='源列：').pack(side='left')
        self.sum_col_entry = ttk.Entry(self.sum_frame, width=5)
        self.sum_col_entry.pack(side='left', padx=(2, 0))
        ttk.Label(self.sum_frame, text='从第').pack(side='left', padx=(4, 0))
        self.sum_row_entry = ttk.Entry(self.sum_frame, width=4)
        self.sum_row_entry.pack(side='left', padx=(2, 0))
        ttk.Label(self.sum_frame, text='行  → 模板单元格：').pack(side='left', padx=(6, 0))
        self.sum_tpl_entry = ttk.Entry(self.sum_frame, width=8)
        self.sum_tpl_entry.pack(side='left', padx=(2, 0))

        # 第二行：操作按钮（独占一行，间距宽松）
        row2 = ttk.Frame(edit)
        row2.pack(fill='x', pady=(6, 4))
        ttk.Button(row2, text='添加', width=6, command=self._add_map).pack(side='left', padx=(0, 6))
        ttk.Button(row2, text='删除', width=6, command=self._del_map).pack(side='left', padx=6)
        ttk.Button(row2, text='上移', width=6, command=lambda: self._move_map(-1)).pack(side='left', padx=6)
        ttk.Button(row2, text='下移', width=6, command=lambda: self._move_map(1)).pack(side='left', padx=6)

        # 第三行：说明（单独一行）
        row3 = ttk.Frame(edit)
        row3.pack(fill='x', pady=(2, 6))
        ttk.Label(row3, text='说明：整列 = 源列从起始行向下读到空单元格为止；目标已有内容时插入空白行，不覆盖。',
                  foreground='#57606a').pack(anchor='w')

        # 映射保存/载入
        profile_row = ttk.Frame(map_frame)
        profile_row.pack(fill='x', padx=8, pady=(4, 8))
        ttk.Button(profile_row, text='保存当前映射…', command=self._save_mapping).pack(side='left')
        ttk.Label(profile_row, text='  已保存映射：').pack(side='left', padx=(8, 2))
        self.map_cb = ttk.Combobox(profile_row, state='readonly', width=18)
        self.map_cb.pack(side='left')
        ttk.Button(profile_row, text='载入', command=self._load_mapping).pack(side='left', padx=2)
        ttk.Button(profile_row, text='删除', command=self._delete_mapping).pack(side='left', padx=2)
        self._refresh_map_combobox()

        # ④ 预览
        prev = ttk.LabelFrame(body.inner, text='④ 预览（输出文件名 = 模板名 + 源文件名）')
        prev.pack(fill='x', padx=8, pady=4)
        pframe = ttk.Frame(prev)
        pframe.pack(fill='x', padx=6, pady=(6, 2))
        self.preview_button = ttk.Button(pframe, text='预览', command=self._do_preview)
        self.preview_button.pack(side='left')
        ttk.Label(pframe, textvariable=self.status_var, foreground='#57606a').pack(side='left', padx=8)

        pbox = ttk.Frame(prev)
        pbox.pack(fill='both', expand=True, padx=6, pady=(2, 6))
        self.preview_text = tk.Text(pbox, height=3, state='disabled', wrap='none',
                                    font=(UI_FONT, 9))
        hsb = ttk.Scrollbar(pbox, orient='horizontal', command=self.preview_text.xview)
        self.preview_text.configure(xscrollcommand=hsb.set)
        hsb.pack(side='bottom', fill='x')
        self.preview_text.pack(side='left', fill='both', expand=True)

        # ⑤ 输出与执行
        out = ttk.LabelFrame(body.inner, text='⑤ 输出文件夹与执行')
        out.pack(fill='x', padx=8, pady=4)
        orow = ttk.Frame(out)
        orow.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(orow, text='输出文件夹：').pack(side='left')
        self.out_var = tk.StringVar()
        ttk.Entry(orow, textvariable=self.out_var).pack(side='left', fill='x', expand=True)
        ttk.Button(orow, text='浏览…', command=self._pick_outdir).pack(side='left', padx=4)

        brow = ttk.Frame(out)
        brow.pack(fill='x', padx=6, pady=(2, 6))
        self.run_button = ttk.Button(brow, text='开始填充', command=self._run)
        self.run_button.pack(side='left')
        self.cancel_button = ttk.Button(brow, text='取消', command=self._cancel, state='disabled')
        self.cancel_button.pack(side='left', padx=4)

        self._build_log_area(self)

    # ---- 映射编辑 ----
    def _pick_template(self):
        p = filedialog.askopenfilename(parent=self, title='选择模板文件', filetypes=EXCEL_TYPES)
        if p:
            self.template_var.set(p)

    def _toggle_map_type(self):
        kind = self.map_type.get()
        # 先收起全部输入框
        for f in (self.cell_frame, self.col_frame, self.sum_frame):
            f.pack_forget()
        if kind.startswith('单元格'):
            self.cell_frame.pack(side='left')
        elif kind.startswith('整列'):
            self.col_frame.pack(side='left')
        else:
            self.sum_frame.pack(side='left')

    def _clear_map_inputs(self):
        for e in (self.src_entry, self.tpl_entry, self.src_col_entry, self.src_row_entry,
                  self.tpl_col_entry, self.tpl_row_entry, self.sum_col_entry,
                  self.sum_row_entry, self.sum_tpl_entry):
            e.delete(0, 'end')

    def _add_map(self):
        kind = self.map_type.get()
        if kind.startswith('单元格'):
            src = self.src_entry.get().strip().upper()
            tpl = self.tpl_entry.get().strip().upper()
            if not (src and tpl):
                messagebox.showinfo('提示', '请填写源单元格与模板单元格', parent=self)
                return
            for name, c in (('源单元格', src), ('模板单元格', tpl)):
                if not excel_core.CELL_RE.match(c):
                    messagebox.showerror('格式错误', f'{name}应为 A1、B3 这样的坐标：{c!r}', parent=self)
                    return
            m = ['cell', src, tpl]
        elif kind.startswith('整列'):
            src_col = self.src_col_entry.get().strip().upper()
            tpl_col = self.tpl_col_entry.get().strip().upper()
            src_row = self.src_row_entry.get().strip()
            tpl_row = self.tpl_row_entry.get().strip()
            if not (src_col and tpl_col and src_row and tpl_row):
                messagebox.showinfo('提示', '请填写源列/模板列及各自的起始行', parent=self)
                return
            if not excel_core.COL_RE.match(src_col) or not excel_core.COL_RE.match(tpl_col):
                messagebox.showerror('格式错误', '列应为 A、B、AA 这样的字母', parent=self)
                return
            try:
                src_row = int(src_row)
                tpl_row = int(tpl_row)
            except ValueError:
                messagebox.showerror('格式错误', '起始行应为正整数', parent=self)
                return
            if src_row < 1 or tpl_row < 1:
                messagebox.showerror('格式错误', '起始行应为正整数', parent=self)
                return
            m = ['column', src_col, src_row, tpl_col, tpl_row]
        else:  # 列求和
            src_col = self.sum_col_entry.get().strip().upper()
            src_row = self.sum_row_entry.get().strip()
            tpl = self.sum_tpl_entry.get().strip().upper()
            if not (src_col and src_row and tpl):
                messagebox.showinfo('提示', '请填写源列、起始行与模板单元格', parent=self)
                return
            if not excel_core.COL_RE.match(src_col):
                messagebox.showerror('格式错误', '列应为 A、B、AA 这样的字母', parent=self)
                return
            try:
                src_row = int(src_row)
            except ValueError:
                messagebox.showerror('格式错误', '起始行应为正整数', parent=self)
                return
            if src_row < 1:
                messagebox.showerror('格式错误', '起始行应为正整数', parent=self)
                return
            if not excel_core.CELL_RE.match(tpl):
                messagebox.showerror('格式错误', f'模板单元格应为 A1、B3 这样的坐标：{tpl!r}', parent=self)
                return
            m = ['sum', src_col, src_row, tpl]
        if m in self._mappings:
            messagebox.showinfo('提示', '该映射已存在', parent=self)
            return
        self._mappings.append(m)
        self._clear_map_inputs()
        self._refresh_map_tree()

    def _refresh_map_tree(self):
        self.tree.delete(*self.tree.get_children())
        for m in self._mappings:
            if m[0] == 'cell':
                self.tree.insert('', 'end', values=('单元格', m[1], m[2]))
            elif m[0] == 'column':
                self.tree.insert('', 'end', values=('整列', f'{m[1]}列 从第{m[2]}行', f'{m[3]}列 从第{m[4]}行'))
            else:
                self.tree.insert('', 'end', values=('求和', f'{m[1]}列 从第{m[2]}行', m[3]))

    def _del_map(self):
        for iid in reversed(self.tree.selection()):
            del self._mappings[self.tree.index(iid)]
        self._refresh_map_tree()

    def _move_map(self, delta: int):
        sel = self.tree.selection()
        if not sel:
            return
        i = self.tree.index(sel[0])
        j = i + delta
        if j < 0 or j >= len(self._mappings):
            return
        self._mappings[i], self._mappings[j] = self._mappings[j], self._mappings[i]
        self._refresh_map_tree()
        new_items = list(self.tree.get_children())
        self.tree.selection_set(new_items[j])
        self.tree.focus(new_items[j])

    def _mapping_rows(self) -> list:
        return [tuple(m) for m in self._mappings]

    # ---- 映射保存/载入 ----
    def _save_mapping(self):
        mapping = self._mapping_rows()
        if not mapping:
            messagebox.showwarning('提示', '请先添加映射（单元格或整列）再保存', parent=self)
            return
        name = simpledialog.askstring('保存映射', '请输入映射名称（留空则用时间命名）：', parent=self)
        if name is None:
            return
        name = name.strip() or time.strftime('%m-%d %H:%M')
        if name in config_store.list_profiles('fill'):
            if not messagebox.askyesno('提示', f'映射「{name}」已存在，是否覆盖？', parent=self):
                return
        config_store.save_profile('fill', name, {'mapping': [list(m) for m in mapping]})
        self._refresh_map_combobox(name)
        self._append_log(f'[提示] 映射「{name}」已保存')

    def _load_mapping(self):
        name = self.map_cb.get()
        if not name:
            messagebox.showinfo('提示', '请先在列表中选择要载入的映射', parent=self)
            return
        prof = config_store.get_profile('fill', name)
        if not prof:
            messagebox.showwarning('提示', f'映射「{name}」不存在', parent=self)
            return
        self._mappings = []
        for m in prof.get('mapping', []):
            m = list(m)
            if m and m[0] == 'cell' and len(m) == 3 \
                    and excel_core.CELL_RE.match(str(m[1])) and excel_core.CELL_RE.match(str(m[2])):
                self._mappings.append(['cell', str(m[1]).upper(), str(m[2]).upper()])
            elif m and m[0] == 'column' and len(m) == 5 \
                    and excel_core.COL_RE.match(str(m[1])) and excel_core.COL_RE.match(str(m[3])):
                try:
                    sr, tr = int(m[2]), int(m[4])
                except (ValueError, TypeError):
                    continue
                if sr >= 1 and tr >= 1:
                    self._mappings.append(['column', str(m[1]).upper(), sr, str(m[3]).upper(), tr])
            elif m and m[0] == 'sum' and len(m) == 4 \
                    and excel_core.COL_RE.match(str(m[1])) and excel_core.CELL_RE.match(str(m[3])):
                try:
                    sr = int(m[2])
                except (ValueError, TypeError):
                    continue
                if sr >= 1:
                    self._mappings.append(['sum', str(m[1]).upper(), sr, str(m[3]).upper()])
        self._refresh_map_tree()
        self._append_log(f'[提示] 已载入映射「{name}」')

    def _delete_mapping(self):
        name = self.map_cb.get()
        if not name:
            return
        if not messagebox.askyesno('删除映射', f'确定删除映射「{name}」吗？', parent=self):
            return
        config_store.delete_profile('fill', name)
        self._refresh_map_combobox()
        self._append_log(f'[提示] 已删除映射「{name}」')

    def _refresh_map_combobox(self, select=None):
        names = config_store.list_profiles('fill')
        self.map_cb['values'] = names
        if select and select in names:
            self.map_cb.set(select)
        elif self.map_cb.get() not in names and names:
            self.map_cb.set(names[0])
        elif not names:
            self.map_cb.set('')

    # ---- 交互 ----
    def _pick_outdir(self):
        d = filedialog.askdirectory(parent=self, title='选择输出文件夹')
        if d:
            self.out_var.set(d)

    def _collect(self):
        files = list(self.picker.files)
        template = self.template_var.get().strip()
        mapping = self._mapping_rows()
        out_dir = self.out_var.get().strip()
        return files, template, mapping, out_dir

    def _validate(self, files, template, mapping, out_dir) -> bool:
        if not files:
            messagebox.showwarning('提示', '请先选择源文件', parent=self)
            return False
        if not template:
            messagebox.showwarning('提示', '请选择模板文件', parent=self)
            return False
        if not Path(template).is_file():
            messagebox.showwarning('提示', '模板文件不存在', parent=self)
            return False
        if not mapping:
            messagebox.showwarning('提示', '请至少添加一条映射', parent=self)
            return False
        if not out_dir:
            messagebox.showwarning('提示', '请选择输出文件夹', parent=self)
            return False
        return True

    def _do_preview(self):
        files, template, mapping, out_dir = self._collect()
        if not self._validate(files, template, mapping, out_dir):
            return
        self._clear_log()
        self.start_worker(self._worker_preview, (files, template, mapping, out_dir))

    def _run(self):
        files, template, mapping, out_dir = self._collect()
        if not self._validate(files, template, mapping, out_dir):
            return
        self._clear_log()
        self.start_worker(self._worker_run, (files, template, mapping, out_dir))

    # ---- 工作线程 ----
    def _worker_preview(self, queue, cancel, files, template, mapping, out_dir):
        plan, errors = excel_core.plan_fill(files, template, mapping, out_dir)
        lines = [f'{item.src.name}  →  {item.dst.name}' for item in plan]
        for f, msg in errors:
            lines.append(f'[错误] {f}：{msg}')
        queue.put(('preview', lines))
        queue.put(('done', None))

    def _worker_run(self, queue, cancel, files, template, mapping, out_dir):
        log = lambda m: queue.put(('log', m))
        plan, errors = excel_core.plan_fill(files, template, mapping, out_dir)
        for f, msg in errors:
            log(f'[失败] {f}：{msg}')
        if not plan:
            queue.put(('error', '没有可处理的文件'))
            queue.put(('done', None))
            return
        summary = excel_core.execute_fill(template, plan, mapping, log,
                                          lambda d, t: queue.put(('progress', (d, t))),
                                          cancel)
        queue.put(('done', summary))

    def _show_preview(self, lines):
        self.preview_text.configure(state='normal')
        self.preview_text.delete('1.0', 'end')
        self.preview_text.insert('end', '\n'.join(lines) + ('\n' if lines else ''))
        self.preview_text.configure(state='disabled')
        self.status_var.set(f'预览：共 {len(lines)} 项')


# ---------------------------------------------------------------------------
# Tab 3：行/列删除
# ---------------------------------------------------------------------------

class DeleteTab(BaseTab):
    def __init__(self, master):
        super().__init__(master)
        self._terms: list = []  # [(值, 匹配方式), ...]
        self._build()

    def _build(self):
        body = ScrollableFrame(self)
        body.pack(fill='both', expand=True)
        # ① 选择文件
        picker = FilePicker(body.inner, text='① 选择要处理的文件（.xlsx / .xlsm / .xls）')
        picker.pack(fill='x', padx=8, pady=(8, 4))
        self.picker = picker

        # ② 删除设置
        cfg = ttk.LabelFrame(body.inner, text='② 删除设置（删除整行/整列中任意单元格匹配条件的行/列）')
        cfg.pack(fill='x', padx=8, pady=4)

        row1 = ttk.Frame(cfg)
        row1.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(row1, text='删除对象：').pack(side='left')
        self.mode_var = tk.StringVar(value='row')
        ttk.Radiobutton(row1, text='删除行', variable=self.mode_var, value='row').pack(side='left', padx=4)
        ttk.Radiobutton(row1, text='删除列', variable=self.mode_var, value='column').pack(side='left', padx=4)

        row2 = ttk.Frame(cfg)
        row2.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(row2, text='查找条件（可添加多条，任一匹配即删除）：').pack(side='left')
        self.case_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text='区分大小写', variable=self.case_var).pack(side='left', padx=8)

        # 条件列表编辑器
        terms_row = ttk.Frame(cfg)
        terms_row.pack(fill='x', padx=6, pady=2)

        list_frame = ttk.Frame(terms_row)
        list_frame.pack(side='left', fill='x', expand=True)
        self.terms_list = tk.Listbox(list_frame, height=3, exportselection=False)
        tsb = ttk.Scrollbar(list_frame, orient='vertical', command=self.terms_list.yview)
        self.terms_list.configure(yscrollcommand=tsb.set)
        tsb.pack(side='right', fill='y')
        self.terms_list.pack(side='left', fill='both', expand=True)

        btn_col = ttk.Frame(terms_row)
        btn_col.pack(side='left', padx=(10, 0))
        ttk.Label(btn_col, text='内容：').pack(anchor='w')
        self.target_entry = ttk.Entry(btn_col, width=12)
        self.target_entry.pack(fill='x', pady=(2, 2))
        self.target_entry.bind('<Return>', lambda e: self._add_term())
        ttk.Label(btn_col, text='匹配：').pack(anchor='w')
        self.match_combo = ttk.Combobox(btn_col, state='readonly', width=8, values=['包含', '等于'])
        self.match_combo.set('包含')
        self.match_combo.pack(fill='x', pady=(2, 4))
        ttk.Button(btn_col, text='添加', width=8, command=self._add_term).pack(pady=1)
        ttk.Button(btn_col, text='删除', width=8, command=self._del_term).pack(pady=1)
        ttk.Button(btn_col, text='上移', width=8, command=lambda: self._move_term(-1)).pack(pady=1)
        ttk.Button(btn_col, text='下移', width=8, command=lambda: self._move_term(1)).pack(pady=1)

        row3 = ttk.Frame(cfg)
        row3.pack(fill='x', padx=6, pady=(2, 2))
        ttk.Label(row3, text='说明：如添加"作废(包含)"和"计算(等于)"，则任一单元格含"作废"或等于"计算"的行都会被删除。',
                  foreground='#57606a').pack(anchor='w')

        # 规则保存/载入
        profile_row = ttk.Frame(cfg)
        profile_row.pack(fill='x', padx=6, pady=(2, 6))
        ttk.Button(profile_row, text='保存当前规则…', command=self._save_rule).pack(side='left')
        ttk.Label(profile_row, text='  已保存规则：').pack(side='left', padx=(8, 2))
        self.rule_cb = ttk.Combobox(profile_row, state='readonly', width=18)
        self.rule_cb.pack(side='left')
        ttk.Button(profile_row, text='载入', command=self._load_rule).pack(side='left', padx=2)
        ttk.Button(profile_row, text='删除', command=self._delete_rule).pack(side='left', padx=2)
        self._refresh_rule_combobox()

        # ③ 预览
        prev = ttk.LabelFrame(body.inner, text='③ 预览（统计每个文件将删除的行/列数）')
        prev.pack(fill='x', padx=8, pady=4)
        pframe = ttk.Frame(prev)
        pframe.pack(fill='x', padx=6, pady=(6, 2))
        self.preview_button = ttk.Button(pframe, text='预览', command=self._do_preview)
        self.preview_button.pack(side='left')
        ttk.Label(pframe, textvariable=self.status_var, foreground='#57606a').pack(side='left', padx=8)

        pbox = ttk.Frame(prev)
        pbox.pack(fill='both', expand=True, padx=6, pady=(2, 6))
        self.preview_text = tk.Text(pbox, height=3, state='disabled', wrap='none',
                                    font=(UI_FONT, 9))
        hsb = ttk.Scrollbar(pbox, orient='horizontal', command=self.preview_text.xview)
        self.preview_text.configure(xscrollcommand=hsb.set)
        hsb.pack(side='bottom', fill='x')
        self.preview_text.pack(side='left', fill='both', expand=True)

        # ④ 输出与执行
        out = ttk.LabelFrame(body.inner, text='④ 输出文件夹与执行（原文件保留不动，删除结果输出到新文件夹）')
        out.pack(fill='x', padx=8, pady=4)
        orow = ttk.Frame(out)
        orow.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(orow, text='输出文件夹：').pack(side='left')
        self.out_var = tk.StringVar()
        ttk.Entry(orow, textvariable=self.out_var).pack(side='left', fill='x', expand=True)
        ttk.Button(orow, text='浏览…', command=self._pick_outdir).pack(side='left', padx=4)

        brow = ttk.Frame(out)
        brow.pack(fill='x', padx=6, pady=(2, 6))
        self.run_button = ttk.Button(brow, text='开始删除', command=self._run)
        self.run_button.pack(side='left')
        self.cancel_button = ttk.Button(brow, text='取消', command=self._cancel, state='disabled')
        self.cancel_button.pack(side='left', padx=4)

        self._build_log_area(self)

    # ---- 条件编辑 ----
    def _add_term(self):
        v = self.target_entry.get().strip()
        if not v:
            messagebox.showinfo('提示', '请输入查找内容', parent=self)
            return
        m = 'contains' if self.match_combo.get() == '包含' else 'equals'
        if (v, m) in self._terms:
            messagebox.showinfo('提示', '该条件已存在', parent=self)
            return
        self._terms.append((v, m))
        self.target_entry.delete(0, 'end')
        self._refresh_terms()

    def _del_term(self):
        for i in reversed(self.terms_list.curselection()):
            del self._terms[i]
        self._refresh_terms()

    def _move_term(self, delta: int):
        sel = self.terms_list.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if j < 0 or j >= len(self._terms):
            return
        self._terms[i], self._terms[j] = self._terms[j], self._terms[i]
        self._refresh_terms()
        self.terms_list.selection_set(j)
        self.terms_list.activate(j)

    def _refresh_terms(self):
        self.terms_list.delete(0, 'end')
        for v, m in self._terms:
            self.terms_list.insert('end', f'{v}（{"包含" if m == "contains" else "等于"}）')

    # ---- 规则保存/载入 ----
    def _save_rule(self):
        if not self._terms:
            messagebox.showwarning('提示', '请先添加至少一条查找条件再保存', parent=self)
            return
        name = simpledialog.askstring('保存规则', '请输入规则名称（留空则用时间命名）：', parent=self)
        if name is None:
            return
        name = name.strip() or time.strftime('%m-%d %H:%M')
        if name in config_store.list_profiles('delete'):
            if not messagebox.askyesno('提示', f'规则「{name}」已存在，是否覆盖？', parent=self):
                return
        config_store.save_profile('delete', name, {
            'mode': self.mode_var.get(),
            'terms': [list(t) for t in self._terms],
            'case_sensitive': bool(self.case_var.get()),
        })
        self._refresh_rule_combobox(name)
        self._append_log(f'[提示] 删除规则「{name}」已保存')

    def _load_rule(self):
        name = self.rule_cb.get()
        if not name:
            messagebox.showinfo('提示', '请先在列表中选择要载入的规则', parent=self)
            return
        prof = config_store.get_profile('delete', name)
        if not prof:
            messagebox.showwarning('提示', f'规则「{name}」不存在', parent=self)
            return
        self.mode_var.set(prof.get('mode', 'row'))
        self._terms = []
        for t in prof.get('terms', []):
            if len(t) == 2:
                self._terms.append((str(t[0]), t[1] if t[1] in ('contains', 'equals') else 'contains'))
        self._refresh_terms()
        self.case_var.set(bool(prof.get('case_sensitive', False)))
        self._append_log(f'[提示] 已载入删除规则「{name}」')

    def _delete_rule(self):
        name = self.rule_cb.get()
        if not name:
            return
        if not messagebox.askyesno('删除规则', f'确定删除规则「{name}」吗？', parent=self):
            return
        config_store.delete_profile('delete', name)
        self._refresh_rule_combobox()
        self._append_log(f'[提示] 已删除规则「{name}」')

    def _refresh_rule_combobox(self, select=None):
        names = config_store.list_profiles('delete')
        self.rule_cb['values'] = names
        if select and select in names:
            self.rule_cb.set(select)
        elif self.rule_cb.get() not in names and names:
            self.rule_cb.set(names[0])
        elif not names:
            self.rule_cb.set('')

    # ---- 交互 ----
    def _pick_outdir(self):
        d = filedialog.askdirectory(parent=self, title='选择输出文件夹')
        if d:
            self.out_var.set(d)

    def _collect(self):
        files = list(self.picker.files)
        mode = self.mode_var.get()
        terms = list(self._terms)
        case_sensitive = bool(self.case_var.get())
        out_dir = self.out_var.get().strip()
        return files, mode, terms, case_sensitive, out_dir

    def _validate(self, files, terms, out_dir, same_dir_warn=True) -> bool:
        if not files:
            messagebox.showwarning('提示', '请先选择要处理的文件', parent=self)
            return False
        if not terms:
            messagebox.showwarning('提示', '请至少添加一条查找条件', parent=self)
            return False
        if not out_dir:
            messagebox.showwarning('提示', '请选择输出文件夹', parent=self)
            return False
        if same_dir_warn:
            from pathlib import Path as _P
            same = [str(f) for f in files if _P(f).parent == _P(out_dir)]
            if same:
                if not messagebox.askyesno('提示',
                        '输出文件夹与部分源文件所在文件夹相同。\n'
                        '为避免覆盖原文件，输出文件会自动改名（如 名称 (1).xlsx）。\n\n'
                        '建议选择不同的输出文件夹。是否继续？', parent=self):
                    return False
        return True

    def _do_preview(self):
        files, mode, terms, case_sensitive, out_dir = self._collect()
        if not self._validate(files, terms, out_dir, same_dir_warn=False):
            return
        self._clear_log()
        self.start_worker(self._worker_preview, (files, mode, terms, case_sensitive))

    def _run(self):
        files, mode, terms, case_sensitive, out_dir = self._collect()
        if not self._validate(files, terms, out_dir, same_dir_warn=True):
            return
        self._clear_log()
        self.start_worker(self._worker_run, (files, mode, terms, case_sensitive, out_dir))

    # ---- 工作线程 ----
    def _worker_preview(self, queue, cancel, files, mode, terms, case_sensitive):
        unit = '行' if mode == 'row' else '列'
        lines = []
        for f in files:
            try:
                deleted, total = excel_core.count_delete(f, mode, terms, case_sensitive)
                lines.append(f'{Path(f).name}  →  将删除 {deleted} {unit}（共 {total} {unit}）')
            except Exception as e:
                lines.append(f'[错误] 无法读取 {f}：{e}')
        queue.put(('preview', lines))
        queue.put(('done', None))

    def _worker_run(self, queue, cancel, files, mode, terms, case_sensitive, out_dir):
        log = lambda m: queue.put(('log', m))
        plan, errors = excel_core.plan_delete(files, mode, terms, out_dir)
        for f, msg in errors:
            log(f'[失败] {f}：{msg}')
        if not plan:
            queue.put(('error', '没有可处理的文件'))
            queue.put(('done', None))
            return
        summary = excel_core.execute_delete(plan, mode, terms, case_sensitive, log,
                                            lambda d, t: queue.put(('progress', (d, t))),
                                            cancel)
        queue.put(('done', summary))

    def _show_preview(self, lines):
        self.preview_text.configure(state='normal')
        self.preview_text.delete('1.0', 'end')
        self.preview_text.insert('end', '\n'.join(lines) + ('\n' if lines else ''))
        self.preview_text.configure(state='disabled')
        self.status_var.set(f'预览：共 {len(lines)} 项')


# ---------------------------------------------------------------------------
# Tab 4：端口拆分（按端口号拆表 + 流量排序 + 标红汇总）
# ---------------------------------------------------------------------------

class SplitTab(BaseTab):
    def __init__(self, master):
        super().__init__(master)
        self._build()

    def _build(self):
        body = ScrollableFrame(self)
        body.pack(fill='both', expand=True)

        # ① 选择文件
        picker = FilePicker(body.inner, text='① 选择要拆分的文件（.xlsx / .xlsm，按第一个工作表处理）')
        picker.pack(fill='x', padx=8, pady=(8, 4))
        self.picker = picker

        # ② 拆分与排序设置
        cfg = ttk.LabelFrame(body.inner, text='② 拆分与排序规则')
        cfg.pack(fill='x', padx=8, pady=4)

        row1 = ttk.Frame(cfg)
        row1.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(row1, text='端口号列名：').pack(side='left')
        self.port_entry = ttk.Entry(row1, width=16)
        self.port_entry.insert(0, excel_core.DEFAULT_PORT_HEADER)
        self.port_entry.pack(side='left')
        ttk.Label(row1, text='  流入速度列名：').pack(side='left', padx=(10, 0))
        self.inflow_entry = ttk.Entry(row1, width=12)
        self.inflow_entry.insert(0, excel_core.DEFAULT_INFLOW_HEADER)
        self.inflow_entry.pack(side='left')
        ttk.Label(row1, text='  流出速度列名：').pack(side='left', padx=(10, 0))
        self.outflow_entry = ttk.Entry(row1, width=12)
        self.outflow_entry.insert(0, excel_core.DEFAULT_OUTFLOW_HEADER)
        self.outflow_entry.pack(side='left')

        row2 = ttk.Frame(cfg)
        row2.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(row2, text='标红行号（含表头，第 1 行为表头）：').pack(side='left')
        self.highlight_var = tk.StringVar(value=str(excel_core.DEFAULT_HIGHLIGHT_ROW))
        ttk.Spinbox(row2, from_=2, to=1000000, width=8,
                    textvariable=self.highlight_var).pack(side='left', padx=4)
        ttk.Label(row2, text='排序后该行整行字体标红，并汇总到新子表',
                  foreground='#57606a').pack(side='left', padx=6)

        note = ttk.Frame(cfg)
        note.pack(fill='x', padx=6, pady=(2, 6))
        ttk.Label(
            note,
            text='规则：按端口号列的不同值拆分（表头保持一致）；分别计算每个子表「流入速度」与'
                 '「流出速度」的总和，按总和较大的一列降序排列；非数值/空值排在最后。',
            foreground='#57606a', wraplength=820, justify='left').pack(anchor='w')

        # ③ 预览
        prev = ttk.LabelFrame(body.inner, text='③ 预览（子表数量、排序依据、标红行）')
        prev.pack(fill='x', padx=8, pady=4)
        pframe = ttk.Frame(prev)
        pframe.pack(fill='x', padx=6, pady=(6, 2))
        self.preview_button = ttk.Button(pframe, text='预览', command=self._do_preview)
        self.preview_button.pack(side='left')
        ttk.Label(pframe, textvariable=self.status_var, foreground='#57606a').pack(side='left', padx=8)

        pbox = ttk.Frame(prev)
        pbox.pack(fill='both', expand=True, padx=6, pady=(2, 6))
        self.preview_text = tk.Text(pbox, height=5, state='disabled', wrap='none',
                                    font=(UI_FONT, 9))
        hsb = ttk.Scrollbar(pbox, orient='horizontal', command=self.preview_text.xview)
        self.preview_text.configure(xscrollcommand=hsb.set)
        hsb.pack(side='bottom', fill='x')
        vsb = ttk.Scrollbar(pbox, orient='vertical', command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self.preview_text.pack(side='left', fill='both', expand=True)

        # ④ 输出与执行
        out = ttk.LabelFrame(body.inner, text='④ 输出文件夹与执行（输出一个 .xlsx，内含各端口子表 + 标红行汇总表）')
        out.pack(fill='x', padx=8, pady=4)
        orow = ttk.Frame(out)
        orow.pack(fill='x', padx=6, pady=(6, 2))
        ttk.Label(orow, text='输出文件夹：').pack(side='left')
        self.out_var = tk.StringVar()
        ttk.Entry(orow, textvariable=self.out_var).pack(side='left', fill='x', expand=True)
        ttk.Button(orow, text='浏览…', command=self._pick_outdir).pack(side='left', padx=4)

        brow = ttk.Frame(out)
        brow.pack(fill='x', padx=6, pady=(2, 6))
        self.run_button = ttk.Button(brow, text='开始拆分', command=self._run)
        self.run_button.pack(side='left')
        self.cancel_button = ttk.Button(brow, text='取消', command=self._cancel, state='disabled')
        self.cancel_button.pack(side='left', padx=4)

        self._build_log_area(self)

    def _pick_outdir(self):
        d = filedialog.askdirectory(parent=self, title='选择输出文件夹')
        if d:
            self.out_var.set(d)

    def _collect(self):
        files = list(self.picker.files)
        port_h = self.port_entry.get().strip()
        inflow_h = self.inflow_entry.get().strip()
        outflow_h = self.outflow_entry.get().strip()
        highlight = self.highlight_var.get().strip()
        out_dir = self.out_var.get().strip()
        return files, port_h, inflow_h, outflow_h, highlight, out_dir

    def _validate(self, files, port_h, inflow_h, outflow_h, highlight, out_dir) -> bool:
        if not files:
            messagebox.showwarning('提示', '请先选择要拆分的文件', parent=self)
            return False
        if not (port_h and inflow_h and outflow_h):
            messagebox.showwarning('提示', '请填写端口号、流入速度、流出速度三个列名', parent=self)
            return False
        if len({port_h, inflow_h, outflow_h}) < 3:
            messagebox.showwarning('提示', '三个列名不能相同', parent=self)
            return False
        try:
            if int(highlight) < 2:
                raise ValueError
        except ValueError:
            messagebox.showwarning('提示', '标红行号必须是 ≥ 2 的整数（第 1 行是表头）', parent=self)
            return False
        if not out_dir:
            messagebox.showwarning('提示', '请选择输出文件夹', parent=self)
            return False
        bad = [f for f in files if Path(f).suffix.lower() not in ('.xlsx', '.xlsm')]
        if bad:
            messagebox.showwarning(
                '提示',
                '该功能仅支持 .xlsx / .xlsm 文件：\n' + '\n'.join(Path(f).name for f in bad[:8]),
                parent=self)
            return False
        return True

    def _do_preview(self):
        files, port_h, inflow_h, outflow_h, highlight, out_dir = self._collect()
        if not self._validate(files, port_h, inflow_h, outflow_h, highlight, out_dir):
            return
        self._clear_log()
        self.start_worker(self._worker_preview,
                          (files, port_h, inflow_h, outflow_h, int(highlight)))

    def _run(self):
        files, port_h, inflow_h, outflow_h, highlight, out_dir = self._collect()
        if not self._validate(files, port_h, inflow_h, outflow_h, highlight, out_dir):
            return
        self._clear_log()
        self.start_worker(self._worker_run,
                          (files, port_h, inflow_h, outflow_h, int(highlight), out_dir))

    # ---- 工作线程 ----
    def _worker_preview(self, queue, cancel, files, port_h, inflow_h, outflow_h, highlight):
        lines = []
        for f in files:
            try:
                lines.extend(excel_core.preview_split(f, port_h, inflow_h, outflow_h, highlight))
            except Exception as e:
                lines.append(f'[错误] 无法读取 {Path(f).name}：{e}')
            lines.append('')
        queue.put(('preview', lines))
        queue.put(('done', None))

    def _worker_run(self, queue, cancel, files, port_h, inflow_h, outflow_h, highlight, out_dir):
        log = lambda m: queue.put(('log', m))
        plan, errors = excel_core.plan_split(files, out_dir)
        for f, msg in errors:
            log(f'[失败] {f}：{msg}')
        if not plan:
            queue.put(('error', '没有可处理的文件'))
            queue.put(('done', None))
            return
        summary = excel_core.execute_split(
            plan, port_h, inflow_h, outflow_h, highlight, log,
            lambda d, t: queue.put(('progress', (d, t))), cancel)
        queue.put(('done', summary))

    def _show_preview(self, lines):
        self.preview_text.configure(state='normal')
        self.preview_text.delete('1.0', 'end')
        self.preview_text.insert('end', '\n'.join(lines) + ('\n' if lines else ''))
        self.preview_text.configure(state='disabled')
        self.status_var.set(f'预览：共 {len(lines)} 行')


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Lsy的Excel-Tool')
        self.geometry('900x740')
        self.minsize(800, 620)

        for name in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont'):
            try:
                tkfont.nametofont(name).configure(family=UI_FONT, size=9)
            except tk.TclError:
                pass

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True)
        self.rename_tab = RenameTab(nb)
        self.fill_tab = FillTab(nb)
        self.delete_tab = DeleteTab(nb)
        self.split_tab = SplitTab(nb)
        nb.add(self.rename_tab, text='批量重命名')
        nb.add(self.fill_tab, text='模板填充')
        nb.add(self.delete_tab, text='行/列删除')
        nb.add(self.split_tab, text='端口拆分')


def main():
    if IS_WIN:
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    sys.excepthook = _excepthook
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
