from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from engine import DATA_PATH, LOG_PATH, SyncService, SyncTask, list_headers, list_sheets
from metadata import APP_NAME, APP_VERSION

SOURCE_MODE_LABELS = {
    'Whole sheet': 'whole_sheet',
    'Custom range': 'custom_range',
}
SOURCE_MODE_VALUES = {value: label for label, value in SOURCE_MODE_LABELS.items()}
TARGET_MODE_LABELS = {
    'Replace target sheet': 'replace_sheet',
    'Write from cell': 'write_from_cell',
}
TARGET_MODE_VALUES = {value: label for label, value in TARGET_MODE_LABELS.items()}


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f'{APP_NAME} v{APP_VERSION}')
        self.root.geometry('1320x820')
        self.service = SyncService(status_callback=self._schedule_refresh)
        self.pending_refresh = False
        self.selected_task_id: str | None = None
        self.header_vars: dict[str, tk.BooleanVar] = {}
        self.columns_window = None
        self._build_vars()
        self._build_ui()
        self.service.start()
        self._refresh()
        self.root.after(1000, self._poll_ui)

    def _build_vars(self):
        self.name_var = tk.StringVar()
        self.enabled_var = tk.BooleanVar(value=True)
        self.source_file_var = tk.StringVar()
        self.source_sheet_var = tk.StringVar()
        self.source_mode_var = tk.StringVar(value='Whole sheet')
        self.source_range_var = tk.StringVar()
        self.target_file_var = tk.StringVar()
        self.target_sheet_var = tk.StringVar()
        self.target_mode_var = tk.StringVar(value='Replace target sheet')
        self.target_start_cell_var = tk.StringVar(value='A1')
        self.header_row_var = tk.StringVar(value='1')
        self.data_start_row_var = tk.StringVar(value='2')
        self.formula_var = tk.StringVar(value='values')

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=2)
        self.root.rowconfigure(1, weight=1)

        top = ttk.Frame(self.root, padding=10)
        top.grid(row=0, column=0, columnspan=2, sticky='ew')
        self.monitor_button = ttk.Button(top, text='Pause monitoring', command=self._toggle_monitoring)
        self.monitor_button.pack(side='left')
        ttk.Button(top, text='New task', command=self._new_task).pack(side='left', padx=(20, 0))
        ttk.Button(top, text='Open log', command=self._open_log).pack(side='left', padx=(8, 0))
        ttk.Button(top, text='Open data file', command=self._open_data_file).pack(side='left', padx=(8, 0))
        ttk.Label(top, text=f'Version {APP_VERSION}').pack(side='right', padx=(0, 16))
        self.status_label = ttk.Label(top, text='Monitoring on')
        self.status_label.pack(side='right')

        left = ttk.LabelFrame(self.root, text='Tasks', padding=10)
        left.grid(row=1, column=0, sticky='nsew', padx=(10, 5), pady=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        cols = ('enabled', 'name', 'status', 'last_startsync', 'last_synced')
        self.tree = ttk.Treeview(left, columns=cols, show='headings')
        for key, title, width in (
            ('enabled', 'On', 50),
            ('name', 'Task', 180),
            ('status', 'Status', 220),
            ('last_startsync', 'Last startsync since', 170),
            ('last_synced', 'Last synced', 170),
        ):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor='center' if key == 'enabled' else 'w')
        tree_scroll = ttk.Scrollbar(left, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        tree_scroll.grid(row=0, column=1, sticky='ns')
        self.tree.bind('<<TreeviewSelect>>', self._on_select)

        right = ttk.Frame(self.root)
        right.grid(row=1, column=1, sticky='nsew', padx=(5, 10), pady=(0, 10))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        editor = ttk.LabelFrame(right, text='Task Editor', padding=12)
        editor.grid(row=0, column=0, sticky='nsew')
        editor.columnconfigure(1, weight=1)
        row = 0

        ttk.Label(editor, text='Task name').grid(row=row, column=0, sticky='w', pady=4)
        ttk.Entry(editor, textvariable=self.name_var).grid(row=row, column=1, sticky='ew', pady=4)
        row += 1

        ttk.Checkbutton(editor, text='Enabled', variable=self.enabled_var).grid(
            row=row, column=0, columnspan=2, sticky='w', pady=4
        )
        row += 1

        ttk.Label(editor, text='Source file').grid(row=row, column=0, sticky='w', pady=4)
        src = ttk.Frame(editor)
        src.grid(row=row, column=1, sticky='ew', pady=4)
        src.columnconfigure(0, weight=1)
        ttk.Entry(src, textvariable=self.source_file_var).grid(row=0, column=0, sticky='ew')
        ttk.Button(src, text='Browse', command=self._pick_source).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(src, text='Load sheets', command=self._load_sheets).grid(row=0, column=2, padx=(8, 0))
        row += 1

        ttk.Label(editor, text='Source sheet').grid(row=row, column=0, sticky='w', pady=4)
        self.sheet_combo = ttk.Combobox(editor, textvariable=self.source_sheet_var, state='readonly')
        self.sheet_combo.grid(row=row, column=1, sticky='ew', pady=4)
        self.sheet_combo.bind('<<ComboboxSelected>>', lambda _e: self._load_headers())
        row += 1

        ttk.Label(editor, text='Source mode').grid(row=row, column=0, sticky='w', pady=4)
        self.source_mode_combo = ttk.Combobox(
            editor,
            textvariable=self.source_mode_var,
            state='readonly',
            values=list(SOURCE_MODE_LABELS.keys()),
        )
        self.source_mode_combo.grid(row=row, column=1, sticky='ew', pady=4)
        self.source_mode_combo.bind('<<ComboboxSelected>>', self._on_source_mode_change)
        row += 1

        self.source_range_label = ttk.Label(editor, text='Source range')
        self.source_range_label.grid(row=row, column=0, sticky='w', pady=4)
        self.source_range_entry = ttk.Entry(editor, textvariable=self.source_range_var)
        self.source_range_entry.grid(row=row, column=1, sticky='ew', pady=4)
        self.source_range_entry.bind('<FocusOut>', lambda _e: self._load_headers())
        self.source_range_entry.bind('<Return>', lambda _e: self._load_headers())
        row += 1

        ttk.Label(editor, text='Target file').grid(row=row, column=0, sticky='w', pady=4)
        dst = ttk.Frame(editor)
        dst.grid(row=row, column=1, sticky='ew', pady=4)
        dst.columnconfigure(0, weight=1)
        ttk.Entry(dst, textvariable=self.target_file_var).grid(row=0, column=0, sticky='ew')
        ttk.Button(dst, text='Browse', command=self._pick_target).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(dst, text='Load sheets', command=self._load_target_sheets).grid(row=0, column=2, padx=(8, 0))
        row += 1

        ttk.Label(editor, text='Target sheet').grid(row=row, column=0, sticky='w', pady=4)
        self.target_sheet_combo = ttk.Combobox(editor, textvariable=self.target_sheet_var, state='normal')
        self.target_sheet_combo.grid(row=row, column=1, sticky='ew', pady=4)
        row += 1

        ttk.Label(editor, text='Target mode').grid(row=row, column=0, sticky='w', pady=4)
        self.target_mode_combo = ttk.Combobox(
            editor,
            textvariable=self.target_mode_var,
            state='readonly',
            values=list(TARGET_MODE_LABELS.keys()),
        )
        self.target_mode_combo.grid(row=row, column=1, sticky='ew', pady=4)
        self.target_mode_combo.bind('<<ComboboxSelected>>', self._on_target_mode_change)
        row += 1

        self.target_start_label = ttk.Label(editor, text='Target start cell')
        self.target_start_label.grid(row=row, column=0, sticky='w', pady=4)
        self.target_start_entry = ttk.Entry(editor, textvariable=self.target_start_cell_var)
        self.target_start_entry.grid(row=row, column=1, sticky='ew', pady=4)
        row += 1

        adv = ttk.LabelFrame(editor, text='Advanced', padding=8)
        adv.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(8, 10))
        ttk.Label(adv, text='Header row').grid(row=0, column=0, sticky='w', pady=2)
        self.header_row_entry = ttk.Entry(adv, textvariable=self.header_row_var, width=10)
        self.header_row_entry.grid(row=0, column=1, sticky='w', pady=2)
        self.header_row_entry.bind('<FocusOut>', lambda _e: self._load_headers())
        self.header_row_entry.bind('<Return>', lambda _e: self._load_headers())
        ttk.Label(adv, text='Data start row').grid(row=1, column=0, sticky='w', pady=2)
        self.data_start_row_entry = ttk.Entry(adv, textvariable=self.data_start_row_var, width=10)
        self.data_start_row_entry.grid(row=1, column=1, sticky='w', pady=2)
        ttk.Label(adv, text='Formula output').grid(row=2, column=0, sticky='w', pady=2)
        ttk.Combobox(
            adv,
            textvariable=self.formula_var,
            state='readonly',
            values=['values', 'formulas'],
            width=12,
        ).grid(row=2, column=1, sticky='w', pady=2)
        row += 1

        ttk.Label(editor, text='Columns').grid(row=row, column=0, sticky='nw', pady=4)
        cols_frame = ttk.Frame(editor)
        cols_frame.grid(row=row, column=1, sticky='nsew')
        editor.rowconfigure(row, weight=1)
        cols_frame.columnconfigure(0, weight=1)
        cols_frame.rowconfigure(1, weight=1)

        selection_actions = ttk.Frame(cols_frame)
        selection_actions.grid(row=0, column=0, sticky='w', pady=(0, 6))
        ttk.Button(selection_actions, text='Select all', command=self._select_all_headers).pack(side='left')
        ttk.Button(selection_actions, text='Clear', command=self._clear_headers).pack(side='left', padx=(8, 0))
        ttk.Button(selection_actions, text='Invert', command=self._invert_headers).pack(side='left', padx=(8, 0))

        self.columns_canvas = tk.Canvas(cols_frame, height=240, highlightthickness=0)
        cols_scroll = ttk.Scrollbar(cols_frame, orient='vertical', command=self.columns_canvas.yview)
        self.columns_canvas.configure(yscrollcommand=cols_scroll.set)
        self.columns_canvas.grid(row=1, column=0, sticky='nsew')
        cols_scroll.grid(row=1, column=1, sticky='ns')
        background = self.columns_canvas.cget('background')
        self.columns_inner = tk.Frame(self.columns_canvas, bg=background)
        self.columns_inner.bind(
            '<Configure>',
            lambda _e: self.columns_canvas.configure(scrollregion=self.columns_canvas.bbox('all')),
        )
        self.columns_window = self.columns_canvas.create_window((0, 0), window=self.columns_inner, anchor='nw')
        self.columns_canvas.bind('<Configure>', self._resize_columns_area)
        row += 1

        actions = ttk.Frame(editor)
        actions.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(10, 0))
        ttk.Button(actions, text='Save task', command=self._save_task).pack(side='left')
        ttk.Button(actions, text='Save && Sync now', command=self._run_now).pack(side='left', padx=(8, 0))
        ttk.Button(actions, text='Copy task', command=self._copy_task).pack(side='left', padx=(8, 0))
        ttk.Button(actions, text='Delete task', command=self._delete_task).pack(side='left', padx=(8, 0))

        logs = ttk.LabelFrame(right, text='Recent log', padding=10)
        logs.grid(row=1, column=0, sticky='nsew', pady=(10, 0))
        logs.columnconfigure(0, weight=1)
        logs.rowconfigure(0, weight=1)
        self.log_text = tk.Text(logs, height=12, wrap='word', state='disabled')
        log_scroll = ttk.Scrollbar(logs, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.grid(row=0, column=0, sticky='nsew')
        log_scroll.grid(row=0, column=1, sticky='ns')

        self._show_columns_placeholder('Load a source sheet to choose columns.')
        self._update_source_mode_ui()
        self._update_target_mode_ui()

    def _schedule_refresh(self):
        self.pending_refresh = True

    def _poll_ui(self):
        if self.pending_refresh:
            self.pending_refresh = False
            self._refresh()
        self.root.after(1000, self._poll_ui)

    def _refresh(self):
        rows = self.service.list_runtime_rows()
        ids = set()
        for task, runtime in rows:
            ids.add(task.id)
            values = (
                'Yes' if task.enabled else 'No',
                task.name,
                runtime.status,
                runtime.last_startsync_at,
                runtime.last_synced_at,
            )
            if self.tree.exists(task.id):
                self.tree.item(task.id, values=values)
            else:
                self.tree.insert('', 'end', iid=task.id, values=values)
        for item_id in self.tree.get_children():
            if item_id not in ids:
                self.tree.delete(item_id)
        running = self.service.is_running()
        self.status_label.configure(text='Monitoring on' if running else 'Monitoring paused')
        self.monitor_button.configure(text='Pause monitoring' if running else 'Resume monitoring')
        self._refresh_log()

    def _refresh_log(self):
        tail = ''
        if LOG_PATH.exists():
            lines = LOG_PATH.read_text(encoding='utf-8', errors='ignore').splitlines()
            tail = '\n'.join(lines[-25:])
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.insert('1.0', tail)
        self.log_text.configure(state='disabled')

    def _update_source_mode_ui(self):
        is_custom = SOURCE_MODE_LABELS.get(self.source_mode_var.get(), 'whole_sheet') == 'custom_range'
        if is_custom:
            self.source_range_label.grid()
            self.source_range_entry.grid()
            self.header_row_entry.state(['disabled'])
            self.data_start_row_entry.state(['disabled'])
        else:
            self.source_range_label.grid_remove()
            self.source_range_entry.grid_remove()
            self.header_row_entry.state(['!disabled'])
            self.data_start_row_entry.state(['!disabled'])

    def _update_target_mode_ui(self):
        write_from_cell = TARGET_MODE_LABELS.get(self.target_mode_var.get(), 'replace_sheet') == 'write_from_cell'
        if write_from_cell:
            self.target_start_label.grid()
            self.target_start_entry.grid()
        else:
            self.target_start_label.grid_remove()
            self.target_start_entry.grid_remove()

    def _on_source_mode_change(self, _event=None):
        self._update_source_mode_ui()
        self._load_headers()

    def _on_target_mode_change(self, _event=None):
        self._update_target_mode_ui()

    def _pick_source(self):
        path = filedialog.askopenfilename(
            title='Choose source workbook',
            filetypes=[('Excel files', '*.xlsx *.xlsm'), ('All files', '*.*')],
        )
        if path:
            self.source_file_var.set(path)
            self._load_sheets()

    def _pick_target(self):
        path = filedialog.asksaveasfilename(
            title='Choose target workbook',
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')],
        )
        if path:
            self.target_file_var.set(path)
            self._load_target_sheets()
            if not self.target_sheet_var.get():
                self.target_sheet_var.set('Export')

    def _load_sheets(self):
        source_file = self.source_file_var.get().strip()
        if not source_file:
            return
        if Path(source_file).name.startswith('~$'):
            messagebox.showerror('Load sheets failed', 'Choose the real workbook, not the temporary lock file.')
            return
        try:
            sheets = list_sheets(source_file)
        except Exception as exc:
            messagebox.showerror('Load sheets failed', str(exc))
            return
        self.sheet_combo['values'] = sheets
        if sheets and self.source_sheet_var.get() not in sheets:
            self.source_sheet_var.set(sheets[0])
        self._load_headers()

    def _load_target_sheets(self):
        target_file = self.target_file_var.get().strip()
        if not target_file:
            self.target_sheet_combo['values'] = []
            return
        target_path = Path(target_file)
        if not target_path.exists():
            self.target_sheet_combo['values'] = []
            if not self.target_sheet_var.get():
                self.target_sheet_var.set('Export')
            return
        try:
            sheets = list_sheets(target_file)
        except Exception as exc:
            messagebox.showerror('Load target sheets failed', str(exc))
            return
        self.target_sheet_combo['values'] = sheets
        if sheets and not self.target_sheet_var.get():
            self.target_sheet_var.set(sheets[0])

    def _show_columns_placeholder(self, text: str):
        for child in self.columns_inner.winfo_children():
            child.destroy()
        self.header_vars = {}
        background = self.columns_inner.cget('bg')
        label = tk.Label(
            self.columns_inner,
            text=text,
            anchor='w',
            justify='left',
            bg=background,
            fg='#666666',
            padx=4,
            pady=8,
        )
        label.grid(row=0, column=0, sticky='w')

    def _load_headers(self):
        source_file = self.source_file_var.get().strip()
        source_sheet = self.source_sheet_var.get().strip()
        if not source_file or not source_sheet:
            self._show_columns_placeholder('Load a source sheet to choose columns.')
            return
        try:
            headers = list_headers(
                source_file,
                source_sheet,
                int(self.header_row_var.get() or '1'),
                source_mode=SOURCE_MODE_LABELS.get(self.source_mode_var.get(), 'whole_sheet'),
                source_range=self.source_range_var.get().strip(),
            )
        except Exception as exc:
            messagebox.showerror('Load columns failed', str(exc))
            return
        current = {header for header, var in self.header_vars.items() if var.get()}
        for child in self.columns_inner.winfo_children():
            child.destroy()
        self.header_vars = {}
        if not headers:
            self._show_columns_placeholder('No headers found on the selected sheet or range.')
            return
        columns_per_row = 2
        background = self.columns_inner.cget('bg')
        for col_idx in range(columns_per_row):
            self.columns_inner.grid_columnconfigure(col_idx, weight=1)
        for idx, header in enumerate(headers):
            var = tk.BooleanVar(value=header in current)
            self.header_vars[header] = var
            checkbox = tk.Checkbutton(
                self.columns_inner,
                text=header,
                variable=var,
                onvalue=True,
                offvalue=False,
                anchor='w',
                justify='left',
                bg=background,
                activebackground=background,
                selectcolor='white',
                highlightthickness=0,
                relief='flat',
                padx=4,
                pady=2,
            )
            checkbox.grid(row=idx // columns_per_row, column=idx % columns_per_row, sticky='ew', padx=(0, 12), pady=2)
        self._resize_columns_area()

    def _resize_columns_area(self, event=None):
        if self.columns_window is None:
            return
        width = event.width if event is not None else self.columns_canvas.winfo_width()
        if width > 1:
            self.columns_canvas.itemconfigure(self.columns_window, width=width)

    def _select_all_headers(self):
        for var in self.header_vars.values():
            var.set(True)

    def _clear_headers(self):
        for var in self.header_vars.values():
            var.set(False)

    def _invert_headers(self):
        for var in self.header_vars.values():
            var.set(not var.get())

    def _on_select(self, _event):
        selection = self.tree.selection()
        if not selection:
            return
        self.selected_task_id = selection[0]
        self._load_task_into_form(self.service.get_task(self.selected_task_id))

    def _load_task_into_form(self, task: SyncTask | None):
        if task is None:
            self.selected_task_id = None
            self.name_var.set('')
            self.enabled_var.set(True)
            self.source_file_var.set('')
            self.source_sheet_var.set('')
            self.source_mode_var.set('Whole sheet')
            self.source_range_var.set('')
            self.target_file_var.set('')
            self.target_sheet_var.set('')
            self.target_mode_var.set('Replace target sheet')
            self.target_start_cell_var.set('A1')
            self.header_row_var.set('1')
            self.data_start_row_var.set('2')
            self.formula_var.set('values')
            self.sheet_combo['values'] = []
            self.target_sheet_combo['values'] = []
            self._update_source_mode_ui()
            self._update_target_mode_ui()
            self._show_columns_placeholder('Load a source sheet to choose columns.')
            return

        self.name_var.set(task.name)
        self.enabled_var.set(task.enabled)
        self.source_file_var.set(task.source_file)
        self.source_sheet_var.set(task.source_sheet)
        self.source_mode_var.set(SOURCE_MODE_VALUES.get(task.source_mode, 'Whole sheet'))
        self.source_range_var.set(task.source_range)
        self.target_file_var.set(task.target_file)
        self.target_sheet_var.set(task.target_sheet)
        self.target_mode_var.set(TARGET_MODE_VALUES.get(task.target_mode, 'Replace target sheet'))
        self.target_start_cell_var.set(task.target_start_cell or 'A1')
        self.header_row_var.set(str(task.header_row))
        self.data_start_row_var.set(str(task.data_start_row))
        self.formula_var.set(task.formula_handling)
        self._update_source_mode_ui()
        self._update_target_mode_ui()
        try:
            sheets = list_sheets(task.source_file) if task.source_file else []
        except Exception:
            sheets = []
        self.sheet_combo['values'] = sheets
        self._load_target_sheets()
        self._load_headers()
        for header, var in self.header_vars.items():
            var.set(header in task.columns_by_header)

    def _build_task(self) -> SyncTask:
        headers = [header for header, var in self.header_vars.items() if var.get()]
        existing = self.service.get_task(self.selected_task_id) if self.selected_task_id else None
        return SyncTask(
            id=self.selected_task_id or SyncTask().id,
            name=self.name_var.get().strip() or 'New Task',
            enabled=self.enabled_var.get(),
            source_file=self.source_file_var.get().strip(),
            source_sheet=self.source_sheet_var.get().strip(),
            source_mode=SOURCE_MODE_LABELS.get(self.source_mode_var.get(), 'whole_sheet'),
            source_range=self.source_range_var.get().strip(),
            target_file=self.target_file_var.get().strip(),
            target_sheet=self.target_sheet_var.get().strip() or 'Export',
            target_mode=TARGET_MODE_LABELS.get(self.target_mode_var.get(), 'replace_sheet'),
            target_start_cell=self.target_start_cell_var.get().strip() or 'A1',
            columns_by_header=headers,
            header_row=int(self.header_row_var.get() or '1'),
            data_start_row=int(self.data_start_row_var.get() or '2'),
            formula_handling=self.formula_var.get().strip() or 'values',
            last_target_rows=existing.last_target_rows if existing else 0,
            last_target_cols=existing.last_target_cols if existing else 0,
        )

    def _persist_task(self) -> SyncTask:
        task = self._build_task()
        tasks = list(self.service.tasks)
        for idx, existing in enumerate(tasks):
            if existing.id == task.id:
                tasks[idx] = task
                break
        else:
            tasks.append(task)
        self.service.set_tasks(tasks)
        self.selected_task_id = task.id
        self._refresh()
        if self.tree.exists(task.id):
            self.tree.selection_set(task.id)
            self.tree.focus(task.id)
        self._load_task_into_form(task)
        return task

    def _save_task(self):
        try:
            self._persist_task()
        except Exception as exc:
            messagebox.showerror('Save task failed', str(exc))

    def _copy_task(self):
        source_task = self.service.get_task(self.selected_task_id) if self.selected_task_id else None
        if source_task is None:
            messagebox.showinfo('Copy task', 'Select a task first.')
            return
        copied = SyncTask(
            name=f'{source_task.name}-副本',
            enabled=source_task.enabled,
            source_file=source_task.source_file,
            source_sheet=source_task.source_sheet,
            source_mode=source_task.source_mode,
            source_range=source_task.source_range,
            target_file=source_task.target_file,
            target_sheet=source_task.target_sheet,
            target_mode=source_task.target_mode,
            target_start_cell=source_task.target_start_cell,
            columns_by_header=list(source_task.columns_by_header),
            header_row=source_task.header_row,
            data_start_row=source_task.data_start_row,
            copy_style=source_task.copy_style,
            copy_column_widths=source_task.copy_column_widths,
            copy_row_heights=source_task.copy_row_heights,
            include_header=source_task.include_header,
            drop_empty_rows=source_task.drop_empty_rows,
            formula_handling=source_task.formula_handling,
            last_target_rows=0,
            last_target_cols=0,
        )
        tasks = list(self.service.tasks)
        tasks.append(copied)
        self.service.set_tasks(tasks)
        self.selected_task_id = copied.id
        self._refresh()
        if self.tree.exists(copied.id):
            self.tree.selection_set(copied.id)
            self.tree.focus(copied.id)
        self._load_task_into_form(copied)

    def _delete_task(self):
        if not self.selected_task_id:
            return
        if not messagebox.askyesno('Delete task', 'Delete this task?'):
            return
        tasks = [task for task in self.service.tasks if task.id != self.selected_task_id]
        self.service.set_tasks(tasks)
        self._new_task()
        self._refresh()

    def _new_task(self):
        for item in self.tree.selection():
            self.tree.selection_remove(item)
        self._load_task_into_form(None)

    def _run_now(self):
        try:
            task = self._persist_task()
        except Exception as exc:
            messagebox.showerror('Save and sync failed', str(exc))
            return
        self.service.run_task_now(task.id)
        self._refresh()

    def _toggle_monitoring(self):
        if self.service.is_running():
            self.service.stop()
        else:
            self.service.start()
        self._refresh()

    def _open_log(self):
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        if not LOG_PATH.exists():
            LOG_PATH.write_text('', encoding='utf-8')
        os.startfile(LOG_PATH)

    def _open_data_file(self):
        if not DATA_PATH.exists():
            DATA_PATH.write_text('{"settings": {}, "tasks": []}', encoding='utf-8')
        os.startfile(DATA_PATH)

    def run(self):
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)
        self.root.mainloop()

    def _on_close(self):
        self.service.stop()
        self.root.destroy()


def main():
    root = tk.Tk()
    ttk.Style().theme_use('clam')
    App(root).run()


if __name__ == '__main__':
    main()
