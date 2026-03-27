from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from engine import DATA_PATH, LOG_PATH, SyncService, SyncTask, list_headers, list_sheets


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title('Excel Sync Manager')
        self.root.geometry('1280x760')
        self.service = SyncService(status_callback=self._schedule_refresh)
        self.pending_refresh = False
        self.selected_task_id: str | None = None
        self.header_vars: dict[str, tk.BooleanVar] = {}
        self._build_vars()
        self._build_ui()
        self._refresh()
        self.root.after(1000, self._poll_ui)

    def _build_vars(self):
        self.name_var = tk.StringVar()
        self.enabled_var = tk.BooleanVar(value=True)
        self.source_file_var = tk.StringVar()
        self.source_sheet_var = tk.StringVar()
        self.target_file_var = tk.StringVar()
        self.target_sheet_var = tk.StringVar()
        self.header_row_var = tk.StringVar(value='1')
        self.data_start_row_var = tk.StringVar(value='2')
        self.formula_var = tk.StringVar(value='values')

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=2)
        self.root.rowconfigure(1, weight=1)

        top = ttk.Frame(self.root, padding=10)
        top.grid(row=0, column=0, columnspan=2, sticky='ew')
        ttk.Button(top, text='Start monitoring', command=self._start).pack(side='left')
        ttk.Button(top, text='Stop monitoring', command=self._stop).pack(side='left', padx=(8, 0))
        ttk.Button(top, text='New task', command=self._new_task).pack(side='left', padx=(20, 0))
        ttk.Button(top, text='Open log', command=self._open_log).pack(side='left', padx=(8, 0))
        ttk.Button(top, text='Open data file', command=self._open_data_file).pack(side='left', padx=(8, 0))
        ttk.Label(top, text='Version 0.1').pack(side='right', padx=(0, 16))
        self.status_label = ttk.Label(top, text='Stopped')
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
            ('status', 'Status', 200),
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
        self.tree.bind('<Double-1>', self._toggle_enabled)

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

        ttk.Label(editor, text='Target file').grid(row=row, column=0, sticky='w', pady=4)
        dst = ttk.Frame(editor)
        dst.grid(row=row, column=1, sticky='ew', pady=4)
        dst.columnconfigure(0, weight=1)
        ttk.Entry(dst, textvariable=self.target_file_var).grid(row=0, column=0, sticky='ew')
        ttk.Button(dst, text='Browse', command=self._pick_target).grid(row=0, column=1, padx=(8, 0))
        row += 1

        ttk.Label(editor, text='Target sheet').grid(row=row, column=0, sticky='w', pady=4)
        ttk.Entry(editor, textvariable=self.target_sheet_var).grid(row=row, column=1, sticky='ew', pady=4)
        row += 1

        adv = ttk.LabelFrame(editor, text='Advanced', padding=8)
        adv.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(8, 10))
        ttk.Label(adv, text='Header row').grid(row=0, column=0, sticky='w', pady=2)
        ttk.Entry(adv, textvariable=self.header_row_var, width=10).grid(row=0, column=1, sticky='w', pady=2)
        ttk.Label(adv, text='Data start row').grid(row=1, column=0, sticky='w', pady=2)
        ttk.Entry(adv, textvariable=self.data_start_row_var, width=10).grid(row=1, column=1, sticky='w', pady=2)
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
        cols_frame.rowconfigure(0, weight=1)
        self.columns_canvas = tk.Canvas(cols_frame, height=220, highlightthickness=0)
        cols_scroll = ttk.Scrollbar(cols_frame, orient='vertical', command=self.columns_canvas.yview)
        self.columns_canvas.configure(yscrollcommand=cols_scroll.set)
        self.columns_canvas.grid(row=0, column=0, sticky='nsew')
        cols_scroll.grid(row=0, column=1, sticky='ns')
        self.columns_inner = ttk.Frame(self.columns_canvas)
        self.columns_inner.bind(
            '<Configure>',
            lambda _e: self.columns_canvas.configure(scrollregion=self.columns_canvas.bbox('all')),
        )
        self.columns_canvas.create_window((0, 0), window=self.columns_inner, anchor='nw')
        row += 1

        actions = ttk.Frame(editor)
        actions.grid(row=row, column=0, columnspan=2, sticky='ew', pady=(10, 0))
        ttk.Button(actions, text='Save task', command=self._save_task).pack(side='left')
        ttk.Button(actions, text='Run now', command=self._run_now).pack(side='left', padx=(8, 0))
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
        self.status_label.configure(text='Running' if self.service.is_running() else 'Stopped')
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

    def _load_headers(self):
        source_file = self.source_file_var.get().strip()
        source_sheet = self.source_sheet_var.get().strip()
        if not source_file or not source_sheet:
            return
        try:
            headers = list_headers(source_file, source_sheet, int(self.header_row_var.get() or '1'))
        except Exception as exc:
            messagebox.showerror('Load columns failed', str(exc))
            return
        current = {header for header, var in self.header_vars.items() if var.get()}
        for child in self.columns_inner.winfo_children():
            child.destroy()
        self.header_vars = {}
        for idx, header in enumerate(headers):
            var = tk.BooleanVar(value=header in current)
            self.header_vars[header] = var
            ttk.Checkbutton(self.columns_inner, text=header, variable=var).grid(
                row=idx, column=0, sticky='w', pady=2
            )

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
            self.target_file_var.set('')
            self.target_sheet_var.set('')
            self.header_row_var.set('1')
            self.data_start_row_var.set('2')
            self.formula_var.set('values')
            self.sheet_combo['values'] = []
            for child in self.columns_inner.winfo_children():
                child.destroy()
            self.header_vars = {}
            return

        self.name_var.set(task.name)
        self.enabled_var.set(task.enabled)
        self.source_file_var.set(task.source_file)
        self.source_sheet_var.set(task.source_sheet)
        self.target_file_var.set(task.target_file)
        self.target_sheet_var.set(task.target_sheet)
        self.header_row_var.set(str(task.header_row))
        self.data_start_row_var.set(str(task.data_start_row))
        self.formula_var.set(task.formula_handling)
        try:
            sheets = list_sheets(task.source_file) if task.source_file else []
        except Exception:
            sheets = []
        self.sheet_combo['values'] = sheets
        self._load_headers()
        for header, var in self.header_vars.items():
            var.set(header in task.columns_by_header)

    def _build_task(self) -> SyncTask:
        headers = [header for header, var in self.header_vars.items() if var.get()]
        return SyncTask(
            id=self.selected_task_id or SyncTask().id,
            name=self.name_var.get().strip() or 'New Task',
            enabled=self.enabled_var.get(),
            source_file=self.source_file_var.get().strip(),
            source_sheet=self.source_sheet_var.get().strip(),
            target_file=self.target_file_var.get().strip(),
            target_sheet=self.target_sheet_var.get().strip() or 'Export',
            columns_by_header=headers,
            header_row=int(self.header_row_var.get() or '1'),
            data_start_row=int(self.data_start_row_var.get() or '2'),
            formula_handling=self.formula_var.get().strip() or 'values',
        )

    def _save_task(self):
        try:
            task = self._build_task()
        except Exception as exc:
            messagebox.showerror('Save task failed', str(exc))
            return
        tasks = list(self.service.tasks)
        for idx, existing in enumerate(tasks):
            if existing.id == task.id:
                tasks[idx] = task
                break
        else:
            tasks.append(task)
        self.service.set_tasks(tasks)
        self.selected_task_id = task.id
        self.tree.selection_set(task.id)
        self._load_task_into_form(task)
        self._refresh()

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
        if not self.selected_task_id:
            messagebox.showinfo('Run task', 'Save the task first.')
            return
        self.service.run_task_now(self.selected_task_id)
        self._refresh()

    def _toggle_enabled(self, _event):
        selection = self.tree.selection()
        if selection:
            self.selected_task_id = selection[0]
        if not self.selected_task_id:
            return
        tasks = list(self.service.tasks)
        for idx, task in enumerate(tasks):
            if task.id == self.selected_task_id:
                replacement = SyncTask(**{**task.__dict__, 'enabled': not task.enabled})
                tasks[idx] = replacement
                self.service.set_tasks(tasks)
                self._load_task_into_form(replacement)
                self._refresh()
                return

    def _start(self):
        self.service.start()
        self._refresh()

    def _stop(self):
        self.service.stop()
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


