from __future__ import annotations

import copy
import json
import logging
import threading
import time
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Callable

import openpyxl
from openpyxl.utils import get_column_letter


APP_DIR = Path(__file__).resolve().parent
DATA_PATH = APP_DIR / 'tasks.json'
LOG_DIR = APP_DIR / '_runtime'
LOG_PATH = LOG_DIR / 'manager.log'

DEFAULT_SETTINGS = {
    'debounce_seconds': 5.0,
    'post_save_delay_seconds': 1.0,
    'read_retry_count': 3,
    'read_retry_delay_seconds': 1.2,
    'scan_interval_seconds': 2.0,
    'retry_locked_file_seconds': 5.0,
}


@dataclass
class SyncTask:
    id: str = field(default_factory=lambda: uuid.uuid4().hex)
    name: str = 'New Task'
    enabled: bool = True
    source_file: str = ''
    source_sheet: str = ''
    target_file: str = ''
    target_sheet: str = ''
    columns_by_header: list[str] = field(default_factory=list)
    header_row: int = 1
    data_start_row: int = 2
    copy_style: bool = True
    copy_column_widths: bool = True
    copy_row_heights: bool = True
    include_header: bool = True
    drop_empty_rows: bool = True
    formula_handling: str = 'values'


@dataclass
class TaskRuntime:
    status: str = 'idle'
    last_startsync_at: str = ''
    last_synced_at: str = ''
    last_error: str = ''
    last_change_at: str = ''
    pending_retry: bool = False


def setup_logging() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(LOG_PATH, encoding='utf-8'),
            logging.StreamHandler(),
        ],
    )


def load_data() -> tuple[dict, list[SyncTask]]:
    if not DATA_PATH.exists():
        save_data(DEFAULT_SETTINGS.copy(), [])
        return DEFAULT_SETTINGS.copy(), []

    raw = json.loads(DATA_PATH.read_text(encoding='utf-8'))
    settings = DEFAULT_SETTINGS.copy()
    settings.update(raw.get('settings', {}))
    tasks = [SyncTask(**item) for item in raw.get('tasks', [])]
    return settings, tasks


def save_data(settings: dict, tasks: list[SyncTask]) -> None:
    payload = {
        'settings': settings,
        'tasks': [asdict(task) for task in tasks],
    }
    DATA_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')


def list_sheets(source_file: str) -> list[str]:
    workbook = openpyxl.load_workbook(source_file, read_only=True, data_only=False)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def list_headers(source_file: str, sheet_name: str, header_row: int) -> list[str]:
    workbook = openpyxl.load_workbook(source_file, read_only=True, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            return []
        ws = workbook[sheet_name]
        headers: list[str] = []
        for col_idx in range(1, ws.max_column + 1):
            value = ws.cell(header_row, col_idx).value
            if value not in (None, ''):
                headers.append(str(value).strip())
        return headers
    finally:
        workbook.close()


def _copy_cell_style(src_cell, dst_cell) -> None:
    dst_cell.font = copy.copy(src_cell.font)
    dst_cell.fill = copy.copy(src_cell.fill)
    dst_cell.border = copy.copy(src_cell.border)
    dst_cell.alignment = copy.copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format
    dst_cell.protection = copy.copy(src_cell.protection)


def _row_is_empty(ws, row_idx: int, source_columns: list[int]) -> bool:
    for source_col in source_columns:
        if ws.cell(row_idx, source_col).value not in (None, ''):
            return False
    return True


def _apply_column_widths(src_ws, dst_ws, source_columns: list[int]) -> None:
    for dst_col_idx, source_col_idx in enumerate(source_columns, start=1):
        source_letter = get_column_letter(source_col_idx)
        target_letter = get_column_letter(dst_col_idx)
        dst_ws.column_dimensions[target_letter].width = src_ws.column_dimensions[source_letter].width
        dst_ws.column_dimensions[target_letter].hidden = src_ws.column_dimensions[source_letter].hidden


def _apply_row_height(src_ws, dst_ws, source_row_idx: int, target_row_idx: int) -> None:
    src_dim = src_ws.row_dimensions[source_row_idx]
    if src_dim.height is not None:
        dst_ws.row_dimensions[target_row_idx].height = src_dim.height
    dst_ws.row_dimensions[target_row_idx].hidden = src_dim.hidden


def _clone_merged_cells(src_ws, dst_ws, source_columns: list[int], target_row_map: dict[int, int]) -> None:
    source_col_map = {source_col: dst_col for dst_col, source_col in enumerate(source_columns, start=1)}
    for merged_range in src_ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        if any(col not in source_col_map for col in range(min_col, max_col + 1)):
            continue
        if min_row not in target_row_map or max_row not in target_row_map:
            continue
        dst_ws.merge_cells(
            start_row=target_row_map[min_row],
            start_column=source_col_map[min_col],
            end_row=target_row_map[max_row],
            end_column=source_col_map[max_col],
        )


def _get_selected_columns(ws, task: SyncTask) -> list[int]:
    header_map: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        header_value = ws.cell(task.header_row, col_idx).value
        if header_value in (None, ''):
            continue
        header_map[str(header_value).strip()] = col_idx

    selected: list[int] = []
    missing: list[str] = []
    for header in task.columns_by_header:
        if header in header_map:
            selected.append(header_map[header])
        else:
            missing.append(header)
    if missing:
        raise ValueError(f"Missing headers: {', '.join(missing)}")
    return selected


def _export_value(src_formula_cell, src_value_cell, task: SyncTask):
    if task.formula_handling == 'formulas':
        return src_formula_cell.value
    if src_formula_cell.data_type == 'f':
        return src_value_cell.value
    return src_formula_cell.value


def sync_task(task: SyncTask, settings: dict) -> Path:
    if not task.source_file or not task.source_sheet:
        raise ValueError('Source file and source sheet are required.')
    if not task.target_file or not task.target_sheet:
        raise ValueError('Target file and target sheet are required.')
    if not task.columns_by_header:
        raise ValueError('Pick at least one column.')

    source_path = Path(task.source_file).resolve()
    if source_path.name.startswith('~$'):
        raise ValueError('Choose the real workbook, not the temporary lock file.')
    target_path = Path(task.target_file).resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)

    post_save_delay = float(settings.get('post_save_delay_seconds', 0.0))
    if post_save_delay > 0:
        time.sleep(post_save_delay)

    retry_count = int(settings.get('read_retry_count', 3))
    retry_delay = float(settings.get('read_retry_delay_seconds', 1.2))
    last_error: Exception | None = None

    for _ in range(retry_count):
        wb = None
        wb_values = None
        try:
            wb = openpyxl.load_workbook(source_path, data_only=False)
            wb_values = openpyxl.load_workbook(source_path, data_only=True)
            if task.source_sheet not in wb.sheetnames:
                raise ValueError(f'Source sheet not found: {task.source_sheet}')

            src_ws = wb[task.source_sheet]
            src_ws_values = wb_values[task.source_sheet]
            source_columns = _get_selected_columns(src_ws, task)

            out_wb = openpyxl.Workbook()
            dst_ws = out_wb.active
            dst_ws.title = task.target_sheet

            target_row_idx = 1
            target_row_map: dict[int, int] = {}

            if task.include_header:
                target_row_map[task.header_row] = target_row_idx
                for dst_col_idx, source_col_idx in enumerate(source_columns, start=1):
                    src_cell = src_ws.cell(task.header_row, source_col_idx)
                    src_value_cell = src_ws_values.cell(task.header_row, source_col_idx)
                    dst_cell = dst_ws.cell(
                        target_row_idx,
                        dst_col_idx,
                        _export_value(src_cell, src_value_cell, task),
                    )
                    if task.copy_style:
                        _copy_cell_style(src_cell, dst_cell)
                if task.copy_row_heights:
                    _apply_row_height(src_ws, dst_ws, task.header_row, target_row_idx)
                target_row_idx += 1

            for source_row_idx in range(task.data_start_row, src_ws.max_row + 1):
                if task.drop_empty_rows and _row_is_empty(src_ws, source_row_idx, source_columns):
                    continue
                target_row_map[source_row_idx] = target_row_idx
                for dst_col_idx, source_col_idx in enumerate(source_columns, start=1):
                    src_cell = src_ws.cell(source_row_idx, source_col_idx)
                    src_value_cell = src_ws_values.cell(source_row_idx, source_col_idx)
                    dst_cell = dst_ws.cell(
                        target_row_idx,
                        dst_col_idx,
                        _export_value(src_cell, src_value_cell, task),
                    )
                    if task.copy_style:
                        _copy_cell_style(src_cell, dst_cell)
                if task.copy_row_heights:
                    _apply_row_height(src_ws, dst_ws, source_row_idx, target_row_idx)
                target_row_idx += 1

            if task.copy_column_widths:
                _apply_column_widths(src_ws, dst_ws, source_columns)
            if task.copy_style:
                _clone_merged_cells(src_ws, dst_ws, source_columns, target_row_map)

            out_wb.save(target_path)
            return target_path
        except PermissionError as exc:
            last_error = exc
            time.sleep(retry_delay)
        finally:
            if wb is not None:
                wb.close()
            if wb_values is not None:
                wb_values.close()

    if last_error is not None:
        raise last_error
    raise RuntimeError('Sync failed.')


class SyncService:
    def __init__(self, status_callback: Callable[[], None] | None = None):
        setup_logging()
        self.settings, tasks = load_data()
        self.tasks: list[SyncTask] = tasks
        self.runtime: dict[str, TaskRuntime] = {task.id: TaskRuntime() for task in tasks}
        self.status_callback = status_callback
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._lock = threading.RLock()
        self._file_signatures: dict[str, tuple[int, int]] = {}
        self._pending_since: dict[str, float] = {}
        self._pending_retry_at: dict[str, float] = {}

    def save(self) -> None:
        with self._lock:
            save_data(self.settings, self.tasks)

    def set_tasks(self, tasks: list[SyncTask]) -> None:
        with self._lock:
            self.tasks = tasks
            existing_runtime = self.runtime
            self.runtime = {task.id: existing_runtime.get(task.id, TaskRuntime()) for task in tasks}
            self.save()
        self._notify()

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()
        logging.info('GUI sync service started.')

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=3)
        logging.info('GUI sync service stopped.')

    def is_running(self) -> bool:
        return bool(self._thread and self._thread.is_alive())

    def run_task_now(self, task_id: str) -> None:
        task = self.get_task(task_id)
        if task is None:
            return
        self._sync_one(task, manual=True)

    def get_task(self, task_id: str) -> SyncTask | None:
        with self._lock:
            for task in self.tasks:
                if task.id == task_id:
                    return task
        return None

    def list_runtime_rows(self) -> list[tuple[SyncTask, TaskRuntime]]:
        with self._lock:
            return [(task, self.runtime.get(task.id, TaskRuntime())) for task in self.tasks]

    def _notify(self) -> None:
        if self.status_callback is not None:
            self.status_callback()

    def _mark_status(self, task_id: str, **updates) -> None:
        with self._lock:
            runtime = self.runtime.setdefault(task_id, TaskRuntime())
            for key, value in updates.items():
                setattr(runtime, key, value)
        self._notify()

    def _run_loop(self) -> None:
        while not self._stop_event.is_set():
            self._scan_for_changes()
            self._process_pending()
            self._process_retries()
            self._stop_event.wait(float(self.settings.get('scan_interval_seconds', 2.0)))

    def _scan_for_changes(self) -> None:
        now = time.monotonic()
        with self._lock:
            tasks = list(self.tasks)
        for task in tasks:
            if not task.enabled or not task.source_file:
                continue
            source_path = Path(task.source_file)
            if not source_path.exists() or source_path.name.startswith('~$'):
                continue
            try:
                stat = source_path.stat()
            except OSError:
                continue
            signature = (stat.st_mtime_ns, stat.st_size)
            previous = self._file_signatures.get(task.id)
            self._file_signatures[task.id] = signature
            if previous is None:
                self._pending_since.setdefault(task.id, now)
                continue
            if signature != previous:
                self._pending_since[task.id] = now
                self._mark_status(
                    task.id,
                    status='change detected',
                    last_change_at=time.strftime('%Y-%m-%d %H:%M:%S'),
                )
                logging.info("Detected file change for task '%s'", task.name)

    def _process_pending(self) -> None:
        now = time.monotonic()
        debounce = float(self.settings.get('debounce_seconds', 5.0))
        due = [task_id for task_id, changed_at in self._pending_since.items() if now - changed_at >= debounce]
        for task_id in due:
            self._pending_since.pop(task_id, None)
            task = self.get_task(task_id)
            if task is None or not task.enabled:
                continue
            self._sync_one(task, manual=False)

    def _process_retries(self) -> None:
        now = time.monotonic()
        due = [task_id for task_id, retry_at in self._pending_retry_at.items() if now >= retry_at]
        for task_id in due:
            self._pending_retry_at.pop(task_id, None)
            task = self.get_task(task_id)
            if task is None or not task.enabled:
                continue
            self._sync_one(task, manual=False)

    def _sync_one(self, task: SyncTask, manual: bool = False) -> None:
        updates = {
            'status': 'syncing',
            'pending_retry': False,
            'last_error': '',
        }
        if manual:
            updates['last_startsync_at'] = time.strftime('%Y-%m-%d %H:%M:%S')
        self._mark_status(task.id, **updates)
        try:
            target_path = sync_task(task, self.settings)
            try:
                synced_at = target_path.stat().st_mtime
                timestamp = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(synced_at))
            except OSError:
                timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
            self._mark_status(
                task.id,
                status=f'ok -> {target_path.name}',
                last_synced_at=timestamp,
                last_error='',
                pending_retry=False,
            )
            logging.info("Task '%s' synced to %s", task.name, target_path)
        except PermissionError:
            retry_delay = float(self.settings.get('retry_locked_file_seconds', 5.0))
            self._pending_retry_at[task.id] = time.monotonic() + retry_delay
            self._mark_status(
                task.id,
                status='waiting for file unlock',
                last_error='Source or target file is locked by Excel/WPS.',
                pending_retry=True,
            )
            logging.warning("Task '%s' is waiting for file unlock.", task.name)
        except Exception as exc:
            self._mark_status(task.id, status='error', last_error=str(exc), pending_retry=False)
            logging.exception("Task '%s' failed", task.name)


