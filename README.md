# Excel Sync Manager V1

This folder contains a GUI-first version of the Excel sync tool.

## What it does

- Lets you create, edit, enable, disable, and delete sync tasks in a desktop window
- Watches source Excel files in the background
- Syncs selected columns from a source sheet into a target workbook
- Exports formula display values by default, which is safer for outward-facing files
- Retries automatically when WPS or Excel is locking the source or target file

## How to use

1. Double-click `start_manager.cmd`
2. Click `New task`
3. Pick the source workbook
4. Click `Load sheets`
5. Choose the source sheet
6. Pick the columns to export
7. Pick the target workbook path
8. Save the task
9. Click `Start monitoring`

## Notes

- Task data is stored in `tasks.json`, but the main workflow is through the GUI
- Log output is written to `_runtime\manager.log`
- Double-click a task row to quickly toggle it on or off
