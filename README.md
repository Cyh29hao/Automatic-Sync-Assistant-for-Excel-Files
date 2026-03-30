# Excel Sync Manager V1

This folder contains the GUI-first version of the Excel sync tool.

## Features

- Create, edit, enable, disable, and delete sync tasks in a desktop window
- Watch source Excel files in the background
- Sync selected columns from a source sheet into a target workbook
- Export formula display values by default
- Retry automatically when WPS or Excel is locking the source or target file

## Versioning

- The single source of truth for app version is `metadata.py`
- Change only `APP_VERSION` there when you want a new release

## Run In Development

1. Double-click `start_manager.cmd`
2. Click `New task`
3. Pick the source workbook
4. Click `Load sheets`
5. Choose the source sheet
6. Pick the columns to export
7. Pick the target workbook path
8. Save the task
9. Click `Start monitoring`

## Build A Release

1. Make sure Python dependencies are installed from `requirements-dev.txt`
2. Update `APP_VERSION` in `metadata.py` if needed
3. Run `build_release.cmd` or `build_release.ps1`
4. After build:
   - `dist\Excel Sync Manager` is the raw PyInstaller output
   - `release\Excel Sync Manager vX.Y` is the clean release folder
   - `release\Excel Sync Manager vX.Y.zip` is the shareable zip

## Release Layout

- `Excel Sync Manager.exe`: the GUI app for end users
- `tasks.json`: saved tasks and settings, stored next to the exe
- `_runtime\manager.log`: runtime log
- `README_RELEASE.md`: short end-user instructions

## Notes

- Development mode reads and writes `tasks.json` in this folder
- Packaged mode reads and writes `tasks.json` next to the exe
- For end users, send the versioned release folder or zip from `release`, not the raw `dist` folder
