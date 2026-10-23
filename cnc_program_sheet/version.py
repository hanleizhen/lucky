"""Single source of truth for the application version."""

__version__ = "1.1.48"
APP_NAME = "CNC 智能程序单"
APP_ICON_FILE = "cnc_smart_sheet_icon.ico"
# Keep this identifier stable across releases.  Windows uses it to associate
# the running process, taskbar button and shortcuts with this application.
APP_USER_MODEL_ID = "CNCProgramSheet.CNCSmartSheet"

# This is deliberately packaged with each installer instead of fetched from
# the internet at startup: a customer can always see what changed even when
# their network is unavailable immediately after an update.  Update this list
# whenever ``__version__`` changes.
RELEASE_NOTES: tuple[str, ...] = (
    "新增：已导入 NC 文件支持多选删除，可右键移除或按 Delete；只从当前程序单移除，原始文件保留。",
    "修复：表格空白单元格首次点击即可选中；点击表格外取消选区，Esc 仍可取消。",
    "提示：同一版本只显示一次。",
)

