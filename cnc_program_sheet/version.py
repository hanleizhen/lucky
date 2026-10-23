"""Single source of truth for the application version."""

__version__ = "1.1.49"
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
    "新增：检测同一 NC 文件中同一刀号的多种 S 转速；相关程序行红色填充提示，导出 Excel 保留标记。",
    "提示：同一版本只显示一次。",
)

