"""Single source of truth for the application version."""

__version__ = "1.1.47"
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
    "新增：安装或升级后，首次打开会显示本次更新内容；同一版本只显示一次。",
    "优化：拖动表格列宽后按 Ctrl+S 保存到当前模板，重启和导出都会保持。",
)

