"""Single source of truth for the application version."""

__version__ = "1.1.59"
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
    "新增：Mastercam 可按 NC 起始文件名查找连续程序组；文件名前缀精确匹配、按数字顺序处理，组大小可在设置中调整。",
    "新增：设置中可选择打印清晰度：标准、高清或最高。",
    "优化：每次升级的首次启动只显示本版本更新内容，不再重复列出上次版本的内容。",
    "新增：可在设置中控制 Mastercam 按钮安装成功后是否弹出提醒；默认不弹出成功提示。",
    "优化：升级日志按版本只显示一次。",
)

