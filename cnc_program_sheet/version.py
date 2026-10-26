"""Single source of truth for the application version."""

__version__ = "1.1.57"
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
    "新增：可在设置中指定 Mastercam NC 默认文件夹，按钮会优先从该目录打开 NC 选择窗口。",
    "修复：Mastercam 安装在嵌套目录时，按钮弹出的 NC 文件选择窗口仍默认打开对应版本的 Mill\\NC 文件夹。",
    "修复：兼容 Mastercam.exe 位于产品子目录中的标准安装布局，解决 Mastercam 2022 未被自动检测的问题。",
    "Mastercam 按钮自动安装仍由设置中的用户选项控制；勾选并重启后会检测并安装到已发现的 Mastercam 2021–2027。",
    "修复：从 Mastercam 重复启动时复用已运行的程序单窗口，避免隐藏图标区重复出现多个托盘图标。",
    "新增：检测同一 NC 文件中同一刀号的多种 S 转速；相关程序行红色填充提示，导出 Excel 保留标记。",
    "提示：同一版本只显示一次。",
)

