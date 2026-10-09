# CNC 程序单自动生成工具

Windows 桌面工具。将一个或多个 `.NC` 文件拖入窗口后，程序会按顺序解析并写入提供的 `CNC程序单.xlsx` 模板：

- 程序名称：`B7:B25`
- 刀具类型：`C7:C25`
- D：`D7:D25`
- R：`E7:E25`
- 加工 DATA：`F7:F25`
- 最低 Z 深度：`J7:J25`

程序不会保存或覆盖源模板。首次启动会把随程序附带的模板复制到 `%LOCALAPPDATA%\CNCProgramSheet\templates\CNC程序单.xlsx`；此后的升级只更新安装目录，绝不会写入这个用户数据目录。

## 本地运行

```powershell
py -3.12 -m pip install -r requirements.txt
py -3.12 main.py
```

拖入 `.NC` 文件后，可以在左侧解析表或右侧 Excel 预览中修正所有自动填写项。点击“完成并保存”会在 Windows 桌面生成一个不重名的新 `.xlsx` 文件。

## 图片插入

1. 在右侧 Excel 预览中点击图片要停靠的单元格。
2. 点击“插入图片”选择 PNG、JPG、JPEG、BMP 或 GIF；也可以直接按 `Ctrl+V` 粘贴 Windows 截图、复制的图片或资源管理器中复制的图片文件。
3. 设定导出尺寸后，预览会在锚点显示缩略图。导出的 Excel 会按相同单元格和尺寸插入原图。

插入或粘贴的图片会复制到 `%LOCALAPPDATA%\CNCProgramSheet\images`，所以原始图片被移动、剪贴板内容被替换或软件升级后，当前会话仍可正常导出。点击“删除图片”可移除当前选中单元格的所有插入图片。点击合并单元格中的任意位置时，程序会自动使用该合并区域的左上角作为 Excel 图片锚点。

## 解析原则

- 只有在程序头或刀具描述有明确依据时才填写刀具类型、D、R。
- `G41/G42 D1` 这类刀补号不会作为刀具直径。
- `G81/G83/G85` 等固定循环的 `R` 平面不会作为刀具半径。
- 扫描所有实际 Z 值，写入最小值。
- 证据不足时显示“未识别”，等待人工修改。

## 版本、GitHub Release 与自动更新

版本号只维护在 [cnc_program_sheet/version.py](cnc_program_sheet/version.py)。发布步骤：

1. 修改 `__version__`，例如 `1.0.1`，并提交代码。
2. 创建并推送匹配的 Git 标签，例如 `v1.0.1`。
3. `.github/workflows/release.yml` 会在 Windows runner 上运行测试、打包无黑窗 EXE、生成 Inno Setup 安装程序、计算 SHA-256 并发布 GitHub Release。

发布工作流将当前仓库名写入安装包内的更新源。软件启动后及运行期间每 6 小时会在后台从 GitHub Release 查询最新版本。发现新版本时，下载与 SHA-256 校验均在后台进行，安装前仍由用户确认，避免中断尚未保存的程序单。

数据隔离规则：

- 安装程序仅写入应用安装目录。
- 用户模板、`settings.json`、日志、更新下载文件均在 `%LOCALAPPDATA%\CNCProgramSheet`。
- 安装程序不包含该数据目录，卸载和升级也不会删除或覆盖它。
- 用户“打开模板”选择的外部 `.xlsx` 始终只读；导出始终另存为桌面的新文件。

首次推送到 GitHub 后，把仓库远程地址添加到本地仓库：

```powershell
git remote add origin https://github.com/<owner>/<repository>.git
git add .
git commit -m "feat: initial CNC program sheet desktop app"
git push -u origin main
```

## 验证

```powershell
py -3.12 -m pytest -q
```

测试涵盖模板单元格映射、模板只读复制、文件名防覆盖、刀具 D/R 安全识别、固定循环 R 平面排除和最低 Z 深度扫描。
