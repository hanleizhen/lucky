$ErrorActionPreference = 'Continue'

function Get-SafePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return '(未提供)' }
    $userProfile = [Environment]::GetFolderPath('UserProfile')
    if ($userProfile -and $Path.StartsWith($userProfile, [StringComparison]::OrdinalIgnoreCase)) {
        return '%USERPROFILE%' + $Path.Substring($userProfile.Length)
    }
    return $Path
}

function Get-UninstallEntries([string]$Pattern) {
    $locations = @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $found = @()
    foreach ($location in $locations) {
        $found += Get-ItemProperty -Path $location -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -like $Pattern }
    }
    return @($found | Sort-Object DisplayName, DisplayVersion, InstallLocation -Unique)
}

function Get-ValueOrMissing($Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '(未设置)' }
    return [string]$Value
}

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('CNC 智能程序单 / Mastercam 按钮安装诊断报告')
$lines.Add(('扫描时间：{0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
$lines.Add('')
$lines.Add('[系统]')
try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $lines.Add(('Windows：{0}，版本 {1}，{2}' -f $os.Caption, $os.Version, $os.OSArchitecture))
} catch {
    $lines.Add(('Windows 信息读取失败：{0}' -f $_.Exception.Message))
}

$lines.Add('')
$lines.Add('[CNC 智能程序单]')
$appEntries = Get-UninstallEntries '*CNC*程序单*'
if ($appEntries.Count -eq 0) {
    $lines.Add('安装注册信息：未找到')
} else {
    foreach ($entry in $appEntries) {
        $lines.Add(('安装项：{0}；版本：{1}；安装目录：{2}' -f $entry.DisplayName, (Get-ValueOrMissing $entry.DisplayVersion), (Get-SafePath $entry.InstallLocation)))
    }
}

$localAppData = [Environment]::GetFolderPath('LocalApplicationData')
$dataDirectory = Join-Path $localAppData 'CNCProgramSheet'
$settingsPath = Join-Path $dataDirectory 'settings.json'
$lines.Add(('设置文件：{0}' -f $(if (Test-Path -LiteralPath $settingsPath) { '存在' } else { '不存在' })))
if (Test-Path -LiteralPath $settingsPath) {
    try {
        $settings = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
        $lines.Add(('Mastercam 自动安装选项：{0}' -f $(if ($settings.install_mastercam_button) { '已勾选' } else { '未勾选' })))
        $lines.Add(('已记录安装的程序版本：{0}' -f (Get-ValueOrMissing $settings.mastercam_addin_installed_version)))
    } catch {
        $lines.Add(('设置文件无法解析：{0}' -f $_.Exception.Message))
    }
} else {
    $lines.Add('Mastercam 自动安装选项：无设置记录（默认关闭）')
}

$bridgePathFile = Join-Path $dataDirectory 'mastercam_bridge_path.txt'
if (Test-Path -LiteralPath $bridgePathFile) {
    $bridgeTarget = (Get-Content -LiteralPath $bridgePathFile -Raw -Encoding UTF8).Trim()
    $lines.Add(('Mastercam 按钮记录的程序路径：{0}' -f $(if (Test-Path -LiteralPath $bridgeTarget) { '存在' } else { '文件不存在' })))
} else {
    $lines.Add('Mastercam 按钮记录的程序路径：未生成')
}

$lines.Add('')
$lines.Add('[Mastercam 安装登记]')
$mcEntries = Get-UninstallEntries '*Mastercam*'
if ($mcEntries.Count -eq 0) {
    $lines.Add('卸载注册信息：未找到')
} else {
    foreach ($entry in $mcEntries) {
        $lines.Add(('登记项：{0}；版本：{1}；安装目录：{2}' -f $entry.DisplayName, (Get-ValueOrMissing $entry.DisplayVersion), (Get-SafePath $entry.InstallLocation)))
    }
}

$documents = [Environment]::GetFolderPath('MyDocuments')
$programRoots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432) |
    Where-Object { $_ } | Select-Object -Unique
$supportedYears = 2021..2027
$detectedCount = 0
$lines.Add('')
$lines.Add('[按钮安装目标检测（按程序使用的标准目录规则）]')
foreach ($year in $supportedYears) {
    $installCandidates = @()
    foreach ($root in $programRoots) {
        $installRoot = Join-Path $root "Mastercam $year"
        foreach ($exePath in @(
            (Join-Path $installRoot 'Mastercam.exe'),
            (Join-Path (Join-Path $installRoot 'Mastercam') 'Mastercam.exe')
        )) {
            if (Test-Path -LiteralPath $exePath -PathType Leaf) {
                $installCandidates += [PSCustomObject]@{
                    Directory = Split-Path -Parent $exePath
                    Executable = $exePath
                }
            }
        }
    }
    $profileMastercam = Join-Path (Join-Path $documents "My Mastercam $year") 'Mastercam'
    $addinDirectory = Join-Path $profileMastercam 'Add-Ins'
    $hasProfile = Test-Path -LiteralPath $profileMastercam
    if ($installCandidates.Count -gt 0) {
        $detectedCount++
        $lines.Add(('Mastercam {0}：程序文件已检测到；用户配置目录{1}；Add-Ins 目录{2}' -f $year, $(if ($hasProfile) { '存在' } else { '缺失' }), $(if (Test-Path -LiteralPath $addinDirectory) { '存在' } else { '不存在' })))
        foreach ($candidate in $installCandidates) {
            $lines.Add(('  程序目录：{0}' -f (Get-SafePath $candidate.Directory)))
            $lines.Add(('  可执行文件：{0}' -f (Get-SafePath $candidate.Executable)))
        }
        if (-not $hasProfile) { $lines.Add(('  需要的用户配置目录：{0}' -f (Get-SafePath $profileMastercam))) }
        if (Test-Path -LiteralPath $addinDirectory) {
            foreach ($fileName in @('GenerateProgramSheet.dll', 'GenerateProgramSheet.ft', 'GenerateProgramSheet.csx')) {
                $addinFile = Join-Path $addinDirectory $fileName
                if (Test-Path -LiteralPath $addinFile) {
                    $fileSize = (Get-Item -LiteralPath $addinFile).Length
                    $lines.Add(('  {0}：已存在，大小 {1} 字节' -f $fileName, $fileSize))
                } else {
                    $lines.Add(('  {0}：缺失' -f $fileName))
                }
            }
        }
    }
}
if ($detectedCount -eq 0) {
    $lines.Add('未检测到可自动安装的版本。程序支持 Program Files 下 Mastercam 年份目录中的 Mastercam.exe（也支持其 Mastercam 子目录），并要求对应的“我的文档\My Mastercam 年份\Mastercam”用户配置目录存在。')
}
$lines.Add(('用户文档位置类型：{0}' -f $(if ($documents) { '已定位（实际用户名路径已隐藏）' } else { '未定位' })))

$lines.Add('')
$lines.Add('[结论提示]')
if (-not (Test-Path -LiteralPath $settingsPath)) {
    $lines.Add('未找到用户设置文件。先打开 CNC 智能程序单一次，再在“设置 → 常规”勾选 Mastercam 自动安装并保存。')
} elseif (-not $settings.install_mastercam_button) {
    $lines.Add('Mastercam 自动安装选项当前未勾选。勾选、保存并重启 CNC 智能程序单后才会安装。')
} elseif ($detectedCount -eq 0) {
    $lines.Add('自动安装已启用，但程序没有找到符合标准目录规则的 Mastercam 安装与用户配置目录。')
} else {
    $lines.Add('自动安装选项已启用且发现标准安装目录。若按钮仍未显示，请检查上述 Add-Ins 文件状态，并在安装成功后重启 Mastercam。')
}
$lines.Add('')
$lines.Add('本报告仅检查系统版本、软件安装登记、自动安装设置及 Mastercam 标准目录/按钮文件；不读取 NC 文件、模板内容或剪贴板历史。')

$report = $lines -join [Environment]::NewLine
$desktop = [Environment]::GetFolderPath('Desktop')
if (-not $desktop) { $desktop = $PSScriptRoot }
$reportPath = Join-Path $desktop 'Mastercam诊断报告.txt'
Set-Content -LiteralPath $reportPath -Value $report -Encoding UTF8

$copied = $false
try {
    Set-Clipboard -Value $report -ErrorAction Stop
    $copied = $true
} catch {
    $copied = $false
}

Write-Host ''
Write-Host ('报告已保存：{0}' -f $reportPath) -ForegroundColor Green
if ($copied) {
    Write-Host '报告内容已复制到剪贴板；可以直接回到聊天窗口按 Ctrl+V 粘贴。' -ForegroundColor Green
} else {
    Write-Host '自动复制失败；请打开报告文件，全选并复制，再粘贴到聊天窗口。' -ForegroundColor Yellow
}
Write-Host ''
Write-Host '按任意键关闭窗口……'
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
