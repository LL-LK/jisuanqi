# tools/create_shortcuts.ps1
# 三早一晚免勤自动统计系统 - PowerShell 快捷方式创建脚本
# Milestone M3 (Shortcut Provisioning & Windows Installer Packaging)

[CmdletBinding()]
param(
    [string]$TargetDir   = "D:\Software\三早一晚免勤自动统计",
    [string]$ExeName     = "三早一晚免勤自动统计.exe",
    [string]$AppName     = "三早一晚免勤自动统计",
    [string]$Description = "三早一晚免勤自动统计系统 - 本地考勤去重与计算工具",
    [switch]$DesktopOnly,
    [switch]$StartMenuOnly
)

$ErrorActionPreference = "Stop"

$TargetExe = Join-Path $TargetDir $ExeName

if (-not (Test-Path $TargetExe)) {
    Write-Error "[ERROR] 目标可执行文件未找到: $TargetExe"
    exit 1
}

$WshShell = New-Object -ComObject WScript.Shell
$CreatedCount = 0

# 1. 桌面快捷方式
if (-not $StartMenuOnly) {
    $DesktopDir = [Environment]::GetFolderPath("Desktop")
    if (-not (Test-Path $DesktopDir)) {
        New-Item -ItemType Directory -Path $DesktopDir -Force | Out-Null
    }
    $DesktopLnk = Join-Path $DesktopDir "$AppName.lnk"
    $ShortcutD = $WshShell.CreateShortcut($DesktopLnk)
    $ShortcutD.TargetPath = $TargetExe
    $ShortcutD.WorkingDirectory = $TargetDir
    $ShortcutD.IconLocation = "$TargetExe,0"
    $ShortcutD.Description = $Description
    $ShortcutD.WindowStyle = 1
    $ShortcutD.Save()

    if (Test-Path $DesktopLnk) {
        Write-Host "[SUCCESS] 桌面快捷方式已创建: $DesktopLnk"
        $CreatedCount++
    } else {
        Write-Error "[ERROR] 桌面快捷方式创建失败，文件未生成: $DesktopLnk"
        exit 1
    }
}

# 2. 开始菜单快捷方式
if (-not $DesktopOnly) {
    $ProgramsDir = [Environment]::GetFolderPath("Programs")
    if (-not (Test-Path $ProgramsDir)) {
        New-Item -ItemType Directory -Path $ProgramsDir -Force | Out-Null
    }
    $StartMenuLnk = Join-Path $ProgramsDir "$AppName.lnk"
    $ShortcutP = $WshShell.CreateShortcut($StartMenuLnk)
    $ShortcutP.TargetPath = $TargetExe
    $ShortcutP.WorkingDirectory = $TargetDir
    $ShortcutP.IconLocation = "$TargetExe,0"
    $ShortcutP.Description = $Description
    $ShortcutP.WindowStyle = 1
    $ShortcutP.Save()

    if (Test-Path $StartMenuLnk) {
        Write-Host "[SUCCESS] 开始菜单快捷方式已创建: $StartMenuLnk"
        $CreatedCount++
    } else {
        Write-Error "[ERROR] 开始菜单快捷方式创建失败，文件未生成: $StartMenuLnk"
        exit 1
    }
}

Write-Host "`n[SUCCESS] 全部快捷方式创建完成！共生成 $CreatedCount 个快捷方式。"
exit 0
