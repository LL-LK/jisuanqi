; Inno Setup 6 Script
; 本地考勤工具安装脚本
;
; 设计要点：
; 1. AppId: 唯一 GUID 标识，升级覆盖不丢失配置；
; 2. 权限自适应: PrivilegesRequired=lowest，普通师生无管理员权限可安装至 %LOCALAPPDATA%，管理员运行安装至 Program Files；
; 3. 词典保护: 免勤词典.xlsx 标记 onlyifdoesntexist uninsneveruninstall，升级/重装/卸载均绝不冲掉用户自定义规则；
; 4. 现代化向导: WizardStyle=modern，应用图标 assets\attendance.ico，关闭运行中进程提示 (CloseApplications=yes)；
; 5. 算法口径: 纯净出勤率公式（学院总实到 / (学院总应到 - 学院总免勤)），全量扣减免勤人数。

#ifndef MyAppVersion
  #define MyAppVersion "1.0.8"
#endif

#define MyAppName "本地考勤工具"
#define MyAppVerName "本地考勤工具"
#define MyAppPublisher "自习考勤自动化工作组 / LL-LK"
#define MyAppURL "https://github.com/LL-LK/sanzao-wan-attendance-web"
#define MyAppExeName "三早一晚免勤自动统计.exe"

[Setup]
; 核心唯一标识
AppId={{8B7C9D2E-4F1A-492B-9B8A-7D1E5C3B8A4F}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppVerName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases

; 默认安装目录：{autopf} 根据当前运行权限自适应
; - 非管理员运行: {localappdata}\Programs\三早一晚免勤自动统计
; - 管理员运行: {commonpf}\三早一晚免勤自动统计 (或 64 位 Program Files)
DefaultDirName={autopf}\三早一晚免勤自动统计
DefaultGroupName={#MyAppName}
AllowNoIcons=yes

; 产物输出路径与文件名规范
OutputDir=发布包
OutputBaseFilename=三早一晚免勤自动统计_Setup
SetupIconFile=assets\attendance.ico

; 高压缩比配置
Compression=lzma2/ultra64
SolidCompression=yes

; 现代化向导外观与架构支持
WizardStyle=modern
WizardResizable=no
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
CloseApplicationsFilter=*.exe

; 卸载程序配置
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppVerName} (卸载)

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
; 核心可执行文件与运行时
Source: "发布版\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; 用户自定义词典与规则表 —— 绝不覆盖已有文件，卸载时不删除用户积累的规则与别名！
Source: "发布版\免勤词典.xlsx"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
; 使用说明与相关文档
Source: "发布版\使用说明.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "发布版\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "发布版\隐私扫描报告.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "发布版\SHA256.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; 开始菜单快捷方式
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
; 桌面快捷方式（可选任务，默认勾选）
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\{#MyAppExeName}"

[Run]
; 安装完成后可选立即启动程序
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
