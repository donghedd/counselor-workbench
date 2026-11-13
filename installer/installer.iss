; ============================================================
;  辅导员工作台 · 安装包脚本（Inno Setup 6）
;  编译：installer\编译安装包.bat   或   ISCC.exe installer.iss
;  产物：dist\辅导员工作台-v2.1.0-安装版.exe（另出一份 ASCII 名，方便 Release 附件）
; ============================================================

#define AppName "辅导员工作台"
; 版本号由 编译安装包.bat 从 version.py 读出后传入：
;   ISCC.exe /DAppVersion=2.1.0 installer.iss
; 直接双击编译时用下面的兜底值。
#ifndef AppVersion
  #define AppVersion "2.1.0"
#endif
#define AppExe "辅导员工作台.exe"
#define AppId "{{8F3D6A21-4C7E-4B95-9D2A-1E7F5C3B8A06}"

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=辅导员工作台
VersionInfoVersion={#AppVersion}.0
VersionInfoDescription={#AppName} 安装程序
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}

; 装到当前用户目录，双击即装、不弹 UAC；数据目录必须可写，所以不放 Program Files
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

; 一律装到上面的固定位置，**不复用上次安装目录**。
; 原因：Inno 会把上次的安装路径记在注册表（同一 AppId 共用一条卸载记录），
; 应用内「检查更新」时安装包是静默启动的、没人选目录，于是新版本会被装到
; 上次那个目录里。如果上次装在了临时目录（比如测试安装），更新就会一直
; “装成功”，但桌面/开始菜单指向的还是老目录，版本永远不变。
UsePreviousAppDir=no

OutputDir=..\dist
OutputBaseFilename={#AppName}-v{#AppVersion}-安装版
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\app.ico
UninstallDisplayName={#AppName}
LicenseFile=assets\使用须知.txt
WizardStyle=modern
DisableWelcomePage=no
WizardImageFile=assets\wizard.bmp,assets\wizard@2x.bmp
WizardSmallImageFile=assets\wizard_small.bmp,assets\wizard_small@2x.bmp
WizardImageStretch=yes
Compression=lzma2/max
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; 程序在跑的时候覆盖安装：先尝试关掉，不强行重启
CloseApplications=yes
RestartApplications=no
AllowNoIcons=yes
MinVersion=10.0

[Languages]
; 中文系统走简体中文；其它语言的系统自动用英文向导，避免字体/排版异常
Name: "cn"; MessagesFile: "assets\ChineseSimplified.isl"
Name: "en"; MessagesFile: "compiler:Default.isl"

[Messages]
cn.WelcomeLabel2=将把 [name/ver] 安装到这台电脑。%n%n默认装在当前用户目录，不需要管理员权限；如果之前装过，原有的学生数据（data、photos_src、备份）不会被覆盖。%n%n继续前建议关闭其他正在运行的程序，然后点「下一步」。

[Tasks]
Name: "desktopicon"; Description: "在桌面创建快捷方式"; Flags: checkedonce

[Dirs]
; 用户数据目录：卸载时一律保留（是否删除由卸载流程单独询问）
Name: "{app}\data"; Flags: uninsneveruninstall
Name: "{app}\photos"; Flags: uninsneveruninstall
Name: "{app}\photos_src"; Flags: uninsneveruninstall
Name: "{app}\备份"; Flags: uninsneveruninstall

[Files]
; 程序本体（exe + py + 模板 + 文档 + 说明）
Source: "..\dist\辅导员工作台-v{#AppVersion}\*"; DestDir: "{app}"; \
    Flags: ignoreversion; \
    Excludes: "data\*,photos\*,photos_src\*,备份\*,*.db,dashboard.html,dashboard_脱敏外发版.html,工作台数据汇总.xlsx,__pycache__\*"

; 空数据表：只在文件不存在时铺一份表头，绝不覆盖已有数据
Source: "..\dist\辅导员工作台-v{#AppVersion}\data\*"; DestDir: "{app}\data"; \
    Flags: onlyifdoesntexist uninsneveruninstall

; 图标单独放一份，保证快捷方式与「应用和功能」里显示正常
Source: "assets\app.ico"; DestDir: "{app}"; DestName: "app.ico"; Flags: ignoreversion

[InstallDelete]
; 看板是生成物：升级时清掉，首启动按新模板重画，避免看到旧版界面
Type: files; Name: "{app}\dashboard.html"
Type: files; Name: "{app}\dashboard_脱敏外发版.html"
Type: filesandordirs; Name: "{app}\__pycache__"

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; WorkingDir: "{app}"; IconFilename: "{app}\app.ico"; Comment: "打开辅导员工作台"
Name: "{autoprograms}\{#AppName} 使用说明"; Filename: "{app}\使用说明.md"; WorkingDir: "{app}"
Name: "{autoprograms}\卸载 {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; WorkingDir: "{app}"; IconFilename: "{app}\app.ico"; Check: WantDesktopIcon

[UninstallDelete]
Type: filesandordirs; Name: "{app}\__pycache__"
Type: files; Name: "{app}\dashboard.html"
Type: files; Name: "{app}\dashboard_脱敏外发版.html"
Type: files; Name: "{app}\工作台数据汇总.xlsx"
Type: files; Name: "{app}\启动日志.txt"
Type: filesandordirs; Name: "{app}\webview-profile"
Type: files; Name: "{app}\photo_index.json"
Type: filesandordirs; Name: "{app}\导出_脱敏版"
Type: dirifempty; Name: "{app}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "立即启动 {#AppName}"; Flags: nowait postinstall skipifsilent; Check: NotSilentInstall

[Code]
var
  DataKept: Boolean;

{ 安装开始前，先把还在运行的旧程序结束掉。
  为什么要自己动手：程序用 PyInstaller 打包，是「引导进程 + 应用进程」两层。
  关窗口时偶尔会留下一个没有窗口、也不退出的引导进程（日志和 %TEMP%\_MEI* 残留可证），
  它占着 exe 让 Restart Manager 关不掉，安装就卡在「正在关闭应用程序…」，
  过一会儿弹出「安装程序无法自动关闭所有应用程序」，得手动点「重试」才能继续。
  这里直接在它之前静默结束同名进程，装机就不会停在那一步。
  直接调 taskkill.exe（不经 cmd），中文进程名不会因代码页出错；没在运行也无害。

  注意：**不能用 /T（结束进程树）**。应用内的「检查更新」是由主程序启动安装包，
  安装包此时是「辅导员工作台.exe」的子进程，/T 会把安装包自己也一起杀掉——
  表现就是点「立即更新」后窗口一闪、什么都没装。
  /F /IM 按映像名就能把引导进程和应用进程都结束，不需要 /T。 }
function InitializeSetup(): Boolean;
var
  RC: Integer;
begin
  Result := True;
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM "{#AppExe}"',
       '', SW_HIDE, ewWaitUntilTerminated, RC);
  { 等文件句柄释放，避免紧接着的复制报「文件正在使用」 }
  Sleep(1200);
end;

{ 目标目录里已经有像样的学生数据吗？（表头之外还有行就算有） }
function HasExistingData(): Boolean;
var
  Lines: TArrayOfString;
  P: String;
begin
  Result := False;
  P := ExpandConstant('{app}\data\01_学生主表.csv');
  if not FileExists(P) then
    Exit;
  if LoadStringsFromFile(P, Lines) then
    Result := GetArrayLength(Lines) > 1;
end;

procedure InitializeWizard();
begin
  DataKept := False;
end;

{ 静默安装/命令行安装时不自动拉起程序，避免批处理里挂住 }
function NotSilentInstall(): Boolean;
begin
  Result := not WizardSilent();
end;

{ 桌面快捷方式要不要建。
  不能只写 Tasks: desktopicon：那样只有「首次安装」勾选过才建，
  升级时任务不再勾选，旧快捷方式就一直指向老目录（换过安装路径就会指错）。
  这里改成：勾了任务就建；或者桌面上已经有快捷方式（说明用户想留着），也重建一遍指向当前目录。
  静默升级时正好走第二条，顺手把指错的快捷方式纠正回来。 }
function WantDesktopIcon(): Boolean;
begin
  Result := WizardIsTaskSelected('desktopicon') or
            FileExists(ExpandConstant('{autodesktop}\{#AppName}.lnk'));
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = wpSelectDir) and HasExistingData() then
  begin
    if MsgBox('检测到这个位置已经有学生数据（data 目录里有内容）。' + #13#10#13#10 +
              '继续安装只会更新程序文件，不会动你的数据，看板会在下次启动时重新生成。' + #13#10#13#10 +
              '是否继续？', mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Answer: Integer;
begin
  if CurUninstallStep <> usUninstall then
    Exit;
  if UninstallSilent then
    Exit;

  Answer := MsgBox('是否同时删除本机保存的学生数据？' + #13#10#13#10 +
                   '包括：' + #13#10 +
                   '  · data\            学生主表、谈话记录、成绩分析' + #13#10 +
                   '  · photos_src\      学生照片' + #13#10 +
                   '  · photos\          照片缩略图' + #13#10 +
                   '  · 备份\            你自己放的备份' + #13#10 +
                   '  · 辅导员工作台.db  数据库副本' + #13#10#13#10 +
                   '选「否」＝只卸载程序，数据留着，以后重装可以接着用（推荐）。' + #13#10 +
                   '选「是」＝连数据一起删掉，无法恢复，请先自己备份。',
                   mbConfirmation, MB_YESNO);
  if Answer = IDYES then
  begin
    DelTree(ExpandConstant('{app}\data'), True, True, True);
    DelTree(ExpandConstant('{app}\photos_src'), True, True, True);
    DelTree(ExpandConstant('{app}\photos'), True, True, True);
    DelTree(ExpandConstant('{app}\备份'), True, True, True);
    DelTree(ExpandConstant('{app}\导出_脱敏版'), True, True, True);
    DeleteFile(ExpandConstant('{app}\辅导员工作台.db'));
    DeleteFile(ExpandConstant('{app}\photo_index.json'));
    DeleteFile(ExpandConstant('{app}\dashboard.html'));
    DeleteFile(ExpandConstant('{app}\dashboard_脱敏外发版.html'));
    DeleteFile(ExpandConstant('{app}\工作台数据汇总.xlsx'));
  end
  else
    DataKept := True;
end;
