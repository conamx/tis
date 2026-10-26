[Setup]
AppName=원고 생성기
AppVersion=260826
AppPublisher=jnc마케팅
DefaultDirName={autopf}\원고생성기
DefaultGroupName=원고생성기
OutputBaseFilename=원고생성기_설치파일
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\원고 생성기.exe
AppMutex=원고생성기_running
CloseApplications=force
RestartApplications=no

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 만들기"; GroupDescription: "추가 작업:"; Flags: unchecked

[Files]
Source: "dist\원고 생성기.exe"; DestDir: "{app}"; Flags: ignoreversion
[Icons]
Name: "{group}\원고 생성기"; Filename: "{app}\원고 생성기.exe"
Name: "{group}\원고 생성기 제거"; Filename: "{uninstallexe}"
Name: "{autodesktop}\원고 생성기"; Filename: "{app}\원고 생성기.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\원고 생성기.exe"; Description: "원고 생성기 실행"; Flags: nowait postinstall skipifsilent

[Code]
procedure InitializeWizard;
begin
end;

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  // CloseApplications가 못 잡아내는 경우 대비, 실행 중인 프로세스 강제 종료
  Exec('taskkill.exe', '/F /IM "원고 생성기.exe"', '', SW_HIDE,
       ewWaitUntilTerminated, ResultCode);
  Result := True;
end;