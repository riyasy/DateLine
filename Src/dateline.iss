[Setup]
AppName=DateLine
AppVerName=DateLine V1.0
AppPublisher=AAA
AppPublisherURL=http://www.AAA.com/
AppSupportURL=http://www.AAA.com/
AppUpdatesURL=http://www.weytec.com/
DisableDirPage=yes
DefaultDirName={pf}\DateLine
CreateAppDir=yes
OutputBaseFilename=DateLineV1.0_Setup
Compression=lzma
SolidCompression=yes
;Overwrite an existing Installation log
UninstallLogMode=overwrite

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

;Copy File to Application Directory
[Files]
Source: "USBMouseSwitching.exe"; DestDir: "{app}"; Flags: ignoreversion; BeforeInstall: BeforeProgInstall();
Source: "..\\bin\\DateLine.exe"; DestDir: "{app}"; Flags: ignoreversion;
Source: "..\\bin\\Interop.Microsoft.Office.Interop.Outlook.dll"; DestDir: "{app}"; Flags: ignoreversion;


;After Install
[Run]
;Register Service
Filename: "{app}\USBMouseSwitching.exe"; Parameters: "/install /silent"; Flags: runhidden; StatusMsg: "Register Service USBMouseSwitch..."
Filename: "{app}\MouseTrackerSwitcher.exe"; Parameters: "/install"; Flags: runhidden; StatusMsg: "Register Service MouseTrackerSwitcher..."

;Before Uninstall
[UninstallRun]
;Stop Service
Filename: "{sys}\net.exe"; Parameters: "stop WEYUSBMouseSwitching"; Flags: runhidden; StatusMsg: "Stop Service USBMouseSwitch..."
Filename: "{sys}\net.exe"; Parameters: "stop MouseTrackerSwitcher"; Flags: runhidden; StatusMsg: "Stop Service USBMouseSwitch..."


;Delete All Directories
[UninstallDelete]
;Delete application Folder
Type: filesandordirs; Name: "{app}"

[Code]
//Stop Service before install
procedure BeforeProgInstall();
var
  ResultCode:Integer;
begin
  //Stop Service
  Exec(ExpandConstant('{sys}\net.exe'), 'stop WEYUSBMouseSwitching', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Exec(ExpandConstant('{sys}\net.exe'), 'stop MouseTrackerSwitcher', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  //Unregister Service
  Exec(ExpandConstant('{app}\USBMouseSwitching.exe'), '/uninstall /silent', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Exec(ExpandConstant('{app}\MouseTrackerSwitcher.exe'), '-u', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;







