; PDFCreator Installation
; Setup created with Inno Setup  4.0.2 beta and ISTool 3.9.9
; Installation from Frank Heindörfer, Philip Chinery

#define GetFileVersionVBExe(str S) Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + Local[3] + Local[4]

#define AppVersion  GetFileVersionVBExe("..\PDFCreator\PDFCreator.exe")
#define Appname         "PDFCreator"
#define Printername     "PDFCreator"
#define Drivername      "PDFCreator"
#define Portname        "PDFCreator:"
#define Monitorname     "PDFCreator"
#define AppExename      "PDFCreator.exe"
#define SpoolerExename  "PDFSpooler.exe"

[_ISTool]
EnableISX=true

[_ISToolPreCompile]
Name: .\upx\upx.exe; Parameters: ..\TransTool\TransTool.exe --best --compress-icons=0
Name: .\upx\upx.exe; Parameters: ..\PDFSpooler\PDFSpooler.exe --best --compress-icons=0
Name: .\upx\upx.exe; Parameters: ..\PDFCreator\PDFCreator.exe --best --compress-icons=0

[Setup]
AllowNoIcons=false
AlwaysRestart=false
AppCopyright=© 2002 - 2003 Philip Chinery, Frank Heindörfer
AppName={#AppName}
AppVerName={#AppName} {#AppVersion}
AppPublisher=Philip Chinery, Frank Heindörfer
AppPublisherURL=http://www.pdfcreator.de.vu
AppSupportURL=http://www.pdfcreator.de.vu
AppUpdatesURL=http://www.pdfcreator.de.vu
Compression=bzip
DefaultDirName={pf}\{#AppName}
DefaultGroupName={#AppName}
DisableStartupPrompt=true
LicenseFile=.\License\readme.rtf
OutputBaseFilename={#AppName}_Setup_{#AppVersion}
OutputDir=Installation
PrivilegesRequired=admin
RestartIfNeededByRun=true
ShowTasksTreeLines=false
SolidCompression=True
WizardImageFile=..\Pictures\PDFCreatorBig.bmp
WizardSmallImageFile=..\Pictures\PDFCreator.bmp

[Files]
Source: ..\SystemFiles\VB6DE.DLL; DestDir: {sys}; Flags: sharedfile; Components: program programm
Source: ..\SystemFiles\STDOLE2.TLB; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib; Components: program programm
Source: ..\SystemFiles\ASYCFILT.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile; Components: program programm
Source: ..\SystemFiles\OLEPRO32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver; Components: program programm
Source: ..\SystemFiles\OLEAUT32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver; Components: program programm
Source: ..\SystemFiles\MSVBVM60.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver; Components: program programm

Source: ..\SystemFiles\MSCC2DE.DLL; DestDir: {sys}; Flags: sharedfile; Components: program programm
Source: ..\SystemFiles\MSCOMCT2.OCX; DestDir: {sys}; Flags: regserver sharedfile; Components: program programm
Source: ..\SystemFiles\CMDLGDE.DLL; DestDir: {sys}; Flags: sharedfile; Components: program programm
Source: ..\SystemFiles\COMDLG32.OCX; DestDir: {sys}; Flags: regserver sharedfile; Components: program programm
Source: ..\SystemFiles\MSCMCDE.DLL; DestDir: {sys}; Flags: sharedfile; Components: program programm
Source: ..\SystemFiles\Mscomctl.ocx; DestDir: {sys}; Flags: regserver sharedfile; Components: program programm

Source: ..\PDFCreator\gsdll32.dll; DestDir: {app}; Components: program programm
Source: ..\PDFCreator\PDFCreator.exe; DestDir: {app}; Components: program programm
Source: .\License\License.txt; DestDir: {app}; Components: program programm
Source: ..\PDFCreator\Languages\*.ini; DestDir: {app}\languages; Components: program programm
Source: ..\PDFCreator\Fonts\*.afm; DestDir: {app}\fonts; Components: program programm
Source: ..\PDFCreator\Fonts\*.pfb; DestDir: {app}\fonts; Components: program programm
Source: ..\PDFCreator\Fonts\fonts.dir; DestDir: {app}\fonts; Components: program programm
Source: ..\PDFCreator\Fonts\fonts.scale; DestDir: {app}\fonts; Components: program programm
Source: ..\PDFCreator\Lib\*.*; DestDir: {app}\lib; Components: program programm
Source: ..\Transtool\TransTool.exe; DestDir: {app}\languages; Components: program programm
Source: ..\PDFSpooler\PDFSpooler.exe; DestDir: {app}; Components: printer drucker
Source: PDFCreatorEnglish.ini; DestDir: {app}; Components: program programm; DestName: PDFCreator.ini; Check: LanguageIs(English)
Source: PDFCreatorGerman.ini; DestDir: {app}; Components: program programm; DestName: PDFCreator.ini; Check: LanguageIs(German)

Source: ..\Printer\Win98\*.*; DestDir: {win}; MinVersion: 4.00.950,0; Components: printer drucker
Source: ..\Printer\Win98\System\*.*; DestDir: {sys}; MinVersion: 4.00.950,0; Components: printer drucker
Source: ..\Printer\WinNt\*.*; DestDir: {sys}\spool\drivers\w32x86; MinVersion: 0,4.00.1381; Components: printer drucker

Source: ..\Printer\Redmon\redmonnt.dll; DestDir: {sys}; Flags: restartreplace  sharedfile; MinVersion: 0,4.00.1381; Components: printer drucker
Source: ..\Printer\Redmon\redmon95.dll; DestDir: {sys}; Flags: restartreplace  sharedfile; MinVersion: 4.00.950,0; Components: printer drucker

[Registry]
;PrinterMonitor
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: string; Valuename: Arguments; ValueData: -PPDFCREATORPRINTER; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: string; Valuename: Command; ValueData: {app}\{#SpoolerExename}; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: dword; Valuename: Delay; ValueData: 300; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: string; Valuename: Description; ValueData: Redirected Port; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: dword; Valuename: LogFileDebug; ValueData: 0; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: dword; Valuename: LogFileUse; ValueData: 0; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: dword; Valuename: Output; ValueData: 0; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: string; Valuename: Printer; ValueData: {#Printername}; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: dword; Valuename: Printerror; ValueData: 0; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: dword; Valuename: Runuser; ValueData: 1; Flags: uninsdeletevalue; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}\Ports\{#Portname}; ValueType: dword; Valuename: ShowWindow; ValueData: 0; Flags: uninsdeletevalue; Components: printer drucker
;PrinterPort
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Ports\{#Portname}; Components: printer drucker

;Uninstall - Deletekey
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Printers\{#Printername}; Flags: uninsdeletekey; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Environments\Windows 4.0\Drivers\{#Drivername}; Flags: uninsdeletekey; MinVersion: 4.00.950,0; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Environments\Windows NT x86\Drivers\{#Drivername}; Flags: uninsdeletekey; MinVersion: 0,4.00.1381; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Ports\{#Portname}; Flags: uninsdeletekey; Components: printer drucker
Root: HKLM; Subkey: System\CurrentControlSet\Control\Print\Monitors\{#Monitorname}; Flags: uninsdeletekey; Components: printer drucker


[INI]
Filename: win.ini; Section: windows; Key: device; String: {#Appname},PSCRIPT,{#Portname}; MinVersion: 4.00.950,0; Components: printer drucker
Filename: win.ini; Section: Devices; Key: {#Appname}; String: PSCRIPT,{#Portname}; MinVersion: 4.00.950,0; Components: printer drucker
Filename: win.ini; Section: Devices; Key: {#Appname}; String: PSCRIPT,{#Portname},15,45; MinVersion: 4.00.950,0; Components: printer drucker
Filename: win.ini; Section: windows; Key: device; String: {#Appname},PSCRIPT,{#Monitorname}; Components: printer drucker


[Icons]
Name: {group}\{#Appname}; Filename: {app}\{#AppExename}; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\License; Filename: {app}\License.txt; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\Translation Tool; Filename: {app}\languages\transtool.exe; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\Uninstall {#Appname}; Filename: {uninstallexe}; IconIndex: 0; Flags: createonlyiffileexists

[Run]
Filename: {app}\PDFCreator.exe; Parameters: -NSTRUE; WorkingDir: {app}; Description: Install printerdriver; StatusMsg: Install PDFCreator printerdriver; Flags: runminimized; Components: printer drucker; Check: InstallPrinterDriverEnglish(English)
Filename: {app}\PDFCreator.exe; Parameters: -NSTRUE; WorkingDir: {app}; Description: Installiere Druckertreiber; StatusMsg: Installiere PDFCreator Druckertreiber; Flags: runminimized; Components: printer drucker; Check: InstallPrinterDriverGerman(German)
;Filename: {app}\{#AppExename}; WorkingDir: {app}; Parameters: -IPTRUE; Flags: postinstall nowait

[UninstallDelete]
Name: {app}; Type: filesandordirs
Name: {%tmp}\{#Appname}; Type: filesandordirs

[UninstallRun]
Filename: {app}\PDFCreator.exe; WorkingDir: {app}; Parameters: -IPFALSE -ULTRUE -NSTRUE; Flags: runminimized


[Languages]
Name: English; MessagesFile: compiler:Default.isl
Name: German; MessagesFile: German-2-4.0.0.isl

[Types]
Name: full; Description: Full installation; Check: LanguageIs(English)
Name: compact; Description: Compact installation; Check: LanguageIs(English)
Name: custom; Description: Custom installation; Flags: iscustom; Check: LanguageIs(English)

Name: voll; Description: Komplette Installation; Check: LanguageIs(German)
Name: minimal; Description: Minimale Installation; Check: LanguageIs(German)
Name: benutzer; Description: Benutzerdefinierte Installation; Flags: iscustom; Check: LanguageIs(German)

[Components]
Name: program; Description: Program Files; Types: full compact custom; Flags: fixed; Check: LanguageIs(English)
Name: printer; Description: Printer Driver; Types: full custom; Check: LanguageIs(English)

Name: programm; Description: Programm Dateien; Types: voll minimal benutzer; Flags: fixed; Check: LanguageIs(German)
Name: drucker; Description: Druckertreiber; Types: voll benutzer; Check: LanguageIs(German)

[Tasks]
Name: desktopicon; Description: Create a &desktop icon; GroupDescription: Additional icons:; Languages: English
Name: desktopicon\common; Description: For all users; GroupDescription: Additional icons:; Flags: exclusive; Languages: English
Name: desktopicon\user; Description: For the current user only; GroupDescription: Additional icons:; Flags: exclusive unchecked; Languages: English
Name: quicklaunchicon; Description: Create a &Quick Launch icon; GroupDescription: Additional icons:; Flags: unchecked; Languages: English
Name: fileassoc; Description: &Associate PDFCreator with the .ps file extension; GroupDescription: Other tasks:; Flags: unchecked; Languages: English

Name: desktopicon; Description: &Desktopsymbol anlegen; GroupDescription: Zusätzliche Symbole:; Languages: German
Name: desktopicon\common; Description: Für &alle Benutzer; GroupDescription: Zusätzliche Symbole:; Flags: exclusive; Languages: German
Name: desktopicon\user; Description: Nur für den angemeldeten &Benutzer; GroupDescription: Zusätzliche Symbole:; Flags: exclusive unchecked; Languages: German
Name: quicklaunchicon; Description: Erzeuge eine Symbol in der &Schnellzugriffsleiste; GroupDescription: Zusätzliche Symbole:; Flags: unchecked; Languages: German
Name: fileassoc; Description: &Verknüpfe PDFCreator mit der dateierweiterung .ps; GroupDescription: Andere Aufgaben:; Flags: unchecked; Languages: German

[Code]
type
 TMonitorInfo2 = record
  pName : PChar;
  pEnvironment : PChar;
  pDLLName : PChar;
 end;
 TDriverInfo3 = record
  cVersion : LongInt;
  pName : PChar;
  pEnvironment : PChar;
  pDriverPath : PChar;
  pDataFile : PChar;
  pConfigFile : PChar;
  pHelpFile : PChar;
  pDependentFiles : PChar;
  pMonitorName : PChar;
  pDefaultDataType : PChar;
 end;
 TPrinterInfo2 = record
  pServerName : PChar;
  pPrinterName : PChar;
  pShareName : PChar;
  pPortName : PChar;
  pDriverName : PChar;
  pComment : PChar;
  pLocation : PChar;
  pDevMode : PChar;
  pSepFile : PChar;
  pPrintProcessor : PChar;
  pDatatype : PChar;
  pParameters : PChar;
  pSecurityDescriptor : PChar;
  Attributes : LongInt;
  Priority : LongInt;
  DefaultPriority : LongInt;
  StartTime : LongInt;
  UntilTime : LongInt;
  Status : LongInt;
  cJobs : LongInt;
  AveragePPM : LongInt;
 end;

function GetWindow (hWnd : LongInt; wCmd : LongInt) : LongInt;
 external 'GetWindow@user32.dll';
function GetWindowLong  (hWnd : LongInt; wIndx : LongInt) : LongInt;
 external 'GetWindowLongA@user32.dll';
function GetWindowText (hWnd : LongInt; lpString : String; cch : LongInt) : LongInt;
 external 'GetWindowTextA@user32.dll';
function GetWindowTextLength (hWnd : LongInt) : LongInt;
 external 'GetWindowTextLengthA@user32.dll';
function GetWindowThreadProcessId (hWnd : Longint;var lpdwProcessId : Longint) : Longint;
 external 'GetWindowThreadProcessId@user32.dll';
function GetParent (var hWnd : Longint) : Longint;
 external 'GetParent@user32.dll';

function OpenProcess (dwDesiredAccess : Longint; bInheritHandle : LongInt; dwProcessId : Longint) : Longint;
 external 'OpenProcess@kernel32.dll';
function TerminateProcess (hProcess : Longint; uExitCode : Longint) : Longint;
 external 'TerminateProcess@kernel32.dll';
function CloseHandle (hObject : Longint) : Longint;
 external 'CloseHandle@kernel32.dll';

function AddMonitor (pName : PChar; Level : LongInt;var pMonitors : TMonitorInfo2): LongInt; external
'AddMonitorA@winspool.drv stdcall';
function AddPort (pName : PChar; hwnd : LongInt; pPort : PChar): LongInt; external
'AddPortA@winspool.drv stdcall';
function AddPrinterDriver (pName : PChar; Level : LongInt;var pDriverInfo : TDriverInfo3) : LongInt; external
'AddPrinterDriverA@winspool.drv stdcall';
function ClosePrinter(pPrinter: LongInt): Boolean; external
'ClosePrinter@winspool.drv stdcall';
function AddPrinter(pName : PChar; Level: Longint; var pPrinter2: TPrinterInfo2): LongInt; external
'AddPrinterA@winspool.drv stdcall';
function GetLastError() : LongInt; external
'GetLastError@kernel32.dll stdcall';


var progTitel, progHandle: TArrayOfString;


procedure ButtonOnClick(Sender: TObject);
begin
  MsgBox('You clicked the button!', mbInformation, mb_Ok);
end;

procedure FormButtonOnClick(Sender: TObject);
var
  Form: TForm;
  Button: TButton;
begin
  Form := TForm.Create(WizardForm);
  Form.Width := 256;
  Form.Height := 256;
  Form.Caption := 'TForm';
  Form.Position := poScreenCenter;

  Button := TButton.Create(Form);
  Button.Parent := Form;
  Button.Left := 8;
  Button.Top := Form.ClientHeight - Button.Height - 10;
  Button.Caption := 'Close';
  Button.ModalResult := mrOk;

  Form.ActiveControl := Button;

  Form.ShowModal();
  Form.Release();
end;



procedure GetActivePDFLoaders;
var
 Style:LongInt;
 GWL_STYLE, GW_HWNDNEXT, WS_VISIBLE, WS_BORDER, procID :LongInt;
 whwnd, res, parent:LongInt;
 Caption:String;
begin
 WS_VISIBLE:= $10000000; WS_BORDER:= $800000;
 GWL_STYLE:=-16; GW_HWNDNEXT := 2;

 whwnd := StrToInt(ExpandConstant('{wizardhwnd}'));
 repeat
  whWnd := GetWindow(whWnd, GW_HWNDNEXT);
  Style:= GetWindowLong(whwnd, GWL_STYLE);
//  Style:= Style And (WS_VISIBLE Or WS_BORDER);
  Style:= WS_BORDER;
  res:=GetWindowTextLength(whwnd);
  if (res>0) and (Style = (WS_VISIBLE Or WS_BORDER)) then begin
   Caption:=StringOfChar(' ',res);
   GetWindowText(whwnd,Caption,res+1);
   if (length(Caption)>0) and (Caption<>'Setup') then begin
    if (UpperCase(Caption)='PDFLOADER') or (UpperCase(Caption)='PDFCREATOR') then begin
     setarraylength(progTitel,getarraylength(progTitel)+1);
     progTitel[getarraylength(progTitel)-1]:=Caption;
     parent:=whWnd;
     repeat
      parent:=GetParent(parent);
     until parent = 0;

     res := GetWindowThreadProcessId(whWnd, procID);
     setarraylength(progHandle,getarraylength(progHandle)+1);
     progHandle[getarraylength(progHandle)-1]:=IntToStr(procID);
    end;
   end;
  end;
 Until whWnd = 0;
end;

procedure KillActivePDFLoaders;
var
 i, c, res, task, inheritHandle, exitCode, PROCESS_TERMINATE, ProcID : LongInt;
begin
 PROCESS_TERMINATE:=1; inheritHandle:=0; exitCode:=1;
 c:=GetArrayLength(progHandle);
 for i:=0 to c-1 do begin
  ProcID := StrToInt(progHandle[i]);
  task   := OpenProcess(PROCESS_TERMINATE, inheritHandle, ProcID);
  res    := TerminateProcess(task, exitCode);
  res    := CloseHandle(task);
 end;
end;

procedure InstallMonitor;
var M2:TMonitorInfo2; res:LongInt;
begin
 M2.pName:=ExpandConstant('{#Monitorname}');
 If UsingWinNT=True then Begin
   M2.pEnvironment:='Windows NT x86';
   M2.pDLLName:='redmonnt.dll'
  end else Begin
   M2.pEnvironment:='Windows 4.0';
   M2.pDLLName:='redmon95.dll'
 end;

 res := AddMonitor(CastIntegerToString(0), 2, M2);
 if res=0 then
   SaveStringToFile(ExpandConstant('{app}') + '\SetupLog.txt', 'InstallMonitor: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10, True)
  else
   SaveStringToFile(ExpandConstant('{app}') + '\SetupLog.txt', 'InstallMonitor: Success' + #13#10, True);
end;

procedure InstallDriver;
var DI3:TDriverInfo3; res:LongInt;
begin

 DI3.pName :=ExpandConstant('{#Drivername}');
 If UsingWinNT=True then Begin
   DI3.cVersion:=1;
   DI3.pConfigFile :='ADOBEPSU.DLL';
   DI3.pDataFile :='ADIST5.PPD';
   DI3.pDriverPath := 'ADOBEPS5.DLL';
   DI3.pEnvironment:='Windows NT x86';
   DI3.pHelpFile :='ADOBEPSU.HLP';
  end else Begin
   DI3.cVersion:=0;
   DI3.pConfigFile :='ADOBEPS4.DRV';
   DI3.pDataFile :='ADIST5.PPD';
   DI3.pDriverPath := 'ADOBEPS4.DRV';
   DI3.pEnvironment:='Windows 4.0';
   DI3.pHelpFile :='ADOBEPS4.HLP';
 end;
 DI3.pDependentFiles :='';
 DI3.pDefaultDataType :='RAW';
 DI3.pMonitorName :='';

 res := AddPrinterDriver(CastIntegerToString(0), 3, DI3);
 if res=0 then
   SaveStringToFile(ExpandConstant('{app}') + '\SetupLog.txt', 'InstallDriver: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10, True)
  else
   SaveStringToFile(ExpandConstant('{app}') + '\SetupLog.txt', 'InstallDriver: Success' + #13#10, True);
end;

procedure InstallPrinter;
var
 P2: TPrinterInfo2; res: LongInt;
begin
  P2.pPrinterName := ExpandConstant('{#Monitorname}');
  P2.pDriverName := ExpandConstant('{#Drivername}');
  P2.pPrintProcessor := 'WinPrint';
  P2.pPortName := ExpandConstant('{#Portname}');
  P2.pComment := 'eDoc Printer';

  res := AddPrinter( CastIntegerToString(0), 2, P2 );

  if res<>0 then begin
    ClosePrinter(res);
    SaveStringToFile(ExpandConstant('{app}') + '\SetupLog.txt', 'InstallPrinter: Success' + #13#10, True)
   end else
    SaveStringToFile(ExpandConstant('{app}') + '\SetupLog.txt', 'InstallPrinter: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10, True);
end;

function LanguageIs(Language: string): boolean;
begin
 Result:=(ActiveLanguage()=Language);
end;

function InstallPrinterDriverEnglish(Language: string): boolean;
begin
 If ActiveLanguage()='English' Then Begin
  InstallMonitor;
  InstallDriver;
  InstallPrinter;
  Result:=True;
 end else
  Result:=False;
end;

function InstallPrinterDriverGerman(Language: string): boolean;
begin
 If ActiveLanguage()='German' Then Begin
  InstallMonitor;
  InstallDriver;
  InstallPrinter;
  Result:=True;
 end else
  Result:=False;
end;

function ScriptDlgPages(CurPage: Integer; BackClicked: Boolean): Boolean;
begin
 if (not BackClicked and (CurPage = wpReady)) or (BackClicked and (CurPage = wpFinished)) then begin
  GetActivePDFLoaders;
  KillActivePDFLoaders;
 end;
 Result := True;
end;


function NextButtonClick(CurPage: Integer): Boolean;
begin
 Result := ScriptDlgPages(CurPage, False);
end;

function BackButtonClick(CurPage: Integer): Boolean;
begin
 Result := ScriptDlgPages(CurPage, True);
end;
