; Transtool stand alone installation
; Setup created with Inno Setup QuickStart Pack 5.1.14 (with ISPP) and ISTool 5.1.8
; Installation from Frank Heindörfer

;#define Test

#ifdef Test
 #define FastCompilation
 #define IncludeToolbar
#else
; #define FastCompilation
 #define CompileHelp
 #define IncludeToolbar
 #define Localization
#endif

#define ProgramLicense "GNU"

#ifdef FastCompilation
 #define CompressionMode="none"
 #define SetupLZMACompressionMode "none"
#else
 #define CompressionMode="lzma/ultra"
 #define SetupLZMACompressionMode "ultra"
#endif

#if (fileexists("..\TransTool\TransTool.exe")==0)
 #error Compile TransTool first!
#endif

;remove the german localization
#IFNDEF Test
 #IFDEF Localization
  #if (fileexists("C:\IPDK\VBLOCAL.EXE")==0)
   #error Please install the IPDK!
  #endif
  #expr Exec("C:\IPDK\VBLOCAL.EXE","..\TransTool\TransTool.exe * 0x409 ~ 0x0",".\")
 #endif
#endif

;add manifest to exe files
#IFNDEF Test
 #if (fileexists("..\ManifestManager\ManifestManager.exe")==0)
  #error Compile ManifestManager first!
 #endif
 #expr Exec("..\ManifestManager\ManifestManager.exe","/ADD""..\TransTool\TransTool.exe""","..\ManifestManager\")
#endif

#define GetFileVersionVBExe(str S)     Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + Local[3] + Local[4]
#define GetFileVersionVBExeLine(str S) Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])-1), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])-1), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + '_' + Local[3] + '_'  + Local[4]

#define Homepage             "http://www.pdfforge.org/projects/transtool"
#define Appname              "TransTool"
#define AppExename           Appname + ".exe"

#define AppVersion           GetFileVersionVBExe("..\TransTool\TransTool.exe")

#define SetupAppVersion      GetFileVersionVBExeLine("..\TransTool\TransTool.exe")
#define TransToolVersion     GetFileVersionVBExe("..\Transtool\Transtool.exe")

#define AppVersionStr       AppVersion
#define SetupAppVersionStr  SetupAppVersion

#define AppIDGUID            "00017DAD-2AB1-48E1-807B-CC8D9FC3757D"
#define AppID                "{" + AppIDGUID + "}"
#define AppIDStr             "{" + AppID
#define AppIDreg             "{" + AppIDGUID + "%7d"
#define TransToolExeID       "{B7BCA0D2-7305-4318-BA7A-01B028D910EB}"
#define TransToolExeIDStr    "{" + TransToolExeID
#define UninstallID          AppID
#define UninstallIDreg       AppIDreg
#define UninstallIDStr       "{"+ UninstallID
#define UninstallIDStr2      "{"+ UninstallIDreg

#define UninstallReg         "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallID
#define UninstallRegStr      "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallIDStr
#define UninstallRegStr2     "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallIDStr2

;#define UpdateIsPossible
#define UpdateIsPossibleMinVersion "2.9.0"

#IFDEF IncludeToolbar
 #include "ToolbarForm.isd"
#ENDIF

[Setup]
AllowNoIcons=true
AlwaysRestart=false
AppContact={#Homepage}
AppCopyright=© Frank Heindörfer, Philip Chinery
AppID={#AppIDStr}
AppName={#AppName}
AppVerName={#AppName} {#AppVersionStr}
AppPublisher=Frank Heindörfer
AppPublisherURL={#Homepage}
AppSupportURL={#Homepage}
AppUpdatesURL={#Homepage}
AppVersion={#AppVersion}
Compression=lzma/ultra
CreateUninstallRegKey=false
DefaultDirName={reg:HKLM\{#UninstallRegStr2},Inno Setup: App Path|{pf}\{#AppName}}
DefaultGroupName={#AppName}
DisableDirPage=false
DisableStartupPrompt=true
ExtraDiskSpaceRequired=10303775
InternalCompressLevel=ultra

LicenseFile=.\License\TransTool license - english.rtf
OutputBaseFilename={#AppName}-{#SetupAppVersionStr}
OutputDir=Installation
RestartIfNeededByRun=true
ShowLanguageDialog=no
ShowTasksTreeLines=false
SolidCompression=true
UsePreviousAppDir=true

VersionInfoVersion={#AppVersion}
VersionInfoCompany=Frank Heindörfer
VersionInfoDescription=TransTool basically is a tool written for PDFCreator, to enable users to translate PDFCreators language files. It can be used to translate all kinds of programs that use INI files.
VersionInfoTextVersion={#AppVersion}

WizardImageFile=..\Pictures\Setup\TransToolBig.bmp
WizardSmallImageFile=..\Pictures\Setup\TransTool.bmp

[InstallDelete]
Name: {app}\unload.tmp; Type: files; Components: program

[Files]
#IFNDEF Test
;We sort all files by extension for a maximal compression
;Systemfiles
Source: ..\SystemFiles\ASYCFILT.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace uninsneveruninstall

;Please use newest MSVBVM60.DLL
;http://support.microsoft.com/default.aspx?scid=kb;en-us;823746
Source: ..\SystemFiles\MSVBVM60.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace regserver uninsneveruninstall

Source: ..\SystemFiles\OLEPRO32.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace regserver uninsneveruninstall; OnlyBelowVersion: 0,6.0
Source: ..\SystemFiles\OLEAUT32.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace regserver uninsneveruninstall; OnlyBelowVersion: 0,6.0

Source: ..\SystemFiles\MSCOMCT2.OCX; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt regserver
Source: ..\SystemFiles\MSCOMCTL.OCX; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt regserver

Source: ..\SystemFiles\STDOLE2.TLB; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace uninsneveruninstall regtypelib; OnlyBelowVersion: 0,6.0

;Program files
Source: ..\Transtool\TransTool.exe; DestDir: {app}; Components: program; Flags: comparetimestamp

;ShFolder for older systems
;http://www.microsoft.com/downloads/release.asp?releaseid=30340
Source: ShFolder\ShFolder.Exe; DestDir: {app}; Components: program; Flags: ignoreversion deleteafterinstall; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 4.1.2222,5.0.2195

Source: License\GNU License.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp

; Toolbar
#IFDEF IncludeToolbar
Source: ..\Pictures\Toolbar\Toolbar.bmp; DestDir: {tmp}; Flags: dontcopy nocompression; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,0
Source: ..\Toolbar\PDFCreator_Toolbar_Setup.exe; DestDir: {tmp}; DestName: PDFCreator_Toolbar_Setup.exe; Components: ietoolbar; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,0
#ENDIF
#ENDIF

[Icons]
Name: {group}\{#Appname}; Filename: {app}\{#AppExename}; WorkingDir: {app}; IconFilename: {app}\{#AppExename}; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\GPL License; Filename: {app}\GNU License.txt; WorkingDir: {app}
Name: {group}\Donate TransTool; Filename: {app}\Donate TransTool.url; WorkingDir: {app}; IconFilename: {app}\TransTool.exe; IconIndex: 2
Name: {group}\{cm:ProgramOnTheWeb,TransTool}; Filename: {app}\TransTool.url; WorkingDir: {app}; IconFilename: {app}\TransTool.exe; IconIndex: 1

Name: {commondesktop}\TransTool; Filename: {app}\TransTool.exe; WorkingDir: {app}; IconIndex: 0; Tasks: desktopicon\common
Name: {userdesktop}\TransTool; Filename: {app}\TransTool.exe; WorkingDir: {app}; IconIndex: 0; Tasks: desktopicon\user
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\TransTool; Filename: {app}\TransTool.exe; WorkingDir: {app}; IconIndex: 0; Tasks: quicklaunchicon

[INI]
Filename: {app}\TransTool.url; Section: InternetShortcut; Key: URL; String: http://www.pdfforge.org; Components: program
Filename: {app}\TransTool.url; Section: InternetShortcut; Key: Iconindex; String: 20; Components: program
Filename: {app}\TransTool.url; Section: InternetShortcut; Key: IconFile; String: {app}\TransTool.exe; Components: program

Filename: {app}\Donate TransTool.url; Section: InternetShortcut; Key: URL; String: http://www.paypal.com/xclick/business=paypal01%40heindoerfer.com&item_name=TransTool&no_note=1&tax=0&currency_code=EUR; Components: program
Filename: {app}\Donate TransTool.url; Section: InternetShortcut; Key: Iconindex; String: 21; Components: program
Filename: {app}\Donate TransTool.url; Section: InternetShortcut; Key: IconFile; String: {app}\TransTool.exe; Components: program

[Registry]
;Uninstall - Software
Root: HKLM; Subkey: {#UninstallRegStr}; Flags: uninsdeletekey
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Comments; Valuedata: TransTool - Opensource; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayIcon; Valuedata: {app}\TransTool.exe; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayName; Valuedata: {#AppName} {#AppVersionStr}; Flags: uninsdeletevalue; MinVersion: 4.0.950,0; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayName; Valuedata: {#AppName}; Flags: uninsdeletevalue; MinVersion: 0,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayVersion; Valuedata: {#AppVersionStr}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: HelpLink; Valuedata: {#Homepage}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: InstallDate; Valuedata: {code:GetDateString}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Publisher; Valuedata: Frank Heindörfer, Philip Chinery; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Readme; Valuedata: {#Homepage}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: URLInfoAbout; Valuedata: {#Homepage}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: URLUpdateInfo; Valuedata: {#Homepage}; Flags: uninsdeletevalue

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: ApplicationVersion; Valuedata: {#AppVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: TranstoolVersion; Valuedata: {#TranstoolVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: UninstallString; Valuedata: {app}\unins000.exe; Flags: uninsdeletevalue

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: App Path; Valuedata: {app}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Components; Valuedata: {code:GetWizardSelectedComponents}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Tasks; Valuedata: {code:GetWizardSelectedTasks}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Silent; Valuedata: {code:GetWizardSilent}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Group; Valuedata: {code:GetWizardGroupValue}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: NoIcons; Valuedata: {code:GetWizardNoIcons}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: SetupType; Valuedata: {code:GetWizardSetupType}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: SetupLanguage; Valuedata: {code:GetActiveLanguage}; Flags: uninsdeletevalue

[Run]
#IFNDEF Test
Filename: {app}\ShFolder.Exe; WorkingDir: {app}; Parameters: /Q:A; Flags: runminimized; Components: program; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 4.1.2222,5.0.2195
Filename: {app}\TransTool.exe; WorkingDir: {app}; Description: {cm:LaunchProgram,{#Appname}}; Flags: postinstall nowait skipifsilent
#ENDIF

#IFDEF IncludeToolbar
Filename: {tmp}\PDFCreator_Toolbar_Setup.exe; Components: ietoolbar
#ENDIF

[UninstallDelete]
Name: {app}\SetupLog.txt; Type: files
Name: {app}\Unload.tmp; Type: files
Name: {app}\TransTool.url; Type: files
Name: {app}\{cm:Donation}.url; Type: files
Name: {app}; Type: dirifempty
;User temp directories

[Messages]
;Remove the 'StatusRunProgram' message
StatusRunProgram=

[Languages]
Name: english; MessagesFile: compiler:Default.isl

[CustomMessages]
#include "Language includes\english.inc"

[Types]
Name: custom; Description: {cm:CustomInstallation}; Flags: iscustom
Name: full; Description: {cm:FullInstallation}
Name: compact; Description: {cm:CompactInstallation}

[Components]
Name: program; Description: {cm:Programfiles}; Types: full compact custom; Flags: fixed

#IFDEF IncludeToolbar
Name: ietoolbar; Description: {cm:Toolbarfiles}; ExtraDiskSpaceRequired: 909077; Types: full custom; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,6.0; Check: IExplorerVersionLower55
Name: ietoolbar; Description: {cm:Toolbarfiles}; ExtraDiskSpaceRequired: 909077; Types: ; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,6.0; Check: Not IExplorerVersionLower55; Flags: fixed
#ENDIF

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Check: UseDesktopiconCommon
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked; Check: Not UseDesktopiconCommon
Name: desktopicon\common; Description: {cm:ForAllUser}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive; Check: UseDesktopiconCommon
Name: desktopicon\common; Description: {cm:ForAllUser}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive unchecked; Check: Not UseDesktopiconCommon
Name: desktopicon\user; Description: {cm:ForTheCurrentUserOnly}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive; Check: UseDesktopiconUser
Name: desktopicon\user; Description: {cm:ForTheCurrentUserOnly}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive unchecked; Check: Not UseDesktopiconUser
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Check: IExplorerVersionGreater3 And UseQuickLaunchIcon
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked; Check: IExplorerVersionGreater3 And Not UseQuickLaunchIcon

[Code]
type
 TAInt = Array of Integer; TAStr = Array of String;

function SearchPath(lpPath : String; lpFilename : String; lpExtension : String; nBufferLength : LongInt; lpBuffer : String; lpFilePart : LongInt) : LongInt;
 external 'SearchPathA@kernel32.dll stdcall';


var msg : TAStr;
    FullInstallation : boolean;
    LogFile, UninstallLogfile : String;

    cmdlSaveInfFile, cmdlLoadInfFile: String;
    cmdlSilent, cmdlVerysilent, cmdlForceInstall: Boolean;

    desktopicon, desktopicon_common, desktopicon_user, quicklaunchicon: Boolean;


function IsX64: Boolean;
begin
 Result:=(ProcessorArchitecture=paX64);
end;

function GetDateString(Default:String):String;
begin
 result:=GetDateTimeString('yyyymmdd',#0,#0)
end;

function GetWizardSelectedComponents(Default:String):String;
begin
 Result:=WizardSelectedComponents(false);
end;

function GetWizardSelectedTasks(Default:String):String;
begin
 Result:=WizardSelectedTasks(false);
end;

function GetWizardSilent(Default:String):String;
begin
 if WizardSilent=true then
   Result:='true'
  else
   Result:='false';
end;

function GetWizardGroupValue(Default:String):String;
begin
 Result:=WizardGroupValue();
end;

function GetWizardNoIcons(Default:String):String;
begin
 if WizardNoIcons=true then
   Result:='true'
  else
   Result:='false';
end;

function GetWizardSetupType(Default:String):String;
begin
 Result:=WizardSetupType(false)
end;

function GetActiveLanguage(Default:string):String;
begin
 Result:=ActiveLanguage();
end;

function FileInPath(Filename:String; Path:String):Boolean;
var
 Buffer: String;
 res, BufferLength, pFilePart: LongInt;
begin
 res:=0;
 if length(Filename)>0 then begin
  BufferLength := 260;
  Buffer:=StringOfChar(#0,BufferLength);
  res:=SearchPath(Path, Filename, #0, BufferLength, Buffer, pFilePart);
 end;
 if res>0 then
   result:=true
  else
   result:=false;
end;

function GetIExplorerVersion(): String;
var
 sVersion:  String;
begin
 RegQueryStringValue(HKLM,'SOFTWARE\Microsoft\Internet Explorer', 'Version', sVersion );
 Result := sVersion;
end;

procedure DecodeVersion( verstr: String; var verint: TAInt);
var
  i,p: Integer; s: string;
begin
  verint := [0,0,0,0];
  i := 0;
  while ( (Length(verstr) > 0) and (i < 4) ) do
  begin
  	p := pos('.', verstr);
  	if p > 0 then
  	begin
      if p = 1 then s:= '0' else s:= Copy( verstr, 1, p - 1 );
  	  verint[i] := StrToInt(s);
  	  i := i + 1;
  	  verstr := Copy( verstr, p+1, Length(verstr));
  	end
  	else
  	begin
  	  verint[i] := StrToInt( verstr );
  	  verstr := '';
  	end;
  end;
end;

function IExplorerVersionGreater3(): Boolean;
var vers: TAInt;
begin
 DecodeVersion(GetIExplorerVersion,vers);
 if vers[0]<4 then
   Result:=false
  else
   Result:=true;
end;

function IExplorerVersionLower55(): Boolean;
var vers: TAInt;
begin
 DecodeVersion(GetIExplorerVersion,vers);
 if vers[0]<5 then
   Result:=false
  else
   if (vers[0]=5) and (vers[1]<5) then
     Result:=false
    else
     Result:=true;
end;

function ProgramIsInstalled(): Boolean;
begin
 if RegKeyExists(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}')=true then
   Result:=true
  else
   Result:=false;
end;

procedure SaveInstallInformations;
begin
 SaveStringToFile(LogFile, '--------------------------------------'#13#10#13#10, True);
 SaveStringToFile(LogFile, 'Windowsversion: '+GetWindowsVersionString+#13#10, True);
 SaveStringToFile(LogFile, 'WinDir: '+GetWinDir+#13#10, True);
 If IsWin64 then
   SaveStringToFile(LogFile, 'Win64: true'+#13#10, True)
  else
   SaveStringToFile(LogFile, 'Win64: false'+#13#10, True);

 Case ProcessorArchitecture of
  paUnknown:
   SaveStringToFile(LogFile, 'ProcessorArchitecture: Unknown'+#13#10, True);
  paX86:
   SaveStringToFile(LogFile, 'ProcessorArchitecture: X86'+#13#10, True);
  paX64:
   SaveStringToFile(LogFile, 'ProcessorArchitecture: X64'+#13#10, True);
  paIA64:
   SaveStringToFile(LogFile, 'ProcessorArchitecture: IA64'+#13#10, True);
 end;

 SaveStringToFile(LogFile, 'SystemDir: '+GetSystemDir+#13#10, True);
 SaveStringToFile(LogFile, 'TempDir: '+GetTempDir+#13#10, True);
 SaveStringToFile(LogFile, 'CurrentDir: '+GetCurrentDir+#13#10, True);
 SaveStringToFile(LogFile, 'Computername: '+GetComputernameString+#13#10, True);
 SaveStringToFile(LogFile, 'Username: '+GetUsernameString+#13#10, True);
 SaveStringToFile(LogFile, 'UILanguage: '+IntToStr(GetUILanguage)+#13#10, True);
 SaveStringToFile(LogFile, 'Internet Explorer version: '+GetIExplorerVersion+#13#10, True);
 SaveStringToFile(LogFile, 'Path: '+Getenv('Path')+#13#10, True);
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
 Result:=False;
 if CurPageID=wpReady then begin
  LogFile:=ExpandConstant('{app}')+'\SetupLog.txt';
 end;
 if CurPageID=wpFinished then
  SaveInstallInformations;
 Result:=True;
end;

function GetInstalledVersion(): String;
var
 instVersion:String;
begin
 if RegQueryStringValue(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}', 'ApplicationVersion', instVersion)=true then begin
   Result:=instVersion;
  end else begin
   Result:='0.0.0';
 end;
end;

function GetInstalledVersionBeta(): String;
var
 instVersion:String;
begin
 if RegQueryStringValue(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}', 'ApplicationVersion', instVersion)=true then
   Result:=instversion
  else
   Result:='0.0.0';
end;

procedure InitMessages();
var
 tmsg:String;
begin
 setArraylength(msg,11);
 tmsg:=ExpandConstant('{cm:ProgramIsRunning}');
 StringChange(tmsg,'%1','Transtool.exe');
 Msg[8]:=tmsg;
end;

function IsFullInstallation(): Boolean;
begin
 result:=FullInstallation;
end;

function CompletePath(Path: String): String;
begin
 if Copy(Path,Length(Path),1)<>'\' then
   result:=Path + '\'
  else
   result:=Path;
end;

procedure Split(Expression, Delimeter : String; var res :TArrayOfString);
var
 i, l :Integer; sl :tStringList;
begin
 SetArrayLength(res, 0);
 sl:=tStringList.Create;
 try
  l:=Length(Delimeter);
  while Length(Expression)>0 do begin
   i:=pos(Delimeter, Expression);
   if i <= 0 then
    i:=Length(Expression)+1;
   sl.Add (Copy(Expression, 1, i-1));
   Delete(Expression, 1, i+l-1);
  end;
  SetArrayLength (res, sl.Count);
  for i:=0 to (sl.Count-1) do
   res[i]:=sl[i];
 finally
  sl.free;
 end;
end;

function AnalyzeCommandlineParameters:Boolean;
var
 i:Longint; cmdParam, pStr: String;
begin
 Result:=false;
 for i:=0 to Paramcount do begin
  if Length(paramstr(i))=1 then begin
   Msgbox('False commandline parameter: ' + paramstr(i),mbError,MB_OK);
   exit;
  end;
  if (paramstr(i)='-?') or (paramstr(i)='/?') then begin
   Msgbox('Additional setup commandline parameters: '#13#10#13#10 +
    '/?'#9#9#9#9'- this help screen'#13#10 +
    '/ForceInstall'#9#9#9'- force the installation'#13#10
    ,mbInformation,MB_OK);
   exit;
  end;

  if uppercase(paramstr(i))='/VERYSILENT' then
   cmdlVerySilent:=true;
  if uppercase(paramstr(i))='/SILENT' then
   cmdlSilent:=true;
  if uppercase(paramstr(i))='/FORCEINSTALL' then
   cmdlForceInstall:=true;

  cmdParam:='/LoadInf';
  pStr:=Copy(paramstr(i),1,Length(cmdParam));
  if uppercase(pstr)=uppercase(cmdParam) then begin
   if Copy(paramstr(i),Length(cmdParam)+1,1)='=' then
     cmdlLoadInfFile:=Copy(paramstr(i),Length(cmdParam)+2,Length(paramstr(i)))
    else
     cmdlLoadInfFile:=Copy(paramstr(i),Length(cmdParam)+1,Length(paramstr(i)));
  end;
  if Length(cmdlLoadInfFile)>0 then
   if Length(ExtractFilePath(cmdlLoadInfFile))=0 then
    cmdlLoadInfFile:=CompletePath(GetCurrentDir) + cmdlLoadInfFile;

  cmdParam:='/SaveInf';
  pStr:=Copy(paramstr(i),1,Length(cmdParam));
  if uppercase(pstr)=uppercase(cmdParam) then begin
   if Copy(paramstr(i),Length(cmdParam)+1,1)='=' then
     cmdlSaveInfFile:=Copy(paramstr(i),Length(cmdParam)+2,Length(paramstr(i)))
    else
     cmdlSaveInfFile:=Copy(paramstr(i),Length(cmdParam)+1,Length(paramstr(i)));
  end;
  if Length(cmdlSaveInfFile)>0 then
   if Length(ExtractFilePath(cmdlSaveInfFile))=0 then
    cmdlSaveInfFile:=CompletePath(GetCurrentDir) + cmdlSaveInfFile;
 end;
 Result:=true;
end;

function UseDesktopIcon: boolean;
begin
 result:=desktopicon;
end;

function UseDesktopiconCommon: boolean;
begin
 result:=desktopicon_common;
end;

function UseDesktopiconUser: boolean;
begin
 result:=desktopicon_user;
end;

function UseQuickLaunchIcon: boolean;
begin
 result:=quicklaunchicon;
end;

procedure LoadInf;
var tasks:string; atasks:TArrayOfString; i:LongInt;
begin
 tasks:='';
 desktopicon:=false;
 desktopicon_common:=false;
 desktopicon_user:=false;
 quicklaunchicon:=false;
 if IniKeyExists('Setup','Tasks',cmdlLoadInfFile) then
  tasks:=GetIniString('Setup', 'Tasks', tasks, cmdlLoadInfFile);
 if length(tasks)>0 then
  Split(tasks,',',atasks);
 for i:=0 to GetArrayLength(atasks)-1 do
  Case lowercase(atasks[i]) of
   'desktopicon':        desktopicon:=true;
   'desktopicon\common': desktopicon_common:=true;
   'desktopicon\user':   desktopicon_user:=true;
   'quicklaunchicon':    quicklaunchicon:=true;
  end
end;

procedure SaveInf;
var res: boolean; tasks: String;
begin
 if IsTaskSelected('desktopicon') then
  tasks:='desktopicon';
 if IsTaskSelected('desktopicon\common') then
  tasks:=tasks + ',desktopicon\common';
 if IsTaskSelected('desktopicon\user') then
  tasks:=tasks + ',desktopicon\user';
 if IsTaskSelected('quicklaunchicon') then
  tasks:=tasks + ',quicklaunchicon';
 if length(tasks)>0 then begin
  if copy(tasks,1,1)=',' then
   tasks:=copy(tasks,2,length(tasks)-1);
  res:=SetIniString('Setup', 'Tasks', tasks, cmdlSaveInfFile)
 end
end;

function InitializeSetup(): Boolean;
var
#ifdef UpdateIsPossible
 cv,a:Longint;  verySilent:boolean;
#else
 a:Longint;
#endif
begin
 InitMessages;

 desktopicon:=true;
 desktopicon_common:=true;

 If AnalyzeCommandlineParameters=false then begin
  result:=false;
  exit
 end;

 If cmdlLoadInfFile<>'' then LoadInf;

 if CheckForMutexes('{#TransToolExeIDStr}')=true then begin
  Repeat
   a:=msgbox(msg[8],mbInformation, MB_OKCancel);
  until (a=IDCancel) or (CheckForMutexes('{#TransToolExeIDStr}')=false);
  if a=IDCancel then exit;
 end;

 Result:=True;
end;

procedure InitializeWizard();
begin
#IFDEF IncludeToolbar
 If InstallOnThisVersion('4.1.1998,5.0.2195','0,6.0')=irInstall then // Not Win95, Not WinNT4, Not Vista
  ToolbarForm_CreatePage(wpSelectDir);
#ENDIF
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
  S: String;
begin
  S := MemoUserInfoInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoDirInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoTypeInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoComponentsInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoGroupInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoTasksInfo;
  Result := S;
end;

function InitializeUninstall(): Boolean;
begin
 UninstallLogFile:=ExpandConstant('{%tmp}')+'\TransToolUninstall.txt';
 SaveStringToFile(UninstallLogFile, 'Start uninstall:' + #13#10, False)
 Result:=True;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
 tStr :String; i:LongInt; silent, verysilent:boolean;
begin
  case CurUninstallStep of
    usUninstall:
      begin
       tStr:=ExpandConstant('{app}')+'\Unload.tmp';
       if fileexists(tStr)=false then
        SaveStringToFile(tStr, '', True);
       silent:=false; verysilent:=false;
       for i:=1 to paramcount do begin
        if lowercase(ParamStr(i))='/silent' then
         silent:=true;
        if lowercase(ParamStr(i))='/verysilent' then
         verysilent:=true;
       end;
       SaveStringToFile(UninstallLogFile, ' Uninstall options:' + #13#10, True)
       if silent=true then
         SaveStringToFile(UninstallLogFile, '  Silent=True' + #13#10, True)
        else
         SaveStringToFile(UninstallLogFile, '  Silent=False' + #13#10, True);
       if verysilent=true then
         SaveStringToFile(UninstallLogFile, '  Verysilent=True' + #13#10, True)
        else
         SaveStringToFile(UninstallLogFile, '  Veryilent=False' + #13#10, True);
      end;
  end;
end;

//Only for debugging.
//#expr savetofile("TransTool-debug.ini")
