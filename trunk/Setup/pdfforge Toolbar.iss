; pdfforge Toolbar Installation
; Setup created with Inno Setup QuickStart Pack 5.3.9 (with ISPP) and ISTool 5.3.0.1
; Installation from pdfforge

;#define Test

#ifdef Test
 #define FastCompilation
#else
; #define FastCompilation
#endif

#define ProgramLicense "GNU"

#ifdef FastCompilation
 #define CompressionMode="none"
 #define SetupLZMACompressionMode "none"
#else
 #define CompressionMode="lzma/ultra"
 #define SetupLZMACompressionMode "ultra"
#endif

#define GetFileVersionVBExe(str S)     Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + Local[3] + Local[4]
#define GetFileVersionVBExeLine(str S) Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])-1), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])-1), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + '_' + Local[3] + '_'  + Local[4]

#define Company              "pdfforge GbR"
#define Homepage             "http://www.pdfforge.org"
#define SourceforgeHomepage  "http://www.sf.net/projects/pdfcreator"
#define Appname              "pdfforge Toolbar"
#define AppExename           "pdfforgeToolbar-stub-1.exe"

;#define Toolbar         "..\Toolbar\pdfforgeToolbar128.exe"
#define Toolbar         "..\Toolbar\" + AppExename

#define AppVersion           GetFileVersionVBExe(Toolbar)

#define PDFCreatorVersion    GetFileVersionVBExe(Toolbar)
#define SetupAppVersion      GetFileVersionVBExeLine(Toolbar)

#define AppID                "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01E}"
#define AppIDStr             "{" + AppID
#define AppIDreg             "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01E%7d"
#define AppVersionStr        AppVersion
#define SetupAppVersionStr   SetupAppVersion
#define UninstallID          "{B8B0FC8B-E69B-4215-AF1A-4BDFF20D794B}"

#define ChannelID 302398

#include "ToolbarForm.isd"

[Setup]
AllowNoIcons=true
AlwaysRestart=false
AppContact={#Homepage}
AppCopyright=© {#Company}
AppID={#AppIDStr}
AppName={#AppName}
AppVerName={#AppName} {#AppVersionStr}
AppPublisher={#Company}
AppPublisherURL={#Homepage}
AppSupportURL={#Homepage}
AppUpdatesURL={#Homepage}
AppVersion={#AppVersion}
ArchitecturesAllowed=x86 x64
ChangesAssociations=true
Compression=none
CreateAppDir=false
CreateUninstallRegKey=false
DisableDirPage=true
DisableStartupPrompt=true
ExtraDiskSpaceRequired=10303775

OutputBaseFilename={#AppName}-{#SetupAppVersionStr}_setup
OutputDir=Installation
RestartIfNeededByRun=true
ShowLanguageDialog=true
ShowTasksTreeLines=false
ShowUndisplayableLanguages=true
SolidCompression=true
UsePreviousAppDir=true

VersionInfoVersion={#AppVersion}
VersionInfoCompany={#Company}
VersionInfoDescription=pdfforge Toolbar
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}
VersionInfoTextVersion={#AppVersion}

WizardImageFile=..\Pictures\Setup\pdfforgeToolbarBig.bmp
WizardSmallImageFile=..\Pictures\Setup\PDFCreator.bmp

MinVersion=0,5.0.2195

[Files]
; Toolbar
Source: ..\Pictures\Toolbar\Toolbar.bmp; DestDir: {tmp}; Flags: dontcopy nocompression; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,0
Source: {#Toolbar}; DestDir: {tmp}; DestName: {#AppExename}; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,0

[Messages]
;Remove the 'StatusRunProgram' message
StatusRunProgram=

[Languages]
#include "languages.inc"

[CustomMessages]
#include "custommessages.inc"

[Run]
Filename: {tmp}\{#AppExename}; Parameters: "/S /V""/qn CHANNEL_ID={#ChannelID} D_WSD=1"" /UM""http://download.mybrowserbar.com/vkits/dlv1/{#ChannelID}/pdfforgeToolbar.msi"""; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,0; Check: Not DontUseYahooSearch
Filename: {tmp}\{#AppExename}; Parameters: "/S /V""/qn CHANNEL_ID={#ChannelID} D_WSD=0"" /UM""http://download.mybrowserbar.com/vkits/dlv1/{#ChannelID}/pdfforgeToolbar.msi"""; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,0; Check: DontUseYahooSearch

[Code]
var
 cmdlDontUseYahooSearch: Boolean;
 LogFile : String;

function ProgramIsInstalled(): Boolean;
begin
 if RegKeyExists(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}')=true then
   Result:=true
  else
   Result:=false;
end;

function DontUseYahooSearch:Boolean;
begin
 Result:=Not chkUseYahoo.Checked;
end;

function AnalyzeCommandlineParameters:Boolean;
var
 i:Longint;
begin
 Result:=false;
 for i:=0 to Paramcount do begin
  if Length(paramstr(i))=1 then begin
   Msgbox('False commandline parameter: ' + paramstr(i),mbError,MB_OK);
   exit;
  end;
  if (paramstr(i)='-?') or (paramstr(i)='/?') then begin
   Msgbox('Additional setup commandline parameters: '#13#10#13#10 +
    '/? - this help screen'#13#10 +
    '/DontUseYahooSearch - Don''t use Yahoo search if installing pdfforge Toolbar'
    ,mbInformation,MB_OK);
   exit;
  end;
  if uppercase(paramstr(i))='/DONTUSEYAHOOSEARCH' then
   cmdlDontUseYahooSearch:=true;
 end;
 Result:=true;
end;

function InitializeSetup(): Boolean;
var
 msg : String;
begin
 If AnalyzeCommandlineParameters=false then begin
  result:=false;
  exit
 end;
 msg:=ExpandConstant('{cm:AlreadyInstalled}');

 If ToolbarIsInstalled then begin
  msgbox(msg,mbInformation, MB_OK);
  Result:=false;
  exit
 end;

 SetupApplication := 'Toolbar';

 Result:=true;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
 Result:=False;
 if CurPageID=wpReady then begin
  LogFile:=ExpandConstant('{tmp}')+'\pdfforge Toolbar - SetupLog.txt';
 end;
 Result:=True;
end;

procedure InitializeWizard();
begin
 ToolbarPage := ToolbarForm_CreatePage(wpSelectDir);
 if (cmdlDontUseYahooSearch) then
  chkUseYahoo.Checked := false;
end;

function GetPDFCreatorToolbar1InstallLocation : String;
var
 uninstallStr, installLocation, resS : String;
begin
 uninstallStr := 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PDFCreator Toolbar';
 installLocation := 'InstallLocation';
 if RegQueryStringValue(HKEY_LOCAL_MACHINE, uninstallStr, installLocation, resS) then
  result := resS;
end;

function GetPDFCreatorToolbar1DllInstallLocation : String;
var
 uninstallStr, installLocation, version, resIlS, resVeS : String;
begin
 uninstallStr := 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PDFCreator Toolbar';
 installLocation := 'InstallLocation';
 version := 'DisplayVersion';
 if RegQueryStringValue(HKEY_LOCAL_MACHINE, uninstallStr, installLocation, resIlS) then
  if RegQueryStringValue(HKEY_LOCAL_MACHINE, uninstallStr, version, resVeS) then
   result := resIlS + '\v' + resVeS;
end;

procedure UnregisterPDFCreatorToolbar1;
var
 resS, dll : String;
 res : Boolean;
begin
 resS := GetPDFCreatorToolbar1DllInstallLocation;
 dll := resS + '\PDFCreator_Toolbar.dll';
 if Length(resS) > 0 then
  if DirExists(resS) then
   if FileExists(dll) then
    begin
     res := UnregisterServer(false, dll,false);
     if (res = true) then
       SaveStringToFile(LogFile, 'Unregister ' + dll + ' res = true' + #13#10, True)
      else
       SaveStringToFile(LogFile, 'Unregister ' + dll + ' res = false' + #13#10, True);
    end;
end;

procedure DeletePDFCreatorToolbar1FirefoxExtension;
var
 str1, str2, resS, resS2 : String;
 res : Boolean;
begin
 str1 := 'SOFTWARE\Mozilla\Mozilla Firefox';
 str2 := 'CurrentVersion';
 if RegQueryStringValue(HKEY_LOCAL_MACHINE, str1, str2, resS) then
  if Length(resS) > 0 then
   if RegQueryStringValue(HKEY_LOCAL_MACHINE, str1 + '\' + resS + '\Main', 'Install Directory', resS2) then
    if FileExists(resS2 + '\extensions\support@pdfcreator-toolbar.org') then begin
      res := DeleteFile(resS2 + '\extensions\support@pdfcreator-toolbar.org');
      if res then
        SaveStringToFile(LogFile, 'Toolbar Firefox extension succesfully deleted.' + #13#10, True)
       else
        SaveStringToFile(LogFile, 'Can''t delete toolbar Firefox extension.' + #13#10, True);
     end;
end;

procedure DeletePDFCreatorToolbar1Diretory;
var
 resS : String;
 res : Boolean;
begin
 resS := GetPDFCreatorToolbar1InstallLocation;
 if Length(resS) > 0 then
  if DirExists(resS) then
   res := DelTree(resS, true, true, true);
end;

procedure DeleteUninstallFile;
var
 uninstallStr, uninstallExe, resS : String;
 exeP : Integer;
 res : Boolean;
begin
 uninstallStr := 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PDFCreator Toolbar';
 uninstallExe := 'UninstallString';
 if RegQueryStringValue(HKEY_LOCAL_MACHINE, uninstallStr, uninstallExe, resS) then
  begin
   resS := AnsiLowercase(resS);
   exeP := Pos('.exe', resS);
   if exeP > 1 Then
   begin
    resS := Copy(resS, 2, exeP + 2);
    if FileExists(resS) then
     res := DeleteFile(resS);
   end;
  end;
end;

procedure DeleteStartMenuEntry;
var
 entry, s1: String;
 res : Boolean;
begin
 entry := ExpandConstant('{commonprograms}') + '\PDFCreator Toolbar';
 s1 := 'Delete toolbar start menu entry: ';
 if DirExists(entry) then begin
   res := DelTree(entry, true, true, true);
   if (res = true) then
     SaveStringToFile(LogFile, s1 + 'true' + #13#10, True)
    else
     SaveStringToFile(LogFile, s1 + 'false' + #13#10, True)
  end else
   SaveStringToFile(LogFile, s1 + 'not found.' + #13#10, True);
end;

procedure RemoveRegistrySettings;
var
 uninstallStr : String;
 rootKey : LongInt;
begin
 rootKey := HKEY_CURRENT_USER;
 uninstallStr := 'Software\Microsoft\Windows\CurrentVersion\Explorer\MenuOrder\Start Menu\Programs\PDFCreator Toolbar';
 SaveStringToFile(LogFile, 'Remove toolbar registry settings.' + #13#10, True);
 if RegKeyExists(rootKey, uninstallStr) then
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, uninstallStr)

 rootKey := HKEY_CURRENT_USER;
 uninstallStr := 'Software\Microsoft\Windows\CurrentVersion\Explorer\MenuOrder\Start Menu2\Programs\PDFCreator Toolbar';
 if RegKeyExists(rootKey, uninstallStr) then
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, uninstallStr)

 rootKey := HKEY_LOCAL_MACHINE;
 uninstallStr := 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PDFCreator Toolbar';
 if RegKeyExists(rootKey, uninstallStr) then
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, uninstallStr)

 rootKey := HKEY_LOCAL_MACHINE;
 uninstallStr := 'SOFTWARE\Microsoft\Internet Explorer\Low Rights\ElevationPolicy\{DCAAA846-F9B9-4E1C-B2FE-CD0045097E76}';
 if RegKeyExists(rootKey, uninstallStr) then
  RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, uninstallStr)
end;
