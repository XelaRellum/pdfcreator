; PDFCreator-Patch Installation
; Setup created with Inno Setup 5.0.5 Beta (ISPack with included ISPP) and ISTool 4.2.7
; Installation from Frank Heindörfer, Philip Chinery

;#define Test
#define SetupLZMACompressionMode "ultra"
;#define SetupLZMACompressionMode "fast"

#define ProgramLicense "GNU"

#define GetFileVersionVBExe(str S)     Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + Local[3] + Local[4]
#define GetFileVersionVBExeLine(str S) Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])-1), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])-1), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + '_' + Local[3] + '_'  + Local[4]

#define Homepage             "http://www.pdfcreator.de.vu"
#define Appname              "PDFCreator"
#define AppExename           "PDFCreator.exe"
#define SpoolerExename       "PDFSpooler.exe"

#define AppVersion           GetFileVersionVBExe("..\PDFCreator\PDFCreator.exe")

#define PDFCreatorVersion    GetFileVersionVBExe("..\PDFCreator\PDFCreator.exe")
#define SetupAppVersion      GetFileVersionVBExeLine("..\PDFCreator\PDFCreator.exe")
#define PDFSpoolerVersion    GetFileVersionVBExe("..\PDFSpooler\PDFSpooler.exe")
#define TransToolVersion     GetFileVersionVBExe("..\Transtool\Transtool.exe")

#define BetaVersion          "0"

#define AppVersionStr        AppVersion
#define SetupAppVersionStr   SetupAppVersion

#define AppID                "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D}"
#define AppIDStr             "{" + AppID
#define AppIDreg             "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D%7d"
#define PDFCreatorExeID      "{A7332D94-E8FE-40B2-937F-8515FC0FF52F}"
#define PDFCreatorExeIDstr   "{" + PDFCreatorExeID
#define TransToolExeID       "{B7BCA0D2-7305-4318-BA7A-01B028D910EB}"
#define TransToolExeIDStr    "{" + TransToolExeID
#define PDFSpoolerExeID      "{C387A397-047A-4354-AE89-F75B1B550257}"
#define PDFSpoolerExeIDStr   "{" + PDFSpoolerExeID
#define UninstallID          AppID
#define UninstallIDreg       AppIDreg
#define UninstallIDStr       "{"+ UninstallID
#define UninstallIDStr2      "{"+ UninstallIDreg

#define UninstallReg         "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallID
#define UninstallRegStr      "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallIDStr
#define UninstallRegStr2     "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallIDStr2

;#define UpdateIsPossible
#define UpdateIsPossibleMinVersion "0.8.1"

[_ISToolPreCompile]
#If (SetupLZMACompressionMode=="ultra")
;Name: .\upx\upx.exe; Parameters: ..\TransTool\TransTool.exe   -d
;Name: .\upx\upx.exe; Parameters: ..\PDFSpooler\PDFSpooler.exe -d
;Name: .\upx\upx.exe; Parameters: ..\PDFCreator\PDFCreator.exe -d

Name: .\upx\upx.exe; Parameters: ..\TransTool\TransTool.exe   --best --compress-icons=0 --crp-ms=999999
Name: .\upx\upx.exe; Parameters: ..\PDFSpooler\PDFSpooler.exe --best --compress-icons=0 --crp-ms=999999
Name: .\upx\upx.exe; Parameters: ..\PDFCreator\PDFCreator.exe --best --compress-icons=0 --crp-ms=999999

;Name: .\upx\upx.exe; Parameters: ..\TransTool\TransTool.exe   -3 --compress-icons=0 --crp-ms=999999
;Name: .\upx\upx.exe; Parameters: ..\PDFSpooler\PDFSpooler.exe -3 --compress-icons=0 --crp-ms=999999
;Name: .\upx\upx.exe; Parameters: ..\PDFCreator\PDFCreator.exe -3 --compress-icons=0 --crp-ms=999999
#Endif

[Setup]
AllowNoIcons=false
AlwaysRestart=false
AppCopyright=© 2002 - 2004 Philip Chinery, Frank Heindörfer
AppID={#AppIDStr}
AppName={#AppName}
AppVerName=Patch-{#AppName} {#AppVersionStr}
AppPublisher=Philip Chinery, Frank Heindörfer
AppPublisherURL={#Homepage}
AppSupportURL={#Homepage}
AppUpdatesURL={#Homepage}
ChangesAssociations=true
Compression=lzma/{#SetupLZMACompressionMode}
CreateUninstallRegKey=false
DefaultDirName={reg:HKLM\{#UninstallRegStr2},Inno Setup: App Path|{pf}\{#AppName}}
DefaultGroupName={#AppName}
DisableDirPage=true
DisableStartupPrompt=true
InternalCompressLevel={#SetupLZMACompressionMode}
;LicenseFile=.\License\GNU Readme.rtf
OutputBaseFilename=Patch-{#AppName}-{#SetupAppVersionStr}
OutputDir=Installation
RestartIfNeededByRun=true
ShowTasksTreeLines=false
SolidCompression=true
UsePreviousAppDir=true

VersionInfoVersion=0.8.1
VersionInfoCompany=Frank Heindörfer, Philip Chinery
VersionInfoDescription=PDFCreator is the easy way of creating PDFs.
VersionInfoTextVersion=0.8.1

WizardImageFile=..\Pictures\Setup\PDFCreatorBigPatch.bmp
WizardSmallImageFile=..\Pictures\Setup\PDFCreator.bmp
Uninstallable=false

[Files]
#IFNDEF Test
;Program files
Source: ..\PDFCreator\PDFCreator.exe; DestDir: {app}; Flags: comparetimestamp
Source: ..\Transtool\TransTool.exe; DestDir: {app}\languages; Flags: comparetimestamp
Source: ..\PDFSpooler\PDFSpooler.exe; DestDir: {sys}; Flags: comparetimestamp
#Endif

[Languages]
#include "languages.inc"

[CustomMessages]
#include "custommessages.inc"

[Code]
type
 TAInt = Array of Integer; TAStr = Array of String;
var
 msg : TAStr;
 Win9x, WinNT, Win2000 : String;

function ProgramIsInstalled(): Boolean;
begin
 if RegKeyExists(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}')=true then
   Result:=true
  else
   Result:=false;
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
 instVersion, BetaVersion:String;
begin
 if RegQueryStringValue(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}', 'ApplicationVersion', instVersion)=true then
   if RegQueryStringValue(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}', 'BetaVersion', BetaVersion)=true then
     if trim(BetaVersion)<>'' then
       Result:=instversion + ' Beta ' + BetaVersion
      else
       Result:=instversion
    else
     Result:=instversion
  else
   Result:='0.0.0';
end;

procedure DecodeVBVersion( verstr: String; var verint: TAInt);
var
  i,p: Integer; s: string;
begin
  // initialize array
  verint := [0,0,0];
  i := 0;
  while ((Length(verstr) > 0) and (i < 3)) do
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

// This function detect a Beta Update
// return 0 = equal beta number
// return 1 = major release or patch release
// return 2 = update a beta
// return 3 = no beta update possible
function BetaUpdate() : LongInt;
var InstBetaNumber, BetaNumber : LongInt;
    InstBetaNumberStr:String;
begin
 if RegQueryStringValue(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}', 'BetaVersion', InstBetaNumberStr)=true then begin
   if length(InstBetaNumberStr)>0 then
     InstBetaNumber:=StrToInt(InstBetaNumberStr)
    else
     InstBetaNumber:=0;
  end else begin
   InstBetaNumber:=0;
 end;
 BetaNumber:=StrToInt('{#BetaVersion}');
 If (InstBetaNumber=BetaNumber) then
   Result:=0 //equal
  else
   if '{#BetaVersion}'='' then
     Result:=1 //major release
    else
     if (InstBetaNumber<BetaNumber) and (InstBetaNumber>=0) then
       Result:=2  //beta update
      else
       Result:=3; //no beta update possible
end;

// This function compares VB version string
// return -1 if ver1 < ver2
// return  0 if ver1 = ver2
// return  1 if ver1 > ver2
function CompareVBVersion(ver1, ver2: String ) : Integer;
var
 verint1, verint2: TAInt; betaUpd:LongInt;
 i: integer;
begin
  SetArrayLength(verint1,3); DecodeVBVersion(ver1,verint1);
  SetArrayLength(verint2,3); DecodeVBVersion(ver2,verint2);
  Result:=0; i:=0;
  while ((Result=0) and (i<3)) do
  begin
  	if verint1[i] > verint2[i] then
  	  Result:=1
     else
      if verint1[i] < verint2[i] then
  	    Result:=-1
  	   else
  	    Result:=0;
  	i:=i+1;
  end;

 betaUpd:=BetaUpdate;
 If Result=0 then
  If (betaupd=1) or (betaUpd=2) then
     Result:=-1
    else
     If betaUpd=3 then
      Result:=1;
end;

procedure InitMessages();
var
 tmsg:String;
begin
 setArraylength(msg,14);
 Msg[0]:=ExpandConstant('{cm:NoAdmin}');

 tmsg:=ExpandConstant('{cm:OldVersion}');
 StringChange(tmsg,'%1',GetInstalledVersion);
 StringChange(tmsg,'%2','{#AppVersionStr}');
 Msg[1]:=tmsg;

 Msg[2]:=ExpandConstant('{cm:NoNoAdmin}');
 Msg[3]:=ExpandConstant('{cm:Update}');
 Msg[4]:=ExpandConstant('{cm:AlreadyInstalled}');

 tmsg:=ExpandConstant('{cm:NewerVersion}');
 StringChange(tmsg,'%1',GetInstalledVersion);
 StringChange(tmsg,'%2','{#AppVersionStr}');
 Msg[5]:=tmsg;

 Msg[6]:=ExpandConstant('{cm:AlreadyInstalledNoUpdate}');

 tmsg:=ExpandConstant('{cm:ProgramIsRunning}');
 StringChange(tmsg,'%1','PDFCreator.exe');
 Msg[7]:=tmsg;

 tmsg:=ExpandConstant('{cm:ProgramIsRunning}');
 StringChange(tmsg,'%1','Transtool.exe');
 Msg[8]:=tmsg;

 tmsg:=ExpandConstant('{cm:ProgramIsRunning}');
 StringChange(tmsg,'%1','PDFSpooler.exe');
 Msg[9]:=tmsg;

 tmsg:=ExpandConstant('{cm:NoUpdate}');
 StringChange(tmsg,'%1',GetInstalledVersionBeta);
 StringChange(tmsg,'%2','{#AppVersionStr}');
 Msg[10]:=tmsg;
 Msg[11]:=ExpandConstant('{cm:PatchProgramIsNotInstalled}');
 Msg[12]:=ExpandConstant('{cm:PatchProgramIsTooOld}');
 Msg[13]:=ExpandConstant('{cm:PatchProgramIsTooNew}');
end;

function InitializeSetup(): Boolean;
var
 cv,a:Longint; verySilent:boolean;
begin
 InitMessages;
 Win9x:='Windows 95, Windows 98, Windows Me';
 WinNt:='Windows NT 4.0';
 Win2000:='Windows 2000';
 verySilent:=false;

 if CheckForMutexes('{#PDFCreatorExeIDStr}')=true then begin
  Repeat
   a:=msgbox(msg[7],mbInformation, MB_OKCancel);
  until (a=IDCancel) or (CheckForMutexes('{#PDFCreatorExeIDStr}')=false);
  if a=IDCancel then exit;
 end;
 if CheckForMutexes('{#TransToolExeIDStr}')=true then begin
  Repeat
   a:=msgbox(msg[8],mbInformation, MB_OKCancel);
  until (a=IDCancel) or (CheckForMutexes('{#TransToolExeIDStr}')=false);
  if a=IDCancel then exit;
 end;
 if CheckForMutexes('{#PDFSpoolerExeIDStr}')=true then begin
  Repeat
   a:=msgbox(msg[9],mbInformation, MB_OKCancel);
  until (a=IDCancel) or (CheckForMutexes('{#PDFSpoolerExeIDStr}')=false);
  if a=IDCancel then exit;
 end;
 for a:=1 to Paramcount do begin
  if uppercase(paramstr(a))='/VERYSILENT' then
   verySilent:=true;
 end;
 If ProgramIsInstalled=true then begin
   cv:=CompareVBVersion(GetInstalledVersion,'{#AppVersion}');
   if cv=0 then begin
    cv:=CompareVBVersion(GetInstalledVersion,'{#UpdateIsPossibleMinVersion}');
    if cv=-1 then begin
      Result:=false;
      msgbox(msg[6],mbConfirmation, MB_OKCancel);
     end else
      Result:=true;
    cv:=-2;
   end;
   if cv=-1 then begin
    msgbox(msg[12],mbInformation, MB_OK);
    Result:=false
   end
   if cv=1 then begin
    msgbox(msg[13],mbInformation, MB_OK);
    Result:=false
   end
  end else begin
   msgbox(msg[11],mbInformation, MB_OK);
   Result:=false
  end;
end;
