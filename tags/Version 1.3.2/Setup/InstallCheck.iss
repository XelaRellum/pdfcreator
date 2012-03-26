[Setup]
AppId={{87FD7F54-DC9A-4311-A746-0C7DFE0A4756}
AppName=InstallCheck
AppVersion=0.1
CreateAppDir=no
OutputBaseFilename=InstallCheck
OutputDir=Installation
Compression=lzma
CreateUninstallRegKey=false
RestartIfNeededByRun=false
ShowLanguageDialog=false
SolidCompression=yes
Uninstallable=No

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Code]
function InternetGetConnectedState(var lpdwFlags: DWORD; dwReserved: DWORD): boolean;
 external 'InternetGetConnectedState@wininet.dll stdcall';
function URLDownloadToFile(pCaller:LongInt; szURL:String; szFileName:String; dwReserved: LongInt; lpfnCB: LongInt): LongInt;
 external 'URLDownloadToFileA@urlmon.dll stdcall';

const
 INTERNET_CONNECTION_OFFLINE = $20;

var
 cmdlVersion, cmdlLanguageCode : string;
 
function AnalyzeCommandlineParameters:Boolean;
var
 i:Longint; cmdParam, pStr: String;
begin
 Result:=false;
 for i:=0 to Paramcount do begin
  cmdParam:='/v';
  pStr:=Copy(paramstr(i),1,Length(cmdParam));
  if uppercase(pstr)=uppercase(cmdParam) then begin
   if Copy(paramstr(i),Length(cmdParam)+1,1)='=' then
     cmdlVersion:=Copy(paramstr(i),Length(cmdParam)+2,Length(paramstr(i)))
    else
     cmdlVersion:=Copy(paramstr(i),Length(cmdParam)+1,Length(paramstr(i)));
  end;
  cmdParam:='/lc';
  pStr:=Copy(paramstr(i),1,Length(cmdParam));
  if uppercase(pstr)=uppercase(cmdParam) then begin
   if Copy(paramstr(i),Length(cmdParam)+1,1)='=' then
     cmdlLanguageCode:=Copy(paramstr(i),Length(cmdParam)+2,Length(paramstr(i)))
    else
     cmdlLanguageCode:=Copy(paramstr(i),Length(cmdParam)+1,Length(paramstr(i)));
  end;
 end;
 
 Result:=true;
end;

function InitializeSetup(): Boolean;
var
 resB: Boolean;
 installCheckResultFile: String;
 ConnectionState: DWORD;
begin
 try
 cmdlVersion := '0.0.0'; cmdlLanguageCode := '-';
 AnalyzeCommandlineParameters();
 resB := InternetGetConnectedState(ConnectionState, 0);
 if (ConnectionState And INTERNET_CONNECTION_OFFLINE) <> INTERNET_CONNECTION_OFFLINE then begin
  installCheckResultFile := ExpandConstant('{tmp}') + '\installCheck.txt';
  if FileExists(installCheckResultFile) then DeleteFile(installCheckResultFile);
  UrlDownloadToFile(0, 'http://piwik.pdfforge.org/check.php?version=' + cmdlVersion + '&lang=' + cmdlLanguageCode, installCheckResultFile, 0, 0);
 end
 except
 end;
 result := false;
end;
