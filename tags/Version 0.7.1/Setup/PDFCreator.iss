; PDFCreator Installation
; Setup created with Inno Setup 4.0.5 beta, ISPP 1.1 Pre-release and ISTool 4.0.4 beta
; Installation from Frank Heindörfer, Philip Chinery

;#define Test

#define GetFileVersionVBExe(str S)     Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + Local[3] + Local[4]
#define GetFileVersionVBExeLine(str S) Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])-1), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])-1), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + '_' + Local[3] + '_'  + Local[4]

#define Homepage             "http://www.pdfcreator.de.vu"
#define Appname              "PDFCreator"
#define Printername          "PDFCreator"
#define Drivername           "PDFCreator"
#define Portname             "PDFCreator:"
#define Monitorname          "PDFCreator"
#define AppExename           "PDFCreator.exe"
#define SpoolerExename       "PDFSpooler.exe"

#define AppVersion           GetFileVersionVBExe("..\PDFCreator\PDFCreator.exe")

#define PDFCreatorVersion    GetFileVersionVBExe("..\PDFCreator\PDFCreator.exe")
#define SetupAppVersion      GetFileVersionVBExeLine("..\PDFCreator\PDFCreator.exe")
#define PDFSpoolerVersion    GetFileVersionVBExe("..\PDFSpooler\PDFSpooler.exe")
#define TransToolVersion     GetFileVersionVBExe("..\Transtool\Transtool.exe")
#define UnInstVersion        GetFileVersionVBExe("..\UnInst\UnInst.exe")
#define GhostscriptVersion   "8.00"

#define BetaVersion          ""

#IF (BetaVersion!="")
 #define AppVersionStr       AppVersion + " Beta " + BetaVersion
 #define SetupAppVersionStr  SetupAppVersion + "_" + "Beta_" + BetaVersion
#ELSE
 #define AppVersionStr       AppVersion
 #define SetupAppVersionStr  SetupAppVersion
#ENDIF

#define AppID                "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D}"
#define AppIDreg             "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D%7d"
#define PDFCreatorExeID      "{A7332D94-E8FE-40B2-937F-8515FC0FF52F}"
#define PDFCreatorExeIDstr   "{" + PDFCreatorExeID
#define TransToolExeID       "{B7BCA0D2-7305-4318-BA7A-01B028D910EB}"
#define TransToolExeIDStr    "{" + TransToolExeID
#define PDFSpoolerExeID      "{C387A397-047A-4354-AE89-F75B1B550257}"
#define PDFSpoolerExeIDStr   "{" + PDFSpoolerExeID
#define UnInstExeID          "{D95872D0-0DE7-4C01-859C-1BAE47FB1C6B}"
#define UnInstExeIDStr       "{" + UnInstExeID
#define UninstallID          AppID
#define UninstallIDreg       AppIDreg
#define UninstallIDStr       "{"+ UninstallID
#define UninstallIDStr2      "{"+ UninstallIDreg

#define UninstallReg         "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallID
#define UninstallRegStr      "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallIDStr
#define UninstallRegStr2     "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + UninstallIDStr2
#define PrintReg             "System\CurrentControlSet\Control\Print\"
#define PrintMonitorReg      PrintReg + "Monitors\" + Monitorname
#define PrintMonitorPortReg  PrintReg + "Monitors\" + Monitorname + "\Ports\" + Portname

#define UpdateIsPossible
#define UpdateIsPossibleMinVersion "0.7.0"

[_ISTool]
EnableISX=true

[_ISToolPreCompile]
Name: .\upx\upx.exe; Parameters: ..\TransTool\TransTool.exe   --best --compress-icons=0
Name: .\upx\upx.exe; Parameters: ..\PDFSpooler\PDFSpooler.exe --best --compress-icons=0
Name: .\upx\upx.exe; Parameters: ..\PDFCreator\PDFCreator.exe --best --compress-icons=0
Name: .\upx\upx.exe; Parameters: ..\UnInst\UnInst.exe         --best --compress-icons=0

[Setup]
AllowNoIcons=false
AlwaysRestart=false
AppCopyright=© 2002 - 2003 Philip Chinery, Frank Heindörfer
AppID={#AppID}
AppName={#AppName}
AppVerName={#AppName} {#AppVersionStr}
AppPublisher=Philip Chinery, Frank Heindörfer
AppPublisherURL={#Homepage}
AppSupportURL={#Homepage}
AppUpdatesURL={#Homepage}
ChangesAssociations=true
Compression=bzip/9
CreateUninstallRegKey=false
DefaultDirName={reg:HKLM\{#UninstallRegStr2},Inno Setup: App Path|{pf}\{#AppName}}
DefaultGroupName={#AppName}
DisableDirPage=false
DisableStartupPrompt=true
InternalCompressLevel=9
LicenseFile=.\License\readme.rtf
OutputBaseFilename={#AppName}-Setup-{#SetupAppVersionStr}
OutputDir=Installation
RestartIfNeededByRun=true
ShowTasksTreeLines=false
SolidCompression=true
UsePreviousAppDir=true
WizardImageFile=..\Pictures\PDFCreatorBig.bmp
WizardSmallImageFile=..\Pictures\PDFCreator.bmp

[InstallDelete]
Name: {app}\fonts\*.*; Type: filesandordirs; Tasks: ghostscript
Name: {app}\lib\*.*; Type: filesandordirs; Tasks: ghostscript
Name: {app}\gsdll32.dll; Type: files; Tasks: ghostscript
Name: {app}\languages\*.ini; Type: files; Components: program

[Files]
#IFNDEF Test
;Systemfiles
Source: ..\SystemFiles\ASYCFILT.DLL; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace uninsneveruninstall
Source: ..\SystemFiles\COMDLG32.OCX; DestDir: {sys}; Components: program; Flags: sharedfile regserver
Source: ..\SystemFiles\MSCOMCT2.OCX; DestDir: {sys}; Components: program; Flags: sharedfile regserver
Source: ..\SystemFiles\MSCOMCTL.OCX; DestDir: {sys}; Components: program; Flags: sharedfile regserver
Source: ..\SystemFiles\MSVBVM60.DLL; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace regserver uninsneveruninstall
Source: ..\SystemFiles\OLEPRO32.DLL; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace regserver uninsneveruninstall
Source: ..\SystemFiles\OLEAUT32.DLL; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace regserver uninsneveruninstall
Source: ..\SystemFiles\STDOLE2.TLB; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace uninsneveruninstall regtypelib

;Systemfiles Language: German
;http://msdn.microsoft.com/vbasic/downloads/tools/ipdk.aspx
Source: C:\IPDK\German\CMDLGDE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile
Source: C:\IPDK\German\MSCC2DE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile
Source: C:\IPDK\German\MSCMCDE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile
Source: C:\IPDK\German\VB6DE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile

;Ghostscript files
Source: ..\PDFCreator\Fonts\*.*; DestDir: {app}\fonts; Components: program; Tasks: ghostscript; Flags: ignoreversion
Source: ..\PDFCreator\Lib\*.*; DestDir: {app}\lib; Components: program; Tasks: ghostscript; Flags: ignoreversion
Source: ..\PDFCreator\gsdll32.dll; DestDir: {app}; Components: program; Tasks: ghostscript

;Program files
Source: ..\PDFCreator\PDFCreator.exe; DestDir: {app}; Components: program; Flags: comparetimestamp
Source: .\License\License.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp
Source: ..\PDFCreator\Languages\*.ini; DestDir: {app}\languages; Components: program; Flags: ignoreversion comparetimestamp
Source: ..\Transtool\TransTool.exe; DestDir: {app}\languages; Components: program; Flags: comparetimestamp
Source: ..\UnInst\UnInst.exe; DestDir: {app}; Components: program; Flags: comparetimestamp

Source: PDFCreator.ini; DestDir: {userappdata}\PDFCreator; Components: program; DestName: PDFCreator.ini; Flags: ignoreversion onlyifdoesntexist uninsneveruninstall
Source: History.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp
Source: ..\PDFSpooler\PDFSpooler.exe; DestDir: {app}; Components: printer; Flags: comparetimestamp

;ShFolder for old systems
;http://www.microsoft.com/downloads/release.asp?releaseid=30340
Source: ShFolder\ShFolder.Exe; DestDir: {app}; Components: program; Flags: ignoreversion deleteafterinstall; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 4.1.2222,5.0.2195

;Printer files
; for Win9x/Me
Source: ..\Printer\Win98\MSIMGSIZ.DAT; DestDir: {win}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Win98\$00379C1.WPX; DestDir: {win}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Win98\AD59A730.MFD; DestDir: {win}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Win98\FONTSDIR.MFD; DestDir: {win}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Win98\System\Psmon.dll; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0
Source: ..\Printer\Win98\System\Adfonts.mfm; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Win98\System\Adobeps4.drv; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0
Source: ..\Printer\Win98\System\Adobeps4.hlp; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Win98\System\Defprtr2.ppd; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion

; for WinNt/2000/XP
Source: ..\Printer\WinNt\PDFCreator.PPD; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.00.1381
Source: ..\Printer\WinNt\AdobePS5.dll; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.00.1381
Source: ..\Printer\WinNt\AdobePS5.ntf; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.00.1381
Source: ..\Printer\WinNt\adobepsu.dll; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.00.1381
Source: ..\Printer\WinNt\Adobepsu.hlp; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.00.1381
Source: ..\Printer\WinNt\Defprtr2.ppd; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.00.1381

;Redmon files
Source: ..\Printer\Redmon\redmonnt.dll; DestDir: {sys}; Components: printer; Flags: sharedfile restartreplace; MinVersion: 0,4.00.1381
Source: ..\Printer\Redmon\redmon95.dll; DestDir: {sys}; Components: printer; Flags: sharedfile restartreplace; MinVersion: 4.00.950,0
#ENDIF

[Icons]
Name: {group}\{#Appname}; Filename: {app}\{#AppExename}; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\License; Filename: {app}\License.txt; Flags: createonlyiffileexists
Name: {group}\History; Filename: {app}\History.txt; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\Translation Tool; Filename: {app}\languages\transtool.exe; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\Uninstall {#Appname}; Filename: {uninstallexe}; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\PDFCreator Homepage; Filename: {app}\PDFCreator.url

Name: {commondesktop}\PDFCreator; Filename: {app}\pdfcreator.exe; Tasks: desktopicon\common
Name: {userdesktop}\PDFCreator; Filename: {app}\pdfcreator.exe; Tasks: desktopicon\user
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\PDFCreator; Filename: {app}\pdfcreator.exe; Tasks: quicklaunchicon

[INI]
Filename: {app}\PDFCreator.url; Section: InternetShortcut; Key: URL; String: http://www.pdfcreator.de.vu/
Filename: {app}\PDFCreator.url; Section: InternetShortcut; Key: Iconindex; String: 1; Components: program
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: AutosaveDirectory; String: {userdocs}; Components: program
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: LastsaveDirectory; String: {userdocs}; Components: program
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: Language; String: english; Components: program; Languages: English
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: Language; String: deutsch; Components: program; Languages: German

[Registry]
;PrinterMonitor
Root: HKLM; Subkey: {#PrintMonitorReg}; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: string; Valuename: Arguments; ValueData: -PPDFCREATORPRINTER; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: string; Valuename: Command; ValueData: {app}\{#SpoolerExename}; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: dword; Valuename: Delay; ValueData: 300; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: string; Valuename: Description; ValueData: Redirected Port; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: dword; Valuename: LogFileDebug; ValueData: 0; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: dword; Valuename: LogFileUse; ValueData: 0; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: dword; Valuename: Output; ValueData: 0; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: string; Valuename: Printer; ValueData: {#Printername}; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: dword; Valuename: Printerror; ValueData: 0; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: dword; Valuename: Runuser; ValueData: 1; Flags: uninsdeletevalue; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; ValueType: dword; Valuename: ShowWindow; ValueData: 0; Flags: uninsdeletevalue; Components: printer
;PrinterPort
Root: HKLM; Subkey: {#PrintMonitorPortReg}; Components: printer

;Uninstall - Deletekey
Root: HKLM; Subkey: {#PrintReg}Printers\{#Printername}; Flags: uninsdeletekey dontcreatekey; Components: printer
Root: HKLM; Subkey: {#PrintReg}Environments\Windows 4.0\Drivers\{#Drivername}; Flags: uninsdeletekey dontcreatekey; MinVersion: 4.00.950,0; Components: printer
Root: HKLM; Subkey: {#PrintReg}Environments\Windows NT x86\Drivers\{#Drivername}; Flags: uninsdeletekey dontcreatekey; MinVersion: 0,4.00.1381; Components: printer
Root: HKLM; Subkey: {#PrintMonitorPortReg}; Flags: uninsdeletekey dontcreatekey; Components: printer
Root: HKLM; Subkey: {#PrintMonitorReg}; Flags: uninsdeletekey dontcreatekey; Components: printer

;File-Assoc
Root: HKCR; SubKey: .ps; ValueType: string; ValueData: PostScript; Flags: uninsdeletekey; Tasks: fileassoc
Root: HKCR; SubKey: PostScript; ValueType: string; ValueData: PostScript; Flags: uninsdeletekey; Tasks: fileassoc
Root: HKCR; SubKey: PostScript\Shell\Open\Command; ValueType: string; ValueData: """{app}\PDFCreator.exe"" -IF""%1"""; Flags: uninsdeletevalue; Tasks: fileassoc
Root: HKCR; Subkey: PostScript\DefaultIcon; ValueType: string; ValueData: {app}\PDFCreator.exe,0; Flags: uninsdeletevalue; Tasks: fileassoc
Root: HKU; Subkey: .DEFAULT\Software\Microsoft\Windows\ShellNoRoam\MUICache; ValueType: string; Valuename: {app}\PDFCreator.exe; ValueData: Create PDF and Bitmap Files with {#Appname}; Tasks: fileassoc; Languages: English; Check: IsAdmin()
Root: HKU; Subkey: .DEFAULT\Software\Microsoft\Windows\ShellNoRoam\MUICache; ValueType: string; Valuename: {app}\PDFCreator.exe; ValueData: Erzeuge PDF and Bilddateien mit {#Appname}; Tasks: fileassoc; Languages: German; Check: IsAdmin()
Root: HKCU; Subkey: Software\Microsoft\Windows\ShellNoRoam\MUICache; ValueType: string; Valuename: {app}\PDFCreator.exe; ValueData: Create PDF and Bitmap Files with {#Appname}; Tasks: fileassoc; Languages: English; Check: IsAdmin()
Root: HKCU; Subkey: Software\Microsoft\Windows\ShellNoRoam\MUICache; ValueType: string; Valuename: {app}\PDFCreator.exe; ValueData: Erzeuge PDF and Bilddateien mit {#Appname}; Tasks: fileassoc; Languages: German; Check: IsAdmin()

;Uninstall - Software
Root: HKLM; Subkey: {#UninstallRegStr}; Flags: uninsdeletekey
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayName; Valuedata: {#AppName} {#AppVersionStr}}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: ApplicationVersion; Valuedata: {#AppVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: BetaVersion; Valuedata: {#BetaVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: PDFCreatorVersion; Valuedata: {#PDFCreatorVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: PDFSpoolerVersion; Valuedata: {#PDFSpoolerVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: TranstoolVersion; Valuedata: {#TranstoolVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptVersion; Valuedata: {#GhostscriptVersion}; Flags: uninsdeletevalue; Tasks: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: HelpLink; Valuedata: {#Homepage}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: UninstallString; Valuedata: {app}\unins000.exe; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: UnInstVersion; Valuedata: {#UnInstVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Publisher; Valuedata: Frank Heindörfer, Philip Chinery; Flags: uninsdeletevalue

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
Filename: {app}\PDFCreator.exe; Parameters: -NSTRUE; WorkingDir: {app}; Description: Install printerdriver; StatusMsg: Install PDFCreator printer; Flags: runminimized; Components: printer; Check: InstallCompletePrinter(); Languages: English
Filename: {app}\PDFCreator.exe; Parameters: -NSTRUE; WorkingDir: {app}; Description: Installiere Druckertreiber; StatusMsg: Installiere PDFCreator Drucker; Flags: runminimized; Components: printer; Check: InstallCompletePrinter(); Languages: German
Filename: {app}\ShFolder.Exe; Parameters: /Q:A; WorkingDir: {app}; Flags: runminimized; Components: program; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 4.1.2222,5.0.2195
#ENDIF

[UninstallDelete]
Name: {app}; Type: filesandordirs
Name: {%tmp}\{#Appname}; Type: filesandordirs

[UninstallRun]
Filename: {app}\PDFCreator.exe; WorkingDir: {app}; Parameters: -IPFALSE -ULTRUE -NSTRUE; Flags: runminimized
Filename: {app}\UnInst.exe; WorkingDir: {app}; Parameters: -UITRUE; Components: program; Check: IsFullInstallation()

[Languages]
Name: English; MessagesFile: compiler:Default.isl
Name: German; MessagesFile: German-2-4.0.0.isl

[Types]
Name: full; Description: Full installation; Check: CanPrinterInstall(); Languages: English
Name: full; Description: Komplette Installation; Check: CanPrinterInstall(); Languages: German

Name: compact; Description: Compact installation; Check: CanPrinterInstall(); Languages: English
Name: compact; Description: Minimale Installation; Check: CanPrinterInstall(); Languages: German

Name: custom; Description: Custom installation; Languages: English; Flags: iscustom
Name: custom; Description: Benutzerdefinierte Installation; Languages: German; Flags: iscustom

[Components]
Name: program; Description: Program Files; Types: full compact custom; Flags: fixed; Languages: English
Name: program; Description: Programm Dateien; Types: full compact custom; Flags: fixed; Languages: German

Name: printer; Description: Printer; Types: full custom; Check: CanPrinterInstall(); Flags: restart; Languages: English
Name: printer; Description: Drucker; Types: full custom; Check: CanPrinterInstall(); Flags: restart; Languages: German

[Tasks]
Name: desktopicon; Description: Create a &desktop icon; GroupDescription: Additional icons:; Languages: English
Name: desktopicon\common; Description: For all users; GroupDescription: Additional icons:; Flags: exclusive; Languages: English
Name: desktopicon\user; Description: For the current user only; GroupDescription: Additional icons:; Flags: exclusive unchecked; Languages: English
Name: quicklaunchicon; Description: Create a &Quick Launch icon; GroupDescription: Additional icons:; Flags: unchecked; Languages: English
Name: ghostscript; Description: Install &Ghostscript Version {#GhostscriptVersion}; GroupDescription: Other tasks:; Flags: exclusive; Languages: English
Name: fileassoc; Description: &Associate PDFCreator with the .ps file extension; GroupDescription: Other tasks:; Flags: unchecked; Languages: English

Name: desktopicon; Description: &Desktopsymbol anlegen; GroupDescription: Zusätzliche Symbole:; Languages: German
Name: desktopicon\common; Description: Für &alle Benutzer; GroupDescription: Zusätzliche Symbole:; Flags: exclusive; Languages: German
Name: desktopicon\user; Description: Nur für den angemeldeten &Benutzer; GroupDescription: Zusätzliche Symbole:; Flags: exclusive unchecked; Languages: German
Name: quicklaunchicon; Description: Erzeuge eine Symbol in der &Schnellzugriffsleiste; GroupDescription: Zusätzliche Symbole:; Flags: unchecked; Check: IExplorerVersionGreater3; Languages: German
Name: ghostscript; Description: Installiere &Ghostscript Version {#GhostscriptVersion}; GroupDescription: Andere Aufgaben:; Flags: exclusive; Languages: German
Name: fileassoc; Description: &Verknüpfe PDFCreator mit der Dateierweiterung .ps; GroupDescription: Andere Aufgaben:; Flags: unchecked; Languages: German

[Code]
var msg :Array of String; FullInstallation : boolean;

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

function AddMonitor (pName:PChar; Level:LongInt; var pMonitors:TMonitorInfo2): LongInt;
 external 'AddMonitorA@winspool.drv stdcall';
function AddPort (pName:PChar; hwnd:LongInt; pPort:PChar): LongInt;
 external 'AddPortA@winspool.drv stdcall';
function AddPrinterDriver (pName : PChar; Level : LongInt; var pDriverInfo : TDriverInfo3) : LongInt;
 external 'AddPrinterDriverA@winspool.drv stdcall';
function ClosePrinter(pPrinter: LongInt): Boolean;
 external 'ClosePrinter@winspool.drv stdcall';
function AddPrinter(pName : PChar; Level: Longint; var pPrinter2: TPrinterInfo2): LongInt;
 external 'AddPrinterA@winspool.drv stdcall';
function GetLastError() : LongInt;
 external 'GetLastError@kernel32.dll stdcall';
function GetPrinterDriverDirectory(pName:PChar; pEnvironment:PChar; Level:LongInt; pDriverDirectory:PChar; cbBuf:LongInt; var pcbNeened:LongInt):Integer;
 external 'GetPrinterDriverDirectoryA@winspool.drv stdcall';


var progTitel, progHandle: TArrayOfString;

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
 Result:=WizardSetupType(false);
end;

function GetActiveLanguage(Default:String):String;
begin
 Result:=ActiveLanguage();
end;

function IsAdmin(Default:String):Boolean;
begin
 Result:=IsAdminLoggedOn();
end;

function GetIExplorerVersion(): String;
var
 sVersion:  String;
begin
 RegQueryStringValue( HKLM, 'SOFTWARE\Microsoft\Internet Explorer', 'Version', sVersion );
 Result := sVersion;
end;

procedure DecodeVersion( verstr: String; var verint: array of Integer );
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
var vers: array of integer;
begin
 DecodeVersion(GetIExplorerVersion,vers);
 if vers[0]<4 then
   Result:=false
  else
   Result:=true;
end;

function PrinterDriverDirectory(Default:String):String;
var sb: LongInt;
	PrDrvDir : String;
	res: Integer;
begin
 res:=GetPrinterDriverDirectory(null, null, 1, null, 0, sb) ;
 PrDrvDir := StringOfChar(' ', sb+1 );
 res:=GetPrinterDriverDirectory( null, null, 1, PrDrvDir, sb, sb) ;
 if res=0 then begin
   PrDrvDir:= Default;
   If Default='Log' then
    SaveStringToFile(ExpandConstant('{app}')+'\SetupLog.txt', 'Printerdriver-Directory: Error '+IntToStr(GetLastError())+' = '+SysErrorMessage(GetLastError())+#13#10, True);
  end else begin
   PrDrvDir:= CastIntegerToString(CastStringToInteger(PrDrvDir));
   If Default='Log' then
    SaveStringToFile(ExpandConstant('{app}')+'\SetupLog.txt', 'Printerdriver-Directory: Success = '+PrDrvDir+#13#10, True);
  end;
 Result:=PrDrvDir;
end;

function ProgramIsInstalled(): Boolean;
begin
 if RegKeyExists(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}')=true then
   Result:=true
  else
   Result:=false;
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
   DI3.pDriverPath := 'ADOBEPS5.DLL';
   DI3.pEnvironment:='Windows NT x86';
   DI3.pHelpFile :='ADOBEPSU.HLP';
   DI3.pDataFile :='PDFCREATOR.PPD';
  end else Begin
   DI3.cVersion:=0;
   DI3.pConfigFile :='ADOBEPS4.DRV';
   DI3.pDriverPath := 'ADOBEPS4.DRV';
   DI3.pEnvironment:='Windows 4.0';
   DI3.pHelpFile :='ADOBEPS4.HLP';
   DI3.pDataFile :='DEFPRTR2.PPD';
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
 P2.pSharename:='';
 P2.Priority:=1;
 P2.DefaultPriority:=1;
 P2.Attributes :=0;
 P2.pDatatype:='RAW';

 res := AddPrinter(CastIntegerToString(0), 2, P2 );

 if res<>0 then begin
   ClosePrinter(res);
   SaveStringToFile(ExpandConstant('{app}') + '\SetupLog.txt', 'InstallPrinter: Success' + #13#10, True)
  end else
   SaveStringToFile(ExpandConstant('{app}') + '\SetupLog.txt', 'InstallPrinter: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10, True);
end;

function CanPrinterInstall(Default: string): boolean;
begin
 If IsAdminLoggedOn=False then
   Result:=False
  else
   If ProgramIsInstalled=true then
     Result:=false
    else
     Result:=true;
end;

function InstallCompletePrinter(Default: string): boolean;
var s:String;
begin
 s:='windows';
#IFNDEF Test
 InstallMonitor;
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(s)); // Ini-Refresh !!! Importent for Win9x/Me
 InstallDriver;
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(s)); // Ini-Refresh !!! Importent for Win9x/Me
 InstallPrinter;
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(s)); // Ini-Refresh !!! Importent for Win9x/Me
 PrinterDriverDirectory('Log');
#ENDIF
 Result:=True;
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

procedure InitMessages();
begin
 setArraylength(msg,11);
 If ActiveLanguage()='English' then begin
  msg[0]:=
  'You are not a member of the administrator group.'#10#13#13#13
  'For creating a pdf-file out of any application, the PDFCreator is using a postscript printerdriver.'#13
  'You have to be a member of the administrator group, to install any printerdriver.'#13#13
  'If you continue the installation, you will only be able to convert existing postscript-files.'#13#13
  'Continue the setup?';
  msg[1]:=
  'There was found an old version ('+GetInstalledVersionBeta+'). It is possible to update this version to version'+ExpandConstant('{#AppVersionStr}')+').'#10#13#10#13;
  msg[2]:=
  'for this you have NOT to be a member of the administrator group.'#10#13#10#13;
  msg[3]:=
  'To update, use <OK>, or cancel the setup and uninstall the older version first.';
  msg[4]:=
  'The program is installed already.'#10#13#10#13
  'For a new installation uninstall the program first.'#10#13#10#13
  'The setup will be cancelled.';
  msg[5]:=
  'The installed version (' + GetInstalledVersion + ') is newer than these setupversion ('+ExpandConstant('{#AppVersionStr}')+ ')!'#10#13#10#13
  'For the installation of an older version, uninstall the program first.'#10#13#10#13
  'The setup will be cancelled.';
  msg[6]:=
  'The program is installed already.'#10#13#10#13
  'An update is not possible! Please uninstall the program first.'#10#13#10#13
  'The setup will be cancelled.';
  msg[7]:=
  'The program <PDFCreator.exe> is running.'#10#13#10#13
  'Please close the program first.'#10#13#10#13;
  msg[8]:=
  'The program <Transtool.exe> is running.'#10#13#10#13
  'Please close the program first.'#10#13#10#13;
  msg[9]:=
  'The program <PDFSpooler.exe> is running.'#10#13#10#13
  'Please wait until all printjobs are finished or delete these printjobs.'#10#13#10#13;
  msg[10]:=
  'The program <UnInst.exe> is running.'#10#13#10#13
  'Please finish the uninstallation first.'#10#13#10#13
 end;
 If ActiveLanguage()='German' then begin
  msg[0]:=
  'Sie gehören nicht zur Gruppe der Administratoren.'#10#13#10#13
  'Um aus jeder Anwendung ein PDF-Dokument zu erzeugen, verwendet PDFCreator einen Postscript Druckertreiber.'#13
  'Für die Installation eines Druckertreibers müssen Sie zur Gruppe der Administratoren gehören.'#13#13
  'Wenn Sie mit der Installation fortfahren, können Sie nur bereits erstellte Postscript-Dateien konvertieren.'#13#13
  'Soll das Setup fortgesetzt werden?';
  msg[1]:=
  'Es wurde eine alte Version ('+GetInstalledVersionBeta+') gefunden! Eine Aktualisierung auf die Version ('+ExpandConstant('{#AppVersionStr}')+') ist möglich.'#10#13#10#13;
  msg[2]:=
  'Für dieses Update müssen Sie kein Mitglied einer Gruppe der Administratoren sein.'#10#13#10#13;
  msg[3]:=
  'Wenn ein Update durchgeführt werden soll, betätigen Sie die <OK> Taste,'#10#13
  'ansonsten brechen Sie die Installation ab und deinstallieren zuvor die alte Version.';
  msg[4]:=
  'Das Programm ist bereits installiert.'#10#13#10#13
  'Für eine Neuinstallation deinstallieren zuvor das Programm.'#10#13#10#13
  'Die Installation wird abgebrochen.';
  msg[5]:=
  'Die installierte Version ('+GetInstalledVersion+') ist aktueller, als diese Installationsversion ('+ExpandConstant('{#AppVersionStr}')+')!'#10#13#10#13
  'Für die Installation einer älteren Version, deinstallieren zuvor das Programm.'#10#13#10#13
  'Die Installation wird abgebrochen.';
  msg[6]:=
  'Das Programm wurde bereits installiert.'#10#13#10#13
  'Ein Update ist nicht möglich! Bitte deinstallieren Sie erst das Programm.'#10#13#10#13
  'Die Installation wird abgebrochen.';
  msg[7]:=
  'Das Programm <PDFCreator.exe> wird gerade verwendet.'#10#13#10#13
  'Bitte beenden Sie das Programm erst.'#10#13#10#13;
  msg[8]:=
  'Das Programm <Transtool.exe> wird gerade verwendet.'#10#13#10#13
  'Bitte beenden Sie das Programm erst.'#10#13#10#13;
  msg[9]:=
  'Das Programm <PDFSpooler.exe> wird gerade verwendet.'#10#13#10#13
  'Bitte warten Sie bis alle Druckaufträge abgearbeitet wurden.'#10#13#10#13;
  msg[10]:=
  'Das Programm <UnInst.exe> wird gerade verwendet.'#10#13#10#13
  'Bitte beenden Sie die Deinstallation erst.'#10#13#10#13;
 end;
end;

procedure DecodeVBVersion( verstr: String; var verint: array of Integer );
var
  i,p: Integer; s: string;
begin
  // initialize array
  verint := [0,0,0];
  i := 0;
  while ( (Length(verstr) > 0) and (i < 3) ) do
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
   InstBetaNumber:=StrToInt(InstBetaNumberStr);
  end else begin
   InstBetaNumber:=0;
 end;
 BetaNumber:=StrToInt(ExpandConstant('{#BetaVersion}'));
 If (InstBetaNumber=BetaNumber) then
   Result:=0 //equal
  else
   if ExpandConstant('{#BetaVersion}')='' then
     Result:=1 //major release
    else
     if (InstBetaNumber<BetaNumber) and (InstBetaNumber>=0) then
       Result:=2 //beta update
      else
       Result:=3; //no beta update possible
end;

// This function compares VB version string
// return -1 if ver1 < ver2
// return  0 if ver1 = ver2
// return  1 if ver1 > ver2
function CompareVBVersion(ver1, ver2: String ) : Integer;
var
  verint1, verint2: array of Integer; betaUpd:LongInt;
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

function IsFullInstallation(Default:String): Boolean;
begin
 result:=FullInstallation;
end;

function InitializeSetup(): Boolean;
var
 cv,a:Longint; tmsg:String; verySilent:boolean;
begin
 InitMessages;
 verySilent:=false;
 if CheckForMutexes(ExpandConstant('{#PDFCreatorExeIDStr}'))=true then begin
  Repeat
   a:=msgbox(msg[7],mbInformation, MB_OKCancel);
  until (a=IDCancel) or (CheckForMutexes(ExpandConstant('{#PDFCreatorExeIDStr}'))=false);
  if a=IDCancel then exit;
 end;
 if CheckForMutexes(ExpandConstant('{#TransToolExeIDStr}'))=true then begin
  Repeat
   a:=msgbox(msg[8],mbInformation, MB_OKCancel);
  until (a=IDCancel) or (CheckForMutexes(ExpandConstant('{#TransToolExeIDStr}'))=false);
  if a=IDCancel then exit;
 end;
 if CheckForMutexes(ExpandConstant('{#PDFSpoolerExeIDStr}'))=true then begin
  Repeat
   a:=msgbox(msg[9],mbInformation, MB_OKCancel);
  until (a=IDCancel) or (CheckForMutexes(ExpandConstant('{#PDFSpoolerExeIDStr}'))=false);
  if a=IDCancel then exit;
 end;
 if CheckForMutexes(ExpandConstant('{#UnInstExeIDStr}'))=true then begin
  Repeat
   a:=msgbox(msg[10],mbInformation, MB_OKCancel);
  until (a=IDCancel) or (CheckForMutexes(ExpandConstant('{#UnInstExeIDStr}'))=false);
  if a=IDCancel then exit;
 end;
 for a:=1 to Paramcount do begin
  if uppercase(paramstr(a))='/VERYSILENT' then
   verySilent:=true;
 end;
#ifdef UpdateIsPossible
 If ProgramIsInstalled=true then begin
   FullInstallation:=false;
   cv:=CompareVBVersion(GetInstalledVersion,ExpandConstant('{#AppVersion}'));
   if cv=-1 then begin
    cv:=CompareVBVersion(GetInstalledVersion,ExpandConstant('{#UpdateIsPossibleMinVersion}'));
    if cv=-1 then begin
      Result:=false;
      msgbox(msg[6],mbConfirmation, MB_OKCancel);
     end else begin
      Result:=true;
      if verySilent=false then begin
       tmsg:=msg[1];
       if UsingWinNt=true then
        tmsg:=tmsg+msg[2];
       tmsg:=tmsg+msg[3];
       a:=msgbox(tmsg,mbConfirmation, MB_OKCancel);
       if a=IDCancel then
         Result:=false
        else
         Result:=true;
      end
    end
    cv:=-1;
   end
   if cv=0 then begin
    msgbox(msg[4],mbInformation, MB_OK);
    Result:=false
   end
   if cv=1 then begin
    a:=msgbox(msg[5],mbInformation, MB_OK);
    Result:=false
   end
  end else begin
   FullInstallation:=true;
   If IsAdminLoggedOn=True then
     Result := True
    else begin
     Result:=true;
     if verySilent=false then begin
      a:=MsgBox(msg[0], mbConfirmation, MB_YesNo);
      If a=IDYES then
        Result:=True
       else
        Result:=False;
     end;
    end;
  end;
#Else
 Result:= true;
 FullInstallation:=true;
 If ProgramIsInstalled=true then begin
   if verySilent=false then
    msgbox(msg[6],mbInformation, MB_OK);
   Result:=false;
  end else
   If IsAdminLoggedOn=True then
     Result := True
    else begin
     Result:=true;
     if verySilent=false then begin
      a:=MsgBox(msg[0], mbConfirmation, MB_YesNo);
      If a=IDYES then
        Result:=True
       else
        Result:=False;
     end;
    end;
#endif
end;
