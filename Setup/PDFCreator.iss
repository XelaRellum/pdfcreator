; PDFCreator Installation
; Setup created with Inno Setup 5.0.4 Beta, ISPP 1.3.0.305 and ISTool 4.2.7
; Installation from Frank Heindörfer, Philip Chinery

;#define Test
#define CompileHelp
#define SetupLZMACompressionMode "ultra"
;#define SetupLZMACompressionMode "fast"

#define IncludeGhostscript

#define ProgramLicense "GNU"
#define GhostscriptLicense "GPL"

#Ifdef IncludeGhostscript
#If (GhostscriptLicense=="GPL")
 #define GhostscriptVersion "8.15"
 #define GhostscriptSetupString "GPLGhostscript"
#ENDIF
#ENDIF

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

#define BetaVersion          ""

#IF (BetaVersion!="")
 #define AppVersionStr       AppVersion + " Beta " + BetaVersion
 #define SetupAppVersionStr  SetupAppVersion + "_" + "Beta_" + BetaVersion
#ELSE
 #define AppVersionStr       AppVersion
 #define SetupAppVersionStr  SetupAppVersion
#ENDIF

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
#define PrintReg             "System\CurrentControlSet\Control\Print\"
#define PrintRegMon          "System\CurrentControlSet\Control\Print\Monitors\"

;#define UpdateIsPossible
#define UpdateIsPossibleMinVersion "0.8.0"

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

#Ifdef CompileHelp
Name: C:\Program Files\HTML Help Workshop\HHC.EXE; Parameters: ..\Help\PDFCreator.hhp
#Endif

[Setup]
AllowNoIcons=false
AlwaysRestart=false
AppCopyright=© 2002 - 2004 Philip Chinery, Frank Heindörfer
AppID={#AppIDStr}
AppName={#AppName}
AppVerName={#AppName} {#AppVersionStr}
AppPublisher=Philip Chinery, Frank Heindörfer
AppPublisherURL={#Homepage}
AppSupportURL={#Homepage}
AppUpdatesURL={#Homepage}
ChangesAssociations=true
Compression=lzma/{#SetupLZMACompressionMode}
CreateUninstallRegKey=false
DefaultDirName={reg:HKLM\{#UninstallRegStr2},Inno Setup: App Path|{pf}\{#AppName}}
DefaultGroupName={#AppName}
DisableDirPage=false
DisableStartupPrompt=true
InternalCompressLevel={#SetupLZMACompressionMode}
LicenseFile=.\License\GNU Readme.rtf
#Ifdef IncludeGhostscript
OutputBaseFilename={#AppName}-{#SetupAppVersionStr}_{#GhostscriptSetupString}
#ELSE
OutputBaseFilename={#AppName}-{#SetupAppVersionStr}_WithoutGhostscript
#ENDIF
OutputDir=Installation
RestartIfNeededByRun=true
ShowTasksTreeLines=false
SolidCompression=true
UsePreviousAppDir=true

VersionInfoVersion=0.8.1
VersionInfoCompany=Frank Heindörfer, Philip Chinery
VersionInfoDescription=PDFCreator is the easy way of creating PDFs.
VersionInfoTextVersion=0.8.1

WizardImageFile=..\Pictures\PDFCreatorBig.bmp
WizardSmallImageFile=..\Pictures\PDFCreator.bmp

[InstallDelete]
#Ifdef GhostscriptVersion
Name: {app}\Gs{#GhostscriptVersion}\Fonts\*.*; Type: filesandordirs; Tasks: ghostscript
Name: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib\*.*; Type: filesandordirs; Tasks: ghostscript
Name: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin\gsdll32.dll; Type: files; Tasks: ghostscript
#ENDIF
Name: {app}\languages\*.ini; Type: files; Components: program

[Files]
#IFNDEF Test
;We sort all files by extension for a maximal compression
;Systemfiles
Source: ..\SystemFiles\ASYCFILT.DLL; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace uninsneveruninstall

;Please use newest MSVBVM60.DLL
;http://support.microsoft.com/default.aspx?scid=kb;en-us;823746
Source: ..\SystemFiles\MSVBVM60.DLL; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace regserver uninsneveruninstall

Source: ..\SystemFiles\MSMPIDE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile
Source: ..\SystemFiles\OLEPRO32.DLL; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace regserver uninsneveruninstall
Source: ..\SystemFiles\OLEAUT32.DLL; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace regserver uninsneveruninstall

;Systemfiles Language: German
;http://msdn.microsoft.com/vbasic/downloads/tools/ipdk.aspx
Source: C:\IPDK\German\CMDLGDE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile
Source: C:\IPDK\German\MSCC2DE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile
Source: C:\IPDK\German\MSCMCDE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile
Source: C:\IPDK\German\VB6DE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile

;Printer DLLs
; for Win9x/Me
Source: ..\Printer\Adobe\Windows\ICONLIB.DLL; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0
Source: ..\Printer\Adobe\Windows\PSMON.DLL; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0
; for WinNt
Source: ..\Printer\Adobe\WinNT\AdobePS5.dll; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.0.1381; OnlyBelowVersion: 0,5.0.2195; Flags: deleteafterinstall
Source: ..\Printer\Adobe\WinNT\AdobePSu.dll; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.0.1381; OnlyBelowVersion: 0,5.0.2195; Flags: deleteafterinstall
; for Win2000
Source: ..\Printer\Adobe\Win2000\PS5UI.DLL; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Flags: deleteafterinstall
Source: ..\Printer\Adobe\Win2000\PSCRIPT5.DLL; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Flags: deleteafterinstall
; for WinXP
Source: ..\Printer\Adobe\WinXP\PS5UI.DLL; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,5.01.2600; Flags: deleteafterinstall
Source: ..\Printer\Adobe\WinXP\PSCRIPT5.DLL; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,5.01.2600; Flags: deleteafterinstall

;Ghostscript
#IFDEF GhostscriptVersion
Source: C:\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin\gsdll32.dll; DestDir: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Components: program; Tasks: ghostscript; Flags: ignoreversion
#ENDIF

;Redmon files
Source: ..\Printer\Redmon\redmonnt.dll; DestDir: {sys}; Components: printer; MinVersion: 0,4.00.1381; DestName: pdfcmnnt.dll
Source: ..\Printer\Redmon\redmon95.dll; DestDir: {sys}; Components: printer; MinVersion: 4.00.950,0; DestName: pdfcmn95.dll

Source: ..\SystemFiles\MSCOMCT2.OCX; DestDir: {sys}; Components: program; Flags: sharedfile regserver
Source: ..\SystemFiles\MSCOMCTL.OCX; DestDir: {sys}; Components: program; Flags: sharedfile regserver
Source: ..\SystemFiles\MSMAPI32.OCX; DestDir: {sys}; Components: program; Flags: sharedfile regserver promptifolder

Source: ..\SystemFiles\STDOLE2.TLB; DestDir: {sys}; Components: program; Flags: sharedfile restartreplace uninsneveruninstall regtypelib

;Program files
Source: ..\PDFCreator\PDFCreator.exe; DestDir: {app}; Components: program; Flags: comparetimestamp
Source: ..\Transtool\TransTool.exe; DestDir: {app}\languages; Components: program; Flags: comparetimestamp
Source: ..\PDFSpooler\PDFSpooler.exe; DestDir: {sys}; Components: printer; Flags: comparetimestamp

;ShFolder for older systems
;http://www.microsoft.com/downloads/release.asp?releaseid=30340
Source: ShFolder\ShFolder.Exe; DestDir: {app}; Components: program; Flags: ignoreversion deleteafterinstall; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 4.1.2222,5.0.2195

;pdfenc
Source: pdfenc\pdfenc.exe; DestDir: {app}; Flags: ignoreversion

;Help file
Source: ..\Help\PDFCreator.chm; DestDir: {app}; Flags: ignoreversion

;#If (GhostscriptLicense=="AFPL")
Source: License\AFPL License.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp
;#ENDIF
Source: License\GNU License.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp
Source: History.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp

Source: ..\PDFCreator\Languages\*.ini; DestDir: {app}\languages; Components: program; Flags: ignoreversion comparetimestamp
Source: PDFCreator.ini; DestDir: {userappdata}\PDFCreator; Components: program; DestName: PDFCreator.ini; Flags: ignoreversion onlyifdoesntexist uninsneveruninstall


;Printer files
Source: ..\Printer\Adobe\PDFCREATOR.PPD; DestName: ADIST5.PPD; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Adobe\PDFCREATOR.PPD; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.00.1381; Flags: ignoreversion deleteafterinstall

;Printer HLPs
; for Win9x/Me
Source: ..\Printer\Adobe\Windows\ADOBEPS4.HLP; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
; for WinNt
Source: ..\Printer\Adobe\WinNT\ADOBEPSU.HLP; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.0.1381; OnlyBelowVersion: 0,5.0.2195; Flags: ignoreversion deleteafterinstall
; for Win2000
Source: ..\Printer\Adobe\Win2000\PSCRIPT.HLP; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Flags: ignoreversion deleteafterinstall
; for WinXP
Source: ..\Printer\Adobe\WinXP\PSCRIPT.HLP; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,5.01.2600; Flags: ignoreversion deleteafterinstall

;Printer others
; for Win9x/Me
Source: ..\Printer\Adobe\Windows\FONTSDIR.MFD; DestDir: {win}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Adobe\Windows\adfonts.mfm; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0; Flags: ignoreversion
Source: ..\Printer\Adobe\Windows\ADOBEPS4.DRV; DestDir: {code:PrinterDriverDirectory|{sys}}; Components: printer; MinVersion: 4.00.950,0

; for WinNt
Source: ..\Printer\Adobe\WinNT\AdobePS5.ntf; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,4.0.1381; OnlyBelowVersion: 0,5.0.2195; Flags: ignoreversion deleteafterinstall

; for Win2000/XP
Source: ..\Printer\Adobe\Win2000\PSCRIPT.NTF; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,0; Flags: ignoreversion deleteafterinstall
;Source: ..\Printer\Adobe\Win2000\PSCRPTFE.NTF; DestDir: {code:PrinterDriverDirectory|{sys}\spool\drivers\w32x86}; Components: printer; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,0; Flags: ignoreversion deleteafterinstall

;Ghostscript
#IFDEF GhostscriptVersion
Source: C:\Gs{#GhostscriptVersion}\Fonts\*.*; DestDir: {app}\Gs{#GhostscriptVersion}\Fonts; Components: program; Tasks: ghostscript; Flags: ignoreversion sortfilesbyextension
Source: C:\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib\*.*; DestDir: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Components: program; Tasks: ghostscript; Flags: ignoreversion sortfilesbyextension
#ENDIF
#ENDIF

[Icons]
Name: {group}\{#Appname}; Filename: {app}\{#AppExename}; IconIndex: 0; Flags: createonlyiffileexists
;#If (GhostscriptLicense=="AFPL")
Name: {group}\AFPL License; Filename: {app}\AFPL License.txt
;#ENDIF
;#If (GhostscriptLicense=="GNU")
Name: {group}\GPL License; Filename: {app}\GNU License.txt
;#ENDIF
Name: {group}\{cm:History}; Filename: {app}\History.txt; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\Translation Tool; Filename: {app}\languages\transtool.exe; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\{cm:UninstallProgram,{#Appname}}; Filename: {uninstallexe}; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\{cm:ProgramOnTheWeb,PDFCreator}; Filename: {app}\PDFCreator.url
Name: {group}\PDFCreator {cm:Help}; Filename: {app}\PDFCreator.chm

Name: {commondesktop}\PDFCreator; Filename: {app}\pdfcreator.exe; Tasks: desktopicon\common
Name: {userdesktop}\PDFCreator; Filename: {app}\pdfcreator.exe; Tasks: desktopicon\user
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\PDFCreator; Filename: {app}\pdfcreator.exe; Tasks: quicklaunchicon

[INI]
Filename: {app}\PDFCreator.url; Section: InternetShortcut; Key: URL; String: http://www.pdfcreator.de.vu/; Components: program
Filename: {app}\PDFCreator.url; Section: InternetShortcut; Key: Iconindex; String: 1; Components: program

Filename: {app}\{cm:Donation}.url; Section: InternetShortcut; Key: URL; String: https://www.paypal.com/xclick/business=paypal01%40heindoerfer.com&item_name=PDFCreator&no_note=1&tax=0&currency_code=EUR; Components: program
Filename: {app}\{cm:Donation}.url; Section: InternetShortcut; Key: Iconindex; String: 1; Components: program

Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: AutosaveDirectory; String: {userdocs}; Components: program
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: LastsaveDirectory; String: {userdocs}; Components: program
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: DirectoryJava; String: {sys}; Components: program
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: Language; String: {code:GetActiveLanguage}

#Ifdef GhostscriptVersion
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: DirectoryGhostscriptBinaries; String: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Components: program
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: DirectoryGhostscriptFonts; String: {app}\Gs{#GhostscriptVersion}\Fonts; Components: program
Filename: {userappdata}\PDFCreator\PDFCreator.ini; Section: Options; Key: DirectoryGhostscriptLibraries; String: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Components: program
#ENDIF

[Registry]
;PrinterMonitor
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}; Components: printer
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: string; Valuename: Arguments; ValueData: -PPDFCREATORPRINTER; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: string; Valuename: Command; ValueData: {code:Shortname|{sys}\{#SpoolerExename}}; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: dword; Valuename: Delay; ValueData: 300; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: string; Valuename: Description; ValueData: PDFCreator Redirected Port; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: dword; Valuename: LogFileDebug; ValueData: 0; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: dword; Valuename: LogFileUse; ValueData: 0; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: dword; Valuename: Output; ValueData: 0; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: string; Valuename: Printer; ValueData: {code:GetPrintername|PDFCreator}; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: dword; Valuename: Printerror; ValueData: 0; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: dword; Valuename: Runuser; ValueData: 0; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; ValueType: dword; Valuename: ShowWindow; ValueData: 0; Flags: uninsdeletevalue; Components: printer; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 0,0

;Uninstall - Deletekey
Root: HKLM; Subkey: {#PrintReg}Printers\{code:GetPrintername|PDFCreator}; Flags: uninsdeletekey dontcreatekey; Components: printer
Root: HKLM; Subkey: {#PrintReg}Environments\Windows 4.0\Drivers\{code:GetPrinterdrivername|PDFCreator}; Flags: uninsdeletekey dontcreatekey; MinVersion: 4.00.950,0; Components: printer
Root: HKLM; Subkey: {#PrintReg}Environments\Windows NT x86\Drivers\{code:GetPrinterdrivername|PDFCreator}; Flags: uninsdeletekey dontcreatekey; MinVersion: 0,4.00.1381; Components: printer
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}\Ports\{code:GetPrinterportname|PDFCreator:}; Flags: uninsdeletekey dontcreatekey; Components: printer
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname|PDFCreator}; Flags: uninsdeletekey dontcreatekey; Components: printer

;File-Assoc
Root: HKCR; SubKey: .ps; ValueType: string; ValueData: PostScript; Flags: uninsdeletekey; Tasks: fileassoc
Root: HKCR; SubKey: PostScript; ValueType: string; ValueData: PostScript; Flags: uninsdeletekey; Tasks: fileassoc
Root: HKCR; SubKey: PostScript\Shell\Open\Command; ValueType: string; ValueData: """{app}\PDFCreator.exe"" -IF""%1"""; Flags: uninsdeletevalue; Tasks: fileassoc
Root: HKCR; Subkey: PostScript\DefaultIcon; ValueType: string; ValueData: {app}\PDFCreator.exe,0; Flags: uninsdeletevalue; Tasks: fileassoc
Root: HKU; Subkey: .DEFAULT\Software\Microsoft\Windows\ShellNoRoam\MUICache; ValueType: string; Valuename: {app}\PDFCreator.exe; ValueData: {cm:WinexplorerEntryCreate,{#Appname}}; Tasks: fileassoc; Check: IsAdminLoggedOn()
Root: HKCU; Subkey: Software\Microsoft\Windows\ShellNoRoam\MUICache; ValueType: string; Valuename: {app}\PDFCreator.exe; ValueData: {cm:WinexplorerEntryCreate,{#Appname}}; Tasks: fileassoc; Check: IsAdminLoggedOn()

;Windows Explorer popup-menu
;Root: HKCR; SubKey: *\shell\{#UninstallIDStr}; ValueType: string; ValueData: Create &PDF with PDFCreator; Flags: uninsdeletekey; Tasks: winexplorer; Languages: English
;Root: HKCR; SubKey: *\shell\{#UninstallIDStr}; ValueType: string; ValueData: Erzeuge &PDF mit PDFCreator; Flags: uninsdeletekey; Tasks: winexplorer; Languages: German
;Root: HKCR; SubKey: *\shell\{#UninstallIDStr}\command; ValueType: string; ValueData: "{app}\pdfcreator.exe -PF""%1"" -NS"; Flags: uninsdeletekey; Tasks: winexplorer

;Uninstall - Software
Root: HKLM; Subkey: {#UninstallRegStr}; Flags: uninsdeletekey
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayName; Valuedata: {#AppName} {#AppVersionStr}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: ApplicationVersion; Valuedata: {#AppVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: BetaVersion; Valuedata: {#BetaVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: PDFCreatorVersion; Valuedata: {#PDFCreatorVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: PDFSpoolerVersion; Valuedata: {#PDFSpoolerVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: TranstoolVersion; Valuedata: {#TranstoolVersion}; Flags: uninsdeletevalue

#Ifdef GhostscriptVersion
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptCopyright; Valuedata: {#GhostscriptLicense}; Flags: uninsdeletevalue; Tasks: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptVersion; Valuedata: {#GhostscriptVersion}; Flags: uninsdeletevalue; Tasks: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptDirectoryBinaries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Flags: uninsdeletevalue; Tasks: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptDirectoryLibraries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Flags: uninsdeletevalue; Tasks: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptDirectoryFonts; Valuedata: {app}\Gs{#GhostscriptVersion}\Fonts; Flags: uninsdeletevalue; Tasks: ghostscript
#Endif

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: HelpLink; Valuedata: {#Homepage}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: UninstallString; Valuedata: {app}\unins000.exe; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Publisher; Valuedata: Frank Heindörfer, Philip Chinery; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Printername; Valuedata: {code:GetPrintername|PDFCreator}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Printerdrivername; Valuedata: {code:GetPrinterdrivername|PDFCreator}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Printerportname; Valuedata: {code:GetPrinterportname|PDFCreator:}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Printermonitorname; Valuedata: {code:GetPrintermonitorname|PDFCreator}; Flags: uninsdeletevalue

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: App Path; Valuedata: {app}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Components; Valuedata: {code:GetWizardSelectedComponents}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Tasks; Valuedata: {code:GetWizardSelectedTasks}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Silent; Valuedata: {code:GetWizardSilent}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Group; Valuedata: {code:GetWizardGroupValue}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: NoIcons; Valuedata: {code:GetWizardNoIcons}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: SetupType; Valuedata: {code:GetWizardSetupType}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: SetupLanguage; Valuedata: {code:GetActiveLanguage}; Flags: uninsdeletevalue

;CustomMessages for uninstall. InnoSetop 5.0.4 Beta doesn't support custom messages for uninstalling.
Root: HKLM; Subkey: {#UninstallRegStr}\CustomMessages; ValueType: string; ValueName: UninstallOptions; Valuedata: {cm:UninstallOptions}; Flags: uninsdeletevalue

[Run]
#IFNDEF Test
;Uninstall old printer
Filename: {app}\PDFCreator.exe; WorkingDir: {app}; Parameters: -IPFALSE -NSTRUE; Flags: runminimized
;Install new printer
Filename: {app}\PDFCreator.exe; WorkingDir: {app}; Parameters: -NSTRUE; Description: {cm:InstallPrinterdriver}; StatusMsg: {cm:InstallPrinter,{#AppName}}; Flags: runminimized; Components: printer; Check: InstallCompletePrinter()
Filename: {app}\ShFolder.Exe; WorkingDir: {app}; Parameters: /Q:A; Flags: runminimized; Components: program; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 4.1.2222,5.0.2195
#ENDIF

[UninstallDelete]
Name: {app}; Type: filesandordirs
;User temp directory
Name: {%tmp}\{#Appname}; Type: filesandordirs

[UninstallRun]
Filename: {app}\PDFCreator.exe; WorkingDir: {app}; Parameters: -IPFALSE -ULTRUE -NSTRUE; Flags: runminimized

[Languages]
Name: Czech; MessagesFile: compiler:Languages\Czech.isl
Name: English; MessagesFile: compiler:Default.isl
Name: German; MessagesFile: compiler:Languages\German.isl
Name: Italian; MessagesFile: Italian-3-4.2.1.isl

[CustomMessages]
#include "czech.inc"
#include "english.inc"
#include "german.inc"
#include "italian.inc"

[Types]
Name: full; Description: {cm:FullInstallation}; Check: CanPrinterInstall()
Name: compact; Description: {cm:CompactInstallation}; Check: CanPrinterInstall()
Name: custom; Description: {cm:CustomInstallation}; Flags: iscustom

[Components]
Name: program; Description: {cm:Programfiles}; Types: full compact custom; Flags: fixed
Name: printer; Description: {cm:Printer}; Types: full custom; Check: CanPrinterInstall(); Flags: restart

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}
Name: desktopicon\common; Description: {cm:ForAllUser}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive
Name: desktopicon\user; Description: {cm:ForTheCurrentUserOnly}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive unchecked
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked; Check: IExplorerVersionGreater3
#Ifdef GhostscriptVersion
Name: ghostscript; Description: {cm:InstallGhostscript,{#GhostscriptLicense},{#GhostscriptVersion}}; GroupDescription: {cm:OtherTasks}; Flags: exclusive
#ENDIF
Name: fileassoc; Description: {cm:AssocFileExtension,PDFCreator,.ps}; GroupDescription: {cm:OtherTasks}; Flags: unchecked
Name: winexplorer; Description: {cm:WinexplorerEntry}; GroupDescription: {cm:OtherTasks}

[Code]
type
 TAInt = Array of Integer; TAStr = Array of String;
 TMonitorInfo2 = record
  pName : String;
  pEnvironment : String;
  pDLLName : String;
 end;
 TDriverInfo3 = record
  cVersion : LongInt;
  pName : String;
  pEnvironment : String;
  pDriverPath : String;
  pDataFile : String;
  pConfigFile : String;
  pHelpFile : String;
  pDependentFiles : String;
  pMonitorName : String;
  pDefaultDataType : String;
 end;
 TPrinterInfo2 = record
  pServerName : String;
  pPrinterName : String;
  pShareName : String;
  pPortName : String;
  pDriverName : String;
  pComment : String;
  pLocation : String;
  pDevMode : String;
  pSepFile : String;
  pPrintProcessor : String;
  pDatatype : String;
  pParameters : String;
  pSecurityDescriptor : String;
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
function GetLastError() : LongInt;
 external 'GetLastError@kernel32.dll stdcall';
function lstrlenA (lpString : LongInt) : LongInt;
 external 'lstrlenA@kernel32.dll stdcall';
function lstrcpyA (lpString1 : String; lpString2 : LongInt) : LongInt;
 external 'lstrcpyA@kernel32.dll stdcall';

function AddMonitor (pName:String; Level:LongInt; var pMonitors:TMonitorInfo2): LongInt;
 external 'AddMonitorA@winspool.drv stdcall';
function AddPort (pName:String; hwnd:LongInt; pPort:String): LongInt;
 external 'AddPortA@winspool.drv stdcall';
function AddPrinterDriver (pName : String; Level : LongInt; var pDriverInfo : TDriverInfo3) : LongInt;
 external 'AddPrinterDriverA@winspool.drv stdcall';
function ClosePrinter(pPrinter: LongInt): Boolean;
 external 'ClosePrinter@winspool.drv stdcall';
function AddPrinter(pName : String; Level: Longint; var pPrinter2: TPrinterInfo2): LongInt;
 external 'AddPrinterA@winspool.drv stdcall';
function GetPrinterDriverDirectory(pName:String; pEnvironment:String; Level:LongInt; pDriverDirectory:String; cbBuf:LongInt; var pcbNeened:LongInt):Integer;
 external 'GetPrinterDriverDirectoryA@winspool.drv stdcall';

function EnumPorts(pName:String; Level:LongInt; lpbPorts:String;
 cbBuf:LongInt; var pcbNeeded:LongInt; var pcbReturned:LongInt):LongInt;
 external 'EnumPortsA@winspool.drv stdcall';
function EnumMonitors(pName:String; Level:LongInt; lpbMonitors:String;
 cbBuf:LongInt; var pcbNeeded:LongInt; var pcbReturned:LongInt):LongInt;
 external 'EnumMonitorsA@winspool.drv stdcall';
function EnumPrinterDrivers(pName:String; pEnvironment:String; Level:LongInt; lpbPrinterdrivers:String;
 cbBuf:LongInt; var pcbNeeded:LongInt; var pcbReturned:LongInt):LongInt;
 external 'EnumPrinterDriversA@winspool.drv stdcall';
function EnumPrinters(flags:LongInt; pName:String; Level:LongInt; lpbPrinters:String;
 cbBuf:LongInt; var pcbNeeded:LongInt; var pcbReturned:LongInt):LongInt;
 external 'EnumPrintersA@winspool.drv stdcall';


var progTitel, progHandle: TArrayOfString;
    msg : TAStr; FullInstallation : boolean;
    Printername, Printerdrivername, Printerportname, Printermonitorname, LogFile, PrintSystem : String;

function Shortname(Default:String):String;
begin
 Result:=GetShortname(Default);
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
 Result:=WizardSetupType(false);
end;

function GetActiveLanguage(Default:String):String;
begin
 Result:=ActiveLanguage();
end;

procedure IntegrateWinexplorer;
 var res: Boolean; keys: TArrayofString;i,c :LongInt;s1,s2,s3:String;
begin
 s3:=ExpandConstant('{cm:WinexplorerEntryCreate}');
 StringChange(s3,'%1',ExpandConstant('{#Appname}'));
 res:=RegGetSubkeyNames(HKEY_CLASSES_ROOT,'',keys);
 If res=true then begin
  c:=GetArrayLength(keys);
  If c>0 then begin
   For i:=0 to c-1 do begin
    If StrGet(keys[i],1)='.' then begin
     res:=RegQueryStringValue(HKEY_CLASSES_ROOT,keys[i],'',s1)
     if Length(s1)>0 then begin
      If RegKeyExists(HKEY_CLASSES_ROOT,s1)=true then begin
       If RegKeyExists(HKEY_CLASSES_ROOT,s1+'\shell\print\command')=true then begin
        res:=RegQueryStringValue(HKEY_CLASSES_ROOT,s1+'\shell\print\command','',s2);
        If res=true then begin
         If Length(s2)>0 then begin
          If RegKeyExists(HKEY_CLASSES_ROOT,s1+'\shell\print\command')=true then begin
           If RegKeyExists(HKEY_CLASSES_ROOT,s1+'\shell\'+'{#UninstallID}')=false then begin
            RegWriteStringValue(HKEY_CLASSES_ROOT,s1+'\shell\'+'{#UninstallID}','',s3);
            RegWriteStringValue(HKEY_CLASSES_ROOT,s1+'\shell\'+'{#UninstallID}'+'\command','',ExpandConstant('{app}')+'\pdfcreator.exe -PF'#34#37+'1'+#34' -NS');
           end;
          end;
         end;
        end;
       end;
      end;
     end;
    end;
   end;
  end;
 end;
end;

function GetStrFromPtrA(lpszA : LongInt) : String;
var
 tStr : String;
begin
 tStr := StringOfChar('A',lstrlenA(lpszA));
 lstrcpyA(tStr,lpszA);
 result:=tStr;
end;

function GetLongFromString(LStr : String; StartPos : LongInt) : LongInt;
var
 cStr : String;
begin
 cStr:=Copy(LStr,StartPos,4);
 result:=Ord(StrGet(cStr,1))       + Ord(StrGet(cStr,2))*256+
         Ord(StrGet(cStr,3))*65536 + Ord(StrGet(cStr,4))*16777216;
end;

function GetPorts(var Ports:Array of String) : LongInt;
var
 res, cbBuf, pcbNeeded, pcbReturned,i : LongInt;
 tArr : Array of String;
 tStr : String;
begin
 Setarraylength(tArr,0); cbBuf:=0; pcbNeeded:=0; pcbReturned:=0;
 res:=EnumPorts(Chr(0), 1, tStr, cbBuf, pcbNeeded, pcbReturned)
 if pcbNeeded>0 then begin
  cbBuf:=pcbNeeded; tStr:=StringOfChar(#0,pcbNeeded);
  res:=EnumPorts(Chr(0), 1, tStr, cbBuf, pcbNeeded, pcbReturned);
  Setarraylength(tArr,pcbReturned);
  For i:=0 To pcbReturned-1 do begin
   tArr[i]:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*4));
  end;
 end;
 Ports:=tArr;
 result:=GetArrayLength(tArr);
end;

function GetMonitors(var Monitors:Array of String) : LongInt;
var
 res, cbBuf, pcbNeeded, pcbReturned,i : LongInt;
 tArr : Array of String;
 tStr : String;
begin
 Setarraylength(tArr,0); cbBuf:=0; pcbNeeded:=0; pcbReturned:=0;
 res:=EnumMonitors(Chr(0), 1, tStr, cbBuf, pcbNeeded, pcbReturned)
 if pcbNeeded>0 then begin
  cbBuf:=pcbNeeded; tStr:=StringOfChar(#0,pcbNeeded);
  res:=EnumMonitors(Chr(0), 1, tStr, cbBuf, pcbNeeded, pcbReturned);
  Setarraylength(tArr,pcbReturned);
  For i:=0 To pcbReturned-1 do begin
   tArr[i]:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*4));
  end;
 end;
 Monitors:=tArr;
 result:=GetArrayLength(tArr);
end;

function GetPrinterdrivers(var Drivers : Array of String) : LongInt;
var
 res, cbBuf, pcbNeeded, pcbReturned,i : LongInt;
 tArr : Array of String;
 tStr : String;
begin
 Setarraylength(tArr,0); cbBuf:=0; pcbNeeded:=0; pcbReturned:=0;
 res:=EnumPrinterdrivers(Chr(0), Chr(0), 1, tStr, cbBuf, pcbNeeded, pcbReturned)
 if pcbNeeded>0 then begin
  cbBuf:=pcbNeeded; tStr:=StringOfChar(#0,pcbNeeded);
  res:=EnumPrinterdrivers(Chr(0), Chr(0), 1, tStr, cbBuf, pcbNeeded, pcbReturned);
  Setarraylength(tArr,pcbReturned);
  For i:=0 To pcbReturned-1 do begin
   tArr[i]:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*4));
  end;
 end;
 Drivers:=tArr;
 result:=GetArrayLength(tArr);
end;

function GetPrinters(var Printers : Array of String) : LongInt;
var
 res, cbBuf, pcbNeeded, pcbReturned,i,sizeofPI, offs : LongInt;
 tArr : Array of String;
 tStr : String;
begin
 Setarraylength(tArr,0); cbBuf:=0; pcbNeeded:=0; pcbReturned:=0;
 if UsingWinnt=true then
   res:=EnumPrinters(2, Chr(0), 4, tStr, cbBuf, pcbNeeded, pcbReturned)
  else
   res:=EnumPrinters(2, Chr(0), 1, tStr, cbBuf, pcbNeeded, pcbReturned);
 if pcbNeeded>0 then begin
  cbBuf:=pcbNeeded; tStr:=StringOfChar(#0,pcbNeeded);
  if UsingWinNt=true then begin
    sizeofPI:=12; offs:=0;
    res:=EnumPrinters(2, Chr(0), 4, tStr, cbBuf, pcbNeeded, pcbReturned);
   end else begin
    sizeofPI:=16; offs:=8;
    res:=EnumPrinters(2, Chr(0), 1, tStr, cbBuf, pcbNeeded, pcbReturned);
   end;
  Setarraylength(tArr,pcbReturned);
  For i:=0 To pcbReturned-1 do begin
   tArr[i]:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*sizeofPI+offs));
  end;
 end;
 Printers:=tArr;
 result:=GetArrayLength(tArr);
end;

function GetPrintername(Default : String): String;
var tStr:String;
begin
 tStr:=Trim(Printername);
 if Length(tStr)=0 then begin
  tStr:=Trim(Default);
  if Length(tStr)=0 then
   tStr:='PDFCreator';
 end;
 result:=tStr;
end;

function GetPrinterdrivername(Default : String): String;
var tStr:String;
begin
 tStr:=Trim(Printerdrivername);
 if Length(tStr)=0 then begin
  tStr:=Trim(Default);
  if Length(tStr)=0 then
   tStr:='PDFCreator';
 end;
 result:=tStr;
end;

function GetPrinterportname(Default : String): String;
var tStr:String;
begin
 tStr:=Trim(Printerportname);
 if Length(tStr)=0 then begin
  tStr:=Trim(Default);
  if Length(tStr)=0 then
   tStr:='PDFCreator';
 end;
 result:=tStr;
end;

function GetPrintermonitorname(Default : String): String;
var tStr:String;
begin
 tStr:=Trim(Printermonitorname);
 if Length(tStr)=0 then begin
  tStr:=Trim(Default);
  if Length(tStr)=0 then
   tStr:='PDFCreator';
 end;
 result:=tStr;
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

function PrinterDriverDirectory(Default:String):String;
var sb: LongInt;
	PrDrvDir : String;
	res: Integer;
begin
 res:=GetPrinterDriverDirectory(chr(0),chr(0), 1,chr(0), 0, sb);
 PrDrvDir := StringOfChar(' ', sb+1 );
 If Default='Log' then begin
  SaveStringToFile(Logfile, 'Printerdriver-Directory:'+#13#10, True)
 end;
 res:=GetPrinterDriverDirectory(chr(0),chr(0), 1, PrDrvDir, sb, sb) ;
 if res=0 then begin
   PrDrvDir:= Default;
   If Default='Log' then
    SaveStringToFile(LogFile, ' Result: Error '+IntToStr(GetLastError())+' = '+SysErrorMessage(GetLastError())+#13#10#13#10, True);
  end else begin
   PrDrvDir:= CastIntegerToString(CastStringToInteger(PrDrvDir));
   If Default='Log' then
    SaveStringToFile(LogFile, ' Result: Success = '+PrDrvDir+#13#10#13#10, True);
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

procedure SaveInstallInformations;
begin
 SaveStringToFile(LogFile, '--------------------------------------'#13#10#13#10, True);
 SaveStringToFile(LogFile, 'Windowsversion: '+GetWindowsVersionString+#13#10, True);
 SaveStringToFile(LogFile, 'WinDir: '+GetWinDir+#13#10, True);
 SaveStringToFile(LogFile, 'SystemDir: '+GetSystemDir+#13#10, True);
 SaveStringToFile(LogFile, 'TempDir: '+GetTempDir+#13#10, True);
 SaveStringToFile(LogFile, 'CurrentDir: '+GetCurrentDir+#13#10, True);
 SaveStringToFile(LogFile, 'Computername: '+GetComputernameString+#13#10, True);
 SaveStringToFile(LogFile, 'Username: '+GetUsernameString+#13#10, True);
 SaveStringToFile(LogFile, 'UILanguage: '+IntToStr(GetUILanguage)+#13#10, True);
 SaveStringToFile(LogFile, 'Internet Explorer version: '+GetIExplorerVersion+#13#10, True);
 SaveStringToFile(LogFile, 'Path: '+Getenv('Path')+#13#10, True);
end;

procedure SavePrinterInformations;
var
 i,c:Longint; pi:Array of String;
begin
 c:=GetPorts(pi);
 SaveStringToFile(LogFile, 'Printerports ['+IntToStr(c)+']:'#13#10, True);
 for i:=1 to c do
  SaveStringToFile(LogFile,' '+pi[i-1]+#13#10, True);
 SaveStringToFile(LogFile, #13#10, True);

 c:=GetMonitors(pi);
 SaveStringToFile(LogFile, 'Printermonitors ['+IntToStr(c)+']:'#13#10, True);
 for i:=1 to c do
  SaveStringToFile(LogFile,' '+pi[i-1]+#13#10, True);
 SaveStringToFile(LogFile, #13#10, True);

 c:=GetPrinterdrivers(pi);
 SaveStringToFile(LogFile, 'Printerdrivers ['+IntToStr(c)+']:'#13#10, True);
 for i:=1 to c do
  SaveStringToFile(LogFile,' '+pi[i-1]+#13#10, True);
 SaveStringToFile(LogFile, #13#10, True);

 c:=GetPrinters(pi);
 SaveStringToFile(LogFile, 'Printers ['+IntToStr(c)+']:'#13#10, True);
 for i:=1 to c do
  SaveStringToFile(LogFile,' '+pi[i-1]+#13#10, True);
 SaveStringToFile(LogFile, #13#10, True);
end;

procedure InstallMonitor;
var M2:TMonitorInfo2; res:LongInt;
begin
 M2.pName:=GetPrintermonitorname('PDFCreator');
 If UsingWinNT=True then Begin
   M2.pEnvironment:='Windows NT x86';
   M2.pDLLName:='pdfcmnnt.dll'
  end else Begin
   M2.pEnvironment:='Windows 4.0';
   M2.pDLLName:='pdfcmn95.dll'
 end;

 SaveStringToFile(LogFile, 'InstallMonitor:' + #13#10, True)
 SaveStringToFile(LogFile, ' Monitorname : ' + M2.pName  + #13#10, True)

 res := AddMonitor(Chr(0), 2, M2);
 if res=0 then
   SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
  else
   SaveStringToFile(LogFile, ' Result: Success' + #13#10#13#10, True);
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(PrintSystem)); // Ini-Refresh !!! Important for Win9x/Me
end;

procedure InstallPort;
var res:Boolean; SubKeyName : String;
begin
 SubKeyName:=ExpandConstant('{#PrintRegMon}')+GetPrintermonitorname('PDFCreator');
 SubKeyName:=SubKeyName+'\Ports\'+GetPrinterportname('PDFCreator:');

 res:=RegWriteStringValue(HKLM,SubKeyName,'Arguments','-PPDFCREATORPRINTER');
 res:=RegWriteStringValue(HKLM,SubKeyName,'Command',Shortname(ExpandConstant('{sys}\{#SpoolerExename}')));
 res:=RegWriteDWordValue(HKLM,SubKeyName,'Delay',300);
 res:=RegWriteStringValue(HKLM,SubKeyName,'Description','PDFCreator Redirected Port');
 res:=RegWriteDWordValue(HKLM,SubKeyName,'LogFileDebug',0);
 res:=RegWriteDWordValue(HKLM,SubKeyName,'LogFileUse',0);
 res:=RegWriteDWordValue(HKLM,SubKeyName,'Output',0);
 res:=RegWriteStringValue(HKLM,SubKeyName,'Printer',GetPrintername('PDFCreator'));
 res:=RegWriteDWordValue(HKLM,SubKeyName,'Printerror',0);
 res:=RegWriteDWordValue(HKLM,SubKeyName,'Runuser',0);
 res:=RegWriteDWordValue(HKLM,SubKeyName,'ShowWindow',0);
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(PrintSystem)); // Ini-Refresh !!! Important for Win9x/Me
end;

procedure InstallDriver;
var DI3:TDriverInfo3; res:LongInt; PrDrDir:String;
begin
 PrDrDir:=PrinterDriverDirectory(ExpandConstant('{sys}') + '\spool\drivers\w32x86') + '\';
 DI3.pName :=GetPrinterdrivername('PDFCreator');
 DI3.pDependentFiles :='';
// Win9x
 If InstallOnThisVersion('4.00.950,0','0,0')=irInstall then begin
  DI3.cVersion:=0;
  DI3.pDependentFiles :=PrDrDir + 'ADOBEPS4.HLP'#0 + PrDrDir + 'ICONLIB.DLL'#0 + PrDrDir + 'PSMON.DLL'#0 + PrDrDir + 'ADFONTS.MFM'#0 + PrDrDir + 'ADOBEPS4.HLP'#0 + PrDrDir + 'ADOBEPS4.DRV'#0 + PrDrDir + 'ADIST5.PPD'#0#0;
  DI3.pConfigFile :='ADOBEPS4.DRV';
  DI3.pDriverPath := 'ADOBEPS4.DRV';
  DI3.pEnvironment:='Windows 4.0';
  DI3.pHelpFile :='ADOBEPS4.HLP';
  DI3.pDataFile :='ADIST5.PPD';
  DI3.cVersion := 3474436;
 end;
// WinNt
 If InstallOnThisVersion('0,4.0.1381','0,5.0.2195')=irInstall then begin
  DI3.cVersion:=2;
  DI3.pDependentFiles :=PrDrDir + 'PDFCREATOR.PPD'#0 + PrDrDir + 'ADOBEPS5.DLL'#0  + PrDrDir + 'ADOBEPSU.DLL'#0  + PrDrDir + 'ADOBEPS5.NTF'#0  + PrDrDir + 'ADOBEPSU.HLP'#0#0;
  DI3.pConfigFile :='ADOBEPSU.DLL';
  DI3.pDriverPath := 'ADOBEPS5.DLL';
  DI3.pEnvironment:='Windows NT x86';
  DI3.pHelpFile :='ADOBEPSU.HLP';
  DI3.pDataFile :='PDFCREATOR.PPD';
 end;
// Win2000/XP
 If InstallOnThisVersion('0,5.0.2195','0,0')=irInstall then begin
  DI3.cVersion:=3;
  DI3.pDependentFiles :=PrDrDir + 'PSCRIPT.NTF'#0#0;
  DI3.pConfigFile :='PS5UI.DLL';
  DI3.pDriverPath := 'PSCRIPT5.DLL';
  DI3.pEnvironment:='Windows NT x86';
  DI3.pHelpFile :='PSCRIPT.HLP';
  DI3.pDataFile :='PDFCREATOR.PPD';
 end;

 DI3.pDefaultDataType :='RAW';
 DI3.pMonitorName :='';

 SaveStringToFile(LogFile, 'InstallDriver:' + #13#10, True)
 SaveStringToFile(LogFile, ' Drivername : ' + DI3.pName  + #13#10, True)

 res := AddPrinterDriver(Chr(0), 3, DI3);

 if res=0 then
   SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
  else
   SaveStringToFile(LogFile, ' Result: Success' + #13#10#13#10, True);
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(PrintSystem)); // Ini-Refresh !!! Important for Win9x/Me
end;

procedure InstallPrinter;
var
 P2: TPrinterInfo2; res: LongInt; Printers : Array of String; cPrinters:LongInt;
begin
 P2.pPrinterName := GetPrintername('PDFCreator');
 P2.pDriverName := GetPrinterdrivername('PDFCreator');
 P2.pPrintProcessor := 'WinPrint';
 P2.pPortName := GetPrinterportname('PDFCreator:');
 P2.pComment := 'eDoc Printer';
 P2.pSharename:='';
 P2.Priority:=1;
 P2.DefaultPriority:=1;
 P2.pDatatype:='RAW';
 P2.Attributes :=0;

 cPrinters:=GetPrinters(Printers);
 If cPrinters=0 then
   P2.Attributes :=4 // Set as defaultprinter
  else
   P2.Attributes :=0;

 SaveStringToFile(LogFile, 'InstallPrinter:' + #13#10, True)
 SaveStringToFile(LogFile, ' Printername: ' + P2.pPrintername + #13#10, True)
 SaveStringToFile(LogFile, ' Drivername : ' + P2.pDrivername  + #13#10, True)
 SaveStringToFile(LogFile, ' Portname   : ' + P2.pPortname    + #13#10, True)

 res := AddPrinter(CastIntegerToString(0), 2, P2 );

 if res<>0 then begin
   ClosePrinter(res);
   SaveStringToFile(LogFile, ' Result: Success' + #13#10, True)
   if cPrinters=0 then begin
    // Set as defaultprinter
    SetIniString('windows','device',GetPrintername('PDFCreator')+',PSCRIPT,'+ GetPrinterMonitorname('PDFCreator'),ExpandConstant('{win}')+'\win.ini')
   end
  end else
   SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10, True);
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(PrintSystem)); // Ini-Refresh !!! Important for Win9x/Me
end;

function CanPrinterInstall(): boolean;
begin
 If IsAdminLoggedOn=False then
   Result:=False
  else
   If ProgramIsInstalled=true then
     Result:=false
    else
     Result:=true;
end;

function InstallCompletePrinter(): boolean;
var s : String; Ports, Monitors, Drivers, Printers : Array of String;
begin
 PrintSystem:='windows';
#IFNDEF Test
 SaveStringToFile(LogFile, 'Printerstatus before installing:' + #13#10, True)
 SavePrinterInformations
 PrinterDriverDirectory('Log');
 GetPorts(Ports);

 InstallMonitor;

 GetMonitors(Monitors);

 InstallPort;

 GetMonitors(Monitors);
 GetPorts(Ports);
 InstallDriver;

 GetPrinterdrivers(Drivers);

 InstallPrinter;

 If UsingWinNT=true then begin
  s:='SYSTEM\CurrentControlSet\Control\Print\Printers\'+GetPrintername('PDFCreator')+'\PrinterDriverData';
  If RegKeyExists(HKLM,s)=true then
    RegWriteDWordValue(HKLM,s,'FreeMem',32767);
 end;
 GetPrinters(Printers);

 SaveStringToFile(LogFile, #13#10+'Printerstatus after installing:' + #13#10, True)
 SavePrinterInformations

 s:=LowerCase(WizardSelectedTasks(false));
 if Pos('winexplorer',s)>0 then
  IntegrateWinexplorer;

#ENDIF
 Result:=True;
end;

function NextButtonClick(CurPage: Integer): Boolean;
begin
// MsgBox(IntToStr(CurPage),mbInformation,MB_OK)
 if CurPage=wpReady then begin
  GetActivePDFLoaders;
  KillActivePDFLoaders;
  LogFile:=ExpandConstant('{app}')+'\SetupLog.txt';
 end;
 if CurPage=wpFinished then
  SaveInstallInformations;
 Result := True;
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
var
 tmsg:String;
begin
 setArraylength(msg,11);
 Msg[0]:=ExpandConstant('{cm:NoAdmin}');

 tmsg:=ExpandConstant('{cm:OldVersion}');
 StringChange(tmsg,'%1',GetInstalledVersion);
 StringChange(tmsg,'%2',ExpandConstant('{#AppVersionStr}'));
 Msg[1]:=tmsg;

 Msg[2]:=ExpandConstant('{cm:NoNoAdmin}');
 Msg[3]:=ExpandConstant('{cm:Update}');
 Msg[4]:=ExpandConstant('{cm:AlreadyInstalled}');

 tmsg:=ExpandConstant('{cm:NewerVersion}');
 StringChange(tmsg,'%1',GetInstalledVersion);
 StringChange(tmsg,'%2',ExpandConstant('{#AppVersionStr}'));
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
 StringChange(tmsg,'%2',ExpandConstant('{#AppVersionStr}'));
 Msg[10]:=tmsg;
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

function IsFullInstallation(): Boolean;
begin
 result:=FullInstallation;
end;

function InitializeSetup(): Boolean;
var
#ifdef UpdateIsPossible
 cv,a:Longint;  verySilent:boolean;
#else
 a:Longint; verySilent:boolean;
#endif
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
   if verySilent=false then begin
    msgbox(msg[4],mbInformation, MB_OK);
   end;
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

#Ifdef IncludeGhostscript
procedure InitializeWizard();
begin
 WizardForm.TASKSLIST.Height := WizardForm.TASKSLIST.Height + 6;
end;
#endif

procedure RemoveProgramSettings();
var
 iniPath:String;
begin
 iniPath:=ExpandConstant('{userappdata}')+'\PDFCreator';
 DelTree(iniPath,true,true,true);
end;

procedure RemoveExplorerIntegretation();
var
 keys: TArrayOfString; i :LongInt;tStr:String;
begin
 if RegGetSubkeyNames(HKEY_CLASSES_ROOT, '', keys) then begin
  for i:=0 to GetArrayLength(keys)-1 do begin
   tStr:=keys[i]+'\shell\'+ExpandConstant('{#UninstallIDStr}');
   if RegKeyExists(HKEY_CLASSES_ROOT,tStr) then begin
    RegDeleteKeyIncludingSubkeys(HKEY_CLASSES_ROOT,tStr);
   end;
  end;
 end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
 tStr,engStr :String; i:LongInt; saveoptions, silent, verysilent:boolean;
begin
  case CurUninstallStep of
    usUninstall:
      begin
       tStr:=ExpandConstant('{app}')+'\Unload.tmp';
       if fileexists(tStr)=false then
        SaveStringToFile(tStr, '', True);
       tStr:='';engStr:='Delete all program settings?';
       if RegQueryStringValue(HKEY_LOCAL_MACHINE,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#UninstallID}\CustomMessages', 'UninstallOptions', tStr)=false then
        tStr:=engStr;
       if length(tStr)=0 then
        tStr:=engStr;
       saveoptions:=false; silent:=false; verysilent:=false;
       for i:=1 to paramcount do begin
        if lowercase(ParamStr(i))='/saveoptions' then
         saveoptions:=true;
        if lowercase(ParamStr(i))='/silent' then
         silent:=true;
        if lowercase(ParamStr(i))='/verysilent' then
         verysilent:=true;
       end;
       if saveoptions=false then
        if (silent=false) and (verysilent=false) then
         if MsgBox(tStr, mbConfirmation, MB_YESNO) = IDYES then
          RemoveProgramSettings;
       RemoveExplorerIntegretation;
      end;
  end;
end;
