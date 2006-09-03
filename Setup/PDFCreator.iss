; PDFCreator Installation
; Setup created with Inno Setup QuickStart Pack 5.1.7 (with ISPP) and ISTool 5.1.6
; Installation from Frank Heindörfer

;#define Test

#ifdef Test
 #define FastCompilation
 #define IncludeToolbar
#else
 #define FastCompilation
 #define CompileHelp
; #define IncludeGhostscript
; #define IncludeToolbar
; #define Localization
#endif

#define ProgramLicense "GNU"
#define GhostscriptLicense "GPL"

#ifdef FastCompilation
 #define CompressionMode="none"
 #define SetupLZMACompressionMode "none"
#else
 #define CompressionMode="lzma/ultra"
 #define SetupLZMACompressionMode "ultra"
#endif

#Ifdef IncludeGhostscript
 #define GhostscriptVersion "8.54"
 #define GhostscriptSetupString "GPLGhostscript"
#ENDIF

#if (fileexists("..\PDFCreator\PDFCreator.exe")==0)
 #error Compile PDFCreator first!
#endif
#if (fileexists("..\PDFSpooler\PDFSpooler.exe")==0)
 #error Compile PDFSpooler first!
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
  #expr Exec("C:\IPDK\VBLOCAL.EXE","..\PDFCreator\PDFCreator.exe * 0x409 ~ 0x0",".\")
  #expr Exec("C:\IPDK\VBLOCAL.EXE","..\PDFSpooler\PDFSpooler.exe * 0x409 ~ 0x0",".\")
  #expr Exec("C:\IPDK\VBLOCAL.EXE","..\TransTool\TransTool.exe * 0x409 ~ 0x0",".\")
 #endif
#endif

;add manifest to exe files
#IFNDEF Test
 #if (fileexists("..\ManifestManager\ManifestManager.exe")==0)
  #error Compile ManifestManager first!
 #endif
 #expr Exec("..\ManifestManager\ManifestManager.exe","/ADD""..\PDFCreator\PDFCreator.exe""","..\ManifestManager\")
 #expr Exec("..\ManifestManager\ManifestManager.exe","/ADD""..\PDFSpooler\PDFSpooler.exe""","..\ManifestManager\")
 #expr Exec("..\ManifestManager\ManifestManager.exe","/ADD""..\TransTool\TransTool.exe""","..\ManifestManager\")
#endif

#ifdef CompileHelp
 #if (fileexists("C:\Program Files\HTML Help Workshop\HHC.EXE")==0)
  #error Please install the "HTML Help Workshop" first!
 #endif
 #expr Exec("C:\Program Files\HTML Help Workshop\HHC.EXE", "..\Help\english\PDFCreator.hhp",".\")
 #expr Exec("C:\Program Files\HTML Help Workshop\HHC.EXE", "..\Help\german\PDFCreator.hhp" ,".\")
 #expr Exec("C:\Program Files\HTML Help Workshop\HHC.EXE", "..\Help\french\PDFCreator.hhp" ,".\")
#endif

#define GetFileVersionVBExe(str S)     Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + Local[3] + Local[4]
#define GetFileVersionVBExeLine(str S) Local[0]=GetFileVersion(S), Local[1]=Copy(Local[0],1,Pos(".",Local[0])-1), Local[2]=Copy(Local[0],Pos(".",Local[0])+1,Len(Local[0])-Pos(".",Local[0])), Local[3]=Copy(Local[2],1,Pos(".",Local[2])-1), Local[4]=Copy(Local[0],RPos(".",Local[0])+1,Len(Local[0])-RPos(".",Local[0])), S = Local[1] + '_' + Local[3] + '_'  + Local[4]

#define Homepage             "http://www.pdfforge.org"
#define SourceforgeHomepage  "http://www.sf.net/projects/pdfcreator"
#define Appname              "PDFCreator"
#define AppExename           "PDFCreator.exe"
#define SpoolerExename       "PDFSpooler.exe"

#define AppVersion           GetFileVersionVBExe("..\PDFCreator\PDFCreator.exe")

#define PDFCreatorVersion    GetFileVersionVBExe("..\PDFCreator\PDFCreator.exe")
#define SetupAppVersion      GetFileVersionVBExeLine("..\PDFCreator\PDFCreator.exe")
#define PDFSpoolerVersion    GetFileVersionVBExe("..\PDFSpooler\PDFSpooler.exe")
#define TransToolVersion     GetFileVersionVBExe("..\Transtool\Transtool.exe")

#define ReleaseCandidate     ""

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

#define DefaultPrinterMonitorname   "PDFCreator"
#define DefaultPrinterPortname      "PDFCreator:"
#define DefaultPrinterDrivername    "PDFCreator"
#define DefaultPrintername          "PDFCreator"

;#define UpdateIsPossible
#define UpdateIsPossibleMinVersion "0.9.2"

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
AppPublisher=Philip Chinery, Frank Heindörfer
AppPublisherURL={#Homepage}
AppSupportURL={#Homepage}
AppUpdatesURL={#Homepage}
AppVersion={#AppVersion}
ArchitecturesAllowed=x86 x64
ChangesAssociations=true
Compression={#CompressionMode}
CreateUninstallRegKey=false
DefaultDirName={reg:HKLM\{#UninstallRegStr2},Inno Setup: App Path|{pf}\{#AppName}}
DefaultGroupName={#AppName}
DisableDirPage=false
DisableStartupPrompt=true
ExtraDiskSpaceRequired=10303775
InternalCompressLevel={#SetupLZMACompressionMode}
LicenseFile=.\License\Program license - english.rtf
#Ifdef IncludeGhostscript
OutputBaseFilename={#AppName}-{#SetupAppVersionStr}_{#GhostscriptSetupString}
#ELSE
OutputBaseFilename={#AppName}-{#SetupAppVersionStr}_WithoutGhostscript
#ENDIF
OutputDir=Installation
RestartIfNeededByRun=true
ShowLanguageDialog=yes
ShowTasksTreeLines=false
SolidCompression=true
UsePreviousAppDir=true

VersionInfoVersion={#AppVersion}
VersionInfoCompany=Frank Heindörfer, Philip Chinery
VersionInfoDescription=PDFCreator is the easy way of creating PDFs.
VersionInfoTextVersion={#AppVersion}

WizardImageFile=..\Pictures\Setup\PDFCreatorBig.bmp
WizardSmallImageFile=..\Pictures\Setup\PDFCreator.bmp

MinVersion=4.0.950,4.0.1381
OnlyBelowVersion=0,6.0

[InstallDelete]
#Ifdef GhostscriptVersion
Name: {app}\Gs{#GhostscriptVersion}\Fonts\*.*; Type: filesandordirs; Components: ghostscript
Name: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib\*.*; Type: filesandordirs; Components: ghostscript
Name: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin\gsdll32.dll; Type: files; Components: ghostscript
#ENDIF
Name: {app}\languages\*.ini; Type: files; Components: program
Name: {app}\unload.tmp; Type: files; Components: program

[Files]
#IFNDEF Test
;We sort all files by extension for a maximal compression
;Systemfiles
Source: ..\SystemFiles\ASYCFILT.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace uninsneveruninstall

;psapi.dll - only for NT4 to enum processes
Source: ..\SystemFiles\PSAPI.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace uninsneveruninstall; MinVersion: 0,4.0.1381; OnlyBelowVersion: 0,5.0.2195

;Please use newest MSVBVM60.DLL
;http://support.microsoft.com/default.aspx?scid=kb;en-us;823746
Source: ..\SystemFiles\MSVBVM60.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace regserver uninsneveruninstall

Source: ..\SystemFiles\MSMPIDE.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt
Source: ..\SystemFiles\OLEPRO32.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace regserver uninsneveruninstall
Source: ..\SystemFiles\OLEAUT32.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace regserver uninsneveruninstall

;Language satellite system files
;http://msdn.microsoft.com/vbasic/downloads/tools/ipdk.aspx
;Language: German
;Source: C:\IPDK\German\CMCT3DE.DLL; DestDir: {sys}; Components: program; Flags: sharedfile uninsnosharedfileprompt; Check: IsLanguage('german')
Source: C:\IPDK\German\MSCC2DE.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('german')
Source: C:\IPDK\German\MSCMCDE.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('german')
Source: C:\IPDK\German\VB6DE.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('german')
;Language: Italian
;Source: C:\IPDK\Italian\CMCT3IT.DLL; DestDir: {sys}; Components: program; Flags: sharedfile uninsnosharedfileprompt; Check: IsLanguage('italian')
Source: C:\IPDK\Italian\MSCC2IT.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('italian')
Source: C:\IPDK\Italian\MSCMCIT.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('italian')
Source: C:\IPDK\Italian\VB6IT.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('italian')
;Language: French
;Source: C:\IPDK\French\CMCT3FR.DLL; DestDir: {sys}; Components: program; Flags: sharedfile uninsnosharedfileprompt; Check: IsLanguage('french')
Source: C:\IPDK\French\MSCC2FR.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('french')
Source: C:\IPDK\French\MSCMCFR.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('french')
Source: C:\IPDK\French\VB6FR.DLL; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt; Check: IsLanguage('french')

;Printerdriver files
;PPD-File
; Win9x/Me
Source: ..\Printer\Adobe\PDFCREATOR_german.PPD; DestName: ADIST5.PPD; DestDir: {code:PrinterDriverDirectory|Windows 4.0}; Components: program; Flags: ignoreversion; Check: InstallWin9xPrinterdriver AND IsLanguage('german') AND NOT UseOwnPPDFile
Source: ..\Printer\Adobe\PDFCREATOR_english.PPD; DestName: ADIST5.PPD; DestDir: {code:PrinterDriverDirectory|Windows 4.0}; Components: program; Flags: ignoreversion; Check: InstallWin9xPrinterdriver AND NOT IsLanguage('german') AND NOT UseOwnPPDFile
Source: {code:GetExternalPPDFile}; DestName: ADIST5.PPD; DestDir: {code:PrinterDriverDirectory|Windows 4.0}; Components: program; Flags: ignoreversion external; Check: UseOwnPPDFile
; WinNt4, Win2k, WinXP, Win2k3 - 32bit
Source: ..\Printer\Adobe\PDFCREATOR_german.PPD; DestName: PDFCREAT.PPD; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: (InstallWinNtPrinterdriver OR InstallWin2kXP2k3Printerdriver32bit) AND IsLanguage('german') AND NOT UseOwnPPDFile
Source: ..\Printer\Adobe\PDFCREATOR_english.PPD; DestName: PDFCREAT.PPD; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: (InstallWinNtPrinterdriver OR InstallWin2kXP2k3Printerdriver32bit) AND NOT IsLanguage('german') AND NOT UseOwnPPDFile
Source: {code:GetExternalPPDFile}; DestName: PDFCREAT.PPD; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion  deleteafterinstall external; Check: UseOwnPPDFile
; WinXP, Win2k3 - 64bit
Source: ..\Printer\Adobe\PDFCREATOR_german.PPD; DestName: PDFCREAT.PPD; DestDir: {code:PrinterDriverDirectory|Windows x64}; Components: program; Flags: ignoreversion deleteafterinstall; Check: (InstallWinXP2k3Printerdriver64bit) AND IsLanguage('german') AND NOT UseOwnPPDFile
Source: ..\Printer\Adobe\PDFCREATOR_english.PPD; DestName: PDFCREAT.PPD; DestDir: {code:PrinterDriverDirectory|Windows x64}; Components: program; Flags: ignoreversion deleteafterinstall; Check: (InstallWinXP2k3Printerdriver64bit) AND NOT IsLanguage('german') AND NOT UseOwnPPDFile
Source: {code:GetExternalPPDFile}; DestName: PDFCREAT.PPD; DestDir: {code:PrinterDriverDirectory|Windows x64}; Components: program; Flags: ignoreversion  deleteafterinstall external; Check: UseOwnPPDFile

;Driver files
; Win9x/Me
Source: ..\Printer\Adobe\Windows\ICONLIB.DLL; DestDir: {code:PrinterDriverDirectory|Windows 4.0}; Components: program; Check: InstallWin9xPrinterdriver
Source: ..\Printer\Adobe\Windows\PSMON.DLL; DestDir: {code:PrinterDriverDirectory|Windows 4.0}; Components: program; Check: InstallWin9xPrinterdriver
Source: ..\Printer\Adobe\Windows\ADOBEPS4.HLP; DestDir: {code:PrinterDriverDirectory|Windows 4.0}; Components: program; Flags: ignoreversion; Check: InstallWin9xPrinterdriver
Source: ..\Printer\Adobe\Windows\FONTSDIR.MFD; DestDir: {win}; Flags: ignoreversion; Components: program; Check: InstallWin9xPrinterdriver
Source: ..\Printer\Adobe\Windows\adfonts.mfm; DestDir: {code:PrinterDriverDirectory|Windows 4.0}; Components: program; Flags: ignoreversion; Check: InstallWin9xPrinterdriver
Source: ..\Printer\Adobe\Windows\ADOBEPS4.DRV; DestDir: {code:PrinterDriverDirectory|Windows 4.0}; Components: program; Check: InstallWin9xPrinterdriver
; WinNt 4.0
Source: ..\Printer\Adobe\WinNT\AdobePS5.dll; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWinNtPrinterdriver
Source: ..\Printer\Adobe\WinNT\AdobePSu.dll; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWinNtPrinterdriver
Source: ..\Printer\Adobe\WinNT\ADOBEPSU.HLP; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWinNtPrinterdriver
Source: ..\Printer\Adobe\WinNT\AdobePS5.ntf; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWinNtPrinterdriver
; Win2000
; Win2000: english files
Source: ..\Printer\Adobe\Win2000\English\PS5UI.DLL; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Languages: Not german
Source: ..\Printer\Adobe\Win2000\English\PSCRIPT5.DLL; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Languages: Not german
Source: ..\Printer\Adobe\Win2000\English\PSCRIPT.HLP; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Languages: Not german
; Win2000: german files
Source: ..\Printer\Adobe\Win2000\German\PS5UI.DLL; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Languages: german
Source: ..\Printer\Adobe\Win2000\German\PSCRIPT5.DLL; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Languages: german
Source: ..\Printer\Adobe\Win2000\German\PSCRIPT.HLP; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600; Languages: german
; Win2000: common files
Source: ..\Printer\Adobe\Win2000\PSCRIPT.NTF; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600
Source: ..\Printer\Adobe\Win2000\PSCRPTFE.NTF; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.0.2195; OnlyBelowVersion: 0,5.01.2600
; WinXP, Win2003 - x86 (32bit)
; WinXP, Win2003 - x86 (32bit): english files
Source: ..\Printer\Adobe\WinXP2k3-x86\English\PS5UI.DLL; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0; Languages: Not german
Source: ..\Printer\Adobe\WinXP2k3-x86\English\PSCRIPT5.DLL; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0; Languages: Not german
Source: ..\Printer\Adobe\WinXP2k3-x86\English\PSCRIPT.HLP; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0; Languages: Not german
; WinXP, Win2003 - x86 (32bit): german files
Source: ..\Printer\Adobe\WinXP2k3-x86\German\PS5UI.DLL; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0; Languages: german
Source: ..\Printer\Adobe\WinXP2k3-x86\German\PSCRIPT5.DLL; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0; Languages: german
Source: ..\Printer\Adobe\WinXP2k3-x86\German\PSCRIPT.HLP; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0; Languages: german
; WinXP, Win2003 - x86 (32bit): common files
Source: ..\Printer\Adobe\WinXP2k3-x86\PSCRIPT.NTF; DestDir: {code:PrinterDriverDirectory|Windows NT x86}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWin2kXP2k3Printerdriver32bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0
; WinXP, Win2003 - x64 (64bit)
Source: ..\Printer\Adobe\WinXP2k3-x64\PS5UI.DLL; DestDir: {code:PrinterDriverDirectory|Windows x64}; Components: program; Flags: deleteafterinstall; Check: InstallWinXP2k3Printerdriver64bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0
Source: ..\Printer\Adobe\WinXP2k3-x64\PSCRIPT5.DLL; DestDir: {code:PrinterDriverDirectory|Windows x64}; Components: program; Flags: deleteafterinstall; Check: InstallWinXP2k3Printerdriver64bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0
Source: ..\Printer\Adobe\WinXP2k3-x64\PSCRIPT.HLP; DestDir: {code:PrinterDriverDirectory|Windows x64}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWinXP2k3Printerdriver64bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0
Source: ..\Printer\Adobe\WinXP2k3-x64\PSCRIPT.NTF; DestDir: {code:PrinterDriverDirectory|Windows x64}; Components: program; Flags: ignoreversion deleteafterinstall; Check: InstallWinXP2k3Printerdriver64bit; MinVersion: 0,5.01.2600; OnlyBelowVersion: 0,0

;Ghostscript
#IFDEF GhostscriptVersion
Source: C:\gs\{#GhostscriptLicense}\gs{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin\gsdll32.dll; DestDir: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Components: ghostscript; Flags: ignoreversion
Source: C:\gs\{#GhostscriptLicense}\gs{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin\gsdll32.lib; DestDir: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Components: ghostscript; Flags: ignoreversion
#ENDIF

;Redmon files
Source: ..\Printer\Redmon\redmon95.dll; Components: program; DestDir: {sys}; MinVersion: 4.00.950,0; DestName: pdfcmn95.dll; Flags: comparetimestamp
Source: ..\Printer\Redmon\redmonnt.dll; Components: program; DestDir: {sys}; MinVersion: 0,4.00.1381; DestName: pdfcmnnt.dll; Check: Not IsWin64; Flags: 32bit comparetimestamp
Source: ..\Printer\Redmon\redmonnt-x64.dll; Components: program; DestDir: {sys}; MinVersion: 0,5.02.3790; DestName: pdfcmnnt.dll; Check: IsX64; Flags: 64bit comparetimestamp

;Source: ..\SystemFiles\COMCT332.OCX; DestDir: {sys}; Components: program; Flags: sharedfile uninsnosharedfileprompt regserver
Source: ..\SystemFiles\MSCOMCT2.OCX; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt regserver
Source: ..\SystemFiles\MSCOMCTL.OCX; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt regserver
Source: ..\SystemFiles\MSMAPI32.OCX; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt regserver

Source: ..\SystemFiles\STDOLE2.TLB; DestDir: {sys}; Components: program; Flags: 32bit sharedfile uninsnosharedfileprompt restartreplace uninsneveruninstall regtypelib

;Program files
Source: ..\PDFCreator\PDFCreator.exe; DestDir: {app}; Components: program; Flags: comparetimestamp
Source: ..\Transtool\TransTool.exe; DestDir: {app}\languages; Components: program; Flags: comparetimestamp
Source: ..\PDFSpooler\PDFSpooler.exe; DestDir: {app}; Components: program; Flags: comparetimestamp

;vblocal.exe from IPDK
Source: C:\IPDK\vblocal.exe; DestDir: {app}; Components: program; Flags: deleteafterinstall overwritereadonly onlyifdoesntexist ignoreversion

;ShFolder for older systems
;http://www.microsoft.com/downloads/release.asp?releaseid=30340
Source: ShFolder\ShFolder.Exe; DestDir: {app}; Components: program; Flags: ignoreversion deleteafterinstall; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 4.1.2222,5.0.2195

;pdfenc
Source: pdfenc\pdfenc.exe; DestDir: {app}; Components: program; Flags: ignoreversion


;Help files
Source: ..\Help\english\PDFCreator_english.chm; DestDir: {app}; Components: program; Flags: ignoreversion
Source: ..\Help\german\PDFCreator_german.chm; DestDir: {app}; Components: program; Flags: ignoreversion
Source: ..\Help\french\PDFCreator_french.chm; DestDir: {app}; Components: program; Flags: ignoreversion

Source: License\AFPL License.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp
Source: License\GNU License.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp
Source: History.txt; DestDir: {app}; Components: program; Flags: ignoreversion comparetimestamp

;Languages
Source: ..\PDFCreator\Languages\catalan.ini; DestDir: {app}\languages; Components: languages\catalan; Flags: ignoreversion
;Source: ..\PDFCreator\Languages\chinese_simplified.ini; DestDir: {app}\languages; Components: languages\chinese_simplified; Flags: ignoreversion
Source: ..\PDFCreator\Languages\czech.ini; DestDir: {app}\languages; Components: languages\czech; Flags: ignoreversion
Source: ..\PDFCreator\Languages\dutch.ini; DestDir: {app}\languages; Components: languages\dutch; Flags: ignoreversion
Source: ..\PDFCreator\Languages\english.ini; DestDir: {app}\languages; Components: languages\english; Flags: ignoreversion
Source: ..\PDFCreator\Languages\french.ini; DestDir: {app}\languages; Components: languages\french; Flags: ignoreversion
Source: ..\PDFCreator\Languages\german.ini; DestDir: {app}\languages; Components: languages\german; Flags: ignoreversion
Source: ..\PDFCreator\Languages\hungarian.ini; DestDir: {app}\languages; Components: languages\hungarian; Flags: ignoreversion
;Source: ..\PDFCreator\Languages\indonesian.ini; DestDir: {app}\languages; Components: languages\indonesian; Flags: ignoreversion
Source: ..\PDFCreator\Languages\italian.ini; DestDir: {app}\languages; Components: languages\italian; Flags: ignoreversion
Source: ..\PDFCreator\Languages\lithuanian.ini; DestDir: {app}\languages; Components: languages\lithuanian; Flags: ignoreversion
Source: ..\PDFCreator\Languages\polish.ini; DestDir: {app}\languages; Components: languages\polish; Flags: ignoreversion
Source: ..\PDFCreator\Languages\romanian.ini; DestDir: {app}\languages; Components: languages\romanian; Flags: ignoreversion
Source: ..\PDFCreator\Languages\russian.ini; DestDir: {app}\languages; Components: languages\russian; Flags: ignoreversion
Source: ..\PDFCreator\Languages\slovak.ini; DestDir: {app}\languages; Components: languages\slovak; Flags: ignoreversion
Source: ..\PDFCreator\Languages\slovenian.ini; DestDir: {app}\languages; Components: languages\slovenian; Flags: ignoreversion
Source: ..\PDFCreator\Languages\spanish.ini; DestDir: {app}\languages; Components: languages\spanish; Flags: ignoreversion
Source: ..\PDFCreator\Languages\swedish.ini; DestDir: {app}\languages; Components: languages\swedish; Flags: ignoreversion
Source: ..\PDFCreator\Languages\turkish.ini; DestDir: {app}\languages; Components: languages\turkish; Flags: ignoreversion
Source: ..\PDFCreator\Languages\valencian.ini; DestDir: {app}\languages; Components: languages\valencian; Flags: ignoreversion

;Ini file
Source: PDFCreator.ini; DestDir: {code:GetIniPath}; Components: program; DestName: PDFCreator.ini; Flags: ignoreversion onlyifdoesntexist uninsneveruninstall; Check: (Not UseOwnINIFile) And UseINI
Source: {code:GetExternalINIFile}; DestName: PDFCreator.ini; DestDir: {code:GetIniPath}; Components: program; Flags: ignoreversion  external; Check: UseOwnINIFile AND UseINI
Source: PDFCreator.ini; DestDir: {code:GetDefaultIniPath}; Components: program; DestName: PDFCreator.ini; Flags: ignoreversion onlyifdoesntexist uninsneveruninstall; Check: (Not UseOwnINIFile) And UseINI
Source: {code:GetExternalINIFile}; DestName: PDFCreator.ini; DestDir: {code:GetDefaultIniPath}; Components: program; Flags: ignoreversion  external; Check: UseOwnINIFile AND UseINI

;Reg file
Source: {code:GetExternalREGFile}; DestName: PDFCreator-external.reg; DestDir: {%tmp}; Components: program; Flags: ignoreversion  external deleteafterinstall; Check: UseOwnREGFile AND (Not UseINI)

;Ghostscript
#IFDEF IncludeGhostscript
Source: C:\gs\{#GhostscriptLicense}\gs{#GhostscriptVersion}\Fonts\*.*; DestDir: {app}\Gs{#GhostscriptVersion}\Fonts; Components: ghostscript; Flags: ignoreversion sortfilesbyextension
Source: C:\gs\{#GhostscriptLicense}\gs{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib\*.*; DestDir: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Components: ghostscript; Flags: ignoreversion sortfilesbyextension
Source: C:\gs\{#GhostscriptLicense}\gs{#GhostscriptVersion}\gs{#GhostscriptVersion}\Resource\*.*; DestDir: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Resource; Components: ghostscript; Flags: ignoreversion sortfilesbyextension recursesubdirs
#ENDIF

; Scripts
; Scripts: RunProgramAfterSaving
Source: ..\Scripts\RunProgramAfterSaving\AddWatermarkToPDF.vbs; DestDir: {app}\Scripts\RunProgramAfterSaving; Components: program; Flags: ignoreversion
Source: ..\Scripts\RunProgramAfterSaving\FTPUpload.vbs; DestDir: {app}\Scripts\RunProgramAfterSaving; Components: program; Flags: ignoreversion
Source: ..\Scripts\RunProgramAfterSaving\Logger.vbs; DestDir: {app}\Scripts\RunProgramAfterSaving; Components: program; Flags: ignoreversion
Source: ..\Scripts\RunProgramAfterSaving\Watermark.pdf; DestDir: {app}\Scripts\RunProgramAfterSaving; Components: program; Flags: ignoreversion
Source: ..\Scripts\RunProgramAfterSaving\NetSend.vbs; DestDir: {app}\Scripts\RunProgramAfterSaving; Components: program; Flags: ignoreversion
Source: ..\Scripts\RunProgramAfterSaving\PopUpMessage.vbs; DestDir: {app}\Scripts\RunProgramAfterSaving; Components: program; Flags: ignoreversion
Source: ..\Scripts\RunProgramAfterSaving\SayIt.vbs; DestDir: {app}\Scripts\RunProgramAfterSaving; Components: program; Flags: ignoreversion
Source: ..\Scripts\RunProgramAfterSaving\MSAgent.vbs; DestDir: {app}\Scripts\RunProgramAfterSaving; Components: program; Flags: ignoreversion
; Scripts: RunProgramBeforSaving
Source: ..\Scripts\RunProgramBeforeSaving\AddBookmarks.vbs; DestDir: {app}\Scripts\RunProgramBeforeSaving; Components: program; Flags: ignoreversion
Source: ..\Scripts\RunProgramBeforeSaving\AddPDFDocview.vbs; DestDir: {app}\Scripts\RunProgramBeforeSaving; Components: program; Flags: ignoreversion
; Samples: Com
Source: ..\COM\Samples\VB6\Sample1\Form1.frm; DestDir: {app}\COM\VB6\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\VB6\Sample1\Form1.frx; DestDir: {app}\COM\VB6\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\VB6\Sample1\Sample1.RES; DestDir: {app}\COM\VB6\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\VB6\Sample1\Sample1.vbp; DestDir: {app}\COM\VB6\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\VB6\Sample1\Sample1.vbw; DestDir: {app}\COM\VB6\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\VB6\Sample2\Form1.frm; DestDir: {app}\COM\VB6\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\VB6\Sample2\Sample2.vbp; DestDir: {app}\COM\VB6\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\VB6\Sample2\Sample2.vbw; DestDir: {app}\COM\VB6\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\DOTNET Scripting Host\readme.txt; DestDir: {app}\COM\DOTNET Scripting Host; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\DOTNET Scripting Host\Sample1.dsh; DestDir: {app}\COM\DOTNET Scripting Host; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\C#\Sample1\Form1.resx; DestDir: {app}\COM\Dot Net\VS2003\C#\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\C#\Sample1\Sample1.csproj; DestDir: {app}\COM\Dot Net\VS2003\C#\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\C#\Sample1\AssemblyInfo.cs; DestDir: {app}\COM\Dot Net\VS2003\C#\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\C#\Sample1\Form1.cs; DestDir: {app}\COM\Dot Net\VS2003\C#\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\C#\Sample2\Form1.resx; DestDir: {app}\COM\Dot Net\VS2003\C#\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\C#\Sample2\Sample2.csproj; DestDir: {app}\COM\Dot Net\VS2003\C#\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\C#\Sample2\AssemblyInfo.cs; DestDir: {app}\COM\Dot Net\VS2003\C#\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\C#\Sample2\Form1.cs; DestDir: {app}\COM\Dot Net\VS2003\C#\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\Visual Basic\Sample1\AssemblyInfo.vb; DestDir: {app}\COM\Dot Net\VS2003\Visual Basic\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\Visual Basic\Sample1\Form1.resx; DestDir: {app}\COM\Dot Net\VS2003\Visual Basic\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\Visual Basic\Sample1\Form1.vb; DestDir: {app}\COM\Dot Net\VS2003\Visual Basic\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\Visual Basic\Sample1\Sample1.vbproj; DestDir: {app}\COM\Dot Net\VS2003\Visual Basic\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\Visual Basic\Sample2\AssemblyInfo.vb; DestDir: {app}\COM\Dot Net\VS2003\Visual Basic\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\Visual Basic\Sample2\Form1.resx; DestDir: {app}\COM\Dot Net\VS2003\Visual Basic\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\Visual Basic\Sample2\Form1.vb; DestDir: {app}\COM\Dot Net\VS2003\Visual Basic\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2003\Visual Basic\Sample2\Sample2.vbproj; DestDir: {app}\COM\Dot Net\VS2003\Visual Basic\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\C#\Sample1\Form1.resx; DestDir: {app}\COM\Dot Net\VS2005\C#\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\C#\Sample1\Sample1.csproj; DestDir: {app}\COM\Dot Net\VS2005\C#\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\C#\Sample1\AssemblyInfo.cs; DestDir: {app}\COM\Dot Net\VS2005\C#\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\C#\Sample1\Form1.cs; DestDir: {app}\COM\Dot Net\VS2005\C#\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\C#\Sample2\Form1.resx; DestDir: {app}\COM\Dot Net\VS2005\C#\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\C#\Sample2\Sample2.csproj; DestDir: {app}\COM\Dot Net\VS2005\C#\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\C#\Sample2\AssemblyInfo.cs; DestDir: {app}\COM\Dot Net\VS2005\C#\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\C#\Sample2\Form1.cs; DestDir: {app}\COM\Dot Net\VS2005\C#\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\Visual Basic\Sample1\AssemblyInfo.vb; DestDir: {app}\COM\Dot Net\VS2005\Visual Basic\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\Visual Basic\Sample1\Form1.resx; DestDir: {app}\COM\Dot Net\VS2005\Visual Basic\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\Visual Basic\Sample1\Form1.vb; DestDir: {app}\COM\Dot Net\VS2005\Visual Basic\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\Visual Basic\Sample1\Sample1.vbproj; DestDir: {app}\COM\Dot Net\VS2005\Visual Basic\Sample1; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\Visual Basic\Sample2\AssemblyInfo.vb; DestDir: {app}\COM\Dot Net\VS2005\Visual Basic\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\Visual Basic\Sample2\Form1.resx; DestDir: {app}\COM\Dot Net\VS2005\Visual Basic\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\Visual Basic\Sample2\Form1.vb; DestDir: {app}\COM\Dot Net\VS2005\Visual Basic\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Dot Net\VS2005\Visual Basic\Sample2\Sample2.vbproj; DestDir: {app}\COM\Dot Net\VS2005\Visual Basic\Sample2; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\MS Office\frmPDFCreatorExcel.frm; DestDir: {app}\COM\MS Office; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\MS Office\frmPDFCreatorExcel.frx; DestDir: {app}\COM\MS Office; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\MS Office\frmPDFCreatorWord.frm; DestDir: {app}\COM\MS Office; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\MS Office\frmPDFCreatorWord.frx; DestDir: {app}\COM\MS Office; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\MS Office\modPDFCreatorAccess2000.bas; DestDir: {app}\COM\MS Office; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\CombineAndAddBookmarks.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\CombineJobs.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\CompareColorCompressionModes.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\Convert2PDF.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\Convert2TIFF.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\Convert2TXT.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\GUI.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\SaveOptionsToFile.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\ShowLogfile.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\ShowOptions.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\ShowPrintjobInfos.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\TestCompression1.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\TestCompression2.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\TestCompression3.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\TestEvents.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\Testpage2PDF.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\Windows Scripting Host\VBScripts\Testpage2PDFSendEmail.vbs; DestDir: {app}\COM\Windows Scripting Host\VBScripts; Components: program; Flags: ignoreversion
Source: ..\COM\Samples\WinBatch\Convert2PDF.wbt; DestDir: {app}\COM\WinBatch; Components: program; Flags: ignoreversion

; Toolbar
#IFDEF IncludeToolbar
Source: ..\Pictures\Toolbar\Toolbar.bmp; DestDir: {tmp}; Flags: dontcopy nocompression; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,0
Source: ..\Toolbar\PDFCreator_Toolbar_Setup.exe; DestDir: {tmp}; DestName: PDFCreator_Toolbar_Setup.exe; Components: ietoolbar; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,0
#ENDIF
#ENDIF

[Dirs]
Name: {code:GetPrinterTemppath}; Flags: uninsalwaysuninstall

[Icons]
Name: {group}\{#Appname}; Filename: {app}\{#AppExename}; WorkingDir: {app}; Flags: createonlyiffileexists
Name: {group}\AFPL License; Filename: {app}\AFPL License.txt; WorkingDir: {app}
Name: {group}\GPL License; Filename: {app}\GNU License.txt; WorkingDir: {app}
Name: {group}\{cm:History}; Filename: {app}\History.txt; WorkingDir: {app}; Flags: createonlyiffileexists
Name: {group}\Translation Tool; Filename: {app}\languages\transtool.exe; WorkingDir: {app}\languages; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\{cm:ProgramOnTheWeb,PDFCreator}; Filename: {app}\PDFCreator.url; WorkingDir: {app}
Name: {group}\PDFCreator {cm:Help}; Filename: {app}\PDFCreator_english.chm; WorkingDir: {app}; Languages: (Not german) AND (Not french)
Name: {group}\PDFCreator {cm:Help}; Filename: {app}\PDFCreator_german.chm; WorkingDir: {app}; Languages: german
Name: {group}\PDFCreator {cm:Help}; Filename: {app}\PDFCreator_french.chm; WorkingDir: {app}; Languages: french

Name: {group}\{cm:Logfile}; Filename: {app}\PDFCreator.exe; Parameters: -ShowOnlyLogfile; WorkingDir: {app}; IconIndex: 0; Check: IsServerInstallation
Name: {group}\{cm:Settings}; Filename: {app}\PDFCreator.exe; Parameters: -ShowOnlyOptions; WorkingDir: {app}; IconIndex: 0; Check: IsServerInstallation

Name: {commondesktop}\PDFCreator; Filename: {app}\PDFCreator.exe; WorkingDir: {app}; IconIndex: 0; Tasks: desktopicon\common
Name: {userdesktop}\PDFCreator; Filename: {app}\PDFCreator.exe; WorkingDir: {app}; IconIndex: 0; Tasks: desktopicon\user
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\PDFCreator; Filename: {app}\PDFCreator.exe; WorkingDir: {app}; IconIndex: 0; Tasks: quicklaunchicon

[INI]
Filename: {app}\PDFCreator.url; Section: InternetShortcut; Key: URL; String: http://www.pdfforge.org; Components: program
Filename: {app}\PDFCreator.url; Section: InternetShortcut; Key: Iconindex; String: 1; Components: program

Filename: {app}\{cm:Donation}.url; Section: InternetShortcut; Key: URL; String: http://www.paypal.com/xclick/business=paypal01%40heindoerfer.com&item_name=PDFCreator&no_note=1&tax=0&currency_code=EUR; Components: program
Filename: {app}\{cm:Donation}.url; Section: InternetShortcut; Key: Iconindex; String: 1; Components: program

Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: AutosaveDirectory; String: <MyFiles>; Components: program; Flags: createkeyifdoesntexist; Check: UseINI  And (Not IsServerInstallation)
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: AutosaveDirectory; String: C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>; Components: program; Flags: createkeyifdoesntexist; Check: UseINI And IsServerInstallation
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: LastsaveDirectory; String: <MyFiles>; Components: program; Flags: createkeyifdoesntexist; Check: UseINI And (Not IsServerInstallation)
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: LastsaveDirectory; String: C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>; Components: program; Flags: createkeyifdoesntexist; Check: UseINI And IsServerInstallation
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: Language; String: {code:GetActiveLanguage}; Flags: createkeyifdoesntexist; Check: UseINI
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: PrinterTemppath; String: <Temp>PDFCreator\; Flags: createkeyifdoesntexist; Check: UseINI And (Not IsServerInstallation)
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: PrinterTemppath; String: {app}\Temp\; Flags: createkeyifdoesntexist; Check: UseINI And IsServerInstallation

#Ifdef GhostscriptVersion
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: DirectoryGhostscriptBinaries; String: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Components: ghostscript; Flags: createkeyifdoesntexist; Check: UseINI
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: DirectoryGhostscriptFonts; String: {app}\Gs{#GhostscriptVersion}\Fonts; Components: ghostscript; Flags: createkeyifdoesntexist; Check: UseINI
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: DirectoryGhostscriptLibraries; String: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Components: ghostscript; Flags: createkeyifdoesntexist; Check: UseINI
Filename: {code:GetIniPath}\PDFCreator.ini; Section: Options; Key: DirectoryGhostscriptResource; String: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Resource; Components: ghostscript; Flags: createkeyifdoesntexist; Check: UseINI
#ENDIF

[Registry]
;PrinterMonitor
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: string; Valuename: Arguments; ValueData: -PPDFCREATORPRINTER; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: string; Valuename: Command; ValueData: {code:GetShortname|{syswow64}\{#SpoolerExename}}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: dword; Valuename: Delay; ValueData: 300; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: string; Valuename: Description; ValueData: PDFCreator Redirected Port; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: dword; Valuename: LogFileDebug; ValueData: 0; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: dword; Valuename: LogFileUse; ValueData: 0; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: dword; Valuename: Output; ValueData: 0; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: string; Valuename: Printer; ValueData: {code:GetPrintername}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: dword; Valuename: Printerror; ValueData: 0; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: dword; Valuename: Runuser; ValueData: 0; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; ValueType: dword; Valuename: ShowWindow; ValueData: 0; Flags: uninsdeletevalue

;Uninstall - Deletekey
Root: HKLM; Subkey: {#PrintReg}Printers\{code:GetPrintername}; Flags: uninsdeletekey dontcreatekey
Root: HKLM; Subkey: {#PrintReg}Environments\Windows 4.0\Drivers\{code:GetPrinterdrivername}; Flags: uninsdeletekey dontcreatekey; MinVersion: 4.00.950,0
Root: HKLM; Subkey: {#PrintReg}Environments\Windows NT x86\Drivers\{code:GetPrinterdrivername}; Flags: uninsdeletekey dontcreatekey; MinVersion: 0,4.00.1381
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}\Ports\{code:GetPrinterportname}; Flags: uninsdeletekey dontcreatekey
Root: HKLM; Subkey: {#PrintRegMon}{code:GetPrintermonitorname}; Flags: uninsdeletekey dontcreatekey

;File-Assoc
Root: HKCR; SubKey: .ps; ValueType: string; ValueData: PostScript; Flags: uninsdeletekeyifempty noerror; Tasks: fileassoc
Root: HKCR; SubKey: PostScript\Shell\Open\Command; ValueType: string; ValueData: """{app}\PDFCreator.exe"" -IF""%1"""; Flags: uninsdeletevalue uninsdeletekeyifempty noerror; Tasks: fileassoc
Root: HKCR; Subkey: PostScript\DefaultIcon; ValueType: string; ValueData: {app}\PDFCreator.exe,0; Flags: uninsdeletevalue uninsdeletekeyifempty noerror; Tasks: fileassoc
Root: HKCR; SubKey: PostScript; ValueType: string; ValueData: PostScript; Flags: uninsdeletekeyifempty noerror; Tasks: fileassoc

;Uninstall - Software
Root: HKLM; Subkey: {#UninstallRegStr}; Flags: uninsdeletekey
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Comments; Valuedata: PDFCreator - Opensource; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayIcon; Valuedata: {app}\PDFCreator.exe; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayName; Valuedata: {#AppName} {#AppVersionStr}; Flags: uninsdeletevalue; MinVersion: 4.0.950,0; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayName; Valuedata: {#AppName}; Flags: uninsdeletevalue; MinVersion: 0,4.0.1381; OnlyBelowVersion: 0,0
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: DisplayVersion; Valuedata: {#AppVersionStr}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: HelpLink; Valuedata: {#SourceforgeHomepage}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: InstallDate; Valuedata: {code:GetDateString}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Publisher; Valuedata: Frank Heindörfer, Philip Chinery; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Readme; Valuedata: {#Homepage}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: URLInfoAbout; Valuedata: {#SourceforgeHomepage}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: URLUpdateInfo; Valuedata: {#SourceforgeHomepage}; Flags: uninsdeletevalue

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: ApplicationVersion; Valuedata: {#AppVersion}; Flags: uninsdeletevalue
#IF (BetaVersion!="")
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: BetaVersion; Valuedata: {#BetaVersion}; Flags: uninsdeletevalue
#ENDIF
#IF (ReleaseCandidate!="")
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: ReleaseCandidate; Valuedata: {#ReleaseCandidate}; Flags: uninsdeletevalue
#ENDIF
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: PatchLevel; Valuedata: ; Flags: uninsdeletevalue

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: PDFCreatorVersion; Valuedata: {#PDFCreatorVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: PDFSpoolerVersion; Valuedata: {#PDFSpoolerVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: TranstoolVersion; Valuedata: {#TranstoolVersion}; Flags: uninsdeletevalue

#Ifdef GhostscriptVersion
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptCopyright; Valuedata: {#GhostscriptLicense}; Flags: uninsdeletevalue; Components: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptVersion; Valuedata: {#GhostscriptVersion}; Flags: uninsdeletevalue; Components: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptDirectoryBinaries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Flags: uninsdeletevalue; Components: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptDirectoryLibraries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Flags: uninsdeletevalue; Components: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptDirectoryFonts; Valuedata: {app}\Gs{#GhostscriptVersion}\Fonts; Flags: uninsdeletevalue; Components: ghostscript
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: GhostscriptDirectoryResource; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Resource; Flags: uninsdeletevalue; Components: ghostscript
#Endif

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: UninstallString; Valuedata: {app}\unins000.exe; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Printername; Valuedata: {code:GetPrintername}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Printerdrivername; Valuedata: {code:GetPrinterdrivername}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Printerportname; Valuedata: {code:GetPrinterportname}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Printermonitorname; Valuedata: {code:GetPrintermonitorname}; Flags: uninsdeletevalue

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: App Path; Valuedata: {app}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Components; Valuedata: {code:GetWizardSelectedComponents}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Tasks; Valuedata: {code:GetWizardSelectedTasks}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Silent; Valuedata: {code:GetWizardSilent}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: Group; Valuedata: {code:GetWizardGroupValue}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: NoIcons; Valuedata: {code:GetWizardNoIcons}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: SetupType; Valuedata: {code:GetWizardSetupType}; Flags: uninsdeletevalue
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: Inno Setup: SetupLanguage; Valuedata: {code:GetActiveLanguage}; Flags: uninsdeletevalue

Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: PDFServer; Valuedata: 1; Flags: uninsdeletevalue; Check: IsServerInstallation
Root: HKLM; Subkey: {#UninstallRegStr}; ValueType: string; ValueName: UseINI; Valuedata: 1; Flags: uninsdeletevalue; Check: UseINI

; Remove keys and/or values on uninstall
Root: HKLM; Subkey: SOFTWARE\Classes\Applications\PDFCreator.exe; Flags: uninsdeletekey noerror dontcreatekey deletekey
Root: HKCU; Subkey: Printers\Settings; ValueName: {code:GetPrintername}; Flags: uninsdeletevalue noerror dontcreatekey
Root: HKCU; Subkey: Software\Microsoft\Windows\CurrentVersion\Explorer\MenuOrder\Start Menu\Programs\PDFCreator; Flags: uninsdeletekey noerror dontcreatekey

;CustomMessages for uninstall. InnoSetup doesn't support custom messages for uninstalling at the moment.
Root: HKLM; Subkey: {#UninstallRegStr}\CustomMessages; ValueType: string; ValueName: UninstallOptions; Valuedata: {cm:UninstallOptions}; Flags: uninsdeletevalue

; Application
Root: HKLM; Subkey: Software\PDFCreator\Program; ValueType: string; ValueName: AutosaveDirectory; Valuedata: C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>; Check: IsServerInstallation; Flags: createvalueifdoesntexist
Root: HKLM; Subkey: Software\PDFCreator\Program; ValueType: string; ValueName: LastsaveDirectory; Valuedata: C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>; Flags: createvalueifdoesntexist; Check: IsServerInstallation
Root: HKLM; Subkey: Software\PDFCreator\Program; ValueType: string; ValueName: Language; Valuedata: {code:GetActiveLanguage}; Flags: createvalueifdoesntexist; Check: IsServerInstallation
Root: HKLM; Subkey: Software\PDFCreator\Program; ValueType: string; ValueName: PrinterTemppath; Valuedata: {app}\Temp\; Flags: createvalueifdoesntexist; Check: IsServerInstallation

Root: HKU; Subkey: .DEFAULT\Software\PDFCreator\Program; ValueType: string; ValueName: AutosaveDirectory; Valuedata: <MyFiles>; MinVersion: 0,4.0.1381; Flags: createvalueifdoesntexist; Check: Not IsServerInstallation
Root: HKU; Subkey: .DEFAULT\Software\PDFCreator\Program; ValueType: string; ValueName: LastsaveDirectory; Valuedata: <MyFiles>; MinVersion: 0,4.0.1381; Flags: createvalueifdoesntexist; Check: Not IsServerInstallation
Root: HKU; Subkey: .DEFAULT\Software\PDFCreator\Program; ValueType: string; ValueName: Language; Valuedata: {code:GetActiveLanguage}; MinVersion: 0,4.0.1381; Flags: createvalueifdoesntexist; Check: Not IsServerInstallation
Root: HKU; Subkey: .DEFAULT\Software\PDFCreator\Program; ValueType: string; ValueName: PrinterTemppath; Valuedata: <Temp>PDFCreator\; MinVersion: 0,4.0.1381; Flags: createvalueifdoesntexist; Check: Not IsServerInstallation

Root: HKCU; Subkey: Software\PDFCreator\Program; ValueType: string; ValueName: AutosaveDirectory; Valuedata: <MyFiles>; Flags: createvalueifdoesntexist; Check: Not IsServerInstallation
Root: HKCU; Subkey: Software\PDFCreator\Program; ValueType: string; ValueName: LastsaveDirectory; Valuedata: <MyFiles>; Flags: createvalueifdoesntexist; Check: Not IsServerInstallation
Root: HKCU; Subkey: Software\PDFCreator\Program; ValueType: string; ValueName: Language; Valuedata: {code:GetActiveLanguage}; Flags: createvalueifdoesntexist; Check: Not IsServerInstallation
Root: HKCU; Subkey: Software\PDFCreator\Program; ValueType: string; ValueName: PrinterTemppath; Valuedata: <Temp>PDFCreator\; Flags: createvalueifdoesntexist; Check: Not IsServerInstallation

#Ifdef GhostscriptVersion
Root: HKLM; Subkey: Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptBinaries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Components: ghostscript; Flags: uninsdeletevalue; Check: IsServerInstallation
Root: HKLM; Subkey: Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptFonts; Valuedata: {app}\Gs{#GhostscriptVersion}\Fonts; Components: ghostscript; Flags: uninsdeletevalue; Check: IsServerInstallation
Root: HKLM; Subkey: Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptLibraries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Components: ghostscript; Flags: uninsdeletevalue; Check: IsServerInstallation
Root: HKLM; Subkey: Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptResource; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Resource; Components: ghostscript; Flags: uninsdeletevalue; Check: IsServerInstallation

Root: HKU; Subkey: .DEFAULT\Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptBinaries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Components: ghostscript; MinVersion: 0,4.0.1381; Flags: uninsdeletevalue; Check: Not IsServerInstallation
Root: HKU; Subkey: .DEFAULT\Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptFonts; Valuedata: {app}\Gs{#GhostscriptVersion}\Fonts; Components: ghostscript; Flags: uninsdeletevalue; MinVersion: 0,4.0.1381; Check: Not IsServerInstallation
Root: HKU; Subkey: .DEFAULT\Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptLibraries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Components: ghostscript; MinVersion: 0,4.0.1381; Flags: uninsdeletevalue; Check: Not IsServerInstallation
Root: HKU; Subkey: .DEFAULT\Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptResource; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Resource; Components: ghostscript; MinVersion: 0,4.0.1381; Flags: uninsdeletevalue; Check: Not IsServerInstallation

Root: HKCU; Subkey: Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptBinaries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Bin; Components: ghostscript; Flags: uninsdeletevalue; Check: Not IsServerInstallation
Root: HKCU; Subkey: Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptFonts; Valuedata: {app}\Gs{#GhostscriptVersion}\Fonts; Components: ghostscript; Flags: uninsdeletevalue; Check: Not IsServerInstallation
Root: HKCU; Subkey: Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptLibraries; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Lib; Components: ghostscript; Flags: uninsdeletevalue; Check: Not IsServerInstallation
Root: HKCU; Subkey: Software\PDFCreator\Ghostscript; ValueType: string; ValueName: DirectoryGhostscriptResource; Valuedata: {app}\GS{#GhostscriptVersion}\gs{#GhostscriptVersion}\Resource; Components: ghostscript; Flags: uninsdeletevalue; Check: Not IsServerInstallation
#ENDIF

[Run]
#IFNDEF Test
;german localization
Filename: {app}\vblocal.Exe; WorkingDir: {syswow64}; Parameters: pdfspooler.exe vb6de.dll 0x407 * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('german')
Filename: {app}\vblocal.Exe; WorkingDir: {app}; Parameters: pdfcreator.exe vb6de.dll 0x407 * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('german')
Filename: {app}\vblocal.Exe; WorkingDir: {app}\Languages; Parameters: transtool.exe vb6de.dll 0x407 * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('german')
;italian localization
Filename: {app}\vblocal.Exe; WorkingDir: {syswow64}; Parameters: pdfspooler.exe vb6it.dll 0x410 * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('italian')
Filename: {app}\vblocal.Exe; WorkingDir: {app}; Parameters: pdfcreator.exe vb6it.dll 0x410 * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('italian')
Filename: {app}\vblocal.Exe; WorkingDir: {app}\Languages; Parameters: transtool.exe vb6it.dll 0x410 * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('italian')
;french localization
Filename: {app}\vblocal.Exe; WorkingDir: {syswow64}; Parameters: pdfspooler.exe vb6fr.dll 0x40C * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('french')
Filename: {app}\vblocal.Exe; WorkingDir: {app}; Parameters: pdfcreator.exe vb6fr.dll 0x40C * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('french')
Filename: {app}\vblocal.Exe; WorkingDir: {app}\Languages; Parameters: transtool.exe vb6fr.dll 0x40C * 0x409; Flags: runhidden; Components: program; Check: IsLanguage('french')

Filename: {app}\ShFolder.Exe; WorkingDir: {app}; Parameters: /Q:A; Flags: runminimized; Components: program; MinVersion: 4.0.950,4.0.1381; OnlyBelowVersion: 4.1.2222,5.0.2195
Filename: {app}\PDFCreator.exe; WorkingDir: {app}; Parameters: /RegServer; Flags: nowait
Filename: {app}\PDFCreator.exe; WorkingDir: {app}; Description: {cm:LaunchProgram,{#Appname}}; Flags: postinstall nowait skipifsilent; Check: IsServerInstallation
Filename: {app}\SetupLog.txt; Description: SetupLog.txt; Flags: postinstall shellexec skipifsilent; Check: Not IsPrinterInstallationSuccessfully

Filename: regedit.exe; WorkingDir: {%tmp}; Parameters: /s {%tmp}\PDFCreator-external.reg; Components: program; Flags: runhidden; Check: UseOwnREGFile AND (Not UseINI)
#ENDIF

#IFDEF IncludeToolbar
Filename: {tmp}\PDFCreator_Toolbar_Setup.exe; Components: ietoolbar; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,0
#ENDIF

[UninstallRun]
Filename: {app}\PDFCreator.exe; Parameters: /UnRegServer; Flags: skipifdoesntexist runhidden

[UninstallDelete]
Name: {app}\SetupLog.txt; Type: files
Name: {app}\Unload.tmp; Type: files
Name: {app}\PDFCreator.url; Type: files
Name: {app}\{cm:Donation}.url; Type: files
Name: {app}\PDFCreatorSpool; Type: filesandordirs
Name: {app}\Temp; Type: filesandordirs
Name: {app}\languages; Type: filesandordirs
Name: {app}; Type: dirifempty
;User temp directories
Name: {%tmp}\{#Appname}; Type: filesandordirs
Name: {%tmp}\PDFCreatorSpool; Type: filesandordirs

[Messages]
;Remove the 'StatusRunProgram' message
StatusRunProgram=

[Languages]
#include "languages.inc"

[CustomMessages]
#include "custommessages.inc"

[Types]
Name: custom; Description: {cm:CustomInstallation}; Flags: iscustom
Name: full; Description: {cm:FullInstallation}
Name: compact; Description: {cm:CompactInstallation}

[Components]
Name: program; Description: {cm:Programfiles}; Types: full compact custom; Flags: fixed
#Ifdef IncludeGhostscript
Name: ghostscript; Description: {#GhostscriptLicense} Ghostscript {#GhostscriptVersion}; Types: full custom; Flags: fixed; Check: IsGhostscriptInstalled(true)
Name: ghostscript; Description: {#GhostscriptLicense} Ghostscript {#GhostscriptVersion}; Types: full custom; Check: IsGhostscriptInstalled(false)
#ENDIF

#IFDEF IncludeToolbar
Name: ietoolbar; Description: {cm:Toolbarfiles}; ExtraDiskSpaceRequired: 900000; Types: full custom; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,0; Check: IExplorerVersionLower55
Name: ietoolbar; Description: {cm:Toolbarfiles}; ExtraDiskSpaceRequired: 900000; Types: ; MinVersion: 4.1.1998,5.0.2195; OnlyBelowVersion: 0,0; Check: Not IExplorerVersionLower55; Flags: fixed
#ENDIF

Name: languages; Description: {cm:Languages}; Types: full custom

Name: languages\catalan; Description: Catalan; Types: full; Check: Not IsLanguage('catalan'); Flags: dontinheritcheck
Name: languages\catalan; Description: Catalan; Types: full custom; Check: IsLanguage('catalan'); Flags: dontinheritcheck

;Name: languages\chinese_simplified; Description: Chinese simplified; Types: full; Check: Not IsLanguage('chinese_simplified'); Flags: dontinheritcheck
;Name: languages\chinese_simplified; Description: Chinese simplified; Types: full custom; Check: IsLanguage('chinese_simplified'); Flags: dontinheritcheck

Name: languages\czech; Description: Czech; Types: full; Check: Not IsLanguage('czech'); Flags: dontinheritcheck
Name: languages\czech; Description: Czech; Types: full custom; Check: IsLanguage('czech'); Flags: dontinheritcheck

Name: languages\dutch; Description: Dutch; Types: full; Check: Not IsLanguage('dutch'); Flags: dontinheritcheck
Name: languages\dutch; Description: Dutch; Types: full custom; Check: IsLanguage('dutch'); Flags: dontinheritcheck

Name: languages\english; Description: English; Types: full compact custom; Flags: fixed dontinheritcheck

Name: languages\french; Description: French; Types: full; Check: Not IsLanguage('french'); Flags: dontinheritcheck
Name: languages\french; Description: French; Types: full custom; Check: IsLanguage('french'); Flags: dontinheritcheck

Name: languages\german; Description: German; Types: full; Check: Not IsLanguage('german'); Flags: dontinheritcheck
Name: languages\german; Description: German; Types: full custom; Check: IsLanguage('german'); Flags: dontinheritcheck

Name: languages\hungarian; Description: Hungarian; Types: full; Check: Not IsLanguage('hungarian'); Flags: dontinheritcheck
Name: languages\hungarian; Description: Hungarian; Types: full custom; Check: IsLanguage('hungarian'); Flags: dontinheritcheck

;Name: languages\indonesian; Description: Indonesian; Types: full; Check: Not IsLanguage('indonesian'); Flags: dontinheritcheck
;Name: languages\indonesian; Description: Indonesian; Types: full custom; Check: IsLanguage('indonesian'); Flags: dontinheritcheck

Name: languages\italian; Description: Italian; Types: full; Check: Not IsLanguage('italian'); Flags: dontinheritcheck
Name: languages\italian; Description: Italian; Types: full custom; Check: IsLanguage('italian'); Flags: dontinheritcheck

Name: languages\lithuanian; Description: Lithuanian; Types: full; Check: Not IsLanguage('lithuanian'); Flags: dontinheritcheck
Name: languages\lithuanian; Description: Lithuanian; Types: full custom; Check: IsLanguage('lithuanian'); Flags: dontinheritcheck

Name: languages\polish; Description: Polish; Types: full; Check: Not IsLanguage('polish'); Flags: dontinheritcheck
Name: languages\polish; Description: Polish; Types: full custom; Check: IsLanguage('polish'); Flags: dontinheritcheck

Name: languages\romanian; Description: Romanian; Types: full; Check: Not IsLanguage('romanian'); Flags: dontinheritcheck
Name: languages\romanian; Description: Romanian; Types: full custom; Check: IsLanguage('romanian'); Flags: dontinheritcheck

Name: languages\russian; Description: Russian; Types: full; Check: Not IsLanguage('russian'); Flags: dontinheritcheck
Name: languages\russian; Description: Russian; Types: full custom; Check: IsLanguage('russian'); Flags: dontinheritcheck

Name: languages\slovak; Description: Slovak; Types: full; Check: Not IsLanguage('slovak'); Flags: dontinheritcheck
Name: languages\slovak; Description: Slovak; Types: full custom; Check: IsLanguage('slovak'); Flags: dontinheritcheck

Name: languages\slovenian; Description: Slovenian; Types: full; Check: Not IsLanguage('slovenian'); Flags: dontinheritcheck
Name: languages\slovenian; Description: Slovenian; Types: full custom; Check: IsLanguage('slovenian'); Flags: dontinheritcheck

Name: languages\spanish; Description: Spanish; Types: full; Check: Not IsLanguage('spanish'); Flags: dontinheritcheck
Name: languages\spanish; Description: Spanish; Types: full custom; Check: IsLanguage('spanish'); Flags: dontinheritcheck

Name: languages\swedish; Description: Swedish; Types: full; Check: Not IsLanguage('swedish'); Flags: dontinheritcheck
Name: languages\swedish; Description: Swedish; Types: full custom; Check: IsLanguage('swedish'); Flags: dontinheritcheck

Name: languages\turkish; Description: Turkish; Types: full; Check: Not IsLanguage('turkish'); Flags: dontinheritcheck
Name: languages\turkish; Description: Turkish; Types: full custom; Check: IsLanguage('turkish'); Flags: dontinheritcheck

Name: languages\valencian; Description: Valencian; Types: full; Flags: dontinheritcheck

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Check: UseDesktopiconCommon
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked; Check: Not UseDesktopiconCommon
Name: desktopicon\common; Description: {cm:ForAllUser}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive; Check: UseDesktopiconCommon
Name: desktopicon\common; Description: {cm:ForAllUser}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive unchecked; Check: Not UseDesktopiconCommon
Name: desktopicon\user; Description: {cm:ForTheCurrentUserOnly}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive; Check: UseDesktopiconUser
Name: desktopicon\user; Description: {cm:ForTheCurrentUserOnly}; GroupDescription: {cm:AdditionalIcons}; Flags: exclusive unchecked; Check: Not UseDesktopiconUser
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Check: IExplorerVersionGreater3 And UseQuickLaunchIcon
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked; Check: IExplorerVersionGreater3 And Not UseQuickLaunchIcon
Name: fileassoc; Description: {cm:AssocFileExtension,PDFCreator,.ps}; GroupDescription: {cm:OtherTasks}; Check: UseFileAssoc
Name: fileassoc; Description: {cm:AssocFileExtension,PDFCreator,.ps}; GroupDescription: {cm:OtherTasks}; Flags: unchecked; Check: Not UseFileAssoc
Name: winexplorer; Description: {cm:WinexplorerEntry}; GroupDescription: {cm:OtherTasks}; Check: UseWinExplorer
Name: winexplorer; Description: {cm:WinexplorerEntry}; GroupDescription: {cm:OtherTasks}; Flags: unchecked; Check: Not UseWinExplorer

[Code]
const
 SIZE_OF_MONITORINFO1 = $4;
 SIZE_OF_PORTINFO2 = $14;
 SIZE_OF_PRINTERINFO2 = $54;
 SIZE_OF_DRIVERINFO3 = $28;
 PRINTER_ENUM_LOCAL = $2;

 STANDARD_RIGHTS_REQUIRED = $F0000;
 PRINTER_ACCESS_ADMINISTER = $4;
 PRINTER_ACCESS_USE = $8;
 PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE);

 SC_MANAGER_ALL_ACCESS = $f003f;
 SERVICE_QUERY_STATUS  = $4;
 SERVICE_RUNNING       = $4;

 ERROR_SUCCESS = 0;
 ERROR_MORE_DATA = 234;
 STANDARD_RIGHTS_ALL = $1F0000;
 KEY_QUERY_VALUE = $1;
 KEY_SET_VALUE = $2;
 KEY_CREATE_SUB_KEY = $4;
 KEY_ENUMERATE_SUB_KEYS = $8;
 KEY_NOTIFY = $10;
 KEY_CREATE_LINK = $20;
 SYNCHRONIZE = $100000;
 KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE));

type
 TAInt = Array of Integer; TAStr = Array of String;
 TPrinterDefaults = record
  pDatatype : LongInt;
  pDevMode : LongInt;
  DesiredAccess : LongInt;
 end;
 TPortInfo2 = record
  pPortName: String;
  pMonitorName: String;
  pDescription: String;
  fPortType: LongInt;
  Reserved: LongInt;
 end;
 TMonitorInfo1 = record
  pName: String;
 end;
 TMonitorInfo2 = record
  pName : String;
  pEnvironment : String;
  pDLLName : String;
 end;
 TDriverInfo3 = record
  cVersion: LongInt;
  pName: String;
  pEnvironment: String;
  pDriverPath: String;
  pDataFile: String;
  pConfigFile: String;
  pHelpFile: String;
  pDependentFiles: String;
  pMonitorName: String;
  pDefaultDataType: String;
 end;
 TPrinterInfo2 = record
  pServerName : String;
  pPrinterName : String;
  pShareName : String;
  pPortName : String;
  pDriverName : String;
  pComment : String;
  pLocation : String;
  pDevMode : LongInt;
  pSepFile : String;
  pPrintProcessor : String;
  pDatatype : String;
  pParameters : String;
  pSecurityDescriptor : LongInt;
  Attributes : LongInt;
  Priority : LongInt;
  DefaultPriority : LongInt;
  StartTime : LongInt;
  UntilTime : LongInt;
  Status : LongInt;
  cJobs : LongInt;
  AveragePPM : LongInt;
 end;
 SERVICE_STATUS = record
  dwServiceType             : cardinal;
  dwCurrentState            : cardinal;
  dwControlsAccepted        : cardinal;
  dwWin32ExitCode           : cardinal;
  dwServiceSpecificExitCode : cardinal;
  dwCheckPoint              : cardinal;
  dwWaitHint                : cardinal;
 end;
 HANDLE = cardinal;

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

function GetEnvironmentStrings(): LongInt;
 external 'GetEnvironmentStringsA@kernel32.dll';
function FreeEnvironmentStrings(lpsz: LongInt): LongInt;
 external 'FreeEnvironmentStringsA@kernel32.dll';
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
function SearchPath(lpPath : String; lpFilename : String; lpExtension : String; nBufferLength : LongInt; lpBuffer : String; lpFilePart : LongInt) : LongInt;
 external 'SearchPathA@kernel32.dll stdcall';

function GetPrinterDriverDirectory(pName:String; pEnvironment:String; Level:LongInt; pDriverDirectory:String; cbBuf:LongInt; var pcbNeened:LongInt):Integer;
 external 'GetPrinterDriverDirectoryA@winspool.drv stdcall';
function GetPrinterDriverDirectory2(pName:String; pEnvironment:String; Level:LongInt; pDriverDirectory: LongInt; cbBuf:LongInt; var pcbNeened:LongInt):Integer;
 external 'GetPrinterDriverDirectoryA@winspool.drv stdcall';

function AddMonitor (pName:String; Level:LongInt; var pMonitors:TMonitorInfo2): LongInt;
 external 'AddMonitorA@winspool.drv stdcall';
function AddPort (pName:String; hwnd:LongInt; pPort:String): LongInt;
 external 'AddPortA@winspool.drv stdcall';
function AddPrinterDriver (pName : String; Level : LongInt; var pDriverInfo : TDriverInfo3) : LongInt;
 external 'AddPrinterDriverA@winspool.drv stdcall';
function AddPrinter(pName : String; Level: Longint; var pPrinter2: TPrinterInfo2): LongInt;
 external 'AddPrinterA@winspool.drv stdcall';

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

function OpenPrinter(pName : String; var phPrinter: Longint; pDefault: TPrinterDefaults): LongInt;
 external 'OpenPrinterA@winspool.drv stdcall';
function ClosePrinter(phPrinter: LongInt): LongInt;
 external 'ClosePrinter@winspool.drv stdcall';

function DeletePrinter(phPrinter: Longint): LongInt;
 external 'DeletePrinter@winspool.drv stdcall';
function DeletePrinterDriver(pName : String; pEnviroment: String; pDriverName: String): LongInt;
 external 'DeletePrinterDriverA@winspool.drv stdcall';
function DeletePort(pName : String; pHwnd: Longint; pPortName : String): LongInt;
 external 'DeletePortA@winspool.drv stdcall';
function DeleteMonitor(pName : String;  pEnviroment: String; pMonitorName : String): LongInt;
 external 'DeleteMonitorA@winspool.drv stdcall';

function OpenSCManager(lpMachineName, lpDatabaseName: string; dwDesiredAccess :cardinal): HANDLE;
 external 'OpenSCManagerA@advapi32.dll stdcall';
function OpenService(hSCManager :HANDLE;lpServiceName: string; dwDesiredAccess :cardinal): HANDLE;
 external 'OpenServiceA@advapi32.dll stdcall';
function CloseServiceHandle(hSCObject :HANDLE): boolean;
 external 'CloseServiceHandle@advapi32.dll stdcall';
function QueryServiceStatus(hService :HANDLE;var ServiceStatus :SERVICE_STATUS) : boolean;
 external 'QueryServiceStatus@advapi32.dll stdcall';

function RegOpenKeyEx(hKey :LongInt; lpValueName: String; ulOptions: LongInt; samDesired :LongInt;var phkResult: LongInt) :Longint;
 external 'RegOpenKeyExA@advapi32.dll';
function RegQueryValueEx(hKey :LongInt; lpValueName: String; lpReserved: LongInt;var lpType :LongInt;var lpData: LongInt; var lpcbData: LongInt) :Longint;
 external 'RegQueryValueExA@advapi32.dll';

var progTitel, progHandle: TArrayOfString;
    msg : TAStr;
    FullInstallation : boolean;
    Printername, Printerdrivername, Printerportname, Printermonitorname,
     LogFile, UninstallLogfile,
     PrintSystem, Win9x, WinNT, Win2000, WinXP, Win2003,
     WinXP2003_32bit, WinXP2003_64bit : String;
    AdditionalPrinterProgressSteps, AdditionalPrinterProgressIndex: LongInt;
    ProgressPage: TOutputProgressWizardPage;

    cmdlPrintername, cmdlPPDFile, cmdlREGFile, cmdlINIFile,
    cmdlSaveInfFile, cmdlLoadInfFile: String;
    cmdlSilent, cmdlVerysilent, cmdlForceInstall, cmdlUseINI: Boolean;

    desktopicon, desktopicon_common, desktopicon_user,
    quicklaunchicon, fileassoc, winexplorer: Boolean;

    SCPage:TWizardPage;
    PrinternamePage: TInputQueryWizardpage;
    PrinterdriverPage : TInputOptionWizardPage;
    Standardmodus: TRadioButton;
    ServerDescriptionPage: TOutputMsgWizardPage;
    PrinterInstallationSuccessfully: Boolean;

function IsX64: Boolean;
begin
 Result:=(ProcessorArchitecture=paX64);
end;

function GetDateString(Default:String):String;
begin
 result:=GetDateTimeString('yyyymmdd',#0,#0)
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

function GetPrinterTemppath(Default:string): String;
var
 TempDir: String;
begin
 If (InstallOnThisVersion('4.00.950,0','0,0')=irInstall) then     // Win9xMe
  TempDir:=ExpandConstant('{%tmp}');
 If InstallOnThisVersion('0,4.0.1381','5.0.2195,0')=irInstall then // WinNt
  TempDir:=ExpandConstant('{userappdata}')+ '\PDFCreator\' + ExpandConstant('{username}');
 If InstallOnThisVersion('0,5.0.2195','0,0')=irInstall then // Win2k and above
  TempDir:=ExpandConstant('{%tmp}')+ '\PDFCreator';
 If Length(TempDir) = 0 Then
  TempDir := ExpandConstant('{app}') + '\Temp';
 Result:=TempDir;
end;

function IsLanguage(LangName: String): Boolean;
begin
 If LowerCase(LangName)=Lowercase(ActiveLanguage) then
  Result:=True;
end;

procedure SetDummyRunOnce;
begin
 RegWriteStringValue(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce', 'PDFCreatorRestart', '');
end;

function UseINI(): Boolean;
begin
 Result:=cmdlUseINI;
end;

function UseOwnREGFile(): Boolean;
begin
 if Length(cmdlREGFile)>0 then
  Result:=True
 else
  Result:=False;
end;

function UseOwnINIFile(): Boolean;
begin
 if Length(cmdlINIFile)>0 then
  Result:=True
 else
  Result:=False;
end;

function UseOwnPPDFile(): Boolean;
begin
 if Length(cmdlPPDFile)>0 then
  Result:=True
 else
  Result:=False;
end;

function GetExternalREGFile(Default:string): String;
begin
 Result:=cmdlREGFile
end;

function GetExternalINIFile(Default:string): String;
begin
 Result:=cmdlINIFile
end;

function GetExternalPPDFile(Default:string): String;
begin
 Result:=cmdlPPDFile
end;

function GetEnvironment(): String;
var res: LongInt; tStr, resStr: String;
begin
 res := GetEnvironmentStrings;
 Repeat
  tStr := CastIntegerToString(res);
  OemToCharBuff(tStr);
  If Length(resStr) = 0 then
    resStr := tStr
   else
    resStr := resStr + #13#10 + tStr;
  res := res + Length(tStr) + 1;
 until Length(CastIntegerToString(res)) = 0;
 FreeEnvironmentStrings(res);
 Result := resStr;
end;

function InstallWin9xPrinterdriver(): Boolean;
begin
 Result:=False;
 If (InstallOnThisVersion('4.00.950,0','0,0')=irInstall) then
  Result:=True;
 If InstallOnThisVersion('0,4.0.1381','0,0')=irInstall then
  If PrinterdriverPage.Values[0] then
   Result:=True
end;

function InstallWinNtPrinterdriver(): Boolean;
begin
 Result:=False;
 If (InstallOnThisVersion('0,4.0.1381','0,5.0.2195')=irInstall) then
  Result:=True;
 If (InstallOnThisVersion('0,5.0.2195','0,0')=irInstall) and PrinterdriverPage.Values[1] then
  Result:=True
end;

function InstallWin2kXP2k3Printerdriver32bit(): Boolean;
begin
 Result:=False;
 If (InstallOnThisVersion('0,5.0.2195','0,0')=irInstall) and Not IsWin64 then
  Result:=True;
 If (InstallOnThisVersion('0,5.01.2600','0,0')=irInstall) and
  IsWin64 and PrinterdriverPage.Values[2] then
  Result:=True
end;

function InstallWinXP2k3Printerdriver64bit(): Boolean;
begin
 Result:=False;
 If (InstallOnThisVersion('0,5.01.2600','0,0')=irInstall) and IsWin64 then
  Result:=True
end;

function GetPrintermonitorname(Default:String): String;
var tStr:String;
begin
 tStr:=Trim(Printermonitorname);
 if Length(tStr)=0 then
  tStr:='{#DefaultPrinterMonitorname}';
 if Length(tStr)=0 then begin
  RaiseException('Error in setup: Empty printer monitorname!'#13#13+
   'The setup will be cancelled.');
 end;
 result:=tStr;
end;

function GetPrinterportname(Default:String): String;
var tStr:String;
begin
 tStr:=Trim(Printerportname);
 if Length(tStr)=0 then
  tStr:='{#DefaultPrinterPortname}';
 if Length(tStr)=0 then begin
  RaiseException('Error in setup: Empty printer portname!'#13#13+
   'The setup will be cancelled.');
 end;
 result:=tStr;
end;

function GetPrinterdrivername(Default:String): String;
var tStr:String;
begin
 tStr:=Trim(Printerdrivername);
 if Length(tStr)=0 then
  tStr:='{#DefaultPrinterDrivername}';
 if Length(tStr)=0 then begin
  RaiseException('Error in setup: Empty printer drivername!'#13#13+
   'The setup will be cancelled.');
 end;
 result:=tStr;
end;

function GetPrintername(Default:String): String;
var tStr:String;
begin
 tStr:=Trim(Printername);
 if Length(tStr)=0 then
  tStr:='{#DefaultPrintername}';
 if Length(tStr)=0 then begin
  RaiseException('Error in setup: Empty printername!'#13#13+
   'The setup will be cancelled.');
 end;
 result:=tStr;
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

function GetPorts(var Ports : Array of TPortInfo2) : LongInt;
var
 PORT_LEVEL, res, cbBuf, pcbNeeded, pcbReturned, i : LongInt;
 tArr: Array of TPortInfo2;
 tStr:String;
begin
 Setarraylength(tArr,0);
 cbBuf:=0; pcbNeeded:=0; pcbReturned:=0;
 PORT_LEVEL:=2;
 res:=EnumPorts('', PORT_LEVEL, tStr, cbBuf, pcbNeeded, pcbReturned);
 if pcbNeeded>0 then begin
  cbBuf:=pcbNeeded;
  tStr:=StringOfChar(#0,pcbNeeded);
  res:=EnumPorts('', PORT_LEVEL, tStr, cbBuf, pcbNeeded, pcbReturned);
  Setarraylength(tArr,pcbReturned);
  for i:=0 to pcbReturned-1 do begin
   tArr[i].pPortName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PORTINFO2));
   tArr[i].pMonitorName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PORTINFO2 + 1*4));
   tArr[i].pDescription:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PORTINFO2 + 2*4));
   tArr[i].fPorttype:=GetLongFromstring(tStr,1+i*SIZE_OF_PORTINFO2                   + 3*4);
   tArr[i].Reserved:=GetLongFromstring(tStr,1+i*SIZE_OF_PORTINFO2                    + 4*4);
  end;
 end;
 Ports:=tArr;
 result:=GetArrayLength(tArr);
end;

function GetMonitors(var Monitors : Array of TMonitorInfo1) : LongInt;
var
 MONITOR_LEVEL, res, cbBuf, pcbNeeded, pcbReturned, i : LongInt;
 tArr: Array of TMonitorInfo1;
 tStr:String;
begin
 Setarraylength(tArr,0);
 cbBuf:=0; pcbNeeded:=0; pcbReturned:=0;
 MONITOR_LEVEL:=1;
 res:=EnumMonitors('', MONITOR_LEVEL, tStr, cbBuf, pcbNeeded, pcbReturned)
 if pcbNeeded>0 then begin
  cbBuf:=pcbNeeded;
  tStr:=StringOfChar(#0,pcbNeeded);
 res:=EnumMonitors('', MONITOR_LEVEL, tStr, cbBuf, pcbNeeded, pcbReturned)
  Setarraylength(tArr,pcbReturned);
  for i:=0 to pcbReturned-1 do
   tArr[i].pName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_MONITORINFO1));
 end;
 Monitors:=tArr;
 result:=GetArrayLength(tArr);
end;

function GetPrinterDrivers(var PrinterDrivers : Array of TDriverInfo3; Environment: String) : LongInt;
var
 PRINTERDRIVER_LEVEL, res, cbBuf, pcbNeeded, pcbReturned, i : LongInt;
 tArr: Array of TDriverInfo3;
 tStr: String;
begin
 Setarraylength(tArr,0);
 cbBuf:=0; pcbNeeded:=0; pcbReturned:=0;
 PRINTERDRIVER_LEVEL:=3;
 res:=EnumPrinterdrivers('', Environment, PRINTERDRIVER_LEVEL, tStr, cbBuf, pcbNeeded, pcbReturned);
 if pcbNeeded>0 then begin
  cbBuf:=pcbNeeded;
  tStr:=StringOfChar(#0,pcbNeeded);
  res:=EnumPrinterdrivers('', Environment, PRINTERDRIVER_LEVEL, tStr, cbBuf, pcbNeeded, pcbReturned);
  Setarraylength(tArr,pcbReturned);
  for i:=0 to pcbReturned-1 do begin
   tArr[i].cVersion:=GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3);
   tArr[i].pName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3            + 1*4));
   tArr[i].pEnvironment:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3     + 2*4));
   tArr[i].pDriverPath:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3      + 3*4));
   tArr[i].pDataFile:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3        + 4*4));
   tArr[i].pConfigFile:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3      + 5*4));
   tArr[i].pHelpFile:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3        + 6*4));
   tArr[i].pDependentFiles:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3  + 7*4));
   tArr[i].pMonitorName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3     + 8*4));
   tArr[i].pDefaultDataType:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_DRIVERINFO3 + 9*4));
  end;
 end;
 PrinterDrivers:=tArr;
 result:=GetArrayLength(tArr);
end;

function GetPrinters(var Printers : Array of TPrinterInfo2) : LongInt;
var
 PRINTER_LEVEL, res, cbBuf, pcbNeeded, pcbReturned, i : LongInt;
 tArr: Array of TPrinterInfo2;
 tStr: String;
begin
 Setarraylength(tArr,0);
 cbBuf:=0; pcbNeeded:=0; pcbReturned:=0;
 PRINTER_LEVEL:=2;
 res:=EnumPrinters(PRINTER_ENUM_LOCAL, '', PRINTER_LEVEL, tStr, cbBuf, pcbNeeded, pcbReturned);
 if pcbNeeded>0 then begin
  cbBuf:=pcbNeeded;
  tStr:=StringOfChar(#0,pcbNeeded);
  res:=EnumPrinters(PRINTER_ENUM_LOCAL, '', PRINTER_LEVEL, tStr, cbBuf, pcbNeeded, pcbReturned);
  Setarraylength(tArr,pcbReturned);
  for i:=0 to pcbReturned-1 do begin
   tArr[i].pServername:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2));
   tArr[i].pPrinterName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2    +  1*4));
   tArr[i].pShareName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2      +  2*4));
   tArr[i].pPortName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2       +  3*4));
   tArr[i].pDriverName:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2     +  4*4));
   tArr[i].pComment:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2        +  5*4));
   tArr[i].pLocation:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2       +  6*4));
   tArr[i].pDevMode:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                       +  7*4);
   tArr[i].pSepFile:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2        +  8*4));
   tArr[i].pPrintProcessor:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2 +  9*4));
   tArr[i].pDatatype:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2       + 10*4));
   tArr[i].pParameters:=GetStrFromPtrA(GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2     + 11*4));
   tArr[i].pSecurityDescriptor:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2            + 12*4);
   tArr[i].Attributes:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                     + 13*4);
   tArr[i].Priority:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                       + 14*4);
   tArr[i].DefaultPriority:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                + 15*4);
   tArr[i].StartTime:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                      + 16*4);
   tArr[i].UntilTime:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                      + 17*4);
   tArr[i].Status:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                         + 18*4);
   tArr[i].cJobs:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                          + 19*4);
   tArr[i].AveragePPM:=GetLongFromstring(tStr,1+i*SIZE_OF_PRINTERINFO2                     + 20*4);
  end;
 end;
 Printers:=tArr;
 result:=GetArrayLength(tArr);
end;

function GetPDFCreatorPrinters(var PDFCreatorPrinters : Array of TPrinterInfo2) : LongInt;
var
 Printers: Array of TPrinterInfo2;
 SubKeys: TArrayOfString;
 i, j, cP, c: LongInt;
begin
 SetArrayLength(PDFCreatorPrinters, 0);
 Result:=0;
 cP:=GetPrinters(Printers);
 if RegGetSubkeyNames(HKLM, 'SYSTEM\CurrentControlSet\Control\Print\Monitors\PDFCreator\Ports', SubKeys) then
  begin
   c:=0;
   for i:=0 to cP-1 do
    for j:=0 to GetArrayLength(SubKeys)-1 do
     if Uppercase(SubKeys[j])=Uppercase(Printers[i].pPortName) then
      c:=c+1;
   if c>0 then begin
    SetArrayLength(PDFCreatorPrinters, c);
    c:=0;
    for i:=0 to cP-1 do
     for j:=0 to GetArrayLength(SubKeys)-1 do
      if Uppercase(SubKeys[j])=Uppercase(Printers[i].pPortName) then begin
       PDFCreatorPrinters[c]:=Printers[i];
       c:=c+1;
      end;
   end;
   Result:=c;
  end;
end;

function InstallMonitor(MonitorName: String):Boolean;
var M2:TMonitorInfo2; res:LongInt;
begin
 M2.pName:=MonitorName;
 If UsingWinNT then Begin
   If IsWin64 then
     M2.pEnvironment:='Windows x64'
    else
     M2.pEnvironment:='Windows NT x86';
   M2.pDLLName:='pdfcmnnt.dll'
  end else Begin
   M2.pEnvironment:='Windows 4.0';
   M2.pDLLName:='pdfcmn95.dll'
 end;

 SaveStringToFile(LogFile, 'InstallMonitor:' + #13#10, True)
 SaveStringToFile(LogFile, ' Monitorname : ' + M2.pName  + #13#10, True)
 SaveStringToFile(LogFile, ' Environment : ' + M2.pEnvironment  + #13#10, True)

 res := AddMonitor(Chr(0), 2, M2);
 if res=0 then begin
   Result:=False;
   SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
  end else begin
   Result:=True;
   SaveStringToFile(LogFile, ' Result: Success' + #13#10#13#10, True)
 end;
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(PrintSystem)); // Ini-Refresh !!! Important for Win9x/Me
end;

function InstallPort:Boolean;
var res, tres:Boolean; SubKeyName : String;
begin
 SaveStringToFile(LogFile, 'Install printerport:' + #13#10, True)
 SaveStringToFile(LogFile, ' Portname : ' + GetPrinterportname('')  + #13#10, True)
 SubKeyName:='{#PrintRegMon}'+GetPrintermonitorname('');
 SubKeyName:=SubKeyName+'\Ports\'+GetPrinterportname('');
 res:=true;
 tres:=RegWriteStringValue(HKLM,SubKeyName,'Arguments','-PPDFCREATORPRINTER');
 res:=res and tres;
 tres:=RegWriteStringValue(HKLM,SubKeyName,'Command',GetShortname(ExpandConstant('{app}')+'\{#SpoolerExename}'));
 res:=res and tres;
 tres:=RegWriteDWordValue(HKLM,SubKeyName,'Delay',300);
 res:=res and tres;
 tres:=RegWriteStringValue(HKLM,SubKeyName,'Description','PDFCreator Redirected Port');
 res:=res and tres;
 tres:=RegWriteDWordValue(HKLM,SubKeyName,'LogFileDebug',0);
 res:=res and tres;
 tres:=RegWriteDWordValue(HKLM,SubKeyName,'LogFileUse',0);
 res:=res and tres;
 tres:=RegWriteDWordValue(HKLM,SubKeyName,'Output',0);
 res:=res and tres;
 tres:=RegWriteStringValue(HKLM,SubKeyName,'Printer',GetPrintername(''));
 res:=res and tres;
 tres:=RegWriteDWordValue(HKLM,SubKeyName,'PrintError',0);
 res:=res and tres;
 tres:=RegWriteDWordValue(HKLM,SubKeyName,'RunUser',0);
 res:=res and tres;
 tres:=RegWriteDWordValue(HKLM,SubKeyName,'ShowWindow',0);
 res:=res and tres;
 if res=false then begin
   SaveStringToFile(LogFile, ' Result: Error ' + #13#10#13#10, True)
  end else
   SaveStringToFile(LogFile, ' Result: Success' + #13#10#13#10, True);
 Result:=res;
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(PrintSystem)); // Ini-Refresh !!! Important for Win9x/Me
end;

function InstallDriver:Boolean;
var DI3:TDriverInfo3; res:LongInt;
begin
 Result:=True;
 DI3.pName :=GetPrinterdrivername('');
 DI3.pDependentFiles :='';
// Win9x
 If InstallWin9xPrinterdriver then begin
  ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'),Win9x);
  AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
  ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);
  DI3.cVersion:=0;
  DI3.pDependentFiles :='ADOBEPS4.HLP'#0 + 'ICONLIB.DLL'#0 + 'PSMON.DLL'#0 + 'ADFONTS.MFM'#0 + 'ADOBEPS4.HLP'#0 + 'ADOBEPS4.DRV'#0 + 'ADIST5.PPD'#0#0;
  DI3.pConfigFile :='ADOBEPS4.DRV';
  DI3.pDriverPath := 'ADOBEPS4.DRV';
  DI3.pEnvironment:='Windows 4.0';
  DI3.pHelpFile :='ADOBEPS4.HLP';
  DI3.pDataFile :='ADIST5.PPD';
  DI3.cVersion := 3474436;
  DI3.pDefaultDataType :='RAW';
  DI3.pMonitorName :='';

  SaveStringToFile(LogFile, 'Install printerdriver for Win95/98/Me:' + #13#10, True)
  SaveStringToFile(LogFile, ' Drivername : ' + DI3.pName  + #13#10, True)
  SaveStringToFile(LogFile, ' Environment : ' + DI3.pEnvironment  + #13#10, True)

  res := AddPrinterDriver(Chr(0), 3, DI3);

  if res=0 then begin
    Result:=False;
    SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
   end else
    SaveStringToFile(LogFile, ' Result: Success' + #13#10#13#10, True);
  If UsingWinNT=false then
   SendMessage(65535, 26, 0, CastStringToInteger(PrintSystem)); // Ini-Refresh !!! Important for Win9x/Me
 end;
// WinNt 4.0
 If InstallWinNtPrinterdriver then begin
  ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'), WinNt);
  AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
  ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);
  DI3.cVersion:=2;
  DI3.pDependentFiles :='PDFCREAT.PPD'#0 + 'ADOBEPS5.DLL'#0 + 'ADOBEPSU.DLL'#0 + 'ADOBEPS5.NTF'#0 + 'ADOBEPSU.HLP'#0#0;
  DI3.pConfigFile :='ADOBEPSU.DLL';
  DI3.pDriverPath := 'ADOBEPS5.DLL';
  DI3.pEnvironment:='Windows NT x86';
  DI3.pHelpFile :='ADOBEPSU.HLP';
  DI3.pDataFile :='PDFCREAT.PPD';
  DI3.pDefaultDataType :='RAW';
  DI3.pMonitorName :='';

  SaveStringToFile(LogFile, 'Install printerdriver for WinNt:' + #13#10, True)
  SaveStringToFile(LogFile, ' Drivername : ' + DI3.pName  + #13#10, True)
  SaveStringToFile(LogFile, ' Environment : ' + DI3.pEnvironment  + #13#10, True)

  res := AddPrinterDriver(Chr(0), 3, DI3);

  if res=0 then begin
    Result:=False;
    SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
   end else
    SaveStringToFile(LogFile, ' Result: Success' + #13#10#13#10, True);
 end;
// Win2000, WinXP, Win2003 - 32bit
 If InstallWin2kXP2k3Printerdriver32bit then begin
  If InstallOnThisVersion('0,5.0.2195','0,5.01.2600')=irInstall then
    ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'),Win2000)
   else If InstallOnThisVersion('0,5.0.2600','0,5.02.3790')=irInstall then
     If IsWin64 then
       ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'),WinXP2003_32bit)
      else
       ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'),WinXP)
    else
     If IsWin64 then
       ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'),WinXP2003_32bit)
      else
       ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'),Win2003);
  AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
  ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);
  DI3.cVersion:=3;
  If InstallOnThisVersion('0,5.0.2195','0,5.01.2600')=irInstall then
    DI3.pDependentFiles :='PSCRPTFE.NTF'#0+'PSCRIPT.NTF'#0#0
   else
    DI3.pDependentFiles :='PSCRIPT.NTF'#0#0;
  DI3.pConfigFile :='PS5UI.DLL';
  DI3.pDriverPath := 'PSCRIPT5.DLL';
  DI3.pEnvironment:='Windows NT x86';
  DI3.pHelpFile :='PSCRIPT.HLP';
  DI3.pDataFile :='PDFCREAT.PPD';
  DI3.pDefaultDataType :='RAW';
  DI3.pMonitorName :='';

  SaveStringToFile(LogFile, 'Install printerdriver for Win2kXP2k3 (32bit):' + #13#10, True)
  SaveStringToFile(LogFile, ' Drivername : ' + DI3.pName  + #13#10, True)
  SaveStringToFile(LogFile, ' Environment : ' + DI3.pEnvironment  + #13#10, True)

  res := AddPrinterDriver(Chr(0), 3, DI3);

  if res=0 then begin
    Result:=False;
    SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
   end else
    SaveStringToFile(LogFile, ' Result: Success' + #13#10#13#10, True);
 end;
// WinXP, Win2003 - 64bit
 If InstallWinXP2k3Printerdriver64bit then begin
  ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'),WinXP2003_64bit);
  AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
  ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);
  DI3.cVersion:=3;
  DI3.pDependentFiles :='PSCRIPT.NTF'#0#0;
  DI3.pConfigFile :='PS5UI.DLL';
  DI3.pDriverPath := 'PSCRIPT5.DLL';
  DI3.pEnvironment:='Windows x64';
  DI3.pHelpFile :='PSCRIPT.HLP';
  DI3.pDataFile :='PDFCREAT.PPD';
  DI3.pDefaultDataType :='RAW';
  DI3.pMonitorName :='';

  SaveStringToFile(LogFile, 'Install printerdriver for WinXP2k3:' + #13#10, True)
  SaveStringToFile(LogFile, ' Drivername : ' + DI3.pName  + #13#10, True)
  SaveStringToFile(LogFile, ' Environment : ' + DI3.pEnvironment  + #13#10, True)

  res := AddPrinterDriver(Chr(0), 3, DI3);

  if res=0 then begin
    Result:=False;
    SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
   end else
    SaveStringToFile(LogFile, ' Result: Success' + #13#10#13#10, True);
 end;
end;

function InstallPrinter:Boolean;
var
 P2: TPrinterInfo2; res: LongInt; Printers : Array of TPrinterInfo2; c:LongInt;
begin
 Result:=True;
 P2.pPrinterName := GetPrintername('');
 P2.pDriverName := GetPrinterdrivername('');
 P2.pPrintProcessor := 'WinPrint';
 P2.pPortName := GetPrinterportname('');
 P2.pComment := 'eDoc Printer';
 P2.pSharename:= GetPrintername('');
 P2.Priority:=1;
 P2.DefaultPriority:=1;
 P2.pDatatype:='RAW';

 c:=GetPrinters(Printers);
 If c=0 then
   P2.Attributes :=4 // Set as defaultprinter
  else
   P2.Attributes :=0;

 SaveStringToFile(LogFile, 'InstallPrinter:' + #13#10, True)
 SaveStringToFile(LogFile, ' Printername: ' + P2.pPrintername + #13#10, True)
 SaveStringToFile(LogFile, ' Drivername : ' + P2.pDrivername  + #13#10, True)
 SaveStringToFile(LogFile, ' Portname   : ' + P2.pPortname    + #13#10, True)

 res := AddPrinter('', 2, P2);

 if res<>0 then begin
   ClosePrinter(res);
   SaveStringToFile(LogFile, ' Result: Success' + #13#10, True)
   if c=0 then begin
    // Set as defaultprinter
    SetIniString('windows','device',GetPrintername('')+',PSCRIPT,'+ GetPrinterMonitorname(''),ExpandConstant('{win}')+'\win.ini')
   end
  end else begin
   Result:=False;
   SaveStringToFile(LogFile, ' Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10, True)
 end;
 If UsingWinNT=false then
  SendMessage(65535, 26, 0, CastStringToInteger(PrintSystem)); // Ini-Refresh !!! Important for Win9x/Me
end;

function DeleteWindowsPrinter(Printername:String; Logfile: String):LongInt;
var
 pd:TPrinterDefaults; res, pHandle:LongInt;
begin
 Result:=0;
 SaveStringToFile(LogFile, ' Uninstall printer:' + #13#10, True)
 SaveStringToFile(LogFile, '  Printername : ' + Printername + #13#10, True)
 pd.pDatatype := 0;
 pd.pDevMode := 0
 pd.DesiredAccess := PRINTER_ALL_ACCESS
 SaveStringToFile(LogFile, '  Open printer' + #13#10, True)
 res := OpenPrinter(Printername, pHandle, pd);
 If res <> 0 Then begin
   SaveStringToFile(LogFile, '   Result: Success' + #13#10, True);
   SaveStringToFile(LogFile, '  Delete printer' + #13#10, True)
   res := DeletePrinter(pHandle)
   If res <> 0 Then begin
     SaveStringToFile(LogFile,  '   Result: Success' + #13#10, True);
     SaveStringToFile(LogFile, '  Close printer' + #13#10, True)
     res := ClosePrinter(pHandle);
     if res <> 0 then
       SaveStringToFile(LogFile, '   Result: Success' + #13#10#13#10, True)
      else begin
       result:=1;
       SaveStringToFile(LogFile, '   Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
      end
    end else begin
     result:=1;
     SaveStringToFile(LogFile, '   Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
    end
  end else begin
   result:=1
   SaveStringToFile(LogFile, '   Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
 end;
end;

function GetPorts2(var Ports2 : TArrayofString; Monitor: String) : LongInt;
var
 SubKeys: TArrayOfString;
 c: LongInt;
begin
 SetArrayLength(Ports2, 0);
 Result:=0;
 if RegGetSubkeyNames(HKLM, 'SYSTEM\CurrentControlSet\Control\Print\Monitors\'+Monitor+'\Ports', SubKeys) then
  begin
   if GetArrayLength(SubKeys)>0 then
    Ports2:=SubKeys;
   Result:=c;
  end;
end;

function IsPrinterdriverInstalled(PrinterdriverName: String; Environment: String): Boolean;
var
 c, i: LongInt;
 PrinterDrivers: Array of TDriverInfo3;
begin
 c:=GetPrinterDrivers(PrinterDrivers, Environment);
 for i:=0 to c-1 do
  If Uppercase(PrinterDrivers[i].pName)=Uppercase(PrinterdriverName) then begin
   result:=true;
   exit
  end
end;

function OpenServiceManager() : HANDLE;
begin
 Result:=0;
 if UsingWinNT() = true then begin
  Result := OpenSCManager('','ServicesActive',SC_MANAGER_ALL_ACCESS);
 end
end;

function IsServiceRunning(ServiceName: string) : boolean;
var
 hSCM, hService	: HANDLE;
 Status	: SERVICE_STATUS;
begin
 hSCM := OpenServiceManager();
 Result := false;
 if hSCM <> 0 then begin
  hService := OpenService(hSCM,ServiceName,SERVICE_QUERY_STATUS);
  if hService <> 0 then begin
   if QueryServiceStatus(hService,Status) then
    Result :=(Status.dwCurrentState = SERVICE_RUNNING);
   CloseServiceHandle(hService)
  end
  CloseServiceHandle(hSCM)
 end
end;

procedure SaveSpoolerServiceInformation(LogFile : String);
begin
 if UsingWinNT then
  if IsServiceRunning('Spooler') then
    SaveStringToFile(LogFile, 'Spooler service: is running'#13#10, True)
   else
    SaveStringToFile(LogFile, 'Spooler service: is NOT running'#13#10, True)
end;

procedure UninstallCompletePrinter(PrinterMonitorname:String; PrinterPortname: String; PrinterDrivername: String; Printername:String; LogFile: String);
var
 res, resUI, c, i: LongInt;
 PDFCreatorPrinters: Array of TPrinterInfo2;
 Ports: TArrayofString; Environment: String;
begin
 SaveStringToFile(LogFile, #13#10, True)

 SaveSpoolerServiceInformation(LogFile);
 c:=GetPDFCreatorPrinters(PDFCreatorPrinters);
 For i:=0 to c-1 do
  resUI:=DeleteWindowsPrinter(PDFCreatorPrinters[i].pPrinterName, UninstallLogfile);

 SaveStringToFile(LogFile, ' Uninstall printer driver for Win95/98/Me:' + #13#10, True)
 SaveStringToFile(LogFile, '  Drivername : ' + PrinterDrivername + #13#10, True)
 Environment:='Windows 4.0';
 If IsPrinterdriverInstalled(PrinterdriverName, Environment) then begin
  res:=DeletePrinterDriver('',Environment, PrinterDrivername);
  if res=0 then begin
    resUI:=resUI+1;
    SaveStringToFile(LogFile, '  Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
   end else
    SaveStringToFile(LogFile, '  Result: Success' + #13#10#13#10, True);
 end;

 SaveStringToFile(LogFile, ' Uninstall printerdriver for WinNT/Win2000/WinXP/Win2003:' + #13#10, True)
 SaveStringToFile(LogFile, '  Drivername : ' + PrinterDrivername + #13#10, True)
 Environment:='Windows NT x86';
 If IsPrinterdriverInstalled(PrinterdriverName, Environment) then begin
  res:=DeletePrinterDriver('',Environment, PrinterDrivername);
  if res=0 then begin
    resUI:=resUI+1;
    SaveStringToFile(LogFile, '  Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
   end else
    SaveStringToFile(LogFile, '  Result: Success' + #13#10#13#10, True);
 end;

 SaveStringToFile(LogFile, ' Uninstall printer ports:' + #13#10, True)
 c:=GetPorts2(Ports, PrinterPortname);
 For i:=0 to c-1 do begin
  SaveStringToFile(LogFile, '  Portname : ' + Ports[i] + #13#10, True)
  res:=DeletePort('',0,Ports[i]);
  if res=0 then begin
    resUI:=resUI+1;
    SaveStringToFile(LogFile, '  Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
   end else
    SaveStringToFile(LogFile, '  Result: Success' + #13#10#13#10, True);
 end;

 SaveStringToFile(LogFile, ' Uninstall printer monitor:' + #13#10, True)
 SaveStringToFile(LogFile, '  Monitorname : ' + PrinterMonitorname + #13#10, True)
 res:=DeleteMonitor('','',PrinterMonitorname);
 if res=0 then begin
   resUI:=resUI+1;
   SaveStringToFile(LogFile, '  Result: Error ' + IntToStr(GetLastError()) + ' = ' + SysErrorMessage(GetLastError()) + #13#10#13#10, True)
  end else
   SaveStringToFile(LogFile, '  Result: Success' + #13#10#13#10, True);
 if resUI>0 then begin
   SetDummyRunOnce;
   SaveStringToFile(LogFile, 'Need restart: True' + #13#10, True)
  end else
   SaveStringToFile(LogFile, 'Need restart: False' + #13#10#, True);
end;

function IsPrinterInstallationSuccessfully:Boolean;
begin
 Result:=PrinterInstallationSuccessfully
end;

function GetDefaultIniPath(Default:String):String;
var
 AppData: String;
begin
 if Not RegQueryStringValue(HKEY_USERS, '.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',
     'AppData', AppData) then
   AppData:=ExpandConstant('{commonappdata}');
 If Not DirExists(Appdata) Then
  ForceDirectories(AppData);
 Result:=AppData
end;

function GetIniPath(Default:String):String;
begin
 if Standardmodus.Checked = True then
   Result:=ExpandConstant('{userappdata}')+'\PDFCreator'
  else
   Result:=ExpandConstant('{commonappdata}');
end;

function IsServerInstallation: Boolean;
begin
 if Standardmodus.Checked = True then
   Result:=false
  else
   Result:=true;
end;

procedure IntegrateWinexplorer;
 var res: Boolean; keys: TArrayofString;i,c :LongInt;s1,s2,s3:String;
begin
 s1:=ExpandConstant('{cm:WinexplorerEntry}');
 StringChange(s1,'&','');
 ProgressPage.Caption:=s1;
 ProgressPage.Description:='';
 ProgressPage.SetText(s1,'');
 ProgressPage.SetProgress(0, 0);
 s3:=ExpandConstant('{cm:WinexplorerEntryCreate}');
 StringChange(s3,'%1','{#Appname}');
 res:=RegGetSubkeyNames(HKEY_CLASSES_ROOT,'',keys);
 ProgressPage.SetProgress(i, c-1);
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
           RegWriteStringValue(HKEY_CLASSES_ROOT,s1+'\shell\'+'{#UninstallID}','',s3);
           RegWriteStringValue(HKEY_CLASSES_ROOT,s1+'\shell\'+'{#UninstallID}'+'\command','',ExpandConstant('{app}')+'\pdfcreator.exe -NOSTART -PF'#34#37+'1'+#34);
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

function IsGhostscriptInstalled(InvertResult : Boolean):Boolean;
var
 subKeys:TArrayOfString; i:LongInt; rootKey, gsdll:String;
begin
 if InvertResult then
   result:=true
  else
   result:=false;
 rootKey:='SOFTWARE\AFPL Ghostscript';
 if RegKeyExists(HKLM,rootKey) then
  if RegGetSubkeyNames(HKLM, rootKey, subKeys) then
    for i:=0 to GetArrayLength(subKeys)-1 do
     if RegQueryStringValue(HKLM, rootKey + '\' + subKeys[i], 'GS_DLL',gsdll) then
      if FileExists(gsdll) then begin
       if InvertResult then
         result:=false
        else
         result:=true;
       exit
      end
 rootKey:='SOFTWARE\GNU Ghostscript';
 if RegKeyExists(HKLM,rootKey) then
  if RegGetSubkeyNames(HKLM, rootKey, subKeys) then
    for i:=0 to GetArrayLength(subKeys)-1 do
     if RegQueryStringValue(HKLM, rootKey + '\' + subKeys[i], 'GS_DLL',gsdll) then
      if FileExists(gsdll) then begin
       if InvertResult then
         result:=false
        else
         result:=true;
       exit
      end
 rootKey:='SOFTWARE\GPL Ghostscript';
 if RegKeyExists(HKLM,rootKey) then
  if RegGetSubkeyNames(HKLM, rootKey, subKeys) then
    for i:=0 to GetArrayLength(subKeys)-1 do
     if RegQueryStringValue(HKLM, rootKey + '\' + subKeys[i], 'GS_DLL',gsdll) then
      if FileExists(gsdll) then begin
       if InvertResult then
         result:=false
        else
         result:=true;
       exit
      end
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

function PrinterDriverDirectory(WinEnvironment:String):String;
var sb: LongInt;
	PrDrvDir : String;
	res: Integer;
begin
 res:=GetPrinterDriverDirectory('',WinEnvironment, 1, '', 0, sb);
 PrDrvDir := StringOfChar(' ', sb+1 );
 res:=GetPrinterDriverDirectory('',WinEnvironment, 1, PrDrvDir, sb, sb) ;
 if res=0 then
   PrDrvDir:= ''
  else begin
   PrDrvDir:= CastIntegerToString(CastStringToInteger(PrDrvDir));
 end;
 Result:=PrDrvDir;
end;

procedure PrinterDriverDirectoryLog(WinEnvironment:String);
var sb: LongInt;
	PrDrvDir : String;
	res: Integer;
begin
 res:=GetPrinterDriverDirectory(chr(0),WinEnvironment, 1,chr(0), 0, sb);
 PrDrvDir := StringOfChar(' ', sb+1 );
 SaveStringToFile(Logfile, 'Printerdriver-Directory (Environment: '+WinEnvironment+'):'+#13#10, True)
 res:=GetPrinterDriverDirectory(chr(0),WinEnvironment, 1, PrDrvDir, sb, sb) ;
 if res=0 then begin
   SaveStringToFile(LogFile, ' Result: Error '+IntToStr(GetLastError())+' = '+SysErrorMessage(GetLastError())+#13#10#13#10, True);
  end else begin
   PrDrvDir:= CastIntegerToString(CastStringToInteger(PrDrvDir));
   SaveStringToFile(LogFile, ' Result: Success = '+PrDrvDir+#13#10#13#10, True);
  end;
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
var fdPath, fdName : String;
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
 if InstallOnThisVersion('0,5.0.2195','0,0')=irInstall then begin
  fdName:='framedyn.dll';
  fdPath:=ExpandConstant('{sys}')+'\Wbem\'+ fdName;
  if fileExists(fdPath) then
    SaveStringToFile(LogFile, fdPath + ': found'+#13#10, True)
   else
    SaveStringToFile(LogFile, fdPath + ': NOT found'+#13#10, True);
  if FileInPath('framedyn.dll','') then
    SaveStringToFile(LogFile, fdName + ': found in path'+#13#10, True)
   else
    SaveStringToFile(LogFile, fdName + ': found NOT in path'+#13#10, True)
 end;
 SaveStringToFile(LogFile, 'Environment:'#13#10 + GetEnvironment, True)
end;

procedure SavePrinterdriverInformations(Environment :String);
var
 i,c:Longint;
 PrinterDrivers: Array of TDriverInfo3;
begin
 c:=GetPrinterdrivers(PrinterDrivers, Environment);
 SaveStringToFile(LogFile, 'Printerdrivers (' + Environment + ') ['+IntToStr(c)+']:'#13#10, True);
 for i:=0 to c-1 do
  SaveStringToFile(LogFile,' '+PrinterDrivers[i].pName+#13#10, True);
 SaveStringToFile(LogFile, #13#10, True);
end;

procedure SavePrinterInformations;
var
 i,c:Longint;
 Monitors: Array of TMonitorInfo1;
 Ports: Array of TPortInfo2;
 Printers: Array of TPrinterInfo2;
begin
 SaveSpoolerServiceInformation(LogFile);
 c:=GetMonitors(Monitors);
 SaveStringToFile(LogFile, 'Printermonitors ['+IntToStr(c)+']:'#13#10, True);
 for i:=0 to c-1 do
  SaveStringToFile(LogFile,' '+Monitors[i].pName+#13#10, True);
 SaveStringToFile(LogFile, #13#10, True);

 c:=GetPorts(Ports);
 SaveStringToFile(LogFile, 'Printerports ['+IntToStr(c)+']:'#13#10, True);
 for i:=0 to c-1 do
  SaveStringToFile(LogFile,' '+Ports[i].pPortname+#13#10, True);
 SaveStringToFile(LogFile, #13#10, True);

 SavePrinterdriverInformations('Windows 4.0');
 SavePrinterdriverInformations('Windows NT x86');
 SavePrinterdriverInformations('Windows x64');
 SavePrinterdriverInformations('Windows IA64');
 SavePrinterdriverInformations('Windows NT Alpha_AXP');

 c:=GetPrinters(Printers);
 SaveStringToFile(LogFile, 'Printers ['+IntToStr(c)+']:'#13#10, True);
 for i:=0 to c-1 do
  SaveStringToFile(LogFile,' '+Printers[i].pPrinterName+#13#10, True);
 SaveStringToFile(LogFile, #13#10, True);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
 if InstallOnThisVersion('4.00.950,0','0,0')=irInstall then
   if (PageID = PrinterdriverPage.ID) or (PageID = ServerDescriptionPage.ID) or
    (PageID = SCPage.ID) then
     Result := True
    else
     Result := False
  else
   if Standardmodus.Checked = True then
     if (PageID = PrinterdriverPage.ID) or (PageID = ServerDescriptionPage.ID) then
       Result := True
      else
       Result := False
    else
     Result:=False
end;

function CheckMonitorname(MonitornameStr: String): Boolean;
var
 Monitors: Array of TMonitorInfo1; c, i: LongInt;
begin
 Result:=False;
 if Length(MonitornameStr)=0 then exit;
 c:=GetMonitors(Monitors);
 for i:=0 to c-1 do
  If Uppercase(Monitors[i].pName)=Uppercase(MonitornameStr) then begin
   Result:=True;
   exit
  end
end;

function CheckPrintername(PrinternameStr: String; ShowMsg: Boolean): Boolean;
var
 Printers: Array of TPrinterInfo2; c, i: LongInt;
begin
 Result:=False;
 if Length(PrinternameStr)=0 then begin
  If ShowMsg then
   MsgBox(ExpandConstant('{cm:FalsePrintername2}'),mbError,MB_OK);
  exit
 end;
 if Length(PrinternameStr)>221 then begin
  If ShowMsg then
   MsgBox(ExpandConstant('{cm:FalsePrintername3}'),mbError,MB_OK);
  exit
 end;
 if (Pos('!',PrinternameStr)>0)or(Pos('\',PrinternameStr)>0)or(Pos(',',PrinternameStr)>0) then begin
  If ShowMsg then
   MsgBox(ExpandConstant('{cm:FalsePrintername1}'),mbError,MB_OK);
  exit
 end;
 c:=GetPrinters(Printers);
 for i:=0 to c-1 do begin
  If Uppercase(Printers[i].pPrinterName)=Uppercase(PrinternameStr) then begin
   If ShowMsg then
    MsgBox(ExpandConstant('{cm:FalsePrintername4}'),mbError,MB_OK);
   exit
  end
 end
 Result:=True;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
 Result:=False;
 if CurPageID=wpWelcome then
  PrinternamePage.Values[0]:=Printername;
 if CurPageID=wpReady then begin
  GetActivePDFLoaders;
  KillActivePDFLoaders;
  LogFile:=ExpandConstant('{app}')+'\SetupLog.txt';
 end;
 if CurPageID=wpFinished then
  SaveInstallInformations;
 if CurPageID = PrinternamePage.ID then begin
  if CheckPrintername(PrinternamePage.Values[0],True)=False then begin
   PrinternamePage.Values[0]:='PDFCreator';
   exit;
  end;
  Printername := PrinternamePage.Values[0];
 end
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

function IsFullInstallation(): Boolean;
begin
 result:=FullInstallation;
end;

function GetEnv2(const EnvVar: String; const System: Boolean):String;
var Rootkey :Integer; SubKeyName, ResultStr : String;
begin
 If System=True Then Begin
   rootKey:=HKLM;
   SubKeyName:='SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
  End Else Begin
   rootKey:=HKCU;
   SubKeyName:='Environment'
 End;
 RegQueryStringValue(RootKey, SubKeyName, EnvVar, ResultStr);
 Result:=ResultStr;
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

function IsDirInSystemEnvironPath(const Directory :String):Boolean;
var
 Path :String; d:array of string; i:Integer;
begin
 Path:=Lowercase(GetEnv2('Path',True));
 Result:=False;
 If Copy(Directory,Length(Directory),1)='\' Then
  Directory:=Copy(Directory,1,Length(Directory)-1);
 If Length(Directory)>0 Then begin
  Split(Path,';',d);
  For i:=0 To GetArrayLength(d)-1 do begin
   If Length(d[i])>0 Then Begin
    If Copy(d[i],Length(d[i]),1)='\' Then
     d[i]:=Copy(d[i],1,Length(d[i])-1);
    If Lowercase(d[i])=Lowercase(Directory) Then begin
     Result:=True;
     Exit
    end
   end
  end
 end
end;

function RegistryValueType(hkey :LongInt; SubKeyName, ValueName :String) :LongInt;
var res, thkey, rType, Data, cbData :LongInt;
begin
 result := -1;
 SubKeyName := RemoveBackslash(SubKeyName);
 if RegKeyExists(hkey, SubKeyName) then
  if RegValueExists(hkey, SubKeyName, ValueName) then begin
   res := RegOpenKeyEx(hkey, SubKeyName, 0, KEY_ALL_ACCESS, thkey);
   if res = ERROR_SUCCESS then begin
    res:=RegQueryValueEx(thkey, ValueName , 0, rType, Data, cbData);
    If (res = ERROR_SUCCESS) or (res = ERROR_MORE_DATA) then
     result := rType;
   end
  end
end;

function IsPathSettingCorrupt(): Boolean;
begin
 if IsDirInSystemEnvironPath(GetEnv('Systemroot')+'\system32\Wbem') or
  IsDirInSystemEnvironPath('%systemroot%\system32\Wbem') then
     Result:=False
    else
     Result:=True;
end;

procedure RepairFalseSystemPathEnvironment;
var Rootkey :Integer; SubKeyName, ResultStr : String;
begin
 rootKey:=HKLM;
 SubKeyName:='SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
 RegQueryStringValue(RootKey, SubKeyName, 'Path', ResultStr);
 If Length(ResultStr) = 0 Then
   ResultStr:='%SystemRoot%\System32\Wbem'
  Else
   ResultStr:='%SystemRoot%\System32\Wbem;' + ResultStr;
 RegWriteExpandStringValue(RootKey, SubKeyName, 'Path', ResultStr);
end;

procedure RepairFalseTypeSystemPathEnvironment;
var Rootkey :Integer; SubKeyName, ResultStr : String;
begin
 rootKey:=HKLM;
 SubKeyName:='SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
 RegQueryStringValue(RootKey, SubKeyName, 'Path', ResultStr);
 RegWriteExpandStringValue(RootKey, SubKeyName, 'Path', ResultStr);
end;

function IsDummyRunOnce:Boolean;
begin
 Result:=RegValueExists(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce', 'PDFCreatorRestart')
end;

function CompletePath(Path: String): String;
begin
 if Copy(Path,Length(Path),1)<>'\' then
   result:=Path + '\'
  else
   result:=Path;
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
    '/ForceInstall'#9#9#9'- force the installation'#13#10 +
    '/Printername=<PrinterName>'#9'- set a different printername'#13#10 +
    '/PPDFile=<PPDFile>'#9#9'- use an own ppd-file'#13#10 +
    '/REGFile=<REGFile>'#9#9'- use an own registry-file'#13#10 +
    '/UseINI'#9#9#9#9'- use an ini-file instead of registry settings'#13#10 +
    '/INIFile=<INIFile>'#9#9#9'- use an own ini-file'
    ,mbInformation,MB_OK);
   exit;
  end;

  if uppercase(paramstr(i))='/VERYSILENT' then
   cmdlVerySilent:=true;
  if uppercase(paramstr(i))='/SILENT' then
   cmdlSilent:=true;
  if uppercase(paramstr(i))='/FORCEINSTALL' then
   cmdlForceInstall:=true;
  if uppercase(paramstr(i))='/USEINI' then
   cmdlUseINI:=true;

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

  cmdParam:='/REGFile';
  pStr:=Copy(paramstr(i),1,Length(cmdParam));
  if uppercase(pstr)=uppercase(cmdParam) then begin
   if Copy(paramstr(i),Length(cmdParam)+1,1)='=' then
     cmdlREGFile:=Copy(paramstr(i),Length(cmdParam)+2,Length(paramstr(i)))
    else
     cmdlREGFile:=Copy(paramstr(i),Length(cmdParam)+1,Length(paramstr(i)));
  end;
  cmdParam:='/INIFile';
  pStr:=Copy(paramstr(i),1,Length(cmdParam));
  if uppercase(pstr)=uppercase(cmdParam) then begin
   if Copy(paramstr(i),Length(cmdParam)+1,1)='=' then
     cmdlINIFile:=Copy(paramstr(i),Length(cmdParam)+2,Length(paramstr(i)))
    else
     cmdlINIFile:=Copy(paramstr(i),Length(cmdParam)+1,Length(paramstr(i)));
  end;
  cmdParam:='/PPDFile';
  pStr:=Copy(paramstr(i),1,Length(cmdParam));
  if uppercase(pstr)=uppercase(cmdParam) then begin
   if Copy(paramstr(i),Length(cmdParam)+1,1)='=' then
     cmdlPPDFile:=Copy(paramstr(i),Length(cmdParam)+2,Length(paramstr(i)))
    else
     cmdlPPDFile:=Copy(paramstr(i),Length(cmdParam)+1,Length(paramstr(i)));
  end;
  cmdParam:='/Printername';
  pStr:=Copy(paramstr(i),1,Length(cmdParam));
  if uppercase(pstr)=uppercase(cmdParam) then begin
   if Copy(paramstr(i),Length(cmdParam)+1,1)='=' then
     cmdlPrintername:=Copy(paramstr(i),Length(cmdParam)+2,Length(paramstr(i)))
    else
     cmdlPrintername:=Copy(paramstr(i),Length(cmdParam)+1,Length(paramstr(i)));
  end;
 end;
 If Length(cmdlPrintername)>0 then begin
  If CheckPrintername(cmdlPrintername, Not cmdlVerySilent)=False then begin
   Result:=False
   exit;
  end;
  Printername:=cmdlPrintername;
 end;

 if Length(cmdlPPDFile)>0 then
  if FileExists(cmdlPPDFile)=False then begin
   pStr:=SetupMessage(msgSourceDoesntExist);
   StringChange(pStr,'%1',cmdlPPDFile);
   if cmdlVerySilent=false then
    msgbox(pStr,mbCriticalError, MB_OK);
   Result:=False
   exit;
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

function UseFileAssoc: boolean;
begin
 result:=fileassoc;
end;

function UseWinExplorer: boolean;
begin
 result:=winexplorer;
end;

procedure LoadInf;
var tasks:string; atasks:TArrayOfString; i:LongInt;
begin
 tasks:='';
 desktopicon:=false;
 desktopicon_common:=false;
 desktopicon_user:=false;
 quicklaunchicon:=false;
 fileassoc:=false;
 winexplorer:=false;
 if IniKeyExists('Setup','Printername',cmdlLoadInfFile) then
  Printername:=GetIniString('Setup', 'Printername', Printername, cmdlLoadInfFile);
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
   'fileassoc':          fileassoc:=true;
   'winexplorer':        winexplorer:=true;
  end
end;

procedure SaveInf;
var res: boolean; tasks: String;
begin
 res:=SetIniString('Setup', 'Printername', Printername, cmdlSaveInfFile)
 if IsTaskSelected('desktopicon') then
  tasks:='desktopicon';
 if IsTaskSelected('desktopicon\common') then
  tasks:=tasks + ',desktopicon\common';
 if IsTaskSelected('desktopicon\user') then
  tasks:=tasks + ',desktopicon\user';
 if IsTaskSelected('quicklaunchicon') then
  tasks:=tasks + ',quicklaunchicon';
 if IsTaskSelected('fileassoc') then
  tasks:=tasks + ',fileassoc';
 if IsTaskSelected('winexplorer') then
  tasks:=tasks + ',winexplorer';
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
 Win9x:=   'Windows 95, Windows 98, Windows Me';
 WinNt:=   'Windows NT 4.0';
 Win2000:= 'Windows 2000';
 WinXP:=   'Windows XP';
 Win2003:= 'Windows 2003';
 WinXP2003_32bit:= 'Windows XP/2003 - 32bit';
 WinXP2003_64bit:= 'Windows XP/2003 - 64bit';
 PrinterMonitorname:= 'PDFCreator';
 PrinterPortname:=    'PDFCreator:';
 PrinterDrivername:=  'PDFCreator';
 Printername:=        'PDFCreator';

 desktopicon:=true;
 desktopicon_common:=true;
 winexplorer:=true;

 If AnalyzeCommandlineParameters=false then begin
  result:=false;
  exit
 end;

 If cmdlLoadInfFile<>'' then LoadInf;

 if not cmdlForceInstall then begin
  If IsDummyRunOnce then begin
   MsgBox(ExpandConstant('{cm:RestartError}'),mbError,MB_OK);
   Result:=False;
   Exit;
  end;
  If InstallOnThisVersion('0,5.01.2600','0,0')=irInstall then begin // XP and above
   If IsPathSettingCorrupt then begin
    If MsgBox(ExpandConstant('{cm:FalseSystemEnvironPath}'),mbCriticalError,MB_OKCANCEL or MB_SETFOREGROUND or MB_DEFBUTTON2)=IDOK then begin
     RepairFalseSystemPathEnvironment;
     SetDummyRunOnce
    end;
    Result:=False;
    Exit
   end
   if fileExists(ExpandConstant('{sys}')+'\Wbem\framedyn.dll') And Not FileInPath('framedyn.dll','') then begin
    a := RegistryValueType(HKLM, 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 'Path');
    if (a>=0) and (a<>2) then begin
     if MsgBox(ExpandConstant('{cm:FalseSystemEnvironPath}'),mbCriticalError,MB_OKCANCEL or MB_SETFOREGROUND or MB_DEFBUTTON2)=IDOK then begin
      RepairFalseTypeSystemPathEnvironment;
      SetDummyRunOnce
     end;
     Result:=False;
     Exit
    end
   end
  end
 end;

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

#ifdef UpdateIsPossible
 If ProgramIsInstalled And not cmdlForceInstall then begin
   FullInstallation:=false;
   cv:=CompareVBVersion(GetInstalledVersion,'{#AppVersion}');
   if cv=-1 then begin
    cv:=CompareVBVersion(GetInstalledVersion,'{#UpdateIsPossibleMinVersion}');
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
 If ProgramIsInstalled and not cmdlForceInstall then begin
   if cmdlVerySilent=false then begin
    msgbox(msg[4],mbInformation, MB_OK);
   end;
   Result:=false;
  end else
   If IsAdminLoggedOn then
     Result := True
    else begin
     Result:=true;
     if cmdlVerySilent=false then begin
      a:=MsgBox(msg[0], mbConfirmation, MB_YesNo);
      If a=IDYES then
        Result:=True
       else
        Result:=False;
     end;
    end;
#endif
end;

function CreateLabel(ALeft, ATop, AWidth, AHeight: Integer; ACaption: String; FontColor: LongInt; Page: TWizardPage):TLabel;
var
 tLbl: TLabel;
begin
 tLbl:=TLabel.Create(WizardForm);
 with tLbl do begin
  Autosize := False;
  Caption := ACaption;
  Font.Color := FontColor;
  Height:=AHeight;
  Left:=ALeft;
  Top:=ATop;
  Width:=AWidth;
  Wordwrap := True;
  Parent := Page.Surface;
 end;
 Result:=tLbl;
end;

function CreateRadioButton(ALeft, ATop, AWidth, AHeight: Integer; ACaption: String; AChecked: Boolean; Page: TWizardPage):TRadioButton;
var
 trb: TRadioButton;
begin
 trb:=TRadioButton.Create(WizardForm);
 with trb do begin
  Caption := ACaption;
  Checked := AChecked;
  Height:=AHeight;
  Left:=ALeft;
  Top:=ATop;
  Width:=AWidth;
  Parent := Page.Surface;
 end;
 Result:=trb;
end;

procedure InitializeWizard();
begin
#IFDEF IncludeToolbar
 If InstallOnThisVersion('4.1.1998,5.0.2195','0,0')=irInstall then // Not Win95, Not WinNT4
  ToolbarForm_CreatePage(wpSelectDir);
#ENDIF

 SCPage:=CreateCustomPage(wpLicense, ExpandConstant('{cm:InstallationType}'),
  ExpandConstant('{cm:InstallationTypeDescription}'));

 CreateLabel(0,10,450,15,ExpandConstant('{cm:InstallationTypeDescription2}'),clWindowText,SCPage);
 Standardmodus:=CreateRadioButton(16,45,200,15,ExpandConstant('{cm:StandardInstallation}'),True,SCPage);
 CreateLabel(35, 65,350,350,ExpandConstant('{cm:StandardInstallationDescription}'),clWindowText,SCPage);
 CreateRadioButton(16,130,200,15,ExpandConstant('{cm:ServerInstallation}'),False,SCPage);
 CreateLabel(35,150,350,350,ExpandConstant('{cm:ServerInstallationDescription}'),clWindowText,SCPage);
 ServerDescriptionPage:=CreateOutputMsgPage(SCPage.ID,ExpandConstant('{cm:ServerMode}'),
  ExpandConstant('{cm:ServerModeDescription}'),ExpandConstant('{cm:ServerModeMessage}'));
 PrinternamePage:=CreateInputQuerypage(ServerDescriptionPage.ID,
  ExpandConstant('{cm:Printername}'), ExpandConstant('{cm:PrinternameDescription}'),
  ExpandConstant('{cm:PrinternameMessage}'));
 PrinternamePage.Add(ExpandConstant('{cm:PrinternameValue}'), False);
 PrinternamePage.Values[0]:=Printername;

 PrinterdriverPage := CreateInputOptionPage(PrinternamePage.ID,
  ExpandConstant('{cm:AdditionalPrinterdriver}'),
  ExpandConstant('{cm:AdditionalPrinterdriverDescription}'),
  ExpandConstant('{cm:AdditionalPrinterdriverMessage}'), False, False);
 PrinterdriverPage.Add('Windows 95, Windows 98, Windows Me');
 if InstallOnThisVersion('0,5.0.2195','0,0')=irInstall then
  PrinterdriverPage.Add('Windows NT 4.0');
 if IsWin64 then
  PrinterdriverPage.Add('Windows 2000/XP/2003 - 32bit');

 ProgressPage := CreateOutputProgressPage(ExpandConstant('{cm:InstallPrinter}'),
  ExpandConstant('{cm:InstallPrinterDescription}'));
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
  S: String; ShowAdditionPrinterdriversInMemo : Boolean;
begin
  S := MemoUserInfoInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoDirInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoTypeInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoComponentsInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  ShowAdditionPrinterdriversInMemo:=False;
  If InstallOnThisVersion('0,4.0.1381','0,0,5.0.2195')=irInstall then
   If PrinterdriverPage.Values[0] then
    ShowAdditionPrinterdriversInMemo:=True
  If InstallOnThisVersion('0,5.0.2195','0,0')=irInstall then
   If PrinterdriverPage.Values[0] Or PrinterdriverPage.Values[1] then
    ShowAdditionPrinterdriversInMemo:=True
  If ShowAdditionPrinterdriversInMemo Then begin
   S := S + ExpandConstant('{cm:AdditionalPrinterdriverCaption}');
   S := S + NewLine;
   If InstallOnThisVersion('0,4.0.1381','0,5.0.2195')=irInstall then
    If PrinterdriverPage.Values[0] then
     S := S + Space + Win9x + NewLine;
   If InstallOnThisVersion('0,5.0.2195','0,5.01.2600,0')=irInstall then begin
    If PrinterdriverPage.Values[0] then
     S := S + Space + Win9x + NewLine;
    If PrinterdriverPage.Values[1] then
     S := S + Space + WinNt + NewLine;
   end
   If (InstallOnThisVersion('0,5.01.2600','0,0')=irInstall) then begin
    If PrinterdriverPage.Values[0] then
     S := S + Space + Win9x + NewLine;
    If PrinterdriverPage.Values[1] then
     S := S + Space + WinNt + NewLine;
    If IsWin64 then
     If PrinterdriverPage.Values[2] then
      S := S + Space + WinXP2003_32bit + NewLine;
   end
   S := S + NewLine;
  end;
  S := S + MemoGroupInfo;
  if length(S)>0 then S := S + NewLine + NewLine;
  S := S + MemoTasksInfo;
  Result := S;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
 s : String;
 Ports: Array of TPortInfo2;
 PrinterDrivers: Array of TDriverInfo3;
 Printers: Array of TPrinterInfo2;
 Monitors : Array of TMonitorInfo1;
 res, tres: Boolean;
begin
  if CurStep = ssPostinstall then begin
   AdditionalPrinterProgressSteps:=5; AdditionalPrinterProgressIndex:=0;
   If InstallWin9xPrinterdriver then
    AdditionalPrinterProgressSteps:=AdditionalPrinterProgressSteps+1;
   If InstallWinNtPrinterdriver then
    AdditionalPrinterProgressSteps:=AdditionalPrinterProgressSteps+1;
   If InstallWin2kXP2k3Printerdriver32bit then
    AdditionalPrinterProgressSteps:=AdditionalPrinterProgressSteps+1;
   If InstallWinXP2k3Printerdriver64bit then
    AdditionalPrinterProgressSteps:=AdditionalPrinterProgressSteps+1;
   ProgressPage.SetProgress(0, 0);
   ProgressPage.Show;

   try
     PrintSystem:='windows';
     SaveStringToFile(LogFile, 'Printerstatus before installing:' + #13#10, True);
     SavePrinterInformations;
     PrinterDriverDirectoryLog('Windows 4.0');
     If UsingWinNT then
      PrinterDriverDirectoryLog('Windows NT x86');
     If IsWin64 then begin
      PrinterDriverDirectoryLog('Windows x64');
      PrinterDriverDirectoryLog('Windows IA64');
     end;
     res:=true;

     AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
     ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);
     ProgressPage.SetText(ExpandConstant('{cm:InstallPrintermonitor}'), GetPrinterMonitorname(''));
     GetPorts(Ports);

     s := GetPrintermonitorname('');
     if Not CheckMonitorname(s) then begin
      tres:=InstallMonitor(s);
      res:=res and tres;
     end else
      SaveStringToFile(LogFile, ' Monitorname : ' + s  + ' already exists.'#13#10, True);
     s :='';

     AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
     ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);

     ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterport}'), GetPrinterPortname(''));
     GetMonitors(Monitors);
     tres:=InstallPort;
     res:=res and tres;
     AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
     ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);

     ProgressPage.SetText(ExpandConstant('{cm:InstallPrinterdriver}'), GetPrinterDrivername(''));
     GetMonitors(Monitors);
     GetPorts(Ports);
     tres:=InstallDriver;
     res:=res and tres;
     AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
     ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);

     ProgressPage.SetText(ExpandConstant('{cm:InstallPrinter}'), GetPrintername(''));
     GetPrinterdrivers(PrinterDrivers,'Windows 4.0');
     GetPrinterdrivers(PrinterDrivers,'Windows NT x86');
     GetPrinterdrivers(PrinterDrivers,'Windows x64');
     tres:=InstallPrinter;
     res:=res and tres;
     AdditionalPrinterProgressIndex:=AdditionalPrinterProgressIndex+1;
     ProgressPage.SetProgress(AdditionalPrinterProgressIndex, AdditionalPrinterProgressSteps);
     If UsingWinNT=true then begin
      s:='SYSTEM\CurrentControlSet\Control\Print\Printers\'+GetPrintername('')+'\PrinterDriverData';
      If RegKeyExists(HKLM,s)=true then
       RegWriteDWordValue(HKLM,s,'FreeMem',32767);
     end
     GetPrinters(Printers);
     SaveStringToFile(LogFile, #13#10+'Printerstatus after installing:' + #13#10, True);
     SavePrinterInformations;
     s:=LowerCase(WizardSelectedTasks(false));
     if Pos('winexplorer',s)>0 then
      IntegrateWinexplorer;
     PrinterInstallationSuccessfully:=res;
     If cmdlSaveInfFile<>'' Then SaveInf;
    finally
      if res=false then
       MsgBox(ExpandConstant('{cm:PrinterInstallationFailed}'),mbError,MB_OK + MB_SETFOREGROUND);
      ProgressPage.Hide;
    end;
 end;
end;

function InitializeUninstall(): Boolean;
begin
 PrinterMonitorname:='PDFCreator';
 PrinterPortname:='PDFCreator:';
 PrinterDrivername:='PDFCreator';
 Printername:='PDFCreator';
 UninstallLogFile:=ExpandConstant('{%tmp}')+'\PDFCreatorUninstall.txt';
 SaveStringToFile(UninstallLogFile, 'Start uninstall:' + #13#10, False)
 Result:=True;
end;

procedure RemoveProgramSettings();
var
 iniPath:String;
begin
 iniPath:=ExpandConstant('{userappdata}')+'\PDFCreator';
 DelTree(iniPath,true,true,true);
 iniPath:=ExpandConstant('{app}')+'\PDFCreator.ini';
 DelTree(iniPath,false,true,false);
 RegDeleteKeyIncludingSubkeys(HKEY_USERS, '.DEFAULT\Software\PDFCreator');
 RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\PDFCreator');
 RegDeleteKeyIncludingSubkeys(HKEY_LOCAL_MACHINE, 'Software\PDFCreator');
end;

procedure RemoveExplorerIntegretation();
var
 keys: TArrayOfString; i :LongInt;tStr:String;
begin
 if RegGetSubkeyNames(HKEY_CLASSES_ROOT, '', keys) then begin
  for i:=0 to GetArrayLength(keys)-1 do begin
   tStr:=keys[i]+'\shell\'+'{#UninstallID}';
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
       SaveStringToFile(UninstallLogFile, ' Uninstall options:' + #13#10, True)
       if saveoptions=true then
         SaveStringToFile(UninstallLogFile, '  Saveoptions=True' + #13#10, True)
        else
         SaveStringToFile(UninstallLogFile, '  Saveoptions=False' + #13#10, True);
       if silent=true then
         SaveStringToFile(UninstallLogFile, '  Silent=True' + #13#10, True)
        else
         SaveStringToFile(UninstallLogFile, '  Silent=False' + #13#10, True);
       if verysilent=true then
         SaveStringToFile(UninstallLogFile, '  Verysilent=True' + #13#10, True)
        else
         SaveStringToFile(UninstallLogFile, '  Veryilent=False' + #13#10, True);
       if saveoptions=false then
        if (silent=false) and (verysilent=false) then
         if MsgBox(tStr, mbConfirmation, MB_YESNO) = IDYES then
          RemoveProgramSettings;
       RemoveExplorerIntegretation;
       UninstallCompletePrinter(PrinterMonitorname, PrinterPortname, PrinterDrivername, Printername, UninstallLogFile)
      end;
  end;
end;

//Only for debugging.
//#expr savetofile("PDFCreator-debug.ini")
