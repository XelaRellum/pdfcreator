; *** Inno Setup version 4.2.2+ Czech messages ***
;
; Copyright (c) 2005 Martin Kozák (martin.kozak@openoffice.cz)
;
; To download user-contributed translations of this file, go to:
;   http://www.jrsoftware.org/is3rdparty.php
;
; Note: When translating this text, do not add periods (.) to the end of
; messages that didn't have them already, because on those messages Inno
; Setup adds the periods automatically (appending a period would result in
; two periods being displayed).
;
; $jrsoftware: issrc/Files/Default.isl,v 1.58 2004/04/07 20:17:13 jr Exp $

[LangOptions]
LanguageName=Czech
LanguageID=$0405
LanguageCodePage=1250
; If the language you are translating to requires special font faces or
; sizes, uncomment any of the following entries and change them accordingly.
;DialogFontName=
;DialogFontSize=8
;WelcomeFontName=Verdana
;WelcomeFontSize=12
;TitleFontName=Arial
;TitleFontSize=29
;CopyrightFontName=Arial
;CopyrightFontSize=8

[Messages]

; *** Application titles
SetupAppTitle=Prùvodce instalací
SetupWindowTitle=Prùvodce instalací aplikace %1
UninstallAppTitle=Prùvodce odinstalací
UninstallAppFullTitle=Prùvodce odinstalací aplikace %1

; *** Misc. common
InformationTitle=Informace
ConfirmTitle=Potvrzení
ErrorTitle=Chyba

; *** SetupLdr messages
SetupLdrStartupMessage=Bude spuštìn prùvodce instalací aplikace %1. Pøejete si pokraèovat?
LdrCannotCreateTemp=Nebylo mono vytvoøit odkládací soubor. Prùvodce instalací bude ukonèen
LdrCannotExecTemp=Nebylo mono spustit soubor v odkládací sloce. Prùvodce instalací bude ukonèen

; *** Startup error messages
LastErrorMessage=%1.%n%nChyba %2: %3
SetupFileMissing=Soubor %1 nebyl v instalaèní sloce nalezen. Opravte prosím problém nebo pouijte jinou kopii aplikace
SetupFileCorrupt=Soubory prùvodce instalací jsou poškozeny. Pouijte prosím jinou kopii aplikace.
SetupFileCorruptOrWrongVer=Soubory prùvodce instalací jsou poškozeny nebo nejsou kompatibilní s touto verzí prùvodce. Opravte prosím problém nebou pouijte jinou kopii aplikace.
NotOnThisPlatform=Aplikace není urèena pro platformu %1.
OnlyOnThisPlatform=Aplikace je urèena pro platformu %1.
WinVersionTooLowError=Aplikace vyaduje verzi %2 systému %1 nebo vyšší.
WinVersionTooHighError=Aplikaci není moné ve verzi %2 systému %2 a vyšších vyuít.
AdminPrivilegesRequired=Pro instalaci aplikace musíte mít práva administrátora.
PowerUserPrivilegesRequired=Pro instalaci aplikace musíte mít práva administrátora nebo bıt èlenem skupiny Power Users.
SetupAppRunningError=Aplikace %1 je spuštìna.%n%nUzavøete prosím všechny instance a klepnìte na tlaèítko 'OK' pro pokraèování nebo 'Zrušit' pro ukonèení prùvodce instalací.
UninstallAppRunningError=Aplikace %1 je spuštìna.%n%nUzavøete prosím všechny instance a klepnìte na tlaèítko 'OK' pro pokraèování nebo 'Zrušit' pro ukonèení prùvodce instalací.

; *** Misc. errors
ErrorCreatingDir=Nebylo moné vytvoøit adresáø "%1"
ErrorTooManyFilesInDir=Nebylo moné vytvoøit soubor v adresáøi "%1". Adresáø obsahuje pøíliš mnoho souborù.

; *** Setup common messages
ExitSetupTitle=Ukonèení prùvodce instalací
ExitSetupMessage=Instalace aplikace nebyla dokonèena. Jestlie ukonèíte prùvodce instalací, aplikace nebude nainstalována.%n%nPro dokonèení instalace je moné prùvodce instalací spustit kdykoliv pozdìji.%n%nUkonèit prùvodce instalací?
AboutSetupMenuItem=&O prùvodci instalací...
AboutSetupTitle=O prùvodci instalací
AboutSetupMessage=%1 verze %2%n%3%n%n%1 internetová adresa:%n%4
AboutSetupNote=

; *** Buttons
ButtonBack=< &Zpìt
ButtonNext=&Další >
ButtonInstall=&Instalovat
ButtonOK=OK
ButtonCancel=Zrušit
ButtonYes=&Ano
ButtonYesToAll=Ano &všem
ButtonNo=&Ne
ButtonNoToAll=N&e všem
ButtonFinish=&Dokonèit
ButtonBrowse=&Procházet...
ButtonWizardBrowse=P&rocházet...
ButtonNewFolder=&Vytvoøit sloku

; *** "Select Language" dialog messages
SelectLanguageTitle=Zvolit jazyk prùvodce instalací
SelectLanguageLabel=Zvolte prosím jazyk prùvodce instalací:

; *** Common wizard text
ClickNext=Pokraèujte klepnutím na tlaèítko 'Další'. Klepnutím na tlaèítko 'Zrušit' prùvodce instalací ukonèíte.
BeveledLabel=
BrowseDialogTitle=Nalézt sloku
BrowseDialogLabel=Vyberte prosím sloku a klepnìte na tlaèítko 'OK'.
NewFolderName=Nová sloka

; *** "Welcome" wizard page
WelcomeLabel1=Vítejte v prùvodci instalací aplikace [name]
WelcomeLabel2=Prùvodce instalací nainstaluje na váš poèítaè aplikaci [name/ver].%n%nPøed pokraèováním instalace je doporuèeno uzavøít ostatní bìící aplikace.

; *** "Password" wizard page
WizardPassword=Heslo
PasswordLabel1=Instalaèní balíèek je chránìn heslem.
PasswordLabel3=Vlote prosím heslo a klepnìte na tlaèítko 'Další'. Ovìøovací proces rozlišuje malá a veká písmena.
PasswordEditLabel=&Heslo:
IncorrectPassword=Vloené heslo nesouhlasí. Opakujte prosím akci znovu.

; *** "License Agreement" wizard page
WizardLicense=Licenèní ujednání
LicenseLabel=Pøed pokraèováním vìnujte prosím pozornost následujícím dùleitım informacím.
LicenseLabel3=Vìnujte prosím pozornost následujícímu licenènímu ujednání. Podmínky tohoto ujednání je nutné pøed pokraèováním pøijmout.
LicenseAccepted=&Pøijmout licenèní ujednání
LicenseNotAccepted=&Odmítnout licenèní ujednání

; *** "Information" wizard pages
WizardInfoBefore=Informace
InfoBeforeLabel=Pøed pokraèováním vìnujte prosím pozornost následujícím dùleitım informacím.
InfoBeforeClickLabel=A bude pøipravení pokraèovat v instalaci, klepnìte na tlaèítko 'Další'.
WizardInfoAfter=Informace
InfoAfterLabel=Pøed dalším pokraèováním vìnujte prosím pozornost následujícím dùleitım informacím.
InfoAfterClickLabel=A bude pøipravení pokraèovat v instalaci, klepnìte na tlaèítko 'Další'.

; *** "User Information" wizard page
WizardUserInfo=Informace o uivateli
UserInfoDesc=Vlote prosím poadované informace.
UserInfoName=&Jméno uivatele:
UserInfoOrg=&Organizace:
UserInfoSerial=&Instalaèní èíslo:
UserInfoNameRequired=Jméno uivatele je vyadováno.

; *** "Select Destination Location" wizard page
WizardSelectDir=Urèení umístìní aplikace
SelectDirDesc=Urèete prosím, kam bude aplikace [name] nainstalována.
SelectDirLabel3=Prùvodce instalací nainstaluje aplikaci [name] do následující sloky.
SelectDirBrowseLabel=Pro pokraèování klepnìte na tlaèítko 'Další'. Pøejete-li si vybrat odlišnou sloku, klepnìte na tlaèítko 'Procházet'.
DiskSpaceMBLabel=Je vyadováno nejménì [mb] MB volného místa na disku.
ToUNCPathname=Aplikaci není moné instalovat do síové sloky. Pøejete-li si aplikaci instalovat na zaøízení v síti, bude nutné pøipojit síovı disk.
InvalidPath=Vloit je nutné plnou cestu s urèením písmena jednotky; napøíklad:%n%nC:\APP%n%npopøípadì UNC cestu ve tvaru:%n%n\\server\share
InvalidDrive=Vybrané zaøízení nebo sdílená cesta UNC neexistuje nebo není pøístupná. Zvolte prosím cestu jinou.
DiskSpaceWarningTitle=Nedostatek místa na disku
DiskSpaceWarning=Prùvodce instalací vyaduje nejménì %1 KB volného místa na disku. Na vybrané jednotce je však k dispozici pouze %2 KB volného místa.%n%nPokraèovat?
DirNameTooLong=Název sloky nebo cesta je pøíliš dlouhá.
InvalidDirName=Název sloky není platnı.
BadDirName32=V názvu sloky nemohou bıt obsaeny následující znaky:%n%n%1
DirExistsTitle=Sloka existuje
DirExists=Sloka:%n%n%1%n%nji existuje. Pokraèovat v instalaci do této sloky?
DirDoesntExistTitle=Sloka neexistuje
DirDoesntExist=Sloka:%n%n%1%n%nneexistuje. Vytvoøit tuto sloku?

; *** "Select Components" wizard page
WizardSelectComponents=Volba souèástí
SelectComponentsDesc=Urèete prosím, které souèásti aplikace budou nainstalovány.
SelectComponentsLabel2=Zvolte prosím souèásti, které budou nainstalovány. Zrušte oznaèení souèástí, které si nepøejete instalovat. Pro pokraèování klepnìte na tlaèítko 'Další'.
FullInstallation=Plná instalace
; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Typická instalace
CustomInstallation=Vlastní instalace
NoUninstallWarningTitle=Souèást existuje
NoUninstallWarning=Souèást:%n%n%1%n%n ji byla nainstalována. Zrušení oznaèení této souèásti nepovede k její odinstalaci.%n%nPokraèovat?
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=Vybrané souèásti vyadují nejménì [mb] MB volného místa na disku.

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Volba doplòujících moností
SelectTasksDesc=Urèete prosím, které další akce budou v prùbìhu instalace provedeny.
SelectTasksLabel2=Zvolte prosím doplòující akce, které budou v prùbìhu instalace aplikace [name] provedeny. Pro pokraèování klepnìte na tlaèítko 'Další'.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=Vıbìr sloky v nabídce Start
SelectStartMenuFolderDesc=Urèete prosím, kam má prùvodce instalací umístit zástupce aplikace.
SelectStartMenuFolderLabel3=Prùvodce instalací vytvoøí zástupce aplikace [name] v uvedené sloce nabdky Start.
SelectStartMenuFolderBrowseLabel=Pro pokraèování klepnìte na tlaèítko 'Další'. Pøejete-li si zvolit jinou sloku, klepnìte na tlaèítko 'Procházet'.
NoIconsCheck=&Nevytváøet ádné zástupce
MustEnterGroupName=Pro pokraèování je nutné vloit název sloky.
GroupNameTooLong=Název sloky je pøíliš dlouhı.
InvalidGroupName=Název sloky není platnı.
BadGroupName=V názvu sloky nemohou bıt obsaeny následující znaky:%n%n%1
NoProgramGroupCheck2=&Nevytváøet zástupce v nabídce Start

; *** "Ready to Install" wizard page
WizardReady=Instalace pøipravena
ReadyLabel1=Prùvodce instalací je pøipraven zahájit instalaci aplikace [name] na váš poèítaè.
ReadyLabel2a=Pro zahájení instalace klepnìte na tlaèítko 'Instalovat'. Vrátit se k pøedchozím krokùm je moné klepnutím na tlaèítko 'Zpìt'.
ReadyLabel2b=Pro zahájení instalace klepnìte na tlaèítko 'Instalovat'.
ReadyMemoUserInfo=Informace o uivateli:
ReadyMemoDir=Umístení aplikace:
ReadyMemoType=Typ instalace:
ReadyMemoComponents=Zvolené souèásti:
ReadyMemoGroup=Sloka v nabídce Start:
ReadyMemoTasks=Doplòující monosti:

; *** "Preparing to Install" wizard page
WizardPreparing=Pøíprava k instalaci
PreparingDesc=Prùvodce instalací pøipravuje instalaci aplikace [name] na váš poèítaè.
PreviousInstallNotCompleted=Pøedchozí instalace/odebrání jiné aplikace nebylo dokonèeno. Pro dokonèení instalace je nutné restartovat poèítaè.%n%nPo restartu spuste prosím prùvodce instalací znovu.
CannotContinue=Prùvodce instalací nemùe v instalaci pokraèovat. Klepnìte prosím na 'Zrušit' a prùvodce instalací ukonèete.

; *** "Installing" wizard page
WizardInstalling=Instalace
InstallingLabel=Èekejte prosím na dokonèení instalace aplikace [name] na Váš poèítaè.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Dokonèení instalace
FinishedLabelNoIcons=Prùvodce instalací dokonèil instalaci aplikace [name] na váš poèítaè.
FinishedLabel=Prùvodce instalací dokonèil instalaci aplikace [name] na váš poèítaè. Aplikace mùe bıt spuštìna pomocí nainstalovanıch zástupcù.
ClickFinish=Pro ukonèení prùvodce instalací klepnìte na tlaèítko 'Dokonèit'.
FinishedRestartLabel=Pro dokonèení instalace aplikace [name] je nutné restartovat váš poèítaè. Restartovat?
FinishedRestartMessage=Pro dokonèení instalace aplikace [name] je nutné restartovat váš poèítaè. %n%nRestartovat?
ShowReadmeCheck=Zobrazit soubor README
YesRadio=&Restartovat poèítaè teï
NoRadio=Poèítaè restartuji &pozdìji
; used for example as 'Run MyProg.exe'
RunEntryExec=Spustit %1
; used for example as 'View Readme.txt'
RunEntryShellExec=Zobrazit %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=Prùvodce instalací potøebuje další disk
SelectDiskLabel2=Vlote prosím disk %1 a klepnìte na tlaèítko 'OK'.%n%n nacházejí-li se soubory v jiné sloce ne ve sloceIf the files on this disk can be found in a folder other than the one displayed below, enter the correct path or click Browse.
PathLabel=&Cesta:
FileNotInDir2=Soubor "%1" nebyl na "%2" nalezen. Ovìøte prosím, e je disk správnì vloen nebo vyberte jinou sloku.
SelectDirectoryLabel=Urèete prosím umístìní dalšího disku.

; *** Installation phase messages
SetupAborted=Prùvodce instalací nebyl dokonèen.%n%nVyøešte prosím nedostatky a opakujte spuštìní prùvodce instalací.
EntryAbortRetryIgnore=Pro opakování akce klepnìte na tlaèítko 'Znovu', pro pokraèování na tlaèítko 'Ignorovat'. Pro zrušení akce klepnìte na tlaèítko 'Pøerušit'.

; *** Installation status messages
StatusCreateDirs=Vytváøení struktury sloek...
StatusExtractFiles=Rozmísování souborù...
StatusCreateIcons=Vytváøení zástupcù...
StatusCreateIniEntries=Vytváøení poloek v souborech INI...
StatusCreateRegistryEntries=Vytváøení poloek v systémovém registru...
StatusRegisterFiles=Registrace souborù...
StatusSavingUninstall=Ukládání informací pro odinstalaci...
StatusRunProgram=Dokonèování instalace...
StatusRollback=Vracení provedenıch zmìn...

; *** Misc. errors
ErrorInternal2=Vnitøní chyba: %1
ErrorFunctionFailedNoCode=Funkce %1 selhala
ErrorFunctionFailed=Funkce %1 selhala; kód selhání %2
ErrorFunctionFailedWithMessage=Funkce %1 selhala; kód selhání %2.%n%3
ErrorExecutingProgram=Nebylo moné spustit soubor:%n%1

; *** Registry errors
ErrorRegOpenKey=Chyba pøi otevírání klíèe systémového registru:%n%1\%2
ErrorRegCreateKey=Chyba pøi vytváøení klíèe systémového registru:%n%1\%2
ErrorRegWriteKey=Chyba pøi zápisu klíèe systémového registru:%n%1\%2

; *** INI errors
ErrorIniEntry=Chyba pøi vytváøení poloky v souboru INI "%1".

; *** File copying errors
FileAbortRetryIgnore=Pro pokraèování klpenìte na 'Pokraèovat', pro pøeskoèení tohoto souboru (není doporuèeno) klepnìte na 'Ignorovat'. Instalaci je moné pøerušit klepnutím na tlaèítko 'Zrušit'.
FileAbortRetryIgnore2=Pro pokraèování klepnìte na 'Pokraèovat', pro pøeskoèení tohoto souboru (není doporuèeno) klepnìte na 'Ignorovat'. Instalaci je moné pøerušit klepnutím na tlaèítko 'Zrušit'.
EntryAbortRetryIgnore=Pro opakování akce klepnìte na tlaèítko 'Znovu', pro pokraèování na tlaèítko 'Ignorovat'. Pro zrušení akce klepnìte na tlaèítko 'Pøerušit'.
SourceIsCorrupted=Zdrojovı soubor je poškozen
SourceDoesntExist=Zdrojovı soubor "%1" nebyl nalezen
ExistingFileReadOnly=Pùvodní soubor je oznaèen pøíznakem jen pro ètení.%n%nPro zrušení pøíznaku jen pro ètení a opakování akce klepnìte na tlaèítko 'Znovu', pro pøeskoèení souboru klepnìte na tlaèítko Ignorovat. Instalaci je moné pøerušit klepnutím na tlaèítko 'Zrušit'.
ErrorReadingExistingDest=Pøi pokusu o ètení souboru došlo k chybì:
FileExists=Soubor ji existuje.%n%nPøepsat?
ExistingFileNewer=Pùvodní soubor je novìjší ne soubor instalovanı prùvodcem instalací. Je doporuèeno zachovat pùvodní soubor.%n%nZachovat pùvodní soubor?
ErrorChangingAttr=Pøi pokusu o zmìnu pøíznakù pùvodního souboru došlo k chybì:
ErrorCreatingTemp=Pøi pokusu o vytvoøení souboru v cílové sloce došlo k chybì:
ErrorReadingSource=Pøi pokusu o ètení souboru došlo k chybì:
ErrorCopying=Pøi pokusu o kopírování souboru došlo k chybì:
ErrorReplacingExistingFile=Pøi pokusu o nahrazení pùvodního souboru došlo k chybì:
ErrorRestartReplace=Funkce RestartReplace selhala:
ErrorRenamingTemp=Pøi pokusu o pøejmenování souboru v cílové sloce došlo k chybì:
ErrorRegisterServer=Unable to register the DLL/OCX: %1
ErrorRegisterServerMissingExport=Funkce DllRegisterServer nebyla nalezena
ErrorRegisterTypeLib=Nebylo moné zaregistrovat knihovnu typù: %1

; *** Post-installation errors
ErrorOpeningReadme=Pøi pokusu o otevøení souboru README došlo k chybì:
ErrorRestartingComputer=Prùvodce instalací nemohl poèítaè restartovat. Restartujte prosím poèítaè ruènì.

; *** Uninstaller messages
UninstallNotFound=Soubor "%1" nebyl nalezen. Odinstalace memùe bıt provedena.
UninstallOpenError=Soubor "%1" nebyl otevøen. Odinstalace memùe bıt provedena
UninstallUnsupportedVer=Odinstalaèní soubor "%1" není ve formátu vyhovujícím této verzi odinstalaèního programu. Odinstalace memùe bıt provedena.
UninstallUnknownEntry=V odinstalaèním souboru (%1) byla nalezena neznámá poloka odinstalace
ConfirmUninstall=Opravdu si pøejete odstranit %1 a všechny jeho komponenty?
OnlyAdminCanUninstall=Odinstalace této instalace programu mùe bıt provedena pouze uivatelem s právy administrátora.
UninstallStatusLabel=Vyèkejte prosím dokud nebude %1 odstranìn z vašeho poèítaèe.
UninstalledAll=%1 byl úspìšnì odinstalován.
UninstalledMost=Odinstalace aplikace %1 byla úspìšnì dokonèena.%n%nNìkteré souèásti vak nebyly odstranìny. Tyto souèásti mohou bıt odstranìny ruènì.
UninstalledAndNeedsRestart=Pro dokonèení odinstalace je nutné %1 aby byl váš poèítaè restartován.%n%nRestartovat?
UninstallDataCorrupted=Soubor "%1" je porušen. Odinstalace memùe bıt provedena

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=Odstranìní sdíleného souboru
ConfirmDeleteSharedFile2=Podle údajù operaèního systému není uvedenı sdílenı soubor vyuíván ádnou aplikací. Odinstalovat a odstranit tento sdílenı soubor?%n%n Jestlie nìkteré programy soubor stále vyuívají, nemusí po jeho odinstalaci pracovat správnì. Nejste-li si jistí, zvolte 'Ne'. Zachováním souboru integrovaného do vašeho operaèního systému nebude mít za následek ádné poškození.
SharedFileNameLabel=File name:
SharedFileNameLabel=Název souboru:
SharedFileLocationLabel=Location:
SharedFileLocationLabel=Umístìní:
WizardUninstalling=Stav odinstalace
StatusUninstalling=Odinstalace aplikace %1...

; The custom messages below aren't used by Setup itself, but if you make
; use of them in your scripts, you'll want to translate them.

[CustomMessages]

NameAndVersion=%1, verze %2
AdditionalIcons=Doplòující zástupci:
CreateDesktopIcon=Vytvoøit zástupce na &ploše
CreateQuickLaunchIcon=Vytvoit zástupce v panelu &rychlého spouštìbí
ProgramOnTheWeb=Aplikace %1 na Internetu
UninstallProgram=Odinstalovat %1
LaunchProgram=Spustit %1
AssocFileExtension=&Pøidruit aplikaci %1 k pøíponì souboru %2
AssocingFileExtension=Pøidruování aplikace %1 k pøíponì souboru %2...
