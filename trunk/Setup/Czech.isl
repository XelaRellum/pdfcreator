; *** Inno Setup version 4.2.2+ Czech messages ***
;
; Copyright (c) 2005 Martin Koz�k (martin.kozak@openoffice.cz)
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
SetupAppTitle=Pr�vodce instalac�
SetupWindowTitle=Pr�vodce instalac� aplikace %1
UninstallAppTitle=Pr�vodce odinstalac�
UninstallAppFullTitle=Pr�vodce odinstalac� aplikace %1

; *** Misc. common
InformationTitle=Informace
ConfirmTitle=Potvrzen�
ErrorTitle=Chyba

; *** SetupLdr messages
SetupLdrStartupMessage=Bude spu�t�n pr�vodce instalac� aplikace %1. P�ejete si pokra�ovat?
LdrCannotCreateTemp=Nebylo mo�no vytvo�it odkl�dac� soubor. Pr�vodce instalac� bude ukon�en
LdrCannotExecTemp=Nebylo mo�no spustit soubor v odkl�dac� slo�ce. Pr�vodce instalac� bude ukon�en

; *** Startup error messages
LastErrorMessage=%1.%n%nChyba %2: %3
SetupFileMissing=Soubor %1 nebyl v instala�n� slo�ce nalezen. Opravte pros�m probl�m nebo pou�ijte jinou kopii aplikace
SetupFileCorrupt=Soubory pr�vodce instalac� jsou po�kozeny. Pou�ijte pros�m jinou kopii aplikace.
SetupFileCorruptOrWrongVer=Soubory pr�vodce instalac� jsou po�kozeny nebo nejsou kompatibiln� s touto verz� pr�vodce. Opravte pros�m probl�m nebou pou�ijte jinou kopii aplikace.
NotOnThisPlatform=Aplikace nen� ur�ena pro platformu %1.
OnlyOnThisPlatform=Aplikace je ur�ena pro platformu %1.
WinVersionTooLowError=Aplikace vy�aduje verzi %2 syst�mu %1 nebo vy���.
WinVersionTooHighError=Aplikaci nen� mo�n� ve verzi %2 syst�mu %2 a vy���ch vyu��t.
AdminPrivilegesRequired=Pro instalaci aplikace mus�te m�t pr�va administr�tora.
PowerUserPrivilegesRequired=Pro instalaci aplikace mus�te m�t pr�va administr�tora nebo b�t �lenem skupiny Power Users.
SetupAppRunningError=Aplikace %1 je spu�t�na.%n%nUzav�ete pros�m v�echny instance a klepn�te na tla��tko 'OK' pro pokra�ov�n� nebo 'Zru�it' pro ukon�en� pr�vodce instalac�.
UninstallAppRunningError=Aplikace %1 je spu�t�na.%n%nUzav�ete pros�m v�echny instance a klepn�te na tla��tko 'OK' pro pokra�ov�n� nebo 'Zru�it' pro ukon�en� pr�vodce instalac�.

; *** Misc. errors
ErrorCreatingDir=Nebylo mo�n� vytvo�it adres�� "%1"
ErrorTooManyFilesInDir=Nebylo mo�n� vytvo�it soubor v adres��i "%1". Adres�� obsahuje p��li� mnoho soubor�.

; *** Setup common messages
ExitSetupTitle=Ukon�en� pr�vodce instalac�
ExitSetupMessage=Instalace aplikace nebyla dokon�ena. Jestli�e ukon��te pr�vodce instalac�, aplikace nebude nainstalov�na.%n%nPro dokon�en� instalace je mo�n� pr�vodce instalac� spustit kdykoliv pozd�ji.%n%nUkon�it pr�vodce instalac�?
AboutSetupMenuItem=&O pr�vodci instalac�...
AboutSetupTitle=O pr�vodci instalac�
AboutSetupMessage=%1 verze %2%n%3%n%n%1 internetov� adresa:%n%4
AboutSetupNote=

; *** Buttons
ButtonBack=< &Zp�t
ButtonNext=&Dal�� >
ButtonInstall=&Instalovat
ButtonOK=OK
ButtonCancel=Zru�it
ButtonYes=&Ano
ButtonYesToAll=Ano &v�em
ButtonNo=&Ne
ButtonNoToAll=N&e v�em
ButtonFinish=&Dokon�it
ButtonBrowse=&Proch�zet...
ButtonWizardBrowse=P&roch�zet...
ButtonNewFolder=&Vytvo�it slo�ku

; *** "Select Language" dialog messages
SelectLanguageTitle=Zvolit jazyk pr�vodce instalac�
SelectLanguageLabel=Zvolte pros�m jazyk pr�vodce instalac�:

; *** Common wizard text
ClickNext=Pokra�ujte klepnut�m na tla��tko 'Dal��'. Klepnut�m na tla��tko 'Zru�it' pr�vodce instalac� ukon��te.
BeveledLabel=
BrowseDialogTitle=Nal�zt slo�ku
BrowseDialogLabel=Vyberte pros�m slo�ku a klepn�te na tla��tko 'OK'.
NewFolderName=Nov� slo�ka

; *** "Welcome" wizard page
WelcomeLabel1=V�tejte v pr�vodci instalac� aplikace [name]
WelcomeLabel2=Pr�vodce instalac� nainstaluje na v� po��ta� aplikaci [name/ver].%n%nP�ed pokra�ov�n�m instalace je doporu�eno uzav��t ostatn� b��c� aplikace.

; *** "Password" wizard page
WizardPassword=Heslo
PasswordLabel1=Instala�n� bal��ek je chr�n�n heslem.
PasswordLabel3=Vlo�te pros�m heslo a klepn�te na tla��tko 'Dal��'. Ov��ovac� proces rozli�uje mal� a vek� p�smena.
PasswordEditLabel=&Heslo:
IncorrectPassword=Vlo�en� heslo nesouhlas�. Opakujte pros�m akci znovu.

; *** "License Agreement" wizard page
WizardLicense=Licen�n� ujedn�n�
LicenseLabel=P�ed pokra�ov�n�m v�nujte pros�m pozornost n�sleduj�c�m d�le�it�m informac�m.
LicenseLabel3=V�nujte pros�m pozornost n�sleduj�c�mu licen�n�mu ujedn�n�. Podm�nky tohoto ujedn�n� je nutn� p�ed pokra�ov�n�m p�ijmout.
LicenseAccepted=&P�ijmout licen�n� ujedn�n�
LicenseNotAccepted=&Odm�tnout licen�n� ujedn�n�

; *** "Information" wizard pages
WizardInfoBefore=Informace
InfoBeforeLabel=P�ed pokra�ov�n�m v�nujte pros�m pozornost n�sleduj�c�m d�le�it�m informac�m.
InfoBeforeClickLabel=A� bude p�ipraven� pokra�ovat v instalaci, klepn�te na tla��tko 'Dal��'.
WizardInfoAfter=Informace
InfoAfterLabel=P�ed dal��m pokra�ov�n�m v�nujte pros�m pozornost n�sleduj�c�m d�le�it�m informac�m.
InfoAfterClickLabel=A� bude p�ipraven� pokra�ovat v instalaci, klepn�te na tla��tko 'Dal��'.

; *** "User Information" wizard page
WizardUserInfo=Informace o u�ivateli
UserInfoDesc=Vlo�te pros�m po�adovan� informace.
UserInfoName=&Jm�no u�ivatele:
UserInfoOrg=&Organizace:
UserInfoSerial=&Instala�n� ��slo:
UserInfoNameRequired=Jm�no u�ivatele je vy�adov�no.

; *** "Select Destination Location" wizard page
WizardSelectDir=Ur�en� um�st�n� aplikace
SelectDirDesc=Ur�ete pros�m, kam bude aplikace [name] nainstalov�na.
SelectDirLabel3=Pr�vodce instalac� nainstaluje aplikaci [name] do n�sleduj�c� slo�ky.
SelectDirBrowseLabel=Pro pokra�ov�n� klepn�te na tla��tko 'Dal��'. P�ejete-li si vybrat odli�nou slo�ku, klepn�te na tla��tko 'Proch�zet'.
DiskSpaceMBLabel=Je vy�adov�no nejm�n� [mb] MB voln�ho m�sta na disku.
ToUNCPathname=Aplikaci nen� mo�n� instalovat do s�ov� slo�ky. P�ejete-li si aplikaci instalovat na za��zen� v s�ti, bude nutn� p�ipojit s�ov� disk.
InvalidPath=Vlo�it je nutn� plnou cestu s ur�en�m p�smena jednotky; nap��klad:%n%nC:\APP%n%npop��pad� UNC cestu ve tvaru:%n%n\\server\share
InvalidDrive=Vybran� za��zen� nebo sd�len� cesta UNC neexistuje nebo nen� p��stupn�. Zvolte pros�m cestu jinou.
DiskSpaceWarningTitle=Nedostatek m�sta na disku
DiskSpaceWarning=Pr�vodce instalac� vy�aduje nejm�n� %1 KB voln�ho m�sta na disku. Na vybran� jednotce je v�ak k dispozici pouze %2 KB voln�ho m�sta.%n%nPokra�ovat?
DirNameTooLong=N�zev slo�ky nebo cesta je p��li� dlouh�.
InvalidDirName=N�zev slo�ky nen� platn�.
BadDirName32=V n�zvu slo�ky nemohou b�t obsa�eny n�sleduj�c� znaky:%n%n%1
DirExistsTitle=Slo�ka existuje
DirExists=Slo�ka:%n%n%1%n%nji� existuje. Pokra�ovat v instalaci do t�to slo�ky?
DirDoesntExistTitle=Slo�ka neexistuje
DirDoesntExist=Slo�ka:%n%n%1%n%nneexistuje. Vytvo�it tuto slo�ku?

; *** "Select Components" wizard page
WizardSelectComponents=Volba sou��st�
SelectComponentsDesc=Ur�ete pros�m, kter� sou��sti aplikace budou nainstalov�ny.
SelectComponentsLabel2=Zvolte pros�m sou��sti, kter� budou nainstalov�ny. Zru�te ozna�en� sou��st�, kter� si nep�ejete instalovat. Pro pokra�ov�n� klepn�te na tla��tko 'Dal��'.
FullInstallation=Pln� instalace
; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Typick� instalace
CustomInstallation=Vlastn� instalace
NoUninstallWarningTitle=Sou��st existuje
NoUninstallWarning=Sou��st:%n%n%1%n%n ji� byla nainstalov�na. Zru�en� ozna�en� t�to sou��sti nepovede k jej� odinstalaci.%n%nPokra�ovat?
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=Vybran� sou��sti vy�aduj� nejm�n� [mb] MB voln�ho m�sta na disku.

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Volba dopl�uj�c�ch mo�nost�
SelectTasksDesc=Ur�ete pros�m, kter� dal�� akce budou v pr�b�hu instalace provedeny.
SelectTasksLabel2=Zvolte pros�m dopl�uj�c� akce, kter� budou v pr�b�hu instalace aplikace [name] provedeny. Pro pokra�ov�n� klepn�te na tla��tko 'Dal��'.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=V�b�r slo�ky v nab�dce Start
SelectStartMenuFolderDesc=Ur�ete pros�m, kam m� pr�vodce instalac� um�stit z�stupce aplikace.
SelectStartMenuFolderLabel3=Pr�vodce instalac� vytvo�� z�stupce aplikace [name] v uveden� slo�ce nabdky Start.
SelectStartMenuFolderBrowseLabel=Pro pokra�ov�n� klepn�te na tla��tko 'Dal��'. P�ejete-li si zvolit jinou slo�ku, klepn�te na tla��tko 'Proch�zet'.
NoIconsCheck=&Nevytv��et ��dn� z�stupce
MustEnterGroupName=Pro pokra�ov�n� je nutn� vlo�it n�zev slo�ky.
GroupNameTooLong=N�zev slo�ky je p��li� dlouh�.
InvalidGroupName=N�zev slo�ky nen� platn�.
BadGroupName=V n�zvu slo�ky nemohou b�t obsa�eny n�sleduj�c� znaky:%n%n%1
NoProgramGroupCheck2=&Nevytv��et z�stupce v nab�dce Start

; *** "Ready to Install" wizard page
WizardReady=Instalace p�ipravena
ReadyLabel1=Pr�vodce instalac� je p�ipraven zah�jit instalaci aplikace [name] na v� po��ta�.
ReadyLabel2a=Pro zah�jen� instalace klepn�te na tla��tko 'Instalovat'. Vr�tit se k p�edchoz�m krok�m je mo�n� klepnut�m na tla��tko 'Zp�t'.
ReadyLabel2b=Pro zah�jen� instalace klepn�te na tla��tko 'Instalovat'.
ReadyMemoUserInfo=Informace o u�ivateli:
ReadyMemoDir=Um�sten� aplikace:
ReadyMemoType=Typ instalace:
ReadyMemoComponents=Zvolen� sou��sti:
ReadyMemoGroup=Slo�ka v nab�dce Start:
ReadyMemoTasks=Dopl�uj�c� mo�nosti:

; *** "Preparing to Install" wizard page
WizardPreparing=P��prava k instalaci
PreparingDesc=Pr�vodce instalac� p�ipravuje instalaci aplikace [name] na v� po��ta�.
PreviousInstallNotCompleted=P�edchoz� instalace/odebr�n� jin� aplikace nebylo dokon�eno. Pro dokon�en� instalace je nutn� restartovat po��ta�.%n%nPo restartu spus�te pros�m pr�vodce instalac� znovu.
CannotContinue=Pr�vodce instalac� nem��e v instalaci pokra�ovat. Klepn�te pros�m na 'Zru�it' a pr�vodce instalac� ukon�ete.

; *** "Installing" wizard page
WizardInstalling=Instalace
InstallingLabel=�ekejte pros�m na dokon�en� instalace aplikace [name] na V� po��ta�.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Dokon�en� instalace
FinishedLabelNoIcons=Pr�vodce instalac� dokon�il instalaci aplikace [name] na v� po��ta�.
FinishedLabel=Pr�vodce instalac� dokon�il instalaci aplikace [name] na v� po��ta�. Aplikace m��e b�t spu�t�na pomoc� nainstalovan�ch z�stupc�.
ClickFinish=Pro ukon�en� pr�vodce instalac� klepn�te na tla��tko 'Dokon�it'.
FinishedRestartLabel=Pro dokon�en� instalace aplikace [name] je nutn� restartovat v� po��ta�. Restartovat?
FinishedRestartMessage=Pro dokon�en� instalace aplikace [name] je nutn� restartovat v� po��ta�. %n%nRestartovat?
ShowReadmeCheck=Zobrazit soubor README
YesRadio=&Restartovat po��ta� te�
NoRadio=Po��ta� restartuji &pozd�ji
; used for example as 'Run MyProg.exe'
RunEntryExec=Spustit %1
; used for example as 'View Readme.txt'
RunEntryShellExec=Zobrazit %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=Pr�vodce instalac� pot�ebuje dal�� disk
SelectDiskLabel2=Vlo�te pros�m disk %1 a klepn�te na tla��tko 'OK'.%n%n nach�zej�-li se soubory v jin� slo�ce ne� ve slo�ceIf the files on this disk can be found in a folder other than the one displayed below, enter the correct path or click Browse.
PathLabel=&Cesta:
FileNotInDir2=Soubor "%1" nebyl na "%2" nalezen. Ov��te pros�m, �e je disk spr�vn� vlo�en nebo vyberte jinou slo�ku.
SelectDirectoryLabel=Ur�ete pros�m um�st�n� dal��ho disku.

; *** Installation phase messages
SetupAborted=Pr�vodce instalac� nebyl dokon�en.%n%nVy�e�te pros�m nedostatky a opakujte spu�t�n� pr�vodce instalac�.
EntryAbortRetryIgnore=Pro opakov�n� akce klepn�te na tla��tko 'Znovu', pro pokra�ov�n� na tla��tko 'Ignorovat'. Pro zru�en� akce klepn�te na tla��tko 'P�eru�it'.

; *** Installation status messages
StatusCreateDirs=Vytv��en� struktury slo�ek...
StatusExtractFiles=Rozm�s�ov�n� soubor�...
StatusCreateIcons=Vytv��en� z�stupc�...
StatusCreateIniEntries=Vytv��en� polo�ek v souborech INI...
StatusCreateRegistryEntries=Vytv��en� polo�ek v syst�mov�m registru...
StatusRegisterFiles=Registrace soubor�...
StatusSavingUninstall=Ukl�d�n� informac� pro odinstalaci...
StatusRunProgram=Dokon�ov�n� instalace...
StatusRollback=Vracen� proveden�ch zm�n...

; *** Misc. errors
ErrorInternal2=Vnit�n� chyba: %1
ErrorFunctionFailedNoCode=Funkce %1 selhala
ErrorFunctionFailed=Funkce %1 selhala; k�d selh�n� %2
ErrorFunctionFailedWithMessage=Funkce %1 selhala; k�d selh�n� %2.%n%3
ErrorExecutingProgram=Nebylo mo�n� spustit soubor:%n%1

; *** Registry errors
ErrorRegOpenKey=Chyba p�i otev�r�n� kl��e syst�mov�ho registru:%n%1\%2
ErrorRegCreateKey=Chyba p�i vytv��en� kl��e syst�mov�ho registru:%n%1\%2
ErrorRegWriteKey=Chyba p�i z�pisu kl��e syst�mov�ho registru:%n%1\%2

; *** INI errors
ErrorIniEntry=Chyba p�i vytv��en� polo�ky v souboru INI "%1".

; *** File copying errors
FileAbortRetryIgnore=Pro pokra�ov�n� klpen�te na 'Pokra�ovat', pro p�esko�en� tohoto souboru (nen� doporu�eno) klepn�te na 'Ignorovat'. Instalaci je mo�n� p�eru�it klepnut�m na tla��tko 'Zru�it'.
FileAbortRetryIgnore2=Pro pokra�ov�n� klepn�te na 'Pokra�ovat', pro p�esko�en� tohoto souboru (nen� doporu�eno) klepn�te na 'Ignorovat'. Instalaci je mo�n� p�eru�it klepnut�m na tla��tko 'Zru�it'.
EntryAbortRetryIgnore=Pro opakov�n� akce klepn�te na tla��tko 'Znovu', pro pokra�ov�n� na tla��tko 'Ignorovat'. Pro zru�en� akce klepn�te na tla��tko 'P�eru�it'.
SourceIsCorrupted=Zdrojov� soubor je po�kozen
SourceDoesntExist=Zdrojov� soubor "%1" nebyl nalezen
ExistingFileReadOnly=P�vodn� soubor je ozna�en p��znakem jen pro �ten�.%n%nPro zru�en� p��znaku jen pro �ten� a opakov�n� akce klepn�te na tla��tko 'Znovu', pro p�esko�en� souboru klepn�te na tla��tko Ignorovat. Instalaci je mo�n� p�eru�it klepnut�m na tla��tko 'Zru�it'.
ErrorReadingExistingDest=P�i pokusu o �ten� souboru do�lo k chyb�:
FileExists=Soubor ji� existuje.%n%nP�epsat?
ExistingFileNewer=P�vodn� soubor je nov�j�� ne� soubor instalovan� pr�vodcem instalac�. Je doporu�eno zachovat p�vodn� soubor.%n%nZachovat p�vodn� soubor?
ErrorChangingAttr=P�i pokusu o zm�nu p��znak� p�vodn�ho souboru do�lo k chyb�:
ErrorCreatingTemp=P�i pokusu o vytvo�en� souboru v c�lov� slo�ce do�lo k chyb�:
ErrorReadingSource=P�i pokusu o �ten� souboru do�lo k chyb�:
ErrorCopying=P�i pokusu o kop�rov�n� souboru do�lo k chyb�:
ErrorReplacingExistingFile=P�i pokusu o nahrazen� p�vodn�ho souboru do�lo k chyb�:
ErrorRestartReplace=Funkce RestartReplace selhala:
ErrorRenamingTemp=P�i pokusu o p�ejmenov�n� souboru v c�lov� slo�ce do�lo k chyb�:
ErrorRegisterServer=Unable to register the DLL/OCX: %1
ErrorRegisterServerMissingExport=Funkce DllRegisterServer nebyla nalezena
ErrorRegisterTypeLib=Nebylo mo�n� zaregistrovat knihovnu typ�: %1

; *** Post-installation errors
ErrorOpeningReadme=P�i pokusu o otev�en� souboru README do�lo k chyb�:
ErrorRestartingComputer=Pr�vodce instalac� nemohl po��ta� restartovat. Restartujte pros�m po��ta� ru�n�.

; *** Uninstaller messages
UninstallNotFound=Soubor "%1" nebyl nalezen. Odinstalace mem��e b�t provedena.
UninstallOpenError=Soubor "%1" nebyl otev�en. Odinstalace mem��e b�t provedena
UninstallUnsupportedVer=Odinstala�n� soubor "%1" nen� ve form�tu vyhovuj�c�m t�to verzi odinstala�n�ho programu. Odinstalace mem��e b�t provedena.
UninstallUnknownEntry=V odinstala�n�m souboru (%1) byla nalezena nezn�m� polo�ka odinstalace
ConfirmUninstall=Opravdu si p�ejete odstranit %1 a v�echny jeho komponenty?
OnlyAdminCanUninstall=Odinstalace t�to instalace programu m��e b�t provedena pouze u�ivatelem s pr�vy administr�tora.
UninstallStatusLabel=Vy�kejte pros�m dokud nebude %1 odstran�n z va�eho po��ta�e.
UninstalledAll=%1 byl �sp�n� odinstalov�n.
UninstalledMost=Odinstalace aplikace %1 byla �sp�n� dokon�ena.%n%nN�kter� sou��sti vak nebyly odstran�ny. Tyto sou��sti mohou b�t odstran�ny ru�n�.
UninstalledAndNeedsRestart=Pro dokon�en� odinstalace je nutn� %1 aby byl v� po��ta� restartov�n.%n%nRestartovat?
UninstallDataCorrupted=Soubor "%1" je poru�en. Odinstalace mem��e b�t provedena

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=Odstran�n� sd�len�ho souboru
ConfirmDeleteSharedFile2=Podle �daj� opera�n�ho syst�mu nen� uveden� sd�len� soubor vyu��v�n ��dnou aplikac�. Odinstalovat a odstranit tento sd�len� soubor?%n%n Jestli�e n�kter� programy soubor st�le vyu��vaj�, nemus� po jeho odinstalaci pracovat spr�vn�. Nejste-li si jist�, zvolte 'Ne'. Zachov�n�m souboru integrovan�ho do va�eho opera�n�ho syst�mu nebude m�t za n�sledek ��dn� po�kozen�.
SharedFileNameLabel=File name:
SharedFileNameLabel=N�zev souboru:
SharedFileLocationLabel=Location:
SharedFileLocationLabel=Um�st�n�:
WizardUninstalling=Stav odinstalace
StatusUninstalling=Odinstalace aplikace %1...

; The custom messages below aren't used by Setup itself, but if you make
; use of them in your scripts, you'll want to translate them.

[CustomMessages]

NameAndVersion=%1, verze %2
AdditionalIcons=Dopl�uj�c� z�stupci:
CreateDesktopIcon=Vytvo�it z�stupce na &plo�e
CreateQuickLaunchIcon=Vytvoit z�stupce v panelu &rychl�ho spou�t�b�
ProgramOnTheWeb=Aplikace %1 na Internetu
UninstallProgram=Odinstalovat %1
LaunchProgram=Spustit %1
AssocFileExtension=&P�idru�it aplikaci %1 k p��pon� souboru %2
AssocingFileExtension=P�idru�ov�n� aplikace %1 k p��pon� souboru %2...
