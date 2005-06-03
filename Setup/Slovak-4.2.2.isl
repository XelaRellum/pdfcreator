; *** Inno Setup version 4.2.2+ English messages ***
;
;
; Translated by:              Juraj Matel
; Contact:          j.matel@orangemail.sk
;                      http://matel.wz.cz
;
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
LanguageName=Slovak
LanguageID=$041B
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
SetupAppTitle=Sprievodca in�tal�ciou
SetupWindowTitle=Sprievodca in�tal�ciou - %1
UninstallAppTitle=Odin�talova�
UninstallAppFullTitle=Odin�talova� - %1

; *** Misc. common
InformationTitle=Inform�cia
ConfirmTitle=Potvrdenie
ErrorTitle=Chyba

; *** SetupLdr messages
SetupLdrStartupMessage=Chyst�te sa nain�talova� program %1. Chcete pokra�ova�?
LdrCannotCreateTemp=Nie je mo�n� vytvori� do�asn� s�bor . In�tal�cia ukon�en�.
LdrCannotExecTemp=Nie je mo�n� spusti� s�bor v do�asnom prie�inku. In�tal�cia ukon�en�.

; *** Startup error messages
LastErrorMessage=%1.%n%nChyba %2: %3
SetupFileMissing=V in�tala�nom prie�inku ch�ba s�bor %1. Ak chcete pokra�ova�, opravte tento probl�m alebo po�iadajte o nov� k�piu programu.
SetupFileCorrupt=In�tala�n� s�bor je po�koden�. Po�iadajte o nov� verziu programu.
SetupFileCorruptOrWrongVer=In�tala�n� s�bor je po�koden� alebo nekompatibiln� s aktu�lnou verziou in�tal�tora. Ak chcete pokra�ova�, opravte tento probl�m alebo po�iadajte o nov� k�piu programu.
NotOnThisPlatform=Tento program sa na %1 ned� spusti�.
OnlyOnThisPlatform=Tento program sa d� spusti� len na %1.
WinVersionTooLowError=Program vy�aduje %1 verzia %2 alebo nov�ia.
WinVersionTooHighError=Tento program nie je mo�n� nain�talova� na %1 verzie %2 alebo nov�ej.
AdminPrivilegesRequired=Ak chcete pokra�ova� v in�tal�cii mus�te by� prihl�sen� ako pou��vate� Administr�tor.
PowerUserPrivilegesRequired=Ak chcete pokra�ova� v in�tal�cii mus�te by� prihl�sen� ako pou��vate� Administr�tor alebo by� skupiny Power Users.
SetupAppRunningError=In�tal�tor zistil, �e program %1 je pr�ve spusten�.%n%nUkon�ite v�etky spusten� aplik�cie. Ak chcete pokra�ova�, kliknite na tla�idlo �alej. Kliknut�m na tla�idlo Zru�i� in�tal�ciu ukon��te.
UninstallAppRunningError=In�tal�tor zistil, �e program %1 je pr�ve spusten�.%n%nUkon�ite v�etky spusten� aplik�cie. Ak chcete pokra�ova�, kliknite na tla�idlo �alej. Kliknut�m na tla�idlo Zru�i� in�tal�ciu ukon��te.

; *** Misc. errors
ErrorCreatingDir=In�tal�tor nemohol vytvori� prie�inok �%1�.
ErrorTooManyFilesInDir=In�tal�tor nemohol vytvori� s�bor v prie�inku �%1�, preto�e obsahuje pr�li� ve�a s�borov.

; *** Setup common messages
ExitSetupTitle=Ukon�enie in�tal�cie
ExitSetupMessage=In�tal�cia nie je dokon�en�. Ak ju teraz ukon��te, program nebude nain�talovan�.%n%nIn�tal�tor m��ete spusti� nesk�r a in�tal�ciu dokon�i�.%n%nChcete naozaj skon�i� in�tal�ciu?
AboutSetupMenuItem=�o je in�tala�n� program...
AboutSetupTitle=�o je in�tala�n� program...
AboutSetupMessage=%1 verzia %2%n%3%n%n%1, domovsk� str�nka:%n%4
AboutSetupNote=

; *** Buttons
ButtonBack=< Nasp�
ButtonNext=�alej >
ButtonInstall=&In�talova�
ButtonOK=OK
ButtonCancel=Zru�i�
ButtonYes=�no
ButtonYesToAll=�no pre v�etky
ButtonNo=&Nie
ButtonNoToAll=Nie pre v�etky
ButtonFinish=&Dokon�i�
ButtonBrowse=Preh�ad�va�...
ButtonWizardBrowse=Preh�ad�va�...
ButtonNewFolder=Vytvori� nov� prie�inok

; *** "Select Language" dialog messages
SelectLanguageTitle=V�ber jazyka
SelectLanguageLabel=Vyberte jazyk, ktor� chcete pou��va� po�as in�tal�cie:

; *** Common wizard text
ClickNext=Ak chcete pokra�ova�, kliknite na tla�idlo �alej. Kliknut�m na tla�idlo Zru�i� in�tal�ciu ukon��te.
BeveledLabel=
BrowseDialogTitle=V�ber prie�inka programu
BrowseDialogLabel=V nasleduj�com zozname vyberte prie�inok a kliknite na tla�idlo OK.
NewFolderName=Nov� prie�inok

; *** "Welcome" wizard page
WelcomeLabel1=V�ta v�s Sprievodca in�tal�ciou programu [name].
WelcomeLabel2=Chyst�te sa nain�talova� program [name/ver] na v� po��ta�.%n%nSk�r ako budete pokra�ova�, odpor��a sa ukon�i� v�etky ostatn� aplik�cie.

; *** "Password" wizard page
WizardPassword=Heslo
PasswordLabel1=In�tal�cia je chr�nen� heslom.
PasswordLabel3=Zadajte heslo a pokra�ujte v in�tal�cii kliknut�m na tla�idlo �alej. Rozli�ujte ve�k� a mal� p�smen�.
PasswordEditLabel=Heslo:
IncorrectPassword=Zadan� heslo nie je spr�vne. Sk�ste to znova.

; *** "License Agreement" wizard page
WizardLicense=Licen�n� zmluva
LicenseLabel=Pre��tajte si tieto d�le�it� inform�cie, pred za�at�m in�tal�cie.
LicenseLabel3=Pre��tajte si t�to Licen�n� zmluvu. Ak chcete pokra�ova� v in�tal�cii, mus�te s�hlasi� so zmluvou.
LicenseAccepted=S�hlas�m so zmluvou
LicenseNotAccepted=Nes�hlas�m so zmluvou

; *** "Information" wizard pages
WizardInfoBefore=Inform�cia
InfoBeforeLabel=Pre��tajte si tieto d�le�it� inform�cie, pred za�at�m in�tal�cie.
InfoBeforeClickLabel=Ak chcete pokra�ova�, kliknite na tla�idlo �alej.
WizardInfoAfter=Inform�cia
InfoAfterLabel=Pre��tajte si tieto d�le�it� inform�cie, pred za�at�m in�tal�cie.
InfoAfterClickLabel=Ak chcete pokra�ova�, kliknite na tla�idlo �alej.

; *** "User Information" wizard page
WizardUserInfo=Inform�cie o pou��vate�ovi
UserInfoDesc=Zadajte inform�cie o pou��vate�ovi.
UserInfoName=Meno pou��vate�a:
UserInfoOrg=Organiz�cia:
UserInfoSerial=S�riov� ��slo:
UserInfoNameRequired=Mus�te zada� meno pou��vate�a.

; *** "Select Destination Location" wizard page
WizardSelectDir=Umiestnenie programu
SelectDirDesc=Zadajte cestu k umiestneniu, kam chcete nain�talova� program [name].
SelectDirLabel3=Program [name] sa nain�taluje do nasleduj�ceho prie�inku.
SelectDirBrowseLabel=Ak chcete pokra�ova�, kliknite na tla�idlo �alej. Ak chcete vybra� in� prie�inok, kliknite na tla�idlo Preh�ad�va�.
DiskSpaceMBLabel=Po�adovan� miesto na disku: [mb] MB
ToUNCPathname=In�tal�tor nem��e pou�i� zadan� cestu UNC. Ak sa pok��ate nain�talova� tento program v sieti, pou�ite niektor� z dostupn�ch sie�ov�ch jednotiek.
InvalidPath=Zadajte �pln� cestu spolu s p�smenom jednotky (p�smeno:\cesta) alebo �pln� cestu spolu so znakom \\ na konci bez n�zvu s�boru (\\server\\zdie�anie).
InvalidDrive=Zadan� zariadenie alebo cesta UNC neexistuje alebo je odpojen�. Vyberte in� zariadenie alebo cestu.
DiskSpaceWarningTitle=Na disku nie je dos� miesta.
DiskSpaceWarning=Na dokon�enie in�tal�cie je potrebn�ch minim�lne %1 kB vo�n�ho miesta na disku, zvolen� jednotka obsahuje len %2 kB vo�n�ho miesta.%n%nNaozaj chcete pokra�ova�?
DirNameTooLong=N�zov prie�inku alebo zadan� cesta je pr�li� dlh�.
InvalidDirName=N�zov prie�inku je neplatn�.
BadDirName32=N�zov prie�inku nesmie obsahova� �iaden z nasleduj�cich znakov:%n%n%1
DirExistsTitle=Prie�inok s t�mto n�zvom u� existuje.
DirExists=Prie�inok %n%n%1%n%u� existuje. Chcete pokra�ova� v in�tal�cii?
DirDoesntExistTitle=Prie�inok s t�mto n�zvom neexistuje.
DirDoesntExist=Prie�inok %n%n%1%n%nneexistuje. Chcete ho vytvori�??

; *** "Select Components" wizard page
WizardSelectComponents=S��asti programu
SelectComponentsDesc=V�ber s��ast�, ktor� sa maj� in�talova�.
SelectComponentsLabel2=Zvo�te si typ in�tal�cie alebo vyberte s��asti programu, ktor� chcete nain�talova�. Ak chcete pokra�ova�, kliknite na tla�idlo �alej.
FullInstallation=�pln� in�tal�cia
; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Kompaktn� in�tal�cia
CustomInstallation=Vlastn� in�tal�cia
NoUninstallWarningTitle=T�to s��as� programu u� existuje.
NoUninstallWarning=In�tal�tor zistil, �e nasleduj�ce s��asti programu s� u� na va�om po��ta�i nain�talovan�:%n%n%1%n%nZru�te v�ber t�ch s��ast�, ktor� nechcete odin�talova�.%n%nChcete aj napriek tomu pokra�ova�?
ComponentSize1=%1 kB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=Po�adovan� miesto na disku: [mb] MB

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=�al�ie �lohy
SelectTasksDesc=Ak� �al�ie �lohy sa maj� vykona�?
SelectTasksLabel2=Vyberte �al�ie �lohy, ktor� sa maj� spolu s programom [name] nain�talova�. Ak chcete pokra�ova�, kliknite na tla�idlo �alej.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=In�tal�cia polo�ky ponuky �tart
SelectStartMenuFolderDesc=Kam chcete aby in�tal�tor vytvoril odkazy na vybrat� polo�ky?
SelectStartMenuFolderLabel3=In�tal�tor vytvor� odkazy na vybrat� polo�ky vo zvolenom prie�inku ponuky �tart.
SelectStartMenuFolderBrowseLabel=Ak chcete pokra�ova�, kliknite na tla�idlo �alej. Ak chcete vybra� in� prie�inok, kliknite na tla�idlo Preh�ad�va�.
NoIconsCheck=Nevytv�ra� �iadne ikony
MustEnterGroupName=Zadajte n�zov prie�inku.
GroupNameTooLong=N�zov prie�inku alebo zadan� cesta je pr�li� dlh�.
InvalidGroupName=N�zov prie�inku je neplatn�.
BadGroupName=N�zov prie�inku nesmie obsahova� �iaden z nasleduj�cich znakov:%n%n%1
NoProgramGroupCheck2=Nevytv�ra� polo�ky ponuky �tart

; *** "Ready to Install" wizard page
WizardReady=Pripraven� na in�tal�ciu
ReadyLabel1=In�tal�tor je teraz pripraven� na in�tal�ciu programu [name] na tento po��ta�.
ReadyLabel2a=V in�tal�cii pokra�ujte kliknut�m na tla�idlo In�talova�. Ak chcete skontrolova� alebo zmeni� ktor�ko�vek nastavenie, kliknite najsk�r na tla�idlo Sp�.
ReadyLabel2b=V in�tal�cii pokra�ujte kliknut�m na tla�idlo In�talova�.
ReadyMemoUserInfo=User information:
ReadyMemoDir=Cie�ov� umiestnenie:
ReadyMemoType=Typ in�tal�cie:
ReadyMemoComponents=Vybran� s��asti:
ReadyMemoGroup=Ponuka �tart:
ReadyMemoTasks=�al�ie �lohy:

; *** "Preparing to Install" wizard page
WizardPreparing=Pr�prava in�tal�cie
PreparingDesc=In�tal�tor pripravuje in�tal�ciu programu [name] na v� po��ta�.
PreviousInstallNotCompleted=In�tal�cia alebo odin�talovanie programu nebolo dokon�en�. Je potrebn� re�tartova� po��ta� na dokon�enie tejto oper�cie.%n%nPo re�tartovan� syst�mu je potrebn� znovu spusti� in�tal�ciu programu [name] a dokon�i� ju.
CannotContinue=In�tal�cia nem��e pokra�ova�. Kliknut�m na tla�idlo Zru�i�, ukon��te in�tal�ciu.

; *** "Installing" wizard page
WizardInstalling=In�tal�cia
InstallingLabel=Po�kajte, k�m in�tal�tor nain�taluje s��asti programu [name]. M��e to trva� nieko�ko min�t.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Dokon�uje sa in�tal�cia programu[name]
FinishedLabelNoIcons=In�tal�tor dokon�il in�tal�ciu programu [name].
FinishedLabel=In�tal�tor dokon�il in�tal�ciu programu [name]. Program spust�te pomocou vytvorenej ikony.
ClickFinish=In�tal�ciu programu ukon��te kliknut�m na tla�idlo Dokon�i�.
FinishedRestartLabel=In�tal�tor mus� re�tartova� po��ta�, aby mohol dokon�i� in�tal�ciu programu [name]. Chcete re�tartova� teraz?
FinishedRestartMessage=In�tal�tor mus� re�tartova� po��ta�, aby mohol dokon�i� in�tal�ciu programu [name].%n%nChcete re�tartova� teraz?
ShowReadmeCheck=�no, chcem zobrazi� s�bor readme.txt.
YesRadio=Re�tartova� teraz
NoRadio=Re�tartova� nesk�r
; used for example as 'Run MyProg.exe'
RunEntryExec=Spusti� program %1
; used for example as 'View Readme.txt'
RunEntryShellExec=Zobrazi� s�bor %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=In�tal�tor potrebuje �al�iu disketu (disk).
SelectDiskLabel2=Vlo�te disketu (disk) s n�zvom %1 a kliknite na tla�idlo OK.%n%n Ak sa s�bory nach�dzaj� na inom disku alebo prie�inku, kliknite na tla�idlo Preh�ad�va�.
PathLabel=Cesta:
FileNotInDir2=S�bor s n�zvom �%1� v �%2� neexistuje. Vlo�te spr�vny disk alebo vyberte in� prie�inok.
SelectDirectoryLabel=Zadajte umiestnenie �al�ej diskety (disku).

; *** Installation phase messages
SetupAborted=In�tal�cia nebola dokon�en�.%n%nAk chcete pokra�ova�, opravte tento probl�m.
EntryAbortRetryIgnore=Ak chcete oper�ciu zopakova�, kliknite na tla�idlo Znova. Ak chcete aj napriek tomu pokra�ova�, kliknite na tla�idlo Ignorova�. Ak ju chcete zru�i�, kliknite na tla�idlo Zru�i�.

; *** Installation status messages
StatusCreateDirs=Vytv�raj� sa prie�inky...
StatusExtractFiles=Extrahuj� sa s�bory...
StatusCreateIcons=Vytv�raj� sa odkazy...
StatusCreateIniEntries=Vytv�raj� sa INI s�bory...
StatusCreateRegistryEntries=Vytv�raj� sa k���e datab�zy Registry...
StatusRegisterFiles=Registr�cia s�borov...
StatusSavingUninstall=Ukladaj� sa �daje pre odin�talovanie...
StatusRunProgram=Dokon�uje sa in�tal�cia...
StatusRollback=Vr�tenie vykonan�ch zmien...

; *** Misc. errors
ErrorInternal2=Vn�torn� chyba: %1
ErrorFunctionFailedNoCode=%1 zlyhala
ErrorFunctionFailed=%1 zlyhala; k�d %2
ErrorFunctionFailedWithMessage=%1 zlyhala; k�d %2.%n%3
ErrorExecutingProgram=Nepodarilo sa spusti� s�bor:%n%1

; *** Registry errors
ErrorRegOpenKey=Chyba pri otv�ran� k���a datab�zy Registry:%n%1\%2
ErrorRegCreateKey=Chyba pri vytv�ran� k���a datab�zy Registry:%n%1\%2
ErrorRegWriteKey=Chyba pri zapisovan� k���a do datab�zy Registry:%n%1\%2

; *** INI errors
ErrorIniEntry=Pri vytv�ran� polo�ky INI v s�bore �%1� sa vyskytla chyba.

; *** File copying errors
FileAbortRetryIgnore=Ak chcete oper�ciu zopakova�, kliknite na tla�idlo Znova. Ak chcete aj napriek tomu pokra�ova�, kliknite na tla�idlo Ignorova�. Ak ju chcete zru�i�, kliknite na tla�idlo Zru�i�.
FileAbortRetryIgnore2=Ak chcete oper�ciu zopakova�, kliknite na tla�idlo Znova. Ak chcete aj napriek tomu pokra�ova�, kliknite na tla�idlo Ignorova� (neodpor��a sa). Ak ju chcete zru�i�, kliknite na tla�idlo Zru�i�.
SourceIsCorrupted=Zdrojov� s�bor je po�koden�.
SourceDoesntExist=Zdrojov� s�bor �%1� neexistuje.
ExistingFileReadOnly=Existuj�ci s�bor je ur�en� len na ��tanie..%n%nAk chcete odstr�ni� atrib�t �Len na ��tanie�, kliknite na tla�idlo Znova. Ak chcete vynecha� tento s�bor, kliknite na tla�idlo Ignorova�. Ak chcete in�tal�ciu zru�i�, kliknite na tla�idlo Zru�i�.
ErrorReadingExistingDest=Pri ��tan� existuj�ceho sa vyskytla chyba. N�zov s�boru:
FileExists=S�bor u� existuje.%n%nChcete ho prep�sa�?
ExistingFileNewer=Existuj�ci s�bor je nov�� ne� ten, ktor� chcete nain�talova�. Odpor��a sa ponecha� existuj�ci s�bor.%n%nChcete ponecha� existuj�ci s�bor?
ErrorChangingAttr=Pri pokuse o zmenu atrib�tov s�boru sa vyskytla chyba. N�zov s�boru:
ErrorCreatingTemp=Pri pokuse o vytvorenie s�boru v cie�ovom prie�inku sa vyskytla chyba. Cie�ov� prie�inok:
ErrorReadingSource=Pri na��tavan� zdrojov�ho s�boru sa vyskytla chyba. Zdrojov� s�bor:
ErrorCopying=Pri kop�rovan� s�boru sa vyskytla chyba. N�zov s�boru:
ErrorReplacingExistingFile=Pri pokuse o prep�sanie s�boru sa vyskytla chyba. N�zov s�boru:
ErrorRestartReplace=Funkcia in�tal�tora �RestartReplace� zlyhala:
ErrorRenamingTemp=Pri pokuse o premenovanie s�boru v cie�ovom prie�inku sa vyskytla chyba. Cie�ov� prie�inok:
ErrorRegisterServer=Ovl�dac� prvok DLL/OCX (%1) nie je mo�n� zaregistrova�.
ErrorRegisterServerMissingExport=Funkcia exportu DllRegisterServer sa nena�la.
ErrorRegisterTypeLib=Nepodarilo sa zaregistrova� kni�nicu typov: %1

; *** Post-installation errors
ErrorOpeningReadme=Pri pokuse o otvorenie s�boru �readme.txt� sa vyskytla chyba.
ErrorRestartingComputer=In�tal�tor nem��e re�tartova� po��ta�. Je potrebn� to urobi� ru�ne.

; *** Uninstaller messages
UninstallNotFound=S�bor �%1� neexistuje. Program sa ned� odin�talova�.
UninstallOpenError=S�bor �%1� sa ned� otvori�. Program sa ned� odin�talova�.
UninstallUnsupportedVer=S�bor denn�ka s inform�ciami o in�tal�cii programu �%1� nie je kompatibiln� s aktu�lnou verziou nain�talovan�ho in�tal�tora. In�tal�tor nem��e odin�talova� tento program.
UninstallUnknownEntry=V denn�ku s inform�ciami o in�tal�cii programu sa vyskytla chyba (%1).
ConfirmUninstall=Naozaj chcete �plne odstr�ni� program %1 a v�etky jeho s��asti?
OnlyAdminCanUninstall=Ak chcete tento program odin�talova� mus�te by� prihl�sen� ako pou��vate� Administr�tor.
UninstallStatusLabel=Po�kajte, pros�m, k�m sa dokon�� odin�talovanie programu %1 z v�ho po��ta�a.
UninstalledAll=Program %1 bol �spe�ne odstr�nen� z tohto po��ta�a.
UninstalledMost=Program %1 bol odstr�nen� z tohto po��ta�a.%n%nNiektor� s��asti sa nedali odstr�ni�. Je potrebn� ich odstr�ni� ru�ne.
UninstalledAndNeedsRestart=In�tal�tor mus� re�tartova� po��ta�, aby mohol dokon�i� odin�talovanie programu [name].%n%nChcete re�tartova� teraz?
UninstallDataCorrupted=S�bor �%1� je po�koden�. Program sa ned� odin�talova�.

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=Chcete odstr�ni� zdie�an� s�bor?
ConfirmDeleteSharedFile2=Nasleduj�ci zdie�an� s�bor sa pr�ve nepou��va �iadnym in�m programom. Chcete odstr�ni� tento zdie�an� s�bor?%n%nNiektor� moment�lne nespusten� programy v�ak po jeho odstr�nen� nemusia pracova� spr�vne. Ak si nie ste ist�, kliknite na tla�idlo Nie.
SharedFileNameLabel=N�zov s�boru:
SharedFileLocationLabel=Umiestnenie:
WizardUninstalling=Odin�talovanie
StatusUninstalling=Odin�talovanie programu %1...

; The custom messages below aren't used by Setup itself, but if you make
; use of them in your scripts, you'll want to translate them.

[CustomMessages]

NameAndVersion=%1 version %2
AdditionalIcons=�al�ie ikony:
CreateDesktopIcon=Vytvori� ikonu na pracovnej ploche
CreateQuickLaunchIcon=Vytvori� pre r�chle spustenie
ProgramOnTheWeb=%1 on the Web
UninstallProgram=Odin�talova� program %1
LaunchProgram=Spusti� program %1
AssocFileExtension=Pr�ponu s�boru %2 priradi� k programu %1
AssocingFileExtension=Prira�uje sa pr�pona s�boru %2 k programu %1...
