; *** Inno Setup version 5.1.0+ Romanian messages (with diacrytics) ***
;
; Romanian translation:
; Perde Marius
; contact: emarius89@gmail.com
;
;
; $jrsoftware: issrc/Files/Default.isl,v 1.58 2004/04/07 20:17:13 jr Exp $

[LangOptions]
LanguageName=Romanian
LanguageID=$0418
LanguageCodePage=0
; If the language you are translating to requires special font faces or
; sizes, uncomment any of the following entries and change them accordingly.
DialogFontName=
DialogFontSize=8
WelcomeFontName=Verdana
WelcomeFontSize=12
TitleFontName=Arial
TitleFontSize=29
CopyrightFontName=Arial
CopyrightFontSize=8
[Messages]

; *** Application titles
SetupAppTitle=Instalare
SetupWindowTitle=Instalare - %1
UninstallAppTitle=Dezinstalare
UninstallAppFullTitle=Dezinstalare %1

; *** Misc. common
InformationTitle=Info
ConfirmTitle=Confirmare
ErrorTitle=Eroare

; *** SetupLdr messages
SetupLdrStartupMessage=Acesta este programul de instalare al %1. Dori�i s� continua�i ?
LdrCannotCreateTemp=Nu pot crea fi�ierele temporare. Instalarea se va incheia aici
LdrCannotExecTemp=Nu pot executa fi�ierul din directorul temporar. Instalarea se va incheia aici

; *** Startup error messages
LastErrorMessage=%1.%n%nEroare %2: %3
SetupFileMissing=Fi�ierul %1 lipseste din directorul de instalare. V� rug�m corecta�i problema sau ob�ine�i o nou� copie a programului.
SetupFileCorrupt=Fi�ierele de instalare a programului sunt corupte. V� rugam ob�ine�i o alt� copie a programului.
SetupFileCorruptOrWrongVer=Fi�ierele de instalare sunt corupte sau sunt incompatibile cu aceast� versiune a programului de instalare. V� rug�m corecta�i problema sau ob�ine�i o nou� copie a programului.
NotOnThisPlatform=Acest program nu ruleaz� sub %1.
OnlyOnThisPlatform=Acest program trebuie s� ruleze sub %1.
OnlyOnTheseArchitectures=Acest program poate fi instalat doar pe versiuni de Windows proiectate pentru urm�toarele arhitecturi de procesoare:%n%n%1
MissingWOW64APIs=Versiunea de Windows care ruleaz� nu include func�iile necesare programului pentru instalarea pe 64 de bi�i. Pentru a corecta aceast� problem�, v� rug�m s� instala�i Service Pack %1.
WinVersionTooLowError=Acest program necesit� %1 versiunea %2 sau ulterioar�.
WinVersionTooHighError=Acest program nu poate fi instalat sub %1 versiunea %2 sau ulterioar�.
AdminPrivilegesRequired=Trebuie s� ave�i drepturi de Administrator pentru a instala acest program.
PowerUserPrivilegesRequired=Trebuie s� ave�i drepturi de Administrator sau Power User pentru a instala acest program.
SetupAppRunningError=S-a detectat c� programul %1 ruleaz�.%n%nV� rug�m �nchide�i toate instan�ele, apoi ap�sa�i OK pentru a continua sau Anuleaz� pentru a p�r�si programul de instalare.
UninstallAppRunningError=S-a detectat c� programul %1 ruleaz�.%n%nV� rug�m �nchide�i toate instan�ele, apoi ap�sa�i OK pentru a continua sau Anuleaz� pentru a p�r�si programul de instalare.

; *** Misc. errors
ErrorCreatingDir=Nu pot crea directorul "%1"
ErrorTooManyFilesInDir=Nu pot crea un fi�ier �n directorul "%1" deoarece acesta con�ine prea multe fi�iere

; *** Setup common messages
ExitSetupTitle=Instalare
ExitSetupMessage=Procesul de instalare nu s-a �ncheiat. Dac� p�r�si�i programul acum, aplica�ia nu se va instala.%n%nPute�i rula ulterior programul de instalare pentru a finaliza procesul.%n%nParasiti instalarea ?
AboutSetupMenuItem=&Despre Setup...
AboutSetupTitle=Despre Setup
AboutSetupMessage=%1 versiunea %2%n%3%n%n%1 pe Internet:%n%4
AboutSetupNote=

; *** Buttons
TranslatorNote=Romanian translation:%nPerde Marius
ButtonBack=< &�napoi
ButtonNext=&Continu� >
ButtonInstall=&Instaleaz�
ButtonOK=OK
ButtonCancel=Anuleaz�
ButtonYes=&Da
ButtonYesToAll=Da tot &timpul
ButtonNo=&Nu
ButtonNoToAll=N&u tot timpul
ButtonFinish=&Finalizare
ButtonBrowse=&Selecteaz�...
ButtonWizardBrowse=&Selecteaz�...
ButtonNewFolder=Creeaz� un director &nou

; *** "Select Language" dialog messages
SelectLanguageTitle=Selectare limb�
SelectLanguageLabel=Selecta�i limba pe care dori�i s� o utilizati �n timpul instal�rii:

; *** Common wizard text
ClickNext=Ap�sa�i Continu� pentru pasul urm�tor sau Anuleaz� pentru a p�r�si programul.
BeveledLabel=
BrowseDialogTitle=Selectare director
BrowseDialogLabel=Selecteaz� un director din lista de mai jos, apoi apas� OK.
NewFolderName=New Folder

; *** "Welcome" wizard page
WelcomeLabel1=Bine a�i venit �n programul de instalare al [name].
WelcomeLabel2=Acesta va instala [name/ver] pe sistemul dumneavoastr�.%n%nEste recomandat s� �nchide�i toate celelalte aplica�ii care ruleaz� �n acest moment, �nainte de a continua.

; *** "Password" wizard page
WizardPassword=Protec�ie
PasswordLabel1=Aceast� instalare este protejat� de o parola.
PasswordLabel3=V� rug�m introduce�i parola, apoi ap�sa�i Continu�. Parola este case-sensitive.
PasswordEditLabel=&Parola:
IncorrectPassword=Parola introdus� este incorect�. Mai �ncerca�i o dat�.

; *** "License Agreement" wizard page
WizardLicense=Acceptul licen�ei de utilizare
LicenseLabel=V� rug�m s� citi�i urm�toarele informa�ii �nainte de a continua.
LicenseLabel3=V� rug�m s� citi�i urm�toarea Licen��. Este necesar s� accepta�i termenii acestei licen�e pentru a putea continua instalarea.
LicenseAccepted=&Accept termenii licen�ei
LicenseNotAccepted=&Nu accept termenii licen�ei

; *** "Information" wizard pages
WizardInfoBefore=Informa�ii
InfoBeforeLabel=V� rug�m s� citi�i aceste informa�ii suplimentare �nainte de a continua.
InfoBeforeClickLabel=C�nd sunte�i gata s� continua�i instalarea, ap�sa�i Continu�.
WizardInfoAfter=Informa�ii
InfoAfterLabel=V� rug�m s� citi�i aceste informa�ii suplimentare �nainte de a continua.
InfoAfterClickLabel=C�nd sunte�i gata s� continua�i instalarea, ap�sa�i Continu�.

; *** "User Information" wizard page
WizardUserInfo=Informa�ii despre utilizator
UserInfoDesc=V� rug�m introduce�i informa�iile despre utilizator.
UserInfoName=Nume &utilizator:
UserInfoOrg=&Organiza�ie:
UserInfoSerial=Num�r &serial:
UserInfoNameRequired=Trebuie s� introduce�i numele.

; *** "Select Destination Directory" wizard page
WizardSelectDir=Selecta�i directorul destina�ie
SelectDirDesc=Unde dori�i s� instala�i [name]?
SelectDirLabel3=Selecta�i directorul �n care dori�i s� instala�i [name], apoi ap�sa�i Continu�.
SelectDirBrowseLabel=Pentru a continua, ap�sa�i Continu�. Dac� dori�i s� selecta�i un alt director, ap�sa�i Selecteaz�.
DiskSpaceMBLabel=Acest program necesit� cel putin [mb] Mb pe disc.
ToUNCPathname=Programul nu poate fi instalat pe o cale de re�ea. Dac� dori�i instalarea pe o cale de re�ea, trebuie s� mapa�i calea la acea unitate de re�ea.
InvalidPath=Trebuie introdus� calea complet�, incluz�nd litera unit��ii.%n%nExemplu:%nC:\APP%n%nsau calea de re�ea de forma:%n\\server\share
InvalidDrive=Unitatea selectat� nu exist� sau nu este accesibil�. V� rug�m selecta�i o alt� unitate.
DiskSpaceWarningTitle=Nu exist� spa�iu suficient pe disc
DiskSpaceWarning=Programul de instalare necesit� un spa�iu minim de %1 KB, dar unitatea selectat� are disponibil doar %2 KB.%n%nDori�i s� continua�i ?
DirNameTooLong=Numele directorului sau calea este prea lung�.
InvalidDirName=Numele directorului nu este valid.
BadDirName32=Numele directorului nu poate con�ine nici unul din urmatoarele caractere: :%n%n%1
DirExistsTitle=Director existent
DirExists=Directorul:%n%n%1%n%ndeja exist�. Dori�i s� instala�i �n acest director ?
DirDoesntExistTitle=Director inexistent
DirDoesntExist=Directorul%n%n%1%n%nnu exist�. Dori�i s� fie creat ?

; *** "Select Components" wizard page
WizardSelectComponents=Selectare componente
SelectComponentsDesc=Care componente ar trebui instalate ?
SelectComponentsLabel2=Selecta�i componentele care dori�i s� fie instalate; deselecta�i componentele pe care nu dori�i s� le instala�i. Ap�sa�i Continu� c�nd sunte�i gata s� continua�i instalarea.
FullInstallation=Instalare complet�
; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Instalare compact�
CustomInstallation=Instalare personalizat�
NoUninstallWarningTitle=Componenta exist�
NoUninstallWarning=S-a detectat c� urmatoarele componente sunt deja instalate �n sistem:%n%n%1%n%nDeselectarea acestor componente nu va duce la dezinstalarea lor din sistem.%n%nDori�i s� continua�i ?
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=Selec�ia curent� necesit� cel putin [mb] MB.

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Op�iuni suplimentare
SelectTasksDesc=Care op�iuni suplimentare dori�i?
SelectTasksLabel2=Selecta�i op�iunile suplimentare dorite pentru instalarea [name], apoi ap�sa�i Continu�.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=Selecta�i directorul din meniul Start
SelectStartMenuFolderDesc=Unde dori�i s� adaug link-urile c�tre aplica�ie?
SelectStartMenuFolderLabel3=Programul de instalare va crea link-urile c�tre program �n urm�torul director din meniul Start.
SelectStartMenuFolderBrowseLabel=Pentru a continua, ap�sa�i Continu�. Dac� dori�i s� selecta�i un alt director, ap�sa�i Selecteaz�.
MustEnterGroupName=Trebuie s� introduce�i numele unui director.
GroupNameTooLong=Numele directorului sau calea este prea lung�.
InvalidGroupName=Numele directorului nu este valid.
BadGroupName=Numele directorului nu poate con�ine nici unul din urm�toarele caractere:%n%n%1
NoProgramGroupCheck2=Nu crea &director �n meniul Start

; *** "Ready to Install" wizard page
WizardReady=Gata de instalare
ReadyLabel1=Programul este �n punctul de a �ncepe instalarea [name] pe acest sistem.
ReadyLabel2a=Ap�sa�i Instaleaz� pentru a continua, sau �napoi dac� dori�i s� revede�i sau s� modifica�i set�rile f�cute anterior.
ReadyLabel2b=Ap�sa�i Instaleaz� pentru a continua.
ReadyMemoUserInfo=Informa�ii utilizator:
ReadyMemoDir=Director destina�ie:
ReadyMemoType=Tipul instal�rii:
ReadyMemoComponents=Componente selectate:
ReadyMemoGroup=Director �n meniu Start:
ReadyMemoTasks=Op�iuni suplimentare:

; *** "Preparing to Install" wizard page
WizardPreparing=Preg�tire instalare
PreparingDesc=Programul de instalare preg�te�te instalarea [name].
PreviousInstallNotCompleted=Instalarea / dezinstalarea versiunii anterioare a programului nu este complet�. Trebuie s� restarta�i sistemul pentru a termina instalarea anterioar�.%n%nDup� restart, rula�i �nc� o data programul de instalare al [name].
CannotContinue=Programul de instalare nu poate continua. Ap�sa�i Anuleaz� pentru a p�r�si instalarea.

; *** "Installing" wizard page
WizardInstalling=Instalare
InstallingLabel=V� rug�m a�tepta�i p�n� c�nd instalarea [name] ia sfar�it.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Finalizare instalare [name]
FinishedLabelNoIcons=Programul a terminat instalarea [name].
FinishedLabel=Programul a terminat instalarea [name]. Aplica�ia poate fi lansat� utiliz�nd link-urile create.
ClickFinish=Ap�sa�i Finalizare pentru a p�r�si programul de instalare.
FinishedRestartLabel=Pentru a completa instalarea [name], sistemul dvs. trebuie restartat. Dori�i s� restarta�i acum?
FinishedRestartMessage=Pentru a completa instalarea [name], sistemul dvs. trebuie restartat. %n%nDori�i s� restarta�i acum?
ShowReadmeCheck=Da, doresc s� citesc fisierul README
YesRadio=&Da, doresc restartarea sistemului acum
NoRadio=&Nu, voi restarta mai tarziu
; used for example as 'Run MyProg.exe'
RunEntryExec=Ruleaz� %1
; used for example as 'View Readme.txt'
RunEntryShellExec=Cite�te %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=Urm�torul disc
SelectDiskLabel2=V� rug�m introduce�i Discul %1 �i ap�sa�i OK.%n%nDac� fi�ierele de pe acest disc se afl� �ntr-un alt director dec�t cel afi�at mai jos, introduce�i calea corect� sau ap�sa�i Selecteaz�.
PathLabel=&Cale:
FileNotInDir2=Fi�ierul "%1" nu poate fi g�sit �n "%2". V� rug�m introduce�i discul corect sau selecta�i un alt director.
SelectDirectoryLabel=V� rug�m specifica�i loca�ia urm�torului disc.

; *** Installation phase messages
SetupAborted=Programul de instalare nu s-a �ncheiat cu succes.%n%nV� rug�m corecta�i problema �i porni�i instalarea din nou.
EntryAbortRetryIgnore=Ap�sa�i 'Retry' pentru a �ncerca �nc� o dat�, 'Ignore' pentru a trece oricum de acest pas sau 'Abort' pentru a opri instalarea.

; *** Installation status messages
StatusCreateDirs=Creare directoare ...
StatusExtractFiles=Extragere fi�iere ...
StatusCreateIcons=Creare link-uri ...
StatusCreateIniEntries=Creare intr�ri INI ...
StatusCreateRegistryEntries=Creare intr�ri �n Registry ...
StatusRegisterFiles=Inregistrare fi�iere ...
StatusSavingUninstall=Salvare informa�ii de dezinstalare ...
StatusRunProgram=Finalizare instalare ...
StatusRollback=Anulare modific�ri ...

; *** Misc. errors
ErrorInternal2=Eroare intern�: %1
ErrorFunctionFailedNoCode=%1 a e�uat
ErrorFunctionFailed=%1 a e�uat; cod %2
ErrorFunctionFailedWithMessage=%1 a e�uat; cod %2.%n%3
ErrorExecutingProgram=Nu pot executa:%n%1

; *** Registry errors
ErrorRegOpenKey=Eroare la deschiderea cheii din regi�tri:%n%1\%2
ErrorRegCreateKey=Eroare la crearea urm�toarei chei �n regi�tri:%n%1\%2
ErrorRegWriteKey=Eroare la scrierea urm�toarei chei �n regi�tri:%n%1\%2

; *** INI errors
ErrorIniEntry=Eroare la crearea �nregistr�rilor INI �n fi�ierul "%1".

; *** File copying errors
FileAbortRetryIgnore=Ap�sa�i 'Retry' pentru a �ncerca �nc� o dat�, 'Ignore' pentru a trece peste acest fi�ier (nerecomandat) sau 'Abort' pentru a opri instalarea.
FileAbortRetryIgnore2=Ap�sa�i 'Retry' pentru a �ncerca �nc� o dat�, 'Ignore' pentru a trece oricum de acest pas (nerecomandat) sau 'Abort' pentru a opri instalarea.
SourceIsCorrupted=Fi�ierul surs� este corupt
SourceDoesntExist=Fi�ierul surs� "%1" nu exist�
ExistingFileReadOnly=Fi�ierul existent este marcat read-only.%n%nAp�sa�i 'Retry' pentru a schimba atributele fi�ierului �i a �ncerca �nc� o dat�, 'Ignore' pentru a trece peste acest fi�ier sau 'Abort' pentru a opri instalarea.
ErrorReadingExistingDest=A ap�rut o eroare �n timp ce citeam fi�ierul:
FileExists=Fi�ierul exist�.%n%nDori�i s� �l suprascrie�i ?
ExistingFileNewer=Fi�ierul existent este mai nou dec�t cel care se instaleaz� acum. Este recomandat s� p�stra�i fi�ierul existent.%n%nDori�i s� p�stra�i fisierul existent ?
ErrorChangingAttr=A ap�rut o eroare �n timp ce �ncercam s� modific atributele fi�ierului:
ErrorCreatingTemp=A ap�rut o eroare �n timp ce �ncercam s� creez un fi�ier �n directorul destina�ie:
ErrorReadingSource=A ap�rut o eroare �n timp ce �ncercam s� citesc fi�ierul surs�:
ErrorCopying=A ap�rut o eroare �n timp ce �ncercam s� copiez fi�ierul:
ErrorReplacingExistingFile=A aparut o eroare in timp ce incercam sa inlocuiesc fisierul:
ErrorRestartReplace=Eroare �nlocuire la restart:
ErrorRenamingTemp=A ap�rut o eroare �n timp ce �ncercam s� redenumesc fi�ierul din directorul destina�ie:
ErrorRegisterServer=Nu pot s� �nregistrez DLL/OCX: %1
ErrorRegisterServerMissingExport=Nu pot g�si DllRegisterServer
ErrorRegisterTypeLib=Nu pot s� �nregistrez tipul de libr�rie: %1

; *** Post-installation errors
ErrorOpeningReadme=A ap�rut o eroare la deschiderea fi�ierului README.
ErrorRestartingComputer=Programul de instalare nu poate restarta sistemul. V� rug�m s� �ncerca�i s� restarta�i manual sistemul.

; *** Uninstaller messages
UninstallNotFound=Fi�ierul "%1" nu exist�. Nu pot dezinstala.
UninstallOpenError=Fi�ierul "%1" nu poate fi deschis. Nu pot dezinstala
UninstallUnsupportedVer=Fi�ierul de dezinstalare "%1" are un format necunoscut acestei versiuni de program. Nu pot dezinstala
UninstallUnknownEntry=O intrare necunoscut� (%1) a fost g�sit� �n fi�ierul de dezinstalare
ConfirmUninstall=Dori�i s� dezinstala�i %1 �i componentele sale adi�ionale?
UninstallOnlyOnWin64=Trebuie s� rula�i o versiune de Windows pe 64 de bi�i pentru a dezinstala acest program.
OnlyAdminCanUninstall=Trebuie s� ave�i drepturi de Administrator pentru a dezinstala acest program.
UninstallStatusLabel=V� rug�m a�tepta�i p�n� c�nd dezinstalarea %1 ia sfar�it.
UninstalledAll=%1 a fost dezinstalat cu succes.
UninstalledMost=%1 a fost dezinstalat.%n%nUnele fi�iere nu au putut fi �terse. Acestea pot fi �terse manual.
UninstalledAndNeedsRestart=Pentru a termina dezinstalarea %1, sistemul trebuie restartat.%n%nDori�i s� restarta�i sistemul acum?
UninstallDataCorrupted=Fi�ierul "%1" este corupt. Nu pot dezinstala

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=�tergere fi�iere 'Shared' ?
ConfirmDeleteSharedFile2=Sistemul indic� faptul c� urm�torul fi�ier nu mai este utilizat de nici un alt program. Dori�i s� �terge�i acest fi�ier ?%n%nDac� acest fi�ier este totu�i utilizat de un alt program, acesta din urm� nu va mai func�iona corect. Dac� nu sunte�i sigur, alege�i 'Nu'. L�s�nd fi�ierul pe sistem nu v� va afecta cu nimic.
SharedFileNameLabel=Nume fi�ier:
SharedFileLocationLabel=Loca�ie:
WizardUninstalling=Progres dezinstalare
StatusUninstalling=Dezinstalare %1...
[CustomMessages]

NameAndVersion=%1 versiunea %2
AdditionalIcons=Iconi�e adi�ionale:
CreateDesktopIcon=Creeaz� o iconi�� pe &desktop
CreateQuickLaunchIcon=Creeaz� o iconi�� &Quick Launch
ProgramOnTheWeb=%1 pe Internet
UninstallProgram=Dezinstalare %1
LaunchProgram=Lanseaz� %1
AssocFileExtension=&Asociaz� %1 cu extensia de fi�iere %2
AssocingFileExtension=Asociere %1 cu extensia de fi�iere %2 ...

