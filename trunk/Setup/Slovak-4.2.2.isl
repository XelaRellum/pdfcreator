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
SetupAppTitle=Sprievodca inštaláciou
SetupWindowTitle=Sprievodca inštaláciou - %1
UninstallAppTitle=Odinštalova
UninstallAppFullTitle=Odinštalova - %1

; *** Misc. common
InformationTitle=Informácia
ConfirmTitle=Potvrdenie
ErrorTitle=Chyba

; *** SetupLdr messages
SetupLdrStartupMessage=Chystáte sa nainštalova program %1. Chcete pokraèova?
LdrCannotCreateTemp=Nie je moné vytvori doèasnı súbor . Inštalácia ukonèená.
LdrCannotExecTemp=Nie je moné spusti súbor v doèasnom prieèinku. Inštalácia ukonèená.

; *** Startup error messages
LastErrorMessage=%1.%n%nChyba %2: %3
SetupFileMissing=V inštalaènom prieèinku chıba súbor %1. Ak chcete pokraèova, opravte tento problém alebo poiadajte o novú kópiu programu.
SetupFileCorrupt=Inštalaènı súbor je poškodenı. Poiadajte o novú verziu programu.
SetupFileCorruptOrWrongVer=Inštalaènı súbor je poškodenı alebo nekompatibilnı s aktuálnou verziou inštalátora. Ak chcete pokraèova, opravte tento problém alebo poiadajte o novú kópiu programu.
NotOnThisPlatform=Tento program sa na %1 nedá spusti.
OnlyOnThisPlatform=Tento program sa dá spusti len na %1.
WinVersionTooLowError=Program vyaduje %1 verzia %2 alebo novšia.
WinVersionTooHighError=Tento program nie je moné nainštalova na %1 verzie %2 alebo novšej.
AdminPrivilegesRequired=Ak chcete pokraèova v inštalácii musíte by prihlásení ako pouívate¾ Administrátor.
PowerUserPrivilegesRequired=Ak chcete pokraèova v inštalácii musíte by prihlásení ako pouívate¾ Administrátor alebo by skupiny Power Users.
SetupAppRunningError=Inštalátor zistil, e program %1 je práve spustenı.%n%nUkonèite všetky spustené aplikácie. Ak chcete pokraèova, kliknite na tlaèidlo Ïalej. Kliknutím na tlaèidlo Zruši inštaláciu ukonèíte.
UninstallAppRunningError=Inštalátor zistil, e program %1 je práve spustenı.%n%nUkonèite všetky spustené aplikácie. Ak chcete pokraèova, kliknite na tlaèidlo Ïalej. Kliknutím na tlaèidlo Zruši inštaláciu ukonèíte.

; *** Misc. errors
ErrorCreatingDir=Inštalátor nemohol vytvori prieèinok „%1“.
ErrorTooManyFilesInDir=Inštalátor nemohol vytvori súbor v prieèinku „%1“, pretoe obsahuje príliš ve¾a súborov.

; *** Setup common messages
ExitSetupTitle=Ukonèenie inštalácie
ExitSetupMessage=Inštalácia nie je dokonèená. Ak ju teraz ukonèíte, program nebude nainštalovanı.%n%nInštalátor môete spusti neskôr a inštaláciu dokonèi.%n%nChcete naozaj skonèi inštaláciu?
AboutSetupMenuItem=Èo je inštalaènı program...
AboutSetupTitle=Èo je inštalaènı program...
AboutSetupMessage=%1 verzia %2%n%3%n%n%1, domovská stránka:%n%4
AboutSetupNote=

; *** Buttons
ButtonBack=< Naspä
ButtonNext=Ïalej >
ButtonInstall=&Inštalova
ButtonOK=OK
ButtonCancel=Zruši
ButtonYes=Áno
ButtonYesToAll=Áno pre všetky
ButtonNo=&Nie
ButtonNoToAll=Nie pre všetky
ButtonFinish=&Dokonèi
ButtonBrowse=Preh¾adáva...
ButtonWizardBrowse=Preh¾adáva...
ButtonNewFolder=Vytvori novı prieèinok

; *** "Select Language" dialog messages
SelectLanguageTitle=Vıber jazyka
SelectLanguageLabel=Vyberte jazyk, ktorı chcete pouíva poèas inštalácie:

; *** Common wizard text
ClickNext=Ak chcete pokraèova, kliknite na tlaèidlo Ïalej. Kliknutím na tlaèidlo Zruši inštaláciu ukonèíte.
BeveledLabel=
BrowseDialogTitle=Vıber prieèinka programu
BrowseDialogLabel=V nasledujúcom zozname vyberte prieèinok a kliknite na tlaèidlo OK.
NewFolderName=Novı prieèinok

; *** "Welcome" wizard page
WelcomeLabel1=Víta vás Sprievodca inštaláciou programu [name].
WelcomeLabel2=Chystáte sa nainštalova program [name/ver] na váš poèítaè.%n%nSkôr ako budete pokraèova, odporúèa sa ukonèi všetky ostatné aplikácie.

; *** "Password" wizard page
WizardPassword=Heslo
PasswordLabel1=Inštalácia je chránená heslom.
PasswordLabel3=Zadajte heslo a pokraèujte v inštalácii kliknutím na tlaèidlo Ïalej. Rozlišujte ve¾ké a malé písmená.
PasswordEditLabel=Heslo:
IncorrectPassword=Zadané heslo nie je správne. Skúste to znova.

; *** "License Agreement" wizard page
WizardLicense=Licenèná zmluva
LicenseLabel=Preèítajte si tieto dôleité informácie, pred zaèatím inštalácie.
LicenseLabel3=Preèítajte si túto Licenènú zmluvu. Ak chcete pokraèova v inštalácii, musíte súhlasi so zmluvou.
LicenseAccepted=Súhlasím so zmluvou
LicenseNotAccepted=Nesúhlasím so zmluvou

; *** "Information" wizard pages
WizardInfoBefore=Informácia
InfoBeforeLabel=Preèítajte si tieto dôleité informácie, pred zaèatím inštalácie.
InfoBeforeClickLabel=Ak chcete pokraèova, kliknite na tlaèidlo Ïalej.
WizardInfoAfter=Informácia
InfoAfterLabel=Preèítajte si tieto dôleité informácie, pred zaèatím inštalácie.
InfoAfterClickLabel=Ak chcete pokraèova, kliknite na tlaèidlo Ïalej.

; *** "User Information" wizard page
WizardUserInfo=Informácie o pouívate¾ovi
UserInfoDesc=Zadajte informácie o pouívate¾ovi.
UserInfoName=Meno pouívate¾a:
UserInfoOrg=Organizácia:
UserInfoSerial=Sériové èíslo:
UserInfoNameRequired=Musíte zada meno pouívate¾a.

; *** "Select Destination Location" wizard page
WizardSelectDir=Umiestnenie programu
SelectDirDesc=Zadajte cestu k umiestneniu, kam chcete nainštalova program [name].
SelectDirLabel3=Program [name] sa nainštaluje do nasledujúceho prieèinku.
SelectDirBrowseLabel=Ak chcete pokraèova, kliknite na tlaèidlo Ïalej. Ak chcete vybra inı prieèinok, kliknite na tlaèidlo Preh¾adáva.
DiskSpaceMBLabel=Poadované miesto na disku: [mb] MB
ToUNCPathname=Inštalátor nemôe poui zadanú cestu UNC. Ak sa pokúšate nainštalova tento program v sieti, pouite niektorú z dostupnıch sieovıch jednotiek.
InvalidPath=Zadajte úplnú cestu spolu s písmenom jednotky (písmeno:\cesta) alebo úplnú cestu spolu so znakom \\ na konci bez názvu súboru (\\server\\zdie¾anie).
InvalidDrive=Zadané zariadenie alebo cesta UNC neexistuje alebo je odpojená. Vyberte iné zariadenie alebo cestu.
DiskSpaceWarningTitle=Na disku nie je dos miesta.
DiskSpaceWarning=Na dokonèenie inštalácie je potrebnıch minimálne %1 kB vo¾ného miesta na disku, zvolená jednotka obsahuje len %2 kB vo¾ného miesta.%n%nNaozaj chcete pokraèova?
DirNameTooLong=Názov prieèinku alebo zadaná cesta je príliš dlhá.
InvalidDirName=Názov prieèinku je neplatnı.
BadDirName32=Názov prieèinku nesmie obsahova iaden z nasledujúcich znakov:%n%n%1
DirExistsTitle=Prieèinok s tımto názvom u existuje.
DirExists=Prieèinok %n%n%1%n%u existuje. Chcete pokraèova v inštalácii?
DirDoesntExistTitle=Prieèinok s tımto názvom neexistuje.
DirDoesntExist=Prieèinok %n%n%1%n%nneexistuje. Chcete ho vytvori??

; *** "Select Components" wizard page
WizardSelectComponents=Súèasti programu
SelectComponentsDesc=Vıber súèastí, ktoré sa majú inštalova.
SelectComponentsLabel2=Zvo¾te si typ inštalácie alebo vyberte súèasti programu, ktoré chcete nainštalova. Ak chcete pokraèova, kliknite na tlaèidlo Ïalej.
FullInstallation=Úplná inštalácia
; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Kompaktná inštalácia
CustomInstallation=Vlastná inštalácia
NoUninstallWarningTitle=Táto súèas programu u existuje.
NoUninstallWarning=Inštalátor zistil, e nasledujúce súèasti programu sú u na vašom poèítaèi nainštalované:%n%n%1%n%nZrušte vıber tıch súèastí, ktoré nechcete odinštalova.%n%nChcete aj napriek tomu pokraèova?
ComponentSize1=%1 kB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=Poadované miesto na disku: [mb] MB

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Ïalšie úlohy
SelectTasksDesc=Aké ïalšie úlohy sa majú vykona?
SelectTasksLabel2=Vyberte ïalšie úlohy, ktoré sa majú spolu s programom [name] nainštalova. Ak chcete pokraèova, kliknite na tlaèidlo Ïalej.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=Inštalácia poloky ponuky Štart
SelectStartMenuFolderDesc=Kam chcete aby inštalátor vytvoril odkazy na vybraté poloky?
SelectStartMenuFolderLabel3=Inštalátor vytvorí odkazy na vybraté poloky vo zvolenom prieèinku ponuky Štart.
SelectStartMenuFolderBrowseLabel=Ak chcete pokraèova, kliknite na tlaèidlo Ïalej. Ak chcete vybra inı prieèinok, kliknite na tlaèidlo Preh¾adáva.
NoIconsCheck=Nevytvára iadne ikony
MustEnterGroupName=Zadajte názov prieèinku.
GroupNameTooLong=Názov prieèinku alebo zadaná cesta je príliš dlhá.
InvalidGroupName=Názov prieèinku je neplatnı.
BadGroupName=Názov prieèinku nesmie obsahova iaden z nasledujúcich znakov:%n%n%1
NoProgramGroupCheck2=Nevytvára poloky ponuky Štart

; *** "Ready to Install" wizard page
WizardReady=Pripravenı na inštaláciu
ReadyLabel1=Inštalátor je teraz pripravenı na inštaláciu programu [name] na tento poèítaè.
ReadyLabel2a=V inštalácii pokraèujte kliknutím na tlaèidlo Inštalova. Ak chcete skontrolova alebo zmeni ktoréko¾vek nastavenie, kliknite najskôr na tlaèidlo Spä.
ReadyLabel2b=V inštalácii pokraèujte kliknutím na tlaèidlo Inštalova.
ReadyMemoUserInfo=User information:
ReadyMemoDir=Cie¾ové umiestnenie:
ReadyMemoType=Typ inštalácie:
ReadyMemoComponents=Vybrané súèasti:
ReadyMemoGroup=Ponuka Štart:
ReadyMemoTasks=Ïalšie úlohy:

; *** "Preparing to Install" wizard page
WizardPreparing=Príprava inštalácie
PreparingDesc=Inštalátor pripravuje inštaláciu programu [name] na váš poèítaè.
PreviousInstallNotCompleted=Inštalácia alebo odinštalovanie programu nebolo dokonèené. Je potrebné reštartova poèítaè na dokonèenie tejto operácie.%n%nPo reštartovaní systému je potrebné znovu spusti inštaláciu programu [name] a dokonèi ju.
CannotContinue=Inštalácia nemôe pokraèova. Kliknutím na tlaèidlo Zruši, ukonèíte inštaláciu.

; *** "Installing" wizard page
WizardInstalling=Inštalácia
InstallingLabel=Poèkajte, kım inštalátor nainštaluje súèasti programu [name]. Môe to trva nieko¾ko minút.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Dokonèuje sa inštalácia programu[name]
FinishedLabelNoIcons=Inštalátor dokonèil inštaláciu programu [name].
FinishedLabel=Inštalátor dokonèil inštaláciu programu [name]. Program spustíte pomocou vytvorenej ikony.
ClickFinish=Inštaláciu programu ukonèíte kliknutím na tlaèidlo Dokonèi.
FinishedRestartLabel=Inštalátor musí reštartova poèítaè, aby mohol dokonèi inštaláciu programu [name]. Chcete reštartova teraz?
FinishedRestartMessage=Inštalátor musí reštartova poèítaè, aby mohol dokonèi inštaláciu programu [name].%n%nChcete reštartova teraz?
ShowReadmeCheck=Áno, chcem zobrazi súbor readme.txt.
YesRadio=Reštartova teraz
NoRadio=Reštartova neskôr
; used for example as 'Run MyProg.exe'
RunEntryExec=Spusti program %1
; used for example as 'View Readme.txt'
RunEntryShellExec=Zobrazi súbor %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=Inštalátor potrebuje ïalšiu disketu (disk).
SelectDiskLabel2=Vlote disketu (disk) s názvom %1 a kliknite na tlaèidlo OK.%n%n Ak sa súbory nachádzajú na inom disku alebo prieèinku, kliknite na tlaèidlo Preh¾adáva.
PathLabel=Cesta:
FileNotInDir2=Súbor s názvom „%1“ v „%2“ neexistuje. Vlote správny disk alebo vyberte inı prieèinok.
SelectDirectoryLabel=Zadajte umiestnenie ïalšej diskety (disku).

; *** Installation phase messages
SetupAborted=Inštalácia nebola dokonèená.%n%nAk chcete pokraèova, opravte tento problém.
EntryAbortRetryIgnore=Ak chcete operáciu zopakova, kliknite na tlaèidlo Znova. Ak chcete aj napriek tomu pokraèova, kliknite na tlaèidlo Ignorova. Ak ju chcete zruši, kliknite na tlaèidlo Zruši.

; *** Installation status messages
StatusCreateDirs=Vytvárajú sa prieèinky...
StatusExtractFiles=Extrahujú sa súbory...
StatusCreateIcons=Vytvárajú sa odkazy...
StatusCreateIniEntries=Vytvárajú sa INI súbory...
StatusCreateRegistryEntries=Vytvárajú sa k¾úèe databázy Registry...
StatusRegisterFiles=Registrácia súborov...
StatusSavingUninstall=Ukladajú sa údaje pre odinštalovanie...
StatusRunProgram=Dokonèuje sa inštalácia...
StatusRollback=Vrátenie vykonanıch zmien...

; *** Misc. errors
ErrorInternal2=Vnútorná chyba: %1
ErrorFunctionFailedNoCode=%1 zlyhala
ErrorFunctionFailed=%1 zlyhala; kód %2
ErrorFunctionFailedWithMessage=%1 zlyhala; kód %2.%n%3
ErrorExecutingProgram=Nepodarilo sa spusti súbor:%n%1

; *** Registry errors
ErrorRegOpenKey=Chyba pri otváraní k¾úèa databázy Registry:%n%1\%2
ErrorRegCreateKey=Chyba pri vytváraní k¾úèa databázy Registry:%n%1\%2
ErrorRegWriteKey=Chyba pri zapisovaní k¾úèa do databázy Registry:%n%1\%2

; *** INI errors
ErrorIniEntry=Pri vytváraní poloky INI v súbore „%1“ sa vyskytla chyba.

; *** File copying errors
FileAbortRetryIgnore=Ak chcete operáciu zopakova, kliknite na tlaèidlo Znova. Ak chcete aj napriek tomu pokraèova, kliknite na tlaèidlo Ignorova. Ak ju chcete zruši, kliknite na tlaèidlo Zruši.
FileAbortRetryIgnore2=Ak chcete operáciu zopakova, kliknite na tlaèidlo Znova. Ak chcete aj napriek tomu pokraèova, kliknite na tlaèidlo Ignorova (neodporúèa sa). Ak ju chcete zruši, kliknite na tlaèidlo Zruši.
SourceIsCorrupted=Zdrojovı súbor je poškodenı.
SourceDoesntExist=Zdrojovı súbor „%1“ neexistuje.
ExistingFileReadOnly=Existujúci súbor je urèenı len na èítanie..%n%nAk chcete odstráni atribút „Len na èítanie“, kliknite na tlaèidlo Znova. Ak chcete vynecha tento súbor, kliknite na tlaèidlo Ignorova. Ak chcete inštaláciu zruši, kliknite na tlaèidlo Zruši.
ErrorReadingExistingDest=Pri èítaní existujúceho sa vyskytla chyba. Názov súboru:
FileExists=Súbor u existuje.%n%nChcete ho prepísa?
ExistingFileNewer=Existujúci súbor je novší ne ten, ktorı chcete nainštalova. Odporúèa sa ponecha existujúci súbor.%n%nChcete ponecha existujúci súbor?
ErrorChangingAttr=Pri pokuse o zmenu atribútov súboru sa vyskytla chyba. Názov súboru:
ErrorCreatingTemp=Pri pokuse o vytvorenie súboru v cie¾ovom prieèinku sa vyskytla chyba. Cie¾ovı prieèinok:
ErrorReadingSource=Pri naèítavaní zdrojového súboru sa vyskytla chyba. Zdrojovı súbor:
ErrorCopying=Pri kopírovaní súboru sa vyskytla chyba. Názov súboru:
ErrorReplacingExistingFile=Pri pokuse o prepísanie súboru sa vyskytla chyba. Názov súboru:
ErrorRestartReplace=Funkcia inštalátora „RestartReplace“ zlyhala:
ErrorRenamingTemp=Pri pokuse o premenovanie súboru v cie¾ovom prieèinku sa vyskytla chyba. Cie¾ovı prieèinok:
ErrorRegisterServer=Ovládací prvok DLL/OCX (%1) nie je moné zaregistrova.
ErrorRegisterServerMissingExport=Funkcia exportu DllRegisterServer sa nenašla.
ErrorRegisterTypeLib=Nepodarilo sa zaregistrova kninicu typov: %1

; *** Post-installation errors
ErrorOpeningReadme=Pri pokuse o otvorenie súboru „readme.txt“ sa vyskytla chyba.
ErrorRestartingComputer=Inštalátor nemôe reštartova poèítaè. Je potrebné to urobi ruène.

; *** Uninstaller messages
UninstallNotFound=Súbor „%1“ neexistuje. Program sa nedá odinštalova.
UninstallOpenError=Súbor „%1“ sa nedá otvori. Program sa nedá odinštalova.
UninstallUnsupportedVer=Súbor denníka s informáciami o inštalácii programu „%1“ nie je kompatibilnı s aktuálnou verziou nainštalovaného inštalátora. Inštalátor nemôe odinštalova tento program.
UninstallUnknownEntry=V denníku s informáciami o inštalácii programu sa vyskytla chyba (%1).
ConfirmUninstall=Naozaj chcete úplne odstráni program %1 a všetky jeho súèasti?
OnlyAdminCanUninstall=Ak chcete tento program odinštalova musíte by prihlásení ako pouívate¾ Administrátor.
UninstallStatusLabel=Poèkajte, prosím, kım sa dokonèí odinštalovanie programu %1 z vášho poèítaèa.
UninstalledAll=Program %1 bol úspešne odstránenı z tohto poèítaèa.
UninstalledMost=Program %1 bol odstránenı z tohto poèítaèa.%n%nNiektoré súèasti sa nedali odstráni. Je potrebné ich odstráni ruène.
UninstalledAndNeedsRestart=Inštalátor musí reštartova poèítaè, aby mohol dokonèi odinštalovanie programu [name].%n%nChcete reštartova teraz?
UninstallDataCorrupted=Súbor „%1“ je poškodenı. Program sa nedá odinštalova.

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=Chcete odstráni zdie¾anı súbor?
ConfirmDeleteSharedFile2=Nasledujúci zdie¾anı súbor sa práve nepouíva iadnym inım programom. Chcete odstráni tento zdie¾anı súbor?%n%nNiektoré momentálne nespustené programy však po jeho odstránení nemusia pracova správne. Ak si nie ste istí, kliknite na tlaèidlo Nie.
SharedFileNameLabel=Názov súboru:
SharedFileLocationLabel=Umiestnenie:
WizardUninstalling=Odinštalovanie
StatusUninstalling=Odinštalovanie programu %1...

; The custom messages below aren't used by Setup itself, but if you make
; use of them in your scripts, you'll want to translate them.

[CustomMessages]

NameAndVersion=%1 version %2
AdditionalIcons=Ïalšie ikony:
CreateDesktopIcon=Vytvori ikonu na pracovnej ploche
CreateQuickLaunchIcon=Vytvori pre rıchle spustenie
ProgramOnTheWeb=%1 on the Web
UninstallProgram=Odinštalova program %1
LaunchProgram=Spusti program %1
AssocFileExtension=Príponu súboru %2 priradi k programu %1
AssocingFileExtension=Priraïuje sa prípona súboru %2 k programu %1...
