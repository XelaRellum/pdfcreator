; *** Inno Setup version 4.2.2+ Slovak messages ***
; Translated by: Ing. Michal Krempa
; Contact: Marek Istvanek (Marek.Istvanek@astra-zlin.cz)
;
; To download user-contributed translations of this file, go to:
;   http://www.jrsoftware.org/is3rdparty.php
;
; Note: When translating this text, do not add periods (.) to the end of
; messages that didn't have them already, because on those messages Inno
; Setup adds the periods automatically (appending a period would result in
; two periods being displayed).
;
; $jrsoftware: issrc/Files/Default.isl,v 1.32 2003/06/18 19:24:07 jr Exp $

[LangOptions]
LanguageName=Sloven<010D>ina
LanguageID=$041B
; If the language you are translating to requires special font faces or
; sizes, uncomment any of the following entries and change them accordingly.
;DialogFontName=MS Shell Dlg
;DialogFontSize=8
;DialogFontStandardHeight=13
;TitleFontName=Arial
;TitleFontSize=29
;WelcomeFontName=Verdana
;WelcomeFontSize=12
;CopyrightFontName=Arial
;CopyrightFontSize=8

[Messages]

; *** Application titles
SetupAppTitle=Sprievodca inštaláciou
SetupWindowTitle=Sprievodca inštaláciou - %1
UninstallAppTitle=Sprievodca odinštaláciou
UninstallAppFullTitle=Sprievodca odinštaláciou - %1

; *** Misc. common
InformationTitle=Informácia
ConfirmTitle=Otázka
ErrorTitle=Chyba

; *** SetupLdr messages
SetupLdrStartupMessage=Toto je sprievodca inštaláciou produktu %1. Prajete si pokraèova?
LdrCannotCreateTemp=Nedá sa vytvori doèasnı súbor. Sprievodca inštaláciou bude ukonèenı
LdrCannotExecTemp=Nedá sa spusti súbor v doèasnej zloke. Sprievodca inštaláciou bude ukonèenı

; *** Startup error messages
LastErrorMessage=%1.%n%nChyba %2: %3
SetupFileMissing=Inštalaèná zloka neobsahuje súbor %1. Opravte, prosím, túto chybu alebo si zaobstarajte novú kópiu tohto produktu.
SetupFileCorrupt=Súbory sprievodcu inštaláciou sú poškodené. Zaobstarajte si, prosím, novú kópiu tohto produktu.
SetupFileCorruptOrWrongVer=Súbory sprievodcu inštaláciou sú poškodené alebo sa nezluèujú s touto verziou sprievodcu inštaláciou. Opravte, prosím, túto chybu alebo si zaobstarajte novú kópiu tohto produktu.
NotOnThisPlatform=Tento produkt sa nedá spusti pod %1.
OnlyOnThisPlatform=Tento produkt musí by spustenı pod %1.
WinVersionTooLowError=Tento produkt vyaduje %1 verzie %2 alebo vyššiu.
WinVersionTooHighError=Tento produkt sa nedá nainštalova v %1 verzie %2 alebo vyššej
AdminPrivilegesRequired=K vykonaniu inštalácie tohto produktu musíte by prihlásenı(á) ako administrátor.
PowerUserPrivilegesRequired=K vykonaniu inštalácie tohto produktu musíte by prihlásenı(á) ako administrátor alebo èlen skupiny Power Users.
SetupAppRunningError=Sprievodca inštaláciou zistil, e %1 je teraz spustenı.%n%nUkonèite, prosím, všetky spustené inštalácie tohto produktu a klepnite na OK pre pokraèovanie alebo na Storno pre ukonèenie.
UninstallAppRunningError=Sprievodca odinštaláciou zistil, e %1 je teraz spustenı.%n%nUkonèite, prosím, všetky spustené inštalácie tohto produktu a klepnite na OK pre pokraèovanie alebo na Storno pre ukonèenie.

; *** Misc. errors
ErrorCreatingDir=Sprievodca inštaláciou nemohol vytvori zloku "%1"
ErrorTooManyFilesInDir=Nedá sa vytvori súbor v zloke "%1", pretoe táto zloka u obsahuje príliš ve¾a súborov

; *** Setup common messages
ExitSetupTitle=Ukonèi sprievodcu inštaláciou
ExitSetupMessage=Inštalacia nebola úplne dokonèená. Ak teraz ukonèíte sprievodcu inštaláciou, produkt nebude nainštalovanı.%n%nSprievodcu inštaláciou môete znovu spusti neskôr a dokonèi tak inštaláciu.%n%nUkonèi sprievodcu inštaláciou?
AboutSetupMenuItem=&O sprievodcovi inštaláciou...
AboutSetupTitle=O sprievodcovi inštaláciou
AboutSetupMessage=%1 verzia %2%n%3%n%n%1 domovská stránka:%n%4
AboutSetupNote=

; *** Buttons
ButtonBack=< &Spä
ButtonNext=&Ïalší >
ButtonInstall=&Inštalova
ButtonOK=OK
ButtonCancel=Storno
ButtonYes=&Áno
ButtonYesToAll=Áno &všetkım
ButtonNo=&Nie
ButtonNoToAll=N&ie všetkım
ButtonFinish=&Dokonèi
ButtonBrowse=&Prechádza...
ButtonWizardBrowse=&Prechádza...
ButtonNewFolder=&Vytvoti novú zloku

; *** Common wizard text
SelectLanguageTitle=Zvoli jazyk sprievodcu inštaláciou
SelectLanguageLabel=Zvo¾te jazyk, ktorı sa má poui pri inštalácii:
ClickNext=Klepnite na Ïalší pre pokraèovanie alebo na Storno pre ukonèenie sprievodcu inštaláciou.
BeveledLabel=
BrowseDialogTitle=Vyh¾ada zloku	
BrowseDialogLabel=Z nišie uvedeného zoznamu vyberte zloku a klepnite na OK.	
NewFolderName=Nová zloka

; *** "Welcome" wizard page
WelcomeLabel1=Víta Vás sprievodca inštaláciou produktu [name].
WelcomeLabel2=[name/ver] bude nainštalovanı na Váš poèítaè.%n%nOdporúèa sa ukonèi všetky spustené aplikácie predtım, ne budete pokraèova.

; *** "Password" wizard page
WizardPassword=Heslo
PasswordLabel1=Táto inštalácia je chránená heslom.
PasswordLabel3=Prosím, zadajte heslo a klepnite na Ïalší pre pokraèovanie. Pri zadávaní hesla rozlišujte malé a ve¾ké písmená.
PasswordEditLabel=&Heslo:
IncorrectPassword=Zadané heslo nie je správne. Prosím, skúste to znovu.

; *** "License Agreement" wizard page
WizardLicense=Licenèná dohoda
LicenseLabel=Prosím, preèítajte si pozorne tieto dôleité informácie predtım, ne budete pokraèova.
LicenseLabel3=Prosím, preèítajte si túto Licenènú dohodu. Musíte súhlasi s podmienkami tejto dohody, aby mohol inštalaènı proces pokraèova.
LicenseAccepted=&Súhlasím s podmienkami Licenènej dohody 
LicenseNotAccepted=&Nesúhlasím s podmienkami Licenènej dohody

; *** "Information" wizard pages
WizardInfoBefore=Informácie
InfoBeforeLabel=Prosím, preèítajte si pozorne tieto dôleité informácie predtım, ne budete pokraèova.
InfoBeforeClickLabel=Klepnite na Ïalší pre pokraèovanie inštalaèného procesu.
WizardInfoAfter=Informácie
InfoAfterLabel=Prosím, preèítajte si pozorne tieto dôleité informácie predtım, ne budete pokraèova.
InfoAfterClickLabel=Klepnite na Ïalší pre pokraèovanie inštalaèného procesu.

; *** "User Information" wizard page
WizardUserInfo=Informácie o uivate¾ovi
UserInfoDesc=Prosím, zadajte poadované informácie.
UserInfoName=&Uívate¾ské meno:
UserInfoOrg=&Organizácia:
UserInfoSerial=&Sériové èíslo:
UserInfoNameRequired=Uívate¾ské meno musí by zadané.

; *** "Select Destination Directory" wizard page
WizardSelectDir=Zvo¾te cie¾ovú zloku
SelectDirDesc=Kam má by [name] nainštalovanı?
SelectDirBrowseLabel=Klepnite na Ïalší pre pokraèovanie. Pokia¾ chcete zvoli inú zloku, klepnite na Prechádza.
SelectDirLabel3=[name] bude nainštalovanı do následujúcej zloky.
;SelectDirLabel2=[name] bude nainštalovanı do následujúcej zloky.%n%nKlepnite na Ïalší pre pokraèovanie.
;SelectDirLabel=Zvo¾te zloku, do ktorej má by [name] nainštalovanı a klepnite na Ïalší.
DiskSpaceMBLabel=Tento produkt vyaduje najmenej [mb] MB miesta na disku.
ToUNCPathname=Sprievodca inštaláciou nemôe inštalova do cesty UNC. Ak sa pokúšate inštalova po sieti, musíte poui niektorú z dostupnıch sieovıch jednotiek.
InvalidPath=Musíte zada úplnú cestu vrátane písmena jednotky; napríklad:%n%nC:\Aplikácia%n%nalebo cestu UNC v tvare:%n%n\\server\zdie¾aná zloka
InvalidDrive=Vami zvolená jednotka alebo cesta UNC neexistuje alebo nie je dostupná. Prosím, zvo¾te iné umiestnenie.
DiskSpaceWarningTitle=Nedostatok miesta na disku
DiskSpaceWarning=Sprievodca inštaláciou vyaduje najmenej %1 KB vo¾ného miesta pre inštaláciu produktu, ale na zvolenej jednotke je dostupnıch len %2 KB.%n%nPrajete si napriek tomu pokraèova?
InvalidDirName=Toto nie je platnı názov zloky.
DirNameTooLong=Názov zloky alebo jej cesta je príliš dlhá.
BadDirName32=Názvy zloiek nemôu obsahova iadny z nasledujúcich znakov:%n%n%1
DirExistsTitle=Zloka existuje
DirExists=Zloka:%n%n%1%n%nu existuje. Má sa napriek tomu inštalova do tejto zloky?
DirDoesntExistTitle=Zloka neexistuje
DirDoesntExist=Zloka:%n%n%1%n%nneexistuje. Má by táto zloka vytvorená?

; *** "Select Components" wizard page
WizardSelectComponents=Vyberte súèasti
SelectComponentsDesc=Aké súèasti majú by nainštalované?
SelectComponentsLabel2=Zaškrtnite súèasti, ktoré majú by nainštalované; súèasti, ktoré sa nemajú inštalova, ponechajte nezaškrtnuté. Klepnite na Ïalší pre pokraèovanie.
FullInstallation=Úplná inštalácia

; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Kompaktná inštalácia
CustomInstallation=Volite¾ná inštalácia
NoUninstallWarningTitle=Súèasti existujú
NoUninstallWarning=Sprievodca inštaláciou zistil, e nasledujúce súèasti sú u na Vašom poèítaèi nainštalované:%n%n%1%n%nNezaškrtnutie tıchto súèastí do vıberu spôsobí, e nebudú neskôr odinštalované.%n%nPrajete si napriek tomu pokraèova?
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=Vybrané súèasti vyadujú najmenej [mb] MB miesta na disku.

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Zvo¾te ïalšie úlohy
SelectTasksDesc=Ktoré ïalšie úlohy majú by vykonané?
SelectTasksLabel2=Zvo¾te ïalšie úlohy, ktoré majú by vykonané v priebehu inštalácie produktu [name] a pokraèujte klepnutím na Ïalší.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=Vyberte zloku v ponuke Štart
SelectStartMenuFolderDesc=Kam majú by sprievodcom inštaláciou umiestnení zástupci aplikácie?
SelectStartMenuFolderBrowseLabel=Klepnite na Ïalší pre pokraèovanie. Pokia¾ chcete zvoli inú zloku, klepnite na Prechádza.
SelectStartMenuFolderLabel3=Zástupci aplikácie budú vytvorené v následujúcej zloke ponuky Štart.
;SelectStartMenuFolderLabel2=Zástupci aplikácie budú vytvorené v následujúcej zloke ponuky Štart.%n%nKlepnite na Ïalší pre pokraèovanie. Pokia¾ chcete zvoli inú zloku, klepnite na Prechádza.
;SelectStartMenuFolderLabel=Vyberte zloku v ponuke Štart, do ktorej majú by sprievodcom inštaláciou umiestnení zástupci aplikácie a pokraèujte klepnutím na Ïalší.
NoIconsCheck=&Nevytvára iadne ikony
MustEnterGroupName=Musíte zada názov zloky.
InvalidGroupName=Toto nie je platnı názov zloky.
GroupNameTooLong=Názov zloky alebo jej cesta je príliš dlhá.
BadGroupName=Názov zloky nemôe obsahova iadny z nasledujúcich znakov:%n%n%1
NoProgramGroupCheck2=&Nevytvára zloku v ponuke Štart

; *** "Ready to Install" wizard page
WizardReady=Inštalácia pripravená
ReadyLabel1=Sprievodca inštaláciou je teraz pripravenı nainštalova [name] na Váš poèítaè.
ReadyLabel2a=Klepnite na Inštalova pre pokraèovanie inštalaèného procesu alebo klepnite na Spä, pokia¾ si prajete zmeni niektoré nastavenia inštalácie.
ReadyLabel2b=Klepnite na Inštalova pre pokraèovanie inštalaèného procesu.
ReadyMemoUserInfo=Informácie o uívate¾ovi:
ReadyMemoDir=Cie¾ová zloka:
ReadyMemoType=Typ inštalácie:
ReadyMemoComponents=Vybrané súèasti:
ReadyMemoGroup=Zloka v ponuke Štart:
ReadyMemoTasks=Ïalšie úlohy:

; *** "Preparing to Install" wizard page
WizardPreparing=Príprava inštalácie
PreparingDesc=Sprievodca inštaláciou pripravuje inštaláciu produktu [name] na Váš poèítaè.
PreviousInstallNotCompleted=Proces inštalácie/odinštalácie predchádzajúceho produktu nebol úplne dokonèenı. Pre dokonèenie tohto procesu je nutné reštartova tento poèítaè.%n%nPo vykonanom reštarte poèítaèa spuste znovu tohto sprievodcu inštaláciou pre dokonèenie inštalácie produktu [name].
CannotContinue=Sprievodca inštaláciou nemôe pokraèova. Prosím, klepnite na Storno pre ukonèenie sprievodcu inštaláciou.

; *** "Installing" wizard page
WizardInstalling=Inštalujem
InstallingLabel=Èakajte prosím, pokia¾ sprievodca inštaláciou nedokonèí inštaláciu produktu [name] na Váš poèítaè.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Dokonèuje sa inštalácia produktu [name]
FinishedLabelNoIcons=Sprievodca inštaláciou dokonèil inštaláciu produktu [name] na Váš poèítaè.
FinishedLabel=Sprievodca inštaláciou dokonèil inštaláciu produktu [name] na Váš poèítaè. Produkt sa dá spusti pomocou nainštalovanıch ikon a zástupcov.
ClickFinish=Klepnite na Dokonèi pre ukonèenie sprievodcu inštaláciou.
FinishedRestartLabel=Pre dokonèenie inštalácie produktu [name] je nutné, aby sprievodca inštaláciou reštartoval Váš poèítaè. Prajete si teraz reštartova Váš poèítaè?
FinishedRestartMessage=Pre dokonèenie inštalácie produktu [name] je nutné, aby sprievodca inštaláciou reštartoval Váš poèítaè.%n%nPrajete si teraz reštartova Váš poèítaè?
ShowReadmeCheck=Áno, chcem zobrazi dokument "ÈTIMNE"
YesRadio=&Áno, chcem teraz reštartova poèítaè
NoRadio=&Nie, poèítaè reštartujem neskôr

; used for example as 'Run MyProg.exe'
RunEntryExec=Spusti %1

; used for example as 'View Readme.txt'
RunEntryShellExec=Zobrazi %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=Sprievodca inštaláciou vyaduje ïalší disk
;SelectDirectory=Vyberte zloku
SelectDiskLabel2=Prosím, vlote disk %1 a klepnite na OK.%n%nAk sa súbory na tomto disku nachádzajú v inej zloke, ne v tej, ktorá je zobrazená nišie, tak zadajte správnu cestu alebo klepnite na Prechádza.
PathLabel=&Cesta:
FileNotInDir2=Súbor "%1" sa nedá nájs v "%2". Prosím vlote správny disk alebo zvo¾te inú zloku.
SelectDirectoryLabel=Prosím, špecifikujte umiestnenie ïalšieho disku.

; *** Installation phase messages
SetupAborted=Inštalácia nebola úplne dokonèená.%n%nProsím, opravte chybu a spuste sprievodcu inštaláciou znovu.
EntryAbortRetryIgnore=Klepnite na Opakova pre zopakovanie akcie, na Preskoèi pre vynechanie akcie alebo na Preruši pre stornovanie inštalácie.

; *** Installation status messages
StatusCreateDirs=Vytvárajú sa zloky...
StatusExtractFiles=Extrahujú sa súbory...
StatusCreateIcons=Vytvárajú sa zástupci...
StatusCreateIniEntries=Vytvárajú sa záznamy v konfiguraènıch súboroch...
StatusCreateRegistryEntries=Vytvárajú sa záznamy v systémovom registri...
StatusRegisterFiles=Registrujú sa súbory...
StatusSavingUninstall=Ukladajú sa informácie nutné pre neskoršiu odinštálaciu produktu...
StatusRunProgram=Dokonèuje sa inštalácia...
StatusRollback=Prebieha spätné vrátenie všetkıch vykonanıch zmien...

; *** Misc. errors
ErrorInternal2=Interná chyba: %1
ErrorFunctionFailedNoCode=%1 zlyhala
ErrorFunctionFailed=%1 zlyhala; kód %2
ErrorFunctionFailedWithMessage=%1 zlyhala; kód %2.%n%3
ErrorExecutingProgram=Nedá sa spusti súbor:%n%1

; *** Registry errors
ErrorRegOpenKey=Došlo k chybe pri otváraní k¾úèa systémového registra:%n%1\%2
ErrorRegCreateKey=Došlo k chybe pri vytváraní k¾úèa systémového registra:%n%1\%2
ErrorRegWriteKey=Došlo k chybe pri zápise do k¾úèa systémového registra:%n%1\%2

; *** INI errors
ErrorIniEntry=Došlo k chybe pri vytváraní záznamu v konfiguraènom súbore "%1".

; *** File copying errors
FileAbortRetryIgnore=Klepnite na Opakova pre zopakovanie akcie, na Preskoèi pre preskoèenie tohto súboru (neodporúèa sa) alebo na Preruši pre stornovanie inštalácie.
FileAbortRetryIgnore2=Klepnite na Opakova pre zopakovanie akcie, na Preskoèi pre pokraèovanie (neodporúèa se) alebo na Preruši pre stornovanie inštalácie.
SourceIsCorrupted=Zdrojovı súbor je poškodenı
SourceDoesntExist=Zdrojovı súbor "%1" neexistuje
ExistingFileReadOnly=Existujúci súbor je urèenı len pre èítanie.%n%nKlepnite na Opakova pre odstránenie atribútu "len pre èítanie" a zopakovanie akcie, na Preskoèi pre preskoèenie tohto súboru alebo na Preruši pre stornovanie inštalácie.
ErrorReadingExistingDest=Došlo k chybe pri pokuse o èítanie existujúceho súboru:
FileExists=Súbor u existuje.%n%nMá by sprievodcom inštaláciou prepísanı?
ExistingFileNewer=Existujúci súbor je novší ne ten, ktorı sa sprievodca inštaláciou pokúša nainštalova. Odporúèa s ponecha existujúci súbor.%n%nPrajete si ponecha existujúci súbor?
ErrorChangingAttr=Došlo k chybe pri pokuse o modifikáciu atribútov existujúceho súboru:
ErrorCreatingTemp=Došlo k chybe pri pokuse o vytvorenie súboru v cie¾ovej zloke:
ErrorReadingSource=Došlo k chybe pri pokuse o èítanie zdrojového súboru:
ErrorCopying=Došlo k chybe pri pokuse o skopírovanie súboru:
ErrorReplacingExistingFile=Došlo k chybe pri pokuse o nahradenie existujúceho súboru:
ErrorRestartReplace=Funkcia sprievodcu inštaláciou "RestartReplace" zlyhala:
ErrorRenamingTemp=Došlo k chybe pri pokuse o premenovanie súboru v cie¾ovej zloke:
ErrorRegisterServer=Nedá sa vykona registráciu DLL/OCX: %1
ErrorRegisterServerMissingExport=Nedá sa nájs export DllRegisterServer
ErrorRegisterTypeLib=Nedá sa vykona registráciu typovej kninice: %1

; *** Post-installation errors
ErrorOpeningReadme=Došlo k chybe pri pokuse o otvorenie dokumentu "ÈTIMNE".
ErrorRestartingComputer=Sprievodcovi inštaláciou sa nepodarilo reštartova Váš poèítaè. Urobte to, prosím, manuálne.

; *** Uninstaller messages
UninstallNotFound=Súbor "%1" neexistuje. Produkt sa nedá odinštalova.
UninstallOpenError=Súbor "%1" sa nedá otvori. Produkt sa nedá odinštalova.
UninstallUnsupportedVer=Sprievodcovi odinštaláciou sa nepodarilo rozpozna formát súboru obsahujúceho informácie pre odinštaláciu produktu "%1". Produkt sa nedá odinštalova
UninstallUnknownEntry=V súbore obsahujúcom informácie pre odinštaláciu produktu bola zistená neznáma poloka (%1)
ConfirmUninstall=Ste si naozaj istı(á), e chcete odinštalova %1 a všetky jeho súèasti?
OnlyAdminCanUninstall=K odinštalovaniu tohto produktu musíte by prihlásenı(á) ako administrátor.
UninstallStatusLabel=Èakajte, prosím, pokia¾ %1 nebude odinštalovanı z Vášho poèítaèa.
UninstalledAll=%1 bol úspìšne odinštalovanı z Vášho poèítaèa.
UninstalledMost=%1 bol odinštalovanı z Vášho poèítaèa.%n%nNiektoré jeho súèasti sa však nepodarilo odinštalova. Tieto môu by odobrané manuálne.
UninstalledAndNeedsRestart=Pre dokonèenie odinštalácie produktu %1 je nutné, aby sprievodca odinštaláciou reštartoval Váš poèítaè.%n%nPrajete si teraz reštartova Váš poèítaè?
UninstallDataCorrupted=Súbor "%1" je poškodenı. Produkt sa nedá odinštalova

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=Odobra zdie¾anı súbor?
ConfirmDeleteSharedFile2=Systém indikuje, e nasledujúci zdie¾anı súbor nie je pouívanı iadnymi inımi aplikáciami. Má by tento zdie¾anı súbor sprievodcom odinštaláciou odstránenı?%n%nAk niektoré  aplikáce tento súbor pouívajú, potom po jeho odstranení nemusia tieto aplikácie pracova správne. Ak si nie ste istı(á), zvo¾te Nie. Ponechanie tohto súboru vo Vašom  systéme nespôsobí iadnu škodu.
SharedFileNameLabel=Názov súboru:
SharedFileLocationLabel=Umiestnenie:
WizardUninstalling=Stav odinštalácie
StatusUninstalling=Odinštalovávam %1...


[CustomMessages]

NameAndVersion=%1 verzia %2
AdditionalIcons=Ïalší zástupci:
CreateDesktopIcon=Vytvori zástupca na &ploche
CreateQuickLaunchIcon=Vytvori zástupca na panelu &Snadné spustenie
ProgramOnTheWeb=Aplikácia %1 na internete

UninstallProgram=Odinstalovat aplikaci %1
LaunchProgram=Spustit aplikaci %1
AssocFileExtension=Vytvoøit &asociaci mezi soubory typu %2 a aplikací %1
AssocingFileExtension=Vytváøí se asociace mezi soubory typu %2 a aplikací %1...