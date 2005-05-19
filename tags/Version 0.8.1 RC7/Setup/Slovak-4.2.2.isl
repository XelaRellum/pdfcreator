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
SetupAppTitle=Sprievodca in�tal�ciou
SetupWindowTitle=Sprievodca in�tal�ciou - %1
UninstallAppTitle=Sprievodca odin�tal�ciou
UninstallAppFullTitle=Sprievodca odin�tal�ciou - %1

; *** Misc. common
InformationTitle=Inform�cia
ConfirmTitle=Ot�zka
ErrorTitle=Chyba

; *** SetupLdr messages
SetupLdrStartupMessage=Toto je sprievodca in�tal�ciou produktu %1. Prajete si pokra�ova�?
LdrCannotCreateTemp=Ned� sa vytvori� do�asn� s�bor. Sprievodca in�tal�ciou bude ukon�en�
LdrCannotExecTemp=Ned� sa spusti� s�bor v do�asnej zlo�ke. Sprievodca in�tal�ciou bude ukon�en�

; *** Startup error messages
LastErrorMessage=%1.%n%nChyba %2: %3
SetupFileMissing=In�tala�n� zlo�ka neobsahuje s�bor %1. Opravte, pros�m, t�to chybu alebo si zaobstarajte nov� k�piu tohto produktu.
SetupFileCorrupt=S�bory sprievodcu in�tal�ciou s� po�koden�. Zaobstarajte si, pros�m, nov� k�piu tohto produktu.
SetupFileCorruptOrWrongVer=S�bory sprievodcu in�tal�ciou s� po�koden� alebo sa nezlu�uj� s touto verziou sprievodcu in�tal�ciou. Opravte, pros�m, t�to chybu alebo si zaobstarajte nov� k�piu tohto produktu.
NotOnThisPlatform=Tento produkt sa ned� spusti� pod %1.
OnlyOnThisPlatform=Tento produkt mus� by� spusten� pod %1.
WinVersionTooLowError=Tento produkt vy�aduje %1 verzie %2 alebo vy��iu.
WinVersionTooHighError=Tento produkt sa ned� nain�talova� v %1 verzie %2 alebo vy��ej
AdminPrivilegesRequired=K vykonaniu in�tal�cie tohto produktu mus�te by� prihl�sen�(�) ako administr�tor.
PowerUserPrivilegesRequired=K vykonaniu in�tal�cie tohto produktu mus�te by� prihl�sen�(�) ako administr�tor alebo �len skupiny Power Users.
SetupAppRunningError=Sprievodca in�tal�ciou zistil, �e %1 je teraz spusten�.%n%nUkon�ite, pros�m, v�etky spusten� in�tal�cie tohto produktu a klepnite na OK pre pokra�ovanie alebo na Storno pre ukon�enie.
UninstallAppRunningError=Sprievodca odin�tal�ciou zistil, �e %1 je teraz spusten�.%n%nUkon�ite, pros�m, v�etky spusten� in�tal�cie tohto produktu a klepnite na OK pre pokra�ovanie alebo na Storno pre ukon�enie.

; *** Misc. errors
ErrorCreatingDir=Sprievodca in�tal�ciou nemohol vytvori� zlo�ku "%1"
ErrorTooManyFilesInDir=Ned� sa vytvori� s�bor v zlo�ke "%1", preto�e t�to zlo�ka u� obsahuje pr�li� ve�a s�borov

; *** Setup common messages
ExitSetupTitle=Ukon�i� sprievodcu in�tal�ciou
ExitSetupMessage=In�talacia nebola �plne dokon�en�. Ak teraz ukon��te sprievodcu in�tal�ciou, produkt nebude nain�talovan�.%n%nSprievodcu in�tal�ciou m��ete znovu spusti� nesk�r a dokon�i� tak in�tal�ciu.%n%nUkon�i� sprievodcu in�tal�ciou?
AboutSetupMenuItem=&O sprievodcovi in�tal�ciou...
AboutSetupTitle=O sprievodcovi in�tal�ciou
AboutSetupMessage=%1 verzia %2%n%3%n%n%1 domovsk� str�nka:%n%4
AboutSetupNote=

; *** Buttons
ButtonBack=< &Sp�
ButtonNext=&�al�� >
ButtonInstall=&In�talova�
ButtonOK=OK
ButtonCancel=Storno
ButtonYes=&�no
ButtonYesToAll=�no &v�etk�m
ButtonNo=&Nie
ButtonNoToAll=N&ie v�etk�m
ButtonFinish=&Dokon�i�
ButtonBrowse=&Prech�dza�...
ButtonWizardBrowse=&Prech�dza�...
ButtonNewFolder=&Vytvoti� nov� zlo�ku

; *** Common wizard text
SelectLanguageTitle=Zvoli� jazyk sprievodcu in�tal�ciou
SelectLanguageLabel=Zvo�te jazyk, ktor� sa m� pou�i� pri in�tal�cii:
ClickNext=Klepnite na �al�� pre pokra�ovanie alebo na Storno pre ukon�enie sprievodcu in�tal�ciou.
BeveledLabel=
BrowseDialogTitle=Vyh�ada� zlo�ku	
BrowseDialogLabel=Z ni��ie uveden�ho zoznamu vyberte zlo�ku a klepnite na OK.	
NewFolderName=Nov� zlo�ka

; *** "Welcome" wizard page
WelcomeLabel1=V�ta V�s sprievodca in�tal�ciou produktu [name].
WelcomeLabel2=[name/ver] bude nain�talovan� na V� po��ta�.%n%nOdpor��a sa ukon�i� v�etky spusten� aplik�cie predt�m, ne� budete pokra�ova�.

; *** "Password" wizard page
WizardPassword=Heslo
PasswordLabel1=T�to in�tal�cia je chr�nen� heslom.
PasswordLabel3=Pros�m, zadajte heslo a klepnite na �al�� pre pokra�ovanie. Pri zad�van� hesla rozli�ujte mal� a ve�k� p�smen�.
PasswordEditLabel=&Heslo:
IncorrectPassword=Zadan� heslo nie je spr�vne. Pros�m, sk�ste to znovu.

; *** "License Agreement" wizard page
WizardLicense=Licen�n� dohoda
LicenseLabel=Pros�m, pre��tajte si pozorne tieto d�le�it� inform�cie predt�m, ne� budete pokra�ova�.
LicenseLabel3=Pros�m, pre��tajte si t�to Licen�n� dohodu. Mus�te s�hlasi� s podmienkami tejto dohody, aby mohol in�tala�n� proces pokra�ova�.
LicenseAccepted=&S�hlas�m s podmienkami Licen�nej dohody 
LicenseNotAccepted=&Nes�hlas�m s podmienkami Licen�nej dohody

; *** "Information" wizard pages
WizardInfoBefore=Inform�cie
InfoBeforeLabel=Pros�m, pre��tajte si pozorne tieto d�le�it� inform�cie predt�m, ne� budete pokra�ova�.
InfoBeforeClickLabel=Klepnite na �al�� pre pokra�ovanie in�tala�n�ho procesu.
WizardInfoAfter=Inform�cie
InfoAfterLabel=Pros�m, pre��tajte si pozorne tieto d�le�it� inform�cie predt�m, ne� budete pokra�ova�.
InfoAfterClickLabel=Klepnite na �al�� pre pokra�ovanie in�tala�n�ho procesu.

; *** "User Information" wizard page
WizardUserInfo=Inform�cie o u�ivate�ovi
UserInfoDesc=Pros�m, zadajte po�adovan� inform�cie.
UserInfoName=&U��vate�sk� meno:
UserInfoOrg=&Organiz�cia:
UserInfoSerial=&S�riov� ��slo:
UserInfoNameRequired=U��vate�sk� meno mus� by� zadan�.

; *** "Select Destination Directory" wizard page
WizardSelectDir=Zvo�te cie�ov� zlo�ku
SelectDirDesc=Kam m� by� [name] nain�talovan�?
SelectDirBrowseLabel=Klepnite na �al�� pre pokra�ovanie. Pokia� chcete zvoli� in� zlo�ku, klepnite na Prech�dza�.
SelectDirLabel3=[name] bude nain�talovan� do n�sleduj�cej zlo�ky.
;SelectDirLabel2=[name] bude nain�talovan� do n�sleduj�cej zlo�ky.%n%nKlepnite na �al�� pre pokra�ovanie.
;SelectDirLabel=Zvo�te zlo�ku, do ktorej m� by� [name] nain�talovan� a klepnite na �al��.
DiskSpaceMBLabel=Tento produkt vy�aduje najmenej [mb] MB miesta na disku.
ToUNCPathname=Sprievodca in�tal�ciou nem��e in�talova� do cesty UNC. Ak sa pok��ate in�talova� po sieti, mus�te pou�i� niektor� z dostupn�ch sie�ov�ch jednotiek.
InvalidPath=Mus�te zada� �pln� cestu vr�tane p�smena jednotky; napr�klad:%n%nC:\Aplik�cia%n%nalebo cestu UNC v tvare:%n%n\\server\zdie�an� zlo�ka
InvalidDrive=Vami zvolen� jednotka alebo cesta UNC neexistuje alebo nie je dostupn�. Pros�m, zvo�te in� umiestnenie.
DiskSpaceWarningTitle=Nedostatok miesta na disku
DiskSpaceWarning=Sprievodca in�tal�ciou vy�aduje najmenej %1 KB vo�n�ho miesta pre in�tal�ciu produktu, ale na zvolenej jednotke je dostupn�ch len %2 KB.%n%nPrajete si napriek tomu pokra�ova�?
InvalidDirName=Toto nie je platn� n�zov zlo�ky.
DirNameTooLong=N�zov zlo�ky alebo jej cesta je pr�li� dlh�.
BadDirName32=N�zvy zlo�iek nem��u obsahova� �iadny z nasleduj�cich znakov:%n%n%1
DirExistsTitle=Zlo�ka existuje
DirExists=Zlo�ka:%n%n%1%n%nu� existuje. M� sa napriek tomu in�talova� do tejto zlo�ky?
DirDoesntExistTitle=Zlo�ka neexistuje
DirDoesntExist=Zlo�ka:%n%n%1%n%nneexistuje. M� by� t�to zlo�ka vytvoren�?

; *** "Select Components" wizard page
WizardSelectComponents=Vyberte s��asti
SelectComponentsDesc=Ak� s��asti maj� by� nain�talovan�?
SelectComponentsLabel2=Za�krtnite s��asti, ktor� maj� by� nain�talovan�; s��asti, ktor� sa nemaj� in�talova�, ponechajte neza�krtnut�. Klepnite na �al�� pre pokra�ovanie.
FullInstallation=�pln� in�tal�cia

; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Kompaktn� in�tal�cia
CustomInstallation=Volite�n� in�tal�cia
NoUninstallWarningTitle=S��asti existuj�
NoUninstallWarning=Sprievodca in�tal�ciou zistil, �e nasleduj�ce s��asti s� u� na Va�om po��ta�i nain�talovan�:%n%n%1%n%nNeza�krtnutie t�chto s��ast� do v�beru sp�sob�, �e nebud� nesk�r odin�talovan�.%n%nPrajete si napriek tomu pokra�ova�?
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=Vybran� s��asti vy�aduj� najmenej [mb] MB miesta na disku.

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Zvo�te �al�ie �lohy
SelectTasksDesc=Ktor� �al�ie �lohy maj� by� vykonan�?
SelectTasksLabel2=Zvo�te �al�ie �lohy, ktor� maj� by� vykonan� v priebehu in�tal�cie produktu [name] a pokra�ujte klepnut�m na �al��.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=Vyberte zlo�ku v ponuke �tart
SelectStartMenuFolderDesc=Kam maj� by� sprievodcom in�tal�ciou umiestnen� z�stupci aplik�cie?
SelectStartMenuFolderBrowseLabel=Klepnite na �al�� pre pokra�ovanie. Pokia� chcete zvoli� in� zlo�ku, klepnite na Prech�dza�.
SelectStartMenuFolderLabel3=Z�stupci aplik�cie bud� vytvoren� v n�sleduj�cej zlo�ke ponuky �tart.
;SelectStartMenuFolderLabel2=Z�stupci aplik�cie bud� vytvoren� v n�sleduj�cej zlo�ke ponuky �tart.%n%nKlepnite na �al�� pre pokra�ovanie. Pokia� chcete zvoli� in� zlo�ku, klepnite na Prech�dza�.
;SelectStartMenuFolderLabel=Vyberte zlo�ku v ponuke �tart, do ktorej maj� by� sprievodcom in�tal�ciou umiestnen� z�stupci aplik�cie a pokra�ujte klepnut�m na �al��.
NoIconsCheck=&Nevytv�ra� �iadne ikony
MustEnterGroupName=Mus�te zada� n�zov zlo�ky.
InvalidGroupName=Toto nie je platn� n�zov zlo�ky.
GroupNameTooLong=N�zov zlo�ky alebo jej cesta je pr�li� dlh�.
BadGroupName=N�zov zlo�ky nem��e obsahova� �iadny z nasleduj�cich znakov:%n%n%1
NoProgramGroupCheck2=&Nevytv�ra� zlo�ku v ponuke �tart

; *** "Ready to Install" wizard page
WizardReady=In�tal�cia pripraven�
ReadyLabel1=Sprievodca in�tal�ciou je teraz pripraven� nain�talova� [name] na V� po��ta�.
ReadyLabel2a=Klepnite na In�talova� pre pokra�ovanie in�tala�n�ho procesu alebo klepnite na Sp�, pokia� si prajete zmeni� niektor� nastavenia in�tal�cie.
ReadyLabel2b=Klepnite na In�talova� pre pokra�ovanie in�tala�n�ho procesu.
ReadyMemoUserInfo=Inform�cie o u��vate�ovi:
ReadyMemoDir=Cie�ov� zlo�ka:
ReadyMemoType=Typ in�tal�cie:
ReadyMemoComponents=Vybran� s��asti:
ReadyMemoGroup=Zlo�ka v ponuke �tart:
ReadyMemoTasks=�al�ie �lohy:

; *** "Preparing to Install" wizard page
WizardPreparing=Pr�prava in�tal�cie
PreparingDesc=Sprievodca in�tal�ciou pripravuje in�tal�ciu produktu [name] na V� po��ta�.
PreviousInstallNotCompleted=Proces in�tal�cie/odin�tal�cie predch�dzaj�ceho produktu nebol �plne dokon�en�. Pre dokon�enie tohto procesu je nutn� re�tartova� tento po��ta�.%n%nPo vykonanom re�tarte po��ta�a spus�te znovu tohto sprievodcu in�tal�ciou pre dokon�enie in�tal�cie produktu [name].
CannotContinue=Sprievodca in�tal�ciou nem��e pokra�ova�. Pros�m, klepnite na Storno pre ukon�enie sprievodcu in�tal�ciou.

; *** "Installing" wizard page
WizardInstalling=In�talujem
InstallingLabel=�akajte pros�m, pokia� sprievodca in�tal�ciou nedokon�� in�tal�ciu produktu [name] na V� po��ta�.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Dokon�uje sa in�tal�cia produktu [name]
FinishedLabelNoIcons=Sprievodca in�tal�ciou dokon�il in�tal�ciu produktu [name] na V� po��ta�.
FinishedLabel=Sprievodca in�tal�ciou dokon�il in�tal�ciu produktu [name] na V� po��ta�. Produkt sa d� spusti� pomocou nain�talovan�ch ikon a z�stupcov.
ClickFinish=Klepnite na Dokon�i� pre ukon�enie sprievodcu in�tal�ciou.
FinishedRestartLabel=Pre dokon�enie in�tal�cie produktu [name] je nutn�, aby sprievodca in�tal�ciou re�tartoval V� po��ta�. Prajete si teraz re�tartova� V� po��ta�?
FinishedRestartMessage=Pre dokon�enie in�tal�cie produktu [name] je nutn�, aby sprievodca in�tal�ciou re�tartoval V� po��ta�.%n%nPrajete si teraz re�tartova� V� po��ta�?
ShowReadmeCheck=�no, chcem zobrazi� dokument "�TIMNE"
YesRadio=&�no, chcem teraz re�tartova� po��ta�
NoRadio=&Nie, po��ta� re�tartujem nesk�r

; used for example as 'Run MyProg.exe'
RunEntryExec=Spusti� %1

; used for example as 'View Readme.txt'
RunEntryShellExec=Zobrazi� %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=Sprievodca in�tal�ciou vy�aduje �al�� disk
;SelectDirectory=Vyberte zlo�ku
SelectDiskLabel2=Pros�m, vlo�te disk %1 a klepnite na OK.%n%nAk sa s�bory na tomto disku nach�dzaj� v inej zlo�ke, ne� v tej, ktor� je zobrazen� ni��ie, tak zadajte spr�vnu cestu alebo klepnite na Prech�dza�.
PathLabel=&Cesta:
FileNotInDir2=S�bor "%1" sa ned� n�js� v "%2". Pros�m vlo�te spr�vny disk alebo zvo�te in� zlo�ku.
SelectDirectoryLabel=Pros�m, �pecifikujte umiestnenie �al�ieho disku.

; *** Installation phase messages
SetupAborted=In�tal�cia nebola �plne dokon�en�.%n%nPros�m, opravte chybu a spus�te sprievodcu in�tal�ciou znovu.
EntryAbortRetryIgnore=Klepnite na Opakova� pre zopakovanie akcie, na Presko�i� pre vynechanie akcie alebo na Preru�i� pre stornovanie in�tal�cie.

; *** Installation status messages
StatusCreateDirs=Vytv�raj� sa zlo�ky...
StatusExtractFiles=Extrahuj� sa s�bory...
StatusCreateIcons=Vytv�raj� sa z�stupci...
StatusCreateIniEntries=Vytv�raj� sa z�znamy v konfigura�n�ch s�boroch...
StatusCreateRegistryEntries=Vytv�raj� sa z�znamy v syst�movom registri...
StatusRegisterFiles=Registruj� sa s�bory...
StatusSavingUninstall=Ukladaj� sa inform�cie nutn� pre neskor�iu odin�t�laciu produktu...
StatusRunProgram=Dokon�uje sa in�tal�cia...
StatusRollback=Prebieha sp�tn� vr�tenie v�etk�ch vykonan�ch zmien...

; *** Misc. errors
ErrorInternal2=Intern� chyba: %1
ErrorFunctionFailedNoCode=%1 zlyhala
ErrorFunctionFailed=%1 zlyhala; k�d %2
ErrorFunctionFailedWithMessage=%1 zlyhala; k�d %2.%n%3
ErrorExecutingProgram=Ned� sa spusti� s�bor:%n%1

; *** Registry errors
ErrorRegOpenKey=Do�lo k chybe pri otv�ran� k���a syst�mov�ho registra:%n%1\%2
ErrorRegCreateKey=Do�lo k chybe pri vytv�ran� k���a syst�mov�ho registra:%n%1\%2
ErrorRegWriteKey=Do�lo k chybe pri z�pise do k���a syst�mov�ho registra:%n%1\%2

; *** INI errors
ErrorIniEntry=Do�lo k chybe pri vytv�ran� z�znamu v konfigura�nom s�bore "%1".

; *** File copying errors
FileAbortRetryIgnore=Klepnite na Opakova� pre zopakovanie akcie, na Presko�i� pre presko�enie tohto s�boru (neodpor��a sa) alebo na Preru�i� pre stornovanie in�tal�cie.
FileAbortRetryIgnore2=Klepnite na Opakova� pre zopakovanie akcie, na Presko�i� pre pokra�ovanie (neodpor��a se) alebo na Preru�i� pre stornovanie in�tal�cie.
SourceIsCorrupted=Zdrojov� s�bor je po�koden�
SourceDoesntExist=Zdrojov� s�bor "%1" neexistuje
ExistingFileReadOnly=Existuj�ci s�bor je ur�en� len pre ��tanie.%n%nKlepnite na Opakova� pre odstr�nenie atrib�tu "len pre ��tanie" a zopakovanie akcie, na Presko�i� pre presko�enie tohto s�boru alebo na Preru�i� pre stornovanie in�tal�cie.
ErrorReadingExistingDest=Do�lo k chybe pri pokuse o ��tanie existuj�ceho s�boru:
FileExists=S�bor u� existuje.%n%nM� by� sprievodcom in�tal�ciou prep�san�?
ExistingFileNewer=Existuj�ci s�bor je nov�� ne� ten, ktor� sa sprievodca in�tal�ciou pok��a nain�talova�. Odpor��a s ponecha� existuj�ci s�bor.%n%nPrajete si ponecha� existuj�ci s�bor?
ErrorChangingAttr=Do�lo k chybe pri pokuse o modifik�ciu atrib�tov existuj�ceho s�boru:
ErrorCreatingTemp=Do�lo k chybe pri pokuse o vytvorenie s�boru v cie�ovej zlo�ke:
ErrorReadingSource=Do�lo k chybe pri pokuse o ��tanie zdrojov�ho s�boru:
ErrorCopying=Do�lo k chybe pri pokuse o skop�rovanie s�boru:
ErrorReplacingExistingFile=Do�lo k chybe pri pokuse o nahradenie existuj�ceho s�boru:
ErrorRestartReplace=Funkcia sprievodcu in�tal�ciou "RestartReplace" zlyhala:
ErrorRenamingTemp=Do�lo k chybe pri pokuse o premenovanie s�boru v cie�ovej zlo�ke:
ErrorRegisterServer=Ned� sa vykona� registr�ciu DLL/OCX: %1
ErrorRegisterServerMissingExport=Ned� sa n�js� export DllRegisterServer
ErrorRegisterTypeLib=Ned� sa vykona� registr�ciu typovej kni�nice: %1

; *** Post-installation errors
ErrorOpeningReadme=Do�lo k chybe pri pokuse o otvorenie dokumentu "�TIMNE".
ErrorRestartingComputer=Sprievodcovi in�tal�ciou sa nepodarilo re�tartova� V� po��ta�. Urobte to, pros�m, manu�lne.

; *** Uninstaller messages
UninstallNotFound=S�bor "%1" neexistuje. Produkt sa ned� odin�talova�.
UninstallOpenError=S�bor "%1" sa ned� otvori�. Produkt sa ned� odin�talova�.
UninstallUnsupportedVer=Sprievodcovi odin�tal�ciou sa nepodarilo rozpozna� form�t s�boru obsahuj�ceho inform�cie pre odin�tal�ciu produktu "%1". Produkt sa ned� odin�talova�
UninstallUnknownEntry=V s�bore obsahuj�com inform�cie pre odin�tal�ciu produktu bola zisten� nezn�ma polo�ka (%1)
ConfirmUninstall=Ste si naozaj ist�(�), �e chcete odin�talova� %1 a v�etky jeho s��asti?
OnlyAdminCanUninstall=K odin�talovaniu tohto produktu mus�te by� prihl�sen�(�) ako administr�tor.
UninstallStatusLabel=�akajte, pros�m, pokia� %1 nebude odin�talovan� z V�ho po��ta�a.
UninstalledAll=%1 bol �sp�ne odin�talovan� z V�ho po��ta�a.
UninstalledMost=%1 bol odin�talovan� z V�ho po��ta�a.%n%nNiektor� jeho s��asti sa v�ak nepodarilo odin�talova�. Tieto m��u by� odobran� manu�lne.
UninstalledAndNeedsRestart=Pre dokon�enie odin�tal�cie produktu %1 je nutn�, aby sprievodca odin�tal�ciou re�tartoval V� po��ta�.%n%nPrajete si teraz re�tartova� V� po��ta�?
UninstallDataCorrupted=S�bor "%1" je po�koden�. Produkt sa ned� odin�talova�

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=Odobra� zdie�an� s�bor?
ConfirmDeleteSharedFile2=Syst�m indikuje, �e nasleduj�ci zdie�an� s�bor nie je pou��van� �iadnymi in�mi aplik�ciami. M� by� tento zdie�an� s�bor sprievodcom odin�tal�ciou odstr�nen�?%n%nAk niektor�  aplik�ce tento s�bor pou��vaj�, potom po jeho odstranen� nemusia tieto aplik�cie pracova� spr�vne. Ak si nie ste ist�(�), zvo�te Nie. Ponechanie tohto s�boru vo Va�om  syst�me nesp�sob� �iadnu �kodu.
SharedFileNameLabel=N�zov s�boru:
SharedFileLocationLabel=Umiestnenie:
WizardUninstalling=Stav odin�tal�cie
StatusUninstalling=Odin�talov�vam %1...


[CustomMessages]

NameAndVersion=%1 verzia %2
AdditionalIcons=�al�� z�stupci:
CreateDesktopIcon=Vytvori� z�stupca na &ploche
CreateQuickLaunchIcon=Vytvori� z�stupca na panelu &Snadn� spustenie
ProgramOnTheWeb=Aplik�cia %1 na internete

UninstallProgram=Odinstalovat aplikaci %1
LaunchProgram=Spustit aplikaci %1
AssocFileExtension=Vytvo�it &asociaci mezi soubory typu %2 a aplikac� %1
AssocingFileExtension=Vytv��� se asociace mezi soubory typu %2 a aplikac� %1...