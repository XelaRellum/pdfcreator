; *** Inno Setup version 5.1.0+ Spanish (Standard) messages ***
;
; To download user-contributed translations of this file, go to:
;   http://www.jrsoftware.org/is3rdparty.php
;
; Note: When translating this text, do not add periods (.) to the end of
; messages that didn't have them already, because on those messages Inno
; Setup adds the periods automatically (appending a period would result in
; two periods being displayed).
;
; Translated by �ngel Mart�n
; e-mail: amartin@gawab.com
;

[LangOptions]
LanguageName=Espa<00F1>ol
LanguageID=$040a
LanguageCodePage=1252
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
SetupAppTitle=Instalaci�n
SetupWindowTitle=Instalaci�n - %1
UninstallAppTitle=Desinstalaci�n
UninstallAppFullTitle=Desinstalar %1

; *** Misc. common
InformationTitle=Informaci�n
ConfirmTitle=Confirmaci�n
ErrorTitle=Error

; *** SetupLdr messages
SetupLdrStartupMessage=Este programa instalar� %1. �Desea continuar?
LdrCannotCreateTemp=No se ha podido crear un archivo temporal. Instalaci�n cancelada
LdrCannotExecTemp=No se ha podido ejecutar el archivo en el directorio temporal. Instalaci�n cancelada

; *** Startup error messages
LastErrorMessage=%1.%n%nError %2: %3
SetupFileMissing=No se encuentra el archivo %1 en la carpeta de instalaci�n. Por favor, corrija el problema u obtenga una nueva copia del programa.
SetupFileCorrupt=Los archivos de instalaci�n est�n da�ados. Por favor, obtenga una nueva copia del programa.
SetupFileCorruptOrWrongVer=Los archivos de instalaci�n est�n da�ados, o son incompatibles con �sta versi�n de la instalaci�n. Por favor, corrija el problema u obtenga una nueva copia del programa.
NotOnThisPlatform=Este programa no funcionar� en %1.
OnlyOnThisPlatform=Este programa debe ejecutarse en %1.
OnlyOnTheseArchitectures=Este programa s�lo puede ser instalado en versiones de Windows dise�adas para las siguientes arquitecturas de procesador:%n%n%1
MissingWOW64APIs=La versi�n de Windows que est� usando no incluye la funcionalidad necesaria para realizar una instalaci�n de 64 bits. Para corregir este problema, por favor, instale el Service Pack %1.
WinVersionTooLowError=Este programa requiere %1 versi�n %2 o posterior.
WinVersionTooHighError=Este programa no puede ser instalado en %1 versi�n %2 o posterior.
AdminPrivilegesRequired=Debe iniciar la sesi�n como administrador para instalar este programa.
PowerUserPrivilegesRequired=Debe iniciar la sesi�n como administrador o miembro del grupo Usuarios Avanzados para instalar este programa.
SetupAppRunningError=El programa de instalaci�n ha detectado que %1 se est� ejecutando actualmente.%n%nPor favor, ci�rrelo y luego haga clic en Aceptar para continuar, o Cancelar para salir.
UninstallAppRunningError=El programa de desinstalaci�n ha detectado que %1 se est� ejecutando actualmente.%n%nPor favor, ci�rrelo y luego haga clic en Aceptar para continuar, o Cancelar para salir.

; *** Misc. errors
ErrorCreatingDir=El programa de instalaci�n no ha podido crear la carpeta "%1"
ErrorTooManyFilesInDir=No se ha podido crear un archivo en la carpeta "%1" porque contiene demasiados archivos.

; *** Setup common messages
ExitSetupTitle=Salir de la Instalaci�n
ExitSetupMessage=La instalaci�n no se ha completado. Si abandona ahora, el programa no quedar� instalado.%n%nPara completarla, podr� ejecutar de nuevo el programa de instalaci�n en otro momento.%n%n�Salir de la Instalaci�n?
AboutSetupMenuItem=&Acerca de la Instalaci�n...
AboutSetupTitle=Acerca de la Instalaci�n
AboutSetupMessage=%1 versi�n %2%n%3%n%nP�gina web de %1:%n%4
AboutSetupNote=
TranslatorNote=Spanish (Standard) translation by �ngel Mart�n (amartin@gawab.com)

; *** Buttons
ButtonBack=< &Atr�s
ButtonNext=&Siguiente >
ButtonInstall=&Instalar
ButtonOK=Aceptar
ButtonCancel=Cancelar
ButtonYes=&S�
ButtonYesToAll=S� a &todo
ButtonNo=&No
ButtonNoToAll=N&o a todo
ButtonFinish=&Terminar
ButtonBrowse=&Examinar...
ButtonWizardBrowse=&Examinar...
ButtonNewFolder=C&rear nueva carpeta...

; *** "Select Language" dialog messages
SelectLanguageTitle=Elija el idioma de instalaci�n
SelectLanguageLabel=Elija el idioma a usar durante la instalaci�n:

; *** Common wizard text
ClickNext=Haga clic en Siguiente para continuar, o en Cancelar para abandonar la instalaci�n.
BeveledLabel=
BrowseDialogTitle=Buscar carpeta
BrowseDialogLabel=Elija una carpeta de la lista, y haga clic en Aceptar.
NewFolderName=Nueva carpeta

; *** "Welcome" wizard page
WelcomeLabel1=Bienvenido al asistente de instalaci�n de [name].
WelcomeLabel2=Este programa instalar� [name/ver] en su equipo.%n%nEs recomendable que cierre el resto de aplicaciones antes de continuar.

; *** "Password" wizard page
WizardPassword=Contrase�a
PasswordLabel1=Esta instalaci�n est� protegida con una contrase�a.
PasswordLabel3=Por favor, introduzca su contrase�a y haga clic en Siguiente para continuar. La contrase�a distingue entre may�sculas y min�sculas.
PasswordEditLabel=&Contrase�a:
IncorrectPassword=La contrase�a introducida no es correcta. Por favor, int�ntelo de nuevo.

; *** "License Agreement" wizard page
WizardLicense=Acuerdo de Licencia
LicenseLabel=Por favor, lea cuidadosamente la siguiente informaci�n antes de continuar.
LicenseLabel3=Por favor, lea cuidadosamente el siguiente acuerdo de licencia. Debe de aceptar los t�rminos de este acuerdo para continuar con la instalaci�n.
LicenseAccepted=A&cepto el acuerdo
LicenseNotAccepted=&No acepto el acuerdo

; *** "Information" wizard pages
WizardInfoBefore=Informaci�n
InfoBeforeLabel=Por favor, lea la siguiente informaci�n antes de continuar.
InfoBeforeClickLabel=Cuando est� listo para continuar con la instalaci�n, haga clic en Siguiente.
WizardInfoAfter=Informaci�n
InfoAfterLabel=Por favor, lea la siguiente informaci�n antes de continuar.
InfoAfterClickLabel=Cuando est� listo para continuar con la instalaci�n, haga clic en Siguiente.

; *** "User Information" wizard page
WizardUserInfo=Informaci�n sobre el usuario
UserInfoDesc=Por favor, introduzca su informaci�n.
UserInfoName=Nombre del &usuario:
UserInfoOrg=&Organizaci�n:
UserInfoSerial=N�mero de &serie:
UserInfoNameRequired=Debe introducir un nombre.

; *** "Select Destination Location" wizard page
WizardSelectDir=Elija la Carpeta de Destino
SelectDirDesc=�D�nde debe instalarse [name]?
SelectDirLabel3=El programa instalar� [name] en la siguiente carpeta.
SelectDirBrowseLabel=Para continuar, haga clic en Siguiente. Si desea elegir una carpeta distinta, haga clic en Examinar.
DiskSpaceMBLabel=Se necesita un m�nimo de [mb] MB de espacio libre en la unidad.
ToUNCPathname=El programa de instalaci�n no puede instalar en un directorio UNC. Si est� tratando de instalar en una red, necesitar� mapear una unidad de red.
InvalidPath=Debe introducir una ruta completa con la letra de unidad; por ejemplo:%n%nC:\Aplicaci�n%n%no una ruta UNC de la siguiente forma:%n%n\\servidor\compartido
InvalidDrive=La unidad o ruta UNC que seleccion� no existe o no es accesible. Por favor, elija otra.
DiskSpaceWarningTitle=No hay suficiente espacio en el disco
DiskSpaceWarning=El programa de instalaci�n necesita al menos %1 KB de espacio libre para la instalaci�n, pero la unidad seleccionada solamente tiene %2 KB disponibles.%n%n�Desea continuar?
DirNameTooLong=El nombre de la carpeta o de la ruta es demasiado largo.
InvalidDirName=El nombre de la carpeta no es v�lido.
BadDirName32=Los nombres de carpeta no pueden incluir ninguno de los siguientes caracteres:%n%n%1
DirExistsTitle=La carpeta existe
DirExists=La carpeta:%n%n%1%n%nya existe. �Desea instalar en dicha carpeta de todas formas?
DirDoesntExistTitle=La carpeta no existe
DirDoesntExist=La carpeta:%n%n%1%n%nno existe. �Desea que dicha carpeta sea creada?

; *** "Select Components" wizard page
WizardSelectComponents=Selecci�n de Componentes
SelectComponentsDesc=�Qu� componentes deben instalarse?
SelectComponentsLabel2=Seleccione los componentes a instalar; desmarque los componentes que no desea instalar. Haga clic en Siguiente cuando est� preparado para continuar.
FullInstallation=Instalaci�n Completa
; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Instalaci�n Compacta
CustomInstallation=Instalaci�n Personalizada
NoUninstallWarningTitle=Componentes Existentes
NoUninstallWarning=El programa de instalaci�n ha detectado que los siguientes componentes est�n instalados en su equipo:%n%n%1%n%nSi estos componentes no est�n seleccionados no ser�n desinstalados.%n%n�Desea continuar de todos modos?
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=La selecci�n actual requiere un m�nimo de [mb] MB de espacio en disco.

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Seleccione Tareas Adicionales
SelectTasksDesc=�Qu� tareas adicionales deben realizarse?
SelectTasksLabel2=Elija las tareas adicionales que desea que se realicen durante la instalaci�n de [name], y despu�s haga clic en Siguiente.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=Seleccione la carpeta del Men� Inicio
SelectStartMenuFolderDesc=�D�nde deben ubicarse los accesos directos del programa?
SelectStartMenuFolderLabel3=El programa de instalaci�n crear� los accesos directos de programa en la siguiente carpeta del Men� Inicio.
SelectStartMenuFolderBrowseLabel=Para continuar, haga clic en Siguiente. Si desea elegir una carpeta distinta, haga clic en Examinar.
MustEnterGroupName=Debe introducir un nombre de carpeta.
GroupNameTooLong=El nombre de la carpeta o su ruta es demasiado largo.
InvalidGroupName=El nombre de la carpeta no es v�lido.
BadGroupName=El nombre de la carpeta no puede contener ninguno de los siguientes caracteres:%n%n%1
NoProgramGroupCheck2=&No crear una carpeta en el Men� Inicio

; *** "Ready to Install" wizard page
WizardReady=Preparado para Instalar
ReadyLabel1=El programa de instalaci�n est� preparado para comenzar la instalaci�n de [name] en su equipo.
ReadyLabel2a=Haga clic en Instalar para continuar con la instalaci�n, o Atr�s si desea revisar o cambiar las opciones de instalaci�n.
ReadyLabel2b=Haga clic en Instalar para continuar con la instalaci�n.
ReadyMemoUserInfo=Informaci�n del usuario:
ReadyMemoDir=Carpeta de destino:
ReadyMemoType=Tipo de instalaci�n:
ReadyMemoComponents=Componentes seleccionados:
ReadyMemoGroup=Carpeta del Men� Inicio:
ReadyMemoTasks=Tareas adicionales:

; *** "Preparing to Install" wizard page
WizardPreparing=Prepar�ndose para Instalar
PreparingDesc=El programa se est� preparando para instalar [name] en su equipo.
PreviousInstallNotCompleted=La instalaci�n/desinstalaci�n previa de un programa no se complet�. Necesitar� reiniciar el equipo para completar esa instalaci�n.%n%nDespu�s de reiniciar el equipo, ejecute �ste programa de nuevo para completar la instalaci�n de [name].
CannotContinue=El programa de instalaci�n no puede continuar. Por favor, haga clic en Cancelar para salir.

; *** "Installing" wizard page
WizardInstalling=Instalando
InstallingLabel=Por favor, espere mientras se instala [name] en su equipo.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Completando el asistente de instalaci�n de [name]
FinishedLabelNoIcons=El programa ha finalizado la instalaci�n de [name] en su equipo.
FinishedLabel=El programa ha finalizado la instalaci�n de [name] en su equipo. Puede iniciar la aplicaci�n seleccionando los iconos instalados.
ClickFinish=Haga clic en Terminar para salir de la instalaci�n.
FinishedRestartLabel=Para completar la instalaci�n de [name], el programa de instalaci�n debe reiniciar su equipo. �Desea reiniciar ahora?
FinishedRestartMessage=Para completar la instalaci�n de [name], el programa de instalaci�n debe reiniciar su equipo.%n%n�Desea reiniciar ahora?
ShowReadmeCheck=S�, deseo ver el archivo L�AME.
YesRadio=&S�, deseo reiniciar el equipo ahora
NoRadio=&No, reiniciar� el equipo m�s tarde
; used for example as 'Run MyProg.exe'
RunEntryExec=Ejecutar %1
; used for example as 'View Readme.txt'
RunEntryShellExec=Ver %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=El prograa de instalaci�n necesita el siguiente disco
SelectDiskLabel2=Por favor, introduzca el Disco %1 y haga clic en Aceptar.%n%nSi los archivos del disco se hallan en una carpeta diferente a la mostrada abajo, introduzca la ruta correcta o haga clic en Examinar.
PathLabel=&Ruta:
FileNotInDir2=No se ha podido encontrar el archivo "%1" en "%2". Por favor, introduzca el disco correcto o seleccione otra carpeta.
SelectDirectoryLabel=Por favor, indique la ubicaci�n del siguiente disco.

; *** Installation phase messages
SetupAborted=La instalaci�n no pudo completarse.%n%nPor favor, corrija el problema y ejecute el programa de instalaci�n de nuevo.
EntryAbortRetryIgnore=Haga clic en Reintentar para intentarlo de nuevo, Ignorar para continuar de todos modos, o Anular para cancelar la instalaci�n.

; *** Installation status messages
StatusCreateDirs=Creando carpetas...
StatusExtractFiles=Extrayendo archivos...
StatusCreateIcons=Creando accesos directos...
StatusCreateIniEntries=Creando entradas de archivo INI...
StatusCreateRegistryEntries=Creando entradas de registro...
StatusRegisterFiles=Registrando archivos...
StatusSavingUninstall=Guardando informaci�n para desinstalar...
StatusRunProgram=Terminando la instalaci�n...
StatusRollback=Deshaciendo cambios...

; *** Misc. errors
ErrorInternal2=Error Interno: %1
ErrorFunctionFailedNoCode=%1 ha fallado
ErrorFunctionFailed=%1 ha fallado; c�digo %2
ErrorFunctionFailedWithMessage=%1 ha fallado; c�digo %2.%n%3
ErrorExecutingProgram=Imposible ejecutar el archivo:%n%1

; *** Registry errors
ErrorRegOpenKey=Error abriendo la clave de registro:%n%1\%2
ErrorRegCreateKey=Error creando la clave de registro:%n%1\%2
ErrorRegWriteKey=Error escribiendo en la clave de registro:%n%1\%2

; *** INI errors
ErrorIniEntry=Error creando entrada en archivo INI "%1".

; *** File copying errors
FileAbortRetryIgnore=Haga clic en Reintentar para intentarlo de nuevo, Ignorar para omitir este archivo (no recomendado), o Anular para cancelar la instalaci�n.
FileAbortRetryIgnore2=Hag clic en Reintentar para intentarlo de nuevo, Ignorar para proceder de todos modos (no recomendado), o Anular para cancelar la instalaci�n.
SourceIsCorrupted=El archivo de origen est� da�ado
SourceDoesntExist=El archivo de origen "%1" no existe
ExistingFileReadOnly=El archivo existente es de s�lo-lectura.%n%nHaga clic en Reintentar para quitar el atributo s�lo-lectura e intentarlo de nuevo, Ignorar para omitir este archivo, o Anular para cancelar la instalaci�n.
ErrorReadingExistingDest=Se produjo un error tratando de leer el archivo existente:
FileExists=El archivo ya existe.%n%n�Desea sobreescribirlo?
ExistingFileNewer=El archivo existente es m�s reciente que el que est� tratando de instalar. Se recomienda que mantenga el archivo existente.%n%n�Desea mantener el archivo existente?
ErrorChangingAttr=Se produjo un error al tratar de cambiar los atributos del archivo:
ErrorCreatingTemp=Se produjo un error al tratar de crear un archivo en la carpeta de destino:
ErrorReadingSource=Se produjo un error al tratar de leer el archivo de origen:
ErrorCopying=Se produjo un error al tratar de copiar un archivo:
ErrorReplacingExistingFile=Se produjo un error al tratar de reemplazar el archivo:
ErrorRestartReplace=Se produjo un fallo al reemplazar:
ErrorRenamingTemp=Se produjo un error al tratar de renombrar un archivo en la carpeta de destino:
ErrorRegisterServer=No se ha podido registrar el DLL/OCX: %1
ErrorRegisterServerMissingExport=No se ha encontrado el exportador DllRegisterServer
ErrorRegisterTypeLib=No se ha podido registrar la librer�a de tipo: %1

; *** Post-installation errors
ErrorOpeningReadme=Se produjo un error al tratar de abrir el archivo L�AME.
ErrorRestartingComputer=El programa de Instalaci�n no ha podido reiniciar el equipo. Por favor, h�galo manualmente.

; *** Uninstaller messages
UninstallNotFound=El archivo "%1" no existe. No se puede desinstalar.
UninstallOpenError=El archivo "%1" no pudo abrirse. No se puede desinstalar
UninstallUnsupportedVer=El archivo de desinstalaci�n "%1" est� en un formato no reconocido por esta versi�n del desinstalador. No se puede desinstalar
UninstallUnknownEntry=Se ha encontrado una entrada desconocida (%1) en el registro de desinstalaci�n
ConfirmUninstall=�Est� seguro que desea eliminar completamente %1 y todos sus componentes?
UninstallOnlyOnWin64=Este programa s�lo puede ser desinstalado en un Windows de 64 bits.
OnlyAdminCanUninstall=Este programa s�lo puede ser desinstalado un usuario con privilegios de administrador.
UninstallStatusLabel=Por favor, espere mientras se elimina %1 de su equipo.
UninstalledAll=%1 se ha eliminado correctamente de su equipo.
UninstalledMost=Desinstalaci�n de %1 completada.%n%nAlgunos elementos no pudieron eliminarse. Puede eliminarlos manualmente.
UninstalledAndNeedsRestart=Para completar la desinstalaci�n de %1, el equipo debe reiniciarse.%n%n�Desea reiniciarlo ahora?
UninstallDataCorrupted=El archivo "%1" est� da�ado. No puede desinstalarse

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=�Eliminar archivos compartidos?
ConfirmDeleteSharedFile2=El sistema indica que el siguiente archivo compartido no es usado por ning�n otro programa. �Desea eliminar este archivo compartido?%n%nSi otros programas usan este archivo y es eliminado, pueden dejar de funcionar correctamente. Si no est� seguro, elija No. Dejar el archivo en su sistema no causar� ning�n da�o.
SharedFileNameLabel=Nombre del archivo:
SharedFileLocationLabel=Ubicaci�n:
WizardUninstalling=Estado de la desinstalaci�n
StatusUninstalling=Desinstalando %1...

; The custom messages below aren't used by Setup itself, but if you make
; use of them in your scripts, you'll want to translate them.

[CustomMessages]

NameAndVersion=%1 versi�n %2
AdditionalIcons=Iconos adicionales:
CreateDesktopIcon=Crear un Acceso directo en el &Escritorio
CreateQuickLaunchIcon=Crear un icono en la barra de inicio &r�pido
ProgramOnTheWeb=%1 en la Web
UninstallProgram=Desinstalar %1
LaunchProgram=Ejecutar %1
AssocFileExtension=&Asociar %1 con la extensi�n de archivo %2
AssocingFileExtension=Asociando %1 con la extensi�n de archivo %2...
