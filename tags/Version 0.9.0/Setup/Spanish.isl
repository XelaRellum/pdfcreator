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
; Translated by Ángel Martín
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
SetupAppTitle=Instalación
SetupWindowTitle=Instalación - %1
UninstallAppTitle=Desinstalación
UninstallAppFullTitle=Desinstalar %1

; *** Misc. common
InformationTitle=Información
ConfirmTitle=Confirmación
ErrorTitle=Error

; *** SetupLdr messages
SetupLdrStartupMessage=Este programa instalará %1. ¿Desea continuar?
LdrCannotCreateTemp=No se ha podido crear un archivo temporal. Instalación cancelada
LdrCannotExecTemp=No se ha podido ejecutar el archivo en el directorio temporal. Instalación cancelada

; *** Startup error messages
LastErrorMessage=%1.%n%nError %2: %3
SetupFileMissing=No se encuentra el archivo %1 en la carpeta de instalación. Por favor, corrija el problema u obtenga una nueva copia del programa.
SetupFileCorrupt=Los archivos de instalación están dañados. Por favor, obtenga una nueva copia del programa.
SetupFileCorruptOrWrongVer=Los archivos de instalación están dañados, o son incompatibles con ésta versión de la instalación. Por favor, corrija el problema u obtenga una nueva copia del programa.
NotOnThisPlatform=Este programa no funcionará en %1.
OnlyOnThisPlatform=Este programa debe ejecutarse en %1.
OnlyOnTheseArchitectures=Este programa sólo puede ser instalado en versiones de Windows diseñadas para las siguientes arquitecturas de procesador:%n%n%1
MissingWOW64APIs=La versión de Windows que está usando no incluye la funcionalidad necesaria para realizar una instalación de 64 bits. Para corregir este problema, por favor, instale el Service Pack %1.
WinVersionTooLowError=Este programa requiere %1 versión %2 o posterior.
WinVersionTooHighError=Este programa no puede ser instalado en %1 versión %2 o posterior.
AdminPrivilegesRequired=Debe iniciar la sesión como administrador para instalar este programa.
PowerUserPrivilegesRequired=Debe iniciar la sesión como administrador o miembro del grupo Usuarios Avanzados para instalar este programa.
SetupAppRunningError=El programa de instalación ha detectado que %1 se está ejecutando actualmente.%n%nPor favor, ciérrelo y luego haga clic en Aceptar para continuar, o Cancelar para salir.
UninstallAppRunningError=El programa de desinstalación ha detectado que %1 se está ejecutando actualmente.%n%nPor favor, ciérrelo y luego haga clic en Aceptar para continuar, o Cancelar para salir.

; *** Misc. errors
ErrorCreatingDir=El programa de instalación no ha podido crear la carpeta "%1"
ErrorTooManyFilesInDir=No se ha podido crear un archivo en la carpeta "%1" porque contiene demasiados archivos.

; *** Setup common messages
ExitSetupTitle=Salir de la Instalación
ExitSetupMessage=La instalación no se ha completado. Si abandona ahora, el programa no quedará instalado.%n%nPara completarla, podrá ejecutar de nuevo el programa de instalación en otro momento.%n%n¿Salir de la Instalación?
AboutSetupMenuItem=&Acerca de la Instalación...
AboutSetupTitle=Acerca de la Instalación
AboutSetupMessage=%1 versión %2%n%3%n%nPágina web de %1:%n%4
AboutSetupNote=
TranslatorNote=Spanish (Standard) translation by Ángel Martín (amartin@gawab.com)

; *** Buttons
ButtonBack=< &Atrás
ButtonNext=&Siguiente >
ButtonInstall=&Instalar
ButtonOK=Aceptar
ButtonCancel=Cancelar
ButtonYes=&Sí
ButtonYesToAll=Sí a &todo
ButtonNo=&No
ButtonNoToAll=N&o a todo
ButtonFinish=&Terminar
ButtonBrowse=&Examinar...
ButtonWizardBrowse=&Examinar...
ButtonNewFolder=C&rear nueva carpeta...

; *** "Select Language" dialog messages
SelectLanguageTitle=Elija el idioma de instalación
SelectLanguageLabel=Elija el idioma a usar durante la instalación:

; *** Common wizard text
ClickNext=Haga clic en Siguiente para continuar, o en Cancelar para abandonar la instalación.
BeveledLabel=
BrowseDialogTitle=Buscar carpeta
BrowseDialogLabel=Elija una carpeta de la lista, y haga clic en Aceptar.
NewFolderName=Nueva carpeta

; *** "Welcome" wizard page
WelcomeLabel1=Bienvenido al asistente de instalación de [name].
WelcomeLabel2=Este programa instalará [name/ver] en su equipo.%n%nEs recomendable que cierre el resto de aplicaciones antes de continuar.

; *** "Password" wizard page
WizardPassword=Contraseña
PasswordLabel1=Esta instalación está protegida con una contraseña.
PasswordLabel3=Por favor, introduzca su contraseña y haga clic en Siguiente para continuar. La contraseña distingue entre mayúsculas y minúsculas.
PasswordEditLabel=&Contraseña:
IncorrectPassword=La contraseña introducida no es correcta. Por favor, inténtelo de nuevo.

; *** "License Agreement" wizard page
WizardLicense=Acuerdo de Licencia
LicenseLabel=Por favor, lea cuidadosamente la siguiente información antes de continuar.
LicenseLabel3=Por favor, lea cuidadosamente el siguiente acuerdo de licencia. Debe de aceptar los términos de este acuerdo para continuar con la instalación.
LicenseAccepted=A&cepto el acuerdo
LicenseNotAccepted=&No acepto el acuerdo

; *** "Information" wizard pages
WizardInfoBefore=Información
InfoBeforeLabel=Por favor, lea la siguiente información antes de continuar.
InfoBeforeClickLabel=Cuando esté listo para continuar con la instalación, haga clic en Siguiente.
WizardInfoAfter=Información
InfoAfterLabel=Por favor, lea la siguiente información antes de continuar.
InfoAfterClickLabel=Cuando esté listo para continuar con la instalación, haga clic en Siguiente.

; *** "User Information" wizard page
WizardUserInfo=Información sobre el usuario
UserInfoDesc=Por favor, introduzca su información.
UserInfoName=Nombre del &usuario:
UserInfoOrg=&Organización:
UserInfoSerial=Número de &serie:
UserInfoNameRequired=Debe introducir un nombre.

; *** "Select Destination Location" wizard page
WizardSelectDir=Elija la Carpeta de Destino
SelectDirDesc=¿Dónde debe instalarse [name]?
SelectDirLabel3=El programa instalará [name] en la siguiente carpeta.
SelectDirBrowseLabel=Para continuar, haga clic en Siguiente. Si desea elegir una carpeta distinta, haga clic en Examinar.
DiskSpaceMBLabel=Se necesita un mínimo de [mb] MB de espacio libre en la unidad.
ToUNCPathname=El programa de instalación no puede instalar en un directorio UNC. Si está tratando de instalar en una red, necesitará mapear una unidad de red.
InvalidPath=Debe introducir una ruta completa con la letra de unidad; por ejemplo:%n%nC:\Aplicación%n%no una ruta UNC de la siguiente forma:%n%n\\servidor\compartido
InvalidDrive=La unidad o ruta UNC que seleccionó no existe o no es accesible. Por favor, elija otra.
DiskSpaceWarningTitle=No hay suficiente espacio en el disco
DiskSpaceWarning=El programa de instalación necesita al menos %1 KB de espacio libre para la instalación, pero la unidad seleccionada solamente tiene %2 KB disponibles.%n%n¿Desea continuar?
DirNameTooLong=El nombre de la carpeta o de la ruta es demasiado largo.
InvalidDirName=El nombre de la carpeta no es válido.
BadDirName32=Los nombres de carpeta no pueden incluir ninguno de los siguientes caracteres:%n%n%1
DirExistsTitle=La carpeta existe
DirExists=La carpeta:%n%n%1%n%nya existe. ¿Desea instalar en dicha carpeta de todas formas?
DirDoesntExistTitle=La carpeta no existe
DirDoesntExist=La carpeta:%n%n%1%n%nno existe. ¿Desea que dicha carpeta sea creada?

; *** "Select Components" wizard page
WizardSelectComponents=Selección de Componentes
SelectComponentsDesc=¿Qué componentes deben instalarse?
SelectComponentsLabel2=Seleccione los componentes a instalar; desmarque los componentes que no desea instalar. Haga clic en Siguiente cuando esté preparado para continuar.
FullInstallation=Instalación Completa
; if possible don't translate 'Compact' as 'Minimal' (I mean 'Minimal' in your language)
CompactInstallation=Instalación Compacta
CustomInstallation=Instalación Personalizada
NoUninstallWarningTitle=Componentes Existentes
NoUninstallWarning=El programa de instalación ha detectado que los siguientes componentes están instalados en su equipo:%n%n%1%n%nSi estos componentes no están seleccionados no serán desinstalados.%n%n¿Desea continuar de todos modos?
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=La selección actual requiere un mínimo de [mb] MB de espacio en disco.

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Seleccione Tareas Adicionales
SelectTasksDesc=¿Qué tareas adicionales deben realizarse?
SelectTasksLabel2=Elija las tareas adicionales que desea que se realicen durante la instalación de [name], y después haga clic en Siguiente.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=Seleccione la carpeta del Menú Inicio
SelectStartMenuFolderDesc=¿Dónde deben ubicarse los accesos directos del programa?
SelectStartMenuFolderLabel3=El programa de instalación creará los accesos directos de programa en la siguiente carpeta del Menú Inicio.
SelectStartMenuFolderBrowseLabel=Para continuar, haga clic en Siguiente. Si desea elegir una carpeta distinta, haga clic en Examinar.
MustEnterGroupName=Debe introducir un nombre de carpeta.
GroupNameTooLong=El nombre de la carpeta o su ruta es demasiado largo.
InvalidGroupName=El nombre de la carpeta no es válido.
BadGroupName=El nombre de la carpeta no puede contener ninguno de los siguientes caracteres:%n%n%1
NoProgramGroupCheck2=&No crear una carpeta en el Menú Inicio

; *** "Ready to Install" wizard page
WizardReady=Preparado para Instalar
ReadyLabel1=El programa de instalación está preparado para comenzar la instalación de [name] en su equipo.
ReadyLabel2a=Haga clic en Instalar para continuar con la instalación, o Atrás si desea revisar o cambiar las opciones de instalación.
ReadyLabel2b=Haga clic en Instalar para continuar con la instalación.
ReadyMemoUserInfo=Información del usuario:
ReadyMemoDir=Carpeta de destino:
ReadyMemoType=Tipo de instalación:
ReadyMemoComponents=Componentes seleccionados:
ReadyMemoGroup=Carpeta del Menú Inicio:
ReadyMemoTasks=Tareas adicionales:

; *** "Preparing to Install" wizard page
WizardPreparing=Preparándose para Instalar
PreparingDesc=El programa se está preparando para instalar [name] en su equipo.
PreviousInstallNotCompleted=La instalación/desinstalación previa de un programa no se completó. Necesitará reiniciar el equipo para completar esa instalación.%n%nDespués de reiniciar el equipo, ejecute éste programa de nuevo para completar la instalación de [name].
CannotContinue=El programa de instalación no puede continuar. Por favor, haga clic en Cancelar para salir.

; *** "Installing" wizard page
WizardInstalling=Instalando
InstallingLabel=Por favor, espere mientras se instala [name] en su equipo.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Completando el asistente de instalación de [name]
FinishedLabelNoIcons=El programa ha finalizado la instalación de [name] en su equipo.
FinishedLabel=El programa ha finalizado la instalación de [name] en su equipo. Puede iniciar la aplicación seleccionando los iconos instalados.
ClickFinish=Haga clic en Terminar para salir de la instalación.
FinishedRestartLabel=Para completar la instalación de [name], el programa de instalación debe reiniciar su equipo. ¿Desea reiniciar ahora?
FinishedRestartMessage=Para completar la instalación de [name], el programa de instalación debe reiniciar su equipo.%n%n¿Desea reiniciar ahora?
ShowReadmeCheck=Sí, deseo ver el archivo LÉAME.
YesRadio=&Sí, deseo reiniciar el equipo ahora
NoRadio=&No, reiniciaré el equipo más tarde
; used for example as 'Run MyProg.exe'
RunEntryExec=Ejecutar %1
; used for example as 'View Readme.txt'
RunEntryShellExec=Ver %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=El prograa de instalación necesita el siguiente disco
SelectDiskLabel2=Por favor, introduzca el Disco %1 y haga clic en Aceptar.%n%nSi los archivos del disco se hallan en una carpeta diferente a la mostrada abajo, introduzca la ruta correcta o haga clic en Examinar.
PathLabel=&Ruta:
FileNotInDir2=No se ha podido encontrar el archivo "%1" en "%2". Por favor, introduzca el disco correcto o seleccione otra carpeta.
SelectDirectoryLabel=Por favor, indique la ubicación del siguiente disco.

; *** Installation phase messages
SetupAborted=La instalación no pudo completarse.%n%nPor favor, corrija el problema y ejecute el programa de instalación de nuevo.
EntryAbortRetryIgnore=Haga clic en Reintentar para intentarlo de nuevo, Ignorar para continuar de todos modos, o Anular para cancelar la instalación.

; *** Installation status messages
StatusCreateDirs=Creando carpetas...
StatusExtractFiles=Extrayendo archivos...
StatusCreateIcons=Creando accesos directos...
StatusCreateIniEntries=Creando entradas de archivo INI...
StatusCreateRegistryEntries=Creando entradas de registro...
StatusRegisterFiles=Registrando archivos...
StatusSavingUninstall=Guardando información para desinstalar...
StatusRunProgram=Terminando la instalación...
StatusRollback=Deshaciendo cambios...

; *** Misc. errors
ErrorInternal2=Error Interno: %1
ErrorFunctionFailedNoCode=%1 ha fallado
ErrorFunctionFailed=%1 ha fallado; código %2
ErrorFunctionFailedWithMessage=%1 ha fallado; código %2.%n%3
ErrorExecutingProgram=Imposible ejecutar el archivo:%n%1

; *** Registry errors
ErrorRegOpenKey=Error abriendo la clave de registro:%n%1\%2
ErrorRegCreateKey=Error creando la clave de registro:%n%1\%2
ErrorRegWriteKey=Error escribiendo en la clave de registro:%n%1\%2

; *** INI errors
ErrorIniEntry=Error creando entrada en archivo INI "%1".

; *** File copying errors
FileAbortRetryIgnore=Haga clic en Reintentar para intentarlo de nuevo, Ignorar para omitir este archivo (no recomendado), o Anular para cancelar la instalación.
FileAbortRetryIgnore2=Hag clic en Reintentar para intentarlo de nuevo, Ignorar para proceder de todos modos (no recomendado), o Anular para cancelar la instalación.
SourceIsCorrupted=El archivo de origen está dañado
SourceDoesntExist=El archivo de origen "%1" no existe
ExistingFileReadOnly=El archivo existente es de sólo-lectura.%n%nHaga clic en Reintentar para quitar el atributo sólo-lectura e intentarlo de nuevo, Ignorar para omitir este archivo, o Anular para cancelar la instalación.
ErrorReadingExistingDest=Se produjo un error tratando de leer el archivo existente:
FileExists=El archivo ya existe.%n%n¿Desea sobreescribirlo?
ExistingFileNewer=El archivo existente es más reciente que el que está tratando de instalar. Se recomienda que mantenga el archivo existente.%n%n¿Desea mantener el archivo existente?
ErrorChangingAttr=Se produjo un error al tratar de cambiar los atributos del archivo:
ErrorCreatingTemp=Se produjo un error al tratar de crear un archivo en la carpeta de destino:
ErrorReadingSource=Se produjo un error al tratar de leer el archivo de origen:
ErrorCopying=Se produjo un error al tratar de copiar un archivo:
ErrorReplacingExistingFile=Se produjo un error al tratar de reemplazar el archivo:
ErrorRestartReplace=Se produjo un fallo al reemplazar:
ErrorRenamingTemp=Se produjo un error al tratar de renombrar un archivo en la carpeta de destino:
ErrorRegisterServer=No se ha podido registrar el DLL/OCX: %1
ErrorRegisterServerMissingExport=No se ha encontrado el exportador DllRegisterServer
ErrorRegisterTypeLib=No se ha podido registrar la librería de tipo: %1

; *** Post-installation errors
ErrorOpeningReadme=Se produjo un error al tratar de abrir el archivo LÉAME.
ErrorRestartingComputer=El programa de Instalación no ha podido reiniciar el equipo. Por favor, hágalo manualmente.

; *** Uninstaller messages
UninstallNotFound=El archivo "%1" no existe. No se puede desinstalar.
UninstallOpenError=El archivo "%1" no pudo abrirse. No se puede desinstalar
UninstallUnsupportedVer=El archivo de desinstalación "%1" está en un formato no reconocido por esta versión del desinstalador. No se puede desinstalar
UninstallUnknownEntry=Se ha encontrado una entrada desconocida (%1) en el registro de desinstalación
ConfirmUninstall=¿Está seguro que desea eliminar completamente %1 y todos sus componentes?
UninstallOnlyOnWin64=Este programa sólo puede ser desinstalado en un Windows de 64 bits.
OnlyAdminCanUninstall=Este programa sólo puede ser desinstalado un usuario con privilegios de administrador.
UninstallStatusLabel=Por favor, espere mientras se elimina %1 de su equipo.
UninstalledAll=%1 se ha eliminado correctamente de su equipo.
UninstalledMost=Desinstalación de %1 completada.%n%nAlgunos elementos no pudieron eliminarse. Puede eliminarlos manualmente.
UninstalledAndNeedsRestart=Para completar la desinstalación de %1, el equipo debe reiniciarse.%n%n¿Desea reiniciarlo ahora?
UninstallDataCorrupted=El archivo "%1" está dañado. No puede desinstalarse

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=¿Eliminar archivos compartidos?
ConfirmDeleteSharedFile2=El sistema indica que el siguiente archivo compartido no es usado por ningún otro programa. ¿Desea eliminar este archivo compartido?%n%nSi otros programas usan este archivo y es eliminado, pueden dejar de funcionar correctamente. Si no está seguro, elija No. Dejar el archivo en su sistema no causará ningún daño.
SharedFileNameLabel=Nombre del archivo:
SharedFileLocationLabel=Ubicación:
WizardUninstalling=Estado de la desinstalación
StatusUninstalling=Desinstalando %1...

; The custom messages below aren't used by Setup itself, but if you make
; use of them in your scripts, you'll want to translate them.

[CustomMessages]

NameAndVersion=%1 versión %2
AdditionalIcons=Iconos adicionales:
CreateDesktopIcon=Crear un Acceso directo en el &Escritorio
CreateQuickLaunchIcon=Crear un icono en la barra de inicio &rápido
ProgramOnTheWeb=%1 en la Web
UninstallProgram=Desinstalar %1
LaunchProgram=Ejecutar %1
AssocFileExtension=&Asociar %1 con la extensión de archivo %2
AssocingFileExtension=Asociando %1 con la extensión de archivo %2...
