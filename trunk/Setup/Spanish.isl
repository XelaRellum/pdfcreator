; *** Inno Setup version 4.1.8+ Spanish messages ***
;
; Note: When translating this text, do not add periods (.) to the end of
; messages that didn't have them already, because on those messages Inno
; Setup adds the periods automatically (appending a period would result in
; two periods being displayed).
;
; $Id: SpanishStd.isl,v 1.52 2004/03/16 9:11 jl Exp $
;
; Versi�n 2: Traducida al Espa�ol por Germ�n Giraldo G. - Colombia
; e-mail: tripleg@tutopia.com
;
; Versi�n 3+: Adaptado al Espa�ol (Espa�a) por Jordi Latorre - Espa�a
;               e-mail: jlatorref@eic.ictnet.es
;             y Ximo Tamarit - Espa�a
;               e-mail: tamarit@mail.ono.es
;

[LangOptions]
LanguageName=Spanish - Traditional Sort
LanguageID=$040a
; Si el lenguaje al cual est� traduciendo requiere un tipo de letra o
; tama�o, quite el comentario de alguna de las siguientes entradas y c�mbielas seg�n el caso.
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
SetupAppTitle=Instalar
SetupWindowTitle=Instalar - %1
UninstallAppTitle=Desinstalar
UninstallAppFullTitle=Desinstalar - %1

; *** Misc. common
InformationTitle=Informaci�n
ConfirmTitle=Confirmar
ErrorTitle=Error

; *** SetupLdr messages
SetupLdrStartupMessage=Se instalar� %1. �Desea continuar?
LdrCannotCreateTemp=No se ha podido crear el archivo temporal. Se cancela la instalaci�n
LdrCannotExecTemp=No se ha podido ejecutar el archivo en el directorio temporal. Se cancela la instalaci�n

; *** Startup error messages
LastErrorMessage=%1.%n%nError %2: %3
SetupFileMissing=El archivo %1 no se encuentra en la carpeta de instalaci�n. Por favor, corrija el problema u obtenga una copia nueva del programa.
SetupFileCorrupt=Los archivos de instalaci�n est�n da�ados. Por favor, obtenga una copia nueva del programa.
SetupFileCorruptOrWrongVer=Los archivos de instalaci�n est�n da�ados, o son incompatibles con su versi�n del Instalador. Por favor, corrija el problema u obtenga una copia nueva del programa.
NotOnThisPlatform=Este programa no se ejecutar� en %1.
OnlyOnThisPlatform=Este programa debe ejecutarse en %1.
WinVersionTooLowError=Este programa requiere %1 versi�n %2 o posterior.
WinVersionTooHighError=Este programa no puede instalarse en %1 versi�n %2 o posterior.
AdminPrivilegesRequired=Debe iniciar la sesi�n como administrador para instalar este programa.
PowerUserPrivilegesRequired=Debe iniciar la sesi�n como administrador o miembro del grupo Usuarios Avanzados para instalar este programa.
SetupAppRunningError=El Instalador ha detectado que %1 se est� ejecutando actualmente.%n%nPor favor, ci�rrelo ahora, luego haga clic en Aceptar para continuar, o Cancelar para salir.
UninstallAppRunningError=La Desinstalador ha detectado que %1 se est� ejecutando actualmente.%n%nPor favor, ci�rrelo ahora, luego haga clic en Aceptar para continuar, o Cancelar para salir.

; *** Misc. errors
ErrorCreatingDir=Imposible crear la carpeta "%1"
ErrorTooManyFilesInDir=Imposible crear un archivo en la carpeta "%1" porque contiene demasiados archivos.

; *** Setup common messages
ExitSetupTitle=Salir de la Instalaci�n
ExitSetupMessage=La instalaci�n no se ha completado. Si abandona ahora, el programa no quedar� instalado.%n%nPodr� ejecutar de nuevo el programa de instalaci�n para completarla.%n%n�Salir de la Instalaci�n?
AboutSetupMenuItem=&Acerca de Instalar...
AboutSetupTitle=Acerca de Instalar
AboutSetupMessage=%1 versi�n %2%n%3%n%n%1 p�gina Web:%n%4
AboutSetupNote=

; *** Buttons
ButtonBack=< &Atr�s
ButtonNext=&Siguiente >
ButtonInstall=&Instalar
ButtonOK=Aceptar
ButtonCancel=Cancelar
ButtonYes=&S�
ButtonYesToAll=S� a &Todo
ButtonNo=&No
ButtonNoToAll=N&o a Todo
ButtonFinish=&Terminar
ButtonBrowse=&Examinar...
ButtonWizardBrowse=E&xaminar...
ButtonNewFolder=&Crear Nueva Carpeta

; *** "Select Language" dialog messages
SelectLanguageTitle=Seleccione el idioma de la instalaci�n
SelectLanguageLabel=Seleccione el idioma que desea usar durante la instalaci�n:

; *** Common wizard text
ClickNext=Haga clic en Siguiente para continuar, Cancelar para salir.
BeveledLabel=
BrowseDialogTitle=Buscar Carpeta
BrowseDialogLabel=Seleccione una Carpeta de la siguiente lista, haga clic en Aceptar.
NewFolderName=Nueva Carpeta

; *** "Welcome" wizard page
WelcomeLabel1=Bienvenido a la instalaci�n de [name].
WelcomeLabel2=Este programa instalar� [name/ver] en su sistema.%n%nSe recomienda que cierre todas las dem�s aplicaciones antes de continuar.

; *** "Password" wizard page
WizardPassword=Contrase�a
PasswordLabel1=Esta instalaci�n est� protegida por contrase�a.
PasswordLabel3=Por favor, suministre su contrase�a, haga clic en Siguiente para continuar. Las contrase�as diferencian entre may�sculas y min�sculas.
PasswordEditLabel=&Contrase�a:
IncorrectPassword=La contrase�a suministrada no es correcta. Por favor, int�ntelo de nuevo.

; *** "License Agreement" wizard page
WizardLicense=Acuerdo de Licencia
LicenseLabel=Por favor, lea la siguiente informaci�n importante antes de continuar.
LicenseLabel3=Por favor, lea detenidamente el siguiente acuerdo de licencia. Debe de aceptar los t�rminos de este acuerdo antes de continuar con la instalaci�n.
LicenseAccepted=A&cepto el acuerdo
LicenseNotAccepted=&No acepto el acuerdo

; *** "Information" wizard pages
WizardInfoBefore=Informaci�n
InfoBeforeLabel=Por favor, lea la siguiente informaci�n importante antes de continuar.
InfoBeforeClickLabel=Cuando est� listo para continuar con la instalaci�n, haga clic en Siguiente.
WizardInfoAfter=Informaci�n
InfoAfterLabel=Por favor, lea la siguiente informaci�n importante antes de continuar.
InfoAfterClickLabel=Cuando est� listo para continuar, haga clic en Siguiente.

; *** "User Information" wizard page
WizardUserInfo=Informaci�n de usuario
UserInfoDesc=Por favor, introduzca su informaci�n.
UserInfoName=Nombre de &Usuario:
UserInfoOrg=&Organizaci�n:
UserInfoSerial=N�mero de &Serie:
UserInfoNameRequired=Debe proporcionar su nombre.

; *** "Select Destination Location" wizard page
WizardSelectDir=Seleccione la Carpeta Destino
SelectDirDesc=�D�nde debe instalarse [name]?
SelectDirLabel3=El programa instalar� [name] en la siguiente carpeta.
SelectDirBrowseLabel=Para continuar, haga clic en Siguiente. Si desea seleccionar una carpeta distinta, haga clic en Examinar.
DiskSpaceMBLabel=Se requieren al menos [mb] MB de espacio libre en el disco.
ToUNCPathname=No se puede instalar en un directorio UNC. Si est� tratando de instalar en una red, necesitar� mapear una unidad de la red.
InvalidPath=Debe proporcionar una ruta completa con la letra de unidad; por ejemplo:%n%nC:\APP%n%no una ruta UNC de la siguiente forma:%n%n\\servidor\compartido
InvalidDrive=La unidad o ruta UNC que seleccion� no existe o no es accesible. Por favor, seleccione otra.
DiskSpaceWarningTitle=No hay suficiente espacio en el disco
DiskSpaceWarning=Se requiere al menos %1 KB de espacio libre para la instalaci�n, pero la unidad seleccionada solamente tiene %2 KB disponibles.%n%n�Desea continuar?
DirNameTooLong=El nombre de la carpeta o su ruta es demasiado largo.
InvalidDirName=El nombre de la carpeta no es v�lido.
BadDirName32=El nombre de una carpeta no puede incluir ninguno de los siguientes caracteres:%n%n%1
DirExistsTitle=La Carpeta Ya Existe
DirExists=La carpeta:%n%n%1%n%nya existe. �Desea instalar en dicha carpeta de todas formas?
DirDoesntExistTitle=La Carpeta No Existe
DirDoesntExist=La carpeta:%n%n%1%n%nno existe. �Desea que se cree dicha carpeta?

; *** "Select Components" wizard page
WizardSelectComponents=Seleccione los Componentes
SelectComponentsDesc=�Qu� componentes deben de instalarse?
SelectComponentsLabel2=Seleccione los componentes a instalar; desactive los componentes que no desea instalar. Haga clic en Siguiente cuando desee continuar.
FullInstallation=Instalaci�n Completa
; Si es posible no traduzca 'Compacta' a 'Minima' (Me refiero a 'Minima' en su lenguaje)
CompactInstallation=Instalaci�n Compacta
CustomInstallation=Instalaci�n Personalizada
NoUninstallWarningTitle=Componentes Existentes
NoUninstallWarning=El Instalador ha detectado que los siguientes componentes ya est�n instalados en su sistema:%n%n%1%n%nQuitando la selecci�n de estos componentes, no ser�n desinstalados.%n%n�Desea continuar de todos modos?
ComponentSize1=%1 KB
ComponentSize2=%1 MB
ComponentsDiskSpaceMBLabel=La selecci�n actual requiere al menos [mb] MB de espacio en disco.

; *** "Select Additional Tasks" wizard page
WizardSelectTasks=Seleccione las Tareas Adicionales
SelectTasksDesc=�Qu� tareas adicionales deben realizarse?
SelectTasksLabel2=Seleccione las tareas adicionales que usted desea que se realicen durante la instalaci�n de [name] y haga clic en Siguiente.

; *** "Select Start Menu Folder" wizard page
WizardSelectProgramGroup=Seleccione la carpeta del Men� de Inicio
SelectStartMenuFolderDesc=�D�nde deben ubicarse los iconos de programa?
SelectStartMenuFolderLabel3=El Instalador crear� los iconos de programa en la siguiente carpeta del Men� de Inicio.
SelectStartMenuFolderBrowseLabel=Para continuar, haga clic en Siguiente. Si desea seleccionar una carpeta distinta, haga clic en Examinar.
NoIconsCheck=&No crear ning�n icono
MustEnterGroupName=Debe proporcionar un nombre de carpeta.
GroupNameTooLong=El nombre de la carpeta o su ruta es demasiado largo.
InvalidGroupName=El nombre de la carpeta no es v�lido.
BadGroupName=El nombre de la carpeta no puede incluir ninguno de los siguientes caracteres:%n%n%1
NoProgramGroupCheck2=&No crear ninguna carpeta en el Men� Inicio

; *** "Ready to Install" wizard page
WizardReady=Listo para Instalar
ReadyLabel1=Ahora el instalador est� listo para iniciar el proceso de instalaci�n de [name] en su sistema.
ReadyLabel2a=Haga clic en Instalar para continuar con el proceso, o haga clic en Atr�s si desea revisar o cambiar alguna configuraci�n.
ReadyLabel2b=Haga clic en Instalar para continuar con el proceso.
ReadyMemoUserInfo=Informaci�n del usuario:
ReadyMemoDir=Carpeta de Destino:
ReadyMemoType=Tipo de Instalaci�n:
ReadyMemoComponents=Componentes Seleccionados:
ReadyMemoGroup=Carpeta del Men� de Inicio:
ReadyMemoTasks=Tareas Adicionales:

; *** "Preparing to Install" wizard page
WizardPreparing=Prepar�ndose para Instalar
PreparingDesc=El programa se est� preparando para instalar [name] en su sistema.
PreviousInstallNotCompleted=La instalaci�n/desinstalaci�n previa del programa no se complet�. Deber� reiniciar el sistema para completar la instalaci�n.%n%nUna vez reiniciado el sistema, ejecute el programa de nuevo para completar la instalaci�n de [name].
CannotContinue=El programa no puede continuar. Por favor, presione Cancelar para salir.

; *** "Installing" wizard page
WizardInstalling=Instalando
InstallingLabel=Por favor, espere mientras se instala [name] en su sistema.

; *** "Setup Completed" wizard page
FinishedHeadingLabel=Completando la instalaci�n de [name]
FinishedLabelNoIcons=El programa termin� la instalaci�n de [name] en su sistema.
FinishedLabel=El programa termin� la instalaci�n de [name] en su sistema. Puede ejecutar la aplicaci�n haciendo clic sobre el icono instalado.
ClickFinish=Haga clic en Terminar para salir de la Instalaci�n.
FinishedRestartLabel=Para completar la instalaci�n de [name], debe reiniciar su sistema. �Desea reiniciar ahora?
FinishedRestartMessage=Para completar la instalaci�n de [name], debe reiniciar su sistema.%n%n�Desea reiniciar ahora?
ShowReadmeCheck=S�, deseo ver el archivo L�AME.
YesRadio=&S�, deseo reiniciar el sistema ahora
NoRadio=&No, yo reiniciar� el sistema m�s tarde
; used for example as 'Run MyProg.exe'
RunEntryExec=Ejecutar %1
; used for example as 'View Readme.txt'
RunEntryShellExec=Ver %1

; *** "Setup Needs the Next Disk" stuff
ChangeDiskTitle=La Instalaci�n Necesita el Siguiente Disco
SelectDiskLabel2=Por favor, inserte el Disco %1 y haga clic en Aceptar.%n%nSi los archivos pueden hallarse en una carpeta diferente a la mostrada abajo, introduzca la ruta correcta o haga clic en Examinar.
PathLabel=&Ruta:
FileNotInDir2=El archivo "%1" no puede localizarse en "%2". Por favor, inserte el disco correcto o seleccione otra carpeta.
SelectDirectoryLabel=Por favor, especifique la localizaci�n del siguiente disco.

; *** Installation phase messages
SetupAborted=La instalaci�n no puede completarse.%n%nPor favor, corrija el problema y ejecute Instalar de nuevo.
EntryAbortRetryIgnore=Haga clic en Reintentar para intentarlo de nuevo, Ignorar para continuar como sea, o Anular para cancelar la instalaci�n.

; *** Installation status messages
StatusCreateDirs=Creando carpetas...
StatusExtractFiles=Copiando archivos...
StatusCreateIcons=Creando accesos directos...
StatusCreateIniEntries=Creando entradas en INI...
StatusCreateRegistryEntries=Creando entradas de registro...
StatusRegisterFiles=Registrando archivos...
StatusSavingUninstall=Guardando informaci�n para desinstalar...
StatusRunProgram=Terminando la instalaci�n...
StatusRollback=Deshaciendo cambios...

; *** Misc. errors
ErrorInternal2=Error Interno: %1
ErrorFunctionFailedNoCode=%1 fall�
ErrorFunctionFailed=%1 fall�; c�digo %2
ErrorFunctionFailedWithMessage=%1 fall�; c�digo %2.%n%3
ErrorExecutingProgram=Imposible ejecutar el archivo:%n%1

; *** Registry errors
ErrorRegOpenKey=Error abriendo clave de registro:%n%1\%2
ErrorRegCreateKey=Error creando clave de registro:%n%1\%2
ErrorRegWriteKey=Error escribiendo en clave de registro:%n%1\%2

; *** INI errors
ErrorIniEntry=Error creando entrada en archivo INI "%1".

; *** File copying errors
FileAbortRetryIgnore=Haga clic en Reintentar para intentarlo de nuevo, Ignorar para omitir este archivo (no recomendado), o Anular para cancelar la instalaci�n.
FileAbortRetryIgnore2=Hag clic en Reintentar para intentarlo de nuevo, Ignorar para proceder de cualquier forma (no recomendado), o Anular para cancelar la instalaci�n.
SourceIsCorrupted=El archivo de origen est� da�ado
SourceDoesntExist=El archivo de origen "%1" no existe
ExistingFileReadOnly=El archivo existente est� marcado como s�lo-lectura.%n%nHaga clic en Reintentar para quitar el atributo s�lo-lectura e intentarlo de nuevo, Ignorar para omitir este archivo, o Anular para cancelar la instalaci�n.
ErrorReadingExistingDest=Ocurri� un error tratando de leer el archivo existente:
FileExists=El archivo ya existe.%n%n�Desea sobreescribirlo?
ExistingFileNewer=El archivo existente es m�s reciente que el que est� tratando de instalar. Se recomienda que mantenga el archivo existente.%n%n�Desea mantener el archivo existente?
ErrorChangingAttr=Ocurri� un error tratando de cambiar los atributos del archivo:
ErrorCreatingTemp=Ocurri� un error tratando de crear un archivo en la carpeta de destino:
ErrorReadingSource=Ocurri� un error tratando de leer el archivo de origen:
ErrorCopying=Ocurri� un error tratando de copiar el archivo:
ErrorReplacingExistingFile=Ocurri� un error tratando de reemplazar el archivo:
ErrorRestartReplace=Fall� reintento de reemplazar:
ErrorRenamingTemp=Ocurri� un error tratando de renombrar un archivo en la carpeta de destino:
ErrorRegisterServer=Imposible registrar el DLL/OCX: %1
ErrorRegisterServerMissingExport=No se encuentra DllRegisterServer export
ErrorRegisterTypeLib=Imposible registrar la librer�a de tipo: %1

; *** Post-installation errors
ErrorOpeningReadme=Ocurri� un error tratando de abrir el archivo L�AME.
ErrorRestartingComputer=El programa de Instalaci�n no puede reiniciar el sistema. Por favor, h�galo manualmente.

; *** Uninstaller messages
UninstallNotFound=El archivo "%1" no existe. No se puede desinstalar.
UninstallOpenError=El archivo "%1" no pudo abrirse. No se puede desinstalar.
UninstallUnsupportedVer=El archivo de bit�cora para desinstalar "%1" est� en un formato no reconocido por esta versi�n de desinstalaci�n. No se puede desinstalar
UninstallUnknownEntry=Una entrada desconocida (%1) se encontr� en el bit�cora para desinstalar
ConfirmUninstall=�Est� seguro que desea eliminar completamente %n%1 y todos sus componentes?
OnlyAdminCanUninstall=Este programa s�lo puede desinstalarlo un usuario con privilegios de administrador.
UninstallStatusLabel=Por favor, espere mientras se elimina %1 de su sistema.
UninstalledAll=%1 se elimin� con �xito de su sistema.
UninstalledMost=La desinstalaci�n de %1 termin�.%n%nAlgunos elementos no pudieron eliminarse. Puede usted eliminarlos manualmente.
UninstalledAndNeedsRestart=Para completar la desinstalaci�n de %1, el sistema debe de reiniciarse.%n%nQuiere reiniciarlo ahora?
UninstallDataCorrupted=El archivo "%1" est� da�ado. No puede desinstalarse

; *** Uninstallation phase messages
ConfirmDeleteSharedFileTitle=�Eliminar Archivos Compartidos?
ConfirmDeleteSharedFile2=El sistema indica que el siguiente archivo compartido no es usado por ning�n otro programa. �Desea eliminar este archivo compartido?%n%nSi otros programas usan este archivo y es eliminado, pueden dejar de funcionar correctamente. Si no est� seguro, elija <No>. Dejar el archivo en su sistema no producir� ning�n da�o.
SharedFileNameLabel=Nombre de archivo:
SharedFileLocationLabel=Localizaci�n:
WizardUninstalling=Estado de la Desinstalaci�n
StatusUninstalling=Desinstalando %1...

[CustomMessages]

AdditionalIcons=Iconos adicionales:
CreateDesktopIcon=Crear un &icono en el escritorio
CreateQuickLaunchIcon=Crear un icono en la &barra de Inicio R�pido
ProgramOnTheWeb=%1 en la Web
UninstallProgram=Desinstalar %1
LaunchProgram=Ejecutar %1
AssocFileExtension=&Asociar %1 con la extensi�n de archivo %2
AssocingFileExtension=Asociando %1 con la extensi�n de archivo %2...
