Name: desktopicon; Description: Create a &desktop icon; GroupDescription: Additional icons:; Languages: English
Name: desktopicon\common; Description: For &all users; GroupDescription: Additional icons:; Flags: exclusive; Languages: English
Name: desktopicon\user; Description: For the &current user only; GroupDescription: Additional icons:; Flags: exclusive unchecked; Languages: English
Name: quicklaunchicon; Description: Create a &Quick Launch icon; GroupDescription: Additional icons:; Flags: unchecked; Languages: English

#Ifdef GhostscriptVersion
Name: ghostscript; Description: Install {#License} &Ghostscript Version {#GhostscriptVersion}; GroupDescription: Other tasks:; Flags: exclusive; Languages: English
#ENDIF

Name: fileassoc; Description: &Associate PDFCreator with the .ps file extension; GroupDescription: Other tasks:; Flags: unchecked; Languages: English
Name: winexplorer; Description: Create a &entry in then windows explorer context menu; GroupDescription: Other tasks:; Languages: English
