  msg[0]:=
  'You are not a member of the administrator group.'#10#13#13#13
  'For creating a pdf-file out of any application, the PDFCreator is using a postscript printerdriver.'#13
  'You have to be a member of the administrator group, to install any printerdriver.'#13#13
  'If you continue the installation, you will only be able to convert existing postscript-files.'#13#13
  'Continue the setup?';
  msg[1]:=
  'There was found an old version '#39+GetInstalledVersionBeta+#39'.'#10#13 + 'It is possible to update this version to version '#39+ExpandConstant('{#AppVersionStr}')+#39'.'#10#13#10#13;
  msg[2]:=
  'For this you have NOT to be a member of the administrator group.'#10#13#10#13;
  msg[3]:=
  'To update, use <OK>, or cancel the setup and uninstall the older version first.';
  msg[4]:=
  'The program is installed already.'#10#13#10#13
  'For a new installation uninstall the program first.'#10#13#10#13
  'The setup will be cancelled.';
  msg[5]:=
  'The installed version '#39 + GetInstalledVersion + #39' is newer than these setupversion '#39+ExpandConstant('{#AppVersionStr}')+#39'!'#10#13#10#13
  'For the installation of an older version, uninstall the program first.'#10#13#10#13
  'The setup will be cancelled.';
  msg[6]:=
  'The program is installed already.'#10#13#10#13
  'An update is not possible! Please uninstall the program first.'#10#13#10#13
  'The setup will be cancelled.';
  msg[7]:=
  'The program <PDFCreator.exe> is running.'#10#13#10#13
  'Please close the program first.'#10#13#10#13;
  msg[8]:=
  'The program <Transtool.exe> is running.'#10#13#10#13
  'Please close the program first.'#10#13#10#13;
  msg[9]:=
  'The program <PDFSpooler.exe> is running.'#10#13#10#13
  'Please wait until all printjobs are finished or delete these printjobs.'#10#13#10#13;
  msg[10]:=
  'The program <UnInst.exe> is running.'#10#13#10#13
  'Please finish the uninstallation first.'#10#13#10#13
  msg[11]:=
  'There was found an old version '#39+GetInstalledVersionBeta+#39'.'#10#13#10#13
  'It is not possible to update this version to version '#39+ExpandConstant('{#AppVersionStr}')+#39'.'#10#13#10#13
  'Please uninstall the program first.';
