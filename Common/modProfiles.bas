Attribute VB_Name = "modProfiles"
Option Explicit

Public Function ProfileAssociatedPrinters(ProfileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PrinterProfiles As Collection, p As Variant, i As Long, tStr As String
50020  Set PrinterProfiles = GetPrinterProfiles
50030
50040  For i = 1 To PrinterProfiles.Count
50050   If StrComp(PrinterProfiles(i)(1), ProfileName, vbTextCompare) = 0 Then
50060    If LenB(tStr) = 0 Then
50070      tStr = PrinterProfiles(i)(0)
50080     Else
50090      tStr = tStr & ", " & PrinterProfiles(i)(0)
50100    End If
50110   End If
50120  Next i
50130  ProfileAssociatedPrinters = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProfiles", "ProfileAssociatedPrinters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function HKLMProfileExists(ByVal ProfileName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, keys As Collection, i As Long
50020
50030  Set reg = New clsRegistry
50040
50050  reg.hkey = HKEY_LOCAL_MACHINE
50060  reg.KeyRoot = "Software\PDFCreator\Profiles\"
50070  Set keys = reg.EnumRegistryKeys(reg.hkey, reg.KeyRoot)
50080
50090  For i = 1 To keys.Count
50100   If LCase$(keys(i)) = LCase$(ProfileName) Then
50110    HKLMProfileExists = True
50120    Exit Function
50130   End If
50140  Next i
50150  HKLMProfileExists = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProfiles", "HKLMProfileExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ProfileExists(ByVal ProfileName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Profiles As Collection, i As Long
50020  Set Profiles = GetProfiles
50030  For i = 1 To Profiles.Count
50040   If LCase$(Profiles(i)) = LCase$(ProfileName) Then
50050    ProfileExists = True
50060    Exit Function
50070   End If
50080  Next i
50090  ProfileExists = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProfiles", "ProfileExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetProfiles() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, tmpProfiles1 As Collection, tmpProfiles2 As Collection, i As Long
50020
50030  Set reg = New clsRegistry
50040
50050  If InstalledAsServer Then
50060    reg.hkey = HKEY_LOCAL_MACHINE
50070    reg.KeyRoot = "Software\PDFCreator\Profiles\"
50080    Set GetProfiles = reg.EnumRegistryKeys(reg.hkey, reg.KeyRoot)
50090   Else
50100    reg.hkey = HKEY_USERS
50110    reg.KeyRoot = ".\Default\Software\PDFCreator\Profiles\"
50120    Set tmpProfiles1 = reg.EnumRegistryKeys(reg.hkey, reg.KeyRoot)
50130    reg.hkey = HKEY_CURRENT_USER
50140    reg.KeyRoot = "Software\PDFCreator\Profiles\"
50150    Set tmpProfiles2 = reg.EnumRegistryKeys(reg.hkey, reg.KeyRoot)
50160    For i = 1 To tmpProfiles2.Count
50170     AddSortedStr tmpProfiles1, tmpProfiles2(i)
50180    Next i
50190    reg.hkey = HKEY_LOCAL_MACHINE
50200    reg.KeyRoot = "Software\PDFCreator\Profiles\"
50210    Set tmpProfiles2 = reg.EnumRegistryKeys(reg.hkey, reg.KeyRoot)
50220    For i = 1 To tmpProfiles2.Count
50230     AddSortedStr tmpProfiles1, tmpProfiles2(i)
50240    Next i
50250    Set GetProfiles = tmpProfiles1
50260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProfiles", "GetProfiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPrinterDefaultProfile(PrinterName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, PrinterProfiles As Collection, i As Long
50020  Set reg = New clsRegistry
50030  reg.KeyRoot = "Software\PDFCreator\Printers\"
50040  If InstalledAsServer Then
50050    reg.hkey = HKEY_LOCAL_MACHINE
50060   Else
50070    reg.hkey = HKEY_CURRENT_USER
50080  End If
50090  Set PrinterProfiles = reg.EnumRegistryValues(reg.hkey, "Software\PDFCreator\Printers\")
50100  For i = 1 To PrinterProfiles.Count
50110   If UCase$(Trim$(PrinterProfiles(i)(0))) = UCase$(Trim$(PrinterName)) Then
50120    GetPrinterDefaultProfile = PrinterProfiles(i)(1)
50130    Exit For
50140   End If
50150  Next i
50160  If InstalledAsServer Then
50170    reg.hkey = HKEY_LOCAL_MACHINE
50180    Set PrinterProfiles = reg.EnumRegistryValues(reg.hkey, "Software\Policies\PDFCreator\Printers\")
50190    For i = 1 To PrinterProfiles.Count
50200     If UCase$(Trim$(PrinterProfiles(i)(0))) = UCase$(Trim$(PrinterName)) Then
50210      GetPrinterDefaultProfile = PrinterProfiles(i)(1)
50220      Exit For
50230     End If
50240    Next i
50250   Else
50260    reg.hkey = HKEY_CURRENT_USER
50270    Set PrinterProfiles = reg.EnumRegistryValues(reg.hkey, "Software\Policies\PDFCreator\Printers\")
50280    For i = 1 To PrinterProfiles.Count
50290     If UCase$(Trim$(PrinterProfiles(i)(0))) = UCase$(Trim$(PrinterName)) Then
50300      GetPrinterDefaultProfile = PrinterProfiles(i)(1)
50310      Exit For
50320     End If
50330    Next i
50340    reg.hkey = HKEY_LOCAL_MACHINE
50350    Set PrinterProfiles = reg.EnumRegistryValues(reg.hkey, "Software\PDFCreator\Printers\")
50360    For i = 1 To PrinterProfiles.Count
50370     If UCase$(Trim$(PrinterProfiles(i)(0))) = UCase$(Trim$(PrinterName)) Then
50380      GetPrinterDefaultProfile = PrinterProfiles(i)(1)
50390      Exit For
50400     End If
50410    Next i
50420    Set PrinterProfiles = reg.EnumRegistryValues(reg.hkey, "Software\Policies\PDFCreator\Printers\")
50430    For i = 1 To PrinterProfiles.Count
50440     If UCase$(Trim$(PrinterProfiles(i)(0))) = UCase$(Trim$(PrinterName)) Then
50450      GetPrinterDefaultProfile = PrinterProfiles(i)(1)
50460      Exit For
50470     End If
50480    Next i
50490  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProfiles", "GetPrinterDefaultProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub DeleteProfile(ProfileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030
50040  ProfileName = Trim$(ProfileName)
50050  reg.KeyRoot = "Software\PDFCreator\Profiles"
50060
50070  If InstalledAsServer Then
50080    reg.hkey = HKEY_LOCAL_MACHINE
50090   Else
50100    reg.hkey = HKEY_CURRENT_USER
50110  End If
50120  reg.DeleteKeyWithSubkeys ProfileName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProfiles", "DeleteProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UpdatePrinterProfiles(ByRef PrinterProfilesMain As Collection, ByVal PrinterProfilesUpdate As Collection)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long
50020  For i = 1 To PrinterProfilesMain.Count
50030   For j = PrinterProfilesUpdate.Count To 1 Step -1
50040    If StrComp(PrinterProfilesMain(i)(0), PrinterProfilesUpdate(j)(0), vbTextCompare) = 0 Then
50050     PrinterProfilesMain.Add PrinterProfilesUpdate(j), , , i
50060     PrinterProfilesMain.Remove i
50070     PrinterProfilesUpdate.Remove j
50080     Exit For
50090    End If
50100   Next j
50110  Next i
50120  For j = 1 To PrinterProfilesUpdate.Count
50130   PrinterProfilesMain.Add PrinterProfilesUpdate(j)
50140  Next j
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProfiles", "UpdatePrinterProfiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetPrinterProfiles() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, PrinterProfiles As Collection, tColl As Collection
50020  Set reg = New clsRegistry
50030  Set PrinterProfiles = New Collection
50040  If InstalledAsServer Then
50050    Set PrinterProfiles = reg.EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\PDFCreator\Printers")
50060   Else
50070    Set PrinterProfiles = reg.EnumRegistryValues(HKEY_CURRENT_USER, "Software\PDFCreator\Printers")
50080    UpdatePrinterProfiles PrinterProfiles, reg.EnumRegistryValues(HKEY_CURRENT_USER, "Software\Policies\PDFCreator\Printers")
50090    UpdatePrinterProfiles PrinterProfiles, reg.EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\PDFCreator\Printers")
50100  End If
50110  UpdatePrinterProfiles PrinterProfiles, reg.EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\Policies\PDFCreator\Printers")
50120  Set GetPrinterProfiles = PrinterProfiles
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProfiles", "GetPrinterProfiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

