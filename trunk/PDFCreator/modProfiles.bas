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
Select Case ErrPtnr.OnError("modProfiles", "ProfileAssociatedPrinter")
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
50010  Dim reg As clsRegistry
50020
50030  Set reg = New clsRegistry
50040  reg.KeyRoot = "Software\PDFCreator\Profiles\"
50050
50060  If InstalledAsServer Then
50070    reg.hkey = HKEY_LOCAL_MACHINE
50080   Else
50090    reg.hkey = HKEY_CURRENT_USER
50100  End If
50110  Set GetProfiles = reg.EnumRegistryKeys(reg.hkey, reg.KeyRoot)
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

Public Function GetPrinterDefaultProfile(Printername As String) As String
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
50110   If UCase$(Trim$(PrinterProfiles(i)(0))) = UCase$(Trim$(Printername)) Then
50120    GetPrinterDefaultProfile = PrinterProfiles(i)(1)
50130    Exit For
50140   End If
50150  Next i
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
50050
50060  reg.KeyRoot = "Software\PDFCreator\Profiles"
50070
50080  If InstalledAsServer Then
50090    reg.hkey = HKEY_LOCAL_MACHINE
50100   Else
50110    reg.hkey = HKEY_CURRENT_USER
50120  End If
50130  reg.DeleteKeyWithSubkeys ProfileName
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

