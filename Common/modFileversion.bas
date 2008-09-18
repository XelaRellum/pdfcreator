Attribute VB_Name = "modFileversion"
Option Explicit

Public Function GetFileVersion(filename As String) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim l As Long, buff() As Byte, Pointer As Long, BuffL As Long, _
  Version As VS_FIXEDFILEINFO, flagsStr As String, osStr As String, _
  Typ As String, STyp As String
50040  l = GetFileVersionInfoSize(filename, 0&)
50050  Set GetFileVersion = New Collection
50060  If l > 0 Then
50070   ReDim buff(l)
50080   Call GetFileVersionInfo(filename, 0&, l, buff(0))
50090   Call VerQueryValue(buff(0), "\", Pointer, BuffL)
50100   Call MoveMemory2(Version, Pointer, Len(Version))
50110   With Version
50120    ' Structur version
50130    GetFileVersion.Add Format$(.dwStrucVersionh) & "." & _
    Format$(.dwStrucVersionl)
50150    ' File version
50160    GetFileVersion.Add Format$(.dwFileVersionMSh) & "." & _
    Format$(.dwFileVersionMSl) & "." & _
    Format$(.dwFileVersionLSh) & "." & _
    Format$(.dwFileVersionLSl)
50200    ' Program version
50210    GetFileVersion.Add Format$(.dwProductVersionMSh) & "." & _
    Format$(.dwProductVersionMSl) & "." & _
    Format$(.dwProductVersionLSh) & "." & _
    Format$(.dwProductVersionLSl)
50250    If .dwFileFlags And VS_FF_DEBUG Then flagsStr = "Debug "
50260    If .dwFileFlags And VS_FF_PRERELEASE Then flagsStr = flagsStr & "PreRel "
50270    If .dwFileFlags And VS_FF_PATCHED Then flagsStr = flagsStr & "Patched "
50280    If .dwFileFlags And VS_FF_PRIVATEBUILD Then flagsStr = flagsStr & "Private "
50290    If .dwFileFlags And VS_FF_INFOINFERRED Then flagsStr = flagsStr & "Info "
50300    If .dwFileFlags And VS_FF_SPECIALBUILD Then flagsStr = flagsStr & "Special "
50310    If .dwFileFlags And VFT2_UNKNOWN Then flagsStr = flagsStr & "Unknown "
50320    ' Flags
50330    GetFileVersion.Add flagsStr
50341    Select Case Version.dwFileOS
          Case VOS_DOS_WINDOWS16: osStr = "DOS-Win16"
50360     Case VOS_DOS_WINDOWS32: osStr = "DOS-Win32"
50370     Case VOS_OS216_PM16:    osStr = "OS/2-16 PM-16"
50380     Case VOS_OS232_PM32:    osStr = "OS/2-16 PM-32"
50390     Case VOS_NT_WINDOWS32:  osStr = "NT-Win32"
50400     Case Else:              osStr = "Unknown"
50410    End Select
50420    ' OS
50430    GetFileVersion.Add osStr
50441    Select Case Version.dwFileType
          Case VFT_APP:                Typ = "App"
50460     Case VFT_DLL:                Typ = "DLL"
50470     Case VFT_DRV:                Typ = "Driver"
50481     Select Case Version.dwFileSubtype
           Case VFT2_DRV_PRINTER:     STyp = "Printer drv"
50500      Case VFT2_DRV_KEYBOARD:    STyp = "Keyboard drv"
50510      Case VFT2_DRV_LANGUAGE:    STyp = "Language drv"
50520      Case VFT2_DRV_DISPLAY:     STyp = "Display drv"
50530      Case VFT2_DRV_MOUSE:       STyp = "Mouse drv"
50540      Case VFT2_DRV_NETWORK:     STyp = "Network drv"
50550      Case VFT2_DRV_SYSTEM:      STyp = "System drv"
50560      Case VFT2_DRV_INSTALLABLE: STyp = "Installable"
50570      Case VFT2_DRV_SOUND:       STyp = "Sound drv"
50580      Case VFT2_DRV_COMM:        STyp = "Comm drv"
50590      Case VFT2_UNKNOWN:         STyp = "Unknown"
50600      End Select
50610     Case VFT_FONT:               Typ = "Font"
50621      Select Case Version.dwFileSubtype
            Case VFT2_FONT_RASTER:     STyp = "Raster Font"
50640       Case VFT2_FONT_VECTOR:     STyp = "Vector Font"
50650       Case VFT2_FONT_TRUETYPE:   STyp = "TrueType Font"
50660      End Select
50670     Case VFT_VXD:                Typ = "VxD"
50680     Case VFT_STATIC_LIB:         Typ = "Lib"
50690     Case Else:                   Typ = "Unbekannt"
50700    End Select
50710   ' Typ
50720    GetFileVersion.Add Typ
50730   ' Subtyp
50740    GetFileVersion.Add STyp
50750   End With
50760  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modFileversion", "GetFileVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
