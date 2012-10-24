Attribute VB_Name = "modPrinterEnums"
Option Explicit

Public Function GetAvailableMonitors() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim pcbNeeded As Long, pcReturned As Long, mi1() As MONITOR_INFO_1, _
  i As Integer, sPortType As String, tColl As Collection
50030
50040  Set tColl = New Collection
50050  EnumMonitors vbNullString, 1, 0, 0, pcbNeeded, pcReturned
50060
50070  If pcbNeeded Then
50080   ReDim mi1((pcbNeeded / SIZEOFMONITOR_INFO_1))
50090   If EnumMonitors(vbNullString, 1, mi1(0), pcbNeeded, pcbNeeded, pcReturned) Then
50100    For i = 0 To (pcReturned - 1)
50110     tColl.Add GetStrFromPtrA(mi1(i).pName)
50120    Next i
50130   End If
50140  End If
50150
50160  Set GetAvailableMonitors = tColl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterEnums", "GetAvailableMonitors")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAvailablePorts() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim pcbNeeded As Long, pcReturned As Long, pi2() As PORT_INFO_2, _
  i As Integer, sPortType As String, tColl As Collection
50030
50040  Set tColl = New Collection
50050  Call EnumPorts(vbNullString, 2, 0, 0, pcbNeeded, pcReturned)
50060  If pcbNeeded Then
50070   ReDim pi2((pcbNeeded / SIZEOFPORT_INFO_2))
50080   If EnumPorts(vbNullString, 2, pi2(0), pcbNeeded, pcbNeeded, pcReturned) Then
50090    For i = 0 To (pcReturned - 1)
50100     tColl.Add GetStrFromPtrA(pi2(i).pPortName)
50110    Next i
50120   End If
50130  End If
50140
50150  Set GetAvailablePorts = tColl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterEnums", "GetAvailablePorts")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAvailablePrinterdrivers() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lngDriverInfo1Level As Long, lngDriverInfo1Needed As Long, _
  lngDriverInfo1Returned As Long, bytDriverInfo1Buffer() As Byte, _
  udtDriverInfo1() As DRIVER_INFO_1, lngDriverInfo1Count As Long, _
  strDriverInfo1Name As String * 128, lngWin32apiResultCode As Long, _
  tColl As Collection
50060
50070  Set tColl = New Collection
50080  lngDriverInfo1Level = 1
50090  lngWin32apiResultCode = EnumPrinterDrivers(vbNullString, vbNullString, _
  lngDriverInfo1Level, ByVal vbNullString, 0, lngDriverInfo1Needed, lngDriverInfo1Returned)
50110  If lngDriverInfo1Needed > 0 Then
50120   ReDim bytDriverInfo1Buffer(lngDriverInfo1Needed - 1)
50130   lngWin32apiResultCode = EnumPrinterDrivers(vbNullString, vbNullString, _
   lngDriverInfo1Level, bytDriverInfo1Buffer(0), lngDriverInfo1Needed, _
   lngDriverInfo1Needed, lngDriverInfo1Returned)
50160   ReDim udtDriverInfo1(lngDriverInfo1Returned - 1)
50170   MoveMemory udtDriverInfo1(0), bytDriverInfo1Buffer(0), Len(udtDriverInfo1(0)) * lngDriverInfo1Returned
50180   For lngDriverInfo1Count = 0 To lngDriverInfo1Returned - 1
50190    lngWin32apiResultCode = lstrcpy(ByVal strDriverInfo1Name, ByVal udtDriverInfo1 _
    (lngDriverInfo1Count).pName)
50210    tColl.Add Left(strDriverInfo1Name, InStr(strDriverInfo1Name, vbNullChar) - 1)
50220   Next lngDriverInfo1Count
50230  End If
50240  Set GetAvailablePrinterdrivers = tColl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterEnums", "GetAvailablePrinterdrivers")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAvailablePrinters() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If IsWin9xMe = True Then
50020    Set GetAvailablePrinters = EnumPrintersWin9x
50030   Else
50040    Set GetAvailablePrinters = EnumPrintersWinNT
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterEnums", "GetAvailablePrinters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function EnumPrintersWinNT() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Success As Boolean, cbRequired As Long, cbBuffer As Long, nEntries As Long, _
  pntr() As PRINTER_INFO_4, c As Long, tColl As Collection
50030
50040  Set tColl = New Collection
50050
50060  Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL4, 0, 0, cbRequired, nEntries)
50080
50090  ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_4))
50100  cbBuffer = cbRequired
50110  If EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL4, pntr(0), cbBuffer, cbRequired, nEntries) Then
50130   For c = 0 To nEntries - 1
50140    tColl.Add GetStrFromPtrA(pntr(c).pPrinterName)
50150   Next c
50160  End If
50170  Set EnumPrintersWinNT = tColl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterEnums", "EnumPrintersWinNT")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function EnumPrintersWin9x() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cbRequired As Long, cbBuffer As Long, pntr() As PRINTER_INFO_1, nEntries As Long, _
  c As Long, sFlags As String, tColl As Collection
50030
50040  Set tColl = New Collection
50050
50060  Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL1, 0, 0, cbRequired, nEntries)
50080
50090  ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_1))
50100  cbBuffer = cbRequired
50110
50120  If EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL1, pntr(0), cbBuffer, cbRequired, nEntries) Then
50140
50150   For c = 0 To nEntries - 1
50160    tColl.Add GetStrFromPtrA(pntr(c).Pane)
50170   Next c
50180  End If
50190  Set EnumPrintersWin9x = tColl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterEnums", "EnumPrintersWin9x")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAvailablePrinters2() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Success As Boolean, cbRequired As Long, cbBuffer As Long, nEntries As Long, _
  pntr() As PRINTER_INFO_2, c As Long, tColl As Collection, dArr(1) As String
50030
50040  Set tColl = New Collection
50050
50060  Call EnumPrinters(PRINTER_ENUM_LOCAL, vbNullString, PRINTER_LEVEL2, 0, 0, cbRequired, nEntries)
50070  ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_2))
50080  cbBuffer = cbRequired
50090  If EnumPrinters(PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL2, pntr(0), cbBuffer, cbRequired, nEntries) Then
50110   For c = 0 To nEntries - 1
50120    dArr(0) = GetStrFromPtrA(pntr(c).pPrinterName)
50130    dArr(1) = GetStrFromPtrA(pntr(c).pPortName)
50140    tColl.Add dArr
50150   Next c
50160  End If
50170  Set GetAvailablePrinters2 = tColl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterEnums", "GetAvailablePrinters2")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

