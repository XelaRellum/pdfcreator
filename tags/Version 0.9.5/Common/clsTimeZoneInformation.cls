VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeZoneInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bDaylight As Boolean
Private dDayLightStartdate As Boolean
Private dDayLightStopdate As Boolean
Private sDayLightDescription As String
Private sNormaltimeDescription As String
Private lDaylightToGMT As Long
Private lNormaltimeToGMT As Long

Public Property Get DayLight() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DayLight = bDaylight
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "DayLight [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get DayLightStartdate() As Date
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DayLightStartdate = dDayLightStartdate
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "DayLightStartdate [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get DayLightStopdate() As Date
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DayLightStopdate = dDayLightStopdate
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "DayLightStopdate [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get DayLightDescription() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DayLightDescription = sDayLightDescription
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "DayLightDescription [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get NormaltimeDescription() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  NormaltimeDescription = sNormaltimeDescription
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "NormaltimeDescription [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get DaylightToGMT() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DaylightToGMT = lDaylightToGMT
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "DaylightToGMT [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get NormaltimeToGMT() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  NormaltimeToGMT = lNormaltimeToGMT
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "NormaltimeToGMT [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub Class_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tzi As TIME_ZONE_INFORMATION, res As Long
50020
50030  dDayLightStartdate = Now
50040  dDayLightStopdate = Now
50050  res = GetTimeZoneInformation(tzi)
50061  Select Case res
        Case TIME_ZONE_ID_UNKNOWN
50080    bDaylight = False
50090   Case TIME_ZONE_ID_STANDARD
50100    bDaylight = False
50110   Case TIME_ZONE_ID_DAYLIGHT
50120    bDaylight = True
50130   Case Else
50140    bDaylight = False
50150  End Select
50160  With tzi
50170   dDayLightStartdate = GetTimezoneChangeDate(.StandardDate, Year(Now))
50180   dDayLightStopdate = GetTimezoneChangeDate(.StandardDate, Year(Now))
50190   sDayLightDescription = Left$(.DaylightName, InStr(.DaylightName, Chr$(0)) - 1)
50200   sNormaltimeDescription = Left$(.StandardName, InStr(.StandardName, Chr$(0)) - 1)
50210   lDaylightToGMT = -(.Bias + .DaylightBias)
50220   lNormaltimeToGMT = -(.Bias + .StandardBias)
50230  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "Class_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetTimezoneChangeDate(Data As SYSTEMTIME, InYear As Long) As Date
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tempDate As Date, MonthFirstWeekday As Long
50020  With Data
50031   Select Case .wDay
         Case 1 To 4
50050     MonthFirstWeekday = Weekday(DateSerial(InYear, .wMonth, 1))
50060     GetTimezoneChangeDate = DateSerial(InYear, .wMonth, _
     .wDayOfWeek - MonthFirstWeekday + .wDay * 7 + 1) + TimeSerial(.wHour, .wMinute, .wSecond)
50080    Case 5
50090     tempDate = DateSerial(InYear, .wMonth + 1, 0)
50100     GetTimezoneChangeDate = DateAdd("d", tempDate, _
     -(Weekday(tempDate) - .wDayOfWeek + 7) Mod 7) + TimeSerial(.wHour, .wMinute, .wSecond)
50120   End Select
50130  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTimeZoneInformation", "GetTimezoneChangeDate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
