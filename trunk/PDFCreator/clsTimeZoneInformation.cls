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
 DayLight = bDaylight
End Property

Public Property Get DayLightStartdate() As Date
 DayLightStartdate = dDayLightStartdate
End Property

Public Property Get DayLightStopdate() As Date
 DayLightStopdate = dDayLightStopdate
End Property

Public Property Get DayLightDescription() As String
 DayLightDescription = sDayLightDescription
End Property

Public Property Get NormaltimeDescription() As Date
 NormaltimeDescription = sNormaltimeDescription
End Property

Public Property Get DaylightToGMT() As Long
 DaylightToGMT = lDaylightToGMT
End Property

Public Property Get NormaltimeToGMT() As Long
 NormaltimeToGMT = lNormaltimeToGMT
End Property

Private Sub Class_Initialize()
 Dim tzi As TIME_ZONE_INFORMATION, res As Long
  
 dDayLightStartdate = Now
 dDayLightStopdate = Now
 res = GetTimeZoneInformation(tzi)
 Select Case res
  Case TIME_ZONE_ID_UNKNOWN
   bDaylight = False
  Case TIME_ZONE_ID_STANDARD
   bDaylight = False
  Case TIME_ZONE_ID_DAYLIGHT
   bDaylight = True
  Case Else
   bDaylight = False
 End Select
 With tzi
  dDayLightStartdate = GetTimezoneChangeDate(.StandardDate, Year(Now))
  dDayLightStopdate = GetTimezoneChangeDate(.StandardDate, Year(Now))
  sDayLightDescription = Left$(.DaylightName, InStr(.DaylightName, Chr$(0)) - 1)
  sNormaltimeDescription = Left$(.StandardName, InStr(.StandardName, Chr$(0)) - 1)
  lDaylightToGMT = -(.Bias + .DaylightBias)
  lNormaltimeToGMT = -(.Bias + .StandardBias)
 End With
End Sub

Private Function GetTimezoneChangeDate(Data As SYSTEMTIME, InYear As Long) As Date
 Dim tempDate As Date, MonthFirstWeekday As Long
 With Data
  Select Case .wDay
   Case 1 To 4
    MonthFirstWeekday = Weekday(DateSerial(InYear, .wMonth, 1))
    GetTimezoneChangeDate = DateSerial(InYear, .wMonth, _
     .wDayOfWeek - MonthFirstWeekday + .wDay * 7 + 1) + TimeSerial(.wHour, .wMinute, .wSecond)
   Case 5
    tempDate = DateSerial(InYear, .wMonth + 1, 0)
    GetTimezoneChangeDate = DateAdd("d", tempDate, _
     -(Weekday(tempDate) - .wDayOfWeek + 7) Mod 7) + TimeSerial(.wHour, .wMinute, .wSecond)
  End Select
 End With
End Function
