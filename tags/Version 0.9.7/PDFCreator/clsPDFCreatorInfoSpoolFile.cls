VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFCreatorInfoSpoolFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public REDMON_PORT As String
Public REDMON_JOB As String
Public REDMON_PRINTER As String
Public REDMON_MACHINE As String
Public REDMON_USER As String
Public REDMON_DOCNAME As String
Public REDMON_FILENAME As String
Public REDMON_SESSIONID As String
Public SpoolFilename As String
Public SpoolerAccount As String
Public Computer As String
Public Created As String
