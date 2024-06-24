VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaecesResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SampleID As String
Public TestName As String
Public Result As String
Public Valid As Integer
Public ValidatedBy As String
Public ValidatedDateTime As String
Public Printed As Integer
Public PrintedBy As String
Public PrintedDateTime As String
Public HealthLink As Integer
Public UserName As String
Public DateTimeOfRecord As String

