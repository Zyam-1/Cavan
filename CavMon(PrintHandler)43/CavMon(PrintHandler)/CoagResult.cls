VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoagResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RunDate As String
Public RunTime As String
Public SampleID As String
Public Code As String
Public Result As String
Public Flag As String
Public OperatorCode As String
Public Printed As Boolean
Public Valid As Boolean
Public Authorised As Boolean
Public Released As Boolean
Public Low As String
Public High As String
Public Units As String
Public TestName As String
Public Printable As Boolean
Public DP As Integer
Public InUse As Boolean
Public PrintRefRange As Boolean
