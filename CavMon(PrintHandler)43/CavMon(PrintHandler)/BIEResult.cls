VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SampleID As String
Public Code As String
Public Result As String
Public Operator As String
Public Units As String
Public Rundate As String
Public RunTime As String
Public Flags As String
Public SampleType As String
Public Analyser As String
Public Low As Single
Public FlagLow As Single
Public PlausibleLow As Single
Public High As Single
Public FlagHigh As Single
Public PlausibleHigh As Single
Public LongName As String
Public ShortName As String
Public DoDelta As Boolean
Public DeltaLimit As Single
Public Category As String

Public Valid As Boolean
Public Printed As Boolean

Public h As Boolean
Public s As Boolean
Public o As Boolean
Public l As Boolean
Public g As Boolean
Public J As Boolean

Public Authorised As Boolean

Public NormalLow As Single
Public NormalHigh As Single
Public NormalUsed As String

Public Comment As String

Public DefIndex As Long

Public PrintSplit As Integer

Public PrintRefRange As Boolean
Public Printformat As Integer

Public Printable As Boolean
