VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LongName        As String
Public ShortName       As String
Public DoDelta         As Integer
Public DeltaLimit      As Single
Public PrintPriority   As Integer
Public DP              As Integer
Public BarCode         As String
Public Units           As String
Public H               As Integer
Public S               As Integer
Public l               As Integer
Public o               As Integer
Public g               As Integer
Public J               As Integer
Public Category        As String
Public Code            As String
Public Printable       As Boolean
Public PlausibleLow    As Single
Public PlausibleHigh   As Single
Public KnownToAnalyser As Boolean
Public SampleType      As String
Public InUse           As Boolean
Public MaleLow         As Single
Public MaleHigh        As Single
Public FemaleLow       As Single
Public FemaleHigh      As Single
Public FlagMaleLow     As Single
Public FlagMaleHigh    As Single
Public FlagFemaleLow   As Single
Public FlagFemaleHigh  As Single
Public LControlLow     As Single
Public LControlHigh    As Single
Public NControlLow     As Single
Public NControlHigh    As Single
Public HControlLow     As Single
Public HControlHigh    As Single
Public AgeFromDays     As Long
Public AgeToDays       As Long
Public AutoValLow      As Single
Public AutoValHigh     As Single
Public Hospital        As String
Public Analyser        As String
Public ImmunoCode      As String
Public SplitList       As Integer
Public EOD             As Boolean
Public LIH             As Integer
'Public ActiveFromDate  As String
'Public ActiveToDate    As String
Public ArchitectCode   As String
Public BIE             As String
Public ReRunDays       As Long
Public AgeFromText     As String
Public AgeToText       As String
Public PrintRefRange   As Integer
Public PrintSplit      As Integer



