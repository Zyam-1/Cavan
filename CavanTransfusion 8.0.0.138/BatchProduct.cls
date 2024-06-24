VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BatchNumber As String
Public Product As String
Public Identifier As String
Public UnitVolume As String
Public DateExpiry  As String
Public DateReceived As String
Public UnitGroup As String
Public Concentration As String

Public Chart As String
Public PatName As String
Public DoB As String
Public Age As String
Public Sex As String
Public Addr0 As String
Public Addr1 As String
Public Addr2 As String
Public Ward As String
Public Clinician As String

Public PatientGroup As String
Public SampleID As String
Public Typenex As String
Public AandE As String
Public EventCode As String
Public Comment As String
Public EventStart As String
Public EventEnd As String
Public UserName As String
Public RecordDateTime As String

Public LabelPrinted As Integer

