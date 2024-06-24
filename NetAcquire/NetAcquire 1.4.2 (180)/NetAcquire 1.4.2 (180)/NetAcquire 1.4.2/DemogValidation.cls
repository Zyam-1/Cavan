VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DemogValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_SampleID As String
Private m_EnteredBy As String
Private m_ValidatedBy As String
Private m_EnteredDateTime As String
Private m_ValidatedDateTime As String


Public Property Get SampleID() As String

11220     SampleID = m_SampleID

End Property

Public Property Let SampleID(ByVal Value As String)

11230     m_SampleID = Value

End Property

Public Property Get EnteredBy() As String

11240     EnteredBy = m_EnteredBy

End Property

Public Property Let EnteredBy(ByVal Value As String)

11250     m_EnteredBy = Value

End Property

Public Property Get ValidatedBy() As String

11260     ValidatedBy = m_ValidatedBy

End Property

Public Property Let ValidatedBy(ByVal Value As String)

11270     m_ValidatedBy = Value

End Property

Public Property Get EnteredDateTime() As String

11280     EnteredDateTime = m_EnteredDateTime

End Property

Public Property Let EnteredDateTime(ByVal Value As String)

11290     m_EnteredDateTime = Value

End Property

Public Property Get ValidatedDateTime() As String

11300     ValidatedDateTime = m_ValidatedDateTime

End Property

Public Property Let ValidatedDateTime(ByVal Value As String)

11310     m_ValidatedDateTime = Value

End Property
