VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Code As String
Private m_Text As String

Public Property Get code() As String

10    code = m_Code

End Property
Public Property Get Text() As String

10    Text = m_Text

End Property

Public Property Let code(ByVal sNewValue As String)

10    m_Code = sNewValue

End Property


Public Property Let Text(ByVal sNewValue As String)

10    m_Text = sNewValue

End Property


