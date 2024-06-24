VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PassWord As String
Private m_UserName As String
Private m_UserCode As String
Private m_InUse As Boolean
Private m_MemberOf As String
Private m_LogOffDelaySec As Long
Private m_ListOrder As Integer
Private m_ExpiryDate As String

Public Property Get PassWord() As String

  PassWord = m_PassWord

End Property

Public Property Let PassWord(ByVal PassWord As String)

  m_PassWord = PassWord

End Property

Public Property Get UserName() As String

  UserName = m_UserName

End Property

Public Property Let UserName(ByVal UserName As String)

  m_UserName = UserName

End Property

Public Property Get UserCode() As String

  UserCode = m_UserCode

End Property

Public Property Let UserCode(ByVal UserCode As String)

  m_UserCode = UserCode

End Property

Public Property Get InUse() As Boolean

  InUse = m_InUse

End Property

Public Property Let InUse(ByVal InUse As Boolean)

  m_InUse = InUse

End Property

Public Property Get MemberOf() As String

  MemberOf = m_MemberOf

End Property

Public Property Let MemberOf(ByVal MemberOf As String)

  m_MemberOf = MemberOf

End Property

Public Property Get LogOffDelaySec() As Long

  LogOffDelaySec = m_LogOffDelaySec

End Property

Public Property Let LogOffDelaySec(ByVal LogOffDelaySec As Long)

  m_LogOffDelaySec = LogOffDelaySec

End Property

Public Property Get ListOrder() As Integer

  ListOrder = m_ListOrder

End Property

Public Property Let ListOrder(ByVal ListOrder As Integer)

  m_ListOrder = ListOrder

End Property

Public Property Get ExpiryDate() As String

  ExpiryDate = m_ExpiryDate

End Property

Public Property Let ExpiryDate(ByVal ExpiryDate As String)

  m_ExpiryDate = ExpiryDate

End Property
