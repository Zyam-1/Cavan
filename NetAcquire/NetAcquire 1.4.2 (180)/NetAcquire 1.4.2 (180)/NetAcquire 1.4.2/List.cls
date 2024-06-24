VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Code As String
Private m_Text As String
Private m_ListType As String
Private m_ListOrder As Integer
Private m_InUse As Boolean
Private m_Default As String


Public Sub GetCode(ByVal ListType As String, ByVal Text As String)

          Dim tb As Recordset
          Dim sql As String
          '      Dim Lx As New List

21640     On Error GoTo GetCode_Error

21650     sql = "SELECT Code, COALESCE(Text, '') Text, " & _
              "COALESCE(ListOrder, 999) ListOrder, " & _
              "COALESCE(InUse, 0) InUse, " & _
              "COALESCE([Default], '') DefaultValue " & _
              "FROM Lists " & _
              "WHERE ListType = '" & ListType & "' " & _
              "AND Text = '" & AddTicks(Text) & "'"
21660     Set tb = New Recordset
21670     RecOpenClient 0, tb, sql
21680     If Not tb.EOF Then
21690         m_Code = tb!Code & ""
21700         m_Text = tb!Text
21710         m_ListType = ListType
21720         m_ListOrder = tb!ListOrder
21730         m_InUse = tb!InUse
21740         m_Default = tb!DefaultValue
21750     End If

21760     Exit Sub

GetCode_Error:

          Dim strES As String
          Dim intEL As Integer

21770     intEL = Erl
21780     strES = Err.Description
21790     LogError "List", "GetCode", intEL, strES, sql

End Sub


Public Property Get Code() As String

21800     Code = m_Code

End Property

Public Property Let Code(ByVal sNewValue As String)

21810     m_Code = sNewValue

End Property

Public Property Get Text() As String

21820     Text = m_Text

End Property

Public Property Get ListType() As String

21830     ListType = m_ListType

End Property

Public Property Let ListType(ByVal sNewValue As String)

21840     m_ListType = sNewValue

End Property
Public Property Let Text(ByVal sNewValue As String)

21850     m_Text = sNewValue

End Property


Public Property Get ListOrder() As Integer

21860     ListOrder = m_ListOrder

End Property

Public Property Let ListOrder(ByVal iNewValue As Integer)

21870     m_ListOrder = iNewValue

End Property

Public Property Get InUse() As Boolean

21880     InUse = m_InUse

End Property

Public Property Let InUse(ByVal bNewValue As Boolean)

21890     m_InUse = bNewValue

End Property

Public Property Get DefaultValue() As String

21900     DefaultValue = m_Default

End Property

Public Property Let DefaultValue(ByVal sNewValue As String)

21910     m_Default = sNewValue

End Property
