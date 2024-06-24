VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private mcolEventCodes As New Collection

Private mEventCodes As New EventCodes

Public Function CodeFor(ByVal Text As String) As String

      Dim retval As String
      Dim ev As EventCode

10    retval = " "

20    Text = UCase$(Trim$(Text))

30    For Each ev In mcolEventCodes
40      If UCase$(ev.Text) = Text Then
50        retval = ev.code
60        Exit For
70      End If
80    Next

90    CodeFor = retval
    
End Function

Public Function Count() As Long
  
10    On Error Resume Next
20    Count = mcolEventCodes.Count

End Function

Public Function Item(ByVal Index As String) As EventCode
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolEventCodes.Item(Index)

End Function

Public Sub Load()

      Dim tb As Recordset
      Dim sql As String
      Dim ev As EventCode
Dim Trial As EventCode
Dim Found As Boolean

10    On Error GoTo Load_Error

20    sql = "SELECT * FROM Lists " & _
            "WHERE ListType = 'EventBarCodes' "
30    Set tb = New Recordset

40    RecOpenClientBB 0, tb, sql
50    Do While Not tb.EOF

60      Set ev = New EventCode
70      ev.code = Trim$(tb!code & "")
80      ev.Text = Trim$(tb!Text & "")
  
90      Found = False
100     For Each Trial In mcolEventCodes
110       If Trial.code = ev.code Then
120         Found = True
130         Exit For
140       End If
150     Next
  
160     If Not Found Then
170       mcolEventCodes.Add ev, ev.code
180     End If
190     tb.MoveNext
200   Loop

210   Exit Sub

Load_Error:

Dim strES As String
Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "EventCodes", "Load", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolEventCodes.[_NewEnum]

End Function



Public Property Get EventCodes() As EventCodes
  
10    Set EventCodes = mEventCodes

End Property

Private Sub Class_Initialize()

10    Load

End Sub


