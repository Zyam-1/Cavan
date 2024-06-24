VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LIHs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolLIHs As New Collection

Private mLIHs As New LIHs

Public Function Count() As Long
  
10    Count = mcolLIHs.Count

End Function

Public Property Get LIHs() As LIHs
  
10    Set LIHs = mLIHs

End Property

Public Function Item(ByVal LIorH As String, ByVal Code As String, ByVal NoPrintOrWarning As String) As LIH
              
      Dim i As Integer
10    On Error GoTo Item_Error

20    Load LIorH, Code, NoPrintOrWarning
30    Set Item = Nothing
40    For i = 1 To mcolLIHs.Count
50        If mcolLIHs.Item(i).LIorH = LIorH And mcolLIHs.Item(i).Code = Code And mcolLIHs.Item(i).NoPrintOrWarning = NoPrintOrWarning Then
60            Set Item = mcolLIHs(i)
70        End If
               
80    Next i
90    Exit Function

Item_Error:

       Dim strES As String
       Dim intEL As Integer

100    intEL = Erl
110    strES = Err.Description
120    LogError "LIHs", "Item", intEL, strES
End Function



Public Sub Load(ByVal LIorH As String, ByVal Code As String, ByVal NoPrintOrWarning As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Lx As LIH

10    On Error GoTo Load_Error

20    Set mcolLIHs = New Collection

30    sql = "SELECT LIH, Code, CutOff, NoPrintOrWarning, " & _
            "UserName, DateTimeOfRecord " & _
            "FROM LIHValues " & _
            "WHERE LIH = '" & LIorH & "' AND Code = '" & Code & "' AND NoPrintOrWarning = '" & NoPrintOrWarning & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    Do While Not tb.EOF
70      Set Lx = New LIH
80      With Lx
90        .LIorH = tb!LIH
100       .Code = tb!Code
110       .CutOff = tb!CutOff
120       .NoPrintOrWarning = tb!NoPrintOrWarning
130       .UserName = tb!UserName
140       .DateTimeOfRecord = tb!DateTimeOfRecord
150       mcolLIHs.Add Lx, .LIorH & CStr(.Code) & .NoPrintOrWarning
160     End With
170     tb.MoveNext
180   Loop

190   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "LIHs", "Load", intEL, strES, sql

End Sub
Public Sub Clear()

      Dim n As Integer

10    For n = 1 To mcolLIHs.Count
20      mcolLIHs.Remove 1
30    Next

End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolLIHs.[_NewEnum]

End Function




Public Sub Save()

10    On Error GoTo Save_Error

      Dim Lx As LIH

20    For Each Lx In mcolLIHs
30      Lx.Save
40    Next

50    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "LIHs", "Save", intEL, strES

End Sub




Private Sub Class_Initialize()

'Load

End Sub


