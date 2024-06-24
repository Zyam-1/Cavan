VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaecesRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolFaecesRequests As New Collection

Private mFaecesRequests As New FaecesRequests

Public Sub Add(ByVal Fx As FaecesRequest)

10    mcolFaecesRequests.Add Fx, Fx.Request

End Sub


Public Function Count() As Long
        
10    Count = mcolFaecesRequests.Count

End Function




Public Sub Delete(ByVal SampleID As String, ByVal Request As String)
        
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM FaecesRequests50 WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Request = '" & Request & "'"
30    Cnxn(0).Execute sql

40    mcolFaecesRequests.Remove Request

50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "FaecesRequests", "Delete", intEL, strES, sql

End Sub

Public Property Get FaecesRequests() As FaecesRequests
        
10    Set FaecesRequests = mFaecesRequests

End Property


Public Function Item(ByVal Request As String) As FaecesRequest
        
10    On Error Resume Next

20    Set Item = mcolFaecesRequests.Item(Request)

End Function





Public Sub Load(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Fx As FaecesRequest

10    On Error GoTo Load_Error

20    sql = "SELECT SampleID, Request, UserName, DateTimeOfRecord, Analyser, COALESCE(Programmed, 0) Programmed " & _
            "FROM FaecesRequests50 " & _
            "WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Fx = New FaecesRequest
70      With Fx
80        .SampleID = tb!SampleID & ""
90        .Request = tb!Request & ""
100       .UserName = tb!UserName & ""
110       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
120       .Analyser = tb!Analyser & ""
130       .Programmed = CBool(tb!Programmed)
140       mcolFaecesRequests.Add Fx, .Request
150     End With
160     tb.MoveNext
170   Loop

180   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "FaecesRequests", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
10    Set NewEnum = mcolFaecesRequests.[_NewEnum]

End Function




Public Sub Save(ByVal SampleID As String)

      Dim sql As String
      Dim Fx As FaecesRequest

10    On Error GoTo Save_Error

20    sql = "Delete from FaecesRequests50 where " & _
            "SampleID = '" & SampleID & "'"
30    Cnxn(0).Execute sql

40    For Each Fx In mcolFaecesRequests

50      sql = "INSERT INTO FaecesRequests50 (SampleID, Request, UserName, Analyser, Programmed) " & _
              "VALUES " & _
              "('" & Fx.SampleID & "', " & _
              " '" & Fx.Request & "', " & _
              " '" & Fx.UserName & "', " & _
              " '" & Fx.Analyser & "', " & _
              " '" & Fx.Programmed & "')"
60      Cnxn(0).Execute sql
70    Next

80    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "FaecesRequests", "Save", intEL, strES, sql

End Sub



