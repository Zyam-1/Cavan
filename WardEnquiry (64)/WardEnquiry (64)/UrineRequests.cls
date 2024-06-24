VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UrineRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolUrineRequests As New Collection

Private mUrineRequests As New UrineRequests

Public Sub Add(ByVal Fx As UrineRequest)

10    mcolUrineRequests.Add Fx, Fx.Request

End Sub

Public Function Count() As Long
  
10    Count = mcolUrineRequests.Count

End Function

Public Sub Delete(ByVal SampleID As String, ByVal Request As String)
  
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM UrineRequests50 WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Request = '" & Request & "'"
30    Cnxn(0).Execute sql

40    mcolUrineRequests.Remove Request

50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "UrineRequests", "Delete", intEL, strES, sql

End Sub
Public Property Get UrineRequests() As UrineRequests
  
10    Set UrineRequests = mUrineRequests

End Property



Public Function Item(ByVal Request As String) As UrineRequest
  


Dim I As Integer
On Error GoTo Item_Error

Set Item = Nothing
For I = 1 To mcolUrineRequests.Count
    If mcolUrineRequests.Item(I).Request = Request Then
        Set Item = mcolUrineRequests(I)
    End If
Next I

Exit Function

Item_Error:

 Dim strES As String
 Dim intEL As Integer

 intEL = Erl
 strES = Err.Description
 LogError "UrineRequests", "Item", intEL, strES
    

End Function


Public Sub Load(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Fx As UrineRequest

10    On Error GoTo Load_Error

20    sql = "SELECT SampleID, Request, UserName, DateTimeOfRecord " & _
            "FROM UrineRequests50 " & _
            "WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Fx = New UrineRequest
70      With Fx
80        .SampleID = tb!SampleID & ""
90        .Request = tb!Request & ""
100       .UserName = tb!UserName & ""
110       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
120       mcolUrineRequests.Add Fx, .Request
130     End With
140     tb.MoveNext
150   Loop

160   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "UrineRequests", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolUrineRequests.[_NewEnum]

End Function

Public Sub Save(ByVal SampleID As String)

      Dim sql As String
      Dim Fx As UrineRequest

10    On Error GoTo Save_Error

20    sql = "Delete from UrineRequests50 where " & _
            "SampleID = '" & SampleID & "'"
30    Cnxn(0).Execute sql

40    For Each Fx In mcolUrineRequests

50      sql = "INSERT INTO UrineRequests50 (SampleID, Request, UserName) " & _
              "VALUES " & _
              "('" & Fx.SampleID & "', " & _
              " '" & Fx.Request & "', " & _
              " '" & AddTicks(Fx.UserName) & "')"
60      Cnxn(0).Execute sql
70    Next

80    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "UrineRequests", "Save", intEL, strES, sql

End Sub




