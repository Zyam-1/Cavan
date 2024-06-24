VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SemenResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolSemenResults As New Collection

Private mSemenResults As New SemenResults

Public Function Count() As Long
  
10    Count = mcolSemenResults.Count

End Function

Public Sub Delete(ByVal SampleID As String, ByVal TestName As String)
  
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM SemenResults50 WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND TestName = '" & TestName & "'"
30    Cnxn(0).Execute sql

40    mcolSemenResults.Remove TestName

50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "SemenResults", "Delete", intEL, strES, sql

End Sub

Public Property Get SemenResults() As SemenResults
  
10    Set SemenResults = mSemenResults

End Property

Public Function Item(ByVal TestName As String) As SemenResult
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolSemenResults.Item(TestName)

End Function

Public Sub Load(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim sx As SemenResult

10    On Error GoTo Load_Error

20    sql = "SELECT SampleID, TestName, Result, " & _
            "UserName, DateTimeOfRecord, COALESCE(Valid, 0) Valid, COALESCE(Printed, 0) Printed, " & _
            "ValidatedBy, ValidatedDateTime, PrintedBy, PrintedDateTime " & _
            "FROM SemenResults50 WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set sx = New SemenResult
70      With sx
80        .SampleID = tb!SampleID & ""
90        .TestName = tb!TestName & ""
100       .Result = tb!Result & ""
110       .Valid = tb!Valid
120       .Printed = tb!Printed
130       .ValidatedBy = tb!ValidatedBy & ""
140       .ValidatedDateTime = tb!ValidatedDateTime & ""
150       .PrintedBy = tb!PrintedBy & ""
160       .PrintedDateTime = tb!PrintedDateTime & ""
170       .UserName = tb!UserName & ""
180       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
190       mcolSemenResults.Add sx, .TestName
200     End With
210     tb.MoveNext
220   Loop

230   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "SemenResults", "Load", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolSemenResults.[_NewEnum]

End Function

Public Sub Save()

      Dim sx As SemenResult

10    On Error GoTo Save_Error

20    For Each sx In mcolSemenResults
30      sx.Save
40    Next

50    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "SemenResults", "Save", intEL, strES

End Sub


