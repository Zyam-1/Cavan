VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaecesResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolFaecesResults As New Collection

Private mFaecesResults As New FaecesResults

Public Function Count() As Long
  
10    Count = mcolFaecesResults.Count

End Function





Public Sub Delete(ByVal SampleID As String, ByVal TestName As String)
  
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM FaecesResults50 WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND TestName = '" & TestName & "'"
30    Cnxn(0).Execute sql

40    mcolFaecesResults.Remove TestName

50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "FaecesResults", "Delete", intEL, strES, sql

End Sub





Public Property Get FaecesResults() As FaecesResults
  
10    Set FaecesResults = mFaecesResults

End Property



Public Function Item(ByVal TestName As String) As FaecesResult
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolFaecesResults.Item(TestName)

End Function





Public Sub Load(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Gx As FaecesResult

10    On Error GoTo Load_Error

20    sql = "SELECT G.SampleID, G.TestName, G.Result, COALESCE(G.HealthLink, 0) Healthlink, " & _
            "G.UserName, G.DateTimeOfRecord, COALESCE(P.Valid, 0) Valid, COALESCE(P.Printed, 0) Printed, " & _
            "P.ValidatedBy, P.ValidatedDateTime, P.PrintedBy, P.PrintedDateTime " & _
            "FROM FaecesResults50 G LEFT JOIN PrintValidLog P " & _
            "ON G.SampleID = P.SampleID " & _
            "WHERE G.SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Gx = New FaecesResult
70      With Gx
80        .SampleID = tb!SampleID & ""
90        .TestName = tb!TestName & ""
100       .Result = tb!Result & ""
110       .Healthlink = tb!Healthlink
120       .Valid = tb!Valid
130       .Printed = tb!Printed
140       .ValidatedBy = tb!ValidatedBy & ""
150       .ValidatedDateTime = tb!ValidatedDateTime & ""
160       .PrintedBy = tb!PrintedBy & ""
170       .PrintedDateTime = tb!PrintedDateTime & ""
180       .UserName = tb!UserName & ""
190       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
200       mcolFaecesResults.Add Gx, .TestName
210     End With
220     tb.MoveNext
230   Loop

240   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "FaecesResults", "Load", intEL, strES, sql


End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolFaecesResults.[_NewEnum]

End Function


Public Sub Save(ByVal Fx As FaecesResult)

      Dim sql As String
Dim VDate As String
Dim PDate As String
Dim VBy As String
Dim PBy As String

10    On Error GoTo Save_Error

20    VDate = IIf(IsDate(Fx.ValidatedDateTime), Format$(Fx.ValidatedDateTime, " 'dd/MMM/yyyy HH:nn:ss'"), "Null")
30    PDate = IIf(IsDate(Fx.PrintedDateTime), Format$(Fx.PrintedDateTime, "'dd/MMM/yyyy HH:nn:ss'"), "Null")
40    VBy = AddTicks(Fx.ValidatedBy)
50    PBy = AddTicks(Fx.PrintedBy)

60    If Trim$(Fx.Result) = "" Then
70      sql = "Delete from FaecesResults50 where " & _
              "SampleID = '" & Fx.SampleID & "' " & _
              "and TestName = '" & Fx.TestName & "'"
80    Else
90      sql = "IF EXISTS (SELECT * FROM FaecesResults50 WHERE " & _
              "           SampleID = '" & Fx.SampleID & "' " & _
              "           AND TestName = '" & Fx.TestName & "') " & _
              "  UPDATE FaecesResults50 " & _
              "  SET Result = '" & Fx.Result & "', " & _
              "  Valid = '" & Fx.Valid & "', " & _
              "  ValidatedBy = '" & VBy & "', " & _
              "  ValidatedDateTime = " & VDate & ", " & _
              "  Printed = '" & Fx.Printed & "', " & _
              "  PrintedBy = '" & PBy & "', " & _
              "  PrintedDateTime = " & IIf(IsDate(Fx.PrintedDateTime), Format$(Fx.PrintedDateTime, "'dd/MMM/yyyy HH:nn:ss'"), "Null") & ", " & _
              "  UserName = '" & AddTicks(Fx.UserName) & "', " & _
              "  HealthLink = '" & Fx.Healthlink & "' " & _
              "  WHERE SampleID = '" & Fx.SampleID & "' " & _
              "  AND TestName = '" & Fx.TestName & "' "
100    sql = sql & "ELSE " & _
              "  INSERT INTO FaecesResults50 (SampleID, TestName, Result, UserName, HealthLink, Valid, ValidatedBy, ValidatedDateTime, Printed, PrintedBy, PrintedDateTime) " & _
              "  VALUES " & _
              "  ('" & Fx.SampleID & "', " & _
              "  '" & Fx.TestName & "', " & _
              "  '" & Fx.Result & "', " & _
              "  '" & AddTicks(Fx.UserName) & "', " & _
              "  '" & Fx.Healthlink & "', " & _
              "  '" & Fx.Valid & "', " & _
              "  '" & VBy & "', " & _
              VDate & ", " & _
              "  '" & Fx.Printed & "', " & _
              "  '" & PBy & "', " & _
              PDate & ")"
110   End If
120   Cnxn(0).Execute sql

130   sql = "IF EXISTS (SELECT * FROM PrintValidLog WHERE " & _
            "           SampleID = '" & Fx.SampleID & "' " & _
            "           AND Department = 'M') " & _
            "  UPDATE PrintValidLog " & _
            "  SET Valid = '" & Fx.Valid & "', " & _
            "  Printed = '" & Fx.Printed & "', " & _
            "  ValidatedBy = '" & AddTicks(Fx.ValidatedBy) & "', " & _
            "  PrintedBy = '" & AddTicks(Fx.PrintedBy) & "', " & _
            "  ValidatedDateTime = '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
            "  PrintedDateTime = '" & Format$(Fx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "' " & _
            "  WHERE SampleID = '" & Fx.SampleID & "' AND Department = 'M' " & _
            "ELSE " & _
            "  INSERT INTO PrintValidLog " & _
            "  (SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime) " & _
            "  VALUES " & _
            "  ('" & Fx.SampleID & "', 'M', '" & Fx.Printed & "', '" & Fx.Valid & "', " & _
            "  '" & AddTicks(Fx.PrintedBy) & "', '" & Format$(Fx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
            "  '" & AddTicks(Fx.ValidatedBy) & "', '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "')"
140   Cnxn(0).Execute sql

150   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "FaecesResults", "Save", intEL, strES, sql

End Sub




