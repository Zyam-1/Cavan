VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolGenericResults As New Collection

Private mGenericResults As New GenericResults

Public Function Count() As Long
  
10    Count = mcolGenericResults.Count

End Function




Public Sub Delete(ByVal SampleID As String, ByVal TestName As String)
  
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM GenericResults WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND TestName = '" & TestName & "'"
30    Cnxn(0).Execute sql

40    mcolGenericResults.Remove TestName

50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "GenericResults", "Delete", intEL, strES, sql

End Sub



Public Property Get GenericResults() As GenericResults
  
10    Set GenericResults = mGenericResults

End Property


Public Function Item(ByVal TestName As String) As GenericResult
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolGenericResults.Item(TestName)

End Function




Public Sub Load(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Gx As GenericResult

10    On Error GoTo Load_Error

20    sql = "SELECT G.SampleID, G.TestName, G.Result, COALESCE(G.HealthLink, 0) Healthlink, " & _
            "G.UserName, G.DateTimeOfRecord, COALESCE(P.Valid, 0) Valid, COALESCE(P.Printed, 0) Printed, " & _
            "P.ValidatedBy, P.ValidatedDateTime, P.PrintedBy, P.PrintedDateTime " & _
            "FROM GenericResults G LEFT JOIN PrintValidLog P " & _
            "ON G.SampleID = P.SampleID " & _
            "WHERE G.SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Gx = New GenericResult
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
200       mcolGenericResults.Add Gx, .TestName
210     End With
220     tb.MoveNext
230   Loop

240   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "GenericResults", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  
10    Set NewEnum = mcolGenericResults.[_NewEnum]

End Function



Public Sub Save(ByVal Gx As GenericResult)

      Dim sql As String

10    On Error GoTo Save_Error

20    If Trim$(Gx.Result) = "" Then
30      sql = "Delete from GenericResults where " & _
              "SampleID = '" & Gx.SampleID & "' " & _
              "and TestName = '" & Gx.TestName & "'"
40    Else
50      sql = "IF EXISTS (SELECT * FROM GenericResults WHERE " & _
              "           SampleID = '" & Gx.SampleID & "' " & _
              "           AND TestName = '" & Gx.TestName & "') " & _
              "  UPDATE GenericResults " & _
              "  SET Result = '" & Gx.Result & "', " & _
              "  UserName = '" & AddTicks(Gx.UserName) & "', " & _
              "  HealthLink = '" & Gx.Healthlink & "', " & _
              "  Valid = '" & Gx.Valid & "', " & _
              "  ValidatedBy = '" & AddTicks(Gx.ValidatedBy) & "', " & _
              "  ValidatedDateTime = " & IIf(IsDate(Gx.ValidatedDateTime), Format$(Gx.ValidatedDateTime, "'dd/MMM/yyyy HH:nn:ss'"), "Null") & " " & _
              "  WHERE SampleID = '" & Gx.SampleID & "' " & _
              "  AND TestName = '" & Gx.TestName & "' "
60    sql = sql & "ELSE " & _
              "  INSERT INTO GenericResults (SampleID, TestName, Result, UserName, HealthLink, Valid, ValidatedBy, ValidatedDateTime ) " & _
              "  VALUES " & _
              "  ('" & Gx.SampleID & "', " & _
              "  '" & Gx.TestName & "', " & _
              "  '" & Gx.Result & "', " & _
              "  '" & AddTicks(Gx.UserName) & "', " & _
              "  '" & Gx.Healthlink & "', " & _
              "  '" & Gx.Valid & "', " & _
              "  '" & AddTicks(Gx.ValidatedBy) & "', " & _
              "  " & IIf(IsDate(Gx.ValidatedDateTime), Format$(Gx.ValidatedDateTime, "'dd/MMM/yyyy HH:nn:ss'"), "Null") & ")"
70    End If
80    Cnxn(0).Execute sql

90    sql = "IF EXISTS (SELECT * FROM PrintValidLog WHERE " & _
            "           SampleID = '" & Gx.SampleID & "' " & _
            "           AND Department = 'M') " & _
            "  UPDATE PrintValidLog " & _
            "  SET Valid = '" & Gx.Valid & "', " & _
            "  Printed = '" & Gx.Printed & "', " & _
            "  ValidatedBy = '" & AddTicks(Gx.ValidatedBy) & "', " & _
            "  PrintedBy = '" & AddTicks(Gx.PrintedBy) & "', " & _
            "  ValidatedDateTime = '" & Format$(Gx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
            "  PrintedDateTime = '" & Format$(Gx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "' " & _
            "  WHERE SampleID = '" & Gx.SampleID & "' AND Department = 'M' " & _
            "ELSE " & _
            "  INSERT INTO PrintValidLog " & _
            "  (SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime) " & _
            "  VALUES " & _
            "  ('" & Gx.SampleID & "', 'M', '" & Gx.Printed & "', '" & Gx.Valid & "', " & _
            "  '" & AddTicks(Gx.PrintedBy) & "', '" & Format$(Gx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
            "  '" & AddTicks(Gx.ValidatedBy) & "', '" & Format$(Gx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "')"
100   Cnxn(0).Execute sql

110   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "GenericResults", "Save", intEL, strES, sql

End Sub


