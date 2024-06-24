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
        
12530     Count = mcolFaecesResults.Count

End Function





Public Sub Delete(ByVal SampleID As String, ByVal TestName As String)
        
          Dim sql As String

12540     On Error GoTo Delete_Error

12550     sql = "DELETE FROM FaecesResults50 WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND TestName = '" & TestName & "'"
12560     Cnxn(0).Execute sql

12570     mcolFaecesResults.Remove TestName

12580     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

12590     intEL = Erl
12600     strES = Err.Description
12610     LogError "FaecesResults", "Delete", intEL, strES, sql

End Sub





Public Property Get FaecesResults() As FaecesResults
        
12620     Set FaecesResults = mFaecesResults

End Property



Public Function Item(ByVal TestName As String) As FaecesResult
Attribute Item.VB_UserMemId = 0
        
12630     On Error Resume Next

12640     Set Item = mcolFaecesResults.Item(TestName)

End Function





Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Gx As FaecesResult

12650     On Error GoTo Load_Error

12660     sql = "SELECT G.SampleID, G.TestName, G.Result, COALESCE(G.HealthLink, 0) Healthlink, " & _
              "G.UserName, G.DateTimeOfRecord, COALESCE(P.Valid, 0) Valid, COALESCE(P.Printed, 0) Printed, " & _
              "P.ValidatedBy, P.ValidatedDateTime, P.PrintedBy, P.PrintedDateTime " & _
              "FROM FaecesResults50 G LEFT JOIN PrintValidLog P " & _
              "ON G.SampleID = P.SampleID " & _
              "WHERE G.SampleID = '" & SampleID & "'"
12670     Set tb = New Recordset
12680     RecOpenClient 0, tb, sql
12690     Do While Not tb.EOF
12700         Set Gx = New FaecesResult
12710         With Gx
12720             .SampleID = tb!SampleID & ""
12730             .TestName = tb!TestName & ""
12740             .Result = tb!Result & ""
12750             .HealthLink = tb!HealthLink
12760             .Valid = tb!Valid
12770             .Printed = tb!Printed
12780             .ValidatedBy = tb!ValidatedBy & ""
12790             .ValidatedDateTime = tb!ValidatedDateTime & ""
12800             .PrintedBy = tb!PrintedBy & ""
12810             .PrintedDateTime = tb!PrintedDateTime & ""
12820             .UserName = tb!UserName & ""
12830             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
12840             mcolFaecesResults.Add Gx, .TestName
12850         End With
12860         tb.MoveNext
12870     Loop

12880     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

12890     intEL = Erl
12900     strES = Err.Description
12910     LogError "FaecesResults", "Load", intEL, strES, sql


End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
12920     Set NewEnum = mcolFaecesResults.[_NewEnum]

End Function


Public Sub Save(ByVal Fx As FaecesResult)

          Dim sql As String
          Dim VDate As String
          Dim PDate As String
          Dim VBy As String
          Dim PBy As String

12930     On Error GoTo Save_Error

12940     VDate = IIf(IsDate(Fx.ValidatedDateTime), Format$(Fx.ValidatedDateTime, " 'dd/MMM/yyyy HH:nn:ss'"), "Null")
12950     PDate = IIf(IsDate(Fx.PrintedDateTime), Format$(Fx.PrintedDateTime, "'dd/MMM/yyyy HH:nn:ss'"), "Null")
12960     VBy = AddTicks(Fx.ValidatedBy)
12970     PBy = AddTicks(Fx.PrintedBy)

12980     If Trim$(Fx.Result) = "" Then
              '70      Sql = "Delete from FaecesResults50 where " & _
              '              "SampleID = '" & Fx.SampleID & "' " & _
              '              "and TestName = '" & Fx.TestName & "'"
              '        Exit Sub
12990     Else
13000         sql = "IF EXISTS (SELECT * FROM FaecesResults50 WHERE " & _
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
                  "  HealthLink = '" & Fx.HealthLink & "' " & _
                  "  WHERE SampleID = '" & Fx.SampleID & "' " & _
                  "  AND TestName = '" & Fx.TestName & "' "
13010         sql = sql & "ELSE " & _
                  "  INSERT INTO FaecesResults50 (SampleID, TestName, Result, UserName, HealthLink, Valid, ValidatedBy, ValidatedDateTime, Printed, PrintedBy, PrintedDateTime) " & _
                  "  VALUES " & _
                  "  ('" & Fx.SampleID & "', " & _
                  "  '" & Fx.TestName & "', " & _
                  "  '" & Fx.Result & "', " & _
                  "  '" & AddTicks(Fx.UserName) & "', " & _
                  "  '" & Fx.HealthLink & "', " & _
                  "  '" & Fx.Valid & "', " & _
                  "  '" & VBy & "', " & _
                  VDate & ", " & _
                  "  '" & Fx.Printed & "', " & _
                  "  '" & PBy & "', " & _
                  PDate & ")"
13020         Cnxn(0).Execute sql
13030     End If

          '+++ Junaid 17+01-2024
          '130   Sql = "IF EXISTS (SELECT * FROM PrintValidLog WHERE " & _
          '            "           SampleID = '" & Fx.SampleID & "' " & _
          '            "           AND Department = 'M') " & _
          '            "  UPDATE PrintValidLog " & _
          '            "  SET Valid = '" & Fx.Valid & "', " & _
          '            "  Printed = '" & Fx.Printed & "', " & _
          '            "  ValidatedBy = '" & AddTicks(Fx.ValidatedBy) & "', " & _
          '            "  PrintedBy = '" & AddTicks(Fx.PrintedBy) & "', " & _
          '            "  ValidatedDateTime = '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
          '            "  PrintedDateTime = '" & Format$(Fx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "' " & _
          '            "  WHERE SampleID = '" & Fx.SampleID & "' AND Department = 'M' " & _
          '            "ELSE " & _
          '            "  INSERT INTO PrintValidLog " & _
          '            "  (SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime) " & _
          '            "  VALUES " & _
          '            "  ('" & Fx.SampleID & "', 'M', '" & Fx.Printed & "', '" & Fx.Valid & "', " & _
          '            "  '" & AddTicks(Fx.PrintedBy) & "', '" & Format$(Fx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
          '            "  '" & AddTicks(Fx.ValidatedBy) & "', '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "')"
          '140   Cnxn(0).Execute Sql

13040     sql = "IF EXISTS (SELECT * FROM PrintValidLog WHERE " & _
              "           SampleID = '" & Fx.SampleID & "' " & _
              "           AND Department = 'M') " & _
              "  UPDATE PrintValidLog " & _
              "  SET Valid = '" & Fx.Valid & "', " & _
              "  Printed = '" & Fx.Printed & "', " & _
              "  ValidatedBy = '" & AddTicks(Fx.ValidatedBy) & "', " & _
              "  PrintedBy = '" & AddTicks(Fx.ValidatedBy) & "', " & _
              "  ValidatedDateTime = '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
              "  PrintedDateTime = '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "' " & _
              "  WHERE SampleID = '" & Fx.SampleID & "' AND Department = 'M' " & _
              "ELSE " & _
              "  INSERT INTO PrintValidLog " & _
              "  (SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime) " & _
              "  VALUES " & _
              "  ('" & Fx.SampleID & "', 'M', '" & Fx.Printed & "', '" & Fx.Valid & "', " & _
              "  '" & AddTicks(Fx.ValidatedBy) & "', '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
              "  '" & AddTicks(Fx.ValidatedBy) & "', '" & Format$(Fx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "')"
13050     Cnxn(0).Execute sql
          '---Junaid
13060     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

13070     intEL = Erl
13080     strES = Err.Description
13090     LogError "FaecesResults", "Save", intEL, strES, sql

End Sub





