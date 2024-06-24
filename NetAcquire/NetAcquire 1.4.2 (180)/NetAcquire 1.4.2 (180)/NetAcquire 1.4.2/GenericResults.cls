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
        
14620     Count = mcolGenericResults.Count

End Function




Public Sub Delete(ByVal SampleID As String, ByVal TestName As String)
        
          Dim sql As String

14630     On Error GoTo Delete_Error

14640     sql = "DELETE FROM GenericResults WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND TestName = '" & TestName & "'"
14650     Cnxn(0).Execute sql

14660     mcolGenericResults.Remove TestName

14670     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

14680     intEL = Erl
14690     strES = Err.Description
14700     LogError "GenericResults", "Delete", intEL, strES, sql

End Sub



Public Property Get GenericResults() As GenericResults
        
14710     Set GenericResults = mGenericResults

End Property


Public Function Item(ByVal TestName As String) As GenericResult
Attribute Item.VB_UserMemId = 0
        
14720     On Error Resume Next

14730     Set Item = mcolGenericResults.Item(TestName)

End Function




Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Gx As GenericResult

14740     On Error GoTo Load_Error

14750     sql = "SELECT G.SampleID, G.TestName, G.Result, COALESCE(G.HealthLink, 0) Healthlink, " & _
              "G.UserName, G.DateTimeOfRecord, COALESCE(P.Valid, 0) Valid, COALESCE(P.Printed, 0) Printed, " & _
              "P.ValidatedBy, P.ValidatedDateTime, P.PrintedBy, P.PrintedDateTime " & _
              "FROM GenericResults G LEFT JOIN PrintValidLog P " & _
              "ON G.SampleID = P.SampleID " & _
              "WHERE G.SampleID = '" & SampleID & "'"
14760     Set tb = New Recordset
14770     RecOpenClient 0, tb, sql
14780     Do While Not tb.EOF
14790         Set Gx = New GenericResult
14800         With Gx
14810             .SampleID = tb!SampleID & ""
14820             .TestName = tb!TestName & ""
14830             .Result = tb!Result & ""
14840             .HealthLink = tb!HealthLink
14850             .Valid = tb!Valid
14860             .Printed = tb!Printed
14870             .ValidatedBy = tb!ValidatedBy & ""
14880             .ValidatedDateTime = tb!ValidatedDateTime & ""
14890             .PrintedBy = tb!PrintedBy & ""
14900             .PrintedDateTime = tb!PrintedDateTime & ""
14910             .UserName = tb!UserName & ""
14920             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
14930             mcolGenericResults.Add Gx, .TestName
14940         End With
14950         tb.MoveNext
14960     Loop

14970     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

14980     intEL = Erl
14990     strES = Err.Description
15000     LogError "GenericResults", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
        
15010     Set NewEnum = mcolGenericResults.[_NewEnum]

End Function



Public Sub Save(ByVal Gx As GenericResult)

          Dim sql As String

15020     On Error GoTo Save_Error

15030     If Trim$(Gx.Result) = "" Then
              '30      Sql = "Delete from GenericResults where " & _
              '              "SampleID = '" & Gx.SampleID & "' " & _
              '              "and TestName = '" & Gx.TestName & "'"
              '        Exit Sub
15040     Else
15050         sql = "IF EXISTS (SELECT * FROM GenericResults WHERE " & _
                  "           SampleID = '" & Gx.SampleID & "' " & _
                  "           AND TestName = '" & Gx.TestName & "') " & _
                  "  UPDATE GenericResults " & _
                  "  SET Result = '" & Gx.Result & "', " & _
                  "  UserName = '" & AddTicks(Gx.UserName) & "', " & _
                  "  HealthLink = '" & Gx.HealthLink & "', " & _
                  "  Valid = '" & Gx.Valid & "', " & _
                  "  ValidatedBy = '" & AddTicks(Gx.ValidatedBy) & "', " & _
                  "  ValidatedDateTime = " & IIf(IsDate(Gx.ValidatedDateTime), Format$(Gx.ValidatedDateTime, "'dd/MMM/yyyy HH:nn:ss'"), "Null") & " " & _
                  "  WHERE SampleID = '" & Gx.SampleID & "' " & _
                  "  AND TestName = '" & Gx.TestName & "' "
15060         sql = sql & "ELSE " & _
                  "  INSERT INTO GenericResults (SampleID, TestName, Result, UserName, HealthLink, Valid, ValidatedBy, ValidatedDateTime ) " & _
                  "  VALUES " & _
                  "  ('" & Gx.SampleID & "', " & _
                  "  '" & Gx.TestName & "', " & _
                  "  '" & Gx.Result & "', " & _
                  "  '" & AddTicks(Gx.UserName) & "', " & _
                  "  '" & Gx.HealthLink & "', " & _
                  "  '" & Gx.Valid & "', " & _
                  "  '" & AddTicks(Gx.ValidatedBy) & "', " & _
                  "  " & IIf(IsDate(Gx.ValidatedDateTime), Format$(Gx.ValidatedDateTime, "'dd/MMM/yyyy HH:nn:ss'"), "Null") & ")"
15070         Cnxn(0).Execute sql
15080     End If

          '+++Junaid 17-01-2024
          '90    Sql = "IF EXISTS (SELECT * FROM PrintValidLog WHERE " & _
          '            "           SampleID = '" & Gx.SampleID & "' " & _
          '            "           AND Department = 'M') " & _
          '            "  UPDATE PrintValidLog " & _
          '            "  SET Valid = '" & Gx.Valid & "', " & _
          '            "  Printed = '" & Gx.Printed & "', " & _
          '            "  ValidatedBy = '" & AddTicks(Gx.ValidatedBy) & "', " & _
          '            "  PrintedBy = '" & AddTicks(Gx.PrintedBy) & "', " & _
          '            "  ValidatedDateTime = '" & Format$(Gx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
          '            "  PrintedDateTime = '" & Format$(Gx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "' " & _
          '            "  WHERE SampleID = '" & Gx.SampleID & "' AND Department = 'M' " & _
          '            "ELSE " & _
          '            "  INSERT INTO PrintValidLog " & _
          '            "  (SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime) " & _
          '            "  VALUES " & _
          '            "  ('" & Gx.SampleID & "', 'M', '" & Gx.Printed & "', '" & Gx.Valid & "', " & _
          '            "  '" & AddTicks(Gx.PrintedBy) & "', '" & Format$(Gx.PrintedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
          '            "  '" & AddTicks(Gx.ValidatedBy) & "', '" & Format$(Gx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "')"
          '100   Cnxn(0).Execute Sql

15090     sql = "IF EXISTS (SELECT * FROM PrintValidLog WHERE " & _
              "           SampleID = '" & Gx.SampleID & "' " & _
              "           AND Department = 'M') " & _
              "  UPDATE PrintValidLog " & _
              "  SET Valid = '" & Gx.Valid & "', " & _
              "  Printed = '" & Gx.Printed & "', " & _
              "  ValidatedBy = '" & AddTicks(Gx.ValidatedBy) & "', " & _
              "  PrintedBy = '" & AddTicks(Gx.ValidatedBy) & "', " & _
              "  ValidatedDateTime = '" & Format$(Gx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
              "  PrintedDateTime = '" & Format$(Gx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
              "  WHERE SampleID = '" & Gx.SampleID & "' AND Department = 'M' " & _
              "ELSE " & _
              "  INSERT INTO PrintValidLog " & _
              "  (SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime) " & _
              "  VALUES " & _
              "  ('" & Gx.SampleID & "', 'M', '" & Gx.Printed & "', '" & Gx.Valid & "', " & _
              "  '" & AddTicks(Gx.ValidatedBy) & "', '" & Format$(Gx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
              "  '" & AddTicks(Gx.ValidatedBy) & "', '" & Format$(Gx.ValidatedDateTime, "dd/MMM/yyyy HH:nn:ss") & "')"
15100     Cnxn(0).Execute sql
          '---Junaid
15110     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

15120     intEL = Erl
15130     strES = Err.Description
15140     LogError "GenericResults", "Save", intEL, strES, sql

End Sub



