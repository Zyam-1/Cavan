Attribute VB_Name = "modUpdateHaem"
Option Explicit

Public Sub CheckForHaemUpdate()

26090 EnsureColsExist

End Sub

Private Sub EnsureColsExist()

      Dim sql As String

26100 On Error GoTo EnsureColsExist_Error

26110 EnsureColumnExists "HaemTestDefinitions", "PrintPriority", "int NOT NULL DEFAULT 1"
26120 EnsureColumnExists "HaemTestDefinitions", "RedWhitePltOther", "char(1) NOT NULL DEFAULT 'O'"
26130 EnsureColumnExists "HaemTestDefinitions", "LongName", "nvarchar(50)"
26140 If EnsureColumnExists("HaemTestDefinitions", "ShortName", "nvarchar(50)") Then
26150   sql = "UPDATE HaemTestDefinitions " & _
              "SET LongName = AnalyteName, ShortName = AnalyteName"
26160   Cnxn(0).Execute sql
26170 End If

26180 Exit Sub

EnsureColsExist_Error:

      Dim strES As String
      Dim intEL As Integer

26190 intEL = Erl
26200 strES = Err.Description
26210 LogError "modUpdateHaem", "EnsureColsExist", intEL, strES, sql


End Sub


Public Sub UpdateHaem()

      Dim n As Integer
      Dim Param As String
      Dim Unit As String

26220 For n = 1 To 8
26230   Param = Choose(n, "WBC", "RBC", "Hgb", "Hct", "MCV", "MCH", "MCHC", "Plt")
26240   Unit = Choose(n, "x10^3/ml", "x10^12/ml", "g/dl", "%", "fl", "pg", "g/dl", "x10^3/ml")
26250   UpdateParameter Param, Unit
26260 Next

26270 For n = 1 To 5
26280   Param = Choose(n, "LymP", "MonoP", "NeutP", "EosP", "BasP")
26290   Unit = "%"
26300   UpdateParameter Param, Unit
26310 Next

26320 For n = 1 To 5
26330   Param = Choose(n, "LymA", "MonoA", "NeutA", "EosA", "BasA")
26340   Unit = "x10^3/ml"
26350   UpdateParameter Param, Unit
26360 Next

26370 For n = 1 To 7
26380   Param = Choose(n, "CD3A", "CD4A", "CD8A", "CD3P", "CD4P", "CD8P", "CD48")
26390   Unit = ""
26400   UpdateParameter Param, Unit
26410 Next

26420 For n = 1 To 3
26430   Param = Choose(n, "Monospot", "Malaria", "Sickledex")
26440   Unit = ""
26450   UpdateParameter Param, Unit
26460 Next

26470 For n = 1 To 15
26480   Param = Choose(n, "MI", "AN", "CA", "VA", "HO", "HE", "LS", "AT", "BL", "PP", "NL", "MN", "WP", "CH", "WB")
26490   Unit = ""
26500   UpdateParameter Param, Unit
26510 Next

26520 For n = 1 To 2
26530   Param = Choose(n, "RDWCV", "RDWSD")
26540   Unit = "%"
26550   UpdateParameter Param, Unit
26560 Next

26570 For n = 1 To 10
26580   Param = Choose(n, "PDW", "MPV", "PLCR", "Retics", "ESR", "Pct", "WIC", "WOC", "RetA", "RetP")
26590   Unit = Choose(n, "", "fl", "%", "", "", "%", "", "", "", "%")
26600   UpdateParameter Param, Unit
26610 Next

26620 For n = 1 To 8
26630   Param = Choose(n, "NRBCA", "NRBCP", "RA", "IRF", "HDW", "LUCP", "LUCA", "LI")
26640   Unit = ""
26650   UpdateParameter Param, Unit
26660 Next

26670 For n = 1 To 11
26680   Param = Choose(n, "MPXI", "tASOT", "tRA", "HYP", "RBCf", "RBCg", "MPO", "IG", "LPLT", "PCLM", "WVF")
26690   Unit = ""
26700   UpdateParameter Param, Unit
26710 Next

End Sub

Private Sub UpdateParameter(ByVal FieldName As String, ByVal Units As String)

      Dim sql As String

26720 On Error GoTo UpdateParameter_Error

26730 sql = "INSERT INTO [Haem50Results] " & _
            "(  [SampleId], " & _
            "   [Code], " & _
            "   [Result], " & _
            "   [Flags], [Units], " & _
            "   [Valid], [Printed], [Faxed], " & _
            "   [RunDateTime], [UserName], [SampleType], [Analyser], " & _
            "   [HealthLinkSent], [DateTimeOfRecord] ) " & _
            "SELECT [SampleId], " & _
            "'" & FieldName & "', " & _
            " " & FieldName & ", " & _
            " 0, '" & Units & "', " & _
            "COALESCE([Valid], 0), COALESCE([Printed], 0), COALESCE([Faxed], 0), " & _
            "[RunDateTime], [Operator], 'WholeBlood', [Analyser], " & _
            "COALESCE([HealthLink], 0), [RunDateTime] " & _
            "FROM HaemResults " & _
            "WHERE COALESCE(" & FieldName & ", '') <> ''"
26740 Cnxn(0).Execute sql

26750 Exit Sub

UpdateParameter_Error:

      Dim strES As String
      Dim intEL As Integer

26760 intEL = Erl
26770 strES = Err.Description
26780 LogError "modUpdateHaem", "UpdateParameter", intEL, strES, sql


End Sub



