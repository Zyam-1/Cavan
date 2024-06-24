Attribute VB_Name = "modErrorHandler"
Option Explicit

Public Sub LogError(ByVal ModuleName As String, _
                    ByVal ProcedureName As String, _
                    ByVal ErrorLineNumber As Integer, _
                    ByVal ErrorDescription As String, _
                    Optional ByVal SQLStatement As String, _
                    Optional ByVal EventDesc As String)

          Dim sql As String
          Dim MyMachineName As String
          Dim Vers As String
          Dim UID As String

10        On Error Resume Next

20        UID = AddTicks(Username)

30        SQLStatement = AddTicks(SQLStatement)

40        ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "[MSSQL]")
50        ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver]", "[SQL]")
60        ErrorDescription = AddTicks(ErrorDescription)

70        Vers = App.Major & "-" & App.Minor & "-" & App.Revision

80        MyMachineName = vbGetComputerName()

90        sql = "IF NOT EXISTS " & _
                "    (SELECT * FROM ErrorLog WHERE " & _
                "     ModuleName = '" & ModuleName & "' " & _
                "     AND ProcedureName = '" & ProcedureName & "' " & _
                "     AND ErrorLineNumber = '" & ErrorLineNumber & "' " & _
                "     AND AppName = '" & App.EXEName & "' " & _
                "     AND AppVersion = '" & Vers & "' ) " & _
                "  INSERT INTO ErrorLog (" & _
                "    ModuleName, ProcedureName, ErrorLineNumber, SQLStatement, " & _
                "    ErrorDescription, UserName, MachineName, Eventdesc, AppName, AppVersion, EventCounter, eMailed) " & _
                "  VALUES  ('" & ModuleName & "', " & _
                "           '" & ProcedureName & "', " & _
                "           '" & ErrorLineNumber & "', " & _
                "           '" & SQLStatement & "', " & _
                "           '" & ErrorDescription & "', " & _
                "           '" & UID & "', " & _
                "           '" & MyMachineName & "', " & _
                "           '" & AddTicks(EventDesc) & "', " & _
                "           '" & App.EXEName & "', " & _
                "           '" & Vers & "', " & _
                "           '1', '0') " & _
      "ELSE "
100       sql = sql & "  UPDATE ErrorLog " & _
                "  SET SQLStatement = '" & SQLStatement & "', " & _
                "  ErrorDescription = '" & ErrorDescription & "', " & _
                "  MachineName = '" & MyMachineName & "', " & _
                "  DateTime = getdate(), " & _
                "  UserName = '" & UID & "', " & _
                "  EventCounter = COALESCE(EventCounter, 0) + 1 " & _
                "  WHERE ModuleName = '" & ModuleName & "' " & _
                "  AND ProcedureName = '" & ProcedureName & "' " & _
                "  AND ErrorLineNumber = '" & ErrorLineNumber & "' " & _
                "  AND AppName = '" & App.EXEName & "' " & _
                "  AND AppVersion = '" & Vers & "'"

110       Cnxn(0).Execute sql

End Sub


