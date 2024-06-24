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

190       On Error Resume Next

200       UID = AddTicks(UserName)

210       SQLStatement = AddTicks(SQLStatement)

220       ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "[MSSQL]")
230       ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver]", "[SQL]")
240       ErrorDescription = AddTicks(ErrorDescription)

250       Vers = App.Major & "-" & App.Minor & "-" & App.Revision

260       MyMachineName = vbGetComputerName()

270       sql = "IF NOT EXISTS " & _
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
280       sql = sql & "  UPDATE ErrorLog " & _
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

290       Cnxn(0).Execute sql

End Sub


