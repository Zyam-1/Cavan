Attribute VB_Name = "modOptionSettings"
Option Explicit

Public Sub SaveOptionSetting(ByVal Description As String, _
                             ByVal Contents As String)
   
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveOptionSetting_Error

20    sql = "SELECT * FROM Options WHERE " & _
            "Description = '" & Description & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      tb.AddNew
70    End If
80    tb!Description = Description
90    tb!Contents = Contents
100   tb.Update

110   Exit Sub

SaveOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modOptionSettings", "SaveOptionSetting", intEL, strES, sql

End Sub


Public Function GetOptionSetting(ByVal Description As String, _
                                 ByVal Default As String) As String
   
      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo GetOptionSetting_Error

20    sql = "SELECT Contents FROM Options WHERE " & _
            "Description = '" & Description & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      RetVal = Default
70    ElseIf Trim$(tb!Contents & "") = "" Then
80      RetVal = Default
90    Else
100     RetVal = tb!Contents
110   End If

120   GetOptionSetting = RetVal

130   Exit Function

GetOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "modOptionSettings", "GetOptionSetting", intEL, strES, sql

End Function


Public Sub SaveUserOptionSetting(ByVal Description As String, _
                             ByVal Contents As String, _
                             ByVal UserName As String)
   
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveUserOptionSetting_Error

20    UserName = AddTicks(UserName)
30    Contents = AddTicks(Contents)

40    sql = "IF EXISTS (SELECT * FROM Options WHERE " & _
            "           Description = '" & Description & "' AND " & _
            "           Username = '" & UserName & "') " & _
            "  UPDATE Options SET Contents = '" & Contents & "' " & _
            "  WHERE Description = '" & Description & "' AND " & _
            "  Username = '" & UserName & "' " & _
            "ELSE " & _
            "  INSERT INTO Options " & _
            "  (Description, Contents, UserName) VALUES ( " & _
            "  '" & Description & "', " & _
            "  '" & Contents & "', " & _
            "  '" & UserName & "')"
50    Cnxn(0).Execute sql

60    Exit Sub

SaveUserOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modOptionSettings", "SaveUserOptionSetting", intEL, strES, sql

End Sub


Public Function GetUserOptionSetting(ByVal Description As String, _
                                 ByVal Default As String, _
                                 ByVal UserName As String) As String
   
      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo GetUserOptionSetting_Error

20    sql = "SELECT Contents FROM Options WHERE " & _
            "Description = '" & Description & "' " & _
            "AND COALESCE(Username, '') = '" & AddTicks(UserName) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      RetVal = Default
70    ElseIf Trim$(tb!Contents & "") = "" Then
80      RetVal = Default
90    Else
100     RetVal = tb!Contents
110   End If

120   GetUserOptionSetting = RetVal

130   Exit Function

GetUserOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "modOptionSettings", "GetUserOptionSetting", intEL, strES, sql

End Function

