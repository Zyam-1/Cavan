Attribute VB_Name = "modTechnicians"
Option Explicit

Public Function TechnicianMemberOf(ByVal TechnicianNameOrCode As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo TechnicianMemberOf_Error

20    sql = "Select MemberOf from Users where " & _
            "Name = '" & TechnicianNameOrCode & "' " & _
            "or Code = '" & TechnicianNameOrCode & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60      TechnicianMemberOf = tb!MemberOf & ""
70    End If

80    Exit Function

TechnicianMemberOf_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "modTechnicians", "TechnicianMemberOf", intEL, strES, sql


End Function

Public Function TechnicianNameForCode(ByVal code As String) As String

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo TechnicianNameForCode_Error

20    sql = "Select * from Users where " & _
            "Code = '" & AddTicks(code) & "'"
30    RecOpenServer 0, tb, sql
40    If Not tb.EOF Then
50      TechnicianNameForCode = tb!Name & ""
60    Else
70      TechnicianNameForCode = "???"
80    End If

90    Exit Function

TechnicianNameForCode_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modTechnicians", "TechnicianNameForCode", intEL, strES, sql


End Function

Public Function TechnicianPasswordForName(ByVal UserName As String) As String

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo TechnicianPasswordForName_Error

20    sql = "Select Password from Users where " & _
            "Name = '" & AddTicks(UserName) & "'"
30    RecOpenServer 0, tb, sql
40    If Not tb.EOF Then
50      TechnicianPasswordForName = tb!Password & ""
60    Else
70      TechnicianPasswordForName = "???"
80    End If

90    Exit Function

TechnicianPasswordForName_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modTechnicians", "TechnicianPasswordForName", intEL, strES, sql


End Function

Public Function TechnicianCodeFor(ByVal NameOrCode As String) As String
      'Returns UserCode given either UserName or UserCode

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo TechnicianCodeFor_Error

20    TechnicianCodeFor = "???"

30    NameOrCode = AddTicks(UCase$(Trim$(NameOrCode)))

40    sql = "Select Code from Users where " & _
            "Code = '" & NameOrCode & "' " & _
            "or Name = '" & NameOrCode & "'"
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then
70      TechnicianCodeFor = Trim$(tb!code & "")
80    End If

90    Exit Function

TechnicianCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modTechnicians", "TechnicianCodeFor", intEL, strES, sql


End Function


Public Function UserPasswordForCode(ByVal strCode As String) As String
      'Returns UserPassword given UserCode

      Dim tb As New Recordset
      Dim sql As String



   On Error GoTo UserPasswordForCode_Error

20    UserPasswordForCode = ""

30    strCode = AddTicks(UCase$(Trim$(strCode)))

40    sql = "Select Password from Users where " & _
            "Code = '" & strCode & "' "
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then
70      UserPasswordForCode = Trim$(tb!Password & "")
80    End If


   Exit Function

UserPasswordForCode_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modTechnicians", "UserPasswordForCode", intEL, strES, sql

End Function

