Attribute VB_Name = "modMicro"
Option Explicit

Public Function GetWBCValue(ByVal S As String) As String

      Dim v As Integer
      Dim RetVal As String

10    v = Val(S)
20    If v = 0 Then
30      RetVal = "Nil"
40    ElseIf v < 101 Then
50      RetVal = Str$(v)
60    Else
70      RetVal = ">100"
80    End If

90    GetWBCValue = RetVal

End Function

Public Function GetPlussesOrNil(ByVal S As String) As String

      Dim RetVal As String

10    RetVal = ""

20    If InStr(S, "-") > 0 Then
30      RetVal = "Nil"
40    Else
50      If InStr(S, "++++") > 0 Then
60        RetVal = "++++"
70      ElseIf InStr(S, "+++") > 0 Then
80        RetVal = "+++"
90      ElseIf InStr(S, "++") > 0 Then
100       RetVal = "++"
110     ElseIf InStr(S, "+") > 0 Then
120       RetVal = "+"
130     End If
140   End If

150   GetPlussesOrNil = RetVal

End Function

Public Function GetPlusses(ByVal S As String) As String

      Dim RetVal As String

10    RetVal = ""

20    If InStr(S, "++++") > 0 Then
30      RetVal = "++++"
40    ElseIf InStr(S, "+++") > 0 Then
50      RetVal = "+++"
60    ElseIf InStr(S, "++") > 0 Then
70      RetVal = "++"
80    ElseIf InStr(S, "+") > 0 Then
90      RetVal = "+"
100   End If

110   GetPlusses = RetVal

End Function

Public Function AntibioticCodeFor(ByVal inABName As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo AntibioticCodeFor_Error

20    sql = "Select Code from Antibiotics where " & _
            "AntibioticName = '" & inABName & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      AntibioticCodeFor = "???"
70    Else
80      If Trim$(tb!Code & "") <> "" Then
90        AntibioticCodeFor = tb!Code
100     Else
110       AntibioticCodeFor = "???"
120     End If
130   End If

140   Exit Function

AntibioticCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "modMicro", "AntibioticCodeFor", intEL, strES, sql

End Function

Public Function AntibioticNameFor(ByVal inABCode As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo AntibioticNameFor_Error

20    sql = "Select AntibioticName from Antibiotics where " & _
            "Code = '" & inABCode & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      AntibioticNameFor = inABCode
70    Else
80      AntibioticNameFor = tb!AntibioticName & ""
90    End If

100   Exit Function

AntibioticNameFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modMicro", "AntibioticNameFor", intEL, strES, sql

End Function



