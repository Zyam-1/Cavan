Attribute VB_Name = "modMicro"
Option Explicit

Public Function GetWBCValue(ByVal s As String) As String

      Dim v As Integer
      Dim RetVal As String

2520  v = Val(s)
2530  If v = 0 Then
2540    RetVal = "Nil"
2550  ElseIf v < 101 Then
2560    RetVal = str$(v)
2570  Else
2580    RetVal = ">100"
2590  End If

2600  GetWBCValue = RetVal

End Function

Public Function GetPlussesOrNil(ByVal s As String) As String

      Dim RetVal As String

2610  RetVal = ""

2620  If InStr(s, "-") > 0 Then
2630    RetVal = "Nil"
2640  Else
2650    If InStr(s, "++++") > 0 Then
2660      RetVal = "++++"
2670    ElseIf InStr(s, "+++") > 0 Then
2680      RetVal = "+++"
2690    ElseIf InStr(s, "++") > 0 Then
2700      RetVal = "++"
2710    ElseIf InStr(s, "+") > 0 Then
2720      RetVal = "+"
2730    End If
2740  End If

2750  GetPlussesOrNil = RetVal

End Function

Public Function GetPlusses(ByVal s As String) As String

      Dim RetVal As String

2760  RetVal = ""

2770  If InStr(s, "++++") > 0 Then
2780    RetVal = "++++"
2790  ElseIf InStr(s, "+++") > 0 Then
2800    RetVal = "+++"
2810  ElseIf InStr(s, "++") > 0 Then
2820    RetVal = "++"
2830  ElseIf InStr(s, "+") > 0 Then
2840    RetVal = "+"
2850  End If

2860  GetPlusses = RetVal

End Function

Public Function AntibioticCodeFor(ByVal inABName As String) As String

      Dim tb As Recordset
      Dim sql As String

2870  On Error GoTo AntibioticCodeFor_Error

2880  sql = "Select Code from Antibiotics where " & _
            "AntibioticName = '" & inABName & "'"
2890  Set tb = New Recordset
2900  RecOpenServer 0, tb, sql
2910  If tb.EOF Then
2920    AntibioticCodeFor = "???"
2930  Else
2940    If Trim$(tb!Code & "") <> "" Then
2950      AntibioticCodeFor = tb!Code
2960    Else
2970      AntibioticCodeFor = "???"
2980    End If
2990  End If

3000  Exit Function

AntibioticCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

3010  intEL = Erl
3020  strES = Err.Description
3030  LogError "modMicro", "AntibioticCodeFor", intEL, strES, sql

End Function

Public Function AntibioticNameFor(ByVal inABCode As String) As String

      Dim tb As Recordset
      Dim sql As String

3040  On Error GoTo AntibioticNameFor_Error

3050  sql = "Select AntibioticName from Antibiotics where " & _
            "Code = '" & inABCode & "'"
3060  Set tb = New Recordset
3070  RecOpenServer 0, tb, sql
3080  If tb.EOF Then
3090    AntibioticNameFor = inABCode
3100  Else
3110    AntibioticNameFor = tb!AntibioticName & ""
3120  End If

3130  Exit Function

AntibioticNameFor_Error:

      Dim strES As String
      Dim intEL As Integer

3140  intEL = Erl
3150  strES = Err.Description
3160  LogError "modMicro", "AntibioticNameFor", intEL, strES, sql

End Function



