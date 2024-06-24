Attribute VB_Name = "modPrintable"
Option Explicit

Public Function IsPrintable(ByVal SampleID As String, _
                            ByVal RunDate As String, _
                            ByVal Department As String) _
                            As Boolean

      Dim RetVal As Boolean
      Dim V7Date As String

10    On Error GoTo IsPrintable_Error

20    RetVal = False

30    If WardEnqForcedPrinter <> "" Then
40      If UserCanPrint Then
50        V7Date = GetOptionSetting("WardEnqV7Date", "01/May/2011", "")
60        If DateDiff("d", V7Date, RunDate) > 0 Then
70          If ReportIsPresent(SampleID, Department) Then
80            RetVal = True
90          End If
100       Else
110         RetVal = AllValid(SampleID, Department)
120       End If
130     End If
140   End If

150   IsPrintable = RetVal

160   Exit Function

IsPrintable_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "modPrintable", "IsPrintable", intEL, strES

End Function


Public Function AllValid(ByVal SampleID As String, _
                         ByVal Department As String) _
                         As Boolean
    
      Dim sql As String
      Dim Dept As String
      Dim tb As Recordset

10    On Error GoTo AllValid_Error

20    Select Case UCase$(Department)
        Case "BIOCHEMISTRY": Dept = "Bio"
30      Case "COAGULATION": Dept = "Coag"
40      Case "HAEMATOLOGY": Dept = "Haem"
50    End Select

60    sql = "IF EXISTS(SELECT * FROM " & Dept & "Results " & _
            "          WHERE SampleID = '" & SampleID & "') " & _
            "  SELECT COUNT(*) Tot " & _
            "  FROM " & Dept & "Results " & _
            "  WHERE SampleID = '" & SampleID & "' " & _
            "  AND COALESCE(Valid, 0) = 0 " & _
            "ELSE " & _
            "  SELECT -1 Tot"
70    Set tb = New Recordset
80    Set tb = Cnxn(0).Execute(sql)
90    AllValid = tb!Tot = 0

100   Exit Function

AllValid_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modPrintable", "AllValid", intEL, strES, sql

End Function

Public Function ReportIsPresent(ByVal SampleID As String, ByVal Department As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo ReportIsPresent_Error

20    sql = "SELECT COUNT(*) Tot FROM Reports " & _
            "WHERE SampleID = '" & SampleID & "' " & _
            "AND Dept = '" & Department & "' " & _
            "AND COALESCE(Hidden, 0) = 0"
30    Set tb = New Recordset
40    Set tb = Cnxn(0).Execute(sql)
50    ReportIsPresent = tb!Tot > 0

60    Exit Function

ReportIsPresent_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modPrintable", "ReportIsPresent", intEL, strES, sql

End Function


