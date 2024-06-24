Attribute VB_Name = "modUpDatePrinted"
Option Explicit


Public Sub UpdatePrinted(ByVal SampleID As String, _
                         ByVal LabelOrForm As String)
       
      Dim sql As String

10    On Error GoTo UpdatePrinted_Error

20    sql = "IF EXISTS (SELECT * FROM PatientDetails " & _
            "           WHERE LabNumber = '" & SampleID & "' ) " & _
            "  UPDATE PatientDetails " & _
            "  SET " & LabelOrForm & "PrintTime = getdate(), " & _
            "      " & LabelOrForm & "PrintedBy = '" & UserName & "', " & _
            "      Valid = 1 " & _
            "  WHERE LabNumber = '" & SampleID & "'"
30    CnxnBB(0).Execute sql

40    Exit Sub

UpdatePrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "modUpDatePrinted", "UpdatePrinted", intEL, strES, sql
       
End Sub



