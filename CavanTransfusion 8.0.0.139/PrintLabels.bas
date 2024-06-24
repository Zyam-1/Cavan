Attribute VB_Name = "PrintLabels"
Option Explicit

Public Sub PrintBatchLabels(ByVal LabNumber As String)

10      frmXMLabel.SampleID = LabNumber
20      frmXMLabel.Show 1

End Sub


Public Sub PrintLabel(ByRef RowNumbersToPrint() As Integer)

10        PrintLabelCavan RowNumbersToPrint
20        UpdatePrinted frmXMLabel.lLabNumber, "Label"

End Sub




Public Sub UpdatePrintedLabels(ByVal SampleID As String, _
                               ByVal UnitNumber As String, _
                               ByVal Expiry As String, _
                               ByVal BarCode As String)
      Dim sql As String

10    On Error GoTo UpdatePrintedLabels_Error

20    sql = "IF NOT EXISTS (SELECT * FROM PrintedLabels " & _
            "               WHERE SampleID = '" & SampleID & "' " & _
            "               AND UnitNumber = '" & UnitNumber & "' " & _
            "               AND BarCode = '" & BarCode & "' " & _
            "               AND Expiry = '" & Format$(Expiry, "dd/MMM/yyyy HH:mm") & "') " & _
            "  INSERT INTO PrintedLabels " & _
            "  ([SampleID], [UnitNumber], [Expiry], [PrintedBy], [BarCode]) " & _
            "  VALUES " & _
            "  ('" & SampleID & "', " & _
            "  '" & UnitNumber & "', " & _
            "  '" & Format$(Expiry, "dd/MMM/yyyy HH:mm") & "', " & _
            "  '" & AddTicks(UserName) & "', " & _
            "  '" & BarCode & "')"
30    CnxnBB(0).Execute sql

40    Exit Sub

UpdatePrintedLabels_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "PrintLabels", "UpdatePrintedLabels", intEL, strES, sql

End Sub


Public Function AlreadyPrintedLabel(ByVal SampleID As String, _
                                    ByVal UnitNumber As String, _
                                    ByVal Expiry As String, _
                                    ByVal BarCode As String) As Boolean
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo AlreadyPrintedLabel_Error

20    sql = "SELECT COUNT(*) Ret FROM PrintedLabels " & _
            "WHERE SampleID = '" & SampleID & "' " & _
            "AND UnitNumber = '" & UnitNumber & "' " & _
            "AND BarCode = '" & BarCode & "' " & _
            "AND Expiry = '" & Format$(Expiry, "dd/MMM/yyyy hh:mm") & "'"
30    Set tb = New Recordset
40    Set tb = CnxnBB(0).Execute(sql)
50    AlreadyPrintedLabel = tb!Ret > 0

60    Exit Function

AlreadyPrintedLabel_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "PrintLabels", "AlreadyPrintedLabel", intEL, strES, sql

End Function

Public Sub MarkAsNotPrinted(ByVal SampleID As String, _
                            ByVal UnitNumber As String, _
                            ByVal Expiry As String, _
                            ByVal BarCode As String)
      Dim sql As String

10    On Error GoTo MarkAsNotPrinted_Error

20    sql = "DELETE FROM PrintedLabels " & _
            "WHERE SampleID = '" & SampleID & "' " & _
            "AND UnitNumber = '" & UnitNumber & "' " & _
            "AND Expiry = '" & Format$(Expiry, "dd/MMM/yyyy hh:mm") & "' " & _
            "AND BarCode = '" & BarCode & "'"
30    CnxnBB(0).Execute sql

40    Exit Sub

MarkAsNotPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "PrintLabels", "MarkAsNotPrinted", intEL, strES, sql

End Sub



