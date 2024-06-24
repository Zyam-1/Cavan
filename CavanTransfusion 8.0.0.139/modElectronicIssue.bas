Attribute VB_Name = "modElectronicIssue"
Option Explicit


Public Function GetEITime(ByVal SampleDate As String, ByVal SampleTime As String, ByVal Chart As String) As String

    Dim SampleDT As String
    Dim ExpiryDate As String
    Dim retval As String

10    On Error GoTo GetEITime_Error

20    If IsDate(SampleDate) Then
30      If IsDate(SampleTime) Then
40          SampleDT = Format$(SampleDate & " " & SampleTime, "dd/MMM/yyyy HH:nn")
50          ExpiryDate = DateAdd("h", GetOptionSetting("TransfusionHoldFor", "72"), SampleDT)
60          If DateDiff("n", ExpiryDate, Now) > 0 Then
70              retval = "Expired"
80          Else
90              retval = ExpiryDate
100         End If
110     Else
120         retval = "Invalid Sample Time"
130     End If
140   Else
150     retval = "Invalid Sample Date"
160   End If

170   GetEITime = retval

180   Exit Function

GetEITime_Error:

    Dim strES As String
    Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modElectronicIssue", "GetEITime", intEL, strES

End Function


Private Function GetFirstTxAfterSampleDate(ByVal Chart As String, ByVal SampleDT As String) As String

      Dim sql As String
      Dim tb As New Recordset
      Dim retval As String

10    On Error GoTo GetFirstTxAfterSampleDate_Error

20    retval = ""

30    sql = "SELECT TOP 1 " & _
            "CASE Event WHEN 'S' THEN EventStart WHEN 'Y' THEN DateTime END T " & _
            "FROM Product " & _
            "WHERE PatID = '" & Chart & "' " & _
            "AND (Event = 'S' OR Event = 'Y') " & _
            "AND DATEDIFF(n, CASE [Event] WHEN 'S' THEN EventStart " & _
            "                             WHEN 'Y' THEN [datetime] END, " & _
            "                   '" & Format$(SampleDT, "dd/MMM/yyyy HH:nn") & "') < 0 " & _
            "ORDER BY T ASC"

40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      retval = tb!t & ""
70    End If

80    GetFirstTxAfterSampleDate = retval

90    Exit Function

GetFirstTxAfterSampleDate_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modElectronicIssue", "GetFirstTxAfterSampleDate", intEL, strES, sql


End Function


Private Function GetInitialExpiry(ByVal Chart As String, ByVal SampleDT As String) As Date

      Dim sql As String
      Dim tb As New Recordset
      Dim retval As Date

10    On Error GoTo GetInitialExpiry_Error

20    sql = "IF EXISTS ( SELECT * FROM Product " & _
            "            WHERE PatID = '" & Chart & "' " & _
            "            AND ([Event] = 'S' OR [Event] = 'Y') " & _
            "            AND DATEDIFF(n, CASE [Event] WHEN 'S' THEN EventStart WHEN 'Y' THEN [datetime] END , '" & SampleDT & "') < 4320 ) " & _
            "  SELECT DATEADD(n, 4320, '" & SampleDT & "') ExDate " & _
            "ELSE " & _
            "  SELECT DATEADD(n, 10080, '" & SampleDT & "') ExDate"

30    Set tb = CnxnBB(0).Execute(sql)

40    retval = Format$(tb!ExDate, "dd/MMM/yyyy HH:nn")

50    GetInitialExpiry = retval

60    Exit Function

GetInitialExpiry_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modElectronicIssue", "GetInitialExpiry", intEL, strES, sql

End Function


Private Function HasLaterTxEvent(ByVal Chart As String, ByVal SampleDT As String) As Boolean

      Dim retval As Boolean
      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo HasLaterTxEvent_Error

20    sql = "SELECT COUNT(*) Tot FROM Product " & _
            "WHERE ([Event] = 'S' OR [Event] = 'Y') " & _
            "AND DATEDIFF(n, CASE [Event] WHEN 'S' THEN EventStart " & _
            "                             WHEN 'Y' THEN [datetime] END , '" & SampleDT & "') < 0 " & _
            "AND PatID = '" & Chart & "'"

30    RecOpenServerBB 0, tb, sql
40    retval = tb!Tot > 0

50    HasLaterTxEvent = retval

60    Exit Function

HasLaterTxEvent_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modElectronicIssue", "HasLaterTxEvent", intEL, strES, sql

End Function
