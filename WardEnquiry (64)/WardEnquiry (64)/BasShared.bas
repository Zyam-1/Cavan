Attribute VB_Name = "BasShared"
Public Sub AddActivity(ByVal SampleID As String, ByVal ActionType As String, ByVal Action As String, Optional ByVal SubmissionID As String, Optional ByVal PatientID As String, _
                       Optional ByVal Reason As String, Optional ByVal Notes As String)
    Dim tb As New Recordset
    Dim sql As String
10    On Error GoTo AddActivity_Error

20    sql = "Select * from ActivityLog"
30    Set tb = New Recordset

40    RecOpenServer 0, tb, sql

50    tb.AddNew
60    tb!SampleID = SampleID
70    If SampleID <> "" And (SubmissionID = "" Or PatientID = "") Then
80      Call getChartFromSampleID(SampleID, PatientID)
90    End If
100   tb!ActionType = ActionType
110   tb!Action = Action
'120   tb!SubmissionID = IIf(IsMissing(SubmissionID), "", SubmissionID)
130   tb!PatientID = IIf(IsMissing(PatientID), "", PatientID)
140   tb!Reason = IIf(IsMissing(Reason), "", Reason)
150   tb!Notes = IIf(IsMissing(Notes), "", Notes)
160   tb!UserName = UserName
170   tb!DateTimeOfRecord = Format(Now, "dd/mmm/yyyy hh:mm:ss")
180   tb!MachineName = vbGetComputerName
190   tb!ApplicationName = "NetAcquire LIS"
200   tb!ApplicationVersion = App.Major & "-" & App.Minor & "-" & App.Revision
210   tb!Createdby = UserName
220   tb.Update
230   Exit Sub

AddActivity_Error:
    Dim strES As String
    Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "basShared", "AddActivity", intEL, strES, sql

End Sub

Public Sub getChartFromSampleID(ByVal strSID As String, ByRef strChart As String)
Dim sql As String
Dim tb As Recordset

10    On Error GoTo getChartFromSampleID_Error

20    sql = "SELECT  chart FROM Demographics WHERE SampleID  = '" & strSID & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not (tb.EOF) Then
70      strChart = tb!Chart & ""
80    End If

90    Exit Sub

getChartFromSampleID_Error:

 Dim strES As String
 Dim intEL As Integer

100    intEL = Erl
110    strES = Err.Description
120    LogError "basShared", "getChartFromSampleID", intEL, strES, sql

End Sub


'+++ Junaid 15-02-2024
Public Function ConvertNull(Data As Variant, Default As Variant) As Variant
    On Error GoTo ERROR_ConvertNull
    If IsNull(Data) = True Then
        ConvertNull = Default
    Else
        ConvertNull = Data
    End If
    Exit Function
ERROR_ConvertNull:
    LogError "Shared", "ConvertNull", Erl, Err.Description
End Function
'--- Junaid
