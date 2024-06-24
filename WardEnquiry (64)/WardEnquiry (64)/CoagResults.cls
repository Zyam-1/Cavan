VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoagResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolCoagResults As New Collection
Private mCoagResults As New CoagResults

Public Function Add(CR As CoagResult) As CoagResult

10    mcolCoagResults.Add CR

20    Set Add = CR

End Function

Public Property Get CoagResults() As CoagResults

10    Set CoagResults = mCoagResults

End Property

Public Sub Clear()

      Dim n As Integer

10    For n = 1 To mcolCoagResults.Count
20        mcolCoagResults.Remove 1
30    Next

End Sub

Public Function Count() As Long

10    Count = mcolCoagResults.Count

End Function


Public Sub Delete(ByVal SampleID As String, _
                  ByVal Code As String, ByVal Units As String)

      Dim sql As String

10    On Error Resume Next

20    sql = "Delete from CoagResults where " & _
            "SampleID = '" & SampleID & "' " & _
            "and Code = '" & Code & "' and units = '" & Units & "'"

30    Cnxn(0).Execute sql

End Sub

Public Function Item(ByVal x As Long) _
       As CoagResult
Attribute Item.VB_UserMemId = 0

10    On Error Resume Next

20    Set Item = mcolCoagResults(x)

End Function


Public Function Load(ByVal SampleID As String, _
                     ByVal intV As Integer, _
                     ByVal intP As Integer, _
                     ByVal strResultsOrRepeats As String, _
                     ByVal intCN As Integer) _
                     As CoagResults

          Dim CRs As New CoagResults
          Dim CR As CoagResult
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo Load_Error

20        sql = "Select R.SampleID, R.Result,  " & _
                "R.RunDate, R.RunTime, R.Valid, R.Analyser, " & _
                "R.Printed, D.Code, D.TestName, " & _
                "D.DP, D.MaleLow as Low, D.MaleHigh as High, " & _
                "D.Units, D.PrintPriority, D.Printable, D.PrintRefRange " & _
                 ",ISNULL(R.SignOff,0) as SignOff , ISNULL(R.SignOffBy,'') as SignOffBy, R.SignOffDateTime as SignOffDateTime " & _
                "from Coag" & strResultsOrRepeats & " as R, CoagTestDefinitions as D where " & _
                "SampleID = '" & SampleID & "' " & _
                "and R.Code = D.Code "

30        If intV = gVALID And intP = gNOTPRINTED Then
40            sql = sql & "and Valid = 1 and Printed = 0 "
50        ElseIf intV = gVALID And intP = gDONTCARE Then
60            sql = sql & "and valid = 1 "
70        ElseIf intV = gNOTVALID And intP = gDONTCARE Then
80            sql = sql & "and valid = 0 "
90        End If

100       sql = sql & " order by PrintPriority"
110       Set tb = New Recordset
120       RecOpenClient intCN, tb, sql

130       Do While Not tb.EOF
140           Set CR = New CoagResult
150           With CR
160               .SampleID = tb!SampleID & ""
170               .Code = tb!Code & ""
180               .Result = tb!Result & ""
                  '.OperatorCode = tb!OperatorCode & ""
190               .RunDate = Format$(tb!RunDate, "dd/mm/yyyy")
200               .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
210               .Printed = tb!Printed
220               .Valid = tb!Valid
                  '.Flag = tb!Flag & ""
230               .Units = tb!Units & ""
240               .TestName = tb!TestName & ""
250               .Low = IIf(IsNull(tb!Low), 0, tb!Low)
260               .High = IIf(IsNull(tb!High), 999, tb!High)
270               .DP = tb!DP
280               .Printable = IIf(IsNull(tb!Printable), True, tb!Printable)
290               .Analyser = tb!Analyser & ""
                  .PrintRefRange = tb!PrintRefRange
300               .SignOff = tb!SignOff
310               .SignOffBy = tb!SignOffBy
320               If Not IsNull(tb!SignOffDateTime) Then
330                   .SignOffDateTime = tb!SignOffDateTime
340               End If

350               CRs.Add CR
360           End With
370           tb.MoveNext
380       Loop

390       If CRs.Count <> 0 Then
400           Set Load = CRs
410       Else
420           Set Load = Nothing
430       End If

440       Exit Function

Load_Error:

          Dim strES As String
          Dim intEL As Integer

450       intEL = Erl
460       strES = Err.Description
470       LogError "CoagResults", "Load", intEL, strES, sql


End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

10    Set NewEnum = mcolCoagResults.[_NewEnum]

End Function



Public Sub Remove(ByVal Index As Long)

10    mcolCoagResults.Remove (Index)

End Sub

Public Sub Save(ByVal CRs As CoagResults)

      Dim sql As String
      Dim tb As Recordset
      Dim CR As CoagResult
      Dim lngSampleID As Long
      Dim RunDate As String

10    On Error GoTo Save_Error

20    For Each CR In CRs
30        With CR
40            If IsNumeric(.Code) Then
50                lngSampleID = Val(.SampleID)
60                sql = "Select * from CoagResults where " & _
                        "SampleID = '" & lngSampleID & "' " & _
                        "and  Code = '" & .Code & "' and units = '" & .Units & "'"
70                Set tb = New Recordset
80                RecOpenServer 0, tb, sql
90                If Not tb.EOF Then
100                   sql = "Select * from CoagRepeats where " & _
                            "SampleID = '" & lngSampleID & "' " & _
                            "and  Code = '" & CR.Code & "' and units = '" & .Units & "'"
110                   Set tb = New Recordset
120                   RecOpenServer 0, tb, sql
130               End If
140               tb.AddNew
150               tb!SampleID = lngSampleID
160               RunDate = Format$(.RunDate, "dd/mmm/yyyy")
170               tb!RunDate = RunDate
180               If Trim$(.RunTime) <> "" Then
190                   tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
200               End If
210               tb!Code = Left$(.Code, 4)
220               If IsNumeric(.Result) Then
230                   tb!Result = Left$(Format$(Val(.Result)), 6)
240               Else
250                   tb!Result = Left$(.Result, 6)
260               End If
270               tb!Printed = IIf(.Printed, 1, 0)
280               tb!Valid = IIf(.Valid, 1, 0)
290               tb!Units = .Units
300               tb.Update
310               sql = "Delete from CoagRequests where " & _
                        "sampleid = '" & lngSampleID & "' " & _
                        "and Code = '" & .Code & "' and units = '" & .Units & "'"
320               Cnxn(0).Execute sql
330           End If
340       End With
350   Next

360   Set CR = Nothing

370   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "CoagResults", "Save", intEL, strES, sql


End Sub


Public Sub LogAsPrinted(ByVal Cn As Integer, ByVal SampleID As String)

      Dim sql As String

10    On Error GoTo LogAsPrinted_Error

20    sql = "Update CoagResults set Printed = 1 where " & _
            "SampleID = '" & SampleID & "'"
30    Cnxn(Cn).Execute sql

40    Exit Sub

LogAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "CoagResults", "LogAsPrinted", intEL, strES, sql


End Sub



