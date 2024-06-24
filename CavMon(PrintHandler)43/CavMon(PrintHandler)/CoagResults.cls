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
20      mcolCoagResults.Remove 1
30    Next

End Sub

Public Function Count() As Long
  
10    Count = mcolCoagResults.Count

End Function


Public Sub Delete(ByVal SampleID As String, _
                  ByVal Code As String)

      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "Delete from CoagResults where " & _
            "SampleID = '" & SampleID & "' " & _
            "and Code = '" & Code & "'"

30    Cnxn(0).Execute sql

40    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "CoagResults", "Delete", intEL, strES, sql

End Sub

Public Function Item(ByVal X As Long) _
                     As CoagResult
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolCoagResults(X)

End Function


Public Function Load(ByVal SampleID As String, _
                     ByVal v As Integer, _
                     ByVal ResultsOrRepeats As String) _
                     As CoagResults

      Dim CRs As New CoagResults
      Dim CR As CoagResult
      Dim sql As String
      Dim tb As Recordset
      Dim Sex As String
      Dim DaysOld As Long
      Dim strNormalRange As String
      Dim rs As Recordset

10    On Error GoTo Load_Error

20    sql = "Select DoB, Sex from Demographics where SampleID = '" & SampleID & "'"
30    Set rs = New Recordset
40    RecOpenServer 0, rs, sql
50    If Not rs.EOF Then
60      Sex = UCase$(Left$(Trim$(rs!Sex & ""), 1))
70      If IsDate(rs!Dob & "") Then
80        DaysOld = DateDiff("d", rs!Dob, Now)
90      Else
100       DaysOld = 0
110     End If
120   End If
130   rs.Close

140   Select Case Sex
        Case "M":  strNormalRange = " D.MaleLow as Low, D.MaleHigh as High, "
150     Case "F":  strNormalRange = " D.FemaleLow as Low, D.FemaleHigh as High, "
160     Case Else: strNormalRange = " D.FemaleLow as Low, D.MaleHigh as High, "
170   End Select

180   sql = "select R.SampleID, R.Result, " & strNormalRange & _
            "R.RunDate, R.RunTime, R.Valid, R.Authorised, D.PrintRefRange, " & _
            "R.Printed, D.TestName, D.code, D.Units, D.Printable, d.DP, D.InUse, R.Released, " & _
            "D.PrintPriority from Coag" & ResultsOrRepeats & " as R, CoagTestDefinitions as D where " & _
            "R.SampleID = '" & SampleID & "' " & _
            "and R.Code = D.Code " & _
            "AND TestName NOT IN " & _
            "  ( SELECT Parameter FROM PrintInhibit WHERE " & _
            "    SampleID = '" & SampleID & "' " & _
            "    AND Discipline = 'Coa') "
190   If v = gVALID Then
200     sql = sql & "and valid = 1  "
210   ElseIf v = gNOTVALID Then
220     sql = sql & "and valid = 0 "
230   End If
240   If DaysOld <> 0 Then
250     sql = sql & "and AgeFromDays <= '" & DaysOld & "' and AgeToDays >= '" & DaysOld & "' "
260   End If
270   sql = sql & "order by PrintPriority"
280   Set tb = New Recordset
290   RecOpenServer 0, tb, sql
300   Debug.Print sql

310   Do While Not tb.EOF
320     Set CR = New CoagResult
330     With CR
340       .SampleID = tb!SampleID & ""
350       .Code = tb!Code & ""
360       .Result = tb!Result & ""
370       .TestName = tb!TestName & ""
380       .RunDate = Format$(tb!RunDate, "dd/mm/yyyy")
390       .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
400       .Printable = IIf(IsNull(tb!Printable), 0, tb!Printable)
410       .Printed = tb!Printed
420       .Valid = tb!Valid
430       .Authorised = IIf(Not IsNull(tb!Authorised), tb!Authorised, 0)
440       .Released = IIf(Not IsNull(tb!Released), tb!Released, 0)
450       .Units = tb!Units & ""
460       .Low = tb!Low & ""
470       .High = tb!High & ""
480       .DP = IIf(IsNull(tb!DP), 1, tb!DP)
490       .InUse = IIf(IsNull(tb!InUse), False, tb!InUse)
500       .PrintRefRange = tb!PrintRefRange
510       CRs.Add CR
520     End With
530     tb.MoveNext
540   Loop

550   If CRs.Count <> 0 Then
560     Set Load = CRs
570   Else
580     Set Load = Nothing
590   End If

600   Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer

610   intEL = Erl
620   strES = Err.Description
630   LogError "CoagResults", "Load", intEL, strES, sql

End Function
Public Function LoadRepeats(ByVal SampleID As String) As CoagResults

      Dim CRs As New CoagResults
      Dim CR As CoagResult
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo LoadRepeats_Error

20    sql = "select R.*, PrintPriority, TestName, Units from CoagRepeats as R, CoagTestDefinitions as D where " & _
            "SampleID = '" & SampleID & "' " & _
            "and R.Code = D.Code " & _
            "order by PrintPriority"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set CR = New CoagResult
70      With CR
80        .TestName = tb!TestName & ""
90        .Units = tb!Units & ""
100       .SampleID = tb!SampleID & ""
110       .Code = tb!Code & ""
120       .Result = tb!Result & ""
          '.OperatorCode = tb!OperatorCode & ""
130       .RunDate = Format$(tb!RunDate, "dd/mm/yyyy")
      '    .RunTime = format$(tb!RunTime, "dd/mm/yyyy hh:mm")
140       .Printed = tb!Printed
150       .Valid = tb!Valid
          '.Flag = tb!Flag & ""
160       CRs.Add CR
170     End With
180     tb.MoveNext
190   Loop

200   If CRs.Count <> 0 Then
210     Set LoadRepeats = CRs
220   Else
230     Set LoadRepeats = Nothing
240   End If

250   Exit Function

LoadRepeats_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "CoagResults", "LoadRepeats", intEL, strES, sql

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolCoagResults.[_NewEnum]

End Function



Public Sub Remove(ByVal index As Long)

10    mcolCoagResults.Remove (index)

End Sub

Public Sub Save(ByVal CRs As CoagResults)

      Dim sql As String
      Dim tb As Recordset
      Dim CR As CoagResult
      Dim RunDate As String
      Dim SampleID As Long

10    On Error GoTo Save_Error

20    For Each CR In CRs
30      With CR
40        If Not IsNumeric(.SampleID) Then
50          SampleID = 0
60        Else
70          SampleID = CLng(.SampleID)
80          If SampleID <> 0 Then
90            sql = "Select * from CoagResults where " & _
                    "SampleID = '" & SampleID & "' " & _
                    "and  Code = '" & .Code & "'"
100           Set tb = New Recordset
110           RecOpenClient 0, tb, sql
120           If Not tb.EOF Then
130             sql = "Select * from CoagRepeats where " & _
                      "SampleID = '" & SampleID & "' " & _
                      "and  Code = '" & .Code & "'"
140             Set tb = New Recordset
150             RecOpenClient 0, tb, sql
160           End If
170           tb.AddNew
180           tb!SampleID = SampleID
190           RunDate = Format$(.RunDate, "dd/mmm/yyyy")
200           tb!RunDate = RunDate
210           If Trim$(.RunTime) <> "" Then
220             tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
230           End If
240           tb!Code = Left$(.Code, 3)
250           If IsNumeric(.Result) Then
260             tb!Result = Left$(Format$(Val(.Result)), 6)
270           Else
280             tb!Result = Left$(.Result, 6)
290           End If
300           tb!Printed = IIf(.Printed, 1, 0)
310           tb!Valid = IIf(.Valid, 1, 0)
320           tb.Update
330           sql = "Delete from CoagRequests where " & _
                    "SampleID = '" & SampleID & "'"
340           Cnxn(0).Execute sql
350         End If
360       End If
370     End With
380   Next

390   If SampleID <> 0 Then LogForCoag SampleID

400   Set CR = Nothing

410   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

420   intEL = Erl
430   strES = Err.Description
440   LogError "CoagResults", "Save", intEL, strES, sql

End Sub


Public Sub LogAsPrinted(ByVal SampleID As String)

      Dim sql As String

10    On Error GoTo LogAsPrinted_Error

20    sql = "Update CoagResults set Printed = 1 where " & _
            "SampleID = '" & SampleID & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

LogAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "CoagResults", "LogAsPrinted", intEL, strES, sql

End Sub



