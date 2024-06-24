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

6700      mcolCoagResults.Add CR

6710      Set Add = CR

End Function

Public Property Get CoagResults() As CoagResults
        
6720      Set CoagResults = mCoagResults

End Property

Public Sub Clear()

          Dim n As Integer

6730      For n = 1 To mcolCoagResults.Count
6740          mcolCoagResults.Remove 1
6750      Next

End Sub

Public Function Count() As Long
        
6760      Count = mcolCoagResults.Count

End Function


Public Sub Delete(ByVal SampleID As String, _
          ByVal Code As String)

          Dim sql As String

6770      On Error GoTo Delete_Error

6780      sql = "Delete from CoagResults where " & _
              "SampleID = '" & SampleID & "' " & _
              "and Code = '" & Code & "'"

6790      Cnxn(0).Execute sql

6800      Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

6810      intEL = Erl
6820      strES = Err.Description
6830      LogError "CoagResults", "Delete", intEL, strES, sql

End Sub

Public Function Item(ByVal X As Long) _
          As CoagResult
Attribute Item.VB_UserMemId = 0
        
6840      On Error Resume Next

6850      Set Item = mcolCoagResults(X)

End Function


Public Function Load(ByVal SampleID As String, _
          ByVal v As Integer, _
          ByVal p As Integer, _
          ByVal ResultsOrRepeats As String, _
          Optional ByVal Connection As Integer) _
          As CoagResults

          Dim CRs As New CoagResults
          Dim CR As CoagResult
          Dim sql As String
          Dim tb As Recordset
          Dim Sex As String
          Dim DaysOld As Long
          Dim strNormalRange As String
          Dim rs As Recordset

6860      On Error GoTo Load_Error

6870      If Val(SampleID) = 0 Then
6880          Set Load = Nothing
6890          Exit Function
6900      End If

6910      sql = "Select DoB, Sex from Demographics where SampleID = '" & SampleID & "'"
6920      Set rs = New Recordset
6930      RecOpenServer Connection, rs, sql
6940      If Not rs.EOF Then
6950          Sex = UCase$(Left$(Trim$(rs!Sex & ""), 1))
6960          If IsDate(rs!DoB & "") Then
6970              DaysOld = DateDiff("d", rs!DoB, Now)
6980          Else
6990              DaysOld = 0
7000          End If
7010      End If
7020      rs.Close

7030      Select Case Sex
              Case "M":  strNormalRange = " D.MaleLow as Low, D.MaleHigh as High, "
7040          Case "F":  strNormalRange = " D.FemaleLow as Low, D.FemaleHigh as High, "
7050          Case Else: strNormalRange = " D.FemaleLow as Low, D.MaleHigh as High, "
7060      End Select

7070      sql = "select R.SampleID, R.Result, R.Analyser, " & strNormalRange & _
              "R.RunDate, R.RunTime, R.Valid, R.Authorised, " & _
              "R.Printed, D.TestName, D.code, D.Units, D.Printable, d.DP, D.InUse, R.Released, D.PlausibleLow, D.PlausibleHigh, " & _
              "D.PrintPriority from Coag" & ResultsOrRepeats & " as R, CoagTestDefinitions as D where " & _
              "R.SampleID = '" & SampleID & "' " & _
              "and R.Code = D.Code "
7080      If v = gVALID And p = gNOTPRINTED Then
7090          sql = sql & "and valid = 1 and printed = 0 "
7100      ElseIf v = gVALID And p = gDONTCARE Then
7110          sql = sql & "and valid = 1 "
7120      ElseIf v = gNOTVALID And p = gDONTCARE Then
7130          sql = sql & "and valid = 0 "
7140      End If
7150      If DaysOld <> 0 Then
7160          sql = sql & "and AgeFromDays <= '" & DaysOld & "' and AgeToDays >= '" & DaysOld & "' "
7170      End If
7180      sql = sql & "order by PrintPriority"
7190      Set tb = New Recordset
7200      RecOpenServer Connection, tb, sql

7210      Do While Not tb.EOF
7220          Set CR = New CoagResult
7230          With CR
7240              .SampleID = tb!SampleID & ""
7250              .Code = tb!Code & ""
7260              .Result = tb!Result & ""
7270              .TestName = tb!TestName & ""
7280              .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
7290              .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
7300              .Printable = IIf(IsNull(tb!Printable), 0, tb!Printable)
7310              .Printed = tb!Printed
7320              .Valid = tb!Valid
7330              .Authorised = IIf(Not IsNull(tb!Authorised), tb!Authorised, 0)
7340              .Released = IIf(Not IsNull(tb!Released), tb!Released, 0)
7350              .Units = tb!Units & ""
7360              .Low = tb!Low & ""
7370              .High = tb!High & ""
                  '+++ Junaid 26-01-2024
7380              .PLow = ConvertNull(tb!PlausibleLow, "")
7390              .PHigh = ConvertNull(tb!PlausibleHigh, "")
                  '--- Junaid
7400              .DP = IIf(IsNull(tb!DP), 1, tb!DP)
7410              .InUse = IIf(IsNull(tb!InUse), False, tb!InUse)
7420              .Analyser = tb!Analyser & ""
7430              CRs.Add CR
7440          End With
7450          tb.MoveNext
7460      Loop

7470      If CRs.Count <> 0 Then
7480          Set Load = CRs
7490      Else
7500          Set Load = Nothing
7510      End If

7520      Exit Function

Load_Error:

          Dim strES As String
          Dim intEL As Integer

7530      intEL = Erl
7540      strES = Err.Description
7550      LogError "CoagResults", "Load", intEL, strES, sql

End Function
Public Function LoadRepeats(ByVal SampleID As String) As CoagResults

          Dim CRs As New CoagResults
          Dim CR As CoagResult
          Dim sql As String
          Dim tb As Recordset

7560      On Error GoTo LoadRepeats_Error

7570      sql = "select R.*, PrintPriority, TestName, Units from CoagRepeats as R, CoagTestDefinitions as D where " & _
              "SampleID = '" & SampleID & "' " & _
              "and R.Code = D.Code " & _
              "order by PrintPriority"
7580      Set tb = New Recordset
7590      RecOpenClient 0, tb, sql
7600      Do While Not tb.EOF
7610          Set CR = New CoagResult
7620          With CR
7630              .TestName = tb!TestName & ""
7640              .Units = tb!Units & ""
7650              .SampleID = tb!SampleID & ""
7660              .Code = tb!Code & ""
7670              .Result = tb!Result & ""
                  '.OperatorCode = tb!OperatorCode & ""
7680              .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
                  '    .RunTime = format$(tb!RunTime, "dd/mm/yyyy hh:mm")
7690              .Printed = tb!Printed
7700              .Valid = tb!Valid
7710              .Analyser = tb!Analyser & ""
                  '.Flag = tb!Flag & ""
7720              CRs.Add CR
7730          End With
7740          tb.MoveNext
7750      Loop

7760      If CRs.Count <> 0 Then
7770          Set LoadRepeats = CRs
7780      Else
7790          Set LoadRepeats = Nothing
7800      End If

7810      Exit Function

LoadRepeats_Error:

          Dim strES As String
          Dim intEL As Integer

7820      intEL = Erl
7830      strES = Err.Description
7840      LogError "CoagResults", "LoadRepeats", intEL, strES, sql

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
7850      Set NewEnum = mcolCoagResults.[_NewEnum]

End Function



Public Sub Remove(ByVal Index As Long)

7860      mcolCoagResults.Remove (Index)

End Sub

Public Sub Save(ByVal CRs As CoagResults)

          Dim sql As String
          Dim tb As Recordset
          Dim CR As CoagResult
          Dim Rundate As String
          Dim SampleID As Long

7870      On Error GoTo Save_Error

7880      For Each CR In CRs
7890          With CR
7900              If Not IsNumeric(.SampleID) Then
7910                  SampleID = 0
7920              Else
7930                  SampleID = CLng(.SampleID)
7940                  If SampleID <> 0 Then
7950                      sql = "Select * from CoagResults where " & _
                              "SampleID = '" & SampleID & "' " & _
                              "and  Code = '" & .Code & "'"
7960                      Set tb = New Recordset
7970                      RecOpenClient 0, tb, sql
7980                      If Not tb.EOF Then
7990                          sql = "Select * from CoagRepeats where " & _
                                  "SampleID = '" & SampleID & "' " & _
                                  "and  Code = '" & .Code & "'"
8000                          Set tb = New Recordset
8010                          RecOpenClient 0, tb, sql
8020                      End If
8030                      tb.AddNew
8040                      tb!SampleID = SampleID
8050                      Rundate = Format$(.Rundate, "dd/mmm/yyyy")
8060                      tb!Rundate = Rundate
8070                      If Trim$(.RunTime) <> "" Then
8080                          tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
8090                      End If
8100                      tb!Code = Left$(.Code, 3)
8110                      If IsNumeric(.Result) Then
8120                          tb!Result = Left$(Format$(Val(.Result)), 6)
8130                      Else
8140                          tb!Result = Left$(.Result, 6)
8150                      End If
8160                      tb!Printed = IIf(.Printed, 1, 0)
8170                      tb!Valid = IIf(.Valid, 1, 0)
8180                      tb.Update
8190                      sql = "Delete from CoagRequests where " & _
                              "SampleID = '" & SampleID & "'"
8200                      Cnxn(0).Execute sql
8210                  End If
8220              End If
8230          End With
8240      Next

8250      Set CR = Nothing

8260      Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

8270      intEL = Erl
8280      strES = Err.Description
8290      LogError "CoagResults", "Save", intEL, strES, sql

End Sub


Public Sub LogAsPrinted(ByVal SampleID As String)

          Dim sql As String

8300      On Error GoTo LogAsPrinted_Error

8310      sql = "Update CoagResults set Printed = 1 where " & _
              "SampleID = '" & SampleID & "'"
8320      Cnxn(0).Execute sql

8330      Exit Sub

LogAsPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

8340      intEL = Erl
8350      strES = Err.Description
8360      LogError "CoagResults", "LogAsPrinted", intEL, strES, sql

End Sub




