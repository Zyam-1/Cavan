VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolBIEResults As New Collection
Private mBIEResults As New BIEResults

Public Function Add(BR As BIEResult) As BIEResult

10    mcolBIEResults.Add BR

20    Set Add = BR

End Function

Public Property Get BIEResults() As BIEResults
  
10    Set BIEResults = mBIEResults

End Property

Public Sub Clear()

      Dim n As Integer

10    For n = 1 To mcolBIEResults.Count
20      mcolBIEResults.Remove 1
30    Next

End Sub

Public Function Count() As Long
  
10    Count = mcolBIEResults.Count

End Function

Public Function Item(ByVal X As Long) _
                     As BIEResult
Attribute Item.VB_UserMemId = 0
        
10    On Error Resume Next

20    Set Item = mcolBIEResults(X)


End Function


Public Function Load(ByVal Discipline As String, _
                     ByVal SampleID As String, _
                     ByVal ResultOrRepeat As String, _
                     ByVal v As Integer, _
                     ByVal p As Integer, _
                     Optional ByVal Cat As String, _
                     Optional ByVal Connection As Integer, _
                     Optional ByVal Sex As String = "", _
                     Optional ByVal Dob As String = "") _
                     As BIEResults
      'Discipline is either "Bio", "Imm" or "End"

      Dim BRs As New BIEResults
      Dim BR As BIEResult
      Dim tb As Recordset
      Dim tbRange As Recordset
      Dim tbNewIDX As Recordset
      Dim sql As String
      Dim DaysOld As Long
      Dim SelectNormalRange As String
      Dim SelectFlagRange As String
      Dim TableName As String
      Dim SelectSexLow As String
      Dim SelectSexHigh As String
      Dim Accredited As String

10    On Error GoTo Load_Error

20    If Cat = "" Then Cat = "Human"

30    If Val(SampleID) = 0 Then
40        Set BRs = Nothing
50        Exit Function
60    End If

70    TableName = Discipline & ResultOrRepeat

80    If Sex = "" And Dob = "" Then
90        sql = "SELECT DoB, Sex FROM Demographics WHERE " & _
                "SampleID = '" & SampleID & "'"
100       Set tb = Cnxn(Connection).Execute(sql)
110       If Not tb.EOF Then
120           Dob = tb!Dob & ""
130           Sex = tb!Sex & ""
140       End If
150   End If

160   If IsDate(Dob) Then
170       Dob = Format$(Dob, "dd/mmm/yyyy")
180       DaysOld = DateDiff("d", Dob, Now)
190   End If

200   Select Case Left$(UCase$(Trim$(Sex)), 1)
      Case "M": SelectSexLow = "Male": SelectSexHigh = "Male"
210   Case "F": SelectSexLow = "Female": SelectSexHigh = "Female"
220   Case Else: SelectSexLow = "Female": SelectSexHigh = "Male"
230   End Select

240   SelectNormalRange = " COALESCE(" & SelectSexLow & "Low, 0) Low, " & _
                          " COALESCE(" & SelectSexHigh & "High, 9999) High, "
250   SelectFlagRange = " COALESCE(Flag" & SelectSexLow & "Low, 0) FlagLow, " & _
                        " COALESCE(Flag" & SelectSexHigh & "High, 9999) FlagHigh, "

260   sql = "UPDATE " & TableName & " " & _
            "SET Result = 'XXXXX' " & _
            "WHERE SampleID = '" & SampleID & "' " & _
            "AND ( Result LIKE 'I_' " & _
            "      OR Result LIKE 'L_' " & _
            "      OR Result LIKE 'H_' )"
270   Cnxn(0).Execute sql

280   If Discipline = "Bio" Then
290       Accredited = ",ISNull(Accredited,0)  as Accredited "
300   Else
310       Accredited = ",0  as Accredited "
320   End If

330   sql = "SELECT R.DefIndex, X.NormalLow Low, X.NormalHigh High, X.FlagLow, X.FlagHigh, X.PlausibleLow, X.PlausibleHigh, " & _
            "LongName, ShortName, DoDelta, DeltaLimit, " & _
            "DP, PrintPriority, COALESCE(PrintSplit, 0) PS, " & _
            "R.SampleID, R.Code, R.Result, " & _
            "COALESCE(R.Valid, 0) AS Valid, COALESCE(R.Printed, 0) Printed, " & _
            "COALESCE(D.Printable, 0) AS Printable, " & _
            "R.RunTime, R.RunDate, R.Operator, R.Flags, R.Units, " & _
            "R.SampleType, R.Analyser, R.Faxed, R.Authorised, " & _
            "R.Comment AS Comment, " & _
            "COALESCE(D.PrintRefRange, 1) PrnRefRange " & Accredited & _
            "FROM " & TableName & " R, " & Discipline & "TestDefinitions D, " & _
            Discipline & "DefIndex X " & _
            "WHERE " & _
            "R.DefIndex = X.Defindex " & _
            "AND SampleID = '" & SampleID & "' AND D.category = '" & Cat & "' " & _
            "AND R.Code = D.Code " & _
            "AND AgeFromDays <= " & DaysOld & " " & _
            "AND AgeToDays >= " & DaysOld & " "

340   If UCase(HospName(0)) <> "CAVAN" Then
350       sql = sql & "AND R.SampleType = D.SampleType "
360   End If
370   If p = gNOTPRINTED And v = gNOTVALID Then
380       sql = sql & "AND Printed = 0 AND Valid = 0 "
390   ElseIf p = gNOTPRINTED And v = gVALID Then
400       sql = sql & "AND Printed = 0 AND Valid = 1 "
410   ElseIf p = gNOTPRINTED And v = gDONTCARE Then
420       sql = sql & "AND Printed = 0 "
430   ElseIf p = gPRINTED And v = gNOTVALID Then
440       sql = sql & "AND Printed = 1 AND Valid = 0 "
450   ElseIf p = gPRINTED And v = gVALID Then
460       sql = sql & "AND Printed = 1 AND Valid = 1 "
470   ElseIf p = gPRINTED And v = gDONTCARE Then
480       sql = sql & "AND Printed = 1 "
490   ElseIf p = gDONTCARE And v = gNOTVALID Then
500       sql = sql & "AND Valid = 0 "
510   ElseIf p = gDONTCARE And v = gVALID Then
520       sql = sql & "AND Valid = 1 "
530   End If

540   sql = sql & "ORDER BY PrintPriority ASC"
550   Set tb = New Recordset
560   RecOpenServer Connection, tb, sql    '  RecOpenClient 0,tb, Sql
570   Do While Not tb.EOF
580       Set BR = New BIEResult
590       With BR
600           .SampleID = tb!SampleID & ""
610           .Code = tb!Code & ""
620           .Result = tb!Result & ""
630           .Operator = tb!Operator & ""
640           .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
650           .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm:ss")
660           .Units = tb!Units & ""
670           .Printed = tb!Printed
680           .Valid = IIf(tb!Valid, True, False)
690           If Not IsNull(tb!Authorised) Then
700               .Authorised = IIf(tb!Authorised, True, False)
710           Else
720               .Authorised = False
730           End If
              '    .Faxed = tb!Faxed
              '    If TB!h & "" <> "" Then .h = TB!h Else .h = False
              '    If TB!g & "" <> "" Then .g = TB!g Else .g = False
              '    If TB!o & "" <> "" Then .o = TB!o Else .o = False
              '    If TB!J & "" <> "" Then .J = TB!h Else .J = False
              '    If TB!l & "" <> "" Then .l = TB!h Else .l = False
              '    If TB!s & "" <> "" Then .s = TB!h Else .s = False
740           .Flags = tb!Flags & ""
750           .SampleType = tb!SampleType & ""
760           .Printformat = IIf(IsNull(tb!DP), 0, tb!DP)
770           .ShortName = tb!ShortName & ""

780           .LongName = tb!LongName & ""
790           If tb!Accredited = 1 Then
800               .LongName = .LongName & " ª"
810           End If
820           If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
830           .DeltaLimit = IIf(IsNull(tb!DeltaLimit), 9999, tb!DeltaLimit)
840           .Analyser = tb!Analyser & ""
850           .Comment = tb!Comment & ""
860           .Low = tb!Low
870           .High = tb!High
880           .FlagLow = tb!FlagLow
890           .FlagHigh = tb!FlagHigh
900           .PlausibleLow = tb!PlausibleLow
910           .PlausibleHigh = tb!PlausibleHigh
920           .DefIndex = tb!DefIndex
930           .PrintSplit = tb!PS
940           .PrintRefRange = tb!PrnRefRange
950           .Printable = tb!Printable
960           If .DefIndex = 0 Then
970               If Dob <> "" And Sex <> "" Then
980                   sql = "SELECT " & _
                            SelectNormalRange & SelectFlagRange & _
                            "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
                            "COALESCE(PlausibleHigh, 9999) PlausibleHigh " & _
                            "FROM " & Discipline & "TestDefinitions  " & _
                            "WHERE category = '" & Cat & "' " & _
                            "AND Code = '" & .Code & "' " & _
                            "AND AgeFromDays <= " & DaysOld & " " & _
                            "AND AgeToDays >= " & DaysOld & " "
990                   Set tbRange = New Recordset
1000                  RecOpenServer 0, tbRange, sql
1010                  If Not tbRange.EOF Then
1020                      .Low = tbRange!Low
1030                      .High = tbRange!High
1040                      .FlagLow = tbRange!FlagLow
1050                      .FlagHigh = tbRange!FlagHigh
                          
                          
1060                      .PlausibleLow = tbRange!PlausibleLow
1070                      .PlausibleHigh = tbRange!PlausibleHigh




1080                      sql = "SELECT * FROM " & Discipline & "DefIndex " & _
                                "WHERE NormalLow = '" & .Low & "' " & _
                                "AND NormalHigh = '" & .High & "' " & _
                                "AND FlagLow = '" & .FlagLow & "' " & _
                                "AND FlagHigh = '" & .FlagHigh & "' " & _
                                "AND PlausibleLow = '" & .PlausibleLow & "' " & _
                                "AND PlausibleHigh = '" & .PlausibleHigh & "' "

1090                      Set tbNewIDX = New Recordset
1100                      RecOpenClient 0, tbNewIDX, sql
1110                      If Not tbNewIDX.EOF Then
1120                          .DefIndex = tbNewIDX!DefIndex
1130                      Else

1140                          sql = "INSERT INTO " & Discipline & "DefIndex " & _
                                    "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
                                    "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
                                    "VALUES ( " & _
                                    .Low & ", " & .High & ", " & .FlagLow & ", " & .FlagHigh & ", " & _
                                    .PlausibleLow & ", " & .PlausibleHigh & ", 0,9999) "
1150                          Cnxn(0).Execute sql

1160                          sql = "SELECT MAX(DefIndex) NewIndex FROM " & Discipline & "DefIndex"
1170                          Set tbNewIDX = New Recordset
1180                          RecOpenClient 0, tbNewIDX, sql
1190                          .DefIndex = tbNewIDX!NewIndex

1200                      End If

1210                      sql = "UPDATE " & TableName & " " & _
                                "SET DefIndex = '" & .DefIndex & "' " & _
                                "WHERE SampleID = '" & .SampleID & "' " & _
                                "AND Code = '" & .Code & "'"
1220                      Cnxn(0).Execute sql

1230                  End If
1240              End If
1250          End If


1260          BRs.Add BR
1270      End With
1280      tb.MoveNext
1290  Loop

1300  If BRs.Count <> 0 Then
1310      Set Load = BRs
1320  Else
1330      Set Load = Nothing
1340  End If
1350  Set BR = Nothing
1360  Set BRs = Nothing

1370  Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer

1380  intEL = Erl
1390  strES = Err.Description
1400  LogError "BIEResults", "Load", intEL, strES, sql

End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolBIEResults.[_NewEnum]

End Function



Public Sub RemoveItem(ByVal X As Long)

10    mcolBIEResults.Remove X

End Sub


Public Sub Save(ByVal Discipline As String, _
                ByVal BRs As BIEResults)
      'Discipline is either "Bio", "Imm" or "End"

      Dim tb As Recordset
      Dim sql As String
      Dim BR As BIEResult
      Dim lngSampleID As Long

10    On Error GoTo Save_Error

20    For Each BR In BRs
30      With BR
40        lngSampleID = Format$(Val(.SampleID))
50        sql = "Select * from " & Discipline & "Results where " & _
                "SampleID = '" & lngSampleID & "' " & _
                "and Code = '" & BR.Code & "'"
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql
80        If Not tb.EOF Then
90          sql = "Select * from " & Discipline & "Repeats where " & _
                  "SampleID = '" & lngSampleID & "'"
100         Set tb = New Recordset
110         RecOpenClient 0, tb, sql
120       End If
130       tb.AddNew
140       tb!SampleID = lngSampleID
150       tb!Rundate = Format$(.Rundate, "dd/mmm/yyyy")
160       If Trim$(.RunTime) <> "" Then
170         tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
180       End If
190       tb!Code = .Code
200       tb!Result = .Result
210       tb!Units = .Units
220       tb!Printed = .Printed
230       tb!Valid = .Valid
240       tb!Faxed = 0
250       tb!Analyser = .Analyser
260       tb!SampleType = .SampleType
270       tb!NormalLow = .NormalLow
280       tb!NormalHigh = .NormalHigh
290       tb!NormalUsed = .NormalUsed
300       tb.Update

310       sql = "Delete from " & Discipline & "Requests where " & _
                "SampleID = '" & lngSampleID & "' " & _
                "and Code = '" & BR.Code & "'"
320       Cnxn(0).Execute sql
    
330     End With
340   Next

350   Set BR = Nothing

360   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "BIEResults", "Save", intEL, strES, sql

End Sub




