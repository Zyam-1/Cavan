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

Public Function Add(Br As BIEResult) As BIEResult

      On Error Resume Next
10    mcolBIEResults.Add Br, Br.Code
      'MsgBox (mcolBIEResults.Add(Br.Code))

20    Set Add = Br

End Function

' Public Function Add(Br As BIEResult) As BIEResult
'    If Not KeyExists(mcolBIEResults, Br.Code) Then
'        mcolBIEResults.Add Br, Br.Code
'        Set Add = Br
'    Else
'        MsgBox "Code already exists."
'        Set Add = Nothing
'    End If
'End Function
'
'Private Function KeyExists(col As Collection, key As String) As Boolean
'    On Error Resume Next
'    Dim temp As Object
'    Set temp = col(key)
'    KeyExists = (Err.Number = 0)
'    On Error GoTo 0
'End Function
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

Public Function Item(ByVal Code As String) _
                     As BIEResult
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolBIEResults(Code)

End Function


Public Function Load(ByVal Discipline As String, _
                     ByVal SampleID As String, _
                     ByVal ResultOrRepeat As String, _
                     ByVal v As Integer, _
                     ByVal P As Integer, _
                     Optional ByVal Cat As String, _
                     Optional ByVal Connection As Integer, _
                     Optional ByVal Sex As String = "", _
                     Optional ByVal DoB As String = "") _
                     As BIEResults
      'Discipline is either "Bio", "Imm" or "End"

      Dim BRs As New BIEResults
      Dim Br As BIEResult
      Dim tb As Recordset
      Dim TempTb As Recordset
      Dim tbRange As Recordset
      Dim tbNewIDX As Recordset
      Dim sql As String
      Dim TempSQL As String
      Dim DaysOld As Long
      Dim SelectNormalRange As String
      Dim SelectFlagRange As String
      Dim TableName As String
      Dim SelectSexLow As String
      Dim SelectSexHigh As String
      Dim Fasting As Boolean
      Dim sampleDate As String
10    On Error GoTo Load_Error

20    If Cat = "" Then Cat = "Human"
'MsgBox SampleID
30    If Val(SampleID) = 0 Then
40        Set BRs = Nothing
50        Exit Function
60    End If

70    TableName = Discipline & ResultOrRepeat

80    If Sex = "" And DoB = "" Then
90        sql = "SELECT DoB, Sex, COALESCE(fasting,0) fasting, SampleDate FROM Demographics WHERE " & _
                "SampleID = '" & SampleID & "'"
100       Set tb = Cnxn(Connection).Execute(sql)
110       If Not tb.EOF Then
120           DoB = tb!DoB & ""
130           Sex = tb!Sex & ""
140           Fasting = tb!Fasting
              sampleDate = tb!sampleDate

150       End If
160   End If

170   If IsDate(DoB) Then
180       DoB = Format$(DoB, "dd/mmm/yyyy")
190       DaysOld = DateDiff("d", DoB, sampleDate)
'MsgBox DaysOld
200   End If

210   Select Case Left$(UCase$(Trim$(Sex)), 1)
      Case "M": SelectSexLow = "Male": SelectSexHigh = "Male"
220   Case "F": SelectSexLow = "Female": SelectSexHigh = "Female"
230   Case Else: SelectSexLow = "Female": SelectSexHigh = "Male"
240   End Select

250   SelectNormalRange = " COALESCE(" & SelectSexLow & "Low, 0) Low, " & _
                          " COALESCE(" & SelectSexHigh & "High, 9999) High, "
260   SelectFlagRange = " COALESCE(Flag" & SelectSexLow & "Low, 0) FlagLow, " & _
                        " COALESCE(Flag" & SelectSexHigh & "High, 9999) FlagHigh, "

270   sql = "UPDATE " & TableName & " " & _
            "SET Result = 'XXXXX' " & _
            "WHERE SampleID = '" & SampleID & "' " & _
            "AND ( Result LIKE 'I_' " & _
            "      OR Result LIKE 'L_' " & _
            "      OR Result LIKE 'H_' )"
280   Cnxn(0).Execute sql

290   sql = "SELECT R.DefIndex, X.NormalLow Low, X.NormalHigh High, X.FlagLow, X.FlagHigh, X.PlausibleLow, X.PlausibleHigh, " & _
            "LongName, ShortName, DoDelta, DeltaLimit, " & _
            "DP, PrintPriority, " & _
            "R.SampleID, R.Code, R.Result, " & _
            "COALESCE(R.Valid, 0) AS Valid, R.Printed, " & _
            "R.RunTime, R.RunDate, R.Operator, R.Flags, R.Units, " & _
            "R.SampleType, R.Analyser, D.Analyser As Anyl, R.Faxed, R.Authorised, " & _
            "R.Comment AS Comment, " & _
            "ISNULL(R.SignOff,0) as SignOff , ISNULL(R.SignOffBy,'') as SignOffBy, R.SignOffDateTime as SignOffDateTime " & _
            "FROM " & TableName & " R, " & Discipline & "TestDefinitions D, " & _
            Discipline & "DefIndex X " & _
            "WHERE " & _
            "R.DefIndex = X.Defindex " & _
            "AND SampleID = '" & SampleID & "' AND D.category = '" & Cat & "' " & _
            "AND R.Code = D.Code " & _
            "AND AgeFromDays <= " & DaysOld & " " & _
            "AND AgeToDays > " & DaysOld & " " & _
            "AND InUse = 1 "

300   If UCase$(HospName(0)) <> "CAVAN" Then
310       sql = sql & "AND R.SampleType = D.SampleType "
320   End If
330   If P = gNOTPRINTED And v = gNOTVALID Then
340       sql = sql & "AND Printed = 0 AND Valid = 0 "
350   ElseIf P = gNOTPRINTED And v = gVALID Then
360       sql = sql & "AND Printed = 0 AND Valid = 1 "
370   ElseIf P = gNOTPRINTED And v = gDONTCARE Then
380       sql = sql & "AND Printed = 0 "
390   ElseIf P = gPRINTED And v = gNOTVALID Then
400       sql = sql & "AND Printed = 1 AND Valid = 0 "
410   ElseIf P = gPRINTED And v = gVALID Then
420       sql = sql & "AND Printed = 1 AND Valid = 1 "
430   ElseIf P = gPRINTED And v = gDONTCARE Then
440       sql = sql & "AND Printed = 1 "
450   ElseIf P = gDONTCARE And v = gNOTVALID Then
460       sql = sql & "AND Valid = 0 "
470   ElseIf P = gDONTCARE And v = gVALID Then
480       sql = sql & "AND Valid = 1 "
490   End If

500   sql = sql & "ORDER BY PrintPriority ASC"
510   Set tb = New Recordset
'MsgBox sql
'frmViewResultsWE.Text1.Text = sql
520   RecOpenServer Connection, tb, sql    '  RecOpenClient 0,tb, Sql
530   Do While Not tb.EOF
540       Set Br = New BIEResult
550       With Br
560           .SampleID = ConvertNull(tb!SampleID, "") & ""
570           .Code = ConvertNull(tb!Code, "") & ""
580           .Result = ConvertNull(tb!Result, "") & ""
590           .Operator = ConvertNull(tb!Operator, "") & ""
600           .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
610           .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm:ss")
620           .Units = ConvertNull(tb!Units, "") & ""
630           .Printed = ConvertNull(tb!Printed, "0")
640           .Valid = IIf(tb!Valid, True, False)
650           If Not IsNull(tb!Authorised) Then
660               .Authorised = IIf(tb!Authorised, True, False)
670           Else
680               .Authorised = False
690           End If
              '    .Faxed = tb!Faxed
              '    If TB!h & "" <> "" Then .h = TB!h Else .h = False
              '    If TB!g & "" <> "" Then .g = TB!g Else .g = False
              '    If TB!o & "" <> "" Then .o = TB!o Else .o = False
              '    If TB!J & "" <> "" Then .J = TB!h Else .J = False
              '    If TB!l & "" <> "" Then .l = TB!h Else .l = False
              '    If TB!s & "" <> "" Then .s = TB!h Else .s = False
700           .Flags = ConvertNull(tb!Flags, "") & ""
710           .SampleType = ConvertNull(tb!SampleType, "") & ""
720           .Printformat = IIf(IsNull(tb!DP), 0, tb!DP)
730           .ShortName = ConvertNull(tb!ShortName, "") & ""
740           .LongName = ConvertNull(tb!LongName, "") & ""
750           If ConvertNull(tb!DoDelta, "") & "" <> "" Then .DoDelta = ConvertNull(tb!DoDelta, "") Else .DoDelta = False
760           .DeltaLimit = IIf(IsNull(tb!DeltaLimit), 9999, tb!DeltaLimit)
770           .Analyser = ConvertNull(tb!Analyser, "") & ""
            '+++ Junaid 15-02-2024
              .Anyl = ConvertNull(tb!Anyl, "") & ""
            '--- Junaid
780           .Comment = ConvertNull(tb!Comment, "") & ""
790           TempSQL = ""
800           If Fasting Then
810               .Fasting = True
820               If ConvertNull(tb!Code, "") = GetOptionSetting("BioCodeForGlucose", "0", "") Then
830                   TempSQL = "SELECT Fastinglow, FastingHigh FROM fastings WHERE testname = 'GLU'"
840               ElseIf ConvertNull(tb!Code, "") = GetOptionSetting("BioCodeForChol", "0", "") Then
850                   TempSQL = "SELECT Fastinglow, FastingHigh FROM fastings WHERE testname = 'CHO'"
860               ElseIf ConvertNull(tb!Code, "") = GetOptionSetting("BioCodeForTrig", "0", "") Then
870                   TempSQL = "SELECT Fastinglow, FastingHigh FROM fastings WHERE testname = 'TRI'"
880               Else
890                   .Low = ConvertNull(tb!Low, "")
900                   .High = ConvertNull(tb!High, "")
910               End If
920               If Len(TempSQL) > 0 Then
930                   Set TempTb = New Recordset
940                   RecOpenServer Connection, TempTb, TempSQL
950                   .Low = TempTb!FastingLow
960                   .High = TempTb!FastingHigh
970               End If

980           Else
990               .Fasting = False
1000              .Low = tb!Low
1010              .High = tb!High
1020          End If
1030          .FlagLow = tb!FlagLow
1040          .FlagHigh = tb!FlagHigh
1050          .PlausibleLow = tb!PlausibleLow
1060          .PlausibleHigh = tb!PlausibleHigh
1070          .DefIndex = tb!DefIndex

1080          .SignOff = tb!SignOff
1090          .SignOffBy = tb!SignOffBy
1100          If Not IsNull(tb!SignOffDateTime) Then
1110              .SignOffDateTime = tb!SignOffDateTime
1120          End If

1130          If .DefIndex = 0 Then
1140              If DoB <> "" And Sex <> "" Then
1150                  sql = "SELECT " & _
                            SelectNormalRange & SelectFlagRange & _
                            "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
                            "COALESCE(PlausibleHigh, 9999) PlausibleHigh " & _
                            "FROM " & Discipline & "TestDefinitions  " & _
                            "WHERE category = '" & Cat & "' " & _
                            "AND Code = '" & .Code & "' " & _
                            "AND AgeFromDays <= " & DaysOld & " " & _
                            "AND AgeToDays >= " & DaysOld & " "
1160                  Set tbRange = New Recordset
1170                  RecOpenServer 0, tbRange, sql
1180                  If Not tbRange.EOF Then
1190                      .Low = tbRange!Low
1200                      .High = tbRange!High
1210                      .FlagLow = tbRange!FlagLow
1220                      .FlagHigh = tbRange!FlagHigh
1230                      .PlausibleLow = tbRange!PlausibleLow
1240                      .PlausibleHigh = tbRange!PlausibleHigh

1250                      sql = "SELECT * FROM " & Discipline & "DefIndex " & _
                                "WHERE NormalLow = '" & .Low & "' " & _
                                "AND NormalHigh = '" & .High & "' " & _
                                "AND FlagLow = '" & .FlagLow & "' " & _
                                "AND FlagHigh = '" & .FlagHigh & "' " & _
                                "AND PlausibleLow = '" & .PlausibleLow & "' " & _
                                "AND PlausibleHigh = '" & .PlausibleHigh & "' "

1260                      Set tbNewIDX = New Recordset
1270                      RecOpenClient 0, tbNewIDX, sql
1280                      If Not tbNewIDX.EOF Then
1290                          .DefIndex = tbNewIDX!DefIndex
1300                      Else

1310                          sql = "INSERT INTO " & Discipline & "DefIndex " & _
                                    "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
                                    "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
                                    "VALUES ( " & _
                                    .Low & ", " & .High & ", " & .FlagLow & ", " & .FlagHigh & ", " & _
                                    .PlausibleLow & ", " & .PlausibleHigh & ", 0,9999) "
1320                          Cnxn(0).Execute sql

1330                          sql = "SELECT MAX(DefIndex) NewIndex FROM " & Discipline & "DefIndex"
1340                          Set tbNewIDX = New Recordset
1350                          RecOpenClient 0, tbNewIDX, sql
1360                          .DefIndex = tbNewIDX!NewIndex

1370                      End If

1380                      sql = "UPDATE " & TableName & " " & _
                                "SET DefIndex = '" & .DefIndex & "' " & _
                                "WHERE SampleID = '" & .SampleID & "' " & _
                                "AND Code = '" & .Code & "'"
1390                      Cnxn(0).Execute sql

1400                  End If
1410              End If
1420          End If


1430          BRs.Add Br
1440      End With
1450      tb.MoveNext
1460  Loop

1470  If BRs.Count <> 0 Then
1480      Set Load = BRs
1490  Else
1500      Set Load = Nothing
1510  End If
1520  Set Br = Nothing
1530  Set BRs = Nothing

1540  Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer

1550  intEL = Erl
1560  strES = Err.Description
1570  LogError "BIEResults", "Load", intEL, strES, sql & " " & TempSQL

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolBIEResults.[_NewEnum]

End Function



Public Sub RemoveItem(ByVal x As Long)

10    mcolBIEResults.Remove x

End Sub


Public Sub Save(ByVal Discipline As String, _
                ByVal BRs As BIEResults)
      'Discipline is either "Bio", "Imm" or "End"

      Dim tb As Recordset
      Dim sql As String
      Dim Br As BIEResult
      Dim lngSampleID As Long

10    On Error GoTo Save_Error

20    For Each Br In BRs
30      With Br
40        lngSampleID = Format$(Val(.SampleID))
50        sql = "Select * from " & Discipline & "Results where " & _
                "SampleID = '" & lngSampleID & "' " & _
                "and Code = '" & Br.Code & "'"
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
270       tb.Update

280       sql = "Delete from " & Discipline & "Requests where " & _
                "SampleID = '" & lngSampleID & "' " & _
                "and Code = '" & Br.Code & "'"
290       Cnxn(0).Execute sql
    
300     End With
310   Next

320   Set Br = Nothing

330   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "BIEResults", "Save", intEL, strES, sql


End Sub



