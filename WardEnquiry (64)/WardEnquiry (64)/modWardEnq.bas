Attribute VB_Name = "modWardEnq"
Option Explicit

Public Cnxn() As Connection
Public CnxnBB() As Connection

Public HospName() As String

Public Const gVALID = 1
Public Const gNOTVALID = 2
Public Const gPRINTED = 1
Public Const gNOTPRINTED = 2
Public Const gDONTCARE = 0

Public Const MaxAgeToDays As Long = 43830

Public UserName As String
Public UserCode As String
Public UserPass As String

Public UserMemberOf As String
Public UserRoleName As String

Public LogOffDelaySecs As Long
Public LogOffDelayMin As Long

Public LogOffNow As Boolean

Public WardEnqForcedPrinter As String
Public RunningInArea As String

Public UserCanPrint As Boolean

Public intOtherHospitalsInGroup As Long

Public Type udtChartDoBName
  Chart As String
  DoB As String
  Name As String
  Hospital As String
  Cn As Integer
End Type
Public Type BTChartDoBName
  Chart As String
  DoB As String
  Name As String
  Addr As String
  Cn As Integer
End Type
Public Enum InputValidation
    NumericFullStopDash = 0
    Char = 1
    YorN = 2
    AlphaNumeric_NoApos = 3
    AlphaNumeric_AllowApos = 4
    Numeric_Only = 5
    AlphaOnly = 6
    NumericSlash = 7
    AlphaAndSpaceonly = 8
    CharNumericDashSlash = 9
    AlphaAndSpaceApos = 10
    DecimalNumericOnly = 11
    CharNumericDashSlashFullStop = 12
    ivSampleID = 13
    AlphaNumeric = 14
    AlphaNumericSpace = 15
    NumericDWMY = 16
    NumericDotLessGreater = 17
    AlphaAndSpaceAposDash = 18
End Enum

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public LiIcHas As New LIHs

Public Function ListTextFor(ByVal ListType As String, ByVal Code As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ListTextFor_Error

20    ListTextFor = ""
30    Code = UCase$(Trim$(Code))

40    sql = "SELECT * FROM Lists " & _
            "WHERE ListType = '" & ListType & "' " & _
            "AND Code = '" & AddTicks(Code) & "' " & _
            "AND InUse = 1"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80      ListTextFor = tb!Text & ""
90    End If

100   Exit Function

ListTextFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modLists", "ListTextFor", intEL, strES, sql

End Function

Public Sub AskUserQuestion(ByVal QuestionCode As String)

Dim tb As Recordset
Dim sql As String
Dim UserReply As String

On Error GoTo AskUserQuestion_Error

sql = "SELECT TOP 1 * FROM UserAcceptance WHERE QuestionCode = '" & QuestionCode & "' AND UserName = '" & UserName & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
    'ask user question
    If iMsg(ListTextFor("UserQuestions", "UQ1"), vbYesNo, "Please confirm") = vbYes Then
        UserReply = "Accepted"
    Else
        UserReply = "Rejected"
    End If
    
    sql = "INSERT INTO [dbo].[UserAcceptance] " & _
        "([UserCode], [UserName], [UserReply], [QuestionCode], [CreatedBy], [CreatedDateTime]) " & _
        " VALUES " & _
        "('" & AddTicks(UserCode) & "', '" & AddTicks(UserName) & "', '" & AddTicks(UserReply) & "', '" & AddTicks(QuestionCode) & "', '" & AddTicks("WardEnquiry") & "', '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "')"
    Cnxn(0).Execute sql
End If


Exit Sub
AskUserQuestion_Error:
   
LogError "modWardEnq", "AskUserQuestion", Erl, Err.Description, sql


End Sub

Function initial2upper(ByVal S As String) As String
    
      Dim n As Integer

10    S = Trim$(S & "")
20    If S = "" Then
30        initial2upper = ""
40        Exit Function
50    End If
  
60    If InStr(UCase$(S), "MAC") > 0 Or InStr(UCase$(S), "MC") > 0 Or InStr(S, "'") > 0 Then
70    S = LCase$(S)
80    S = UCase$(Left$(S, 1)) & Mid$(S, 2)

90    For n = 1 To Len(S) - 1
100       If Mid$(S, n, 1) = " " Or Mid$(S, n, 1) = "'" Then
110           S = Left$(S, n) & UCase$(Mid$(S, n + 1, 1)) & Mid$(S, n + 2)
120       End If
130       If n > 1 Then
140           If Mid$(S, n, 1) = "c" And Mid$(S, n - 1, 1) = "M" Then
150               S = Left$(S, n) & UCase$(Mid$(S, n + 1, 1)) & Mid$(S, n + 2)
160           End If
170       End If
180   Next
190   Else
200     S = StrConv(S, vbProperCase)
210   End If
220   initial2upper = S

End Function

Public Function Bar2Group(ByVal Group As String) As String

      Dim S As String

10       On Error GoTo Bar2Group_Error

20    Select Case Group
      Case "51": S = "O Pos"
30    Case "62": S = "A Pos"
40    Case "73": S = "B Pos"
50    Case "84": S = "AB Pos"
60    Case "95": S = "O Neg"
70    Case "06": S = "A Neg"
80    Case "17": S = "B Neg"
90    Case "28": S = "AB Neg"
100   Case Else: S = ""
110   End Select

120   Bar2Group = S

130      Exit Function

Bar2Group_Error:
Dim strES As String
Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "modWardEnq", "Bar2Group", intEL, strES

End Function
Public Function GetWBCValue(ByVal S As String) As String

      Dim v As Integer
      Dim RetVal As String

10    v = Val(S)
20    If v = 0 Then
30      RetVal = "Nil"
40    ElseIf v < 101 Then
50      RetVal = Str$(v)
60    Else
70      RetVal = ">100"
80    End If

90    GetWBCValue = RetVal

End Function

Public Function GetPlussesOrNil(ByVal S As String) As String

      Dim RetVal As String

10    RetVal = ""

20    If InStr(S, "-") > 0 Then
30      RetVal = "Nil"
40    Else
50      If InStr(S, "++++") > 0 Then
60        RetVal = "++++"
70      ElseIf InStr(S, "+++") > 0 Then
80        RetVal = "+++"
90      ElseIf InStr(S, "++") > 0 Then
100       RetVal = "++"
110     ElseIf InStr(S, "+") > 0 Then
120       RetVal = "+"
130     End If
140   End If

150   GetPlussesOrNil = RetVal

End Function

Public Function GetPlusses(ByVal S As String) As String

      Dim RetVal As String

10    RetVal = ""

20    If InStr(S, "++++") > 0 Then
30      RetVal = "++++"
40    ElseIf InStr(S, "+++") > 0 Then
50      RetVal = "+++"
60    ElseIf InStr(S, "++") > 0 Then
70      RetVal = "++"
80    ElseIf InStr(S, "+") > 0 Then
90      RetVal = "+"
100   End If

110   GetPlusses = RetVal

End Function


Public Function GetEGFRComment(ByVal SampleID As String, ByRef S() As String) As Boolean
      'Returns True if Comment Present

      Dim CodeForEGFR As String
      Dim BRs As New BIEResults
      Dim Br As BIEResult

10    GetEGFRComment = False

20    CodeForEGFR = UCase$(GetOptionSetting("BioCodeForEGFR", "5555", ""))

30    Set BRs = BRs.Load("Bio", SampleID, "Results", gDONTCARE, gDONTCARE)
40    If Not BRs Is Nothing Then
50      For Each Br In BRs
60        If Br.Code = CodeForEGFR Then
70          If Br.Valid Then
80            GetEGFRComment = True
90            ReDim S(0 To 11) As String
100           S(0) = "eGFR Interpretation:"
110           Select Case Val(Br.Result)
                Case Is >= 90:
120               S(1) = "CKD Stage 1"
130               S(2) = "eGFR >=90 Normal in the absence of other evidence of kidney damage."
    Case 60 To 89:
140               S(1) = "CKD Stage 2"
150               S(2) = "eGFR 60-89 Slight decrease in GFR. Not CKD in absence of other evidence of kidney damage."
    Case 45 To 59:
160               S(1) = "CKD Stage 3A"
170               S(2) = "eGFR 45-59 Moderate decrease in GFR with or without other evidence of kidney damage."
    Case 30 To 44:
180               S(1) = "CKD Stage 3B"
190               S(2) = "eGFR 30-44 Moderate decrease in GFR with or without other evidence of kidney damage."
200             Case 15 To 29
210               S(1) = "CKD Stage 4"
220               S(2) = "eGFR 15-29 Severe decrease in GFR with or without other evidence of kidney damage."
    Case Is < 15:
230               S(1) = "CKD Stage 5"
240               S(2) = "eGFR <15   Established renal failure."
250           End Select
260           S(3) = ""
270           S(4) = "The Laboratory uses the abbreviated four variable MDRD formula to derive the eGFR. The only"
280           S(5) = "correction the user need apply is multiply the result by 1.21 for patients of African origin."
290           S(6) = ""
300           S(7) = "Limitations of eGRF measurements:-"
310           S(8) = "eGFR is an estimate not a measurement and falls down in extremes, it is not useful in"
320           S(9) = "severely ill patients, those undergoing dialysis, in extremes of muscle mass or children."
330          S(10) = "It is subject to both the biological and analytical variability in creatinine measurement"
340           S(11) = ""
350         End If

360         Exit For
370       End If
380     Next
390   End If

End Function

Public Sub LogAsViewed(ByVal Discipline As String, _
                       ByVal SampleID As String, _
                       ByVal Chart As String)

Dim sql As String

'Discipline will be one of:
'A Results OverView
'B Bio Result
'C Coag Results
'D Bio History
'E Coag History
'F Haem History
'G Haem Graphs
'H Haem Cumulative
'I Bio/Imm/End Print
'J Haem Print
'K Coag Print
'L Log On
'M Manual Log Off
'N Micro Print
'O Auto Log Off
'P
'Q
'R Haem Result
'S
'T
'U
'V
'W
'X Close Program
'Y
'Z

On Error GoTo LogAsViewed_Error

sql = "INSERT INTO ViewedReports " & _
      "(Discipline, DateTime, Viewer, SampleID, Chart, Usercode) VALUES " & _
      "('" & Discipline & "', " & _
      "'" & Format(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
      "'" & AddTicks(UserName) & "', " & _
      "'" & SampleID & "', " & _
      "'" & Chart & "', " & _
      "'" & UserCode & "')"
Cnxn(0).Execute sql

Exit Sub

LogAsViewed_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modWardEnq", "LogAsViewed", intEL, strES, sql

End Sub

Public Function AreFlagsPresent(f() As Integer) As Boolean

      Dim n As Integer

10    AreFlagsPresent = False

20    For n = 0 To 5
30      If f(n) Then
40        AreFlagsPresent = True
50        Exit Function
60      End If
70    Next

End Function


Public Sub SplitLine(ByVal S As String, _
                     ByRef res() As String, _
                     ByVal MaxLen As Integer)

      Dim LineCounter As Integer
      Dim n As Integer
      Dim SpaceFound As Boolean

10    LineCounter = 1

20    Do While LineCounter <= UBound(res)
30      S = Trim$(S)
40      If Len(S) > MaxLen Then
50        SpaceFound = False
60        For n = MaxLen To 1 Step -1
70          If Mid$(S, n, 1) = " " Then
80            SpaceFound = True
90            res(LineCounter) = Trim$(Left$(S, n))
100           S = Trim$(Mid$(S, n))
110           LineCounter = LineCounter + 1
120           Exit For
130         End If
140       Next
150       If Not SpaceFound Then
160         res(LineCounter) = Left$(S, MaxLen)
170         S = Mid$(S, MaxLen + 1)
180         LineCounter = LineCounter + 1
190       End If
200     Else
210       res(LineCounter) = S
220       Exit Do
230     End If
240   Loop

End Sub


Function InterpH(ByVal Value As Single, _
                 ByVal Analyte As String, _
                 ByVal Sex As String, _
                 ByVal DoB As String, ByVal sampleDate As String) _
                 As String

      Dim sql As String
      Dim tb As Recordset
      Dim DaysOld As Long
      Dim SexSQL As String
      

10    Select Case Left$(UCase$(Sex), 1)
        Case "M"
20        SexSQL = "MaleLow as Low, MaleHigh as High "
30      Case "F"
40        SexSQL = "FemaleLow as Low, FemaleHigh as High "
50      Case Else
60        SexSQL = "FemaleLow as Low, MaleHigh as High "
70    End Select

80    If IsDate(DoB) Then

90      DaysOld = Abs(DateDiff("d", IIf(sampleDate <> "", sampleDate, Now), DoB))

100     sql = "Select top 1 PlausibleLow, PlausibleHigh, " & _
              SexSQL & _
              "from HaemTestDefinitions where " & _
              "AnalyteName = '" & Analyte & "' and AgeFromDays <= '" & DaysOld & "' " & _
              "and AgeToDays >= '" & DaysOld & "' " & _
              "order by AgeFromDays desc, AgeToDays asc"
110   Else
120     sql = "Select top 1 PlausibleLow, PlausibleHigh, " & _
              SexSQL & _
              "from HaemTestDefinitions where Analytename = '" & Analyte & "' " & _
              "and AgeFromDays = '0' " & _
              "and AgeToDays = '43830'"
130   End If

140   Set tb = New Recordset
150   RecOpenClient 0, tb, sql
160   If Not tb.EOF Then
  
170     If Value > tb!PlausibleHigh Then
180       InterpH = "X"
190       Exit Function
200     ElseIf Value < tb!PlausibleLow Then
210       InterpH = "X"
220       Exit Function
230     End If

240     If Value > tb!High Then
250       InterpH = "H"
260     ElseIf Value < tb!Low Then
270       InterpH = "L"
280     Else
290       InterpH = " "
300     End If
310   Else
320     InterpH = " "
330   End If

End Function
Public Function AddTicks(ByVal S As String) As String

10    S = Trim$(S)

20    S = Replace(S, "'", "''")

30    AddTicks = S

End Function

Public Function MaskInhibit(ByVal Br As BIEResult, ByVal BRs As BIEResults) As String

      Dim Lx As LIH
      Dim RetVal As String
      Dim Result As Single
      Dim BRLIH As BIEResult
      Dim CutOffForThisParameter As Single
      Dim LIHValue As Single


10    On Error GoTo MaskInhibit_Error

20    RetVal = ""

30    Set Lx = LiIcHas.Item("L", Br.Code, "P")
40    If Not Lx Is Nothing Then
50      Set BRLIH = BRs.Item("1071")
60      If Not BRLIH Is Nothing Then
70        CutOffForThisParameter = Lx.CutOff
80        If CutOffForThisParameter > 0 Then
90          LIHValue = BRLIH.Result
100         If LIHValue >= CutOffForThisParameter Then
110           RetVal = "XL"
120         End If
130       End If
140     End If
150   End If

160   If RetVal = "" Then
170     Set Lx = LiIcHas.Item("I", Br.Code, "P")
180     If Not Lx Is Nothing Then
190       Set BRLIH = BRs.Item("1072")
200       If Not BRLIH Is Nothing Then
210         CutOffForThisParameter = Lx.CutOff
220         If CutOffForThisParameter > 0 Then
230           LIHValue = BRLIH.Result
240           If LIHValue >= CutOffForThisParameter Then
250             RetVal = "XI"
260           End If
270         End If
280       End If
290     End If
300   End If

310   If RetVal = "" Then
320     Set Lx = LiIcHas.Item("H", Br.Code, "P")
330     If Not Lx Is Nothing Then
340       Set BRLIH = BRs.Item("1073")
350       If Not BRLIH Is Nothing Then
360         CutOffForThisParameter = Lx.CutOff
370         If CutOffForThisParameter > 0 Then
380           LIHValue = BRLIH.Result
390           If LIHValue >= CutOffForThisParameter Then
400             RetVal = "XH"
410           End If
420         End If
430       End If
440     End If
450   End If

460   MaskInhibit = RetVal

470   Exit Function

MaskInhibit_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "modWardEnq", "MaskInhibit", intEL, strES

End Function


'Public Function MaskInhibit(ByVal Br As BIEResult, ByVal BRs As BIEResults) As String
'
'      Dim Lx As LIH
'      Dim RetVal As String
'      Dim Result As Single
'      Dim BRLIH As BIEResult
'      Dim CutOffForThisParameter As Single
'      Dim LIHValue As Single
'
'10    On Error GoTo MaskInhibit_Error
'
'20    RetVal = ""
'
'30    Set Lx = LiIcHas.Item("L", Br.Code, "P")
'40    If Not Lx Is Nothing Then
'50      Set BRLIH = BRs.Item("1071")
'60      If Not BRLIH Is Nothing Then
'70        CutOffForThisParameter = Lx.CutOff
'80        If CutOffForThisParameter > 0 Then
'90          LIHValue = BRLIH.Result
'100         If LIHValue >= CutOffForThisParameter Then
'110           RetVal = "XL"
'120         End If
'130       End If
'140     End If
'150   End If
'
'160   If RetVal = "" Then
'170     Set Lx = LiIcHas.Item("I", Br.Code, "P")
'180     If Not Lx Is Nothing Then
'190       Set BRLIH = BRs.Item("1072")
'200       If Not BRLIH Is Nothing Then
'210         CutOffForThisParameter = Lx.CutOff
'220         If CutOffForThisParameter > 0 Then
'230           LIHValue = BRLIH.Result
'240           If LIHValue >= CutOffForThisParameter Then
'250             RetVal = "XI"
'260           End If
'270         End If
'280       End If
'290     End If
'300   End If
'
'310   If RetVal = "" Then
'320     Set Lx = LiIcHas.Item("H", Br.Code, "P")
'330     If Not Lx Is Nothing Then
'340       Set BRLIH = BRs.Item("1073")
'350       If Not BRLIH Is Nothing Then
'360         CutOffForThisParameter = Lx.CutOff
'370         If CutOffForThisParameter > 0 Then
'380           LIHValue = BRLIH.Result
'390           If LIHValue >= CutOffForThisParameter Then
'400             RetVal = "XH"
'410           End If
'420         End If
'430       End If
'440     End If
'450   End If
'
'460   MaskInhibit = RetVal
'
'470   Exit Function
'
'MaskInhibit_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'480   intEL = Erl
'490   strES = Err.Description
'500   LogError "modWardEnq", "MaskInhibit", intEL, strES
'
'End Function



Public Sub SingleUserUpdateLoggedOn(ByVal UserName As String)

      Dim sql As String

10    On Error GoTo SingleUserUpdateLoggedOn_Error

20    If Trim$(UserName) <> "" Then
30      sql = "IF EXISTS(SELECT * FROM WardEnqUsers WHERE " & _
              "          UserName = '" & AddTicks(UserName) & "') " & _
              "  UPDATE WardEnqUsers " & _
              "  SET DateTimeOfRecord = GETDATE() " & _
              "  WHERE UserName = '" & AddTicks(UserName) & "' " & _
              "ELSE " & _
              "  INSERT INTO WardEnqUsers " & _
              "  (UserName) VALUES ('" & AddTicks(UserName) & "')"
40      Cnxn(0).Execute sql
50    End If

60    Exit Sub

SingleUserUpdateLoggedOn_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modWardEnq", "SingleUserUpdateLoggedOn", intEL, strES, sql

End Sub


Public Function CheckAutoComments(ByVal SampleID As String, ByVal Index As Integer) As String

      Dim tb As Recordset
      Dim sql As String
      Dim ShortDisc As String
      Dim Discipline As String
      Dim RetVal As String

10    On Error GoTo CheckAutoComments_Error

20    RetVal = ""

30    If Index = 2 Then
40      Discipline = "Biochemistry"
50      ShortDisc = "Bio"
60    Else
70      Discipline = "Coagulation"
80      ShortDisc = "Coag"
90    End If

100   sql = "SELECT 'Output' = " & _
            "CASE WHEN ISNUMERIC(R.Result) = 1 AND R.Result <> '.' " & _
            "  THEN " & _
            "    CASE " & _
            "      WHEN Criteria = 'Present' THEN A.Comment " & _
            "      WHEN Criteria = 'Equal to' AND CONVERT(float, R.Result) = CONVERT(float, A.Value0) THEN A.Comment " & _
            "      WHEN Criteria = 'Less than' AND CONVERT(float, R.Result) < CONVERT(float, A.Value0) THEN A.Comment " & _
            "      WHEN Criteria = 'Greater than' AND CONVERT(float, R.Result) > CONVERT(float, A.Value0) THEN A.Comment " & _
            "      WHEN Criteria = 'Between' AND CONVERT(float, R.Result) > CONVERT(float, A.Value0) AND CONVERT(float, R.Result) < CONVERT(float, A.Value1) THEN A.Comment " & _
            "      WHEN Criteria = 'Not between' AND (CONVERT(float, R.Result) < CONVERT(float, A.Value0) OR CONVERT(float, R.Result) > CONVERT(float, A.Value1)) THEN A.Comment " & _
            "      ELSE '' " & _
            "    END " & _
            "  ELSE " & _
            "    CASE " & _
            "      WHEN Criteria = 'Contains Text' AND CHARINDEX( A.Value0, R.Result) > 0 THEN A.Comment " & _
            "      WHEN Criteria = 'Starts with' AND LEFT(R.Result, 1) = A.Value0 THEN A.Comment " & _
            "      ELSE '' " & _
            "    END " & _
            "END "
110   sql = sql & "FROM AutoComments A JOIN " & ShortDisc & "Results R ON " & _
            "R.Code = (SELECT TOP 1 Code FROM " & ShortDisc & "TestDefinitions " & _
            "          WHERE ShortName = A.Parameter " & _
            "          AND InUse = 1 ) " & _
            "WHERE A.Discipline = '" & Discipline & "' " & _
            "AND R.SampleID = '" & SampleID & "' "
            '& _
            '"AND A.Parameter = '" & ShortName & "'"

120   Set tb = New Recordset
130   RecOpenClient 0, tb, sql
140   Do While Not tb.EOF
150     If Trim$(tb!Output & "") <> "" Then
160       RetVal = RetVal & tb!Output & vbCrLf
170     End If
180   tb.MoveNext
190   Loop

200   CheckAutoComments = Trim$(RetVal)

210   Exit Function

CheckAutoComments_Error:

Dim strES As String
Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "modWardEnq", "CheckAutoComments", intEL, strES, sql

End Function
Public Function ForeNameFor(ByVal FullName As String) As String

      Dim n As Integer

10    ForeNameFor = ""

20    FullName = Trim$(FullName)
30    If FullName = "" Then Exit Function
40    n = InStr(FullName, " ")
50    If n = 0 Then
60      ForeNameFor = ""
70      Exit Function
80    End If

90    For n = Len(FullName) To 1 Step -1
100     If Mid$(FullName, n, 1) = " " Then
110       Exit For
120     End If
130   Next

140   ForeNameFor = Mid$(FullName, n + 1)

End Function

Public Sub LogError(ByVal ModuleName As String, _
                    ByVal ProcedureName As String, _
                    ByVal ErrorLineNumber As Integer, _
                    ByVal ErrorDescription As String, _
                    Optional ByVal SQLStatement As String, _
                    Optional ByVal EventDesc As String)

      Dim sql As String
      Dim MyMachineName As String
      Dim Vers As String
      Dim UID As String

10    On Error Resume Next

20    UID = AddTicks(UserName)

30    SQLStatement = AddTicks(SQLStatement)

40    ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "[MSSQL]")
50    ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver]", "[SQL]")
60    ErrorDescription = AddTicks(ErrorDescription)

70    Vers = App.Major & "-" & App.Minor & "-" & App.Revision

80    MyMachineName = vbGetComputerName()

90    sql = "IF NOT EXISTS " & _
      "    (SELECT * FROM ErrorLog WHERE " & _
      "     ModuleName = '" & ModuleName & "' " & _
      "     AND ProcedureName = '" & ProcedureName & "' " & _
      "     AND ErrorLineNumber = '" & ErrorLineNumber & "' " & _
      "     AND AppName = '" & App.EXEName & "' " & _
      "     AND AppVersion = '" & Vers & "' ) " & _
      "  INSERT INTO ErrorLog (" & _
      "    ModuleName, ProcedureName, ErrorLineNumber, SQLStatement, " & _
      "    ErrorDescription, UserName, MachineName, Eventdesc, AppName, AppVersion, EventCounter, eMailed) " & _
      "  VALUES  ('" & ModuleName & "', " & _
      "           '" & ProcedureName & "', " & _
      "           '" & ErrorLineNumber & "', " & _
      "           '" & SQLStatement & "', " & _
      "           '" & ErrorDescription & "', " & _
      "           '" & UID & "', " & _
      "           '" & MyMachineName & "', " & _
      "           '" & AddTicks(EventDesc) & "', " & _
      "           '" & App.EXEName & "', " & _
      "           '" & Vers & "', " & _
      "           '1', '0') " & _
      "ELSE "
100   sql = sql & "  UPDATE ErrorLog " & _
      "  SET SQLStatement = '" & SQLStatement & "', " & _
      "  ErrorDescription = '" & ErrorDescription & "', " & _
      "  MachineName = '" & MyMachineName & "', " & _
      "  DateTime = getdate(), " & _
      "  UserName = '" & UID & "', " & _
      "  EventCounter = COALESCE(EventCounter, 0) + 1 " & _
      "  WHERE ModuleName = '" & ModuleName & "' " & _
      "  AND ProcedureName = '" & ProcedureName & "' " & _
      "  AND ErrorLineNumber = '" & ErrorLineNumber & "' " & _
      "  AND AppName = '" & App.EXEName & "' " & _
      "  AND AppVersion = '" & Vers & "'"

110   Cnxn(0).Execute sql

End Sub

Public Function vbGetComputerName() As String
  
      'Gets the name of the machine
      Const MAXSIZE As Integer = 256
      Dim sTmp As String * MAXSIZE
      Dim lLen As Long
 
10    lLen = MAXSIZE - 1
20    If (GetComputerName(sTmp, lLen)) Then
30      vbGetComputerName = Left$(sTmp, lLen)
40    Else
50      vbGetComputerName = ""
60    End If

End Function

Public Function QuickInterpBio(ByVal Result As BIEResult) _
                               As String

10    With Result
20      If Val(.Result) < .Low Then
30        QuickInterpBio = "Low "
40      ElseIf Val(.Result) > .High Then
50        QuickInterpBio = "High"
60      Else
70        QuickInterpBio = "    "
80      End If
90    End With

End Function


Public Function GetMicroSiteDetails(SampleIDWithOffset As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetMicroSiteDetails_Error

20    sql = "Select * From MicroSiteDetails Where SampleID = '" & SampleIDWithOffset & "'"
30    Set tb = New Recordset

40    RecOpenClient 0, tb, sql

50    If Not tb.EOF Then
60        GetMicroSiteDetails = tb!Site & " " & tb!SiteDetails
70    Else
80        GetMicroSiteDetails = ""
90    End If


100   Exit Function

GetMicroSiteDetails_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modWardEnq", "GetMicroSiteDetails", intEL, strES, sql




End Function

Public Function VI(KeyAscii As Integer, _
                   iv As InputValidation, _
                   Optional NextFieldOnEnter As Boolean) As Integer

          Dim sTemp As String

10        sTemp = Chr$(KeyAscii)
20        If KeyAscii = 13 Then    'Enter Key
30            If NextFieldOnEnter = True Then
40                VI = 9    'Return Tab Keyascii if User Selected NextFieldOnEnter Option
50            Else
60                VI = 13
70            End If
80            Exit Function
90        ElseIf KeyAscii = 8 Then    'BackSpace
100           VI = 8
110           Exit Function
120       End If

          ' turn input to upper case

130       Select Case iv
          Case InputValidation.NumericFullStopDash:
140           Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", "-"
150               VI = Asc(sTemp)
160           Case Else
170               VI = 0
180           End Select

190       Case InputValidation.ivSampleID
200           Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
210               VI = Asc(sTemp)
220           Case "A" To "Z"
230               VI = Asc(sTemp)
240           Case "a" To "z"
250               VI = Asc(sTemp) - 32    'Convert to upper case
260           Case Else
270               VI = 0
280           End Select

290       Case InputValidation.NumericDotLessGreater
300           Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", ">", "<"
310               VI = Asc(sTemp)
320           Case Else
330               VI = 0
340           End Select

350       Case InputValidation.AlphaNumeric
360           Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
370               VI = Asc(sTemp)
380           Case "A" To "Z"
390               VI = Asc(sTemp)
400           Case "a" To "z"
410               VI = Asc(sTemp)
420           Case Else
430               VI = 0
440           End Select

450       Case InputValidation.AlphaNumericSpace
460           Select Case sTemp
              Case " ", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "<", ">"
470               VI = Asc(sTemp)
480           Case "A" To "Z"
490               VI = Asc(sTemp)
500           Case "a" To "z"
510               VI = Asc(sTemp)
520           Case Else
530               VI = 0
540           End Select

550       Case InputValidation.Char
560           Select Case sTemp
              Case " ", "-"
570               VI = Asc(sTemp)
580           Case "A" To "Z"
590               VI = Asc(sTemp)
600           Case "a" To "z"
610               VI = Asc(sTemp)
620           Case Else
630               VI = 0
640           End Select

650       Case InputValidation.YorN
660           sTemp = UCase(Chr$(KeyAscii))
670           Select Case sTemp
              Case "Y", "N"
680               VI = Asc(sTemp)
690           Case Else
700               VI = 0
710           End Select

720       Case InputValidation.AlphaNumeric_NoApos
730           Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", _
                   " ", "/", ";", ":", "\", "-", ">", "<", "(", ")", "@", _
                   "%", "!", """", "+", "^", "~", "`", "Ç", "´", "Ã", "Á", _
                   "Â", "È", "É", "Ê", "Ì", "Í", "Î", "Ò", "Ó", "Ô", "Õ", _
                   "Ù", "Ú", "Û", "Ü", "à", "á", "â", "ã", "ç", "è", "é", _
                   "ê", "ì", "í", "î", "ò", "ó", "ô", "õ", "ö", "ù", "ú", _
                   "û", "ü", "Æ", "æ", ",", "?"
740               VI = Asc(sTemp)
750           Case "A" To "Z"
760               VI = Asc(sTemp)
770           Case "a" To "z"
780               VI = Asc(sTemp)
790           Case Else
800               VI = 0
810           End Select

820       Case InputValidation.AlphaNumeric_AllowApos
830           Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", " ", "'"
840               VI = Asc(sTemp)
850           Case "A" To "Z"
860               VI = Asc(sTemp)
870           Case "a" To "z"
880               VI = Asc(sTemp)
890           Case Else
900               VI = 0
910           End Select

920       Case InputValidation.Numeric_Only
930           Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
940               VI = Asc(sTemp)
950           Case Else
960               VI = 0
970           End Select

980       Case InputValidation.AlphaOnly
990           Select Case sTemp
              Case "A" To "Z"
1000              VI = Asc(sTemp)
1010          Case "a" To "z"
1020              VI = Asc(sTemp)
1030          Case Else
1040              VI = 0
1050          End Select

1060      Case InputValidation.NumericSlash
1070          Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/"
1080              VI = Asc(sTemp)
1090          Case Else
1100              VI = 0
1110          End Select

1120      Case InputValidation.AlphaAndSpaceonly
1130          Select Case sTemp
              Case " "
1140              VI = Asc(sTemp)
1150          Case "A" To "Z"
1160              VI = Asc(sTemp)
1170          Case "a" To "z"
1180              VI = Asc(sTemp)
1190          Case Else
1200              VI = 0
1210          End Select

1220      Case InputValidation.CharNumericDashSlash
1230          Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/", "-"
1240              VI = Asc(sTemp)
1250          Case "A" To "Z"
1260              VI = Asc(sTemp)
1270          Case "a" To "z"
1280              VI = Asc(sTemp) - 32    'Convert to upper case
1290          Case Else
1300              VI = 0
1310          End Select

1320      Case InputValidation.AlphaAndSpaceApos
1330          Select Case sTemp
              Case " ", "'"
1340              VI = Asc(sTemp)
1350          Case "A" To "Z"
1360              VI = Asc(sTemp)
1370          Case "a" To "z"
1380              VI = Asc(sTemp)
1390          Case Else
1400              VI = 0
1410          End Select

1420      Case InputValidation.DecimalNumericOnly
1430          Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "."
1440              VI = Asc(sTemp)
1450          Case Else
1460              VI = 0
1470          End Select

1480      Case InputValidation.CharNumericDashSlashFullStop
1490          Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/", "-", "."
1500              VI = Asc(sTemp)
1510          Case "A" To "Z"
1520              VI = Asc(sTemp)
1530          Case "a" To "z"
1540              VI = Asc(sTemp)
1550          Case Else
1560              VI = 0
1570          End Select

1580      Case InputValidation.NumericDWMY
1590          Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "D", "M", "Y", "d", "m", "y", "w", "W", "s", "S"
1600              VI = Asc(sTemp)
1610          Case Else
1620              VI = 0
1630          End Select
          
1640      Case InputValidation.AlphaAndSpaceAposDash
1650          Select Case sTemp
              Case " ", "'", "-"
1660              VI = Asc(sTemp)
1670          Case "A" To "Z"
1680              VI = Asc(sTemp)
1690          Case "a" To "z"
1700              VI = Asc(sTemp)
1710          Case Else
1720              VI = 0
1730          End Select

1740      End Select

1750      If VI = 0 Then Beep

End Function

Public Function UserHasAuthority(ByVal MemberOf As String, SystemRole As String) As Boolean

10        On Error GoTo UserHasAuthority_Error

          
20        UserHasAuthority = False
30        Exit Function

UserHasAuthority_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "Shared", "UserHasAuthority", intEL, strES

End Function

Public Function FixComboWidth(Combo As ComboBox) As Boolean

    Dim i As Integer
    Dim ScrollWidth As Long

10  With Combo
20      For i = 0 To .ListCount
30          If .Parent.TextWidth(.List(i)) > ScrollWidth Then
40              ScrollWidth = .Parent.TextWidth(.List(i))
50          End If
60      Next i
70      FixComboWidth = SendMessage(.hwnd, CB_SETDROPPEDWIDTH, _
                                    ScrollWidth / 15 + 30, 0) > 0

80  End With

End Function

Public Function CheckForUnsignedMicro(ByVal chartNumber As String) As Boolean

          Dim tb As Recordset
          Dim sql As String
          
10        On Error GoTo CheckForUnsignedMicro_Error
          
20        sql = "Select sampleId From PrintValidLog Where SampleID in (select sampleid from demographics where Chart ='" & chartNumber & "') and (SignOff is NULL or SignOff = 0)"
30        Set tb = New Recordset

40        RecOpenClient 0, tb, sql

50        If Not tb.EOF Then
60            CheckForUnsignedMicro = True
70        Else
80            CheckForUnsignedMicro = False
90        End If
100       Exit Function

CheckForUnsignedMicro_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "modWardEnq", "CheckForUnsignedMicro", intEL, strES, sql

End Function


Public Function ListCodeFor(ByVal ListType As String, ByVal Text As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ListCodeFor_Error

20    ListCodeFor = ""
30    Text = UCase$(Trim$(Text))

40    sql = "SELECT * FROM Lists " & _
            "WHERE ListType = '" & ListType & "' " & _
            "AND Code = '" & AddTicks(Text) & "' " & _
            "AND InUse = 1"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80      ListCodeFor = tb!Code & ""
90    End If

100   Exit Function

ListCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modLists", "ListCodeFor", intEL, strES, sql

End Function

