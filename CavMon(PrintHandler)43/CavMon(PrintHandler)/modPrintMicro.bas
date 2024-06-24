Attribute VB_Name = "modPrintMicro"
Option Explicit

Public Type OrgGroup
    OrgGroup As String
    OrgName As String
    ShortName As String
    ReportName As String
    Qualifier As String
End Type

Public Type ABResult
    AntibioticName As String
    AntibioticCode As String
    Result(1 To 8) As String
    Report(1 To 8) As Boolean
    RSI(1 To 8) As String
    CPO(1 To 8) As String
End Type

Public ValidatedBy As String

Private Function CountLines(ByVal strIP As String) As Integer

10    ReDim Comments(1 To MicroCommentLineCount) As String
      Dim n As Integer

20    FillCommentLines strIP, MicroCommentLineCount, Comments()

30    For n = MicroCommentLineCount To 1 Step -1
40        If Trim$(Comments(n)) <> "" Then
50            CountLines = n
60            Exit For
70        End If
80    Next

End Function

Public Function GetMicroscopyLineCount(ByVal SampleID As String) As Integer

      Dim URS As New UrineResults
      Dim UR As UrineResult
      Dim retval As Integer

10    On Error GoTo GetMicroscopyLineCount_Error

20    retval = 0

URS.LoadSedimax Val(SampleID)
URS.CheckForAllResults

Set URS = New UrineResults
30    URS.Load Val(SampleID) + sysOptMicroOffset(0)
40    If URS.Count > 0 Then
50      Set UR = URS("Bacteria")
60      If Not UR Is Nothing Then
70        retval = 1
80      Else
90        Set UR = URS("Crystals")
100       If Not UR Is Nothing Then
110         retval = 1
120       End If
130     End If

140     Set UR = URS("WCC")
150     If Not UR Is Nothing Then
160       retval = retval + 1
170     Else
180       Set UR = URS("Casts")
190       If Not UR Is Nothing Then
200         retval = retval + 1
210       End If
220     End If
  
230     Set UR = URS("RCC")
240     If Not UR Is Nothing Then
250       retval = retval + 1
260     Else
270       Set UR = URS("Misc0")
280       If Not UR Is Nothing Then
290         retval = retval + 1
300       End If
310     End If

320   End If

330   GetMicroscopyLineCount = retval

340   Exit Function

GetMicroscopyLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "modPrintMicro", "GetMicroscopyLineCount", intEL, strES

End Function

Public Function GetPregnancyLineCount(ByVal SampleID As String) As Integer

      Dim URS As New UrineResults
      Dim UR As UrineResult
      Dim retval As Integer

10    On Error GoTo GetPregnancyLineCount_Error

20    retval = 0

30    URS.Load Val(SampleID) + sysOptMicroOffset(0)
40    If URS.Count > 0 Then
50      Set UR = URS("Pregnancy")
60      If Not UR Is Nothing Then
70        retval = 1
80      End If
90    End If

100     Set UR = URS("HCGLevel")
110     If Not UR Is Nothing Then
120       retval = retval + 1
130     End If
  
140   GetPregnancyLineCount = retval

150   Exit Function

GetPregnancyLineCount_Error:

Dim strES As String
Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modPrintMicro", "GetPregnancyLineCount", intEL, strES

End Function

Public Function GetCSFCount(ByVal SampleID As String) As Integer

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetCSFCount_Error

20    GetCSFCount = 0
30    sql = "Select SampleID from CSFResults where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then
70        GetCSFCount = 8
80    End If

90    Exit Function

GetCSFCount_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modPrintMicro", "GetCSFCount", intEL, strES, sql

End Function
Public Function GetIsolateCount(ByVal SampleID As String) As Integer

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetIsolateCount_Error

20    sql = "Select Count(DISTINCT IsolateNumber) as tot from Isolates where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    GetIsolateCount = tb!Tot

60    Exit Function

GetIsolateCount_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modPrintMicro", "GetIsolateCount", intEL, strES, sql


End Function
Public Function GetABCount(ByVal SampleID As String, _
                            ByVal OrgNumbers As String) _
                            As Integer

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

10    On Error GoTo GetABCount_Error

20    sql = "Select Count(distinct AntibioticCode) as tot from Sensitivities where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "' and ("
30    For n = 1 To Len(OrgNumbers)
40        sql = sql & "IsolateNumber = '" & Mid$(OrgNumbers, n, 1) & "' or "
50    Next
60    sql = Left$(sql, Len(sql) - 3) & ") " & _
            "and Report = 1"

70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    GetABCount = tb!Tot

100   Exit Function

GetABCount_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modPrintMicro", "GetABCount", intEL, strES, sql

End Function


Public Function GetCommentLineCount(ByVal SampleID As String) As Integer

      Dim n As Integer
      Dim OBs As Observations
      Dim OB As Observation

10    On Error GoTo GetCommentLineCount_Error

20    GetCommentLineCount = 0
30    n = 0

40    Set OBs = New Observations
50    Set OBs = OBs.Load(Val(SampleID) + sysOptMicroOffset(0), "Demographic", "MicroCS", "MicroConsultant", "MicroGeneral", "MicroCSAutoComment")
60    If Not OBs Is Nothing Then
70      For Each OB In OBs
80        n = n + CountLines(OB.Comment)
90      Next
100   End If

110   GetCommentLineCount = n

120   Exit Function

GetCommentLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "modPrintMicro", "GetCommentLineCount", intEL, strES

End Function

Public Function GetMiscLineCount(ByVal SampleID As String) As Long

      'FOB+CDiff+Rota/Adeno+OP

      Dim intCount As Integer
      Dim Fxs As New FaecesResults
      Dim Gxs As New GenericResults

10    On Error GoTo GetMiscLineCount_Error

20    Fxs.Load Val(SampleID) + sysOptMicroOffset(0)
30    intCount = Fxs.Count

40    Gxs.Load Val(SampleID) + sysOptMicroOffset(0)
50    If Not Gxs("RSV") Is Nothing Then
60      intCount = intCount + 1
70    End If
80    If Not Gxs("RedSub") Is Nothing Then
90      intCount = intCount + 1
100   End If
110   If Not Gxs("cDiffPCR") Is Nothing Then
120     intCount = intCount + 1
130   End If

140   GetMiscLineCount = intCount

150   Exit Function

GetMiscLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modPrintMicro", "GetMiscLineCount", intEL, strES

End Function

Private Function IsForcedTo(ByVal TrueOrFalse As String, _
                            ByVal ABName As String, _
                            ByVal SID As Variant, _
                            ByVal index As Integer) _
                            As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo IsForcedTo_Error

20    sql = "Select * from ForcedABReport where " & _
            "SampleID = " & SID & " " & _
            "and ABName = '" & Trim$(ABName) & "' " & _
            "and Report = '" & IIf(TrueOrFalse = "Yes", "1", "0") & "' " & _
            "and [Index] = " & index
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    IsForcedTo = Not tb.EOF

60    Exit Function

IsForcedTo_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modPrintMicro", "IsForcedTo", intEL, strES, sql


End Function

Public Function IsNegativeResults(ByVal SampleID As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo IsNegativeResults_Error

20    IsNegativeResults = False

30    sql = "Select OrganismGroup from Isolates where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then

70        If UCase$(tb!OrganismGroup & "") = "_NO GROWTH_" Or _
             UCase$(tb!OrganismGroup & "") = "_NEGATIVE RESULTS_" Then

80            IsNegativeResults = True

90        End If

100   End If

110   Exit Function

IsNegativeResults_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modPrintMicro", "IsNegativeResults", intEL, strES, sql

End Function

Public Sub LoadResultArray(ByVal SampleIDWithOffset As Variant, _
                           ByRef ResultArray() As ABResult)

      Dim tb As Recordset
      Dim tbR As Recordset
      Dim sql As String
      Dim U As Integer
      Dim ReportThis As Boolean
      Dim NewABAdded As Boolean
      Dim IsolateNumber As Integer

10    On Error GoTo LoadResultArray_Error

20    sql = "Select Code, AntibioticName, MAX(ListOrder) AS M from Antibiotics where " & _
            "Code in ( " & _
            "         Select distinct AntibioticCode from Sensitivities where " & _
            "         SampleID = '" & SampleIDWithOffset & "' and Report = 1 " & _
            "        ) " & _
            "GROUP BY Code, AntiBioticName Order by M"

30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60        Debug.Print tb!AntibioticName
70        sql = "Select * from Sensitivities where " & _
                "AntibioticCode = '" & tb!Code & "' " & _
                "and SampleID = " & SampleIDWithOffset
80        Set tbR = New Recordset
90        RecOpenServer 0, tbR, sql
100       NewABAdded = False
110       Do While Not tbR.EOF
120           ReportThis = False
130           If Not IsForcedTo("No", tb!AntibioticName, SampleIDWithOffset, tbR!IsolateNumber) Then
140               ReportThis = True
150           End If
              '    Else
              '      If IsForcedTo("Yes", tb!AntibioticName, SampleIDWithOffset, tbR!IsolateNumber) Then
              '        ReportThis = True
              '      End If
              '    End If
160           If ReportThis Then
170               If Not NewABAdded Then
180                   U = UBound(ResultArray) + 1
190                   ReDim Preserve ResultArray(0 To U)
200                   ResultArray(U).AntibioticCode = tb!Code
210                   ResultArray(U).AntibioticName = Trim$(tb!AntibioticName)
220                   NewABAdded = True
230               End If
240               IsolateNumber = tbR!IsolateNumber
250               ResultArray(U).RSI(IsolateNumber) = tbR!RSI & ""
260           End If
270           tbR.MoveNext
280       Loop
290       tb.MoveNext
300   Loop

310   Exit Sub

LoadResultArray_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "modPrintMicro", "LoadResultArray", intEL, strES, sql

End Sub


Public Function GetPDefault(ByVal SampleID As String) As Integer

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetPDefault_Error

20    sql = "SELECT L.[Default] " & _
            "FROM Lists L JOIN SiteDetails50 M " & _
            "ON L.Text = M.Site " & _
            "WHERE M.SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "' " & _
            "AND L.ListType = 'SI' "
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    GetPDefault = 3
60    If Not tb.EOF Then
70        GetPDefault = Val(tb!Default)
80    End If

90    Exit Function

GetPDefault_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modPrintMicro", "GetPDefault", intEL, strES, sql

End Function
Public Function FillOrgGroups(ByRef strGroup() As OrgGroup, _
                              ByVal SampleIDWithOffset As Variant) _
                              As Integer

      Dim tb As Recordset
      Dim tbO As Recordset
      Dim sql As String
      Dim n As Integer
      Dim IsoNum As Integer

10    On Error GoTo FillOrgGroups_Error

20    sql = "Select OrganismGroup, OrganismName, Qualifier, IsolateNumber " & _
            "from Isolates where " & _
            "SampleID = '" & SampleIDWithOffset & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    n = 1
60    Do While Not tb.EOF
70        IsoNum = tb!IsolateNumber
80        With strGroup(IsoNum)
90            .OrgGroup = tb!OrganismGroup & ""
100           .OrgName = tb!OrganismName & ""
110           .Qualifier = tb!Qualifier & ""
120           sql = "Select ShortName, ReportName from Organisms where " & _
                    "Name = '" & tb!OrganismName & "' AND GroupName = '" & tb!OrganismGroup & "'"
130           Set tbO = New Recordset
140           RecOpenClient 0, tbO, sql
150           If Not tbO.EOF Then
160               .ShortName = tbO!ShortName & ""
170               .ReportName = Trim$(tbO!ReportName & "")
180           Else
190               .ShortName = Trim$(tb!OrganismName & "")
200               .ReportName = Trim$(tb!OrganismName & "")
210           End If
220           If .ReportName = "" Then
230               .ReportName = .OrgName
240           End If
250       End With
260       n = n + 1
270       tb.MoveNext
280   Loop

290   FillOrgGroups = n - 1

300   Exit Function

FillOrgGroups_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "modPrintMicro", "FillOrgGroups", intEL, strES, sql

End Function


Public Sub UpdatePrintValid(ByVal SampleID As Variant, _
                            ByVal Dept As String, _
                            ByVal LogAsValid As Boolean, _
                            ByVal LogAsPrinted As Boolean)

      Dim tb As Recordset
      Dim sql As String
      Dim NewValue As Long

10    On Error GoTo UpdatePrintValid_Error

20    Select Case UCase$(Dept)
      Case "REDSUB": NewValue = 1
30    Case "RSV": NewValue = 2
40    Case "OP": NewValue = 4
50    Case "CDIFF": NewValue = 8
60    Case "ROTAADENO": NewValue = 16
70    Case "FOB": NewValue = 32
80    Case "URINE": NewValue = 64
90    Case "CANDS": NewValue = 128
100   Case "CSF": NewValue = 256
110   End Select

120   sql = "IF EXISTS(SELECT * FROM PrintValid WHERE " & _
            "          SampleID = '" & SampleID & "') " & _
            "  UPDATE PrintValid "
130   If LogAsValid Or LogAsPrinted Then
140     If LogAsValid And LogAsPrinted Then
150       sql = sql & "  SET V = CONVERT(int, V) | " & NewValue & ", P = CONVERT(int, P) | " & NewValue & " "
160     Else
170       If LogAsValid Then
180         sql = sql & "  SET V = CONVERT(int, V) | " & NewValue & " "
190       End If
200       If LogAsPrinted Then
210         sql = sql & "  SET P = CONVERT(int, P) | " & NewValue & " "
220       End If
230     End If
240   Else
250     sql = sql & "SET V = 0, P = 0 "
260   End If
270   sql = sql & "  WHERE SampleID = '" & SampleID & "' " & _
                  "ELSE " & _
                  "  INSERT INTO PrintValid " & _
                  "  (SampleID, P, V) VALUES " & _
                  "  ('" & SampleID & "', " & _
                  "  " & IIf(LogAsPrinted, NewValue, 0) & ", " & _
                  "  " & IIf(LogAsValid, NewValue, 0) & ") "
280   Cnxn(0).Execute sql
290   If UCase(RP.PrintAction) <> "SAVE" Then
300       UpdatePrintValidLog SampleID, "MICRO"
310   End If

320   Exit Sub

UpdatePrintValid_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "modPrintMicro", "UpdatePrintValid", intEL, strES, sql

End Sub


Public Sub UpdatePrintValidLog(ByVal SampleID As Variant, _
                               ByVal Dept As String)

      Dim tb As Recordset
      Dim sql As String
      Dim LogDept As String

      'B Biochemistry
      'C Coagulation
      'E Endocrinology
      'H Haematology
      'I Immunology
      'M Micro
      'S ESR
      'X External

10    On Error GoTo UpdatePrintValidLog_Error

20    Select Case UCase$(Dept)
      Case "MICRO": LogDept = "M"
          '  Case "RSV":       LogDept = "V"
          '  Case "OP":        LogDept = "O"
          '  Case "CDIFF":     LogDept = "G"
          '  Case "ROTAADENO": LogDept = "A"
          '  Case "FOB":       LogDept = "F"
          '  Case "URINE":     LogDept = "U"
          '  Case "CANDS":     LogDept = "D"
30    End Select

40    sql = "SELECT * FROM PrintValidLog WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Department = '" & LogDept & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If tb.EOF Then
80        tb.AddNew
90    Else
100       ValidatedBy = tb!ValidatedBy & ""
110       sql = "INSERT INTO PrintValidLogArc " & _
                "  SELECT PrintValidLog.*, " & _
                "  'PrintHandler', " & _
                "  '" & Format$(Now, "dd/MMM/yyyy hh:mm:ss") & "' " & _
                "  FROM PrintValidLog WHERE " & _
                "  SampleID = '" & SampleID & "' " & _
                "  AND Department = '" & LogDept & "' "
120       Cnxn(0).Execute sql
130   End If
140   tb!SampleID = SampleID
150   tb!Department = LogDept
160   tb!Printed = 1
170   tb!Valid = 1
180   tb!PrintedBy = RP.Initiator
190   tb!PrintedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")


      'tb!ValidatedBy = ValidatedBy
      '
      'If Not IsNull(tb!ValidatedDateTime) Then
      '  If Not IsDate(tb!ValidatedDateTime) Then
      '    tb!ValidatedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
      '  End If
      'Else
      '  tb!ValidatedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
      'End If
200   tb.Update

210   Exit Sub

UpdatePrintValidLog_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "modPrintMicro", "UpdatePrintValidLog", intEL, strES, sql

End Sub

