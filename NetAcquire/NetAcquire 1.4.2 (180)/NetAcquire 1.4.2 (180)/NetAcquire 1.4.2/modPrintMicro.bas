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

5150  ReDim Comments(1 To MicroCommentLineCount) As String
      Dim n As Integer

5160  FillCommentLines strIP, MicroCommentLineCount, Comments()

5170  For n = MicroCommentLineCount To 1 Step -1
5180      If Trim$(Comments(n)) <> "" Then
5190          CountLines = n
5200          Exit For
5210      End If
5220  Next

End Function

Public Function GetMicroscopyLineCount(ByVal SampleID As String) As Integer

      Dim URS As New UrineResults
      Dim UR As UrineResult
      Dim RetVal As Integer

5230  On Error GoTo GetMicroscopyLineCount_Error

5240  RetVal = 0

5250  URS.LoadSedimax Val(SampleID)
5260  URS.CheckForAllResults

5270  Set URS = New UrineResults
5280  URS.Load Val(SampleID) ' + sysOptMicroOffset(0)
5290  If URS.Count > 0 Then
5300    Set UR = URS("Bacteria")
5310    If Not UR Is Nothing Then
5320      RetVal = 1
5330    Else
5340      Set UR = URS("Crystals")
5350      If Not UR Is Nothing Then
5360        RetVal = 1
5370      End If
5380    End If

5390    Set UR = URS("WCC")
5400    If Not UR Is Nothing Then
5410      RetVal = RetVal + 1
5420    Else
5430      Set UR = URS("Casts")
5440      If Not UR Is Nothing Then
5450        RetVal = RetVal + 1
5460      End If
5470    End If
        
5480    Set UR = URS("RCC")
5490    If Not UR Is Nothing Then
5500      RetVal = RetVal + 1
5510    Else
5520      Set UR = URS("Misc0")
5530      If Not UR Is Nothing Then
5540        RetVal = RetVal + 1
5550      End If
5560    End If

5570  End If

5580  GetMicroscopyLineCount = RetVal

5590  Exit Function

GetMicroscopyLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

5600  intEL = Erl
5610  strES = Err.Description
5620  LogError "modPrintMicro", "GetMicroscopyLineCount", intEL, strES

End Function

Public Function GetPregnancyLineCount(ByVal SampleID As String) As Integer

      Dim URS As New UrineResults
      Dim UR As UrineResult
      Dim RetVal As Integer

5630  On Error GoTo GetPregnancyLineCount_Error

5640  RetVal = 0

5650  URS.Load Val(SampleID) ' + sysOptMicroOffset(0)
5660  If URS.Count > 0 Then
5670    Set UR = URS("Pregnancy")
5680    If Not UR Is Nothing Then
5690      RetVal = 1
5700    End If
5710  End If

5720    Set UR = URS("HCGLevel")
5730    If Not UR Is Nothing Then
5740      RetVal = RetVal + 1
5750    End If
        
5760  GetPregnancyLineCount = RetVal

5770  Exit Function

GetPregnancyLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

5780  intEL = Erl
5790  strES = Err.Description
5800  LogError "modPrintMicro", "GetPregnancyLineCount", intEL, strES

End Function

Public Function GetCSFCount(ByVal SampleID As String) As Integer

      Dim sql As String
      Dim tb As Recordset

5810  On Error GoTo GetCSFCount_Error

5820  GetCSFCount = 0
      '30    sql = "Select SampleID from CSFResults where " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
5830  sql = "Select SampleID from CSFResults where " & _
            "SampleID = '" & Val(SampleID) & "'"
5840  Set tb = New Recordset
5850  RecOpenClient 0, tb, sql
5860  If Not tb.EOF Then
5870      GetCSFCount = 8
5880  End If

5890  Exit Function

GetCSFCount_Error:

      Dim strES As String
      Dim intEL As Integer

5900  intEL = Erl
5910  strES = Err.Description
5920  LogError "modPrintMicro", "GetCSFCount", intEL, strES, sql

End Function
Public Function GetIsolateCount(ByVal SampleID As String) As Integer

      Dim sql As String
      Dim tb As Recordset

5930  On Error GoTo GetIsolateCount_Error

      '20    sql = "Select Count(DISTINCT IsolateNumber) as tot from Isolates where " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
5940  sql = "Select Count(DISTINCT IsolateNumber) as tot from Isolates where " & _
            "SampleID = '" & Val(SampleID) & "'"
5950  Set tb = New Recordset
5960  RecOpenClient 0, tb, sql
5970  GetIsolateCount = tb!Tot

5980  Exit Function

GetIsolateCount_Error:

      Dim strES As String
      Dim intEL As Integer

5990  intEL = Erl
6000  strES = Err.Description
6010  LogError "modPrintMicro", "GetIsolateCount", intEL, strES, sql


End Function
Public Function GetABCount(ByVal SampleID As String, _
                            ByVal OrgNumbers As String) _
                            As Integer

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

6020  On Error GoTo GetABCount_Error

      '20    sql = "Select Count(distinct AntibioticCode) as tot from Sensitivities where " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "' and ("
6030  sql = "Select Count(distinct AntibioticCode) as tot from Sensitivities where " & _
            "SampleID = '" & Val(SampleID) & "' and ("
6040  For n = 1 To Len(OrgNumbers)
6050      sql = sql & "IsolateNumber = '" & Mid$(OrgNumbers, n, 1) & "' or "
6060  Next
6070  sql = Left$(sql, Len(sql) - 3) & ") " & _
            "and Report = 1"

6080  Set tb = New Recordset
6090  RecOpenClient 0, tb, sql
6100  GetABCount = tb!Tot

6110  Exit Function

GetABCount_Error:

      Dim strES As String
      Dim intEL As Integer

6120  intEL = Erl
6130  strES = Err.Description
6140  LogError "modPrintMicro", "GetABCount", intEL, strES, sql

End Function


Public Function GetCommentLineCount(ByVal SampleID As String) As Integer

      Dim n As Integer
      Dim OBs As Observations
      Dim OB As Observation

6150  On Error GoTo GetCommentLineCount_Error

6160  GetCommentLineCount = 0
6170  n = 0

6180  Set OBs = New Observations
      '50    Set OBs = OBs.Load(Val(SampleID) + sysOptMicroOffset(0), "Demographic", "MicroCS", "MicroConsultant", "MicroGeneral", "MicroCSAutoComment")
6190  Set OBs = OBs.Load(Val(SampleID), "Demographic", "MicroCS", "MicroConsultant", "MicroGeneral", "MicroCSAutoComment")
6200  If Not OBs Is Nothing Then
6210    For Each OB In OBs
6220      n = n + CountLines(OB.Comment)
6230    Next
6240  End If

6250  GetCommentLineCount = n

6260  Exit Function

GetCommentLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

6270  intEL = Erl
6280  strES = Err.Description
6290  LogError "modPrintMicro", "GetCommentLineCount", intEL, strES

End Function

Public Function GetMiscLineCount(ByVal SampleID As String) As Long

      'FOB+CDiff+Rota/Adeno+OP

      Dim intCount As Integer
      Dim Fxs As New FaecesResults
      Dim GXs As New GenericResults

6300  On Error GoTo GetMiscLineCount_Error

6310  Fxs.Load Val(SampleID) ' + sysOptMicroOffset(0)
6320  intCount = Fxs.Count

6330  GXs.Load Val(SampleID) ' + sysOptMicroOffset(0)
6340  If Not GXs("RSV") Is Nothing Then
6350    intCount = intCount + 1
6360  End If
6370  If Not GXs("RedSub") Is Nothing Then
6380    intCount = intCount + 1
6390  End If
6400  If Not GXs("cDiffPCR") Is Nothing Then
6410    intCount = intCount + 1
6420  End If

6430  GetMiscLineCount = intCount

6440  Exit Function

GetMiscLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

6450  intEL = Erl
6460  strES = Err.Description
6470  LogError "modPrintMicro", "GetMiscLineCount", intEL, strES

End Function

Private Function IsForcedTo(ByVal TrueOrFalse As String, _
                            ByVal ABName As String, _
                            ByVal SID As Variant, _
                            ByVal Index As Integer) _
                            As Boolean

      Dim tb As Recordset
      Dim sql As String

6480  On Error GoTo IsForcedTo_Error

6490  sql = "Select * from ForcedABReport where " & _
            "SampleID = " & SID & " " & _
            "and ABName = '" & Trim$(ABName) & "' " & _
            "and Report = '" & IIf(TrueOrFalse = "Yes", "1", "0") & "' " & _
            "and [Index] = " & Index
6500  Set tb = New Recordset
6510  RecOpenServer 0, tb, sql
6520  IsForcedTo = Not tb.EOF

6530  Exit Function

IsForcedTo_Error:

      Dim strES As String
      Dim intEL As Integer

6540  intEL = Erl
6550  strES = Err.Description
6560  LogError "modPrintMicro", "IsForcedTo", intEL, strES, sql


End Function

Public Function IsNegativeResults(ByVal SampleID As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

6570  On Error GoTo IsNegativeResults_Error

6580  IsNegativeResults = False

      '30    sql = "Select OrganismGroup from Isolates where " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
6590  sql = "Select OrganismGroup from Isolates where " & _
            "SampleID = '" & Val(SampleID) & "'"
6600  Set tb = New Recordset
6610  RecOpenClient 0, tb, sql
6620  If Not tb.EOF Then

6630      If UCase$(tb!OrganismGroup & "") = "NO GROWTH" Or _
             UCase$(tb!OrganismGroup & "") = "NEGATIVE RESULTS" Then

6640          IsNegativeResults = True

6650      End If

6660  End If

6670  Exit Function

IsNegativeResults_Error:

      Dim strES As String
      Dim intEL As Integer

6680  intEL = Erl
6690  strES = Err.Description
6700  LogError "modPrintMicro", "IsNegativeResults", intEL, strES, sql

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

6710  On Error GoTo LoadResultArray_Error

6720  sql = "Select Code, AntibioticName, MAX(ListOrder) AS M from Antibiotics where " & _
            "Code in ( " & _
            "         Select distinct AntibioticCode from Sensitivities where " & _
            "         SampleID = '" & SampleIDWithOffset & "' and Report = 1 " & _
            "        ) " & _
            "GROUP BY Code, AntiBioticName Order by M"

6730  Set tb = New Recordset
6740  RecOpenClient 0, tb, sql
6750  Do While Not tb.EOF
6760      Debug.Print tb!AntibioticName
6770      sql = "Select * from Sensitivities where " & _
                "AntibioticCode = '" & tb!Code & "' " & _
                "and SampleID = " & SampleIDWithOffset
6780      Set tbR = New Recordset
6790      RecOpenServer 0, tbR, sql
6800      NewABAdded = False
6810      Do While Not tbR.EOF
6820          ReportThis = False
6830          If Not IsForcedTo("No", tb!AntibioticName, SampleIDWithOffset, tbR!IsolateNumber) Then
6840              ReportThis = True
6850          End If
              '    Else
              '      If IsForcedTo("Yes", tb!AntibioticName, SampleIDWithOffset, tbR!IsolateNumber) Then
              '        ReportThis = True
              '      End If
              '    End If
6860          If ReportThis Then
6870              If Not NewABAdded Then
6880                  U = UBound(ResultArray) + 1
6890                  ReDim Preserve ResultArray(0 To U)
6900                  ResultArray(U).AntibioticCode = tb!Code
6910                  ResultArray(U).AntibioticName = Trim$(tb!AntibioticName)
6920                  NewABAdded = True
6930              End If
6940              IsolateNumber = tbR!IsolateNumber
6950              ResultArray(U).RSI(IsolateNumber) = tbR!RSI & ""
6960          End If
6970          tbR.MoveNext
6980      Loop
6990      tb.MoveNext
7000  Loop

7010  Exit Sub

LoadResultArray_Error:

      Dim strES As String
      Dim intEL As Integer

7020  intEL = Erl
7030  strES = Err.Description
7040  LogError "modPrintMicro", "LoadResultArray", intEL, strES, sql

End Sub


Public Function GetPDefault(ByVal SampleID As String) As Integer

      Dim tb As Recordset
      Dim sql As String

7050  On Error GoTo GetPDefault_Error

      '20    sql = "SELECT L.[Default] " & _
      '            "FROM Lists L JOIN SiteDetails50 M " & _
      '            "ON L.Text = M.Site " & _
      '            "WHERE M.SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "' " & _
      '            "AND L.ListType = 'SI' "
7060  sql = "SELECT L.[Default] " & _
            "FROM Lists L JOIN SiteDetails50 M " & _
            "ON L.Text = M.Site " & _
            "WHERE M.SampleID = '" & Val(SampleID) & "' " & _
            "AND L.ListType = 'SI' "
7070  Set tb = New Recordset
7080  RecOpenClient 0, tb, sql

7090  GetPDefault = 3
7100  If Not tb.EOF Then
7110      GetPDefault = Val(tb!Default)
7120  End If

7130  Exit Function

GetPDefault_Error:

      Dim strES As String
      Dim intEL As Integer

7140  intEL = Erl
7150  strES = Err.Description
7160  LogError "modPrintMicro", "GetPDefault", intEL, strES, sql

End Function
Public Function FillOrgGroups(ByRef strGroup() As OrgGroup, _
                              ByVal SampleIDWithOffset As Variant) _
                              As Integer

      Dim tb As Recordset
      Dim tbO As Recordset
      Dim sql As String
      Dim n As Integer
      Dim IsoNum As Integer

7170  On Error GoTo FillOrgGroups_Error

7180  sql = "Select OrganismGroup, OrganismName, Qualifier, IsolateNumber " & _
            "from Isolates where " & _
            "SampleID = '" & SampleIDWithOffset & "'"
7190  Set tb = New Recordset
7200  RecOpenClient 0, tb, sql
7210  n = 1
7220  Do While Not tb.EOF
7230      IsoNum = tb!IsolateNumber
7240      With strGroup(IsoNum)
7250          .OrgGroup = tb!OrganismGroup & ""
7260          .OrgName = tb!OrganismName & ""
7270          .Qualifier = tb!Qualifier & ""
7280          sql = "Select ShortName, ReportName from Organisms where " & _
                    "Name = '" & tb!OrganismName & "' AND GroupName = '" & tb!OrganismGroup & "'"
7290          Set tbO = New Recordset
7300          RecOpenClient 0, tbO, sql
7310          If Not tbO.EOF Then
7320              .ShortName = tbO!ShortName & ""
7330              .ReportName = Trim$(tbO!ReportName & "")
7340          Else
7350              .ShortName = Trim$(tb!OrganismName & "")
7360              .ReportName = Trim$(tb!OrganismName & "")
7370          End If
7380          If .ReportName = "" Then
7390              .ReportName = .OrgName
7400          End If
7410      End With
7420      n = n + 1
7430      tb.MoveNext
7440  Loop

7450  FillOrgGroups = n - 1

7460  Exit Function

FillOrgGroups_Error:

      Dim strES As String
      Dim intEL As Integer

7470  intEL = Erl
7480  strES = Err.Description
7490  LogError "modPrintMicro", "FillOrgGroups", intEL, strES, sql

End Function


'Public Sub UpdatePrintValid(ByVal SampleID As Variant, _
'                            ByVal Dept As String, _
'                            ByVal LogAsValid As Boolean, _
'                            ByVal LogAsPrinted As Boolean)
'
'      Dim tb As Recordset
'      Dim sql As String
'      Dim NewValue As Long
'
'10    On Error GoTo UpdatePrintValid_Error
'
'20    Select Case UCase$(Dept)
'      Case "REDSUB": NewValue = 1
'30    Case "RSV": NewValue = 2
'40    Case "OP": NewValue = 4
'50    Case "CDIFF": NewValue = 8
'60    Case "ROTAADENO": NewValue = 16
'70    Case "FOB": NewValue = 32
'80    Case "URINE": NewValue = 64
'90    Case "CANDS": NewValue = 128
'100   Case "CSF": NewValue = 256
'110   End Select
'
'120   sql = "IF EXISTS(SELECT * FROM PrintValid WHERE " & _
'            "          SampleID = '" & SampleID & "') " & _
'            "  UPDATE PrintValid "
'130   If LogAsValid Or LogAsPrinted Then
'140     If LogAsValid And LogAsPrinted Then
'150       sql = sql & "  SET V = CONVERT(int, V) | " & NewValue & ", P = CONVERT(int, P) | " & NewValue & " "
'160     Else
'170       If LogAsValid Then
'180         sql = sql & "  SET V = CONVERT(int, V) | " & NewValue & " "
'190       End If
'200       If LogAsPrinted Then
'210         sql = sql & "  SET P = CONVERT(int, P) | " & NewValue & " "
'220       End If
'230     End If
'240   Else
'250     sql = sql & "SET V = 0, P = 0 "
'260   End If
'270   sql = sql & "  WHERE SampleID = '" & SampleID & "' " & _
'                  "ELSE " & _
'                  "  INSERT INTO PrintValid " & _
'                  "  (SampleID, P, V) VALUES " & _
'                  "  ('" & SampleID & "', " & _
'                  "  " & IIf(LogAsPrinted, NewValue, 0) & ", " & _
'                  "  " & IIf(LogAsValid, NewValue, 0) & ") "
'280   Cnxn(0).Execute Sql
'290   If UCase(RP.PrintAction) <> "SAVE" Then
'300       UpdatePrintValidLog SampleID, "MICRO"
'310   End If
'
'320   Exit Sub
'
'UpdatePrintValid_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'330   intEL = Erl
'340   strES = Err.Description
'350   LogError "modPrintMicro", "UpdatePrintValid", intEL, strES, sql
'
'End Sub
'
'
'Public Sub UpdatePrintValidLog(ByVal SampleID As Variant, _
'                               ByVal Dept As String)
'
'      Dim tb As Recordset
'      Dim sql As String
'      Dim LogDept As String
'
'      'B Biochemistry
'      'C Coagulation
'      'E Endocrinology
'      'H Haematology
'      'I Immunology
'      'M Micro
'      'S ESR
'      'X External
'
'10    On Error GoTo UpdatePrintValidLog_Error
'
'20    Select Case UCase$(Dept)
'      Case "MICRO": LogDept = "M"
'          '  Case "RSV":       LogDept = "V"
'          '  Case "OP":        LogDept = "O"
'          '  Case "CDIFF":     LogDept = "G"
'          '  Case "ROTAADENO": LogDept = "A"
'          '  Case "FOB":       LogDept = "F"
'          '  Case "URINE":     LogDept = "U"
'          '  Case "CANDS":     LogDept = "D"
'30    End Select
'
'40    sql = "SELECT * FROM PrintValidLog WHERE " & _
'            "SampleID = '" & SampleID & "' " & _
'            "AND Department = '" & LogDept & "'"
'50    Set tb = New Recordset
'60    RecOpenServer 0, tb, sql
'70    If tb.EOF Then
'80        tb.AddNew
'90    Else
'100       ValidatedBy = tb!ValidatedBy & ""
'110       sql = "INSERT INTO PrintValidLogArc " & _
'                "  SELECT PrintValidLog.*, " & _
'                "  'PrintHandler', " & _
'                "  '" & Format$(Now, "dd/MMM/yyyy hh:mm:ss") & "' " & _
'                "  FROM PrintValidLog WHERE " & _
'                "  SampleID = '" & SampleID & "' " & _
'                "  AND Department = '" & LogDept & "' "
'120       Cnxn(0).Execute Sql
'130   End If
'140   tb!SampleID = SampleID
'150   tb!Department = LogDept
'160   tb!Printed = 1
'170   tb!Valid = 1
'180   tb!PrintedBy = RP.Initiator
'190   tb!PrintedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
'
'
'      'tb!ValidatedBy = ValidatedBy
'      '
'      'If Not IsNull(tb!ValidatedDateTime) Then
'      '  If Not IsDate(tb!ValidatedDateTime) Then
'      '    tb!ValidatedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
'      '  End If
'      'Else
'      '  tb!ValidatedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
'      'End If
'200   tb.Update
'
'210   Exit Sub
'
'UpdatePrintValidLog_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'220   intEL = Erl
'230   strES = Err.Description
'240   LogError "modPrintMicro", "UpdatePrintValidLog", intEL, strES, sql
'
'End Sub
'
