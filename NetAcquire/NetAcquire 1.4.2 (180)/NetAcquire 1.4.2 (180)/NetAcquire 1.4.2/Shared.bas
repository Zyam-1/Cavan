Attribute VB_Name = "Shared"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public StrEvent As String

Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As Any) As Long
Public Const HH_HELP_CONTEXT = &HF
Public Const CB_FINDSTRING = &H14C


Public UserName As String
Public UserCode As String
Public UserMemberOf As String

Public LogOffDelayMin As Long
Public LogOffDelaySecs As Long

Public Const gVALID = 1
Public Const gNOTVALID = 0
Public Const gPRINTED = 1
Public Const gNOTPRINTED = 0
Public Const gDONTCARE = 2
Public Const gONLYVALID = 1

Public Const gYES = True
Public Const gNO = False
Public Const gNOCHANGE = 4

'Constants for calulating dates
Public Const FORWARD = 1  'used for expiry dates etc
Public Const BACKWARD = 2    'used for DoB etc

Public Cnxn() As Connection
Public CnxnBB() As Connection
Public CnxnRemote() As Connection
Public CnxnRemoteBB() As Connection

Public HospName() As String
Public Entity As String
Public RemoteEntity As String
Public LabNoUpdatePrviousData As String ' Masood 23-09-2014
Public ShowShortScreen As Boolean
Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Const MaxAgeToDays As Long = 43830

Public InterpList(0 To 24) As Single
Public MultiSelectedDemoForLabNoUpdate As String
Public Remote As String
Public LabIDAssigned As Boolean

Public Type udtHaem
    High As Single
    Low As Single
    PlausibleHigh As Single
    PlausibleLow As Single
    DoDelta As Boolean
    DeltaValue As Single
    DeltaDaysBackLimit As Single
End Type

Public Enum PrintAlignContants
    AlignLeft = 0
    AlignCenter = 1
    AlignRight = 2
End Enum

Public intOtherHospitalsInGroup As Integer

Public SampleProcessingError As String

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


Public Function Split_Comm(ByVal Comm As String) As String

          Dim s As String

27970     On Error GoTo Split_Comm_Error

27980     s = Replace(Comm, vbLf, "")
27990     s = Replace(s, vbCr, vbCrLf)

28000     Split_Comm = s

28010     Exit Function

Split_Comm_Error:

          Dim strES As String
          Dim intEL As Integer

28020     intEL = Erl
28030     strES = Err.Description
28040     LogError "Shared", "Split_Comm", intEL, strES

End Function


Public Function FndMaxID(TableName As String, Feild As String, Condition As String) As String
          Dim sql As String
          Dim tb As New ADODB.Recordset
28050     On Error GoTo FndMaxID_Error

          Dim MaxReseveredID As Double
          Dim MaxDemoID As Double

28060     FndMaxID = 0

28070     sql = "Select MAX(CAST(" & Feild & " AS FLOAT)) as MaxID from " & TableName & Condition
28080     Set tb = New Recordset
28090     RecOpenServer 0, tb, sql
28100     If Not tb.EOF Then
28110         If IsNull(tb!MaxID) Then
28120             MaxDemoID = 1
28130         Else
28140             MaxDemoID = tb!MaxID
28150         End If
28160     End If

28170     MaxReseveredID = GetOptionSetting("ReservedLabID", 0)

28180     If MaxDemoID >= MaxReseveredID Then
28190         FndMaxID = MaxDemoID + 1

28200     Else
28210         FndMaxID = MaxReseveredID + 1
28220     End If

28230     SaveOptionSetting "ReservedLabID", FndMaxID
          
          
28240     Exit Function


FndMaxID_Error:

          Dim strES As String
          Dim intEL As Integer

28250     intEL = Erl
28260     strES = Err.Description
28270     LogError "frmEditAll", "FndMaxID", intEL, strES, sql
End Function



Public Function CalcpAge(ByVal DoB As String) As String

          Dim Diff As Long
          Dim DobYr As Single

28280     On Error GoTo CalcpAge_Error

28290     DoB = Format$(DoB, "dd/mm/yyyy")
28300     If IsDate(DoB) Then
28310         Diff = DateDiff("d", (DoB), (Now))
28320         DobYr = Diff / 365.25
28330         If DobYr > 1 Then
28340             CalcpAge = Int(DobYr)
28350         ElseIf Diff < 30.43 Then
28360             CalcpAge = Diff
28370         Else
28380             CalcpAge = Int(Diff / 30.43)
28390         End If
28400     Else
28410         CalcpAge = ""
28420     End If

28430     Exit Function

CalcpAge_Error:

          Dim strES As String
          Dim intEL As Integer

28440     intEL = Erl
28450     strES = Err.Description
28460     LogError "Shared", "CalcpAge", intEL, strES

End Function

Public Sub FillGenericList(ByRef cmb As ComboBox, ByVal ListType As String, Optional ByVal AddEmpty As Boolean = False)

          Dim tb As Recordset
          Dim sql As String

28470     On Error GoTo FillGenericList_Error

28480     cmb.Clear
28490     sql = "SELECT * FROM Lists WHERE " & _
                "ListType = '" & ListType & "' " & _
                "AND InUse = 1 " & _
                "ORDER BY ListOrder"
28500     Set tb = New Recordset
28510     RecOpenServer 0, tb, sql
28520     If AddEmpty Then cmb.AddItem ""
28530     Do While Not tb.EOF
      '        If Len(tb!Text) > 40 Then
      '            cmb.AddItem left(tb!Text, 40) & "..."
      '        Else
28540             cmb.AddItem tb!Text & ""
      '        End If
28550         tb.MoveNext
28560     Loop

28570     Exit Sub

FillGenericList_Error:

          Dim strES As String
          Dim intEL As Integer

28580     intEL = Erl
28590     strES = Err.Description
28600     LogError "Shared", "FillGenericList", intEL, strES, sql

End Sub


Public Function vbGetComputerName() As String

      'Gets the name of the machine
          Const MAXSIZE As Integer = 256
          Dim sTmp As String * MAXSIZE
          Dim lLen As Long

28610     lLen = MAXSIZE - 1
28620     If (GetComputerName(sTmp, lLen)) Then
28630         vbGetComputerName = Left$(sTmp, lLen)
28640     Else
28650         vbGetComputerName = ""
28660     End If

End Function

Public Sub UpdatePrintValidLog(ByVal SampleID As Long, _
                               ByVal Dept As String, _
                               ByVal LogAsValid As Boolean, _
                               ByVal LogAsPrinted As Boolean)

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

28670     On Error GoTo UpdatePrintValidLog_Error

28680     Select Case UCase$(Dept)
          Case "MICRO": LogDept = "M"
28690     Case "SEMEN": LogDept = "Z"
              '  Case "REDSUB":    LogDept = "R"
              '  Case "RSV":       LogDept = "V"
              '  Case "OP":        LogDept = "O"
              '  Case "CDIFF":     LogDept = "G"
              '  Case "FOB":       LogDept = "F"
              '  Case "URINE":     LogDept = "U"
              '  Case "CANDS":     LogDept = "D"
28700     End Select
          Dim v As Integer
          Dim VBy As String
          Dim VDT As String
          Dim p As Integer
          Dim PBy As String
          Dim PDT As String

28710     If LogAsValid Then
28720         v = 1
28730         VBy = AddTicks(UserName)
28740         VDT = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
28750     Else
28760         v = 0
28770         VBy = ""
28780         VDT = ""
28790     End If
28800     If LogAsPrinted Then
28810         p = 1
28820         PBy = AddTicks(UserName)
28830         PDT = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
28840     Else
28850         p = 0
28860         PBy = ""
28870         PDT = ""
28880     End If
      '    sql = "IF NOT EXISTS(SELECT * FROM PrintValidLog WHERE " & _
      '          "              SampleID = '" & SampleID & "' " & _
      '          "              AND Department = '" & LogDept & "') " & _
      '          "  INSERT INTO PrintValidLog " & _
      '          "  (SampleID, Department, Printed, PrintedBy, PrintedDateTime, Valid, ValidatedBy, ValidatedDateTime) VALUES " & _
      '          "  ('" & SampleID & "', " & _
      '          "  '" & LogDept & "', " & _
      '          "  '" & p & "', '" & PBy & "', '" & PDT & "', " & _
      '          "  '" & v & "', '" & VBy & "', '" & VDT & "' ) " & _
      '          "ELSE " & _
      '          "  INSERT INTO PrintValidLogArc " & _
      '          "  SELECT PrintValidLog.*, " & _
      '          "  '" & AddTicks(UserName) & "', " & _
      '          "  '" & Format$(Now, "dd/MMM/yyyy hh:mm:ss") & "' " & _
      '          "  FROM PrintValidLog WHERE " & _
      '          "  SampleID = '" & SampleID & "' " & _
      '          "  AND Department = '" & LogDept & "' " & _
      '          "  UPDATE PrintValidLog SET Valid = " & v & ", ValidatedBy = '" & VBy & "', " & _
      '          "  ValidatedDateTime = " & IIf(v = 1, "'" & VDT & "'", "NULL") & " " & _
      '          "  WHERE   SampleID = '" & SampleID & "' AND Department = '" & LogDept & "'"
      Dim PrintValidLogColums As String
      Dim PrintValidLogArcColums As String

28890 PrintValidLogColums = "SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime"
28900 PrintValidLogArcColums = "SELECT     SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime,SignOff, SignOffBy, SignOffDateTime, ArchivedBy, ArchivedDateTime "
          
          '+++ Junaid 17-01-2024
      '250       Sql = "IF NOT EXISTS(SELECT * FROM PrintValidLog WHERE " & _
      '                "              SampleID = '" & SampleID & "' " & _
      '                "              AND Department = '" & LogDept & "') " & _
      '                "  INSERT INTO PrintValidLog " & _
      '                "  (SampleID, Department, Printed, PrintedBy, PrintedDateTime, Valid, ValidatedBy, ValidatedDateTime) VALUES " & _
      '                "  ('" & SampleID & "', " & _
      '                "  '" & LogDept & "', " & _
      '                "  '" & p & "', '" & PBy & "', '" & PDT & "', " & _
      '                "  '" & v & "', '" & VBy & "', '" & VDT & "' ) " & _
      '                "ELSE " & _
      '                "  INSERT INTO PrintValidLogArc " & _
      '                "  SELECT " & PrintValidLogColums & " " & _
      '                " , SignOff, SignOffBy, SignOffDateTime" & _
      '                                " , '" & AddTicks(UserName) & "', " & _
      '                "  '" & Format$(Now, "dd/MMM/yyyy hh:mm:ss") & "' " & _
      '                "  FROM PrintValidLog WHERE " & _
      '                "  SampleID = '" & SampleID & "' " & _
      '                "  AND Department = '" & LogDept & "' " & _
      '                "  UPDATE PrintValidLog SET Valid = " & v & ", ValidatedBy = '" & VBy & "', " & _
      '                "  ValidatedDateTime = " & IIf(v = 1, "'" & VDT & "'", "NULL") & " " & _
      '                "  WHERE   SampleID = '" & SampleID & "' AND Department = '" & LogDept & "'"
      '
      '260             If v = 0 Then
      '270               Sql = Sql & " UPDATE PrintValidLog SET SignOff = 0, SignOffBy =  NULL , SignOffDateTime = NULL "
      '280               Sql = Sql & "  WHERE   SampleID = '" & SampleID & "' AND Department = '" & LogDept & "'"
      '290             End If
      '____________

28910     sql = "IF NOT EXISTS(SELECT * FROM PrintValidLog WHERE " & _
                "              SampleID = '" & SampleID & "' " & _
                "              AND Department = '" & LogDept & "') " & _
                "  INSERT INTO PrintValidLog " & _
                "  (SampleID, Department, Printed, PrintedBy, PrintedDateTime, Valid, ValidatedBy, ValidatedDateTime) VALUES " & _
                "  ('" & SampleID & "', " & _
                "  '" & LogDept & "', " & _
                "  '" & p & "', '" & VBy & "', '" & VDT & "', " & _
                "  '" & v & "', '" & VBy & "', '" & VDT & "' ) " & _
                "ELSE " & _
                "  INSERT INTO PrintValidLogArc " & _
                "  SELECT " & PrintValidLogColums & " " & _
                " , SignOff, SignOffBy, SignOffDateTime" & _
                                " , '" & AddTicks(UserName) & "', " & _
                "  '" & Format$(Now, "dd/MMM/yyyy hh:mm:ss") & "' " & _
                "  FROM PrintValidLog WHERE " & _
                "  SampleID = '" & SampleID & "' " & _
                "  AND Department = '" & LogDept & "' " & _
                "  UPDATE PrintValidLog SET Valid = " & v & ", ValidatedBy = '" & VBy & "', " & _
                "  ValidatedDateTime = " & IIf(v = 1, "'" & VDT & "'", "NULL") & " " & _
                "  WHERE   SampleID = '" & SampleID & "' AND Department = '" & LogDept & "'"
                
28920           If v = 0 Then
28930             sql = sql & " UPDATE PrintValidLog SET SignOff = 0, SignOffBy =  NULL , SignOffDateTime = NULL "
28940             sql = sql & "  WHERE   SampleID = '" & SampleID & "' AND Department = '" & LogDept & "'"
28950           End If
      '--- Junaid
                
                
28960     Cnxn(0).Execute sql

28970     Exit Sub

UpdatePrintValidLog_Error:

          Dim strES As String
          Dim intEL As Integer

28980     intEL = Erl
28990     strES = Err.Description
29000     LogError "Shared", "UpdatePrintValidLog", intEL, strES, sql


End Sub


Public Sub UpdatePrintValidLog_AutoSignOff(ByVal SampleID As Long, _
                                           ByVal Dept As String)
          Dim sql As String
          Dim LogDept As String
          Dim tb As Recordset
          Dim blnSignOff As Boolean
          'B Biochemistry   'C Coagulation    'E Endocrinology    'H Haematology
          'I Immunology     'M Micro          'S ESR              'X External
29010    On Error GoTo UpdatePrintValidLog_AutoSignOff_Error

29020     Select Case UCase$(Dept)
          Case "MICRO": LogDept = "M"
29030     Case "SEMEN": LogDept = "Z"
29040     End Select

29050     blnSignOff = False 'Default - Don't SignOff
29060     sql = "SELECT OrganismGroup FROM Isolates WHERE SampleID = '" & SampleID & "' order by IsolateNumber"
29070     Set tb = New Recordset
29080     RecOpenServer 0, tb, sql
29090     Do While Not tb.EOF
29100         If UCase(Trim$(tb!OrganismGroup & "")) = "NEGATIVE RESULTS" Then
29110             blnSignOff = True 'NEGATIVE RESULTS - can be Signed off automatically
29120         Else
29130             blnSignOff = False 'Non Negative Results - Don't Sign off
29140             Exit Do
29150         End If
29160         tb.MoveNext
29170     Loop

          'If all Qualifier are "NOT DETECTED" THEN can Auto Sign Off
29180     If Not blnSignOff Then 'Criteria for Auto Sign Off not met yet, but will check again below
29190         sql = "SELECT Qualifier FROM Isolates WHERE SampleID = '" & SampleID & "' order by IsolateNumber"
29200         Set tb = New Recordset
29210         RecOpenServer 0, tb, sql
29220         Do While Not tb.EOF
29230             If InStr(UCase(Trim$(tb!Qualifier & "")), "NOT DETECTED") > 0 Then
29240                 blnSignOff = True 'NEGATIVE RESULTS/NOT DETECTED - can be Signed off automatically
29250             Else
29260                 blnSignOff = False 'Non Negative Results - Don't Sign off
29270                 Exit Do
29280             End If
29290             tb.MoveNext
29300         Loop
29310     End If

29320     If blnSignOff Then 'If True SignOff Sample Id
29330         sql = "UPDATE PrintValidLog SET SignOff = 1, SignOffBy =  'AV' , SignOffDateTime = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' "
29340         sql = sql & " WHERE   SampleID = '" & SampleID & "' AND Department = '" & LogDept & "'"
29350         Cnxn(0).Execute sql
29360     End If

29370    Exit Sub

UpdatePrintValidLog_AutoSignOff_Error:

          Dim strES As String
          Dim intEL As Integer

29380 intEL = Erl
29390 strES = Err.Description
29400 LogError "Shared", "UpdatePrintValidLog_AutoSignOff", intEL, strES, sql
End Sub

Public Sub UpdateFaxLog(ByVal SampleID As String, _
                        ByVal Discipline As String, _
                        ByVal FaxNumber As String)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Disc As String

29410     On Error GoTo UpdateFaxLog_Error

29420     Disc = ""
29430     For n = 0 To 7
29440         If InStr(Discipline, Mid$("HBCIGEMD", n + 1, 1)) Then
29450             Disc = Disc & Discipline
29460         Else
29470             Disc = Disc & " "
29480         End If
29490     Next
          'Check if at least one discipline is selected
29500     If Disc = "" Then Exit Sub

29510     sql = "Select * from FaxLog where " & _
                "SampleID = '" & SampleID & "'"
29520     Set tb = New Recordset
29530     RecOpenServer 0, tb, sql
29540     tb.AddNew
29550     tb!SampleID = SampleID
29560     tb!FaxedTo = FaxNumber
29570     tb!FaxedBy = UserName
29580     tb!DateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
29590     tb!Comment = ""
29600     tb!Discipline = Disc
29610     tb.Update

29620     Exit Sub

UpdateFaxLog_Error:

          Dim strES As String
          Dim intEL As Integer

29630     intEL = Erl
29640     strES = Err.Description
29650     LogError "Shared", "UpdateFaxLog", intEL, strES, sql


End Sub


Public Function IsFaxable(ByVal Source As String, _
                          ByVal Specific As String) As String

          Dim tb As Recordset
          Dim sql As String

29660     On Error GoTo IsFaxable_Error

29670     sql = "Select Fax from " & Source & " where " & _
                "Text = '" & AddTicks(Specific) & "' " & _
                "and Fax <> '' and Fax is not null"
29680     Set tb = New Recordset
29690     RecOpenServer 0, tb, sql

29700     If Not tb.EOF Then
29710         IsFaxable = tb!FAX & ""
29720     Else
29730         IsFaxable = ""
29740     End If

29750     Exit Function

IsFaxable_Error:

          Dim strES As String
          Dim intEL As Integer

29760     intEL = Erl
29770     strES = Err.Description
29780     LogError "Shared", "IsFaxable", intEL, strES, sql


End Function

Public Function GetWardLocation(ByVal Ward As String) As String

      Dim sql As String
      Dim tb As Recordset

29790 On Error GoTo GetWardLocation_Error

29800 GetWardLocation = ""

29810 sql = "SELECT Location FROM Wards WHERE Text = '" & AddTicks(Ward) & "'"
29820 Set tb = New Recordset
29830 RecOpenServer 0, tb, sql
29840 If Not tb.EOF Then
29850     GetWardLocation = tb!Location & ""
29860 End If
          

29870 Exit Function

GetWardLocation_Error:

       Dim strES As String
       Dim intEL As Integer

29880  intEL = Erl
29890  strES = Err.Description
29900  LogError "Shared", "GetWardLocation", intEL, strES, sql
          
End Function

Public Function GetWard(ByVal CodeOrText As String, _
                        ByVal HospCode As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

29910     On Error GoTo GetWard_Error

29920     s = AddTicks(Trim$(CodeOrText))

29930     sql = "Select [Text] from Wards where " & _
                "HospitalCode = '" & HospCode & "' " & _
                "and Inuse = 1 " & _
                "and (Code = '" & s & "' or [Text] = '" & s & "')"
29940     Set tb = New Recordset
29950     RecOpenServer 0, tb, sql
29960     If Not tb.EOF Then
29970         GetWard = tb!Text & ""
29980     Else
29990         GetWard = CodeOrText
30000     End If

30010     Exit Function

GetWard_Error:

          Dim strES As String
          Dim intEL As Integer

30020     intEL = Erl
30030     strES = Err.Description
30040     LogError "Shared", "GetWard", intEL, strES, sql


End Function


Public Function GetClinician(ByVal CodeOrText As String, _
                             ByVal HospCode As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

30050     On Error GoTo GetClinician_Error

30060     s = AddTicks(Trim$(CodeOrText))

30070     sql = "Select [Text] from Clinicians where " & _
                "HospitalCode = '" & HospCode & "' " & _
                "and Inuse = 1 " & _
                "and (Code = '" & s & "' or [Text] = '" & s & "')"
30080     Set tb = New Recordset
30090     RecOpenServer 0, tb, sql
30100     If Not tb.EOF Then
30110         GetClinician = tb!Text & ""
30120     Else
30130         GetClinician = CodeOrText
30140     End If

30150     Exit Function

GetClinician_Error:

          Dim strES As String
          Dim intEL As Integer

30160     intEL = Erl
30170     strES = Err.Description
30180     LogError "Shared", "GetClinician", intEL, strES, sql


End Function

Public Sub Archive(ByVal FromSQL As String, ByVal ToTable As String)

          Dim tbTo As Recordset
          Dim tbFrom As Recordset
          Dim f As Field
          Dim sql As String

30190     On Error GoTo Archive_Error

30200     Set tbFrom = New Recordset
30210     RecOpenServer 0, tbFrom, FromSQL
30220     Do While Not tbFrom.EOF

30230         sql = "SELECT * FROM " & ToTable & " WHERE 0 = 1"
30240         Set tbTo = New Recordset
30250         RecOpenServer 0, tbTo, sql

30260         tbTo.AddNew
30270         For Each f In tbTo.Fields
30280             If UCase$(f.Name) = "DATETIMEOFARCHIVE" Then
30290                 tbTo!DateTimeOfArchive = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
30300             ElseIf UCase$(f.Name) = "ARCHIVEDATETIME" Then
30310                 tbTo!ArchiveDateTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
30320             ElseIf UCase$(f.Name) = "ARCHIVEDBY" Then
30330                 tbTo!ArchivedBy = UserName
30340             ElseIf UCase$(f.Name) <> "ROWGUID" And UCase(f.Name) <> "PKID" Then
30350                 tbTo(f.Name) = tbFrom(f.Name)
30360             End If

30370         Next
30380         tbTo.Update

30390         tbFrom.MoveNext

30400     Loop

30410     Exit Sub

Archive_Error:

          Dim strES As String
          Dim intEL As Integer

30420     intEL = Erl
30430     strES = Err.Description
30440     LogError "Shared", "Archive", intEL, strES, sql

End Sub

Public Function BioLongNameFor(ByVal LongOrShortName As String) As String

          Dim tb As Recordset
          Dim sql As String

30450     On Error GoTo BioLongNameFor_Error

30460     sql = "Select LongName from BioTestDefinitions where " & _
                "ShortName = '" & AddTicks(LongOrShortName) & "' " & _
                "or LongName = '" & AddTicks(LongOrShortName) & "'"
30470     Set tb = New Recordset
30480     RecOpenServer 0, tb, sql
30490     If Not tb.EOF Then
30500         BioLongNameFor = tb!LongName & ""
30510     Else
30520         BioLongNameFor = "???"
30530     End If

30540     Exit Function

BioLongNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

30550     intEL = Erl
30560     strES = Err.Description
30570     LogError "Shared", "BioLongNameFor", intEL, strES, sql


End Function


Public Function CoagNameFor(ByVal Code As String) As String

          Dim tb As Recordset
          Dim sql As String

30580     On Error GoTo CoagNameFor_Error

30590     sql = "Select TestName from CoagTestDefinitions where " & _
                "Code = '" & Code & "'"
30600     Set tb = New Recordset
30610     RecOpenServer 0, tb, sql
30620     If Not tb.EOF Then
30630         CoagNameFor = tb!TestName & ""
30640     Else
30650         CoagNameFor = "???"
30660     End If

30670     Exit Function

CoagNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

30680     intEL = Erl
30690     strES = Err.Description
30700     LogError "Shared", "CoagNameFor", intEL, strES, sql


End Function
Public Function BioShortNameFor(ByVal LongOrShortName As String) As String

          Dim tb As Recordset
          Dim sql As String

30710     On Error GoTo BioShortNameFor_Error

30720     sql = "SELECT ShortName FROM BioTestDefinitions WHERE " & _
                "ShortName = '" & AddTicks(LongOrShortName) & "' " & _
                "OR LongName = '" & AddTicks(LongOrShortName) & "'"
30730     Set tb = New Recordset
30740     RecOpenServer 0, tb, sql
30750     If Not tb.EOF Then
30760         BioShortNameFor = tb!ShortName & ""
30770     Else
30780         BioShortNameFor = "???"
30790     End If

30800     Exit Function

BioShortNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

30810     intEL = Erl
30820     strES = Err.Description
30830     LogError "Shared", "BioShortNameFor", intEL, strES, sql

End Function


Public Function CoagCodeForTestName(ByVal TestName As String) As String

          Dim tb As Recordset
          Dim sql As String

30840     On Error GoTo CoagCodeForTestName_Error

30850     sql = "Select Code from CoagTestDefinitions where " & _
                "TestName = '" & TestName & "'"
30860     Set tb = New Recordset
30870     RecOpenServer 0, tb, sql
30880     If Not tb.EOF Then
30890         CoagCodeForTestName = tb!Code & ""
30900     Else
30910         CoagCodeForTestName = "???"
30920     End If

30930     Exit Function

CoagCodeForTestName_Error:

          Dim strES As String
          Dim intEL As Integer

30940     intEL = Erl
30950     strES = Err.Description
30960     LogError "Shared", "CoagCodeForTestName", intEL, strES, sql


End Function

Public Function dmyFromCount(ByVal Days As Long) As String

          Dim D As Long
          Dim m As Long
          Dim Y As Long
          Dim s As String

30970     Y = Int(Days / 365)

30980     Days = Days - (Y * 365)

30990     m = Days \ 30

31000     D = Days - (m * 30)

31010     If Y > 0 Then
31020         s = Format$(Y) & "Y "
31030     End If

31040     If m > 0 Then
31050         s = s & Format$(m) & "M "
31060     End If

31070     dmyFromCount = s & Format$(D, "0") & "D"

          'y = Int(Days / 365)
          '
          'Days = Days - (y * 365.25)
          '
          'm = Days \ 30.42
          '
          'd = Days - (m * 30.42)
          '
          'If y > 0 Then
          '  s = Format$(y) & "Y "
          'End If
          '
          'If m > 0 Then
          '  s = s & Format$(m) & "M "
          'End If
          '
          'dmyFromCount = s & Format$(d, "0") & "D"


End Function


'---------------------------------------------------------------------------------------
' Procedure : AddTicks
' Author    : Masood
' Date      : 22/Sep/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function AddTicks(ByVal s As String) As String

31080     On Error GoTo AddTicks_Error


31090     s = Trim$(s)

31100     s = Replace(s, "'", "''")

          '    AddTicks = s

          ' Masood 16 Sep 2015
      '    Dim FirstSpace As Integer
      '    Dim SecondSpace As Integer
      '    Dim o As String
      '
      '    FirstSpace = InStr(s, " ")
      '    SecondSpace = InStrRev(s, " ")
      '
      '    If Len(s) > 2 Then
      '        o = Mid(s, Len(s) - FirstSpace, 2)
      '        If Trim(o) = "O" Then
      '            s = Right(s, Len(s) - FirstSpace)
      '            s = Replace(s, " ", "")
      '            '        MsgBox s
      '        ElseIf FirstSpace > 0 Then
      '            '            s = Mid(s, FirstSpace - 1, 1)
      '            If Trim("O") = Trim(Mid(s, FirstSpace - 1, 1)) Then
      '                s = Replace(s, " ", "")
      '            End If
      '        End If
      '    End If

31110     AddTicks = s
          ' Masood 16 Sep 2015

31120     Exit Function


AddTicks_Error:

          Dim strES As String
          Dim intEL As Integer

31130     intEL = Erl
31140     strES = Err.Description
31150     LogError "Shared", "AddTicks", intEL, strES

End Function

'---------------------------------------------------------------------------------------
' Procedure : NameO
' Author    : Masood
' Date      : 27/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function NameO(s As String) As String
          
          
          Dim o As String
          Dim FirstSpace As Integer
          Dim SecondSpace As Integer
          Dim FirstName As String
          Dim sn As String
          Dim FTowChar As String
          
31160     On Error GoTo NameO_Error




31170     FirstSpace = InStr(s, " ")
31180     FirstName = Left(s, (Len(s) - FirstSpace))
31190     sn = Right(s, (Len(s) - FirstSpace))
31200     FTowChar = Left(sn, 2)

31210     If InStr(sn, "O") Then
31220         If Len(sn) > 2 Then
31230             sn = Right(sn, Len(sn) - 2)
31240         End If
31250         FTowChar = Replace(FTowChar, "'", "")
31260         FTowChar = Replace(FTowChar, " ", "")

31270         sn = FirstName & " " & FTowChar & sn
31280         NameO = sn
31290     Else
31300         NameO = s
31310     End If


31320     Exit Function


NameO_Error:

          Dim strES As String
          Dim intEL As Integer

31330     intEL = Erl
31340     strES = Err.Description
31350     LogError "Shared", "NameO", intEL, strES
End Function


Public Function AreMicroResultsPresent(ByVal SampleIDWithOffset As String) As Integer

          Dim tb As Recordset
          Dim sql As String

31360     On Error GoTo AreMicroResultsPresent_Error

31370     sql = "SELECT COUNT(*) Tot " & _
                "FROM FaecesResults50 F, " & _
                "SiteDetails50 MSD, " & _
                "UrineResults50 U, " & _
                "UrineIdent50 UI WHERE " & _
                "F.SampleID = '" & SampleIDWithOffset & "' " & _
                "OR MSD.SampleID = '" & SampleIDWithOffset & "' " & _
                "OR U.SampleID = '" & SampleIDWithOffset & "' " & _
                "OR UI.SampleID = '" & SampleIDWithOffset & "' "

31380     Set tb = New Recordset
31390     Set tb = Cnxn(0).Execute(sql)

31400     AreMicroResultsPresent = Sgn(tb!Tot)

31410     Exit Function

AreMicroResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

31420     intEL = Erl
31430     strES = Err.Description
31440     LogError "Shared", "AreMicroResultsPresent", intEL, strES, sql

End Function


Public Sub GetWardClinGP(ByVal SampleID As Long, _
                         ByRef Ward As String, _
                         ByRef Clin As String, _
                         ByRef GP As String)

          Dim tb As Recordset
          Dim sql As String

31450     On Error GoTo GetWardClinGP_Error

31460     sql = "Select Ward, Clinician, GP from Demographics where " & _
                "SampleID = '" & SampleID & "'"
31470     Set tb = New Recordset
31480     RecOpenServer 0, tb, sql
31490     If Not tb.EOF Then
31500         Ward = tb!Ward & ""
31510         Clin = tb!Clinician & ""
31520         GP = tb!GP & ""
31530     Else
31540         Ward = ""
31550         Clin = ""
31560         GP = ""
31570     End If

31580     Exit Sub

GetWardClinGP_Error:

          Dim strES As String
          Dim intEL As Integer

31590     intEL = Erl
31600     strES = Err.Description
31610     LogError "Shared", "GetWardClinGP", intEL, strES, sql


End Sub

Public Function QuickInterpBio(ByVal Result As BIEResult) _
       As String

31620     With Result
31630         If Val(.Result) < .Low Then
31640             QuickInterpBio = "Low "
31650         ElseIf Val(.Result) > .High Then
31660             QuickInterpBio = "High"
31670         Else
31680             QuickInterpBio = "    "
31690         End If
31700     End With

End Function

Public Function CountDays(ByVal Number As Long, ByVal Interval As String) As Long

31710     Select Case Interval
          Case "Days": CountDays = Number
31720     Case "Months": CountDays = Number * (365.25 / 12)
31730     Case "Years": CountDays = Number * 365.25
31740     End Select

End Function

Public Sub FillCommentLines(ByVal FullComment As String, _
                            ByVal NumberOfLines As Integer, _
                            ByRef Comments() As String, _
                            Optional ByVal MaxLen As Integer = 80)

          Dim n As Integer
          Dim CurrentLine As Integer
          Dim X As Integer
          Dim ThisLine As String
          Dim SpaceFound As Boolean

31750     On Error GoTo ErrorHandler

31760     For n = 1 To UBound(Comments)
31770         Comments(n) = ""
31780     Next

31790     CurrentLine = 0
31800     FullComment = Trim$(FullComment)
31810     n = Len(FullComment)

31820     For X = n - 1 To 1 Step -1
31830         If Mid$(FullComment, X, 1) = vbCr Or Mid$(FullComment, X, 1) = vbLf Or Mid$(FullComment, X, 1) = vbTab Then
31840             Mid$(FullComment, X, 1) = " "
31850         End If
31860     Next

31870     For X = n - 3 To 1 Step -1
31880         If Mid$(FullComment, X, 2) = "  " Then
31890             FullComment = Left$(FullComment, X) & Mid$(FullComment, X + 2)
31900         End If
31910     Next
31920     n = Len(FullComment)

31930     Do While n > MaxLen
31940         SpaceFound = False
31950         For X = MaxLen To 1 Step -1
31960             If Mid$(FullComment, X, 1) = " " Then
31970                 ThisLine = Left$(FullComment, X - 1)
31980                 FullComment = Mid$(FullComment, X + 1)

31990                 CurrentLine = CurrentLine + 1
32000                 If CurrentLine <= NumberOfLines Then
32010                     Comments(CurrentLine) = ThisLine
32020                 End If
32030                 SpaceFound = True
32040                 Exit For
32050             End If
32060         Next
32070         If Not SpaceFound Then
32080             ThisLine = Left$(FullComment, MaxLen)
32090             FullComment = Mid$(FullComment, MaxLen + 1)

32100             CurrentLine = CurrentLine + 1
32110             If CurrentLine <= NumberOfLines Then
32120                 Comments(CurrentLine) = ThisLine
32130             End If
32140         End If
32150         n = Len(FullComment)
32160     Loop

32170     CurrentLine = CurrentLine + 1
32180     If CurrentLine <= NumberOfLines Then
32190         Comments(CurrentLine) = FullComment
32200     End If
32210     Exit Sub
ErrorHandler:
          '    MsgBox Err.Description, vbInformation
End Sub

Public Sub LogTimeOfPrinting(ByVal SampleID As String, _
                             ByVal Dept As String)
      'Dept is "H", "B" or "C" or "D"

          Dim sql As String

32220     On Error GoTo LogTimeOfPrinting_Error

32230     Select Case UCase$(Left$(Dept, 1))
          Case "H": Dept = "DateTimeHaemPrinted"
32240     Case "B": Dept = "DateTimeBioPrinted"
32250     Case "C": Dept = "DateTimeCoagPrinted"
32260     Case "D": Dept = "DateTimeDemographics"
32270     Case Else: Exit Sub
32280     End Select

32290     sql = "update Demographics set " & _
                Dept & " = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' " & _
                "where SampleID = '" & SampleID & "' " & _
                "and " & Dept & " is null"
32300     Cnxn(0).Execute sql

32310     Exit Sub

LogTimeOfPrinting_Error:

          Dim strES As String
          Dim intEL As Integer

32320     intEL = Erl
32330     strES = Err.Description
32340     LogError "Shared", "LogTimeOfPrinting", intEL, strES, sql


End Sub

Public Function QueryKnown(ByVal CodeOrText As String, _
                           Optional ByVal Hospital As String) _
                           As String
      'Returns either "" = not known
      '        or CodeOrText = known

          Dim HospCode As String
          Dim Original As String
          Dim sql As String
          Dim tb As Recordset

32350     On Error GoTo QueryKnown_Error

32360     QueryKnown = ""
32370     Original = CodeOrText

32380     CodeOrText = Trim$(UCase$(AddTicks(CodeOrText)))
32390     If CodeOrText = "" Then Exit Function

32400     If Hospital <> "" Then
32410         sql = "Select * from Lists where " & _
                    "ListType = 'HO' " & _
                    "and Text = '" & Hospital & "' and InUse = 1"
32420         Set tb = New Recordset
32430         RecOpenServer 0, tb, sql
32440         If Not tb.EOF Then
32450             HospCode = tb!Code & ""
32460         End If
32470     End If

32480     sql = "Select * from Clinicians where " & _
                "(Code = '" & CodeOrText & "' " & _
                "or Text = '" & CodeOrText & "') and InUse = 1"

32490     If Hospital <> "" Then
32500         sql = sql & "And HospitalCode = '" & HospCode & "' and InUse = 1"
32510     End If

32520     Set tb = New Recordset
32530     RecOpenServer 0, tb, sql
32540     If Not tb.EOF Then
32550         QueryKnown = tb!Text & ""
32560     Else
32570         QueryKnown = ""    'Original
32580     End If

32590     Exit Function

QueryKnown_Error:

          Dim strES As String
          Dim intEL As Integer

32600     intEL = Erl
32610     strES = Err.Description
32620     LogError "Shared", "QueryKnown", intEL, strES, sql

End Function

Public Function ParseForeName(ByVal Name As String) As String

          Dim strNameParts
          Dim intU As Integer
          Dim strTemp As String

32630     Name = Trim$(Name)
32640     If Name = "" Then
32650         ParseForeName = ""
32660     Else
32670         strNameParts = Split(Name)
32680         intU = UBound(strNameParts)
32690         strTemp = strNameParts(intU)
32700         If strTemp Like "*[!'!a-z!A-Z]*" Or Len(strTemp) = 1 Then
32710             ParseForeName = ""
32720         Else
32730             ParseForeName = strTemp
32740         End If
32750     End If

End Function

Public Function ParseSurName(ByVal Name As String) As String

          Dim strNameParts
          Dim strTemp As String
          Dim n As Integer

32760     Name = Trim$(Name)
32770     If Name = "" Then
32780         ParseSurName = ""
32790     Else
32800         strNameParts = Split(Name)

32810         For n = 0 To UBound(strNameParts) - 1
32820             strTemp = strTemp & strNameParts(n) & " "
32830         Next
32840         strTemp = Trim$(strTemp)

32850         If strTemp Like "*[!'!a-z!A-Z ]*" Or Len(strTemp) = 1 Then
32860             ParseSurName = ""
32870         Else
32880             ParseSurName = strTemp
32890         End If
32900     End If

End Function

Public Function BetweenDates(ByVal Index As Integer, _
                             ByRef UpTo As String) _
                             As String

          Dim From As String
          Dim m As Integer

32910     Select Case Index
          Case 0:    'last week
32920         From = Format$(DateAdd("ww", -1, Now), "dd/mm/yyyy")
32930         UpTo = Format$(Now, "dd/mm/yyyy")
32940     Case 1:    'last month
32950         From = Format$(DateAdd("m", -1, Now), "dd/mm/yyyy")
32960         UpTo = Format$(Now, "dd/mm/yyyy")
32970     Case 2:    'last fullmonth
32980         From = Format$(DateAdd("m", -1, Now), "dd/mm/yyyy")
32990         From = "01/" & Mid$(From, 4)
33000         UpTo = DateAdd("m", 1, From)
33010         UpTo = Format$(DateAdd("d", -1, UpTo), "dd/mm/yyyy")
33020     Case 3:    'last quarter
33030         From = Format$(DateAdd("q", -1, Now), "dd/mm/yyyy")
33040         UpTo = Format$(Now, "dd/mm/yyyy")
33050     Case 4:    'last full quarter
33060         From = Format$(DateAdd("q", -1, Now), "dd/mm/yyyy")
33070         m = Val(Mid$(From, 4, 2))
33080         m = ((m - 1) \ 3) * 3 + 1
33090         From = "01/" & Format$(m, "00") & Mid$(From, 6)
33100         UpTo = DateAdd("q", 1, From)
33110         UpTo = Format$(DateAdd("d", -1, UpTo), "dd/mm/yyyy")
33120     Case 5:    'year to date
33130         From = "01/01/" & Format$(Now, "yyyy")
33140         UpTo = Format$(Now, "dd/mm/yyyy")
33150     Case 6:    'today
33160         From = Format$(Now, "dd/mm/yyyy")
33170         UpTo = From
33180     End Select

33190     BetweenDates = From

End Function


'---------------------------------------------------------------------------------------
' Procedure : CalcAgeToDays
' Author    : Masood
' Date      : 13/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function CalcAgeToDays(cYears As Integer, cMonths As Integer, cDays As Integer) As Double

33200     On Error GoTo CalcAgeToDays_Error


33210     CalcAgeToDays = (Val(cYears) * 365.25) + (Val(cMonths) * 30.42) + Val(cDays)

33220     Exit Function


CalcAgeToDays_Error:

          Dim strES As String
          Dim intEL As Integer

33230     intEL = Erl
33240     strES = Err.Description
33250     LogError "Shared", "CalcAgeToDays", intEL, strES
End Function




Public Function CalcAge(ByVal DoB As String, ByVal SampleDate As String) As String

      Dim Diffdays As Long

33260     On Error GoTo CalcAge_Error

33270     CalcAge = ""
          Dim CDob, CsDate, dobFormat, sDateFormat As String
33280     CDob = CDate(DoB)
33290     CsDate = CDate(SampleDate)
33300     dobFormat = Format(CDob, "dd/mm/yyyy")
33310     sDateFormat = Format(CsDate, "dd/mm/yyyy")
          
          
          
33320     If IsDate(DoB) And IsDate(SampleDate) Then
33330         Diffdays = DateDiff("d", dobFormat, sDateFormat)
33340         If Diffdays < 30 Then
33350             CalcAge = Diffdays & " D"
      '70            ElseIf Diffdays < 57 Then
      '80                CalcAge = Diffdays \ 7 & " W"
33360         ElseIf Diffdays < 364 Then
33370             CalcAge = Diffdays \ 30 & " M"
33380         Else
33390             CalcAge = Diffdays \ 365.25 & " Y"
33400         End If
33410     End If

33420     Exit Function

CalcAge_Error:

          Dim strES As String
          Dim intEL As Integer

33430     intEL = Erl
33440     strES = Err.Description
      '          MsgBox strES
33450     LogError "Shared", "CalcAge", intEL, strES


End Function

Sub LogBioAsPrinted(ByVal SampleID As String, _
                    ByVal TestCode As String)

          Dim sql As String

33460     On Error GoTo LogBioAsPrinted_Error

33470     sql = "update BioResults " & _
                "set valid = 1, printed = 1 where " & _
                "SampleID = '" & SampleID & "' " & _
                "and code = '" & TestCode & "'"

33480     Cnxn(0).Execute sql

33490     Exit Sub

LogBioAsPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

33500     intEL = Erl
33510     strES = Err.Description
33520     LogError "Shared", "LogBioAsPrinted", intEL, strES, sql

End Sub

Public Sub PrintResultBioWin(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Ward As String
          Dim Clin As String
          Dim GP As String

33530     On Error GoTo PrintResultBioWin_Error

33540     GetWardClinGP SampleID, Ward, Clin, GP

33550     sql = "Select * from PrintPending where " & _
                "Department = 'B' " & _
                "and SampleID = '" & SampleID & "'"
33560     Set tb = New Recordset
33570     RecOpenClient 0, tb, sql
33580     If tb.EOF Then
33590         tb.AddNew
33600     End If
33610     tb!SampleID = SampleID
33620     tb!Ward = Ward
33630     tb!Clinician = Clin
33640     tb!GP = GP
33650     tb!Department = "B"
33660     tb!Initiator = UserName
33670     tb.Update

33680     Exit Sub

PrintResultBioWin_Error:

          Dim strES As String
          Dim intEL As Integer

33690     intEL = Erl
33700     strES = Err.Description
33710     LogError "Shared", "PrintResultBioWin", intEL, strES, sql

End Sub


Function nr(ByVal Analyte As String, _
            ByVal Sex As String, _
            ByVal DoB As String) As String

          Dim l As String * 4
          Dim H As String * 4
          Dim fMat As String
          Dim DaysOld As Long
          Dim SelectNormalRange As String
          Dim sql As String
          Dim tb As Recordset

33720     On Error GoTo nr_Error

33730     If IsDate(DoB) Then
33740         DaysOld = Abs(DateDiff("d", Now, DoB))
33750     Else
33760         DaysOld = 7300    'Default Age 20 years
33770     End If

33780     nr = "(    -    )"

33790     Select Case Left$(UCase$(Trim$(Sex)), 1)
          Case "M": SelectNormalRange = "MaleLow as Low, MaleHigh as High"
33800     Case "F": SelectNormalRange = "FemaleLow as Low, FemaleHigh as High"
33810     Case Else: SelectNormalRange = "FemaleLow as Low, MaleHigh as High"
33820     End Select

33830     sql = "Select top 1 " & SelectNormalRange & " from HaemTestDefinitions where " & _
                "AgeFromDays <= " & DaysOld & " " & _
                "and AgeToDays >= " & DaysOld & " " & _
                "and AnalyteName = '" & Analyte & "' " & _
                "order by AgeFromDays Asc, AgeToDays Desc"
33840     Set tb = New Recordset
33850     RecOpenServer 0, tb, sql
33860     If tb.EOF Then Exit Function

33870     If tb!High = 999 Then
33880         nr = "           "
33890         Exit Function
33900     End If

33910     Select Case UCase$(Analyte)
          Case "WBC", "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC", "RDWCV", "FIB", "NEUTA", "LYMA", "MONOA", "EOSA", "BASA": fMat = "#0.0"
33920     Case Else: fMat = "####"
33930     End Select

33940     RSet l = Format$(tb!Low, fMat)
33950     Mid$(nr, 2, 4) = l
33960     LSet H = Format$(tb!High, fMat)
33970     Mid$(nr, 7, 4) = H

33980     Exit Function

nr_Error:

          Dim strES As String
          Dim intEL As Integer

33990     intEL = Erl
34000     strES = Err.Description
34010     LogError "Shared", "nr", intEL, strES, sql


End Function
Sub buildinterp(tb As Recordset, i() As String)

          Dim n As Integer
          Dim R As String
          Dim l As Integer

34020     l = True
34030     n = 0

34040     If Val(tb!NeutA & "") <> 0 And Val(tb!NeutP & "") <> 0 Then
34050         R = interp(0, Val(tb!NeutA))
34060         If R <> "" Then
34070             i(0) = R
34080             l = False
34090         Else
34100             R = interp(1, Val(tb!NeutP))
34110             If R <> "" Then
34120                 i(0) = R
34130                 l = False
34140             End If
34150         End If
34160         R = interp(2, Val(tb!NeutA))
34170         If R <> "" Then
34180             i(0) = R
34190             l = False
34200         Else
34210             R = interp(3, Val(tb!NeutP))
34220             If R <> "" Then
34230                 i(0) = R
34240                 l = False
34250             End If
34260         End If
34270     End If

34280     If Val(tb!LymA & "") <> 0 And Val(tb!LymP & "") <> 0 Then
34290         R = interp(4, Val(tb!LymA))
34300         If R <> "" Then
34310             i(0) = i(0) & R
34320             l = Not l
34330             If l Then n = n + 1
34340         Else
34350             R = interp(5, Val(tb!LymP))
34360             If R <> "" Then
34370                 i(0) = i(0) & R
34380                 l = Not l
34390                 If l Then n = n + 1
34400             End If
34410         End If
34420         R = interp(6, Val(tb!LymA))
34430         If R <> "" Then
34440             i(0) = i(0) & R
34450             l = Not l
34460             If l Then n = n + 1
34470         Else
34480             R = interp(7, Val(tb!LymP))
34490             If R <> "" Then
34500                 i(0) = i(0) & R
34510                 l = Not l
34520                 If l Then n = n + 1
34530             End If
34540         End If
34550     End If



34560     If Val(tb!MonoA & "") <> 0 And Val(tb!MonoP & "") <> 0 Then
34570         R = interp(8, Val(tb!MonoA))
34580         If R <> "" Then
34590             i(n) = i(n) & R
34600             l = Not l
34610             If l Then n = n + 1
34620         Else
34630             R = interp(9, Val(tb!MonoP))
34640             If R <> "" Then
34650                 i(n) = i(n) & R
34660                 l = Not l
34670                 If l Then n = n + 1
34680             End If
34690         End If
34700     End If



34710     If Val(tb!EosA & "") <> 0 And Val(tb!EosP & "") <> 0 Then
34720         R = interp(10, Val(tb!EosA))
34730         If R <> "" Then
34740             i(n) = i(n) & R
34750             l = Not l
34760             If l Then n = n + 1
34770         Else
34780             R = interp(11, Val(tb!EosP))
34790             If R <> "" Then
34800                 i(n) = i(n) & R
34810                 l = Not l
34820                 If l Then n = n + 1
34830             End If
34840         End If
34850     End If



34860     If Val(tb!BasA & "") <> 0 And Val(tb!BasP & "") <> 0 Then
34870         R = interp(12, Val(tb!BasA))
34880         If R <> "" Then
34890             i(n) = i(n) & R
34900             l = Not l
34910             If l Then n = n + 1
34920         Else
34930             R = interp(13, Val(tb!BasP))
34940             If R <> "" Then
34950                 i(n) = i(n) & R
34960                 l = Not l
34970                 If l Then n = n + 1
34980             End If
34990         End If
35000     End If




35010     If Val(tb!WBC & "") <> 0 Then
35020         R = interp(14, Val(tb!WBC))
35030         If R <> "" Then
35040             i(n) = i(n) & R
35050             l = Not l
35060             If l Then n = n + 1
35070         Else
35080             R = interp(15, Val(tb!WBC))
35090             If R <> "" Then
35100                 i(n) = i(n) & R
35110                 l = Not l
35120                 If l Then n = n + 1
35130             End If
35140         End If
35150     End If



35160     If (Val(tb!rdwsd & "") <> 0) And (Val(tb!RDWCV & "") <> 0) Then
35170         R = interp(16, Val(tb!rdwsd))
35180         If R <> "" Then
35190             i(n) = i(n) & R
35200             l = Not l
35210             If l Then n = n + 1
35220         Else
35230             R = interp(17, Val(tb!RDWCV))
35240             If R <> "" Then
35250                 i(n) = i(n) & R
35260                 l = Not l
35270                 If l Then n = n + 1
35280             End If
35290         End If
35300     End If


35310     If Val(tb!MCV & "") <> 0 Then
35320         R = interp(18, Val(tb!MCV))
35330         If R <> "" Then
35340             i(n) = i(n) & R
35350             l = Not l
35360             If l Then n = n + 1
35370         Else
35380             R = interp(19, Val(tb!MCV))
35390             If R <> "" Then
35400                 i(n) = i(n) & R
35410                 l = Not l
35420                 If l Then n = n + 1
35430             End If
35440         End If
35450     End If



35460     If Val(tb!mchc & "") <> 0 Then
35470         R = interp(20, Val(tb!mchc))
35480         If R <> "" Then
35490             i(n) = i(n) & R
35500             l = Not l
35510             If l Then n = n + 1
35520         End If
35530     End If



35540     If Val(tb!Hgb & "") <> 0 Then
35550         R = interp(21, Val(tb!Hgb))
35560         If R <> "" Then
35570             i(n) = i(n) & R
35580             l = Not l
35590             If l Then n = n + 1
35600         End If
35610     End If


35620     If Val(tb!rbc & "") <> 0 Then
35630         R = interp(22, Val(tb!rbc))
35640         If R <> "" Then
35650             i(n) = i(n) & R
35660             l = Not l
35670             If l Then n = n + 1
35680         End If
35690     End If



35700     If Val(tb!plt & "") <> 0 Then
35710         R = interp(23, Val(tb!plt))
35720         If R <> "" Then
35730             i(n) = i(n) & R
35740             l = Not l
35750             If l Then n = n + 1
35760         Else
35770             R = interp(24, Val(tb!plt))
35780             If R <> "" Then
35790                 i(n) = i(n) & R
35800                 l = Not l
35810                 If l Then n = n + 1
35820             End If
35830         End If
35840     End If

End Sub

Function calcmean(v() As Single) As Single

          Dim mean As Single
          Dim n As Integer
          Dim Entries As Integer

35850     Entries = (UBound(v) - LBound(v)) + 1
35860     mean = 0
35870     For n = LBound(v) To UBound(v)
35880         mean = mean + v(n)
35890     Next
35900     mean = mean / Entries

35910     calcmean = mean

End Function


Public Function AreResultsPresent(ByVal Dept As String, ByVal SampleID As String) _
       As Boolean

      'dept = "BGA", "Bio", "Haem", "Coag", "Histo", "Cyto", "Imm", "Ext"

          Dim tb As Recordset
          Dim sql As String

35920     On Error GoTo AreResultsPresent_Error

35930     sql = "Select count(*) as tot from " & Dept & "Results where " & _
                "SampleID = '" & Val(SampleID) & "'"
35940     Set tb = New Recordset
35950     Set tb = Cnxn(0).Execute(sql)

35960     AreResultsPresent = Sgn(tb!Tot)

35970     Exit Function

AreResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

35980     intEL = Erl
35990     strES = Err.Description
36000     LogError "Shared", "AreResultsPresent", intEL, strES, sql


End Function



Public Function QueryTwo(ByVal in1 As String, _
                         ByVal in2 As String, _
                         ByVal inCaption As String, _
                         ByVal ShowNeitherButton As Boolean) _
                         As String
      'Given two options, return one
          Dim f As New frmQueryTwo

36010     f.Caption = inCaption
36020     f.ShowNeither = ShowNeitherButton
36030     f.cmdSelect(0).Caption = in1
36040     f.cmdSelect(1).Caption = in2
36050     f.Show 1

36060     QueryTwo = f.ReturnVal

36070     Unload f
36080     Set f = Nothing

End Function

Function AreFlagsPresent(f() As Integer) As Boolean

          Dim n As Integer

36090     AreFlagsPresent = False

36100     For n = 0 To 5
36110         If f(n) Then
36120             AreFlagsPresent = True
36130             Exit Function
36140         End If
36150     Next

End Function


Function interp(ByVal p As Integer, ByVal v As Single) As String

36160     On Error GoTo interp_Error

36170     Select Case p
          Case 0: If v < InterpList(0) Then interp = "Neutropaenia   "
36180     Case 1: If v < InterpList(1) Then interp = "Neutropaenia   "
36190     Case 2: If v > InterpList(2) Then interp = "Neutrophilia   "
36200     Case 3: If v > InterpList(3) Then interp = "Neutrophilia   "
36210     Case 4: If v < InterpList(4) Then interp = "Lymphopaenia   "
36220     Case 5: If v < InterpList(5) Then interp = "Lymphopaenia   "
36230     Case 6: If v > InterpList(6) Then interp = "Lymphocytosis  "
36240     Case 7: If v > InterpList(7) Then interp = "Lymphocytosis  "
36250     Case 8: If v > InterpList(8) Then interp = "Monocytosis    "
36260     Case 9: If v > InterpList(9) Then interp = "Monocytosis    "
36270     Case 10: If v > InterpList(10) Then interp = "Eosinophilia   "
36280     Case 11: If v > InterpList(11) Then interp = "Eosinophilia   "
36290     Case 12: If v > InterpList(12) Then interp = "Basophilia     "
36300     Case 13: If v > InterpList(13) Then interp = "Basophilia     "
36310     Case 14: If v > InterpList(14) Then interp = "Leucocytosis   "
36320     Case 15: If v < InterpList(15) Then interp = "Leucopaenia    "
36330     Case 16: If v > InterpList(16) Then interp = "Anisocytosis   "
36340     Case 17: If v > InterpList(17) Then interp = "Anisocytosis   "
36350     Case 18: If v < InterpList(18) Then interp = "Microcytosis   "
36360     Case 19: If v > InterpList(19) Then interp = "Macrocytosis   "
36370     Case 20: If v < InterpList(20) Then interp = "Hypochromia    "
36380     Case 21: If v < InterpList(21) Then interp = "Anaemia        "
36390     Case 22: If v > InterpList(22) Then interp = "Erythrocytosis "
36400     Case 23: If v < InterpList(23) Then interp = "Thrombopaenia  "
36410     Case 24: If v > InterpList(24) Then interp = "Thrombocytosis "
36420     End Select

36430     Exit Function

interp_Error:

          Dim strES As String
          Dim intEL As Integer

36440     intEL = Erl
36450     strES = Err.Description
36460     LogError "Shared", "interp", intEL, strES

End Function




Public Sub FillInterpTable()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String

36470     On Error GoTo FillInterpTable_Error

36480     sql = "select * from interp"

36490     Set tb = New Recordset
36500     RecOpenClient 0, tb, sql
36510     If Not tb.EOF Then
36520         For n = 0 To 24
36530             InterpList(n) = Val(tb(n))
36540         Next
36550     End If

          'InterpList(0) = "Neutropaenia   "
          'InterpList(1) = "Neutropaenia   "
          'InterpList(2) = "Neutrophilia   "
          'InterpList(3) = "Neutrophilia   "
          'InterpList(4) = "Lymphopaenia   "
          'InterpList(5) = "Lymphopaenia   "
          'InterpList(6) = "Lymphocytosis  "
          'InterpList(7) = "Lymphocytosis  "
          'InterpList(8) = "Monocytosis    "
          'InterpList(9) = "Monocytosis    "
          'InterpList(10) = "Eosinophilia   "
          'InterpList(11) = "Eosinophilia   "
          'InterpList(12) = "Basophilia     "
          'InterpList(13) = "Basophilia     "
          'InterpList(14) = "Leucocytosis   "
          'InterpList(15) = "Leucopaenia    "
          'InterpList(16) = "Anisocytosis   "
          'InterpList(17) = "Anisocytosis   "
          'InterpList(18) = "Microcytosis   "
          'InterpList(19) = "Macrocytosis   "
          'InterpList(20) = "Hypochromia    "
          'InterpList(21) = "Anaemia        "
          'InterpList(22) = "Erythrocytosis "
          'InterpList(23) = "Thrombopaenia  "
          'InterpList(24) = "Thrombocytosis "

36560     Exit Sub

FillInterpTable_Error:

          Dim strES As String
          Dim intEL As Integer

36570     intEL = Erl
36580     strES = Err.Description
36590     LogError "Shared", "FillInterpTable", intEL, strES, sql

End Sub

Function InterpH(ByVal Value As Single, _
                 ByVal Analyte As String, _
                 ByVal Sex As String, _
                 ByVal DoB As String) _
                 As String

          Dim sql As String
          Dim tb As Recordset
          Dim DaysOld As Long
          Dim SexSQL As String

36600     On Error GoTo InterpH_Error

36610     Select Case Left$(UCase$(Sex), 1)
          Case "M"
36620         SexSQL = "MaleLow as Low, MaleHigh as High "
36630     Case "F"
36640         SexSQL = "FemaleLow as Low, FemaleHigh as High "
36650     Case Else
36660         SexSQL = "FemaleLow as Low, MaleHigh as High "
36670     End Select

36680     If IsDate(DoB) Then

36690         DaysOld = Abs(DateDiff("d", Now, DoB))

36700         sql = "Select top 1 PlausibleLow, PlausibleHigh, " & _
                    SexSQL & _
                    "from HaemTestDefinitions where " & _
                    "AnalyteName = '" & Analyte & "' and AgeFromDays <= '" & DaysOld & "' " & _
                    "and AgeToDays >= '" & DaysOld & "' " & _
                    "order by AgeFromDays desc, AgeToDays asc"
36710     Else
36720         sql = "Select top 1 PlausibleLow, PlausibleHigh, " & _
                    SexSQL & _
                    "from HaemTestDefinitions where Analytename = '" & Analyte & "' " & _
                    "and AgeFromDays = '0' " & _
                    "and AgeToDays = '43830'"
36730     End If

36740     Set tb = New Recordset
36750     RecOpenClient 0, tb, sql
36760     If Not tb.EOF Then

36770         If Value > tb!PlausibleHigh Then
36780             InterpH = "X"
36790             Exit Function
36800         ElseIf Value < tb!PlausibleLow Then
36810             InterpH = "X"
36820             Exit Function
36830         End If

36840         If Value > tb!High Then
36850             InterpH = "H"
36860         ElseIf Value < tb!Low Then
36870             InterpH = "L"
36880         Else
36890             InterpH = " "
36900         End If
36910     Else
36920         InterpH = " "
36930     End If

36940     Exit Function

InterpH_Error:

          Dim strES As String
          Dim intEL As Integer

36950     intEL = Erl
36960     strES = Err.Description
36970     LogError "Shared", "InterpH", intEL, strES, sql


End Function



Function LowAge(ByVal Age As String) As Integer

36980     LowAge = False
36990     If Val(Age) <> 0 Then
37000         If InStr(Age, "D") Then
37010             LowAge = True
37020         ElseIf InStr(Age, "M") Then
37030             LowAge = True
37040         ElseIf Val(Age) < 15 Then
37050             LowAge = True
37060         End If
37070     End If

End Function

Public Function Initial2Upper(ByVal s As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim SurName As String
          Dim ForeName As String

37080     On Error GoTo Initial2Upper_Error

37090     s = StrConv(s, vbProperCase)

37100     For n = 1 To Len(s) - 1
37110         If Mid$(s, n, 1) = "'" Then
37120             s = Left$(s, n) & UCase$(Mid$(s, n + 1, 1)) & Mid$(s, n + 2)
37130         End If
37140         If n > 1 Then
37150             If Mid$(s, n - 1, 1) = "M" And Mid$(s, n, 1) = "c" Then
37160                 s = Left$(s, n) & UCase$(Mid$(s, n + 1, 1)) & Mid$(s, n + 2)
37170             End If
37180         End If
37190     Next

37200     Initial2Upper = s

37210     Exit Function

37220     s = Trim$(s)
37230     If s = "" Then
37240         Initial2Upper = ""
37250         Exit Function
37260     End If

37270     s = LCase$(s)
37280     s = UCase$(Left$(s, 1)) & Mid$(s, 2)

37290     If Len(s) > 1 Then
37300         If Left$(s, 1) = "O" Then
37310             Mid$(s, 2, 1) = UCase$(Mid$(s, 2, 1))
37320         End If
37330     End If

37340     For n = 1 To Len(s) - 1
37350         If Mid$(s, n, 1) = " " Or Mid$(s, n, 1) = "'" Then
37360             s = Left$(s, n) & UCase$(Mid$(s, n + 1, 1)) & Mid$(s, n + 2)
37370         End If
37380         If n > 1 Then
37390             If Mid$(s, n, 1) = "c" And Mid$(s, n - 1, 1) = "M" Then
37400                 s = Left$(s, n) & UCase$(Mid$(s, n + 1, 1)) & Mid$(s, n + 2)
37410             End If
37420         End If
37430     Next

37440     SurName = ParseSurName(s)
37450     ForeName = ParseForeName(s)

37460     sql = "Select * from NameExclusions where " & _
                "GivenName = '" & AddTicks(SurName) & "'"
37470     Set tb = New Recordset
37480     RecOpenServer 0, tb, sql
37490     If Not tb.EOF Then
37500         SurName = tb!ReportName & ""
37510     Else
37520         sql = "Select * from NameExclusions where " & _
                    "GivenName = '" & AddTicks(ForeName) & "'"
37530         Set tb = New Recordset
37540         RecOpenServer 0, tb, sql
37550         If Not tb.EOF Then
37560             ForeName = tb!ReportName & ""
37570         End If
37580     End If

37590     Initial2Upper = SurName & " " & ForeName

37600     Exit Function

Initial2Upper_Error:

          Dim strES As String
          Dim intEL As Integer

37610     intEL = Erl
37620     strES = Err.Description
37630     LogError "Shared", "Initial2Upper", intEL, strES, sql


End Function

Public Function IsRoutine() As Boolean

      'Returns True if time now is between
      '09:30 and 16:30 Mon to Fri
      'else returns False

37640     IsRoutine = False

37650     If Weekday(Now) <> vbSaturday And Weekday(Now) <> vbSunday Then
37660         If TimeValue(Now) > TimeValue("09:29") And _
                 TimeValue(Now) < TimeValue("16:31") Then
37670             IsRoutine = True
37680         End If
37690     End If

End Function


Public Function Convert62Date(ByVal s As String, _
                              ByVal Direction As Integer) _
                              As String

          Dim D As String

37700     If Len(s) <> 6 Then
37710         Convert62Date = s
37720         Exit Function
37730     End If

37740     D = Left$(s, 2) & "/" & Mid$(s, 3, 2) & "/" & Right$(s, 2)
37750     If IsDate(D) Then
37760         Select Case Direction
              Case BACKWARD:
37770             If DateValue(D) > DateValue(Now) Then
37780                 D = DateAdd("yyyy", -100, D)
37790             End If
37800             Convert62Date = Format$(D, "dd/mm/yyyy")
37810         Case FORWARD:
37820             If DateValue(D) < Now Then
37830                 D = DateAdd("yyyy", 100, D)
37840             End If
37850             Convert62Date = Format$(D, "dd/mm/yyyy")
37860         Case gDONTCARE:
37870             Convert62Date = Format$(D, "dd/mm/yyyy")
37880         End Select
37890     Else
37900         Convert62Date = s
37910     End If

End Function




Public Sub PrintResultHaemWin(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Ward As String
          Dim Clin As String
          Dim GP As String

37920     On Error GoTo PrintResultHaemWin_Error

37930     GetWardClinGP SampleID, Ward, Clin, GP

37940     sql = "Select * from PrintPending where " & _
                "Department = 'H' " & _
                "and SampleID = '" & SampleID & "'"
37950     Set tb = New Recordset
37960     RecOpenClient 0, tb, sql
37970     If tb.EOF Then
37980         tb.AddNew
37990     End If
38000     tb!SampleID = SampleID
38010     tb!Ward = Ward
38020     tb!Clinician = Clin
38030     tb!GP = GP
38040     tb!Department = "H"
38050     tb!Initiator = UserName
38060     tb.Update
38070     tb.Close

38080     Set tb = Nothing

38090     sql = "Update HaemResults set " & _
                "Valid = 1, " & _
                "Printed = 1 " & _
                "where SampleID = '" & SampleID & "'"
38100     Cnxn(0).Execute sql

38110     Exit Sub

PrintResultHaemWin_Error:

          Dim strES As String
          Dim intEL As Integer

38120     intEL = Erl
38130     strES = Err.Description
38140     LogError "Shared", "PrintResultHaemWin", intEL, strES, sql

End Sub


Public Function TechnicianCodeFor(ByVal CodeOrName As String)

          Dim tb As Recordset
          Dim sql As String

38150     On Error GoTo TechnicianCodeFor_Error

38160     CodeOrName = Trim$(AddTicks(CodeOrName))

38170     sql = "Select Code from Users where " & _
                "Name = '" & CodeOrName & "' " & _
                "or Code = '" & CodeOrName & "'"
38180     Set tb = New Recordset
38190     RecOpenServer 0, tb, sql
38200     If Not tb.EOF Then
38210         TechnicianCodeFor = tb!Code & ""
38220     End If

38230     Exit Function

TechnicianCodeFor_Error:

          Dim strES As String
          Dim intEL As Integer

38240     intEL = Erl
38250     strES = Err.Description
38260     LogError "Shared", "TechnicianCodeFor", intEL, strES, sql


End Function


Public Function TechnicianNameFor(ByVal CodeOrName As String)

          Dim tb As Recordset
          Dim sql As String

38270     On Error GoTo TechnicianNameFor_Error

38280     CodeOrName = Trim$(AddTicks(CodeOrName))

38290     sql = "Select Name from Users where " & _
                "Name = '" & CodeOrName & "' " & _
                "or Code = '" & CodeOrName & "'"
38300     Set tb = New Recordset
38310     RecOpenServer 0, tb, sql
38320     If Not tb.EOF Then
38330         TechnicianNameFor = tb!Name & ""
38340     End If

38350     Exit Function

TechnicianNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

38360     intEL = Erl
38370     strES = Err.Description
38380     LogError "Shared", "TechnicianNameFor", intEL, strES, sql


End Function



Public Function TechnicianPassFor(ByVal CodeOrName As String)

          Dim tb As Recordset
          Dim sql As String

38390     On Error GoTo TechnicianPassFor_Error

38400     CodeOrName = Trim$(AddTicks(CodeOrName))

38410     sql = "Select PassWord from Users where " & _
                "Name = '" & CodeOrName & "' " & _
                "or Code = '" & CodeOrName & "'"
38420     Set tb = New Recordset
38430     RecOpenServer 0, tb, sql
38440     If Not tb.EOF Then
38450         TechnicianPassFor = tb!Password & ""
38460     End If

38470     Exit Function

TechnicianPassFor_Error:

          Dim strES As String
          Dim intEL As Integer

38480     intEL = Erl
38490     strES = Err.Description
38500     LogError "Shared", "TechnicianPassFor", intEL, strES, sql


End Function

Function ForeName(ByVal s As String) As String

          Dim p As Integer

38510     s = Trim$(s)
38520     If s = "" Then
38530         ForeName = ""
38540     Else
38550         p = InStr(s, " ")
38560         If p = 0 Then
38570             ForeName = ""
38580         Else
38590             ForeName = Trim$(Mid$(s, p))
38600         End If
38610     End If

End Function
Function SurName(ByVal s As String) As String

          Dim p As Integer

38620     s = Trim$(s)
38630     If s = "" Then
38640         SurName = ""
38650     Else
38660         p = InStr(s, " ")
38670         If p = 0 Then
38680             SurName = ""
38690         Else
38700             SurName = Trim$(Left$(s, p))
38710         End If
38720     End If

End Function

Public Function FormatString(strDestString As String, _
                             intNumChars As Integer, _
                             Optional strSeperator As String = "", _
                             Optional intAlign As PrintAlignContants = AlignLeft) As String

      '**************intAlign = 0 --> Left Align
      '**************intAlign = 1 --> Center Align
      '**************intAlign = 2 --> Right Align
          Dim intPadding As Integer
38730     intPadding = 0

38740     If Len(strDestString) > intNumChars Then
38750         FormatString = Mid(strDestString, 1, intNumChars) & strSeperator
38760     ElseIf Len(strDestString) < intNumChars Then
              Dim i As Integer
              Dim intStringLength As String
38770         intStringLength = Len(strDestString)
38780         intPadding = intNumChars - intStringLength

38790         If intAlign = PrintAlignContants.AlignLeft Then
38800             strDestString = strDestString & String(intPadding, " ")  '& " "
38810         ElseIf intAlign = PrintAlignContants.AlignCenter Then
38820             If (intPadding Mod 2) = 0 Then
38830                 strDestString = String(intPadding / 2, " ") & strDestString & String(intPadding / 2, " ")
38840             Else
38850                 strDestString = String((intPadding - 1) / 2, " ") & strDestString & String((intPadding - 1) / 2 + 1, " ")
38860             End If
38870         ElseIf intAlign = PrintAlignContants.AlignRight Then
38880             strDestString = String(intPadding, " ") & strDestString
38890         End If

38900         strDestString = strDestString & strSeperator
38910         FormatString = strDestString
38920     Else
38930         strDestString = strDestString & strSeperator
38940         FormatString = strDestString
38950     End If



End Function

Public Function SaveOptionSetting(ByVal Description As String, _
                                  ByVal Contents As String)

          Dim sql As String

38960     On Error GoTo SaveOptionSetting_Error

38970     sql = "IF EXISTS (SELECT * FROM Options WHERE " & _
                "           Description = '" & Description & "') " & _
                "  UPDATE Options " & _
                "  SET Contents = '" & AddTicks(Contents) & "' " & _
                "  WHERE Description = '" & Description & "' " & _
                "ELSE " & _
                "  INSERT INTO Options (Description, Contents) VALUES " & _
                "  ('" & Description & "', " & _
                "  '" & AddTicks(Contents) & "')"
38980     Cnxn(0).Execute sql

38990     Exit Function

SaveOptionSetting_Error:

          Dim strES As String
          Dim intEL As Integer

39000     intEL = Erl
39010     strES = Err.Description
39020     LogError "modOptionSettings", "SaveOptionSetting", intEL, strES, sql

End Function


Public Function GetOptionSetting(ByVal Description As String, _
                                 ByVal Default As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim RetVal As String

39030     On Error GoTo GetOptionSetting_Error

39040     sql = "SELECT Contents FROM Options WHERE " & _
                "Description = '" & Description & "'"
39050     Set tb = New Recordset
39060     RecOpenServer 0, tb, sql
39070     If tb.EOF Then
39080         RetVal = Default
39090     ElseIf Trim$(tb!Contents & "") = "" Then
39100         RetVal = Default
39110     Else
39120         RetVal = tb!Contents
39130     End If

39140     GetOptionSetting = RetVal

39150     Exit Function

GetOptionSetting_Error:

          Dim strES As String
          Dim intEL As Integer

39160     intEL = Erl
39170     strES = Err.Description
39180     LogError "modOptionSettings", "GetOptionSetting", intEL, strES, sql

End Function

Public Function ChangeComboHeight(cmb As ComboBox, numItemsToDisplay As Integer) As Boolean

          Dim newHeight As Long
          Dim itemHeight As Long


39190     itemHeight = SendMessage(cmb.hWnd, CB_GETITEMHEIGHT, 0, ByVal 0)
39200     newHeight = itemHeight * (numItemsToDisplay + 2)
39210     Call MoveWindow(cmb.hWnd, cmb.Left / 15, cmb.Top / 15, cmb.width / 15, newHeight, True)


End Function
Public Function FixComboWidth(Combo As ComboBox) As Boolean

          Dim i As Integer
          Dim ScrollWidth As Long

39220     With Combo
39230         For i = 0 To .ListCount
39240             If .Parent.TextWidth(.List(i)) > ScrollWidth Then
39250                 ScrollWidth = .Parent.TextWidth(.List(i))
39260             End If
39270         Next i
39280         FixComboWidth = SendMessage(.hWnd, CB_SETDROPPEDWIDTH, _
                                          ScrollWidth / 15 + 30, 0) > 0

39290     End With

End Function


'Public Sub SetDatesColour(ByVal f As Form)
'
'On Error GoTo SetDatesColour_Error
'
'With f
'  If CheckDateSequence(.dtSampleDate, .dtRecDate, .dtRunDate, .tSampleTime, .tRecTime) Then
'    .fraDate.ForeColor = vbButtonText
'    .fraDate.Font.Bold = False
'    .lblDate(0).ForeColor = vbButtonText
'    .lblDate(0).Font.Bold = False
'    .lblDate(1).ForeColor = vbButtonText
'    .lblDate(1).Font.Bold = False
'    .lblDate(2).ForeColor = vbButtonText
'    .lblDate(2).Font.Bold = False
'    .lblDateError.Visible = False
'  Else
'    .fraDate.ForeColor = vbRed
'    .fraDate.Font.Bold = True
'    .lblDate(0).ForeColor = vbRed
'    .lblDate(0).Font.Bold = True
'    .lblDate(1).ForeColor = vbRed
'    .lblDate(1).Font.Bold = True
'    .lblDate(2).ForeColor = vbRed
'    .lblDate(2).Font.Bold = True
'    .lblDateError.Visible = True
'  End If
'End With
'
'Exit Sub
'
'SetDatesColour_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "basShared", "SetDatesColour", intEL, strES
'
'End Sub

'Zyam 15-06-24
Public Function CheckDateSequence(ByRef SampleDate As String, _
          ByRef ReceivedDate As String, _
          ByRef Rundate As String, _
          ByRef SampleTime As String, _
          ByRef ReceivedTime As String) _
          As Boolean
        'Returns True if ok
        Dim RetVal As Boolean
        Dim SampleDateTime As Date
        Dim ReceivedDateTime As Date

39300     On Error GoTo CheckDateSequence_Error

39310     RetVal = True

          ' Ensure the dates and times are valid
39320     If Not IsDate(SampleDate) Or Not IsDate(ReceivedDate) Or Not IsDate(Rundate) Then
39330         RetVal = False
39340         GoTo CheckDateSequence_Exit
39350     End If

          ' Concatenate and convert to Date type
39360     SampleDateTime = CDate(SampleDate & " " & SampleTime)
39370     ReceivedDateTime = CDate(ReceivedDate & " " & ReceivedTime)

          ' Check date sequences
39380     If DateDiff("d", ReceivedDate, Rundate) < 0 Then
39390         RetVal = False
39400     End If

39410     If DateDiff("d", SampleDate, ReceivedDate) < 0 Then
39420         RetVal = False
39430     End If

39440     If DateDiff("n", SampleDateTime, ReceivedDateTime) < 0 Then
39450         RetVal = False
39460     End If

39470     If SampleDateTime > Now Then
39480         RetVal = False
39490     End If

39500     If ReceivedDateTime > Now Then
39510         RetVal = False
39520     End If

CheckDateSequence_Exit:
39530     CheckDateSequence = RetVal
39540     Exit Function

CheckDateSequence_Error:
          Dim strES As String
          Dim intEL As Integer
39550     intEL = Erl
39560     strES = Err.Description
39570     LogError "basShared", "CheckDateSequence", intEL, strES
39580     RetVal = False
39590     Resume CheckDateSequence_Exit

End Function

'Zyam 15-06-24



'Public Function CheckDateSequence(ByRef SampleDate As String, _
'                                  ByRef ReceivedDate As String, _
'                                  ByRef Rundate As String, _
'                                  ByRef SampleTime As String, _
'                                  ByRef ReceivedTime As String) _
'                                  As Boolean
'      'Returns True if ok
'          Dim RetVal As Boolean
'
'10        On Error GoTo CheckDateSequence_Error
'
'20        RetVal = True
'
'30        If DateDiff("d", ReceivedDate, Rundate) < 0 Then
'40            RetVal = False
'50        End If
'60        If DateDiff("d", SampleDate, ReceivedDate) < 0 Then
'70            RetVal = False
'80        End If
'90        If IsDate(SampleTime) And IsDate(ReceivedTime) Then
'100           If DateDiff("n", SampleDate & " " & SampleTime, ReceivedDate & " " & ReceivedTime) < 0 Then
'110               RetVal = False
'120           End If
'130       End If
'
'140       If Format$(SampleDate & " " & SampleTime, "dd/mmm/yyyy HH:mm") > Format$(Now, "dd/mmm/yyyy hh:mm") Then
'150           RetVal = False
'160       End If
'
'170       If Format$(ReceivedDate & " " & ReceivedTime, "dd/mmm/yyyy HH:mm") > Format$(Now, "dd/mmm/yyyy hh:mm") Then
'180           RetVal = False
'190       End If
'
'200       CheckDateSequence = RetVal
'
'210       Exit Function
'
'CheckDateSequence_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'220       intEL = Erl
'230       strES = Err.Description
'240       LogError "basShared", "CheckDateSequence", intEL, strES
'
'End Function

Public Function PrintText(ByVal Text As String, _
                          Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                          Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                          Optional FontColor As ColorConstants = vbBlack, _
                          Optional EnterCrLf As Boolean = False)


      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------
39600     With Printer
39610         .Font.size = FontSize
39620         .Font.Bold = FontBold
39630         .Font.Italic = FontItalic
39640         .Font.Underline = FontUnderLine
39650         .ForeColor = FontColor
39660         If EnterCrLf Then
39670             Printer.Print Text
39680         Else
39690             Printer.Print Text;
39700         End If
39710     End With
End Function


Public Function PrintTextRTB(rtb As RichTextBox, ByVal Text As String, _
                             Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                             Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                             Optional FontColor As ColorConstants = vbBlack)


      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------
39720     With rtb

39730         .SelFontSize = FontSize
39740         .SelBold = FontBold
39750         .SelItalic = FontItalic
39760         .SelUnderline = FontUnderLine
39770         .SelColor = FontColor
39780         .SelText = Text
39790     End With
End Function

Public Function EditGrid(ByVal g As MSFlexGrid, _
                         ByVal KeyCode As Integer, _
                         ByVal Shift As Integer) _
                         As Boolean

      'returns true if grid changed

          Dim ShiftDown As Boolean

39800     EditGrid = False

39810     If g.row < g.FixedRows Then
39820         Exit Function
39830     ElseIf g.Col < g.FixedCols Then
39840         Exit Function
39850     End If
39860     ShiftDown = (Shift And vbShiftMask) > 0


39870     Select Case KeyCode
          Case vbKeyA To vbKeyZ:
39880         If ShiftDown Then
39890             g = g & Chr(KeyCode)
39900             EditGrid = True
39910         Else
39920             g = g & Chr(KeyCode + 32)
39930             EditGrid = True
39940         End If

39950     Case vbKey0 To vbKey9:
39960         g = g & Chr(KeyCode)
39970         EditGrid = True

39980     Case vbKeyBack:
39990         If Len(g) > 0 Then
40000             g = Left$(g, Len(g) - 1)
40010             EditGrid = True
40020         End If

40030     Case &HBE, vbKeyDecimal:
40040         g = g & "."
40050         EditGrid = True

40060     Case vbKeySpace:
40070         g = g & " "
40080         EditGrid = True

40090     Case vbKeyNumpad0 To vbKeyNumpad9:
40100     Case vbKeyDelete:
40110     Case vbKeyLeft:
40120     Case vbKeyRight:
40130     Case vbKeyUp:
40140     Case vbKeyDown:
40150     Case vbKeyTab:
40160     End Select

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

40170 On Error GoTo LogAsViewed_Error

40180 sql = "INSERT INTO ViewedReports " & _
            "(Discipline, DateTime, Viewer, SampleID, Chart, Usercode) VALUES " & _
            "('" & Discipline & "', " & _
            "'" & Format(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
            "'" & AddTicks(UserName) & "', " & _
            "'" & SampleID & "', " & _
            "'" & Chart & "', " & _
            "'" & UserCode & "')"
40190 Cnxn(0).Execute sql
40200 Exit Sub

LogAsViewed_Error:

      Dim strES As String
      Dim intEL As Integer

40210 intEL = Erl
40220 strES = Err.Description
40230 LogError "modWardEnq", "LogAsViewed", intEL, strES, sql

End Sub
Public Function UserHasAuthority(ByVal MemberOf As String, SystemRole As String) As Boolean

40240     On Error GoTo UserHasAuthority_Error

          Dim UR As New UserRole
40250     If Not UR.GetUserRole(MemberOf, SystemRole, "") Then
40260         UserHasAuthority = False
40270     Else
40280         UserHasAuthority = (UR.Enabled > 0)
40290     End If

40300     Exit Function

UserHasAuthority_Error:

          Dim strES As String
          Dim intEL As Integer

40310     intEL = Erl
40320     strES = Err.Description
40330     LogError "Shared", "UserHasAuthority", intEL, strES

End Function
Public Sub ClearFGrid(ByVal g As MSFlexGrid)

40340     On Error GoTo ClearFGrid_Error

40350     With g
40360         .Rows = .FixedRows + 1
40370         .AddItem ""
40380         .RemoveItem .FixedRows
40390         .Visible = False
40400     End With

40410     Exit Sub

ClearFGrid_Error:

          Dim strES As String
          Dim intEL As Integer



40420     intEL = Erl
40430     strES = Err.Description
40440     LogError "basLibrary", "ClearFGrid", intEL, strES

End Sub
Public Function Swap_Year(ByVal Hyear As String) As String

40450     On Error GoTo Swap_Year_Error

40460     Swap_Year = Right(Hyear, 1) & Mid(Hyear, 3, 1) & Mid(Hyear, 2, 1) & Left(Hyear, 1)

40470     Exit Function

Swap_Year_Error:

          Dim strES As String
          Dim intEL As Integer



40480     intEL = Erl
40490     strES = Err.Description
40500     LogError "basHistology", "Swap_Year", intEL, strES


End Function
Public Sub FixG(ByVal g As MSFlexGrid)

40510     On Error GoTo FixG_Error

40520     With g
40530         .Visible = True
40540         If .Rows > .FixedRows + 1 And .TextMatrix(.FixedRows, 0) = "" Then
40550             .RemoveItem .FixedRows
40560         End If
40570     End With

40580     Exit Sub

FixG_Error:

          Dim strES As String
          Dim intEL As Integer

40590     intEL = Erl
40600     strES = Err.Description
40610     LogError "basLibrary", "FixG", intEL, strES

End Sub

Public Function AutoSel(cmb As ComboBox, KeyCode As Integer)
          
40620     Debug.Print KeyCode
       
40630     If KeyCode = vbEnter Then Exit Function
40640     If KeyCode = 8 Then Exit Function    'Backspace
40650     If KeyCode = 37 Then Exit Function  'left key
40660     If KeyCode = 38 Then Exit Function 'up arrow key
40670     If KeyCode = 39 Then Exit Function  'right key
40680     If KeyCode = 40 Then Exit Function  'down arrow key
40690     If KeyCode = 46 Then Exit Function  'delete key
40700     If KeyCode = 33 Then Exit Function  'page up key
40710     If KeyCode = 34 Then Exit Function  'page down key
40720     If KeyCode = 35 Then Exit Function  'end key
40730     If KeyCode = 36 Then Exit Function  'home key
          
          
          Dim Text As String
40740     Text = cmb.Text
          
          Dim i As Long
          Dim Temp As String
          
          
40750     For i = 0 To cmb.ListCount
40760         Temp = Left(cmb.List(i), Len(Text))
40770         If LCase(Temp) = LCase(Text) Then
40780             cmb.Text = cmb.List(i)
40790             cmb.ListIndex = i
40800             cmb.SelStart = Len(Text)
40810             cmb.SelLength = Len(cmb.List(i))
                  'Cmb.SetFocus
40820         End If
40830     Next
          
End Function

Function AutoComplete(cmbCombo As ComboBox, sKeyAscii As Integer, Optional bUpperCase As Boolean = True) As Integer
          Dim lngFind As Long, intPos As Integer, intLength As Integer
          Dim tStr As String
40840     On Error GoTo AutoComplete_Error

40850     If sKeyAscii = 8 Or sKeyAscii = 13 Then
40860         AutoComplete = sKeyAscii
40870         Exit Function
40880     End If

40890     With cmbCombo
40900         If sKeyAscii = 8 Then
40910             If .SelStart = 0 Then Exit Function
40920             .SelStart = .SelStart - 1
40930             .SelLength = 32000
40940             .SelText = ""
40950         Else
40960             intPos = .SelStart    '// save intial cursor position
40970             tStr = .Text    '// save string
40980             If bUpperCase = True Then
40990                 .SelText = UCase(Chr(sKeyAscii))    '// change string. (uppercase only)
41000             Else
41010                 .SelText = Chr(sKeyAscii)    '// change string. (leave case alone)
41020             End If
41030         End If

41040         lngFind = SendMessage(.hWnd, CB_FINDSTRING, 0, ByVal .Text)    '// Find string in combobox
41050         If lngFind = -1 Then    '// if string not found
41060             .Text = tStr    '// set old string (used for boxes that require charachter monitoring
41070             .SelStart = intPos    '// set cursor position
41080             .SelLength = (Len(.Text) - intPos)    '// set selected length
41090             AutoComplete = sKeyAscii    '// return 0 value to KeyAscii
41100             Exit Function

41110         Else    '// If string found
41120             intPos = .SelStart    '// save cursor position
41130             intLength = Len(.List(lngFind)) - Len(.Text)    '// save remaining highlighted text length
41140             .SelText = .SelText & Right(.List(lngFind), intLength)    '// change new text in string
                  '.Text = .List(lngFind)'// Use this instead of the above .Seltext line to set the text typed to the exact case of the item selected in the combo box.
41150             .SelStart = intPos    '// set cursor position
41160             .SelLength = intLength    '// set selected length
41170         End If
41180     End With


41190     Exit Function

AutoComplete_Error:

          Dim strES As String
          Dim intEL As Integer

41200     intEL = Erl
41210     strES = Err.Description
41220     LogError "Shared", "AutoComplete", intEL, strES


End Function

Public Function QueryCombo(cmbCombo As ComboBox) As String

          Dim strFoundString As String
          Dim lngFind As Long

41230     On Error GoTo QueryCombo_Error

41240     With cmbCombo
41250         lngFind = SendMessage(.hWnd, CB_FINDSTRING, 0, ByVal .Text)
41260         If lngFind = -1 Then
41270             QueryCombo = ""
41280         Else
41290             QueryCombo = .List(lngFind)

41300         End If
41310     End With


41320     Exit Function

QueryCombo_Error:

          Dim strES As String
          Dim intEL As Integer

41330     intEL = Erl
41340     strES = Err.Description
41350     LogError "Shared", "QueryCombo", intEL, strES

End Function

Public Function CheckScanViewLog(ByVal SampleID As String, ByVal Department As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

41360 On Error GoTo CheckScanViewLog_Error

41370 sql = "SELECT Count(*) AS Cnt FROM ScanViewLog WHERE SampleID = '" & SampleID & "' AND Department = '" & Department & "'"
41380 Set tb = New Recordset
41390 RecOpenServer 0, tb, sql
41400 CheckScanViewLog = (tb!Cnt > 0)

41410 Exit Function

CheckScanViewLog_Error:

       Dim strES As String
       Dim intEL As Integer

41420  intEL = Erl
41430  strES = Err.Description
41440  LogError "Shared", "CheckScanViewLog", intEL, strES, sql
          
End Function




'---------------------------------------------------------------------------------------
' Procedure : PrintPictureToFitPage
' Author    : XPMUser
' Date      : 20/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub PrintPictureToFitPage(pic As Picture)
      Dim PicRatio As Double
      Dim printerWidth As Double
      Dim printerHeight As Double
      Dim printerRatio As Double
      Dim printerPicWidth As Double
      Dim printerPicHeight As Double

      ' Determine if picture should be printed in landscape or portrait
      ' and set the orientation.
41450 On Error GoTo PrintPictureToFitPage_Error


41460 If pic.height >= pic.width Then
41470     Printer.Orientation = vbPRORPortrait    ' Taller than wide.
41480 Else
41490     Printer.Orientation = vbPRORLandscape    ' Wider than tall.
41500 End If
      ' Calculate device independent Width-to-Height ratio for picture.
41510 PicRatio = pic.width / pic.height
      ' Calculate the dimentions of the printable area in HiMetric.
41520 printerWidth = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbHimetric)
41530 printerHeight = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbHimetric)
      ' Calculate device independent Width to Height ratio for printer.
41540 printerRatio = printerWidth / printerHeight
      ' Scale the output to the printable area.
41550 If PicRatio >= printerRatio Then
          ' Scale picture to fit full width of printable area.
41560     printerPicWidth = Printer.ScaleX(printerWidth, vbHimetric, Printer.ScaleMode)
41570     printerPicHeight = Printer.ScaleY(printerWidth / PicRatio, vbHimetric, Printer.ScaleMode)
41580 Else
          ' Scale picture to fit full height of printable area.
41590     printerPicHeight = Printer.ScaleY(printerHeight, vbHimetric, Printer.ScaleMode)
41600     printerPicWidth = Printer.ScaleX(printerHeight * PicRatio, vbHimetric, Printer.ScaleMode)
41610 End If
      ' Print the picture using the PaintPicture method.
41620 Printer.Print ;
41630 Printer.PaintPicture pic, 0, 0, printerPicWidth, printerPicHeight
41640 Printer.EndDoc

41650 Exit Sub


PrintPictureToFitPage_Error:

      Dim strES As String
      Dim intEL As Integer

41660 intEL = Erl
41670 strES = Err.Description
41680 LogError "Shared", "PrintPictureToFitPage", intEL, strES
End Sub

Public Function VI(KeyAscii As Integer, _
                   iv As InputValidation, _
                   Optional NextFieldOnEnter As Boolean) As Integer

          Dim sTemp As String

41690     sTemp = Chr$(KeyAscii)
41700     If KeyAscii = 13 Then    'Enter Key
41710         If NextFieldOnEnter = True Then
41720             VI = 9    'Return Tab Keyascii if User Selected NextFieldOnEnter Option
41730         Else
41740             VI = 13
41750         End If
41760         Exit Function
41770     ElseIf KeyAscii = 8 Then    'BackSpace
41780         VI = 8
41790         Exit Function
41800     End If

          ' turn input to upper case

41810     Select Case iv
          Case InputValidation.NumericFullStopDash:
41820         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", "-"
41830             VI = Asc(sTemp)
41840         Case Else
41850             VI = 0
41860         End Select

41870     Case InputValidation.ivSampleID
41880         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
41890             VI = Asc(sTemp)
41900         Case "A" To "Z"
41910             VI = Asc(sTemp)
41920         Case "a" To "z"
41930             VI = Asc(sTemp) - 32    'Convert to upper case
41940         Case Else
41950             VI = 0
41960         End Select

41970     Case InputValidation.NumericDotLessGreater
41980         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", ">", "<"
41990             VI = Asc(sTemp)
42000         Case Else
42010             VI = 0
42020         End Select

42030     Case InputValidation.AlphaNumeric
42040         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
42050             VI = Asc(sTemp)
42060         Case "A" To "Z"
42070             VI = Asc(sTemp)
42080         Case "a" To "z"
42090             VI = Asc(sTemp)
42100         Case Else
42110             VI = 0
42120         End Select

42130     Case InputValidation.AlphaNumericSpace
42140         Select Case sTemp
              Case " ", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "<", ">"
42150             VI = Asc(sTemp)
42160         Case "A" To "Z"
42170             VI = Asc(sTemp)
42180         Case "a" To "z"
42190             VI = Asc(sTemp)
42200         Case Else
42210             VI = 0
42220         End Select

42230     Case InputValidation.Char
42240         Select Case sTemp
              Case " ", "-"
42250             VI = Asc(sTemp)
42260         Case "A" To "Z"
42270             VI = Asc(sTemp)
42280         Case "a" To "z"
42290             VI = Asc(sTemp)
42300         Case Else
42310             VI = 0
42320         End Select

42330     Case InputValidation.YorN
42340         sTemp = UCase(Chr$(KeyAscii))
42350         Select Case sTemp
              Case "Y", "N"
42360             VI = Asc(sTemp)
42370         Case Else
42380             VI = 0
42390         End Select

42400     Case InputValidation.AlphaNumeric_NoApos
42410         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", _
                   " ", "/", ";", ":", "\", "-", ">", "<", "(", ")", "@", _
                   "%", "!", """", "+", "^", "~", "`", "Ç", "´", "Ã", "Á", _
                   "Â", "È", "É", "Ê", "Ì", "Í", "Î", "Ò", "Ó", "Ô", "Õ", _
                   "Ù", "Ú", "Û", "Ü", "à", "á", "â", "ã", "ç", "è", "é", _
                   "ê", "ì", "í", "î", "ò", "ó", "ô", "õ", "ö", "ù", "ú", _
                   "û", "ü", "Æ", "æ", ",", "?"
42420             VI = Asc(sTemp)
42430         Case "A" To "Z"
42440             VI = Asc(sTemp)
42450         Case "a" To "z"
42460             VI = Asc(sTemp)
42470         Case Else
42480             VI = 0
42490         End Select

42500     Case InputValidation.AlphaNumeric_AllowApos
42510         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", " ", "'"
42520             VI = Asc(sTemp)
42530         Case "A" To "Z"
42540             VI = Asc(sTemp)
42550         Case "a" To "z"
42560             VI = Asc(sTemp)
42570         Case Else
42580             VI = 0
42590         End Select

42600     Case InputValidation.Numeric_Only
42610         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
42620             VI = Asc(sTemp)
42630         Case Else
42640             VI = 0
42650         End Select

42660     Case InputValidation.AlphaOnly
42670         Select Case sTemp
              Case "A" To "Z"
42680             VI = Asc(sTemp)
42690         Case "a" To "z"
42700             VI = Asc(sTemp)
42710         Case Else
42720             VI = 0
42730         End Select

42740     Case InputValidation.NumericSlash
42750         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/"
42760             VI = Asc(sTemp)
42770         Case Else
42780             VI = 0
42790         End Select

42800     Case InputValidation.AlphaAndSpaceonly
42810         Select Case sTemp
              Case " "
42820             VI = Asc(sTemp)
42830         Case "A" To "Z"
42840             VI = Asc(sTemp)
42850         Case "a" To "z"
42860             VI = Asc(sTemp)
42870         Case Else
42880             VI = 0
42890         End Select

42900     Case InputValidation.CharNumericDashSlash
42910         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/", "-"
42920             VI = Asc(sTemp)
42930         Case "A" To "Z"
42940             VI = Asc(sTemp)
42950         Case "a" To "z"
42960             VI = Asc(sTemp) - 32    'Convert to upper case
42970         Case Else
42980             VI = 0
42990         End Select

43000     Case InputValidation.AlphaAndSpaceApos
43010         Select Case sTemp
              Case " ", "'"
43020             VI = Asc(sTemp)
43030         Case "A" To "Z"
43040             VI = Asc(sTemp)
43050         Case "a" To "z"
43060             VI = Asc(sTemp)
43070         Case Else
43080             VI = 0
43090         End Select

43100     Case InputValidation.DecimalNumericOnly
43110         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "."
43120             VI = Asc(sTemp)
43130         Case Else
43140             VI = 0
43150         End Select

43160     Case InputValidation.CharNumericDashSlashFullStop
43170         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/", "-", "."
43180             VI = Asc(sTemp)
43190         Case "A" To "Z"
43200             VI = Asc(sTemp)
43210         Case "a" To "z"
43220             VI = Asc(sTemp)
43230         Case Else
43240             VI = 0
43250         End Select

43260     Case InputValidation.NumericDWMY
43270         Select Case sTemp
              Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "D", "M", "Y", "d", "m", "y", "w", "W", "s", "S"
43280             VI = Asc(sTemp)
43290         Case Else
43300             VI = 0
43310         End Select
          
43320     Case InputValidation.AlphaAndSpaceAposDash
43330         Select Case sTemp
              Case " ", "'", "-"
43340             VI = Asc(sTemp)
43350         Case "A" To "Z"
43360             VI = Asc(sTemp)
43370         Case "a" To "z"
43380             VI = Asc(sTemp)
43390         Case Else
43400             VI = 0
43410         End Select

43420     End Select

43430     If VI = 0 Then Beep

End Function


Public Function GetValidatorUser(ByVal strSID As String, _
                        ByVal strDept As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim strTable As String
          Dim strField As String

43440     On Error GoTo GetValidatorUser_Error

43450 If strDept = "B" Then
43460     strTable = "BioResults"
43470     strField = "Operator"
43480 ElseIf strDept = "H" Then
43490     strTable = "HaemResults"
43500     strField = "Operator"
43510 ElseIf strDept = "D" Then
43520     strTable = "CoagResults"
43530     strField = "UserName"
43540 End If

43550 sql = "select top 1 " & strField & " from " & strTable & " where sampleid = '" & strSID & "'"

43560 Set tb = New Recordset
43570 RecOpenServer 0, tb, sql
43580 If Not tb.EOF Then
43590     If strDept = "B" Or strDept = "H" Then
43600       GetValidatorUser = tb!Operator & ""
43610     ElseIf strDept = "D" Then
43620       GetValidatorUser = tb!UserName & "" 'TechnicianNameFor(tb!UserName & "")
43630     End If
43640 Else
43650     GetValidatorUser = ""
43660 End If

43670     Exit Function

GetValidatorUser_Error:

          Dim strES As String
          Dim intEL As Integer

43680     intEL = Erl
43690     strES = Err.Description
43700     LogError "Shared", "GetValidatorUser", intEL, strES, sql


End Function




'---------------------------------------------------------------------------------------
' Procedure : IsNotpadExists
' Author    : Masood
' Date      : 18/Nov/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function IsNotepadExists(SampleID As String, Dept As String) As Boolean
          Dim cond As String
          Dim sql As String
          Dim tb As New ADODB.Recordset

43710     On Error GoTo IsNotepadExists_Error


43720     IsNotepadExists = False


43730     sql = "Select * from PatientNotePad where "
43740     sql = sql & " SampleID = '" & SampleID & "'"
43750     If Dept <> "" Then
43760         sql = sql & " AND Descipline = '" & Dept & "'"
43770     End If
43780     Set tb = New Recordset
43790     RecOpenClient 0, tb, sql
43800     If tb.EOF = False Then
43810         IsNotepadExists = True
43820     End If


43830     Exit Function


IsNotepadExists_Error:

          Dim strES As String
          Dim intEL As Integer

43840     intEL = Erl
43850     strES = Err.Description
43860     LogError "Shared", "IsNotepadExists", intEL, strES, sql

End Function

'---------------------------------------------------------------------------------------
' Procedure : FndOptionSetting
' Author    : XPMUser
' Date      : 22/Jan/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function FndOptionSettingGlucose(TestCode As String) As String
43870     On Error GoTo FndOptionSetting_Error

          Dim sql As String
          Dim tb As Recordset

43880     sql = "SELECT Contents FROM Options WHERE Description Like 'GlucoseCode%' And Contents ='" & TestCode & "'"
43890     Set tb = New Recordset
43900     RecOpenServer 0, tb, sql
43910     If Not tb.EOF Then
43920         FndOptionSettingGlucose = tb!Contents & ""
43930         FndOptionSettingGlucose = RTrim(LTrim(FndOptionSettingGlucose))
43940     End If


43950     Exit Function


FndOptionSetting_Error:

          Dim strES As String
          Dim intEL As Integer

43960     intEL = Erl
43970     strES = Err.Description
43980     LogError "frmNewOrder", "FndOptionSetting", intEL, strES, sql

End Function



'---------------------------------------------------------------------------------------
' Procedure : FindFeildValue
' Author    : Masood
' Date      : 15/Jan/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function FindFeildValue(TableName As String, FeildName As String, Condition As String) As String

43990 On Error GoTo FindFeildValue_Error


      Dim sql As String
      Dim tb As Recordset

44000 sql = "SELECT " & FeildName & " FROM " & TableName & " " & Condition
44010 Set tb = New Recordset
44020 RecOpenServer 0, tb, sql
44030 If Not tb.EOF Then
44040     FindFeildValue = tb(0)
44050 End If



44060 Exit Function


FindFeildValue_Error:

      Dim strES As String
      Dim intEL As Integer

44070 intEL = Erl
44080 strES = Err.Description
44090 LogError "Shared", "FindFeildValue", intEL, strES, sql
End Function


'---------------------------------------------------------------------------------------
' Procedure : FindFeildValueByQuer
' Author    : Masood
' Date      : 15/Jan/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetMicroReportType(SampleIDWithOffset As String) As String
      Dim sql As String
      Dim tb As Recordset

44100 On Error GoTo GetMicroReportType_Error

44110 If Val(SampleIDWithOffset) = 0 Then
44120   Exit Function
44130 End If
44140 sql = "SELECT TOP 1 ReportType FROM REPORTS WHERE SampleID =" & SampleIDWithOffset & " ORDER BY PRINTTIME DESC"
44150 Set tb = New Recordset
44160 RecOpenServer 0, tb, sql
44170 If Not tb Is Nothing Then
44180     If Not tb.EOF Then
44190         GetMicroReportType = ConvertNull(tb!ReportType, "") & ""
44200     Else
44210         GetMicroReportType = ""
44220      End If
44230 End If


44240 Exit Function


GetMicroReportType_Error:

      Dim strES As String
      Dim intEL As Integer

44250 intEL = Erl
44260 strES = Err.Description
44270 LogError "Shared", "GetMicroReportType", intEL, strES, sql

End Function

'---------------------------------------------------------------------------------------
' Procedure : IsWardInternal
' Author    : m
' Date      : 3/2/2017
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function IsWardInternal(Ward As String) As Boolean

          Dim tw As Recordset
          Dim sql As String




44280    On Error GoTo IsWardInternal_Error

44290     sql = "SELECT Location " & _
                "FROM Wards " & _
              " WHERE Text = '" & AddTicks(Ward) & "'"
44300     Set tw = New Recordset
44310     RecOpenServer 0, tw, sql
44320     Do While Not tw.EOF
44330         If UCase(tw!Location) = UCase("In-House") Then
44340             IsWardInternal = True
44350         End If
44360         tw.MoveNext
44370     Loop




44380    Exit Function

IsWardInternal_Error:
      Dim strES As String
      Dim intEL As Integer

44390 intEL = Erl
44400 strES = Err.Description
44410 LogError "Shared", "IsWardInternal", intEL, strES, sql

End Function
Public Sub AddActivity(ByVal SampleID As String, ByVal ActionType As String, ByVal Action As String, Optional ByVal SubmissionID As String, Optional ByVal PatientID As String, _
                       Optional ByVal Reason As String, Optional ByVal Notes As String)
          Dim tb As New Recordset
          Dim sql As String
44420 On Error GoTo AddActivity_Error

44430 sql = "Select * from ActivityLog"
44440 Set tb = New Recordset

44450 RecOpenServer 0, tb, sql

44460 tb.AddNew
44470 tb!SampleID = SampleID
44480 If SampleID <> "" And (SubmissionID = "" Or PatientID = "") Then
44490   Call getChartFromSampleID(SampleID, PatientID)
44500 End If
44510 tb!ActionType = ActionType
44520 tb!Action = Action
      '120   tb!SubmissionID = IIf(IsMissing(SubmissionID), "", SubmissionID)
44530 tb!PatientID = IIf(IsMissing(PatientID), "", PatientID)
44540 tb!Reason = IIf(IsMissing(Reason), "", Reason)
44550 tb!Notes = IIf(IsMissing(Notes), "", Notes)
44560 tb!UserName = UserName
44570 tb!DateTimeOfRecord = Format(Now, "dd/mmm/yyyy hh:mm:ss")
44580 tb!MachineName = vbGetComputerName
44590 tb!ApplicationName = "NetAcquire LIS"
44600 tb!ApplicationVersion = App.Major & "-" & App.Minor & "-" & App.Revision
44610 tb!Createdby = UserName
44620 tb.Update
44630 Exit Sub

AddActivity_Error:
          Dim strES As String
          Dim intEL As Integer

44640 intEL = Erl
44650 strES = Err.Description
44660 LogError "Shared", "AddActivity", intEL, strES, sql

End Sub
Public Sub getChartFromSampleID(ByVal strSID As String, ByRef strCHART As String)
      Dim sql As String
      Dim tb As Recordset

44670 On Error GoTo getChartFromSampleID_Error

44680 sql = "SELECT  chart FROM Demographics WHERE SampleID  = '" & strSID & "'"
44690 Set tb = New Recordset
44700 RecOpenServer 0, tb, sql
44710 If Not (tb.EOF) Then
44720   strCHART = tb!Chart & ""
44730 End If

44740 Exit Sub

getChartFromSampleID_Error:

       Dim strES As String
       Dim intEL As Integer

44750  intEL = Erl
44760  strES = Err.Description
44770  LogError "Shared", "getChartFromSampleID", intEL, strES, sql

End Sub

Public Function GetOCMMapping(ByVal MappingType As String, ByVal TargetHospital As String, ByVal SourceValue As String) As String

      Dim sql As String
      Dim tb As Recordset

44780 On Error GoTo GetOCMMapping_Error

44790 sql = "SELECT TargetValue FROM ocmMapping WHERE MappingType = '" & MappingType & "' AND TargetHospital = '" & TargetHospital & "' AND SourceValue = '" & SourceValue & "' "
44800 Set tb = New Recordset
44810 RecOpenServer 0, tb, sql
44820 If tb.EOF Then
44830     GetOCMMapping = ""
44840 Else
44850     GetOCMMapping = tb!TargetValue & ""
44860 End If

44870 Exit Function
GetOCMMapping_Error:
         
44880 LogError "Shared", "GetOCMMapping", Erl, Err.Description, sql
End Function
                        
'+++ Junaid
Public Function ConvertNull(Data As Variant, Default As Variant) As Variant
44890     On Error GoTo ERROR_ConvertNull
44900     If IsNull(Data) = True Then
44910         ConvertNull = Default
44920     Else
44930         ConvertNull = Data
44940     End If
44950     Exit Function
ERROR_ConvertNull:
44960     LogError "Shared", "ConvertNull", Erl, Err.Description
End Function
'--- Junaid

'+++ Junaid 21-05-2024
Public Sub SaveAuditDemo(ByVal strSID As String)
          Dim sql As String

44970     On Error GoTo SaveAuditDemo_Error

44980     sql = "Insert Into demographicsAudit (SampleID, Chart, PatName, Age, Sex, ForHaem, ForBio, TimeTaken, ForHba1c,"
44990     sql = sql & " ForFerritin, ForPSA, Source, RunDate, DoB, Addr0, Addr1, Ward, Clinician, GP, SampleDate, HaemComments, "
45000     sql = sql & " BioComments, HaemComments2, HaemComments3, HaemComments4, ClDetails, Hospital, BioComments1, RooH, FAXed, "
45010     sql = sql & " ForCoag, ForESR, NinNumber, Fasting, OnWarfarin, DateTimeDemographics, DateTimeHaemPrinted, DateTimeBioPrinted, "
45020     sql = sql & " DateTimeCoagPrinted, Pregnant, AandE, NOPAS, RecDate, ForImm, RecordDateTime, Operator, ForBGA, Category, "
45030     sql = sql & " ForHisto, ForCyto, HistoValid, CytoValid, Mrn, ForExt, ForEnd, ForPgp, Username, Urgent, Valid, HYear, "
45040     sql = sql & " ForMicro, ForSemen, SentToEMedRenal, AssID, SurName, ForeName, ExtSampleID, Healthlink, MicroHealthLinkReleaseTime, LabNo, ArchivedBy, ArchiveDateTime) "
45050     sql = sql & " (Select SampleID, Chart, PatName, Age, Sex, ForHaem, ForBio, TimeTaken, ForHba1c, ForFerritin, ForPSA, Source, "
45060     sql = sql & " RunDate, DoB, Addr0, Addr1, Ward, Clinician, GP, SampleDate, HaemComments, BioComments, HaemComments2, "
45070     sql = sql & " HaemComments3, HaemComments4, ClDetails, Hospital, BioComments1, RooH, FAXed, ForCoag, ForESR, NinNumber, "
45080     sql = sql & " Fasting, OnWarfarin, DateTimeDemographics, DateTimeHaemPrinted, DateTimeBioPrinted, DateTimeCoagPrinted, "
45090     sql = sql & " Pregnant, AandE, NOPAS, RecDate, ForImm, RecordDateTime, Operator, ForBGA, Category, ForHisto, ForCyto, "
45100     sql = sql & " HistoValid, CytoValid, Mrn, ForExt, ForEnd, ForPgp, Username, Urgent, Valid, HYear, ForMicro, ForSemen, "
45110     sql = sql & " SentToEMedRenal, AssID, SurName, ForeName, ExtSampleID, Healthlink, MicroHealthLinkReleaseTime, LabNo, "
45120     sql = sql & " '" & UserName & "', GetDate() From demographics Where SampleID = '" & strSID & "')"
45130     Cnxn(0).Execute sql

45140     Exit Sub

SaveAuditDemo_Error:

          Dim strES As String
          Dim intEL As Integer

45150     intEL = Erl
45160     strES = Err.Description
45170     LogError "Shared", "SaveAuditDemo", intEL, strES, sql

End Sub
'--- Junaid
