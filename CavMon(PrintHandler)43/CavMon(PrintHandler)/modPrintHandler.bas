Attribute VB_Name = "Module1"
Option Explicit

'API's Function Declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As String) As Long
'API Constants
Public Const GWL_STYLE = -16
Public Const WS_DISABLED = &H8000000
Public Const WM_CANCELMODE = &H1F
Public Const WM_CLOSE = &H10
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public HospName(0 To 0) As String
'Public dbConnect As String
'Public dbConnectBB As String
Public Hosp As String

Public Const MaxAgeToDays As Long = 43830

Public Const gVALID = 1
Public Const gNOTVALID = 2
Public Const gPRINTED = 1
Public Const gNOTPRINTED = 2
Public Const gDONTCARE = 0
Public Const gNOCHANGE = 3

Public Cnxn(0 To 0) As Connection
Public CnxnBB(0 To 0) As Connection
'Public CnxnRemote() As Connection
'Public CnxnRemoteBB() As Connection

Public gData(1 To 365, 1 To 3) As Variant    '(n,1)=rundate, (n,2)=INR, (n,3)=Warfarin

Public LatestINR As String

Public CurrentDose As String
Public pLatest As String
Public pEarliest As String

Public pLowerTarget As String
Public pUpperTarget As String
Public pCondition As String

Public pForcePrintTo As String

Public Type PrintLine
    Analyte As String * 16
    Analyte20 As String * 20
    Result As String * 6
    Flag As String * 3
    Units As String * 7
    NormalRange As String * 13
    Fasting As String * 9
    Comment As String
End Type

Public Type ReportToPrint
    Department As String
    SampleID As String
    Initiator As String
    Ward As String
    Clinician As String
    GP As String
    FaxNumber As String
    UsePrinter As String
    ThisIsCopy As Boolean
    SendCopyTo As String
    PrintAction As String
    'WardPrint As Boolean
End Type

Public gPrintCopyReport As Boolean

Public Enum PrintAlignContants
    AlignLeft = 0
    AlignCenter = 1
    AlignRight = 2
End Enum

Public OriginalPrinter As String

Public RP As ReportToPrint

Public sAppName As String
Public sAppPath As String

Public intOtherHospitalsInGroup As Integer

Public Const UserName As String = "PrintHandler"
Public Const UserCode As String = "PrH"

'PrintSplit 0 is General default
'           1 to 10 are user defined
'           11 is TDM (Gentamicin or Tobramicin)
'           12 is eGFR
Private PrintSplit(0 To 12) As Boolean
Public Function CommentsPresent(ByVal SampleID As String, ByVal Discipline As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CommentsPresent_Error

20    sql = "SELECT COUNT(*) Tot FROM Observations " & _
            "WHERE SampleID = '" & SampleID & "' " & _
            "AND Discipline = '" & Discipline & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    CommentsPresent = tb!Tot > 0

60    Exit Function

CommentsPresent_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "Module1", "CommentsPresent", intEL, strES, sql

End Function

Public Function IsInhibited(ByVal Discipline As String, ByVal ShortName As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo IsInhibited_Error

20    IsInhibited = False

30    sql = "SELECT COUNT(*) Tot FROM PrintInhibit WHERE " & _
            "Discipline = '" & Discipline & "' " & _
            "AND Parameter = '" & ShortName & "' " & _
            "AND SampleID = '" & RP.SampleID & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    IsInhibited = tb!Tot > 0

70    Exit Function

IsInhibited_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "Module1", "IsInhibited", intEL, strES, sql

End Function

Public Function SetPrinter(ByVal CHNameOfPrinter As String) As Boolean

      Dim TargetPrinter As String
      Dim xFound As Boolean
      Dim Px As Printer

10    On Error GoTo SetPrinter_Error

20    SetPrinter = False

30    OriginalPrinter = Printer.DeviceName
40    If pForcePrintTo = "" Then
50        xFound = False
60        TargetPrinter = PrinterName(CHNameOfPrinter)
70        For Each Px In Printers
80            If UCase$(Px.DeviceName) = TargetPrinter Then
90                Set Printer = Px
100               Printer.Print ;
110               xFound = True
120               Exit For
130           End If
140       Next
150       If Not xFound Then
160           LogError "Module1", "SetPrinter", 0, "Can't find " & TargetPrinter
170           Exit Function
180       End If
190   End If

200   SetPrinter = True

210   Exit Function

SetPrinter_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "Module1", "SetPrinter", intEL, strES
250   SetPrinter = False

End Function

Public Sub ReSetPrinter()

      Dim Px As Printer

10    On Error GoTo ReSetPrinter_Error

20    For Each Px In Printers
30        If UCase$(Px.DeviceName) = OriginalPrinter Then
40            Set Printer = Px
50            Exit For
60        End If
70    Next

80    Exit Sub

ReSetPrinter_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "Module1", "ReSetPrinter", intEL, strES

End Sub


Public Function GetEGFRComment(ByVal SampleID As String, ByRef S() As String) As Boolean
      'Returns True if Comment Present

      Dim CodeForEGFR As String
      Dim BRs As New BIEResults
      Dim BR As BIEResult

10    GetEGFRComment = False

20    CodeForEGFR = UCase$(GetOptionSetting("BioCodeForEGFR", "5555"))

30    Set BRs = BRs.Load("Bio", SampleID, "Results", gDONTCARE, gDONTCARE)
40    If Not BRs Is Nothing Then
50        For Each BR In BRs
60            If BR.Code = CodeForEGFR Then
                  ' If BR.Valid Then
70                GetEGFRComment = True
80                ReDim S(0 To 11) As String
90                S(0) = "eGFR Interpretation:"
100               Select Case Val(BR.Result)
                  Case Is >= 90:
110                   S(1) = "CKD Stage 1"
120                   S(2) = "eGFR >=90 Normal in the absence of other evidence of kidney damage."
130               Case 60 To 89:
140                   S(1) = "CKD Stage 2"
150                   S(2) = "eGFR 60-89 Slight decrease in GFR. Not CKD in absence of other evidence of kidney damage."
160               Case 45 To 59:
170                   S(1) = "CKD Stage 3A"
180                   S(2) = "eGFR 45-59 Moderate decrease in GFR with or without other evidence of kidney damage."
190               Case 30 To 44:
200                   S(1) = "CKD Stage 3B"
210                   S(2) = "eGFR 30-44 Moderate decrease in GFR with or without other evidence of kidney damage."
220               Case 15 To 29
230                   S(1) = "CKD Stage 4"
240                   S(2) = "eGFR 15-29 Severe decrease in GFR with or without other evidence of kidney damage."
250               Case Is < 15:
260                   S(1) = "CKD Stage 5"
270                   S(2) = "eGFR <15   Established renal failure."
280               End Select
290               S(3) = ""
300               S(4) = "The Laboratory uses the abbreviated four variable MDRD formula to derive the eGFR. The only"
310               S(5) = "correction the user need apply is multiply the result by 1.21 for patients of African origin."
320               S(6) = ""
330               S(7) = "Limitations of eGRF measurements:-"
340               S(8) = "eGFR is an estimate not a measurement and falls down in extremes, it is not useful in"
350               S(9) = "severely ill patients, those undergoing dialysis, in extremes of muscle mass or children."
360               S(10) = "It is subject to both the biological and analytical variability in creatinine measurement"
370               S(11) = ""
                  '    End If

380               Exit For
390           End If
400       Next
410   End If

End Function
Public Function GetAutoCommentso(ByVal ShortName As String) As String

      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo GetAutoComments_Error

20    RetVal = ""

30    sql = "SELECT 'Output' = " & _
            "CASE WHEN ISNUMERIC(R.Result) = 1 " & _
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
            "  ELSE '' " & _
            "END " & _
            "FROM AutoComments A JOIN BioResults R ON " & _
            "R.Code = (SELECT TOP 1 Code FROM BioTestDefinitions " & _
            "          WHERE ShortName = A.Parameter " & _
            "          AND InUse = 1) " & _
            "WHERE A.Discipline = 'Biochemistry' " & _
            "AND R.SampleID = '" & RP.SampleID & "' " & _
            "AND A.Parameter = '" & ShortName & "' ORDER BY A.ListOrder"

40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    Do While Not tb.EOF
70        RetVal = RetVal & tb!Output & vbCrLf
80        tb.MoveNext
90    Loop
100   RetVal = Trim(RetVal)

110   If Right$(RetVal, 2) = vbCrLf Then
120       RetVal = Left$(RetVal, Len(RetVal) - 2)
130   End If

140   GetAutoCommentso = RetVal

150   Exit Function

GetAutoComments_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "Module1", "GetAutoComments", intEL, strES, sql

End Function
Private Function isTDM(ByVal SampleID As String) As Boolean
      'Returns true if Result is either Gentamicin or Tobramicin

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo isTDM_Error

20    isTDM = False

30    sql = "SELECT COUNT(*) Tot FROM BioResults " & _
            "WHERE Code In (SELECT Contents FROM Options " & _
            "               WHERE Description = 'BioCodeForGentamicin' " & _
            "               OR Description = 'BioCodeForTobramicin') " & _
            "AND SampleID = '" & SampleID & "'"
40    Set tb = New Recordset
50    Set tb = Cnxn(0).Execute(sql)

60    isTDM = tb!Tot > 0

70    Exit Function

isTDM_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "Module1", "isTDM", intEL, strES

End Function

Private Function isEGFR(ByVal SampleID As String) As Boolean
      'Returns true if Result is eGFR

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo isEGFR_Error

20    isEGFR = False

30    sql = "SELECT COUNT(*) Tot FROM BioResults " & _
            "WHERE Code IN (SELECT Contents FROM Options " & _
            "               WHERE Description = 'BioCodeForEGFR') " & _
            "AND SampleID = '" & SampleID & "'"
40    Set tb = New Recordset
50    Set tb = Cnxn(0).Execute(sql)

60    isEGFR = tb!Tot > 0

70    Exit Function

isEGFR_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "Module1", "isEGFR", intEL, strES, sql

End Function


Private Sub PopulatePrintSplit(ByVal SampleID As String)

      Dim n As Integer
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo PopulatePrintSplit_Error

20    For n = 0 To 12
30        PrintSplit(n) = False
40    Next

50    sql = "SELECT DISTINCT COALESCE(T.PrintSplit, 0) PS " & _
            "FROM BioTestDefinitions T JOIN BioResults R " & _
            "ON T.Code = R.Code " & _
            "WHERE R.SampleID = '" & SampleID & "'"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    If Not tb.EOF Then
90        Do While Not tb.EOF
100           PrintSplit(tb!PS) = True
110           tb.MoveNext
120       Loop
130   Else
140       If CommentsPresent(SampleID, "Biochemistry") Then
150           PrintSplit(0) = True
160       End If
170   End If

      '120   If isTDM(SampleID) Then
      '130     PrintSplit(11) = True
      '140   End If

      '180   If isEGFR(SampleID) Then
      '190       PrintSplit(12) = True
      '200   End If

180   Exit Sub

PopulatePrintSplit_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "Module1", "PopulatePrintSplit", intEL, strES, sql

End Sub

Public Function PrinterName(ByVal strMappedTo As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrinterName_Error

20    sql = "Select * from Printers where " & _
            "MappedTo = '" & strMappedTo & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        PrinterName = Trim$(UCase$(tb!PrinterName & ""))
70    Else
80        PrinterName = ""
90    End If

100   Exit Function

PrinterName_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "Module1", "PrinterName", intEL, strES, sql


End Function

Public Sub LogForCoag(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LogForCoag_Error

20    sql = "Select * from Demographics where " & _
            "SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If tb.EOF Then
60        tb.AddNew
70        tb!SampleID = SampleID
80        tb!Fasting = 0
90        tb!ForESR = 0
100       tb!ForPSA = 0
110       tb!ForBio = 0
120       tb!ForHaem = 0
130       tb!Faxed = 0
140       tb!ForHbA1c = 0
150       tb!ForFerritin = 0
160       tb!RooH = 1
170   End If
180   tb!ForCoag = 1

190   tb.Update

200   Exit Sub

LogForCoag_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "Module1", "LogForCoag", intEL, strES, sql

End Sub








Public Sub PrintNoSexDoB(ByVal Sex As String, ByVal Dob As String)

10    If Not IsDate(Dob) And Trim$(Sex) = "" Then
20        Printer.ForeColor = vbBlue
30        Printer.CurrentY = 6950
40        Printer.Print Tab(24); "No Sex/DoB given. Normal ranges may not be relevant"
50    ElseIf Not IsDate(Dob) Then
60        Printer.ForeColor = vbBlue
70        Printer.CurrentY = 6950
80        Printer.Print Tab(24); "No DoB given. Normal ranges may not be relevant"
90    ElseIf Trim$(Sex) = "" Then
100       Printer.ForeColor = vbBlue
110       Printer.CurrentY = 6950
120       Printer.Print Tab(24); "No Sex given. Normal ranges may not be relevant"
130   End If

End Sub


Public Sub PrintRecord(ByVal VV As Integer)

Dim n          As Integer

Select Case RP.Department
    Case "A":
        PrintResultACClinic
    Case "B":
        PopulatePrintSplit RP.SampleID
        For n = 0 To 10
            If PrintSplit(n) Then
                RTFPrintBioSplit n                ' Apply CheckDisk
            End If
        Next
        'If PrintSplit(12) Then RTFPrintEGFR RP.SampleID
    Case "I":
        PrintResultBioSideBySide "Imm", VV        ' Apply CheckDisk
    Case "C"
        RTFPrintCoag
    Case "D"
        RTFPrintCoag  'force print all
    Case "G"
        RTFPrintGTT
    Case "H"
        RTFPrintHaem
    Case "J"
        RTFPrintHaemSpecific "Sickledex"
    Case "K"
        RTFPrintHaemSpecific "ESR"
    Case "P"
        RTFPrintHaemSpecific "MonoSpot"
    Case "L"
        RTFPrintHaemSpecific "Malaria"
    Case "M"
        RTFPrintResultMicro
    Case "Z"
        RTFPrintSAReport
    Case "R", "T"
        RTFPrintCreatinine  'R-Urine T-Serum Number Given
    Case "S"
        RTFPrintGlucoseSeries
    Case "U"
        RTFPrintUPro
End Select

End Sub

Public Function SaveOptionSetting(ByVal Description As String, _
                                  ByVal Contents As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveOptionSetting_Error

20    sql = "SELECT * FROM Options WHERE " & _
            "Description = '" & Description & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        tb.AddNew
70    End If
80    tb!Description = Description
90    tb!Contents = Contents
100   tb.Update

110   Exit Function

SaveOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "Module1", "SaveOptionSetting", intEL, strES, sql

End Function

Public Function GetOptionSetting(ByVal Description As String, _
                                 ByVal Default As String) As String

      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo GetOptionSetting_Error

20    sql = "SELECT Contents FROM Options WHERE " & _
            "Description = '" & Description & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        RetVal = Default
70    ElseIf Trim$(tb!Contents & "") = "" Then
80        RetVal = Default
90    Else
100       RetVal = tb!Contents
110   End If

120   GetOptionSetting = RetVal

130   Exit Function

GetOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "Module1", "GetOptionSetting", intEL, strES, sql

End Function


'Public Sub FaxRecord()
'
'          Dim X As Word.Document
'          Dim FAXIndex As String
'          Dim FaxFileName As String
'          Dim Recipient As String
'          Dim ZetaFaxFolderFilePathName As String
'          Dim ZetaFaxDocFilePathName As String
'          'Set x = CreateObject(Word.Document)
'
'
'10        On Error GoTo FaxRecord_Error
'          '
'20        Set X = New Word.Document
'
'30        Select Case RP.Department
'              'Case "A":    PrintResultACClinic
'              Case "B":
'40                If frmMain.optSecondPage Then
'50                    WordPrintResultBioVert X
'60                Else
'70                    WordPrintResultBioSideBySide X
'80                End If
'
'90            Case "C": WordPrintCoag X
'100           Case "D": WordPrintCoag X  'force print all
'                  'Case "G": PrintGTT
'110           Case "H": WordPrintResultHaem X
'                  'Case "M": PrintResultMicro
'                  'Case "R", "T": PrintCreatinine  'R-Urine T-Serum Number Given
'                  'Case "S": PrintGlucoseSeries
'120       End Select
'
'130       FAXIndex = Format(Now, "yymmddhhmmss") & RP.Department
'140       FaxFileName = "C:\FAX\FAX" & FAXIndex & ".DOC"
'
'150       ZetaFaxDocFilePathName = frmMain.lblDocument & "\FAX" & FAXIndex & ".DOC"
'160       ZetaFaxFolderFilePathName = frmMain.lblZSubmit & "\FAX" & FAXIndex & ".SUB"
'
'170       X.SaveAs ZetaFaxDocFilePathName
'180       X.SaveAs FaxFileName
'190       X.Close
'
'200       Set X = Nothing
'
'210       If IsTaskRunning(sAppName) Then
'220           Call EndTask(sAppName)
'230       End If
'
'240       Call Send_EMail_Notes("Laboratory Report", RP.FaxNumber, FaxFileName)
'
'250       If IsTaskRunning(sAppName) Then
'260           Call EndTask(sAppName)
'270       End If
'
'280       If Trim$(RP.GP) <> "" Then
'290           Recipient = RP.GP
'300       ElseIf Trim$(RP.Clinician) <> "" Then
'310           Recipient = RP.Clinician
'320       ElseIf Trim$(RP.Ward) <> "" Then
'330           Recipient = RP.Ward
'340       Else
'350           Recipient = "Unknown"
'360       End If
'370       SendZetaFax Recipient, RP.FaxNumber, ZetaFaxFolderFilePathName, ZetaFaxDocFilePathName
'
'380       Exit Sub
'
'FaxRecord_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'390       intEL = Erl
'400       strES = Err.Description
'410       LogError "Module1", "FaxRecord", intEL, strES
'
'End Sub

Public Function EndTask(sWindowName As String) As Integer

      Dim X As Long
      Dim TargetHwnd As Long

10    On Error GoTo EndTask_Error

20    EndTask = False

      'find handle of the application
30    TargetHwnd = FindWindow(0&, sWindowName)

40    If TargetHwnd = 0 Then
50        Exit Function
60    Else
          'close application
70        If Not (GetWindowLong(TargetHwnd, GWL_STYLE) And WS_DISABLED) Then
80            X = PostMessage(TargetHwnd, WM_CLOSE, 0, 0&)
90            DoEvents
100       End If
110   End If

120   EndTask = True

130   Exit Function

EndTask_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "Module1", "EndTask", intEL, strES

End Function
Public Function IsTaskRunning(sWindowName As String) As Boolean

      Dim hwnd As Long

      'get handle of the application
      'if handle is 0 the application is currently not running
10    On Error GoTo IsTaskRunning_Error

20    hwnd = FindWindow(0&, sWindowName)
30    If hwnd = 0 Then
40        IsTaskRunning = False
50        Exit Function
60    Else
70        IsTaskRunning = True
80    End If

90    Exit Function

IsTaskRunning_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Module1", "IsTaskRunning", intEL, strES

End Function
Private Sub Send_EMail_Notes(in_subject As String, _
                             in_recipients As String, _
                             Optional in_file_attachement_path)

      ' This routine is used to compose a new notes document and send the mail (for faxing in this case)
      ' There is no error checking to ensure that numbers are entered as oppossed to e-mail addresses,
      ' therefore the code can be used for either.
      '
      Dim notesApp As Object
      Dim notesDB As Object
      Dim notesDoc As Object
      Dim notesAttachment As Object
      Dim strServer As String
      Dim strFile As String

      ' Verify recipients are listed
10    On Error GoTo Send_EMail_Notes_Error

20    If in_recipients = "" Then Exit Sub

      ' Establish notes session with server, open mail file
30    Set notesApp = CreateObject("Notes.NotesSession")
40    strServer = notesApp.GetEnvironmentstring$("", "")
50    strFile = notesApp.GetEnvironmentstring$("MailFile", True)
60    Set notesDB = notesApp.GetDatabase(strServer, strFile)
70    If notesDB.IsOpen = False Then notesDB.OPENMAIL

      ' Compose a new notes document
80    Set notesDoc = notesDB.CreateDocument
90    With notesDoc
100       .SendTo = in_recipients
110       .Subject = in_subject
120       If Not IsMissing(in_file_attachement_path) Then
130           If Dir(in_file_attachement_path) <> "" Then
140               Set notesAttachment = .CreateRichTextItem("Attachment")
150               notesAttachment.EmbedObject 1454, "", in_file_attachement_path
160           End If
170       End If
180       .posteddate = Now
190       .Save True, True, True
200       .Send False
210   End With

220   Set notesAttachment = Nothing
230   Set notesDoc = Nothing
240   Set notesDB = Nothing
250   Set notesApp = Nothing

260   Exit Sub

Send_EMail_Notes_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "Module1", "Send_EMail_Notes", intEL, strES

End Sub
Public Sub PrintResultACClinic()

      Dim tb As Recordset
      Dim SampleDate As String
      Dim Rundate As String
      Dim sql As String
      Dim Dob As String

10    On Error GoTo PrintResultACClinic_Error

20    If Trim$(RP.SampleID) = "" Then Exit Sub

30    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql

60    If tb.EOF Then Exit Sub

70    If IsDate(tb!Dob) Then
80        Dob = Format(tb!Dob, "dd/mmm/yyyy")
90    Else
100       Dob = ""
110   End If

120   If IsDate(tb!SampleDate) Then
130       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
140   Else
150       SampleDate = ""
160   End If
170   If IsDate(tb!Rundate) Then
180       Rundate = Format(tb!Rundate, "dd/mmm/yyyy hh:mm")
190   Else
200       Rundate = ""
210   End If

220   If Not SetPrinter("CHCOAG") Then Exit Sub

230   frmMain.DrawPicture tb!Chart & ""

240   frmMain.pb.Picture = frmMain.pb.Image
250   PrintHeading "Haematology", tb!PatName & "", Dob, tb!Chart & "", _
                   tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

260   Printer.PaintPicture frmMain.pb.Picture, 500, 1800

270   Printer.CurrentY = 2500
280   Printer.Font.Size = 10
290   Printer.ForeColor = vbBlack
300   Printer.Print Tab(79); "Legend:-"
310   Printer.ForeColor = vbBlue
320   Printer.Print Tab(79); "INR Level"
330   Printer.ForeColor = vbRed
340   Printer.Font.Bold = False
350   Printer.Print Tab(79); "INR Target "
360   Printer.Font.Bold = True
370   Printer.Print Tab(79); pLowerTarget & " - " & pUpperTarget
380   Printer.ForeColor = vbGreen
390   Printer.Font.Bold = False
400   Printer.Print Tab(79); "Warfarin Dose "
410   Printer.Font.Bold = True
420   Printer.Print Tab(79); CurrentDose
430   Printer.ForeColor = vbBlack
440   Printer.Font.Bold = False
450   Printer.Print Tab(79); pCondition

460   Printer.CurrentY = 5600
470   Printer.Font.Size = 14
480   Printer.Font.Bold = True
490   Printer.ForeColor = vbBlack
500   Printer.Print Tab(5); "INR Result for " & pLatest & " = " & LatestINR

510   Printer.DrawWidth = 3
520   Printer.Line (500, 5500)-(6500, 6100), , B

530   PrintFooter "INR", RP.Initiator, SampleDate, Rundate

540   Printer.EndDoc

      '##############################

550   PrintHeading "Haematology", tb!PatName & "", Dob, tb!Chart & "", _
                   tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

560   Printer.Font.Bold = True
570   Printer.Font.Size = 16
580   Printer.Print
590   Printer.Print
600   Printer.Print
610   Printer.ForeColor = vbBlack
620   Printer.Print Tab(10); "Your next appointment is for"
630   Printer.Print
640   Printer.Print
650   Printer.Print Tab(17); "Your Warfarin dose is"
660   Printer.Print
670   Printer.Print
680   Printer.ForeColor = vbRed
690   Printer.Print " Keep this form safe and bring it to your next appointment."

700   Printer.ForeColor = vbBlack
710   Printer.Line (7500, 2500)-(9500, 3100), , B
720   Printer.Line (7500, 3550)-(9500, 4150), , B

730   Do While Printer.CurrentY < 7200
740       Printer.Print
750   Loop

760   Printer.ForeColor = vbRed
770   Printer.Font.Size = 4
780   Printer.Print String$(250, "-")

790   Printer.Font.Size = 16
800   Printer.Font.Bold = True
810   Printer.Print Tab(20); "INR REQUEST FORM"

820   Printer.EndDoc

830   ReSetPrinter

840   Exit Sub

PrintResultACClinic_Error:

      Dim strES As String
      Dim intEL As Integer

850   intEL = Erl
860   strES = Err.Description
870   LogError "Module1", "PrintResultACClinic", intEL, strES, sql

End Sub


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
30        vbGetComputerName = Left$(sTmp, lLen)
40    Else
50        vbGetComputerName = ""
60    End If

End Function

Public Function AddTicks(ByVal S As String) As String

10    AddTicks = Replace(S, "'", "''")

End Function

Public Function IsTableInDatabase(ByVal TableName As String) As Boolean

      Dim tbExists As Recordset
      Dim sql As String
      Dim RetVal As Boolean

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist
      'if it has a record then the table does exist.

10    On Error GoTo IsTableInDatabase_Error

20    sql = "SELECT name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = '" & TableName & "'"
30    Set tbExists = Cnxn(0).Execute(sql)

40    RetVal = True

50    If tbExists.EOF Then    'There is no table <TableName> in database
60        RetVal = False
70    End If
80    IsTableInDatabase = RetVal

90    Exit Function

IsTableInDatabase_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Module1", "IsTableInDatabase", intEL, strES, sql

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

10    On Error GoTo FillCommentLines_Error

20    For n = 1 To UBound(Comments)
30        Comments(n) = ""
40    Next

50    CurrentLine = 0
60    FullComment = Trim$(FullComment)
70    n = Len(FullComment)

80    For X = n - 1 To 1 Step -1
90        If Mid$(FullComment, X, 1) = vbCr Or Mid$(FullComment, X, 1) = vbLf Or Mid$(FullComment, X, 1) = vbTab Then
100           Mid$(FullComment, X, 1) = " "
110       End If
120   Next

130   For X = n - 3 To 1 Step -1
140       If Mid$(FullComment, X, 2) = "  " Then
150           FullComment = Left$(FullComment, X) & Mid$(FullComment, X + 2)
160       End If
170   Next
180   n = Len(FullComment)

190   Do While n > MaxLen
200       SpaceFound = False
210       For X = MaxLen To 1 Step -1
220           If Mid$(FullComment, X, 1) = " " Then
230               ThisLine = Left$(FullComment, X - 1)
240               FullComment = Mid$(FullComment, X + 1)

250               CurrentLine = CurrentLine + 1
260               If CurrentLine <= NumberOfLines Then
270                   Comments(CurrentLine) = ThisLine
280               End If
290               SpaceFound = True
300               Exit For
310           End If
320       Next
330       If Not SpaceFound Then
340           ThisLine = Left$(FullComment, MaxLen)
350           FullComment = Mid$(FullComment, MaxLen + 1)

360           CurrentLine = CurrentLine + 1
370           If CurrentLine <= NumberOfLines Then
380               Comments(CurrentLine) = ThisLine
390           End If
400       End If
410       n = Len(FullComment)
420   Loop

430   CurrentLine = CurrentLine + 1
440   If CurrentLine <= NumberOfLines Then
450       Comments(CurrentLine) = FullComment
460   End If

470   Exit Sub

FillCommentLines_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "Module1", "FillCommentLines", intEL, strES

End Sub


Function InterpH(ByVal Value As Single, _
                 ByVal Analyte As String, _
                 ByVal Sex As String, _
                 ByVal Dob As String) _
                 As String

      Dim sql As String
      Dim tb As Recordset
      Dim DaysOld As Long
      Dim SexSQL As String

10    On Error GoTo InterpH_Error

20    Select Case Left$(UCase$(Sex), 1)
      Case "M"
30        SexSQL = "MaleLow as Low, MaleHigh as High "
40    Case "F"
50        SexSQL = "FemaleLow as Low, FemaleHigh as High "
60    Case Else
70        SexSQL = "FemaleLow as Low, MaleHigh as High "
80    End Select

90    If IsDate(Dob) Then

100       DaysOld = Abs(DateDiff("d", Now, Dob))

110       sql = "Select top 1 PlausibleLow, PlausibleHigh, " & _
                SexSQL & _
                "from HaemTestDefinitions where " & _
                "AnalyteName = '" & Analyte & "' and AgeFromDays <= '" & DaysOld & "' " & _
                "and AgeToDays >= '" & DaysOld & "' " & _
                "order by AgeFromDays desc, AgeToDays asc"
120   Else
130       sql = "Select top 1 PlausibleLow, PlausibleHigh, " & _
                SexSQL & _
                "from HaemTestDefinitions where " & _
                "AgeFromDays = '0' " & _
                "and AgeToDays = '43830'"
140   End If

150   Set tb = New Recordset
160   RecOpenClient 0, tb, sql
170   If Not tb.EOF Then

180       If Value > tb!PlausibleHigh Then
190           InterpH = "X"
200           Exit Function
210       ElseIf Value < tb!PlausibleLow Then
220           InterpH = "X"
230           Exit Function
240       End If

250       If Value > tb!High Then
260           InterpH = "H"
270       ElseIf Value < tb!Low Then
280           InterpH = "L"
290       Else
300           InterpH = " "
310       End If
320   Else
330       InterpH = " "
340   End If

350   Exit Function

InterpH_Error:

      Dim strES As String
      Dim intEL As Integer

360   intEL = Erl
370   strES = Err.Description
380   LogError "Module1", "InterpH", intEL, strES, sql

End Function

Public Function HaemNormalRange(ByVal Analyte As String, _
                                ByVal Sex As String, _
                                ByVal Dob As String) _
                                As String

      Dim sql As String
      Dim tb As Recordset
      Dim DaysOld As Long
      Dim SexSQL As String
      Dim strFormat As String
      Dim strL As String * 4
      Dim strH As String * 4
      Dim strRange As String

10    On Error GoTo HaemNormalRange_Error

20    Select Case Left$(UCase$(Sex), 1)
      Case "M"
30        SexSQL = "MaleLow as Low, MaleHigh as High "
40    Case "F"
50        SexSQL = "FemaleLow as Low, FemaleHigh as High "
60    Case Else
70        SexSQL = "FemaleLow as Low, MaleHigh as High "
80    End Select

90    If IsDate(Dob) Then

100       DaysOld = Abs(DateDiff("d", Now, Dob))

110       sql = "Select top 1 PrintFormat, " & _
                SexSQL & _
                "from HaemTestDefinitions where " & _
                "AnalyteName = '" & Analyte & "' and AgeFromDays <= '" & DaysOld & "' " & _
                "and AgeToDays >= '" & DaysOld & "' " & _
                "order by AgeFromDays desc, AgeToDays asc"
120   Else
130       sql = "Select top 1 PrintFormat, " & _
                SexSQL & _
                "from HaemTestDefinitions where " & _
                "AnalyteName = '" & Analyte & "' and AgeFromDays = '0' " & _
                "and AgeToDays = '43830'"
140   End If

150   Set tb = New Recordset
160   RecOpenClient 0, tb, sql

170   strRange = "(    -    )"

180   If Not tb.EOF Then
190       Select Case tb!Printformat
          Case 0: strFormat = "0"
200       Case 1: strFormat = "0.0"
210       Case 2: strFormat = "0.00"
220       Case 3: strFormat = "0.000"
230       End Select

240       If tb!High <> 999 Then
250           RSet strL = Format(tb!Low, strFormat)
260           Mid$(strRange, 2, 4) = strL
270           LSet strH = Format(tb!High, strFormat)
280           Mid$(strRange, 7, 4) = strH
290       End If
300   Else
310       strRange = "(    -    )"
320   End If

330   HaemNormalRange = strRange

340   Exit Function

HaemNormalRange_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "Module1", "HaemNormalRange", intEL, strES, sql

End Function


Public Sub PrintFooter(ByVal Dept As String, _
                       ByVal Initiator As String, _
                       ByVal SampleDate As String, _
                       ByVal Rundate As String, _
                       Optional ByVal SplitName As String = "")

      Dim S As String
      Dim sql As String
      Dim tb As Recordset
      Dim DisciplineCode As String
      Dim FColour As Long

10    On Error GoTo PrintFooter_Error

20    Printer.Font.Name = "Courier New"

30    Printer.CurrentY = 7100

40    If gPrintCopyReport = 0 Then
50        Printer.Font.Size = 4
60        Printer.Font.Bold = True
70        Printer.Print String$(250, "-")
80    Else
90        Printer.Font.Size = 8
100       Printer.Font.Bold = True
110       S = "- THIS IS A COPY REPORT - NOT FOR FILING -"
120       S = S & S
130       S = S & "- THIS IS A COPY REPORT -"
140       Printer.ForeColor = vbRed
150       Printer.Print S
160   End If

170   Select Case Dept
      Case "Haematology":
180       FColour = vbRed
190       DisciplineCode = "J"
200   Case "Biochemistry", "Creat Clearance", "Glucose Series", "Gluc. Tolerance"
210       FColour = vbGreen
220       DisciplineCode = "I"
230   Case "Microbiology":
240       DisciplineCode = "N"
250       FColour = vbYellow
260   Case "Blood Transfusion":
270       FColour = vbRed
280       DisciplineCode = "N"
290   Case "Coagulation":
300       DisciplineCode = "K"
310       FColour = vbRed        'RGB(80, 46, 107) 'Purple
320   End Select

330   Printer.ForeColor = FColour
340   Printer.Font.Size = 10
350   Printer.Font.Bold = False

360   Printer.ForeColor = vbBlack
370   If Trim$(SplitName) <> "" Then
380       Printer.Print Left$(SplitName, 16);
390   Else
400       Printer.Print Dept;
410       If UCase$(Dept) = "HAEMATOLOGY" Then
420           Printer.Print " Whole Blood";
430       ElseIf UCase$(Dept) = "COAGULATION" Then
440           Printer.Print " Citrate Plasma";
450       End If
460   End If
470   Printer.ForeColor = FColour

480   If Format(SampleDate, "hh:mm") <> "00:00" Then
490       Printer.Print " Sample Date/Time:"; Format(SampleDate, "dd/mm/yy HH:nn");
500   Else
510       Printer.Print " Sample Date:"; Format(SampleDate, "dd/MM/yy");
520   End If

530   Printer.Print " Tested:"; Format(Rundate, "dd/mm/yyyy hh:mm");

540   If gPrintCopyReport = 0 Then
550       If Trim$(Initiator) <> "" Then
560           Printer.Print " Validated by "; TechnicianCodeFor(Initiator);
570       End If
580   Else
590       sql = "SELECT TOP 1 Viewer FROM ViewedReports WHERE " & _
                "SampleID = '" & RP.SampleID & "' " & _
                "AND Discipline = '" & DisciplineCode & "' " & _
                "AND DATEDIFF(minute, [datetime], getdate()) < 2 " & _
                "ORDER BY [DateTime] DESC"
600       Set tb = New Recordset
610       RecOpenServer 0, tb, sql
620       If Not tb.EOF Then
630           Printer.Print " Printed by "; tb!Viewer & "";
640       Else
650           Printer.Print " Printed by "; Left$(TechnicianCodeFor(Initiator), 14);
660       End If
670   End If

680   Exit Sub

PrintFooter_Error:

      Dim strES As String
      Dim intEL As Integer

690   intEL = Erl
700   strES = Err.Description
710   LogError "Module1", "PrintFooter", intEL, strES

End Sub

Public Sub PrintHeading(ByVal Dept As String, _
                        ByVal Name As String, _
                        ByVal Dob As String, _
                        ByVal Chart As String, _
                        ByVal Address0 As String, _
                        ByVal Address1 As String, _
                        ByVal Sex As String, _
                        ByVal Hospital As String)

      Dim S As String
      Dim n As Integer
      Dim BioPhone As String
      Dim HaemPhone As String

10    On Error GoTo PrintHeading_Error

20    Printer.Font.Name = "Courier New"
30    Printer.Font.Size = 16
40    Printer.Font.Bold = True
50    Printer.Font.Name = "Courier New"
60    Printer.Font.Size = 16
70    Printer.Font.Bold = True

80    Select Case Dept
      Case "Haematology":
90        Printer.ForeColor = vbRed
100   Case "Biochemistry":
110       Printer.ForeColor = vbGreen
120   Case "Blood Transfusion":
130       Printer.ForeColor = vbRed
140   Case "Coagulation":
150       Printer.ForeColor = vbRed           'RGB(80, 46, 107) 'Purple
160   Case "Microbiology"
170       Printer.ForeColor = vbBlue
          '180       Printer.Line (0, 0)-(11250, 350), vbYellow, BF
          '190       Printer.CurrentX = 0
          '200       Printer.CurrentY = 0
180   End Select
190   Dept = Dept & " Laboratory"
200   Printer.Print "CAVAN GENERAL HOSPITAL : " & Dept;
210   Printer.Font.Size = 10
220   Printer.CurrentY = 100
230   Select Case Dept
      Case "Haematology Laboratory":
240       HaemPhone = GetOptionSetting("HaemPhone", "")
250       If HaemPhone <> "" Then
260           Printer.Print " Phone " & HaemPhone;
270       End If
280   Case "Biochemistry Laboratory":
290       BioPhone = GetOptionSetting("BioPhone", "")
300       If BioPhone <> "" Then
310           Printer.Print " Phone " & BioPhone;
320       End If
330   Case "Blood Transfusion Laboratory":
340       Printer.Print " Phone 38830";
350   Case "Microbiology Laboratory":
360       Printer.Print ;
370   End Select
380   Printer.Print

390   Printer.CurrentY = 320

400   Printer.Font.Size = 4
410   If gPrintCopyReport = 0 Then
420       Printer.Print String$(250, "-")
430   Else
440       S = "-- THIS IS A COPY REPORT -- NOT FOR FILING --"
450       Printer.ForeColor = vbRed
460       For n = 1 To 5
470           Printer.Print S;
480       Next
490       Printer.Print
500   End If

510   Printer.ForeColor = vbBlack

520   Printer.Font.Name = "Courier New"
530   Printer.Font.Size = 12
540   Printer.Font.Bold = False

550   Printer.Print " Sample ID:";
      'Printer.Font.bold = True
560   Printer.Print RP.SampleID;
      'Printer.Font.bold = False

570   Printer.Print Tab(35); "Name:";
580   Printer.Font.Bold = True
590   Printer.Font.Size = 14
600   Printer.Print Left$(Name, 27)
610   Printer.Font.Size = 12
620   Printer.Font.Bold = False

630   Printer.Print "      Ward:";
      'Printer.Font.bold = True
640   Printer.Print RP.Ward;
      'Printer.Font.bold = False

650   Printer.Print Tab(35); " DOB:";
      'Printer.Font.bold = True
660   Printer.Print Format(Dob, "dd/mm/yyyy");
      'Printer.Font.bold = False
670   Printer.Print Tab(63); "Chart #:";
      'Printer.Font.bold = True
680   Printer.Print Left$(Hospital & " ", 1) & " ";
690   Printer.Print Chart
      'Printer.Font.bold = False

700   Printer.Print "Consultant:";
      'Printer.Font.bold = True
710   Printer.Print RP.Clinician;
      'Printer.Font.bold = False
720   Printer.Print Tab(35); "Addr:";
      'Printer.Font.bold = True
730   Printer.Print Address0;
      'Printer.Font.bold = False
740   Printer.Print Tab(63); "    Sex:";
750   Select Case Left$(UCase$(Trim$(Sex)), 1)
      Case "M": Printer.Print "Male"
760   Case "F": Printer.Print "Female"
770   Case Else: Printer.Print
780   End Select

790   Printer.Print "        GP:";
      'Printer.Font.bold = True
800   Printer.Print RP.GP;
      'Printer.Font.bold = False
810   Printer.Print Tab(35); "     ";
820   Printer.Print Address1

830   Printer.Font.Size = 4
840   Printer.Font.Bold = True
850   If gPrintCopyReport = 0 Then
860       Printer.Print String$(250, "-")
870   Else
880       S = "-- THIS IS A COPY REPORT -- NOT FOR FILING --"
890       Printer.ForeColor = vbRed
900       For n = 1 To 5
910           Printer.Print S;
920       Next
930       Printer.Print
940   End If
950   Printer.ForeColor = vbBlack
960   Printer.Font.Bold = False

970   Exit Sub

PrintHeading_Error:

      Dim strES As String
      Dim intEL As Integer

980   intEL = Erl
990   strES = Err.Description
1000  LogError "Module1", "PrintHeading", intEL, strES

End Sub

Public Function EnsureColumnExists(ByVal TableName As String, _
                                   ByVal ColumnName As String, _
                                   ByVal Definition As String) _
                                   As Boolean

      'Return 1 if column created
      '       0 if column already exists

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo EnsureColumnExists_Error

20    sql = "IF NOT EXISTS " & _
            "    (SELECT * FROM syscolumns WHERE " & _
            "    id = object_id('" & TableName & "') " & _
            "    AND name = '" & ColumnName & "') " & _
            "  BEGIN " & _
            "    ALTER TABLE " & TableName & " " & _
            "    ADD " & ColumnName & " " & Definition & " " & _
            "    SELECT 1 AS RetVal " & _
            "  END " & _
            "ELSE " & _
            "  SELECT 0 AS RetVal"

30    Set tb = Cnxn(0).Execute(sql)

40    EnsureColumnExists = tb!RetVal

50    Exit Function

60    Exit Function

EnsureColumnExists_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "Module1", "EnsureColumnExists", intEL, strES, sql

End Function

Public Function LoadLIH(ByVal SampleID As Long, _
                        ByRef Lipaemic As Integer, _
                        ByRef Icteric As Integer, _
                        ByRef Haemolysed As Integer) _
                        As Boolean

      Dim sql As String
      Dim tb As Recordset
      Dim RetVal As Boolean
      Dim LIHVal As Integer

10    On Error GoTo LoadLIH_Error

20    RetVal = False

30    sql = "Select * from Masks where " & _
            "SampleID = " & SampleID & " " & _
            "AND (LIH <> 0 OR O IS NOT NULL)"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then
70        RetVal = True
80        LIHVal = tb!LIH
90        If LIHVal > 99 Then
100           Lipaemic = LIHVal \ 100
110           If Lipaemic > 5 Then
120               Lipaemic = 0
130           End If
140           LIHVal = LIHVal Mod 100
150       End If
160       If LIHVal > 9 Then
170           Icteric = LIHVal \ 10
180           If Icteric > 5 Then
190               Icteric = 0
200           End If
210           LIHVal = LIHVal Mod 10
220       End If
230       Haemolysed = LIHVal
240       If Haemolysed > 5 Then
250           Haemolysed = 0
260       End If
270   End If

280   LoadLIH = RetVal

290   Exit Function

LoadLIH_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "Module1", "LoadLIH", intEL, strES, sql

End Function
Public Function LIHEffects(ByVal Code As String, _
                           ByVal Lipaemic As Integer, _
                           ByVal Icteric As Integer, _
                           ByVal Haemolysed As Integer) As Boolean

      Dim RetVal As Boolean
      Dim sql As String
      Dim tb As Recordset
      Dim l As Integer
      Dim i As Integer
      Dim h As Integer
      Dim LIHVal As Integer

10    On Error GoTo LIHEffects_Error

20    RetVal = False

30    If HospName(0) = "Monaghan" Then
40        If Lipaemic <> 0 Or Icteric <> 0 Or Haemolysed <> 0 Then
50            sql = "SELECT TOP 1 LIH FROM BioTestDefinitions WHERE " & _
                    "Code = '" & Code & "' " & _
                    "AND ( LIH <> 0 OR O = 1 )"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            If Not tb.EOF Then
90                LIHVal = tb!LIH
100               If LIHVal > 99 Then
110                   l = LIHVal \ 100
120                   If l > 5 Then l = 0
130                   If Lipaemic <> 0 And Lipaemic >= l Then
140                       RetVal = True
150                   End If
160                   LIHVal = LIHVal Mod 100
170               End If
180               If LIHVal > 9 Then
190                   i = LIHVal \ 10
200                   If i > 5 Then i = 0
210                   If Icteric <> 0 And Icteric >= i Then
220                       RetVal = True
230                   End If
240                   LIHVal = LIHVal Mod 10
250               End If
260               h = LIHVal
270               If h > 5 Then h = 0
280               If Haemolysed <> 0 And Haemolysed >= h Then
290                   RetVal = True
300               End If
310           End If
320       End If
330   End If

340   LIHEffects = RetVal

350   Exit Function

LIHEffects_Error:

      Dim strES As String
      Dim intEL As Integer

360   intEL = Erl
370   strES = Err.Description
380   LogError "Module1", "LIHEffects", intEL, strES, sql

End Function


Public Sub PrintResultBioSideBySide(ByVal Dept As String, ByVal VV As Integer)
      'Dept is either "Bio" or "Imm"
      'VALID = 1 NOTVALID = 2
      'DONTCARE = 0

      Dim tb As Recordset
      Dim tbF As Recordset
      Dim tbUN As Recordset
      Dim sql As String
      Dim Sex As String
      Dim lpc As Integer
      Dim cUnits As String
      Dim Flag As String
      Dim n As Integer
      Dim v As String
      Dim Low As Single
      Dim High As Single
      Dim strLow As String * 5
      Dim strHigh As String * 5
      Dim BRs As New BIEResults
      Dim BR As BIEResult
      Dim TestCount As Integer
      Dim SampleType As String
      Dim ResultsPresent As Boolean
      Dim OBs As Observations
10    On Error GoTo PrintResultBioSideBySide_Error

20    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim Dob As String
      Dim RunTime As String
      Dim Fasting As String
      Dim udtPrintLine(0 To 60) As PrintLine    'max 30 result lines per page
      Dim strFormat As String
      Dim MultiColumn As Boolean
      Dim BioComment As String
      Dim DemoComment As String
      Dim FullDept As String
      Dim CodeForChol As String
      Dim CodeForGlucose As String
      Dim CodeForTrig As String

30    CodeForChol = UCase$(GetOptionSetting("BioCodeForChol", ""))
40    CodeForGlucose = GetOptionSetting("BioCodeForGlucose", "")
50    CodeForTrig = GetOptionSetting("BioCodeForTrig", "")

60    If Dept = "Bio" Then
70        FullDept = "Biochemistry"
80    ElseIf Dept = "Imm" Then
90        FullDept = "Immunology"
100   End If

110   For n = 0 To 60
120       udtPrintLine(n).Analyte = ""
130       udtPrintLine(n).Result = ""
140       udtPrintLine(n).Flag = ""
150       udtPrintLine(n).Units = ""
160       udtPrintLine(n).NormalRange = ""
170       udtPrintLine(n).Fasting = ""
180   Next

190   sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
200   Set tb = New Recordset
210   RecOpenClient 0, tb, sql

220   If tb.EOF Then
230       Exit Sub
240   End If

250   If Not IsNull(tb!Fasting) Then
260       Fasting = tb!Fasting
270   Else
280       Fasting = False
290   End If

300   If IsDate(tb!Dob) Then
310       Dob = Format(tb!Dob, "dd/mmm/yyyy")
320   Else
330       Dob = ""
340   End If

350   ResultsPresent = False
360   Set BRs = BRs.Load(Dept, RP.SampleID, "Results", VV, gDONTCARE)
370   If Not BRs Is Nothing Then
380       TestCount = BRs.Count
390       If TestCount <> 0 Then
400           ResultsPresent = True
410           SampleType = BRs(1).SampleType
420           If Trim$(SampleType) = "" Then SampleType = "S"
430       End If
440   End If

450   If Not SetPrinter("CHBIO") Then Exit Sub

460   lpc = 0
470   If ResultsPresent Then
480       For Each BR In BRs
490           RunTime = BR.RunTime
500           v = BR.Result

510           If BR.Code = CodeForGlucose Or _
                 BR.Code = CodeForChol Or _
                 BR.Code = CodeForTrig Then
520               If Fasting Then
530                   If BR.Code = CodeForGlucose Then
540                       sql = "Select * from Fastings where " & _
                                "TestName = 'GLU'"
550                       Set tbF = New Recordset
560                       RecOpenServer 0, tbF, sql
570                   ElseIf BR.Code = CodeForChol Then
580                       sql = "Select * from Fastings where " & _
                                "TestName = 'CHO'"
590                       Set tbF = New Recordset
600                       RecOpenServer 0, tbF, sql
610                   ElseIf BR.Code = CodeForTrig Then
620                       sql = "Select * from Fastings where " & _
                                "TestName = 'TRI'"
630                       Set tbF = New Recordset
640                       RecOpenServer 0, tbF, sql
650                   End If
660                   If Not tbF.EOF Then
670                       High = tbF!FastingHigh
680                       Low = tbF!FastingLow
690                   Else
700                       High = Val(BR.High)
710                       Low = Val(BR.Low)
720                   End If
730               Else
740                   High = Val(BR.High)
750                   Low = Val(BR.Low)
760               End If
770           Else
780               High = Val(BR.High)
790               Low = Val(BR.Low)
800           End If

810           If Low < 10 Then
820               strLow = Format(Low, "0.00")
830           ElseIf Low < 100 Then
840               strLow = Format(Low, "##.0")
850           Else
860               strLow = Format(Low, " ###")
870           End If
880           If High < 10 Then
890               strHigh = Format(High, "0.00")
900           ElseIf High < 100 Then
910               strHigh = Format(High, "##.0")
920           Else
930               strHigh = Format(High, "### ")
940           End If

950           If IsNumeric(v) Then
960               If Val(v) > BR.PlausibleHigh Then
970                   udtPrintLine(lpc).Flag = " X "
980                   Flag = " X"
990               ElseIf Val(v) < BR.PlausibleLow Then
1000                  udtPrintLine(lpc).Flag = " X "
1010                  Flag = " X"
1020              ElseIf Val(v) > BR.FlagHigh Then
1030                  udtPrintLine(lpc).Flag = " H "
1040                  Flag = " H"
1050              ElseIf Val(v) < BR.FlagLow Then
1060                  udtPrintLine(lpc).Flag = " L "
1070                  Flag = " L"
1080              Else
1090                  udtPrintLine(lpc).Flag = "   "
1100                  Flag = "  "
1110              End If
1120          Else
1130              udtPrintLine(lpc).Flag = "   "
1140              Flag = "  "
1150          End If
1160          udtPrintLine(lpc).Analyte = Left$(BR.LongName & Space(16), 16)

1170          If IsNumeric(v) Then
1180              Select Case BR.Printformat
                  Case 0: strFormat = "######"
1190              Case 1: strFormat = "###0.0"
1200              Case 2: strFormat = "##0.00"
1210              Case 3: strFormat = "#0.000"
1220              End Select
1230              udtPrintLine(lpc).Result = Format(v, strFormat)
1240          Else
1250              udtPrintLine(lpc).Result = v
1260          End If

1270          If udtPrintLine(lpc).Flag = " X " Then
1280              udtPrintLine(lpc).Result = "XXXX"
1290          End If

1300          sql = "Select * from Lists where " & _
                    "ListType = 'UN' and Code = '" & BR.Units & "'"
1310          Set tbUN = Cnxn(0).Execute(sql)
1320          If Not tbUN.EOF Then
1330              cUnits = Left$(tbUN!Text & Space(6), 6)
1340          Else
1350              cUnits = Left$(BR.Units & Space(6), 6)
1360          End If
1370          udtPrintLine(lpc).Units = cUnits
1380          udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"

1390          udtPrintLine(lpc).Fasting = ""
1400          If tb!Fasting Then
1410              udtPrintLine(lpc).Fasting = "(Fasting)"
1420          End If

1430          LogBioAsPrinted RP.SampleID, BR.Code

1440          lpc = lpc + 1
1450      Next
1460  End If

1470  PrintHeading FullDept, tb!PatName & "", Dob, tb!Chart & "", _
                   tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

1480  Sex = tb!Sex & ""

1490  Printer.Print

1500  If TestCount <= Val(frmMain.txtMoreThan) Then
1510      MultiColumn = False
          '1660    Printer.CurrentY = 2500 + (20 - TestCount) * 100
1520  Else
1530      MultiColumn = True
          '1690    Printer.CurrentY = 2500
1540  End If

1550  Printer.Font.Size = 10

1560  If MultiColumn Then
1570      For n = 0 To Val(frmMain.txtMoreThan) - 1
1580          Printer.Font.Bold = False
1590          Printer.Print udtPrintLine(n).Analyte;
1600          If udtPrintLine(n).Flag <> "   " Then
1610              Printer.Font.Bold = True
1620          End If
1630          Printer.Print udtPrintLine(n).Result;
1640          Printer.Print udtPrintLine(n).Flag;
1650          Printer.Font.Bold = False
1660          Printer.Font.Size = 8
1670          Printer.Print udtPrintLine(n).Units;
1680          Printer.Print udtPrintLine(n).NormalRange;
1690          Printer.Font.Size = 10
              'Now Right Hand Column
1700          Printer.Print Tab(45);
1710          Printer.Print udtPrintLine(n + Val(frmMain.txtMoreThan)).Analyte;
1720          If udtPrintLine(n + Val(frmMain.txtMoreThan)).Flag <> "   " Then
1730              Printer.Font.Bold = True
1740          End If
1750          Printer.Print udtPrintLine(n + Val(frmMain.txtMoreThan)).Result;
1760          Printer.Print udtPrintLine(n + Val(frmMain.txtMoreThan)).Flag;
1770          Printer.Font.Bold = False
1780          Printer.Font.Size = 8
1790          Printer.Print udtPrintLine(n + Val(frmMain.txtMoreThan)).Units;
1800          Printer.Print udtPrintLine(n + Val(frmMain.txtMoreThan)).NormalRange
1810          Printer.Font.Size = 10
1820      Next
1830      If Fasting Then
1840          Printer.Print "(All above relate to Normal Fasting Ranges.)"
1850      End If
1860  Else
1870      For n = 0 To 35
1880          If Trim$(udtPrintLine(n).Analyte) <> "" Then
1890              Printer.Print Tab(20);
1900              Printer.Font.Bold = False
1910              Printer.Print udtPrintLine(n).Analyte;
1920              If udtPrintLine(n).Flag <> "   " Then
1930                  Printer.Font.Bold = True
1940              End If
1950              Printer.Print udtPrintLine(n).Result;
1960              Printer.Print udtPrintLine(n).Flag;
1970              Printer.Font.Bold = False
1980              Printer.Print udtPrintLine(n).Units;
1990              Printer.Print udtPrintLine(n).NormalRange;
2000              Printer.Print udtPrintLine(n).Fasting
2010          End If
2020      Next
2030  End If

2040  Set OBs = New Observations
2050  Set OBs = OBs.Load(RP.SampleID, FullDept)
2060  If Not OBs Is Nothing Then
2070      FillCommentLines OBs(1).Comment, 4, Comments(), 97
2080      For n = 1 To 4
2090          Printer.Print Comments(n)
2100      Next
2110  End If

2120  Set OBs = New Observations
2130  Set OBs = OBs.Load(RP.SampleID, "Demographic")
2140  If Not OBs Is Nothing Then
2150      FillCommentLines OBs(1).Comment, 4, Comments(), 97
2160      For n = 1 To 4
2170          Printer.Print Comments(n)
2180      Next
2190  End If

2200  PrintNoSexDoB Sex, Dob

2210  Printer.ForeColor = vbBlack

2220  If IsDate(tb!SampleDate) Then
2230      SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
2240  Else
2250      SampleDate = ""
2260  End If
2270  If IsDate(RunTime) Then
2280      Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
2290  Else
2300      If IsDate(tb!Rundate) Then
2310          Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
2320      Else
2330          Rundate = ""
2340      End If
2350  End If

2360  PrintFooter FullDept, RP.Initiator, SampleDate, Rundate

2370  Printer.EndDoc



2380  ReSetPrinter

2390  sql = "UPDATE " & Dept & "Results SET Printed = '1' WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "AND Code NOT IN " & _
            "( SELECT DISTINCT(Code) FROM " & Dept & "TestDefinitions D JOIN PrintInhibit P " & _
            "  ON D.ShortName = P.Parameter " & _
            "  WHERE SampleID = '" & RP.SampleID & "' " & _
            "  AND Discipline = '" & Dept & "')"
2400  Cnxn(0).Execute sql

2410  Exit Sub

PrintResultBioSideBySide_Error:

      Dim strES As String
      Dim intEL As Integer

2420  intEL = Erl
2430  strES = Err.Description
2440  LogError "Module1", "PrintResultBioSideBySide", intEL, strES, sql, "SampleID = " & RP.SampleID

End Sub
Sub LogBioAsPrinted(ByVal SampleID As String, _
                    ByVal TestCode As String)

      Dim sql As String

10    On Error GoTo LogBioAsPrinted_Error

20    sql = "update BioResults " & _
            "set valid = 1, printed = 1 where " & _
            "SampleID = '" & SampleID & "' " & _
            "and code = '" & TestCode & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

LogBioAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "Module1", "LogBioAsPrinted", intEL, strES, sql

End Sub


Sub LogCoagAsPrinted(ByVal SampleID As String)

      Dim sql As String

10    On Error GoTo LogCoagAsPrinted_Error

20    sql = "update CoagResults " & _
            "set valid = 1, printed = 1 where " & _
            "SampleID = '" & SampleID & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

LogCoagAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "Module1", "LogCoagAsPrinted", intEL, strES, sql

End Sub



Public Function CheckAutoComments(ByVal SampleID As String, ByVal ShortName As String, ByVal index As Integer) As String

      Dim tb As Recordset
      Dim sql As String
      Dim ShortDisc As String
      Dim Discipline As String
      Dim RetVal As String

10    On Error GoTo CheckAutoComments_Error

20    RetVal = ""

30    If index = 2 Then
40        Discipline = "Biochemistry"
50        ShortDisc = "Bio"
60    Else
70        Discipline = "Coagulation"
80        ShortDisc = "Coag"
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
            "AND R.SampleID = '" & SampleID & "' " & _
            "AND A.Parameter = '" & ShortName & "' " & _
            "ORDER BY A.ListOrder"

120   Set tb = New Recordset
130   RecOpenClient 0, tb, sql
140   Do While Not tb.EOF
150       If Trim$(tb!Output & "") <> "" Then
160           RetVal = RetVal & tb!Output & vbCrLf
170       End If
180       tb.MoveNext
190   Loop

200   CheckAutoComments = Trim$(RetVal)

210   Exit Function

CheckAutoComments_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "Module1", "CheckAutoComments", intEL, strES, sql

End Function

Private Sub SendZetaFax(ByVal Recipient As String, _
                        ByVal FaxNumber As String, _
                        ByVal ZetaFaxFolderFilePathName As String, _
                        ByVal ZetaFaxDocFilePathName As String)

      Dim Message As String
      Dim f As Integer

10    On Error GoTo SendZetaFax_Error

20    Message = "%%[MESSAGE]" & vbCrLf & _
                "FROM: Laboratory General Hospital CAVAN" & vbCrLf & _
                "TO: " & Recipient & vbCrLf & _
                "FAX: " & FaxNumber & vbCrLf & _
                "%%[FILE]" & vbCrLf & _
                ZetaFaxDocFilePathName & vbCrLf

30    f = FreeFile()
40    Open ZetaFaxFolderFilePathName For Output As f
50    Print #f, Message
60    Close f

70    Exit Sub

SendZetaFax_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "Module1", "SendZetaFax", intEL, strES

End Sub

Public Function TechnicianCodeFor(ByVal CodeOrName As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo TechnicianCodeFor_Error

20    CodeOrName = Trim$(AddTicks(CodeOrName))

30    sql = "Select Code from Users where " & _
            "Name = '" & CodeOrName & "' " & _
            "or Code = '" & CodeOrName & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        TechnicianCodeFor = tb!Code & ""
80    End If

90    Exit Function

TechnicianCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Module1", "TechnicianCodeFor", intEL, strES, sql

End Function

Public Sub PrintText(ByVal Text As String, _
                     Optional FontSize As Integer = 9, _
                     Optional FontBold As Boolean = False, _
                     Optional FontItalic As Boolean = False, _
                     Optional FontUnderLine As Boolean = False, _
                     Optional FontColor As ColorConstants = vbBlack)


      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------

10    Printer.FontSize = FontSize
20    Printer.FontBold = FontBold
30    Printer.FontItalic = FontItalic
40    Printer.FontUnderLine = FontUnderLine
50    Printer.ForeColor = FontColor
60    Printer.Print Text;

End Sub
Public Function FormatString(strDestString As String, _
                             intNumChars As Integer, _
                             Optional strSeperator As String = "", _
                             Optional intAlign As PrintAlignContants = AlignLeft) _
                             As String

      '**************intAlign = 0 --> Left Align
      '**************intAlign = 1 --> Center Align
      '**************intAlign = 2 --> Right Align
      Dim intPadding As Integer

10    On Error GoTo FormatString_Error

20    intPadding = 0

30    If Len(strDestString) > intNumChars Then
40        FormatString = Mid$(strDestString, 1, intNumChars) & strSeperator
50    ElseIf Len(strDestString) < intNumChars Then
          Dim i As Integer
          Dim intStringLength As String
60        intStringLength = Len(strDestString)
70        intPadding = intNumChars - intStringLength

80        If intAlign = PrintAlignContants.AlignLeft Then
90            strDestString = strDestString & String$(intPadding, " ")  '& " "
100       ElseIf intAlign = PrintAlignContants.AlignCenter Then
110           If (intPadding Mod 2) = 0 Then
120               strDestString = String$(intPadding / 2, " ") & strDestString & String$(intPadding / 2, " ")
130           Else
140               strDestString = String$((intPadding - 1) / 2, " ") & strDestString & String$((intPadding - 1) / 2 + 1, " ")
150           End If
160       ElseIf intAlign = PrintAlignContants.AlignRight Then
170           strDestString = String$(intPadding, " ") & strDestString
180       End If

190       strDestString = strDestString & strSeperator
200       FormatString = strDestString
210   Else
220       strDestString = strDestString & strSeperator
230       FormatString = strDestString
240   End If

250   Exit Function

FormatString_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "Module1", "FormatString", intEL, strES

End Function

Public Function CheckDisablePrinting(ByVal GPName As String, Department As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CheckDisablePrinting_Error

20    CheckDisablePrinting = False
      'If RP.WardPrint = True Then Exit Function

30    sql = "SELECT * from DisablePrinting WHERE " & _
            "Department = '" & Department & "' " & _
            "AND GPName = '" & AddTicks(GPName) & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then
70        CheckDisablePrinting = True
80    End If

90    Exit Function

CheckDisablePrinting_Error:
      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Other", "CheckDisablePrinting", intEL, strES, sql

End Function

