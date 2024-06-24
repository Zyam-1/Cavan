Attribute VB_Name = "modRTFMicro"
Option Explicit

Public MicroCommentLineCount As Integer
Public MicroTotalLineCount As Integer
Public SedimexResultsExist As Boolean


Private Sub MicroPrintAndStore(ByVal PageNumber As Integer)

      Dim sql As String
      Dim Gx As New GP
      Dim PrintHardCopy As Boolean

10    On Error GoTo MicroPrintAndStore_Error

15  If InStr(frmMain.rtb.TextRTF, RP.SampleID) Then 'Double check that report saved is for this Sample Id

20    PrintHardCopy = True


30    Gx.LoadName RP.GP

      'WRITE ALL RULES HERE WHEN YOU DONT WANT TO PRINT HARD COPY

40    If (Gx.PrintReport = False And RP.Ward = "GP") Then PrintHardCopy = False
50    If UCase(RP.PrintAction) = UCase("Save") Then PrintHardCopy = False
60    If UCase(RP.Initiator) = "BIOMNIS IRELAND" Then PrintHardCopy = False

      'END WRITING RULES FOR PRINTING HARD COPY
70    If PageNumber = 1 Then
80        sql = "UPDATE Reports Set Hidden = 1 WHERE SampleID = '" & RP.SampleID & "' AND Dept = 'Microbiology' "
90        Cnxn(0).Execute sql
100   End If

110   With frmMain.rtb
120       .SelStart = 0

130       sql = "INSERT INTO Reports " & _
                "(SampleID, Dept, Initiator, PrintTime, ReportNumber, PageNumber, Report, Printer, ReportType) " & _
                "VALUES " & _
                "( '" & RP.SampleID & "', " & _
                "  'Microbiology', " & _
                "  '" & RP.Initiator & "', " & _
                "   getdate(), " & _
                "  '" & RP.SampleID & "M" & "', " & _
                "  '" & PageNumber & "', " & _
                "  '" & AddTicks(.TextRTF) & "', " & _
                "  '" & IIf(PrintHardCopy, Printer.DeviceName, "None") & "', " & _
                "  '" & IIf(UCase(RP.PrintAction) = UCase("Save"), "Interim Report", "Final Report") & "')"

140       Cnxn(0).Execute sql

150       If PrintHardCopy Then

160           .SelPrint Printer.hDC
170       End If

180   End With
190 Else
200        LogError "modRTF", "PrintAndStore", 200, RP.SampleID & " not found in RTF report", sql
210 End If
      
220   Exit Sub

MicroPrintAndStore_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "modRTFMicro", "MicroPrintAndStore", intEL, strES, sql

End Sub

Public Sub RTFPrintText(ByVal Text As String, _
                        Optional FontSize As Integer = 9, _
                        Optional FontBold As Boolean = False, _
                        Optional FontItalic As Boolean = False, _
                        Optional FontUnderLine As Boolean = False, _
                        Optional FontColor As ColorConstants = vbBlack)

10    On Error GoTo RTFPrintText_Error

20    With frmMain.rtb
30        .SelFontSize = FontSize
40        .SelBold = FontBold
50        .SelItalic = FontItalic
60        .SelUnderline = FontUnderLine
70        .SelColor = FontColor
80        .SelText = Text
90    End With

100   Exit Sub

RTFPrintText_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modRTFMicro", "RTFPrintText", intEL, strES

End Sub


Public Sub RTFPrintResultMicro()

      Dim PageCount As Integer
      Dim PageNumber As Integer
      Dim pDefault As Integer
      Dim ABCount As Integer
      Dim CommentLineCount As Integer
      Dim MicroscopyLineCount As Integer
      Dim TotalLines As Integer
      Dim IsolateCount As Integer
      Dim NegativeResults As Boolean
      Dim CommentsPresent As Boolean
      Dim MiscLineCount As Integer
      Dim CSFResultsCount As Integer
      Dim PregnancyLineCount As Integer


10    On Error GoTo RTFPrintResultMicro_Error

20    MicroTotalLineCount = Val(GetOptionSetting("PerPageTotalLineMicro", "72"))
30    MicroCommentLineCount = Val(GetOptionSetting("MicroCommentLineCount", "4"))

40    MiscLineCount = GetMiscLineCount(RP.SampleID)    'FOB+CDiff+Rota/Adeno+OP+RSV
50    CommentLineCount = GetCommentLineCount(RP.SampleID)
60    CommentsPresent = CommentLineCount > 0

70    If Not SetPrinter("MICRO") Then Exit Sub

80    If isBDMaxReport(RP.SampleID) Then
90        RTFPrintBDMaxReport "HICSDOPF", RP, 1, 1, "1234", CommentsPresent
100       MicroPrintAndStore 1
110   Else
120       MicroscopyLineCount = GetMicroscopyLineCount(RP.SampleID)
130       IsolateCount = GetIsolateCount(RP.SampleID)
140       NegativeResults = IsNegativeResults(RP.SampleID)
150       pDefault = GetPDefault(RP.SampleID)
160       CSFResultsCount = GetCSFCount(RP.SampleID)
170       PregnancyLineCount = GetPregnancyLineCount(RP.SampleID)

180       Select Case IsolateCount
              'A Current Antibiotics
              'B Occult Blood
              'C Clin Details
              'D Demographic Comment
              'E RSV
              'F Footer
              'G Pregnancy
              'H Heading
              'K Reducing Substances
              'M Microscopy
              'N Negative Results
              'O Consultant Comments
              'P Specimen Comments
              'R Rota/Adeno
              'S Sensitivities
              'T CDiff
              'V Ova/Parasites
              'W Blood Culture Associated SampleID
              'X MRSA / VRE Associated SampleIDs
              'Y Urine Comment
              'Z GramStain + WetPrep
          Case 0:
190           PageCount = 1
200           PageNumber = 1
210           TotalLines = CSFResultsCount + CommentLineCount + MicroscopyLineCount + MiscLineCount + PregnancyLineCount
220           If TotalLines > 0 Then
230               RTFPrintMicroPage "HICDLAMZBGRTVEKPYOWXF", RP, PageNumber, PageCount, "", CommentsPresent
240               MicroPrintAndStore 1
250           End If

260       Case 1, 2, 3, 4:
270           PageCount = 1
280           PageNumber = 1
290           If NegativeResults Then
300               RTFPrintMicroPage "HICDLAMZBGRTVEJKNPYOWXF", RP, PageNumber, PageCount, "", CommentsPresent
310               MicroPrintAndStore 1
320           Else
330               ABCount = GetABCount(RP.SampleID, "1234")
340               TotalLines = CSFResultsCount + CommentLineCount + MicroscopyLineCount + IsolateCount + ABCount + MiscLineCount

350               If TotalLines > MicroTotalLineCount Then
360                   PageCount = 2
370                   PageNumber = 1
380                   RTFPrintMicroPage "HIDLJSBGRTVEJKPYOWXF", RP, PageNumber, PageCount, "1234", CommentsPresent

390                   MicroPrintAndStore 1

400                   PageNumber = 2
410                   RTFPrintMicroPage "HICDLAMZPYOWXF", RP, PageNumber, PageCount, "", CommentsPresent
420                   MicroPrintAndStore 2
430               Else
440                   PageCount = 1
450                   RTFPrintMicroPage "HICDLAMZBTRVEKJSPYOWXF", RP, PageNumber, PageCount, "1234", CommentsPresent
460                   MicroPrintAndStore 1
470               End If
480           End If

490       Case 5, 6:
500           PageCount = 2
510           ABCount = GetABCount(RP.SampleID, "123456")
520           TotalLines = CSFResultsCount + CommentLineCount + MicroscopyLineCount + IsolateCount + ABCount + MiscLineCount
530           If TotalLines > MicroTotalLineCount Then
540               PageCount = 3
550               PageNumber = 1
560               RTFPrintMicroPage "HIDLSBGRTVEKPYOWXF", RP, PageNumber, PageCount, "123", CommentsPresent
570               MicroPrintAndStore 1

580               PageNumber = 2
590               RTFPrintMicroPage "HIJSF", RP, PageNumber, PageCount, "456", CommentsPresent
600               MicroPrintAndStore 2

610               PageNumber = 3
620               RTFPrintMicroPage "HICAMZF", RP, PageNumber, PageCount, "", CommentsPresent
630               MicroPrintAndStore 3
640           Else
650               PageNumber = 1
660               RTFPrintMicroPage "HICDLAMZSPYOWXF", RP, PageNumber, PageCount, "123", CommentsPresent
670               MicroPrintAndStore 1

680               PageNumber = 2
690               RTFPrintMicroPage "HIJSF", RP, PageNumber, PageCount, "456", CommentsPresent
700               MicroPrintAndStore 2
710           End If

720       End Select

730   End If

740   ReSetPrinter

750   Exit Sub

RTFPrintResultMicro_Error:

      Dim strES As String
      Dim intEL As Integer

760   intEL = Erl
770   strES = Err.Description
780   LogError "modRTFMicro", "RTFPrintResultMicro", intEL, strES

End Sub

Private Sub RTFPrintBDMaxReport(ByVal HCLM As String, ByRef RP As ReportToPrint, _
                                ByVal PageNumber As String, _
                                ByVal PageCount As String, _
                                ByVal IsolateNumbers As String, _
                                ByVal CommentsPresent As Boolean)

      Dim n As Integer
      Dim sql As String
      Dim tb As Recordset
      Dim PatName As String
      Dim Dob As String
      Dim Chart As String
      Dim Address0 As String
      Dim Address1 As String
      Dim Sex As String
      Dim Hospital As String

10    On Error GoTo RTFPrintBDMaxReport_Error

20    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID + sysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        PatName = tb!PatName & ""
70        If IsDate(tb!Dob) Then
80            Dob = Format(tb!Dob, "dd/mmm/yyyy")
90        End If
100       Chart = tb!Chart & ""
110       Address0 = tb!Addr0 & ""
120       Address1 = tb!Addr1 & ""
130       Sex = tb!Sex & ""
140       Hospital = tb!Hospital & ""
150   End If

160   For n = 1 To Len(HCLM)
170       Select Case Mid$(HCLM, n, 1)
          Case "A": RTFPrintMicroCurrentABs RP.SampleID
180       Case "C": RTFPrintMicroClinDetails RP.SampleID
190       Case "D": RTFPrintMicroComment RP.SampleID, "Demographics"
200       Case "F": RTFPrintMicroFooter RP
210       Case "H": RTFPrintHeading "Microbiology", PatName, Dob, Chart, Address0, Address1, Sex, Hospital
220       Case "I": RTFPrintSpecType RP, PageNumber, PageCount
230       Case "O": RTFPrintMicroComment RP.SampleID, "Consultant"
240       Case "P": RTFPrintMicroComment RP.SampleID, "Med Scientist"
250       Case "S": RTFPrintMicroBDMax RP.SampleID, IsolateNumbers
260       End Select
270   Next

280   Exit Sub

RTFPrintBDMaxReport_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "modRTFMicro", "RTFPrintBDMaxReport", intEL, strES, sql

End Sub

Private Sub RTFPrintMicroPage(ByVal HCLM As String, _
                              ByRef RP As ReportToPrint, _
                              ByVal PageNumber As String, _
                              ByVal PageCount As String, _
                              ByVal IsolateNumbers As String, _
                              ByVal CommentsPresent As Boolean)

      'A Current Antibiotics
      'B Occult Blood
      'C Clin Details
      'D Demographic Comment
      'E RSV
      'F Footer
      'G Pregnancy
      'H Heading
      'I Micro Specific Heading
      'K Reducing Substances
      'M Microscopy
      'N Negative Results
      'O Consultant Comments
      'P Specimen Comments
      'R Rota/Adeno
      'S Sensitivities
      'T CDiff
      'V Ova/Parasites
      'W Blood Culture Associated SampleID
      'X MRSA / VRE Associated SampleIDs
      'Y Urine Comment
      'Z GramStain + WetPrep

      Dim n As Integer
      Dim sql As String
      Dim tb As Recordset
      Dim PatName As String
      Dim Dob As String
      Dim Chart As String
      Dim Address0 As String
      Dim Address1 As String
      Dim Sex As String
      Dim Hospital As String

10    On Error GoTo RTFPrintMicroPage_Error

20    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID + sysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        PatName = tb!PatName & ""
70        If IsDate(tb!Dob) Then
80            Dob = Format(tb!Dob, "dd/mmm/yyyy")
90        End If
100       Chart = tb!Chart & ""
110       Address0 = tb!Addr0 & ""
120       Address1 = tb!Addr1 & ""
130       Sex = tb!Sex & ""
140       Hospital = tb!Hospital & ""
150   End If

160   For n = 1 To Len(HCLM)
170       Debug.Print Mid$(HCLM, n, 1)
180       Select Case Mid$(HCLM, n, 1)

          Case "A": RTFPrintMicroCurrentABs RP.SampleID
190       Case "B": RTFPrintMicroOccultBlood RP.SampleID
200       Case "C": RTFPrintMicroClinDetails RP.SampleID
210       Case "D": RTFPrintMicroComment RP.SampleID, "Demographics"
220       Case "L": RTFPrintMicroComment RP.SampleID, "L"
230       Case "E": RTFPrintMicroRSV RP.SampleID
240       Case "F": RTFPrintMicroFooter RP
250       Case "G": RTFPrintMicroPregnancy RP.SampleID
260       Case "H": RTFPrintHeading "Microbiology", PatName, Dob, Chart, Address0, Address1, Sex, Hospital, CInt(PageCount), CInt(PageNumber)
270       Case "I": RTFPrintSpecType RP, PageNumber, PageCount
280       Case "J": RTFPrintMicroCSF RP.SampleID
290       Case "K": RTFPrintMicroRedSub RP.SampleID
300       Case "M": RTFPrintMicroscopy RP.SampleID
310       Case "N": RTFPrintNegativeResults RP.SampleID
320       Case "O": RTFPrintMicroComment RP.SampleID, "Consultant"
330       Case "P": RTFPrintMicroComment RP.SampleID, "Med Scientist"
340       Case "R": RTFPrintMicroRotaAdeno RP.SampleID
350       Case "S": RTFPrintMicroSensitivities RP.SampleID, IsolateNumbers, CommentsPresent
360       Case "T": RTFPrintMicroCDiff RP.SampleID
370       Case "V": RTFPrintMicroOvaParasites RP.SampleID
380       Case "W": RTFPrintMicroAssIDBC RP.SampleID
390       Case "X": RTFPrintMicroAssIDMRSA RP.SampleID
400       Case "Y": 'RTFPrintMicroUrineComment RP.SampleID
410       Case "Z": RTFPrintMicroGramWetPrep RP.SampleID

420       End Select
430   Next

440   Exit Sub

RTFPrintMicroPage_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "modRTFMicro", "RTFPrintMicroPage", intEL, strES

End Sub

Public Sub RTFPrintMicroFooter(ByRef RP As ReportToPrint)

      Dim tb As Recordset
      Dim sql As String
      Dim RecDate As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim Authoriser As String
      Dim Operator As String
      Dim ValidatedDate As String
      Dim S As String
      Dim x() As String
      Dim LL As Integer
      Dim AccreditationText As String

10    On Error GoTo RTFPrintMicroFooter_Error

20    sql = "Select AuthoriserCode from Sensitivities where " & _
            "SampleID = '" & Val(RP.SampleID) + sysOptMicroOffset(0) & "' " & _
            "and (AuthoriserCode <> '' or AuthoriserCode is not null)"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        sql = "Select Name from Users where Code = '" & tb!AuthoriserCode & "'"
70        Set tb = New Recordset
80        RecOpenClient 0, tb, sql
90        If Not tb.EOF Then
100           Authoriser = tb!Name & ""
110       End If
120   End If

130   sql = "Select * from Demographics where " & _
            "SampleID = '" & Val(RP.SampleID) + sysOptMicroOffset(0) & "'"
140   Set tb = New Recordset
150   RecOpenClient 0, tb, sql
160   If Not tb.EOF Then
170       Operator = tb!Operator & ""
180       If Not IsNull(tb!RecDate) Then
190           RecDate = Format(tb!RecDate, "dd/mm/yy hh:mm")
200       Else
210           RecDate = ""
220       End If
230       If Not IsNull(tb!SampleDate) Then
240           SampleDate = Format(tb!SampleDate, "dd/mm/yy hh:mm")
250           If Right$(SampleDate, 5) = "00:00" Then
260               SampleDate = Format$(SampleDate, "dd/mm/yy")
270           End If
280       Else
290           SampleDate = ""
300       End If
310       If Not IsNull(tb!Rundate) Then
320           Rundate = Format(tb!Rundate, "dd/mm/yy")
330       Else
340           Rundate = ""
350       End If
360   End If

370   sql = "SELECT * FROM PrintValidLog WHERE " & _
            "SampleID = '" & Val(RP.SampleID) + sysOptMicroOffset(0) & "'"
380   Set tb = New Recordset
390   RecOpenClient 0, tb, sql
400   If Not tb.EOF Then
410       If IsDate(tb!ValidatedDateTime) Then
420           ValidatedDate = Format$(tb!ValidatedDateTime, "dd/MM/yy hh:mm")
430       Else
440           ValidatedDate = Format$(Now, "dd/MM/yy hh:mm")
450       End If
460       Operator = tb!ValidatedBy & ""
470   Else
480       ValidatedDate = Format$(Now, "dd/MM/yy hh:mm")
490   End If

500   With frmMain.rtb
510       .SelFontName = "Courier New"
520       .SetFocus
530       .SelFontSize = 10
540       .SelColor = vbBlack

550       If Trim$(RP.SendCopyTo) <> "" Then
560           .SelText = RP.Clinician & " Requested copy to be sent to " & RP.SendCopyTo
570       End If
580       .SelText = vbCrLf
590       .SelFontSize = 10

600       .SetFocus
610       x = Split(.Text, vbCr)
620       LL = UBound(x)


630       Do While LL < MicroTotalLineCount
640           .SelFontSize = 10
650           .SelText = vbCrLf
660           x = Split(.Text, vbCr)
670           LL = UBound(x)
680       Loop

690       .SelFontName = "Courier New"
700       .SelAlignment = rtfCenter
710       .SelBold = False
720       If gPrintCopyReport = 0 Then
730           .SelFontSize = 4
740           .SelText = String$(200, "-") & vbCrLf
750       Else
760           .SelFontSize = 8
770           S = "- THIS IS A COPY REPORT - NOT FOR FILING -"
780           S = S & S
790           S = S & "- THIS IS A COPY REPORT -"
800           .SelColor = vbRed
810           .SelText = S & vbCrLf
820       End If

830       .SelFontName = "Courier New"
840       .SelFontSize = 10
850       .SelBold = False
860       .SelAlignment = rtfLeft
870       .SelColor = vbBlack
880       .SelText = "Micro "
890       .SelColor = vbBlue
900       .SelFontName = "MS Sans Serif"
910       .SelBold = False
920       .SelFontSize = 10

930       .SelText = "Specimen Date:" & SampleDate
940       .SelText = " Received:" & RecDate
950       .SelText = " Reported:" & Format$(ValidatedDate, "dd/mm/yy hh:mm")

960       If gPrintCopyReport = 0 Then
970           .SelText = " Reported by " & Operator
980       Else
990           sql = "SELECT TOP 1 Viewer FROM ViewedReports WHERE " & _
                    "SampleID = '" & RP.SampleID & "' " & _
                    "AND Discipline = 'N' " & _
                    "ORDER BY DateTime DESC"
1000          Set tb = New Recordset
1010          RecOpenServer 0, tb, sql
1020          If Not tb.EOF Then
1030              .SelText = " Printed by " & tb!Viewer & ""
1040          Else
1050              .SelText = " Printed by " & Left$(RP.Initiator, 14)
1060          End If
1070      End If
1080      AccreditationText = GetOptionSetting("MicroAccreditation", "")
          
      '    If Not isBDMaxReport(RP.SampleID) Then
      '        AccreditationText = GetOptionSetting("MicroAccreditation", "")
      '    End If

1090      If AccreditationText <> "" Then
1100          .SelText = vbNewLine
1110          .SelText = Space$(10) & String$(75, "-") & vbCrLf

1120          .SelColor = vbRed
1130          .SelAlignment = rtfLeft

1140          .SelFontName = "Courier New"
1150          .SelFontSize = 10
1160          .SelBold = True
1170          .SelText = Space$(5) & AccreditationText
1180          .SelText = vbCrLf
1190      End If

1200  End With

1210  Exit Sub

RTFPrintMicroFooter_Error:

      Dim strES As String
      Dim intEL As Integer

1220  intEL = Erl
1230  strES = Err.Description
1240  LogError "modRTFMicro", "RTFPrintMicroFooter", intEL, strES, sql

End Sub

Private Function isBDMaxReport(ByVal strSID As String) As Boolean
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo isBDMaxReport_Error

20    sql = "SELECT SampleID From Isolates WHERE SampleID = '" & Val(strSID) + sysOptMicroOffset(0) & "' " & _
            "and OrganismGroup like 'BD Max%'"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        isBDMaxReport = False
70    Else
80        isBDMaxReport = True
90    End If

100   Exit Function

isBDMaxReport_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modRTFMicro", "isBDMaxReport", intEL, strES, sql

End Function

Private Sub RTFPrintMicroGramWetPrep(ByVal SampleID As String)

      Dim ID As IdentResult
      Dim IDs As New IdentResults
      Dim Title As String

10    On Error GoTo RTFPrintMicroGramWetPrep_Error

20    With frmMain.rtb
30        IDs.Load (SampleID)
40        If IDs.Count > 0 Then
50            Title = "Gram Stain: "
60            For Each ID In IDs
70                If UCase$(ID.TestType) = "GRAMSTAIN" Then
80                    .SelBold = False
90                    .SelFontSize = 10
100                   .SelText = Space$(10) & Title
110                   .SelBold = True
120                   .SelText = ID.TestName & " " & ID.Result
130                   .SelText = vbCrLf
140                   Title = "            "
150               End If
160           Next

170           Title = "  Wet Prep: "
180           For Each ID In IDs
190               If UCase$(ID.TestType) = "WETPREP" Then
200                   .SelBold = False
210                   .SelFontSize = 10
220                   .SelText = Space$(10) & Title
230                   .SelBold = True
240                   .SelText = ID.TestName & " " & ID.Result
250                   .SelText = vbCrLf
260                   Title = "            "
270               End If
280           Next
290       End If

300       .SelText = vbCrLf

310       .SelBold = False
320       .SelFontSize = 10
330   End With

340   Exit Sub

RTFPrintMicroGramWetPrep_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "modRTFMicro", "RTFPrintMicroGramWetPrep", intEL, strES

End Sub


Private Sub RTFPrintMicroUrineComment(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim SDS As New SiteDetails

10    On Error GoTo RTFPrintMicroUrineComment_Error

20    SDS.Load Val(SampleID) + sysOptMicroOffset(0)
30    If SDS.Count > 0 Then
40        If UCase$(SDS(1).Site) <> "URINE" Then
50            Exit Sub
60        End If
70    End If

80    sql = "Select * from Isolates where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "' " & _
            "AND OrganismGroup <> 'Negative results' " & _
            "AND OrganismName <> ''"
90    Set tb = New Recordset
100   RecOpenServer 0, tb, sql
110   If tb.EOF Then Exit Sub

120   With frmMain.rtb
130       .SelBold = False
140       .SelFontSize = 10
150       .SelText = "Positive cultures "
160       .SelUnderline = True
170       .SelText = "must"
180       .SelUnderline = False
190       .SelText = " be correlated with signs and symptoms of UTI"
200       .SelText = vbCrLf
210       .SelText = "Particularly with low colony counts" & vbCrLf
220   End With

230   Exit Sub

RTFPrintMicroUrineComment_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "modRTFMicro", "RTFPrintMicroUrineComment", intEL, strES, sql

End Sub


Private Sub RTFPrintMicroAssIDMRSA(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
10    On Error GoTo RTFPrintMicroAssIDMRSA_Error

20    ReDim AssID(0 To 0) As Variant
      Dim ThisID As String
      Dim n As Integer
      Dim S As String
      Dim x As Integer
      Dim Found As Boolean

30    sql = "SELECT AssID FROM AssociatedIDs WHERE " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    n = -1
70    Do While Not tb.EOF
80        n = n + 1
90        ReDim Preserve AssID(0 To n) As Variant
100       AssID(n) = tb!AssID
110       tb.MoveNext
120   Loop
130   sql = "SELECT SampleID FROM AssociatedIDs WHERE " & _
            "AssID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
140   Set tb = New Recordset
150   RecOpenServer 0, tb, sql
160   Do While Not tb.EOF
170       Found = False
180       For x = 0 To UBound(AssID)
190           If AssID(x) = tb!SampleID Then
200               Found = True
210               Exit For
220           End If
230       Next
240       If Not Found Then
250           n = n + 1
260           ReDim Preserve AssID(0 To n) As Variant
270           AssID(n) = tb!SampleID
280       End If
290       tb.MoveNext
300   Loop

310   If n = -1 Then Exit Sub

320   With frmMain.rtb
330       .SelFontSize = 10
340       .SelBold = False

350       .SelText = "This Result relates to the Site specified on this form only."
360       .SelText = vbCrLf
370       .SelText = "Please refer to Results for Lab numbers "
380       S = ""
390       For n = 0 To UBound(AssID)
400           ThisID = Format$(AssID(n) - sysOptMicroOffset(0))
410           S = S & ThisID & ", "
420       Next
430       S = Left$(S, Len(S) - 2)
440       .SelText = S & " as part of this series of screens."
450       .SelText = vbCrLf
460   End With

470   Exit Sub

RTFPrintMicroAssIDMRSA_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "modRTFMicro", "RTFPrintMicroAssIDMRSA", intEL, strES, sql

End Sub

Private Sub RTFPrintMicroAssIDBC(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim AssID As String

10    On Error GoTo RTFPrintMicroAssIDBC_Error

20    sql = "SELECT AssID FROM Demographics WHERE " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If Not tb.EOF Then

60        If Trim$(tb!AssID & "") <> "" Then

70            With frmMain.rtb
80                .SelFontSize = 10
90                .SelBold = False

100               AssID = Format$(tb!AssID - sysOptMicroOffset(0))
110               .SelText = "Please refer to Lab number " & AssID & _
                             " for associated Lab Result."
120               .SelText = vbCrLf
130           End With

140       End If

150   End If

160   Exit Sub

RTFPrintMicroAssIDBC_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "modRTFMicro", "RTFPrintMicroAssIDBC", intEL, strES, sql

End Sub



Private Sub RTFPrintMicroCDiff(ByVal SampleID As String)

      Dim ResultPCR As String
      Dim ResultToxin As String
      Dim Gx As GenericResult
      Dim Gxs As New GenericResults
      Dim Fx As FaecesResult
      Dim Fxs As New FaecesResults

10    On Error GoTo RTFPrintMicroCDiff_Error

20    ResultPCR = ""
30    ResultToxin = ""

40    Gxs.Load Val(SampleID) + sysOptMicroOffset(0)
50    Set Gx = Gxs("cDiffPCR")
60    If Not Gx Is Nothing Then
70        ResultPCR = Gx.Result
80    End If
90    Fxs.Load Val(SampleID) + sysOptMicroOffset(0)
100   Set Fx = Fxs("ToxinAL")
110   If Not Fx Is Nothing Then
120       ResultToxin = Fx.Result
130   End If

140   If ResultPCR = "" And ResultToxin = "" Then
150       Exit Sub
160   End If

170   With frmMain.rtb
180       .SelText = vbCrLf
190       .SelBold = False
200       .SelFontSize = 10

210       If Trim$(ResultPCR) <> "" Then
220           .SelText = Space$(10) & "C. difficile PCR : "
230           .SelBold = True
240           .SelText = ResultPCR
250           .SelText = vbCrLf
260       End If

270       If Trim$(ResultToxin) <> "" Then
280           .SelBold = False
290           .SelText = Space$(10) & "Clostridium difficile Toxin A/B : "
300           .SelBold = True
310           If UCase$(ResultToxin) = "N" Then
320               .SelText = "Not Detected"
330           ElseIf UCase$(ResultToxin) = "P" Then
340               .SelText = "Positive"
350           ElseIf UCase$(ResultToxin) = "I" Then
360               .SelText = "Inconclusive"
370           ElseIf UCase$(ResultToxin) = "R" Then
380               .SelText = "Sample Rejected"
390               .SelBold = False
400               .SelFontSize = 10
410               .SelText = "I wish to remind you that "
420               .SelItalic = True
430               .SelText = "C. difficile"
440               .SelItalic = False
450               .SelText = " should be requested only when there is a high index"
460               .SelText = vbCrLf
470               .SelText = "of suspicion. The clinical details received with this "
480               .SelText = "test request fail to meet the criteria for"
490               .SelText = vbCrLf
500               .SelText = "testing and as such has been deemed unsuitable for analysis. "
510               .SelText = vbCrLf
520               .SelText = "Please refer to the following"
530               .SelText = vbCrLf
540               .SelText = "guidelines for requesting "
550               .SelItalic = True
560               .SelText = "C. difficile"
570               .SelItalic = False
580               .SelText = " toxin testing."
590               .SelText = vbCrLf
600               .SelItalic = True
610               .SelBold = True
620               .SelUnderline = True
630               .SelText = Space$(30) & "C. difficile"
640               .SelItalic = False
650               .SelText = " toxin"
660               .SelText = vbCrLf
670               .SelBold = False
680               .SelUnderline = False

690               .SelText = "Acute onset of loose stools (more than three within a "
700               .SelText = "24-hour period) for two days"
710               .SelText = vbCrLf

720               .SelText = "without another aetiology, onset after >3days in hospital,"
730               .SelText = " and a history of antibiotic use"
740               .SelText = vbCrLf

750               .SelText = "or chemotherapy;"
760               .SelText = vbCrLf

770               .SelBold = True
780               .SelUnderline = True
790               .SelText = Space$(42) & "or"
800               .SelText = vbCrLf
810               .SelBold = False
820               .SelUnderline = False

830               .SelText = "Recurrence of diarrhoea within eight weeks of the end of "
840               .SelText = "previous treatment of "
850               .SelItalic = True
860               .SelText = "C. difficile"
870               .SelText = vbCrLf

880               .SelItalic = False
890               .SelText = "infection."
900               .SelText = vbCrLf
910           Else
920               .SelText = vbCrLf
930           End If
940       End If
950   End With

960   UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "CDIFF", True, True

970   Exit Sub

RTFPrintMicroCDiff_Error:

      Dim strES As String
      Dim intEL As Integer

980   intEL = Erl
990   strES = Err.Description
1000  LogError "modRTFMicro", "RTFPrintMicroCDiff", intEL, strES

End Sub
Private Sub ClearPrintingRule(ByRef rtb As RichTextBox)
On Error GoTo ClearPrintingRule_Error

rtb.SelBold = False
rtb.SelItalic = False
rtb.SelUnderline = False

Exit Sub
ClearPrintingRule_Error:
   
LogError "modRTFMicro", "ClearPrintingRule", Erl, Err.Description


End Sub

Private Sub ApplyPrintingRule(ByRef rtb As RichTextBox, ByVal TestName As String, ByVal TestType As String, ByVal Value As String)

Dim tb As Recordset
Dim sql As String

On Error GoTo ApplyPrintingRule_Error

sql = "SELECT * FROM PrintingRules WHERE TestName = '" & TestName & "' AND Type = '" & TestType & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
    If tb!Criteria & "" = Value Then
        rtb.SelBold = tb!Bold
        rtb.SelItalic = tb!Italic
        rtb.SelUnderline = tb!Underline
    End If
End If

Exit Sub
ApplyPrintingRule_Error:
   
LogError "modRTFMicro", "ApplyPrintingRule", Erl, Err.Description, sql


End Sub


Private Sub RTFPrintMicroSensitivities(ByVal SampleID As String, _
                                       ByVal IsolateNumbers As String, _
                                       ByVal CommentsPresent As Boolean)

      Dim tb         As Recordset
      Dim sql        As String
      Dim strGroup(1 To 8) As OrgGroup
      Dim ABCount    As Integer
      Dim n          As Integer
      Dim x          As Integer
      Dim Y          As Integer
      Dim MaxIsolates As Integer
      Dim SampleIDWithOffset As Variant
      Dim SensPrintMax As Integer
      Dim Site       As String
      Dim SDS        As New SiteDetails

10    On Error GoTo RTFPrintMicroSensitivities_Error

20    ReDim ResultArray(0 To 0) As ABResult

30    SampleIDWithOffset = Val(SampleID) + sysOptMicroOffset(0)

40    SensPrintMax = 3

50    SDS.Load SampleIDWithOffset
60    If SDS.Count > 0 Then
70        Site = SDS(1).Site

80        sql = "Select [Default] as D from Lists where " & _
                "ListType = 'SI' " & _
                "and Text = '" & Site & "'"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       If Not tb.EOF Then
120           SensPrintMax = Val(tb!D & "")
130       End If
140   End If

150   LoadResultArray SampleIDWithOffset, ResultArray()

160   ABCount = UBound(ResultArray())

170   MaxIsolates = FillOrgGroups(strGroup(), SampleIDWithOffset)

180   With frmMain.rtb
190       If Not CommentsPresent Then
200           For n = 1 To 10 - ABCount
210               .SelText = vbCrLf
220           Next
230       End If

240       .SelBold = False
250       .SelFontSize = 10
260       .SelUnderline = True
270       .SelText = "Isolates  "
280       .SelUnderline = False
290       .SelText = vbCrLf

300       For x = 1 To 4
310           n = Val(Mid$(IsolateNumbers, x, 1))
320           If strGroup(n).OrgName <> "" Then
330               .SelText = Left$(x & ": " & strGroup(n).ReportName & Space$(40), 40)
340               ApplyPrintingRule frmMain.rtb, strGroup(n).OrgName, "Qualifier", strGroup(n).Qualifier
350               .SelText = strGroup(n).Qualifier
360               ClearPrintingRule frmMain.rtb
370               .SelText = vbCrLf
                  '.SelText = Space$(10)
380           Else
390               .SelText = Left$(strGroup(n).OrgGroup & Space$(40), 40)
400               .SelText = vbCrLf
410           End If
420       Next

430       If ABCount > 0 Then
440           .SelFontSize = 10
              '.SelBold = False
              '.SelText = vbCrLf

450           .SelUnderline = True
460           .SelText = Left$("Sensitivities" & Space$(20), 20)

470           .SelText = IIf(strGroup(1).OrgName <> "", Left$("1" & Space(19), 19), Space(19))
480           .SelText = IIf(strGroup(2).OrgName <> "", Left$("2" & Space(19), 19), Space(19))
490           .SelText = IIf(strGroup(3).OrgName <> "", Left$("3" & Space(19), 19), Space(19))
500           .SelText = IIf(strGroup(4).OrgName <> "", Left$("4" & Space(19), 19), Space(19))
510           .SelText = vbCrLf
520           Select Case IsolateNumbers
                  Case "1234":
530                   .SelText = Left$("" & Space$(20), 20)
540                   .SelText = Left$(strGroup(1).ShortName & Space(19), 19)
550                   .SelText = Left$(strGroup(2).ShortName & Space(19), 19)
560                   .SelText = Left$(strGroup(3).ShortName & Space(19), 19)
570                   .SelText = Left$(strGroup(4).ShortName & Space(19), 19)
580                   .SelText = vbCrLf
590               Case "5678":
600                   .SelText = Left$("" & Space$(20), 20)
610                   .SelText = Left$(strGroup(5).ShortName & Space(19), 19)
620                   .SelText = Left$(strGroup(6).ShortName & Space(19), 19)
630                   .SelText = Left$(strGroup(7).ShortName & Space(19), 19)
640                   .SelText = Left$(strGroup(8).ShortName & Space(19), 19)
650                   .SelText = vbCrLf
660           End Select
670           .SelUnderline = False

              Dim Sxs As New Sensitivities
              Dim sx As Sensitivity

680           Sxs.Load Val(SampleID) + sysOptMicroOffset(0)

690           For Y = 1 To ABCount
700               .SelColor = vbBlack
710               .SelText = Left$(ResultArray(Y).AntibioticName & Space$(20), 20)
                  'IsolateNumbers = "1234"
720               For x = 1 To 4
730                   n = Val(Mid$(IsolateNumbers, x, 1))
740                   If Trim$(strGroup(n).ReportName) <> "" Then
750                       Set sx = Sxs.Item(n, ResultArray(Y).AntibioticCode)
760                       If Not sx Is Nothing Then
770                           If sx.Report Then
780                               If sx.RSI = "R" Then
790                                   .SelColor = vbRed
800                                   .SelText = Left$("Resistant" & Space(19), 19)
810                               ElseIf sx.RSI = "S" Then
820                                   .SelColor = vbGreen
830                                   .SelText = Left$("Sensitive" & Space(19), 19)
840                               ElseIf sx.RSI = "I" Then
850                                   .SelColor = vbBlack
860                                   .SelText = Left$("Intermediate" & Space(19), 19)
870                               Else
880                                   .SelText = Left$(sx.RSI & Space(19), 19)
890                               End If
900                           Else
910                               .SelText = Space(19)
920                           End If
930                       Else
940                           .SelText = Space$(19)
950                       End If
960                   Else
970                       .SelText = Space$(19)
980                   End If
990               Next
1000              .SelText = vbCrLf
1010          Next
1020      End If




1030      UpdatePrintValid SampleIDWithOffset, "CANDS", True, True

1040      .SelColor = vbBlack

1050  End With

1060  Exit Sub

RTFPrintMicroSensitivities_Error:

      Dim strES      As String
      Dim intEL      As Integer

1070  intEL = Erl
1080  strES = Err.Description
1090  LogError "modRTFMicro", "RTFPrintMicroSensitivities", intEL, strES, sql

End Sub


Private Sub RTFPrintMicroBDMax(ByVal SampleID As String, _
                               ByVal IsolateNumbers As String)

      Dim strGroup(1 To 8) As OrgGroup
      Dim ABCount As Integer
      Dim n As Integer
      Dim x As Integer
      Dim MaxIsolates As Integer
      Dim SampleIDWithOffset As Variant

10    ReDim ResultArray(0 To 0) As ABResult

20    SampleIDWithOffset = Val(SampleID) + sysOptMicroOffset(0)

30    MaxIsolates = FillOrgGroups(strGroup(), SampleIDWithOffset)

40    With frmMain.rtb
50        .SelBold = True
60        .SelFontSize = 10
70        .SelText = vbCrLf
80        .SelText = Space$(10)
90        For x = 1 To 4
100           n = Val(Mid$(IsolateNumbers, x, 1))
110           If strGroup(n).OrgName <> "" Then
120               .SelText = Left$(strGroup(n).ReportName & Space$(30), 30)
130               .SelText = strGroup(n).Qualifier
140               .SelText = vbCrLf
150               .SelText = Space$(10)
160           Else
170               .SelText = Left$(strGroup(n).OrgGroup & Space$(30), 30)
180               .SelText = vbCrLf
190           End If
200       Next
210       .SelFontSize = 10
220       .SelBold = False
230       .SelText = vbCrLf

240       UpdatePrintValid SampleIDWithOffset, "CANDS", True, True

250       .SelColor = vbBlack

260   End With


End Sub



Private Sub RTFPrintMicroRotaAdeno(ByVal SampleID As String)

      Dim Fxs As New FaecesResults

10    On Error GoTo RTFPrintMicroRotaAdeno_Error

20    Fxs.Load Val(SampleID) + sysOptMicroOffset(0)
30    If Fxs("Rota") Is Nothing And Fxs("Adeno") Is Nothing Then Exit Sub

40    With frmMain.rtb
50        .SelText = vbCrLf
60        .SelText = vbCrLf

70        .SelBold = False
80        .SelFontSize = 10

90        If Not Fxs("Rota") Is Nothing Then
100           .SelBold = False
110           .SelText = Space$(10) & " Rota Virus : "
120           .SelBold = True
130           If UCase$(Fxs("Rota").Result) = "N" Then
140               .SelText = "Negative"
150           ElseIf UCase$(Fxs("Rota").Result) = "P" Then
160               .SelText = "Positive"
170           End If
180           .SelText = vbCrLf
190       End If

200       If Not Fxs("Adeno") Is Nothing Then
210           .SelBold = False
220           .SelText = Space$(10) & "Adeno Virus : "
230           .SelBold = True
240           If UCase$(Fxs("Adeno").Result) = "N" Then
250               .SelText = "Negative"
260           ElseIf UCase$(Fxs("Adeno").Result) = "P" Then
270               .SelText = "Positive"
280           End If
290           .SelText = vbCrLf
300       End If

310       .SelBold = False
320   End With

330   UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "ROTAADENO", True, True

340   Exit Sub

RTFPrintMicroRotaAdeno_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "modRTFMicro", "RTFPrintMicroRotaAdeno", intEL, strES

End Sub


Private Sub RTFPrintNegativeResults(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim CultPrinted As Boolean

10    On Error GoTo RTFPrintNegativeResults_Error

20    sql = "Select * from Isolates where " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then

60        With frmMain.rtb

70            .SelText = vbCrLf
80            .SelBold = True
90            .SelFontSize = 10

100           CultPrinted = False
110           Do While Not tb.EOF
120               If Not CultPrinted Then
130                   .SelText = Space$(10) & "CULTURE: " & tb!OrganismName & " " & tb!Qualifier & ""
140                   CultPrinted = True
150               Else
160                   .SelText = Space$(10) & "         " & tb!OrganismName & " " & tb!Qualifier & ""
170               End If
180               .SelText = vbCrLf
190               tb.MoveNext
200           Loop
210       End With
220   End If

230   UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "CANDS", True, True

240   Exit Sub

RTFPrintNegativeResults_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "modRTFMicro", "RTFPrintNegativeResults", intEL, strES, sql

End Sub

Public Sub RTFPrintMicroscopy(ByVal SampleID As String)

      Dim URS As New UrineResults
      Dim UR As UrineResult
      Dim Res As String
      Dim n As Integer
      Dim Found As Boolean

10    On Error GoTo RTFPrintMicroscopy_Error

20    URS.Load Val(SampleID) + sysOptMicroOffset(0)
30    If URS.Count = 0 Then
40        Exit Sub
50    End If

60    For Each UR In URS
70        If UR.TestName = "Bacteria" Or _
             UR.TestName = "Crystals" Or _
             UR.TestName = "RCC" Or _
             UR.TestName = "WCC" Or _
             UR.TestName = "Casts" Or _
             UR.TestName = "Misc0" Or _
             UR.TestName = "Misc1" Or _
             UR.TestName = "Misc2" Then
80            Found = True
90            Exit For
100       End If
110   Next
120   If Not Found Then
130       Exit Sub
140   End If

150   With frmMain.rtb
160       .SelBold = False
170       .SelText = "Microscopy" & IIf(SedimexResultsExist, "(Not accredited):", ":") & " Bacteria:"
180       .SelBold = True
190       Res = ""
200       Set UR = URS.Item("Bacteria")
210       If Not UR Is Nothing Then
220           Res = UR.Result
230       End If
240       .SelText = Left$(Res & Space$(14), 14)
250       .SelBold = False
260       .SelText = "Crystals:"
270       .SelBold = True
280       Res = ""
290       Set UR = URS.Item("Crystals")
300       If Not UR Is Nothing Then
310           Res = UR.Result
320       End If
330       If Res = "" Then
340           .SelText = "Nil"
350       Else
360           .SelText = Res
370       End If
380       .SelText = vbCrLf

390       .SelBold = False
400       .SelText = "                 WCC:"
410       .SelBold = True
420       Res = ""
430       Set UR = URS.Item("WCC")
440       If Not UR Is Nothing Then
450           Res = UR.Result
460       End If
470       .SelText = Left$(Res & " /cmm" & Space$(14), 14)
480       .SelBold = False
490       Res = ""
500       Set UR = URS.Item("Casts")
510       If Not UR Is Nothing Then
520           Res = UR.Result
530       End If

540       .SelText = "   Casts:"
550       .SelBold = True
560       If Res = "" Then
570           .SelText = "Nil"
580       Else
590           .SelText = Res
600       End If
610       .SelText = vbCrLf

620       .SelBold = False
630       .SelText = "                 RCC:"
640       .SelBold = True
650       Res = ""
660       Set UR = URS.Item("RCC")
670       If Not UR Is Nothing Then
680           Res = UR.Result
690       End If
700       .SelText = Left$(Res & Space$(14), 14)
710       .SelBold = False
720       .SelText = "    Misc:"
730       .SelBold = True

740       Res = ""
750       For n = 0 To 2
760           Set UR = URS.Item("Misc" & Format$(n))
770           If Not UR Is Nothing Then
780               Res = Res & UR.Result & " "
790           End If
800       Next

810       If Trim$(Res) = "" Then
820           .SelText = "Nil"
830       Else
840           .SelText = Res
850       End If
860       .SelText = vbCrLf
870       .SelBold = False
880   End With

890   UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "URINE", True, True

900   Exit Sub

RTFPrintMicroscopy_Error:

      Dim strES As String
      Dim intEL As Integer

910   intEL = Erl
920   strES = Err.Description
930   LogError "modRTFMicro", "RTFPrintMicroscopy", intEL, strES

End Sub

Private Sub RTFPrintMicroRedSub(ByVal SampleID As String)

      Dim Gxs As New GenericResults
      Dim Gx As GenericResult

10    On Error GoTo RTFPrintMicroRedSub_Error

20    Gxs.Load Val(SampleID) + sysOptMicroOffset(0)
30    If Gxs.Count = 0 Then
40        Exit Sub
50    End If
60    Set Gx = Gxs("RedSub")
70    If Gx Is Nothing Then
80        Exit Sub
90    End If

100   With frmMain.rtb

110       .SelText = vbCrLf
120       .SelBold = False
130       .SelFontSize = 10

140       .SelText = Space$(5) & "Reducing Substances : "
150       .SelBold = True
160       .SelText = Gx.Result
170       .SelText = vbCrLf
180       .SelBold = False
190       .SelFontSize = 10
200   End With

210   UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "REDSUB", True, True

220   Exit Sub

RTFPrintMicroRedSub_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "modRTFMicro", "RTFPrintMicroRedSub", intEL, strES

End Sub

Private Sub RTFPrintMicroCSF(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim S As String
      Dim Found As Boolean

10    On Error GoTo RTFPrintMicroCSF_Error

20    sql = "SELECT * FROM CSFResults WHERE " & _
            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub
60    Found = False
70    If Trim$(tb!Gram & tb!WCCDiff0 & tb!WCCDiff1 & "") <> "" Then
80        Found = True
90    End If
100   If Not Found Then
110       For n = 0 To 2
120           If Trim$(tb("Appearance" & Format$(n)) & _
                       tb("WCC" & Format$(n)) & _
                       tb("RCC" & Format$(n)) & "") <> "" Then
130               Found = True
140               Exit For
150           End If
160       Next
170   End If
180   If Not Found Then Exit Sub

190   With frmMain.rtb
200       .SelFontSize = 10
210       .SelText = vbCrLf
220       .SelBold = True
230       .SelText = "Appearance: "
240       .SelBold = False
250       .SelText = Space$(12) & "Sample 1: " & tb!Appearance0 & ""
260       .SelText = vbCrLf
270       .SelText = Space$(12) & "Sample 2: " & tb!Appearance1 & ""
280       .SelText = vbCrLf
290       .SelText = Space$(12) & "Sample 3: " & tb!Appearance2 & ""
300       .SelText = vbCrLf

310       .SelBold = True
320       .SelText = "Gram Stain: "
330       .SelBold = False
340       .SelText = tb!Gram & ""
350       .SelText = vbCrLf

360       .SelText = "         "
370       .SelUnderline = True
380       .SelText = "Sample 1"
390       .SelUnderline = False
400       .SelText = "        "
410       .SelUnderline = True
420       .SelText = "Sample 2"
430       .SelUnderline = False
440       .SelText = "        "
450       .SelUnderline = True
460       .SelText = "Sample 3"
470       .SelText = vbCrLf
480       .SelUnderline = False

490       .SelBold = True
500       .SelText = "WCC/cmm    "
510       .SelBold = False
520       S = Left$(tb!WCC0 & Space(16), 16) & _
              Left$(tb!WCC1 & Space(16), 16) & _
              tb!WCC2 & ""
530       .SelText = S
540       .SelText = vbCrLf

550       .SelBold = True
560       .SelText = "RCC/cmm    "
570       .SelBold = False

580       S = Left$(tb!RCC0 & Space(16), 16) & _
              Left$(tb!RCC1 & Space(16), 16) & _
              tb!RCC2 & ""
590       .SelText = S
600       .SelText = vbCrLf

610       .SelBold = True
620       .SelText = "White Cell Differential: "
630       .SelBold = False
640       .SelText = tb!WCCDiff0 & " % Neutrophils " & tb!WCCDiff1 & "% Mononuclear Cells"
650       .SelText = vbCrLf

660       UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "FOB", True, True
670   End With

680   Exit Sub

RTFPrintMicroCSF_Error:

      Dim strES As String
      Dim intEL As Integer

690   intEL = Erl
700   strES = Err.Description
710   LogError "modRTFMicro", "RTFPrintMicroCSF", intEL, strES, sql

End Sub


Public Sub RTFPrintSpecType(ByRef RP As ReportToPrint, _
                            ByVal CurrentPage As Integer, _
                            ByVal TotalPages As Integer)

      Dim SiteDetails As String
      Dim Site As String
      Dim SDS As New SiteDetails

10    On Error GoTo RTFPrintSpecType_Error

20    SDS.Load Val(RP.SampleID) + sysOptMicroOffset(0)
30    If SDS.Count > 0 Then
40        Site = SDS(1).Site
50        SiteDetails = SDS(1).SiteDetails
60    End If

70    With frmMain.rtb
80        .SelColor = vbBlack
90        .SelFontSize = 10
100       .SelBold = False
110       .SelText = "Specimen Type:"
120       .SelBold = True
130       .SelText = Site & " " & SiteDetails & " "

140       .SelColor = vbRed
          '150       If TotalPages > 1 Then
          '160           .SelText = "Page " & CurrentPage & " of " & TotalPages & " "
          '170       End If
150       If RP.ThisIsCopy Then
160           .SelBold = False
170           .SelText = "This is a COPY Report for Attention of "
180           .SelBold = True
190           .SelText = Trim$(RP.SendCopyTo)
200       End If
210       .SelText = vbCrLf

220       .SelColor = vbBlack
230       .SelBold = False
240   End With

250   Exit Sub

RTFPrintSpecType_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "modRTFMicro", "RTFPrintSpecType", intEL, strES

End Sub


Private Sub RTFPrintMicroPregnancy(ByVal SampleID As String)

      Dim URS As New UrineResults
      Dim UR As UrineResult
      Dim Preg As String
      Dim HCG As String

10    On Error GoTo RTFPrintMicroPregnancy_Error

20    Preg = ""
30    URS.Load Val(SampleID) + sysOptMicroOffset(0)
40    If URS.Count = 0 Then Exit Sub
50    Set UR = URS("Pregnancy")
60    If Not UR Is Nothing Then
70        Preg = UR.Result
80    End If

90    HCG = ""
100   Set UR = URS("HCGLevel")
110   If Not UR Is Nothing Then
120       HCG = UR.Result
130   End If

140   If Preg <> "" Or HCG <> "" Then

150       With frmMain.rtb

160           .SelText = vbCrLf

170           .SelBold = False
180           .SelFontSize = 10
190           .SelText = Space$(10) & "Pregnancy Test: "
200           .SelBold = True
210           .SelText = Preg
220           .SelText = vbCrLf
230           .SelBold = False
240           .SelText = Space$(10) & "    HCG Level : "
250           .SelBold = True
260           .SelText = HCG & " IU/L"
270           .SelText = vbCrLf

280           RTFPrintMicroComment RP.SampleID, "Pregnancy Comment"

290           .SelBold = False

300           UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "URINE", True, True

310       End With

320   End If

330   Exit Sub

RTFPrintMicroPregnancy_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "modRTFMicro", "RTFPrintMicroPregnancy", intEL, strES

End Sub


Private Sub RTFPrintMicroRSV(ByVal SampleID As String)

      Dim Gxs As New GenericResults
      Dim Gx As GenericResult

10    On Error GoTo RTFPrintMicroRSV_Error

20    Gxs.Load Val(SampleID) + sysOptMicroOffset(0)
30    If Gxs.Count = 0 Then
40        Exit Sub
50    End If
60    Set Gx = Gxs("RSV")
70    If Gx Is Nothing Then
80        Exit Sub
90    End If

100   With frmMain.rtb
110       .SelText = vbCrLf
120       .SelBold = False
130       .SelFontSize = 10

140       .SelText = Space$(10) & "RSV : "
150       .SelBold = True
160       .SelText = Gx.Result
170       .SelText = vbCrLf
180       .SelBold = False
190   End With

200   UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "RSV", True, True

210   Exit Sub

RTFPrintMicroRSV_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "modRTFMicro", "RTFPrintMicroRSV", intEL, strES

End Sub


Private Sub RTFPrintMicroComment(ByVal SampleID As String, _
                                 ByVal Source As String)

      Dim OBs As Observations
      Dim OffSet As Variant
      Dim n As Integer
      Dim pSource As String

10    On Error GoTo RTFPrintMicroComment_Error

20    ReDim Comments(1 To MicroCommentLineCount) As String

30    OffSet = sysOptMicroOffset(0)

40    Select Case UCase$(Left$(Source, 1))
      Case "D": Source = "Demographic": pSource = "Demographics Comment:"
50    Case "L": Source = "MicroCSAutoComment": pSource = "Auto Comment:"
60    Case "C": Source = "MicroConsultant": pSource = "Consultant Comment:"
70    Case "M": Source = "MicroCS": pSource = "Medical Scientist Comment:"
80    Case "P": Source = "MicroGeneral": pSource = ""
90    Case "I": Source = "Semen": pSource = "Infertility Comment:": OffSet = sysOptSemenOffset
100   Case "P": Source = "Semen": pSource = "Post Vasectomy": OffSet = sysOptSemenOffset

110   End Select

120   Set OBs = New Observations
130   Set OBs = OBs.Load(Val(SampleID + OffSet), Source)

140   With frmMain.rtb
150       .SelFontName = "Courier New"
160       .SelFontSize = 9
170       .SelColor = vbBlack
180       If Not OBs Is Nothing Then
190           FillCommentLines pSource & OBs(1).Comment, MicroCommentLineCount, Comments(), 95
200           .SelBold = False
210           .SelText = pSource
220           .SelBold = False
230           .SelText = Mid$(Comments(1), Len(pSource) + 1)
240           .SelText = vbCrLf
250           For n = 2 To MicroCommentLineCount
260               If Trim$(Comments(n)) <> "" Then
270                   .SelText = Comments(n)
280                   .SelText = vbCrLf
290                   .SelColor = vbBlack
300               End If
310           Next
320       End If
330   End With




      '
      '    Set OBs = New Observations
      '    Set OBs = OBs.Load(Val(SampleID + OffSet), "MicroCSAutoComment")
      '
      '    With frmMain.rtb
      '        .SelFontName = "Courier New"
      '        .SelFontSize = 9
      '        .SelColor = vbBlack
      '        If Not OBs Is Nothing Then
      '            FillCommentLines "Micro Auto Comment:" & OBs(1).Comment, 4, Comments(), 95
      '            .SelBold = False
      '            .SelText = pSource
      '            .SelBold = False
      '            .SelText = Mid$(Comments(1), Len(pSource) + 1)
      '            .SelText = vbCrLf
      '            For n = 2 To 4
      '                If Trim$(Comments(n)) <> "" Then
      '                    .SelText = Comments(n)
      '                    .SelText = vbCrLf
      '                    .SelColor = vbBlack
      '                End If
      '            Next
      '        End If
      '    End With





340   Exit Sub

RTFPrintMicroComment_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "modRTFMicro", "RTFPrintMicroComment", intEL, strES

End Sub
Private Sub RTFPrintMicroOvaParasites(ByVal SampleID As String)

      Dim n As Integer
      Dim blnRejectFound As Boolean
      Dim blnHeadingPrinted As Boolean
      Dim Fx As FaecesResult
      Dim Fxs As New FaecesResults

10    On Error GoTo RTFPrintMicroOvaParasites_Error

20    Fxs.Load Val(SampleID) + sysOptMicroOffset(0)
30    If Fxs.Count = 0 Then
40        Exit Sub
50    End If

60    With frmMain.rtb
70        .SelFontSize = 10
80        .SelText = vbCrLf

90        Set Fx = Fxs.Item("AUS")
100       If Not Fx Is Nothing Then
110           .SelBold = False
120           .SelText = Space$(10) & "Cryptosporidium : "
130           .SelBold = True
140           If UCase$(Fx.Result) = "N" Then
150               .SelText = "Negative"
160           ElseIf UCase$(Fx.Result) = "P" Then
170               .SelText = "Positive"
180           End If
190           .SelText = vbCrLf
200           .SelBold = False
210       End If

220       blnRejectFound = False
230       For n = 0 To 2
240           Set Fx = Fxs.Item("OP" & Format$(n))
250           If Not Fx Is Nothing Then
260               If InStr(UCase$(Fx.Result), "REJECTED") <> 0 Then
270                   blnRejectFound = True
280                   Exit For
290               End If
300           End If
310       Next
320       If blnRejectFound Then
330           .SelText = Space$(10) & "Ova and Parasites : "
340           .SelBold = True
350           .SelText = "Sample Rejected"
360           .SelText = vbCrLf
370           .SelBold = False
380           .SelFontSize = 10
390           .SelText = "I wish to remind you that "
400           .SelItalic = True
410           .SelText = "Ova and Parasites"
420           .SelItalic = False
430           .SelText = " should be requested only when there is a"
440           .SelText = vbCrLf
450           .SelText = "high index of suspicion. The clinical details received "
460           .SelText = "with this test request fail to meet the"
470           .SelText = vbCrLf
480           .SelText = "criteria for testing and as such has been deemed "
490           .SelText = "unsuitable for analysis. Please refer to"
500           .SelText = vbCrLf
510           .SelText = "the following guidelines for requesting "
520           .SelItalic = True
530           .SelText = "Ova and Parasites."
540           .SelText = vbCrLf
550           .SelItalic = False

560           .SelBold = True
570           .SelUnderline = True
580           .SelText = "Ova and Parasites"
590           .SelText = vbCrLf
600           .SelUnderline = False
610           .SelBold = False

620           .SelText = "Submit one stool sample if: Persistent diarrhoea > 7 days ;"
630           .SelText = vbCrLf
640           .SelText = Space$(25)
650           .SelBold = True
660           .SelUnderline = True
670           .SelText = "or"
680           .SelUnderline = False
690           .SelBold = False
700           .SelText = " Patient is immunocompromised;"
710           .SelText = vbCrLf
720           .SelText = Space$(25)
730           .SelBold = True
740           .SelUnderline = True
750           .SelText = "or"
760           .SelUnderline = False
770           .SelBold = False

780           .SelText = " Patient has visited a developing country"

790       Else
800           blnHeadingPrinted = False
810           For n = 0 To 2
820               Set Fx = Fxs.Item("OP" & Format(n))
830               If Not Fx Is Nothing Then
840                   .SelText = vbCrLf
850                   .SelBold = False
860                   If Not blnHeadingPrinted Then
870                       .SelText = Space$(10) & "Ova and Parasites : "
880                       blnHeadingPrinted = True
890                   Else
900                       .SelText = Space$(10) & "                    "
910                   End If
920                   .SelBold = True
930                   .SelText = Trim$(Fx.Result)
940                   .SelText = vbCrLf
950                   .SelBold = False
960               End If
970           Next
980           .SelText = vbCrLf
990       End If

1000      .SelFontSize = 10
1010  End With

1020  UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "OP", True, True

1030  Exit Sub

RTFPrintMicroOvaParasites_Error:

      Dim strES As String
      Dim intEL As Integer

1040  intEL = Erl
1050  strES = Err.Description
1060  LogError "modRTFMicro", "RTFPrintMicroOvaParasites", intEL, strES

End Sub


Private Sub RTFPrintMicroClinDetails(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim CLD As String

10    On Error GoTo RTFPrintMicroClinDetails_Error

20    With frmMain.rtb
30        .SelFontSize = 10
40        .SelText = "Clinical Details:"
50        sql = "Select ClDetails from Demographics where " & _
                "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql
80        If Not tb.EOF Then
90            CLD = tb!ClDetails & ""
100           CLD = Replace(CLD, vbCr, " ")
110           CLD = Replace(CLD, vbLf, " ")
120           .SelText = CLD
130       End If
140       .SelText = vbCrLf
150   End With

160   Exit Sub

RTFPrintMicroClinDetails_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "modRTFMicro", "RTFPrintMicroClinDetails", intEL, strES, sql

End Sub


Private Sub RTFPrintMicroOccultBlood(ByVal SampleID As String)

      Dim n As Integer
      Dim Fx As FaecesResult
      Dim Fxs As New FaecesResults

10    On Error GoTo RTFPrintMicroOccultBlood_Error

20    Fxs.Load Val(SampleID) + sysOptMicroOffset(0)
30    If Fxs.Count = 0 Then
40        Exit Sub
50    End If

60    With frmMain.rtb
70        .SelFontSize = 10
80        .SelText = vbCrLf

90        For n = 0 To 2
100           Set Fx = Fxs.Item("OB" & Format$(n))
110           If Not Fx Is Nothing Then
120               If Trim$(Fx.Result) <> "" Then
130                   .SelBold = False
140                   .SelText = Space$(10) & "Occult Blood ( " & Format$(n + 1) & " ) : "
150                   .SelBold = True
160                   If UCase$(Fx.Result) = "N" Then
170                       .SelText = "Negative"
180                   ElseIf UCase$(Fx.Result) = "P" Then
190                       .SelText = "Positive"
200                   End If
210                   .SelText = vbCrLf
220               End If
230           End If
240       Next
250       .SelBold = False
260       .SelFontSize = 10
270   End With

280   UpdatePrintValid Val(SampleID) + sysOptMicroOffset(0), "FOB", True, True

290   Exit Sub

RTFPrintMicroOccultBlood_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "modRTFMicro", "RTFPrintMicroOccultBlood", intEL, strES

End Sub
Private Sub RTFPrintMicroCurrentABs(ByVal SampleID As String)

      Dim S As String
      Dim CurABs As New CurrentAntibiotics
      Dim CurAB As CurrentAntibiotic

10    On Error GoTo RTFPrintMicroCurrentABs_Error

20    CurABs.Load Val(SampleID) + sysOptMicroOffset(0)

30    S = ""
40    For Each CurAB In CurABs
50        S = CurAB.Antibiotic & " "
60    Next
70    With frmMain.rtb
80        .SelFontSize = 10
90        .SelText = "Current Antibiotics:" & S
100       .SelText = vbCrLf
110       .SelFontSize = 4
120       .SelAlignment = rtfCenter
130       .SelText = String$(200, "-")
140       .SelText = vbCrLf
150       .SelFontSize = 10
160       .SelAlignment = rtfLeft
170   End With

180   Exit Sub

RTFPrintMicroCurrentABs_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modRTFMicro", "RTFPrintMicroCurrentABs", intEL, strES

End Sub

