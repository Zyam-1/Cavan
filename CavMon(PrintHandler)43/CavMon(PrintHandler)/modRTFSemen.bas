Attribute VB_Name = "modRTFSemen"
Option Explicit

Public Sub RTFPrintSAReport()

    Dim tb As Recordset
    Dim sql As String
    Dim Dob As String
    Dim SampleDate As String
    Dim ReceivedDate As String
    Dim Rundate As String
    Dim PrintTime As String
    Dim Srs As New SemenResults
    Dim SR As SemenResult

10  On Error GoTo RTFPrintSAReport_Error

20  If InStr(frmMain.rtb.TextRTF, RP.SampleID) Then 'Double check that report saved is for this Sample Id
30      Srs.Load RP.SampleID + sysOptSemenOffset
40      If Srs.Count > 0 Then
50          Set SR = Srs("SpecimenType")
60          If Not SR Is Nothing Then
70              If Not SR.Valid = 1 Then
80                  Exit Sub
90              End If
100         End If
110     End If

120     PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

130     If Not SetPrinter("CHSEMEN") Then Exit Sub

140     sql = "SELECT PatName, Dob, Chart, Addr0, Addr1, Sex, Hospital, RunDate, SampleDate, ClDetails FROM Demographics WHERE SampleID = '" & SR.SampleID & "'"
150     Set tb = New Recordset
160     RecOpenServer 0, tb, sql
170     If tb.EOF Then
180         Exit Sub
190     End If

200     Dob = Format(tb!Dob, "dd/MMM/yyyy")
210     Rundate = Format(tb!Rundate, "dd/MMM/yyyy")
220     SampleDate = Format(tb!SampleDate, "dd/MMM/yyyy hh:mm:ss")
230     ReceivedDate = Format(tb!RecDate, "dd/MMM/yyyy hh:mm:ss")
240     RTFPrintHeading "Microbiology", tb!PatName, Dob, tb!Chart, tb!Addr0, tb!Addr1, tb!Sex, tb!Hospital

250     RTFPrintText vbCrLf
260     RTFPrintText "Cl Details:", , True
270     RTFPrintText tb!ClDetails & vbCrLf
280     RTFPrintSemenComment SR.SampleID, "D"
290     RTFPrintText vbCrLf

300     If SR.Result = "Infertility Analysis" Then
310         RTFPrintSAInfertility Srs
320     ElseIf SR.Result = "Post Vasectomy" Then
330         RTFPrintSAVasectomy
340     End If

350     RTFPrintFooter "Semen", RP.Initiator, SampleDate, Rundate, ReceivedDate

360     With frmMain.rtb
370         .SelStart = 0
380         .SelPrint Printer.hDC

390         sql = "INSERT INTO Reports " & _
                  "(SampleID, Dept, Initiator, PrintTime, ReportNumber, PageNumber, Report, Printer) " & _
                  "VALUES " & _
                  "( '" & RP.SampleID & "', " & _
                  "  'Semen', " & _
                  "  '" & RP.Initiator & "', " & _
                  "   getdate(), " & _
                  "  '" & RP.SampleID & "S" & "', " & _
                  "  '1', " & _
                  "  '" & AddTicks(.TextRTF) & "', " & _
                  "  '" & Printer.DeviceName & "')"
400         Cnxn(0).Execute sql
410     End With

420     For Each SR In Srs
430         SR.Printed = 1
440         SR.PrintedBy = RP.Initiator
450         SR.PrintedDateTime = Now
460         SR.Save
470     Next
480 Else
490     LogError "modRTF", "PrintAndStore", 490, RP.SampleID & " not found in RTF report", sql
500 End If

510 Exit Sub

RTFPrintSAReport_Error:

    Dim strES As String
    Dim intEL As Integer

520 intEL = Erl
530 strES = Err.Description
540 LogError "modRTFSemen", "RTFPrintSAReport", intEL, strES, sql

End Sub
Private Sub RTFPrintSAInfertility(ByVal Srs As SemenResults)

    Dim SR As SemenResult
    Dim Result As String

10  On Error GoTo RTFPrintSAInfertility_Error

20  RTFPrintText FormatString("Specimen Type : ", 20, , AlignRight)
30  RTFPrintText FormatString("Semen Infertility Analysis", 30, , AlignLeft), , True
40  RTFPrintText vbCrLf & vbCrLf

50  RTFPrintText Space(15)
60  RTFPrintText FormatString("Test Values", 15, , AlignCenter), , True, , True
70  RTFPrintText Space(30)
80  RTFPrintText FormatString("Reference Value", 15, , AlignCenter), , True, , True
90  RTFPrintText vbCrLf & vbCrLf

100 Result = ""
110 Set SR = Srs("pH")
120 If Not SR Is Nothing Then
130     Result = SR.Result
140 End If
150 RTFPrintText FormatString("pH: ", 15, , AlignRight), , True
160 RTFPrintText Left$(Result & Space(40), 40)
170 RTFPrintText "(pH:     7.2 or more)"
180 RTFPrintText vbCrLf

190 Result = ""
200 Set SR = Srs("Volume")
210 If Not SR Is Nothing Then
220     Result = SR.Result
230 End If
240 RTFPrintText FormatString("Volume: ", 15, , AlignRight), , True
250 RTFPrintText Left$(Result & Space$(7), 7)
260 RTFPrintText Left$("mls" & Space$(33), 33)
270 RTFPrintText "(Volume: >2.0 mls)"
280 RTFPrintText vbCrLf

290 Result = ""
300 Set SR = Srs("Consistency")
310 If Not SR Is Nothing Then
320     Result = SR.Result
330 End If
340 RTFPrintText FormatString("Viscosity: ", 15, , AlignRight), , True
350 RTFPrintText Result
360 RTFPrintText vbCrLf

370 RTFPrintText FormatString("Motility: ", 15, , AlignRight), , True
380 RTFPrintText Space(40)
390 RTFPrintText "(Motility: % Grades A+B >50%)"
400 RTFPrintText vbCrLf

410 Result = ""
420 Set SR = Srs("GradeA")
430 If Not SR Is Nothing Then
440     Result = SR.Result
450 End If
460 RTFPrintText FormatString("  Grade A: ", 15, , AlignRight), , True
470 RTFPrintText Left$(Result & Space$(7), 7)
480 RTFPrintText "% (Fast progressive)"
490 RTFPrintText vbCrLf

500 Result = ""
510 Set SR = Srs("GradeB")
520 If Not SR Is Nothing Then
530     Result = SR.Result
540 End If
550 RTFPrintText FormatString("  Grade B: ", 15, , AlignRight), , True
560 RTFPrintText Left$(Result & Space$(7), 7)
570 RTFPrintText "% (Slow progressive)"
580 RTFPrintText vbCrLf

590 Result = ""
600 Set SR = Srs("GradeC")
610 If Not SR Is Nothing Then
620     Result = SR.Result
630 End If
640 RTFPrintText FormatString("  Grade C: ", 15, , AlignRight), , True
650 RTFPrintText Left$(Result & Space$(7), 7)
660 RTFPrintText "% (motile non progressive)"
670 RTFPrintText vbCrLf

680 Result = ""
690 Set SR = Srs("GradeD")
700 If Not SR Is Nothing Then
710     Result = SR.Result
720 End If
730 RTFPrintText FormatString("  Grade D: ", 15, , AlignRight), , True
740 RTFPrintText Left$(Result & Space$(7), 7)
750 RTFPrintText "% (non motile)"
760 RTFPrintText vbCrLf

770 Result = ""
780 Set SR = Srs("Morphology")
790 If Not SR Is Nothing Then
800     Result = SR.Result
810 End If
820 RTFPrintText FormatString("Morphology: ", 15, , AlignRight), , True
830 RTFPrintText Left$(Result & Space$(7), 7)
840 RTFPrintText Left$("% Normal" & Space$(33), 33)
850 RTFPrintText "(Morphology: >15% Normal)"
860 RTFPrintText vbCrLf

870 Result = ""
880 Set SR = Srs("SemenCount")
890 If Not SR Is Nothing Then
900     Result = SR.Result
910 End If
920 RTFPrintText FormatString("Sperm Count: ", 15, , AlignRight), , True
930 RTFPrintText Left$(Result & Space$(7), 7)
940 RTFPrintText Left$("million/ml" & Space$(33), 33)
950 RTFPrintText "(Sperm Count: >20 million/ml)"
960 RTFPrintText vbCrLf & vbCrLf

970 RTFPrintSemenComment Srs("SpecimenType").SampleID, "I"

980 RTFPrintText "Semen Analysis Test Values lower than the Reference Values are ", 9
990 RTFPrintText "ASSOCIATED", 9, True, , True
1000 RTFPrintText " with decreased Fertility.", 9

1010 Exit Sub

RTFPrintSAInfertility_Error:

    Dim strES As String
    Dim intEL As Integer

1020 intEL = Erl
1030 strES = Err.Description
1040 LogError "modRTFMicro", "RTFPrintSAInfertility", intEL, strES

End Sub


Private Sub RTFPrintSAVasectomy()

10  On Error GoTo RTFPrintSAVasectomy_Error

20  RTFPrintText FormatString("Specimen Type : ", 20, , AlignRight)
30  RTFPrintText FormatString("Semen Post Vasectomy Analysis", 40, , AlignLeft), , True
40  RTFPrintText vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf

50  RTFPrintSemenComment RP.SampleID + sysOptSemenOffset, "P"

60  Exit Sub

RTFPrintSAVasectomy_Error:

    Dim strES As String
    Dim intEL As Integer

70  intEL = Erl
80  strES = Err.Description
90  LogError "modRTFMicro", "RTFPrintSAVasectomy", intEL, strES

End Sub



Private Sub RTFPrintSemenComment(ByVal SampleIDWithOffset As String, _
                                 ByVal Source As String)

    Dim n As Integer
    Dim pSource As String
    Dim OBs As Observations

10  On Error GoTo RTFPrintSemenComment_Error

20  ReDim Comments(1 To 4) As String

30  Select Case UCase$(Left$(Source, 1))
        Case "I": Source = "Semen": pSource = "Infertility Comment:"
40      Case "P": Source = "Semen": pSource = "Post Vasectomy:"
50      Case "D": Source = "Demographic": pSource = "Demographic Comment:"
60  End Select

70  Set OBs = New Observations
80  Set OBs = OBs.Load(SampleIDWithOffset, Source)
90  If Not OBs Is Nothing Then
100     With frmMain.rtb
110         .SelFontSize = 10
120         FillCommentLines pSource & OBs.Item(1).Comment, 4, Comments(), 97
130         .SelBold = True
140         .SelText = pSource
150         .SelBold = False
160         .SelText = Mid$(Comments(1), Len(pSource) + 1)
170         .SelText = vbCrLf
180         For n = 2 To 4
190             If Trim$(Comments(n)) <> "" Then
200                 .SelText = Comments(n)
210                 .SelText = vbCrLf
220             End If
230         Next
240     End With
250 End If

260 Exit Sub

RTFPrintSemenComment_Error:

    Dim strES As String
    Dim intEL As Integer

270 intEL = Erl
280 strES = Err.Description
290 LogError "modRTFMicro", "RTFPrintSemenComment", intEL, strES

End Sub




