Attribute VB_Name = "modSpecificCavan"
Option Explicit
Public Sub PrintLabelCavan(ByRef RowNumbersToPrint() As Integer)

      Dim y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim sql As String
      Dim RecsEff As Long

10    On Error GoTo PrintLabelCavan_Error

20    OriginalPrinter = Printer.DeviceName
30    If Not SetLabelPrinter() Then Exit Sub

40    With frmXMLabel
50        Printer.Font.Name = "Courier New"

60        For y = 1 To UBound(RowNumbersToPrint)

70            sql = "Update Latest " & _
                    "Set Operator = '" & UserCode & "' where " & _
                    "Number = '" & .g.TextMatrix(RowNumbersToPrint(y), 0) & "' " & _
                    "and BarCode = '04333' " & _
                    "and Operator = 'Auto' "
80            CnxnBB(0).Execute sql, RecsEff
90            If RecsEff <> 0 Then
100               sql = "Insert into Product " & _
                        " Select * from Latest where " & _
                        " Number = '" & .g.TextMatrix(RowNumbersToPrint(y), 0) & "' " & _
                        " and BarCode = '04333' "
110               CnxnBB(0).Execute sql
120               .g.TextMatrix(RowNumbersToPrint(y), 4) = UserCode
130               .g.Refresh
140           End If

150           Printer.Font.Size = 10
160           Printer.Font.Bold = False

170           Printer.Print .g.TextMatrix(RowNumbersToPrint(y), 3)    'product

180           Printer.Print "Unit No ";
190           Printer.Font.Bold = True
200           Printer.Print .g.TextMatrix(RowNumbersToPrint(y), 0);    'unit no
210           Printer.Font.Bold = False
220           Printer.Print Tab(20); "Expiry ";
230           Printer.Font.Bold = True
240           Printer.Print .g.TextMatrix(RowNumbersToPrint(y), 2)  'expiry

250           Printer.Print "Group ";
260           Printer.Font.Size = 14
270           Printer.Print Tab(10); .g.TextMatrix(RowNumbersToPrint(y), 1)    'GroupRh

280           Printer.Font.Bold = False
290           Printer.Font.Size = 10
300           Printer.Print .tcompat

310           Printer.Print

320           Printer.Print "Name ";
330           Printer.Font.Size = 14
340           Printer.Font.Bold = True
350           Printer.Print UCase$(.txtName)

360           Printer.Font.Size = 10
370           Printer.Print
380           Printer.Font.Bold = False
390           Printer.Print "DoB "; .tdob;
400           Printer.Print "    Hosp No "; .txtChart

410           Printer.Print "Spec Date "; .tSampleDate;
420           Printer.Print "   Ward "; .tward

430           Printer.Print "Spec Grp "; .tgroup;
440           Printer.Print "     Lab Ref No "; .lLabNumber

450           Printer.Print "Issued by ";
460           Printer.Print .g.TextMatrix(RowNumbersToPrint(y), 4);    'xmatch by
470           Printer.Print " on ";
480           Printer.Print Format$(.g.TextMatrix(RowNumbersToPrint(y), 5), "dd/mm/yyyy");  'date/time
490           Printer.Print " at ";
500           Printer.Print Format$(.g.TextMatrix(RowNumbersToPrint(y), 5), "hh:mm")

510           Printer.Print "Checked in ward by "

520           Printer.EndDoc


              '########################################

530           Printer.Font.Size = 10
540           Printer.Font.Bold = False

550           Printer.Print .g.TextMatrix(RowNumbersToPrint(y), 3)    'product

560           Printer.Print "Unit No ";
570           Printer.Font.Bold = True
580           Printer.Print .g.TextMatrix(RowNumbersToPrint(y), 0);    'unit no
590           Printer.Font.Bold = False
600           Printer.Print Tab(20); "Expiry ";
610           Printer.Font.Bold = True
620           Printer.Print .g.TextMatrix(RowNumbersToPrint(y), 2)  'expiry

630           Printer.Print "Group ";
640           Printer.Font.Size = 14
650           Printer.Print Tab(10); .g.TextMatrix(RowNumbersToPrint(y), 1)    'GroupRh

660           Printer.Font.Bold = False
670           Printer.Font.Size = 10

680           Printer.Print

690           Printer.Print "For ";
700           Printer.Font.Size = 14
710           Printer.Font.Bold = True
720           Printer.Print UCase$(.txtName)

730           Printer.Font.Size = 10
740           Printer.Print
750           Printer.Font.Bold = False
760           Printer.Print "DoB "; .tdob;
770           Printer.Print "    Hosp No "; .txtChart

780           Printer.Print "Spec Date "; .tSampleDate;
790           Printer.Print "   Ward "; .tward

800           Printer.Print "Spec Grp "; .tgroup;
810           Printer.Print "   Lab Ref No "; .lLabNumber

820           Printer.Print
830           Printer.Print "Removed from Blood Bank"
840           Printer.Print
850           Printer.Print "By:"
860           Printer.Print
870           Printer.Print "Time:        Date:"
880           Printer.EndDoc
890       Next

900   End With

910   For Each Px In Printers
920       If Px.DeviceName = OriginalPrinter Then
930           Set Printer = Px
940           Exit For
950       End If
960   Next

970   Exit Sub

PrintLabelCavan_Error:

      Dim strES As String
      Dim intEL As Integer

980   intEL = Erl
990   strES = Err.Description
1000  LogError "modSpecificCavan", "PrintLabelCavan", intEL, strES, sql


End Sub
Public Sub PrintLabelCavanWithPreview(ByRef RowNumbersToPrint() As Integer, _
                                      ByVal blnpreview As Boolean)

      Dim y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim sql As String
      Dim RecsEff As Long
      Dim f As Form
      Dim s As String

10    On Error GoTo PrintLabelCavanWithPreview_Error

20    OriginalPrinter = Printer.DeviceName
30    If Not SetLabelPrinter() Then Exit Sub

40    Set f = New frmPreviewRTF
50    f.Dept = "XL"
60    f.SampleID = frmXMLabel.lLabNumber
70    f.AdjustPaperSize "110x54"
80    f.Clear

90    With frmXMLabel.g

100       For y = 1 To UBound(RowNumbersToPrint)

110           sql = "Update Latest " & _
                    "Set Operator = '" & UserCode & "' where " & _
                    "Number = '" & .TextMatrix(RowNumbersToPrint(y), 0) & "' " & _
                    "and BarCode = '04333' " & _
                    "and Operator = 'Auto'"
120           CnxnBB(0).Execute sql, RecsEff
130           If RecsEff <> 0 Then
140               sql = "Insert into Product " & _
                        " Select * from Latest where " & _
                        " Number = '" & .TextMatrix(RowNumbersToPrint(y), 0) & "' " & _
                        " and BarCode = '04333'"
150               CnxnBB(0).Execute sql
160               .TextMatrix(RowNumbersToPrint(y), 4) = UserCode
170               .Refresh
180           End If

190           f.WriteFormattedText .TextMatrix(RowNumbersToPrint(y), 3), False, 10, vbBlack, , "Courier New"

200           f.WriteFormattedText "Unit No ;", False, 10, vbBlack, , "Courier New"
210           s = Left$(.TextMatrix(RowNumbersToPrint(y), 0) & Space$(12), 12) & ";"
220           f.WriteFormattedText s, True, 10, vbBlack, , "Courier New"
230           f.WriteFormattedText "Expiry ;", False, 10, vbBlack, , "Courier New"
240           f.WriteFormattedText .TextMatrix(RowNumbersToPrint(y), 2), True, 10, vbBlack, , "Courier New"

250           f.WriteFormattedText "Group     ;", False, 10, vbBlack, , "Courier New"
260           f.WriteFormattedText .TextMatrix(RowNumbersToPrint(y), 1), True, 14, vbBlack, , "Courier New"

270           f.WriteFormattedText "   " & frmXMLabel.tcompat, False, 10, vbBlack, , "Courier New"

280           f.WriteText vbCrLf

290           f.WriteFormattedText "Name ;", False, 10, vbBlack, , "Courier New"
300           f.WriteFormattedText UCase$(frmXMLabel.txtName), True, 14, vbBlack, , "Courier New"

310           f.WriteFormattedText "DoB ;", False, 10, vbBlack, , "Courier New"
320           f.WriteFormattedText frmXMLabel.tdob & ";", True, 10, vbBlack, , "Courier New"
330           f.WriteFormattedText "    Hosp No ;", False, 10, vbBlack, , "Courier New"
340           f.WriteFormattedText frmXMLabel.txtChart, True, 10, vbBlack, , "Courier New"

350           f.WriteFormattedText "Spec Date ;", False, 10, vbBlack, , "Courier New"
360           f.WriteFormattedText frmXMLabel.tSampleDate & ";", True, 10, vbBlack, , "Courier New"
370           f.WriteFormattedText "    Ward ;", False, 10, vbBlack, , "Courier New"
380           f.WriteFormattedText frmXMLabel.tward, True, 10, vbBlack, , "Courier New"

390           f.WriteFormattedText "Spec Grp ;", False, 10, vbBlack, , "Courier New"
400           f.WriteFormattedText frmXMLabel.tgroup & ";", True, 10, vbBlack, , "Courier New"
410           f.WriteFormattedText "     Lab Ref No ;", False, 10, vbBlack, , "Courier New"
420           f.WriteFormattedText frmXMLabel.lLabNumber, True, 10, vbBlack, , "Courier New"

430           f.WriteFormattedText "Issued by ;", False, 10, vbBlack, , "Courier New"
440           f.WriteFormattedText .TextMatrix(RowNumbersToPrint(y), 4) & ";", False, 10, vbBlack, , "Courier New"
450           f.WriteFormattedText " on ;", False, 10, vbBlack, , "Courier New"
460           f.WriteFormattedText Format$(.TextMatrix(RowNumbersToPrint(y), 5), "dd/mm/yyyy") & ";", False, 10, vbBlack, , "Courier New"
470           f.WriteFormattedText " at ;", False, 10, vbBlack, , "Courier New"
480           f.WriteFormattedText Format$(.TextMatrix(RowNumbersToPrint(y), 5), "hh:mm"), False, 10, vbBlack, , "Courier New"

490           f.WriteFormattedText "Checked in ward by", False, 10, vbBlack, , "Courier New"

500           f.ForceNewPage

              '########################################

510           f.WriteFormattedText .TextMatrix(RowNumbersToPrint(y), 3), False, 10, vbBlack, , "Courier New"

520           f.WriteFormattedText "Unit No ;", False, 10, vbBlack, , "Courier New"
530           s = Left$(.TextMatrix(RowNumbersToPrint(y), 0) & Space$(12), 12) & ";"
540           f.WriteFormattedText s, True, 10, vbBlack, , "Courier New"
550           f.WriteFormattedText "Expiry ;", False, 10, vbBlack, , "Courier New"
560           f.WriteFormattedText .TextMatrix(RowNumbersToPrint(y), 2), True, 10, vbBlack, , "Courier New"

570           f.WriteFormattedText "Group     ;", False, 10, vbBlack, , "Courier New"
580           f.WriteFormattedText .TextMatrix(RowNumbersToPrint(y), 1), True, 14, vbBlack, , "Courier New"

590           f.WriteText vbCrLf

600           f.WriteFormattedText "For ;", False, 10, vbBlack, , "Courier New"
610           f.WriteFormattedText UCase$(frmXMLabel.txtName), True, 14, vbBlack, , "Courier New"

620           f.WriteFormattedText "DoB ;", False, 10, vbBlack, , "Courier New"
630           f.WriteFormattedText frmXMLabel.tdob & ";", True, 10, vbBlack, , "Courier New"
640           f.WriteFormattedText "    Hosp No ;", False, 10, vbBlack, , "Courier New"
650           f.WriteFormattedText frmXMLabel.txtChart, True, 10, vbBlack, , "Courier New"

660           f.WriteFormattedText "Spec Date ;", False, 10, vbBlack, , "Courier New"
670           f.WriteFormattedText frmXMLabel.tSampleDate & ";", True, 10, vbBlack, , "Courier New"
680           f.WriteFormattedText "    Ward ;", False, 10, vbBlack, , "Courier New"
690           f.WriteFormattedText frmXMLabel.tward, True, 10, vbBlack, , "Courier New"

700           f.WriteFormattedText "Spec Grp ;", False, 10, vbBlack, , "Courier New"
710           f.WriteFormattedText frmXMLabel.tgroup & ";", True, 10, vbBlack, , "Courier New"
720           f.WriteFormattedText "     Lab Ref No ;", False, 10, vbBlack, , "Courier New"
730           f.WriteFormattedText frmXMLabel.lLabNumber, True, 10, vbBlack, , "Courier New"

740           f.WriteText vbCrLf

750           f.WriteFormattedText "Removed from Blood Bank", False, 10, vbBlack, , "Courier New"
760           f.WriteText vbCrLf
770           f.WriteFormattedText "By:", False, 10, vbBlack, , "Courier New"
780           f.WriteText vbCrLf
790           f.WriteFormattedText "Time:        Date:", False, 10, vbBlack, , "Courier New"

800           If y < UBound(RowNumbersToPrint) Then
810               f.ForceNewPage
820           End If

830       Next

840   End With

850   If blnpreview Then
860       f.Show 1
870   Else
880       f.PrintRTB
890   End If

900   Set f = Nothing

910   For Each Px In Printers
920       If Px.DeviceName = OriginalPrinter Then
930           Set Printer = Px
940           Exit For
950       End If
960   Next

970   Exit Sub

PrintLabelCavanWithPreview_Error:

      Dim strES As String
      Dim intEL As Integer

980   intEL = Erl
990   strES = Err.Description
1000  LogError "modSpecificCavan", "PrintLabelCavanWithPreview", intEL, strES, sql


End Sub
Public Sub PrintANFormCavan(ByVal SampleID As String, _
                            ByVal SampleDate As String)

          Dim TempTb As Recordset
          Dim Name As String
          Dim sql As String
          Dim tb As Recordset
          Dim Px As Printer
          Dim OriginalPrinter As String
          Dim lines As Integer

10        OriginalPrinter = Printer.DeviceName
20        If Not SetFormPrinter() Then Exit Sub

30        PrintHeadingCavan SampleID
40        With frmRTB

50            PrintTextRTB .rtb, FormatString(" Sample Type: EDTA", 18, , Alignleft), 12, False, , , vbRed
60            PrintTextRTB .rtb, FormatString(" Specimen Taken: " & Format(SampleDate, "dd/mm/yy hh:mm"), 31, , Alignleft), 12, False, , , vbRed
70            PrintTextRTB .rtb, FormatString(" Rec'd Date: " & Format(CurrentReceivedDate, "dd/mm/yy hh:mm"), 30, , Alignleft) & vbCrLf & vbCrLf, 12, False, , , vbRed
80            PrintTextRTB .rtb, FormatString(" Blood Group: ", 15, , Alignleft), 12, False, , , vbBlack

90            If InStr(fPrintForm.tgroup, "P") Or InStr(fPrintForm.tgroup, "N") Then
100               PrintTextRTB .rtb, FormatString(Trim$(Left$(fPrintForm.tgroup, 2)) & " Rh", 4, , Alignleft), 20, True, , , vbBlack
110               PrintTextRTB .rtb, FormatString("(D)", 3, , Alignleft), 10, True, , , vbBlack
120               PrintTextRTB .rtb, FormatString(IIf(InStr(fPrintForm.tgroup, "P"), "Positive", "Negative") & "  ", 10, , Alignleft), 20, True, , , vbBlack
130           Else
140               PrintTextRTB .rtb, FormatString("", 15, , Alignleft), 20, True, , , vbBlack
150           End If
170           PrintTextRTB .rtb, FormatString(" Antibody Screen: " & fPrintForm.lAB, 30, , Alignleft) & vbCrLf, 14, True, , , vbBlack
180           If frmxmatch.tSampleComment <> "" Then
190               PrintTextRTB .rtb, FormatString("Comment:" & frmxmatch.tSampleComment, 20, , Alignleft) & vbCrLf, 14, True, , , vbBlack
200           End If
210           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
220           PrintTextRTB .rtb, FormatString("Ante Natal Report", 70, , AlignCenter), 14, True, , , vbBlack
230           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
240           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
250           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
260           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
270           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
280           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
290           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
300           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
310           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
320           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
330           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
340           PrintFooterGHCavan SampleID
350           .rtb.SelPrint Printer.hDC

              ''''''''''
360           sql = "Select * from PatientDetails where " & _
                    "LabNumber = '" & SampleID & "'"
370           Set TempTb = New Recordset
380           RecOpenServerBB 0, TempTb, sql
390           If TempTb.EOF Then Exit Sub
400           Name = TempTb!Name
              ''''''''''
410           sql = "SELECT * FROM Reports WHERE 0 = 1"
420           Set tb = New Recordset
430           RecOpenServerBB 0, tb, sql
440           tb.AddNew
450           tb!SampleID = SampleID
460           tb!Name = Name
470           tb!Dept = "AN Report"
480           tb!Initiator = UserName
490           tb!PrintTime = Now    'PrintTime
500           tb!RepNo = "AN" & SampleID & Format(Now, "ddMMyyyyhhmmss")
510           tb!pagenumber = 1
520           tb!Report = .rtb.TextRTF
530           tb!Printer = Printer.DeviceName
540           tb.Update

550       End With
560       For Each Px In Printers
570           If Px.DeviceName = OriginalPrinter Then
580               Set Printer = Px
590               Exit For
600           End If
610       Next

End Sub
Public Sub PrintANFormCavanWithPreview(ByVal SampleID As String, _
                                       ByVal blnpreview As Boolean)

      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim f As Form
      Dim s As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Set f = New frmPreviewRTF

40    PrintHeadingCavanWithPreview SampleID, f
50    f.Dept = "AN"
60    f.SampleID = SampleID

70    f.WriteFormattedText "  Spec Grp: ;", , 12, vbBlack, , "Courier New"
80    If InStr(fPrintForm.tgroup, "P") Or InStr(fPrintForm.tgroup, "N") Then
90        f.WriteFormattedText Trim$(Left$(fPrintForm.tgroup, 2)) & " Rh;", True, 20, vbBlack, , "Courier New"
100       f.WriteFormattedText "(D) ;", , 10, vbBlack, , "Courier New"
110       If InStr(fPrintForm.tgroup, "P") Then
120           s = "Positive"
130       ElseIf InStr(fPrintForm.tgroup, "N") Then
140           s = "Negative"
150       End If
160   Else
170       s = " "
180   End If
190   f.WriteFormattedText s, , 20, vbBlack, , "Courier New"

200   f.WriteFormattedText fPrintForm.lAB, , 14, vbBlack, , "Courier New"

210   If frmxmatch.tSampleComment <> "" Then
220       f.WriteFormattedText "Comment:" & frmxmatch.tSampleComment, , 14, vbBlack, , "Courier New"
230   End If

240   f.WriteText vbCrLf
250   f.WriteFormattedText Space$(25) & "Ante Natal Report", , 14, vbBlack, , "Courier New"
260   f.WriteText vbCrLf

270   PrintFooterGHCavanWithPreview f

280   If blnpreview Then
290       f.Show 1
300   Else
310       f.PrintRTB
320   End If

330   Set f = Nothing

340   For Each Px In Printers
350       If Px.DeviceName = OriginalPrinter Then
360           Set Printer = Px
370           Exit For
380       End If
390   Next

End Sub
Public Sub PrintFooterGHCavan(ByVal SampleID As String)

      Dim sn As Recordset
      Dim sql As String
      Dim strAccreditation As String

10    On Error GoTo PrintFooterGHCavan_Error
20    With frmRTB
30        .Font.Name = "Courier New"

          '30    Do While Printer.CurrentY < 6700
          '40      .SelText = ""
          '50    Loop

40        strAccreditation = GetOptionSetting("TransfusionAccreditation", "Blood Transfusion at CGH is accredited by INAB to ISO 15189, detailed in scope Registration Number 231MT") & vbCrLf

50        PrintTextRTB .rtb, FormatString(String(248, "-"), 248, , Alignleft) & vbCrLf, 4, , , , vbBlack

60        sql = "select * from patientdetails where " & _
                "labnumber = '" & SampleID & "'"
70        Set sn = New Recordset
80        RecOpenServerBB 0, sn, sql


90        PrintTextRTB .rtb, FormatString("   " & strAccreditation, 120, , Alignleft) & vbCrLf, 6, , , , vbRed
100       PrintTextRTB .rtb, FormatString(" Report Date:" & Format(Now, "dd/mm/yyyy hh:mm"), 40, , Alignleft), 10, , , , vbRed
110       PrintTextRTB .rtb, FormatString("Issued By " & TechnicianNameForCode(sn!Operator & ""), 30, , AlignRight), 10, , , , vbRed

120   End With
130   Exit Sub

PrintFooterGHCavan_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "modSpecificCavan", "PrintFooterGHCavan", intEL, strES, sql

End Sub

Public Sub PrintFooterGHCavanWithPreview(ByRef f As Form)

      Dim sn As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo PrintFooterGHCavanWithPreview_Error

20    Do While f.LineCounter < 32
30        f.WriteText vbCrLf
40    Loop

50    f.WriteFormattedText String$(230, "-"), , 4, vbRed, , "Courier New"

60    s = Left$("Sample Date:" & Format(fPrintForm.tSampleDate, "dd/mm/yyyy") & Space$(38), 38) & ";"
70    f.WriteFormattedText s, , 10, vbBlack, , "Courier New"

80    s = "Report Date:" & Format(Now, "dd/mm/yyyy") & ";"
90    f.WriteFormattedText s, , 10, vbBlack, , "Courier New"

100   sql = "select * from patientdetails where " & _
            "labnumber = '" & fPrintForm.lLabNumber & "'"
110   Set sn = New Recordset
120   RecOpenServerBB 0, sn, sql

130   s = "    Issued By " & TechnicianNameForCode(sn!Operator & "")
140   f.WriteFormattedText s, , 10, vbBlack, , "Courier New"

150   Exit Sub

PrintFooterGHCavanWithPreview_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modSpecificCavan", "PrintFooterGHCavanWithPreview", intEL, strES, sql

End Sub

'Public Sub PrintGHFormCavan(ByVal SampleID As String, ByVal AvailableForDays As Integer, ByVal SampleDate As String)
'
'      Dim Px As Printer
'      Dim OriginalPrinter As String
'
'10    OriginalPrinter = Printer.DeviceName
'20    If Not SetFormPrinter() Then Exit Sub
'
'30    PrintHeadingCavan SampleID
'
'40    Printer.Font.Name = "Courier New"
'50    Printer.Font.Size = 12
'
'60    Printer.FontBold = False
'70    Printer.ForeColor = vbRed
'80    Printer.Print " Sample Type: EDTA";
'90    Printer.Print " Specimen Taken: "; Format(SampleDate, "dd/mm/yy hh:mm");
'100   Printer.Print " Rec'd Date: "; Format(CurrentReceivedDate, "dd/mm/yy hh:mm")
'110   Printer.ForeColor = vbBlack
'
'
'120   Printer.CurrentY = Printer.CurrentY + 100
'130   Printer.Print "  Spec Grp: ";
'140   Printer.CurrentY = Printer.CurrentY - 100
'150   Printer.Font.Bold = True
'160   Printer.Font.Size = 20
'170   If InStr(fPrintForm.tgroup, "P") Or InStr(fPrintForm.tgroup, "N") Then
'180     Printer.Print Trim$(Left$(fPrintForm.tgroup, 2)); " Rh";
'190     Printer.Font.Size = 10
'200     Printer.CurrentY = Printer.CurrentY + 150
'210     Printer.Print "(D) ";
'220     Printer.Font.Size = 20
'230     Printer.CurrentY = Printer.CurrentY - 150
'240     Printer.Print IIf(InStr(fPrintForm.tgroup, "P"), "Positive", "Negative"); "  ";
'250   Else
'260     Printer.Print
'270   End If
'280   Printer.Font.Size = 14
'290   Printer.Print
'300   Printer.Print
'310   Printer.Font.Bold = False
'320   Printer.Print "  Antibody Screen: ";
'330   Printer.Font.Bold = True
'340   Printer.Print fPrintForm.lAB
'350   If frmxmatch.tSampleComment <> "" Then
'360     Printer.Print "Comment:"; frmxmatch.tSampleComment
'370   End If
'
'380   Printer.Print
'390   Printer.Print
'400   Printer.Print
'410   Printer.Print Tab(25); "Group and Screen Only"
'420   Printer.Print
'430   Printer.Print Tab(13); "Serum Available until ";
'440   Printer.Print Format(DateAdd("d", AvailableForDays, fPrintForm.tSampleDate), "dd/mm/yyyy");
'450   Printer.Print " for Crossmatch."
'
'460   PrintFooterGHCavan SampleID
'
'470   Printer.EndDoc
'
'480   For Each Px In Printers
'490     If Px.DeviceName = OriginalPrinter Then
'500       Set Printer = Px
'510       Exit For
'520     End If
'530   Next
'
'End Sub
Public Sub PrintGHFormCavanWithPreview(ByVal SampleID As String, _
                                       ByVal AvailableForDays As Integer, _
                                       ByVal blnpreview As Boolean)

      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim f As Form
      Dim s As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Set f = New frmPreviewRTF

40    PrintHeadingCavanWithPreview SampleID, f
50    f.Dept = "GH"
60    f.SampleID = SampleID

70    f.WriteFormattedText "  Spec Grp: ;", False, 12, vbBlack, , "Courier New"
80    If InStr(fPrintForm.tgroup, "P") Or InStr(fPrintForm.tgroup, "N") Then
90        f.WriteFormattedText Trim$(Left$(fPrintForm.tgroup, 2)) & " Rh;", True, 20, vbBlack, , "Courier New"
100       f.WriteFormattedText "(D) ;", , 10, vbBlack, , "Courier New"
110       If InStr(fPrintForm.tgroup, "P") Then
120           s = "Positive"
130       ElseIf InStr(fPrintForm.tgroup, "N") Then
140           s = "Negative"
150       End If
160   Else
170       s = " "
180   End If
190   f.WriteFormattedText s, , 20, vbBlack, , "Courier New"

200   f.WriteFormattedText fPrintForm.lAB, False, 14, vbBlack, , "Courier New"
210   If frmxmatch.tSampleComment <> "" Then
220       f.WriteFormattedText "Comment:" & frmxmatch.tSampleComment, False, 14, vbBlack, , "Courier New"
230   End If

240   f.WriteText vbCrLf
250   f.WriteText vbCrLf
260   f.WriteText vbCrLf
270   f.WriteText vbCrLf
280   f.WriteText vbCrLf
290   f.WriteText vbCrLf
300   f.WriteFormattedText Space$(25) & "Group and Screen Only", , 14, vbBlack, , "Courier New"
310   f.WriteText vbCrLf

320   f.WriteFormattedText Space$(13) & "Serum Available until ;", , 14, vbBlack, , "Courier New"
330   f.WriteFormattedText Format(DateAdd("d", AvailableForDays, fPrintForm.tSampleDate), "dd/mm/yyyy") & ";", , 14, vbBlack, , "Courier New"
340   f.WriteFormattedText " for Crossmatch.", , 14, vbBlack, , "Courier New"
350   f.WriteText vbCrLf

360   PrintFooterGHCavanWithPreview f

370   If blnpreview Then
380       f.Show 1
390   Else
400       f.PrintRTB
410   End If

420   Set f = Nothing

430   For Each Px In Printers
440       If Px.DeviceName = OriginalPrinter Then
450           Set Printer = Px
460           Exit For
470       End If
480   Next

End Sub

Public Sub PrintGHFormCavan(ByVal SampleID As String, _
                            ByVal AvailableForDays As Integer, _
                            ByVal SampleDate As String)
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim tb As Recordset
      Dim TempTb As Recordset
      Dim sql As String
      Dim Name As String
10    With frmRTB
20        .rtb.Text = ""
          'frmRTB.Show 1
30        OriginalPrinter = Printer.DeviceName
40        If Not SetFormPrinter() Then Exit Sub

50        PrintHeadingCavan SampleID


60        PrintTextRTB .rtb, FormatString(" Sample Type: EDTA", 18, , Alignleft), 12, , , , vbRed
70        PrintTextRTB .rtb, FormatString(" Specimen Taken: " & Format(SampleDate, "dd/mm/yy hh:mm"), 31, , Alignleft), 12, , , , vbRed
80        PrintTextRTB .rtb, FormatString(" Rec'd Date: " & Format(CurrentReceivedDate, "dd/mm/yy hh:mm") & vbCrLf, 30, , Alignleft) & vbCrLf, 12, , , , vbRed
90        PrintTextRTB .rtb, FormatString("  Blood Group: ", 15, , Alignleft), 12, , , , vbBlack

100       If InStr(fPrintForm.tgroup, "P") Or InStr(fPrintForm.tgroup, "N") Then
110           PrintTextRTB .rtb, FormatString(Trim$(Left$(fPrintForm.tgroup, 2)) & " Rh", 4, , Alignleft), 20, True, , , vbBlack
120           PrintTextRTB .rtb, FormatString("(D) ", 3, , Alignleft), 10, , , , vbBlack
130           PrintTextRTB .rtb, FormatString(IIf(InStr(fPrintForm.tgroup, "P"), "Positive", "Negative"), 8, , Alignleft) & vbCrLf, 20, True, , , vbBlack

140       Else
150           PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 20, True, , , vbBlack

160       End If
170       .rtb.SelFontSize = 12
          'Printer.CurrentY = Printer.CurrentY + 100

180       If fPrintForm.lAB.Caption <> "" Then

190           PrintTextRTB .rtb, vbCrLf & FormatString("  Antibody Screen: ", 20, , Alignleft), 12, False, , , vbBlack
200       End If
210       PrintTextRTB .rtb, FormatString(fPrintForm.lAB, 30, , Alignleft) & vbCrLf, 12, True, , , vbBlack

220       If frmxmatch.tSampleComment <> "" Then
230           PrintTextRTB .rtb, FormatString("Comment:" & frmxmatch.tSampleComment, 50, , Alignleft), 12, , , , vbBlack
240       End If
250       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 12, , , , vbBlack
260       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 12, , , , vbBlack
270       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 12, , , , vbBlack
280       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 12, , , , vbBlack
290       PrintTextRTB .rtb, FormatString("Group and Screen Only", 80, , AlignCenter) & vbCrLf, 12, True, , , vbBlack
300       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 12, , , , vbBlack
310       PrintTextRTB .rtb, FormatString("Serum Available until " & Format(DateAdd("d", AvailableForDays, fPrintForm.tSampleDate), "dd/mm/yyyy") & " for Crossmatch.", 80, , AlignCenter) & vbCrLf & vbCrLf, 12, True, , , vbBlack
320       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 20, , , , vbBlack
330       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 20, , , , vbBlack
340       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 20, , , , vbBlack
350       PrintTextRTB .rtb, FormatString("", 30, , Alignleft) & vbCrLf, 20, , , , vbBlack
360       PrintFooterGHCavan SampleID
370       .rtb.SelPrint Printer.hDC
          ''''''''''
380       sql = "Select * from PatientDetails where " & _
                "LabNumber = '" & SampleID & "'"
390       Set TempTb = New Recordset
400       RecOpenServerBB 0, TempTb, sql
410       If TempTb.EOF Then Exit Sub
420       Name = TempTb!Name
          ''''''''''
430       sql = "SELECT * FROM Reports WHERE 0 = 1"
440       Set tb = New Recordset
450       RecOpenServerBB 0, tb, sql
460       tb.AddNew
470       tb!SampleID = SampleID
480       tb!Name = Name
490       tb!Dept = "GS Form"
500       tb!Initiator = UserName
510       tb!PrintTime = Now    'PrintTime
520       tb!RepNo = "GS" & SampleID & Format(Now, "ddMMyyyyhhmmss")
530       tb!pagenumber = 1
540       tb!Report = .rtb.TextRTF
550       tb!Printer = Printer.DeviceName
560       tb.Update


570   End With

      '500   Printer.EndDoc
      'frmRTB.Show 1
580   For Each Px In Printers
590       If Px.DeviceName = OriginalPrinter Then
600           Set Printer = Px
610           Exit For
620       End If
630   Next

End Sub


Public Sub PrintXMFormCavan(ByVal SampleID As String, _
                            ByRef RowNumbersToPrint() As Integer, _
                            ByRef SampleDate As String, _
                            ByVal HoldFor As Integer, _
                            ByVal Comment As String, _
                            Index As Integer)

          Dim y As Integer
          Dim Px As Printer
          Dim OriginalPrinter As String
          Dim sql As String
          Dim tb As Recordset
          Dim TempTb As Recordset
          Dim Name As String
          Dim fGroup As String
          Dim SampleComment As String
          Dim Antibodies As String
          Dim RecsEff As Long
          Dim R As Integer
          Dim Generic As String
          Dim TxAfter As String
          Dim lines As Integer

10        On Error GoTo PrintXMFormCavan_Error

20        OriginalPrinter = Printer.DeviceName
30        If Not SetFormPrinter() Then Exit Sub
40        With frmRTB
50            PrintHeadingCavan SampleID

60            sql = "Select * from PatientDetails where " & _
                    "LabNumber = '" & SampleID & "'"
70            Set tb = New Recordset
80            RecOpenClientBB 0, tb, sql
90            If Not tb.EOF Then
100               SampleComment = tb!SampleComment & ""
110               fGroup = tb!fGroup & ""
120           End If

130           PrintTextRTB .rtb, FormatString(" Sample Type: EDTA", 18, , Alignleft), 12, , , , vbRed
140           PrintTextRTB .rtb, FormatString(" Specimen Taken: " & Format(SampleDate, "dd/mm/yy hh:mm"), 31, , Alignleft), 12, , , , vbRed
150           PrintTextRTB .rtb, FormatString(" Rec'd Date: " & Format(CurrentReceivedDate, "dd/mm/yy hh:mm") & vbCrLf, 30, , Alignleft) & vbCrLf, 12, , , , vbRed
160           PrintTextRTB .rtb, FormatString("  Blood Group: ", 15, , Alignleft), 12, , , , vbBlack

170           If InStr(fPrintForm.tgroup, "P") Or InStr(fPrintForm.tgroup, "N") Then
180               PrintTextRTB .rtb, FormatString(Trim$(Left$(fPrintForm.tgroup, 2)) & " Rh", 4, , Alignleft), 20, True, , , vbBlack
190               PrintTextRTB .rtb, FormatString("(D) ", 3, , Alignleft), 10, , , , vbBlack
200               PrintTextRTB .rtb, FormatString(IIf(InStr(fPrintForm.tgroup, "P"), "Positive", "Negative"), 10, , Alignleft), 20, True, , , vbBlack

210           Else
220               PrintTextRTB .rtb, FormatString("", 30, , Alignleft), 20, True, , , vbBlack

230           End If

240           Antibodies = ""
250           If Trim$(tb!Anti3Reported & "") <> "" Then Antibodies = "Antibodies: " & tb!Anti3Reported
260           If Trim$(tb!AIDS & "") <> "" Then Antibodies = "Antibody Screen: " & tb!AIDS
270           If Trim$(tb!AIDR & "") <> "" Then Antibodies = "Antibody Screen: " & tb!AIDR
280           If Antibodies = "Antibody Screen: Negative" Then
290               Antibodies = "No Atypical Antibodies detected."
300           End If
310           PrintTextRTB .rtb, FormatString(Antibodies, 40, , Alignleft) & vbCrLf, 14, True, , , vbBlack

320           If Trim$(tb!SampleComment & "") <> "" Then
330               PrintTextRTB .rtb, FormatString("Comment:" & tb!SampleComment, 40, , Alignleft) & vbCrLf, 20, True, , , vbBlack
340           Else
350               PrintTextRTB .rtb, FormatString("", 40, , Alignleft) & vbCrLf, 20, True, , , vbBlack
360           End If


370           fPrintForm.g.row = RowNumbersToPrint(1)
380           fPrintForm.g.col = 3
390           PrintTextRTB .rtb, FormatString(fPrintForm.g, 40, , Alignleft) & vbCrLf, 12, True, , , vbBlack
400           PrintTextRTB .rtb, FormatString(Comment, 40, , Alignleft) & vbCrLf, 12, False, , , vbBlack

410           TxAfter = fPrintForm.g.TextMatrix(fPrintForm.g.row, 6)
420           If TxAfter = "" Then
430               Generic = UCase$(ProductGenericFor(ProductBarCodeFor(fPrintForm.g.TextMatrix(RowNumbersToPrint(1), 3))))
440               If UCase(fPrintForm.g) = "OCTAPLAS" Or UCase(fPrintForm.g) = "UNIPLAS" Or InStr(fPrintForm.g, "PLASMA") Or Generic = "LG OCTAPLAS" Then
450                   TxAfter = DateAdd("H", 4, fPrintForm.tSampleDate)
460               ElseIf Generic = "RED CELLS" Or Generic = "WHOLE BLOOD" Then
470                   TxAfter = DateAdd("H", HoldFor, fPrintForm.tSampleDate)
480               Else
490                   TxAfter = "Expiry Date"
500               End If
510               If IsDate(TxAfter) Then
520                   TxAfter = Format$(MinDate(TxAfter, fPrintForm.g.TextMatrix(fPrintForm.g.row, 2)), "dd/MM/yyyy HH:nn")
530               End If
540           End If

550           PrintTextRTB .rtb, FormatString("Unit No.         Group  Expiry        Date/Time                             ", 80, , Alignleft) & vbCrLf, 12, True, , , vbBlack
560           For y = 1 To UBound(RowNumbersToPrint())
570               R = RowNumbersToPrint(y)
580               If R <> 0 Then
590                   sql = "Update Latest " & _
                            "Set Operator = '" & UserCode & "' where " & _
                            "ISBT128 = '" & fPrintForm.g.TextMatrix(R, 0) & "' " & _
                            "and BarCode = '04333' " & _
                            "and Operator = 'Auto' "
600                   CnxnBB(0).Execute sql, RecsEff
610                   If RecsEff <> 0 Then
620                       sql = "Insert into Product " & _
                              " Select * from Latest where " & _
                              " ISBT128 = '" & fPrintForm.g.TextMatrix(R, 0) & "' " & _
                              " and BarCode = '04333' "
630                       CnxnBB(0).Execute sql
640                       fPrintForm.g.TextMatrix(R, 4) = UserCode
650                       fPrintForm.g.Refresh
660                   End If
670                   fPrintForm.g.row = R
680                   fPrintForm.g.col = 0
690                   PrintTextRTB .rtb, FormatString(fPrintForm.g.Text, 17, , Alignleft), 12, , , , vbBlack
700                   fPrintForm.g.col = 1
710                   PrintTextRTB .rtb, FormatString(fPrintForm.g.Text, 7, , Alignleft), 12, , , , vbBlack
720                   fPrintForm.g.col = 2
730                   PrintTextRTB .rtb, FormatString(Format(fPrintForm.g.Text, "dd/mm/yy hh:mm"), 30, , Alignleft) & vbCrLf, 12, , , , vbBlack
740                   lines = lines + 1
750               End If
760           Next

770           PrintTextRTB .rtb, FormatString("Do not commence Transfusion after " & TxAfter, 80, , Alignleft) & vbCrLf, 12, True, , , vbBlack
772           If Index = 1 Then
775             PrintTextRTB .rtb, FormatString("These units have been electronically issued.", 80, , Alignleft) & vbCrLf, 12, True, , , vbBlack
777           End If
              
780           For y = 1 To (10 - lines)
790               PrintTextRTB .rtb, FormatString("", 80, , Alignleft) & vbCrLf, 12, True, , , vbBlack
800           Next

810           PrintFooterCavan HoldFor
820           .rtb.SelPrint Printer.hDC

              ''''''''''
830           sql = "Select * from PatientDetails where " & _
                    "LabNumber = '" & SampleID & "'"
840           Set TempTb = New Recordset
850           RecOpenServerBB 0, TempTb, sql
860           If TempTb.EOF Then Exit Sub
870           Name = TempTb!Name
              ''''''''''
880           sql = "SELECT * FROM Reports WHERE 0 = 1"
890           Set tb = New Recordset
900           RecOpenServerBB 0, tb, sql
910           tb.AddNew
920           tb!SampleID = SampleID
930           tb!Name = Name
940           tb!Dept = "XM Form"
950           tb!Initiator = UserName
960           tb!PrintTime = Now    'PrintTime
970           tb!RepNo = "XM" & SampleID & Format(Now, "ddMMyyyyhhmmss")
980           tb!pagenumber = 1
990           tb!Report = .rtb.TextRTF
1000          tb!Printer = Printer.DeviceName
1010          tb.Update

1020      End With

1030      For Each Px In Printers
1040          If Px.DeviceName = OriginalPrinter Then
1050              Set Printer = Px
1060              Exit For
1070          End If
1080      Next

1090      Exit Sub


PrintXMFormCavan_Error:

          Dim strES As String
          Dim intEL As Integer

1100      intEL = Erl
1110      strES = Err.Description
1120      LogError "modSpecificCavan", "PrintXMFormCavan", intEL, strES, sql

End Sub

Public Sub PrintXMFormCavanWithPreview(ByVal SampleID As String, _
                                       ByRef RowNumbersToPrint() As Integer, _
                                       ByRef SampleDate As String, _
                                       ByVal HoldFor As Integer, _
                                       ByVal blnpreview As Boolean)

      Dim y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim sql As String
      Dim tb As Recordset
      Dim fGroup As String
      Dim SampleComment As String
      Dim Antibodies As String
      Dim RecsEff As Long
      Dim f As Form
      Dim s As String

10    On Error GoTo PrintXMFormCavanWithPreview_Error

20    OriginalPrinter = Printer.DeviceName
30    If Not SetFormPrinter() Then Exit Sub

40    Set f = New frmPreviewRTF

50    PrintHeadingCavanWithPreview SampleID, f
60    f.Dept = "XM"
70    f.SampleID = SampleID

80    sql = "Select * from PatientDetails where " & _
            "LabNumber = '" & SampleID & "'"
90    Set tb = New Recordset
100   RecOpenClientBB 0, tb, sql
110   If Not tb.EOF Then
120       SampleComment = tb!SampleComment & ""
130       fGroup = tb!fGroup & ""
140   End If

150   f.WriteFormattedText "  Spec Grp: ;", False, 12, vbBlack, , "Courier New"

160   If InStr(fGroup, "P") Or InStr(fGroup, "N") Then
170       f.WriteFormattedText Trim$(Left$(fGroup, 2)) & " Rh;", True, 20, vbBlack, , "Courier New"
180       f.WriteFormattedText "(D) ;", False, 10, vbBlack, , "Courier New"
190       If InStr(fGroup, "P") Then
200           s = "Positive"
210       ElseIf InStr(fGroup, "N") Then
220           s = "Negative"
230       End If
240   Else
250       s = " "
260   End If
270   f.WriteFormattedText s, True, 20, vbBlack, , "Courier New"

280   Antibodies = ""
290   If Trim$(tb!Anti3Reported & "") <> "" Then Antibodies = "Antibodies: " & tb!Anti3Reported
300   If Trim$(tb!AIDS & "") <> "" Then Antibodies = "Antibodies: " & tb!AIDS
310   If Trim$(tb!AIDR & "") <> "" Then Antibodies = "Antibodies: " & tb!AIDR
320   If Antibodies = "Antibodies: Negative" Then
330       Antibodies = "No Atypical Antibodies detected."
340   End If
350   f.WriteFormattedText Antibodies, False, 14, vbBlack, , "Courier New"

360   If Trim$(tb!SampleComment & "") <> "" Then
370       f.WriteFormattedText "Comment:" & tb!SampleComment, False, 14, vbBlack, , "Courier New"
380   End If

390   With fPrintForm
400       f.WriteFormattedText .g.TextMatrix(RowNumbersToPrint(1), 3) & " Issued", , 12, vbBlack, , "Courier New"
410       f.WriteFormattedText "Unit No.  Group    Expiry         Date/Time                              ", , 12, vbBlack, , "Courier New"

420       For y = 1 To UBound(RowNumbersToPrint())
430           sql = "Update Latest " & _
                    "Set Operator = '" & UserCode & "' where " & _
                    "Number = '" & .g.TextMatrix(RowNumbersToPrint(y), 0) & "' " & _
                    "and BarCode = '04333' " & _
                    "and Operator = 'Auto' "
440           CnxnBB(0).Execute sql, RecsEff
450           If RecsEff <> 0 Then
460               sql = "Insert into Product " & _
                        " Select * from Latest where " & _
                        " Number = '" & .g.TextMatrix(RowNumbersToPrint(y), 0) & "' " & _
                        " and BarCode = '04333' "
470               CnxnBB(0).Execute sql
480               .g.TextMatrix(RowNumbersToPrint(y), 4) = UserCode
490               .g.Refresh
500           End If
510           f.WriteText vbCrLf
520           s = Left$(.g.TextMatrix(RowNumbersToPrint(y), 0) & Space$(10), 10)
530           s = s & Left$(.g.TextMatrix(RowNumbersToPrint(y), 1) & Space$(9), 9)
540           s = s & .g.TextMatrix(RowNumbersToPrint(y), 2)
550           f.WriteFormattedText s, , 12, vbBlack, , "Courier New"
560       Next
570   End With

580   PrintFooterXMCavanWithPreview SampleDate, HoldFor, f

590   If blnpreview Then
600       f.Show 1
610   Else
620       f.PrintRTB
630   End If

640   Set f = Nothing

650   For Each Px In Printers
660       If Px.DeviceName = OriginalPrinter Then
670           Set Printer = Px
680           Exit For
690       End If
700   Next

710   Exit Sub

PrintXMFormCavanWithPreview_Error:

      Dim strES As String
      Dim intEL As Integer

720   intEL = Erl
730   strES = Err.Description
740   LogError "modSpecificCavan", "PrintXMFormCavanWithPreview", intEL, strES, sql


End Sub

Public Sub PrintFooterCavan(ByVal HoldFor As Integer)
          Dim strAccreditation As String
10        With frmRTB
20            strAccreditation = GetOptionSetting("TransfusionAccreditation", "Blood Transfusion at CGH is accredited by INAB to ISO 15189, detailed in scope Registration Number 231MT")

30            PrintTextRTB .rtb, FormatString("Please check Patient Identity Prior to Transfusion of each Unit.", 80, , Alignleft) & vbCrLf, 10, , , , vbGreen
40            PrintTextRTB .rtb, FormatString("Crossmatched Blood held for " & Format$(HoldFor) & " Hours only. All Units to be Signed out in BT Register.", 80, , Alignleft) & vbCrLf, 10, , , , vbGreen
50            PrintTextRTB .rtb, FormatString(String(248, "-"), 248, , Alignleft) & vbCrLf, 4, , , , vbRed
60            PrintTextRTB .rtb, FormatString("   " & strAccreditation, 120, , Alignleft) & vbCrLf, 6, , , , vbRed
70            PrintTextRTB .rtb, FormatString(" Report Date:" & Format(Now, "dd/mm/yyyy hh:mm"), 45, , Alignleft), 9, , , , vbRed
80            PrintTextRTB .rtb, FormatString("Issued By " & UserName, 30, , AlignRight), 9, , , , vbRed

90        End With
End Sub

Public Sub PrintFooterCavanConfirmed(ByVal SampleDate As String)

10    Printer.Font.Name = "Courier New"
20    Printer.Font.Size = 10
30    Printer.ForeColor = vbGreen

40    Do While Printer.CurrentY < 6700
50        Printer.Print
60    Loop

70    Printer.ForeColor = vbRed
80    Printer.Font.Size = 4
90    Printer.Print String$(250, "-")

100   Printer.Font.Size = 10
110   Printer.Font.Bold = False

120   Printer.Print "Sample Date:"; Format(SampleDate, "dd/mm/yyyy");

130   Printer.Print Tab(38); "Report Date:"; Format(Now, "dd/mm/yyyy");

140   Printer.Print "    Issued By "; UserName

End Sub


Public Sub PrintFooterXMCavanWithPreview(ByVal SampleDate As String, _
                                         ByVal HoldFor As Integer, _
                                         ByRef f As Form)

      Dim s As String
      Dim strAccreditation As String

10    strAccreditation = GetOptionSetting("TransfusionAccreditation", "Blood Transfusion at CGH is accredited by INAB to ISO 15189, detailed in scope Registration Number 231MT")

20    Do While f.LineCounter < 29
30        f.WriteText vbCrLf
40    Loop

50    f.WriteFormattedText "Please check Patient Identity Prior to Transfusion of each Unit.", , 10, vbGreen, , "Courier New"
60    f.WriteFormattedText "Crossmatched Blood held for ;", , 10, vbGreen, , "Courier New"
70    f.WriteFormattedText Format$(HoldFor) & ";", , 10, vbGreen, , "Courier New"
80    f.WriteFormattedText " Hours only. All Units to be Signed out in BT Register.", , 10, vbGreen, , "Courier New"

90    f.WriteFormattedText String(230, "-"), , 4, vbRed, , "Courier New"
100   f.WriteFormattedText strAccreditation, , 6, vbRed, , "Courier New"

110   Printer.Font.Size = 10

120   s = Left$("Sample Date:" & Format(SampleDate, "dd/mm/yyyy") & Space$(38), 38) & ";"
130   f.WriteFormattedText s, , 10, vbBlack, , "Courier New"

140   s = "Report Date:" & Format(Now, "dd/mm/yyyy") & ";"
150   f.WriteFormattedText s, , 10, vbBlack, , "Courier New"

160   f.WriteFormattedText "    Issued By " & UserName, , 10, vbBlack, , "Courier New"

End Sub

Public Sub PrintCordFormCavan(ByVal SampleID As String, _
                              ByVal SampleDate As String)

      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim prnted As Boolean
      Dim TempTb As Recordset
      Dim tb As Recordset
      Dim sql As String
      Dim Name As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub
30    With frmRTB
          'For n = 0 To 1
40        PrintHeadingCavan SampleID
50        .Font = " Courier New"

60        PrintTextRTB .rtb, FormatString(" Sample Type: EDTA", 18, , Alignleft), 12, False, , , vbRed
70        PrintTextRTB .rtb, FormatString(" Specimen Taken: " & Format(SampleDate, "dd/mm/yy hh:mm"), 31, , Alignleft), 12, False, , , vbRed
80        PrintTextRTB .rtb, FormatString(" Rec'd Date: " & Format(CurrentReceivedDate, "dd/mm/yy hh:mm"), 30, , Alignleft) & vbCrLf & vbCrLf, 12, False, , , vbRed
90        Printer.ForeColor = vbBlack
100           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
110           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
120           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
130       PrintTextRTB .rtb, FormatString(" Blood Group:", 15, , Alignleft), 12, False, , , vbBlack

140       If InStr(fPrintForm.tgroup, "P") Or InStr(fPrintForm.tgroup, "N") Then
150           PrintTextRTB .rtb, FormatString(Trim$(Left$(fPrintForm.tgroup, 2)) & " Rh", 4, , Alignleft), 20, True, , , vbBlack
160           PrintTextRTB .rtb, FormatString("(D)", 3, , Alignleft), 10, True, , , vbBlack
170           PrintTextRTB .rtb, FormatString(IIf(InStr(fPrintForm.tgroup, "P"), "Positive", "Negative") & "  ", 10, , Alignleft), 20, True, , , vbBlack

180       Else
190           PrintTextRTB .rtb, FormatString("", 15, , Alignleft), 20, True, , , vbBlack
200       End If

220       prnted = False
230       frmxmatch.gDAT.col = 1
240       frmxmatch.gDAT.row = 1
250       If frmxmatch.gDAT.CellPicture = frmxmatch.imgSquareCross.Picture Then
255           PrintTextRTB .rtb, FormatString("Direct Coombs ", 15, , Alignleft), 14, True, , , vbBlack
260           PrintTextRTB .rtb, FormatString("Positive", 15, , Alignleft), 14, True, , , vbBlack
270           prnted = True
280       Else
290           frmxmatch.gDAT.col = 2
300           If frmxmatch.gDAT.CellPicture = frmxmatch.imgSquareCross.Picture Then
305               PrintTextRTB .rtb, FormatString("Direct Coombs ", 15, , Alignleft), 14, True, , , vbBlack
310               PrintTextRTB .rtb, FormatString("Negative", 15, , Alignleft), 14, True, , , vbBlack
320               prnted = True
330           End If
340       End If
'350       If Not prnted Then
'360           PrintTextRTB .rtb, FormatString("########", 15, , Alignleft), 14, True, , , vbBlack
'370       End If

380       PrintTextRTB .rtb, FormatString("", 15, , Alignleft) & vbCrLf, 14, False, , , vbBlack
390       PrintTextRTB .rtb, FormatString("", 15, , Alignleft) & vbCrLf, 14, False, , , vbBlack
400       If Trim$(frmxmatch.tSampleComment) <> "" Then
410           PrintTextRTB .rtb, FormatString("Comment: " & frmxmatch.tSampleComment, 50, , Alignleft), 14, False, , , vbBlack
420       End If
430           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
440           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
450           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
460           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
470           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
480           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
490           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
500           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
510           PrintTextRTB .rtb, FormatString("", 80, , AlignCenter) & vbCrLf, 14, True, , , vbBlack
520       PrintFooterGHCavan SampleID

530       .rtb.SelPrint Printer.hDC

          ''''''''''
540       sql = "Select * from PatientDetails where " & _
          "LabNumber = '" & SampleID & "'"
550       Set TempTb = New Recordset
560       RecOpenServerBB 0, TempTb, sql
570       If TempTb.EOF Then Exit Sub
580       Name = TempTb!Name
          ''''''''''
590    sql = "SELECT * FROM Reports WHERE 0 = 1"
600       Set tb = New Recordset
610       RecOpenServerBB 0, tb, sql
620       tb.AddNew
630       tb!SampleID = SampleID
640       tb!Name = Name
650       tb!Dept = "Cord Report"
660       tb!Initiator = UserName
670       tb!PrintTime = Now 'PrintTime
680       tb!RepNo = "Cord" & SampleID & Format(Now, "ddMMyyyyhhmmss")
690       tb!pagenumber = 1
700       tb!Report = .rtb.TextRTF
710       tb!Printer = Printer.DeviceName
720       tb.Update

730   End With
740   For Each Px In Printers
750       If Px.DeviceName = OriginalPrinter Then
760           Set Printer = Px
770           Exit For
780       End If
790   Next

End Sub

Public Sub PrintCordFormCavanWithPreview(ByVal SampleID As String, _
                                         ByVal blnpreview As Boolean)

      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim prnted As Boolean
      Dim f As Form
      Dim s As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Set f = New frmPreviewRTF

40    PrintHeadingCavanWithPreview SampleID, f
50    f.Dept = "CD"
60    f.SampleID = SampleID

70    f.WriteText vbCrLf
80    f.WriteText vbCrLf
90    f.WriteText vbCrLf
100   f.WriteText vbCrLf
110   f.WriteText vbCrLf
120   f.WriteText vbCrLf

130   f.WriteFormattedText "  Spec Grp: ;", , 12, vbBlack, , "Courier New"
140   If InStr(fPrintForm.tgroup, "P") Or InStr(fPrintForm.tgroup, "N") Then
150       f.WriteFormattedText Trim$(Left$(fPrintForm.tgroup, 2)) & " Rh;", True, 20, vbBlack, , "Courier New"
160       f.WriteFormattedText "(D) ;", , 10, vbBlack, , "Courier New"
170       If InStr(fPrintForm.tgroup, "P") Then
180           s = "Positive"
190       ElseIf InStr(fPrintForm.tgroup, "N") Then
200           s = "Negative"
210       End If
220   Else
230       s = " "
240   End If
250   f.WriteFormattedText s, , 20, vbBlack, , "Courier New"

260   f.WriteFormattedText "Direct Coombs ;", False, 14, vbBlack, , "Courier New"

270   prnted = False
280   frmxmatch.gDAT.col = 1
290   frmxmatch.gDAT.row = 1
300   If frmxmatch.gDAT.CellPicture = frmxmatch.imgSquareCross.Picture Then
310       f.WriteFormattedText "Positive", False, 14, vbBlack, , "Courier New"
320       prnted = True
330   Else
340       frmxmatch.gDAT.col = 2
350       If frmxmatch.gDAT.CellPicture = frmxmatch.imgSquareCross.Picture Then
360           f.WriteFormattedText "Negative", False, 14, vbBlack, , "Courier New"
370           prnted = True
380       End If
390   End If
400   If Not prnted Then
410       f.WriteFormattedText "########", False, 14, vbBlack, , "Courier New"
420   End If
430   f.WriteText vbCrLf
440   f.WriteText vbCrLf
450   If Trim$(frmxmatch.tSampleComment) <> "" Then
460       f.WriteFormattedText "Comment: " & frmxmatch.tSampleComment, False, 14, vbBlack, , "Courier New"
470   End If

480   PrintFooterGHCavanWithPreview f

490   If blnpreview Then
500       f.Show 1
510   Else
520       f.PrintRTB
530   End If

540   Set f = Nothing
550   For Each Px In Printers
560       If Px.DeviceName = OriginalPrinter Then
570           Set Printer = Px
580           Exit For
590       End If
600   Next

End Sub
Public Sub PrintHeadingCavan(ByVal SampleID As String, Optional Heading As String)

      Dim tb As Recordset
      Dim sql As String
      Dim TempStr As String
10    On Error GoTo PrintHeadingCavan_Error
20    With frmRTB
30    .rtb.Text = ""
40        sql = "Select * from PatientDetails where " & _
                "LabNumber = '" & SampleID & "'"
50        Set tb = New Recordset
60        RecOpenServerBB 0, tb, sql
70        If tb.EOF Then Exit Sub

80        If Heading = "" Then Heading = "Blood Transfusion Laboratory"

90        .rtb.SelFontName = "Courier New"
100       PrintTextRTB .rtb, FormatString("CAVAN GENERAL HOSPITAL : " & Heading, 70, , Alignleft) & vbCrLf, 14, True, , , vbRed


110       PrintTextRTB .rtb, FormatString(String$(248, "-"), 248, , AlignCenter) & vbCrLf, 4, , , , vbBlue

120       PrintTextRTB .rtb, FormatString(" Sample ID:" & SampleID, 30, , Alignleft), 12, , , , vbBlack
130       PrintTextRTB .rtb, FormatString(" Name:", 7, , Alignleft), 12, , False, , vbBlack
140       PrintTextRTB .rtb, FormatString(tb!Name, 25, , Alignleft) & vbCrLf, 14, True, , , vbBlack
150       PrintTextRTB .rtb, FormatString("      Ward:" & tb!Ward, 31, , Alignleft), 12, , False, , vbBlack
160       PrintTextRTB .rtb, FormatString(" DOB:" & Format(tb!DoB, "dd/mm/yyyy"), 30, , Alignleft), 12, , False, , vbBlack
170       PrintTextRTB .rtb, FormatString(" Chart #:" & tb!patnum, 30, , Alignleft) & vbCrLf, 12, , False, , vbBlack




180       If Trim$(tb!Clinician & "") <> "" Then
190           PrintTextRTB .rtb, FormatString(" Consultant:" & tb!Clinician, 30, , Alignleft), 12, , False, , vbBlack
200       Else
210           PrintTextRTB .rtb, FormatString("        GP:" & tb!GP, 30, , Alignleft), 12, , False, , vbBlack
220       End If
230       PrintTextRTB .rtb, FormatString(" Addr:" & tb!Addr1, 35, , Alignleft), 12, , False, , vbBlack


240       Select Case Left$(UCase$(Trim$(tb!Sex & "")), 1)
          Case "M": TempStr = "Male"
250       Case "F": TempStr = "Female"
260       Case Else: TempStr = ""
270       End Select
280       PrintTextRTB .rtb, FormatString(" Sex:" & TempStr, 30, , Alignleft) & vbCrLf, 12, , False, , vbBlack
290       PrintTextRTB .rtb, FormatString(".", 80, , AlignCenter) & vbCrLf, 12, , False, , vbBlack
300       PrintTextRTB .rtb, FormatString(String$(248, "-"), 248, , AlignCenter) & vbCrLf, 4, , False, , vbBlack

310   End With
320   Exit Sub

PrintHeadingCavan_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "modSpecificCavan", "PrintHeadingCavan", intEL, strES, sql

End Sub

Public Sub PrintHeadingCavanWithPreview(ByVal SampleID As String, _
                                        ByRef f As frmPreviewRTF)

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo PrintHeadingCavanWithPreview_Error

20    sql = "Select * from PatientDetails where " & _
            "LabNumber = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If tb.EOF Then Exit Sub

60    With f
70        .AdjustPaperSize "A5Land"
80        .Clear
90        .WriteFormattedText "CAVAN GENERAL HOSPITAL : Blood Transfusion Laboratory", True, 14, vbRed, , "Courier New"
100       .WriteFormattedText String(230, "-"), False, 4, vbRed, , "Courier New"

110       s = Left$(" Sample ID:" & SampleID & Space$(35), 35)
120       .WriteFormattedText s & ";", False, 12, vbBlack, , "Courier New"
130       .WriteFormattedText "Name:;", False, 12, vbBlack, , "Courier New"
140       .WriteFormattedText tb!Name & "", True, 14, vbBlack, , "Courier New"

150       s = Left$("      Ward:" & tb!Ward & Space$(35), 35)
160       .WriteFormattedText s & ";", False, 12, vbBlack, , "Courier New"
170       s = Left$(" DoB:" & Format$(tb!DoB, "dd/mm/yyyy") & Space$(25), 25)
180       .WriteFormattedText s & ";", False, 12, vbBlack, , "Courier New"
190       .WriteFormattedText "Chart #:" & tb!patnum & "", False, 12, vbBlack, , "Courier New"


200       If Trim$(tb!Clinician & "") <> "" Then
210           s = Left$("Consultant:" & tb!Clinician & Space$(35), 35)
220       Else
230           s = Left$("        GP:" & tb!GP & Space$(35), 35)
240       End If
250       .WriteFormattedText s & ";", False, 12, vbBlack, , "Courier New"
260       s = Left$("Addr:" & tb!Addr1 & Space$(25), 25)
270       .WriteFormattedText s & ";", False, 12, vbBlack, , "Courier New"
280       s = "    Sex:"
290       Select Case Left$(UCase$(Trim$(tb!Sex & "")), 1)
          Case "M": s = s & "Male"
300       Case "F": s = s & "Female"
310       End Select
320       .WriteFormattedText s, False, 12, vbBlack, , "Courier New"

330       .WriteFormattedText Space$(40) & tb!Addr2 & "", False, 12, vbBlack, , "Courier New"

340       .WriteFormattedText String(230, "-"), , 4, vbBlack, , "Courier New"

350   End With

360   Exit Sub

PrintHeadingCavanWithPreview_Error:

      Dim strES As String
      Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "modSpecificCavan", "PrintHeadingCavanWithPreview", intEL, strES, sql


End Sub

Public Sub PrintTransfusedConfirmationCavan(ByRef RowsToPrint() As Integer, ByVal SampleDate As String)

      Dim y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim R As Integer

10    On Error GoTo PrintTransfusedConfirmationCavan_Error

20    OriginalPrinter = Printer.DeviceName
30    If Not SetFormPrinter() Then Exit Sub

40    PrintHeadingCavan fPrintForm.lLabNumber, "Transfusion Confirmation"
50    Printer.ForeColor = vbBlack
60    Printer.Font.Name = "Courier New"
70    Printer.Font.Size = 14
80    Printer.Font.Bold = True
90    Printer.Print "                                  Start      Finish       NAME     "
100   Printer.Print "    Unit No.       Transfused   Date/Time  Date/Time BLOCK CAPITALS"
110   For y = 1 To UBound(RowsToPrint())
120       R = RowsToPrint(y)
130       If R <> 0 Then
140           Printer.Print
150           Printer.Print "   "; Left$(fPrintForm.g.TextMatrix(R, 0) & Space$(17), 17);
160           Printer.Print "Yes";
170           Printer.Font.Name = "Wingdings"
180           Printer.Print "o";
190           Printer.Font.Name = "Courier New"
200           Printer.Print " No";
210           Printer.Font.Name = "Wingdings"
220           Printer.Print "o";
230           Printer.Font.Name = "Courier New"
240           Printer.Print "   .........  ......... .............."
250       End If
260   Next
270   Printer.ForeColor = vbRed
280   Printer.Print
290   Printer.Print Space$(10); "Please return this form to the Laboratory"
300   Printer.Print Space$(10); "on completion of this transfusion episode."
310   Printer.Print
320   Printer.Font.Size = 9
330   Printer.ForeColor = vbBlack
340   Printer.Print "    Was there a reaction to any of the above products?";
350   Printer.Print Space$(2); "Yes";
360   Printer.Font.Name = "Wingdings"
370   Printer.Print "o";
380   Printer.Font.Name = "Courier New"
390   Printer.Print " No";
400   Printer.Font.Name = "Wingdings"
410   Printer.Print "o";
420   Printer.Font.Name = "Courier New"
430   Printer.Print Space$(2); "Unit Number________Signature________";

440   Printer.Font.Size = 10


450   PrintFooterCavanConfirmed SampleDate

460   Printer.ForeColor = vbBlack

470   Printer.EndDoc
480   CurrentReceivedDate = ""
490   For Each Px In Printers
500       If Px.DeviceName = OriginalPrinter Then
510           Set Printer = Px
520           Exit For
530       End If
540   Next

550   Exit Sub

PrintTransfusedConfirmationCavan_Error:

      Dim strES As String
      Dim intEL As Integer

560   intEL = Erl
570   strES = Err.Description
580   LogError "modSpecificCavan", "PrintTransfusedConfirmationCavan", intEL, strES

End Sub
Public Function PrintTextRTB(rtb As RichTextBox, ByVal Text As String, _
                             Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                             Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                             Optional FontColor As ColorConstants = vbBlack, _
                             Optional SuperScript As Boolean = False)

      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------
10    On Error GoTo PrintTextRTB_Error

20    With rtb

30        .SelFontSize = FontSize
40        .SelBold = FontBold
50        .SelItalic = FontItalic
60        .SelUnderline = FontUnderLine
70        .SelColor = FontColor
80        If SuperScript Then
90            .SelCharOffset = 40
100       Else
110           .SelCharOffset = 0
120       End If
130       .SelText = Text
140   End With

150   Exit Function

PrintTextRTB_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "Other", "PrintTextRTB", intEL, strES

End Function



