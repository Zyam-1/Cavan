Attribute VB_Name = "modPrintForms"
Option Explicit

Public Sub PrintBatchForm(ByVal SampleID As String, _
                          ByVal TimeIssued As String)

      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim tb As Recordset
      Dim tP As Recordset
      Dim sql As String
      Dim n As Integer
10    ReDim UnitNumbers(0 To 0) As String
      Dim ub As Integer

20    On Error GoTo PrintBatchForm_Error

30    OriginalPrinter = Printer.DeviceName
40    If Not SetFormPrinter() Then Exit Sub

50    TimeIssued = Format(TimeIssued, "dd/mmm/yyyy hh:mm:ss")
60    sql = "Select * from BatchDetails where " & _
            "SampleID = '" & SampleID & "' " & _
            "and date = '" & TimeIssued & "'"

70    Set tb = New Recordset
80    RecOpenServerBB 0, tb, sql
90    If tb.EOF Then Exit Sub

100   sql = "Select * from BatchProductList where " & _
            "BatchNumber = '" & tb!BatchNumber & "' " & _
            "and Product = '" & tb!Product & "'"
110   Set tP = New Recordset
120   RecOpenServerBB 0, tP, sql
130   If tP.EOF Then Exit Sub

140     PrintHeadingCavan SampleID
        
150   Printer.Font.Name = "Courier New"
160   Printer.Font.Size = 12

170   Printer.Print
180   Printer.Font.Bold = False

190   Printer.Print "                                       Date/Time  Checked   Given   Date/Time"
200   Printer.Print "                                      Transfusion   By       By    Transfusion"
210   Printer.Print "                                         Start                       Complete"

220   Printer.Print "Product Issued: ";
230   Printer.Font.Bold = True
240   Printer.Print tb!Product & "";
250   If Trim$(tP!Group & "") <> "" Then
260     Printer.Font.Bold = False
270     Printer.Print " Group ";
280     Printer.Font.Bold = True
290     Printer.Print tP!Group;
300   End If
310   Printer.Print
320   Printer.Print

330   Printer.Font.Bold = False
340   Printer.Print "No. of Units Issued: ";
350   Printer.Font.Bold = True
360   Printer.Print tb!Bottles

370   ub = 0
380   For n = 1 To Val(tb!Bottles)
390     Printer.Print
400     Printer.Font.Bold = False
410     Printer.Print "Batch No.: ";
420     Printer.Font.Bold = True
430     Printer.Print tP!BatchNumber & "";
440     ub = ub + 1
450     ReDim Preserve UnitNumbers(0 To ub)
460     UnitNumbers(ub) = tP!BatchNumber & ""
470     Printer.Print "  Exp: "; tP!DateExpiry
480   Next

490   Printer.Print
500   If InStr(UCase$(tP!Product), "PLAS") <> 0 Then
510     If Val(tb!Bottles) > 1 Then
520       Printer.Print "These units";
530     Else
540       Printer.Print "This Unit";
550     End If
560     Printer.Print " must be used before "; Format(DateAdd("H", 4, TimeIssued), "dd/mm/yy hh:mm")
570   End If

580   PrintFooterCavan 48

590   Printer.EndDoc

600   CurrentReceivedDate = ""
610   For Each Px In Printers
620     If Px.DeviceName = OriginalPrinter Then
630       Set Printer = Px
640       Exit For
650     End If
660   Next

670   Exit Sub

PrintBatchForm_Error:

      Dim strES As String
      Dim intEL As Integer

680   intEL = Erl
690   strES = Err.Description
700   LogError "modPrintForms", "PrintBatchForm", intEL, strES, sql

End Sub







Public Sub PrintKleihauerFormCavan_OLD(ByVal SampleID As String)

      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    PrintHeadingCavan SampleID

40    Printer.Font.Name = "Courier New"
50    Printer.Font.Size = 12
60    Printer.Print
70    Printer.Print
80    Printer.Print
90    Printer.Print

100   Printer.Font.Size = 20
110   Printer.Print "   Patients Rhesus : Rh";
120   Printer.Font.Size = 10
130   Printer.CurrentY = Printer.CurrentY + 150
140   Printer.Print "(D) ";
150   Printer.Font.Size = 20
160   Printer.CurrentY = Printer.CurrentY - 150
170   Printer.Print frmKleihauer.lblRh
180   Printer.Print
190   Printer.Print "   " & frmKleihauer.lblFMHReport
200   Printer.Print
210   Printer.Print

220   Printer.Font.Size = 12

230   Printer.Print "A 1,500 IU dose of Rhophlyac is sufficient to cover a fetomaternal bleed 12 ml"
240   Printer.Print "within 6 weeks of administration. If a bleed >12 ml is suspected, further "
250   Printer.Print "Anti-D Ig may be required."
260   Printer.Print
270   Printer.Print "Any Queries on the Administration of Anti-D Ig Contact Consultant Haematologist"


280   PrintFooterGHCavan SampleID

290   Printer.EndDoc

300   For Each Px In Printers
310     If Px.DeviceName = OriginalPrinter Then
320       Set Printer = Px
330       Exit For
340     End If
350   Next

End Sub



Public Sub PrintKleihauerFormCavan(ByVal SampleID As String, ByVal intFetal As Integer, _
                                   ByVal strMessageType As String, ByVal strSampleDate As String, _
                                   ByVal strCurrentReceivedDate As String)

      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim sngFMH As Single
           Dim sql As String
           Dim TempTb As Recordset
           Dim tb As Recordset
           Dim strName As String
           
10    On Error GoTo PrintKleihauerFormCavan_Error

20    OriginalPrinter = Printer.DeviceName
30    If Not SetFormPrinter() Then Exit Sub

40         With frmRTB
50    PrintHeadingCavan SampleID

      '40    Printer.Font.Name = "Courier New"
      '50    Printer.Font.Size = 12

      '60    Printer.FontBold = False
      '70    Printer.ForeColor = vbRed
      '80    Printer.Print "  Sample Type: EDTA";
      '90    Printer.Print "  Specimen Taken: "; Format(strSampleDate, "dd/mm/yy hh:mm");
      '100   Printer.Print "  Rec'd Date: "; Format(strCurrentReceivedDate, "dd/mm/yy hh:mm")
      'Printer.ForeColor = vbBlack
      'Printer.Print Tab(65); "Page 1 of 1"


60        PrintTextRTB .rtb, FormatString(" Sample Type: EDTA", 18, , Alignleft), 12, , , , vbRed
70        PrintTextRTB .rtb, FormatString(" Specimen Taken: " & Format(strSampleDate, "dd/mm/yy hh:mm"), 31, , Alignleft), 12, , , , vbRed
80        PrintTextRTB .rtb, FormatString(" Rec'd Date: " & Format(strCurrentReceivedDate, "dd/mm/yy hh:mm") & vbCrLf, 30, , Alignleft) & vbCrLf, 12, , , , vbRed
          
90        PrintTextRTB .rtb, FormatString(" Page 1 of 1" & vbCrLf, 30, , Alignleft) & vbCrLf & vbCrLf & vbCrLf, 12, , , , vbBlack



      'Printer.Print
      'Printer.Print

      'Printer.Font.Size = 20
      'Printer.Print "   Patients Rhesus : Rh";
100   PrintTextRTB .rtb, FormatString("   Patients Rhesus : Rh", 23, , Alignleft), 20, , , , vbBlack

      'Printer.Font.Size = 10
      'Printer.CurrentY = Printer.CurrentY + 150
      'Printer.Print "(D) ";

110   PrintTextRTB .rtb, FormatString("(D) ", 3, , Alignleft), 10, , , , vbBlack


      'Printer.Font.Size = 20
      'Printer.CurrentY = Printer.CurrentY - 150
      'Printer.Print frmKleihauer.lblRh
120   PrintTextRTB .rtb, FormatString(" " & frmKleihauer.lblRh & vbCrLf & vbCrLf, 22, , Alignleft), 20, , , , vbBlack


      'Printer.Print
      'Printer.Print "   " & frmKleihauer.lblFMHReport
      'Printer.Print
130   PrintTextRTB .rtb, FormatString(frmKleihauer.lblFMHReport & vbCrLf, 84, , Alignleft), 20, , , , vbBlack


      'Printer.Font.Size = 12

140   Select Case strMessageType
      Case "M1":  'Printer.Print Tab(2); Chr(149) & " No fetal cells seen."
150               PrintTextRTB .rtb, FormatString("  " & Chr(149) & " No fetal cells seen.", 50, , Alignleft), 12, , , , vbBlack

        
160   Case "M2":  'Printer.Print Tab(2); Chr(149) & " <2ml - 1500iu Anti D Is sufficient. No further testing is required."
170               PrintTextRTB .rtb, FormatString("  " & Chr(149) & " <2ml - 1500iu Anti D Is sufficient. No further testing is required.", 80, , Alignleft), 12, , , , vbBlack

180   Case "M3":  sngFMH = Val(intFetal) * 0.4
      '      Printer.Print Tab(2); Chr(149) & " " & sngFMH & "ml - 1500iu Anti D is sufficient. Repeat Kleihauer testing is"
190         PrintTextRTB .rtb, FormatString("  " & Chr(149) & " " & sngFMH & "ml - 1500iu Anti D is sufficient. Repeat Kleihauer testing is", 80, , Alignleft), 12, , , , vbBlack
200         PrintTextRTB .rtb, vbCrLf, 12, , , , vbBlack
      '      Printer.Print Tab(2); "required 72hrs post administration of Anti D."
210         PrintTextRTB .rtb, FormatString(Chr(149) & " required 72hrs post administration of Anti D.", 80, , Alignleft), 12, , , , vbBlack


220   Case "M4":
230         sngFMH = Val(intFetal) * 0.4
            'Printer.Print Tab(2); Chr(149) & " " & sngFMH & "ml - Send sample for flow cytometry urgently. Repeat Kleihauer and"
240          PrintTextRTB .rtb, FormatString("  " & Chr(149) & " " & sngFMH & "ml - Send sample for flow cytometry urgently. Repeat Kleihauer and" & vbCrLf, 80, , Alignleft), 12, , , , vbBlack

            'Printer.Print Tab(2); "flow cytometry samples required 72hrs post administration of Anti D."
250         PrintTextRTB .rtb, FormatString("flow cytometry samples required 72hrs post administration of Anti D.", 80, , Alignleft), 12, , , , vbBlack

            
260   Case "M5":
270         sngFMH = Val(intFetal) * 0.4
            'Printer.Print Tab(2); Chr(149) & " " & sngFMH & "ml - Send sample for flow cytometry urgently. Further Anti D required"
280         PrintTextRTB .rtb, FormatString("  " & Chr(149) & " " & sngFMH & "ml - Send sample for flow cytometry urgently. Further Anti D required" & vbCrLf, 80, , Alignleft), 12, , , , vbBlack
            
            'Printer.Print Tab(2); "discuss with Consultant Haematologist. Repeat Kleihauer and flow cytometry"
290         PrintTextRTB .rtb, FormatString("  discuss with Consultant Haematologist. Repeat Kleihauer and flow cytometry" & vbCrLf, 80, , Alignleft), 12, , , , vbBlack
            
            'printer.Print Tab(2); "samples required 72hrs post administration of Anti D."
300         PrintTextRTB .rtb, FormatString("samples required 72hrs post administration of Anti D.", 80, , Alignleft), 12, , , , vbBlack
310   End Select

320   PrintFooterGHCavan SampleID

330           .rtb.SelPrint Printer.hDC

340           sql = "Select * from PatientDetails where " & _
                    "LabNumber = '" & SampleID & "'"
350           Set TempTb = New Recordset
360           RecOpenServerBB 0, TempTb, sql
370           If TempTb.EOF Then Exit Sub
380           strName = TempTb!Name
          ''''''''''
390           sql = "SELECT * FROM Reports WHERE 0 = 1"
400           Set tb = New Recordset
410           RecOpenServerBB 0, tb, sql
420           tb.AddNew
430           tb!SampleID = SampleID
440           tb!Name = strName
450           tb!Dept = "Kleihauer Report"
460           tb!Initiator = UserName
470           tb!PrintTime = Now    'PrintTime
480           tb!RepNo = "KLE" & SampleID & Format(Now, "ddMMyyyyhhmmss")
490           tb!pagenumber = 1
500           tb!Report = .rtb.TextRTF
510           tb!Printer = Printer.DeviceName
520           tb.Update

      'Printer.EndDoc
530   End With

540   For Each Px In Printers
550     If Px.DeviceName = OriginalPrinter Then
560       Set Printer = Px
570       Exit For
580     End If
590   Next

600   Exit Sub

PrintKleihauerFormCavan_Error:

       Dim strES As String
       Dim intEL As Integer

610    intEL = Erl
620    strES = Err.Description
630    LogError "modPrintForms", "PrintKleihauerFormCavan", intEL, strES, sql

End Sub



Public Sub PrintXMForm(ByVal SampleID As String, _
                       ByRef RowNumbersToPrint() As Integer, _
                       ByVal SampleDate As String, _
                       ByVal HoldFor As Integer, _
                       ByVal Comment As String, _
                       Index As Integer)

      Dim n As Integer
      Dim RNTP(1 To 4) As Integer

10    For n = 1 To UBound(RowNumbersToPrint) Step 4

20      If UBound(RowNumbersToPrint) >= n Then
30        RNTP(1) = RowNumbersToPrint(n)
40      Else
50        RNTP(1) = 0
60      End If
70      If UBound(RowNumbersToPrint) >= n + 1 Then
80        RNTP(2) = RowNumbersToPrint(n + 1)
90      Else
100       RNTP(2) = 0
110     End If
120     If UBound(RowNumbersToPrint) >= n + 2 Then
130       RNTP(3) = RowNumbersToPrint(n + 2)
140     Else
150       RNTP(3) = 0
160     End If
170     If UBound(RowNumbersToPrint) >= n + 3 Then
180       RNTP(4) = RowNumbersToPrint(n + 3)
190     Else
200       RNTP(4) = 0
210     End If
220     PrintXMFormCavan SampleID, RNTP(), SampleDate, HoldFor, Comment, Index
230     UpdatePrinted SampleID, "Form"

240   Next

End Sub









