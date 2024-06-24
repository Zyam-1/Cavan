Attribute VB_Name = "Cytology"
Option Explicit

Public Sub PrintResultCyto(ByVal FullRunNumber As String)

      Dim Hosp As String
      Dim n As Integer
      Dim PL(1 To 23) As String
      Dim tb As Recordset
      Dim sql As String

54930 On Error GoTo PrintResultCyto_Error

54940 sql = "Select * from Demographics where " & _
            "runnumber = '" & FullRunNumber & "'"

54950 Set tb = New Recordset
54960 RecOpenServer 0, tb, sql
54970 If tb.EOF Then Exit Sub

54980 Screen.MousePointer = 11

      'Select Case tb!Hospital & ""
      '  Case "T": Hosp = "Tullamore"
      '  Case "P": Hosp = "Portlaoise"
      '  Case "M": Hosp = "Mullingar"
      '  Case Else: Hosp = ""
      'End Select

54990 Printer.ColorMode = vbPRCMColor
55000 Printer.ScaleMode = vbTwips

55010 Printer.Print ""
55020 Printer.FontName = "Courier New"
55030 Printer.FontSize = 14
55040 Printer.ForeColor = QBColor(4)

55050 Printer.Font.Bold = True
55060 Printer.Print Tab(17); "Midland Regional Hospital at Tullamore"
55070 Printer.Print Tab(17); "       Cytology Laboratory"
55080 Printer.Font.Bold = False
55090 Printer.Print

55100 Printer.FontName = "Courier New"
55110 Printer.FontSize = 10
55120 Printer.ForeColor = QBColor(0)

55130 Printer.Font.Bold = False
55140 Printer.Print Tab(3); " Name: ";
55150 Printer.Font.Bold = True
55160 Printer.Print tb!PatName & "";
55170 Printer.Font.Bold = False
55180 Printer.Print Tab(72); "Lab #: ";
55190 Printer.Font.Bold = True
55200 Printer.Print FullRunNumber

55210 Printer.Font.Bold = False
55220 Printer.Print Tab(3); " Addr: ";
55230 Printer.Font.Bold = True
55240 Printer.Print tb!Addr0 & " " & tb!Addr1 & "";
55250 Printer.Font.Bold = False
55260 Printer.Print Tab(73); "Cons: ";
55270 Printer.Font.Bold = True
55280 Printer.Print tb!Clinician & ""

55290 Printer.Font.Bold = False
55300 Printer.Print Tab(3); "  DoB: ";
55310 Printer.Font.Bold = True
55320 If Not IsNull(tb!DoB) Then
55330   If IsDate(tb!DoB) Then Printer.Print Format(tb!DoB, "dd/mm/yyyy");
55340 End If
55350 Printer.Font.Bold = False
55360 Printer.Print Tab(23); "Sex: ";
55370 Printer.Font.Bold = True
55380 Select Case Left(tb!Sex & " ", 1)
        Case "M": Printer.Print "Male";
55390   Case "F": Printer.Print "Female";
55400 End Select
55410 Printer.Font.Bold = False
55420 Printer.Print Tab(73); "Hosp: ";
55430 Printer.Font.Bold = True
55440 Printer.Print Hosp

55450 Printer.Font.Bold = False
55460 Printer.Print Tab(3); "Chart: ";
55470 Printer.Font.Bold = True
55480 Printer.Print tb!Chart & "";
55490 Printer.Font.Bold = False
55500 Printer.Print Tab(22); "Ward: ";
55510 Printer.Font.Bold = True
55520 Printer.Print tb!Ward & "";
55530 Printer.Font.Bold = False
55540 Printer.Print Tab(75); "GP: ";
55550 Printer.Font.Bold = True
55560 Printer.Print tb!GP & ""

55570 Printer.Font.Bold = False
55580 Printer.Print Tab(3); "Nature of Specimen: ";
55590 Printer.Font.Bold = True
55600 Printer.Print tb!natureofspecimen & "";
55610 Printer.Font.Bold = False
55620 Printer.Print Tab(68); "Date Recd: ";
55630 Printer.Font.Bold = True
55640 Printer.Print tb!GlobalSampleDateTime

55650 If tb!natureofspecimen1 & "" <> "" Then
55660     Printer.Font.Bold = False
55670     Printer.Print Tab(3); "Nature of Specimen: ";
55680     Printer.Font.Bold = True
55690     Printer.Print tb!natureofspecimen1 & "";
55700 End If

55710 If tb!natureofspecimen2 & "" <> "" Then
55720     Printer.Font.Bold = False
55730     Printer.Print Tab(3); "Nature of Specimen: ";
55740     Printer.Font.Bold = True
55750     Printer.Print tb!natureofspecimen2 & "";
55760 End If

55770 If tb!natureofspecimen3 & "" <> "" Then
55780     Printer.Font.Bold = False
55790     Printer.Print Tab(3); "Nature of Specimen: ";
55800     Printer.Font.Bold = True
55810     Printer.Print tb!natureofspecimen3 & "";
55820 End If

55830 Printer.Print Tab(3); String$(93, "-")

      'HistoSplitT tb!cytoreport & "", pl

55840 For n = 1 To 23
55850   If UCase(Left(PL(n), 24)) = "MICROSCOPIC EXAMINATION:" Or _
           UCase(Left(PL(n), 29)) = "BONE MARROW ASPIRATE & BIOPSY" Or _
           UCase(Left(PL(n), 18)) = "GROSS EXAMINATION:" Or _
           UCase(Left(PL(n), 21)) = "SUPPLEMENTARY REPORT:" Or _
           UCase(Left(PL(n), 14)) = "DR. K. CUNNANE" Or _
           UCase(Left(PL(n), 17)) = "DR. KEVIN CUNNANE" Or _
           UCase(Left(PL(n), 11)) = "DR GERARD C" Or _
           UCase(Left(PL(n), 11)) = "PATHOLOGIST" Or _
           UCase(Left(PL(n), 18)) = "DR. J. D. GILSENAN" Or _
           UCase(Left(PL(n), 14)) = "FURTHER REPORT" Or _
           UCase(Left(PL(n), 10)) = "APPEARANCE" Or _
           UCase(Left(PL(n), 23)) = "MICROSCOPIC EXAMINATION" Or _
           UCase(Left(PL(n), 10)) = "CONSULTANT" Or _
           UCase(Left(PL(n), 7)) = "COMMENT" Or _
           UCase(Left(PL(n), 21)) = "SUPPLEMENTARY REPORT" Then
55860     Printer.Font.Bold = True
55870   Else
55880     Printer.Font.Bold = False
55890   End If
55900   Printer.Print Tab(3); PL(n)
55910 Next

55920 Printer.EndDoc

55930 Screen.MousePointer = 0

55940 Exit Sub

PrintResultCyto_Error:

      Dim strES As String
      Dim intEL As Integer

55950 intEL = Erl
55960 strES = Err.Description
55970 LogError "Cytology", "PrintResultCyto", intEL, strES, sql

End Sub


