Attribute VB_Name = "Histology"
Option Explicit
Public gPaperSize As String

Sub PrintResultHisto(ByVal FullRunNumber As String)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
55980 On Error GoTo PrintResultHisto_Error

55990 ReDim PL(1 To 1) As String
      Dim plCounter As Integer
      Dim crPos As Integer
      Dim HR As String
      Dim LinesAllowed As Integer
      Dim TotalPages As Integer
      Dim ThisPage As Integer
      Dim TopLine As Integer
      Dim BottomLine As Integer
      Dim crlfFound As Boolean

56000 If gPaperSize = "A5" Then
56010   LinesAllowed = 23
56020 Else
56030   LinesAllowed = 56
56040 End If

56050 sql = "Select * from historesults, demographics where " & _
            "demographics.sampleid = '" & FullRunNumber & "' " & _
            "and demographics.sampleid = historesults.sampleid"
56060 Set tb = New Recordset
56070 RecOpenServer 0, tb, sql

56080 If tb.EOF Then Exit Sub

56090 Screen.MousePointer = 11



56100 HR = Trim(tb!historeport & "")
56110 crlfFound = True
56120 Do While crlfFound
56130   HR = RTrim(HR)
56140   crlfFound = False
56150   If Right(HR, 1) = vbCr Or Right(HR, 1) = vbLf Then
56160     HR = Left(HR, Len(HR) - 1)
56170     crlfFound = True
56180   End If
56190 Loop

56200 plCounter = 0
56210 Do While Len(HR) > 0
56220   crPos = InStr(HR, vbCr)
56230   If crPos > 0 And crPos < 91 Then
56240     plCounter = plCounter + 1
56250     ReDim Preserve PL(1 To plCounter)
56260     PL(plCounter) = Left(HR, crPos - 1)
56270     HR = Mid(HR, crPos + 2)
56280   Else
56290     If Len(HR) > 91 Then
56300       For n = 91 To 1 Step -1
56310         If Mid(HR, n, 1) = " " Then
56320           Exit For
56330         End If
56340       Next
56350       plCounter = plCounter + 1
56360       ReDim Preserve PL(1 To plCounter)
56370       PL(plCounter) = Left(HR, n)
56380       HR = Mid(HR, n + 1)
56390     Else
56400       plCounter = plCounter + 1
56410       ReDim Preserve PL(1 To plCounter)
56420       PL(plCounter) = HR
56430       Exit Do
56440     End If
56450   End If
56460 Loop

56470 TotalPages = Int((plCounter - 1) / LinesAllowed) + 1
56480 If TotalPages = 0 Then TotalPages = 1
56490 For ThisPage = 1 To TotalPages
56500   PrintHistoHeading tb, FullRunNumber, ThisPage, TotalPages
56510   TopLine = (ThisPage - 1) * LinesAllowed + 1
56520   BottomLine = (ThisPage - 1) * LinesAllowed + LinesAllowed
56530   If BottomLine > plCounter Then
56540     BottomLine = plCounter
56550   End If
56560   For n = TopLine To BottomLine
56570     If UCase(Left(PL(n), 24)) = "MICROSCOPIC EXAMINATION:" Or _
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
56580       Printer.Font.Bold = True
56590     Else
56600       Printer.Font.Bold = False
56610     End If
56620     Printer.Print Tab(3); PL(n)
56630   Next
56640   Printer.EndDoc
56650 Next

56660 Screen.MousePointer = 0

56670 Exit Sub

PrintResultHisto_Error:

      Dim strES As String
      Dim intEL As Integer

56680 intEL = Erl
56690 strES = Err.Description
56700 LogError "Histology", "PrintResultHisto", intEL, strES, sql

End Sub

Public Sub PrintHistoHeading(ByVal tb As Recordset, _
                             ByVal FullRunNumber As String, _
                             ByVal ThisPage As Integer, _
                             ByVal TotalPages As Integer)

      Dim Hosp As String



      'Printer.ColorMode = vbPRCMColor
56710 Printer.ScaleMode = vbTwips

56720 If TotalPages = 1 Then
56730   Printer.Print
56740 Else
56750   Printer.Font.Bold = True
56760   Printer.Print "Page " & Format(ThisPage) & " of " & Format(TotalPages)
56770 End If
56780 Printer.FontName = "Courier New"
56790 Printer.FontSize = 14
56800 Printer.ForeColor = QBColor(4)

56810 Printer.Font.Bold = True
56820 Printer.Print Tab(17); HospName(0) & " Histopathology  Laboratory"
56830 Printer.Font.Bold = False

56840 Printer.FontName = "Courier New"
56850 Printer.FontSize = 10
56860 Printer.ForeColor = QBColor(0)

56870 Printer.Font.Bold = False
56880 Printer.Print Tab(72); "Lab #: ";
56890 Printer.Font.Bold = True
56900 Printer.Print FullRunNumber

56910 Printer.Font.Bold = False
56920 Printer.Print Tab(3); " Name: ";
56930 Printer.Font.Bold = True
56940 Printer.Print tb!PatName & "";
56950 Printer.Font.Bold = False
56960 Printer.Print Tab(73); "Cons: ";
56970 Printer.Font.Bold = True
56980 Printer.Print tb!Clinician & ""

56990 Printer.Font.Bold = False
57000 Printer.Print Tab(3); " Addr: ";
57010 Printer.Font.Bold = True
57020 Printer.Print tb!Addr0 & " " & tb!Addr1 & "";
57030 Printer.Font.Bold = False
57040 Printer.Print Tab(73); "Hosp: ";
57050 Printer.Font.Bold = True
57060 Printer.Print Hosp

57070 Printer.Font.Bold = False
57080 Printer.Print Tab(3); "  DoB: ";
57090 Printer.Font.Bold = True
57100 If Not IsNull(tb!DoB) Then
57110   If IsDate(tb!DoB) Then Printer.Print Format(tb!DoB, "dd/mm/yyyy");
57120 End If
57130 Printer.Font.Bold = False
57140 Printer.Print Tab(23); "Sex: ";
57150 Printer.Font.Bold = True
57160 Select Case Left(tb!Sex & " ", 1)
        Case "M": Printer.Print "Male";
57170   Case "F": Printer.Print "Female";
57180 End Select
57190 Printer.Font.Bold = False
57200 Printer.Print Tab(75); "GP: ";
57210 Printer.Font.Bold = True
57220 Printer.Print tb!GP & ""

57230 Printer.Font.Bold = False
57240 Printer.Print Tab(3); "Chart: ";
57250 Printer.Font.Bold = True
57260 Printer.Print tb!Chart & "";
57270 Printer.Font.Bold = False
57280 Printer.Print Tab(22); "Ward: ";
57290 Printer.Font.Bold = True
57300 Printer.Print tb!Ward & "";
57310 Printer.Font.Bold = False
57320 Printer.Print Tab(68); "Date Recd: ";
57330 Printer.Font.Bold = True
57340 Printer.Print tb!Rundate & " " & tb!TimeTaken

57350 Printer.Font.Bold = False
57360 Printer.Print Tab(3); "Nature of Specimen [A]: ";
57370 Printer.Font.Bold = True
57380 Printer.Print tb!natureofspecimen & "";
57390 If tb!natureofspecimenB & "" <> "" Then
57400   Printer.Font.Bold = False
57410   Printer.Print Tab(55); "[B]: ";
57420   Printer.Font.Bold = True
57430   Printer.Print tb!natureofspecimenB & ""
57440 Else
57450   Printer.Print
57460 End If

57470 If tb!natureofspecimenC & "" <> "" Then
57480   Printer.Font.Bold = False
57490   Printer.Print Tab(22); "[C]: ";
57500   Printer.Font.Bold = True
57510   Printer.Print tb!natureofspecimenC & "";
57520   Printer.Font.Bold = False
57530   Printer.Print Tab(55); "[D]: ";
57540   Printer.Font.Bold = True
57550   Printer.Print tb!natureofspecimenD & ""
57560 Else
57570   Printer.Print
57580 End If

57590 Printer.Print Tab(3); String$(90, "-")

End Sub

