Attribute VB_Name = "modRTFMicro"
Option Explicit

Public MicroCommentLineCount As Integer
Public MicroTotalLineCount As Integer
Public SedimexResultsExist As Boolean

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
Dim RP As ReportToPrint

Private Sub MicroPrintAndStore(ByVal PageNumber As Integer)

      Dim sql As String
      Dim Gx As New GP
      Dim PrintHardCopy As Boolean

7500  On Error GoTo MicroPrintAndStore_Error
          '+++ Junaid
      '        If RP.sampleid = "" Then
      '            RP.sampleid = sampleid
      '        End If
          '--- Junaid
      '    MsgBox "Start"
      '    MsgBox InStr(frmMicroReport.rtb.TextRTF, RP.sampleid)
      '    MsgBox RP.sampleid
      '    MsgBox frmMicroReport.rtb.TextRTF
7510  If InStr(frmMicroReport.rtb.TextRTF, RP.SampleID) Then 'Double check that report saved is for this Sample Id
      'MsgBox "1"
7520  PrintHardCopy = True
      'MsgBox "2"

7530  Gx.LoadName RP.GP
      'MsgBox "3"
      'WRITE ALL RULES HERE WHEN YOU DONT WANT TO PRINT HARD COPY

7540  If (Gx.PrintReport = False And RP.Ward = "GP") Then PrintHardCopy = False * -1E+20
7550  If UCase(RP.PrintAction) = UCase("Save") Then PrintHardCopy = False
7560  If UCase(RP.Initiator) = "BIOMNIS IRELAND" Then PrintHardCopy = False

      'END WRITING RULES FOR PRINTING HARD COPY
7570  If PageNumber = 1 Then
7580      sql = "UPDATE Reports Set Hidden = 1 WHERE SampleID = '" & RP.SampleID & "' AND Dept = 'Microbiology' "
7590      Cnxn(0).Execute sql
7600  End If

7610  With frmMicroReport.rtb
7620      .SelStart = 0

7630      sql = "INSERT INTO Reports " & _
                "(SampleID, Dept, Initiator, PrintTime, ReportNumber, PageNumber, Report, Printer, ReportType) " & _
                "VALUES " & _
                "( '" & RP.SampleID & "', " & _
                "  'Microbiology', " & _
                "  '" & RP.Initiator & "', " & _
                "   replace(convert(NVARCHAR, getdate(), 106), ' ', '/'), " & _
                "  '" & RP.SampleID & "M" & "', " & _
                "  '" & PageNumber & "', " & _
                "  '" & AddTicks(.TextRTF) & "', " & _
                "  '" & IIf(PrintHardCopy, Printer.DeviceName, "None") & "', " & _
                "  '" & IIf(UCase(RP.PrintAction) = UCase("Save"), "Interim Report", "Final Report") & "')"

7640      Cnxn(0).Execute sql

7650      If PrintHardCopy Then

7660          .SelPrint Printer.hdc
7670      End If

7680  End With
7690  Else
      '+++ Junaid 21-11-2022
      '200        LogError "modRTF", "PrintAndStore", 200, RP.sampleid & " not found in RTF report", sql
7700    MsgBox RP.SampleID & " not found in RTF report", vbInformation
      '--- Junaid 21-11-2022
7710  End If

7720  Exit Sub

MicroPrintAndStore_Error:

      Dim strES As String
      Dim intEL As Integer

7730  intEL = Erl
7740  strES = Err.Description
7750  LogError "modRTFMicro", "MicroPrintAndStore", intEL, strES, sql

End Sub

Public Sub RTFPrintText(ByVal Text As String, _
                        Optional FontSize As Integer = 9, _
                        Optional FontBold As Boolean = False, _
                        Optional FontItalic As Boolean = False, _
                        Optional FontUnderLine As Boolean = False, _
                        Optional FontColor As ColorConstants = vbBlack)

7760  On Error GoTo RTFPrintText_Error

7770  With frmMicroReport.rtb
7780      .SelFontSize = FontSize
7790      .SelBold = FontBold
7800      .SelItalic = FontItalic
7810      .SelUnderline = FontUnderLine
7820      .SelColor = FontColor
7830      .SelText = Text
7840  End With

7850  Exit Sub

RTFPrintText_Error:

      Dim strES As String
      Dim intEL As Integer

7860  intEL = Erl
7870  strES = Err.Description
7880  LogError "modRTFMicro", "RTFPrintText", intEL, strES

End Sub


Public Sub RTFPrintResultMicro(SampleID As String)

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


7890  On Error GoTo RTFPrintResultMicro_Error

7900  MicroTotalLineCount = Val(GetOptionSetting("PerPageTotalLineMicro", "72"))
7910  MicroCommentLineCount = Val(GetOptionSetting("MicroCommentLineCount", "4"))
      'MsgBox RP.sampleid
7920  RP.SampleID = SampleID
      'MsgBox RP.sampleid
7930  frmMicroReport.rtb.Text = ""
7940  MiscLineCount = GetMiscLineCount(RP.SampleID)    'FOB+CDiff+Rota/Adeno+OP+RSV
7950  CommentLineCount = GetCommentLineCount(RP.SampleID)
7960  CommentsPresent = CommentLineCount > 0
7970    With frmMicroReport.rtb
7980        .SelText = vbCrLf
7990        .SelFontSize = 10
8000        .SelText = "Sample ID: " & RP.SampleID
8010        .SelText = vbCrLf
8020        .SelAlignment = rtfLeft
8030    End With
      'If Not SetPrinter("MICRO") Then Exit Sub

8040  If isBDMaxReport(RP.SampleID) Then
      'MsgBox "IF"
8050      RTFPrintBDMaxReport "HICSDOPF", RP, 1, 1, "1234", CommentsPresent
8060      MicroPrintAndStore 1
8070  Else
      'MsgBox "else"
8080      MicroscopyLineCount = GetMicroscopyLineCount(RP.SampleID)
8090      IsolateCount = GetIsolateCount(RP.SampleID)
8100      NegativeResults = IsNegativeResults(RP.SampleID)
8110      pDefault = GetPDefault(RP.SampleID)
8120      CSFResultsCount = GetCSFCount(RP.SampleID)
8130      PregnancyLineCount = GetPregnancyLineCount(RP.SampleID)

8140      Select Case IsolateCount
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
8150          PageCount = 1
8160          PageNumber = 1
8170          TotalLines = CSFResultsCount + CommentLineCount + MicroscopyLineCount + MiscLineCount + PregnancyLineCount
8180          If TotalLines > 0 Then
8190              RTFPrintMicroPage "HICDLAMZBGRTVEKPYOWXF", RP, PageNumber, PageCount, "", CommentsPresent
8200              MicroPrintAndStore 1
8210          End If
      'MsgBox "0"
8220      Case 1, 2, 3, 4:
8230          PageCount = 1
8240          PageNumber = 1
8250          If NegativeResults Then
                  
8260              RTFPrintMicroPage "HICDLAMZBGRTVEJKNPYOWXF", RP, PageNumber, PageCount, "", CommentsPresent
8270              MicroPrintAndStore 1
      'MsgBox "1234"
8280          Else
8290              ABCount = GetABCount(RP.SampleID, "1234")
8300              TotalLines = CSFResultsCount + CommentLineCount + MicroscopyLineCount + IsolateCount + ABCount + MiscLineCount

8310              If TotalLines > MicroTotalLineCount Then
8320                  PageCount = 2
8330                  PageNumber = 1
8340                  RTFPrintMicroPage "HIDLJSBGRTVEJKPYOWXF", RP, PageNumber, PageCount, "1234", CommentsPresent

8350                  MicroPrintAndStore 1

8360                  PageNumber = 2
8370                  RTFPrintMicroPage "HICDLAMZPYOWXF", RP, PageNumber, PageCount, "", CommentsPresent
8380                  MicroPrintAndStore 2
8390              Else
8400                  PageCount = 1
8410                  RTFPrintMicroPage "HICDLAMZBTRVEKJSPYOWXF", RP, PageNumber, PageCount, "1234", CommentsPresent
8420                  MicroPrintAndStore 1
8430              End If
      'MsgBox "1234 Else"
8440          End If

8450      Case 5, 6:
8460          PageCount = 2
8470          ABCount = GetABCount(RP.SampleID, "123456")
8480          TotalLines = CSFResultsCount + CommentLineCount + MicroscopyLineCount + IsolateCount + ABCount + MiscLineCount
8490          If TotalLines > MicroTotalLineCount Then
8500              PageCount = 3
8510              PageNumber = 1
8520              RTFPrintMicroPage "HIDLSBGRTVEKPYOWXF", RP, PageNumber, PageCount, "123", CommentsPresent
8530              MicroPrintAndStore 1
      'MsgBox "56"
8540              PageNumber = 2
8550              RTFPrintMicroPage "HIJSF", RP, PageNumber, PageCount, "456", CommentsPresent
8560              MicroPrintAndStore 2

8570              PageNumber = 3
8580              RTFPrintMicroPage "HICAMZF", RP, PageNumber, PageCount, "", CommentsPresent
8590              MicroPrintAndStore 3
8600          Else
8610              PageNumber = 1
8620              RTFPrintMicroPage "HICDLAMZSPYOWXF", RP, PageNumber, PageCount, "123", CommentsPresent
8630              MicroPrintAndStore 1
      'MsgBox "56 Else"
8640              PageNumber = 2
8650              RTFPrintMicroPage "HIJSF", RP, PageNumber, PageCount, "456", CommentsPresent
8660              MicroPrintAndStore 2
8670          End If

8680      End Select

8690  End If

      '740   ReSetPrinter

8700  Exit Sub

RTFPrintResultMicro_Error:

      Dim strES As String
      Dim intEL As Integer

8710  intEL = Erl
8720  strES = Err.Description
      '        MsgBox strES
8730  LogError "modRTFMicro", "RTFPrintResultMicro", intEL, strES

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
      Dim DoB As String
      Dim Chart As String
      Dim Address0 As String
      Dim Address1 As String
      Dim Sex As String
      Dim Hospital As String

8740  On Error GoTo RTFPrintBDMaxReport_Error

      '20    sql = "Select * from Demographics where " & _
      '            "SampleID = '" & RP.SampleID + sysOptMicroOffset(0) & "'"
8750  sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
8760  Set tb = New Recordset
8770  RecOpenClient 0, tb, sql
8780  If Not tb.EOF Then
8790      PatName = tb!PatName & ""
8800      If IsDate(tb!DoB) Then
8810          DoB = Format(tb!DoB, "dd/mmm/yyyy")
8820      End If
8830      Chart = tb!Chart & ""
8840      Address0 = tb!Addr0 & ""
8850      Address1 = tb!Addr1 & ""
8860      Sex = tb!Sex & ""
8870      Hospital = tb!Hospital & ""
8880  End If

8890  For n = 1 To Len(HCLM)
8900      Select Case Mid$(HCLM, n, 1)
          Case "A": RTFPrintMicroCurrentABs RP.SampleID
8910      Case "C": RTFPrintMicroClinDetails RP.SampleID
8920      Case "D": RTFPrintMicroComment RP.SampleID, "Demographics"
8930      Case "F": RTFPrintMicroFooter RP
          'Case "H": RTFPrintHeading "Microbiology", PatName, Dob, Chart, Address0, Address1, Sex, Hospital
8940      Case "I": RTFPrintSpecType RP, PageNumber, PageCount
8950      Case "O": RTFPrintMicroComment RP.SampleID, "Consultant"
8960      Case "P": RTFPrintMicroComment RP.SampleID, "Med Scientist"
8970      Case "S": RTFPrintMicroBDMax RP.SampleID, IsolateNumbers
8980      End Select
8990  Next

9000  Exit Sub

RTFPrintBDMaxReport_Error:

      Dim strES As String
      Dim intEL As Integer

9010  intEL = Erl
9020  strES = Err.Description
9030  LogError "modRTFMicro", "RTFPrintBDMaxReport", intEL, strES, sql

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
      Dim DoB As String
      Dim Chart As String
      Dim Address0 As String
      Dim Address1 As String
      Dim Sex As String
      Dim Hospital As String

9040  On Error GoTo RTFPrintMicroPage_Error

      '20    sql = "Select * from Demographics where " & _
      '            "SampleID = '" & RP.SampleID + sysOptMicroOffset(0) & "'"
9050  sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
9060  Set tb = New Recordset
9070  RecOpenClient 0, tb, sql
9080  If Not tb.EOF Then
9090      PatName = tb!PatName & ""
9100      If IsDate(tb!DoB) Then
9110          DoB = Format(tb!DoB, "dd/mmm/yyyy")
9120      End If
9130      Chart = tb!Chart & ""
9140      Address0 = tb!Addr0 & ""
9150      Address1 = tb!Addr1 & ""
9160      Sex = tb!Sex & ""
9170      Hospital = tb!Hospital & ""
9180  End If

9190  For n = 1 To Len(HCLM)
9200      Debug.Print Mid$(HCLM, n, 1)
9210      Select Case Mid$(HCLM, n, 1)

          Case "A": RTFPrintMicroCurrentABs RP.SampleID
9220      Case "B": RTFPrintMicroOccultBlood RP.SampleID
9230      Case "C": RTFPrintMicroClinDetails RP.SampleID
9240      Case "D": RTFPrintMicroComment RP.SampleID, "Demographics"
9250      Case "L": RTFPrintMicroComment RP.SampleID, "L"
9260      Case "E": RTFPrintMicroRSV RP.SampleID
9270      Case "F": RTFPrintMicroFooter RP
9280      Case "G": RTFPrintMicroPregnancy RP.SampleID
9290      Case "H": 'RTFPrintHeading "Microbiology", PatName, Dob, Chart, Address0, Address1, Sex, Hospital, CInt(PageCount), CInt(PageNumber)
9300      Case "I": RTFPrintSpecType RP, PageNumber, PageCount
9310      Case "J": RTFPrintMicroCSF RP.SampleID
9320      Case "K": RTFPrintMicroRedSub RP.SampleID
9330      Case "M": RTFPrintMicroscopy RP.SampleID
9340      Case "N": RTFPrintNegativeResults RP.SampleID
9350      Case "O": RTFPrintMicroComment RP.SampleID, "Consultant"
9360      Case "P": RTFPrintMicroComment RP.SampleID, "Med Scientist"
9370      Case "R": RTFPrintMicroRotaAdeno RP.SampleID
9380      Case "S": RTFPrintMicroSensitivities RP.SampleID, IsolateNumbers, CommentsPresent
9390      Case "T": RTFPrintMicroCDiff RP.SampleID
9400      Case "V": RTFPrintMicroOvaParasites RP.SampleID
9410      Case "W": RTFPrintMicroAssIDBC RP.SampleID
9420      Case "X": RTFPrintMicroAssIDMRSA RP.SampleID
9430      Case "Y": 'RTFPrintMicroUrineComment RP.SampleID
9440      Case "Z": RTFPrintMicroGramWetPrep RP.SampleID

9450      End Select
9460  Next
9470  Debug.Print "Finish"
9480  Exit Sub

RTFPrintMicroPage_Error:

      Dim strES As String
      Dim intEL As Integer

9490  intEL = Erl
9500  strES = Err.Description
9510  LogError "modRTFMicro", "RTFPrintMicroPage", intEL, strES

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
      Dim s As String
      Dim X() As String
      Dim LL As Integer
      Dim AccreditationText As String

9520  On Error GoTo RTFPrintMicroFooter_Error

      '20    sql = "Select AuthoriserCode from Sensitivities where " & _
      '            "SampleID = '" & Val(RP.SampleID) + sysOptMicroOffset(0) & "' " & _
      '            "and (AuthoriserCode <> '' or AuthoriserCode is not null)"
9530  sql = "Select AuthoriserCode from Sensitivities where " & _
            "SampleID = '" & Val(RP.SampleID) & "' " & _
            "and (AuthoriserCode <> '' or AuthoriserCode is not null)"
9540  Set tb = New Recordset
9550  RecOpenClient 0, tb, sql
9560  If Not tb.EOF Then
9570      sql = "Select Name from Users where Code = '" & tb!AuthoriserCode & "'"
9580      Set tb = New Recordset
9590      RecOpenClient 0, tb, sql
9600      If Not tb.EOF Then
9610          Authoriser = tb!Name & ""
9620      End If
9630  End If

      '130   sql = "Select * from Demographics where " & _
      '            "SampleID = '" & Val(RP.SampleID) + sysOptMicroOffset(0) & "'"
9640  sql = "Select * from Demographics where " & _
            "SampleID = '" & Val(RP.SampleID) & "'"
9650  Set tb = New Recordset
9660  RecOpenClient 0, tb, sql
9670  If Not tb.EOF Then
9680      Operator = tb!Operator & ""
9690      If Not IsNull(tb!RecDate) Then
9700          RecDate = Format(tb!RecDate, "dd/mm/yy hh:mm")
9710      Else
9720          RecDate = ""
9730      End If
9740      If Not IsNull(tb!SampleDate) Then
9750          SampleDate = Format(tb!SampleDate, "dd/mm/yy hh:mm")
9760          If Right$(SampleDate, 5) = "00:00" Then
9770              SampleDate = Format$(SampleDate, "dd/mm/yy")
9780          End If
9790      Else
9800          SampleDate = ""
9810      End If
9820      If Not IsNull(tb!Rundate) Then
9830          Rundate = Format(tb!Rundate, "dd/mm/yy")
9840      Else
9850          Rundate = ""
9860      End If
9870  End If

      '370   sql = "SELECT * FROM PrintValidLog WHERE " & _
      '            "SampleID = '" & Val(RP.SampleID) + sysOptMicroOffset(0) & "'"
9880  sql = "SELECT * FROM PrintValidLog WHERE " & _
            "SampleID = '" & Val(RP.SampleID) & "'"
9890  Set tb = New Recordset
9900  RecOpenClient 0, tb, sql
9910  If Not tb.EOF Then
9920      If IsDate(tb!ValidatedDateTime) Then
9930          ValidatedDate = Format$(tb!ValidatedDateTime, "dd/MM/yy hh:mm")
9940      Else
9950          ValidatedDate = Format$(Now, "dd/MM/yy hh:mm")
9960      End If
9970      Operator = tb!ValidatedBy & ""
9980  Else
9990      ValidatedDate = Format$(Now, "dd/MM/yy hh:mm")
10000 End If

10010 With frmMicroReport.rtb
10020     .SelFontName = "Courier New"
      '520       .SetFocus
10030     .SelFontSize = 10
10040     .SelColor = vbBlack

10050     If Trim$(RP.SendCopyTo) <> "" Then
10060         .SelText = RP.Clinician & " Requested copy to be sent to " & RP.SendCopyTo
10070     End If
10080     .SelText = vbCrLf
10090     .SelFontSize = 10

      '600       .SetFocus
10100     X = Split(.Text, vbCr)
10110     LL = UBound(X)


10120     Do While LL < MicroTotalLineCount
10130         .SelFontSize = 10
10140         .SelText = vbCrLf
10150         X = Split(.Text, vbCr)
10160         LL = UBound(X)
10170     Loop

10180     .SelFontName = "Courier New"
10190     .SelAlignment = rtfCenter
10200     .SelBold = False
      '720       If gPrintCopyReport = 0 Then
      '730           .SelFontSize = 4
      '740           .SelText = String$(200, "-") & vbCrLf
      '750       Else
10210         .SelFontSize = 8
10220         s = "- THIS IS A COPY REPORT - NOT FOR FILING -"
10230         s = s & s
10240         s = s & "- THIS IS A COPY REPORT -"
10250         .SelColor = vbRed
10260         .SelText = s & vbCrLf
      '820       End If

10270     .SelFontName = "Courier New"
10280     .SelFontSize = 10
10290     .SelBold = False
10300     .SelAlignment = rtfLeft
10310     .SelColor = vbBlack
10320     .SelText = "Micro "
10330     .SelColor = vbBlue
10340     .SelFontName = "MS Sans Serif"
10350     .SelBold = False
10360     .SelFontSize = 10

10370     .SelText = "Specimen Date:" & SampleDate
10380     .SelText = " Received:" & RecDate
10390     .SelText = " Reported:" & Format$(ValidatedDate, "dd/mm/yy hh:mm")

      '960       If gPrintCopyReport = 0 Then
      '970           .SelText = " Reported by " & Operator
      '980       Else
10400         sql = "SELECT TOP 1 Viewer FROM ViewedReports WHERE " & _
                    "SampleID = '" & RP.SampleID & "' " & _
                    "AND Discipline = 'N' " & _
                    "ORDER BY DateTime DESC"
10410         Set tb = New Recordset
10420         RecOpenServer 0, tb, sql
10430         If Not tb.EOF Then
10440             .SelText = " preview by " & tb!Viewer & ""
10450         Else
10460             .SelText = " preview by " & Left$(RP.Initiator, 14)
10470         End If
      '1070      End If
10480     AccreditationText = GetOptionSetting("MicroAccreditation", "")
          
      '    If Not isBDMaxReport(RP.SampleID) Then
      '        AccreditationText = GetOptionSetting("MicroAccreditation", "")
      '    End If

10490     If AccreditationText <> "" Then
10500         .SelText = vbNewLine
10510         .SelText = Space$(10) & String$(75, "-") & vbCrLf

10520         .SelColor = vbRed
10530         .SelAlignment = rtfLeft

10540         .SelFontName = "Courier New"
10550         .SelFontSize = 10
10560         .SelBold = True
10570         .SelText = Space$(5) & AccreditationText
10580         .SelText = vbCrLf
10590     End If

10600 End With

10610 Exit Sub

RTFPrintMicroFooter_Error:

      Dim strES As String
      Dim intEL As Integer

10620 intEL = Erl
10630 strES = Err.Description
10640 LogError "modRTFMicro", "RTFPrintMicroFooter", intEL, strES, sql

End Sub

Private Function isBDMaxReport(ByVal strSID As String) As Boolean
      Dim tb As Recordset
      Dim sql As String

10650 On Error GoTo isBDMaxReport_Error

      '20    sql = "SELECT SampleID From Isolates WHERE SampleID = '" & Val(strSID) + sysOptMicroOffset(0) & "' " & _
      '            "and OrganismGroup like 'BD Max%'"
10660 sql = "SELECT SampleID From Isolates WHERE SampleID = '" & Val(strSID) & "' " & _
            "and OrganismGroup like 'BD Max%'"

10670 Set tb = New Recordset
10680 RecOpenServer 0, tb, sql
10690 If tb.EOF Then
10700     isBDMaxReport = False
10710 Else
10720     isBDMaxReport = True
10730 End If

10740 Exit Function

isBDMaxReport_Error:

      Dim strES As String
      Dim intEL As Integer

10750 intEL = Erl
10760 strES = Err.Description
10770 LogError "modRTFMicro", "isBDMaxReport", intEL, strES, sql

End Function

Private Sub RTFPrintMicroGramWetPrep(ByVal SampleID As String)

      Dim ID As IdentResult
      Dim IDs As New IdentResults
      Dim Title As String

10780 On Error GoTo RTFPrintMicroGramWetPrep_Error

10790 With frmMicroReport.rtb
10800     IDs.Load (SampleID)
10810     If IDs.Count > 0 Then
10820         Title = "Gram Stain: "
10830         For Each ID In IDs
10840             If UCase$(ID.TestType) = "GRAMSTAIN" Then
10850                 .SelBold = False
10860                 .SelFontSize = 10
10870                 .SelText = Space$(10) & Title
10880                 .SelBold = True
10890                 .SelText = ID.TestName & " " & ID.Result
10900                 .SelText = vbCrLf
10910                 Title = "            "
10920             End If
10930         Next

10940         Title = "  Wet Prep: "
10950         For Each ID In IDs
10960             If UCase$(ID.TestType) = "WETPREP" Then
10970                 .SelBold = False
10980                 .SelFontSize = 10
10990                 .SelText = Space$(10) & Title
11000                 .SelBold = True
11010                 .SelText = ID.TestName & " " & ID.Result
11020                 .SelText = vbCrLf
11030                 Title = "            "
11040             End If
11050         Next
11060     End If

11070     .SelText = vbCrLf

11080     .SelBold = False
11090     .SelFontSize = 10
11100 End With

11110 Exit Sub

RTFPrintMicroGramWetPrep_Error:

      Dim strES As String
      Dim intEL As Integer

11120 intEL = Erl
11130 strES = Err.Description
11140 LogError "modRTFMicro", "RTFPrintMicroGramWetPrep", intEL, strES

End Sub


Private Sub RTFPrintMicroUrineComment(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim SDS As New SiteDetails

11150 On Error GoTo RTFPrintMicroUrineComment_Error

11160 SDS.Load Val(SampleID) ' + sysOptMicroOffset(0)
11170 If SDS.Count > 0 Then
11180     If UCase$(SDS(1).Site) <> "URINE" Then
11190         Exit Sub
11200     End If
11210 End If

      '80    sql = "Select * from Isolates where " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "' " & _
      '            "AND OrganismGroup <> 'Negative results' " & _
      '            "AND OrganismName <> ''"
11220 sql = "Select * from Isolates where " & _
            "SampleID = '" & Val(SampleID) & "' " & _
            "AND OrganismGroup <> 'Negative results' " & _
            "AND OrganismName <> ''"
11230 Set tb = New Recordset
11240 RecOpenServer 0, tb, sql
11250 If tb.EOF Then Exit Sub

11260 With frmMicroReport.rtb
11270     .SelBold = False
11280     .SelFontSize = 10
11290     .SelText = "Positive cultures "
11300     .SelUnderline = True
11310     .SelText = "must"
11320     .SelUnderline = False
11330     .SelText = " be correlated with signs and symptoms of UTI"
11340     .SelText = vbCrLf
11350     .SelText = "Particularly with low colony counts" & vbCrLf
11360 End With

11370 Exit Sub

RTFPrintMicroUrineComment_Error:

      Dim strES As String
      Dim intEL As Integer

11380 intEL = Erl
11390 strES = Err.Description
11400 LogError "modRTFMicro", "RTFPrintMicroUrineComment", intEL, strES, sql

End Sub


Private Sub RTFPrintMicroAssIDMRSA(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
11410 On Error GoTo RTFPrintMicroAssIDMRSA_Error

11420 ReDim AssID(0 To 0) As Variant
      Dim ThisID As String
      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim Found As Boolean
      '
      '30    sql = "SELECT AssID FROM AssociatedIDs WHERE " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
11430 sql = "SELECT AssID FROM AssociatedIDs WHERE " & _
            "SampleID = '" & Val(SampleID) & "'"
11440 Set tb = New Recordset
11450 RecOpenServer 0, tb, sql
11460 n = -1
11470 Do While Not tb.EOF
11480     n = n + 1
11490     ReDim Preserve AssID(0 To n) As Variant
11500     AssID(n) = tb!AssID
11510     tb.MoveNext
11520 Loop
      '130   sql = "SELECT SampleID FROM AssociatedIDs WHERE " & _
      '            "AssID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
11530 sql = "SELECT SampleID FROM AssociatedIDs WHERE " & _
            "AssID = '" & Val(SampleID) & "'"
11540 Set tb = New Recordset
11550 RecOpenServer 0, tb, sql
11560 Do While Not tb.EOF
11570     Found = False
11580     For X = 0 To UBound(AssID)
11590         If AssID(X) = tb!SampleID Then
11600             Found = True
11610             Exit For
11620         End If
11630     Next
11640     If Not Found Then
11650         n = n + 1
11660         ReDim Preserve AssID(0 To n) As Variant
11670         AssID(n) = tb!SampleID
11680     End If
11690     tb.MoveNext
11700 Loop

11710 If n = -1 Then Exit Sub

11720 With frmMicroReport.rtb
11730     .SelFontSize = 10
11740     .SelBold = False

11750     .SelText = "This Result relates to the Site specified on this form only."
11760     .SelText = vbCrLf
11770     .SelText = "Please refer to Results for Lab numbers "
11780     s = ""
11790     For n = 0 To UBound(AssID)
      '400           ThisID = Format$(AssID(n) - sysOptMicroOffset(0))
11800         ThisID = Format$(AssID(n))
11810         s = s & ThisID & ", "
11820     Next
11830     s = Left$(s, Len(s) - 2)
11840     .SelText = s & " as part of this series of screens."
11850     .SelText = vbCrLf
11860 End With

11870 Exit Sub

RTFPrintMicroAssIDMRSA_Error:

      Dim strES As String
      Dim intEL As Integer

11880 intEL = Erl
11890 strES = Err.Description
11900 LogError "modRTFMicro", "RTFPrintMicroAssIDMRSA", intEL, strES, sql

End Sub

Private Sub RTFPrintMicroAssIDBC(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim AssID As String

11910 On Error GoTo RTFPrintMicroAssIDBC_Error

      '20    sql = "SELECT AssID FROM Demographics WHERE " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
11920 sql = "SELECT AssID FROM Demographics WHERE " & _
            "SampleID = '" & Val(SampleID) & "'"
11930 Set tb = New Recordset
11940 RecOpenServer 0, tb, sql

11950 If Not tb.EOF Then

11960     If Trim$(tb!AssID & "") <> "" Then

11970         With frmMicroReport.rtb
11980             .SelFontSize = 10
11990             .SelBold = False

      '100               AssID = Format$(tb!AssID - sysOptMicroOffset(0))
12000             AssID = Format$(tb!AssID)
12010             .SelText = "Please refer to Lab number " & AssID & _
                             " for associated Lab Result."
12020             .SelText = vbCrLf
12030         End With

12040     End If

12050 End If

12060 Exit Sub

RTFPrintMicroAssIDBC_Error:

      Dim strES As String
      Dim intEL As Integer

12070 intEL = Erl
12080 strES = Err.Description
12090 LogError "modRTFMicro", "RTFPrintMicroAssIDBC", intEL, strES, sql

End Sub



Private Sub RTFPrintMicroCDiff(ByVal SampleID As String)

      Dim ResultPCR As String
      Dim ResultToxin As String
      Dim Gx As GenericResult
      Dim GXs As New GenericResults
      Dim Fx As FaecesResult
      Dim Fxs As New FaecesResults

12100 On Error GoTo RTFPrintMicroCDiff_Error

12110 ResultPCR = ""
12120 ResultToxin = ""

12130 GXs.Load Val(SampleID) ' + sysOptMicroOffset(0)
12140 Set Gx = GXs("cDiffPCR")
12150 If Not Gx Is Nothing Then
12160     ResultPCR = Gx.Result
12170 End If
12180 Fxs.Load Val(SampleID) ' + sysOptMicroOffset(0)
12190 Set Fx = Fxs("ToxinAL")
12200 If Not Fx Is Nothing Then
12210     ResultToxin = Fx.Result
12220 End If

12230 If ResultPCR = "" And ResultToxin = "" Then
12240     Exit Sub
12250 End If

12260 With frmMicroReport.rtb
12270     .SelText = vbCrLf
12280     .SelBold = False
12290     .SelFontSize = 10

12300     If Trim$(ResultPCR) <> "" Then
12310         .SelText = Space$(10) & "C. difficile PCR : "
12320         .SelBold = True
12330         .SelText = ResultPCR
12340         .SelText = vbCrLf
12350     End If

12360     If Trim$(ResultToxin) <> "" Then
12370         .SelBold = False
12380         .SelText = Space$(10) & "Clostridium difficile Toxin A/B : "
12390         .SelBold = True
12400         If UCase$(ResultToxin) = "N" Then
12410             .SelText = "Not Detected"
12420         ElseIf UCase$(ResultToxin) = "P" Then
12430             .SelText = "Positive"
12440         ElseIf UCase$(ResultToxin) = "I" Then
12450             .SelText = "Inconclusive"
12460         ElseIf UCase$(ResultToxin) = "R" Then
12470             .SelText = "Sample Rejected"
12480             .SelBold = False
12490             .SelFontSize = 10
12500             .SelText = "I wish to remind you that "
12510             .SelItalic = True
12520             .SelText = "C. difficile"
12530             .SelItalic = False
12540             .SelText = " should be requested only when there is a high index"
12550             .SelText = vbCrLf
12560             .SelText = "of suspicion. The clinical details received with this "
12570             .SelText = "test request fail to meet the criteria for"
12580             .SelText = vbCrLf
12590             .SelText = "testing and as such has been deemed unsuitable for analysis. "
12600             .SelText = vbCrLf
12610             .SelText = "Please refer to the following"
12620             .SelText = vbCrLf
12630             .SelText = "guidelines for requesting "
12640             .SelItalic = True
12650             .SelText = "C. difficile"
12660             .SelItalic = False
12670             .SelText = " toxin testing."
12680             .SelText = vbCrLf
12690             .SelItalic = True
12700             .SelBold = True
12710             .SelUnderline = True
12720             .SelText = Space$(30) & "C. difficile"
12730             .SelItalic = False
12740             .SelText = " toxin"
12750             .SelText = vbCrLf
12760             .SelBold = False
12770             .SelUnderline = False

12780             .SelText = "Acute onset of loose stools (more than three within a "
12790             .SelText = "24-hour period) for two days"
12800             .SelText = vbCrLf

12810             .SelText = "without another aetiology, onset after >3days in hospital,"
12820             .SelText = " and a history of antibiotic use"
12830             .SelText = vbCrLf

12840             .SelText = "or chemotherapy;"
12850             .SelText = vbCrLf

12860             .SelBold = True
12870             .SelUnderline = True
12880             .SelText = Space$(42) & "or"
12890             .SelText = vbCrLf
12900             .SelBold = False
12910             .SelUnderline = False

12920             .SelText = "Recurrence of diarrhoea within eight weeks of the end of "
12930             .SelText = "previous treatment of "
12940             .SelItalic = True
12950             .SelText = "C. difficile"
12960             .SelText = vbCrLf

12970             .SelItalic = False
12980             .SelText = "infection."
12990             .SelText = vbCrLf
13000         Else
13010             .SelText = vbCrLf
13020         End If
13030     End If
13040 End With

      'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "CDIFF", True, True

13050 Exit Sub

RTFPrintMicroCDiff_Error:

      Dim strES As String
      Dim intEL As Integer

13060 intEL = Erl
13070 strES = Err.Description
13080 LogError "modRTFMicro", "RTFPrintMicroCDiff", intEL, strES

End Sub
Private Sub ClearPrintingRule(ByRef rtb As RichTextBox)
13090 On Error GoTo ClearPrintingRule_Error

13100 rtb.SelBold = False
13110 rtb.SelItalic = False
13120 rtb.SelUnderline = False

13130 Exit Sub
ClearPrintingRule_Error:
         
13140 LogError "modRTFMicro", "ClearPrintingRule", Erl, Err.Description


End Sub

Private Sub ApplyPrintingRule(ByRef rtb As RichTextBox, ByVal TestName As String, ByVal TestType As String, ByVal Value As String)

      Dim tb As Recordset
      Dim sql As String

13150 On Error GoTo ApplyPrintingRule_Error

13160 sql = "SELECT * FROM PrintingRules WHERE TestName = '" & TestName & "' AND Type = '" & TestType & "'"
13170 Set tb = New Recordset
13180 RecOpenServer 0, tb, sql
13190 If Not tb.EOF Then
13200     If tb!Criteria & "" = Value Then
13210         rtb.SelBold = tb!Bold
13220         rtb.SelItalic = tb!Italic
13230         rtb.SelUnderline = tb!Underline
13240     End If
13250 End If

13260 Exit Sub
ApplyPrintingRule_Error:
         
13270 LogError "modRTFMicro", "ApplyPrintingRule", Erl, Err.Description, sql


End Sub


Private Sub RTFPrintMicroSensitivities(ByVal SampleID As String, _
                                       ByVal IsolateNumbers As String, _
                                       ByVal CommentsPresent As Boolean)

      Dim tb         As Recordset
      Dim sql        As String
      Dim strGroup(1 To 8) As OrgGroup
      Dim ABCount    As Integer
      Dim n          As Integer
      Dim X          As Integer
      Dim Y          As Integer
      Dim MaxIsolates As Integer
      Dim SampleIDWithOffset As Variant
      Dim SensPrintMax As Integer
      Dim Site       As String
      Dim SDS        As New SiteDetails

13280 On Error GoTo RTFPrintMicroSensitivities_Error

13290 ReDim ResultArray(0 To 0) As ABResult
      '+++ Junaid 20-05-2024
      '30    SampleIDWithOffset = Val(SampleID) + sysOptMicroOffset(0)
13300 SampleIDWithOffset = Val(SampleID)
      '--- Junaid
13310 SensPrintMax = 3

13320 SDS.Load SampleIDWithOffset
13330 If SDS.Count > 0 Then
13340     Site = SDS(1).Site

13350     sql = "Select [Default] as D from Lists where " & _
                "ListType = 'SI' " & _
                "and Text = '" & Site & "'"
13360     Set tb = New Recordset
13370     RecOpenServer 0, tb, sql
13380     If Not tb.EOF Then
13390         SensPrintMax = Val(tb!D & "")
13400     End If
13410 End If

13420 LoadResultArray SampleIDWithOffset, ResultArray()

13430 ABCount = UBound(ResultArray())

13440 MaxIsolates = FillOrgGroups(strGroup(), SampleIDWithOffset)

13450 With frmMicroReport.rtb
13460     If Not CommentsPresent Then
13470         For n = 1 To 10 - ABCount
13480             .SelText = vbCrLf
13490         Next
13500     End If

13510     .SelBold = False
13520     .SelFontSize = 10
13530     .SelUnderline = True
13540     .SelText = "Isolates  "
13550     .SelUnderline = False
13560     .SelText = vbCrLf

13570     For X = 1 To 4
13580         n = Val(Mid$(IsolateNumbers, X, 1))
13590         If strGroup(n).OrgName <> "" Then
13600             .SelText = Left$(X & ": " & strGroup(n).ReportName & Space$(40), 40)
13610             ApplyPrintingRule frmMicroReport.rtb, strGroup(n).OrgName, "Qualifier", strGroup(n).Qualifier
13620             .SelText = strGroup(n).Qualifier
13630             ClearPrintingRule frmMicroReport.rtb
13640             .SelText = vbCrLf
                  '.SelText = Space$(10)
13650         Else
13660             .SelText = Left$(strGroup(n).OrgGroup & Space$(40), 40)
13670             .SelText = vbCrLf
13680         End If
13690     Next

13700     If ABCount > 0 Then
13710         .SelFontSize = 10
              '.SelBold = False
              '.SelText = vbCrLf

13720         .SelUnderline = True
13730         .SelText = Left$("Sensitivities" & Space$(20), 20)

13740         .SelText = IIf(strGroup(1).OrgName <> "", Left$("1" & Space(19), 19), Space(19))
13750         .SelText = IIf(strGroup(2).OrgName <> "", Left$("2" & Space(19), 19), Space(19))
13760         .SelText = IIf(strGroup(3).OrgName <> "", Left$("3" & Space(19), 19), Space(19))
13770         .SelText = IIf(strGroup(4).OrgName <> "", Left$("4" & Space(19), 19), Space(19))
13780         .SelText = vbCrLf
13790         Select Case IsolateNumbers
                  Case "1234":
13800                 .SelText = Left$("" & Space$(20), 20)
13810                 .SelText = Left$(strGroup(1).ShortName & Space(19), 19)
13820                 .SelText = Left$(strGroup(2).ShortName & Space(19), 19)
13830                 .SelText = Left$(strGroup(3).ShortName & Space(19), 19)
13840                 .SelText = Left$(strGroup(4).ShortName & Space(19), 19)
13850                 .SelText = vbCrLf
13860             Case "5678":
13870                 .SelText = Left$("" & Space$(20), 20)
13880                 .SelText = Left$(strGroup(5).ShortName & Space(19), 19)
13890                 .SelText = Left$(strGroup(6).ShortName & Space(19), 19)
13900                 .SelText = Left$(strGroup(7).ShortName & Space(19), 19)
13910                 .SelText = Left$(strGroup(8).ShortName & Space(19), 19)
13920                 .SelText = vbCrLf
13930         End Select
13940         .SelUnderline = False

              Dim Sxs As New Sensitivities
              Dim sx As Sensitivity

13950         Sxs.Load Val(SampleID) ' + sysOptMicroOffset(0)

13960         For Y = 1 To ABCount
13970             .SelColor = vbBlack
13980             .SelText = Left$(ResultArray(Y).AntibioticName & Space$(20), 20)
                  'IsolateNumbers = "1234"
13990             For X = 1 To 4
14000                 n = Val(Mid$(IsolateNumbers, X, 1))
14010                 If Trim$(strGroup(n).ReportName) <> "" Then
14020                     Set sx = Sxs.Item(n, ResultArray(Y).AntibioticCode)
14030                     If Not sx Is Nothing Then
14040                         If sx.Report Then
14050                             If sx.RSI = "R" Then
14060                                 .SelColor = vbRed
14070                                 .SelText = Left$("Resistant" & Space(19), 19)
14080                             ElseIf sx.RSI = "S" Then
14090                                 .SelColor = vbGreen
14100                                 .SelText = Left$("Sensitive" & Space(19), 19)
14110                             ElseIf sx.RSI = "I" Then
14120                                 .SelColor = vbBlack
14130                                 .SelText = Left$("Intermediate" & Space(19), 19)
14140                             Else
14150                                 .SelText = Left$(sx.RSI & Space(19), 19)
14160                             End If
14170                         Else
14180                             .SelText = Space(19)
14190                         End If
14200                     Else
14210                         .SelText = Space$(19)
14220                     End If
14230                 Else
14240                     .SelText = Space$(19)
14250                 End If
14260             Next
14270             .SelText = vbCrLf
14280         Next
14290     End If




          'UpdatePrintValid SampleIDWithOffset, "CANDS", True, True

14300     .SelColor = vbBlack

14310 End With

14320 Exit Sub

RTFPrintMicroSensitivities_Error:

      Dim strES      As String
      Dim intEL      As Integer

14330 intEL = Erl
14340 strES = Err.Description
14350 LogError "modRTFMicro", "RTFPrintMicroSensitivities", intEL, strES, sql

End Sub


Private Sub RTFPrintMicroBDMax(ByVal SampleID As String, _
                               ByVal IsolateNumbers As String)

      Dim strGroup(1 To 8) As OrgGroup
      Dim ABCount As Integer
      Dim n As Integer
      Dim X As Integer
      Dim MaxIsolates As Integer
      Dim SampleIDWithOffset As Variant

14360 ReDim ResultArray(0 To 0) As ABResult
      '+++ Junaid 20-05-2024
      '20    SampleIDWithOffset = Val(SampleID) + sysOptMicroOffset(0)
14370 SampleIDWithOffset = Val(SampleID)
      '--- Junaid
14380 MaxIsolates = FillOrgGroups(strGroup(), SampleIDWithOffset)

14390 With frmMicroReport.rtb
14400     .SelBold = True
14410     .SelFontSize = 10
14420     .SelText = vbCrLf
14430     .SelText = Space$(10)
14440     For X = 1 To 4
14450         n = Val(Mid$(IsolateNumbers, X, 1))
14460         If strGroup(n).OrgName <> "" Then
14470             .SelText = Left$(strGroup(n).ReportName & Space$(30), 30)
14480             .SelText = strGroup(n).Qualifier
14490             .SelText = vbCrLf
14500             .SelText = Space$(10)
14510         Else
14520             .SelText = Left$(strGroup(n).OrgGroup & Space$(30), 30)
14530             .SelText = vbCrLf
14540         End If
14550     Next
14560     .SelFontSize = 10
14570     .SelBold = False
14580     .SelText = vbCrLf

      '240       UpdatePrintValid SampleIDWithOffset, "CANDS", True, True

14590     .SelColor = vbBlack

14600 End With


End Sub



Private Sub RTFPrintMicroRotaAdeno(ByVal SampleID As String)

      Dim Fxs As New FaecesResults

14610 On Error GoTo RTFPrintMicroRotaAdeno_Error

14620 Fxs.Load Val(SampleID) ' + sysOptMicroOffset(0)
14630 If Fxs("Rota") Is Nothing And Fxs("Adeno") Is Nothing Then Exit Sub

14640 With frmMicroReport.rtb
14650     .SelText = vbCrLf
14660     .SelText = vbCrLf

14670     .SelBold = False
14680     .SelFontSize = 10

14690     If Not Fxs("Rota") Is Nothing Then
14700         .SelBold = False
14710         .SelText = Space$(10) & " Rota Virus : "
14720         .SelBold = True
14730         If UCase$(Fxs("Rota").Result) = "N" Then
14740             .SelText = "Negative"
14750         ElseIf UCase$(Fxs("Rota").Result) = "P" Then
14760             .SelText = "Positive"
14770         End If
14780         .SelText = vbCrLf
14790     End If

14800     If Not Fxs("Adeno") Is Nothing Then
14810         .SelBold = False
14820         .SelText = Space$(10) & "Adeno Virus : "
14830         .SelBold = True
14840         If UCase$(Fxs("Adeno").Result) = "N" Then
14850             .SelText = "Negative"
14860         ElseIf UCase$(Fxs("Adeno").Result) = "P" Then
14870             .SelText = "Positive"
14880         End If
14890         .SelText = vbCrLf
14900     End If

14910     .SelBold = False
14920 End With

      'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "ROTAADENO", True, True

14930 Exit Sub

RTFPrintMicroRotaAdeno_Error:

      Dim strES As String
      Dim intEL As Integer

14940 intEL = Erl
14950 strES = Err.Description
14960 LogError "modRTFMicro", "RTFPrintMicroRotaAdeno", intEL, strES

End Sub


Private Sub RTFPrintNegativeResults(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim CultPrinted As Boolean

14970 On Error GoTo RTFPrintNegativeResults_Error

      '20    sql = "Select * from Isolates where " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
14980 sql = "Select * from Isolates where " & _
            "SampleID = '" & Val(SampleID) & "'"
14990 Set tb = New Recordset
15000 RecOpenServer 0, tb, sql
15010 If Not tb.EOF Then

15020     With frmMicroReport.rtb

15030         .SelText = vbCrLf
15040         .SelBold = True
15050         .SelFontSize = 10

15060         CultPrinted = False
15070         Do While Not tb.EOF
15080             If Not CultPrinted Then
15090                 .SelText = Space$(10) & "CULTURE: " & tb!OrganismName & " " & tb!Qualifier & ""
15100                 CultPrinted = True
15110             Else
15120                 .SelText = Space$(10) & "         " & tb!OrganismName & " " & tb!Qualifier & ""
15130             End If
15140             .SelText = vbCrLf
15150             tb.MoveNext
15160         Loop
15170     End With
15180 End If

      '230   UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "CANDS", True, True

15190 Exit Sub

RTFPrintNegativeResults_Error:

      Dim strES As String
      Dim intEL As Integer

15200 intEL = Erl
15210 strES = Err.Description
15220 LogError "modRTFMicro", "RTFPrintNegativeResults", intEL, strES, sql

End Sub

Public Sub RTFPrintMicroscopy(ByVal SampleID As String)

      Dim URS As New UrineResults
      Dim UR As UrineResult
      Dim Res As String
      Dim n As Integer
      Dim Found As Boolean

15230 On Error GoTo RTFPrintMicroscopy_Error

15240 URS.Load Val(SampleID) ' + sysOptMicroOffset(0)
15250 If URS.Count = 0 Then
15260     Exit Sub
15270 End If

15280 For Each UR In URS
15290     If UR.TestName = "Bacteria" Or _
             UR.TestName = "Crystals" Or _
             UR.TestName = "RCC" Or _
             UR.TestName = "WCC" Or _
             UR.TestName = "Casts" Or _
             UR.TestName = "Misc0" Or _
             UR.TestName = "Misc1" Or _
             UR.TestName = "Misc2" Then
15300         Found = True
15310         Exit For
15320     End If
15330 Next
15340 If Not Found Then
15350     Exit Sub
15360 End If

15370 With frmMicroReport.rtb
15380     .SelBold = False
15390     .SelText = "Microscopy: Bacteria:"
15400     .SelBold = True
15410     Res = ""
15420     Set UR = URS.Item("Bacteria")
15430     If Not UR Is Nothing Then
15440         Res = UR.Result
15450     End If
15460     .SelText = Left$(Res & Space$(14), 14)
15470     .SelBold = False
15480     .SelText = "Crystals:"
15490     .SelBold = True
15500     Res = ""
15510     Set UR = URS.Item("Crystals")
15520     If Not UR Is Nothing Then
15530         Res = UR.Result
15540     End If
15550     If Res = "" Then
15560         .SelText = "Nil"
15570     Else
15580         .SelText = Res
15590     End If
15600     .SelText = vbCrLf

15610     .SelBold = False
15620     .SelText = "                 WCC:"
15630     .SelBold = True
15640     Res = ""
15650     Set UR = URS.Item("WCC")
15660     If Not UR Is Nothing Then
15670         Res = UR.Result
15680     End If
15690     .SelText = Left$(Res & " /cmm" & Space$(14), 14)
15700     .SelBold = False
15710     Res = ""
15720     Set UR = URS.Item("Casts")
15730     If Not UR Is Nothing Then
15740         Res = UR.Result
15750     End If

15760     .SelText = "   Casts:"
15770     .SelBold = True
15780     If Res = "" Then
15790         .SelText = "Nil"
15800     Else
15810         .SelText = Res
15820     End If
15830     .SelText = vbCrLf

15840     .SelBold = False
15850     .SelText = "                 RCC:"
15860     .SelBold = True
15870     Res = ""
15880     Set UR = URS.Item("RCC")
15890     If Not UR Is Nothing Then
15900         Res = UR.Result
15910     End If
15920     .SelText = Left$(Res & Space$(14), 14)
15930     .SelBold = False
15940     .SelText = "    Misc:"
15950     .SelBold = True

15960     Res = ""
15970     For n = 0 To 2
15980         Set UR = URS.Item("Misc" & Format$(n))
15990         If Not UR Is Nothing Then
16000             Res = Res & UR.Result & " "
16010         End If
16020     Next

16030     If Trim$(Res) = "" Then
16040         .SelText = "Nil"
16050     Else
16060         .SelText = Res
16070     End If
16080     .SelText = vbCrLf
16090     .SelBold = False
16100 End With

      'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "URINE", True, True

16110 Exit Sub

RTFPrintMicroscopy_Error:

      Dim strES As String
      Dim intEL As Integer

16120 intEL = Erl
16130 strES = Err.Description
16140 LogError "modRTFMicro", "RTFPrintMicroscopy", intEL, strES

End Sub

Private Sub RTFPrintMicroRedSub(ByVal SampleID As String)

      Dim GXs As New GenericResults
      Dim Gx As GenericResult

16150 On Error GoTo RTFPrintMicroRedSub_Error

16160 GXs.Load Val(SampleID) ' + sysOptMicroOffset(0)
16170 If GXs.Count = 0 Then
16180     Exit Sub
16190 End If
16200 Set Gx = GXs("RedSub")
16210 If Gx Is Nothing Then
16220     Exit Sub
16230 End If

16240 With frmMicroReport.rtb

16250     .SelText = vbCrLf
16260     .SelBold = False
16270     .SelFontSize = 10

16280     .SelText = Space$(5) & "Reducing Substances : "
16290     .SelBold = True
16300     .SelText = Gx.Result
16310     .SelText = vbCrLf
16320     .SelBold = False
16330     .SelFontSize = 10
16340 End With

      'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "REDSUB", True, True

16350 Exit Sub

RTFPrintMicroRedSub_Error:

      Dim strES As String
      Dim intEL As Integer

16360 intEL = Erl
16370 strES = Err.Description
16380 LogError "modRTFMicro", "RTFPrintMicroRedSub", intEL, strES

End Sub

Private Sub RTFPrintMicroCSF(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim s As String
      Dim Found As Boolean

16390 On Error GoTo RTFPrintMicroCSF_Error

      '20    sql = "SELECT * FROM CSFResults WHERE " & _
      '            "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
16400 sql = "SELECT * FROM CSFResults WHERE " & _
            "SampleID = '" & Val(SampleID) & "'"
16410 Set tb = New Recordset
16420 RecOpenServer 0, tb, sql
16430 If tb.EOF Then Exit Sub
16440 Found = False
16450 If Trim$(tb!Gram & tb!WCCDiff0 & tb!WCCDiff1 & "") <> "" Then
16460     Found = True
16470 End If
16480 If Not Found Then
16490     For n = 0 To 2
16500         If Trim$(tb("Appearance" & Format$(n)) & _
                       tb("WCC" & Format$(n)) & _
                       tb("RCC" & Format$(n)) & "") <> "" Then
16510             Found = True
16520             Exit For
16530         End If
16540     Next
16550 End If
16560 If Not Found Then Exit Sub

16570 With frmMicroReport.rtb
16580     .SelFontSize = 10
16590     .SelText = vbCrLf
16600     .SelBold = True
16610     .SelText = "Appearance: "
16620     .SelBold = False
16630     .SelText = Space$(12) & "Sample 1: " & tb!Appearance0 & ""
16640     .SelText = vbCrLf
16650     .SelText = Space$(12) & "Sample 2: " & tb!Appearance1 & ""
16660     .SelText = vbCrLf
16670     .SelText = Space$(12) & "Sample 3: " & tb!Appearance2 & ""
16680     .SelText = vbCrLf

16690     .SelBold = True
16700     .SelText = "Gram Stain: "
16710     .SelBold = False
16720     .SelText = tb!Gram & ""
16730     .SelText = vbCrLf

16740     .SelText = "         "
16750     .SelUnderline = True
16760     .SelText = "Sample 1"
16770     .SelUnderline = False
16780     .SelText = "        "
16790     .SelUnderline = True
16800     .SelText = "Sample 2"
16810     .SelUnderline = False
16820     .SelText = "        "
16830     .SelUnderline = True
16840     .SelText = "Sample 3"
16850     .SelText = vbCrLf
16860     .SelUnderline = False

16870     .SelBold = True
16880     .SelText = "WCC/cmm    "
16890     .SelBold = False
16900     s = Left$(tb!WCC0 & Space(16), 16) & _
              Left$(tb!WCC1 & Space(16), 16) & _
              tb!WCC2 & ""
16910     .SelText = s
16920     .SelText = vbCrLf

16930     .SelBold = True
16940     .SelText = "RCC/cmm    "
16950     .SelBold = False

16960     s = Left$(tb!RCC0 & Space(16), 16) & _
              Left$(tb!RCC1 & Space(16), 16) & _
              tb!RCC2 & ""
16970     .SelText = s
16980     .SelText = vbCrLf

16990     .SelBold = True
17000     .SelText = "White Cell Differential: "
17010     .SelBold = False
17020     .SelText = tb!WCCDiff0 & " % Neutrophils " & tb!WCCDiff1 & "% Mononuclear Cells"
17030     .SelText = vbCrLf

          'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "FOB", True, True
17040 End With

17050 Exit Sub

RTFPrintMicroCSF_Error:

      Dim strES As String
      Dim intEL As Integer

17060 intEL = Erl
17070 strES = Err.Description
17080 LogError "modRTFMicro", "RTFPrintMicroCSF", intEL, strES, sql

End Sub


Public Sub RTFPrintSpecType(ByRef RP As ReportToPrint, _
                            ByVal CurrentPage As Integer, _
                            ByVal TotalPages As Integer)

      Dim SiteDetails As String
      Dim Site As String
      Dim SDS As New SiteDetails

17090 On Error GoTo RTFPrintSpecType_Error

17100 SDS.Load Val(RP.SampleID) ' + sysOptMicroOffset(0)
17110 If SDS.Count > 0 Then
17120     Site = SDS(1).Site
17130     SiteDetails = SDS(1).SiteDetails
17140 End If

17150 With frmMicroReport.rtb
17160     .SelColor = vbBlack
17170     .SelFontSize = 10
17180     .SelBold = False
17190     .SelText = "Specimen Type:"
17200     .SelBold = True
17210     .SelText = Site & " " & SiteDetails & " "

17220     .SelColor = vbRed
          '150       If TotalPages > 1 Then
          '160           .SelText = "Page " & CurrentPage & " of " & TotalPages & " "
          '170       End If
17230     If RP.ThisIsCopy Then
17240         .SelBold = False
17250         .SelText = "This is a COPY Report for Attention of "
17260         .SelBold = True
17270         .SelText = Trim$(RP.SendCopyTo)
17280     End If
17290     .SelText = vbCrLf

17300     .SelColor = vbBlack
17310     .SelBold = False
17320 End With

17330 Exit Sub

RTFPrintSpecType_Error:

      Dim strES As String
      Dim intEL As Integer

17340 intEL = Erl
17350 strES = Err.Description
17360 LogError "modRTFMicro", "RTFPrintSpecType", intEL, strES

End Sub


Private Sub RTFPrintMicroPregnancy(ByVal SampleID As String)

      Dim URS As New UrineResults
      Dim UR As UrineResult
      Dim Preg As String
      Dim HCG As String

17370 On Error GoTo RTFPrintMicroPregnancy_Error

17380 Preg = ""
17390 URS.Load Val(SampleID) ' + sysOptMicroOffset(0)
17400 If URS.Count = 0 Then Exit Sub
17410 Set UR = URS("Pregnancy")
17420 If Not UR Is Nothing Then
17430     Preg = UR.Result
17440 End If

17450 HCG = ""
17460 Set UR = URS("HCGLevel")
17470 If Not UR Is Nothing Then
17480     HCG = UR.Result
17490 End If

17500 If Preg <> "" Or HCG <> "" Then

17510     With frmMicroReport.rtb

17520         .SelText = vbCrLf

17530         .SelBold = False
17540         .SelFontSize = 10
17550         .SelText = Space$(10) & "Pregnancy Test: "
17560         .SelBold = True
17570         .SelText = Preg
17580         .SelText = vbCrLf
17590         .SelBold = False
17600         .SelText = Space$(10) & "    HCG Level : "
17610         .SelBold = True
17620         .SelText = HCG & " IU/L"
17630         .SelText = vbCrLf

17640         RTFPrintMicroComment RP.SampleID, "Pregnancy Comment"

17650         .SelBold = False

              'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "URINE", True, True

17660     End With

17670 End If

17680 Exit Sub

RTFPrintMicroPregnancy_Error:

      Dim strES As String
      Dim intEL As Integer

17690 intEL = Erl
17700 strES = Err.Description
17710 LogError "modRTFMicro", "RTFPrintMicroPregnancy", intEL, strES

End Sub


Private Sub RTFPrintMicroRSV(ByVal SampleID As String)

      Dim GXs As New GenericResults
      Dim Gx As GenericResult

17720 On Error GoTo RTFPrintMicroRSV_Error

17730 GXs.Load Val(SampleID) ' + sysOptMicroOffset(0)
17740 If GXs.Count = 0 Then
17750     Exit Sub
17760 End If
17770 Set Gx = GXs("RSV")
17780 If Gx Is Nothing Then
17790     Exit Sub
17800 End If

17810 With frmMicroReport.rtb
17820     .SelText = vbCrLf
17830     .SelBold = False
17840     .SelFontSize = 10

17850     .SelText = Space$(10) & "RSV : "
17860     .SelBold = True
17870     .SelText = Gx.Result
17880     .SelText = vbCrLf
17890     .SelBold = False
17900 End With

      'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "RSV", True, True

17910 Exit Sub

RTFPrintMicroRSV_Error:

      Dim strES As String
      Dim intEL As Integer

17920 intEL = Erl
17930 strES = Err.Description
17940 LogError "modRTFMicro", "RTFPrintMicroRSV", intEL, strES

End Sub


Private Sub RTFPrintMicroComment(ByVal SampleID As String, _
                                 ByVal Source As String)

      Dim OBs As Observations
      Dim OffSet As Variant
      Dim n As Integer
      Dim pSource As String

17950 On Error GoTo RTFPrintMicroComment_Error

17960 ReDim Comments(1 To MicroCommentLineCount) As String

17970 OffSet = 0 'sysOptMicroOffset(0)

17980 Select Case UCase$(Left$(Source, 1))
      Case "D": Source = "Demographic": pSource = "Demographics Comment:"
17990 Case "L": Source = "MicroCSAutoComment": pSource = "Auto Comment:"
18000 Case "C": Source = "MicroConsultant": pSource = "Consultant Comment:"
18010 Case "M": Source = "MicroCS": pSource = "Medical Scientist Comment:"
18020 Case "P": Source = "MicroGeneral": pSource = ""
18030 Case "I": Source = "Semen": pSource = "Infertility Comment:": OffSet = sysOptSemenOffset
18040 Case "P": Source = "Semen": pSource = "Post Vasectomy": OffSet = sysOptSemenOffset

18050 End Select

18060 Set OBs = New Observations
18070 Set OBs = OBs.Load(Val(SampleID + OffSet), Source)

18080 With frmMicroReport.rtb
18090     .SelFontName = "Courier New"
18100     .SelFontSize = 9
18110     .SelColor = vbBlack
18120     If Not OBs Is Nothing Then
        '+++ Junaid
      '190           FillCommentLines pSource & OBs(1).Comment, MicroCommentLineCount, Comments(), 90
18130       FillCommentLines pSource, MicroCommentLineCount, Comments(), 90
        '--- Junaid
18140         .SelBold = False
18150         .SelText = pSource
18160         .SelBold = False
18170         .SelText = Mid$(Comments(1), Len(pSource) + 1)
18180         .SelText = vbCrLf
18190         For n = 2 To MicroCommentLineCount
18200             If Trim$(Comments(n)) <> "" Then
18210                 .SelText = Comments(n)
18220                 .SelText = vbCrLf
18230                 .SelColor = vbBlack
18240             End If
18250         Next
18260     End If
18270 End With




      '
      '    Set OBs = New Observations
      '    Set OBs = OBs.Load(Val(SampleID + OffSet), "MicroCSAutoComment")
      '
      '    With frmMicroReport.rtb
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





18280 Exit Sub

RTFPrintMicroComment_Error:

      Dim strES As String
      Dim intEL As Integer

18290 intEL = Erl
18300 strES = Err.Description
18310 LogError "modRTFMicro", "RTFPrintMicroComment", intEL, strES

End Sub
Private Sub RTFPrintMicroOvaParasites(ByVal SampleID As String)

      Dim n As Integer
      Dim blnRejectFound As Boolean
      Dim blnHeadingPrinted As Boolean
      Dim Fx As FaecesResult
      Dim Fxs As New FaecesResults

18320 On Error GoTo RTFPrintMicroOvaParasites_Error

18330 Fxs.Load Val(SampleID) ' + sysOptMicroOffset(0)
18340 If Fxs.Count = 0 Then
18350     Exit Sub
18360 End If

18370 With frmMicroReport.rtb
18380     .SelFontSize = 10
18390     .SelText = vbCrLf

18400     Set Fx = Fxs.Item("AUS")
18410     If Not Fx Is Nothing Then
18420         .SelBold = False
18430         .SelText = Space$(10) & "Cryptosporidium : "
18440         .SelBold = True
18450         If UCase$(Fx.Result) = "N" Then
18460             .SelText = "Negative"
18470         ElseIf UCase$(Fx.Result) = "P" Then
18480             .SelText = "Positive"
18490         End If
18500         .SelText = vbCrLf
18510         .SelBold = False
18520     End If

18530     blnRejectFound = False
18540     For n = 0 To 2
18550         Set Fx = Fxs.Item("OP" & Format$(n))
18560         If Not Fx Is Nothing Then
18570             If InStr(UCase$(Fx.Result), "REJECTED") <> 0 Then
18580                 blnRejectFound = True
18590                 Exit For
18600             End If
18610         End If
18620     Next
18630     If blnRejectFound Then
18640         .SelText = Space$(10) & "Ova and Parasites : "
18650         .SelBold = True
18660         .SelText = "Sample Rejected"
18670         .SelText = vbCrLf
18680         .SelBold = False
18690         .SelFontSize = 10
18700         .SelText = "I wish to remind you that "
18710         .SelItalic = True
18720         .SelText = "Ova and Parasites"
18730         .SelItalic = False
18740         .SelText = " should be requested only when there is a"
18750         .SelText = vbCrLf
18760         .SelText = "high index of suspicion. The clinical details received "
18770         .SelText = "with this test request fail to meet the"
18780         .SelText = vbCrLf
18790         .SelText = "criteria for testing and as such has been deemed "
18800         .SelText = "unsuitable for analysis. Please refer to"
18810         .SelText = vbCrLf
18820         .SelText = "the following guidelines for requesting "
18830         .SelItalic = True
18840         .SelText = "Ova and Parasites."
18850         .SelText = vbCrLf
18860         .SelItalic = False

18870         .SelBold = True
18880         .SelUnderline = True
18890         .SelText = "Ova and Parasites"
18900         .SelText = vbCrLf
18910         .SelUnderline = False
18920         .SelBold = False

18930         .SelText = "Submit one stool sample if: Persistent diarrhoea > 7 days ;"
18940         .SelText = vbCrLf
18950         .SelText = Space$(25)
18960         .SelBold = True
18970         .SelUnderline = True
18980         .SelText = "or"
18990         .SelUnderline = False
19000         .SelBold = False
19010         .SelText = " Patient is immunocompromised;"
19020         .SelText = vbCrLf
19030         .SelText = Space$(25)
19040         .SelBold = True
19050         .SelUnderline = True
19060         .SelText = "or"
19070         .SelUnderline = False
19080         .SelBold = False

19090         .SelText = " Patient has visited a developing country"

19100     Else
19110         blnHeadingPrinted = False
19120         For n = 0 To 2
19130             Set Fx = Fxs.Item("OP" & Format(n))
19140             If Not Fx Is Nothing Then
19150                 .SelText = vbCrLf
19160                 .SelBold = False
19170                 If Not blnHeadingPrinted Then
19180                     .SelText = Space$(10) & "Ova and Parasites : "
19190                     blnHeadingPrinted = True
19200                 Else
19210                     .SelText = Space$(10) & "                    "
19220                 End If
19230                 .SelBold = True
19240                 .SelText = Trim$(Fx.Result)
19250                 .SelText = vbCrLf
19260                 .SelBold = False
19270             End If
19280         Next
19290         .SelText = vbCrLf
19300     End If

19310     .SelFontSize = 10
19320 End With

      'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "OP", True, True

19330 Exit Sub

RTFPrintMicroOvaParasites_Error:

      Dim strES As String
      Dim intEL As Integer

19340 intEL = Erl
19350 strES = Err.Description
19360 LogError "modRTFMicro", "RTFPrintMicroOvaParasites", intEL, strES

End Sub


Private Sub RTFPrintMicroClinDetails(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim CLD As String

19370 On Error GoTo RTFPrintMicroClinDetails_Error

19380 With frmMicroReport.rtb
19390     .SelFontSize = 10
19400     .SelText = "Clinical Details:"
      '50        sql = "Select ClDetails from Demographics where " & _
      '                "SampleID = '" & Val(SampleID) + sysOptMicroOffset(0) & "'"
19410     sql = "Select ClDetails from Demographics where " & _
                "SampleID = '" & Val(SampleID) & "'"
19420     Set tb = New Recordset
19430     RecOpenClient 0, tb, sql
19440     If Not tb.EOF Then
19450         CLD = tb!ClDetails & ""
19460         CLD = Replace(CLD, vbCr, " ")
19470         CLD = Replace(CLD, vbLf, " ")
19480         .SelText = CLD
19490     End If
19500     .SelText = vbCrLf
19510 End With

19520 Exit Sub

RTFPrintMicroClinDetails_Error:

      Dim strES As String
      Dim intEL As Integer

19530 intEL = Erl
19540 strES = Err.Description
19550 LogError "modRTFMicro", "RTFPrintMicroClinDetails", intEL, strES, sql

End Sub


Private Sub RTFPrintMicroOccultBlood(ByVal SampleID As String)

      Dim n As Integer
      Dim Fx As FaecesResult
      Dim Fxs As New FaecesResults

19560 On Error GoTo RTFPrintMicroOccultBlood_Error

19570 Fxs.Load Val(SampleID) ' + sysOptMicroOffset(0)
19580 If Fxs.Count = 0 Then
19590     Exit Sub
19600 End If

19610 With frmMicroReport.rtb
19620     .SelFontSize = 10
19630     .SelText = vbCrLf

19640     For n = 0 To 2
19650         Set Fx = Fxs.Item("OB" & Format$(n))
19660         If Not Fx Is Nothing Then
19670             If Trim$(Fx.Result) <> "" Then
19680                 .SelBold = False
19690                 .SelText = Space$(10) & "Occult Blood ( " & Format$(n + 1) & " ) : "
19700                 .SelBold = True
19710                 If UCase$(Fx.Result) = "N" Then
19720                     .SelText = "Negative"
19730                 ElseIf UCase$(Fx.Result) = "P" Then
19740                     .SelText = "Positive"
19750                 End If
19760                 .SelText = vbCrLf
19770             End If
19780         End If
19790     Next
19800     .SelBold = False
19810     .SelFontSize = 10
19820 End With

      'UpdatePrintValid Val(sampleid) + sysOptMicroOffset(0), "FOB", True, True

19830 Exit Sub

RTFPrintMicroOccultBlood_Error:

      Dim strES As String
      Dim intEL As Integer

19840 intEL = Erl
19850 strES = Err.Description
19860 LogError "modRTFMicro", "RTFPrintMicroOccultBlood", intEL, strES

End Sub
Private Sub RTFPrintMicroCurrentABs(ByVal SampleID As String)

      Dim s As String
      Dim CurABs As New CurrentAntibiotics
      Dim CurAB As CurrentAntibiotic

19870 On Error GoTo RTFPrintMicroCurrentABs_Error

19880 CurABs.Load Val(SampleID) ' + sysOptMicroOffset(0)

19890 s = ""
19900 For Each CurAB In CurABs
19910     s = CurAB.Antibiotic & " "
19920 Next
19930 With frmMicroReport.rtb
19940     .SelFontSize = 10
19950     .SelText = "Current Antibiotics:" & s
19960     .SelText = vbCrLf
19970     .SelFontSize = 4
19980     .SelAlignment = rtfCenter
19990     .SelText = String$(200, "-")
20000     .SelText = vbCrLf
20010     .SelFontSize = 10
20020     .SelAlignment = rtfLeft
20030 End With

20040 Exit Sub

RTFPrintMicroCurrentABs_Error:

      Dim strES As String
      Dim intEL As Integer

20050 intEL = Erl
20060 strES = Err.Description
20070 LogError "modRTFMicro", "RTFPrintMicroCurrentABs", intEL, strES

End Sub

