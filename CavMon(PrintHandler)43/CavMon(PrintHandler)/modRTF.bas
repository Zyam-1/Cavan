Attribute VB_Name = "modRTF"
Option Explicit

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type


Private Sub PrintFlag(ByVal Flag As String)

10    On Error GoTo PrintFlag_Error

20    With frmMain.rtb
30        .SelBold = True
40        .SelFontSize = 9
50        .SelText = Flag
60        .SelBold = False
70        .SelFontSize = 10
80    End With

90    Exit Sub

PrintFlag_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modRTF", "PrintFlag", intEL, strES

End Sub

Public Sub RTFPrintCreatinine()

      Dim tb As Recordset
      Dim tc As Recordset
      Dim sql As String
      Dim Sex As String
      Dim n As Integer
      Dim OBs As Observations
10    On Error GoTo RTFPrintCreatinine_Error

20    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim ReceivedDate As String
      Dim Dob As String
      Dim RunTime As String
      Dim SorU As String
      Dim BioComment As String
      Dim DemoComment As String


30    If RP.Department = "R" Then
40        SorU = "Urine"
50    Else
60        SorU = "Serum"
70    End If

80    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
90    Set tb = New Recordset
100   RecOpenClient 0, tb, sql

110   If tb.EOF Then
120       Exit Sub
130   End If

140   If IsDate(tb!Dob) Then
150       Dob = Format(tb!Dob, "dd/mmm/yyyy")
160   Else
170       Dob = ""
180   End If

190   sql = "Select * from Creatinine where " & _
            SorU & "Number = '" & RP.SampleID & "'"
200   Set tc = New Recordset
210   RecOpenServer 0, tc, sql

220   If Not SetPrinter("CHBIO") Then Exit Sub

230   RTFPrintHeading "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
                      tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

240   With frmMain.rtb
250       .SelFontSize = 10

260       .SelText = vbCrLf
270       .SelText = Space$(15) & "Creatinine Clearance Test"
280       .SelText = vbCrLf
290       .SelText = vbCrLf
300       .SelText = Space$(15) & Left$("Volume Collected:" & Space$(25), 25) & Format(tc!UrineVolume & "", "#####") & " mL"
310       .SelText = vbCrLf
320       .SelText = vbCrLf
330       .SelText = Space$(15) & Left$("Plasma Creatinine:" & Space$(25), 25) & Format(tc!SerumCreat & "", "#####.0") & " umol/L"
340       .SelText = vbCrLf
350       .SelText = vbCrLf
360       .SelText = Space$(15) & Left$("Urinary Creatinine:" & Space$(25), 25) & Format(tc!UrineCreat & "", "0.00") & " umol/L"
370       .SelText = vbCrLf
380       .SelText = vbCrLf
390       .SelText = Space$(15) & Left$("Clearance:" & Space$(25), 25)
400       If Val(tc!CCl & "") > 0 Then
410           If Val(tc!CCl) < 1000 Then
420               .SelText = Format(Val(tc!CCl & ""), "####") & " ml/min"
430           Else
440               .SelText = Format(Val(tc!CCl & "") / 1000, "####") & " ml/min"
450           End If
460       End If
470       .SelText = vbCrLf
480       .SelText = vbCrLf

490       If Trim$(tc!UrineProL & "") <> "" Then
500           .SelText = Space$(15) & Left$("Protein Concentration:" & Space$(25), 25) & Format(tc!UrineProL & "", "0.000") & " g/L"
510           .SelText = vbCrLf
520           .SelText = vbCrLf
530           .SelText = Space$(15) & Space$(25) & Format(tc!UrinePro24Hr & "", "#0.000") & " g/24Hr"
540           .SelText = vbCrLf
550       End If
560       .SelText = vbCrLf
570       .SelText = Space$(15) & Left$("Report Date:" & Space$(25), 25) & Format(tb!Rundate, "dd/mm/yyyy")
580       .SelText = vbCrLf

590       Set OBs = New Observations
600       Set OBs = OBs.Load(RP.SampleID, "Demographic")
610       If Not OBs Is Nothing Then
620           FillCommentLines OBs(1).Comment, 4, Comments(), 97
630           For n = 1 To 4
640               .SelColor = vbBlack
650               .SelText = Space$(15) & Comments(n) & vbCrLf
660           Next
670       End If

680       Set OBs = New Observations
690       Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
700       If Not OBs Is Nothing Then
710           FillCommentLines OBs(1).Comment, 4, Comments(), 97
720           For n = 1 To 4
730               .SelColor = vbBlack
740               .SelText = Space$(15) & Comments(n) & vbCrLf
750           Next
760       End If

770       If IsDate(tb!SampleDate) Then
780           SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy")
790       Else
800           SampleDate = ""
810       End If

820       If IsDate(tb!RecDate) Then
830           ReceivedDate = Format(tb!RecDate, "dd/mmm/yyyy")
840       Else
850           ReceivedDate = ""
860       End If

870       If IsDate(RunTime) Then
880           Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
890       Else
900           If IsDate(tb!Rundate) Then
910               Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
920           Else
930               Rundate = ""
940           End If
950       End If

960       RTFPrintFooter "Creat Clearance", RP.Initiator, SampleDate, Rundate, ReceivedDate

970       PrintAndStore "Biochemistry", "BCreat"

980   End With

990   Exit Sub

RTFPrintCreatinine_Error:

      Dim strES As String
      Dim intEL As Integer

1000  intEL = Erl
1010  strES = Err.Description
1020  LogError "modRTF", "RTFPrintCreatinine", intEL, strES, sql

End Sub
Private Sub PrintX10(ByVal SS As String)

      Dim SuperScriptOffset As Integer

10    On Error GoTo PrintX10_Error

20    SuperScriptOffset = frmMain.TextHeight("H") / 3

30    With frmMain.rtb
40        .SelText = " x10"
50        .SelCharOffset = SuperScriptOffset
60        .SelFontSize = 6
70        .SelText = Left$(SS & "  ", 2)
80        .SelCharOffset = 0
90        .SelFontSize = 10
100       .SelText = "/l" & Space$(2)
110   End With

120   Exit Sub

PrintX10_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "modRTF", "PrintX10", intEL, strES

End Sub

Public Sub RTFPrintHaem()

      Dim tb As Recordset
      Dim tbh As Recordset
      Dim n As Integer
      Dim Sex As String
      Dim fbc As Integer
      Dim TotalRetics As Long
      Dim Dob As String
      Dim Flag As String
      Dim sql As String
      Dim OBs As Observations

10    On Error GoTo RTFPrintHaem_Error

20    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim ReceivedDate As String

30    If Not SetPrinter("CHHAEM") Then Exit Sub

40    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql
70    If tb.EOF Then Exit Sub

80    If IsDate(tb!SampleDate) Then
90        SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
100   Else
110       SampleDate = ""
120   End If

130   If IsDate(tb!RecDate) Then
140       ReceivedDate = Format(tb!RecDate, "dd/mmm/yyyy hh:mm")
150   Else
160       ReceivedDate = ""
170   End If

180   Dob = tb!Dob & ""

190   Select Case Left$(UCase$(tb!Sex & ""), 1)
      Case "M": Sex = "M"
200   Case "F": Sex = "F"
210   Case Else: Sex = ""
220   End Select

230   RTFPrintHeading "Haematology", tb!PatName & "", Dob, tb!Chart & "", _
                      tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""


240   sql = "Select * from HaemResults where " & _
            "SampleID = '" & RP.SampleID & "'"
250   Set tbh = New Recordset
260   RecOpenClient 0, tbh, sql
270   If Not tbh.EOF Then

280       fbc = Trim$(tbh!wbc & "") <> ""

290       With frmMain.rtb
300           .SelFontName = "Courier New"
310           .SelFontSize = 10

320           If fbc Then
330               .SelText = Space$(10) & "WBC   "
340               Flag = InterpH(Val(tbh!wbc & ""), "WBC", Sex, Dob)
350               PrintValueOrXXX Flag, tbh!wbc & "", 5
360               PrintX10 "9"
370               PrintFlag Flag
380               .SelText = HaemNormalRange("WBC", Sex, Dob)

390               Flag = InterpH(Val(tbh!NeutP & ""), "NEUTP", Sex, Dob)
400               .SelText = Space$(5) & "Neut  "
410               PrintValueOrXXX Flag, tbh!NeutP & "", 4
420               .SelText = "% = "
430               Flag = InterpH(Val(tbh!neuta & ""), "NEUTA", Sex, Dob)
440               PrintValueOrXXX Flag, tbh!neuta & "", 5
450               PrintX10 "9"
460               PrintFlag Flag
470               .SelText = HaemNormalRange("NEUTA", Sex, Dob)
480               .SelText = vbCrLf
490               .SelFontSize = 10

500               Flag = InterpH(Val(tbh!LymP & ""), "LYMP", Sex, Dob)
510               .SelText = Space$(47) & "Lymph "
520               PrintValueOrXXX Flag, tbh!LymP & "", 4
530               .SelText = "% = "
540               Flag = InterpH(Val(tbh!lyma & ""), "LYMA", Sex, Dob)
550               PrintValueOrXXX Flag, tbh!lyma & "", 5
560               PrintX10 "9"
570               PrintFlag Flag
580               .SelText = HaemNormalRange("LYMA", Sex, Dob)
590               .SelText = vbCrLf
600               .SelFontSize = 10

610               .SelText = Space$(10) & "RBC   "
620               Flag = InterpH(Val(tbh!rbc & ""), "RBC", Sex, Dob)
630               PrintValueOrXXX Flag, tbh!rbc & "", 5
640               PrintX10 "12"
650               PrintFlag Flag
660               .SelText = HaemNormalRange("RBC", Sex, Dob)
670               Flag = InterpH(Val(tbh!MonoP & ""), "MONOP", Sex, Dob)
680               .SelText = Space$(5) & "Mono  "
690               PrintValueOrXXX Flag, tbh!MonoP & "", 4
700               .SelText = "% = "
710               Flag = InterpH(Val(tbh!monoa & ""), "MONOA", Sex, Dob)
720               PrintValueOrXXX Flag, tbh!monoa & "", 5
730               PrintX10 "9"
740               PrintFlag Flag
750               .SelText = HaemNormalRange("MONOA", Sex, Dob)
760               .SelText = vbCrLf
770               .SelFontSize = 10

780               Flag = InterpH(Val(tbh!eosP & ""), "EOSP", Sex, Dob)
790               .SelText = Space$(47) & "Eos   "
800               PrintValueOrXXX Flag, tbh!eosP & "", 4
810               .SelText = "% = "
820               Flag = InterpH(Val(tbh!eosa & ""), "EOSA", Sex, Dob)
830               PrintValueOrXXX Flag, tbh!eosa & "", 5
840               PrintX10 "9"
850               PrintFlag Flag
860               .SelText = HaemNormalRange("EOSA", Sex, Dob)
870               .SelText = vbCrLf
880               .SelFontSize = 10

890               .SelText = Space(10) & "Hgb   "
900               Flag = InterpH(Val(tbh!Hgb & ""), "Hgb", Sex, Dob)
910               PrintValueOrXXX Flag, tbh!Hgb & "", 5
920               .SelText = " g/dl    "
930               PrintFlag Flag
940               .SelText = HaemNormalRange("Hgb", Sex, Dob)
950               Flag = InterpH(Val(tbh!basP & ""), "BASP", Sex, Dob)
960               .SelText = Space(5) & "Bas   "
970               Flag = InterpH(Val(tbh!basP & ""), "BASP", Sex, Dob)
980               PrintValueOrXXX Flag, tbh!basP & "", 4
990               .SelText = "% = "
1000              Flag = InterpH(Val(tbh!basa & ""), "BASA", Sex, Dob)
1010              PrintValueOrXXX Flag, tbh!basa & "", 5
1020              PrintX10 "9"
1030              PrintFlag Flag
1040              .SelText = HaemNormalRange("BASA", Sex, Dob)
1050              .SelText = vbCrLf


1060              .SelText = Space$(47)
                  '-------------FARHAN----------------
1070              If Len(Trim(tbh!RetA)) > 0 Then
1080                  .SelText = "Ret   "
1090                  Flag = InterpH(Val(tbh!RetP & ""), "RetP", Sex, Dob)
1100                  PrintValueOrXXX Flag, tbh!RetP & "", 4
1110                  .SelText = "% = "
1120                  Flag = InterpH(Val(tbh!RetA & ""), "RetA", Sex, Dob)
1130                  PrintValueOrXXX Flag, tbh!RetA & "", 5
1140                  PrintX10 "9"
1150                  PrintFlag Flag
1160                  .SelText = HaemNormalRange("RetA", Sex, Dob)
1170              End If
                  '===========FARHAN==================
1180              .SelText = vbCrLf

                  '                .SelFontSize = 10
                  '                .SelText = " " & vbCrLf
1190              .SelFontSize = 10
1200              .SelColor = vbBlack
1210              .SelFontName = "Courier New"

1220              .SelText = Space$(10) & "Hct   "
1230              Flag = InterpH(Val(tbh!hct & ""), "Hct", Sex, Dob)
1240              PrintValueOrXXX Flag, tbh!hct & "", 5
1250              .SelText = " l/l     "
1260              PrintFlag Flag
1270              .SelText = HaemNormalRange("Hct", Sex, Dob)




1280              .SelText = vbCrLf
1290              .SelFontSize = 10

1300              .SelText = Space$(10) & "MCV   "
1310              Flag = InterpH(Val(tbh!mcv & ""), "MCV", Sex, Dob)
1320              PrintValueOrXXX Flag, tbh!mcv & "", 5
1330              .SelText = " fl      "
1340              PrintFlag Flag
1350              .SelText = HaemNormalRange("MCV", Sex, Dob)
1360              .SelText = vbCrLf
1370              .SelFontSize = 10

1380              .SelText = Space$(10) & "MCH   "
1390              Flag = InterpH(Val(tbh!mch & ""), "MCH", Sex, Dob)
1400              PrintValueOrXXX Flag, tbh!mch & "", 5
1410              .SelText = " pg      "
1420              PrintFlag Flag
1430              .SelText = HaemNormalRange("MCH", Sex, Dob)
1440              If Trim$(tbh!monospot & "") <> "" Then
1450                  .SelText = Space$(5) & "Infectious Mono Screen "
1460                  If tbh!monospot = "N" Then
1470                      .SelText = "Negative."
1480                  ElseIf tbh!monospot = "P" Then
1490                      .SelText = "Positive."
1500                  Else
1510                      .SelText = tbh!monospot & ""
1520                  End If
1530              End If
1540              .SelText = vbCrLf
1550              .SelFontSize = 10

1560              .SelText = Space$(10) & "MCHC  "
1570              Flag = InterpH(Val(tbh!mchc & ""), "MCHC", Sex, Dob)
1580              PrintValueOrXXX Flag, tbh!mchc & "", 5
1590              .SelText = " g/dl    "
1600              PrintFlag Flag
1610              .SelText = HaemNormalRange("MCHC", Sex, Dob)
1620              If Trim$(tbh!esr & "") <> "" Then
1630                  .SelText = Space(5) & "ESR    "
1640                  Flag = InterpH(Val(tbh!esr & ""), "ESR", Sex, Dob)
1650                  PrintValueOrXXX Flag, tbh!esr & "", 3
1660                  .SelText = " mm/hr "
1670                  PrintFlag Flag
1680                  .SelText = HaemNormalRange("ESR", Sex, Dob)
1690              End If
                  '1590            .SelText = vbCrLf
                  '1600            .SelFontSize = 10

                  '1610            .SelText = Space$(10) & "RDW   "
                  '1620            Flag = InterpH(Val(tbh!rdwcv & ""), "RDWCV", Sex, Dob)
                  '1630            PrintValueOrXXX Flag, tbh!rdwcv & "", 5
                  '1640            .SelText = " %       "
                  '1650            PrintFlag Flag
                  '1660            .SelText = HaemNormalRange("RDWCV", Sex, Dob)
                  '1670            If Trim$(tbh!retp & "") <> "" Then
                  '1680                .SelText = Space$(5) & "Retics   "
                  '1690                Flag = InterpH(TotalRetics, "RET", Sex, Dob)
                  '1700                PrintValueOrXXX Flag, tbh!retp & "", 4
                  '1710                .SelText = " %"
                  '1720            End If
1700              .SelText = vbCrLf
1710              .SelFontSize = 10

1720              .SelText = Space$(10) & "Plt   "
1730              Flag = InterpH(Val(tbh!Plt & ""), "Plt", Sex, Dob)
1740              PrintValueOrXXX Flag, tbh!Plt & "", 5
1750              PrintX10 "9"
1760              PrintFlag Flag
1770              .SelText = HaemNormalRange("Plt", Sex, Dob)
1780              If Trim$(tbh!Sickledex & "") <> "" Then
1790                  .SelText = Space$(5) & "Sickledex test for HbS = " & tbh!Sickledex
1800              End If
1810              .SelText = vbCrLf
1820              .SelFontSize = 10

1830              If Trim$(tbh!Malaria & "") <> "" Then
1840                  .SelText = Space$(47) & "Malaria Screening Kit = " & tbh!Malaria
1850                  .SelText = vbCrLf
1860                  .SelFontSize = 10
1870                  .SelText = vbCrLf
1880                  .SelFontSize = 10
1890              End If





1900          End If
1910      End With
1920  End If

1930  With frmMain.rtb
1940      Set OBs = New Observations
1950      Set OBs = OBs.Load(RP.SampleID, "Haematology")
1960      If Not OBs Is Nothing Then
1970          FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
1980          For n = 1 To 4
1990              If Trim$(Comments(n)) <> "" Then
2000                  .SelColor = vbBlack
2010                  .SelText = Comments(n)
2020                  .SelText = vbCrLf
2030                  .SelFontSize = 10
2040              End If
2050          Next
2060      End If

2070      Set OBs = New Observations
2080      Set OBs = OBs.Load(RP.SampleID, "Demographic")
2090      If Not OBs Is Nothing Then
2100          FillCommentLines OBs.Item(1).Comment, 2, Comments(), 97
2110          For n = 1 To 2
2120              If Trim$(Comments(n)) <> "" Then
2130                  .SelColor = vbBlack
2140                  .SelText = Comments(n)
2150                  .SelText = vbCrLf
2160                  .SelFontSize = 10
2170              End If
2180          Next
2190      End If

2200      Set OBs = New Observations
2210      Set OBs = OBs.Load(RP.SampleID, "Film")
2220      If Not OBs Is Nothing Then
2230          FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
2240          For n = 1 To 4
2250              If Trim$(Comments(n)) <> "" Then
2260                  .SelColor = vbBlack
2270                  .SelText = Comments(n)
2280                  .SelText = vbCrLf
2290                  .SelFontSize = 10
2300              End If
2310          Next
2320      End If

2330      RTFPrintNoSexDoB Sex, Dob
2340  End With

2350  If Not IsNull(tbh!ValidateTime) Then
2360      If IsDate(tbh!ValidateTime) Then
2370          Rundate = Format(tbh!ValidateTime, "dd/mmm/yyyy hh:mm:ss")
2380      Else
2390          Rundate = ""
2400      End If
2410  End If

2420  RTFPrintFooter "Haematology", RP.Initiator, SampleDate, Rundate, ReceivedDate


2430  PrintAndStore "Haematology", "H"

2440  sql = "Update HaemResults " & _
            "set Printed = 1, Valid = 1 " & _
            "where SampleID = '" & RP.SampleID & "'"
2450  Cnxn(0).Execute sql

2460  Exit Sub

RTFPrintHaem_Error:

      Dim strES As String
      Dim intEL As Integer

2470  intEL = Erl
2480  strES = Err.Description
2490  LogError "modRTF", "RTFPrintHaem", intEL, strES, sql

End Sub

Public Sub RTFPrintEGFR(ByVal SampleID As String)

      Dim tb As Recordset
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
      Dim strLow As String * 4
      Dim strHigh As String * 4
      Dim BRs As New BIEResults
      Dim BR As BIEResult
      Dim TestCount As Integer
      Dim SampleType As String
      Dim ResultsPresent As Boolean
      Dim SampleDate As String
      Dim ReceivedDate As String
      Dim Rundate As String
      Dim Dob As String
      Dim RunTime As String
      Dim Fasting As String
      Dim udtPrintLine(0 To 35) As PrintLine
      Dim strFormat As String
      Dim Lipaemic As Integer
      Dim Icteric As Integer
      Dim Haemolysed As Integer
      Dim EGFRComment() As String
      Dim X As Integer
      Dim CodeForCreatinine As String
      Dim CodeForEGFR As String
10    On Error GoTo RTFPrintEGFR_Error

20    ReDim Comments(1 To 4) As String

30    ReDim lp(0 To 35) As String

40    For n = 0 To 35
50        udtPrintLine(n).Analyte = ""
60        udtPrintLine(n).Result = ""
70        udtPrintLine(n).Flag = ""
80        udtPrintLine(n).Units = ""
90        udtPrintLine(n).NormalRange = ""
100       udtPrintLine(n).Fasting = ""
110   Next

120   sql = "Select * from Demographics where " & _
            "SampleID = '" & SampleID & "'"
130   Set tb = New Recordset
140   RecOpenClient 0, tb, sql

150   If tb.EOF Then
160       Exit Sub
170   End If

180   If Not IsNull(tb!Fasting) Then
190       Fasting = tb!Fasting
200   Else
210       Fasting = False
220   End If

230   If IsDate(tb!Dob) Then
240       Dob = Format(tb!Dob, "dd/mmm/yyyy")
250   Else
260       Dob = ""
270   End If

280   CodeForCreatinine = UCase$(GetOptionSetting("BioCodeForCreatinine", "CreJC"))
290   CodeForEGFR = UCase$(GetOptionSetting("BioCodeForEGFR", "5555"))

300   ResultsPresent = False
310   Set BRs = BRs.Load("Bio", RP.SampleID, "Results", gDONTCARE, gDONTCARE)
320   If Not BRs Is Nothing Then
330       For X = BRs.Count To 1 Step -1
340           If UCase$(BRs(X).Code) <> CodeForCreatinine _
                 And UCase$(BRs(X).Code) <> CodeForEGFR Then
350               BRs.RemoveItem X
360           End If
370       Next

380       TestCount = BRs.Count
390       If TestCount <> 0 Then
400           ResultsPresent = True
410           SampleType = BRs(1).SampleType
420           If Trim$(SampleType) = "" Then SampleType = "S"
430       End If
440   End If

450   If Not ResultsPresent Then Exit Sub

460   If Not SetPrinter("CHBIO") Then Exit Sub

470   lpc = 0
480   LoadLIH RP.SampleID, Lipaemic, Icteric, Haemolysed
490   For Each BR In BRs
500       RunTime = BR.RunTime

510       If LIHEffects(BR.Code, Lipaemic, Icteric, Haemolysed) Then
520           v = "*****"
530       Else
540           v = BR.Result
550       End If

560       High = Val(BR.High)
570       Low = Val(BR.Low)

580       If Low < 10 Then
590           strLow = Format(Low, "0.00")
600       ElseIf Low < 100 Then
610           strLow = Format(Low, "##.0")
620       Else
630           strLow = Format(Low, " ###")
640       End If
650       If High < 10 Then
660           strHigh = Format(High, "0.00")
670       ElseIf High < 100 Then
680           strHigh = Format(High, "##.0")
690       Else
700           strHigh = Format(High, "### ")
710       End If

720       If IsNumeric(v) Then
730           If Val(v) > BR.PlausibleHigh Then
740               udtPrintLine(lpc).Flag = " X "
750               lp(lpc) = "  "
760               Flag = " X"
770           ElseIf Val(v) < BR.PlausibleLow Then
780               udtPrintLine(lpc).Flag = " X "
790               lp(lpc) = "  "
800               Flag = " X"
810           ElseIf Val(v) > BR.FlagHigh Then
820               udtPrintLine(lpc).Flag = " H "
830               lp(lpc) = "  "
840               Flag = " H"
850           ElseIf Val(v) < BR.FlagLow Then
860               udtPrintLine(lpc).Flag = " L "
870               lp(lpc) = "  "
880               Flag = " L"
890           Else
900               udtPrintLine(lpc).Flag = "   "
910               lp(lpc) = "  "
920               Flag = "  "
930           End If
940       Else
950           udtPrintLine(lpc).Flag = "   "
960           lp(lpc) = "  "
970           Flag = "  "
980       End If
990       lp(lpc) = lp(lpc) & "  "  'was "    "
1000      lp(lpc) = lp(lpc) & Left$(BR.LongName & Space(20), 20)  '20
1010      udtPrintLine(lpc).Analyte = Left$(BR.LongName & Space(16), 16)  '16

1020      If IsNumeric(v) Then
1030          Select Case BR.Printformat
              Case 0: strFormat = "#####0"
1040          Case 1: strFormat = "###0.0"
1050          Case 2: strFormat = "##0.00"
1060          Case 3: strFormat = "#0.000"
1070          End Select
1080          lp(lpc) = lp(lpc) & " " & Right$(Space(8) & Format(v, strFormat), 8)    'was 6
1090          udtPrintLine(lpc).Result = Format(v, strFormat)
1100      Else
1110          lp(lpc) = lp(lpc) & " " & Right$(Space(8) & v, 8)    'was 6
1120          udtPrintLine(lpc).Result = v
1130      End If
          'Else
          '  lp(lpc) = lp(lpc) & " XXXXXXX"
          'End If
1140      lp(lpc) = lp(lpc) & Flag & " "

1150      sql = "Select * from Lists where " & _
                "ListType = 'UN' and Code = '" & BR.Units & "'"
1160      Set tbUN = Cnxn(0).Execute(sql)
1170      If Not tbUN.EOF Then
1180          cUnits = Left$(tbUN!Text & Space(6), 6)
1190      Else
1200          cUnits = Left$(BR.Units & Space(6), 6)
1210      End If
1220      udtPrintLine(lpc).Units = cUnits
1230      lp(lpc) = lp(lpc) & cUnits

1240      If BR.PrintRefRange Then
1250          lp(lpc) = lp(lpc) & "   ("
1260          lp(lpc) = lp(lpc) & strLow & "-"
1270          lp(lpc) = lp(lpc) & strHigh & ")"
1280          udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
1290      Else
1300          lp(lpc) = lp(lpc) & "              "
1310          udtPrintLine(lpc).NormalRange = "           "
1320      End If

1330      udtPrintLine(lpc).Fasting = ""
1340      If Not IsNull(tb!Fasting) Then
1350          If tb!Fasting Then
1360              udtPrintLine(lpc).Fasting = "(Fasting)"
1370              lp(lpc) = lp(lpc) & "(Fasting)"
1380          End If
1390      End If

1400      If Flag <> " X" Then
1410          LogBioAsPrinted RP.SampleID, BR.Code
1420      End If

1430      lpc = lpc + 1

1440      sql = "UPDATE BioResults SET Printed = '1', Valid = '1' " & _
                "WHERE SampleID = '" & RP.SampleID & "' " & _
                "AND Code = '" & BR.Code & "'"
1450      Cnxn(0).Execute sql

1460  Next

1470  RTFPrintHeading "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
                      tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

1480  Sex = tb!Sex & ""

1490  With frmMain.rtb

1500      .SelFontSize = 10

1510      For n = 0 To 19
1520          If Trim$(lp(n)) <> "" Then
1530              If InStr(lp(n), " L ") Or InStr(lp(n), " H ") Then
1540                  .SelText = Space$(18) & Left$(lp(n), 33)
1550                  .SelBold = True
1560                  .SelFontSize = 9
1570                  .SelText = Mid$(lp(n), 34, 3)
1580                  .SelBold = False
1590                  .SelFontSize = 10
1600                  .SelText = Mid$(lp(n), 37)
1610                  .SelText = vbCrLf
1620              Else
1630                  .SelText = Space$(18) & lp(n)
1640                  .SelText = vbCrLf
1650              End If
1660          End If
1670      Next

1680      .SelText = vbCrLf
1690      .SelText = vbCrLf

1700      If GetEGFRComment(SampleID, EGFRComment) Then
1710          For n = 0 To UBound(EGFRComment)
1720              .SelFontSize = 9
1730              .SelColor = vbBlack
1740              .SelText = EGFRComment(n) & vbCrLf
1750          Next
1760      End If

          Dim OBs As Observations

1770      Set OBs = New Observations
1780      Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
1790      If Not OBs Is Nothing Then
1800          FillCommentLines OBs(1).Comment, 4, Comments(), 97
1810          For n = 1 To 4
1820              .SelFontSize = 9
1830              .SelColor = vbBlack
1840              .SelText = Comments(n) & vbCrLf
1850          Next
1860      End If

1870      Set OBs = New Observations
1880      Set OBs = OBs.Load(RP.SampleID, "Demographic")
1890      If Not OBs Is Nothing Then
1900          FillCommentLines OBs(1).Comment, 4, Comments(), 97
1910          For n = 1 To 4
1920              .SelFontSize = 9
1930              .SelColor = vbBlack
1940              .SelText = Comments(n) & vbCrLf
1950          Next
1960      End If

1970      .SelFontSize = 10

1980      RTFPrintNoSexDoB Sex, Dob

1990      If IsDate(tb!SampleDate) Then
2000          SampleDate = Format(tb!SampleDate, "dd/MMM/yyyy HH:mm")
2010      Else
2020          SampleDate = ""
2030      End If
2040      If IsDate(tb!RecDate) Then
2050          ReceivedDate = Format(tb!RecDate, "dd/MMM/yyyy HH:mm")
2060      Else
2070          ReceivedDate = ""
2080      End If
2090      If IsDate(RunTime) Then
2100          Rundate = Format(RunTime, "dd/MMM/yyyy HH:mm")
2110      Else
2120          If IsDate(tb!Rundate) Then
2130              Rundate = Format(tb!Rundate, "dd/MMM/yyyy")
2140          Else
2150              Rundate = ""
2160          End If
2170      End If

2180      RTFPrintFooter "Biochemistry", RP.Initiator, SampleDate, Rundate, "eGFR"

2190      PrintAndStore "Biochemistry", "Begfr"

2200  End With

2210  Exit Sub

RTFPrintEGFR_Error:

      Dim strES As String
      Dim intEL As Integer

2220  intEL = Erl
2230  strES = Err.Description
2240  LogError "modRTF", "RTFPrintEGFR", intEL, strES, sql

End Sub

Public Sub RTFPrintCoag()

      Dim tb As Recordset
      Dim n As Integer
      Dim Sex As String
      Dim Dob As String
      Dim sql As String
      Dim OBs As Observations
      Dim AutoComment As String

10    On Error GoTo RTFPrintCoag_Error

20    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim ReceivedDate As String
      Dim strNormalRange As String
      Dim DaysOld As Long
      Dim CR As CoagResult
      Dim CRs As CoagResults
      Dim Flag As String

30    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If tb.EOF Then Exit Sub

70    If IsDate(tb!SampleDate) Then
80        SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
90    Else
100       SampleDate = ""
110   End If

120   If IsDate(tb!RecDate) Then
130       ReceivedDate = Format(tb!RecDate, "dd/mmm/yyyy hh:mm")
140   Else
150       ReceivedDate = ""
160   End If

170   Dob = tb!Dob & ""
180   If IsDate(Dob) Then
190       DaysOld = DateDiff("d", Dob, Now)
200   End If

210   Select Case Left$(UCase$(tb!Sex & ""), 1)
      Case "M": Sex = "M"
220   Case "F": Sex = "F"
230   Case Else: Sex = ""
240   End Select

250   If Not SetPrinter("CHCOAG") Then Exit Sub

260   RTFPrintHeading "Coagulation", tb!PatName & "", Dob, tb!Chart & "", _
                      tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

270   With frmMain.rtb
280       .SelFontName = "Courier New"
290       .SelFontSize = 10
300       .SelText = vbCrLf
310       .SelFontSize = 10

320       Set CRs = New CoagResults
330       Set CRs = CRs.Load(RP.SampleID, gDONTCARE, "Results")
340       If Not CRs Is Nothing Then
350           For Each CR In CRs
360               If Not IsDate(Rundate) Then
370                   Rundate = CR.RunTime
380               End If
390               If CR.Printable Then
400                   .SelColor = vbBlack
410                   .SelText = Space$(18) & Left$(CR.TestName & Space$(11), 11)
420                   .SelText = Left$(CR.Result & Space$(6), 6)
430                   .SelText = Left$(CR.Units & Space$(10), 10)
                      '-----------------------------
440                   Flag = " "
450                   If CR.Low = 0 And (CR.High = 999 Or CR.High = 0 Or CR.High = 9999) Then
460                   Else
470                       If IsNumeric(CR.Result) Then
480                           If Val(CR.Result) <= Val(CR.Low) Then
490                               Flag = "L"
500                           ElseIf Val(CR.Result) >= Val(CR.High) Then
510                               Flag = "H"
520                           End If
530                       End If
540                   End If
550                   .SelBold = True
560                   .SelText = Left$(Flag & Space$(4), 4)
570                   .SelBold = False
                      '=============================
580                   If CR.PrintRefRange Then
590                       If Val(CR.Low) = 0 And Val(CR.High) > 0 Then
600                           .SelText = "( <" & CR.High & ")"
610                       ElseIf (Val(CR.High) = 999 Or Val(CR.High) = 9999) And Val(CR.Low) > 0 Then
620                           .SelText = "( >" & CR.Low & ")"
630                       Else
640                           .SelText = "(" & CR.Low & "-" & CR.High & ")"
650                       End If
660                   End If
670                   .SelText = vbCrLf
680                   .SelFontSize = 10
690                   AutoComment = CheckAutoComments(RP.SampleID, CR.TestName, 3)
700                   If Trim$(AutoComment) <> "" Then
710                       .SelText = Space$(10) & "Comment:" & AutoComment
720                       .SelText = vbCrLf
730                       .SelFontSize = 10
740                   End If
750               End If
760           Next
770           sql = "UPDATE CoagResults " & _
                    "SET Printed = 1 " & _
                    "WHERE SampleID = '" & RP.SampleID & "'"
780           Cnxn(0).Execute sql
790       End If

800       .SelText = vbCrLf
810       .SelFontSize = 10

820       Set OBs = New Observations
830       Set OBs = OBs.Load(RP.SampleID, "Demographic")
840       If Not OBs Is Nothing Then
850           FillCommentLines OBs(1).Comment, 2, Comments(), 97
860           For n = 1 To 2
870               If Trim$(Comments(n)) <> "" Then
880                   .SelColor = vbBlack
890                   .SelText = Comments(n)
900                   .SelText = vbCrLf
910                   .SelFontSize = 10
920               End If
930           Next
940       End If

950       Set OBs = New Observations
960       Set OBs = OBs.Load(RP.SampleID, "Coagulation")
970       If Not OBs Is Nothing Then
980           FillCommentLines OBs(1).Comment, 2, Comments(), 97
990           For n = 1 To 2
1000              If Trim$(Comments(n)) <> "" Then
1010                  .SelColor = vbBlack
1020                  .SelText = Comments(n)
1030                  .SelText = vbCrLf
1040                  .SelFontSize = 10
1050              End If
1060          Next
1070      End If

          '570   sql = "UPDATE CoagResults " & _
           '            "SET Printed = 1 " & _
           '            "WHERE SampleID = '" & RP.SampleID & "' " & _
           '            "AND Code NOT IN " & _
           '            "( SELECT DISTINCT(Code) FROM CoagTestDefinitions D JOIN PrintInhibit P " & _
           '            "  ON D.TestName = P.Parameter " & _
           '            "  WHERE SampleID = '" & RP.SampleID & "' " & _
           '            "  AND Discipline = 'Coa')"
          '580   Cnxn(0).Execute sql

1080      .SelText = vbCrLf
1090      .SelFontSize = 10

1100      RTFPrintNoSexDoB Sex, Dob

1110      Rundate = Format(Rundate, "dd/MMM/yyyy HH:mm")

1120      RTFPrintFooter "Coagulation", RP.Initiator, SampleDate, Rundate, ReceivedDate

1130      PrintAndStore "Coagulation", "C"

1140  End With

1150  Exit Sub

RTFPrintCoag_Error:

      Dim strES As String
      Dim intEL As Integer

1160  intEL = Erl
1170  strES = Err.Description
1180  LogError "modRTF", "RTFPrintCoag", intEL, strES, sql

End Sub
Public Sub RTFPrintHaemSpecific(ByVal MonoSpotOrESR As String)

    Dim tb As Recordset
    Dim tbh As Recordset
    Dim sql As String
    Dim n As Integer
    Dim Dob As String
    Dim OB As Observation
    Dim OBs As Observations
    Dim RunDateTime As String
    Dim ReceivedDate As String
    Dim Operator As String
    Dim HaemComment As String
    Dim FilmComment As String
    Dim DemoComment As String
    Dim Sex As String
    'Dim dob As Date

10  On Error GoTo RTFPrintHaemSpecific_Error

20  ReDim CommentLines(1 To 2) As String

30  sql = "SELECT * FROM Demographics " & _
          "WHERE SampleID = '" & RP.SampleID & "'"
40  Set tb = New Recordset
50  RecOpenClient 0, tb, sql
60  If tb.EOF Then
70      Exit Sub
Else
Sex = tb!Sex
'dob = tb!dob

80  End If

90  If Not SetPrinter("CHHAEM") Then Exit Sub

100 sql = "SELECT " & MonoSpotOrESR & " Result, RunDateTime, Operator " & _
          "FROM HaemResults " & _
          "WHERE SampleID = '" & RP.SampleID & "'"

110 Set tbh = New Recordset
120 RecOpenClient 0, tbh, sql
130 If Not tbh.EOF Then
140     If tbh!Result & "" <> "" Then

150         sql = "UPDATE HaemResults " & _
                  "SET Printed = 1, Valid = 1 " & _
                  "WHERE SampleID = '" & RP.SampleID & "'"
160         Cnxn(0).Execute sql

170         RunDateTime = Format(tbh!RunDateTime, "dd/mm/yyyy hh:mm")
180         Operator = tbh!Operator & ""
190         Dob = ""
200         If Not IsNull(tb!Dob) Then
210             If IsDate(tb!Dob) Then
220                 Dob = tb!Dob
230             End If
240         End If

250         RTFPrintHeading "Haematology", tb!PatName & "", Dob, tb!Chart & "", _
                            tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

260         With frmMain.rtb
270             .SelFontSize = 10
280             .SelText = vbCrLf
290             .SelText = vbCrLf
300             .SelText = vbCrLf

310             .SelText = vbCrLf
320             .SelText = vbCrLf
330             .SelText = Space$(20)

340             If UCase$(MonoSpotOrESR) = "ESR" Then
350                 .SelText = "ESR   " & Left$(tbh!Result & Space(7), 7) & "mm/hr   " & HaemNormalRange("ESR", Sex, Dob)

360             ElseIf UCase$(MonoSpotOrESR) = UCase("Malaria") Then
370                 .SelText = MonoSpotOrESR & " Screening Kit  "
380                 .SelBold = True
390                 .SelFontSize = 9
400                 If UCase$(Left$(tbh!Result & "", 1)) = "P" Then
410                     .SelText = "Positive"
420                 ElseIf UCase$(Left$(tbh!Result & "", 1)) = "N" Then
430                     .SelText = "Negative"
440                 End If
450             Else
457                 If UCase(MonoSpotOrESR) = "MONOSPOT" Then
458                     .SelText = "Infectious Mono Screen " & "   "
459                 Else
460                     .SelText = MonoSpotOrESR & "   "
461                 End If
470                 .SelBold = True
480                 .SelFontSize = 9
490                 If UCase$(Left$(tbh!Result & "", 1)) = "P" Then
500                     .SelText = "Positive"
510                 ElseIf UCase$(Left$(tbh!Result & "", 1)) = "N" Then
520                     .SelText = "Negative"
530                 End If
540             End If
550             .SelBold = False
560             .SelFontSize = 10

570             .SelText = vbCrLf
580             .SelText = vbCrLf
590             .SelText = vbCrLf
600             .SelText = vbCrLf

610             Set OBs = New Observations
620             Set OBs = OBs.Load(RP.SampleID, "Haematology", "Film", "Demographic")
630             If Not OBs Is Nothing Then
640                 For Each OB In OBs
650                     If UCase$(OB.Discipline) = "HAEMATOLOGY" Then
660                         HaemComment = OB.Comment
670                     ElseIf UCase$(OB.Discipline) = "FILM" Then
680                         FilmComment = OB.Comment
690                     ElseIf UCase$(OB.Discipline) = "DEMOGRAPHIC" Then
700                         DemoComment = OB.Comment
710                     End If
720                 Next
730                 FillCommentLines Trim$(HaemComment & " " & FilmComment), 2, CommentLines(), 97
740                 For n = 1 To 2
750                     .SelColor = vbBlack
760                     .SelText = CommentLines(n) & vbCrLf
770                 Next
780                 FillCommentLines DemoComment, 2, CommentLines(), 97
790                 For n = 1 To 2
800                     .SelColor = vbBlack
810                     .SelText = CommentLines(n) & vbCrLf
820                 Next
830             End If

840             RTFPrintFooter "Haematology", RP.Initiator, tb!SampleDate & "", RunDateTime, tb!RecDate & ""

850             PrintAndStore "Haematology", "H"

860         End With
870     End If
880 End If

890 Exit Sub

RTFPrintHaemSpecific_Error:

    Dim strES As String
    Dim intEL As Integer

900 intEL = Erl
910 strES = Err.Description
920 LogError "modRTF", "RTFPrintHaemSpecific", intEL, strES, sql

End Sub

Public Sub RTFPrintTDM(ByVal SampleID As String, ByVal SplitNumber As Integer)

      Dim BRs As New BIEResults
      Dim tb As Recordset
      Dim sql As String
      Dim Sex As String
      Dim lpc As Integer
      Dim n As Integer
      Dim BR As BIEResult
      Dim TestCount As Integer
      Dim SampleType As String
      Dim OBs As Observations
10    On Error GoTo RTFPrintTDM_Error

20    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim ReceivedDate As String
      Dim Rundate As String
      Dim Dob As String
      Dim RunTime As String
      Dim Fasting As String
      Dim udtPrintLine(0 To 3) As PrintLine    'max 30 result lines per page
      Dim strFormat As String
      Dim GentamicinTrough As String
      Dim GentamicinPeak As String
      Dim TobramicinTrough As String
      Dim TobramicinPeak As String
      Dim BioComment As String
      Dim DemoComment As String
      Dim CodeForGentamicin As String
      Dim CodeForTobramicin As String
      Dim SplitName As String
      Dim PeakTime As String
      Dim PeakSID As String
      Dim TroughTime As String
      Dim TroughSID As String

30    CodeForGentamicin = UCase$(GetOptionSetting("BioCodeForGentamicin", ""))
40    CodeForTobramicin = UCase$(GetOptionSetting("BioCodeForTobramicin", ""))

50    For n = 0 To 3
60        udtPrintLine(n).Analyte20 = ""
70        udtPrintLine(n).Result = ""
80        udtPrintLine(n).Flag = ""
90        udtPrintLine(n).Units = ""
100       udtPrintLine(n).NormalRange = ""
110       udtPrintLine(n).Fasting = ""
120   Next

130   sql = "Select * from Demographics where " & _
            "SampleID = '" & SampleID & "'"
140   Set tb = New Recordset
150   RecOpenClient 0, tb, sql

160   If tb.EOF Then
170       Exit Sub
180   End If

190   If Not IsNull(tb!Fasting) Then
200       Fasting = tb!Fasting
210   Else
220       Fasting = False
230   End If

240   If IsDate(tb!Dob) Then
250       Dob = Format(tb!Dob, "dd/mmm/yyyy")
260   Else
270       Dob = ""
280   End If

290   Set BRs = BRs.Load("Bio", SampleID, "Results", gDONTCARE, gDONTCARE)

300   TestCount = BRs.Count
310   SampleType = BRs(1).SampleType
320   If Trim$(SampleType) = "" Then SampleType = "S"

330   If Not SetPrinter("CHBIO") Then Exit Sub

340   GetPeaksAndTroughs GentamicinTrough, GentamicinPeak, _
                         TobramicinTrough, TobramicinPeak, _
                         BRs, tb!PatName & "", _
                         PeakSID, PeakTime, _
                         TroughSID, TroughTime

350   If IsDate(PeakTime) Then
360       If Format$(PeakTime, "HH:nn") = "00:00" Then
370           PeakTime = Format(PeakTime, "dd/MM/yyyy")
380       Else
390           PeakTime = Format$(PeakTime, "dd/MM/yyyy HH:nn")
400       End If
410   End If

420   If IsDate(TroughTime) Then
430       If Format$(TroughTime, "HH:nn") = "00:00" Then
440           TroughTime = Format(TroughTime, "dd/MM/yyyy")
450       Else
460           TroughTime = Format$(TroughTime, "dd/MM/yyyy HH:nn")
470       End If
480   End If

490   lpc = 0
500   For Each BR In BRs
510       RunTime = BR.RunTime

520       If BR.Code = CodeForGentamicin Then
530           If GentamicinTrough <> "" And GentamicinPeak <> "" Then
540               udtPrintLine(lpc).Analyte20 = "Gentamicin Trough"
550               udtPrintLine(lpc).Result = Format$(GentamicinTrough, "0.0")
560               udtPrintLine(lpc).Units = BR.Units
570               udtPrintLine(lpc).NormalRange = "(     < 2.0 )"
580               If Val(GentamicinTrough) > BR.PlausibleHigh Then
590                   udtPrintLine(lpc).Flag = " X "
600               ElseIf Val(GentamicinTrough) < BR.PlausibleLow Then
610                   udtPrintLine(lpc).Flag = " X "
620               ElseIf Val(GentamicinTrough) > 2 Then
630                   udtPrintLine(lpc).Flag = " H "
640               Else
650                   udtPrintLine(lpc).Flag = "   "
660               End If
670               udtPrintLine(lpc).Comment = "Sample Time : " & TroughTime
680               lpc = lpc + 1

690               udtPrintLine(lpc).Analyte20 = "Gentamicin Peak"
700               udtPrintLine(lpc).Result = Format$(GentamicinPeak, "0.0")
710               udtPrintLine(lpc).Units = BR.Units
720               udtPrintLine(lpc).NormalRange = "( 5.0 - 10.0)"
730               If Val(GentamicinPeak) > BR.PlausibleHigh Then
740                   udtPrintLine(lpc).Flag = " X "
750               ElseIf Val(GentamicinPeak) < BR.PlausibleLow Then
760                   udtPrintLine(lpc).Flag = " X "
770               ElseIf Val(GentamicinPeak) > 10 Then
780                   udtPrintLine(lpc).Flag = " H "
790               ElseIf Val(GentamicinPeak) < 5 Then
800                   udtPrintLine(lpc).Flag = " L "
810               Else
820                   udtPrintLine(lpc).Flag = "   "
830               End If
840               udtPrintLine(lpc).Comment = "Sample Time : " & PeakTime
850               lpc = lpc + 1

860           ElseIf GentamicinTrough <> "" Then
870               udtPrintLine(lpc).Analyte20 = "Gentamicin"
880               udtPrintLine(lpc).Result = Format$(GentamicinTrough, "0.0")
890               udtPrintLine(lpc).Units = BR.Units
900               udtPrintLine(lpc).NormalRange = "(    -    )"
910               If Val(GentamicinTrough) > BR.PlausibleHigh Then
920                   udtPrintLine(lpc).Flag = " X "
930               ElseIf Val(GentamicinTrough) < BR.PlausibleLow Then
940                   udtPrintLine(lpc).Flag = " X "
950               ElseIf Val(GentamicinTrough) > BR.FlagHigh Then
960                   udtPrintLine(lpc).Flag = " H "
970               Else
980                   udtPrintLine(lpc).Flag = "   "
990               End If
1000              lpc = lpc + 1

1010          End If

1020          LogBioAsPrinted RP.SampleID, BR.Code

1030      ElseIf BR.Code = CodeForTobramicin Then
1040          If TobramicinTrough <> "" And TobramicinPeak <> "" Then
1050              udtPrintLine(lpc).Analyte20 = "Tobramicin Trough"
1060              udtPrintLine(lpc).Result = Format$(TobramicinTrough, "0.0")
1070              udtPrintLine(lpc).Units = BR.Units
1080              udtPrintLine(lpc).NormalRange = "(     < 2.0 )"
1090              If Val(TobramicinTrough) > BR.PlausibleHigh Then
1100                  udtPrintLine(lpc).Flag = " X "
1110              ElseIf Val(TobramicinTrough) < BR.PlausibleLow Then
1120                  udtPrintLine(lpc).Flag = " X "
1130              ElseIf Val(TobramicinTrough) > 2 Then
1140                  udtPrintLine(lpc).Flag = " H "
1150              Else
1160                  udtPrintLine(lpc).Flag = "   "
1170              End If
1180              udtPrintLine(lpc).Comment = "Sample Time : " & TroughTime
1190              lpc = lpc + 1

1200              udtPrintLine(lpc).Analyte20 = "Tobramicin Peak"
1210              udtPrintLine(lpc).Result = Format$(TobramicinPeak, "0.0")
1220              udtPrintLine(lpc).Units = BR.Units
1230              udtPrintLine(lpc).NormalRange = "( 6.0 - 10.0)"
1240              If Val(TobramicinPeak) > BR.PlausibleHigh Then
1250                  udtPrintLine(lpc).Flag = " X "
1260              ElseIf Val(TobramicinPeak) < BR.PlausibleLow Then
1270                  udtPrintLine(lpc).Flag = " X "
1280              ElseIf Val(TobramicinPeak) > 10 Then
1290                  udtPrintLine(lpc).Flag = " H "
1300              ElseIf Val(TobramicinPeak) < 6 Then
1310                  udtPrintLine(lpc).Flag = " L "
1320              Else
1330                  udtPrintLine(lpc).Flag = "   "
1340              End If
1350              udtPrintLine(lpc).Comment = "Sample Time : " & PeakTime
1360              lpc = lpc + 1

1370          ElseIf TobramicinTrough <> "" Then
1380              udtPrintLine(lpc).Analyte20 = "Tobramicin"
1390              udtPrintLine(lpc).Result = Format$(TobramicinTrough, "0.0")
1400              udtPrintLine(lpc).Units = BR.Units
1410              udtPrintLine(lpc).NormalRange = "(     -    )"
1420              If Val(TobramicinTrough) > BR.PlausibleHigh Then
1430                  udtPrintLine(lpc).Flag = " X "
1440              ElseIf Val(TobramicinTrough) < BR.PlausibleLow Then
1450                  udtPrintLine(lpc).Flag = " X "
1460              Else
1470                  udtPrintLine(lpc).Flag = "   "
1480              End If
1490              lpc = lpc + 1

1500          End If

1510          LogBioAsPrinted RP.SampleID, BR.Code
1520      End If

1530  Next

1540  RTFPrintHeading "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
                      tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

1550  Sex = tb!Sex & ""

1560  With frmMain.rtb

1570      .SelFontSize = 10

1580      For n = 0 To 3
1590          If Trim$(udtPrintLine(n).Analyte20) <> "" Then
1600              .SelText = Space$(16)
1610              .SelBold = False
1620              .SelText = udtPrintLine(n).Analyte20
1630              If udtPrintLine(n).Flag <> "   " Then
1640                  .SelBold = True
1650              End If
1660              .SelText = udtPrintLine(n).Result
1670              .SelText = udtPrintLine(n).Flag
1680              .SelBold = False
1690              .SelText = udtPrintLine(n).Units
1700              .SelText = udtPrintLine(n).NormalRange
1710              .SelText = udtPrintLine(n).Fasting
1720              .SelText = vbCrLf
1730              .SelText = Space$(16)
1740              .SelText = udtPrintLine(n).Comment
1750              .SelText = vbCrLf
1760              .SelText = vbCrLf
1770          End If
1780      Next

1790      Set OBs = New Observations
1800      Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
1810      If Not OBs Is Nothing Then
1820          FillCommentLines OBs(1).Comment, 4, Comments(), 97
1830          For n = 1 To 4
1840              .SelColor = vbBlack
1850              .SelText = Comments(n)
1860              .SelText = vbCrLf
1870          Next
1880      End If

1890      Set OBs = New Observations
1900      Set OBs = OBs.Load(RP.SampleID, "Demographic")
1910      If Not OBs Is Nothing Then
1920          FillCommentLines OBs(1).Comment, 4, Comments(), 97
1930          For n = 1 To 4
1940              .SelColor = vbBlack
1950              .SelText = Comments(n)
1960              .SelText = vbCrLf
1970          Next
1980      End If

1990      .SelColor = vbBlack
2000      RTFPrintNoSexDoB Sex, Dob

2010      If IsDate(tb!SampleDate) Then
2020          SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
2030      Else
2040          SampleDate = ""
2050      End If
2060      If IsDate(tb!RecDate) Then
2070          ReceivedDate = Format(tb!RecDate, "dd/mmm/yyyy hh:mm")
2080      Else
2090          ReceivedDate = ""
2100      End If
2110      If IsDate(RunTime) Then
2120          Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
2130      Else
2140          If IsDate(tb!Rundate) Then
2150              Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
2160          Else
2170              Rundate = ""
2180          End If
2190      End If

2200      SplitName = GetOptionSetting("PrintBioSplitName" & Format$(SplitNumber), "Biochemistry")

2210      RTFPrintFooter "Biochemistry", RP.Initiator, SampleDate, Rundate, ReceivedDate, SplitName, SplitNumber

2220      PrintAndStore "Biochemistry", "BTDM"

2230  End With

2240  sql = "Update BioResults set Printed = '1' where " & _
            "SampleID = '" & RP.SampleID & "'"
2250  Cnxn(0).Execute sql

2260  Exit Sub

RTFPrintTDM_Error:

      Dim strES As String
      Dim intEL As Integer

2270  intEL = Erl
2280  strES = Err.Description
2290  LogError "modRTF", "RTFPrintTDM", intEL, strES, sql

End Sub


Private Sub GetPeaksAndTroughs(ByRef GentamicinTrough As String, _
                               ByRef GentamicinPeak As String, _
                               ByRef TobramicinTrough As String, _
                               ByRef TobramicinPeak As String, _
                               ByVal CurrentBRs As BIEResults, _
                               ByVal PatientName As String, _
                               ByRef PeakSID As String, _
                               ByRef PeakTime As String, _
                               ByRef TroughSID As String, _
                               ByRef TroughTime As String)

      Dim tb As Recordset
      Dim tbR As Recordset
      Dim sql As String
      Dim GentCurrentValue As String
      Dim GentAssValue As String
      Dim TOBCurrentValue As String
      Dim TOBAssValue As String
      Dim n As Integer
      Dim SampleID As String
      Dim CodeForGentamicin As String
      Dim CodeForTobramicin As String

      'Gentamicin and Tobramicin

10    On Error GoTo GetPeaksAndTroughs_Error

20    CodeForGentamicin = UCase$(GetOptionSetting("BioCodeForGentamicin", ""))
30    CodeForTobramicin = UCase$(GetOptionSetting("BioCodeForTobramicin", ""))

40    SampleID = CurrentBRs(1).SampleID

50    GentamicinTrough = ""
60    GentamicinPeak = ""
70    TobramicinTrough = ""
80    TobramicinPeak = ""

      'If CurrentBRs.Count < 3 Then
90    For n = 1 To CurrentBRs.Count
100       If CurrentBRs(n).Code = CodeForGentamicin Then
110           GentamicinTrough = CurrentBRs(n).Result
120           GentCurrentValue = CurrentBRs(n).Result


130           sql = "Select distinct D.SampleID " & _
                    "from Demographics as D " & _
                    "where D.sampleid in " & _
                    "  (  select SampleID from BioResults where " & _
                    "     (SampleID = '" & Val(SampleID) - 1 & "' or SampleID = '" & Val(SampleID) + 1 & "') " & _
                    "     and Code = '" & CodeForGentamicin & "'  ) " & _
                    "and D.PatName = '" & AddTicks(PatientName) & "' " & _
                    "and (D.SampleID = '" & Val(SampleID) - 1 & "' or SampleID = '" & Val(SampleID) + 1 & "')"
140           Set tb = New Recordset
150           RecOpenServer 0, tb, sql
160           If Not tb.EOF Then
170               sql = "Select Result from BioResults where " & _
                        "SampleID = '" & tb!SampleID & "' " & _
                        "and Code = '" & CodeForGentamicin & "'"
180               Set tbR = New Recordset
190               RecOpenServer 0, tbR, sql
200               If Not tbR.EOF Then
210                   GentAssValue = tbR!Result & ""

220                   If Val(GentAssValue) < Val(GentCurrentValue) Or InStr(GentAssValue, "<") <> 0 Then
230                       GentamicinTrough = Format$(GentAssValue, "0.0")
240                       GentamicinPeak = Format$(GentCurrentValue, "0.0")
250                       PeakSID = CurrentBRs(n).SampleID
260                       PeakTime = GetSampleTime(CurrentBRs(n).SampleID)
270                       TroughSID = tb!SampleID & ""
280                       TroughTime = GetSampleTime(tb!SampleID)
290                   Else
300                       GentamicinPeak = Format$(GentAssValue, "0.0")
310                       GentamicinTrough = Format$(GentCurrentValue, "0.0")
320                       PeakSID = tb!SampleID & ""
330                       PeakTime = GetSampleTime(tb!SampleID)
340                       TroughSID = CurrentBRs(n).SampleID
350                       TroughTime = GetSampleTime(CurrentBRs(n).SampleID)
360                   End If
370               End If
380           End If
390       ElseIf CurrentBRs(n).Code = CodeForTobramicin Then
400           TobramicinTrough = CurrentBRs(n).Result
410           TOBCurrentValue = CurrentBRs(n).Result
420           sql = "Select distinct D.SampleID " & _
                    "from Demographics as D " & _
                    "where D.sampleid in " & _
                    "  (  select SampleID from BioResults where " & _
                    "     (SampleID = '" & Val(SampleID) - 1 & "' or SampleID = '" & Val(SampleID) + 1 & "') " & _
                    "     and Code = '" & CodeForTobramicin & "'  ) " & _
                    "and D.PatName = '" & AddTicks(PatientName) & "' " & _
                    "and (D.SampleID = '" & Val(SampleID) - 1 & "' or SampleID = '" & Val(SampleID) + 1 & "')"
430           Set tb = New Recordset
440           RecOpenServer 0, tb, sql
450           If Not tb.EOF Then
460               sql = "Select Result from BioResults where " & _
                        "SampleID = '" & tb!SampleID & "' " & _
                        "and Code = '" & CodeForTobramicin & "'"
470               Set tbR = New Recordset
480               RecOpenServer 0, tbR, sql
490               If Not tbR.EOF Then
500                   TOBAssValue = tbR!Result & ""
510                   If Val(TOBAssValue) < Val(TOBCurrentValue) Or InStr(TOBAssValue, "<") <> 0 Then
520                       TobramicinTrough = Format$(TOBAssValue, "0.0")
530                       TobramicinPeak = Format$(TOBCurrentValue, "0.0")
540                       PeakSID = CurrentBRs(n).SampleID
550                       PeakTime = GetSampleTime(CurrentBRs(n).SampleID)
560                       TroughSID = tb!SampleID & ""
570                       TroughTime = GetSampleTime(tb!SampleID)
580                   Else
590                       TobramicinPeak = Format$(TOBAssValue, "0.0")
600                       TobramicinTrough = Format$(TOBCurrentValue, "0.0")
610                       PeakSID = tb!SampleID & ""
620                       PeakTime = GetSampleTime(tb!SampleID)
630                       TroughSID = CurrentBRs(n).SampleID
640                       TroughTime = GetSampleTime(CurrentBRs(n).SampleID)
650                   End If
660               End If
670           End If
680       End If
690   Next
      'End If

700   Exit Sub

GetPeaksAndTroughs_Error:

      Dim strES As String
      Dim intEL As Integer

710   intEL = Erl
720   strES = Err.Description
730   LogError "modRTF", "GetPeaksAndTroughs", intEL, strES, sql

End Sub

Private Function GetSampleTime(ByVal SID As String) As String

      Dim RetVal As String
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetSampleTime_Error

20    sql = "SELECT SampleDate FROM Demographics WHERE " & _
            "SampleID = '" & SID & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        If IsDate(tb!SampleDate) Then
70            If Format$(tb!SampleDate, "HH:nn") <> "00:00" Then
80                RetVal = Format$(tb!SampleDate, "dd/MM/yyyy HH:nn")
90            Else
100               RetVal = Format$(tb!SampleDate, "dd/MM/yyyy")
110           End If
120       Else
130           RetVal = ""
140       End If
150   Else
160       RetVal = ""
170   End If

180   GetSampleTime = RetVal

190   Exit Function

GetSampleTime_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "modRTF", "GetSampleTime", intEL, strES, sql

End Function


Private Sub PrintValueOrXXX(ByVal Flag As String, ByVal Value As String, ByVal FlagWidth As Integer)

10    On Error GoTo PrintValueOrXXX_Error

20    With frmMain.rtb
30        If Flag <> "X" Then
40            .SelText = Right$("     " & Value & "", FlagWidth)
50        Else
60            .SelText = String(FlagWidth, "X")
70        End If
80    End With

90    Exit Sub

PrintValueOrXXX_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modRTF", "PrintValueOrXXX", intEL, strES

End Sub

Public Sub RTFPrintBioSplit(ByVal SplitNumber As Integer)

      Dim tb As Recordset
      Dim tbUN As Recordset
      Dim tbF As Recordset
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
      Dim SampleDate As String
      Dim ReceivedDate As String
      Dim Rundate As String
      Dim Dob As String
      Dim RunTime As String
      Dim Fasting As String
      Dim udtPrintLine(0 To 50) As PrintLine
      Dim strFormat As String
      Dim SplitName As String
      Dim X As Integer

      Dim CodeForEGFR As String
      Dim CodeForChol As String
      Dim CodeForGlucose As String
      Dim CodeForTrig As String
      Dim CodeForProgesterone As String
      Dim CodeForFSH As String
      Dim CodeForLH As String
      Dim CodeForOestradiol As String
      Dim FoundProgesterone As Boolean
      Dim FoundFSH As Boolean
      Dim FoundLH As Boolean
      Dim FoundOestradiol As Boolean
      Dim AutoComments As String
      Dim CodeForGentamicin As String
      Dim CodeForTobramicin As String
      Dim TDMFound As Boolean
      Dim Code As String
      Dim InhibitMask As String

10    On Error GoTo RTFPrintBioSplit_Error

20    SplitName = GetOptionSetting("PrintBioSplitName" & Format$(SplitNumber), "Biochemistry")
30    CodeForGlucose = GetOptionSetting("BioCodeForGlucose", "")
40    CodeForTrig = GetOptionSetting("BioCodeForTrig", "")

50    CodeForEGFR = UCase$(GetOptionSetting("BioCodeForEGFR", "5555"))
60    CodeForChol = UCase$(GetOptionSetting("BioCodeForChol", ""))
70    CodeForProgesterone = UCase$(GetOptionSetting("BioCodeForProgesterone", "191"))
80    CodeForFSH = UCase$(GetOptionSetting("BioCodeForFSH", "81"))
90    CodeForLH = UCase$(GetOptionSetting("BioCodeForLH", ""))
100   CodeForOestradiol = UCase$(GetOptionSetting("BioCodeForOestradiol", "Oestrad"))

110   CodeForGentamicin = UCase$(GetOptionSetting("BioCodeForGentamicin", ""))
120   CodeForTobramicin = UCase$(GetOptionSetting("BioCodeForTobramicin", ""))

130   ReDim Comments(1 To 4) As String
140   ReDim lp(0 To 50) As String

150   For n = 0 To 50
160       udtPrintLine(n).Analyte = ""
170       udtPrintLine(n).Result = ""
180       udtPrintLine(n).Flag = ""
190       udtPrintLine(n).Units = ""
200       udtPrintLine(n).NormalRange = ""
210       udtPrintLine(n).Fasting = ""
220   Next

230   sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
240   Set tb = New Recordset
250   RecOpenClient 0, tb, sql

260   If tb.EOF Then
270       Exit Sub
280   End If

290   If Not IsNull(tb!Fasting) Then
300       Fasting = tb!Fasting
310   Else
320       Fasting = False
330   End If

340   If IsDate(tb!Dob) Then
350       Dob = Format(tb!Dob, "dd/mmm/yyyy")
360   Else
370       Dob = ""
380   End If

390   ResultsPresent = False
400   Set BRs = BRs.Load("Bio", RP.SampleID, "Results", gDONTCARE, gDONTCARE)
410   TDMFound = False
420   If Not BRs Is Nothing Then
430       For X = BRs.Count To 1 Step -1

440           Code = BRs(X).Code

450           If BRs(X).PrintSplit = SplitNumber And _
                 (Code = CodeForGentamicin Or _
                  Code = CodeForTobramicin) Then
460               TDMFound = False
470           End If

480           If BRs(X).ShortName <> "H" And BRs(X).ShortName <> "I" And BRs(X).ShortName <> "L" Then

                  '            If BRs(X).PrintSplit <> SplitNumber Or _
                               '               Code = CodeForTobramicin Or _
                               '               Code = CodeForTobramicin Or _
                               '               BRs(X).Printable = False Or _
                               '               IsInhibited("Bio", BRs(X).ShortName) Then
                  '                BRs.RemoveItem X
                  '            End If
490               If BRs(X).PrintSplit <> SplitNumber Or _
                     BRs(X).Printable = False Or _
                     IsInhibited("Bio", BRs(X).ShortName) Then
500                   BRs.RemoveItem X
510               End If
520           End If
530       Next


540       If TDMFound Then
550           RTFPrintTDM RP.SampleID, SplitNumber
560       End If

570       TestCount = BRs.Count
580       If TestCount <> 0 Then
590           ResultsPresent = True
600           SampleType = BRs(1).SampleType
610           If Trim$(SampleType) = "" Then SampleType = "S"
620       End If
630   End If

640   If Not ResultsPresent And Not CommentsPresent(RP.SampleID, "Biochemistry") Then
650       Exit Sub
660   End If

670   If Not SetPrinter("CHBIO") Then Exit Sub

680   lpc = 0
      '670   LoadLIH RP.SampleID, Lipaemic, Icteric, Haemolysed

690   FoundFSH = False
700   FoundLH = False
710   FoundOestradiol = False
720   FoundProgesterone = False
730   AutoComments = ""
740   For Each BR In BRs
750       If BR.ShortName <> "H" And BR.ShortName <> "I" And BR.ShortName <> "L" Then

760           AutoComments = Trim$(AutoComments & " " & CheckAutoComments(RP.SampleID, BR.ShortName, 2))

770           Code = BR.Code

780           If Code = CodeForProgesterone Then
790               FoundProgesterone = True
800           ElseIf Code = CodeForFSH Then
810               FoundFSH = True
820           ElseIf Code = CodeForLH Then
830               FoundLH = True
840           ElseIf Code = CodeForOestradiol Then
850               FoundOestradiol = True
860           End If

870           RunTime = BR.RunTime


880           InhibitMask = MaskInhibit(BR, BRs)
890           If InhibitMask = "XH" Or InhibitMask = "XI" Or InhibitMask = "XL" Then
900               v = "*****"
910           Else
920               v = BR.Result
930           End If
              '
              '860     If LIHEffects(Code, Lipaemic, Icteric, Haemolysed) Then
              '870       v = "*****"
              '880     Else
              '890       v = BR.Result
              '900     End If

940           If Code = CodeForGlucose Or _
                 Code = CodeForChol Or _
                 Code = CodeForTrig Then
950               If Fasting Then
960                   If Code = CodeForGlucose Then
970                       sql = "Select * from Fastings where " & _
                                "TestName = 'GLU'"
980                       Set tbF = New Recordset
990                       RecOpenServer 0, tbF, sql
1000                  ElseIf Code = CodeForChol Then
1010                      sql = "Select * from Fastings where " & _
                                "TestName = 'CHO'"
1020                      Set tbF = New Recordset
1030                      RecOpenServer 0, tbF, sql
1040                  ElseIf Code = CodeForTrig Then
1050                      sql = "Select * from Fastings where " & _
                                "TestName = 'TRI'"
1060                      Set tbF = New Recordset
1070                      RecOpenServer 0, tbF, sql
1080                  End If
1090                  If Not tbF.EOF Then
1100                      High = tbF!FastingHigh
1110                      Low = tbF!FastingLow
1120                  Else
1130                      High = Val(BR.High)
1140                      Low = Val(BR.Low)
1150                  End If
1160              Else
1170                  High = Val(BR.High)
1180                  Low = Val(BR.Low)
1190              End If
1200          Else
1210              High = Val(BR.High)
1220              Low = Val(BR.Low)
1230          End If

1240          If Low = 0 And High = 9999 Then    ' MASOOD 30-SEP-2015
1250              strLow = (" ")
1260          ElseIf Low < 10 Then
1270              strLow = Format(Low, " 0.00")
1280          ElseIf Low < 100 Then
1290              strLow = Format(Low, " ##.0")
1300          Else
1310              strLow = Format(Low, "#####")
1320          End If

1330          If Low = 0 And High = 9999 Then    ' MASOOD 30-SEP-2015
1340              strLow = (" ")
1350          ElseIf High < 10 Then
1360              strHigh = Format(High, "0.00 ")
1370          ElseIf High < 100 Then
1380              strHigh = Format(High, "##.0 ")
1390          ElseIf High < 1000 Then
1400              strHigh = Format(High, "###  ")
1410          Else
1420              strHigh = Format(High, "#####")
1430          End If



1440          If IsNumeric(v) Then

1450              If Fasting = True And (Code = CodeForGlucose Or Code = CodeForChol Or Code = CodeForTrig) Then
1460                  If Val(v) > High Then
1470                      udtPrintLine(lpc).Flag = " H "
1480                      lp(lpc) = "  "
1490                      Flag = " H"
1500                  ElseIf Val(v) < Low Then
1510                      udtPrintLine(lpc).Flag = " L "
1520                      lp(lpc) = "  "
1530                      Flag = " L"
1540                  End If
1550              Else
1560                  If Val(v) > BR.PlausibleHigh Then
1570                      v = "*****"
1580                      udtPrintLine(lpc).Result = "*****"
1590                      udtPrintLine(lpc).Flag = " X "
1600                      lp(lpc) = "  "
1610                      Flag = " X"
1620                  ElseIf Val(v) < BR.PlausibleLow Then
1630                      v = "*****"
1640                      udtPrintLine(lpc).Result = "*****"
1650                      udtPrintLine(lpc).Flag = " X "
1660                      lp(lpc) = "  "
1670                      Flag = " X"
1680                  ElseIf Val(v) >= BR.FlagHigh Then
1690                      udtPrintLine(lpc).Flag = " H "
1700                      lp(lpc) = "  "
1710                      Flag = " H"
1720                  ElseIf Val(v) <= BR.FlagLow Then
1730                      udtPrintLine(lpc).Flag = " L "
1740                      lp(lpc) = "  "
1750                      Flag = " L"
1760                  Else
1770                      udtPrintLine(lpc).Flag = "   "
1780                      lp(lpc) = "  "
1790                      Flag = "  "
1800                  End If
1810              End If
1820          Else
1830              udtPrintLine(lpc).Flag = "   "
1840              lp(lpc) = "  "
1850              Flag = "  "
1860          End If
1870          lp(lpc) = lp(lpc) & "  "
1880          lp(lpc) = lp(lpc) & Left$(BR.LongName & Space(20), 20)  '20
1890          udtPrintLine(lpc).Analyte = Left$(BR.LongName & Space(16), 16)  '16

1900          If IsNumeric(v) Then
1910              Select Case BR.Printformat
                  Case 0: strFormat = "#####0"
1920              Case 1: strFormat = "###0.0"
1930              Case 2: strFormat = "##0.00"
1940              Case 3: strFormat = "#0.000"
1950              End Select
1960              lp(lpc) = lp(lpc) & " " & Right$(Space(8) & Format(v, strFormat), 8)    'was 6
1970              udtPrintLine(lpc).Result = Format(v, strFormat)
1980          Else
1990              lp(lpc) = lp(lpc) & " " & Right$(Space(8) & v, 8)    'was 6
2000              udtPrintLine(lpc).Result = v
2010          End If
              'Else
              '  lp(lpc) = lp(lpc) & " XXXXXXX"
              'End If
2020          lp(lpc) = lp(lpc) & Flag & " "

2030          sql = "Select * from Lists where " & _
                    "ListType = 'UN' and Code = '" & BR.Units & "'"
2040          Set tbUN = Cnxn(0).Execute(sql)
2050          If Not tbUN.EOF Then
2060              cUnits = Left$(tbUN!Text & Space(16), 16)
2070          Else
2080              cUnits = Left$(BR.Units & Space(16), 16)
2090          End If
2100          udtPrintLine(lpc).Units = cUnits
2110          lp(lpc) = lp(lpc) & cUnits

2120          If BR.PrintRefRange Then
2130              If Low = 0 And High = 9999 Then     ' MASOOD 30-SEP-2015
2140                  lp(lpc) = lp(lpc) & "               "
2150                  udtPrintLine(lpc).NormalRange = "             "
2160              Else
2170                  lp(lpc) = lp(lpc) & "   ("
2180                  lp(lpc) = lp(lpc) & strLow & "-"
2190                  lp(lpc) = lp(lpc) & strHigh & ")"
2200                  udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
2210              End If
2220          Else
2230              lp(lpc) = lp(lpc) & "               "
2240              udtPrintLine(lpc).NormalRange = "             "
2250          End If

2260          udtPrintLine(lpc).Fasting = ""
2270          If Not IsNull(tb!Fasting) Then
2280              If tb!Fasting Then
2290                  udtPrintLine(lpc).Fasting = "(Fasting)"
2300                  lp(lpc) = lp(lpc) & "(Fasting)"
2310              End If
2320          End If

2330          If Flag <> " X" Then
2340              LogBioAsPrinted RP.SampleID, BR.Code
2350          End If

2360          lpc = lpc + 1

2370          sql = "UPDATE BioResults SET Printed = '1', Valid = '1' " & _
                    "WHERE SampleID = '" & RP.SampleID & "' " & _
                    "AND Code = '" & Code & "'"
2380          Cnxn(0).Execute sql
2390      End If
2400  Next

2410  If lpc > 0 Then
2420      RTFPrintHeading "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
                          tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

2430      Sex = tb!Sex & ""

2440      With frmMain.rtb
2450          .SelFontSize = 10
2460          For n = 0 To 19
2470              If Trim$(lp(n)) <> "" Then
2480                  If InStr(lp(n), " L ") Or InStr(lp(n), " H ") Then
2490                      .SelText = Space$(14) & Left$(lp(n), 33)
2500                      .SelBold = True
2510                      .SelFontSize = 9
2520                      .SelText = Mid$(lp(n), 34, 3)
2530                      .SelBold = False
2540                      .SelFontSize = 10
2550                      .SelText = Mid$(lp(n), 37)
2560                  Else
2570                      .SelText = Space$(14) & lp(n)
2580                  End If
2590                  .SelText = vbCrLf
2600                  .SelFontSize = 10
2610              End If
2620          Next

              Dim xx() As String
              Dim LL As Integer
2630          xx = Split(.Text, vbCr)
2640          LL = UBound(xx)
2650          Debug.Print "Bio end of results LL=" & LL

2660          If FoundFSH Or FoundLH Or FoundOestradiol Or FoundProgesterone Then
2670              .SelText = vbCrLf
2680              .SelFontSize = 10
2690              .SelText = Space$(10) & "Normal Ranges  Follicular      Mid Cycle       Luteal"
2700              .SelText = vbCrLf
2710              .SelFontSize = 10
2720              .SelText = Space$(25) & "Day 1 ~ 13       Day 14      Day 15 ~ 28"
2730              .SelText = vbCrLf
2740              .SelFontSize = 10
2750              If FoundFSH Then
2760                  .SelText = Space$(10) & "FSH" & Space$(12) & "3.0 - 8.1      2.5 - 16.7     1.4 - 5.5       "    'Wrike 388584845
2770                  .SelText = vbCrLf
2780                  .SelFontSize = 10
2790              End If
2800              If FoundLH Then
2810                  .SelText = Space$(10) & "LH" & Space$(13) & "1.8 - 11.8     7.6 - 89.1      0.6 - 14.0        "
2820                  .SelText = vbCrLf
2830                  .SelFontSize = 10
2840              End If
2850              If FoundOestradiol Then
2860                  .SelText = Space$(10) & "Oestradiol      77 - 921      139 - 2381      77 - 1140        "    'Wrike 388584845
2870                  .SelText = vbCrLf
2880                  .SelFontSize = 10
2890              End If
2900              If FoundProgesterone Then
2910                  .SelText = Space$(10) & "Progesterone                                   4 - 50"
2920                  .SelText = vbCrLf
2930                  .SelFontSize = 10
2940              End If
2950          End If

2960          If Trim$(AutoComments) <> "" Then
2970              .SelText = AutoComments
2980              .SelText = vbCrLf
2990              .SelFontSize = 10
3000          End If

3010          Set OBs = New Observations
3020          Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
3030          If Not OBs Is Nothing Then
3040              FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
3050              For n = 1 To 4
3060                  If Trim$(Comments(n)) <> "" Then
3070                      .SelText = Comments(n)
3080                      .SelText = vbCrLf
3090                      .SelFontSize = 10
3100                  End If
3110              Next
3120          End If

3130          Set OBs = New Observations
3140          Set OBs = OBs.Load(RP.SampleID, "Demographic")
3150          If Not OBs Is Nothing Then
3160              FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
3170              For n = 1 To 4
3180                  If Trim$(Comments(n)) <> "" Then
3190                      .SelText = Comments(n)
3200                      .SelText = vbCrLf
3210                      .SelFontSize = 10
3220                  End If
3230              Next
3240          End If

3250          RTFPrintNoSexDoB Sex, Dob

3260          .SelColor = vbBlack

3270          If IsDate(tb!RecDate) Then
3280              ReceivedDate = Format(tb!RecDate, "dd/MMM/yyyy HH:mm")
3290          Else
3300              ReceivedDate = ""
3310          End If
3320          If IsDate(tb!SampleDate) Then
3330              SampleDate = Format(tb!SampleDate, "dd/MMM/yyyy HH:mm")
3340          Else
3350              SampleDate = ""
3360          End If
3370          If IsDate(RunTime) Then
3380              Rundate = Format(RunTime, "dd/MMM/yyyy HH:mm")
3390          Else
3400              If IsDate(tb!Rundate) Then
3410                  Rundate = Format(tb!Rundate, "dd/MMM/yyyy")
3420              Else
3430                  Rundate = ""
3440              End If
3450          End If

3460          RTFPrintFooter "Biochemistry", RP.Initiator, SampleDate, Rundate, ReceivedDate, SplitName, SplitNumber

3470          PrintAndStore "Biochemistry", "B" & Format(SplitNumber)


3480      End With

3490  End If

3500  Exit Sub

RTFPrintBioSplit_Error:

      Dim strES As String
      Dim intEL As Integer

3510  intEL = Erl
3520  strES = Err.Description
3530  LogError "modRTF", "RTFPrintBioSplit", intEL, strES, sql

End Sub
Private Function MaskInhibit(ByVal BR As BIEResult, ByVal BRs As BIEResults) As String

      Dim Lx As LIH
      Dim RetVal As String
      Dim Result As Single
      Dim BRLIH As BIEResult
      Dim CutOffForThisParameter As Single
      Dim LIHValue As Single
      Dim BR1 As BIEResult
      Dim LiIcHas As New LIHs

10    RetVal = ""

20    Set Lx = LiIcHas.Item("L", BR.Code, "P")
30    If Not Lx Is Nothing Then
40        For Each BR1 In BRs
50            If BR1.ShortName = "L" Then
60                Set BRLIH = BR1
70                Exit For
80            End If
90        Next
100       If Not BRLIH Is Nothing Then
110           CutOffForThisParameter = Lx.CutOff
120           If CutOffForThisParameter > 0 Then
130               LIHValue = BRLIH.Result
140               If LIHValue >= CutOffForThisParameter Then
150                   RetVal = "XL"
160               End If
170           End If
180       End If
190   End If

200   If RetVal = "" Then
210       Set Lx = LiIcHas.Item("I", BR.Code, "P")
220       If Not Lx Is Nothing Then
230           For Each BR1 In BRs
240               If BR1.ShortName = "I" Then
250                   Set BRLIH = BR1
260                   Exit For
270               End If
280           Next
290           If Not BRLIH Is Nothing Then
300               CutOffForThisParameter = Lx.CutOff
310               If CutOffForThisParameter > 0 Then
320                   LIHValue = BRLIH.Result
330                   If LIHValue >= CutOffForThisParameter Then
340                       RetVal = "XI"
350                   End If
360               End If
370           End If
380       End If
390   End If

400   If RetVal = "" Then
410       Set Lx = LiIcHas.Item("H", BR.Code, "P")
420       If Not Lx Is Nothing Then
430           For Each BR1 In BRs
440               If BR1.ShortName = "H" Then
450                   Set BRLIH = BR1
460                   Exit For
470               End If
480           Next
490           If Not BRLIH Is Nothing Then
500               CutOffForThisParameter = Lx.CutOff
510               If CutOffForThisParameter > 0 Then
520                   LIHValue = BRLIH.Result
530                   If LIHValue >= CutOffForThisParameter Then
540                       RetVal = "XH"
550                   End If
560               End If
570           End If
580       End If
590   End If

600   MaskInhibit = RetVal

End Function

Public Sub RTFPrintFooter(ByVal Dept As String, _
                          ByVal Initiator As String, _
                          ByVal SampleDate As String, _
                          ByVal Rundate As String, _
                          ByVal ReceivedDate As String, _
                          Optional ByVal SplitName As String = "", _
                          Optional ByVal SplitNumber As String = "")

      Dim S As String
      Dim sql As String
      Dim tb As Recordset
      Dim DisciplineCode As String
      Dim FColour As Long
      Dim LL As Integer
      Dim X() As String
      Dim Y As Integer
      Dim iresult As Integer
      Dim PosX As POINTAPI
      Dim SampleType As String
      Dim ValidatedBy As String
      Dim AccreditationDept As String
      Dim AccreditationText As String
10    On Error GoTo RTFPrintFooter_Error

20    AccreditationDept = Dept
30    With frmMain.rtb
40        .SelFontName = "Courier New"
50        .SetFocus
60        iresult = GetCaretPos(PosX)
70        Debug.Print Dept & " X = " & PosX.X & " Y = " & PosX.Y
80        .SelFontSize = 9
90        .SelColor = vbBlack

100       If UCase(Dept) = "HAEMATOLOGY" Then
              'print Pregnancy comment
110           .SelText = "Pregnancy Specific Reference Ranges available in the Hospital /External Laboratory User Manuals"
120           .SelFontSize = 10
130           .SelColor = vbBlack
140       End If
          
          'print send copy to
150       If Trim$(RP.SendCopyTo) <> "" Then
160           .SelText = RP.Clinician & " Requested copy to be sent to " & RP.SendCopyTo
170       End If
180       .SelText = vbCrLf
190       .SelFontSize = 10


          'go to foooter location
200       .SetFocus
210       X = Split(.Text, vbCr)
220       LL = UBound(X)

230       Do While LL < 33
240           .SelFontSize = 10
250           .SelText = vbCrLf
260           X = Split(.Text, vbCr)
270           LL = UBound(X)
280       Loop

          'Start printing footer
290       .SelFontName = "Courier New"
300       .SelAlignment = rtfCenter
310       .SelBold = False
320       iresult = GetCaretPos(PosX)
330       Debug.Print Dept & "before line X = " & PosX.X & " Y = " & PosX.Y

          'line or copy of report text
340       If gPrintCopyReport = 0 Then
350           .SelFontSize = 4
360           .SelText = String$(200, "-") & vbCrLf
370       Else
380           .SelFontSize = 8
390           S = "- THIS IS A COPY REPORT - NOT FOR FILING -"
400           S = S & S
410           S = S & "- THIS IS A COPY REPORT -"
420           .SelColor = vbRed
430           .SelText = S & vbCrLf
440       End If

450       .SelFontName = "Courier New"
460       .SelAlignment = rtfLeft

470       Select Case Dept
          Case "Haematology":
480           FColour = vbRed
490           DisciplineCode = "J"
500           SampleType = "EDTA Whole Blood"
510       Case "Biochemistry"
520           FColour = vbGreen
530           DisciplineCode = "I"
540           SampleType = "Serum"
550       Case "Creat Clearance"
560           FColour = vbGreen
570           DisciplineCode = "I"
580           SampleType = "Serum / Urine"
590       Case "Glucose Series"
600           FColour = vbGreen
610           DisciplineCode = "I"
620           SampleType = "Fluoride Plasma"
630       Case "Gluc. Tolerance"
640           FColour = vbGreen
650           DisciplineCode = "I"
660           SampleType = "Fluoride Plasma"
670       Case "Microbiology":
680           DisciplineCode = "N"
690           FColour = vbYellow
700           SampleType = ""
710       Case "Blood Transfusion":
720           FColour = vbRed
730           DisciplineCode = "N"
740           SampleType = ""
750       Case "Coagulation":
760           DisciplineCode = "K"
770           FColour = vbRed        'RGB(80, 46, 107) 'Purple
780           SampleType = "Sodium Citrated Plasma"
790       End Select

800       RTFPrintText "    "
          'SampleType
810       If Trim$(SplitName) <> "" Then
820           SampleType = Left$(SplitName & Space(16), 16)
830       End If
840       RTFPrintText FormatString(SampleType, 27, " ", AlignLeft), 10

          'Empty space
850       RTFPrintText FormatString(" ", 27, " ", AlignLeft), 10

          'Validated By
860       If gPrintCopyReport = 0 Then
870           If Trim$(Initiator) <> "" Then
880               ValidatedBy = TechnicianCodeFor(Initiator)
890           End If
900       Else
910           sql = "SELECT TOP 1 Viewer FROM ViewedReports WHERE " & _
                    "SampleID = '" & RP.SampleID & "' " & _
                    "AND Discipline = '" & DisciplineCode & "' " & _
                    "AND DATEDIFF(minute, [datetime], getdate()) < 2 " & _
                    "ORDER BY [DateTime] DESC"
920           Set tb = New Recordset
930           RecOpenServer 0, tb, sql
940           If Not tb.EOF Then
950               ValidatedBy = tb!Viewer & ""
960           Else
970               ValidatedBy = TechnicianCodeFor(Initiator)
980           End If
990       End If
1000      RTFPrintText FormatString("Validated by: " & ValidatedBy, 27, " ", AlignRight), 10, , , , FColour

          'New Line
1010      RTFPrintText vbCrLf

1020      RTFPrintText "    "
          'Sample Datetime
1030      RTFPrintText FormatString("Sample Taken:" & Format(SampleDate, "dd/mm/yy HH:nn"), 27, " ", AlignLeft), 10, , , , FColour

          'Received DateTime
1040      RTFPrintText FormatString("Received:" & Format(ReceivedDate, "dd/mm/yy HH:nn"), 27, " ", AlignCenter), 10, , , , FColour

          'Tested DateTime
1050      RTFPrintText FormatString("Tested:" & Rundate, 27, " ", AlignRight), 10, , , , FColour



          '    If Format(SampleDate, "hh:mm") <> "00:00" Then
          '        RTFPrintText FormatString("Sample Date/Time:", 17, " ", AlignLeft), 10, , , , FColour
          '    Else
          '        RTFPrintText FormatString("Sample Date:", 17, " ", AlignLeft), 10, , , , FColour
          '    End If
          '    RTFPrintText FormatString(Format(SampleDate, "dd/mm/yy HH:nn"), 16, " ", AlignLeft), 10, , , , FColour



1060      Select Case AccreditationDept
          Case "Haematology":
1070          AccreditationText = GetOptionSetting("HaemAccreditation", "")
1080      Case "Biochemistry":
1090          AccreditationText = GetOptionSetting("BioAccreditation" & Format$(SplitNumber), "")
1100      Case "Coagulation":
1110          AccreditationText = GetOptionSetting("CoagAccreditation", "")
1120      Case "Microbiology":
1130          AccreditationText = GetOptionSetting("MicroAccreditation", "")
1140      End Select

1150      If AccreditationText <> "" Then
1160          .SelText = vbNewLine
              '        .SelText = Space$(10) & String$(75, "-") & vbCrLf

1170          .SelColor = vbRed
1180          .SelAlignment = rtfLeft

1190          .SelFontName = "Courier New"
1200          .SelFontSize = 10
1210          .SelBold = True
1220          .SelText = Space$(5) & AccreditationText
1230          .SelText = vbCrLf
1240      End If


1250  End With

1260  Exit Sub

RTFPrintFooter_Error:

      Dim strES As String
      Dim intEL As Integer

1270  intEL = Erl
1280  strES = Err.Description
1290  LogError "modRTF", "RTFPrintFooter", intEL, strES, sql

End Sub

'Public Sub RTFPrintFooter(ByVal Dept As String, _
 '                          ByVal Initiator As String, _
 '                          ByVal SampleDate As String, _
 '                          ByVal Rundate As String, _
 '                          Optional ByVal SplitName As String = "")
'
'Dim s As String
'Dim sql As String
'Dim tb As Recordset
'Dim DisciplineCode As String
'Dim FColour As Long
'Dim LL As Integer
'Dim X() As String
'Dim Y As Integer
'Dim iresult As Integer
'Dim PosX As POINTAPI
'
'On Error GoTo RTFPrintFooter_Error
'
'With frmMain.rtb
'    .SelFontName = "Courier New"
'    .SetFocus
'    iresult = GetCaretPos(PosX)
'    Debug.Print Dept & " X = " & PosX.X & " Y = " & PosX.Y
'    .SelFontSize = 10
'    .SelColor = vbBlack
'
'    'print send copy to
'    If Trim$(RP.SendCopyTo) <> "" Then
'        .SelText = RP.Clinician & " Requested copy to be sent to " & RP.SendCopyTo
'    End If
'    .SelText = vbCrLf
'    .SelFontSize = 10
'
'
'    'go to foooter location
'    .SetFocus
'    X = Split(.Text, vbCr)
'    LL = UBound(X)
'
'    Do While LL < 33
'        .SelFontSize = 10
'        .SelText = vbCrLf
'        X = Split(.Text, vbCr)
'        LL = UBound(X)
'    Loop
'
'    'Start printing footer
'    .SelFontName = "Courier New"
'    .SelAlignment = rtfCenter
'    .SelBold = False
'    iresult = GetCaretPos(PosX)
'    Debug.Print Dept & "before line X = " & PosX.X & " Y = " & PosX.Y
'
'    'line or copy of report text
'    If gPrintCopyReport = 0 Then
'        .SelFontSize = 4
'        .SelText = String$(200, "-") & vbCrLf
'    Else
'        .SelFontSize = 8
'        s = "- THIS IS A COPY REPORT - NOT FOR FILING -"
'        s = s & s
'        s = s & "- THIS IS A COPY REPORT -"
'        .SelColor = vbRed
'        .SelText = s & vbCrLf
'    End If
'
'    .SelFontName = "Courier New"
'    .SelAlignment = rtfLeft
'
'    Select Case Dept
'        Case "Haematology":
'            FColour = vbRed
'            DisciplineCode = "J"
'        Case "Biochemistry", "Creat Clearance", "Glucose Series", "Gluc. Tolerance"
'            FColour = vbGreen
'            DisciplineCode = "I"
'        Case "Microbiology":
'            DisciplineCode = "N"
'            FColour = vbYellow
'        Case "Blood Transfusion":
'            FColour = vbRed
'            DisciplineCode = "N"
'        Case "Coagulation":
'            DisciplineCode = "K"
'            FColour = vbRed        'RGB(80, 46, 107) 'Purple
'    End Select
'
'    .SelColor = FColour
'    .SelFontSize = 10
'    .SelBold = False
'
'    .SelColor = vbBlack
'    If Trim$(SplitName) <> "" Then
'        .SelText = Left$(SplitName & Space(16), 16)
'    Else
'        .SelText = Dept
'        If UCase$(Dept) = "HAEMATOLOGY" Then
'            .SelText = " Whole Blood"
'        ElseIf UCase$(Dept) = "COAGULATION" Then
'            .SelText = " Plasma"
'        End If
'    End If
'    .SelColor = FColour
'
'    If Format(SampleDate, "hh:mm") <> "00:00" Then
'        .SelText = " Sample Date/Time:" & Format(SampleDate, "dd/mm/yy HH:nn")
'    Else
'        .SelText = " Sample Date:" & Format(SampleDate, "dd/MM/yy")
'    End If
'
'    .SelText = " Tested:" & Format(Rundate, "dd/mm/yy hh:mm")
'
'    If gPrintCopyReport = 0 Then
'        If Trim$(Initiator) <> "" Then
'            .SelText = " Validated by " & TechnicianCodeFor(Initiator)
'        End If
'    Else
'        sql = "SELECT TOP 1 Viewer FROM ViewedReports WHERE " & _
         '              "SampleID = '" & RP.SampleID & "' " & _
         '              "AND Discipline = '" & DisciplineCode & "' " & _
         '              "AND DATEDIFF(minute, [datetime], getdate()) < 2 " & _
         '              "ORDER BY [DateTime] DESC"
'        Set tb = New Recordset
'        RecOpenServer 0, tb, sql
'        If Not tb.EOF Then
'            .SelText = " Printed by " & tb!Viewer & ""
'        Else
'            .SelText = " Printed by " & Left$(TechnicianCodeFor(Initiator), 14)
'        End If
'    End If
'
'End With
'
'Exit Sub
'
'RTFPrintFooter_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "modRTF", "RTFPrintFooter", intEL, strES, sql
'
'End Sub
Public Sub RTFPrintGTT()

      Dim tb As Recordset
      Dim tbDems As Recordset
      Dim tbRes As Recordset
      Dim sql As String
      Dim Dob As String
      Dim NameToFind As String
      Dim SampleDate As String
      Dim ReceivedDate As String
      Dim Rundate As String
      Dim CodeForGlucose As String
10    On Error GoTo RTFPrintGTT_Error

20    ReDim Comments(1 To 4) As String
      Dim n As Integer

30    CodeForGlucose = GetOptionSetting("BioCodeForGlucose", "")

40    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql
70    If tb.EOF Then Exit Sub

80    If IsDate(tb!SampleDate) Then
90        SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
100   Else
110       SampleDate = ""
120   End If
130   If IsDate(tb!RecDate) Then
140       ReceivedDate = Format(tb!RecDate, "dd/mmm/yyyy hh:mm")
150   Else
160       ReceivedDate = ""
170   End If
180   If IsDate(tb!Rundate) Then
190       Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
200   Else
210       Rundate = ""
220   End If

230   NameToFind = tb!PatName & ""

240   If IsDate(tb!Dob) Then
250       Dob = Format(tb!Dob, "dd/mmm/yyyy")
260   Else
270       Dob = ""
280   End If

290   sql = "select * from demographics where " & _
            "patname = '" & AddTicks(NameToFind) & "' " & _
            "and rundate = '" & Format(tb!Rundate, "dd/mmm/yyyy") & "' "
300   If IsDate(Dob) Then
310       sql = sql & "and DoB = '" & Format(Dob, "dd/mmm/yyyy") & "' "
320   End If
330   sql = sql & "order by SampleDate"
340   Set tbDems = New Recordset
350   RecOpenClient 0, tbDems, sql

360   If tbDems.EOF Then
370       Exit Sub
380   End If

390   If Not SetPrinter("CHBIO") Then Exit Sub

400   RTFPrintHeading "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
                      tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

410   With frmMain.rtb
420       .SelFontSize = 10
430       .SelText = vbCrLf
440       .SelText = vbCrLf
450       .SelText = vbCrLf

460       .SelColor = vbGreen
470       .SelFontSize = 12
480       .SelText = Space$(30) & "Glucose Tolerance Test"
490       .SelText = vbCrLf
500       .SelColor = vbBlack
510       .SelFontSize = 10
520       .SelText = vbCrLf
530       .SelText = vbCrLf
540       .SelText = Space$(25) & "Sample #      Date/Time" & Space$(10) & "Serum mmol/L"
550       .SelText = vbCrLf
560       .SelText = vbCrLf

570       Do While Not tbDems.EOF
580           sql = "Select * from BioResults where " & _
                    "SampleID = '" & tbDems!SampleID & "' " & _
                    "and Code = '" & CodeForGlucose & "'"
590           Set tbRes = New Recordset
600           RecOpenClient 0, tbRes, sql

610           If Not tbRes.EOF Then
620               LogBioAsPrinted tbDems!SampleID & "", CodeForGlucose
630               .SelText = Space$(25) & Left$(tbRes!SampleID & Space$(14), 14)
640               .SelText = Format(tbDems!SampleDate & "", "dd/mm/yyyy hh:mm")
650               .SelText = Space$(7) & Format(tbRes!Result, "0.0")
660               .SelText = vbCrLf
670               .SelText = vbCrLf
680           End If
690           tbDems.MoveNext
700       Loop

          Dim OBs As Observations

710       Set OBs = New Observations
720       Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
730       If Not OBs Is Nothing Then
740           FillCommentLines OBs(1).Comment, 4, Comments(), 97
750           For n = 1 To 4
760               .SelColor = vbBlack
770               .SelText = Comments(n) & vbCrLf
780           Next
790       End If

800       Set OBs = New Observations
810       Set OBs = OBs.Load(RP.SampleID, "Demographic")
820       If Not OBs Is Nothing Then
830           FillCommentLines OBs(1).Comment, 4, Comments(), 97
840           For n = 1 To 4
850               .SelColor = vbBlack
860               .SelText = Comments(n) & vbCrLf
870           Next
880       End If
890       RTFPrintFooter "Gluc. Tolerance", RP.Initiator, SampleDate, Rundate, ReceivedDate

900       PrintAndStore "Biochemistry", "BGlu"

910   End With

920   Exit Sub

RTFPrintGTT_Error:

      Dim strES As String
      Dim intEL As Integer

930   intEL = Erl
940   strES = Err.Description
950   LogError "modRTF", "RTFPrintGTT", intEL, strES, sql

End Sub

Public Sub RTFPrintUPro()

      Dim tb As Recordset
      Dim tU As Recordset
      Dim sql As String
      Dim Sex As String
      Dim n As Integer
      Dim OBs As Observations

10    On Error GoTo RTFPrintUPro_Error

20    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim ReceivedDate As String
      Dim Rundate As String
      Dim Dob As String
      Dim RunTime As String

30    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql

60    If tb.EOF Then
70        Exit Sub
80    End If

90    If IsDate(tb!Dob) Then
100       Dob = Format(tb!Dob, "dd/mmm/yyyy")
110   Else
120       Dob = ""
130   End If

140   sql = "SELECT * FROM UPro WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
150   Set tU = New Recordset
160   RecOpenServer 0, tU, sql
170   If tU.EOF Then
180       Exit Sub
190   End If

200   If Not SetPrinter("CHBIO") Then Exit Sub

210   RTFPrintHeading "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
                      tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

220   With frmMain.rtb

230       .SelFontSize = 10
240       .SelText = vbCrLf

250       .SelText = "Urinary Protein"
260       .SelText = vbCrLf
270       .SelText = vbCrLf
280       .SelText = Left$("Collection Period:" & Space$(25), 25) & tU!CollectionPeriod & " Hours"
290       .SelText = vbCrLf
300       .SelText = vbCrLf
310       .SelText = Left$("Volume Collected:" & Space$(25), 25) & tU!totalVolume & " ml"
320       .SelText = vbCrLf
330       .SelText = vbCrLf
340       .SelText = Left$("Urinary Protein:" & Space$(25), 25) & tU!UPgPerL & " g/l"
350       .SelText = vbCrLf
360       .SelText = vbCrLf
370       .SelText = Space$(25) & tU!UP24H & " g/24Hr"
380       .SelText = vbCrLf
390       .SelText = vbCrLf
400       .SelText = vbCrLf
410       .SelText = vbCrLf

420       Set OBs = New Observations
430       Set OBs = OBs.Load(RP.SampleID, "Demographic")
440       If Not OBs Is Nothing Then
450           FillCommentLines OBs(1).Comment, 4, Comments(), 97
460           For n = 1 To 4
470               If Trim$(Comments(n)) <> "" Then
480                   .SelColor = vbBlack
490                   .SelText = Comments(n) & vbCrLf
500               End If
510           Next
520       End If

530       Set OBs = New Observations
540       Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
550       If Not OBs Is Nothing Then
560           FillCommentLines OBs(1).Comment, 4, Comments(), 97
570           For n = 1 To 4
580               If Trim$(Comments(n)) <> "" Then
590                   .SelColor = vbBlack
600                   .SelText = Comments(n) & vbCrLf
610               End If
620           Next
630       End If

640       If IsDate(tb!SampleDate) Then
650           SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy")
660       Else
670           SampleDate = ""
680       End If
690       If IsDate(tb!RecDate) Then
700           ReceivedDate = Format(tb!RecDate, "dd/mmm/yyyy")
710       Else
720           ReceivedDate = ""
730       End If
740       If IsDate(RunTime) Then
750           Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
760       Else
770           If IsDate(tb!Rundate) Then
780               Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
790           Else
800               Rundate = ""
810           End If
820       End If

830       RTFPrintFooter "Biochemistry", RP.Initiator, SampleDate, Rundate, ReceivedDate

840       PrintAndStore "Biochemistry", "BUPro"

850   End With

860   Exit Sub

RTFPrintUPro_Error:

      Dim strES As String
      Dim intEL As Integer

870   intEL = Erl
880   strES = Err.Description
890   LogError "modRTF", "RTFPrintUPro", intEL, strES, sql

End Sub
Public Sub RTFPrintGlucoseSeries()

      Dim tb As Recordset
      Dim tbDems As Recordset
      Dim tbRes As Recordset
      Dim sql As String
      Dim Dob As String
      Dim NameToFind As String
      Dim SampleDate As String
      Dim ReceivedDate As String
      Dim Rundate As String
      Dim Code As Integer
10    On Error GoTo RTFPrintGlucoseSeries_Error

20    ReDim Comments(1 To 4) As String
      Dim n As Integer

30    Code = GetOptionSetting("BioCodeForGlucose", "1069")

40    sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql
70    If tb.EOF Then Exit Sub

80    If IsDate(tb!SampleDate) Then
90        SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
100   Else
110       SampleDate = ""
120   End If
130   If IsDate(tb!RecDate) Then
140       ReceivedDate = Format(tb!RecDate, "dd/mmm/yyyy hh:mm")
150   Else
160       ReceivedDate = ""
170   End If

180   If IsDate(tb!Rundate) Then
190       Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
200   Else
210       Rundate = ""
220   End If

230   NameToFind = tb!PatName & ""

240   If IsDate(tb!Dob) Then
250       Dob = Format(tb!Dob, "dd/mmm/yyyy")
260   Else
270       Dob = ""
280   End If

290   If Not SetPrinter("CHBIO") Then Exit Sub

300   sql = "select * from demographics where " & _
            "patname = '" & NameToFind & "' " & _
            "and rundate = '" & Format(tb!Rundate, "dd/mmm/yyyy") & "' "
310   If IsDate(Dob) Then
320       sql = sql & "and DoB = '" & Format(Dob, "dd/mmm/yyyy") & "' "
330   End If
340   sql = sql & "order by SampleDate"
350   Set tbDems = New Recordset
360   RecOpenClient 0, tbDems, sql

370   If tbDems.EOF Then
380       Exit Sub
390   End If

400   RTFPrintHeading "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
                      tb!Addr0 & "", tb!Addr1 & "", tb!Sex & "", tb!Hospital & ""

410   With frmMain.rtb
420       .SelFontSize = 10

430       .SelText = vbCrLf
440       .SelText = vbCrLf
450       .SelText = vbCrLf

460       .SelColor = vbGreen
470       .SelFontSize = 12
480       .SelText = Space$(20) & "Glucose Series"
490       .SelText = vbCrLf
500       .SelColor = vbBlack
510       .SelFontSize = 10
520       .SelText = vbCrLf
530       .SelText = vbCrLf
540       .SelText = Space$(25) & "Sample #      Date/Time" & Space$(10) & "Serum mmol/L"
550       .SelText = vbCrLf
560       .SelText = vbCrLf

570       Do While Not tbDems.EOF
580           sql = "Select * from BioResults where " & _
                    "SampleID = '" & tbDems!SampleID & "' " & _
                    "and Code = '" & Code & "'"
590           Set tbRes = New Recordset
600           RecOpenClient 0, tbRes, sql

610           If Not tbRes.EOF Then
620               LogBioAsPrinted tbDems!SampleID & "", Code
630               .SelText = Space$(25) & Left$(tbRes!SampleID & Space$(12), 12)
640               .SelText = Format(tbDems!SampleDate & "", "dd/mm/yyyy hh:mm")
650               .SelText = Space$(7) & Format(tbRes!Result, "0.0")
660               .SelText = vbCrLf
670               .SelText = vbCrLf
680           End If
690           tbDems.MoveNext
700       Loop

          Dim OBs As Observations

710       Set OBs = New Observations
720       Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
730       If Not OBs Is Nothing Then
740           FillCommentLines OBs(1).Comment, 4, Comments(), 97
750           For n = 1 To 4
760               .SelColor = vbBlack
770               .SelText = Comments(n) & vbCrLf
780           Next
790       End If

800       Set OBs = New Observations
810       Set OBs = OBs.Load(RP.SampleID, "Demographic")
820       If Not OBs Is Nothing Then
830           FillCommentLines OBs(1).Comment, 4, Comments(), 97
840           For n = 1 To 4
850               .SelColor = vbBlack
860               .SelText = Comments(n) & vbCrLf
870           Next
880       End If

890       RTFPrintFooter "Glucose Series", RP.Initiator, SampleDate, Rundate, ReceivedDate

900       PrintAndStore "Biochemistry", "Bglu"

910   End With

920   Exit Sub

RTFPrintGlucoseSeries_Error:

      Dim strES As String
      Dim intEL As Integer

930   intEL = Erl
940   strES = Err.Description
950   LogError "modRTF", "RTFPrintGlucoseSeries", intEL, strES, sql

End Sub

Private Sub PrintAndStore(ByVal Department As String, ByVal Dept As String)

Dim sql        As String
Dim Gx         As New GP
Dim PrinterName As String

On Error GoTo PrintAndStore_Error

If InStr(frmMain.rtb.TextRTF, RP.SampleID) Then    'Double check that report saved is for this Sample Id
    
    If Not CheckDisablePrinting(RP.GP, Department) And Not CheckDisablePrinting(RP.Ward, Department) Then       'if printing is not disabled
    '***PRINT REPORT
        PrinterName = Printer.DeviceName
        If UCase(RP.Initiator) <> "BIOMNIS IRELAND" Then
            frmMain.rtb.SelStart = 0
            Gx.LoadName RP.GP
            If (Gx.PrintReport And RP.Ward = "GP") Or RP.Ward <> "GP" Then
                frmMain.rtb.SelPrint Printer.hDC
            Else
                PrinterName = "None"
            End If
        End If
    End If
    '***SAVE REPORT
    sql = "INSERT INTO Reports " & _
          "(SampleID, Dept, Initiator, PrintTime, ReportNumber, PageNumber, Report, Printer) " & _
          "VALUES " & _
          "( '" & RP.SampleID & "', " & _
          "  '" & Department & "', " & _
          "  '" & RP.Initiator & "', " & _
          "   getdate(), " & _
          "  '" & RP.SampleID & Dept & "', " & _
          "  '1', " & _
          "  '" & AddTicks(frmMain.rtb.TextRTF) & "', " & _
          "  '" & PrinterName & "')"
    Cnxn(0).Execute sql
Else
    LogError "modRTF", "PrintAndStore", 134, RP.SampleID & " not found in RTF report", sql
End If

Exit Sub

PrintAndStore_Error:

Dim strES      As String
Dim intEL      As Integer

intEL = Erl
strES = Err.Description
LogError "modRTF", "PrintAndStore", intEL, strES, sql

End Sub

Public Sub RTFPrintNoSexDoB(ByVal Sex As String, ByVal Dob As String)

10    On Error GoTo RTFPrintNoSexDoB_Error

20    With frmMain.rtb
30        If Not IsDate(Dob) And Trim$(Sex) = "" Then
40            .SelColor = vbBlue
50            .SelText = Space$(24) & "No Sex/DoB given. Normal ranges may not be relevant"
60            .SelText = vbCrLf
70            .SelFontSize = 10
80        ElseIf Not IsDate(Dob) Then
90            .SelColor = vbBlue
100           .SelText = Space$(24) & "No DoB given. Normal ranges may not be relevant"
110           .SelText = vbCrLf
120           .SelFontSize = 10
130       ElseIf Trim$(Sex) = "" Then
140           .SelColor = vbBlue
150           .SelText = Space$(24) & "No Sex given. Normal ranges may not be relevant"
160           .SelText = vbCrLf
170           .SelFontSize = 10
180       End If
190   End With

200   Exit Sub

RTFPrintNoSexDoB_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "modRTF", "RTFPrintNoSexDoB", intEL, strES

End Sub

Public Sub RTFPrintHeading(ByVal Dept As String, _
                           ByVal Name As String, _
                           ByVal Dob As String, _
                           ByVal Chart As String, _
                           ByVal Address0 As String, _
                           ByVal Address1 As String, _
                           ByVal Sex As String, _
                           ByVal Hospital As String, _
                           Optional TotalPage As Integer = 1, _
                           Optional CurrentPage As Integer = 1)

      Dim S As String
      Dim n As Integer
      Dim BioPhone As String
      Dim HaemPhone As String
      Dim AccreditationText As String
      Dim AccreditationDept As String
10    On Error GoTo RTFPrintHeading_Error

20    Printer.Print "";

30    With frmMain.rtb
40        AccreditationDept = Dept
50        .TextRTF = ""
60        .SelFontName = "Courier New"
70        .SelFontSize = 16
80        .SelBold = True

90        Select Case Dept
          Case "Haematology":
100           .SelColor = vbRed
110       Case "Biochemistry":
120           .SelColor = vbGreen
130       Case "Blood Transfusion":
140           .SelColor = vbRed
150       Case "Coagulation":
160           .SelColor = vbRed
170       Case "Microbiology"
180           .SelColor = vbBlue
190       End Select
200       Dept = Dept & " Laboratory"
210       .SelText = "CAVAN GENERAL HOSPITAL : " & Dept
220       .SelFontSize = 10
          'Printer.CurrentY = 100

230       Select Case Dept
          Case "Haematology Laboratory":
240           HaemPhone = GetOptionSetting("HaemPhone", "")
250           If HaemPhone <> "" Then
260               .SelText = " Phone " & HaemPhone
270           End If
280       Case "Biochemistry Laboratory":
290           BioPhone = GetOptionSetting("BioPhone", "")
300           If BioPhone <> "" Then
310               .SelText = " Phone " & BioPhone
320           End If
330       Case "Blood Transfusion Laboratory":
340           .SelText = " Phone 38830"
350       Case "Microbiology Laboratory":
360           .SelText = " 049 4376053"
370       End Select
380       .SelText = vbCrLf

          'Printer.CurrentY = 320




390       .SelFontSize = 4
400       .SelFontName = "Courier New"
410       .SelAlignment = rtfCenter
420       .SelBold = False
430       If gPrintCopyReport = 0 Then
440           .SelText = String$(200, "-")
450       Else
460           S = "-- THIS IS A COPY REPORT -- NOT FOR FILING --"
470           .SelColor = vbRed
480           For n = 1 To 5
490               .SelText = S
500           Next
510       End If
520       .SelText = vbCrLf


          '    Select Case AccreditationDept
          '    Case "Haematology":
          '        AccreditationText = GetOptionSetting("HaemAccreditation", "")
          '    Case "Biochemistry":
          '        AccreditationText = GetOptionSetting("BioAccreditation", "")
          '    Case "Coagulation":
          '        AccreditationText = GetOptionSetting("CoagAccreditation", "")
          '    Case "Microbiology":
          '        AccreditationText = GetOptionSetting("MicroAccreditation", "")
          '    End Select
          '
          '    If AccreditationText <> "" Then
          '        .SelColor = vbRed
          '        .SelAlignment = rtfLeft
          '
          '        .SelFontName = "Courier New"
          '        .SelFontSize = 10
          '        .SelBold = True
          '
          '        .SelText = AccreditationText
          '        .SelText = vbCrLf
          '    End If

530       .SelColor = vbBlack
540       .SelAlignment = rtfLeft

550       .SelFontName = "Courier New"
560       .SelFontSize = 12
570       .SelBold = False

580       .SelText = " Sample ID:"
590       .SelText = Left$(RP.SampleID & Space$(10), 10)

600       .SelText = Space$(14) & "Name:"
610       .SelBold = True
620       .SelFontSize = 14
630       .SelText = Left$(Name, 27)
640       .SelFontSize = 12
650       .SelBold = False
660       .SelText = vbCrLf

670       .SelText = "      Ward:"
680       .SelText = Left$(RP.Ward & Space$(20), 20)

690       .SelText = Space$(4) & " DOB:"
700       .SelText = Format(Dob, "dd/mm/yyyy")

710       .SelText = Space$(11) & "Chart:"
720       .SelText = Left$(Hospital & " ", 1) & " "
730       .SelText = Chart
740       .SelText = vbCrLf

750       .SelText = "Consultant:"
760       .SelText = Left$(RP.Clinician & Space$(23), 23)
770       .SelText = " Addr:"
780       .SelText = Left$(Address0 & Space$(21), 21)
790       .SelText = "  Sex:"
800       Select Case Left$(UCase$(Trim$(Sex)), 1)
          Case "M": .SelText = "Male"
810       Case "F": .SelText = "Female"
820       End Select
830       .SelText = vbCrLf

840       .SelText = "        GP:"
850       .SelText = Left$(RP.GP & Space$(25), 25)
860       .SelText = "    " & Address1
870       .SelText = vbCrLf

880       .SelFontSize = 4
890       .SelFontName = "Courier New"
900       .SelAlignment = rtfCenter
910       .SelBold = False
920       If gPrintCopyReport = 0 Then
930           .SelText = String$(200, "-")
940       Else
950           S = "-- THIS IS A COPY REPORT -- NOT FOR FILING --"
960           .SelColor = vbRed
970           For n = 1 To 5
980               .SelText = S
990           Next
1000      End If
1010      .SelText = vbCrLf
1020      .SelFontSize = 12
1030      .SelFontName = "Courier New"
1040      .SelBold = True
1050      .SelText = Left(IIf((UCase(RP.PrintAction) = UCase("save")), "Interim Report", "Final Report") & Space$(55), 55) & "Page " & CurrentPage & " of " & TotalPage
1060      .SelText = vbCrLf
1070      .SelFontSize = 4
1080      .SelFontName = "Courier New"
1090      .SelAlignment = rtfCenter
1100      .SelBold = False
1110      .SelText = String$(200, "-")
1120      .SelText = vbCrLf
1130      .SelAlignment = rtfLeft
1140      .SelColor = vbBlack
1150      .SelBold = False

1160  End With

1170  Exit Sub

RTFPrintHeading_Error:

      Dim strES As String
      Dim intEL As Integer

1180  intEL = Erl
1190  strES = Err.Description
1200  LogError "modRTF", "RTFPrintHeading", intEL, strES


End Sub


