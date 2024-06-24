Attribute VB_Name = "modWord"
'Option Explicit
'
'
'Public Sub WordPrintResultHaem(ByRef X As Word.Document)
'
'      Dim tb As Recordset
'      Dim tbh As Recordset
'      Dim n As Integer
'      Dim Sex As String
'      Dim blnFBC As Boolean
'      Dim TotalRetics As Long
'      Dim Dob As String
'      Dim Flag As String
'      Dim sql As String
'      Dim OBs As Observations
'      'Set x = CreateObject(word.Document)
'10    On Error GoTo WordPrintResultHaem_Error
'
'20    ReDim Comments(1 To 4) As String
'      Dim SampleDate As String
'      Dim Rundate As String
'
'30    sql = "Select * from Demographics where " & _
'            "SampleID = '" & RP.SampleID & "'"
'40    Set tb = New Recordset
'50    RecOpenClient 0, tb, sql
'60    If tb.EOF Then Exit Sub
'
'70    If IsDate(tb!SampleDate) Then
'80      SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
'90    Else
'100     SampleDate = ""
'110   End If
'120   If IsDate(tb!Rundate) Then
'130     Rundate = Format(tb!Rundate, "dd/mmm/yyyy hh:mm")
'140   Else
'150     Rundate = ""
'160   End If
'
'170   sql = "SELECT * FROM HaemResults WHERE " & _
'            "SampleID = '" & RP.SampleID & "' " & _
'            "AND Valid = 1"
'180   Set tbh = New Recordset
'190   RecOpenClient 0, tbh, sql
'200   If tbh.EOF Then Exit Sub
'
'210   Dob = tb!Dob & ""
'
'220   blnFBC = Trim$(tbh!wbc & "") <> ""
'
'230   Select Case Left$(UCase$(tb!Sex & ""), 1)
'        Case "M": Sex = "M"
'240     Case "F": Sex = "F"
'250     Case Else: Sex = ""
'260   End Select
'      '-------------------------------------
'
'270   With X.ActiveWindow.Selection
'280     .Font.Name = "Courier New"
'290     .Font.Size = 16
'300     .Font.Bold = True
'310     .Font.Color = wdColorDarkRed
'320     .TypeText "CAVAN GENERAL HOSPITAL : Haematology Laboratory"
'330     .Font.Size = 10
'340     .TypeText " Phone " & GetOptionSetting("HaemPhone", "")
'350     .TypeText vbCrLf
'
'360     .Font.Size = 4
'370     .TypeText String$(250, "-") & vbCrLf
'
'380     .Font.Color = wdColorBlack
'
'390     .Font.Name = "Courier New"
'400     .Font.Size = 12
'410     .Font.Bold = False
'
'420     .TypeText "Sample ID:"
'430     .Font.Bold = True
'440     .TypeText Left$(RP.SampleID & Space$(25), 25)
'450     .Font.Bold = False
'460     .TypeText "Name:"
'470     .Font.Bold = True
'480     .Font.Size = 14
'490     .TypeText Left$(tb!PatName & "", 27) & vbCrLf
'
'500     .Font.Size = 12
'510     .Font.Bold = False
'520     .TypeText "     Ward:"
'530     .Font.Bold = True
'540     .TypeText Left$(RP.Ward & Space$(25), 25)
'550     .Font.Bold = False
'560     .TypeText " DOB:"
'570     .Font.Bold = True
'580     .TypeText Left$(Format(Dob, "dd/mm/yyyy") & Space$(21), 21)
'590     .Font.Bold = False
'600     .TypeText "Chart #:"
'610     .Font.Bold = True
'620     .TypeText tb!Chart & vbCrLf
'
'630     .Font.Bold = False
'
'640     .TypeText "Clinician:"
'650     .Font.Bold = True
'660     .TypeText Left$(RP.Clinician & Space$(25), 25)
'670     .Font.Bold = False
'680     .TypeText "Addr:"
'690     .Font.Bold = True
'700     .TypeText Left$(tb!Addr0 & Space$(25), 25)
'710     .Font.Bold = False
'720     .TypeText "Sex:"
'730     Select Case Left$(UCase$(Trim$(tb!Sex & "")), 1)
'          Case "M": .TypeText "Male"
'740       Case "F": .TypeText "Female"
'750     End Select
'760     .TypeText vbCrLf
'
'770     .TypeText "       GP:"
'780     .Font.Bold = True
'790     .TypeText Left$(RP.GP & Space$(30), 30)
'800     .Font.Bold = False
'810     .TypeText tb!Addr1 & vbCrLf
'
'
'820     .Font.Size = 4
'830     .TypeText String$(250, "-") & vbCrLf
'
'      'End With
'      '-------------------------------------------
'      'WordPrintHeading x, "Haematology", tb!PatName & "", DoB, tb!Chart & "", _
'       '                tb!Addr0 & "", tb!Addr1 & "", Sex
'
'      'MsgBox "Heading printed"
'
'840     .Font.Name = "Courier New"
'850     .Font.Size = 10
'
'860     If blnFBC Then
'870       Flag = InterpH(Val(tbh!wbc & ""), "WBC", Sex, Dob)
'880       .TypeText "          WBC   "
'890       If Flag <> "X" Then
'900         .TypeText Right("     " & tbh!wbc & "", 5)
'910       Else
'920         .TypeText "XXXXX"
'930       End If
'940       .TypeText " x10"
'950       .Font.Superscript = True
'960       .TypeText "9"
'970       .Font.Superscript = False
'980       .Font.Size = 10
'990       .TypeText "/l "
'1000      .Font.Superscript = True
'1010      .TypeText " "
'1020      .Font.Superscript = False
'1030      .Font.Bold = True
'1040      .TypeText Flag
'1050      .Font.Bold = False
'1060      .TypeText "  " & HaemNormalRange("WBC", Sex, Dob)
'
'1070      .TypeText "  "
'1080      Flag = InterpH(Val(tbh!NeutP & ""), "NEUTP", Sex, Dob)
'1090      .TypeText "Neut  "
'1100      If Flag <> "X" Then
'1110        .TypeText Right("    " & tbh!NeutP & "", 4)
'1120      Else
'1130        .TypeText "XXXX"
'1140      End If
'1150      .TypeText "% = "
'1160      Flag = InterpH(Val(tbh!neuta & ""), "NEUTA", Sex, Dob)
'1170      If Flag <> "X" Then
'1180        .TypeText Right("     " & tbh!neuta & "", 5)
'1190      Else
'1200        .TypeText "XXXXX"
'1210      End If
'1220      .TypeText " x10"
'1230      .Font.Superscript = True
'1240      .TypeText "9"
'1250      .Font.Superscript = False
'1260      .Font.Size = 10
'1270      .TypeText "/l "
'1280      .Font.Superscript = True
'1290      .TypeText " "
'1300      .Font.Superscript = False
'1310      .Font.Bold = True
'1320      .TypeText Flag
'1330      .Font.Bold = False
'1340      .TypeText HaemNormalRange("NEUTA", Sex, Dob)
'1350      .TypeText vbCrLf
'
'1360      Flag = InterpH(Val(tbh!LymP & ""), "LYMP", Sex, Dob)
'1370      .TypeText Space$(45) & "Lymph "
'1380      If Flag <> "X" Then
'1390        .TypeText Right("    " & tbh!LymP & "", 4)
'1400      Else
'1410        .TypeText "XXXX"
'1420      End If
'1430      .TypeText "% = "
'1440      Flag = InterpH(Val(tbh!lyma & ""), "LYMA", Sex, Dob)
'1450      If Flag <> "X" Then
'1460        .TypeText Right("     " & tbh!lyma & "", 5)
'1470      Else
'1480        .TypeText "XXXXX"
'1490      End If
'1500      .TypeText " x10"
'1510      .Font.Superscript = True
'1520      .TypeText "9"
'1530      .Font.Superscript = False
'1540      .TypeText "/l "
'1550      .Font.Superscript = True
'1560      .TypeText " "
'1570      .Font.Superscript = False
'1580      .Font.Bold = True
'1590      .TypeText Flag
'1600      .Font.Bold = False
'1610      .TypeText HaemNormalRange("LYMA", Sex, Dob)
'1620      .TypeText vbCrLf
'
'
'1630      Flag = InterpH(Val(tbh!rbc & ""), "RBC", Sex, Dob)
'1640      .TypeText "          RBC   "
'1650      If Flag <> "X" Then
'1660        .TypeText Right("     " & tbh!rbc & "", 5)
'1670      Else
'1680        .TypeText "XXXXX"
'1690      End If
'1700      .TypeText " x10"
'1710      .Font.Superscript = True
'1720      .TypeText "12"
'1730      .Font.Superscript = False
'1740      .TypeText "/l "
'1750      .Font.Bold = True
'1760      .TypeText Flag
'1770      .Font.Bold = False
'1780      .TypeText "  " & HaemNormalRange("RBC", Sex, Dob)
'1790      Flag = InterpH(Val(tbh!MonoP & ""), "MONOP", Sex, Dob)
'1800      .TypeText "  "
'1810      .TypeText "Mono  "
'1820      If Flag <> "X" Then
'1830        .TypeText Right("    " & tbh!MonoP & "", 4)
'1840      Else
'1850        .TypeText "XXXX"
'1860      End If
'1870      .TypeText "% = "
'1880      Flag = InterpH(Val(tbh!monoa & ""), "MONOA", Sex, Dob)
'1890      If Flag <> "X" Then
'1900        .TypeText Right("     " & tbh!monoa & "", 5)
'1910      Else
'1920        .TypeText "XXXXX"
'1930      End If
'1940      .TypeText " x10"
'1950      .Font.Superscript = True
'1960      .TypeText "9"
'1970      .Font.Superscript = False
'1980      .TypeText "/l "
'1990      .Font.Superscript = True
'2000      .TypeText " "
'2010      .Font.Superscript = False
'2020      .Font.Bold = True
'2030      .TypeText Flag
'2040      .Font.Bold = False
'2050      .TypeText HaemNormalRange("MONOA", Sex, Dob)
'2060      .TypeText vbCrLf
'
'
'2070      Flag = InterpH(Val(tbh!eosP & ""), "EOSP", Sex, Dob)
'2080      .TypeText Space$(45) & "Eos   "
'2090      If Flag <> "X" Then
'2100        .TypeText Right("    " & tbh!eosP & "", 4)
'2110      Else
'2120        .TypeText "XXXX"
'2130      End If
'2140      .TypeText "% = "
'2150      Flag = InterpH(Val(tbh!eosa & ""), "EOSA", Sex, Dob)
'2160      If Flag <> "X" Then
'2170        .TypeText Right("     " & tbh!eosa & "", 5)
'2180      Else
'2190        .TypeText "XXXXX"
'2200      End If
'2210      .TypeText " x10"
'2220      .Font.Superscript = True
'2230      .TypeText "9"
'2240      .Font.Superscript = False
'2250      .TypeText "/l "
'2260      .Font.Superscript = True
'2270      .TypeText " "
'2280      .Font.Superscript = False
'2290      .Font.Bold = True
'2300      .TypeText Flag
'2310      .Font.Bold = False
'2320      .TypeText HaemNormalRange("EOSA", Sex, Dob)
'2330      .TypeText vbCrLf
'
'2340      Flag = InterpH(Val(tbh!Hgb & ""), "Hgb", Sex, Dob)
'2350      .TypeText "          Hgb   "
'2360      If Flag <> "X" Then
'2370        .TypeText Right("     " & tbh!Hgb & "", 5)
'2380      Else
'2390        .TypeText "XXXXX"
'2400      End If
'2410      .TypeText " g/dl  "
'2420      .Font.Superscript = True
'2430      .TypeText "  "
'2440      .Font.Superscript = False
'2450      .Font.Bold = True
'2460      .TypeText Flag
'2470      .Font.Bold = False
'2480      .TypeText "  " & HaemNormalRange("Hgb", Sex, Dob)
'2490      Flag = InterpH(Val(tbh!basP & ""), "BASP", Sex, Dob)
'2500      .TypeText "  " & "Bas   "
'2510      If Flag <> "X" Then
'2520        .TypeText Right("    " & tbh!basP & "", 4)
'2530      Else
'2540        .TypeText "XXXX"
'2550      End If
'2560      .TypeText "% = "
'2570      Flag = InterpH(Val(tbh!basa & ""), "BASA", Sex, Dob)
'2580      If Flag <> "X" Then
'2590        .TypeText Right("     " & tbh!basa & "", 5)
'2600      Else
'2610        .TypeText "XXXXX"
'2620      End If
'2630      .TypeText " x10"
'2640      .Font.Superscript = True
'2650      .TypeText "9"
'2660      .Font.Superscript = False
'2670      .TypeText "/l "
'2680      .Font.Superscript = True
'2690      .TypeText " "
'2700      .Font.Superscript = False
'2710      .Font.Bold = True
'2720      .TypeText Flag
'2730      .Font.Bold = False
'2740      .TypeText HaemNormalRange("BASA", Sex, Dob)
'2750      .TypeText vbCrLf
'
'2760      .TypeText vbCrLf
'
'2770      Flag = InterpH(Val(tbh!hct & ""), "Hct", Sex, Dob)
'2780      .TypeText "          Hct   "
'2790      If Flag <> "X" Then
'2800        .TypeText Right("     " & tbh!hct & "", 5)
'2810      Else
'2820        .TypeText "XXXXX"
'2830      End If
'2840      .TypeText " l/l   "
'2850      .Font.Superscript = True
'2860      .TypeText "  "
'2870      .Font.Superscript = False
'2880      .Font.Bold = True
'2890      .TypeText Flag
'2900      .Font.Bold = False
'2910      .TypeText "  " & HaemNormalRange("Hct", Sex, Dob)
'2920      .TypeText vbCrLf
'2930    End If
'
'2940    .TypeText "  " & tbh!md0 & vbCrLf
'
'2950    If blnFBC Then
'2960      Flag = InterpH(Val(tbh!mcv & ""), "MCV", Sex, Dob)
'2970      .TypeText "          MCV   "
'2980      If Flag <> "X" Then
'2990        .TypeText Right("     " & tbh!mcv & "", 5)
'3000      Else
'3010        .TypeText "XXXXX"
'3020      End If
'3030      .TypeText " fl    "
'3040      .Font.Superscript = True
'3050      .TypeText "  "
'3060      .Font.Superscript = False
'3070      .Font.Bold = True
'3080      .TypeText Flag
'3090      .Font.Bold = False
'3100      .TypeText "  " & HaemNormalRange("MCV", Sex, Dob)
'3110      .TypeText "  " & tbh!md1 & vbCrLf
'3120    End If
'
'3130    .TypeText Space$(45) & tbh!md2 & vbCrLf
'
'3140    If blnFBC Then
'3150      Flag = InterpH(Val(tbh!mch & ""), "MCH", Sex, Dob)
'3160      .TypeText "          MCH   "
'3170      If Flag <> "X" Then
'3180        .TypeText Right("     " & tbh!mch & "", 5)
'3190      Else
'3200        .TypeText "XXXXX"
'3210      End If
'3220      .TypeText " pg    "
'3230      .Font.Superscript = True
'3240      .TypeText "  "
'3250      .Font.Superscript = False
'3260      .Font.Bold = True
'3270      .TypeText Flag
'3280      .Font.Bold = False
'3290      .TypeText "  " & HaemNormalRange("MCH", Sex, Dob)
'3300      .TypeText "  " & tbh!md3 & vbCrLf
'3310    End If
'
'3320    .TypeText Space$(45) & tbh!md4 & vbCrLf
'
'3330    If blnFBC Then
'3340      Flag = InterpH(Val(tbh!mchc & ""), "MCHC", Sex, Dob)
'3350      .TypeText "          MCHC  "
'3360      If Flag <> "X" Then
'3370        .TypeText Right("     " & tbh!mchc & "", 5)
'3380      Else
'3390        .TypeText "XXXXX"
'3400      End If
'3410      .TypeText " g/dl  "
'3420      .Font.Superscript = True
'3430      .TypeText "  "
'3440      .Font.Superscript = False
'3450      .Font.Bold = True
'3460      .TypeText Flag
'3470      .Font.Bold = False
'3480      .TypeText "  " & HaemNormalRange("MCHC", Sex, Dob)
'3490      .TypeText "  " & tbh!md5 & vbCrLf
'3500    End If
'
'
'3510    If Trim$(tbh!esr & "") <> "" And Trim$(tbh!Malaria & "") <> "" Then
'3520      Flag = InterpH(Val(tbh!esr & ""), "ESR", Sex, Dob)
'3530      .TypeText "ESR "
'3540      If Flag <> "X" Then
'3550        .TypeText tbh!esr
'3560      Else
'3570        .TypeText "XXX"
'3580      End If
'3590      .TypeText " mm/hr "
'3600      .Font.Bold = True
'3610      .TypeText Flag
'3620      .Font.Bold = False
'3630      .TypeText "  " & HaemNormalRange("ESR", Sex, Dob)
'3640      .TypeText "  Malaria Screening Kit = " & tbh!Malaria & vbCrLf
'3650    ElseIf Trim$(tbh!esr & "") <> "" Then
'3660      Flag = InterpH(Val(tbh!esr & ""), "ESR", Sex, Dob)
'3670      .TypeText Space$(45) & "ESR      "
'3680      If Flag <> "X" Then
'3690        .TypeText tbh!esr
'3700      Else
'3710        .TypeText "XXX"
'3720      End If
'3730      .TypeText " mm/hr "
'3740      .Font.Bold = True
'3750      .TypeText Flag
'3760      .Font.Bold = False
'3770      .TypeText HaemNormalRange("ESR", Sex, Dob) & vbCrLf
'3780    ElseIf Trim$(tbh!Malaria & "") <> "" Then
'3790      .TypeText "          Malaria Screening Kit = " & tbh!Malaria & vbCrLf
'3800    Else
'3810      .TypeText vbCrLf
'3820    End If
'
'3830    If blnFBC Then
'3840      Flag = InterpH(Val(tbh!rdwcv & ""), "RDWCV", Sex, Dob)
'3850      .TypeText "          RDW   "
'3860      If Flag <> "X" Then
'3870        .TypeText Right("     " & tbh!rdwcv & "", 5)
'3880      Else
'3890        .TypeText "XXXXX"
'3900      End If
'3910      .TypeText " %     "
'3920      .Font.Superscript = True
'3930      .TypeText "  "
'3940      .Font.Superscript = False
'3950      .Font.Bold = True
'3960      .TypeText Flag
'3970      .Font.Bold = False
'3980      .TypeText "  " & HaemNormalRange("RDWCV", Sex, Dob)
'3990      .TypeText "  "
'4000      If Trim$(tbh!RetP & "") <> "" Then
'4010        Flag = InterpH(TotalRetics, "RET", Sex, Dob)
'4020        .TypeText "Retics   "
'4030        If Flag <> "X" Then
'4040          .TypeText tbh!RetP
'4050        Else
'4060          .TypeText "XXX"
'4070        End If
'4080        .TypeText " %  Total Retics "
'4090        TotalRetics = tbh!RetA & "" 'Val((tbH!retics) * Val(tbH!RBC) * 10)
'4100        If Flag <> "X" Then
'4110          .TypeText Format(TotalRetics, "###0")
'4120        Else
'4130          .TypeText "XXXX"
'4140        End If
'4150        .TypeText "x10"
'4160        .Font.Superscript = True
'4170        .TypeText "9"
'4180        .Font.Superscript = False
'4190        .TypeText "/l "
'4200        .Font.Superscript = True
'4210        .TypeText " "
'4220        .Font.Superscript = False
'4230        .Font.Bold = True
'4240        .TypeText Flag
'4250        .Font.Bold = False
'4260        .TypeText HaemNormalRange("RET", Sex, Dob) & vbCrLf
'4270      Else
'4280        .TypeText vbCrLf
'4290      End If
'4300    End If
'
'4310    If Trim$(tbh!monospot & "") <> "" And Trim$(tbh!Sickledex & "") <> "" Then
'4320      .TypeText "Infectious Mono Screen  "
'4330      If tbh!monospot = "N" Then
'4340        .TypeText "Negative.  "
'4350      ElseIf tbh!monospot = "P" Then
'4360        .TypeText "Positive.  "
'4370      Else
'4380        .TypeText tbh!monospot & "  "
'4390      End If
'4400      .TypeText "Sickledex test for HbS = " & tbh!Sickledex & vbCrLf
'4410    ElseIf Trim$(tbh!monospot & "") <> "" Then
'4420      .TypeText Space$(45) & "Infectious Mono Screen  "
'4430      If tbh!monospot = "N" Then
'4440        .TypeText "Negative." & vbCrLf
'4450      ElseIf tbh!monospot = "P" Then
'4460        .TypeText "Positive." & vbCrLf
'4470      Else
'4480        .TypeText tbh!monospot & vbCrLf
'4490      End If
'4500    ElseIf Trim$(tbh!Sickledex & "") <> "" Then
'4510      .TypeText "Sickledex test for HbS = " & tbh!Sickledex & vbCrLf
'4520    Else
'4530      .TypeText vbCrLf
'4540    End If
'
'
'4550    If blnFBC Then
'4560      Flag = InterpH(Val(tbh!Plt & ""), "Plt", Sex, Dob)
'4570      .TypeText "          Plt   "
'4580      If Flag <> "X" Then
'4590        .TypeText Right("     " & tbh!Plt & "", 5)
'4600      Else
'4610        .TypeText "XXXXX"
'4620      End If
'4630      .TypeText " x10"
'4640      .Font.Superscript = True
'4650      .TypeText "9"
'4660      .Font.Superscript = False
'4670      .TypeText "/l "
'4680      .Font.Superscript = True
'4690      .TypeText " "
'4700      .Font.Superscript = False
'4710      .Font.Bold = True
'4720      .TypeText Flag
'4730      .Font.Bold = False
'4740      .TypeText "  " & HaemNormalRange("Plt", Sex, Dob) & vbCrLf
'4750    End If
'
'4760    .TypeText vbCrLf
'
'4770    Set OBs = New Observations
'4780    Set OBs = OBs.Load(RP.SampleID, "Haematology")
'4790    If Not OBs Is Nothing Then
'4800      FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
'4810      For n = 1 To 4
'4820        .TypeText Comments(n) & vbCrLf
'4830      Next
'4840    End If
'
'4850    Set OBs = New Observations
'4860    Set OBs = OBs.Load(RP.SampleID, "Demographic")
'4870    If Not OBs Is Nothing Then
'4880      FillCommentLines OBs.Item(1).Comment, 2, Comments(), 97
'4890      For n = 1 To 2
'4900        .TypeText Comments(n) & vbCrLf
'4910      Next
'4920    End If
'
'4930    Set OBs = New Observations
'4940    Set OBs = OBs.Load(RP.SampleID, "Film")
'4950    If Not OBs Is Nothing Then
'4960      FillCommentLines OBs.Item(1).Comment, 2, Comments(), 97
'4970      For n = 1 To 2
'4980        .TypeText Comments(n) & vbCrLf
'4990      Next
'5000    End If
'
'5010    If Not IsDate(tb!Dob) And Trim$(Sex) = "" Then
'5020      .Font.Color = vbBlue
'5030      .TypeText "No Sex/DoB given. Normal ranges may not be relevant" & vbCrLf
'5040    ElseIf Not IsDate(tb!Dob) Then
'5050      .Font.Color = vbBlue
'5060      .TypeText "No DoB given. Normal ranges may not be relevant" & vbCrLf
'5070    ElseIf Trim$(Sex) = "" Then
'5080      .Font.Color = vbBlue
'5090      .TypeText "No Sex given. Normal ranges may not be relevant" & vbCrLf
'5100    End If
'
'        'now the footer
'5110    .Font.Color = vbRed
'5120    .Font.Size = 4
'5130    .TypeText String$(250, "-") & vbCrLf
'5140    .Font.Size = 10
'5150    .Font.Bold = False
'5160    .TypeText "Sample Date:" & Format(SampleDate, "dd/mm/yyyy")
'5170    If Format(SampleDate, "hh:mm") <> "00:00" Then
'5180      .TypeText "   Sample Time:"
'5190      .TypeText Format(SampleDate, "hh:mm")
'5200    End If
'5210    .TypeText "  Tested:" & Format(Rundate, "dd/mm/yyyy hh:mm")
'5220    If Trim$(RP.Initiator) <> "" Then
'5230      .TypeText " Validated by " & TechnicianCodeFor(RP.Initiator)
'5240    End If
'
'5250  End With
'
'      'WordPrintFooter x, "Haematology", RP.Initiator, SampleDate, RunDate
'      '---------------------------
'
'5260  Exit Sub
'
'WordPrintResultHaem_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'5270  intEL = Erl
'5280  strES = Err.Description
'5290  LogError "modWord", "WordPrintResultHaem", intEL, strES, sql
'
'End Sub
'
'
'Public Sub WordPrintResultBioSideBySide(ByRef X As Word.Document)
'
'
'
'          Dim tb As Recordset
'          Dim tbF As Recordset
'          Dim tbUN As Recordset
'          Dim sql As String
'          Dim Sex As String
'          Dim lpc As Integer
'          Dim cUnits As String
'          Dim Flag As String
'          Dim n As Integer
'          Dim v As String
'          Dim Low As Single
'          Dim High As Single
'          Dim strLow As String * 4
'          Dim strHigh As String * 4
'          Dim BRs As New BIEResults
'          Dim BR As BIEResult
'          Dim TestCount As Integer
'          Dim SampleType As String
'          Dim ResultsPresent As Boolean
'          Dim OBs As Observations
'          Dim SampleDate As String
'          Dim Rundate As String
'          Dim Dob As String
'          Dim RunTime As String
'          Dim Fasting As String
'          Dim udtPrintLine(0 To 60) As PrintLine 'max 30 result lines per page
'          Dim strFormat As String
'          Dim MultiColumn As Boolean
'          Dim CodeForChol As String
'          Dim CodeForGlucose As String
'          Dim CodeForTrig As String
'10        'Set X = CreateObject(Word.Document)
'
'20        On Error GoTo WordPrintResultBioSideBySide_Error
'
'30        CodeForChol = UCase$(GetOptionSetting("BioCodeForChol", ""))
'40        CodeForGlucose = GetOptionSetting("BioCodeForGlucose", "")
'50        CodeForTrig = GetOptionSetting("BioCodeForTrig", "")
'
'60        ReDim Comments(1 To 4) As String
'
'70        For n = 0 To 60
'80            udtPrintLine(n).Analyte = ""
'90            udtPrintLine(n).Result = ""
'100           udtPrintLine(n).Flag = ""
'110           udtPrintLine(n).Units = ""
'120           udtPrintLine(n).NormalRange = ""
'130           udtPrintLine(n).Fasting = ""
'140       Next
'
'150       sql = "Select * from Demographics where " & _
'              "SampleID = '" & RP.SampleID & "'"
'160       Set tb = New Recordset
'170       RecOpenClient 0, tb, sql
'
'180       If tb.EOF Then
'190           Exit Sub
'200       End If
'
'210       Fasting = IIf(IsNull(tb!Fasting), False, tb!Fasting)
'220       Dob = IIf(IsDate(tb!Dob), Format(tb!Dob, "dd/mmm/yyyy"), "")
'
'230       ResultsPresent = False
'240       Set BRs = BRs.Load("Bio", RP.SampleID, "Results", gDONTCARE, gDONTCARE)
'250       If Not BRs Is Nothing Then
'260           TestCount = BRs.Count
'270           If TestCount <> 0 Then
'280               ResultsPresent = True
'290               SampleType = BRs(1).SampleType
'300               If Trim$(SampleType) = "" Then SampleType = "S"
'310           End If
'320       End If
'
'330       lpc = 0
'340       If ResultsPresent Then
'350           For Each BR In BRs
'360               RunTime = BR.RunTime
'370               v = BR.Result
'
'380               If BR.Code = CodeForGlucose Or _
'                      BR.Code = CodeForChol Or _
'                      BR.Code = CodeForTrig Then
'390                   If Fasting Then
'400                       If BR.Code = CodeForGlucose Then
'410                           sql = "Select * from Fastings where " & _
'                                  "TestName = 'GLU'"
'420                           Set tbF = New Recordset
'430                           RecOpenServer 0, tbF, sql
'440                       ElseIf BR.Code = CodeForChol Then
'450                           sql = "Select * from Fastings where " & _
'                                  "TestName = 'CHO'"
'460                           Set tbF = New Recordset
'470                           RecOpenServer 0, tbF, sql
'480                       ElseIf BR.Code = CodeForTrig Then
'490                           sql = "Select * from Fastings where " & _
'                                  "TestName = 'TRI'"
'500                           Set tbF = New Recordset
'510                           RecOpenServer 0, tbF, sql
'520                       End If
'530                       If Not tbF.EOF Then
'540                           High = tbF!FastingHigh
'550                           Low = tbF!FastingLow
'560                       Else
'570                           High = Val(BR.High)
'580                           Low = Val(BR.Low)
'590                       End If
'600                   Else
'610                       High = Val(BR.High)
'620                       Low = Val(BR.Low)
'630                   End If
'640               Else
'650                   High = Val(BR.High)
'660                   Low = Val(BR.Low)
'670               End If
'
'680               If Low < 10 Then
'690                   strLow = Format(Low, "0.00")
'700               ElseIf Low < 100 Then
'710                   strLow = Format(Low, "##.0")
'720               Else
'730                   strLow = Format(Low, " ###")
'740               End If
'750               If High < 10 Then
'760                   strHigh = Format(High, "0.00")
'770               ElseIf High < 100 Then
'780                   strHigh = Format(High, "##.0")
'790               Else
'800                   strHigh = Format(High, "### ")
'810               End If
'
'820               If IsNumeric(v) Then
'830                   If Val(v) > BR.PlausibleHigh Then
'840                       udtPrintLine(lpc).Flag = " X "
'850                       Flag = " X"
'860                   ElseIf Val(v) < BR.PlausibleLow Then
'870                       udtPrintLine(lpc).Flag = " X "
'880                       Flag = " X"
'890                   ElseIf Val(v) > BR.FlagHigh Then
'900                       udtPrintLine(lpc).Flag = " H "
'910                       Flag = " H"
'920                   ElseIf Val(v) < BR.FlagLow Then
'930                       udtPrintLine(lpc).Flag = " L "
'940                       Flag = " L"
'950                   Else
'960                       udtPrintLine(lpc).Flag = "   "
'970                       Flag = "  "
'980                   End If
'990               Else
'1000                  udtPrintLine(lpc).Flag = "   "
'1010                  Flag = "  "
'1020              End If
'1030              udtPrintLine(lpc).Analyte = Left$(BR.LongName & Space(16), 16)
'
'1040              If IsNumeric(v) Then
'1050                  Select Case BR.Printformat
'                          Case 0: strFormat = "######"
'1060                      Case 1: strFormat = "###0.0"
'1070                      Case 2: strFormat = "##0.00"
'1080                      Case 3: strFormat = "#0.000"
'1090                  End Select
'1100                  udtPrintLine(lpc).Result = Format(v, strFormat)
'1110              Else
'1120                  udtPrintLine(lpc).Result = v
'1130              End If
'
'1140              sql = "Select * from Lists where " & _
'                      "ListType = 'UN' and Code = '" & BR.Units & "'"
'1150              Set tbUN = Cnxn(0).Execute(sql)
'1160              If Not tbUN.EOF Then
'1170                  cUnits = Left$(tbUN!Text & Space(6), 6)
'1180              Else
'1190                  cUnits = Left$(BR.Units & Space(6), 6)
'1200              End If
'1210              udtPrintLine(lpc).Units = cUnits
'1220              udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
'
'1230              udtPrintLine(lpc).Fasting = ""
'1240              If tb!Fasting Then
'1250                  udtPrintLine(lpc).Fasting = "(Fasting)"
'1260              End If
'
'1270              lpc = lpc + 1
'1280          Next
'1290      End If
'
'1300      Sex = tb!Sex & ""
'1310      WordPrintHeading X, "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
'              tb!Addr0 & "", tb!Addr1 & "", Sex
'
'
'1320      If TestCount <= Val(frmMain.txtMoreThan) Then
'1330          MultiColumn = False
'1340      Else
'1350          MultiColumn = True
'1360      End If
'
'1370      With X.ActiveWindow.Selection
'1380          .Font.Size = 10
'
'1390          If MultiColumn Then
'1400              For n = 0 To Val(frmMain.txtMoreThan) - 1
'1410                  .Font.Bold = False
'1420                  .TypeText udtPrintLine(n).Analyte
'1430                  If udtPrintLine(n).Flag <> "   " Then
'1440                      .Font.Bold = True
'1450                  End If
'1460                  .TypeText udtPrintLine(n).Result
'1470                  .TypeText udtPrintLine(n).Flag
'1480                  .Font.Bold = False
'1490                  .Font.Size = 8
'1500                  .TypeText udtPrintLine(n).Units
'1510                  .TypeText udtPrintLine(n).NormalRange
'1520                  .Font.Size = 10
'                      'Now Right Hand Column
'1530                  .TypeText "  "
'1540                  .TypeText udtPrintLine(n + Val(frmMain.txtMoreThan)).Analyte
'1550                  If udtPrintLine(n + Val(frmMain.txtMoreThan)).Flag <> "   " Then
'1560                      .Font.Bold = True
'1570                  End If
'1580                  .TypeText udtPrintLine(n + Val(frmMain.txtMoreThan)).Result
'1590                  .TypeText udtPrintLine(n + Val(frmMain.txtMoreThan)).Flag
'1600                  .Font.Bold = False
'1610                  .Font.Size = 8
'1620                  .TypeText udtPrintLine(n + Val(frmMain.txtMoreThan)).Units
'1630                  .TypeText udtPrintLine(n + Val(frmMain.txtMoreThan)).NormalRange
'1640                  .Font.Size = 10
'1650                  .TypeText vbCrLf
'1660              Next
'1670              If Fasting Then
'1680                  .TypeText "(All above relate to Normal Fasting Ranges.)" & vbCrLf
'1690              End If
'1700          Else
'1710              For n = 0 To 35
'1720                  If Trim$(udtPrintLine(n).Analyte) <> "" Then
'1730                      .TypeText Space$(20)
'1740                      .Font.Bold = False
'1750                      .TypeText udtPrintLine(n).Analyte
'1760                      If udtPrintLine(n).Flag <> "   " Then
'1770                          .Font.Bold = True
'1780                      End If
'1790                      .TypeText udtPrintLine(n).Result
'1800                      .TypeText udtPrintLine(n).Flag
'1810                      .Font.Bold = False
'1820                      .TypeText udtPrintLine(n).Units
'1830                      .TypeText udtPrintLine(n).NormalRange
'1840                      .TypeText udtPrintLine(n).Fasting & vbCrLf
'1850                  End If
'1860              Next
'1870          End If
'1880          Set OBs = New Observations
'1890          Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
'1900          If Not OBs Is Nothing Then
'1910              FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
'1920              For n = 1 To 4
'1930                  .TypeText Comments(n) & vbCrLf
'1940              Next
'1950          End If
'
'1960          Set OBs = New Observations
'1970          Set OBs = OBs.Load(RP.SampleID, "Demographic")
'1980          If Not OBs Is Nothing Then
'1990              FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
'2000              For n = 1 To 4
'2010                  .TypeText Comments(n) & vbCrLf
'2020              Next
'2030          End If
'
'2040          If Not IsDate(tb!Dob) And Trim$(Sex) = "" Then
'2050              .Font.Color = vbBlue
'2060              .TypeText "          No Sex/DoB given. Normal ranges may not be relevant" & vbCrLf
'2070          ElseIf Not IsDate(tb!Dob) Then
'2080              .Font.Color = vbBlue
'2090              .TypeText "          No DoB given. Normal ranges may not be relevant" & vbCrLf
'2100          ElseIf Trim$(Sex) = "" Then
'2110              .Font.Color = vbBlue
'2120              .TypeText "          No Sex given. Normal ranges may not be relevant" & vbCrLf
'2130          End If
'
'2140          .Font.Color = vbBlack
'
'2150          If IsDate(tb!SampleDate) Then
'2160              SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
'2170          Else
'2180              SampleDate = ""
'2190          End If
'2200          If IsDate(RunTime) Then
'2210              Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
'2220          Else
'2230              If IsDate(tb!Rundate) Then
'2240                  Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
'2250              Else
'2260                  Rundate = ""
'2270              End If
'2280          End If
'
'2290          WordPrintFooter X, "Biochemistry", RP.Initiator, SampleDate, Rundate
'
'2300      End With
'
'2310      Exit Sub
'
'WordPrintResultBioSideBySide_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'2320      intEL = Erl
'2330      strES = Err.Description
'2340      LogError "modWord", "WordPrintResultBioSideBySide", intEL, strES, sql
'
'End Sub
'
'
'Public Sub WordPrintFooter(ByRef X As Word.Document, _
'          ByVal Dept As String, _
'          ByVal Initiator As String, _
'          ByVal SampleDate As String, _
'          ByVal Rundate As String)
'10        'Set x = CreateObject(word.Document)
'
'20        On Error GoTo WordPrintFooter_Error
'
'30        With X.ActiveWindow.Selection
'40            .Font.Name = "Courier New"
'
'50            Select Case Dept
'                  Case "Haematology":
'60                    .Font.Color = vbRed
'70                Case "Biochemistry":
'80                    .Font.Color = vbGreen
'90                Case "Microbiology":
'100                   .Font.Color = vbYellow
'110               Case "Blood Transfusion":
'120                   .Font.Color = vbBlue
'130           End Select
'
'140           .Font.Size = 4
'150           .TypeText String$(250, "-") & vbCrLf
'
'160           .Font.Size = 10
'170           .Font.Bold = False
'
'180           .TypeText "Sample Date:" & Format(SampleDate, "dd/mm/yyyy")
'190           If Format(SampleDate, "hh:mm") <> "00:00" Then
'200               .TypeText "   Sample Time:"
'210               .TypeText Format(SampleDate, "hh:mm")
'220           End If
'
'230           .TypeText "  Tested:" & Format(Rundate, "dd/mm/yyyy hh:mm")
'
'240           If Trim$(Initiator) <> "" Then
'250               .TypeText " Validated by " & TechnicianCodeFor(Initiator)
'260           End If
'
'270       End With
'
'280       Exit Sub
'
'WordPrintFooter_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'290       intEL = Erl
'300       strES = Err.Description
'310       LogError "modWord", "WordPrintFooter", intEL, strES
'
'End Sub
'
'
'Public Sub WordPrintHeading(ByRef X As Word.Document, _
'          ByVal Dept As String, _
'          ByVal Name As String, _
'          ByVal Dob As String, _
'          ByVal Chart As String, _
'          ByVal Address0 As String, _
'          ByVal Address1 As String, _
'          ByVal Sex As String)
'          Dim BioPhone As String
'          Dim HaemPhone As String
'10        'Set x = CreateObject(word.Document)
'
'20        On Error GoTo WordPrintHeading_Error
'
'30        With X.ActiveWindow.Selection
'40            .Font.Name = "Courier New"
'50            .Font.Size = 16
'60            .Font.Bold = True
'
'70            Select Case Dept
'                  Case "Haematology":
'80                    .Font.Color = wdColorDarkRed
'90                Case "Biochemistry":
'100                   .Font.Color = wdColorBrightGreen
'110               Case "Blood Transfusion":
'120                   .Font.Color = wdColorBlue
'130               Case "Microbiology":
'140                   .Font.Color = wdColorDarkYellow
'150           End Select
'160           Dept = Dept & " Laboratory"
'170           .TypeText "CAVAN GENERAL HOSPITAL : " & Dept
'
'180           .Font.Size = 10
'190           Select Case Dept
'                  Case "Haematology Laboratory":
'200                   HaemPhone = GetOptionSetting("HaemPhone", "")
'210                   If HaemPhone <> "" Then .TypeText " Phone " & HaemPhone
'220               Case "Biochemistry Laboratory":
'230                   BioPhone = GetOptionSetting("BioPhone", "")
'240                   If BioPhone <> "" Then .TypeText " Phone " & BioPhone
'250               Case "Blood Transfusion Laboratory":
'260                   .TypeText " Phone 38830"
'270               Case "Microbiology Laboratory":
'280           End Select
'290           .TypeText vbCrLf
'
'300           .Font.Size = 4
'310           .TypeText String$(250, "-") & vbCrLf
'
'320           .Font.Color = wdColorBlack
'
'330           .Font.Name = "Courier New"
'340           .Font.Size = 12
'350           .Font.Bold = False
'
'360           .TypeText "Sample ID:"
'370           .Font.Bold = True
'380           .TypeText Left$(RP.SampleID & Space$(25), 25)
'390           .Font.Bold = False
'400           .TypeText "Name:"
'410           .Font.Bold = True
'420           .Font.Size = 14
'430           .TypeText Left$(Name, 27) & vbCrLf
'
'440           .Font.Size = 12
'450           .Font.Bold = False
'460           .TypeText "     Ward:"
'470           .Font.Bold = True
'480           .TypeText Left$(RP.Ward & Space$(25), 25)
'490           .Font.Bold = False
'500           .TypeText " DOB:"
'510           .Font.Bold = True
'520           .TypeText Left$(Format(Dob, "dd/mm/yyyy") & Space$(21), 21)
'530           .Font.Bold = False
'540           .TypeText "Chart #:"
'550           .Font.Bold = True
'560           .TypeText Chart & vbCrLf
'
'570           .Font.Bold = False
'
'580           .TypeText "Clinician:"
'590           .Font.Bold = True
'600           .TypeText Left$(RP.Clinician & Space$(25), 25)
'610           .Font.Bold = False
'620           .TypeText "Addr:"
'630           .Font.Bold = True
'640           .TypeText Left$(Address0 & Space$(25), 25)
'650           .Font.Bold = False
'660           .TypeText "Sex:"
'670           Select Case Left$(UCase$(Trim$(Sex)), 1)
'                  Case "M": .TypeText "Male"
'680               Case "F": .TypeText "Female"
'690           End Select
'700           .TypeText vbCrLf
'
'710           .TypeText "       GP:"
'720           .Font.Bold = True
'730           .TypeText Left$(RP.GP & Space$(30), 30)
'740           .Font.Bold = False
'750           .TypeText Address1 & vbCrLf
'
'
'760           .Font.Size = 4
'770           .TypeText String$(250, "-") & vbCrLf
'
'780       End With
'
'790       Exit Sub
'
'WordPrintHeading_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'800       intEL = Erl
'810       strES = Err.Description
'820       LogError "modWord", "WordPrintHeading", intEL, strES
'
'End Sub
'
'Public Sub WordPrintCoag(ByRef X As Word.Document)
'
'          Dim tb As Recordset
'          Dim sql As String
'          Dim n As Integer
'          Dim Sex As String
'          Dim Dob As String
'          Dim CRs As New CoagResults
'          Dim CR As CoagResult
'          Dim APTTPresent As Boolean
'          Dim INRPresent As Boolean
'          Dim PTPresent As Boolean
'          Dim OnWarfarin As Boolean
'          Dim OBs As Observations
'          Dim SampleDate As String
'          Dim Rundate As String
'          Dim PrintIt As Boolean
'10        'Set x = CreateObject(word.Document)
'
'20        On Error GoTo WordPrintCoag_Error
'
'30        ReDim CommentLines(1 To 2) As String
'
'40        Set CRs = CRs.Load(RP.SampleID, gDONTCARE, "Results")
'
'50        sql = "Select * from Demographics where " & _
'              "SampleID = '" & RP.SampleID & "'"
'60        Set tb = New Recordset
'70        RecOpenClient 0, tb, sql
'
'80        If tb.EOF Then Exit Sub
'90        If CRs.Count = 0 Then Exit Sub
'
'100       If Not IsNull(tb!SampleDate) Then
'110           SampleDate = tb!SampleDate
'120       End If
'130       Dob = tb!Dob & ""
'140       If Not IsNull(tb!OnWarfarin) Then
'150           OnWarfarin = tb!OnWarfarin = 1
'160       Else
'170           OnWarfarin = False
'180       End If
'
'190       Select Case Left$(UCase$(tb!Sex & ""), 1)
'              Case "M": Sex = "M"
'200           Case "F": Sex = "F"
'210           Case Else: Sex = ""
'220       End Select
'
'230       WordPrintHeading X, "Coagulation", tb!PatName & "", tb!Dob & "", _
'              tb!Chart & "", tb!Addr0 & "", tb!Addr1 & "", Sex
'
'240       With X.ActiveWindow.Selection
'250           .Font.Name = "Courier New"
'260           .Font.Size = 10
'
'270           .TypeText vbCrLf
'
'280           APTTPresent = False
'290           INRPresent = False
'300           PTPresent = False
'310           For Each CR In CRs
'320               If CR.TestName = "APTT" Then
'330                   APTTPresent = True
'340               ElseIf CR.TestName = "INR" Then
'350                   INRPresent = True
'360               ElseIf CR.TestName = "PT" Then
'370                   PTPresent = True
'380               End If
'390           Next
'
'400           For Each CR In CRs
'410               If IsDate(CR.RunTime) Then
'420                   Rundate = Format(CR.RunTime, "dd/mm/yyyy hh:mm")
'430               ElseIf IsDate(CR.Rundate) Then
'440                   Rundate = Format(CR.Rundate, "dd/mm/yyyy")
'450               End If
'460               PrintIt = False
'470               If CR.Printable Then
'480                   PrintIt = True
'490                   If RP.Department = "D" Then
'500                       PrintIt = True
'510                   Else
'520                       If CR.TestName = "INR" And Not APTTPresent And PTPresent Then
'530                           PrintIt = True
'540                       ElseIf CR.TestName = "PT" And INRPresent And APTTPresent Then
'550                           PrintIt = True
'560                       ElseIf CR.TestName = "APTT" And INRPresent And PTPresent Then
'570                           PrintIt = True
'580                       End If
'590                   End If
'600               End If
'610               If PrintIt Then
'620                   .TypeText Space$(18)
'630                   .TypeText Left$(CR.TestName & Space$(10), 10)
'640                   If IsNumeric(CR.Result) Then
'650                       .TypeText FormatNumber(Val(CR.Result), CR.DP)
'660                   Else
'670                       .TypeText CR.Result
'680                   End If
'690                   .TypeText " " & CR.Units
'700                   .TypeText "   (" & CR.Low & "-" & CR.High & ")"
'710                   .TypeText vbCrLf
'720               End If
'730           Next
'
'740           .TypeText vbCrLf
'
'750           Set OBs = New Observations
'760           Set OBs = OBs.Load(RP.SampleID, "Coagulation")
'770           If Not OBs Is Nothing Then
'780               FillCommentLines OBs.Item(1).Comment, 4, CommentLines(), 97
'790               For n = 1 To 2
'800                   .TypeText CommentLines(n) & vbCrLf
'810               Next
'820           End If
'
'830           .Font.Color = vbBlue
'840           .TypeText vbCrLf
'850       End With
'
'860       WordPrintFooter X, "Coagulation", CRs(1).OperatorCode, SampleDate, Rundate
'
'870       Exit Sub
'
'WordPrintCoag_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'880       intEL = Erl
'890       strES = Err.Description
'900       LogError "modWord", "WordPrintCoag", intEL, strES, sql
'
'End Sub
'
'Public Sub WordPrintResultBioVert(ByRef X As Word.Document)
'
'          Dim tb As Recordset
'          Dim tbUN As Recordset
'          Dim tbF As Recordset
'          Dim sql As String
'          Dim Sex As String
'          Dim lpc As Integer
'          Dim cUnits As String
'          Dim n As Integer
'          Dim v As String
'          Dim Low As Single
'          Dim High As Single
'          Dim strLow As String * 4
'          Dim strHigh As String * 4
'          Dim BRs As New BIEResults
'          Dim BR As BIEResult
'          Dim TestCount As Integer
'          Dim SampleType As String
'          Dim ResultsPresent As Boolean
'          Dim OBs As Observations
'          Dim SampleDate As String
'          Dim Rundate As String
'          Dim Dob As String
'          Dim RunTime As String
'          Dim Fasting As String
'          Dim strFormat As String
'          Dim CodeForChol As String
'          Dim CodeForGlucose As String
'          Dim CodeForTrig As String
'10        'Set x = CreateObject(word.Document)
'
'20        On Error GoTo WordPrintResultBioVert_Error
'
'30        CodeForChol = UCase$(GetOptionSetting("BioCodeForChol", ""))
'40        CodeForTrig = UCase$(GetOptionSetting("BioCodeForTrig", ""))
'
'50        ReDim Comments(1 To 4) As String
'60        ReDim udtPrintLine(0 To 0) As PrintLine
'
'70        sql = "Select * from Demographics where " & _
'              "SampleID = '" & RP.SampleID & "'"
'80        Set tb = New Recordset
'90        RecOpenClient 0, tb, sql
'
'100       If tb.EOF Then
'110           Exit Sub
'120       End If
'
'130       Fasting = IIf(IsNull(tb!Fasting), False, tb!Fasting)
'
'140       Dob = IIf(IsDate(tb!Dob), Format(tb!Dob, "dd/mmm/yyyy"), "")
'
'150       ResultsPresent = False
'160       Set BRs = BRs.Load("Bio", RP.SampleID, "Results", gDONTCARE, gDONTCARE)
'170       If Not BRs Is Nothing Then
'180           TestCount = BRs.Count
'190           If TestCount <> 0 Then
'200               ResultsPresent = True
'210               SampleType = BRs(1).SampleType
'220               If Trim$(SampleType) = "" Then SampleType = "S"
'230           End If
'240       End If
'
'250       lpc = -1
'260       If ResultsPresent Then
'270           For Each BR In BRs
'280               lpc = lpc + 1
'290               ReDim Preserve udtPrintLine(0 To lpc)
'300               RunTime = BR.RunTime
'310               v = BR.Result
'
'320               If BR.Code = CodeForGlucose Or _
'                      BR.Code = CodeForChol Or _
'                      BR.Code = CodeForTrig Then
'330                   If Fasting Then
'340                       If BR.Code = CodeForGlucose Then
'350                           sql = "Select * from Fastings where " & _
'                                  "TestName = 'GLU'"
'360                           Set tbF = New Recordset
'370                           RecOpenServer 0, tbF, sql
'380                       ElseIf BR.Code = CodeForChol Then
'390                           sql = "Select * from Fastings where " & _
'                                  "TestName = 'CHO'"
'400                           Set tbF = New Recordset
'410                           RecOpenServer 0, tbF, sql
'420                       ElseIf BR.Code = CodeForTrig Then
'430                           sql = "Select * from Fastings where " & _
'                                  "TestName = 'TRI'"
'440                           Set tbF = New Recordset
'450                           RecOpenServer 0, tbF, sql
'460                       End If
'470                       If Not tbF.EOF Then
'480                           High = tbF!FastingHigh
'490                           Low = tbF!FastingLow
'500                       Else
'510                           High = Val(BR.High)
'520                           Low = Val(BR.Low)
'530                       End If
'540                   Else
'550                       High = Val(BR.High)
'560                       Low = Val(BR.Low)
'570                   End If
'580               Else
'590                   High = Val(BR.High)
'600                   Low = Val(BR.Low)
'610               End If
'
'620               If Low < 10 Then
'630                   strLow = Format(Low, "0.00")
'640               ElseIf Low < 100 Then
'650                   strLow = Format(Low, "##.0")
'660               Else
'670                   strLow = Format(Low, " ###")
'680               End If
'690               If High < 10 Then
'700                   strHigh = Format(High, "0.00")
'710               ElseIf High < 100 Then
'720                   strHigh = Format(High, "##.0")
'730               Else
'740                   strHigh = Format(High, "### ")
'750               End If
'
'760               If IsNumeric(v) Then
'770                   If Val(v) > BR.PlausibleHigh Then
'780                       udtPrintLine(lpc).Flag = " X "
'790                   ElseIf Val(v) < BR.PlausibleLow Then
'800                       udtPrintLine(lpc).Flag = " X "
'810                   ElseIf Val(v) > BR.FlagHigh Then
'820                       udtPrintLine(lpc).Flag = " H "
'830                   ElseIf Val(v) < BR.FlagLow Then
'840                       udtPrintLine(lpc).Flag = " L "
'850                   Else
'860                       udtPrintLine(lpc).Flag = "   "
'870                   End If
'880               Else
'890                   udtPrintLine(lpc).Flag = "   "
'900               End If
'910               udtPrintLine(lpc).Analyte = Left$(BR.LongName & Space(16), 16)
'
'920               If IsNumeric(v) Then
'930                   Select Case BR.Printformat
'                          Case 0: strFormat = "######"
'940                       Case 1: strFormat = "###0.0"
'950                       Case 2: strFormat = "##0.00"
'960                       Case 3: strFormat = "#0.000"
'970                   End Select
'980                   udtPrintLine(lpc).Result = Format(v, strFormat)
'990               Else
'1000                  udtPrintLine(lpc).Result = v
'1010              End If
'
'1020              sql = "Select * from Lists where " & _
'                      "ListType = 'UN' and Code = '" & BR.Units & "'"
'1030              Set tbUN = Cnxn(0).Execute(sql)
'1040              If Not tbUN.EOF Then
'1050                  cUnits = Left$(tbUN!Text & Space(6), 6)
'1060              Else
'1070                  cUnits = Left$(BR.Units & Space(6), 6)
'1080              End If
'1090              udtPrintLine(lpc).Units = cUnits
'1100              udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
'1110              udtPrintLine(lpc).Fasting = ""
'1120              If Not IsNull(tb!Fasting) Then
'1130                  If tb!Fasting Then
'1140                      udtPrintLine(lpc).Fasting = "(Fasting)"
'1150                  End If
'1160              End If
'
'1170          Next
'1180      End If
'
'1190      Sex = tb!Sex & ""
'1200      WordPrintHeading X, "Biochemistry", tb!PatName & "", Dob, tb!Chart & "", _
'              tb!Addr0 & "", tb!Addr1 & "", Sex
'
'1210      With X.ActiveWindow.Selection
'1220          .Font.Size = 10
'1230          For n = 0 To lpc
'1240              If udtPrintLine(n).Flag = " L " Or udtPrintLine(n).Flag = " H " Then
'1250                  .Font.Bold = True
'1260              Else
'1270                  .Font.Bold = False
'1280              End If
'1290              .TypeText udtPrintLine(n).Analyte
'1300              .TypeText udtPrintLine(n).Result
'1310              .TypeText udtPrintLine(n).Flag
'1320              .TypeText udtPrintLine(n).Units
'1330              .TypeText udtPrintLine(n).NormalRange
'1340              .TypeText udtPrintLine(n).Fasting & vbCrLf
'1350          Next
'1360          .Font.Bold = False
'
'1370          Set OBs = New Observations
'1380          Set OBs = OBs.Load(RP.SampleID, "Biochemistry")
'1390          If Not OBs Is Nothing Then
'1400              FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
'1410              For n = 1 To 4
'1420                  .TypeText Comments(n) & vbCrLf
'1430              Next
'1440          End If
'
'1450          Set OBs = New Observations
'1460          Set OBs = OBs.Load(RP.SampleID, "Demographic")
'1470          If Not OBs Is Nothing Then
'1480              FillCommentLines OBs.Item(1).Comment, 4, Comments(), 97
'1490              For n = 1 To 4
'1500                  .TypeText Comments(n) & vbCrLf
'1510              Next
'1520          End If
'
'1530          If Not IsDate(tb!Dob) And Trim$(Sex) = "" Then
'1540              .Font.Color = vbBlue
'1550              .TypeText "No Sex/DoB given. Normal ranges may not be relevant" & vbCrLf
'1560          ElseIf Not IsDate(tb!Dob) Then
'1570              .Font.Color = vbBlue
'1580              .TypeText "No DoB given. Normal ranges may not be relevant" & vbCrLf
'1590          ElseIf Trim$(Sex) = "" Then
'1600              .Font.Color = vbBlue
'1610              .TypeText "No Sex given. Normal ranges may not be relevant" & vbCrLf
'1620          End If
'
'1630          .Font.Color = vbBlack
'
    '1640          If IsDate(tb!SampleDate) Then
'1650              SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
'1660          Else
'1670              SampleDate = ""
'1680          End If
'1690          If IsDate(RunTime) Then
'1700              Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
'1710          Else
'1720              If IsDate(tb!Rundate) Then
'1730                  Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
'1740              Else
'1750                  Rundate = ""
'1760              End If
'1770          End If
'
'1780          WordPrintFooter X, "Biochemistry", RP.Initiator, SampleDate, Rundate
'
'1790      End With
'
'1800      Exit Sub
'
'WordPrintResultBioVert_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'1810      intEL = Erl
'1820      strES = Err.Description
'1830      LogError "modWord", "WordPrintResultBioVert", intEL, strES, sql
'
'End Sub
'
'
'
'
