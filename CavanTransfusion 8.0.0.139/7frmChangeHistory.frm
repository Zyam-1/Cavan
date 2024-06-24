VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmChangeHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change History"
   ClientHeight    =   7860
   ClientLeft      =   1440
   ClientTop       =   1845
   ClientWidth     =   12795
   ControlBox      =   0   'False
   DrawWidth       =   10
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "7frmChangeHistory.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7860
   ScaleWidth      =   12795
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   900
      Left            =   11685
      Picture         =   "7frmChangeHistory.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "bCancel"
      Top             =   90
      Width           =   900
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   900
      Left            =   10575
      Picture         =   "7frmChangeHistory.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Width           =   900
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   900
      Left            =   5040
      Picture         =   "7frmChangeHistory.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Width           =   900
   End
   Begin VB.TextBox treport 
      Height          =   585
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4740
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtLabNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   420
      Left            =   2580
      MaxLength       =   14
      TabIndex        =   0
      Top             =   270
      Width           =   2325
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   7515
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6345
      Left            =   60
      TabIndex        =   4
      Top             =   1065
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   11192
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"7frmChangeHistory.frx":1A29
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Change history for sample number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   390
      Width           =   2385
   End
End
Attribute VB_Name = "frmChangeHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillDetails()

      Dim tbDem As Recordset
      Dim sn As Recordset
      Dim sql As String
10    On Error GoTo FillDetails_Error

20    ReDim Previous(0 To 88) As String
30    ReDim Current(0 To 88) As String
40    ReDim fieldname(0 To 88) As String
      Dim n As Integer
      Dim s As String

50    g.Rows = 2
60    g.AddItem ""
70    g.RemoveItem 1

80    treport = ""
90    If Trim$(txtLabNum) = "" Then Exit Sub

100   fieldname(0) = "Patient Number"
110   fieldname(1) = "patient Name"
120   fieldname(2) = "Ward"
130   fieldname(3) = "Clinician"
140   fieldname(4) = "Procedure"
150   fieldname(5) = "Conditions"
160   fieldname(6) = "Addr 1"
170   fieldname(7) = "Addr 2"
180   fieldname(8) = "Addr 3"
190   fieldname(9) = "Addr 4"
200   fieldname(10) = "Special Product"
210   fieldname(11) = "Previous Transfusions"
220   fieldname(12) = "Previous Reaction"
230   fieldname(13) = "Previous Pregnancy"
240   fieldname(14) = "Requested"
250   fieldname(15) = "Product"
260   fieldname(16) = "FG Pattern"
270   fieldname(17) = "FG Suggested"
280   fieldname(18) = "Forward Group"
290   fieldname(19) = "RG Pattern"
300   fieldname(20) = "RG Suggested"
310   fieldname(21) = "Reverse Group"
320   fieldname(22) = "Sex"
330   fieldname(23) = "Previous Group"
340   fieldname(24) = "Autoantibodies"
350   fieldname(25) = "Antibody Screen Lot Number"
360   fieldname(26) = "Antibody Screen Coombs Pattern"
370   fieldname(27) = "Antibody Screen Enzyme Pattern"
380   fieldname(28) = "Antibody Screen"
390   fieldname(29) = "AM/PM"
400   fieldname(30) = "Lab Number"
410   fieldname(31) = "Age"
420   fieldname(32) = "Previous Rh"
430   fieldname(33) = "Hold"
440   fieldname(34) = "Request From"
450   fieldname(35) = "Autologous"
460   fieldname(36) = "Comment"
470   fieldname(37) = "Antibody ID Suggested"
480   fieldname(38) = "Antibody ID Reported"
490   fieldname(39) = "Maiden Name"
500   fieldname(40) = "Coombs"
510   fieldname(41) = "DAT0"
520   fieldname(42) = "DAT1"
530   fieldname(43) = "DAT2"
540   fieldname(44) = "DAT3"
550   fieldname(45) = "DAT4"
560   fieldname(46) = "DAT5"
570   fieldname(47) = "DAT6"
580   fieldname(48) = "DAT7"
590   fieldname(49) = "DAT8"
600   fieldname(50) = "DAT9"
610   fieldname(51) = "EDD"
620   fieldname(52) = "Date Required"
630   fieldname(53) = "Sample Date"
640   fieldname(54) = "Date of Birth"
650   fieldname(55) = "Date/Time"
660   fieldname(56) = "Operator"
670   fieldname(57) = "BarCode"
680   fieldname(58) = "NOPAS"
690   fieldname(59) = "AandE"
700   fieldname(60) = "TYPENEX"
710   fieldname(61) = "DateReceived"
720   fieldname(62) = "DAT10"
730   fieldname(63) = "DAT11"
740   fieldname(64) = "Sample Comment"
750   fieldname(65) = "Checker"
760   fieldname(66) = "Routine"
770   fieldname(67) = "Hospital"
780   fieldname(68) = "GP"
790   fieldname(69) = "RequestChart"
800   fieldname(70) = "RequestDoB"
810   fieldname(71) = "RequestAge"
820   fieldname(72) = "RequestSex"
830   fieldname(73) = "RequestAddress0"
840   fieldname(74) = "RequestAddress1"
850   fieldname(75) = "RequestWard"
860   fieldname(76) = "RequestClinician"
870   fieldname(77) = "RequestGP"
880   fieldname(78) = "SampleChart"
890   fieldname(79) = "SampleName"
900   fieldname(80) = "SampleDoB"
910   fieldname(81) = "SampleAge"
920   fieldname(82) = "SampleSex"
930   fieldname(83) = "SampleAddress0"
940   fieldname(84) = "SampleAddress1"
950   fieldname(85) = "SampleWard"
960   fieldname(86) = "SampleID"
970   fieldname(87) = "SampleSigned"
980   fieldname(88) = "RequestName"



990   sql = "select * from PatientDetails where " & _
            "labnumber = '" & txtLabNum & "'"
1000  Set tbDem = New Recordset
1010  RecOpenClientBB 0, tbDem, sql
1020  If tbDem.EOF Then
1030    treport = "No Record of that Lab Number!"
1040    Exit Sub
1050  End If

1060  sql = "select * from PatientDetailsAudit where " & _
            "labnumber = '" & txtLabNum & "' " & _
            "order by DateTime desc"
1070  Set sn = New Recordset
1080  RecOpenClientBB 0, sn, sql
1090  If sn.EOF Then
1100    treport = "No Record of any changes to this Lab Number."
1110    Exit Sub
1120  End If

1130  Current(0) = tbDem!Patnum & ""
1140  Current(1) = tbDem!Name & ""
1150  Current(2) = tbDem!Ward & ""
1160  Current(3) = tbDem!Clinician & ""
1170  Current(4) = tbDem!Procedure & ""
1180  Current(5) = tbDem!Conditions & ""
1190  Current(6) = tbDem!Addr1 & ""
1200  Current(7) = tbDem!Addr2 & ""
1210  Current(8) = tbDem!Addr3 & ""
1220  Current(9) = tbDem!addr4 & ""
1230  Current(10) = tbDem!specialprod & ""
1240  Current(11) = tbDem!prevtrans & ""
1250  Current(12) = tbDem!prevreact & ""
1260  Current(13) = tbDem!prevpreg & ""
1270  Current(14) = tbDem!Requested & ""
1280  Current(15) = tbDem!Product3 & ""
1290  Current(16) = tbDem!fgpattern & ""
1300  Current(17) = tbDem!fgsuggest & ""
1310  Current(18) = tbDem!fGroup & ""
1320  Current(19) = tbDem!rgpattern & ""
1330  Current(20) = tbDem!rgSuggest & ""
1340  Current(21) = tbDem!rgroup & ""
1350  Current(22) = tbDem!Sex & ""
1360  Current(23) = tbDem!PrevGroup & ""
1370  Current(24) = tbDem!AutoAnt & ""
1380  Current(25) = tbDem!anti3lot & ""
1390  Current(26) = tbDem!anti3c & ""
1400  Current(27) = tbDem!anti3e & ""
1410  Current(28) = tbDem!Anti3Reported & ""
1420  Current(29) = tbDem!ampm & ""
1430  Current(30) = tbDem!LabNumber & ""
1440  Current(31) = tbDem!Age & ""
1450  Current(32) = tbDem!previousrh & ""
1460  Current(33) = tbDem!Hold & ""
1470  Current(34) = tbDem!requestfrom & ""
1480  Current(35) = tbDem!Autolog & ""
1490  Current(36) = tbDem!Comment & ""
1500  Current(37) = tbDem!AIDS & ""
1510  Current(38) = tbDem!AIDR & ""
1520  Current(39) = tbDem!maiden & ""
1530  Current(40) = tbDem!coombs & ""
1540  Current(41) = tbDem!DAT0 & ""
1550  Current(42) = tbDem!DAT1 & ""
1560  Current(43) = tbDem!DAT2 & ""
1570  Current(44) = tbDem!DAT3 & ""
1580  Current(45) = tbDem!DAT4 & ""
1590  Current(46) = tbDem!DAT5 & ""
1600  Current(47) = tbDem!DAT6 & ""
1610  Current(48) = tbDem!DAT7 & ""
1620  Current(49) = tbDem!DAT8 & ""
1630  Current(50) = tbDem!DAT9 & ""
1640  If Not IsNull(tbDem!edd) Then
1650    Current(51) = Format(tbDem!edd, "dd/mm/yyyy")
1660  Else
1670    Current(51) = ""
1680  End If
1690  If Not IsNull(tbDem!daterequired) Then
1700    Current(52) = Format(tbDem!daterequired, "dd/mm/yyyy")
1710  Else
1720    Current(52) = ""
1730  End If
1740  If Not IsNull(tbDem!SampleDate) Then
1750    Current(53) = Format(tbDem!SampleDate, "dd/mm/yyyy HH:nn:ss")
1760  Else
1770    Current(53) = ""
1780  End If
1790  If Not IsNull(tbDem!DoB) Then
1800    Current(54) = Format(tbDem!DoB, "dd/mm/yyyy")
1810  Else
1820    Current(54) = ""
1830  End If
1840  Current(55) = tbDem!DateTime & ""
1850  Current(56) = tbDem!Operator & ""
1860  Current(57) = tbDem!BarCode & ""
1870  Current(58) = tbDem!NOPAS & ""
1880  Current(59) = tbDem!AandE & ""
1890  Current(60) = tbDem!Typenex & ""
1900  Current(61) = tbDem!DateReceived & ""
1910  Current(62) = tbDem!DAT10 & ""
1920  Current(63) = tbDem!DAT11 & ""
1930  Current(64) = tbDem!SampleComment & ""
1940  Current(65) = tbDem!Checker & ""
1950  Current(66) = tbDem!RooH & ""
1960  Current(67) = tbDem!Hospital & ""
1970  Current(68) = tbDem!GP & ""
1980  Current(69) = tbDem!requestChart & ""
1990  Current(70) = tbDem!RequestDoB & ""
2000  Current(71) = tbDem!Requestage & ""
2010  Current(72) = tbDem!RequestSex & ""
2020  Current(73) = tbDem!RequestAddress0 & ""
2030  Current(74) = tbDem!RequestAddress1 & ""
2040  Current(75) = tbDem!RequestWard & ""
2050  Current(76) = tbDem!RequestClinician & ""
2060  Current(77) = tbDem!RequestGP & ""
2070  Current(78) = tbDem!sampleChart & ""
2080  Current(79) = tbDem!sampleName & ""
2090  Current(80) = tbDem!sampleDoB & ""
2100  Current(81) = tbDem!sampleage & ""
2110  Current(82) = tbDem!sampleSex & ""
2120  Current(83) = tbDem!sampleAddress0 & ""
2130  Current(84) = tbDem!sampleAddress1 & ""
2140  Current(85) = tbDem!sampleWard & ""
2150  Current(86) = tbDem!SampleID & ""
2160  Current(87) = tbDem!samplesigned & ""
2170  Current(88) = tbDem!RequestName & ""


2180  Previous(0) = sn!Patnum & ""
2190  Previous(1) = sn!Name & ""
2200  Previous(2) = sn!Ward & ""
2210  Previous(3) = sn!Clinician & ""
2220  Previous(4) = sn!Procedure & ""
2230  Previous(5) = sn!Conditions & ""
2240  Previous(6) = sn!Addr1 & ""
2250  Previous(7) = sn!Addr2 & ""
2260  Previous(8) = sn!Addr3 & ""
2270  Previous(9) = sn!addr4 & ""
2280  Previous(10) = sn!specialprod & ""
2290  Previous(11) = sn!prevtrans & ""
2300  Previous(12) = sn!prevreact & ""
2310  Previous(13) = sn!prevpreg & ""
2320  Previous(14) = sn!Requested & ""
2330  Previous(15) = sn!Product3 & ""
2340  Previous(16) = sn!fgpattern & ""
2350  Previous(17) = sn!fgsuggest & ""
2360  Previous(18) = sn!fGroup & ""
2370  Previous(19) = sn!rgpattern & ""
2380  Previous(20) = sn!rgSuggest & ""
2390  Previous(21) = sn!rgroup & ""
2400  Previous(22) = sn!Sex & ""
2410  Previous(23) = sn!PrevGroup & ""
2420  Previous(24) = sn!AutoAnt & ""
2430  Previous(25) = sn!anti3lot & ""
2440  Previous(26) = sn!anti3c & ""
2450  Previous(27) = sn!anti3e & ""
2460  Previous(28) = sn!Anti3Reported & ""
2470  Previous(29) = sn!ampm & ""
2480  Previous(30) = sn!LabNumber & ""
2490  Previous(31) = sn!Age & ""
2500  Previous(32) = sn!previousrh & ""
2510  Previous(33) = sn!Hold & ""
2520  Previous(34) = sn!requestfrom & ""
2530  Previous(35) = sn!Autolog & ""
2540  Previous(36) = sn!Comment & ""
2550  Previous(37) = sn!AIDS & ""
2560  Previous(38) = sn!AIDR & ""
2570  Previous(39) = sn!maiden & ""
2580  Previous(40) = sn!coombs & ""
2590  Previous(41) = sn!DAT0 & ""
2600  Previous(42) = sn!DAT1 & ""
2610  Previous(43) = sn!DAT2 & ""
2620  Previous(44) = sn!DAT3 & ""
2630  Previous(45) = sn!DAT4 & ""
2640  Previous(46) = sn!DAT5 & ""
2650  Previous(47) = sn!DAT6 & ""
2660  Previous(48) = sn!DAT7 & ""
2670  Previous(49) = sn!DAT8 & ""
2680  Previous(50) = sn!DAT9 & ""
2690  If Not IsNull(sn!edd) Then
2700    Previous(51) = Format(sn!edd, "dd/mm/yyyy")
2710  Else
2720    Previous(51) = ""
2730  End If
2740  If Not IsNull(sn!daterequired) Then
2750    Previous(52) = Format(sn!daterequired, "dd/mm/yyyy")
2760  Else
2770    Previous(52) = ""
2780  End If
2790  If Not IsNull(sn!SampleDate) Then
2800    Previous(53) = Format(sn!SampleDate, "dd/mm/yyyy HH:nn:ss")
2810  Else
2820    Previous(53) = ""
2830  End If
2840  If Not IsNull(sn!DoB) Then
2850    Previous(54) = Format(sn!DoB, "dd/mm/yyyy")
2860  Else
2870    Previous(54) = ""
2880  End If
2890  Previous(55) = sn!DateTime & ""
2900  Previous(56) = sn!Operator & ""
2910  Previous(57) = sn!BarCode & ""
2920  Previous(58) = sn!NOPAS & ""
2930  Previous(59) = sn!AandE & ""
2940  Previous(60) = sn!Typenex & ""
2950  Previous(61) = sn!DateReceived & ""
2960  Previous(62) = sn!DAT10 & ""
2970  Previous(63) = sn!DAT11 & ""
2980  Previous(64) = sn!SampleComment & ""
2990  Previous(65) = sn!Checker & ""
3000  Previous(66) = sn!RooH & ""
3010  Previous(67) = sn!Hospital & ""
3020  Previous(68) = sn!GP & ""
3030  Previous(69) = sn!requestChart & ""
3040  Previous(70) = sn!RequestDoB & ""
3050  Previous(71) = sn!Requestage & ""
3060  Previous(72) = sn!RequestSex & ""
3070  Previous(73) = sn!RequestAddress0 & ""
3080  Previous(74) = sn!RequestAddress1 & ""
3090  Previous(75) = sn!RequestWard & ""
3100  Previous(76) = sn!RequestClinician & ""
3110  Previous(77) = sn!RequestGP & ""
3120  Previous(78) = sn!sampleChart & ""
3130  Previous(79) = sn!sampleName & ""
3140  Previous(80) = sn!sampleDoB & ""
3150  Previous(81) = sn!sampleage & ""
3160  Previous(82) = sn!sampleSex & ""
3170  Previous(83) = sn!sampleAddress0 & ""
3180  Previous(84) = sn!sampleAddress1 & ""
3190  Previous(85) = sn!sampleWard & ""
3200  Previous(86) = sn!SampleID & ""
3210  Previous(87) = sn!samplesigned & ""
3220  Previous(88) = sn!RequestName & ""

3230  treport = "Change History For Sample Number : " & txtLabNum & vbCrLf
3240  treport = treport & "------------------------------------------" & vbCrLf & vbCrLf

      Dim strItem As String

3250  For n = 0 To 88
3260      If Current(n) <> Previous(n) Then
3270          s = Format(Current(55), "dd/mm/yy hh:mm:ss") & " " & _
                  fieldname(n) & "    '"
3280              strItem = Current(55) & vbTab & fieldname(n) & vbTab
3290          If Previous(n) <> "" Then
3300              s = s & Previous(n)
3310              strItem = strItem & Previous(n)
3320          Else
3330              strItem = strItem & "*NULL*"
3340              s = s & "*NULL*"
3350          End If
    
3360          s = s & "'"
3370          s = s & " changed to '"
3380          strItem = strItem & " changed to "
3390          If Current(n) <> "" Then
3400              s = s & Current(n)
3410              strItem = strItem & Current(n)
3420          Else
3430              s = s & "*NULL*"
3440              strItem = strItem & "*NULL*"
3450          End If
    
3460          strItem = strItem & vbTab
3470          s = s & "' by " & Current(56) & vbCrLf
3480          strItem = strItem & Current(56)
3490          treport = treport & s
3500          g.AddItem strItem, g.row
3510      End If
3520  Next

3530  For n = 0 To 88
3540     Current(n) = Previous(n)
3550  Next

3560  sn.MoveNext
3570  Do While Not sn.EOF
3580      Previous(0) = sn!Patnum & ""
3590      Previous(1) = sn!Name & ""
3600      Previous(2) = sn!Ward & ""
3610      Previous(3) = sn!Clinician & ""
3620      Previous(4) = sn!Procedure & ""
3630      Previous(5) = sn!Conditions & ""
3640      Previous(6) = sn!Addr1 & ""
3650      Previous(7) = sn!Addr2 & ""
3660      Previous(8) = sn!Addr3 & ""
3670      Previous(9) = sn!addr4 & ""
3680      Previous(10) = sn!specialprod & ""
3690      Previous(11) = sn!prevtrans & ""
3700      Previous(12) = sn!prevreact & ""
3710      Previous(13) = sn!prevpreg & ""
3720      Previous(14) = sn!Requested & ""
3730      Previous(15) = sn!Product3 & ""
3740      Previous(16) = sn!fgpattern & ""
3750      Previous(17) = sn!fgsuggest & ""
3760      Previous(18) = sn!fGroup & ""
3770      Previous(19) = sn!rgpattern & ""
3780      Previous(20) = sn!rgSuggest & ""
3790      Previous(21) = sn!rgroup & ""
3800      Previous(22) = sn!Sex & ""
3810      Previous(23) = sn!PrevGroup & ""
3820      Previous(24) = sn!AutoAnt & ""
3830      Previous(25) = sn!anti3lot & ""
3840      Previous(26) = sn!anti3c & ""
3850      Previous(27) = sn!anti3e & ""
3860      Previous(28) = sn!Anti3Reported & ""
3870      Previous(29) = sn!ampm & ""
3880      Previous(30) = sn!LabNumber & ""
3890      Previous(31) = sn!Age & ""
3900      Previous(32) = sn!previousrh & ""
3910      Previous(33) = sn!Hold & ""
3920      Previous(34) = sn!requestfrom & ""
3930      Previous(35) = sn!Autolog & ""
3940      Previous(36) = sn!Comment & ""
3950      Previous(37) = sn!AIDS & ""
3960      Previous(38) = sn!AIDR & ""
3970      Previous(39) = sn!maiden & ""
3980      Previous(40) = sn!coombs & ""
3990      Previous(41) = sn!DAT0 & ""
4000      Previous(42) = sn!DAT1 & ""
4010      Previous(43) = sn!DAT2 & ""
4020      Previous(44) = sn!DAT3 & ""
4030      Previous(45) = sn!DAT4 & ""
4040      Previous(46) = sn!DAT5 & ""
4050      Previous(47) = sn!DAT6 & ""
4060      Previous(48) = sn!DAT7 & ""
4070      Previous(49) = sn!DAT8 & ""
4080      Previous(50) = sn!DAT9 & ""
4090      If Not IsNull(sn!edd) Then
4100        Previous(51) = Format(sn!edd, "dd/mm/yyyy")
4110      Else
4120        Previous(51) = ""
4130      End If
4140      If Not IsNull(sn!daterequired) Then
4150        Previous(52) = Format(sn!daterequired, "dd/mm/yyyy")
4160      Else
4170        Previous(52) = ""
4180      End If
4190      If Not IsNull(sn!SampleDate) Then
4200        Previous(53) = Format(sn!SampleDate, "dd/mm/yyyy HH:nn:ss")
4210      Else
4220        Previous(53) = ""
4230      End If
4240      If Not IsNull(sn!DoB) Then
4250        Previous(54) = Format(sn!DoB, "dd/mm/yyyy")
4260      Else
4270        Previous(54) = ""
4280      End If
4290      Previous(55) = sn!DateTime & ""
4300      Previous(56) = sn!Operator & ""
4310      Previous(57) = sn!BarCode & ""
4320      Previous(58) = sn!NOPAS & ""
4330      Previous(59) = sn!AandE & ""
4340      Previous(60) = sn!Typenex & ""
4350      Previous(61) = sn!DateReceived & ""
4360      Previous(62) = sn!DAT10 & ""
4370      Previous(63) = sn!DAT11 & ""
4380      Previous(64) = sn!SampleComment & ""
4390      Previous(65) = sn!Checker & ""
4400      Previous(66) = sn!RooH & ""
4410      Previous(67) = sn!Hospital & ""
4420      Previous(68) = sn!GP & ""
4430      Previous(69) = sn!requestChart & ""
4440      Previous(70) = sn!RequestDoB & ""
4450      Previous(71) = sn!Requestage & ""
4460      Previous(72) = sn!RequestSex & ""
4470      Previous(73) = sn!RequestAddress0 & ""
4480      Previous(74) = sn!RequestAddress1 & ""
4490      Previous(75) = sn!RequestWard & ""
4500      Previous(76) = sn!RequestClinician & ""
4510      Previous(77) = sn!RequestGP & ""
4520      Previous(78) = sn!sampleChart & ""
4530      Previous(79) = sn!sampleName & ""
4540      Previous(80) = sn!sampleDoB & ""
4550      Previous(81) = sn!sampleage & ""
4560      Previous(82) = sn!sampleSex & ""
4570      Previous(83) = sn!sampleAddress0 & ""
4580      Previous(84) = sn!sampleAddress1 & ""
4590      Previous(85) = sn!sampleWard & ""
4600      Previous(86) = sn!SampleID & ""
4610      Previous(87) = sn!samplesigned & ""
4620      Previous(88) = sn!RequestName & ""
  
          '  For n = 0 To 88
          '    If Current(n) <> Previous(n) Then
          '      s = Format(Current(55), "dd/mm/yy hh:mm:ss") & " " & _
          '          fieldname(n) & "    '"
          '      If Previous(n) <> "" Then
          '        s = s & Previous(n)
          '      Else
          '        s = s & "*NULL*"
          '      End If
          '      s = s & "'"
          '      s = s & " changed to '"
          '      If Current(n) <> "" Then
          '        s = s & Current(n)
          '      Else
          '        s = s & "*NULL*"
          '      End If
          '      s = s & "' by " & Current(56) & vbCrLf
          '      treport = treport & s
          '    End If
          '  Next
          '  For n = 0 To 88
          '    Current(n) = Previous(n)
          '  Next
4630      For n = 0 To 88
4640          If Current(n) <> Previous(n) Then
4650              s = Format(Current(55), "dd/mm/yy hh:mm:ss") & " " & _
                      fieldname(n) & "    '"
4660                  strItem = Current(55) & vbTab & fieldname(n) & vbTab
4670              If Previous(n) <> "" Then
4680                s = s & Previous(n)
4690                strItem = strItem & Previous(n)
4700              Else
4710                strItem = strItem & "*NULL*"
4720                s = s & "*NULL*"
4730              End If
    
4740              s = s & "'"
4750              s = s & " changed to '"
4760              strItem = strItem & " changed to "
4770              If Current(n) <> "" Then
4780                s = s & Current(n)
4790                strItem = strItem & Current(n)
4800              Else
4810                s = s & "*NULL*"
4820                strItem = strItem & "*NULL*"
4830              End If

4840              strItem = strItem & vbTab
4850              s = s & "' by " & Current(56) & vbCrLf
4860              strItem = strItem & Current(56)
4870              treport = treport & s
4880              g.AddItem strItem, g.row
4890          End If
4900      Next
4910      sn.MoveNext
4920      For n = 0 To 88
4930          Current(n) = Previous(n)
4940      Next
4950  Loop
4960  treport = treport & vbCrLf & "End of Report."

4970  Exit Sub

FillDetails_Error:

      Dim strES As String
      Dim intEL As Integer

4980  intEL = Erl
4990  strES = Err.Description
5000  LogError "frmChangeHistory", "FillDetails", intEL, strES, sql


End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()

      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.Font.Size = 10
50    Printer.Orientation = vbPRORLandscape
60    Printer.Font.Bold = True

      'Heading Section
70    Printer.Print
80    Printer.Print FormatString("Change History For Sample Number " & txtLabNum, 140, , AlignCenter)
90    Printer.Font.Size = 9

100   For i = 1 To 152
110       Printer.Print "-";
120   Next i
130   Printer.Print

140   Printer.Print FormatString("", 0, "|");
150   Printer.Print FormatString("Changed To", 20, "|", AlignCenter); 'Changed To
160   Printer.Print FormatString("Field Name", 15, "|", AlignCenter); 'Field Name
170   Printer.Print FormatString("Change History", 107, "|", AlignCenter); 'Event
180   Printer.Print FormatString("Op", 5, "|", AlignCenter) 'Opertor

190   Printer.Font.Bold = False
200   For i = 1 To 152
210       Printer.Print "-";
220   Next i
230   Printer.Print

240   For i = 1 To g.Rows - 1
250       Printer.Print FormatString("", 0, "|");
260       Printer.Print FormatString(g.TextMatrix(i, 0), 20, "|"); 'Changed To
270       Printer.Print FormatString(g.TextMatrix(i, 1), 15, "|"); 'Field Name
280       Printer.Print FormatString(g.TextMatrix(i, 2), 107, "|"); 'Event
290       Printer.Print FormatString(g.TextMatrix(i, 3), 5, "|") 'Opertor
300   Next i

310   Printer.EndDoc



320   For Each Px In Printers
330     If Px.DeviceName = OriginalPrinter Then
340       Set Printer = Px
350       Exit For
360     End If
370   Next

End Sub



Private Sub cmdSearch_Click()

10    Me.MousePointer = vbHourglass
20    FillDetails
30    Me.MousePointer = vbDefault

End Sub

Private Sub Form_Deactivate()

10    txtLabNum = ""

End Sub


Private Sub Form_Load()
10    If txtLabNum <> "" Then
20      FillDetails
30    End If
End Sub

Private Sub txtLabNum_LostFocus()

10    FillDetails

End Sub

