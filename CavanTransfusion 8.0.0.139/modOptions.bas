Attribute VB_Name = "modOptions"
Option Explicit

Public Type udtOptionList
  Description As String
  Value As String
  DefinedAs As String 'Boolean/String/Single/Long/Integer etc
End Type

Public sysOptSoundCritical() As String
Public sysOptSoundInformation() As String
Public sysOptSoundQuestion() As String
Public sysOptSoundSevere() As String

Public sysOptTransfusionExpiry() As String

Public SysLastPASWarningPeriod As Integer

Public RBCRT011_ReturnValue() As String

Public Sub Load_RBCRT011()
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo Load_RBCRT011_Error

20    ReDim RBCRT011_ReturnValue(0 To 99) As String

      'Set default values
30    RBCRT011_ReturnValue(0) = ""
40    RBCRT011_ReturnValue(1) = "Ena"
50    RBCRT011_ReturnValue(2) = "‘N’"
60    RBCRT011_ReturnValue(3) = "Vw"
70    RBCRT011_ReturnValue(4) = "Mur*"
80    RBCRT011_ReturnValue(5) = "Hut"
90    RBCRT011_ReturnValue(6) = "Hil"
100   RBCRT011_ReturnValue(7) = "P"
110   RBCRT011_ReturnValue(8) = "PP1Pk"
120   RBCRT011_ReturnValue(9) = "hrS"
130   RBCRT011_ReturnValue(10) = "hrB"
140   RBCRT011_ReturnValue(11) = "f"
150   RBCRT011_ReturnValue(12) = "Ce"
160   RBCRT011_ReturnValue(13) = "G"
170   RBCRT011_ReturnValue(14) = "Hro"
180   RBCRT011_ReturnValue(15) = "CE"
190   RBCRT011_ReturnValue(16) = "cE"
200   RBCRT011_ReturnValue(17) = "Cx"
210   RBCRT011_ReturnValue(18) = "Ew"
220   RBCRT011_ReturnValue(19) = "Dw"
230   RBCRT011_ReturnValue(20) = "hrH"
240   RBCRT011_ReturnValue(21) = "Goa"
250   RBCRT011_ReturnValue(22) = "Rh32"
260   RBCRT011_ReturnValue(23) = "Rh33"
270   RBCRT011_ReturnValue(24) = "Tar"
280   RBCRT011_ReturnValue(25) = "Kpb"
290   RBCRT011_ReturnValue(26) = "Kpc"
300   RBCRT011_ReturnValue(27) = "Jsb"
310   RBCRT011_ReturnValue(28) = "Ula"
320   RBCRT011_ReturnValue(29) = "K11"
330   RBCRT011_ReturnValue(30) = "K12"
340   RBCRT011_ReturnValue(31) = "K13"
350   RBCRT011_ReturnValue(32) = "K14"
360   RBCRT011_ReturnValue(33) = "K17"
370   RBCRT011_ReturnValue(34) = "K18"
380   RBCRT011_ReturnValue(35) = "K19"
390   RBCRT011_ReturnValue(36) = "K22"
400   RBCRT011_ReturnValue(37) = "K23"
410   RBCRT011_ReturnValue(38) = "K24"
420   RBCRT011_ReturnValue(39) = "Lub"
430   RBCRT011_ReturnValue(40) = "Lu3"
440   RBCRT011_ReturnValue(41) = "Lu4"
450   RBCRT011_ReturnValue(42) = "Lu5"
460   RBCRT011_ReturnValue(43) = "Lu6"
470   RBCRT011_ReturnValue(44) = "Lu7"
480   RBCRT011_ReturnValue(45) = "Lu8"
490   RBCRT011_ReturnValue(46) = "Lu11"
500   RBCRT011_ReturnValue(47) = "Lu12"
510   RBCRT011_ReturnValue(48) = "Lu13"
520   RBCRT011_ReturnValue(49) = "Lu20"
530   RBCRT011_ReturnValue(50) = "Aua"
540   RBCRT011_ReturnValue(51) = "Aub"
550   RBCRT011_ReturnValue(52) = "Fy4"
560   RBCRT011_ReturnValue(53) = "Fy5"
570   RBCRT011_ReturnValue(54) = "Fy6"
580   RBCRT011_ReturnValue(55) = "Dib"
590   RBCRT011_ReturnValue(56) = "Sda"
600   RBCRT011_ReturnValue(57) = "Wrb"
610   RBCRT011_ReturnValue(58) = "Ytb"
620   RBCRT011_ReturnValue(59) = "Xga"
630   RBCRT011_ReturnValue(60) = "Sc1"
640   RBCRT011_ReturnValue(61) = "Sc2"
650   RBCRT011_ReturnValue(62) = "Sc3"
660   RBCRT011_ReturnValue(63) = "Joa"
670   RBCRT011_ReturnValue(64) = "" '   removed
680   RBCRT011_ReturnValue(65) = "Hy"
690   RBCRT011_ReturnValue(66) = "Gya"
700   RBCRT011_ReturnValue(67) = "Co3"
710   RBCRT011_ReturnValue(68) = "LWa"
720   RBCRT011_ReturnValue(69) = "LWb"
730   RBCRT011_ReturnValue(70) = "Kx"
740   RBCRT011_ReturnValue(71) = "Ge2"
750   RBCRT011_ReturnValue(72) = "Ge3"
760   RBCRT011_ReturnValue(73) = "Wb"
770   RBCRT011_ReturnValue(74) = "Lsa"
780   RBCRT011_ReturnValue(75) = "Ana"
790   RBCRT011_ReturnValue(76) = "Dha"
800   RBCRT011_ReturnValue(77) = "Cra"
810   RBCRT011_ReturnValue(78) = "IFC"
820   RBCRT011_ReturnValue(79) = "Kna"
830   RBCRT011_ReturnValue(80) = "Inb"
840   RBCRT011_ReturnValue(81) = "Csa"
850   RBCRT011_ReturnValue(82) = "I"
860   RBCRT011_ReturnValue(83) = "Era"
870   RBCRT011_ReturnValue(84) = "Vel"
880   RBCRT011_ReturnValue(85) = "Lan"
890   RBCRT011_ReturnValue(86) = "Ata"
900   RBCRT011_ReturnValue(87) = "Jra"
910   RBCRT011_ReturnValue(88) = "Oka"
920   RBCRT011_ReturnValue(89) = "Wra"
930   RBCRT011_ReturnValue(90) = ""
940   RBCRT011_ReturnValue(91) = ""
950   RBCRT011_ReturnValue(92) = ""
960   RBCRT011_ReturnValue(93) = ""
970   RBCRT011_ReturnValue(94) = ""
980   RBCRT011_ReturnValue(95) = ""
990   RBCRT011_ReturnValue(96) = "HbS-"  'Hemoglobin S negative
1000  RBCRT011_ReturnValue(97) = "parvovirus B19 antibody present"
1010  RBCRT011_ReturnValue(98) = "IgA deficient"
1020  RBCRT011_ReturnValue(99) = ""    'No Information Provided )

1030  sql = "Select * from options where description like 'RBCRT011_ReturnValue%' order by Listorder"
1040  Set tb = New Recordset
1050  RecOpenServer 0, tb, sql
1060  Do While Not tb.EOF
1070      Select Case Right(Trim(tb!Description), 2)
              Case "00"
1080              RBCRT011_ReturnValue(0) = tb!Contents & ""
1090          Case "01"
1100              RBCRT011_ReturnValue(1) = tb!Contents & ""
1110          Case "02"
1120              RBCRT011_ReturnValue(2) = tb!Contents & ""
1130          Case "03"
1140              RBCRT011_ReturnValue(3) = tb!Contents & ""
1150          Case "04"
1160              RBCRT011_ReturnValue(4) = tb!Contents & ""
1170          Case "05"
1180              RBCRT011_ReturnValue(5) = tb!Contents & ""
1190          Case "06"
1200              RBCRT011_ReturnValue(6) = tb!Contents & ""
1210          Case "07"
1220              RBCRT011_ReturnValue(7) = tb!Contents & ""
1230          Case "08"
1240              RBCRT011_ReturnValue(8) = tb!Contents & ""
1250          Case "09"
1260              RBCRT011_ReturnValue(9) = tb!Contents & ""
1270          Case "10"
1280              RBCRT011_ReturnValue(10) = tb!Contents & ""
1290          Case "11"
1300              RBCRT011_ReturnValue(11) = tb!Contents & ""
1310          Case "12"
1320              RBCRT011_ReturnValue(12) = tb!Contents & ""
1330          Case "13"
1340              RBCRT011_ReturnValue(13) = tb!Contents & ""
1350          Case "14"
1360              RBCRT011_ReturnValue(14) = tb!Contents & ""
1370          Case "15"
1380              RBCRT011_ReturnValue(15) = tb!Contents & ""
1390          Case "16"
1400              RBCRT011_ReturnValue(16) = tb!Contents & ""
1410          Case "17"
1420              RBCRT011_ReturnValue(17) = tb!Contents & ""
1430          Case "18"
1440              RBCRT011_ReturnValue(18) = tb!Contents & ""
1450          Case "19"
1460              RBCRT011_ReturnValue(19) = tb!Contents & ""
1470          Case "20"
1480              RBCRT011_ReturnValue(20) = tb!Contents & ""
1490          Case "21"
1500              RBCRT011_ReturnValue(21) = tb!Contents & ""
1510          Case "22"
1520              RBCRT011_ReturnValue(22) = tb!Contents & ""
1530          Case "23"
1540              RBCRT011_ReturnValue(23) = tb!Contents & ""
                  
1550          Case "24"
1560              RBCRT011_ReturnValue(24) = tb!Contents & ""
1570          Case "25"
1580              RBCRT011_ReturnValue(25) = tb!Contents & ""
1590          Case "26"
1600              RBCRT011_ReturnValue(26) = tb!Contents & ""
1610          Case "27"
1620              RBCRT011_ReturnValue(27) = tb!Contents & ""
1630          Case "28"
1640              RBCRT011_ReturnValue(28) = tb!Contents & ""
1650          Case "29"
1660              RBCRT011_ReturnValue(29) = tb!Contents & ""
1670          Case "30"
1680              RBCRT011_ReturnValue(30) = tb!Contents & ""
1690          Case "31"
1700              RBCRT011_ReturnValue(31) = tb!Contents & ""
1710          Case "32"
1720              RBCRT011_ReturnValue(32) = tb!Contents & ""
1730          Case "33"
1740              RBCRT011_ReturnValue(33) = tb!Contents & ""
1750          Case "34"
1760              RBCRT011_ReturnValue(34) = tb!Contents & ""
1770          Case "35"
1780              RBCRT011_ReturnValue(35) = tb!Contents & ""
1790          Case "36"
1800              RBCRT011_ReturnValue(36) = tb!Contents & ""
1810          Case "37"
1820              RBCRT011_ReturnValue(37) = tb!Contents & ""
1830          Case "38"
1840              RBCRT011_ReturnValue(38) = tb!Contents & ""
1850          Case "39"
1860              RBCRT011_ReturnValue(39) = tb!Contents & ""
1870          Case "40"
1880              RBCRT011_ReturnValue(40) = tb!Contents & ""
1890          Case "41"
1900              RBCRT011_ReturnValue(41) = tb!Contents & ""
1910          Case "42"
1920              RBCRT011_ReturnValue(42) = tb!Contents & ""
1930          Case "43"
1940              RBCRT011_ReturnValue(43) = tb!Contents & ""
1950          Case "44"
1960              RBCRT011_ReturnValue(44) = tb!Contents & ""
1970          Case "45"
1980              RBCRT011_ReturnValue(45) = tb!Contents & ""
1990          Case "46"
2000              RBCRT011_ReturnValue(46) = tb!Contents & ""
2010          Case "47"
2020              RBCRT011_ReturnValue(47) = tb!Contents & ""
2030          Case "48"
2040              RBCRT011_ReturnValue(48) = tb!Contents & ""
2050          Case "49"
2060              RBCRT011_ReturnValue(49) = tb!Contents & ""
2070          Case "50"
2080              RBCRT011_ReturnValue(50) = tb!Contents & ""
2090          Case "51"
2100              RBCRT011_ReturnValue(51) = tb!Contents & ""
2110          Case "52"
2120              RBCRT011_ReturnValue(52) = tb!Contents & ""
2130          Case "53"
2140              RBCRT011_ReturnValue(53) = tb!Contents & ""
2150          Case "54"
2160              RBCRT011_ReturnValue(54) = tb!Contents & ""
2170          Case "55"
2180              RBCRT011_ReturnValue(55) = tb!Contents & ""
2190          Case "56"
2200              RBCRT011_ReturnValue(56) = tb!Contents & ""
2210          Case "57"
2220              RBCRT011_ReturnValue(57) = tb!Contents & ""
2230          Case "58"
2240              RBCRT011_ReturnValue(58) = tb!Contents & ""
2250          Case "59"
2260              RBCRT011_ReturnValue(59) = tb!Contents & ""
2270          Case "60"
2280              RBCRT011_ReturnValue(60) = tb!Contents & ""
2290          Case "61"
2300              RBCRT011_ReturnValue(61) = tb!Contents & ""
2310          Case "62"
2320              RBCRT011_ReturnValue(62) = tb!Contents & ""
2330          Case "63"
2340              RBCRT011_ReturnValue(63) = tb!Contents & ""
2350          Case "64"
2360              RBCRT011_ReturnValue(64) = tb!Contents & ""
2370          Case "65"
2380              RBCRT011_ReturnValue(65) = tb!Contents & ""
2390          Case "66"
2400              RBCRT011_ReturnValue(66) = tb!Contents & ""
2410          Case "67"
2420              RBCRT011_ReturnValue(67) = tb!Contents & ""
2430          Case "68"
2440              RBCRT011_ReturnValue(68) = tb!Contents & ""
2450          Case "69"
2460              RBCRT011_ReturnValue(69) = tb!Contents & ""
2470          Case "70"
2480              RBCRT011_ReturnValue(70) = tb!Contents & ""
2490          Case "71"
2500              RBCRT011_ReturnValue(71) = tb!Contents & ""
2510          Case "72"
2520              RBCRT011_ReturnValue(72) = tb!Contents & ""
2530          Case "73"
2540              RBCRT011_ReturnValue(73) = tb!Contents & ""
2550          Case "74"
2560              RBCRT011_ReturnValue(74) = tb!Contents & ""
2570          Case "75"
2580              RBCRT011_ReturnValue(75) = tb!Contents & ""
2590          Case "76"
2600              RBCRT011_ReturnValue(76) = tb!Contents & ""
2610          Case "77"
2620              RBCRT011_ReturnValue(77) = tb!Contents & ""
2630          Case "78"
2640              RBCRT011_ReturnValue(78) = tb!Contents & ""
2650          Case "79"
2660              RBCRT011_ReturnValue(79) = tb!Contents & ""
2670          Case "80"
2680              RBCRT011_ReturnValue(80) = tb!Contents & ""
2690          Case "81"
2700              RBCRT011_ReturnValue(81) = tb!Contents & ""
2710          Case "82"
2720              RBCRT011_ReturnValue(82) = tb!Contents & ""
2730          Case "83"
2740              RBCRT011_ReturnValue(83) = tb!Contents & ""
2750          Case "84"
2760              RBCRT011_ReturnValue(84) = tb!Contents & ""
2770          Case "85"
2780              RBCRT011_ReturnValue(85) = tb!Contents & ""
2790          Case "86"
2800              RBCRT011_ReturnValue(86) = tb!Contents & ""
2810          Case "87"
2820              RBCRT011_ReturnValue(87) = tb!Contents & ""
2830          Case "88"
2840              RBCRT011_ReturnValue(88) = tb!Contents & ""
2850          Case "89"
2860              RBCRT011_ReturnValue(89) = tb!Contents & ""
2870          Case "90"
2880              RBCRT011_ReturnValue(90) = tb!Contents & ""
2890          Case "91"
2900              RBCRT011_ReturnValue(91) = tb!Contents & ""
2910          Case "92"
2920              RBCRT011_ReturnValue(92) = tb!Contents & ""
2930          Case "93"
2940              RBCRT011_ReturnValue(93) = tb!Contents & ""
2950          Case "94"
2960              RBCRT011_ReturnValue(94) = tb!Contents & ""
2970          Case "95"
2980              RBCRT011_ReturnValue(95) = tb!Contents & ""
2990          Case "96"
3000              RBCRT011_ReturnValue(96) = tb!Contents & ""
3010          Case "97"
3020              RBCRT011_ReturnValue(97) = tb!Contents & ""
3030          Case "98"
3040              RBCRT011_ReturnValue(98) = tb!Contents & ""
3050          Case "99"
3060              RBCRT011_ReturnValue(99) = tb!Contents & ""
3070      End Select
3080  tb.MoveNext
3090  Loop

3100  Exit Sub

Load_RBCRT011_Error:

       Dim strES As String
       Dim intEL As Integer

3110   intEL = Erl
3120   strES = Err.Description
3130   LogError "modOptions", "Load_RBCRT011", intEL, strES, sql

End Sub


Public Sub LoadOptions()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long

10    On Error GoTo LoadOptions_Error

20    ReDimOptions

      'For n = 0 To intOtherHospitalsInGroup
30    n = 0
40      sql = "Select * from Options " & _
              "order by ListOrder"
  
50      Set tb = New Recordset
60      RecOpenClient n, tb, sql
70      Do While Not tb.EOF
80        Select Case UCase$(Trim$(tb!Description & ""))
            Case "SOUNDCRITICAL": sysOptSoundCritical(n) = Trim$(tb!Contents & "")
90          Case "SOUNDINFORMATION": sysOptSoundInformation(n) = Trim$(tb!Contents & "")
100         Case "SOUNDQUESTION": sysOptSoundQuestion(n) = Trim$(tb!Contents & "")
110         Case "SOUNDSEVERE": sysOptSoundSevere(n) = Trim$(tb!Contents & "")
120         Case "TRANSFUSIONEXPIRY": sysOptTransfusionExpiry(n) = Trim$(tb!Contents & "")
130   Case "LASTPASWARNINGPERIOD":
140                   SysLastPASWarningPeriod = Val(tb!Contents & "")

150       End Select
160       tb.MoveNext
170     Loop
      'Next

180   Exit Sub

LoadOptions_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modOptions", "LoadOptions", intEL, strES, sql

End Sub


Private Sub ReDimOptions()

10    ReDim sysOptSoundCritical(0 To intOtherHospitalsInGroup) As String
20    ReDim sysOptSoundInformation(0 To intOtherHospitalsInGroup) As String
30    ReDim sysOptSoundQuestion(0 To intOtherHospitalsInGroup) As String
40    ReDim sysOptSoundSevere(0 To intOtherHospitalsInGroup) As String

50    ReDim sysOptTransfusionExpiry(0 To intOtherHospitalsInGroup)

End Sub


