VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form fgenotype 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Genotype"
   ClientHeight    =   4065
   ClientLeft      =   1275
   ClientTop       =   2460
   ClientWidth     =   7350
   ControlBox      =   0   'False
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
   Icon            =   "FGENOTYP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4065
   ScaleWidth      =   7350
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2145
      Left            =   2880
      TabIndex        =   17
      Top             =   300
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   3784
      _Version        =   393216
      Rows            =   9
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   2
      AllowUserResizing=   3
      FormatString    =   "^x/y / y/x |^Genotype |^Population Probability |^Patient Probability "
   End
   Begin VB.ComboBox crace 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3000
      Width           =   2535
   End
   Begin VB.ComboBox crace 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2700
      Width           =   2535
   End
   Begin VB.TextBox p 
      Height          =   285
      Index           =   4
      Left            =   1020
      TabIndex        =   10
      Top             =   1980
      Width           =   255
   End
   Begin VB.TextBox p 
      Height          =   285
      Index           =   3
      Left            =   1020
      TabIndex        =   9
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox p 
      Height          =   285
      Index           =   2
      Left            =   1020
      TabIndex        =   8
      Top             =   1380
      Width           =   255
   End
   Begin VB.TextBox p 
      Height          =   285
      Index           =   1
      Left            =   1020
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox p 
      Height          =   285
      Index           =   0
      Left            =   1020
      TabIndex        =   6
      Top             =   780
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   5910
      TabIndex        =   0
      Top             =   3345
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   225
      TabIndex        =   19
      Top             =   3870
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Lab Number"
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
      Left            =   2970
      TabIndex        =   18
      Top             =   2940
      Width           =   870
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Racial Type"
      Height          =   195
      Left            =   270
      TabIndex        =   16
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label llabnumber 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3885
      TabIndex        =   15
      Top             =   2910
      Width           =   2085
   End
   Begin VB.Label lpheno 
      Caption         =   "Phenotype"
      Height          =   555
      Left            =   1620
      TabIndex        =   12
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reactions"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   420
      TabIndex        =   11
      Top             =   420
      Width           =   870
   End
   Begin VB.Shape Shape1 
      Height          =   2115
      Left            =   240
      Top             =   300
      Width           =   2535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Anti-D"
      Height          =   195
      Left            =   420
      TabIndex        =   5
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Anti-e"
      Height          =   195
      Left            =   420
      TabIndex        =   4
      Top             =   1740
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Anti-E"
      Height          =   195
      Left            =   420
      TabIndex        =   3
      Top             =   1440
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Anti-c"
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   1140
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Anti-C"
      Height          =   195
      Left            =   420
      TabIndex        =   1
      Top             =   840
      Width           =   525
   End
End
Attribute VB_Name = "fgenotype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gf(0 To 1, 0 To 7) As Single

Private Sub cmdCancel_Click()


      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo cmdCancel_Click_Error

20    sql = "select * from genotype where labnumber='" & lLabNumber & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If tb.EOF Then tb.AddNew
60    tb!LabNumber = lLabNumber
70    For n = 0 To 4
80      tb(n + 1) = p(n)
90    Next
100   tb.Update

110   Unload Me

120   Exit Sub

cmdCancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "fgenotype", "cmdCancel_Click", intEL, strES, sql


End Sub

Private Sub CalcGeno(Pheno As String)

      Dim n As Integer
      Dim Temp As Single
10    ReDim pop(0 To 11) As Single
20    ReDim Pat(0 To 11) As Single
30    ReDim ph(0 To 7) As String
      Dim Diff As Integer

40    If crace(0).ListIndex = -1 Or crace(1).ListIndex = -1 Then
50      iMsg "Select racial type.", vbExclamation
60      If TimedOut Then Unload Me: Exit Sub
70      Exit Sub
80    End If

90    Diff = True
100   If crace(0).ListIndex = crace(1).ListIndex Then Diff = False

110   For n = 0 To 7
120     ph(n) = Choose(n + 1, "cde", "cdE", "cDe", "cDE", "Cde", "CdE", "CDe", "CDE")
130   Next

140   g.Rows = 2
150   g.AddItem ""
160   g.RemoveItem 1

170   Select Case Pheno
        Case "ccdee":
180       pop(0) = gf(0, 0) * gf(1, 0)
190       Pat(0) = 1
200       FillGridLine ph(0), ph(0), pop(0), Pat(0), 0
  
210     Case "ccdEe":
220       If Diff Then
230         pop(0) = gf(0, 1) * gf(1, 0)
240         pop(1) = gf(1, 0) * gf(1, 1)
250         Temp = pop(0) + pop(1)
260         Pat(0) = pop(0) / Temp
270         Pat(1) = pop(1) / Temp
280         FillGridLine ph(1), ph(0), pop(0), Pat(0), 0
290         FillGridLine ph(1), ph(0), pop(1), Pat(1), 1
300       Else
310         pop(0) = gf(0, 1) * gf(0, 0) * 2
320         Pat(0) = 1
330         FillGridLine ph(1), ph(0), pop(0), Pat(0), 0
340       End If

350     Case "ccdEE":
360       pop(0) = gf(0, 1) * gf(1, 1)
370       Pat(0) = 1
380       FillGridLine ph(1), ph(1), pop(0), Pat(0), 0
  
390     Case "ccDee":
400       If Diff Then
410         pop(0) = gf(0, 3) * gf(1, 0)
420         pop(1) = gf(0, 0) * gf(1, 3)
430         pop(2) = gf(0, 3) * gf(1, 2)
440         pop(3) = gf(0, 2) * gf(1, 3)
450         Temp = pop(0) + pop(1) + pop(2) + pop(3)
460         For n = 0 To 3: Pat(n) = pop(n) / Temp: Next
470         FillGridLine ph(2), ph(0), pop(0), Pat(0), 0
480         FillGridLine ph(2), ph(0), pop(1), Pat(1), 1
490         FillGridLine ph(2), ph(2), pop(2), Pat(2), 0
500         FillGridLine ph(2), ph(2), pop(3), Pat(3), 1
510       Else
520         pop(0) = gf(0, 3) * gf(0, 0) * 2
530         pop(1) = gf(0, 3) * gf(0, 2) * 2
540         Temp = pop(0) + pop(1)
550         For n = 0 To 1: Pat(n) = pop(n) / Temp: Next
560         FillGridLine ph(2), ph(0), pop(0), Pat(0), 0
570         FillGridLine ph(2), ph(2), pop(1), Pat(1), 0
580       End If

590     Case "ccDEe":
600       If Diff Then
610         pop(0) = gf(0, 3) * gf(1, 2)
620         pop(1) = gf(0, 2) * gf(1, 3)
630         pop(2) = gf(0, 3) * gf(1, 0)
640         pop(3) = gf(0, 0) * gf(1, 3)
650         pop(4) = gf(0, 2) * gf(1, 1)
660         pop(5) = gf(0, 1) * gf(1, 2)
670         Temp = pop(0) + pop(1) + pop(2) + pop(3) + pop(4) + pop(5)
680         For n = 0 To 5: Pat(n) = pop(n) / Temp: Next
690         FillGridLine ph(3), ph(2), pop(0), Pat(0), 0
700         FillGridLine ph(3), ph(2), pop(1), Pat(1), 1
710         FillGridLine ph(3), ph(0), pop(2), Pat(2), 0
720         FillGridLine ph(3), ph(0), pop(3), Pat(3), 1
730         FillGridLine ph(2), ph(1), pop(4), Pat(4), 0
740         FillGridLine ph(2), ph(1), pop(5), Pat(5), 1
750       Else
760         pop(0) = gf(0, 3) * gf(0, 2) * 2
770         pop(1) = gf(0, 3) * gf(0, 0) * 2
780         pop(2) = gf(0, 2) * gf(0, 1) * 2
790         Temp = pop(0) + pop(1) + pop(2)
800         For n = 0 To 2: Pat(n) = pop(n) / Temp: Next
810         FillGridLine ph(3), ph(2), pop(0), Pat(0), 0
820         FillGridLine ph(3), ph(0), pop(1), Pat(1), 0
830         FillGridLine ph(2), ph(1), pop(2), Pat(2), 0
840       End If

850     Case "ccDEE":
860       If Diff Then
870         pop(0) = gf(0, 3) * gf(1, 3)
880         pop(1) = gf(0, 3) * gf(1, 1)
890         pop(2) = gf(0, 1) * gf(1, 3)
900         Temp = pop(0) + pop(1) + pop(2)
910         For n = 0 To 3: Pat(n) = pop(n) / Temp: Next
920         FillGridLine ph(3), ph(3), pop(0), Pat(0), 0
930         FillGridLine ph(3), ph(1), pop(1), Pat(1), 0
940         FillGridLine ph(3), ph(1), pop(2), Pat(2), 1
950       Else
960         pop(0) = gf(0, 3) * gf(0, 3)
970         pop(1) = gf(0, 3) * gf(0, 1) * 2
980         Temp = pop(0) + pop(1)
990         Pat(0) = pop(0) / Temp
1000        Pat(1) = pop(1) / Temp
1010        FillGridLine ph(3), ph(3), pop(0), Pat(0), 0
1020        FillGridLine ph(3), ph(1), pop(1), Pat(1), 0
1030       End If

1040    Case "Ccdee":
1050      If Diff Then
1060        pop(0) = gf(0, 4) * gf(1, 0)
1070        pop(1) = gf(0, 0) * gf(1, 4)
1080        Temp = pop(0) + pop(1)
1090        For n = 0 To 1: Pat(n) = pop(n) / Temp: Next
1100        FillGridLine ph(4), ph(0), pop(0), Pat(0), 0
1110        FillGridLine ph(4), ph(0), pop(1), Pat(1), 1
1120      Else
1130        pop(0) = gf(0, 4) * gf(0, 0) * 2
1140        Pat(0) = 1
1150        FillGridLine ph(4), ph(0), pop(0), Pat(0), 0
1160      End If

1170    Case "CcdEe":
1180      If Diff Then
1190        pop(0) = gf(0, 4) * gf(1, 1)
1200        pop(1) = gf(0, 1) * gf(1, 4)
1210        pop(2) = gf(0, 5) * gf(1, 0)
1220        pop(3) = gf(0, 0) * gf(1, 5)
1230        Temp = pop(0) + pop(1) + pop(2) + pop(3)
1240        For n = 0 To 3: Pat(n) = pop(n) / Temp: Next
1250        FillGridLine ph(4), ph(1), pop(0), Pat(0), 0
1260        FillGridLine ph(4), ph(1), pop(1), Pat(1), 1
1270        FillGridLine ph(5), ph(0), pop(2), Pat(2), 0
1280        FillGridLine ph(5), ph(0), pop(3), Pat(3), 1
1290      Else
1300        pop(0) = gf(0, 4) * gf(0, 1) * 2
1310        pop(1) = gf(0, 5) * gf(0, 0) * 2
1320        Temp = pop(0) + pop(1)
1330        Pat(0) = pop(0) / Temp
1340        Pat(1) = pop(1) / Temp
1350        FillGridLine ph(4), ph(1), pop(0), Pat(0), 0
1360        FillGridLine ph(5), ph(0), pop(1), Pat(1), 0
1370      End If

1380    Case "CcdEE":
1390      If Diff Then
1400        pop(0) = gf(0, 5) * gf(1, 1)
1410        pop(1) = gf(0, 1) * gf(1, 5)
1420        Temp = pop(0) + pop(1)
1430        For n = 0 To 1: Pat(n) = pop(n) / Temp: Next
1440        FillGridLine ph(5), ph(1), pop(0), Pat(0), 0
1450        FillGridLine ph(5), ph(1), pop(1), Pat(1), 1
1460      Else
1470        pop(0) = gf(0, 5) * gf(0, 1) * 2
1480        Pat(0) = 1
1490        FillGridLine ph(5), ph(1), pop(0), Pat(0), 0
1500      End If

1510    Case "CcDee":
1520      If Diff Then
1530        pop(0) = gf(0, 6) * gf(1, 0)
1540        pop(1) = gf(0, 0) * gf(1, 6)
1550        pop(2) = gf(0, 6) * gf(1, 2)
1560        pop(3) = gf(0, 2) * gf(1, 6)
1570        Temp = pop(0) + pop(1) + pop(2) + pop(3)
1580        For n = 0 To 3: Pat(n) = pop(n) / Temp: Next
1590        FillGridLine ph(6), ph(0), pop(0), Pat(0), 0
1600        FillGridLine ph(6), ph(0), pop(1), Pat(1), 1
1610        FillGridLine ph(6), ph(2), pop(2), Pat(2), 0
1620        FillGridLine ph(6), ph(2), pop(3), Pat(3), 1
1630      Else
1640        pop(0) = gf(0, 6) * gf(0, 0) * 2
1650        pop(1) = gf(0, 6) * gf(0, 2) * 2
1660        Temp = pop(0) + pop(1)
1670        Pat(0) = pop(0) / Temp
1680        Pat(1) = pop(1) / Temp
1690        FillGridLine ph(6), ph(0), pop(0), Pat(0), 0
1700        FillGridLine ph(6), ph(2), pop(1), Pat(1), 0
1710      End If

1720    Case "CcDEe":
1730      If Diff Then
1740        pop(0) = gf(0, 6) * gf(1, 3)
1750        pop(1) = gf(0, 3) * gf(1, 6)
1760        pop(2) = gf(0, 6) * gf(1, 1)
1770        pop(3) = gf(0, 1) * gf(1, 6)
1780        pop(4) = gf(0, 7) * gf(1, 2)
1790        pop(5) = gf(0, 2) * gf(1, 7)
1800        pop(6) = gf(0, 7) * gf(1, 0)
1810        pop(7) = gf(0, 0) * gf(1, 7)
1820        pop(8) = gf(0, 4) * gf(1, 3)
1830        pop(9) = gf(0, 3) * gf(1, 4)
1840        pop(10) = gf(0, 5) * gf(1, 2)
1850        pop(11) = gf(0, 2) * gf(1, 5)
1860        Temp = 0
1870        For n = 0 To 11: Temp = Temp + pop(n): Next
1880        For n = 0 To 11: Pat(n) = pop(n) / Temp: Next
1890        FillGridLine ph(6), ph(3), pop(0), Pat(0), 0
1900        FillGridLine ph(6), ph(3), pop(1), Pat(1), 1
1910        FillGridLine ph(6), ph(1), pop(2), Pat(2), 0
1920        FillGridLine ph(6), ph(1), pop(3), Pat(3), 1
1930        FillGridLine ph(7), ph(2), pop(4), Pat(4), 0
1940        FillGridLine ph(7), ph(2), pop(5), Pat(5), 1
1950        FillGridLine ph(7), ph(0), pop(6), Pat(6), 0
1960        FillGridLine ph(7), ph(0), pop(7), Pat(7), 1
1970        FillGridLine ph(4), ph(3), pop(8), Pat(8), 0
1980        FillGridLine ph(4), ph(3), pop(9), Pat(9), 1
1990        FillGridLine ph(5), ph(2), pop(10), Pat(10), 0
2000        FillGridLine ph(5), ph(2), pop(11), Pat(11), 1
2010      Else
2020        pop(0) = gf(0, 6) * gf(0, 3) * 2
2030        pop(1) = gf(0, 6) * gf(0, 1) * 2
2040        pop(2) = gf(0, 7) * gf(0, 2) * 2
2050        pop(3) = gf(0, 7) * gf(0, 0) * 2
2060        pop(4) = gf(0, 4) * gf(0, 3) * 2
2070        pop(5) = gf(0, 5) * gf(0, 2) * 2
2080        Temp = pop(0) + pop(1) + pop(2) + pop(3) + pop(4) + pop(5)
2090        For n = 0 To 5: Pat(n) = pop(n) / Temp: Next
2100        FillGridLine ph(6), ph(3), pop(0), Pat(0), 0
2110        FillGridLine ph(6), ph(1), pop(1), Pat(1), 0
2120        FillGridLine ph(7), ph(2), pop(2), Pat(2), 0
2130        FillGridLine ph(7), ph(0), pop(3), Pat(3), 0
2140        FillGridLine ph(4), ph(3), pop(4), Pat(4), 0
2150        FillGridLine ph(5), ph(2), pop(5), Pat(5), 0
2160      End If

2170    Case "CcDEE":
2180      If Diff Then
2190        pop(0) = gf(0, 7) * gf(1, 3)
2200        pop(1) = gf(0, 3) * gf(1, 7)
2210        pop(2) = gf(0, 7) * gf(1, 1)
2220        pop(3) = gf(0, 1) * gf(1, 7)
2230        pop(4) = gf(0, 5) * gf(1, 3)
2240        pop(5) = gf(0, 3) * gf(1, 5)
2250        Temp = pop(0) + pop(1) + pop(2) + pop(3) + pop(4) + pop(5)
2260        For n = 0 To 5: Pat(n) = pop(n) / Temp: Next
2270        FillGridLine ph(7), ph(3), pop(0), Pat(0), 0
2280        FillGridLine ph(7), ph(3), pop(1), Pat(1), 1
2290        FillGridLine ph(7), ph(1), pop(2), Pat(2), 0
2300        FillGridLine ph(7), ph(1), pop(3), Pat(3), 1
2310        FillGridLine ph(5), ph(3), pop(4), Pat(4), 0
2320        FillGridLine ph(5), ph(3), pop(5), Pat(5), 1
2330      Else
2340        pop(0) = gf(0, 7) * gf(0, 3) * 2
2350        pop(1) = gf(0, 7) * gf(0, 1) * 2
2360        pop(2) = gf(0, 5) * gf(0, 3) * 2
2370        Temp = pop(0) + pop(1) + pop(2)
2380        Pat(0) = pop(0) / Temp
2390        Pat(1) = pop(1) / Temp
2400        Pat(2) = pop(2) / Temp
2410        FillGridLine ph(7), ph(3), pop(0), Pat(0), 0
2420        FillGridLine ph(7), ph(1), pop(1), Pat(1), 0
2430        FillGridLine ph(5), ph(3), pop(2), Pat(2), 0
2440      End If

2450    Case "CCdee":
2460      pop(0) = gf(0, 4) * gf(1, 4)
2470      Pat(0) = 1
2480      FillGridLine ph(4), ph(4), pop(0), Pat(0), 0
  
2490    Case "CCdEe":
2500      If Diff Then
2510        pop(0) = gf(0, 5) * gf(1, 4)
2520        pop(1) = gf(1, 4) * gf(1, 5)
2530        Temp = pop(0) + pop(1)
2540        For n = 0 To 1: Pat(n) = pop(n) / Temp: Next
2550        FillGridLine ph(5), ph(4), pop(0), Pat(0), 0
2560        FillGridLine ph(5), ph(4), pop(1), Pat(1), 1
2570      Else
2580        pop(0) = gf(0, 5) * gf(0, 4) * 2
2590        Pat(0) = 1
2600        FillGridLine ph(5), ph(4), pop(0), Pat(0), 0
2610      End If

2620    Case "CCdEE":
2630      pop(0) = gf(0, 5) * gf(1, 5)
2640      Pat(0) = 1
2650      FillGridLine ph(5), ph(5), pop(0), Pat(0), 0
  
2660    Case "CCDee":
2670      If Diff Then
2680        pop(0) = gf(0, 6) * gf(1, 6)
2690        pop(1) = gf(0, 6) * gf(1, 4)
2700        pop(2) = gf(0, 4) * gf(1, 6)
2710        Temp = pop(0) + pop(1) + pop(2)
2720        For n = 0 To 2: Pat(n) = pop(n) / Temp: Next
2730        FillGridLine ph(6), ph(6), pop(0), Pat(0), 0
2740        FillGridLine ph(6), ph(4), pop(1), Pat(1), 0
2750        FillGridLine ph(6), ph(4), pop(2), Pat(2), 1
2760      Else
2770        pop(0) = gf(0, 6) * gf(0, 6)
2780        pop(1) = gf(0, 6) * gf(0, 4) * 2
2790        Temp = pop(0) + pop(1)
2800        Pat(0) = pop(0) / Temp
2810        Pat(1) = pop(1) / Temp
2820        FillGridLine ph(6), ph(6), pop(0), Pat(0), 0
2830        FillGridLine ph(6), ph(4), pop(1), Pat(1), 0
2840      End If

2850    Case "CCDEe":
2860      If Diff Then
2870        pop(0) = gf(0, 7) * gf(1, 6)
2880        pop(1) = gf(0, 6) * gf(1, 7)
2890        pop(2) = gf(0, 7) * gf(1, 4)
2900        pop(3) = gf(0, 4) * gf(1, 7)
2910        pop(4) = gf(0, 5) * gf(1, 6)
2920        pop(5) = gf(0, 6) * gf(1, 5)
2930        Temp = pop(0) + pop(1) + pop(2) + pop(3) + pop(4) + pop(5)
2940        For n = 0 To 5: Pat(n) = pop(n) / Temp: Next
2950        FillGridLine ph(7), ph(6), pop(0), Pat(0), 0
2960        FillGridLine ph(7), ph(6), pop(1), Pat(1), 1
2970        FillGridLine ph(7), ph(4), pop(2), Pat(2), 0
2980        FillGridLine ph(7), ph(4), pop(3), Pat(3), 1
2990        FillGridLine ph(6), ph(5), pop(4), Pat(4), 0
3000        FillGridLine ph(6), ph(5), pop(5), Pat(5), 1
3010      Else
3020        pop(0) = gf(0, 7) * gf(0, 6) * 2
3030        pop(1) = gf(0, 7) * gf(0, 4) * 2
3040        pop(2) = gf(0, 6) * gf(0, 5) * 2
3050        Temp = pop(0) + pop(1) + pop(2)
3060        For n = 0 To 2: Pat(n) = pop(n) / Temp: Next
3070        FillGridLine ph(7), ph(6), pop(0), Pat(0), 0
3080        FillGridLine ph(7), ph(4), pop(1), Pat(1), 0
3090        FillGridLine ph(6), ph(5), pop(2), Pat(2), 0
3100      End If

3110    Case "CCDEE":
3120      If Diff Then
3130        pop(0) = gf(0, 7) * gf(1, 7)
3140        pop(1) = gf(0, 7) * gf(1, 5)
3150        pop(2) = gf(0, 5) * gf(1, 7)
3160        Temp = pop(0) + pop(1) + pop(2)
3170        For n = 0 To 2: Pat(n) = pop(n) / Temp: Next
3180        FillGridLine ph(7), ph(7), pop(0), Pat(0), 0
3190        FillGridLine ph(7), ph(5), pop(1), Pat(1), 0
3200        FillGridLine ph(7), ph(5), pop(2), Pat(2), 1
3210      Else
3220        pop(0) = gf(0, 7) * gf(0, 7)
3230        pop(1) = gf(0, 7) * gf(0, 5) * 2
3240        Temp = pop(0) + pop(1)
3250        Pat(0) = pop(0) / Temp
3260        Pat(1) = pop(1) / Temp
3270        FillGridLine ph(7), ph(7), pop(0), Pat(0), 0
3280        FillGridLine ph(7), ph(5), pop(1), Pat(1), 0
3290      End If

3300  End Select

3310  If g.Rows > 2 Then
3320    g.RemoveItem 1
3330    g.col = 3
3340    g.Sort = flexSortNumericDescending
3350  End If

End Sub

Private Sub crace_Click(Index As Integer)


      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo crace_Click_Error

20    sql = "Select * from GeneFrequency where " & _
            "Race = '" & crace(Index).Text & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If tb.EOF Then
60      iMsg "Details not found.", vbInformation
70      If TimedOut Then Unload Me: Exit Sub
80      Exit Sub
90    End If

100   For n = 0 To 7
110    gf(Index, n) = tb(n + 1)
120   Next

130   If Index = 0 Then
140     crace(1).ListIndex = crace(0).ListIndex
150   End If

160   Call p_change(0)

170   Exit Sub

crace_Click_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "fgenotype", "crace_Click", intEL, strES, sql


End Sub

Private Sub FillGridLine(p1 As String, _
                         p2 As String, _
                         pop As Single, _
                         Pat As Single, _
                         Flag As Integer)

      Dim s As String

10    If Flag = 0 Then
20      s = Left$(crace(0), 3) & "/" & Left$(crace(1), 3)
30    Else
40      s = Left$(crace(1), 3) & "/" & Left$(crace(0), 3)
50    End If

60    s = s & vbTab & p1 & "/" & p2 & vbTab & _
              Format(pop, "##0.0000%") & vbTab & _
              Format(Pat, "##0.0000%")
70    g.AddItem s

End Sub

Private Sub Form_Load()


      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo Form_Load_Error

20    sql = "Select * from GeneFrequency"

30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    Do While Not tb.EOF
60      crace(0).AddItem tb!Race & ""
70      crace(1).AddItem tb!Race & ""
80      tb.MoveNext
90    Loop

100   g.RowHeight(0) = 480
110   g.ColWidth(2) = 1000
120   g.ColWidth(3) = 1000

130   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "fgenotype", "Form_Load", intEL, strES, sql


End Sub

Private Sub p_change(Index As Integer)

      Dim n As Integer
      Dim s As String

10    For n = 0 To 4
20      If Trim$(p(n)) = "" Then Exit Sub
30      If p(n) <> "O" And p(n) <> "+" Then Exit Sub
40    Next

50    If p(0) <> "+" And p(1) <> "+" Then
60      lpheno = ""
70      Exit Sub
80    End If

90    If p(2) <> "+" And p(3) <> "+" Then
100     lpheno = ""
110     Exit Sub
120   End If

130   If p(0) = "+" Then s = "C" Else s = "c"
140   If p(1) = "+" Then s = s & "c" Else s = s & "C"
150   If p(4) = "+" Then s = s & "D" Else s = s & "d"
160   If p(2) = "+" Then s = s & "E" Else s = s & "e"
170   If p(3) = "+" Then s = s & "e" Else s = s & "E"

180   lpheno = "Phenotype " & s

190   CalcGeno s

End Sub

Private Sub p_Click(Index As Integer)

10    Select Case Trim$(p(Index))
        Case "": p(Index) = "O"
20      Case "O": p(Index) = "+"
30      Case "+": p(Index) = "O"
40      Case Else: p(Index) = "O"
50    End Select

End Sub

