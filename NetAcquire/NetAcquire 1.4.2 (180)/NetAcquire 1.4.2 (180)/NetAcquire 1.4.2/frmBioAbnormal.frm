VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmBioAbnormals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire 6.9 - Search for Abnormals - Biochemistry"
   ClientHeight    =   7410
   ClientLeft      =   195
   ClientTop       =   480
   ClientWidth     =   10605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   3630
      Picture         =   "frmBioAbnormal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6030
      Width           =   975
   End
   Begin VB.CommandButton cmdPrintList 
      Caption         =   "&Print List"
      Height          =   735
      Left            =   1140
      Picture         =   "frmBioAbnormal.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6030
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   2370
      Picture         =   "frmBioAbnormal.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6030
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7155
      Left            =   5190
      TabIndex        =   30
      Top             =   150
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   12621
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Run Date       |<Sample ID   |<Chart #    |<Age    |<Result   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Analyte"
      Height          =   3555
      Left            =   90
      TabIndex        =   1
      Top             =   1470
      Width           =   4965
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "&Recalculate"
         Height          =   885
         Left            =   3450
         Picture         =   "frmBioAbnormal.frx":0FDE
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2430
         Width           =   1005
      End
      Begin VB.TextBox tHigh 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3540
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1200
         Width           =   825
      End
      Begin VB.TextBox tLow 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3540
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1500
         Width           =   825
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   5
         Left            =   2250
         TabIndex        =   9
         Top             =   2490
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         Value           =   -1  'True
         PictureDnChange =   2
         PictureUp       =   "frmBioAbnormal.frx":12E8
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   4
         Left            =   1395
         TabIndex        =   8
         Top             =   2490
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmBioAbnormal.frx":18A6
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   3
         Left            =   540
         TabIndex        =   7
         Top             =   2490
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmBioAbnormal.frx":1ED8
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   2
         Left            =   2220
         TabIndex        =   6
         Top             =   900
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmBioAbnormal.frx":2496
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   1
         Left            =   1395
         TabIndex        =   5
         Top             =   900
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmBioAbnormal.frx":2A54
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   4
         Top             =   900
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmBioAbnormal.frx":3086
      End
      Begin VB.ComboBox lAnalyte 
         Height          =   315
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "De-Select Range"
         Height          =   195
         Left            =   3390
         TabIndex        =   27
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Flag Ranges"
         Height          =   195
         Left            =   1260
         TabIndex        =   26
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   2250
         TabIndex        =   25
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   2850
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   1410
         TabIndex        =   21
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   540
         TabIndex        =   20
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   3090
         Width           =   300
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   2250
         TabIndex        =   18
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   540
         TabIndex        =   17
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   1410
         TabIndex        =   16
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   2220
         TabIndex        =   15
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1410
         TabIndex        =   14
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   13
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   2220
         TabIndex        =   12
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1410
         TabIndex        =   11
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   10
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Normal Ranges"
         Height          =   195
         Left            =   1170
         TabIndex        =   3
         Top             =   690
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   765
      Left            =   870
      TabIndex        =   0
      Top             =   180
      Width           =   3315
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1740
         TabIndex        =   35
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219021313
         CurrentDate     =   38055
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   34
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219021313
         CurrentDate     =   38055
      End
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   195
      Left            =   90
      TabIndex        =   36
      Top             =   1110
      Visible         =   0   'False
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3450
      TabIndex        =   42
      Top             =   6780
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Over Range"
      Height          =   195
      Left            =   2760
      TabIndex        =   40
      Top             =   5490
      Width           =   885
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Under Range"
      Height          =   195
      Left            =   210
      TabIndex        =   39
      Top             =   5490
      Width           =   975
   End
   Begin VB.Label lOver 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3660
      TabIndex        =   38
      Top             =   5430
      Width           =   900
   End
   Begin VB.Label lUnder 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   37
      Top             =   5430
      Width           =   900
   End
End
Attribute VB_Name = "frmBioAbnormals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  '© Custom Software 2001


Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim TestNumber As String
          Dim s As String
          Dim DP As Integer
          Dim strFormat As String

1440      On Error GoTo FillG_Error

1450      g.Rows = 2
1460      g.AddItem ""
1470      g.RemoveItem 1
1480      lUnder = ""
1490      lOver = ""

1500      If lAnalyte = "" Then Exit Sub

1510      TestNumber = ""
1520      sql = "Select Code, DP from BioTestDefinitions where " & _
              "LongName = '" & lAnalyte & "'"
1530      Set tb = New Recordset
1540      RecOpenClient 0, tb, sql
1550      If Not tb.EOF Then
1560          TestNumber = Trim$(tb!Code & "")
1570          Select Case tb!DP
                  Case 0: strFormat = "####0"
1580              Case 1: strFormat = "###0.0"
1590              Case 2: strFormat = "##0.00"
1600              Case 3: strFormat = "#0.000"
1610              Case 4: strFormat = "0.0000"
1620              Case Else: strFormat = "###0.0"
1630          End Select
1640      End If

1650      If TestNumber = "" Then Exit Sub

1660      Screen.MousePointer = 11

1670      Cnxn(0).Execute "DROP TABLE NoCharResults"

1680      Cnxn(0).Execute "SELECT B.RunTime, B.Result, B.SampleID, " & _
              "D.Age, D.Chart " & _
              "INTO NoCharResults from BioResults as B, Demographics as D " & _
              "WHERE " & _
              "D.SampleID = B.SampleID " & _
              "and (B.RunDate between '" & _
              Format(dtFrom, "dd/mmm/yyyy") & "' and '" & _
              Format(dtTo, "dd/mmm/yyyy") & "') " & _
              "and Code = '" & TestNumber & "' " & _
              "and Result like '%[0-9]%' " & _
              "and Result <> '' " & _
              "and Result not like '%[^0-9. ]%' " & _
              "order by B.RunTime "

1690      sql = "select * from NoCharResults where " & _
              "cast(Result as Float) < " & Val(tLow) & " " & _
              "or cast(Result as Float) > " & Val(tHigh)
1700      Set tb = New Recordset
1710      RecOpenClient 0, tb, sql

1720      If tb.EOF Then
1730          Screen.MousePointer = 0
1740          Exit Sub
1750      End If

1760      g.Visible = False

1770      pb.Visible = True
1780      pb.max = tb.RecordCount
1790      pb = 0
1800      g.Col = 4
1810      Do While Not tb.EOF
1820          pb = pb + 1
1830          s = Format(tb!RunTime, "dd/MM/yyyy") & vbTab & _
                  tb!SampleID & vbTab & _
                  tb!Chart & vbTab & _
                  tb!Age & vbTab & _
                  Format$(tb!Result, strFormat)
1840          g.AddItem s
1850          g.row = g.Rows - 1
1860          If tb!Result < Val(tLow) Then
1870              g.CellBackColor = &HFFFFC0
1880              lUnder = Val(lUnder) + 1
1890          Else
1900              g.CellBackColor = &H8080FF
1910              lOver = Val(lOver) + 1
1920          End If
1930          tb.MoveNext
1940      Loop

1950      g.Visible = True
1960      pb.Visible = False

1970      If g.Rows > 2 Then
1980          g.RemoveItem 1
1990      End If

          'CalcOutsideRange

2000      Screen.MousePointer = 0

2010      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

2020      intEL = Erl
2030      strES = Err.Description
2040      LogError "frmBioAbnormals", "FillG", intEL, strES, sql

End Sub

Private Sub b_Click(Index As Integer, Value As Integer)

2050      If Value = False Then
2060          tLow.Enabled = True
2070          tHigh.Enabled = True
2080          tLow.SelStart = 0
2090          tLow.SelLength = Len(tLow)
2100          tLow.SetFocus
2110      Else
2120          tLow.Enabled = False
2130          tHigh.Enabled = False
2140          tLow = l((Index * 2) + 1)
2150          tHigh = l(Index * 2)
2160          FillG
2170      End If

End Sub

Private Sub cmdCancel_Click()

2180      Unload Me

End Sub


Private Sub cmdPrintList_Click()

          Dim n As Integer
          Dim s As String

2190      On Error GoTo ehCPL

2200      For n = 0 To g.Rows - 1
2210          s = g.TextMatrix(n, 0) & vbTab & _
                  g.TextMatrix(n, 1) & vbTab & _
                  g.TextMatrix(n, 2) & vbTab & _
                  g.TextMatrix(n, 3) & vbTab & _
                  g.TextMatrix(n, 4)
2220          Printer.Print s
2230      Next

2240      Printer.EndDoc

2250      Exit Sub

ehCPL:

          Dim er As Long
          Dim es As String

2260      er = Err.Number
2270      es = Err.Description
2280      iMsg es
2290      Exit Sub

End Sub

Private Sub cmdRecalc_Click()

2300      FillG

End Sub

Private Sub cmdXL_Click()

2310      ExportFlexGrid g, Me

End Sub

Private Sub dtFrom_CloseUp()

2320      FillG

End Sub

Private Sub dtTo_CloseUp()

2330      FillG

End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

2340      On Error GoTo Form_Load_Error

2350      dtFrom = Format$(DateAdd("m", -1, Now), "dd/mm/yyyy")
2360      dtTo = Format$(Now, "dd/mm/yyyy")

2370      lAnalyte.Clear
2380      sql = "Select distinct LongName, PrintPriority " & _
              "from BioTestDefinitions " & _
              "order by PrintPriority"
2390      Set tb = New Recordset
2400      RecOpenServer 0, tb, sql
2410      Do While Not tb.EOF
2420          lAnalyte.AddItem tb!LongName
2430          tb.MoveNext
2440      Loop

2450      Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

2460      intEL = Erl
2470      strES = Err.Description
2480      LogError "frmBioAbnormals", "Form_Load", intEL, strES, sql

        
End Sub


Private Sub g_Click()

          Dim tb As Recordset
          Dim s As String
          Dim sql As String

2490      On Error GoTo g_Click_Error

2500      If g.MouseRow = 0 Then Exit Sub
2510      If Val(g.TextMatrix(g.row, 1)) = 0 Then Exit Sub

2520      sql = "Select PatName, Chart, DoB from Demographics where " & _
              "SampleID = '" & g.TextMatrix(g.row, 1) & "'"
2530      Set tb = New Recordset
2540      RecOpenServer 0, tb, sql

2550      If Not tb.EOF Then
2560          s = "    Sample ID : " & g & vbCrLf & _
                  "      Patient : " & tb!PatName & vbCrLf & _
                  "        Chart : " & tb!Chart & vbCrLf & _
                  "Date of Birth : " & Format(tb!DoB, "dd/MM/yyyy") & ""
2570          tb.Close
2580          iMsg s, vbInformation
2590      End If


2600      Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

2610      intEL = Erl
2620      strES = Err.Description
2630      LogError "frmBioAbnormals", "g_Click", intEL, strES, sql

       
End Sub

Private Sub lanalyte_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

2640      On Error GoTo lanalyte_Click_Error

2650      sql = "Select * from BioTestDefinitions where " & _
              "LongName = '" & lAnalyte & "'"
2660      Set tb = New Recordset
2670      RecOpenServer 0, tb, sql
2680      If Not tb.EOF Then
2690          With tb
2700              l(0) = !MaleHigh
2710              l(1) = !MaleLow
2720              l(2) = !FemaleHigh
2730              l(3) = !FemaleLow
2740              l(4) = IIf(Val(!MaleHigh) > Val(!FemaleHigh), !MaleHigh, !FemaleHigh)
2750              l(5) = IIf(Val(!MaleLow) < Val(!FemaleLow), !MaleLow, !FemaleLow)
          
2760              l(6) = !FlagMaleHigh
2770              l(7) = !FlagMaleLow
2780              l(8) = !FlagFemaleHigh
2790              l(9) = !FlagFemaleLow
2800              l(10) = IIf(Val(!FlagMaleHigh) > Val(!FlagFemaleHigh), !FlagMaleHigh, !FlagFemaleHigh)
2810              l(11) = IIf(Val(!FlagMaleLow) < Val(!FlagFemaleLow), !FlagMaleLow, !FlagFemaleLow)
2820          End With
2830      End If

2840      For n = 0 To 5
2850          If b(n) Then
2860              tLow = l((n * 2) + 1)
2870              tHigh = l(n * 2)
2880          End If
2890      Next

2900      FillG

2910      Exit Sub

lanalyte_Click_Error:

          Dim strES As String
          Dim intEL As Integer

2920      intEL = Erl
2930      strES = Err.Description
2940      LogError "frmBioAbnormals", "lanalyte_Click", intEL, strES, sql


End Sub





