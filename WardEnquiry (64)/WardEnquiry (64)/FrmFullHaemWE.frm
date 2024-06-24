VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullHaemWE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Haematology History"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   11190
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
   HelpContextID   =   10028
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7875
   ScaleWidth      =   11190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9780
      Picture         =   "FrmFullHaemWE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5370
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plot between"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   150
      TabIndex        =   11
      Top             =   5280
      Width           =   1785
      Begin VB.ComboBox cmbPlotTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Text            =   "cmbPlotTo"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cmbPlotFrom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Text            =   "cmbPlotFrom"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H80000016&
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   570
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1830
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   150
         TabIndex        =   20
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   150
         TabIndex        =   19
         Top             =   270
         Width           =   345
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   3780
      TabIndex        =   10
      Top             =   750
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7530
      Top             =   510
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   2325
      Left            =   2040
      ScaleHeight     =   2265
      ScaleWidth      =   6435
      TabIndex        =   8
      Top             =   5370
      Width           =   6495
      Begin VB.Label lblTest 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2490
         TabIndex        =   21
         Top             =   90
         Width           =   75
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4185
      Left            =   180
      TabIndex        =   7
      Top             =   960
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7382
      _Version        =   393216
      Rows            =   4
      FixedRows       =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9780
      Picture         =   "FrmFullHaemWE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7170
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdDemog 
      Height          =   2175
      Left            =   11460
      TabIndex        =   18
      Top             =   1620
      Visible         =   0   'False
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "<SampleID   |<Cnxn    |<RunDate    |<TimeTaken                  "
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   9780
      TabIndex        =   23
      Top             =   6030
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   8580
      TabIndex        =   17
      Top             =   7470
      Width           =   510
   End
   Begin VB.Label lblMeanVal 
      Alignment       =   2  'Center
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
      Left            =   8580
      TabIndex        =   16
      Top             =   6450
      Width           =   510
   End
   Begin VB.Label lblMaxVal 
      Alignment       =   2  'Center
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
      Left            =   8580
      TabIndex        =   15
      Top             =   5400
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "FrmFullHaemWE.frx":0974
      Top             =   450
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on Parameter to show Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   9
      Top             =   150
      Width           =   2865
   End
   Begin VB.Label lblChartTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "NOPAS"
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
      Left            =   3180
      TabIndex        =   6
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
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
      Left            =   5580
      TabIndex        =   5
      Top             =   180
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3780
      TabIndex        =   4
      Top             =   480
      Width           =   3705
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5940
      TabIndex        =   3
      Top             =   150
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name"
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
      Left            =   3330
      TabIndex        =   2
      Top             =   510
      Width           =   420
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3780
      TabIndex        =   1
      Top             =   150
      Width           =   1545
   End
End
Attribute VB_Name = "frmFullHaemWE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ChartPosition
    xPos As Long
    yPos As Long
    Value As Single
    Date As String
End Type

Private ChartPositions() As ChartPosition

Private NumberOfDays As Long


Private Sub cmdXL_Click()

ExportFlexGrid g, Me

End Sub

Private Sub DrawChart()

          Dim n As Integer
          Dim Counter As Integer
          Dim DaysInterval As Long
          Dim x As Integer
          Dim y As Integer
          Dim PixelsPerDay As Single
          Dim PixelsPerPointY As Single
          Dim FirstDayFilled As Boolean
          Dim MaxVal As Single
          Dim cVal As Single
          Dim StartGridX As Integer
          Dim StopGridX As Integer
          Dim plottedPoints As Object
10        Set plottedPoints = CreateObject("Scripting.Dictionary")
          Dim daysCount As Integer
          
          Dim Count As Integer

20        On Error GoTo DrawChart_Error

30        MaxVal = 0
40        lblMaxVal = ""
50        lblMeanVal = ""
60        lblTest = g.TextMatrix(g.Row, 0)

70        PB.Cls
80        PB.Picture = LoadPicture("")

90        daysCount = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
100       NumberOfDays = g.Cols - 2
110       Count = NumberOfDays
120       If NumberOfDays < 2 Then Exit Sub

130       ReDim ChartPositions(0 To NumberOfDays)

140       For n = 1 To NumberOfDays
150           ChartPositions(n).xPos = 0
160           ChartPositions(n).yPos = 0
170           ChartPositions(n).Value = 0
180           ChartPositions(n).Date = ""
190       Next
          '
          '170       For n = 1 To g.Cols - 1
          '180           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
          '190           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
          '200       Next
200       StartGridX = 1
210       StopGridX = g.Cols - 1

220       FirstDayFilled = False
230       Counter = 0
240       For x = StartGridX To StopGridX
250           If g.TextMatrix(g.Row, x) <> "" Then
260               If Not FirstDayFilled Then
270                   FirstDayFilled = True
280                   MaxVal = Val(g.TextMatrix(g.Row, x))
                      'ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy")
290                   ChartPositions(Count).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy") & " " & g.TextMatrix(2, x)
300                   ChartPositions(Count).Value = Val(g.TextMatrix(g.Row, x))
310                   Count = Count - 1
320               Else
330                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")))
340                   ChartPositions(Count).Date = g.TextMatrix(1, x) & " " & g.TextMatrix(2, x)
350                   cVal = Val(g.TextMatrix(g.Row, x))
360                   ChartPositions(Count).Value = cVal
370                   If cVal > MaxVal Then MaxVal = cVal
380                   Count = Count - 1
390               End If
400           End If
410       Next

420       PixelsPerDay = (PB.Width - 1060) / NumberOfDays
430       MaxVal = MaxVal * 1.1
440       If MaxVal = 0 Then Exit Sub
450       PixelsPerPointY = PB.Height / MaxVal

460       x = 580 + (NumberOfDays * PixelsPerDay)
470       y = PB.Height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
480       ChartPositions(NumberOfDays).yPos = y
490       ChartPositions(NumberOfDays).xPos = x
          'MsgBox (x & " " & y)

500       PB.ForeColor = vbBlue
510       PB.Circle (x, y), 30
520       PB.Line (x - 15, y - 15)-(x + 15, y + 15), vbBlue, BF
530       PB.PSet (x, y)
         
          



540       For n = NumberOfDays - 1 To 0 Step -1
              '490       For n = NumberOfDays To 0
              'MsgBox (ChartPositions(n).Date & " " & ChartPositions(n).Value)
550           If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
                  '                  Dim uniqueKey As String
                  '560               uniqueKey = ChartPositions(n).Date
                  '570               If Not plottedPoints.Exists(uniqueKey) Then

                  '580                   DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
560               DaysInterval = n
570               x = 580 + (DaysInterval * PixelsPerDay)
                  'MsgBox (DaysInterval)
580               ChartPositions(n).xPos = x
590               y = PB.Height - (ChartPositions(n).Value * PixelsPerPointY)
600               ChartPositions(n).yPos = y
                  'MsgBox (x & " " & y)
610               PB.Line -(x, y)
620               PB.Line (x - 15, y - 15)-(x + 15, y + 15), vbBlue, BF
630               PB.Circle (x, y), 30
640               PB.PSet (x, y)
                  '670                   plottedPoints.Add uniqueKey, True


                  '680               End If
650           End If
660       Next

670       PB.Line (0, PB.Height / 2)-(PB.Width, PB.Height / 2), vbBlack, BF

680       lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
690       lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")

700       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

710       intEL = Erl
720       strES = Err.Description
730       LogError "fFullHaemWE", "DrawChart", intEL, strES


End Sub



Private Sub bcancel_Click()

Unload Me

End Sub

Private Sub FillG()

Dim tb As Recordset
Dim sql As String
Dim x As Integer
Dim xrun As String
Dim xdate As String
Dim Cn As Integer

On Error GoTo FillG_Error

g.Rows = 4
g.AddItem ""
g.RemoveItem 3

LoadInitialDemographics

g.Cols = grdDemog.Rows

For x = 1 To grdDemog.Rows - 1
    g.ColWidth(x) = 1095
    g.col = x
    xrun = grdDemog.TextMatrix(x, 0)
    g.TextMatrix(0, x) = xrun
    xdate = Format$(grdDemog.TextMatrix(x, 3), "dd/mm/yy")
    g.TextMatrix(1, x) = xdate
    If IsDate(grdDemog.TextMatrix(x, 3)) Then
        g.TextMatrix(2, x) = Format(grdDemog.TextMatrix(x, 3), "hh:mm")
    Else
        g.TextMatrix(2, x) = ""
    End If
Next

g.AddItem "WBC"
g.AddItem "RBC"
g.AddItem "Hgb"
g.AddItem "Hct"
g.AddItem "MCV"
g.AddItem "Plt"
g.AddItem "RDW"
g.AddItem "Lymp A"
g.AddItem "Mono A"
g.AddItem "Neut A"
g.AddItem "Eos A"
g.AddItem "Bas A"
g.AddItem "ESR"

For x = 1 To g.Cols - 1
    Cn = Val(grdDemog.TextMatrix(x, 1))
    sql = "Select * from HaemResults where " & _
          "SampleID = '" & g.TextMatrix(0, x) & "'"
    Set tb = New Recordset
    RecOpenServer Cn, tb, sql
    If Not tb.EOF Then
        If Not IsNull(tb!Valid) And tb!Valid Then
            g.TextMatrix(4, x) = tb!WBC & ""
            g.TextMatrix(5, x) = tb!rbc & ""
            g.TextMatrix(6, x) = tb!Hgb & ""
            g.TextMatrix(7, x) = tb!Hct & ""
            g.TextMatrix(8, x) = tb!MCV & ""
            g.TextMatrix(9, x) = tb!plt & ""
            g.TextMatrix(10, x) = tb!RDWCV & ""
            g.TextMatrix(11, x) = tb!LymA & ""
            g.TextMatrix(12, x) = tb!MonoA & ""
            g.TextMatrix(13, x) = tb!NeutA & ""
            g.TextMatrix(14, x) = tb!EosA & ""
            g.TextMatrix(15, x) = tb!BasA & ""
            g.TextMatrix(16, x) = tb!ESR & ""
        Else
            g.TextMatrix(4, x) = "Not"
            g.TextMatrix(5, x) = "Valid"
        End If
    End If
Next

If g.Rows > 4 Then
    g.RemoveItem 3
End If
g.Visible = True

Exit Sub

FillG_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "fFullHaemWE", "FillG", intEL, strES, sql


End Sub
Private Sub LoadInitialDemographics()

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer
      Dim S As String

10    On Error GoTo LoadInitialDemographics_Error

20    With frmViewResultsWE
30        For n = 1 To .grd.Rows - 1
40            If n > 1 Then sql = sql & " Union "
50            sql = sql & "SELECT D.SampleID, D.RunDate, D.SampleDate " & _
                    "FROM Demographics D JOIN HaemResults R " & _
                    "ON D.SampleID = R.SampleID " & _
                    "WHERE "
60            sql = sql & "D.Chart = '" & .grd.TextMatrix(n, 0) & "' "
70            sql = sql & "AND D.PatName = '" & AddTicks(.grd.TextMatrix(n, 2)) & "' "
80            If IsDate(.grd.TextMatrix(n, 1)) Then
90                sql = sql & "AND D.DoB = '" & Format(.grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
100           Else
110               sql = sql & "AND (D.DoB = '' or DoB is Null) "
120           End If
              'sql = sql & " AND SampleID < " & sysOptMicroOffset(0) & " "
130           sql = sql & "AND D.RunDate > '" & Format(Now - Val(frmMain.txtLookBack), "dd/MMM/yyyy") & "' "
140       Next n
150       sql = sql & "ORDER BY D.SampleDate DESC,D.SampleID DESC"
160   End With


      'sql = "SELECT D.SampleID, D.RunDate, D.SampleDate " & _
       '      "FROM Demographics D JOIN HaemResults R " & _
       '      "ON D.SampleID = R.SampleID " & _
       '      "WHERE D.Chart = '" & lblChart & "' " & _
       '      "AND D.PatName = '" & AddTicks(lblName) & "' " & _
       '      "AND D.DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "'"
      'If IsDate(lblDoB) Then
      '    sql = sql & "and DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "'"
      'Else
      '    sql = sql & "and (DoB is null or DoB = '')"
      'End If
170   For n = 0 To intOtherHospitalsInGroup
180       Set tb = New Recordset
190       RecOpenClient n, tb, sql
200       Do While Not tb.EOF
210           S = tb!SampleID & vbTab & _
                  Format$(n) & vbTab & _
                  tb!Rundate & vbTab
220           S = S & Format(tb!sampleDate, "dd/MM/yyyy HH:nn")
230           grdDemog.AddItem S
240           tb.MoveNext
250       Loop
260   Next
270   With grdDemog
280       If .Rows > 2 Then
290           .RemoveItem 1
300           .col = 2
310           .Sort = 9
320       End If
330   End With

340   Exit Sub

LoadInitialDemographics_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "fFullHaemWE", "LoadInitialDemographics", intEL, strES, sql


End Sub


Private Sub cmdGo_Click()

DrawChart

End Sub

Private Sub Form_Activate()

If LogOffNow Then
    Unload Me
End If

FillG
FillCombos

PBar.Max = LogOffDelaySecs
PBar = 0
SingleUserUpdateLoggedOn UserName

Timer1.Enabled = True

End Sub

Private Sub Form_Deactivate()

Timer1.Enabled = False

End Sub


Private Sub Form_Load()

PBar.Max = LogOffDelaySecs
PBar = 0
LogAsViewed "F", "", frmMain.txtChart

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

PBar = 0

End Sub

Private Sub g_Click()
10        On Error GoTo g_Click_Error
20        DrawChart
30        Exit Sub
g_Click_Error:

40        LogError "frmFullHaemWE", "g_Click", Erl, Err.Description

End Sub




Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

PBar = 0

End Sub

Private Sub grdDemog_Compare(ByVal Row1 As Long, ByVal Row2 As Long, cmp As Integer)

Dim d1 As Date
Dim d2 As Date
Dim Column As Integer

With grdDemog
    Column = .col
    cmp = 0
    If IsDate(.TextMatrix(Row1, Column)) Then
        d1 = Format(.TextMatrix(Row1, Column), "dd/mmm/yyyy")
        If IsDate(.TextMatrix(Row2, Column)) Then
            d2 = Format(.TextMatrix(Row2, Column), "dd/mmm/yyyy")
            cmp = Sgn(DateDiff("d", d1, d2))
        End If
    End If
End With

End Sub


Private Sub pb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Integer
Dim CurrentDistance As Long
Dim BestDistance As Long
Dim BestIndex As Integer

On Error GoTo pbmm

PBar = 0

If NumberOfDays = 0 Then Exit Sub

BestIndex = -1
BestDistance = 99999
For i = 0 To NumberOfDays
    CurrentDistance = ((x - ChartPositions(i).xPos) ^ 2 + (y - ChartPositions(i).yPos) ^ 2) ^ (1 / 2)
    If i = 0 Or CurrentDistance < BestDistance Then
        BestDistance = CurrentDistance
        BestIndex = i
    End If
Next

If BestIndex <> -1 Then
    PB.ToolTipText = Format$(ChartPositions(BestIndex).Date, "dd/mmm/yyyy") & " " & ChartPositions(BestIndex).Value
End If

Exit Sub

pbmm:
Exit Sub

End Sub

Private Sub FillCombos()

Dim x As Integer

On Error GoTo FillCombos_Error

cmbPlotFrom.Clear
cmbPlotTo.Clear

For x = 1 To g.Cols - 1
    cmbPlotFrom.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
    cmbPlotTo.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
Next

cmbPlotTo = Format$(g.TextMatrix(1, 1), "dd/mmm/yyyy")

For x = g.Cols - 1 To 1 Step -1
    If DateDiff("d", Format$(g.TextMatrix(1, x), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
        cmbPlotFrom = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
        Exit For
    End If
Next

Exit Sub

FillCombos_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "fFullHaemWE", "FillCombos", intEL, strES


End Sub


Private Sub Timer1_Timer()

'tmrRefresh.Interval set to 1000
PBar = PBar + 1

If PBar = PBar.Max Then
    LogOffNow = True
    Unload Me
End If

End Sub


