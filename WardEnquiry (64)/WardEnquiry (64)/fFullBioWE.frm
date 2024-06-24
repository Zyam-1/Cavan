VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fFullBioWE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Biochemistry History"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   13590
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
   HelpContextID   =   10022
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7215
   ScaleWidth      =   13590
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Records"
      Height          =   885
      Left            =   8460
      TabIndex        =   30
      Top             =   5520
      Width           =   1095
      Begin VB.TextBox txtRecords 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "fFullBioWE.frx":0000
         Top             =   240
         Width           =   510
      End
      Begin ComCtl2.UpDown udRecords 
         Height          =   285
         Left            =   300
         TabIndex        =   32
         Top             =   480
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   327681
         Value           =   25
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRecords"
         BuddyDispid     =   196610
         OrigLeft        =   150
         OrigTop         =   450
         OrigRight       =   915
         OrigBottom      =   690
         Increment       =   20
         Max             =   9999
         Min             =   5
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
   End
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
      Height          =   705
      Left            =   10200
      Picture         =   "fFullBioWE.frx":0005
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5220
      Width           =   1245
   End
   Begin VB.CommandButton bPrint 
      Cancel          =   -1  'True
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   10200
      Picture         =   "fFullBioWE.frx":030F
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4500
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
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
      Height          =   705
      Left            =   10200
      Picture         =   "fFullBioWE.frx":0979
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6150
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
      Left            =   60
      TabIndex        =   9
      Top             =   4410
      Width           =   2040
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
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   555
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
         Left            =   180
         TabIndex        =   11
         Text            =   "cmbPlotFrom"
         Top             =   540
         Width           =   1455
      End
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
         Left            =   180
         TabIndex        =   10
         Text            =   "cmbPlotTo"
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Left            =   210
         TabIndex        =   21
         Top             =   300
         Width           =   345
      End
      Begin VB.Label T 
         AutoSize        =   -1  'True
         Caption         =   "to"
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
         Left            =   210
         TabIndex        =   20
         Top             =   900
         Width           =   135
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10260
      Top             =   6810
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   2325
      Left            =   1860
      ScaleHeight     =   2265
      ScaleWidth      =   5505
      TabIndex        =   7
      Top             =   4500
      Width           =   5565
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   315
         Left            =   2640
         TabIndex        =   19
         Top             =   0
         Width           =   105
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3735
      Left            =   30
      TabIndex        =   6
      ToolTipText     =   "Click on Parameter to show Graph"
      Top             =   660
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   4
      Cols            =   4
      FixedRows       =   3
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FormatString    =   "<Code   |<TestName   |<Normal Range |"
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   30
      TabIndex        =   8
      Top             =   6900
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   3600
      TabIndex        =   23
      Top             =   1140
      Width           =   2985
   End
   Begin MSFlexGridLib.MSFlexGrid grdDemog 
      Height          =   4935
      Left            =   11490
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "<SampleID   |<Cnxn    |<RunDate    |<TimeTaken                  "
   End
   Begin VB.CheckBox chkIgnorePOCT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ignore POCT results"
      Height          =   240
      Left            =   9060
      TabIndex        =   29
      Top             =   420
      Width           =   2115
   End
   Begin VB.Image imgSideArrow 
      Height          =   240
      Left            =   11160
      Picture         =   "fFullBioWE.frx":0FE3
      Top             =   420
      Width           =   60
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   7500
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
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
      Left            =   10080
      TabIndex        =   26
      Top             =   60
      Width           =   270
   End
   Begin VB.Label lblSex 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10380
      TabIndex        =   25
      Top             =   30
      Width           =   1035
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   5340
      TabIndex        =   22
      Top             =   2400
      Width           =   315
   End
   Begin VB.Label lblNopas 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   17
      Top             =   5190
      Width           =   1545
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
      Left            =   7500
      TabIndex        =   15
      Top             =   4500
      Width           =   480
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
      Left            =   7500
      TabIndex        =   14
      Top             =   5370
      Width           =   480
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
      Left            =   7500
      TabIndex        =   13
      Top             =   6570
      Width           =   480
   End
   Begin VB.Label Label1 
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
      Left            =   4920
      TabIndex        =   5
      Top             =   90
      Width           =   420
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
      Left            =   2460
      TabIndex        =   4
      Top             =   90
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   30
      Width           =   3765
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2850
      TabIndex        =   2
      Top             =   30
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
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
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   30
      Width           =   1545
   End
End
Attribute VB_Name = "fFullBioWE"
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
Dim LongName(0 To 1000) As String



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
          
          Dim Count As Integer

20        On Error GoTo DrawChart_Error

30        MaxVal = 0
40        lblMaxVal = ""
50        lblMeanVal = ""
60        lblTest = g.TextMatrix(g.Row, 1)

70        PB.Cls
80        PB.Picture = LoadPicture("")

90        NumberOfDays = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
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
    '200       For n = 2 To g.Cols - 1
    '210           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then
    '220               StartGridX = n
    '230               Exit For
    '240           End If
    '250       Next
    '260       For n = 2 To g.Cols - 1
    '270           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then
    '280               StopGridX = n
    '290               Exit For
    '300           End If
    '310       Next
200       StartGridX = 2
210       StopGridX = g.Cols - 1

320       FirstDayFilled = False
330       Counter = 0
340       For x = StartGridX To StopGridX
350           If g.TextMatrix(g.Row, x) <> "" Then
360               If Not FirstDayFilled Then
370                   FirstDayFilled = True
380                   MaxVal = Val(g.TextMatrix(g.Row, x))
                      'ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy")
390                   ChartPositions(Count).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy") & " " & g.TextMatrix(2, x)
400                   ChartPositions(Count).Value = Val(g.TextMatrix(g.Row, x))
410                   Count = Count - 1
420               Else
430                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")))
440                   ChartPositions(Count).Date = g.TextMatrix(1, x) & " " & g.TextMatrix(2, x)
450                   cVal = Val(g.TextMatrix(g.Row, x))
460                   ChartPositions(Count).Value = cVal
470                   If cVal > MaxVal Then MaxVal = cVal
480                   Count = Count - 1
490               End If
500           End If

510       Next

520       PixelsPerDay = (PB.Width - 1060) / NumberOfDays
530       MaxVal = MaxVal * 1.1
540       If MaxVal = 0 Then Exit Sub
550       PixelsPerPointY = PB.Height / MaxVal


560       x = 580 + (NumberOfDays * PixelsPerDay)
570       y = PB.Height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
580       ChartPositions(NumberOfDays).yPos = y
590       ChartPositions(NumberOfDays).xPos = x

600       PB.ForeColor = vbBlue
610       PB.Circle (x, y), 30
620       PB.Line (x - 15, y - 15)-(x + 15, y + 15), vbBlue, BF
630       PB.PSet (x, y)
         
          


640       For n = NumberOfDays - 1 To 0 Step -1
              '490       For n = NumberOfDays To 0
              'MsgBox (ChartPositions(n).Date & " " & ChartPositions(n).Value)
650           If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
                  Dim uniqueKey As String
660               uniqueKey = ChartPositions(n).Date
670               If Not plottedPoints.Exists(uniqueKey) Then

                      'DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
680                   DaysInterval = n
                      'MsgBox (DaysInterval)
690                   x = 580 + (DaysInterval * PixelsPerDay)

700                   ChartPositions(n).xPos = x
710                   y = PB.Height - (ChartPositions(n).Value * PixelsPerPointY)
720                   ChartPositions(n).yPos = y
730                   PB.Line -(x, y)
                      'MsgBox (x & " " & y)

740                   PB.Line (x - 15, y - 15)-(x + 15, y + 15), vbBlue, BF
750                   PB.Circle (x, y), 30
760                   PB.PSet (x, y)
770                   plottedPoints.Add uniqueKey, True


780               End If
790           End If
800       Next

810       PB.Line (0, PB.Height / 2)-(PB.Width, PB.Height / 2), vbBlack, BF

820       lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
830       lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")

840       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

850       intEL = Erl
860       strES = Err.Description
870       LogError "fFullBioWE", "DrawChart", intEL, strES


End Sub

Private Sub FillCombos()

      Dim x As Integer

10    On Error GoTo FillCombos_Error

20    cmbPlotFrom.Clear
30    cmbPlotTo.Clear

40    For x = 3 To g.Cols - 1
50        cmbPlotFrom.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
60        cmbPlotTo.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
70    Next

80    If g.Cols > 3 Then
90        cmbPlotTo = Format$(g.TextMatrix(1, 3), "dd/mmm/yyyy")

100       For x = g.Cols - 1 To 3 Step -1
110           If Trim$(g.TextMatrix(1, x)) <> "" Then
120               If DateDiff("d", Format$(g.TextMatrix(1, x), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
130                   cmbPlotFrom = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
140                   Exit For
150               End If
160           End If
170       Next
180   End If

190   Exit Sub

FillCombos_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "fFullBioWE", "FillCombos", intEL, strES


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
50            sql = sql & "SELECT DISTINCT TOP " & Format$(Val(txtRecords)) & " D.SampleID, D.SampleDate, D.TimeTaken FROM Demographics D JOIN BioResults R ON D.SampleID = R.SampleID WHERE "
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


      'sql = "SELECT DISTINCT D.SampleID, D.SampleDate, D.TimeTaken FROM Demographics D JOIN BioResults R ON D.SampleID = R.SampleID WHERE " & _
       '"D.Chart = '" & lblChart & "' " & _
       '"AND PatName = '" & AddTicks(lblName) & "' "
      'If IsDate(lblDoB) Then
      '    sql = sql & "and DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
      'Else
      '    sql = sql & "and (DoB is null or DoB = '') "
      'End If
      'sql = sql & "ORDER BY D.SampleDate DESC,D.SampleID DESC"
170   grdDemog.Clear
180   grdDemog.Rows = 2
190   grdDemog.Cols = 4
200   grdDemog.FormatString = "<SampleID   |<Cnxn    |<RunDate    |<TimeTaken                  "

210   For n = 0 To intOtherHospitalsInGroup
220       Set tb = New Recordset
230       sql = "SELECT TOP " & Format$(Val(txtRecords)) & " * FROM (" & sql & ") AS CR"
240       RecOpenClient n, tb, sql
250       Do While Not tb.EOF
260           S = tb!SampleID & vbTab & _
                  Format$(n) & vbTab & _
                  Format$(tb!sampleDate, "dd/MM/yyyy") & vbTab
270           If Format$(tb!sampleDate, "HH:nn") <> "00:00" Then
280               S = S & Format(tb!sampleDate, "dd/mmm/yyyy hh:mm")
290           End If
300           If chkIgnorePOCT.Value = 0 Then
310               grdDemog.AddItem S
320           Else
330               If tb!SampleID < 3000000 Or tb!SampleID > 4000000 Then
                      'only add samples if it's not POCT autogenerated sampleid. 1
                      '1000 to 99999 are auto generated sample ids for POCT
340                   grdDemog.AddItem S
350               End If
360           End If
370           tb.MoveNext
380       Loop
390   Next
400   With grdDemog
410       If .Rows > 2 Then
420           .RemoveItem 1
430           .col = 2
440           .Sort = 9
450       End If
460   End With

470   Exit Sub

LoadInitialDemographics_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "fFullBioWE", "LoadInitialDemographics", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

10    Unload Me

End Sub

Private Sub FillG()

Dim snr As Recordset
Dim sql As String
Dim x As Integer
Dim xrun As String
Dim xdate As String
Dim Cn As Integer
Dim DaysOld As Long
Dim SelectNormalRange As String
Dim Found As Boolean
Dim n As Integer
Dim BRs As New BIEResults
Dim Br As BIEResult
Dim MaskFlag As String
Dim l_Result As String
'+++ Junaid 20-02-2024
Dim l_rst As Recordset
Dim l_SQL As String
'--- Junaid

On Error GoTo FillG_Error


'g.BackColorFixed = vbBlue
'g.ForeColorFixed = vbWhite
g.Clear
g.Rows = 4
g.Cols = 4
g.FormatString = "<Code   |<TestName       |<Norm Range|"
LoadInitialDemographics


g.Visible = False
g.Cols = grdDemog.Rows + 2
g.ColWidth(0) = 0
g.ColWidth(1) = 1700
g.ColWidth(2) = 0 '1150

Select Case Left$(UCase$(Trim$(lblSex)), 1)
Case "M": SelectNormalRange = " MaleLow as Low, MaleHigh as High, "
Case "F": SelectNormalRange = " FemaleLow as Low, FemaleHigh as High, "
Case Else: SelectNormalRange = " FemaleLow as Low, MaleHigh as High, "
End Select
If IsDate(lblDoB) Then
    DaysOld = DateDiff("d", lblDoB, Now)
End If

'SampleID and sampledate across
For x = 3 To g.Cols - 1
    g.ColWidth(x) = 1000
    g.col = x
    xrun = grdDemog.TextMatrix(x - 2, 0)
    g.TextMatrix(0, x) = xrun
    xdate = Format$(grdDemog.TextMatrix(x - 2, 2), "dd/mm/yy")
    g.TextMatrix(1, x) = xdate
    If IsDate(grdDemog.TextMatrix(x - 2, 3)) Then
        g.TextMatrix(2, x) = Format(grdDemog.TextMatrix(x - 2, 3), "hh:mm")
    Else
        g.TextMatrix(2, x) = ""
    End If
    Cn = Val(grdDemog.TextMatrix(x - 2, 1))
    'fill list with test names
            
    sql = "Select  " & SelectNormalRange & "  BR.*, TD.PrintPriority, TD.ShortName,TD.LongName, TD.Code " & _
          "from BioResults as BR, BioTestDefinitions as TD " & _
          "where SampleID = '" & xrun & "' " & _
          "and TD.Code = BR.Code " & _
          "and TD.Hospital = '" & HospName(Cn) & "' " & _
          "and TD.AgeFromDays <= " & DaysOld & " " & _
          "and TD.AgeToDays >= " & DaysOld & " "
          If chkIgnorePOCT.Value = vbChecked Then
            sql = sql & "And TD.Analyser <> 'POCT'" & " "
          End If
          sql = sql & "order by TD.PrintPriority"
    Set snr = New Recordset
    RecOpenClient Cn, snr, sql
    Do While Not snr.EOF

        If snr!ShortName <> "H" And snr!ShortName <> "I" And snr!ShortName <> "L" Then

            Found = False
            For n = 3 To g.Rows - 1
                If g.TextMatrix(n, 0) = Trim$(snr!Code) Then
                    Found = True
                    Exit For
                End If
            Next
            If Not Found Then
                g.AddItem Trim$(snr!Code) & vbTab & Trim$(snr!ShortName) '& vbTab & snr!Low & "-" & snr!High & ""
                LongName(n - 3) = snr!LongName
                If Cn <> 0 Then
                    g.Row = g.Rows - 1
                    g.col = 1
                    g.CellBackColor = vbGreen
                End If
            End If
        End If
        snr.MoveNext
    Loop
Next

'fill in results
For x = 3 To g.Cols - 1
    xdate = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
    xrun = g.TextMatrix(0, x)
    Cn = grdDemog.TextMatrix(x - 2, 1)
    Set BRs = BRs.Load("Bio", xrun, "Results", gDONTCARE, gDONTCARE, "", Cn)

    For Each Br In BRs
        g.Row = GetRow(Br.Code)
        g.ColAlignment(x) = flexAlignRightCenter
        If g.Row > 2 Then
            
            MaskFlag = MaskInhibit(Br, BRs)
            If MaskFlag <> "" Then
                g.TextMatrix(g.Row, x) = "*****"
            ElseIf Br.Valid Then
                If Br.Code = "19" And Val(Br.Result) < 5 Then
                    g.TextMatrix(g.Row, x) = "<5"
                Else
                    Select Case Br.Printformat
                    Case 0: g.TextMatrix(g.Row, x) = Format$(Br.Result, "0")
                    Case 1: g.TextMatrix(g.Row, x) = Format$(Br.Result, "0.0")
                    Case 2: g.TextMatrix(g.Row, x) = Format$(Br.Result, "0.00")
                    Case 3: g.TextMatrix(g.Row, x) = Format$(Br.Result, "0.000")
                    
                    
                    End Select
                    If Val(Br.Result) <> 0 Then
                        If Val(Br.Result) < Br.Low Then
                            g.col = x
                            g.CellBackColor = sysOptLowBack(0)
                            g.CellForeColor = sysOptLowFore(0)
                            g.CellFontBold = True
                        ElseIf Val(Br.Result) > Br.High Then
                            g.col = x
                            g.CellBackColor = sysOptHighBack(0)
                            g.CellForeColor = sysOptHighFore(0)
                            g.CellFontBold = True
                        End If
                    End If
                End If
            Else
                g = "NV"
            End If
'            If (Br.SampleID >= 3000000) And (Br.SampleID <= 4000000) Then
'            If Br.Anyl = "POCT" Then
            l_SQL = "Select Analyser from BioTestDefinitions Where Code = '" & Br.Code & "'"
            Set l_rst = New Recordset
            RecOpenClient Cn, l_rst, l_SQL
            If Not l_rst Is Nothing Then
                If Not l_rst.EOF Then
                    If ConvertNull(l_rst!Analyser, "") = "POCT" Then
                        g.col = x
                        Set g.CellPicture = imgSideArrow.Picture
                    End If
                End If
            End If
            
                'they are auto generated POCT results so show a marker with them
'                g.col = x
'                Set g.CellPicture = imgSideArrow.Picture

'            End If

        End If
    Next
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
LogError "fFullBioWE", "FillG", intEL, strES, sql

End Sub

Private Sub bprint_Click()

      Dim x As Long
      Dim n As Long
      Dim z As Long
      Dim lngStart As Long
      Dim lngStop As Long
      Dim AllDone As Boolean
      Dim Px As Long

10    x = g.Cols

20    Printer.Orientation = vbPRORLandscape

30    Printer.Font.Size = 16
40    Printer.Print Tab(15); "Cumulative Report from Biochemistry Dept."
50    Printer.Print

60    Printer.Font.Size = 14
70    Printer.Print Tab(10); "Name : " & lblName;
80    Printer.Print Tab(40); "Dob  : " & lblDoB
90    Printer.Print
100   Printer.Print

110   Printer.Font.Size = 10

120   AllDone = False
130   lngStart = 3
140   lngStop = 9
150   If g.Cols < lngStop - 1 Then
160       lngStop = g.Cols - 1
170   End If
180   Do While Not AllDone
190       Printer.Print
200       For n = 0 To g.Rows - 1
210           For z = 1 To 2
220               Printer.Print Tab(12 * (z - 1)); g.TextMatrix(n, z);
230           Next
240           Px = 24
250           For z = lngStart To lngStop
260               Printer.Print Tab(Px); g.TextMatrix(n, z);
270               Px = Px + 12
280           Next
290           Printer.Print
300       Next
310       If lngStop >= g.Cols - 1 Then
320           AllDone = True
330       Else
340           lngStart = lngStop + 1
350           lngStop = lngStart + 6
360           If g.Cols < lngStop - 1 Then
370               lngStop = g.Cols - 1
380           End If
390       End If
400   Loop

410   Printer.Print Tab(30); "----End of Report----"

420   Printer.EndDoc

End Sub

Private Sub chkIgnorePOCT_Click()

10    On Error GoTo chkIgnorePOCT_Click_Error

20    FillG
30    FillCombos

40    Exit Sub

chkIgnorePOCT_Click_Error:

       Dim strES As String
       Dim intEL As Integer

50     intEL = Erl
60     strES = Err.Description
70     LogError "fFullBioWE", "chkIgnorePOCT_Click", intEL, strES
          
End Sub

Private Sub cmbPlotFrom_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub cmbPlotTo_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub cmdGo_Click()

10    DrawChart

End Sub

Private Sub cmdXL_Click()

10    ExportFlexGrid g, Me

End Sub

Private Sub Form_Activate()
'11595
'frame1 2040 LF 60
'Grid 7305 LF 1860
'
10    FillG
20    FillCombos

30    PBar.Max = LogOffDelaySecs
40    PBar = 0
50    SingleUserUpdateLoggedOn UserName

60    Timer1.Enabled = True

70    If sysOptDisableWardPrinting(0) Then
80        bPrint.Enabled = False
90    End If

End Sub

Private Function GetRow(ByVal TestCode As String) As Integer

      Dim n As Integer

10    On Error GoTo GetRow_Error

20    For n = 3 To g.Rows - 1
30        If UCase(Trim(TestCode)) = UCase(Trim(g.TextMatrix(n, 0))) Then
40            GetRow = n
50            Exit For
60        End If
70    Next

80    Exit Function

GetRow_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "fFullBioWE", "GetRow", intEL, strES


End Function

Private Sub Form_Deactivate()

10    Timer1.Enabled = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub

Private Sub g_Click()
10        On Error GoTo g_Click_Error
20        DrawChart
30        Exit Sub
g_Click_Error:

40        LogError "fFullBioWE", "g_Click", Erl, Err.Description

End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0
20    If g.MouseRow > 2 Then
30        If g.MouseCol = 1 Then
40            g.ToolTipText = LongName(g.MouseRow - 2)
50        Else
60            g.ToolTipText = g.TextMatrix(g.MouseRow, g.MouseCol)
70        End If
80    End If


      '20    If gBio.MouseCol = 1 And gBio.MouseRow > 0 Then
      '30        gBio.ToolTipText = gBio.TextMatrix(gBio.MouseRow, gBio.MouseCol)
End Sub


Private Sub grdDemog_Compare(ByVal Row1 As Long, ByVal Row2 As Long, cmp As Integer)

      Dim d1 As Date
      Dim d2 As Date
      Dim Column As Integer

10    With grdDemog
20        Column = .col
30        cmp = 0
40        If IsDate(.TextMatrix(Row1, Column)) Then
50            d1 = Format(.TextMatrix(Row1, Column), "dd/mmm/yyyy")
60            If IsDate(.TextMatrix(Row2, Column)) Then
70                d2 = Format(.TextMatrix(Row2, Column), "dd/mmm/yyyy")
80                cmp = Sgn(DateDiff("d", d1, d2))
90            End If
100       End If
110   End With

End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub pb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

      Dim i As Integer
      Dim CurrentDistance As Long
      Dim BestDistance As Long
      Dim BestIndex As Integer

10    On Error GoTo pbmm

20    PBar = 0

30    If NumberOfDays = 0 Then Exit Sub

40    BestIndex = -1
50    BestDistance = 99999
60    For i = 0 To NumberOfDays
70        CurrentDistance = ((x - ChartPositions(i).xPos) ^ 2 + (y - ChartPositions(i).yPos) ^ 2) ^ (1 / 2)
80        If i = 0 Or CurrentDistance < BestDistance Then
90            BestDistance = CurrentDistance
100           BestIndex = i
110       End If
120   Next

130   If BestIndex <> -1 Then
140       PB.ToolTipText = Format$(ChartPositions(BestIndex).Date, "dd/mmm/yyyy") & " " & ChartPositions(BestIndex).Value
150   End If

160   Exit Sub

pbmm:
170   Exit Sub

End Sub


Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10    PBar = PBar + 1

20    If PBar = PBar.Max Then
30        Unload Me
40    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtRecords_Change
' Author    : Masood
' Date      : 05/Aug/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtRecords_Change()
10      On Error GoTo txtRecords_Change_Error

20    If Val(txtRecords) <> 0 Then
30        FillG
40    End If
       
50    Exit Sub

       
txtRecords_Change_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "fFullBioWE", "txtRecords_Change", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : udRecords_MouseUp
' Author    : Masood
' Date      : 11/Aug/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub udRecords_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10        On Error GoTo udRecords_MouseUp_Error



20        txtRecords = udRecords.Value

30        Exit Sub


udRecords_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "fFullBioWE", "udRecords_MouseUp", intEL, strES
End Sub
