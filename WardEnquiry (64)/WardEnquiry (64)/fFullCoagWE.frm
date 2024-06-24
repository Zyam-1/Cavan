VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fFullCoagWE 
   Caption         =   "NetAcquire - Full Coagulation History"
   ClientHeight    =   7005
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11610
   HelpContextID   =   10031
   Icon            =   "fFullCoagWE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11610
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10320
      TabIndex        =   24
      Top             =   5640
      Width           =   975
      Begin VB.TextBox txtRecords 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "fFullCoagWE.frx":0ECA
         Top             =   240
         Width           =   510
      End
      Begin ComCtl2.UpDown udRecords 
         Height          =   285
         Left            =   180
         TabIndex        =   26
         Top             =   540
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   327681
         Value           =   25
         AutoBuddy       =   -1  'True
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
      Height          =   615
      Left            =   10260
      Picture         =   "fFullCoagWE.frx":0ECF
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4290
      Width           =   1245
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   2325
      Left            =   1800
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   13
      Top             =   4440
      Width           =   4185
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1950
         TabIndex        =   14
         Top             =   0
         Width           =   75
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10320
      Top             =   4980
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plot between"
      Height          =   2325
      Left            =   60
      TabIndex        =   9
      Top             =   4320
      Width           =   1785
      Begin VB.ComboBox cmbPlotTo 
         Height          =   315
         Left            =   150
         TabIndex        =   12
         Text            =   "cmbPlotTo"
         Top             =   870
         Width           =   1455
      End
      Begin VB.ComboBox cmbPlotFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Text            =   "cmbPlotFrom"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H80000016&
         Caption         =   "Go"
         Height          =   315
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   555
      End
   End
   Begin VB.CommandButton bcancel 
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   6930
      Picture         =   "fFullCoagWE.frx":11D9
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5910
      Width           =   1245
   End
   Begin VB.CommandButton bPrint 
      Cancel          =   -1  'True
      Caption         =   "Print"
      Height          =   825
      Left            =   6900
      Picture         =   "fFullCoagWE.frx":1843
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   6376
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
      FormatString    =   "<Code   |<TestName  |<Normal Range  |          "
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   30
      TabIndex        =   15
      Top             =   6720
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdDemog 
      Height          =   5445
      Left            =   11640
      TabIndex        =   21
      Top             =   990
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   9604
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
      Left            =   9150
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   2940
      TabIndex        =   20
      Top             =   1950
      Width           =   315
   End
   Begin VB.Label lblMaxVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   6390
      Width           =   480
   End
   Begin VB.Label lblMeanVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   5310
      Width           =   480
   End
   Begin VB.Label lblNopas 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3300
      TabIndex        =   16
      Top             =   5100
      Width           =   1545
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   6
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   150
      Width           =   375
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5790
      TabIndex        =   3
      Top             =   120
      Width           =   3765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   2850
      TabIndex        =   2
      Top             =   180
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   5310
      TabIndex        =   1
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "fFullCoagWE"
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

10    ExportFlexGrid g, Me

End Sub

Private Sub bcancel_Click()

10    Unload Me

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
40    Printer.Print Tab(15); "Cumulative Report from Coagulation Dept."
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

Private Sub cmdGo_Click()

10    DrawChart

End Sub

Private Sub Form_Activate()

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

Private Sub Form_Deactivate()

10    Timer1.Enabled = False

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

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
          '          Dim plottedPoints As Object
          '10        Set plottedPoints = CreateObject("Scripting.Dictionary")
          
          Dim Count As Integer

10        On Error GoTo DrawChart_Error

20        MaxVal = 0
30        lblMaxVal = ""
40        lblMeanVal = ""
50        lblTest = g.TextMatrix(g.Row, 1)

60        pb.Cls
70        pb.Picture = LoadPicture("")

80        NumberOfDays = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
90        NumberOfDays = g.Cols - 2
100       Count = NumberOfDays
110       If NumberOfDays < 2 Then Exit Sub

120       ReDim ChartPositions(0 To NumberOfDays)

130       For n = 1 To NumberOfDays
140           ChartPositions(n).xPos = 0
150           ChartPositions(n).yPos = 0
160           ChartPositions(n).Value = 0
170           ChartPositions(n).Date = ""
180       Next
          '
'190       For n = 1 To g.Cols - 1
'200           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
'210           Exit For
'220       Next
'230       For n = 1 To g.Cols - 1
'240           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
'250           Exit For
'260       Next
270       StartGridX = 2
280       StopGridX = g.Cols - 1

290       FirstDayFilled = False
300       Counter = 0
310       For x = StartGridX To StopGridX
320           If g.TextMatrix(g.Row, x) <> "" Then
330               If Not FirstDayFilled Then
340                   FirstDayFilled = True
350                   MaxVal = Val(g.TextMatrix(g.Row, x))
                      'ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy")
360                   ChartPositions(Count).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy") & " " & g.TextMatrix(2, x)
370                   ChartPositions(Count).Value = Val(g.TextMatrix(g.Row, x))
380                   Count = Count - 1
390               Else
400                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")))
410                   ChartPositions(Count).Date = g.TextMatrix(1, x) & " " & g.TextMatrix(2, x)
420                   cVal = Val(g.TextMatrix(g.Row, x))
430                   ChartPositions(Count).Value = cVal
440                   If cVal > MaxVal Then MaxVal = cVal
450                   Count = Count - 1
460               End If
470           End If
480       Next

490       PixelsPerDay = (pb.Width - 1060) / NumberOfDays
500       MaxVal = MaxVal * 1.1
510       If MaxVal = 0 Then Exit Sub
520       PixelsPerPointY = pb.Height / MaxVal

530       x = 580 + (NumberOfDays * PixelsPerDay)
540       y = pb.Height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
550       ChartPositions(NumberOfDays).yPos = y
560       ChartPositions(NumberOfDays).xPos = x

570       pb.ForeColor = vbBlue
580       pb.Circle (x, y), 30
590       pb.Line (x - 15, y - 15)-(x + 15, y + 15), vbBlue, BF
600       pb.PSet (x, y)
         
          


610       For n = NumberOfDays - 1 To 0 Step -1
              '490       For n = NumberOfDays To 0
              'MsgBox (ChartPositions(n).Date & " " & ChartPositions(n).Value)
620           If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
                  '                  Dim uniqueKey As String
                  '550               uniqueKey = ChartPositions(n).Date
                  '560               If Not plottedPoints.Exists(uniqueKey) Then

'630               DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
                  DaysInterval = n
640               x = 580 + (DaysInterval * PixelsPerDay)
650               ChartPositions(n).xPos = x
660               y = pb.Height - (ChartPositions(n).Value * PixelsPerPointY)
670               ChartPositions(n).yPos = y
680               pb.Line -(x, y)
690               pb.Line (x - 15, y - 15)-(x + 15, y + 15), vbBlue, BF
700               pb.Circle (x, y), 30
710               pb.PSet (x, y)
                  '630               plottedPoints.Add uniqueKey, True


                  '670               End If
720           End If
730       Next

740       pb.Line (0, pb.Height / 2)-(pb.Width, pb.Height / 2), vbBlack, BF

750       lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
760       lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")

770       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

780       intEL = Erl
790       strES = Err.Description
800       LogError "fFullCoagWE", "DrawChart", intEL, strES


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

80    cmbPlotTo = Format$(g.TextMatrix(1, 3), "dd/mmm/yyyy")
90    If Not IsDate(cmbPlotTo) Then Exit Sub

100   For x = g.Cols - 1 To 3 Step -1
110       If DateDiff("d", Format$(g.TextMatrix(1, x), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
120           cmbPlotFrom = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
130           Exit For
140       End If
150   Next

160   Exit Sub

FillCombos_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fFullCoagWE", "FillCombos", intEL, strES


End Sub


Private Sub FillG()

      Dim snr As Recordset
      Dim sql As String
      Dim x As Integer
      Dim xrun As String
      Dim xdate As String
      Dim Flag As String
      Dim Code As String
      Dim Cn As Integer
      Dim n As Integer
      Dim Found As Boolean
      Dim CR As CoagResult
      Dim CRs As CoagResults

10    On Error GoTo FillG_Error

20    LoadInitialDemographics

30    g.Visible = False
        g.Clear
        g.Rows = 4
        g.Cols = 4
40    g.Cols = grdDemog.Rows + 2
50    g.ColWidth(0) = 0
60    g.ColWidth(1) = 1095
70    g.ColWidth(2) = 0
'      g.Rows = 0
80    For x = 1 To grdDemog.Rows - 1
90        g.ColWidth(x + 2) = 1095
100       g.col = x
110       xrun = grdDemog.TextMatrix(x, 0)
120       g.TextMatrix(0, x + 2) = xrun
130       xdate = Format$(grdDemog.TextMatrix(x, 2), "dd/mm/yy")
140       g.TextMatrix(1, x + 2) = xdate
150       If IsDate(grdDemog.TextMatrix(x, 2)) Then
160           g.TextMatrix(2, x + 2) = Format(grdDemog.TextMatrix(x, 2), "hh:mm")
170       Else
180           g.TextMatrix(2, x + 2) = ""
190       End If
200       Cn = grdDemog.TextMatrix(x, 1)

          'fill list with test names
210       sql = "Select  R.*, PrintPriority, dp, TestName, D.Units, " & _
                "D.MaleLow as low, D.MaleHigh as High " & _
                "from CoagResults as R, CoagTestDefinitions as D " & _
                "where SampleID = '" & xrun & "' " & _
                "and D.Code = R.Code " & _
                "and D.Hospital = '" & HospName(Cn) & "' " & _
                "order by PrintPriority"
220       Set snr = New Recordset
230       RecOpenClient Cn, snr, sql
240       Do While Not snr.EOF
250           If Trim$(snr!Units) = "INR" Then
260               Code = "INR"
270           Else
280               Code = Trim$(snr!Code)
290           End If
300           Found = False
310           For n = 3 To g.Rows - 1
320               If g.TextMatrix(n, 0) = Code Then
330                   Found = True
340                   Exit For
350               End If
360           Next
370           If Not Found Then
380               If UCase(snr!TestName) = "INR" Then
390                   g.AddItem Code & vbTab & snr!TestName & vbTab & ""
400               Else
410                   g.AddItem Code & vbTab & snr!TestName '& vbTab & snr!Low & "-" & snr!High & ""
420               End If
430           End If
440           snr.MoveNext
450       Loop
460   Next

      'fill in results
'      g.Clear
470   For x = 3 To g.Cols - 1
480       g.col = x
'490       g.Row = 1
500       xdate = Format$(g, "dd/mmm/yyyy")
510       g.Row = 0
520       xrun = g

530       Set CRs = New CoagResults
540       Set CRs = CRs.Load(xrun, gDONTCARE, gDONTCARE, "Results", Cn)
550       If Not CRs Is Nothing Then
560           For Each CR In CRs
570               If Trim$(CR.Units) = "INR" Then
580                   Code = "INR"
590               Else
600                   Code = Trim$(CR.Code)
610               End If
620               g.Row = GetRow(Code)
630               If g.Row <> 0 Then
640                   If CR.Valid Then
650                       Select Case CR.DP
                          Case 0: g = Format$(CR.Result, "0")
660                       Case 1: g = Format$(CR.Result, "0.0")
670                       Case 2: g = Format$(CR.Result, "0.00")
680                       End Select
690                       If Val(CR.Result) > Val(CR.High) Then
700                           g.CellBackColor = sysOptHighBack(0)
710                           g.CellForeColor = sysOptHighFore(0)
720                           g.CellFontBold = True
730                       ElseIf Val(CR.Result) < Val(CR.Low) Then
740                           g.CellBackColor = sysOptLowBack(0)
750                           g.CellForeColor = sysOptLowFore(0)
760                           g.CellFontBold = True
770                       Else
780                           g.CellBackColor = vbWhite
790                           g.CellFontBold = False
800                       End If
810                       Flag = ""
820                   Else
830                       g = "NV"
840                   End If
850               End If
860           Next
870       End If
880   Next

890   g.Visible = True
900   If g.Rows > 4 Then
910       g.RemoveItem 3
920   End If

930   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

940   intEL = Erl
950   strES = Err.Description
960   LogError "fFullCoagWE", "FillG", intEL, strES, sql

End Sub
Private Sub LoadInitialDemographics()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim S As String

10        On Error GoTo LoadInitialDemographics_Error
20        grdDemog.Clear
30        grdDemog.FormatString = "<SampleID   |<Cnxn    |<RunDate    |<TimeTaken                  "
40        grdDemog.Rows = 2
50        With frmViewResultsWE
60            For n = 1 To .grd.Rows - 1
70                If n > 1 Then sql = sql & " Union "
80                sql = sql & "SELECT DISTINCT TOP " & Format$(Val(txtRecords)) & "  D.SampleID, D.RunDate, D.TimeTaken, D.SampleDate " & _
                        "FROM Demographics D JOIN CoagResults R " & _
                        "ON D.SampleID = R.SampleID " & _
                        "WHERE "
90                sql = sql & "D.Chart = '" & .grd.TextMatrix(n, 0) & "' "
100               sql = sql & "AND D.PatName = '" & AddTicks(.grd.TextMatrix(n, 2)) & "' "
110               If IsDate(.grd.TextMatrix(n, 1)) Then
120                   sql = sql & "AND D.DoB = '" & Format(.grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
130               Else
140                   sql = sql & "AND (D.DoB = '' or DoB is Null) "
150               End If
                  'sql = sql & " AND SampleID < " & sysOptMicroOffset(0) & " "
160               sql = sql & "AND D.RunDate > '" & Format(Now - Val(frmMain.txtLookBack), "dd/MMM/yyyy") & "' "
170           Next n
180           sql = sql & "ORDER BY D.SampleDate DESC,D.SampleID DESC"
190       End With


          'sql = "SELECT DISTINCT D.SampleID, D.RunDate, D.TimeTaken " & _
           '      "FROM Demographics D JOIN CoagResults R " & _
           '      "ON D.SampleID = R.SampleID " & _
           '      "WHERE D.Chart = '" & lblChart & "' " & _
           '      "AND D.PatName = '" & AddTicks(lblName) & "' " & _
           '      "AND D.DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "'"
200       For n = 0 To intOtherHospitalsInGroup
210           Set tb = New Recordset
220           sql = "SELECT TOP " & Format$(Val(txtRecords)) & " * FROM (" & sql & ") AS CR"
230           RecOpenClient n, tb, sql
              
240           Do While Not tb.EOF
250               S = tb!SampleID & vbTab & _
                      Format$(n) & vbTab & _
                      tb!sampleDate & vbTab
260               If Not IsNull(tb!TimeTaken) Then
270                   S = S & Format(tb!TimeTaken, "dd/mmm/yyyy hh:mm")
280               End If
290               grdDemog.AddItem S
300               tb.MoveNext
310           Loop
320       Next
330       With grdDemog
340           If .Rows > 2 Then
350               .RemoveItem 1
360               .col = 2
370               .Sort = 9
380           End If
390       End With

400       Exit Sub

LoadInitialDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "fFullCoagWE", "LoadInitialDemographics", intEL, strES, sql

End Sub

Private Function GetRow(ByVal TestCode As String) As Integer

      Dim n As Integer

10    For n = 3 To g.Rows - 1
20        If TestCode = g.TextMatrix(n, 0) Then
30            GetRow = n
40            Exit For
50        End If
60    Next

End Function




Private Sub g_Click()
10        On Error GoTo g_Click_Error
20        DrawChart
30        Exit Sub
g_Click_Error:
40        LogError "fFullCoagWE", "g_Click", Erl, Err.Description
End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

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




Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub lEarliest_Click()

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
140       pb.ToolTipText = Format$(ChartPositions(BestIndex).Date, "dd/mmm/yyyy") & " " & ChartPositions(BestIndex).Value
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


Private Sub txtRecords_Change()
10        On Error GoTo txtRecords_Change_Error

20        If Val(txtRecords) <> 0 Then
30            FillG
40        End If

50        Exit Sub


txtRecords_Change_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "fFullCoagWE", "txtRecords_Change", intEL, strES
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
20        If Val(txtRecords) <> 0 Then
30            txtRecords = udRecords.Value
40        End If
50        Exit Sub


udRecords_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "fFullCoagWE", "udRecords_MouseUp", intEL, strES
End Sub
