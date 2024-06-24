VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullBGA 
   Caption         =   "NetAcquire - Full Blood Gas History"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   11595
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   2325
      Left            =   6840
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   6
      Top             =   2220
      Width           =   4185
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10530
      Top             =   5310
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plot between"
      Height          =   705
      Left            =   6840
      TabIndex        =   1
      Top             =   1440
      Width           =   4185
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H80000016&
         Caption         =   "Go"
         Height          =   315
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   555
      End
      Begin VB.ComboBox cmbPlotFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Text            =   "cmbPlotFrom"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbPlotTo 
         Height          =   315
         Left            =   2010
         TabIndex        =   2
         Text            =   "cmbPlotTo"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton bcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   585
      Left            =   8010
      Picture         =   "frmFullBGA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4830
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   6750
      TabIndex        =   5
      Top             =   5460
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5625
      Left            =   60
      TabIndex        =   7
      Top             =   30
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9922
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
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7260
      TabIndex        =   17
      Top             =   60
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   6840
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9480
      TabIndex        =   15
      Top             =   60
      Width           =   1545
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7260
      TabIndex        =   14
      Top             =   360
      Width           =   3765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   9120
      TabIndex        =   13
      Top             =   90
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   6810
      TabIndex        =   12
      Top             =   360
      Width           =   420
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on Parameter to show Graph"
      Height          =   285
      Left            =   7260
      TabIndex        =   11
      Top             =   840
      Width           =   2865
   End
   Begin VB.Label lblMaxVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11040
      TabIndex        =   10
      Top             =   2220
      Width           =   510
   End
   Begin VB.Label lblMeanVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11040
      TabIndex        =   9
      Top             =   3270
      Width           =   510
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   11040
      TabIndex        =   8
      Top             =   4290
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6780
      Picture         =   "frmFullBGA.frx":066A
      Top             =   750
      Width           =   480
   End
End
Attribute VB_Name = "frmFullBGA"
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

Private Sub DrawChart()

          Dim n As Integer
          Dim Counter As Integer
          Dim DaysInterval As Long
          Dim X As Integer
          Dim Y As Integer
          Dim PixelsPerDay As Single
          Dim PixelsPerPointY As Single
          Dim FirstDayFilled As Boolean
          Dim MaxVal As Single
          Dim cVal As Single
          Dim StartGridX As Integer
          Dim StopGridX As Integer

41810     On Error GoTo DrawChart_Error

41820     MaxVal = 0
41830     lblMaxVal = ""
41840     lblMeanVal = ""

41850     pb.Cls
41860     pb.Picture = LoadPicture("")

41870     NumberOfDays = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
41880     If NumberOfDays = 0 Then Exit Sub

41890     ReDim ChartPositions(0 To NumberOfDays)

41900     For n = 1 To NumberOfDays
41910         ChartPositions(n).xPos = 0
41920         ChartPositions(n).yPos = 0
41930         ChartPositions(n).Value = 0
41940         ChartPositions(n).Date = ""
41950     Next

41960     For n = 1 To g.Cols - 1
41970         If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
41980         If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
41990     Next

42000     FirstDayFilled = False
42010     Counter = 0
42020     For X = StartGridX To StopGridX
42030         If g.TextMatrix(g.row, X) <> "" Then
42040             If Not FirstDayFilled Then
42050                 FirstDayFilled = True
42060                 MaxVal = Val(g.TextMatrix(g.row, X))
42070                 ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, X), "dd/mmm/yyyy")
42080                 ChartPositions(NumberOfDays).Value = Val(g.TextMatrix(g.row, X))
42090             Else
42100                 DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")))
42110                 ChartPositions(NumberOfDays - DaysInterval).Date = g.TextMatrix(1, X)
42120                 cVal = Val(g.TextMatrix(g.row, X))
42130                 ChartPositions(NumberOfDays - DaysInterval).Value = cVal
42140                 If cVal > MaxVal Then MaxVal = cVal
42150             End If
42160         End If
42170     Next

42180     PixelsPerDay = (pb.width - 1060) / NumberOfDays
42190     MaxVal = MaxVal * 1.1
42200     If MaxVal = 0 Then Exit Sub
42210     PixelsPerPointY = pb.height / MaxVal

42220     X = 580 + (NumberOfDays * PixelsPerDay)
42230     Y = pb.height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
42240     ChartPositions(NumberOfDays).yPos = Y
42250     ChartPositions(NumberOfDays).xPos = X

42260     pb.ForeColor = vbBlue
42270     pb.Circle (X, Y), 30
42280     pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
42290     pb.PSet (X, Y)

42300     For n = NumberOfDays - 1 To 0 Step -1
42310         If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
42320             DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
42330             X = 580 + (DaysInterval * PixelsPerDay)
42340             ChartPositions(n).xPos = X
42350             Y = pb.height - (ChartPositions(n).Value * PixelsPerPointY)
42360             ChartPositions(n).yPos = Y
42370             pb.Line -(X, Y)
42380             pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
42390             pb.Circle (X, Y), 30
42400             pb.PSet (X, Y)
42410         End If
42420     Next

42430     pb.Line (0, pb.height / 2)-(pb.width, pb.height / 2), vbBlack, BF

42440     lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
42450     lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")

42460     Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

42470     intEL = Erl
42480     strES = Err.Description
42490     LogError "fFullBGA", "DrawChart", intEL, strES

End Sub

Private Sub DrawHeadings()

42500     g.TextMatrix(0, 0) = "SampleID"
42510     g.TextMatrix(1, 0) = "Run Date"
42520     g.TextMatrix(2, 0) = "Run Time"
42530     g.TextMatrix(3, 0) = "pH"
42540     g.TextMatrix(4, 0) = "PCO2"
42550     g.TextMatrix(5, 0) = "PO2"
42560     g.TextMatrix(6, 0) = "HCO3"
42570     g.TextMatrix(7, 0) = "BE"
42580     g.TextMatrix(8, 0) = "O2Sat"
42590     g.TextMatrix(9, 0) = "Tot CO2"

End Sub

Private Sub FillCombos()

          Dim X As Integer

42600     cmbPlotFrom.Clear
42610     cmbPlotTo.Clear

42620     For X = 1 To g.Cols - 1
42630         cmbPlotFrom.AddItem Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
42640         cmbPlotTo.AddItem Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
42650     Next

42660     cmbPlotTo = Format$(g.TextMatrix(1, 1), "dd/mmm/yyyy")

42670     For X = g.Cols - 1 To 1 Step -1
42680         If DateDiff("d", Format$(g.TextMatrix(1, X), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
42690             cmbPlotFrom = Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
42700             Exit For
42710         End If
42720     Next

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim gcolumns As Integer
          Dim X As Integer
        
42730     On Error GoTo FillG_Error

42740     g.Rows = 4
42750     g.AddItem ""
42760     g.RemoveItem 3
42770     g.Rows = 10
42780     DrawHeadings

42790     sql = "select B.* from BGAResults as B, Demographics as D where " & _
              "D.PatName = '" & AddTicks(lblName) & "' " & _
              "and B.SampleID = D.SampleID " & _
              "order by B.RunDateTime, B.SampleID"

42800     Set tb = New Recordset
42810     RecOpenClient 0, tb, sql
42820     If Not tb.EOF Then
42830         g.Visible = False
42840         tb.MoveLast
42850         gcolumns = tb.RecordCount
42860         g.Cols = gcolumns + 1
42870         g.ColWidth(0) = 1095

              'SampleID and sampledate across
42880         For X = 1 To gcolumns
42890             g.ColWidth(X) = 1095
42900             g.Col = X
42910             g.TextMatrix(0, X) = Trim$(tb!SampleID & "")
42920             g.TextMatrix(1, X) = Format(tb!Rundate, "dd/mm/yy")
42930             g.TextMatrix(2, X) = Format(tb!RunDateTime, "hh:mm")
42940             g.TextMatrix(3, X) = tb!pH & ""
42950             g.TextMatrix(4, X) = tb!PCO2 & ""
42960             g.TextMatrix(5, X) = tb!PO2 & ""
42970             g.TextMatrix(6, X) = tb!HCO3 & ""
42980             g.TextMatrix(7, X) = tb!BE & ""
42990             g.TextMatrix(8, X) = tb!O2SAT & ""
43000             g.TextMatrix(9, X) = tb!TotCO2 & ""
43010             tb.MovePrevious
43020         Next
43030     End If

43040     g.Visible = True

43050     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

43060     intEL = Erl
43070     strES = Err.Description
43080     LogError "fFullBGA", "FillG", intEL, strES, sql

End Sub


Private Sub bcancel_Click()

43090     Unload Me

End Sub


Private Sub cmdGo_Click()

43100     DrawChart

End Sub


Private Sub Form_Activate()

43110     FillG
43120     FillCombos

43130     pBar.max = LogOffDelaySecs
43140     pBar = 0

43150     Timer1.Enabled = True

End Sub

Private Sub Form_Deactivate()

43160     Timer1.Enabled = False

End Sub


Private Sub Form_Load()

43170     pBar.max = LogOffDelaySecs
43180     pBar = 0

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

43190     pBar = 0

End Sub


Private Sub g_Click()

43200     DrawChart

End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

43210     pBar = 0

End Sub


Private Sub pb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim i As Integer
          Dim CurrentDistance As Long
          Dim BestDistance As Long
          Dim BestIndex As Integer

43220     On Error GoTo pbmm

43230     pBar = 0

43240     If NumberOfDays = 0 Then Exit Sub

43250     BestIndex = -1
43260     BestDistance = 99999
43270     For i = 0 To NumberOfDays
43280         CurrentDistance = ((X - ChartPositions(i).xPos) ^ 2 + (Y - ChartPositions(i).yPos) ^ 2) ^ (1 / 2)
43290         If i = 0 Or CurrentDistance < BestDistance Then
43300             BestDistance = CurrentDistance
43310             BestIndex = i
43320         End If
43330     Next

43340     If BestIndex <> -1 Then
43350         pb.ToolTipText = Format$(ChartPositions(BestIndex).Date, "dd/mmm/yyyy") & " " & ChartPositions(BestIndex).Value
43360     End If

43370     Exit Sub

pbmm:
43380     Exit Sub

End Sub


Private Sub Timer1_Timer()

          'tmrRefresh.Interval set to 1000
43390     pBar = pBar + 1
        
43400     If pBar = pBar.max Then
43410         Unload Me
43420     End If

End Sub


