VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullHaem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Haematology History"
   ClientHeight    =   5955
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   11760
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5955
   ScaleWidth      =   11760
   Begin VB.CommandButton bcancel 
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
      Height          =   585
      Left            =   8010
      Picture         =   "frmFullHaem.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4860
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
      Height          =   705
      Left            =   6840
      TabIndex        =   11
      Top             =   1470
      Width           =   4185
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
         Left            =   2010
         TabIndex        =   14
         Text            =   "cmbPlotTo"
         Top             =   240
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
         Top             =   240
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
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   555
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   6750
      TabIndex        =   10
      Top             =   5490
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10530
      Top             =   5340
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   2325
      Left            =   6840
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   8
      Top             =   2250
      Width           =   4185
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5625
      Left            =   30
      TabIndex        =   7
      Top             =   60
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
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   -120
      TabIndex        =   6
      Top             =   2010
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6780
      Picture         =   "frmFullHaem.frx":066A
      Top             =   780
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
      Left            =   11040
      TabIndex        =   17
      Top             =   4320
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
      Left            =   11040
      TabIndex        =   16
      Top             =   3300
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
      Left            =   11040
      TabIndex        =   15
      Top             =   2250
      Width           =   510
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
      Left            =   7260
      TabIndex        =   9
      Top             =   870
      Width           =   2865
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
      Left            =   6810
      TabIndex        =   5
      Top             =   390
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
      Left            =   9120
      TabIndex        =   4
      Top             =   120
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7260
      TabIndex        =   3
      Top             =   390
      Width           =   3765
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9480
      TabIndex        =   2
      Top             =   90
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
      Left            =   6840
      TabIndex        =   1
      Top             =   150
      Width           =   375
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7260
      TabIndex        =   0
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "frmFullHaem"
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

43430     On Error GoTo DrawChart_Error

43440     MaxVal = 0
43450     lblMaxVal = ""
43460     lblMeanVal = ""

43470     pb.Cls
43480     pb.Picture = LoadPicture("")

43490     NumberOfDays = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
43500     If NumberOfDays = 0 Then Exit Sub
43510     ReDim ChartPositions(0 To NumberOfDays)

43520     For n = 1 To NumberOfDays
43530         ChartPositions(n).xPos = 0
43540         ChartPositions(n).yPos = 0
43550         ChartPositions(n).Value = 0
43560         ChartPositions(n).Date = ""
43570     Next

43580     For n = 1 To g.Cols - 1
43590         If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
43600         If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
43610     Next

43620     FirstDayFilled = False
43630     Counter = 0
43640     For X = StartGridX To StopGridX
43650         If g.TextMatrix(g.row, X) <> "" Then
43660             If Not FirstDayFilled Then
43670                 FirstDayFilled = True
43680                 MaxVal = Val(g.TextMatrix(g.row, X))
43690                 ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, X), "dd/mmm/yyyy")
43700                 ChartPositions(NumberOfDays).Value = Val(g.TextMatrix(g.row, X))
43710             Else
43720                 DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")))
43730                 ChartPositions(NumberOfDays - DaysInterval).Date = g.TextMatrix(1, X)
43740                 cVal = Val(g.TextMatrix(g.row, X))
43750                 ChartPositions(NumberOfDays - DaysInterval).Value = cVal
43760                 If cVal > MaxVal Then MaxVal = cVal
43770             End If
43780         End If
43790     Next

43800     PixelsPerDay = (pb.width - 1060) / NumberOfDays
43810     MaxVal = MaxVal * 1.1
43820     If MaxVal = 0 Then Exit Sub
43830     PixelsPerPointY = pb.height / MaxVal

43840     X = 580 + (NumberOfDays * PixelsPerDay)
43850     Y = pb.height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
43860     ChartPositions(NumberOfDays).yPos = Y
43870     ChartPositions(NumberOfDays).xPos = X

43880     pb.ForeColor = vbBlue
43890     pb.Circle (X, Y), 30
43900     pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
43910     pb.PSet (X, Y)

43920     For n = NumberOfDays - 1 To 0 Step -1
43930         If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
43940             DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
43950             X = 580 + (DaysInterval * PixelsPerDay)
43960             ChartPositions(n).xPos = X
43970             Y = pb.height - (ChartPositions(n).Value * PixelsPerPointY)
43980             ChartPositions(n).yPos = Y
43990             pb.Line -(X, Y)
44000             pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
44010             pb.Circle (X, Y), 30
44020             pb.PSet (X, Y)
44030         End If
44040     Next

44050     pb.Line (0, pb.height / 2)-(pb.width, pb.height / 2), vbBlack, BF

44060     lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
44070     lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")

44080     Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

44090     intEL = Erl
44100     strES = Err.Description
44110     LogError "fFullHaem", "DrawChart", intEL, strES

End Sub

Private Sub bcancel_Click()

44120     Unload Me

End Sub

Private Sub FillG()

          Dim sn As Recordset
          Dim tb As Recordset
          Dim sql As String
          Dim gcolumns As Integer
          Dim X As Integer
          Dim xrun As String
          Dim xdate As String
        
44130     On Error GoTo FillG_Error

44140     g.Rows = 4
44150     g.AddItem ""
44160     g.RemoveItem 3

44170     sql = "SELECT D.SampleID, D.RunDate, D.SampleDate " & _
              "FROM Demographics D JOIN HaemResults H " & _
              "ON D.SampleID = H.SampleID " & _
              "WHERE D.Chart = '" & lblChart & "' " & _
              "AND D.PatName = '" & AddTicks(lblName) & "' "
44180     If IsDate(lblDoB) Then
44190         sql = sql & "AND D.DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
44200     Else
44210         sql = sql & "AND (D.DoB IS NULL OR D.DoB = '') "
44220     End If
44230     sql = sql & "ORDER BY D.SampleDate ASC,D.SampleID ASC"
44240     Set sn = New Recordset
44250     RecOpenClient 0, sn, sql
44260     If Not sn.EOF Then
44270         g.Visible = False
44280         sn.MoveLast
44290         gcolumns = sn.RecordCount
44300         g.Cols = gcolumns + 1
44310         g.ColWidth(0) = 1095

              'SampleID and sampledate across
44320         For X = 1 To gcolumns
44330             g.ColWidth(X) = 1095
44340             g.Col = X
44350             xrun = sn!SampleID & ""
44360             g.TextMatrix(0, X) = xrun
44370             If Not IsNull(sn!Rundate) Then
44380                 xdate = Format(sn!Rundate, "dd/mm/yy")
44390             Else
44400                 xdate = ""
44410             End If
44420             g.TextMatrix(1, X) = xdate
44430             If IsDate(sn!SampleDate) Then
44440                 g.TextMatrix(2, X) = Format(sn!SampleDate, "hh:mm")
44450             Else
44460                 g.TextMatrix(2, X) = ""
44470             End If
                  'fill list with test names
44480             sn.MovePrevious
44490         Next
44500     End If

44510     g.AddItem "WBC"
44520     g.AddItem "RBC"
44530     g.AddItem "Hgb"
44540     g.AddItem "Hct"
44550     g.AddItem "MCV"
44560     g.AddItem "Plt"
44570     g.AddItem "RDW"
44580     g.AddItem "Lymp A"
44590     g.AddItem "Mono A"
44600     g.AddItem "Neut A"
44610     g.AddItem "Eos A"
44620     g.AddItem "Bas A"
44630     g.AddItem "ESR"

44640     For X = 1 To g.Cols - 1
44650         sql = "Select * from HaemResults where " & _
                  "SampleID = '" & g.TextMatrix(0, X) & "'"
44660         Set tb = New Recordset
44670         RecOpenServer 0, tb, sql
44680         If Not tb.EOF Then
44690             If Not IsNull(tb!Valid) And tb!Valid Then
44700                 g.TextMatrix(4, X) = tb!WBC & ""
44710                 g.TextMatrix(5, X) = tb!rbc & ""
44720                 g.TextMatrix(6, X) = tb!Hgb & ""
44730                 g.TextMatrix(7, X) = tb!hct & ""
44740                 g.TextMatrix(8, X) = tb!MCV & ""
44750                 g.TextMatrix(9, X) = tb!plt & ""
44760                 g.TextMatrix(10, X) = tb!RDWCV & ""
44770                 g.TextMatrix(11, X) = tb!LymA & ""
44780                 g.TextMatrix(12, X) = tb!MonoA & ""
44790                 g.TextMatrix(13, X) = tb!NeutA & ""
44800                 g.TextMatrix(14, X) = tb!EosA & ""
44810                 g.TextMatrix(15, X) = tb!BasA & ""
44820                 g.TextMatrix(16, X) = tb!ESR & ""
44830             Else
44840                 g.TextMatrix(4, X) = "Not"
44850                 g.TextMatrix(5, X) = "Valid"
44860             End If
44870         End If
44880     Next

44890     If g.Rows > 4 Then
44900         g.RemoveItem 3
44910     End If
44920     g.Visible = True

44930     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

44940     intEL = Erl
44950     strES = Err.Description
44960     LogError "fFullHaem", "FillG", intEL, strES, sql

End Sub

Private Sub cmdGo_Click()

44970     DrawChart

End Sub

Private Sub Form_Activate()

44980     FillG
44990     FillCombos

45000     pBar.max = LogOffDelaySecs
45010     pBar = 0

45020     Timer1.Enabled = True

End Sub

Private Sub Form_Deactivate()

45030     Timer1.Enabled = False

End Sub


Private Sub Form_Load()

45040     pBar.max = LogOffDelaySecs
45050     pBar = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

45060     pBar = 0

End Sub

Private Sub g_Click()

45070     DrawChart

End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

45080     pBar = 0

End Sub

Private Sub pb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim i As Integer
          Dim CurrentDistance As Long
          Dim BestDistance As Long
          Dim BestIndex As Integer

45090     On Error GoTo pbmm

45100     pBar = 0

45110     If NumberOfDays = 0 Then Exit Sub

45120     BestIndex = -1
45130     BestDistance = 99999
45140     For i = 0 To NumberOfDays
45150         CurrentDistance = ((X - ChartPositions(i).xPos) ^ 2 + (Y - ChartPositions(i).yPos) ^ 2) ^ (1 / 2)
45160         If i = 0 Or CurrentDistance < BestDistance Then
45170             BestDistance = CurrentDistance
45180             BestIndex = i
45190         End If
45200     Next

45210     If BestIndex <> -1 Then
45220         pb.ToolTipText = Format$(ChartPositions(BestIndex).Date, "dd/mmm/yyyy") & " " & ChartPositions(BestIndex).Value
45230     End If

45240     Exit Sub

pbmm:
45250     Exit Sub

End Sub

Private Sub FillCombos()

          Dim X As Integer

45260     cmbPlotFrom.Clear
45270     cmbPlotTo.Clear

45280     For X = 1 To g.Cols - 1
45290         cmbPlotFrom.AddItem Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
45300         cmbPlotTo.AddItem Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
45310     Next

45320     cmbPlotTo = Format$(g.TextMatrix(1, 1), "dd/mmm/yyyy")

45330     For X = g.Cols - 1 To 1 Step -1
45340         If DateDiff("d", Format$(g.TextMatrix(1, X), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
45350             cmbPlotFrom = Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
45360             Exit For
45370         End If
45380     Next

End Sub


Private Sub Timer1_Timer()

          'tmrRefresh.Interval set to 1000
45390     pBar = pBar + 1
        
45400     If pBar = pBar.max Then
45410         Unload Me
45420     End If

End Sub


