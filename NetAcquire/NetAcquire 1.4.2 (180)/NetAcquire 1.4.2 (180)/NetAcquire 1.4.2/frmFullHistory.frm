VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullHistory 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Coagulation History"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIgnorePOCT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ignore POCT results"
      Height          =   240
      Left            =   4380
      TabIndex        =   37
      Top             =   180
      Width           =   2115
   End
   Begin VB.Frame Frame3 
      Caption         =   "Plot between"
      Height          =   1620
      Left            =   6870
      TabIndex        =   30
      Top             =   5100
      Width           =   2295
      Begin VB.CommandButton cmdLeft 
         Height          =   555
         Left            =   150
         Picture         =   "frmFullHistory.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdRight 
         Height          =   555
         Left            =   1380
         Picture         =   "frmFullHistory.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   570
         Width           =   675
      End
      Begin VB.Label lblPlotFrom 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   150
         TabIndex        =   34
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label lblPlotTo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   150
         TabIndex        =   33
         Top             =   1200
         Width           =   1905
      End
   End
   Begin VB.Frame fraHL 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   3315
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "LOW"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2730
         TabIndex        =   29
         Top             =   30
         Width           =   405
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000FF&
         Caption         =   "HIGH"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2250
         TabIndex        =   28
         Top             =   30
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Abnormal Results shown "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   30
         Width           =   2145
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11400
      Top             =   -30
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   3525
      Left            =   6870
      ScaleHeight     =   3465
      ScaleMode       =   0  'User
      ScaleWidth      =   4995
      TabIndex        =   9
      Top             =   1440
      Width           =   5055
   End
   Begin VB.CommandButton bcancel 
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   11280
      Picture         =   "frmFullHistory.frx":3304
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1245
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   645
      Left            =   9420
      Picture         =   "frmFullHistory.frx":41CE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Cumulative Report"
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   6585
      Begin VB.ComboBox cmbPrinter 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Text            =   "cmbPrinter"
         Top             =   240
         Width           =   3195
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   435
         Left            =   5880
         Picture         =   "frmFullHistory.frx":44D8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Print Cumulative Report"
         Top             =   180
         Width           =   555
      End
      Begin MSComCtl2.UpDown udPrevious 
         Height          =   315
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblPrevious"
         BuddyDispid     =   196628
         OrigLeft        =   3510
         OrigTop         =   930
         OrigRight       =   4050
         OrigBottom      =   1215
         Max             =   1
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Use Printer"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Current and Previous"
         Height          =   195
         Left            =   4470
         TabIndex        =   5
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label lblPrevious 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         Height          =   315
         Left            =   4320
         TabIndex        =   4
         Top             =   240
         Width           =   585
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   420
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      FixedRows       =   4
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FormatString    =   "<Code"
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   6780
      TabIndex        =   11
      Top             =   60
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gDem 
      Height          =   3135
      Left            =   9960
      TabIndex        =   23
      Top             =   7230
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FormatString    =   "<SampleID    |<SampleDate   |<SampleTime |<Cnxn "
   End
   Begin VB.Image imgSideArrow 
      Height          =   240
      Left            =   6480
      Picture         =   "frmFullHistory.frx":4B42
      Top             =   180
      Width           =   240
   End
   Begin VB.Label lblNormalRange 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   10800
      TabIndex        =   36
      Top             =   1020
      Width           =   1125
   End
   Begin VB.Label lblParameter 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8970
      TabIndex        =   35
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Label lblSex 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11220
      TabIndex        =   25
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   10860
      TabIndex        =   24
      Top             =   390
      Width           =   270
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on Parameter to show Graph"
      Height          =   465
      Left            =   7290
      TabIndex        =   22
      Top             =   930
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   6810
      TabIndex        =   21
      Top             =   660
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   9120
      TabIndex        =   20
      Top             =   390
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7290
      TabIndex        =   19
      Top             =   660
      Width           =   4635
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88/88/8888"
      Height          =   255
      Left            =   9450
      TabIndex        =   18
      Top             =   360
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   6870
      TabIndex        =   17
      Top             =   390
      Width           =   375
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7290
      TabIndex        =   16
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lblMaxVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11940
      TabIndex        =   15
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label lblMeanVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11940
      TabIndex        =   14
      Top             =   3060
      Width           =   540
   End
   Begin VB.Label lblMinVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   11955
      TabIndex        =   13
      Top             =   4710
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6810
      Picture         =   "frmFullHistory.frx":4F99
      Top             =   930
      Width           =   480
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   9450
      TabIndex        =   12
      Top             =   6090
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmFullHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ChartPosition
    xPos As Double
    yPos As Double
    Value As Single
    Date As String
End Type

Private ChartPositions() As ChartPosition

Dim gArray() As String

Dim MaxPrevious As Integer


Private GridLeftCol As Integer
Private GridRightCol As Integer

Private LatestDateTime As Date
Private OldestDateTime As Date


Private PLow As Single
Private PHigh As Single

Private mDept As String

Public Property Let Dept(ByVal strNewValue As String)

45430     mDept = UCase$(strNewValue)

End Property

Private Sub DrawBetweenDates()

          Dim n As Long
          Dim X As Double
          Dim Y As Double
          Dim MaxVal As Single
          Dim MinVal As Single
          Dim DiffVal As Single
          Dim TotalVal As Single
          Dim GraphPoints As Single
          Dim cVal As Single
          Dim Fcolour As Long
          Dim NumberOfRecords As Integer
          Dim NumberOfMinutes As Long
          Dim MinutesOffset As Double
          Dim ValOffset As Single
          Dim BoxX As Single
          Dim BoxY As Single

45440     On Error GoTo DrawBetweenDates_Error

45450     MaxVal = 0
45460     MinVal = 9999
45470     lblMaxVal = ""
45480     lblMeanVal = ""
45490     lblMinVal = "0"
45500     TotalVal = 0
45510     GraphPoints = 0

45520     pb.AutoRedraw = False
45530     pb.Cls
45540     pb.AutoRedraw = True

45550     pb.Cls
45560     pb.Picture = LoadPicture("")

45570     NumberOfMinutes = Abs(DateDiff("n", OldestDateTime, LatestDateTime))
45580     NumberOfRecords = Abs(GridRightCol - GridLeftCol) + 1

45590     MaxVal = 0
45600     MinVal = 9999

45610     For n = GridRightCol To GridLeftCol Step -1
45620         If IsNumeric(g.TextMatrix(g.row, n)) Then
45630             cVal = Val(g.TextMatrix(g.row, n))
45640             If cVal > MaxVal Then MaxVal = cVal
45650             If cVal < MinVal Then MinVal = cVal
45660         End If
45670     Next
45680     DiffVal = MaxVal - MinVal
45690     If MaxVal = 0 And MinVal = 0 Then Exit Sub

45700     If NumberOfRecords = 1 Or NumberOfMinutes = 0 Then
45710         If NumberOfRecords = 1 Then
45720             pb.Scale (0, MaxVal * 1.05)-(100, MaxVal - (MaxVal * 0.05))
45730         Else
45740             pb.Scale (0, MaxVal * 1.05)-(100, MaxVal * 1.05 - (DiffVal * 1.15))
45750         End If
45760     Else
45770         pb.Scale (-NumberOfMinutes * 0.05, MaxVal * 1.05)-(NumberOfMinutes * 1.1, MinVal - (MaxVal * 0.05))
45780     End If

45790     pb.Line (pb.ScaleLeft, PHigh)-(pb.ScaleWidth, pb.ScaleTop), &HC0C0FF, BF
45800     pb.Line (pb.ScaleLeft, PHigh)-(pb.ScaleWidth, PHigh), vbRed, BF

45810     pb.Line (pb.ScaleLeft, PLow)-(pb.ScaleWidth, pb.ScaleHeight), &HFFFFC0, BF
45820     pb.Line (pb.ScaleLeft, PLow)-(pb.ScaleWidth, PLow), vbBlue, BF

45830     X = 0
45840     Y = 0

45850     BoxX = pb.ScaleWidth * 0.01
45860     BoxY = pb.ScaleHeight * 0.02
45870     ReDim ChartPositions(0 To NumberOfRecords - 1)
45880     For n = 0 To NumberOfRecords - 1
45890         If IsNumeric(g.TextMatrix(g.row, n + GridLeftCol)) Then
45900             MinutesOffset = Abs(DateDiff("n", g.TextMatrix(2, n + GridLeftCol) & " " & g.TextMatrix(3, n + GridLeftCol), LatestDateTime))
45910             ValOffset = g.TextMatrix(g.row, n + GridLeftCol)
45920             ChartPositions(n).xPos = pb.ScaleWidth - MinutesOffset - (pb.ScaleWidth * 0.1)
45930             ChartPositions(n).yPos = ValOffset
45940             ChartPositions(n).Value = g.TextMatrix(g.row, n + GridLeftCol)
45950             ChartPositions(n).Date = g.TextMatrix(2, n + GridLeftCol) & " " & g.TextMatrix(3, n + GridLeftCol)
45960             If ChartPositions(n).Value > PHigh Then
45970                 Fcolour = vbRed
45980             ElseIf ChartPositions(n).Value < PLow Then
45990                 Fcolour = vbBlue
46000             Else
46010                 Fcolour = vbBlack
46020             End If
46030             pb.ForeColor = Fcolour
46040             If X <> 0 Then
46050                 pb.Line (X, Y)-(ChartPositions(n).xPos, ValOffset), vbBlack
46060             End If
46070             X = ChartPositions(n).xPos
46080             Y = ValOffset
46090             pb.Circle (X, Y), BoxX
46100             pb.Line (X - BoxX / 2, Y - BoxY / 2)-(X + BoxX / 2, Y + BoxY / 2), Fcolour, BF
46110         End If
46120     Next

46130     lblMaxVal = Format$(MaxVal * 1.05, "###0.0")
46140     lblMeanVal = Format$(((MaxVal - MinVal) / 2) + MinVal, "###0.0")
46150     lblMinVal = Format$(MinVal - (MaxVal * 0.05), "###0.0")

46160     Exit Sub

DrawBetweenDates_Error:

          Dim strES As String
          Dim intEL As Integer

46170     intEL = Erl
46180     strES = Err.Description
46190     LogError "frmFullHistory", "DrawBetweenDates", intEL, strES

End Sub

Private Sub FillCodeGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim Cx As Integer
          Dim s As String
          Dim Y As Integer
          Dim Found As Boolean

46200     On Error GoTo FillCodeGrid_Error

46210     With g
46220         .Rows = 5
46230         .AddItem ""
46240         .RemoveItem 4
46250     End With

          'Get Codes and ShortNames
46260     sql = "SELECT D.PrintPriority, D.ShortName, D.Code " & _
              "FROM " & mDept & "Results R, " & mDept & "TestDefinitions D, Demographics G " & _
              "WHERE R.Code = D.Code " & _
              "AND R.SampleID = G.SampleID " & _
              "AND Chart = '" & lblChart & "' " & _
              "AND PatName = '" & AddTicks(lblName) & "' "
46270     If IsDate(lblDoB) Then
46280         sql = sql & "and DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
46290     Else
46300         sql = sql & "and (DoB is null or DoB = '') "
46310     End If
46320     sql = sql & "GROUP BY D.PrintPriority, D.ShortName, D.Code"

46330     For Cx = 0 To intOtherHospitalsInGroup
46340         Set tb = New Recordset
46350         RecOpenClient Cx, tb, sql
46360         Do While Not tb.EOF
46370             If Cx = 0 Then
46380                 s = tb!Code & vbTab & _
                          tb!ShortName & ""
46390                 g.AddItem s
46400             Else
46410                 Found = False
46420                 For Y = 4 To g.Rows - 1
46430                     If g.TextMatrix(Y, 0) = tb!Code Then
46440                         Found = True
46450                         Exit For
46460                     End If
46470                 Next
46480                 If Not Found Then
46490                     s = tb!Code & vbTab & _
                              tb!TestName & ""
46500                     g.AddItem s
46510                 End If
46520             End If
46530             tb.MoveNext
46540         Loop
46550     Next
46560     If g.Rows > 5 Then
46570         g.RemoveItem 4
46580     End If

46590     Exit Sub

FillCodeGrid_Error:

          Dim strES As String
          Dim intEL As Integer

46600     intEL = Erl
46610     strES = Err.Description
46620     LogError "frmFullHistory", "FillCodeGrid", intEL, strES, sql

End Sub

Private Sub FillResultGrid()

          Dim intCols As Integer

46630     On Error GoTo FillResultGrid_Error

46640     intCols = GetColumnCount()
46650     MaxPrevious = intCols - 3
46660     udPrevious.max = MaxPrevious
46670     If MaxPrevious < 7 Then
46680         lblPrevious = Format$(MaxPrevious)
46690         udPrevious.Value = MaxPrevious
46700     Else
46710         lblPrevious = "6"
46720         udPrevious.Value = 6
46730     End If

46740     g.Cols = intCols

46750     Exit Sub

FillResultGrid_Error:

          Dim strES As String
          Dim intEL As Integer

46760     intEL = Erl
46770     strES = Err.Description
46780     LogError "frmFullHistory", "FillResultGrid", intEL, strES

End Sub
Private Sub FillAllResults()

          Dim tb As Recordset
          Dim sql As String
          Dim X As Integer
          Dim Y As Integer
          Dim SampleID As Long
          Dim Cx As Integer
          Dim strSex As String
          Dim PerCent As Integer

46790     On Error GoTo FillAllResults_Error


46800     If Left$(lblSex, 1) = "M" Then
46810         strSex = "CT.MaleLow as Low, CT.MaleHigh as High "
46820     ElseIf Left$(lblSex, 1) = "F" Then
46830         strSex = "CT.FemaleLow as Low, CT.FemaleHigh as High "
46840     Else
46850         strSex = "CT.FemaleLow as Low, CT.MaleHigh as High "
46860     End If

46870     PerCent = 0

46880     For X = 2 To g.Cols - 1
46890         g.ColAlignment(X) = flexAlignRightCenter
46900         If X = 8 Then
46910             Me.Refresh
46920         End If

46930         PerCent = (X / g.Cols) * 100

46940         Cx = Val(g.TextMatrix(0, X))
46950         If Cx = 0 Then
46960             g.TextMatrix(0, X) = ""
46970         Else
46980             g.TextMatrix(0, X) = HospName(Cx)
46990         End If

47000         SampleID = Val(g.TextMatrix(1, X))

47010         sql = "SELECT " & _
                  "CASE WHEN ISNUMERIC(CR.Result) = 1 AND CR.Result <> '.' THEN " & _
                  "  STR(CR.Result, 6, CT.DP) " & _
                  "ELSE " & _
                  "  CR.Result " & _
                  "END Result, CR.Valid, CT.ShortName, CR.Code, " & _
                  "CT.PlausibleLow, CT.PlausibleHigh, " & _
                  strSex & _
                  "FROM " & mDept & "Results CR, " & mDept & "TestDefinitions CT " & _
                  "WHERE SampleID = '" & SampleID & "' " & _
                  "AND CR.Code = CT.Code"
47020         Set tb = New Recordset
              '    MsgBox Sql
47030         RecOpenClient Cx, tb, sql
47040         Do While Not tb.EOF
47050             For Y = 4 To g.Rows - 1
                      '            MsgBox tb!Code & "=" & g.TextMatrix(Y, 0)
47060                 If tb!Code = g.TextMatrix(Y, 0) Then
                          '                MsgBox tb!Result
47070                     If tb!Valid Then
47080                         g.TextMatrix(Y, X) = ConvertNull(tb!Result, "")
47090                         If IsNumeric(Trim(tb!Result)) Then
                                  '                        MsgBox Val(Trim(tb!Result)) & ">" & Val(Trim(tb!PlausibleHigh))
                                  '                        MsgBox Val(Trim(tb!Result)) & "<" & Val(Trim(tb!PlausibleLow))
                                  '                        MsgBox Val(Trim(tb!Result)) & ">" & Val(Trim(tb!High))
                                  '                        MsgBox Val(Trim(tb!Result)) & "<" & Val(Trim(tb!Low))
47100                             If Val(Trim(ConvertNull(tb!Result, ""))) > Val(Trim(ConvertNull(tb!PlausibleHigh, ""))) Then
47110                                 g.Col = X
47120                                 g.row = Y
47130                                 g.CellBackColor = vbBlue
47140                                 g.CellForeColor = vbYellow
47150                                 g.TextMatrix(Y, X) = "*****"
47160                             ElseIf Val(Trim(ConvertNull(tb!Result, ""))) < Val(Trim(ConvertNull(tb!PlausibleLow, ""))) Then
47170                                 g.Col = X
47180                                 g.row = Y
47190                                 g.CellBackColor = vbBlack
47200                                 g.CellForeColor = vbYellow
47210                                 g.TextMatrix(Y, X) = "*****"
47220                             ElseIf Val(Trim(ConvertNull(tb!Result, ""))) > Val(Trim(ConvertNull(tb!High, ""))) Then
47230                                 fraHL.Visible = True
47240                                 g.Col = X
47250                                 g.row = Y
47260                                 g.CellBackColor = vbRed
47270                                 g.CellForeColor = vbYellow
47280                             ElseIf Val(Trim(ConvertNull(tb!Result, ""))) < Val(Trim(ConvertNull(tb!Low, ""))) Then
47290                                 fraHL.Visible = True
47300                                 g.Col = X
47310                                 g.row = Y
47320                                 g.CellBackColor = vbBlue
47330                                 g.CellForeColor = vbYellow
47340                             Else
47350                                 g.Col = X
47360                                 g.row = Y
47370                                 g.CellBackColor = &H80000018
47380                                 g.CellForeColor = &H8000000D
47390                             End If
47400                         End If
47410                     Else
47420                         g.TextMatrix(Y, X) = "NV"
47430                     End If
47440                     If SampleID < 100000 Then
47450                         g.row = Y
47460                         g.Col = X
47470                         Set g.CellPicture = imgSideArrow.Picture
47480                     End If
47490                     Exit For
47500                 End If
47510             Next
47520             tb.MoveNext
47530         Loop
47540     Next

47550     Exit Sub

FillAllResults_Error:

          Dim strES As String
          Dim intEL As Integer

47560     intEL = Erl
47570     strES = Err.Description
47580     LogError "frmFullHistory", "FillAllResults", intEL, strES, sql

End Sub

Private Function GetColumnCount() As Integer

47590     GetColumnCount = gDem.Rows + 1

End Function


Private Sub TransferDems()

          Dim gDemY As Integer
          Dim gResX As Integer

47600     On Error GoTo TransferDems_Error

47610     For gDemY = 1 To gDem.Rows - 1
47620         gResX = gDemY + 1
47630         g.TextMatrix(0, gResX) = gDem.TextMatrix(gDemY, 3)    'Cnxn
47640         g.TextMatrix(1, gResX) = gDem.TextMatrix(gDemY, 0)    'SampleID
47650         g.TextMatrix(2, gResX) = gDem.TextMatrix(gDemY, 1)    'SampleDate
47660         g.TextMatrix(3, gResX) = gDem.TextMatrix(gDemY, 2)    'SampleTime
47670     Next

47680     Exit Sub

TransferDems_Error:

          Dim strES As String
          Dim intEL As Integer

47690     intEL = Erl
47700     strES = Err.Description
47710     LogError "frmFullHistory", "TransferDems", intEL, strES

End Sub

Private Sub bcancel_Click()

47720     Unload Me

End Sub


Private Sub chkIgnorePOCT_Click()
47730     On Error GoTo chkIgnorePOCT_Click_Error

47740     Me.Refresh
47750     FillgDem
47760     FillResultGrid
47770     FillCodeGrid
47780     TransferDems
47790     Me.Refresh
47800     FillAllResults

47810     Exit Sub

chkIgnorePOCT_Click_Error:

          Dim strES As String
          Dim intEL As Integer

47820     intEL = Erl
47830     strES = Err.Description
47840     LogError "frmFullHistory", "chkIgnorePOCT_Click", intEL, strES
          
End Sub

Private Sub cmdLeft_Click()

47850     On Error GoTo cmdLeft_Click_Error

47860     If GridRightCol < g.Cols - 1 Then
47870         GridRightCol = GridRightCol + 1
47880         OldestDateTime = Format$(g.TextMatrix(2, GridRightCol), "dd/mmm/yyyy" & " " & g.TextMatrix(3, GridRightCol))
47890         lblPlotFrom = OldestDateTime
47900         DrawBetweenDates
47910     End If

47920     Exit Sub

cmdLeft_Click_Error:

          Dim strES As String
          Dim intEL As Integer

47930     intEL = Erl
47940     strES = Err.Description
47950     LogError "frmFullHistory", "cmdLeft_Click", intEL, strES

End Sub

Private Sub cmdPrint_Click()

          Dim Px As Printer
          Dim X As Integer
          Dim Y As Integer
          Dim PageCounter As Integer
          Dim CurrentPage As Integer
          Dim TabPos As Integer
          Dim ArrayStart As Integer
          Dim ArrayStop As Integer

47960     On Error GoTo cmdPrint_Click_Error

47970     For Each Px In Printers
47980         If Px.DeviceName = cmbPrinter.Text Then
47990             Set Printer = Px
48000             Exit For
48010         End If
48020     Next

48030     PageCounter = (Val(lblPrevious) \ 7) + 1
48040     CurrentPage = 1

48050     Do While CurrentPage <= PageCounter
48060         Printer.Font.Name = "Courier New"
48070         Printer.Font.size = 14
48080         Printer.Font.Bold = True
48090         Printer.Print
48100         Printer.Print "Cumulative Biochemistry Report              ";
48110         Printer.Font.size = 10
48120         Printer.Font.Bold = False
48130         Printer.Print "Page "; Format$(CurrentPage); " of "; PageCounter
48140         Printer.Print

48150         Printer.Font.size = 14
48160         Printer.Font.Bold = True
48170         Printer.Print " Patient Name:"; lblName
48180         If lblDoB <> "" Then
48190             Printer.Print "Date of Birth:"; Format$(lblDoB, "dd/mm/yyyy")
48200         End If
48210         Printer.Print "        Chart:"; lblChart
48220         Printer.Print
48230         Printer.Print

48240         Printer.Font.size = 10
48250         Printer.Font.Bold = False

48260         For Y = 0 To UBound(gArray, 1) - 1
48270             Printer.Print gArray(Y, 0);
48280             TabPos = 10

48290             ArrayStart = ((CurrentPage - 1) * 7) + 3
48300             ArrayStop = ArrayStart + 6

                  'MaxCol = ((CurrentPage - 1) * 7) + 7
48310             If ArrayStop > Val(lblPrevious) + 3 Then
48320                 ArrayStop = Val(lblPrevious) + 3
48330             End If
48340             For X = ArrayStart To ArrayStop
48350                 If X <= UBound(gArray, 2) Then
48360                     Printer.Print Tab(TabPos); gArray(Y, X);
48370                 End If
48380                 TabPos = TabPos + 10
48390             Next
48400             Printer.Print
48410         Next

48420         Printer.EndDoc

48430         CurrentPage = CurrentPage + 1

48440     Loop

48450     Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

48460     intEL = Erl
48470     strES = Err.Description
48480     LogError "frmFullHistory", "cmdPrint_Click", intEL, strES

End Sub


Private Sub cmdRight_Click()

48490     On Error GoTo cmdRight_Click_Error

48500     If GridRightCol > 4 Then
48510         GridRightCol = GridRightCol - 1
48520         OldestDateTime = Format$(g.TextMatrix(2, GridRightCol), "dd/mmm/yyyy" & " " & g.TextMatrix(3, GridRightCol))
48530         lblPlotFrom = OldestDateTime
48540         DrawBetweenDates
48550     End If

48560     Exit Sub

cmdRight_Click_Error:

          Dim strES As String
          Dim intEL As Integer

48570     intEL = Erl
48580     strES = Err.Description
48590     LogError "frmFullHistory", "cmdRight_Click", intEL, strES

End Sub

Private Sub cmdXL_Click()

48600     ExportFlexGrid g, Me

End Sub


Private Sub Form_Activate()

48610     On Error GoTo Form_Activate_Error

48620     If UCase$(mDept) = "COAG" Then
48630         Me.Caption = "NetAcquire - Coagulation History"
48640     ElseIf UCase$(mDept) = "BIO" Then
48650         Me.Caption = "NetAcquire - Biochemistry History"
48660     End If

48670     Me.Refresh
48680     FillgDem
48690     FillResultGrid
48700     FillCodeGrid
48710     TransferDems
48720     Me.Refresh
48730     FillAllResults

48740     pBar.max = LogOffDelaySecs
48750     pBar = 0

48760     Timer1.Enabled = True

48770     Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

48780     intEL = Erl
48790     strES = Err.Description
48800     LogError "frmFullHistory", "Form_Activate", intEL, strES

End Sub


Private Sub FillgDem()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim s As String

48810     On Error GoTo FillgDem_Error

48820     With gDem
48830         .Visible = False
48840         .Rows = 2
48850         .AddItem ""
48860         .RemoveItem 1
48870     End With

48880     sql = "SELECT DISTINCT D.SampleID, SampleDate FROM Demographics D, " & mDept & "Results R WHERE " & _
              "Chart = '" & lblChart & "' " & _
              "AND PatName = '" & AddTicks(lblName) & "' "
48890     If IsDate(lblDoB) Then
48900         sql = sql & "AND DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
48910     Else
48920         sql = sql & "AND (DoB IS NULL OR DoB = '') "
48930     End If
48940     sql = sql & "AND R.SampleID = D.SampleID ORDER BY SampleDate desc "

48950     For n = 0 To intOtherHospitalsInGroup
48960         Set tb = New Recordset
48970         RecOpenClient n, tb, sql
48980         Do While Not tb.EOF
48990             s = tb!SampleID & vbTab & _
                      Format$(tb!SampleDate, "dd/MM/yy") & vbTab
49000             If Format(tb!SampleDate, "HH:mm") <> "00:00" Then
49010                 s = s & Format$(tb!SampleDate, "HH:mm")
49020             End If
49030             s = s & vbTab & Format$(n)
                  'gDem.AddItem S
49040             If chkIgnorePOCT.Value = 0 Then
49050                 gDem.AddItem s
49060             Else
49070                 If tb!SampleID > 99999 Then
                          'only add samples if it's not POCT autogenerated sampleid. 1
                          '1000 to 99999 are auto generated sample ids for POCT
49080                     gDem.AddItem s
49090                 End If
49100             End If
49110             tb.MoveNext
49120         Loop
49130     Next

49140     With gDem
49150         If .Rows > 2 Then
49160             .RemoveItem 1
49170             .Visible = True
49180         End If
49190     End With

49200     Exit Sub

FillgDem_Error:

          Dim strES As String
          Dim intEL As Integer

49210     intEL = Erl
49220     strES = Err.Description
49230     LogError "frmFullHistory", "FillgDem", intEL, strES, sql

End Sub

Private Sub Form_Deactivate()

49240     Timer1.Enabled = False

End Sub


Private Sub Form_Load()

49250     g.ColWidth(0) = 0
49260     ReDim ChartPositions(0 To 0)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

49270     pBar = 0
49280     pb.AutoRedraw = False
49290     pb.Cls

End Sub


Private Sub g_Click()

          Dim X As Long
          Dim Y As Long
          Dim sql As String
          Dim tb As Recordset

49300     On Error GoTo g_Click_Error

49310     Y = g.row
49320     X = g.Col

49330     If Y < 4 Then Exit Sub

49340     If Trim(g.TextMatrix(Y, X)) <> "" Then g.ToolTipText = g.TextMatrix(Y, X)
49350     lblParameter = g.TextMatrix(Y, 1)

49360     sql = "SELECT MaleLow, MaleHigh FROM " & mDept & "TestDefinitions WHERE Code = '" & g.TextMatrix(Y, 0) & "'"
49370     Set tb = New Recordset
49380     RecOpenClient 0, tb, sql
49390     If Not tb.EOF Then
49400         PLow = tb!MaleLow
49410         PHigh = tb!MaleHigh
49420     Else
49430         PLow = 0
49440         PHigh = 9999
49450     End If
49460     lblNormalRange = "Normal Range " & PLow & " - " & PHigh

49470     DrawChartNew

49480     Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

49490     intEL = Erl
49500     strES = Err.Description
49510     LogError "frmFullHistory", "g_Click", intEL, strES, sql

End Sub


Private Sub DrawChartNew()

          Dim MaxVal As Single
          Dim MinVal As Single
          Dim TotalVal As Single
          Dim GraphPoints As Single

49520     On Error GoTo DrawChartNew_Error

49530     MaxVal = 0
49540     MinVal = 9999
49550     lblMaxVal = ""
49560     lblMeanVal = ""
49570     lblMinVal = "0"
49580     TotalVal = 0
49590     GraphPoints = 0

49600     pb.AutoRedraw = False
49610     pb.Cls
49620     pb.AutoRedraw = True

49630     pb.Cls
49640     pb.Picture = LoadPicture("")

49650     GridLeftCol = 0
49660     GridRightCol = 0

49670     SetDates

49680     If GridLeftCol = 0 Or GridRightCol = 0 Then Exit Sub

49690     DrawBetweenDates

49700     Exit Sub

DrawChartNew_Error:

          Dim strES As String
          Dim intEL As Integer

49710     intEL = Erl
49720     strES = Err.Description
49730     LogError "frmFullHistory", "DrawChartNew", intEL, strES

End Sub

Private Sub SetDates()

          Dim n As Integer

49740     On Error GoTo SetDates_Error

49750     For n = 2 To g.Cols - 1
49760         If IsNumeric(g.TextMatrix(g.row, n)) Then
49770             GridLeftCol = n
49780             LatestDateTime = Format$(g.TextMatrix(2, n), "dd/mmm/yyyy" & " " & g.TextMatrix(3, n))
49790             lblPlotTo = LatestDateTime
49800             Exit For
49810         End If
49820     Next
49830     For n = g.Cols - 1 To 2 Step -1
49840         If IsNumeric(g.TextMatrix(g.row, n)) Then
49850             GridRightCol = n
49860             OldestDateTime = Format$(g.TextMatrix(2, n), "dd/mmm/yyyy" & " " & g.TextMatrix(3, n))
49870             lblPlotFrom = OldestDateTime
49880             Exit For
49890         End If
49900     Next

49910     Exit Sub

SetDates_Error:

          Dim strES As String
          Dim intEL As Integer

49920     intEL = Erl
49930     strES = Err.Description
49940     LogError "frmFullHistory", "SetDates", intEL, strES

End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

49950     pBar = 0

End Sub



Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

49960     pBar = 0

End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

49970     pBar = 0

End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

49980     pBar = 0

End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

49990     pBar = 0

End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

50000     pBar = 0

End Sub


Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

50010     pBar = 0

End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

50020     pBar = 0

End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

50030     pBar = 0

End Sub


Private Sub pb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim i As Long
          Dim CurrentDistance As Double
          Dim BestDistance As Double
          Dim BestIndex As Integer

50040     On Error GoTo pb_MouseMove_Error

50050     pb.ToolTipText = ""
50060     If UBound(ChartPositions) = 0 Then Exit Sub

50070     BestIndex = -1
50080     BestDistance = 99999999
50090     For i = 0 To UBound(ChartPositions)
50100         CurrentDistance = (((((X - ChartPositions(i).xPos) / pb.ScaleWidth) ^ 2) + ((Y - ChartPositions(i).yPos) / pb.ScaleHeight) ^ 2)) ^ (1 / 2)
50110         If i = 0 Or CurrentDistance < BestDistance Then
50120             BestDistance = CurrentDistance
50130             BestIndex = i
50140         End If
50150     Next

50160     If BestIndex <> -1 Then
50170         pb.ToolTipText = "Date:" & ChartPositions(BestIndex).Date & " Result: " & ChartPositions(BestIndex).Value
50180         pb.AutoRedraw = False
50190         pb.Cls
50200         pb.DrawWidth = 2
50210         pb.Line (X, Y)-(ChartPositions(BestIndex).xPos, ChartPositions(BestIndex).yPos), vbYellow
50220         pb.DrawWidth = 1
50230     End If

50240     Exit Sub

pb_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

50250     intEL = Erl
50260     strES = Err.Description
50270     LogError "frmFullHistory", "pb_MouseMove", intEL, strES

End Sub


Private Sub Timer1_Timer()

          'tmrRefresh.Interval set to 1000
50280     pBar = pBar + 1

50290     If pBar = pBar.max Then
50300         Unload Me
50310     End If

End Sub


