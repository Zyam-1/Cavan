VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTestSequence 
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Test Names"
      Height          =   765
      Left            =   5730
      TabIndex        =   8
      Top             =   240
      Width           =   1365
      Begin VB.OptionButton optName 
         Caption         =   "Long"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
      Begin VB.OptionButton optName 
         Caption         =   "Short"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "&Export to Excel"
      Height          =   945
      Left            =   5970
      Picture         =   "frmTestSequence.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4860
      Width           =   945
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6150
      Top             =   1710
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6150
      Top             =   2490
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   1065
      Left            =   5970
      Picture         =   "frmTestSequence.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6180
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   765
      Left            =   5610
      Picture         =   "frmTestSequence.frx":6ADC
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2220
      Width           =   525
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5610
      Picture         =   "frmTestSequence.frx":845E
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1590
      Width           =   525
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   5970
      Picture         =   "frmTestSequence.frx":9DE0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7320
      Width           =   945
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   1065
      Left            =   5970
      Picture         =   "frmTestSequence.frx":ACAA
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3480
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   8055
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   330
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   14208
      _Version        =   393216
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
      FormatString    =   "<Analyte                                                       |^Printable "
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
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmTestSequence.frx":BB74
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmTestSequence.frx":BE4A
      Top             =   0
      Visible         =   0   'False
      Width           =   210
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
      Height          =   285
      Left            =   5730
      TabIndex        =   7
      Top             =   5850
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmTestSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String

Private sDiscipline As String

Private Activated As Boolean

Private FireCounter As Integer

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim Names As String

20060 On Error GoTo FillG_Error

20070 g.Visible = False
20080 g.Rows = 2
20090 g.AddItem ""
20100 g.RemoveItem 1

20110 Names = IIf(optName(0), "Long", "Short")
        
20120 sql = "SELECT DISTINCT " & Names & "Name N, PrintPriority, Printable " & _
            "FROM " & pDiscipline & "TestDefinitions " & _
            "ORDER BY PrintPriority"
20130 Set tb = New Recordset
20140 RecOpenClient 0, tb, sql
20150 Do While Not tb.EOF
20160   g.AddItem tb!n & ""
20170   g.row = g.Rows - 1
20180   g.Col = 1
20190   If tb!Printable <> 0 Then
20200     Set g.CellPicture = imgGreenTick.Picture
20210   Else
20220     Set g.CellPicture = imgRedCross.Picture
20230   End If
20240   g.CellPictureAlignment = flexAlignCenterCenter
20250   tb.MoveNext
20260 Loop

20270 If g.Rows > 2 Then
20280   g.RemoveItem 1
20290 End If
20300 g.Visible = True

20310 cmdSave.Visible = False

20320 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

20330 intEL = Erl
20340 strES = Err.Description
20350 LogError "frmTestSequence", "FillG", intEL, strES, sql
20360 g.Visible = True

End Sub

Private Sub cmdCancel_Click()

20370 Unload Me

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20380 FireDown

20390 tmrDown.Interval = 250
20400 FireCounter = 0

20410 tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

20420 tmrDown.Enabled = False

20430 cmdSave.Visible = True
20440 optName(0).Enabled = False
20450 optName(1).Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20460 FireUp

20470 tmrUp.Interval = 250
20480 FireCounter = 0

20490 tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

20500 tmrUp.Enabled = False

20510 cmdSave.Visible = True
20520 optName(0).Enabled = False
20530 optName(1).Enabled = False

End Sub


Private Sub cmdPrint_Click()

20540 On Error GoTo cmdPrint_Click_Error

20550 Screen.MousePointer = vbHourglass

20560 Printer.Print

20570 Printer.Print "List of "; sDiscipline; " Sequence"

20580 g.Col = 0
20590 g.row = 1
20600 g.ColSel = g.Cols - 1
20610 g.RowSel = g.Rows - 1

20620 Printer.Print g.Clip

20630 Printer.EndDoc

20640 Screen.MousePointer = vbDefault

20650 Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

20660 intEL = Erl
20670 strES = Err.Description
20680 LogError "frmTestSequence", "cmdPrint_Click", intEL, strES

20690 Screen.MousePointer = vbDefault

End Sub


Private Sub cmdSave_Click()

      Dim sql As String
      Dim Names As String
      Dim n As Integer
      Dim Printable As Integer

20700 On Error GoTo cmdSave_Click_Error

20710 Names = IIf(optName(0), "Long", "Short")
        
20720 For n = 1 To g.Rows - 1
20730   g.row = n
20740   g.Col = 1
20750   Printable = IIf(g.CellPicture = imgGreenTick.Picture, 1, 0)
20760   sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
              "SET PrintPriority = '" & n & "', " & _
              "Printable = " & Printable & " " & _
              "WHERE " & Names & "Name = '" & AddTicks(g.TextMatrix(n, 0)) & "'"
20770   Cnxn(0).Execute sql
20780 Next

20790 cmdSave.Visible = False
20800 optName(0).Enabled = True
20810 optName(1).Enabled = True

20820 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

20830 intEL = Erl
20840 strES = Err.Description
20850 LogError "frmTestSequence", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()

20860 ExportFlexGrid g, Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

20870 If cmdSave.Visible Then
20880   If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
20890     Cancel = True
20900     Exit Sub
20910   End If
20920 End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

20930 Activated = False

End Sub

Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer

20940 On Error GoTo g_Click_Error

20950 ySave = g.row

20960 If g.row > 0 And g.Col = 1 Then
20970   If g.CellPicture = imgRedCross.Picture Then
20980     Set g.CellPicture = imgGreenTick.Picture
20990   Else
21000     Set g.CellPicture = imgRedCross.Picture
21010   End If
21020   cmdSave.Visible = True
21030 End If

21040 g.Visible = False
21050 g.Col = 0
21060 For Y = 1 To g.Rows - 1
21070   g.row = Y
21080   If g.CellBackColor = vbYellow Then
21090     For X = 0 To g.Cols - 1
21100       g.Col = X
21110       g.CellBackColor = 0
21120     Next
21130     Exit For
21140   End If
21150 Next
21160 g.row = ySave
21170 g.Visible = True

21180 If g.MouseRow = 0 Then
21190   If SortOrder Then
21200     g.Sort = flexSortGenericAscending
21210   Else
21220     g.Sort = flexSortGenericDescending
21230   End If
21240   SortOrder = Not SortOrder
21250   Exit Sub
21260 End If

21270 For X = 0 To g.Cols - 1
21280   g.Col = X
21290   g.CellBackColor = vbYellow
21300 Next

21310 cmdMoveUp.Enabled = True
21320 cmdMoveDown.Enabled = True

21330 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

21340 intEL = Erl
21350 strES = Err.Description
21360 LogError "frmListsGeneric", "g_Click", intEL, strES

End Sub



Private Sub optName_Click(Index As Integer)

21370 FillG

End Sub

Private Sub tmrDown_Timer()

21380 FireDown

End Sub


Private Sub tmrUp_Timer()

21390 FireUp

End Sub

Private Sub Form_Activate()

21400 Me.Caption = "NetAcquire - " & sDiscipline & " Test Sequence"

21410 If Not Activated Then
21420   FillG
21430   Activated = True
21440 End If

End Sub

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

21450 If g.row = g.Rows - 1 Then Exit Sub
21460 n = g.row

21470 FireCounter = FireCounter + 1
21480 If FireCounter > 5 Then
21490   tmrDown.Interval = 50
21500 End If

21510 VisibleRows = g.height \ g.RowHeight(1) - 1

21520 g.Visible = False

21530 s = ""
21540 For X = 0 To g.Cols - 1
21550   s = s & g.TextMatrix(n, X) & vbTab
21560 Next
21570 s = Left$(s, Len(s) - 1)

21580 g.RemoveItem n
21590 If n < g.Rows Then
21600   g.AddItem s, n + 1
21610   g.row = n + 1
21620 Else
21630   g.AddItem s
21640   g.row = g.Rows - 1
21650 End If

21660 For X = 0 To g.Cols - 1
21670   g.Col = X
21680   g.CellBackColor = vbYellow
21690 Next

21700 If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
21710   If g.row - VisibleRows + 1 > 0 Then
21720     g.TopRow = g.row - VisibleRows + 1
21730   End If
21740 End If

21750 g.Visible = True

21760 cmdSave.Visible = True

End Sub

Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

21770 If g.row = 1 Then Exit Sub

21780 FireCounter = FireCounter + 1
21790 If FireCounter > 5 Then
21800   tmrUp.Interval = 50
21810 End If

21820 n = g.row

21830 g.Visible = False

21840 s = ""
21850 For X = 0 To g.Cols - 1
21860   s = s & g.TextMatrix(n, X) & vbTab
21870 Next
21880 s = Left$(s, Len(s) - 1)

21890 g.RemoveItem n
21900 g.AddItem s, n - 1

21910 g.row = n - 1
21920 For X = 0 To g.Cols - 1
21930   g.Col = X
21940   g.CellBackColor = vbYellow
21950 Next

21960 If Not g.RowIsVisible(g.row) Then
21970   g.TopRow = g.row
21980 End If

21990 g.Visible = True

22000 cmdSave.Visible = True

End Sub

Public Property Get Discipline() As String

22010 Discipline = pDiscipline

End Property

Public Property Let Discipline(ByVal sNewValue As String)

      'Haem, Bio, Imm, Coag etc

22020 pDiscipline = UCase$(sNewValue)

22030 Select Case pDiscipline
        Case "HAEM": sDiscipline = "Haematology"
22040   Case "BIO": sDiscipline = "Biochemistry"
22050   Case "IMM": sDiscipline = "Immunology"
22060   Case "END": sDiscipline = "Endocrinology"
22070   Case "COAG": sDiscipline = "Coagulation"
22080 End Select

End Property
