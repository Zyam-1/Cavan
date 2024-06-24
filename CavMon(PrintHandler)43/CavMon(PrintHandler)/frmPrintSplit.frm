VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPrintSplit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Biochemistry Splits"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   14835
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13950
      Top             =   3390
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   10
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7140
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   9
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   8
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5820
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   7
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   6
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4500
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   5
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":320A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3180
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   4
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":3C0C
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   3
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":460E
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   2
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":5010
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   645
      Index           =   0
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":5A12
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   1
      Left            =   3300
      TabIndex        =   2
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Split 1"
      Enabled         =   0   'False
      Height          =   645
      Index           =   1
      Left            =   2280
      Picture         =   "frmPrintSplit.frx":6414
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   13770
      Picture         =   "frmPrintSplit.frx":6E16
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7170
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   2
      Left            =   5370
      TabIndex        =   4
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   3
      Left            =   7440
      TabIndex        =   5
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   4
      Left            =   9510
      TabIndex        =   6
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   7245
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   12779
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   5
      Left            =   11580
      TabIndex        =   8
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   6
      Left            =   3300
      TabIndex        =   14
      Top             =   4500
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   7
      Left            =   5400
      TabIndex        =   16
      Top             =   4500
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   8
      Left            =   7470
      TabIndex        =   17
      Top             =   4500
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   9
      Left            =   9540
      TabIndex        =   18
      Top             =   4500
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   3285
      Index           =   10
      Left            =   11610
      TabIndex        =   19
      Top             =   4500
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   5794
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type       |<Analyte       "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "seconds"
      Height          =   195
      Left            =   8760
      TabIndex        =   36
      Top             =   3930
      Width           =   600
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8250
      TabIndex        =   35
      Top             =   3870
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "This screen will automatically close in"
      Height          =   195
      Left            =   5550
      TabIndex        =   34
      Top             =   3930
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "General"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Index           =   10
      Left            =   11610
      TabIndex        =   23
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Index           =   9
      Left            =   9540
      TabIndex        =   22
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   7470
      TabIndex        =   21
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   20
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   3330
      TabIndex        =   15
      Top             =   4200
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   3300
      TabIndex        =   13
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   5370
      TabIndex        =   12
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   11
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   9510
      TabIndex        =   10
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   11580
      TabIndex        =   9
      Top             =   240
      Width           =   2040
   End
End
Attribute VB_Name = "frmPrintSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intLastGridUsed As Integer

Private CloseTimer As Integer

Private Const CLOSEDEFAULT As Integer = 20
Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub FillGrids()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleType As String
      Dim intY As Integer

10    On Error GoTo FillGrids_Error

20    For intY = 0 To 10
30      grdSplit(intY).Rows = 2
40      grdSplit(intY).AddItem ""
50      grdSplit(intY).RemoveItem 1
60    Next

70    sql = "SELECT DISTINCT LongName, SampleType, PrintPriority, " & _
            "COALESCE(PrintSplit, 0) SplitList " & _
            "FROM BioTestDefinitions " & _
            "ORDER BY PrintPriority"

80    Set tb = New Recordset
90    RecOpenClient 0, tb, sql
100   Do While Not tb.EOF
110     Select Case tb!SampleType & ""
          Case "S": SampleType = "Serum"
120       Case "C": SampleType = "CSF"
130       Case "B": SampleType = "Blood"
140       Case "U": SampleType = "Urine"
150     End Select
160     grdSplit(tb!SplitList).AddItem SampleType & vbTab & tb!LongName
170     tb.MoveNext
180   Loop

190   For intY = 0 To 10
200     If grdSplit(intY).Rows > 2 Then
210       grdSplit(intY).RemoveItem 1
220     End If
230   Next

240   Exit Sub

FillGrids_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmPrintSplit", "FillGrids", intEL, strES, sql

End Sub


Private Sub cmdMove_Click(Index As Integer)

      'Index is 'Move To'

      Dim sql As String
      Dim intFromIndex As Integer
      Dim strAnalyte As String
      Dim strSampleType As String
      Dim n As Integer

10    On Error GoTo cmdMove_Click_Error

20    If Index = 0 Then
30      intFromIndex = intLastGridUsed
40    Else
50      intFromIndex = 0
60    End If

70    With grdSplit(intFromIndex)
80      strAnalyte = .TextMatrix(.Row, 1)
90      strSampleType = Left$(.TextMatrix(.Row, 0), 1)
100   End With

110   sql = "UPDATE BioTestDefinitions " & _
            "SET PrintSplit = " & Index & " " & _
            "WHERE LongName = '" & strAnalyte & "' " & _
            "AND SampleType = '" & strSampleType & "'"
120   Cnxn(0).Execute sql

130   For n = 1 To 10
140     cmdMove(n).Enabled = False
150   Next

160   FillGrids

170   Exit Sub

cmdMove_Click_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmPrintSplit", "cmdMove_Click", intEL, strES, sql

End Sub


Private Sub cmdMove_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CloseTimer = CLOSEDEFAULT

End Sub


Private Sub Form_Load()

      Dim n As Integer

10    For n = 1 To 10
20      lblSplit(n).Caption = GetOptionSetting("PrintBioSplitName" & Format$(n), Format$(n))
30      cmdMove(n).Caption = lblSplit(n).Caption
40    Next

50    FillGrids

CloseTimer = CLOSEDEFAULT

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CloseTimer = CLOSEDEFAULT

End Sub


Private Sub grdSplit_Click(Index As Integer)

      Dim intY As Integer
      Dim intRowSave As Integer
      Dim intActive As Integer
      Dim n As Integer

10    On Error GoTo grdSplit_Click_Error

20    If grdSplit(Index).MouseRow = 0 Then Exit Sub
30    If grdSplit(Index).TextMatrix(1, 0) = "" Then Exit Sub

40    intLastGridUsed = Index

50    intRowSave = grdSplit(Index).Row

60    For intActive = 0 To 10
70      If intActive <> Index Then
80        With grdSplit(intActive)
90          For intY = 1 To .Rows - 1
100           .Row = intY
110           .Col = 0
120           .CellBackColor = &H80000018
130           .Col = 1
140           .CellBackColor = &H80000018
150         Next
160       End With
170     End If
180   Next

190   With grdSplit(Index)
200     For intY = 1 To .Rows - 1
210       .Row = intY
220       .Col = 0
230       .CellBackColor = &H80000018
240       .Col = 1
250       .CellBackColor = &H80000018
260     Next
270     .Row = intRowSave
280     .Col = 0
290     .CellBackColor = vbYellow
300     .Col = 1
310     .CellBackColor = vbYellow
320   End With

330   If Index = 0 Then
340     cmdMove(0).Enabled = False
350     For n = 1 To 10
360       cmdMove(n).Enabled = True
370     Next
380   Else
390     cmdMove(0).Enabled = True
400     For n = 1 To 10
410       cmdMove(n).Enabled = False
420     Next
430   End If

440   Exit Sub

grdSplit_Click_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "frmPrintSplit", "grdSplit_Click", intEL, strES

End Sub


Private Sub grdSplit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CloseTimer = CLOSEDEFAULT

End Sub


Private Sub lblSplit_Click(Index As Integer)

10    If lblSplit(Index).Caption = "" Then
20      lblSplit(Index).Caption = Format$(Index)
30    End If

40    lblSplit(Index).Caption = iBOX("Enter Title for Print Split " & lblSplit(Index).Caption)
50    cmdMove(Index).Caption = lblSplit(Index).Caption

60    SaveOptionSetting "PrintBioSplitName" & Format$(Index), lblSplit(Index).Caption

End Sub


Private Sub Timer1_Timer()

CloseTimer = CloseTimer - 1
If CloseTimer < 1 Then
  Unload Me
End If

lblClose.Caption = Format$(CloseTimer)

lblClose.BackColor = vbYellow
If CloseTimer < 11 Then
  lblClose.BackColor = vbRed
End If

End Sub


