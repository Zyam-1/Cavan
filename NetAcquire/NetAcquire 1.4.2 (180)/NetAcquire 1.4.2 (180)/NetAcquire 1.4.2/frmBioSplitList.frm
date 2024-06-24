VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBioSplitList 
   Caption         =   "NetAcquire - Biochemistry Splits"
   ClientHeight    =   7725
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   15780
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move to Split 6"
      Enabled         =   0   'False
      Height          =   855
      Index           =   6
      Left            =   2280
      Picture         =   "frmBioSplitList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4890
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move to Split 5"
      Enabled         =   0   'False
      Height          =   855
      Index           =   5
      Left            =   2280
      Picture         =   "frmBioSplitList.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4020
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move to Split 4"
      Enabled         =   0   'False
      Height          =   855
      Index           =   4
      Left            =   2280
      Picture         =   "frmBioSplitList.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3150
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move to Split 3"
      Enabled         =   0   'False
      Height          =   855
      Index           =   3
      Left            =   2280
      Picture         =   "frmBioSplitList.frx":3186
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Remove Item"
      Enabled         =   0   'False
      Height          =   825
      Index           =   0
      Left            =   2280
      Picture         =   "frmBioSplitList.frx":4208
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5940
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   7065
      Index           =   1
      Left            =   3300
      TabIndex        =   3
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   12462
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
      Caption         =   "Move to Split 2"
      Enabled         =   0   'False
      Height          =   855
      Index           =   2
      Left            =   2280
      Picture         =   "frmBioSplitList.frx":528A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1410
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move to Split 1"
      Enabled         =   0   'False
      Height          =   855
      Index           =   1
      Left            =   2280
      Picture         =   "frmBioSplitList.frx":630C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   2280
      Picture         =   "frmBioSplitList.frx":738E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6930
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   7065
      Index           =   2
      Left            =   5370
      TabIndex        =   5
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   12462
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
      Height          =   7065
      Index           =   3
      Left            =   7440
      TabIndex        =   6
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   12462
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
      Height          =   7065
      Index           =   4
      Left            =   9510
      TabIndex        =   7
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   12462
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
      Height          =   7425
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   150
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   13097
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
      Height          =   7065
      Index           =   5
      Left            =   11580
      TabIndex        =   12
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   12462
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
      Height          =   7065
      Index           =   6
      Left            =   13620
      TabIndex        =   18
      Top             =   540
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   12462
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
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Split 6"
      Height          =   255
      Index           =   6
      Left            =   13650
      TabIndex        =   19
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Split 1"
      Height          =   255
      Index           =   1
      Left            =   3300
      TabIndex        =   17
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Split 2"
      Height          =   255
      Index           =   2
      Left            =   5370
      TabIndex        =   16
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Split 3"
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   15
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Split 4"
      Height          =   255
      Index           =   4
      Left            =   9510
      TabIndex        =   14
      Top             =   240
      Width           =   2040
   End
   Begin VB.Label lblSplit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Split 5"
      Height          =   255
      Index           =   5
      Left            =   11580
      TabIndex        =   13
      Top             =   240
      Width           =   2040
   End
End
Attribute VB_Name = "frmBioSplitList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intLastGridUsed As Integer

Private Sub cmdCancel_Click()

10850     Unload Me

End Sub


Private Sub FillGrids()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleType As String
          Dim intY As Integer

10860     On Error GoTo FillGrids_Error

10870     For intY = 0 To 6
10880         grdSplit(intY).Rows = 2
10890         grdSplit(intY).AddItem ""
10900         grdSplit(intY).RemoveItem 1
10910     Next

10920     sql = "SELECT DISTINCT LongName, SampleType, PrintPriority, " & _
              "COALESCE(SplitList, 0) SplitList " & _
              "FROM BioTestDefinitions " & _
              "ORDER BY PrintPriority"

10930     Set tb = New Recordset
10940     RecOpenClient 0, tb, sql
10950     Do While Not tb.EOF
10960         Select Case tb!SampleType & ""
                  Case "S": SampleType = "Serum"
10970             Case "C": SampleType = "CSF"
10980             Case "B": SampleType = "Blood"
10990             Case "U": SampleType = "Urine"
11000         End Select
11010         grdSplit(tb!SplitList).AddItem SampleType & vbTab & tb!LongName
11020         tb.MoveNext
11030     Loop

11040     For intY = 0 To 6
11050         If grdSplit(intY).Rows > 2 Then
11060             grdSplit(intY).RemoveItem 1
11070         End If
11080     Next

11090     Exit Sub

FillGrids_Error:

          Dim strES As String
          Dim intEL As Integer

11100     intEL = Erl
11110     strES = Err.Description
11120     LogError "frmBioSplitList", "FillGrids", intEL, strES, sql

End Sub


Private Sub cmdMove_Click(Index As Integer)

          'Index is 'Move To'

          Dim sql As String
          Dim intFromIndex As Integer
          Dim strAnalyte As String
          Dim strSampleType As String
          Dim n As Integer

11130     On Error GoTo cmdMove_Click_Error

11140     If Index = 0 Then
11150         intFromIndex = intLastGridUsed
11160     Else
11170         intFromIndex = 0
11180     End If

11190     With grdSplit(intFromIndex)
11200         strAnalyte = .TextMatrix(.row, 1)
11210         strSampleType = Left$(.TextMatrix(.row, 0), 1)
11220     End With

11230     sql = "UPDATE BioTestDefinitions " & _
              "SET Splitlist = " & Index & " " & _
              "WHERE LongName = '" & strAnalyte & "' " & _
              "AND SampleType = '" & strSampleType & "'"
11240     Cnxn(0).Execute sql

11250     For n = 1 To 6
11260         cmdMove(n).Enabled = False
11270     Next

11280     FillGrids

11290     Exit Sub

cmdMove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

11300     intEL = Erl
11310     strES = Err.Description
11320     LogError "frmBioSplitList", "cmdMove_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

          Dim n As Integer

11330     For n = 1 To 6
11340         lblSplit(n).Caption = GetOptionSetting("BioSplitName" & Format$(n), "Split " & Format$(n))
11350         cmdMove(n).Caption = "Move to " & lblSplit(n).Caption
11360     Next

11370     FillGrids

End Sub

Private Sub grdSplit_Click(Index As Integer)

          Dim intY As Integer
          Dim intRowSave As Integer
          Dim intActive As Integer
          Dim n As Integer

11380     If grdSplit(Index).MouseRow = 0 Then Exit Sub
11390     If grdSplit(Index).TextMatrix(1, 0) = "" Then Exit Sub

11400     intLastGridUsed = Index

11410     intRowSave = grdSplit(Index).row

11420     For intActive = 0 To 6
11430         If intActive <> Index Then
11440             With grdSplit(intActive)
11450                 For intY = 1 To .Rows - 1
11460                     .row = intY
11470                     .Col = 0
11480                     .CellBackColor = &H80000018
11490                     .Col = 1
11500                     .CellBackColor = &H80000018
11510                 Next
11520             End With
11530         End If
11540     Next

11550     With grdSplit(Index)
11560         For intY = 1 To .Rows - 1
11570             .row = intY
11580             .Col = 0
11590             .CellBackColor = &H80000018
11600             .Col = 1
11610             .CellBackColor = &H80000018
11620         Next
11630         .row = intRowSave
11640         .Col = 0
11650         .CellBackColor = vbYellow
11660         .Col = 1
11670         .CellBackColor = vbYellow
11680     End With

11690     If Index = 0 Then
11700         cmdMove(0).Enabled = False
11710         For n = 1 To 6
11720             cmdMove(n).Enabled = True
11730         Next
11740     Else
11750         cmdMove(0).Enabled = True
11760         For n = 1 To 6
11770             cmdMove(n).Enabled = False
11780         Next
11790     End If

End Sub


Private Sub lblSplit_Click(Index As Integer)

11800     If lblSplit(Index).Caption = "" Then
11810         lblSplit(Index).Caption = "Split " & Format$(Index)
11820     End If

11830     lblSplit(Index).Caption = iBOX("Enter Title for " & lblSplit(Index).Caption)
11840     cmdMove(Index).Caption = "Move to " & lblSplit(Index).Caption

11850     SaveOptionSetting "BioSplitName" & Format$(Index), lblSplit(Index).Caption

End Sub


