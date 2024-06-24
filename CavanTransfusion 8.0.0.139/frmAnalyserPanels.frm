VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmAnalyserPanels 
   Caption         =   "ID Panels"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   10800
      Picture         =   "frmAnalyserPanels.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox txtPanelComment 
      Enabled         =   0   'False
      Height          =   3450
      Left            =   4365
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2655
      Width           =   5475
   End
   Begin VB.Frame Frame 
      Height          =   1620
      Left            =   5595
      TabIndex        =   5
      Top             =   690
      Width           =   6105
      Begin VB.Label lblAnalyserValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1530
         TabIndex        =   15
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label lblAnalyser 
         Caption         =   "Analyser:"
         Height          =   240
         Left            =   840
         TabIndex        =   14
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label lblCassetteLotNoValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4485
         TabIndex        =   13
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label lblCassetteLotNo 
         Caption         =   "Cassette Lot No:"
         Height          =   240
         Left            =   3240
         TabIndex        =   12
         Top             =   1005
         Width           =   1230
      End
      Begin VB.Label lblCassetteExpiryValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4485
         TabIndex        =   11
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label lblCassetteExpiry 
         Caption         =   "Cassette Expiry:"
         Height          =   240
         Left            =   3285
         TabIndex        =   10
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label lblCassetteIdValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4485
         TabIndex        =   9
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblCassetteId 
         Caption         =   "Cassette Id:"
         Height          =   240
         Left            =   3570
         TabIndex        =   8
         Top             =   255
         Width           =   1230
      End
      Begin VB.Label lblReagentExpiryValue 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1530
         TabIndex        =   7
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblReagentExpiry 
         Caption         =   "Reagent Expiry:"
         Height          =   240
         Left            =   315
         TabIndex        =   6
         Top             =   255
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   10800
      Picture         =   "frmAnalyserPanels.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5310
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid grdPanels 
      Height          =   1605
      Left            =   240
      TabIndex        =   1
      Top             =   750
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   2831
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
      AllowUserResizing=   1
      FormatString    =   "<Panel Name                                            |<Lot Number                  "
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   6315
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdPanelResults 
      Height          =   3480
      Left            =   240
      TabIndex        =   16
      Top             =   2655
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   6138
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
      AllowUserResizing=   1
      FormatString    =   "<Well Name                         |<Result Value     "
   End
   Begin VB.Label lblSampleId 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1485
      TabIndex        =   4
      Top             =   165
      Width           =   1485
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   3
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "frmAnalyserPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String

Public Property Let SampleID(ByVal sNewValue As String)

10    mSampleID = sNewValue

End Property


Private Sub cmdCancel_Click()
10    Unload Me
End Sub


Private Sub cmdSave_Click()
Dim Y As Integer
Dim strPanelName As String
Dim strLotNumber As String
Dim Sql As String

On Error GoTo cmdSave_Click_Error

strPanelName = ""
strLotNumber = ""
grdPanels.col = 0
50    For Y = 1 To grdPanels.Rows - 1
60      grdPanels.row = Y
70      If grdPanels.CellBackColor = vbYellow Then
            strPanelName = grdPanels.TextMatrix(Y, 0)
            strLotNumber = grdPanels.TextMatrix(Y, 1)
120       Exit For
130     End If
140   Next

If strPanelName <> "" And strLotNumber <> "" Then
     Sql = "UPDATE AnalyserIDPanels SET " & _
              "PanelComments = '" & Trim$(txtPanelComment) & "' WHERE " & _
              "TypeofCassette = '" & strPanelName & "' and ReagentLotNo = '" & strLotNumber & "'"
     CnxnBB(0).Execute Sql
End If

Exit Sub

cmdSave_Click_Error:

 Dim strES As String
 Dim intEL As Integer

 intEL = Erl
 strES = Err.Description
 LogError "frmAnalyserPanels", "cmdSave_Click", intEL, strES, Sql

End Sub

Private Sub Form_Load()
10    lblSampleId = mSampleID

20    Fill_TypeofCassette

End Sub


Private Sub Fill_TypeofCassette()

      Dim Sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo Fill_TypeofCassette_Error

20    grdPanels.Rows = 2
30    grdPanels.AddItem ""
40    grdPanels.RemoveItem 1

50    Sql = "SELECT DISTINCT TypeofCassette, ReagentLotNo From AnalyserIDPanels WHERE (SampleID = '" & lblSampleId & "') "

60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, Sql
80    Do While Not tb.EOF
90      s = tb!TypeofCassette & "" & vbTab & _
            tb!ReagentLotNo & ""
100     grdPanels.AddItem s
110     tb.MoveNext
120   Loop

130   If grdPanels.Rows > 2 Then
140     grdPanels.RemoveItem 1
150   End If

160   Exit Sub

Fill_TypeofCassette_Error:

       Dim strES As String
       Dim intEL As Integer

170    intEL = Erl
180    strES = Err.Description
190    LogError "frmAnalyserPanels", "Fill_TypeofCassette", intEL, strES, Sql

End Sub

Private Sub grdPanels_Click()
Dim Y As Integer
Dim X As Integer
Dim ySave As Integer
Dim Sql As String
Dim tb As Recordset
Dim s As String

10    On Error GoTo grdPanels_Click_Error

20    ySave = grdPanels.row

30    grdPanels.Visible = False
40    grdPanels.col = 0
50    For Y = 1 To grdPanels.Rows - 1
60      grdPanels.row = Y
70      If grdPanels.CellBackColor = vbYellow Then
80        For X = 0 To grdPanels.Cols - 1
90          grdPanels.col = X
100         grdPanels.CellBackColor = 0
110       Next
120       Exit For
130     End If
140   Next
150   grdPanels.row = ySave
160   grdPanels.Visible = True

170   For X = 0 To grdPanels.Cols - 1
180     grdPanels.col = X
190     grdPanels.CellBackColor = vbYellow
200   Next


210   grdPanelResults.Rows = 2
220   grdPanelResults.AddItem ""
230   grdPanelResults.RemoveItem 1

240   Sql = "SELECT * From AnalyserIDPanels WHERE (SampleID = '" & lblSampleId & "') and TypeofCassette = '" & grdPanels.TextMatrix(ySave, 0) & "' and ReagentLotNo = '" & grdPanels.TextMatrix(ySave, 1) & "' order by testorder"

250   Set tb = New Recordset
260   RecOpenServerBB 0, tb, Sql

270   If Not tb.EOF Then
280       lblReagentExpiryValue = tb!ReagentExpiry
290       lblCassetteIdValue = tb!CassetteIDNumber & ""
300       lblCassetteExpiryValue = tb!CassetteExpirationDate & ""
310       lblCassetteLotNoValue = tb!CassetteLotNumber & ""
    
320       lblAnalyserValue = tb!Analyser & ""
325       txtPanelComment = tb!PanelComments & ""
330   End If

340   Do While Not tb.EOF
350     s = tb!ResultWellName & "" & vbTab & _
      tb!TestResult & ""
360     grdPanelResults.AddItem s
370     tb.MoveNext
380   Loop

390   If grdPanelResults.Rows > 2 Then
400     grdPanelResults.RemoveItem 1
410   End If

415   txtPanelComment.Enabled = True

420   Exit Sub

grdPanels_Click_Error:

 Dim strES As String
 Dim intEL As Integer

430    intEL = Erl
440    strES = Err.Description
450    LogError "frmAnalyserPanels", "grdPanels_Click", intEL, strES, Sql

End Sub

Private Sub txtPanelComment_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True

End Sub
