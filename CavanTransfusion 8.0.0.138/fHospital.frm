VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fHospital 
   Caption         =   "NetAcquire - Hospitals"
   ClientHeight    =   7050
   ClientLeft      =   1830
   ClientTop       =   1110
   ClientWidth     =   7245
   Icon            =   "fHospital.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7245
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   6150
      Picture         =   "fHospital.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1350
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   6150
      Picture         =   "fHospital.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6150
      Picture         =   "fHospital.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3630
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6150
      Picture         =   "fHospital.frx":19E0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4470
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6150
      Picture         =   "fHospital.frx":1E22
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5970
      Width           =   795
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add New Hospital"
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   5925
      Begin VB.TextBox tCode 
         Height          =   285
         Left            =   810
         MaxLength       =   5
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox tText 
         Height          =   285
         Left            =   810
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3645
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   4860
         TabIndex        =   1
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   390
         TabIndex        =   4
         Top             =   630
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5445
      Left            =   90
      TabIndex        =   6
      Top             =   1350
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   9604
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
      AllowUserResizing=   1
      FormatString    =   "<Code       |<Text                                                                     "
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   210
      TabIndex        =   12
      Top             =   6840
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "fHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from Lists where " & _
            "ListType = 'HO' " & _
            "order by ListOrder"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    Do While Not tb.EOF
90      s = tb!code & vbTab & tb!Text & ""
100     g.AddItem s
110     tb.MoveNext
120   Loop

130   If g.Rows > 2 Then
140     g.RemoveItem 1
150   End If

160   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fHospital", "FillG", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

10    tCode = Trim$(UCase$(tCode))
20    tText = Trim$(tText)

30    If tCode = "" Then
40      Exit Sub
50    End If

60    If tText = "" Then Exit Sub

70    g.AddItem tCode & vbTab & tText

80    tCode = ""
90    tText = ""

100   cmdSave.Enabled = True

End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bMoveDown_Click()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

10    If g.Row = g.Rows - 1 Then Exit Sub
20    n = g.Row

30    s = ""
40    For X = 0 To g.Cols - 1
50      s = s & g.TextMatrix(n, X) & vbTab
60    Next
70    s = Left$(s, Len(s) - 1)

80    g.RemoveItem n
90    If n < g.Rows Then
100     g.AddItem s, n + 1
110     g.Row = n + 1
120   Else
130     g.AddItem s
140     g.Row = g.Rows - 1
150   End If

160   For X = 0 To g.Cols - 1
170     g.Col = X
180     g.CellBackColor = vbYellow
190   Next

200   cmdSave.Enabled = True

End Sub


Private Sub bMoveUp_Click()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

10    If g.Row = 1 Then Exit Sub

20    n = g.Row

30    s = ""
40    For X = 0 To g.Cols - 1
50      s = s & g.TextMatrix(n, X) & vbTab
60    Next
70    s = Left$(s, Len(s) - 1)

80    g.RemoveItem n
90    g.AddItem s, n - 1

100   g.Row = n - 1
110   For X = 0 To g.Cols - 1
120     g.Col = X
130     g.CellBackColor = vbYellow
140   Next

150   cmdSave.Enabled = True

End Sub


Private Sub bprint_Click()

      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Print

40    Printer.Print "List of Hospitals"

50    g.Col = 0
60    g.Row = 1
70    g.ColSel = g.Cols - 1
80    g.RowSel = g.Rows - 1

90    Printer.Print g.Clip

100   Printer.EndDoc

110   For Each Px In Printers
120     If Px.DeviceName = OriginalPrinter Then
130       Set Printer = Px
140       Exit For
150     End If
160   Next

End Sub


Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim Y As Integer

10    On Error GoTo cmdSave_Click_Error

20    For Y = 1 To g.Rows - 1
30      sql = "Select * from Lists where " & _
              "ListType = 'HO' " & _
              "and Code = '" & g.TextMatrix(Y, 0) & "'"
40      Set tb = New Recordset
50      RecOpenServer 0, tb, sql
60      If tb.EOF Then
70        tb.AddNew
80      End If
90      tb!code = g.TextMatrix(Y, 0)
100     tb!ListType = "HO"
110     tb!Text = g.TextMatrix(Y, 1)
120     tb!ListOrder = Y
130     tb!InUse = 1
140     tb.Update
  
150   Next

160   FillG

170   tCode = ""
180   tText = ""
190   tCode.SetFocus
200   bMoveUp.Enabled = False
210   bMoveDown.Enabled = False
220   cmdSave.Enabled = False

230   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "fHospital", "cmdSave_Click", intEL, strES, sql


End Sub




Private Sub Form_Load()

10    g.Font.Bold = True

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
20        FillG
      '**************************************

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Enabled Then
30      Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70        Exit Sub
80      End If
90    End If

End Sub


Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer

10    ySave = g.Row

20    g.Visible = False
30    g.Col = 0
40    For Y = 1 To g.Rows - 1
50      g.Row = Y
60      If g.CellBackColor = vbYellow Then
70        For X = 0 To g.Cols - 1
80          g.Col = X
90          g.CellBackColor = 0
100       Next
110       Exit For
120     End If
130   Next
140   g.Row = ySave
150   g.Visible = True

160   If g.MouseRow = 0 Then
170     If SortOrder Then
180       g.Sort = flexSortGenericAscending
190     Else
200       g.Sort = flexSortGenericDescending
210     End If
220     SortOrder = Not SortOrder
230     Exit Sub
240   End If

250   For X = 0 To g.Cols - 1
260     g.Col = X
270     g.CellBackColor = vbYellow
280   Next

290   bMoveUp.Enabled = True
300   bMoveDown.Enabled = True

End Sub


