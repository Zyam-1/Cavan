VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAntigens 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Antigen List"
   ClientHeight    =   8415
   ClientLeft      =   585
   ClientTop       =   600
   ClientWidth     =   7875
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmAntigens.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8415
   ScaleWidth      =   7875
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   6420
      Picture         =   "frmAntigens.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1830
      Width           =   1245
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6900
      Top             =   4620
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6900
      Top             =   5190
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   735
      Left            =   6480
      Picture         =   "frmAntigens.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move &Down"
      Enabled         =   0   'False
      Height          =   885
      Left            =   6420
      Picture         =   "frmAntigens.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   765
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move &Up"
      Enabled         =   0   'False
      Height          =   885
      Left            =   6420
      Picture         =   "frmAntigens.frx":1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   450
      TabIndex        =   5
      Top             =   150
      Width           =   5865
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Default         =   -1  'True
         Height          =   855
         Left            =   4920
         Picture         =   "frmAntigens.frx":1AC2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   765
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   210
         TabIndex        =   1
         Top             =   1050
         Width           =   4545
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H80000014&
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   0
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   210
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6195
      Left            =   450
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1830
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   10927
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
      FormatString    =   "<Code     |<Description                                                            "
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   6450
      Picture         =   "frmAntigens.frx":1F04
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7290
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   30
      TabIndex        =   11
      Top             =   8130
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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
      Height          =   345
      Left            =   6420
      TabIndex        =   13
      Top             =   2580
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "frmAntigens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

10    With g
20      If .Row = .Rows - 1 Then Exit Sub
30      n = .Row
  
40      FireCounter = FireCounter + 1
50      If FireCounter > 5 Then
60        tmrDown.Interval = 100
70      End If
  
80      VisibleRows = .Height \ .RowHeight(1) - 1
  
90      .Visible = False
  
100     s = ""
110     For X = 0 To .Cols - 1
120       s = s & .TextMatrix(n, X) & vbTab
130     Next
140     s = Left$(s, Len(s) - 1)
  
150     .RemoveItem n
160     If n < .Rows Then
170       .AddItem s, n + 1
180       .Row = n + 1
190     Else
200       .AddItem s
210       .Row = .Rows - 1
220     End If
  
230     For X = 0 To .Cols - 1
240       .Col = X
250       .CellBackColor = vbYellow
260     Next
  
270     If Not .RowIsVisible(.Row) Or .Row = .Rows - 1 Then
280       If .Row - VisibleRows + 1 > 0 Then
290         .TopRow = .Row - VisibleRows + 1
300       End If
310     End If
  
320     .Visible = True
330   End With

340   cmdSave.Visible = True

End Sub
Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

10    With g
20      If .Row = 1 Then Exit Sub
  
30      FireCounter = FireCounter + 1
40      If FireCounter > 5 Then
50        tmrUp.Interval = 100
60      End If
  
70      n = .Row
  
80      .Visible = False
  
90      s = ""
100     For X = 0 To .Cols - 1
110       s = s & .TextMatrix(n, X) & vbTab
120     Next
130     s = Left$(s, Len(s) - 1)
  
140     .RemoveItem n
150     .AddItem s, n - 1
  
160     .Row = n - 1
170     For X = 0 To .Cols - 1
180       .Col = X
190       .CellBackColor = vbYellow
200     Next
  
210     If Not .RowIsVisible(.Row) Then
220       .TopRow = .Row
230     End If
  
240     .Visible = True
  
250     cmdSave.Visible = True
260   End With

End Sub





Private Sub SaveDetails()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo SaveDetails_Error

20    For n = 1 To g.Rows - 1
30        g.Row = n
40        If g.CellBackColor = vbRed Then
50            sql = "Delete from Lists where " & _
                  "ListType = 'PC' " & _
                  "and Code = '" & AddTicks(g.TextMatrix(g.Row, 0)) & "'"
60            CnxnBB(0).Execute sql
70        Else
    
80            sql = "Select * from Lists where " & _
                    "ListType = 'PC' and Code = '" & AddTicks(g.TextMatrix(n, 0)) & "'"
90            Set tb = New Recordset
100           RecOpenServerBB 0, tb, sql
110           If tb.EOF Then tb.AddNew
120           tb!ListType = "PC"
130           tb!code = g.TextMatrix(n, 0)
140           tb!Text = g.TextMatrix(n, 1)
150           tb!ListOrder = n
160           tb!InUse = 1
170           tb.Update
180       End If
190   Next

200   FillG
210   txtText = ""
220   txtCode = ""
230   cmdSave.Visible = False

240   Exit Sub

SaveDetails_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmAntigens", "SaveDetails", intEL, strES, sql


End Sub

Private Sub cmdadd_Click()

10    txtCode = Trim$(txtCode)
20    txtText = Trim$(txtText)
30    If txtText = "" Then
40      iMsg "Enter Description!", vbExclamation
50      If TimedOut Then Unload Me: Exit Sub
60      Exit Sub
70    End If
80    If txtCode = "" Then
90      iMsg "Enter Code!", vbExclamation
100     If TimedOut Then Unload Me: Exit Sub
110     Exit Sub
120   End If
  
      'Change: If fitem already exists in grid then edit it or add new
      Dim boolItemFound As Boolean
130   boolItemFound = False
      Dim X As Integer
140   For X = 1 To g.Rows - 1
150       If txtCode = g.TextMatrix(X, 0) Then
160           boolItemFound = True 'item found
170           Exit For
    
180       End If
190   Next X
200   If boolItemFound Then
210       Call MarkGridRow(g, g.Row, vbYellow, vbBlack, False, True, False)
220       g.TextMatrix(g.Row, 0) = txtCode
230       g.TextMatrix(g.Row, 1) = txtText
    
240   Else
250       g.AddItem txtCode & vbTab & txtText
260   End If
  


270   txtText = ""
280   txtCode = ""

290   txtCode.SetFocus
300   cmdSave.Visible = True

End Sub



Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub FillG()

      Dim s As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from Lists where " & _
            "ListType = 'PC' Order by ListOrder"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF
90      s = tb!code & vbTab & _
            tb!Text & ""
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
190   LogError "frmAntigens", "FillG", intEL, strES, sql


End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FireDown

20    tmrDown.Interval = 250
30    FireCounter = 0

40    tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FireUp

20    tmrUp.Interval = 250
30    FireCounter = 0

40    tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    tmrUp.Enabled = False

End Sub


Private Sub cmdSave_Click()

10    SaveDetails

End Sub


Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Antigen List" & vbCr
20    strHeading = strHeading & " " & vbCr


30    ExportFlexGrid g, Me, strHeading
End Sub

Private Sub Form_Load()

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        FillG
      '**************************************

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Visible Then
30      Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70      End If
80    End If

End Sub


Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer

10    If g.MouseRow = 0 Then
20      If SortOrder Then
30        g.Sort = flexSortGenericAscending
40      Else
50        g.Sort = flexSortGenericDescending
60      End If
70      SortOrder = Not SortOrder
80      cmdMoveUp.Enabled = False
90      cmdMoveDown.Enabled = False
100     Exit Sub
110   End If

120   If g.Col = 0 Then
130     g.Enabled = False
140     If iMsg("Edit/Remove this line?", vbQuestion + vbYesNo) = vbYes Then
150       txtCode = g.TextMatrix(g.Row, 0)
160       txtText = g.TextMatrix(g.Row, 1)
170       Call MarkGridRow(g, g.Row, vbRed, vbYellow, True, True, True)
180       cmdSave.Visible = True
190     End If
200     g.Enabled = True
210     Exit Sub
220   End If

230   ySave = g.Row

240   g.Visible = False
250   g.Col = 0
260   For Y = 1 To g.Rows - 1
270     g.Row = Y
280     If g.CellBackColor = vbYellow Then
290       For X = 0 To g.Cols - 1
300   g.Col = X
310   g.CellBackColor = 0
320       Next
330       Exit For
340     End If
350   Next
360   g.Row = ySave
370   g.Visible = True


380   For X = 0 To g.Cols - 1
390     g.Col = X
400     g.CellBackColor = vbYellow
410   Next



420   cmdMoveUp.Enabled = True
430   cmdMoveDown.Enabled = True

End Sub

Private Sub tmrDown_Timer()

10    FireDown

End Sub

Private Sub tmrUp_Timer()

10    FireUp

End Sub


