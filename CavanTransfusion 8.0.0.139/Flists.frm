VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form flists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listings"
   ClientHeight    =   8325
   ClientLeft      =   1050
   ClientTop       =   450
   ClientWidth     =   11805
   ForeColor       =   &H80000008&
   Icon            =   "Flists.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8325
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1155
      Left            =   10395
      Picture         =   "Flists.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3870
      Width           =   1245
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9735
      Picture         =   "Flists.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6870
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9735
      Picture         =   "Flists.frx":1BD6
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7440
      Width           =   465
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9675
      Top             =   4560
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9675
      Top             =   5040
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1155
      Left            =   10395
      Picture         =   "Flists.frx":2018
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7050
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   1155
      Left            =   10395
      Picture         =   "Flists.frx":2EE2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5730
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame fraList 
      Caption         =   "List"
      Height          =   2415
      Left            =   9735
      TabIndex        =   8
      Top             =   90
      Width           =   1905
      Begin VB.OptionButton oList 
         Caption         =   "AutoVue Tests"
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   24
         Top             =   2040
         Width           =   1425
      End
      Begin VB.OptionButton oList 
         Caption         =   "Reason for Call"
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   23
         Top             =   1770
         Width           =   1425
      End
      Begin VB.OptionButton oList 
         Caption         =   "Comments"
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   19
         Top             =   1500
         Width           =   1635
      End
      Begin VB.OptionButton oList 
         Caption         =   "Batched Products"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   13
         Top             =   1260
         Width           =   1695
      End
      Begin VB.OptionButton oList 
         Caption         =   "Reagents"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   12
         Top             =   1020
         Width           =   1035
      End
      Begin VB.OptionButton oList 
         Caption         =   "Special Products"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   11
         Top             =   780
         Width           =   1575
      End
      Begin VB.OptionButton oList 
         Caption         =   "Surgical Procedures"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   540
         Width           =   1815
      End
      Begin VB.OptionButton oList 
         Caption         =   "Clinical Conditions"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   300
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New"
      Height          =   1035
      Left            =   150
      TabIndex        =   2
      Top             =   90
      Width           =   8415
      Begin VB.CommandButton badd 
         Caption         =   "Add"
         Height          =   825
         Left            =   7500
         Picture         =   "Flists.frx":3DAC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   150
         Width           =   765
      End
      Begin VB.TextBox tinputcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         MaxLength       =   3
         TabIndex        =   4
         Top             =   540
         Width           =   675
      End
      Begin VB.TextBox tInputText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   930
         MaxLength       =   50
         TabIndex        =   3
         Top             =   540
         Width           =   6495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   1020
         TabIndex        =   6
         Top             =   330
         Width           =   315
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6765
      Left            =   150
      TabIndex        =   1
      Top             =   1170
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   11933
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
      GridLines       =   2
      FormatString    =   $"Flists.frx":572E
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1155
      Left            =   10395
      Picture         =   "Flists.frx":57C3
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2610
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pbSave 
      Height          =   225
      Left            =   10395
      TabIndex        =   14
      Top             =   5370
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   150
      TabIndex        =   20
      Top             =   7965
      Width           =   9540
      _ExtentX        =   16828
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
      Height          =   345
      Left            =   10395
      TabIndex        =   22
      Top             =   5010
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "flists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mListName As String


Private FireCounter As Integer

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

10    With g
20      If .row = .Rows - 1 Then Exit Sub
30      n = .row
  
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
180       .row = n + 1
190     Else
200       .AddItem s
210       .row = .Rows - 1
220     End If
  
230     For X = 0 To .Cols - 1
240       .col = X
250       .CellBackColor = vbYellow
260     Next
  
270     If Not .RowIsVisible(.row) Or .row = .Rows - 1 Then
280       If .row - VisibleRows + 1 > 0 Then
290         .TopRow = .row - VisibleRows + 1
300       End If
310     End If
  
320     .Visible = True
330   End With

340   cmdSave.Visible = True
350   fraList.Enabled = False

End Sub
Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

10    With g
20      If .row = 1 Then Exit Sub
  
30      FireCounter = FireCounter + 1
40      If FireCounter > 5 Then
50        tmrUp.Interval = 100
60      End If
  
70      n = .row
  
80      .Visible = False
  
90      s = ""
100     For X = 0 To .Cols - 1
110       s = s & .TextMatrix(n, X) & vbTab
120     Next
130     s = Left$(s, Len(s) - 1)
  
140     .RemoveItem n
150     .AddItem s, n - 1
  
160     .row = n - 1
170     For X = 0 To .Cols - 1
180       .col = X
190       .CellBackColor = vbYellow
200     Next
  
210     If Not .RowIsVisible(.row) Then
220       .TopRow = .row
230     End If
  
240     .Visible = True
  
250     cmdSave.Visible = True
260     fraList.Enabled = False
270   End With

End Sub



Public Property Let ListName(ByVal Ln As String)

10    mListName = Ln

End Property

Private Sub SaveDetails()

      Dim tb As Recordset
      Dim sql As String
      Dim y As Integer

10    On Error GoTo SaveDetails_Error

20    For y = 1 To g.Rows - 1
          'Change: if row is marked as red then delete it
30        g.row = y
40        If g.CellBackColor = vbRed Then
50            sql = "Delete from Lists where " & _
                  "ListType = '" & mListName & "' " & _
                  "and Code = '" & g.TextMatrix(g.row, 0) & "'"
60            CnxnBB(0).Execute sql
70        Else
80            If Trim$(g.TextMatrix(y, 0)) <> "" Then
90              sql = "Select * from Lists where " & _
                      "ListType = '" & mListName & "' " & _
                      "and Code = '" & g.TextMatrix(y, 0) & "'"
100             Set tb = New Recordset
110             RecOpenServerBB 0, tb, sql
120             If tb.EOF Then tb.AddNew
130             tb!code = g.TextMatrix(y, 0)
140             tb!Text = g.TextMatrix(y, 1)
150             tb!ListType = mListName
160             tb!ListOrder = y
170             tb!InUse = 1
180             tb.Update
190           End If
200       End If
210   Next

220   cmdSave.Visible = False
230   fraList.Enabled = True
240   cmdMoveUp.Enabled = False
250   cmdMoveDown.Enabled = False
260   tinputcode = ""
270   tInputText = ""
280   Exit Sub

SaveDetails_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "flists", "SaveDetails", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

10    tinputcode = UCase$(Trim$(tinputcode))
20    tInputText = Trim$(tInputText)

30    If tinputcode = "" Then
40      tinputcode.SetFocus
50      Exit Sub
60    End If

70    If tInputText = "" Then
80      tInputText.SetFocus
90      Exit Sub
100   End If

      'Change: If item already exists in grid then edit it or add new
      Dim boolItemFound As Boolean
110   boolItemFound = False
      Dim X As Integer
120   For X = 1 To g.Rows - 1
130       If tinputcode = g.TextMatrix(X, 0) Then
140           boolItemFound = True 'item found
150           Exit For
    
160       End If
170   Next X
180   If boolItemFound Then
190       Call MarkGridRow(g, g.row, vbYellow, vbBlack, False, True, False)
200       g.TextMatrix(g.row, 1) = tInputText
210   Else
220       g.AddItem tinputcode & vbTab & tInputText
230   End If

240   tinputcode = ""
250   tInputText = ""
260   tinputcode.SetFocus

270   cmdSave.Visible = True
280   fraList.Enabled = False

End Sub

Private Sub cmdCancel_Click()

10    If cmdSave.Visible Then
20      Answer = iMsg("Save changes?", vbQuestion + vbYesNo)
30      If TimedOut Then Unload Me: Exit Sub
40      If Answer = vbYes Then
50        SaveDetails
60      End If
70    End If

80    Unload Me

End Sub





Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

10    FireDown

20    tmrDown.Interval = 250
30    FireCounter = 0

40    tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

10    tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

10    FireUp

20    tmrUp.Interval = 250
30    FireCounter = 0

40    tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

10    tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

      Dim s As String
      Dim n As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    On Error GoTo cmdPrint_Click_Error

20    OriginalPrinter = Printer.DeviceName
30    If Not SetFormPrinter() Then Exit Sub
40    Printer.Orientation = vbPRORPortrait

50    Printer.Print
60    Printer.Font.Name = "Courier New"
70    Printer.Font.Size = 10
80    Printer.Font.Bold = True
90    For n = 0 To 6
100     If oList(n) Then s = Choose(n + 1, "Clinical Conditions", _
                                           "Surgical Procedures", _
                                           "Special Products", _
                                           "Reagents", _
                                           "Batched Products", _
                                           "Cross Match Comments", _
                                           "Reasons for Call", _
                                           "AutoVue Tests")
110   Next

120   Printer.Print FormatString("List of " & s, 99, , AlignCenter)
130   Printer.Font.Size = 9
140   For n = 1 To 108
150       Printer.Print "-";
160   Next n
170   Printer.Print
180   Printer.Print FormatString("", 0, "|");
190   Printer.Print FormatString("Code", 15, "|", AlignCenter);
200   Printer.Print FormatString("Text", 90, "|", AlignCenter)
210   For n = 1 To 108
220       Printer.Print "-";
230   Next n
240   Printer.Print
250   Printer.Font.Bold = False
260   For n = 1 To g.Rows - 1
270     Printer.Print FormatString("", 0, "|");
280     Printer.Print FormatString(g.TextMatrix(n, 0), 15, "|");
290     Printer.Print FormatString(g.TextMatrix(n, 1), 90, "|")
300   Next

310   Printer.EndDoc

320   For Each Px In Printers
330     If Px.DeviceName = OriginalPrinter Then
340       Set Printer = Px
350       Exit For
360     End If
370   Next

380   Exit Sub

cmdPrint_Click_Error:

Dim strES As String
Dim intEL As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "flists", "cmdPrint_Click", intEL, strES

End Sub


Private Sub FillGrid()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillGrid_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from Lists where " & _
            "ListType = '" & mListName & "' " & _
            "Order by ListOrder"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF
90      s = tb!code & vbTab & tb!Text & ""
100     g.AddItem s
110     tb.MoveNext
120   Loop

130   If g.Rows > 2 Then
140     g.RemoveItem 1
150   End If

160   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "flists", "FillGrid", intEL, strES, sql


End Sub

Private Sub cmdSave_Click()

10    SaveDetails
20    FillGrid
End Sub

Private Sub cmdXL_Click()

      Dim strHeading As String

10    On Error GoTo cmdXL_Click_Error

20    strHeading = ""
30    If oList(0).Value = True Then
40        strHeading = strHeading & "Clinical Conditions"
50    ElseIf oList(1).Value = True Then
60        strHeading = strHeading & "Surgical Procedures"
70    ElseIf oList(2).Value = True Then
80        strHeading = strHeading & "Special Products"
90    ElseIf oList(3).Value = True Then
100       strHeading = strHeading & "Reagents"
110   ElseIf oList(4).Value = True Then
120       strHeading = strHeading & "Batched Products"
130   ElseIf oList(5).Value = True Then
140       strHeading = strHeading & "Comments"
150   ElseIf oList(6).Value = True Then
160       strHeading = strHeading & "Reason for Call"
170   ElseIf oList(7).Value = True Then
180       strHeading = strHeading & "AutoVue Tests"
190   End If

      'strHeading = strHeading & vbCr

200   strHeading = strHeading & vbCr & " " & vbCr

210   ExportFlexGrid g, Me, strHeading

220   Exit Sub

cmdXL_Click_Error:

Dim strES As String
Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "flists", "cmdXL_Click", intEL, strES

End Sub



Private Sub Form_Load()


      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        FillGrid
      '**************************************
End Sub

Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim y As Integer
      Dim ySave As Integer

10    On Error GoTo g_Click_Error

20    ySave = g.row

30    If g.MouseRow = 0 Then
40      If SortOrder Then
50        g.Sort = flexSortGenericAscending
60      Else
70        g.Sort = flexSortGenericDescending
80      End If
90      SortOrder = Not SortOrder
100     cmdMoveUp.Enabled = False
110     cmdMoveDown.Enabled = False
120     Exit Sub
130   End If

140   If g.col = 0 Then
150     g.Enabled = False
160     If iMsg("Edit/Remove this line?", vbQuestion + vbYesNo) = vbYes Then
170       tinputcode = g.TextMatrix(g.row, 0)
180       tInputText = g.TextMatrix(g.row, 1)
190       Call MarkGridRow(g, g.row, vbRed, vbYellow, True, True, True)
200       cmdSave.Visible = True
210     End If
220     g.Enabled = True
230     Exit Sub
240   End If

250   g.Visible = False
260   g.col = 0
270   For y = 1 To g.Rows - 1
280     g.row = y
290     If g.CellBackColor = vbYellow Then
300       For X = 0 To g.Cols - 1
310         g.col = X
320         g.CellBackColor = 0
330       Next
340       Exit For
350     End If
360   Next
370   g.row = ySave
380   g.Visible = True

390   For X = 0 To g.Cols - 1
400     g.col = X
410     g.CellBackColor = vbYellow
420   Next

430   cmdMoveUp.Enabled = True
440   cmdMoveDown.Enabled = True

450   Exit Sub

g_Click_Error:

Dim strES As String
Dim intEL As Integer

460   intEL = Erl
470   strES = Err.Description
480   LogError "flists", "g_Click", intEL, strES

End Sub

Private Sub olist_Click(Index As Integer)

10    If cmdSave.Visible Then
20      Answer = iMsg("Save these details?", vbQuestion + vbYesNo)
30      If TimedOut Then Unload Me: Exit Sub
40      If Answer = vbYes Then
50        SaveDetails
60      End If
70    End If

80    mListName = Choose(Index + 1, "X", "P", "S", "R", "B", "XC", "RFC", "AV")

90    FillGrid

End Sub

Private Sub tmrDown_Timer()

10    FireDown

End Sub


Private Sub tmrUp_Timer()

10    FireUp

End Sub


