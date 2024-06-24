VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmXMABControl 
   Caption         =   "NetAcquire"
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "7frmXMABControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7815
   Visible         =   0   'False
   Begin VB.TextBox txtComment 
      Height          =   585
      Left            =   360
      TabIndex        =   6
      Top             =   5190
      Width           =   5685
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Previous"
      Height          =   885
      Left            =   6300
      Picture         =   "7frmXMABControl.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2430
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   375
      Left            =   3540
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92536833
      CurrentDate     =   37735
   End
   Begin MSFlexGridLib.MSFlexGrid gDFya 
      Height          =   1305
      Left            =   390
      TabIndex        =   3
      Top             =   3450
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2302
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      RowHeightMin    =   400
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   "<                          |^Weak Anti D   |^Weak Anti Fya |^AB Serum    ;|D Cells|Fya Cells"
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   885
      Left            =   6300
      Picture         =   "7frmXMABControl.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3450
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   885
      Left            =   6330
      Picture         =   "7frmXMABControl.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid gCards 
      Height          =   2925
      Left            =   390
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5159
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      RowHeightMin    =   400
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   $"7frmXMABControl.frx":1C08
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
      Height          =   225
      Left            =   315
      TabIndex        =   8
      Top             =   6090
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Comment"
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   4950
      Width           =   1365
   End
End
Attribute VB_Name = "frmXMABControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intCardsX As Integer
Private intCardsY As Integer

Private Function CheckPattern() As Boolean

      Dim strPattern As String
      Dim blnMatching As Boolean

10    blnMatching = True

20    With gDFya
30      strPattern = .TextMatrix(1, 1) & .TextMatrix(1, 2) & .TextMatrix(1, 3)
40      If strPattern <> "+XO" Then blnMatching = False
50      strPattern = .TextMatrix(2, 1) & .TextMatrix(2, 2) & .TextMatrix(2, 3)
60      If strPattern <> "X+O" Then blnMatching = False
70    End With

      'With gGrouping
      '  strPattern = .TextMatrix(1, 1) & .TextMatrix(1, 2) & .TextMatrix(1, 3) & .TextMatrix(1, 4)
      '  If strPattern <> "+OXX" Then blnMatching = False
      '  strPattern = .TextMatrix(2, 1) & .TextMatrix(2, 2) & .TextMatrix(2, 3) & .TextMatrix(2, 4)
      '  If strPattern <> "O+XX" Then blnMatching = False
      '  strPattern = .TextMatrix(3, 1) & .TextMatrix(3, 2) & .TextMatrix(3, 3) & .TextMatrix(3, 4)
      '  If strPattern <> "XX+O" Then blnMatching = False
      '  strPattern = .TextMatrix(4, 1) & .TextMatrix(4, 2) & .TextMatrix(4, 3) & .TextMatrix(4, 4)
      '  If strPattern <> "XX+O" Then blnMatching = False
      'End With

80    CheckPattern = blnMatching

End Function

Private Sub LoadPrevious()

      Dim tb As Recordset
      Dim sql As String
      Dim intY As Integer

10    On Error GoTo LoadPrevious_Error

20    sql = "Select top 1 DateTime from XMABControlCards " & _
            "Order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      sql = "Select * from XMABControlCards where " & _
              "DateTime = '" & Format(tb!DateTime, "dd/mmm/yyyy hh:mm:ss") & "'"
70      Set tb = New Recordset
80      RecOpenServerBB 0, tb, sql
90      Do While Not tb.EOF
100       For intY = 1 To gCards.Rows - 1
110         If gCards.TextMatrix(intY, 0) = tb!Type Then
120           gCards.TextMatrix(intY, 1) = tb!Batch & ""
130           If IsDate(tb!Expiry) Then
140             gCards.TextMatrix(intY, 2) = Format(tb!Expiry, "dd/mm/yyyy")
150           End If
160           gCards.TextMatrix(intY, 3) = tb!Manufacturer & ""
170           Exit For
180         End If
190       Next
200       tb.MoveNext
210     Loop
220   End If
  
230   sql = "Select top 1 * from XMABControlPatterns " & _
            "Order by DateTime desc"
240   Set tb = New Recordset
250   RecOpenServerBB 0, tb, sql
260   If Not tb.EOF Then
270     gDFya.TextMatrix(1, 1) = tb!dcwad & ""
280     gDFya.TextMatrix(1, 3) = tb!dcabs & ""
290     gDFya.TextMatrix(2, 2) = tb!fcwaf & ""
300     gDFya.TextMatrix(2, 3) = tb!fcabs & ""
      '  gGrouping.TextMatrix(1, 1) = tb!aaa2 & ""
      '  gGrouping.TextMatrix(1, 2) = tb!aab & ""
      '  gGrouping.TextMatrix(2, 1) = tb!aba2 & ""
      '  gGrouping.TextMatrix(2, 2) = tb!abb & ""
      '  gGrouping.TextMatrix(3, 3) = tb!adr1r1 & ""
      '  gGrouping.TextMatrix(3, 4) = tb!adrr & ""
      '  gGrouping.TextMatrix(4, 3) = tb!ad1r1r1 & ""
      '  gGrouping.TextMatrix(4, 4) = tb!ad1rr & ""
      '  txtComment = tb!Comment & ""
310   End If

320   Exit Sub

LoadPrevious_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmXMABControl", "LoadPrevious", intEL, strES, sql


End Sub

Private Sub SaveDetails()

      Dim tb As Recordset
      Dim sql As String
      Dim intY As Integer
      Dim strEntryDate As String

10    On Error GoTo SaveDetails_Error

20    strEntryDate = Format(Now, "dd/mmm/yyyy hh:mm:ss")

30    sql = "Select * from XMABControlCards " & _
            "where Expiry is null Order by DateTime desc"
40    Set tb = New Recordset
50    RecOpenClientBB 0, tb, sql
60    For intY = 1 To gCards.Rows - 1
70      tb.AddNew
80      tb!DateTime = strEntryDate
90      tb!Operator = UserName
100     tb!Type = gCards.TextMatrix(intY, 0)
110     tb!Batch = gCards.TextMatrix(intY, 1)
120     If IsDate(gCards.TextMatrix(intY, 2)) Then
130       tb!Expiry = gCards.TextMatrix(intY, 2)
140     Else
150       tb!Expiry = Null
160     End If
170     tb!Manufacturer = gCards.TextMatrix(intY, 3)
180     tb.Update
190   Next

200   sql = "Select * from XMABControlPatterns " & _
            "Where DateTime = '01/01/2001'"
210   Set tb = New Recordset
220   RecOpenServerBB 0, tb, sql

230   tb.AddNew
240   tb!DateTime = strEntryDate
250   tb!Operator = UserName
260   tb!dcwad = gDFya.TextMatrix(1, 1)
270   tb!dcabs = gDFya.TextMatrix(1, 3)
280   tb!fcwaf = gDFya.TextMatrix(2, 2)
290   tb!fcabs = gDFya.TextMatrix(2, 3)
      'tb!aaa2 = gGrouping.TextMatrix(1, 1)
      'tb!aab = gGrouping.TextMatrix(1, 2)
      'tb!aba2 = gGrouping.TextMatrix(2, 1)
      'tb!abb = gGrouping.TextMatrix(2, 2)
      'tb!adr1r1 = gGrouping.TextMatrix(3, 3)
      'tb!adrr = gGrouping.TextMatrix(3, 4)
      'tb!ad1r1r1 = gGrouping.TextMatrix(4, 3)
      'tb!ad1rr = gGrouping.TextMatrix(4, 4)
300   tb!Comment = txtComment & " "
310   tb.Update

320   Exit Sub

SaveDetails_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmXMABControl", "SaveDetails", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdLoad_Click()

10    LoadPrevious

End Sub

Private Sub cmdSave_Click()

      Dim s As String

10    If CheckPattern() Then
20      SaveDetails
30      Unload Me
40    Else
50      Answer = iMsg("Patterns not Correct!" & vbCrLf & "Proceed with Save?", vbQuestion + vbYesNo, , vbRed)
60      If TimedOut Then Unload Me: Exit Sub
70      If Answer = vbYes Then
80        s = iBOX("Why are patterns not correct?")
90        If TimedOut Then Unload Me: Exit Sub
100       If Trim$(s) = "" Then
110         iMsg "No reason given - not saved", vbExclamation
120         If TimedOut Then Unload Me: Exit Sub
130         Exit Sub
140       Else
150         LogReasonWhy "Incorrect QC Pattern. " & s, "QC"
160         SaveDetails
170         Unload Me
180       End If
190     End If
200   End If

End Sub

Private Sub dt_CloseUp()
    
10    gCards.TextMatrix(intCardsY, intCardsX) = Format(dt, "dd/mm/yyyy")
20    dt.Visible = False
30    gCards.Enabled = True
40    gDFya.Enabled = True
      'gGrouping.Enabled = True
50    cmdCancel.Enabled = True
60    cmdSave.Enabled = True
70    cmdLoad.Enabled = True

End Sub


Private Sub Form_Load()

10    With gCards
20      .FormatString = "<                          |^Batch Number   |^Expiry Date |^Manufacturer  " & _
                        ";Type|Diamed Cards|D Cells|Fya Cells|Weak Anti D|Weak Anti Fya|AB Serum"
30      .TextMatrix(1, 3) = "Diamed"
40    End With

50    With gDFya
60      .FormatString = "<                          |^Weak Anti D   |^Weak Anti Fya |^AB Serum     " & _
                        ";|D Cells|Fya Cells"
70      .Row = 1
80      .Col = 2
90      .CellBackColor = &H8000000F
100     .Text = "X"
110     .Row = 2
120     .Col = 1
130     .CellBackColor = &H8000000F
140     .Text = "X"
150   End With

      'With gGrouping
      '  .FormatString = "<                          |^A2            |^B             |^R1R1       |^rr            " & _
      '                  ";|Anti A|Anti B|Anti D|Anti D"
      '  .Row = 1
      '  .Col = 3
      '  .CellBackColor = &H8000000F
      '  .Text = "X"
      '  .Col = 4
      '  .CellBackColor = &H8000000F
      '  .Text = "X"
      '  .Row = 2
      '  .Col = 3
      '  .CellBackColor = &H8000000F
      '  .Text = "X"
      '  .Col = 4
      '  .CellBackColor = &H8000000F
      '  .Text = "X"
      '
      '  .Row = 3
      '  .Col = 1
      '  .CellBackColor = &H8000000F
      '  .Text = "X"
      '  .Col = 2
      '  .CellBackColor = &H8000000F
      '  .Text = "X"
      '  .Row = 4
      '  .Col = 1
      '  .CellBackColor = &H8000000F
      '  .Text = "X"
      '  .Col = 2
      '  .CellBackColor = &H8000000F
      '  .Text = "X"
      'End With

End Sub


Private Sub gCards_Click()

      Dim strPrompt As String
      Dim strIP As String

10    With gCards
20      If .MouseCol * .MouseRow = 0 Then Exit Sub
  
30      If .Col = 2 Then 'Expiry
40        .Enabled = False
50        intCardsX = .Col
60        intCardsY = .Row
70        gDFya.Enabled = False
      '    gGrouping.Enabled = False
80        cmdCancel.Enabled = False
90        cmdSave.Enabled = False
100       cmdLoad.Enabled = False
110       If IsDate(.TextMatrix(.Row, .Col)) Then
120         dt = .TextMatrix(.Row, .Col)
130       Else
140         dt = Format(Now, "dd/mm/yyyy")
150       End If
160       dt.Top = .CellTop + dt.Height
170       dt.Visible = True
180       dt.SetFocus
190     Else
200       .Enabled = False
210       strPrompt = "Enter " & .TextMatrix(0, gCards.Col) & " for" & vbCrLf & .TextMatrix(gCards.Row, 0)
220       strIP = iBOX(strPrompt, , .TextMatrix(.Row, .Col))
230       If TimedOut Then Unload Me: Exit Sub
240       .TextMatrix(.Row, .Col) = strIP
250       .Enabled = True
260     End If
270   End With

End Sub


Private Sub gDFya_Click()

10    With gDFya
20      If .MouseCol * .MouseRow = 0 Then Exit Sub
30      If .TextMatrix(.Row, .Col) = "X" Then Exit Sub
40      If .TextMatrix(.Row, .Col) = "O" Then
50        .TextMatrix(.Row, .Col) = "+"
60      Else
70        .TextMatrix(.Row, .Col) = "O"
80      End If
90    End With
    
End Sub




