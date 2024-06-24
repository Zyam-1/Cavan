VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fWardList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Ward List"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   195
      Left            =   150
      TabIndex        =   24
      Top             =   1380
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ComboBox cmbListItems 
      Height          =   315
      ItemData        =   "fWardList.frx":0000
      Left            =   8520
      List            =   "fWardList.frx":0002
      TabIndex        =   22
      Top             =   720
      Width           =   2235
   End
   Begin VB.CommandButton cmdDelete 
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      Height          =   705
      Left            =   10500
      Picture         =   "fWardList.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1260
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   10500
      Picture         =   "fWardList.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10560
      Top             =   4740
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10560
      Top             =   4170
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   150
      TabIndex        =   15
      Top             =   7290
      Visible         =   0   'False
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   705
      Left            =   10590
      Picture         =   "fWardList.frx":0BD8
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   525
      Left            =   10110
      Picture         =   "fWardList.frx":101A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4890
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   10110
      Picture         =   "fWardList.frx":145C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   465
   End
   Begin VB.ListBox lHospital 
      Height          =   1230
      Left            =   6270
      TabIndex        =   9
      Top             =   150
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add New Ward"
      Height          =   1365
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   5925
      Begin VB.TextBox txtPrinter 
         Height          =   285
         Left            =   810
         TabIndex        =   16
         Top             =   960
         Width           =   4755
      End
      Begin VB.TextBox txtFAX 
         Height          =   285
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   11
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Top             =   420
         Width           =   705
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   810
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   3645
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   810
         MaxLength       =   12
         TabIndex        =   4
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printer"
         Height          =   195
         Left            =   330
         TabIndex        =   17
         Top             =   990
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   1860
         TabIndex        =   10
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Left            =   390
         TabIndex        =   7
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   270
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdWardList 
      Height          =   5445
      Left            =   150
      TabIndex        =   2
      Top             =   1560
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   9604
      _Version        =   393216
      Cols            =   6
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
      FormatString    =   $"fWardList.frx":189E
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   10620
      Picture         =   "fWardList.frx":1958
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6300
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   945
      Left            =   10500
      MaskColor       =   &H8000000F&
      Picture         =   "fWardList.frx":1FC2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2130
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Items to be displayed in list"
      Height          =   195
      Left            =   8520
      TabIndex        =   23
      Top             =   480
      Width           =   1875
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Click on Code to Edit/Remove record."
      Height          =   195
      Left            =   150
      TabIndex        =   21
      Top             =   7050
      Width           =   3255
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
      Left            =   10320
      TabIndex        =   19
      Top             =   3870
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "fWardList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer






Private Sub cmbListItems_Click()
10    cmdSave.Visible = True
End Sub

Private Sub cmbListItems_KeyPress(KeyAscii As Integer)
10    KeyAscii = 0
End Sub




Private Sub cmdDelete_Click()
      'check if record doesn't exist in demographics
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdDelete_Click_Error

20    If grdWardList.Row = 0 Or grdWardList.Rows <= 2 Then Exit Sub

30    sql = "SELECT Count(*) as RC FROM Demographics WHERE Ward = '" & grdWardList.TextMatrix(grdWardList.Row, 2) & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If tb!rc > 0 Then
70        iMsg "Reference to " & grdWardList.TextMatrix(grdWardList.Row, 2) & " is in use so cannot be deleted"
80        Exit Sub
90    Else
100       If iMsg("Are you sure you want to delete " & grdWardList.TextMatrix(grdWardList.Row, 2) & "?", vbYesNo) = vbYes Then
110           Cnxn(0).Execute "DELETE FROM Wards WHERE " & _
                      "Code = '" & grdWardList.TextMatrix(grdWardList.Row, 1) & "' AND Text = '" & grdWardList.TextMatrix(grdWardList.Row, 2) & "'"
120           FillG
130       End If
    
140   End If

150   Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "fWardList", "cmdDelete_Click", intEL, strES, sql

    
End Sub

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

10    With grdWardList
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

10    With grdWardList
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



Private Sub cmdCancel_Click()

10    Unload Me

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


Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.Orientation = vbPRORPortrait


      '****Report heading
50    Printer.FontSize = 10
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print FormatString("List Of Wards (" & lHospital & ")", 99, , AlignCenter)

      '****Report body heading

90    Printer.Font.Size = 9
100   For i = 1 To 108
110       Printer.Print "-";
120   Next i
130   Printer.Print


140   Printer.Print FormatString("", 0, "|");
150   Printer.Print FormatString("In Use", 6, "|", AlignCenter);
160   Printer.Print FormatString("Code", 10, "|", AlignCenter);
170   Printer.Print FormatString("Description", 88, "|", AlignCenter)
      '****Report body

180   Printer.Font.Bold = False

190   For i = 1 To 108
200       Printer.Print "-";
210   Next i
220   Printer.Print
230   For Y = 1 To grdWardList.Rows - 1
240       Printer.Print FormatString("", 0, "|");
250       Printer.Print FormatString(grdWardList.TextMatrix(Y, 0), 6, "|", AlignLeft);
260       Printer.Print FormatString(grdWardList.TextMatrix(Y, 1), 10, "|", AlignLeft);
270       Printer.Print FormatString(grdWardList.TextMatrix(Y, 2), 88, "|", AlignLeft)

280   Next

290   Printer.EndDoc

300   For Each Px In Printers
310     If Px.DeviceName = OriginalPrinter Then
320       Set Printer = Px
330       Exit For
340     End If
350   Next

End Sub

Private Sub cmdadd_Click()

10    txtCode = UCase$(Trim$(txtCode))
20    If txtCode = "" Then
30      iMsg "Enter Code.", vbCritical
40      Exit Sub
50    End If

60    txtText = Trim$(txtText)
70    If txtText = "" Then
80      iMsg "Enter Ward.", vbCritical
90      Exit Sub
100   End If

110   grdWardList.AddItem "Yes" & vbTab & _
                txtCode & vbTab & _
                txtText & vbTab & _
                txtFAX & vbTab & _
                txtPrinter

120   txtCode = ""
130   txtText = ""
140   txtFAX = ""
150   txtPrinter = ""

160   cmdSave.Visible = True

End Sub

Private Sub cmdSave_Click()

      Dim Hosp As String
      Dim Y As Integer
      Dim tb As Recordset
      Dim sql As String

10    sql = "Select * from Lists where " & _
            "ListType = 'HO' " & _
            "and Text = '" & lHospital & "' and InUse = 1"
20    Set tb = New Recordset
30    RecOpenServer 0, tb, sql
40    If Not tb.EOF Then
50      Hosp = tb!code & ""
60    End If

70    PB.max = grdWardList.Rows - 1
80    PB.Visible = True
90    cmdSave.Caption = "Saving..."

100   For Y = 1 To grdWardList.Rows - 1
110     PB = Y
120     sql = "Select * from Wards where " & _
              "Code = '" & grdWardList.TextMatrix(Y, 1) & "' " & _
              "and HospitalCode = '" & Hosp & "'"
130     Set tb = New Recordset
140     RecOpenServer 0, tb, sql
150     If tb.EOF Then tb.AddNew
160     With tb
170       !code = grdWardList.TextMatrix(Y, 1)
180       !HospitalCode = Hosp
190       !InUse = grdWardList.TextMatrix(Y, 0) = "Yes"
200       !Text = grdWardList.TextMatrix(Y, 2)
210       !FAX = grdWardList.TextMatrix(Y, 3)
220       !PrinterAddress = grdWardList.TextMatrix(Y, 4)
230       !ListOrder = Y
240       !Location = grdWardList.TextMatrix(Y, 5)
250       .Update
260     End With
270   Next

280   Call SaveOptionSetting("WardListLength", cmbListItems)

290   PB.Visible = False
300   cmdSave.Visible = False
310   cmdSave.Caption = "Save"

End Sub


Private Sub cmdXL_Click()

10    ExportFlexGrid grdWardList, Me
  
End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

10    lHospital.Clear

20    sql = "Select * from Lists where " & _
            "ListType = 'HO' and InUse = 1 " & _
            "order by ListOrder"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    Do While Not tb.EOF
60      lHospital.AddItem tb!Text & ""
70      tb.MoveNext
80    Loop
90    If lHospital.ListCount > 0 Then
100     lHospital.ListIndex = 0
110   End If

120   FillG

      Dim i  As Integer
130   cmbListItems.Clear
140   For i = 8 To 32 Step 8
150       cmbListItems.AddItem i
160   Next i
170   cmbListItems.Text = GetOptionSetting("WardListLength", 8)

End Sub

Private Sub FillG()

      Dim s As String
      Dim Hosp As String
      Dim sql As String
      Dim tb As Recordset

10    sql = "Select * from Lists where " & _
            "ListType = 'HO' " & _
            "and Text = '" & lHospital & "' and InUse = 1"
20    Set tb = New Recordset
30    RecOpenServer 0, tb, sql
40    If Not tb.EOF Then
50      Hosp = tb!code & ""
60    End If

70    grdWardList.Rows = 2
80    grdWardList.AddItem ""
90    grdWardList.RemoveItem 1

100   sql = "Select * from Wards where " & _
            "HospitalCode = '" & Hosp & "'"
110   Set tb = New Recordset
120   RecOpenServer 0, tb, sql

130   Do While Not tb.EOF
140     With tb
150       s = IIf(!InUse, "Yes", "No") & vbTab & _
              !code & vbTab & _
              !Text & vbTab & _
              !FAX & vbTab & _
              !PrinterAddress & vbTab & _
              !Location & ""
160       grdWardList.AddItem s
170     End With
180     tb.MoveNext
190   Loop

200   If grdWardList.Rows > 2 Then
210     grdWardList.RemoveItem 1
220   End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If cmdSave.Visible Then
20      If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
30        Cancel = True
40      End If
50    End If

End Sub

Private Sub grdWardList_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer
      Dim tb As Recordset
      Dim sql As String


10    ySave = grdWardList.Row

20    If grdWardList.MouseRow = 0 Then
30      If SortOrder Then
40        grdWardList.Sort = flexSortGenericAscending
50      Else
60        grdWardList.Sort = flexSortGenericDescending
70      End If
80      SortOrder = Not SortOrder
90      cmdMoveUp.Enabled = False
100     cmdMoveDown.Enabled = False
110     cmdSave.Visible = True
120     Exit Sub
130   End If


140   If grdWardList.Col = 0 Then
150     grdWardList = IIf(grdWardList = "No", "Yes", "No")
160     cmdSave.Visible = True
170     Exit Sub
180   End If

190   If grdWardList.Col = 1 Then
200       sql = "SELECT Count(*) as RC FROM Demographics WHERE Ward = '" & grdWardList.TextMatrix(grdWardList.Row, 2) & "'"
210       Set tb = New Recordset
220       RecOpenClient 0, tb, sql
230       If tb!rc > 0 Then
240           iMsg "Reference to " & grdWardList.TextMatrix(grdWardList.Row, 2) & " is in use so cannot be deleted"
250           Exit Sub
260       End If

270     grdWardList.Enabled = False
280     If iMsg("Edit this line?", vbQuestion + vbYesNo) = vbYes Then
290       txtCode = grdWardList.TextMatrix(grdWardList.Row, 1)
300       txtText = grdWardList.TextMatrix(grdWardList.Row, 2)
310       txtFAX = grdWardList.TextMatrix(grdWardList.Row, 3)
320       txtPrinter = grdWardList.TextMatrix(grdWardList.Row, 4)
330       grdWardList.RemoveItem grdWardList.Row
340       cmdSave.Visible = True
350     End If
360     grdWardList.Enabled = True
370     Exit Sub
380   End If
    
390   If grdWardList.Col = 5 Then
400     Select Case grdWardList.TextMatrix(grdWardList.Row, 5)
          Case "": grdWardList.TextMatrix(grdWardList.Row, 5) = "In-House"
410       Case "In-House": grdWardList.TextMatrix(grdWardList.Row, 5) = "External"
420       Case "External": grdWardList.TextMatrix(grdWardList.Row, 5) = ""
430       Case "Else": grdWardList.TextMatrix(grdWardList.Row, 5) = ""
440     End Select
450     cmdSave.Visible = True
460     Exit Sub
470   End If

480   grdWardList.Visible = False
490   grdWardList.Col = 0
500   For Y = 1 To grdWardList.Rows - 1
510     grdWardList.Row = Y
520     If grdWardList.CellBackColor = vbYellow Then
530       For X = 0 To grdWardList.Cols - 1
540         grdWardList.Col = X
550         grdWardList.CellBackColor = 0
560       Next
570       Exit For
580     End If
590   Next
600   grdWardList.Row = ySave
610   grdWardList.Visible = True

620   For X = 0 To grdWardList.Cols - 1
630     grdWardList.Col = X
640     grdWardList.CellBackColor = vbYellow
650   Next

660   cmdMoveUp.Enabled = True
670   cmdMoveDown.Enabled = True

End Sub

Private Sub grdWardList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    If grdWardList.MouseRow = 0 Then
20      grdWardList.ToolTipText = ""
30    ElseIf grdWardList.MouseCol = 0 Then
40      grdWardList.ToolTipText = "Click to Toggle"
50    ElseIf grdWardList.MouseCol = 1 Then
60      grdWardList.ToolTipText = "Click to Edit"
70    ElseIf grdWardList.MouseCol = 5 Then
80      grdWardList.ToolTipText = "Click to Set"
90    Else
100     grdWardList.ToolTipText = "Click to Move"
110   End If

End Sub


Private Sub lHospital_Click()

10    FillG

End Sub

Private Sub tmrDown_Timer()

10    FireDown

End Sub


Private Sub tmrUp_Timer()

10    FireUp

End Sub


