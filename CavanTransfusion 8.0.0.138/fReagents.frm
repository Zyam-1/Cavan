VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form fReagents 
   Caption         =   "Reagents"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   720
   ClientWidth     =   13515
   Icon            =   "fReagents.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   13515
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   795
      Left            =   12630
      Picture         =   "fReagents.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2610
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   12570
      Picture         =   "fReagents.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4200
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   525
      Left            =   12570
      Picture         =   "fReagents.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4770
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save List Order"
      Height          =   795
      Left            =   12630
      Picture         =   "fReagents.frx":1458
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6030
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   12600
      Top             =   3780
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   12570
      Top             =   5280
   End
   Begin VB.CommandButton bConfirm 
      Caption         =   "Confirm Current Lot &Numbers"
      Height          =   945
      Left            =   12210
      Picture         =   "fReagents.frx":189A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   510
      Width           =   1215
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   795
      Left            =   12630
      Picture         =   "fReagents.frx":1CDC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame2 
      Caption         =   "Show"
      Height          =   1665
      Left            =   8400
      TabIndex        =   15
      Top             =   60
      Width           =   3405
      Begin VB.CheckBox chkIncludeExpired 
         Caption         =   "Include Expired"
         Height          =   225
         Left            =   1860
         TabIndex        =   27
         Top             =   210
         Width           =   1395
      End
      Begin VB.OptionButton o 
         Caption         =   "In Use"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   20
         Top             =   210
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton o 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   210
         Width           =   465
      End
      Begin VB.ListBox lShow 
         Height          =   1110
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   16
         Top             =   450
         Width           =   3075
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   12630
      Picture         =   "fReagents.frx":2346
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6870
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Reagent"
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   8235
      Begin VB.CommandButton bAdd 
         Caption         =   "&Add"
         Height          =   945
         Left            =   7230
         Picture         =   "fReagents.frx":29B0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   420
         Width           =   765
      End
      Begin MSComCtl2.UpDown udQuantity 
         Height          =   285
         Left            =   6420
         TabIndex        =   11
         Top             =   1080
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtQuantity"
         BuddyDispid     =   196625
         OrigLeft        =   6480
         OrigTop         =   990
         OrigRight       =   6900
         OrigBottom      =   1275
         Max             =   999
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   5820
         TabIndex        =   10
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dExpiry 
         Height          =   315
         Left            =   5160
         TabIndex        =   7
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   36818
      End
      Begin VB.TextBox tLot 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1050
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   5115
      End
      Begin VB.ComboBox cBlock 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "cBlock"
         Top             =   360
         Width           =   3045
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Block"
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   420
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Quantity Received"
         Height          =   195
         Left            =   4440
         TabIndex        =   9
         Top             =   1110
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Expiry"
         Height          =   195
         Left            =   4650
         TabIndex        =   8
         Top             =   390
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lot"
         Height          =   195
         Left            =   990
         TabIndex        =   5
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reagent Name"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   750
         Width           =   1080
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5865
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   10345
      _Version        =   393216
      Cols            =   12
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
      AllowUserResizing=   1
      FormatString    =   $"fReagents.frx":2DF2
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
      Left            =   120
      TabIndex        =   24
      Top             =   7830
      Width           =   13320
      _ExtentX        =   23495
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
      Height          =   285
      Left            =   12570
      TabIndex        =   26
      Top             =   3420
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "fReagents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean


Private FireCounter As Integer



Private Sub chkIncludeExpired_Click()

10    FillG

End Sub

Private Sub cmdXL_Click()

10    ExportFlexGrid g, Me

End Sub

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





Private Sub FillDetails()

      Dim n As Integer
      Dim Found As Boolean
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillDetails_Error

20    cBlock.Clear
30    lShow.Clear

40    sql = "SELECT * FROM Reagents WHERE " & _
            "Block IS NOT NULL " & _
            "ORDER BY ListOrder"
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    Do While Not tb.EOF
80      Found = False
90      For n = 0 To cBlock.ListCount - 1
100       If cBlock.List(n) = tb!Block & "" Then
110         Found = True
120         Exit For
130       End If
140     Next
150     If Not Found Then
160       cBlock.AddItem tb!Block & ""
170       lShow.AddItem tb!Block & ""
180     End If
190     tb.MoveNext
200   Loop

210   cBlock.ListIndex = -1
220   lShow.AddItem "All", 0
230   lShow.ListIndex = 0

240   dExpiry = Format(Now, "dd/MM/yyyy")

250   Exit Sub

FillDetails_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "fReagents", "FillDetails", intEL, strES, sql


End Sub

Private Sub FillG()

      Dim s As String
      Dim tb As Recordset
      Dim sql As String
      Dim Temp As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1
      '^In Use|<Reagent Name                       |
      ' <Lot Number                     |<Expiry Date        |
      ' <Quantity Received |<Received By |<Received At |
      ' <Batch Validation By |<Batch Validation At |<Comments

50    sql = "Select * from Reagents "
60    If chkIncludeExpired = 0 Then
70      sql = sql & "WHERE Expiry > getdate() "
80    End If
90    sql = sql & "order by listorder"
100   Set tb = New Recordset
110   RecOpenServerBB 0, tb, sql
120   Do While Not tb.EOF
130     If lShow = "All" Or tb!Block = lShow Then
140       If o(0) Or tb!InUse Then

150         s = IIf(tb!InUse, "Yes", "No") & vbTab & _
                tb!Name & vbTab & _
                tb!lot & vbTab & _
                Format$(tb!Expiry, "dd/MM/yy") & vbTab & _
                tb!QuantityReceived & vbTab & _
                TechnicianCodeFor(tb!LoggedInBy & "") & vbTab & _
                Format$(tb!LoggedInDateTime, "dd/MM/yy HH:nn:ss") & vbTab
160         Temp = TechnicianCodeFor(tb!OpenedBy & "")
170         If Temp <> "???" Then
180           s = s & Temp
190         End If
200         s = s & vbTab & _
                Format$(tb!Opened, "dd/MM/yy HH:nn:ss") & vbTab
210         Temp = TechnicianCodeFor(tb!ValidationBy & "")
220         If Temp <> "???" Then
230           s = s & Temp
240         End If
250         s = s & vbTab & _
                Format$(tb!ValidationDateTime, "dd/MM/yy HH:nn:ss") & vbTab & _
                tb!Comments & ""

260         g.AddItem s
270       End If
280     End If
290     tb.MoveNext
300   Loop

310   If g.Rows > 2 Then
320     g.RemoveItem 1
330   End If

340   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "fReagents", "FillG", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo bAdd_Click_Error

20    If Format$(dExpiry, "dd/MM/yyyy") = Format$(Now, "dd/MM/yyyy") Then
30      Answer = iMsg("Expiry Date is today!" & vbCrLf & _
                "Is this correct?", vbQuestion + vbYesNo, , vbRed)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Exit Sub
70      End If
80    End If

90    If Val(txtQuantity) = 0 Then
100     iMsg "Specify Quantity Received!", vbCritical
110     If TimedOut Then Unload Me: Exit Sub
120     Exit Sub
130   End If

140   If cBlock = "" Then
150     iMsg "Specify Block", vbCritical
160     If TimedOut Then Unload Me: Exit Sub
170     Exit Sub
180   End If

190   If Trim$(txtName) = "" Then
200     iMsg "Specify Reagent Name", vbCritical
210     If TimedOut Then Unload Me: Exit Sub
220     Exit Sub
230   End If

240   If Trim$(tLot) = "" Then
250     iMsg "Specify Lot Number", vbCritical
260     If TimedOut Then Unload Me: Exit Sub
270     Exit Sub
280   End If

290   sql = "Select * from Reagents where 0 = 1"

300   Set tb = New Recordset
310   RecOpenServerBB 0, tb, sql
      'If tb.EOF Then
320     tb.AddNew
      'End If
330   tb!Block = cBlock
340   tb!Name = txtName
350   tb!lot = tLot
360   tb!Expiry = Format(dExpiry, "dd/MMM/yyyy")
370   tb!InUse = 0
380   tb!ListOrder = 999
390   tb!QuantityReceived = Val(txtQuantity)
400   tb!LoggedInBy = UserName
410   tb!LoggedInDateTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
420   tb.Update

430   FillDetails
440   FillG
450   txtName = ""
460   tLot = ""
470   txtQuantity = "0"

480   Exit Sub

bAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

490   intEL = Erl
500   strES = Err.Description
510   LogError "fReagents", "bAdd_Click", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub bConfirm_Click()

      Dim tbRx As Recordset
      Dim tb As Recordset
      Dim TimeNow As String
      Dim sql As String

10    On Error GoTo bConfirm_Click_Error

20    TimeNow = Format(Now, "dd/MMM/yyyy HH:mm:ss")

30    sql = "Select * from Reagents where InUse = 1"
40    Set tbRx = New Recordset
50    RecOpenServerBB 0, tbRx, sql

60    sql = "Select * from ReagentQC where Block = 'x'"
70    Set tb = New Recordset
80    RecOpenServerBB 0, tb, sql
90    Do While Not tbRx.EOF
100     tb.AddNew
110     tb!Block = tbRx!Block
120     tb!Name = tbRx!Name
130     tb!lot = tbRx!lot
140     tb!DateTime = TimeNow
150     tb!Operator = UserCode
160     tb.Update
170     tbRx.MoveNext
180   Loop

190   Unload Me

200   Exit Sub

bConfirm_Click_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "fReagents", "bConfirm_Click", intEL, strES, sql


End Sub

Private Sub bprint_Click()

      Dim n As Integer
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Orientation = vbPRORLandscape
40    Printer.Font.Name = "Courier New"
50    Printer.Font.Size = 14
60    Printer.Font.Bold = True
70    Printer.ForeColor = vbRed
80    Printer.Print "Cavan General Hospital : Blood Transfusion Laboratory";
90    Printer.Font.Size = 10
100   Printer.CurrentY = 100

110   Printer.Print

120   Printer.CurrentY = 320

130   Printer.Font.Size = 4
140   Printer.Print String$(230, "-")

150   Printer.ForeColor = vbBlack

160   Printer.Font.Name = "Courier New"
170   Printer.Font.Size = 12
180   Printer.Font.Bold = False
190   Printer.Print "Reagents in use " & Format(Now, "dd/mmmm/yyyy")
200   Printer.Print

210   For n = 1 To lShow.ListCount - 1
220     lShow.Selected(n) = True
230     Printer.Font.Bold = True
240     Printer.Print UCase$(lShow)
250     Printer.Font.Bold = False
260     For Y = 1 To g.Rows - 1
270       Printer.Print g.TextMatrix(Y, 1);
280       Printer.Print Tab(30); g.TextMatrix(Y, 2);
290       Printer.Print Tab(50); g.TextMatrix(Y, 3)
300     Next
310     Printer.Print
320   Next
330   Printer.EndDoc

340   lShow.Selected(0) = True

350   For Each Px In Printers
360     If Px.DeviceName = OriginalPrinter Then
370       Set Printer = Px
380       Exit For
390     End If
400   Next

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

      Dim Y As Integer
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    cmdSave.Caption = "Saving..."

30    For Y = 1 To g.Rows - 1
40      sql = "Update Reagents set ListOrder = '" & Y & "' where " & _
              "Name = '" & Trim$(g.TextMatrix(Y, 1)) & "' " & _
              "and Lot = '" & Trim$(g.TextMatrix(Y, 2)) & "'"
50      CnxnBB(0).Execute sql
60    Next

70    cmdSave.Visible = False
80    cmdSave.Caption = "Save"

90    Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "fReagents", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

10    FillDetails

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Visible Then
30      Answer = iMsg("Cancel without saving List Order?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70      End If
80    End If

End Sub


Private Sub g_Click()

      Dim sql As String
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer
      Dim ValBy As String
      Dim OpenBy As String
      Dim Comment As String

10    On Error GoTo g_Click_Error

20    If g.MouseRow = 0 Then
30      If g.MouseCol = 3 Then
40        g.Sort = 9
50      Else
60        If SortOrder Then
70          g.Sort = flexSortGenericAscending
80        Else
90          g.Sort = flexSortGenericDescending
100       End If
110     End If
120     SortOrder = Not SortOrder
130     Exit Sub
140   End If

150   If g.Col = 0 Then
160     If g.TextMatrix(g.Row, 9) <> "" And g.TextMatrix(g.Row, 10) <> "" Then
170       g = IIf(g = "Yes", "No", "Yes")

180       sql = "UPDATE Reagents SET InUse = " & _
                IIf(g.TextMatrix(g.Row, 0) = "Yes", 1, 0) & " " & _
                "WHERE Name = '" & Trim$(g.TextMatrix(g.Row, 1)) & "' " & _
                "and Lot = '" & Trim$(g.TextMatrix(g.Row, 2)) & "'"
190       CnxnBB(0).Execute sql
200     Else
210       iMsg "Batch is not Validated.", vbCritical
220       If TimedOut Then Exit Sub: Unload Me
230     End If
240   ElseIf g.Col = 7 And g.TextMatrix(g.Row, 7) = "" Then
250     OpenBy = Trim$(iBOX("Batch Opened by", , UserName))
260     If TimedOut Then Exit Sub: Unload Me
270     If OpenBy <> "" Then
280       g.TextMatrix(g.Row, 7) = OpenBy
290       g.TextMatrix(g.Row, 8) = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
300       sql = "UPDATE Reagents SET " & _
                "OpenedBy = '" & OpenBy & "', " & _
                "Opened = '" & g.TextMatrix(g.Row, 8) & "' " & _
                "WHERE Name = '" & Trim$(g.TextMatrix(g.Row, 1)) & "' " & _
                "and Lot = '" & Trim$(g.TextMatrix(g.Row, 2)) & "'"
310       CnxnBB(0).Execute sql
320     End If
330   ElseIf g.Col = 9 And g.TextMatrix(g.Row, 9) = "" Then
340     ValBy = Trim$(iBOX("Batch Validation by", , UserName))
350     If TimedOut Then Exit Sub: Unload Me
360     If ValBy <> "" Then
370       g.TextMatrix(g.Row, 9) = ValBy
380       g.TextMatrix(g.Row, 10) = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
390       sql = "UPDATE Reagents SET " & _
                "ValidationBy = '" & ValBy & "', " & _
                "ValidationDateTime = '" & g.TextMatrix(g.Row, 10) & "' " & _
                "WHERE Name = '" & Trim$(g.TextMatrix(g.Row, 1)) & "' " & _
                "and Lot = '" & Trim$(g.TextMatrix(g.Row, 2)) & "'"
400       CnxnBB(0).Execute sql
410     End If
420   ElseIf g.Col = 11 Then
430     Comment = iBOX("Comment")
440     If TimedOut Then Exit Sub: Unload Me
450     If Trim$(Comment) <> "" Then
460       g.TextMatrix(g.Row, 11) = Comment
470       sql = "UPDATE Reagents SET " & _
                "Comments = '" & Comment & "' " & _
                "WHERE Name = '" & Trim$(g.TextMatrix(g.Row, 1)) & "' " & _
                "and Lot = '" & Trim$(g.TextMatrix(g.Row, 2)) & "'"
480       CnxnBB(0).Execute sql
490     End If
500   Else
510     ySave = g.Row
520     g.Visible = False
530     g.Col = 0
540     For Y = 1 To g.Rows - 1
550       g.Row = Y
560       If g.CellBackColor = vbYellow Then
570         For X = 0 To g.Cols - 1
580           g.Col = X
590           g.CellBackColor = 0
600         Next
610         Exit For
620       End If
630     Next
640     g.Row = ySave
650     g.Visible = True
  
660     For X = 0 To g.Cols - 1
670       g.Col = X
680       g.CellBackColor = vbYellow
690     Next
  
700     cmdMoveUp.Enabled = True
710     cmdMoveDown.Enabled = True
720   End If

730   Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

740   intEL = Erl
750   strES = Err.Description
760   LogError "fReagents", "g_Click", intEL, strES, sql


End Sub

Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    If Not IsDate(g.TextMatrix(Row1, 3)) Then
20      Cmp = 0
30      Exit Sub
40    End If

50    If Not IsDate(g.TextMatrix(Row2, 3)) Then
60      Cmp = 0
70      Exit Sub
80    End If

90    d1 = Format(g.TextMatrix(Row1, 3), "dd/MMM/yyyy HH:mm:ss")
100   d2 = Format(g.TextMatrix(Row2, 3), "dd/MMM/yyyy HH:mm:ss")

110   If SortOrder Then
120     Cmp = Sgn(DateDiff("D", d1, d2))
130   Else
140     Cmp = Sgn(DateDiff("D", d2, d1))
150   End If

End Sub

Private Sub lShow_Click()

10    bprint.Visible = lShow = "All"
      'g.Enabled = lShow <> "All"

20    FillG

End Sub

Private Sub o_Click(Index As Integer)

10    FillG

End Sub

Private Sub tmrDown_Timer()

10    FireDown

End Sub

Private Sub tmrUp_Timer()

10    FireUp

End Sub

Private Sub txtName_LostFocus()

      Dim QueryDate As String
    
10    dExpiry = Format$(Now, "dd/MM/yyyy")

20    If Len(txtName) = 14 Then
30      If UCase$(Left$(txtName, 1)) = "A" And UCase$(Right$(txtName, 1)) = "B" Then
40        cBlock = "Internal Quality Control"
50        tLot = Mid$(txtName, 11, 3)
60        txtName = "Ortho CQI7"
70      Else
          '06012181061113
80        If Mid$(txtName, 4, 2) = "12" Then
90          cBlock = "Cell Panel"
100         tLot = Left$(txtName, 5) & "." & Mid$(txtName, 6, 2) & "." & Mid$(txtName, 8, 1)
110         dExpiry = Right$(txtName, 2) & "/" & Mid$(txtName, 11, 2) & "/" & Mid$(txtName, 9, 2)
120         txtName = "DiaCell A1 (Reverse Grouping Cells)"
130       ElseIf Mid$(txtName, 4, 2) = "22" Then
140         cBlock = "Cell Panel"
150         tLot = Left$(txtName, 5) & "." & Mid$(txtName, 6, 2) & "." & Mid$(txtName, 8, 1)
160         dExpiry = Right$(txtName, 2) & "/" & Mid$(txtName, 11, 2) & "/" & Mid$(txtName, 9, 2)
170         txtName = "DiaCell A2 (Reverse Grouping Cells)"
180       ElseIf Mid$(txtName, 4, 2) = "32" Then
190         cBlock = "Cell Panel"
200         tLot = Left$(txtName, 5) & "." & Mid$(txtName, 6, 2) & "." & Mid$(txtName, 8, 1)
210         dExpiry = Right$(txtName, 2) & "/" & Mid$(txtName, 11, 2) & "/" & Mid$(txtName, 9, 2)
220         txtName = "DiaCell B (Reverse Grouping Cells)"
230       End If
240     End If
250   ElseIf Len(txtName) = 10 Then
        '8976140578
        '8906310433
        '9186510201
        '0603732180
260     If Mid$(txtName, 5, 3) = "621" Then
270       cBlock = "Antibody Panel"
280       tLot = "8RC" & Right$(txtName, 3)
290       txtName = "0.8% Resolve Panel C UNTREATED"
300     ElseIf Mid$(txtName, 5, 3) = "632" Or Mid$(txtName, 5, 3) = "732" Then
310       cBlock = "Antibody Panel"
320       tLot = "8RC" & Right$(txtName, 3)
330       txtName = "0.8% Resolve Panel C Ficin Treated"
340     ElseIf Mid$(txtName, 5, 3) = "620" Then
350       cBlock = "Antibody Panel"
360       tLot = "8RB" & Right$(txtName, 3)
370       txtName = "0.8% Resolve Panel B"
380     ElseIf Mid$(txtName, 5, 3) = "510" Then
390       cBlock = "Antibody Panel"
400       tLot = "8RA" & Right$(txtName, 3)
410       txtName = "0.8% Resolve Panel A"
420     ElseIf Mid$(txtName, 5, 3) = "140" Then
430       cBlock = "Cell Panel"
440       tLot = "A" & Right$(txtName, 3) & "Z"
450       txtName = "Affirmagen Cells"
460     ElseIf Mid$(txtName, 5, 3) = "120" Then
470       cBlock = "Cell Panel"
480       tLot = "A" & Right$(txtName, 3) & "Z"
490       txtName = "Affirmagen Cells"
500     ElseIf Mid$(txtName, 5, 3) = "310" Then
510       cBlock = "Surgiscreen Cells"
520       tLot = "8SS" & Right$(txtName, 3)
530       txtName = "Surgiscreen Cells"
540     ElseIf Mid$(txtName, 5, 3) = "320" Then
550       cBlock = "Surgiscreen Cells"
560       tLot = "8SS" & Right$(txtName, 3)
570       txtName = "Surgiscreen Cells"
580     ElseIf Mid$(txtName, 5, 3) = "330" Then
590       cBlock = "Surgiscreen Cells"
600       tLot = "8SS" & Right$(txtName, 3)
610       txtName = "Surgiscreen Cells"
620     End If
630   ElseIf Len(txtName) = 20 Then
      '18020722092243875103  17030733070691373101  25050766029948940103
640     QueryDate = Left$(txtName, 2) & "/" & Mid$(txtName, 3, 2) & "/" & Mid$(txtName, 5, 2)
650     If IsDate(QueryDate) Then
660       dExpiry = QueryDate
670     End If
680     Select Case Mid$(txtName, 7, 2)
          Case "22": cBlock = "OrthoBioVue AHG Cards"
690                  tLot = "AHC" & Mid$(txtName, 15, 3) & "A"
700                  txtName = "Anti-IgG,-C3d polyspecific"
710       Case "33": cBlock = "OrthoBioVue AHG Cards"
720                  tLot = "IGC" & Mid$(txtName, 15, 3) & "A"
730                  txtName = "Anti-IgG"
740       Case "66": cBlock = "Ortho BioVue Reverse Grouping"
750                  tLot = "RDC" & Mid$(txtName, 15, 3) & "A"
760                  txtName = "Reverse Diluent"
770       Case "40": cBlock = "Ortho Biovue Forward grouping cards"
780                  tLot = "ADK" & Mid$(txtName, 15, 3) & "A"
790                  txtName = "ADK Cards"
800       Case "48": cBlock = "Ortho Biovue Forward grouping cards"
810                  tLot = "ADD" & Mid$(txtName, 15, 3) & "A"
820                  txtName = "ADD Cards"
830       Case "10": cBlock = "Ortho Biovue Forward grouping cards"
840                  tLot = "ACC" & Mid$(txtName, 15, 3) & "A"
850                  txtName = "Unit Group Check Cards"
860       Case "77": cBlock = "Ortho Biovue RhK cards"
870                  tLot = "RHP" & Mid$(txtName, 15, 3) & "B"
880                  txtName = "Phenotyping cards"
890       Case "30": cBlock = "Ortho BioVue DAT"
900                  tLot = "DAT" & Mid$(txtName, 15, 3) & "A"
910                  txtName = "Anti-IgG, C3d, Control"
920       Case "88": cBlock = "Ortho BioVue Enzyme Card"
930                  tLot = "NEC" & Mid$(txtName, 15, 3) & "A"
940                  txtName = "Neutral Card"
950     End Select

960   ElseIf Len(txtName) = 19 Then
  
        '5001212010609424542   5020031010608508144
        'set to last day of month: 1st + 1month - 1day
        'set to first day of month
970     QueryDate = "01/" & Mid$(txtName, 12, 2) & "/" & Mid$(txtName, 10, 2)
980     If IsDate(QueryDate) Then
          'add one month
990       QueryDate = DateAdd("m", 1, QueryDate)
          'less one day
1000      QueryDate = DateAdd("d", -1, QueryDate)
1010      dExpiry = QueryDate
1020    End If
1030    Select Case Left$(txtName, 5)
          Case "50012": tLot = "50012." & Mid$(txtName, 6, 2) & "." & Mid$(txtName, 8, 2)
1040                    cBlock = "Diaclon Cards"
1050                    txtName = "Diaclon ABO/Rh for Patients"
1060      Case "50200": tLot = "50200." & Mid$(txtName, 6, 2) & "." & Mid$(txtName, 8, 2)
1070                    cBlock = "Diaclon Cards"
1080                    txtName = "Diaclon Anti-K"
1090    End Select

1100  ElseIf Len(txtName) = 12 Then
1110    QueryDate = Mid$(txtName, 7, 2) & "/" & Mid$(txtName, 9, 2) & "/" & Mid$(txtName, 11, 2)
1120    If IsDate(QueryDate) Then
1130      dExpiry = QueryDate
1140    End If
1150    cBlock = "Red Cells"
1160    tLot = Left$(txtName, 6)
1170    Select Case Left$(txtName, 3)
          Case "111": txtName = "Immucor A1 Cells"
1180      Case "112": txtName = "Immucor A2 Cells"
1190      Case "113": txtName = "Immucor B Cells"
1200    End Select
1210  End If

End Sub

