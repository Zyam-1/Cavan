VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmstock 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Ordering"
   ClientHeight    =   8505
   ClientLeft      =   150
   ClientTop       =   465
   ClientWidth     =   12075
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "Frmstock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8505
   ScaleWidth      =   12075
   Begin VB.ComboBox cFAXNumber 
      Height          =   315
      ItemData        =   "Frmstock.frx":08CA
      Left            =   8970
      List            =   "Frmstock.frx":08CC
      TabIndex        =   15
      Text            =   "016674794"
      Top             =   6840
      Width           =   1245
   End
   Begin VB.TextBox txtMessage 
      Height          =   1065
      Left            =   6270
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   7140
      Width           =   3945
   End
   Begin VB.Frame Frame2 
      Caption         =   "Display"
      Height          =   945
      Left            =   8700
      TabIndex        =   10
      Top             =   4380
      Width           =   1545
      Begin VB.OptionButton oView 
         Caption         =   "Minimum"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   570
         Width           =   1005
      End
      Begin VB.OptionButton oView 
         Caption         =   "In Stock"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "View"
      Height          =   945
      Left            =   8700
      TabIndex        =   9
      Top             =   5610
      Width           =   1545
      Begin VB.CheckBox chkShow 
         Caption         =   "Crossmatched"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   17
         Top             =   600
         Width           =   1305
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "In Free Stock"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   300
         Value           =   1  'Checked
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gOrder 
      Height          =   3945
      Left            =   90
      TabIndex        =   8
      Top             =   4260
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6959
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Product                                                                |^Suggest  |^In Stock  |^Min        "
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4065
      Left            =   90
      TabIndex        =   7
      Top             =   90
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   7170
      _Version        =   393216
      Cols            =   9
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      ForeColorFixed  =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"Frmstock.frx":08CE
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "&E-mail Order"
      Height          =   1095
      Left            =   10290
      Picture         =   "Frmstock.frx":097C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7140
      Width           =   735
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Order"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   6300
      Picture         =   "Frmstock.frx":0DBE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4260
      Width           =   735
   End
   Begin VB.CommandButton cmdFAX 
      Caption         =   "&FAX Order"
      Height          =   1095
      Left            =   11070
      Picture         =   "Frmstock.frx":1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7140
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add to Order"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   7080
      Picture         =   "Frmstock.frx":1642
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4260
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Order Form"
      Height          =   1095
      Left            =   11100
      Picture         =   "Frmstock.frx":1A84
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3060
      Width           =   735
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Print &Report"
      Height          =   1095
      Left            =   10320
      Picture         =   "Frmstock.frx":20EE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3060
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1095
      Left            =   10680
      Picture         =   "Frmstock.frx":2530
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5130
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   330
      TabIndex        =   18
      Top             =   8280
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FAX Message                FAX Number"
      Height          =   195
      Left            =   6330
      TabIndex        =   14
      Top             =   6900
      Width           =   2610
   End
End
Attribute VB_Name = "frmstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private InStock() As Integer

Private Sub FillDetails()

10    If oView(0) Then
20      chkShow(0) = 1
30      FillStockGrid
40    Else
50      chkShow(0) = False
60      chkShow(1) = False
70      FillMinGrid
80    End If

End Sub

Private Sub cmdadd_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim q As Integer
      Dim X As Integer
      Dim Y As Integer
      Dim BarCode As String
      Dim MinLevel As Integer

10    On Error GoTo cmdAdd_Click_Error

20    cmdRemove.Enabled = False

30    Y = g.Row
40    X = g.Col
50    If X * Y = 0 Then Exit Sub

60    g.Col = 0
70    BarCode = ProductBarCodeFor(g)
80    sql = "Select [" & _
            Choose(X, "OP", "AP", "BP", "ABP", "ON", "AN", "BN", "ABN") & _
            "] as ProdGroup from Minimum where " & _
            "BarCode = '" & BarCode & "'"
90    Set tb = New Recordset
100   RecOpenServerBB 0, tb, sql
110   If Not tb.EOF Then
120     MinLevel = tb!ProdGroup
130   Else
140     MinLevel = 0
150   End If

160   s = g & " " & _
          Choose(X, "O Pos", "A Pos", "B Pos", _
          "AB Pos", "O Neg", "A Neg", _
          "B Neg", "AB Neg") & vbTab
170   q = MinLevel - InStock(X, Y)
180   If q < 1 Then q = 1
190   s = s & Format(q) & vbTab & _
              Format(InStock(X, Y), "0") & vbTab & _
              Format(MinLevel, "0")
200   gOrder.AddItem s

210   cmdAdd.Enabled = False

220   Exit Sub

cmdAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmstock", "cmdAdd_Click", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdreport_Click()

      Dim X As Integer
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.Orientation = vbPRORLandscape
50    Printer.Font.Size = 8

60    If oView(0) Then
70      Printer.Print "Units in Free Stock ";
80    Else
90      Printer.Print "Minimum Stock Level ";
100   End If
110   Printer.Print Format(Now, "dd/mmm/yyyy")
120   Printer.Print

130   chkShow(1).Value = 0

140   For Y = 0 To g.Rows - 1
150     g.Row = Y
160     For X = 0 To g.Cols - 1
170       g.Col = X
180       Printer.Print g;
190       Printer.Print Tab(Choose(X + 1, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130));
200     Next
210     Printer.Print
220   Next
230   Printer.Print

240   Printer.FontSize = 10
250   Printer.Print "Suggested order."
260   Printer.Print

270   Printer.FontSize = 7
280   For Y = 0 To gOrder.Rows - 1
290     gOrder.Row = Y
300     For X = 0 To gOrder.Cols - 1
310       gOrder.Col = X
320       Printer.Print gOrder;
330       Printer.Print Tab(Choose(X + 1, 40, 60, 80, 100));
340     Next
350     Printer.Print
360   Next
370   Printer.Print

380   Printer.EndDoc

390   Printer.Orientation = vbPRORLandscape

400   For Each Px In Printers
410     If Px.DeviceName = OriginalPrinter Then
420       Set Printer = Px
430       Exit For
440     End If
450   Next

End Sub

Private Sub cmdRemove_Click()

10    cmdRemove.Enabled = False

20    If gOrder.Row = 0 Then Exit Sub

30    If gOrder.Rows = 2 Then
40      gOrder.AddItem ""
50      gOrder.RemoveItem 1
60    Else
70      gOrder.RemoveItem gOrder.Row
80    End If

End Sub

Private Sub FillMinGrid()

      Dim tb As Recordset
      Dim sql As String
      Dim BarCode As String
      Dim Y As Integer

10    On Error GoTo FillMinGrid_Error

20    For Y = 1 To g.Rows - 1
30      g.Row = Y
40      g.Col = 0
50      BarCode = ProductBarCodeFor(g)
60      sql = "Select * from Minimum where " & _
              "BarCode = '" & BarCode & "'"
70      Set tb = New Recordset
80      RecOpenServerBB 0, tb, sql
90      If Not tb.EOF Then
100       If tb!Op <> 0 Then g.Col = 1: g = tb!Op
110       If tb!AP <> 0 Then g.Col = 2: g = tb!AP
120       If tb!BP <> 0 Then g.Col = 3: g = tb!BP
130       If tb!ABP <> 0 Then g.Col = 4: g = tb!ABP
140       If tb!On <> 0 Then g.Col = 5: g = tb!On
150       If tb!AN <> 0 Then g.Col = 6: g = tb!AN
160       If tb!bn <> 0 Then g.Col = 7: g = tb!bn
170       If tb!ABN <> 0 Then g.Col = 8: g = tb!ABN
180     End If
190   Next

200   Exit Sub

FillMinGrid_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmstock", "FillMinGrid", intEL, strES, sql


End Sub

Private Sub FillOrderGrid()

      Dim s As String
      Dim BarCode As String
      Dim X As Integer
      Dim Y As Integer
      Dim MinLevel As Integer
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillOrderGrid_Error

20    gOrder.Rows = 2
30    gOrder.AddItem ""
40    gOrder.RemoveItem 1

50    For Y = 1 To g.Rows - 1
60      BarCode = ProductBarCodeFor(g.TextMatrix(Y, 0))
70      sql = "Select * from Minimum where " & _
              "BarCode = '" & BarCode & "'"
80      Set tb = New Recordset
90      RecOpenServerBB 0, tb, sql
100     If Not tb.EOF Then
110       With tb
120         For X = 1 To 8
130           MinLevel = Choose(X, !Op, !AP, !BP, !ABP, !On, !AN, !bn, !ABN)
140           If InStock(X, Y) < MinLevel Then
150             s = g.TextMatrix(Y, 0) & " " & _
                  Choose(X, "O Pos", "A Pos", "B Pos", "AB Pos", "O Neg", "A Neg", "B Neg", "AB Neg") & vbTab & _
                  Format(MinLevel - InStock(X, Y)) & vbTab & _
                  Format(InStock(X, Y), "0") & vbTab & _
                  Format(MinLevel)
160             gOrder.AddItem s
170           End If
180         Next
190       End With
200     End If
210   Next

220   If gOrder.Rows > 2 Then
230     gOrder.RemoveItem 1
240   End If

250   Exit Sub

FillOrderGrid_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmstock", "FillOrderGrid", intEL, strES, sql


End Sub

Private Sub FillStockGrid()

      Dim tb As Recordset
      Dim pc As String
      Dim sql As String
      Dim X As Integer
      Dim Y As Integer

10    On Error GoTo FillStockGrid_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from ProductList order by ListOrder"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF
90      g.AddItem tb!Wording
100     tb.MoveNext
110   Loop

120   ReDim InStock(1 To 8, 1 To g.Rows)
130   For X = 1 To 8
140     For Y = 1 To g.Rows
150       InStock(X, Y) = 0
160     Next
170   Next

180   For Y = 2 To g.Rows - 1
190     g.Col = 0
200     g.Row = Y
210     pc = ProductBarCodeFor(g)
220     For X = 1 To 8
230       sql = "SELECT distinct ISBT128 FROM latest where " & _
                "dateexpiry >= '" & Format(Now, "dd/mmm/yyyy HH:mm") & "' " & _
                "and barcode = '" & pc & "' " & _
                "and grouprh = '" & _
                Choose(X, "51", "62", "73", "84", "95", "06", "17", "28") & "' "
240       sql = sql & "and (event = 'C' Or event = 'R' "
250       If chkShow(1) = 1 Then
260         sql = sql & "or event = 'X' or event = 'I'"
270       End If
280       sql = sql & ")"
290       Set tb = New Recordset
300       RecOpenClientBB 0, tb, sql
310       If Not tb.EOF Then
320         g.Col = X
330         g.Row = Y
340         InStock(X, Y) = tb.RecordCount
350         g = Format(InStock(X, Y))
360       End If
370     Next
380   Next

390   If g.Rows > 2 Then
400     g.RemoveItem 1
410   End If

420   FillOrderGrid

430   Exit Sub

FillStockGrid_Error:

      Dim strES As String
      Dim intEL As Integer

440   intEL = Erl
450   strES = Err.Description
460   LogError "frmstock", "FillStockGrid", intEL, strES, sql


End Sub
Private Sub chkShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillDetails

20    cmdRemove.Enabled = False
30    cmdAdd.Enabled = False

End Sub


Private Sub Form_Load()

10    With cFAXNumber
20      .Clear
30        .AddItem "0214966905"
40        .AddItem "016674794"
50        .AddItem "04777930"

60      .ListIndex = 0
70    End With

80    FillStockGrid

End Sub

Private Sub g_Click()

      Dim MinLevel As Integer

10    cmdRemove.Enabled = False

20    If g.MouseRow = 0 Then Exit Sub
30    If g.MouseCol = 0 Then Exit Sub

40    If Not oView(1) Then Exit Sub

50    g.Enabled = False

60    MinLevel = Val(iBOX("Minimum Stock Level for " & vbCrLf & _
                      g.TextMatrix(g.Row, 0) & vbCrLf & _
                      g.TextMatrix(0, g.Col) & " ?", , _
                      g.TextMatrix(g.Row, g.Col)))
70    If TimedOut Then Unload Me: Exit Sub
    
80    g = MinLevel

90    SaveMinLevel

100   g.Enabled = True
110   cmdAdd.Enabled = True

End Sub

Private Sub gorder_Click()

      Dim OrderLevel As Integer
      Dim strTemp As String

10    strTemp = ""
20    cmdAdd.Enabled = False

30    If gOrder.MouseRow = 0 Then Exit Sub
40    If gOrder.TextMatrix(gOrder.Row, 0) = "" Then Exit Sub
50    If gOrder.MouseCol = 0 Then
60      cmdRemove.Enabled = True
70      Exit Sub
80    End If
90    If gOrder.MouseCol <> 1 Then Exit Sub

100   gOrder.Enabled = False

110   strTemp = iBOX("Quantity of  " & vbCrLf & _
                      gOrder.TextMatrix(gOrder.Row, 0) & vbCrLf & _
                      "to Order ?", , _
                      gOrder.TextMatrix(gOrder.Row, 1))
120   If TimedOut Then Unload Me: Exit Sub
    
130   If Len(strTemp) > 0 Then
140     OrderLevel = strTemp
150     gOrder.TextMatrix(gOrder.Row, 1) = OrderLevel
160     gOrder.Enabled = True
170   End If
    
180   gOrder.Enabled = True

End Sub

Private Sub SaveMinLevel()

      Dim tb As Recordset
      Dim sql As String
      Dim BarCode As String

10    On Error GoTo SaveMinLevel_Error

20    BarCode = ProductBarCodeFor(g.TextMatrix(g.Row, 0))

30    sql = "Select * from Minimum where " & _
            "BarCode = '" & BarCode & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If tb.EOF Then tb.AddNew
70    tb!BarCode = BarCode
80    tb!Op = Val(g.TextMatrix(g.Row, 1))
90    tb!AP = Val(g.TextMatrix(g.Row, 2))
100   tb!BP = Val(g.TextMatrix(g.Row, 3))
110   tb!ABP = Val(g.TextMatrix(g.Row, 4))
120   tb!On = Val(g.TextMatrix(g.Row, 5))
130   tb!AN = Val(g.TextMatrix(g.Row, 6))
140   tb!bn = Val(g.TextMatrix(g.Row, 7))
150   tb!ABN = Val(g.TextMatrix(g.Row, 8))
160   tb.Update

170   Exit Sub

SaveMinLevel_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmstock", "SaveMinLevel", intEL, strES, sql


End Sub

Private Sub oView_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillDetails

20    cmdRemove.Enabled = False
30    cmdAdd.Enabled = False

End Sub


