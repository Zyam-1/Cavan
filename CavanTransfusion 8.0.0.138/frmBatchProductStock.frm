VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchProductStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Batched Product Stock"
   ClientHeight    =   8475
   ClientLeft      =   240
   ClientTop       =   465
   ClientWidth     =   11235
   ControlBox      =   0   'False
   Icon            =   "frmBatchProductStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Expired"
      Height          =   915
      Left            =   9510
      TabIndex        =   8
      Top             =   1590
      Width           =   1485
      Begin VB.OptionButton optHideExpired 
         Caption         =   "Hide"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   570
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton optShowExpired 
         Caption         =   "Show"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show"
      Height          =   915
      Left            =   9510
      TabIndex        =   7
      Top             =   510
      Width           =   1485
      Begin VB.OptionButton optShowAll 
         Caption         =   "Show All"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   540
         Width           =   915
      End
      Begin VB.OptionButton optShowStock 
         Caption         =   "Only in Stock"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   705
      Left            =   9510
      Picture         =   "frmBatchProductStock.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3060
      Width           =   1485
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   705
      Left            =   9510
      Picture         =   "frmBatchProductStock.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4350
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7485
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   13203
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
      FormatString    =   "<Product               |<Batch Number              |^Group |^Date Recd (latest)|^Expiry Date      |^Current Stock "
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
      Height          =   705
      Left            =   9510
      Picture         =   "frmBatchProductStock.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7410
      Width           =   1485
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   8130
      Width           =   9195
      _ExtentX        =   16219
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
      Left            =   9510
      TabIndex        =   5
      Top             =   3780
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on Batch Number to show History"
      Height          =   495
      Left            =   1830
      TabIndex        =   2
      Top             =   90
      Width           =   2145
   End
End
Attribute VB_Name = "frmBatchProductStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SortOrder As Boolean

Private Sub FillG()

Dim s As String
Dim BPs As New BatchProducts
Dim BP As BatchProduct
Dim CurrentProduct As String
Dim CurrentBatch As String

10    On Error GoTo FillG_Error

20    If optShowStock And optShowExpired Then
30      BPs.LoadStockWithExpired
40    ElseIf optShowStock And optHideExpired Then
50      BPs.LoadStockNotExpired
60    ElseIf optShowAll And optShowExpired Then
70      BPs.LoadStockAll
80    ElseIf optShowAll And optHideExpired Then
90      BPs.LoadStockAllNotExpired
100   End If

110   g.Rows = 2
120   g.AddItem ""
130   g.RemoveItem 1

140   CurrentProduct = ""
150   CurrentBatch = ""

160   For Each BP In BPs
170     If BP.Product <> CurrentProduct Or BP.BatchNumber <> CurrentBatch Then
180       CurrentProduct = BP.Product
190       CurrentBatch = BP.BatchNumber
200       s = BP.Product & vbTab & _
        BP.BatchNumber & vbTab & _
        BP.UnitGroup & vbTab & _
        BP.DateReceived & vbTab & _
        BP.DateExpiry & vbTab
210       If optHideExpired Then
220   s = s & BPs.CountProductBatchInStockNotExpired(BP.Product, BP.BatchNumber)
230       Else
240   s = s & BPs.CountProductBatchInStock(BP.Product, BP.BatchNumber)
250       End If
260       g.AddItem s
270     End If
280   Next

290   If g.Rows > 2 Then
300     g.RemoveItem 1
310   End If

320   Exit Sub

FillG_Error:

Dim strES As String
Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmBatchProductStock", "FillG", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    On Error GoTo cmdPrint_Click_Error

20    OriginalPrinter = Printer.DeviceName

30    If Not SetLabelPrinter() Then Exit Sub

40    Printer.FontName = "Courier New"
50    Printer.FontSize = 9
60    Printer.Orientation = vbPRORPortrait

      '****Report heading
70    Printer.Font.Bold = True
80    Printer.Print
90    Printer.Print "                                               Batch Product Report"

      '****Report body

100   For i = 1 To 108
110       Printer.Print "_";
120   Next i
130   Printer.Print
140   Printer.Print FormatString("Product", 35, "|");
150   Printer.Print FormatString("Batch Number", 15, "|");
160   Printer.Print FormatString("Group", 5, "|");
170   Printer.Print FormatString("Date Rcvd", 16, "|");
180   Printer.Print FormatString("Expiry", 16, "|");
190   Printer.Print FormatString("Stock", 15, "|")
200   Printer.Font.Bold = False
210   For i = 1 To 108
220       Printer.Print "-";
230   Next i
240   Printer.Print
250   For Y = 1 To g.Rows - 1
260       Printer.Print FormatString(g.TextMatrix(Y, 0), 35, "|");
270       Printer.Print FormatString(g.TextMatrix(Y, 1), 15, "|");
280       Printer.Print FormatString(g.TextMatrix(Y, 2), 5, "|");
290       Printer.Print FormatString(g.TextMatrix(Y, 3), 16, "|");
300       Printer.Print FormatString(g.TextMatrix(Y, 4), 16, "|");
310       Printer.Print FormatString(g.TextMatrix(Y, 5), 15, "|")
 
320   Next


330   Printer.EndDoc

340   For Each Px In Printers
350     If Px.DeviceName = OriginalPrinter Then
360       Set Printer = Px
370       Exit For
380     End If
390   Next

400   Exit Sub

cmdPrint_Click_Error:

Dim strES As String
Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "frmBatchProductStock", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdXL_Click()

10    ExportFlexGrid g, Me

End Sub

Private Sub Form_Load()

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        FillG
      '**************************************

End Sub

Private Sub g_Click()

10    On Error GoTo g_Click_Error

20    If g.MouseRow = 0 Then
30      If InStr(g.TextMatrix(0, g.Col), "Date") <> 0 Then
40        g.Sort = 9
50      ElseIf SortOrder Then
60        g.Sort = flexSortGenericAscending
70      Else
80        g.Sort = flexSortGenericDescending
90      End If
100     SortOrder = Not SortOrder
110     Exit Sub
120   End If

130   With frmBatchProductHistory
140     .BatchNumber = g.TextMatrix(g.Row, 1)
150     .Show 1
160   End With

170   Exit Sub

g_Click_Error:

Dim strES As String
Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmBatchProductStock", "g_Click", intEL, strES

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
20      Cmp = 0
30      Exit Sub
40    End If

50    If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
60      Cmp = 0
70      Exit Sub
80    End If

90    d1 = Format(g.TextMatrix(Row1, g.Col), "dd/mmm/yyyy hh:mm:ss")
100   d2 = Format(g.TextMatrix(Row2, g.Col), "dd/mmm/yyyy hh:mm:ss")

110   If SortOrder Then
120     Cmp = Sgn(DateDiff("D", d1, d2))
130   Else
140     Cmp = Sgn(DateDiff("D", d2, d1))
150   End If

End Sub


Private Sub optHideExpired_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillG

End Sub

Private Sub optShowAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillG

End Sub

Private Sub optShowExpired_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillG

End Sub

Private Sub optShowStock_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillG

End Sub

