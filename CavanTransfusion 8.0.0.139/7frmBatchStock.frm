VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Batched Product Stock"
   ClientHeight    =   8475
   ClientLeft      =   240
   ClientTop       =   465
   ClientWidth     =   11235
   ControlBox      =   0   'False
   Icon            =   "7frmBatchStock.frx":0000
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
      Picture         =   "7frmBatchStock.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3060
      Width           =   1485
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   705
      Left            =   9510
      Picture         =   "7frmBatchStock.frx":0BD4
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
      Picture         =   "7frmBatchStock.frx":123E
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
Attribute VB_Name = "frmBatchStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SortOrder As Boolean

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    If optShowStock And optShowExpired Then
30      sql = "Select * from BatchProductList " & _
              "Where CurrentStock > 0 " & _
              "Order by DateReceived"
40    ElseIf optShowStock And optHideExpired Then
50      sql = "Select * from BatchProductList " & _
              "Where CurrentStock > 0 " & _
              "AND DATEDIFF(day, getdate(), DateExpiry) >= 0 " & _
              "Order by DateReceived"
60    ElseIf optShowAll And optShowExpired Then
70      sql = "Select * from BatchProductList " & _
              "Order by DateReceived"
80    ElseIf optShowAll And optHideExpired Then
90      sql = "Select * from BatchProductList " & _
              "WHERE DATEDIFF(day, getdate(), DateExpiry) >= 0 " & _
              "Order by DateReceived"
100   End If
110   Set tb = New Recordset
120   RecOpenServerBB 0, tb, sql

130   g.Rows = 2
140   g.AddItem ""
150   g.RemoveItem 1

160   Do While Not tb.EOF
170     s = tb!Product & vbTab & _
            tb!BatchNumber & vbTab & _
            tb!Group & vbTab & _
            tb!DateReceived & vbTab & _
            tb!DateExpiry & vbTab & _
            tb!CurrentStock
180     g.AddItem s
190     tb.MoveNext
200   Loop

210   If g.Rows > 2 Then
220     g.RemoveItem 1
230   End If

240   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmBatchStock", "FillG", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetLabelPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.FontSize = 9
50    Printer.Orientation = vbPRORPortrait

      '****Report heading
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print "                                               Batch Product Report"

      '****Report body

90    For i = 1 To 108
100       Printer.Print "_";
110   Next i
120   Printer.Print
130   Printer.Print FormatString("Product", 35, "|");
140   Printer.Print FormatString("Batch Number", 15, "|");
150   Printer.Print FormatString("Group", 5, "|");
160   Printer.Print FormatString("Date Rcvd", 16, "|");
170   Printer.Print FormatString("Expiry", 16, "|");
180   Printer.Print FormatString("Stock", 15, "|")
190   Printer.Font.Bold = False
200   For i = 1 To 108
210       Printer.Print "-";
220   Next i
230   Printer.Print
240   For Y = 1 To g.Rows - 1
250       Printer.Print FormatString(g.TextMatrix(Y, 0), 35, "|");
260       Printer.Print FormatString(g.TextMatrix(Y, 1), 15, "|");
270       Printer.Print FormatString(g.TextMatrix(Y, 2), 5, "|");
280       Printer.Print FormatString(g.TextMatrix(Y, 3), 16, "|");
290       Printer.Print FormatString(g.TextMatrix(Y, 4), 16, "|");
300       Printer.Print FormatString(g.TextMatrix(Y, 5), 15, "|")
 
310   Next


320   Printer.EndDoc

330   For Each Px In Printers
340     If Px.DeviceName = OriginalPrinter Then
350       Set Printer = Px
360       Exit For
370     End If
380   Next

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

10    If g.MouseRow = 0 Then
20      If InStr(g.TextMatrix(0, g.Col), "Date") <> 0 Then
30        g.Sort = 9
40      ElseIf SortOrder Then
50        g.Sort = flexSortGenericAscending
60      Else
70        g.Sort = flexSortGenericDescending
80      End If
90      SortOrder = Not SortOrder
100     Exit Sub
110   End If

120   With frmBatchHistory
130     .tBatchNumber = g.TextMatrix(g.Row, 1)
140     .Show 1
150   End With

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

