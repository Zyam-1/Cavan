VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFreeStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Free Stock"
   ClientHeight    =   6420
   ClientLeft      =   495
   ClientTop       =   1680
   ClientWidth     =   9585
   Icon            =   "frmFreeStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   6525
      Picture         =   "frmFreeStock.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print History"
      Top             =   5490
      Width           =   1080
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   705
      Left            =   1950
      Picture         =   "frmFreeStock.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1245
   End
   Begin VB.CheckBox chkExpired 
      Caption         =   "Include Expired"
      Height          =   225
      Left            =   330
      TabIndex        =   2
      Top             =   5520
      Width           =   1485
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   8100
      TabIndex        =   1
      Top             =   5490
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4635
      Left            =   270
      TabIndex        =   0
      Top             =   330
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   8176
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Unit                       |<Expiry                     |<Group       |<Product                             |<Status             "
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
      Left            =   270
      TabIndex        =   6
      Top             =   4980
      Width           =   9075
      _ExtentX        =   16007
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
      Left            =   3210
      TabIndex        =   4
      Top             =   5550
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmFreeStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG()

    Dim sn As Recordset
    Dim sql As String
    Dim s As String

10  On Error GoTo FillG_Error

20  g.Rows = 2
30  g.AddItem ""
40  g.RemoveItem 1

50  sql = "Select P.Wording, L.* from Latest as L, ProductList as P where " & _
          "(event = 'R' or event = 'C' or event = 'X' or event = 'P') " & _
          "and P.BarCode = L.BarCode "
60  If chkExpired = 0 Then
70      sql = sql & "and DateExpiry > '" & Format(Now, "dd/mmm/yyyy hh:mm") & "' "
80  End If
90  sql = sql & "order by GroupRh, DateExpiry, ISBT128"

100 Set sn = New Recordset
110 RecOpenClientBB 0, sn, sql
120 If sn.EOF Then
130     iMsg "No Records found.", vbInformation
140     If TimedOut Then Unload Me: Exit Sub
150     Exit Sub
160 End If

170 g.Visible = False
180 Do While Not sn.EOF
190     s = sn!ISBT128 & "" & vbTab & _
            Format(sn!DateExpiry, "dd/mmm/yyyy hh:mm") & vbTab & _
            Bar2Group(sn!GroupRh & "") & vbTab & _
            sn!Wording & vbTab
200     Select Case UCase$(sn!Event)
        Case "R", "C": s = s & "Free"
210     Case "X": s = s & "XMatched"
220     Case "P": s = s & "Pending"
230     End Select
240     g.AddItem s
250     If DateDiff("n", Format(sn!DateExpiry, "dd/mmm/yyyy hh:mm"), Format(Now, "dd/mmm/yyyy hh:mm")) > 0 Then
260         g.Row = g.Rows - 1
270         g.Col = 1
280         g.CellBackColor = vbRed
290         g.CellForeColor = vbYellow
300     End If
310     sn.MoveNext
320 Loop

330 If g.Rows > 2 Then
340     g.RemoveItem 1
350 End If
360 g.Visible = True

370 Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

380 intEL = Erl
390 strES = Err.Description
400 LogError "frmFreeStock", "FillG", intEL, strES, sql


End Sub

Private Sub btnCancel_Click()

10  Unload Me

End Sub


Private Sub chkExpired_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10  FillG

End Sub


Private Sub cmdPrint_Click()

    Dim Y As Integer
    Dim OriginalPrinter As String
    Dim Px As Printer

10  On Error GoTo cmdPrint_Click_Error

20  OriginalPrinter = Printer.DeviceName
30  If Not SetFormPrinter() Then Exit Sub

40  Printer.Print
50  Printer.Font.Size = 12
60  Printer.Print "Free Stock : "; Format(Now, "dd/mmm/yyyy hh:mm");
70  Printer.Print
80  Printer.Print
90  Printer.Font.Size = 8
100 Printer.Print "Unit"; Tab(20); "Expiry"; Tab(35); "Group"; Tab(45); "Product"; Tab(100); "Status"
110 Printer.Print
120 For Y = 1 To g.Rows - 1
130     Printer.Print g.TextMatrix(Y, 0);    'unit
140     Printer.Print Tab(20); g.TextMatrix(Y, 1);    'expiry
150     Printer.Print Tab(35); g.TextMatrix(Y, 2);    'group
160     Printer.Print Tab(45); Left$(g.TextMatrix(Y, 3), 40);    'Product
170     Printer.Print Tab(100); g.TextMatrix(Y, 4)    'Status
180 Next

190 Printer.Print

200 Printer.EndDoc

210 For Each Px In Printers
220     If Px.DeviceName = OriginalPrinter Then
230         Set Printer = Px
240         Exit For
250     End If
260 Next


270 Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

280 intEL = Erl
290 strES = Err.Description
300 LogError "frmFreeStock", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdXL_Click()

10  ExportFlexGrid g, Me

End Sub



Private Sub Form_Load()

'*****NOTE
'FillG might be dependent on many components so for any future
'update in code try to keep FillG on bottom most line of form load.
10  FillG
    '**************************************
End Sub


Private Sub g_Click()

10  If g.MouseRow = 0 Then
20      If g.Col <> 1 Then
30          If SortOrder Then
40              g.Sort = flexSortGenericAscending
50          Else
60              g.Sort = flexSortGenericDescending
70          End If
80      Else
90          g.Sort = 9
100     End If
110     SortOrder = Not SortOrder
120 End If

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

    Dim d1 As String
    Dim d2 As String
    Dim GC As Integer

10  GC = g.Col

20  If Not IsDate(g.TextMatrix(Row1, GC)) Then
30      Cmp = 0
40      Exit Sub
50  End If

60  If Not IsDate(g.TextMatrix(Row2, GC)) Then
70      Cmp = 0
80      Exit Sub
90  End If

100 d1 = Format(g.TextMatrix(Row1, GC), "dd/mmm/yyyy hh:mm:ss")
110 d2 = Format(g.TextMatrix(Row2, GC), "dd/mmm/yyyy hh:mm:ss")

120 If SortOrder Then
130     Cmp = Sgn(DateDiff("D", d1, d2))
140 Else
150     Cmp = Sgn(DateDiff("D", d2, d1))
160 End If

End Sub

