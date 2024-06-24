VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatListByProduct 
   Caption         =   "NetAcquire - Products Issued or Transfused"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14265
   Icon            =   "frmPatListByProduct.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbProduct 
      Height          =   315
      Left            =   5190
      TabIndex        =   10
      Text            =   "cmbProduct"
      Top             =   90
      Width           =   4770
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   825
      Left            =   12810
      Picture         =   "frmPatListByProduct.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2130
      Width           =   1245
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "&Start"
      Height          =   825
      Left            =   12780
      Picture         =   "frmPatListByProduct.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   570
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   12840
      Picture         =   "frmPatListByProduct.frx":1376
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1245
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   12810
      Picture         =   "frmPatListByProduct.frx":19E0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3030
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6615
      Left            =   90
      TabIndex        =   1
      Top             =   510
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmPatListByProduct.frx":1CEA
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
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   183631873
      CurrentDate     =   37509
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   3120
      TabIndex        =   6
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   183631873
      CurrentDate     =   37509
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   105
      TabIndex        =   12
      Top             =   7260
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   4590
      TabIndex        =   11
      Top             =   150
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Between"
      Height          =   195
      Left            =   780
      TabIndex        =   9
      Top             =   150
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "and"
      Height          =   195
      Left            =   2790
      TabIndex        =   8
      Top             =   150
      Width           =   270
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   12810
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "frmPatListByProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SortOrder As Boolean
Private Sub ClearG()

10  g.Rows = 2
20  g.AddItem ""
30  g.RemoveItem 1

End Sub

Private Sub FillProducts()

    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo FillProducts_Error

20  cmbProduct.Clear

30  sql = "SELECT Wording FROM ProductList WHERE " & _
          "BarCode IN " & _
        "  ( SELECT DISTINCT BarCode FROM Latest WHERE " & _
        "    DateTime BETWEEN '" & Format$(dtFrom, "Long Date") & "' " & _
        "    AND '" & Format$(dtTo, "Long Date") & " 23:59:59' )"
40  Set tb = New Recordset
50  RecOpenServerBB 0, tb, sql
60  Do While Not tb.EOF
70      cmbProduct.AddItem tb!Wording & ""
80      tb.MoveNext
90  Loop

100 Exit Sub

FillProducts_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmPatListByProduct", "FillProducts", intEL, strES, sql


End Sub



Private Sub cmbProduct_Click()

10  ClearG

End Sub


Private Sub cmdCancel_Click()

10  Unload Me

End Sub

Private Sub cmdPrint_Click()

    Dim Y As Integer
    Dim OriginalPrinter As String
    Dim Px As Printer
    Dim i As Integer

10  OriginalPrinter = Printer.DeviceName
20  If Not SetFormPrinter() Then
30      iMsg "Can't set Form Printer!"
40      If TimedOut Then Unload Me: Exit Sub
50      Exit Sub
60  End If

70  Printer.Orientation = vbPRORLandscape

80  Printer.Font.Name = "Courier New"
90  Printer.Font.Size = 10
100 Printer.Font.Bold = True

110 Printer.Print
120 Printer.Print FormatString("Products Issued Or Transfused", 140, , AlignCenter)
130 Printer.Print FormatString(cmbProduct, 140, , AlignCenter)
140 Printer.Print FormatString("Between " & dtFrom & " And " & dtTo, 140, , AlignCenter)
150 Printer.Print

160 Printer.Font.Size = 9

170 For i = 1 To 152
180     Printer.Print "-";
190 Next i
200 Printer.Print

210 Printer.Print FormatString("Date/Time", 16, "|", AlignCenter);
220 Printer.Print FormatString("Patient Name", 40, "|", AlignCenter);
230 Printer.Print FormatString("D.o.B.", 10, "|", AlignCenter);
240 Printer.Print FormatString("Chart", 10, "|", AlignCenter);
250 Printer.Print FormatString("AandE", 10, "|", AlignCenter);
260 Printer.Print FormatString("Pat.", 9, "|", AlignCenter);
270 Printer.Print FormatString("Unit", 16, "|", AlignCenter);
280 Printer.Print FormatString("Group", 5, "|", AlignCenter);
290 Printer.Print FormatString("Date Exp.", 10, "|", AlignCenter);
300 Printer.Print FormatString("Operator", 16, "|", AlignCenter)

310 For i = 1 To 152
320     Printer.Print "-";
330 Next i
340 Printer.Print

350 Printer.Font.Bold = False
360 For Y = 1 To g.Rows - 1
370     Printer.Print FormatString(Format$(g.TextMatrix(Y, 0), "Short Date"), 16, "|");
380     Printer.Print FormatString(g.TextMatrix(Y, 1), 40, "|");
390     Printer.Print FormatString(g.TextMatrix(Y, 2), 10, "|");
400     Printer.Print FormatString(g.TextMatrix(Y, 3), 10, "|");
410     Printer.Print FormatString(g.TextMatrix(Y, 4), 10, "|");
420     Printer.Print FormatString(g.TextMatrix(Y, 5), 9, "|");
430     Printer.Print FormatString(g.TextMatrix(Y, 6), 16, "|");
440     Printer.Print FormatString(g.TextMatrix(Y, 7), 5, "|");
450     Printer.Print FormatString(g.TextMatrix(Y, 8), 10, "|");
460     Printer.Print FormatString(g.TextMatrix(Y, 9), 16, "|")

470 Next


480 Printer.EndDoc

490 For Each Px In Printers
500     If Px.DeviceName = OriginalPrinter Then
510         Set Printer = Px
520         Exit For
530     End If
540 Next


End Sub

Private Sub cmdStart_Click()

    Dim sql As String
    Dim tb As Recordset
    Dim tbDem As Recordset
    Dim BarCode As String
    Dim s As String
    Dim fGroup As String
    Dim DoB As String
    Dim AandE As String

10  On Error GoTo cmdstart_Click_Error

20  ClearG

30  g.Visible = False

40  BarCode = ProductBarCodeFor(cmbProduct)

50  sql = "SELECT LabNumber, PatName, DateTime, PatID, Number, ISBT128, GroupRh, Operator, DateExpiry " & _
          "FROM Latest WHERE " & _
          "(Event ='S' OR Event = 'I') " & _
          "AND DateTime BETWEEN '" & Format$(dtFrom, "Long Date") & "' " & _
        "                 AND '" & Format$(dtTo, "Long Date") & " 23:59:59' " & _
          "AND BarCode = '" & BarCode & "'"
60  Set tb = New Recordset
70  RecOpenServerBB 0, tb, sql
80  Do While Not tb.EOF
90      sql = "SELECT COALESCE(DoB,'') AS DoB, fGroup, AandE FROM PatientDetails WHERE " & _
              "LabNumber = '" & tb!LabNumber & "'"
100     Set tbDem = New Recordset
110     RecOpenServerBB 0, tbDem, sql
120     If Not tbDem.EOF Then
130         If IsDate(tbDem!DoB) Then
140             DoB = Format$(tbDem!DoB, "Short Date")
150         Else
160             DoB = ""
170         End If
180         fGroup = tbDem!fGroup & ""
190         AandE = tbDem!AandE & ""
200     Else
210         DoB = ""
220         fGroup = ""
230         AandE = ""
240     End If
250     s = Format$(tb!DateTime, "dd/MM/yyyy HH:nn:ss") & vbTab & _
            tb!PatName & vbTab & _
            DoB & vbTab & _
            tb!Patid & vbTab & _
            AandE & vbTab & _
            fGroup & vbTab
260     If Len(Trim((tb!ISBT128 & ""))) > 0 Then
270         s = s & tb!ISBT128 & ""
280     Else
290         s = s & tb!Number & ""
300     End If
310     s = s & vbTab & _
            Bar2Group(tb!GroupRh & "") & vbTab & _
            Format(tb!DateExpiry, "dd/mm/yyyy HH:mm") & vbTab & _
            TechnicianNameForCode(tb!Operator & "")
320     g.AddItem s
330     tb.MoveNext
340 Loop

350 If g.Rows > 2 Then
360     g.RemoveItem 1
370 End If

380 g.Visible = True

390 Exit Sub

cmdstart_Click_Error:

    Dim strES As String
    Dim intEL As Integer

400 intEL = Erl
410 strES = Err.Description
420 LogError "frmPatListByProduct", "cmdstart_Click", intEL, strES, sql
430 g.Visible = True

End Sub

Private Sub cmdXL_Click()
    Dim strHeading As String
10  strHeading = "Products Issued Or Transfused" & vbCr
20  strHeading = strHeading & cmbProduct & vbCr
30  strHeading = strHeading & "Between " & dtFrom & " And " & dtTo & vbCr
40  strHeading = strHeading & " " & vbCr
50  ExportFlexGrid g, Me, strHeading

End Sub


Private Sub dtFrom_CloseUp()

10  FillProducts
20  ClearG

End Sub


Private Sub dtTo_CloseUp()

10  FillProducts
20  ClearG

End Sub


Private Sub Form_Load()

10  dtFrom = Format(Now - 30, "dd/mm/yyyy")
20  dtTo = Format(Now, "dd/mm/yyyy")

30  FillProducts

End Sub

Private Sub g_Click()

10  If g.MouseRow = 0 Then
20      If InStr(g.TextMatrix(0, g.MouseCol), "Date") = 0 Then
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

10  GC = g.col

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
