VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmABOAntiDBloodGroupControlHistory 
   Caption         =   "ABO-Anti D Blood Group Control History"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   Icon            =   "frmABOAntiDBloodGroupControlHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOperator 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   960
      TabIndex        =   21
      Top             =   7860
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "ABO/D + Reverse Grouping Card"
      Height          =   1530
      Left            =   180
      TabIndex        =   13
      Top             =   1275
      Width           =   2715
      Begin VB.TextBox tCardLot 
         Height          =   288
         Left            =   90
         TabIndex        =   15
         Top             =   570
         Width           =   2490
      End
      Begin VB.TextBox tCardExpiry 
         Height          =   288
         Left            =   90
         TabIndex        =   14
         Top             =   1095
         Width           =   1236
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Batch No.:"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   915
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cell Panel"
      Height          =   1545
      Left            =   3060
      TabIndex        =   8
      Top             =   1275
      Width           =   2715
      Begin VB.TextBox tCellsLot 
         Height          =   288
         Left            =   90
         TabIndex        =   10
         Top             =   555
         Width           =   2520
      End
      Begin VB.TextBox tCellsExpiry 
         Height          =   288
         Left            =   90
         TabIndex        =   9
         Top             =   1125
         Width           =   1236
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lot No.:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   375
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   945
         Width           =   870
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Reactions"
      Height          =   3510
      Left            =   195
      TabIndex        =   6
      Top             =   2940
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   2985
         Left            =   165
         TabIndex        =   7
         Top             =   315
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   5265
         _Version        =   393216
         Cols            =   5
         RowHeightMin    =   400
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   0
         FormatString    =   $"frmABOAntiDBloodGroupControlHistory.frx":08CA
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
   End
   Begin VB.TextBox txtComment 
      Height          =   1035
      Left            =   195
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6720
      Width           =   6255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   6045
      Picture         =   "frmABOAntiDBloodGroupControlHistory.frx":0925
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   1035
   End
   Begin VB.ListBox lstTime 
      Height          =   780
      IntegralHeight  =   0   'False
      ItemData        =   "frmABOAntiDBloodGroupControlHistory.frx":0F8F
      Left            =   1665
      List            =   "frmABOAntiDBloodGroupControlHistory.frx":0F9F
      TabIndex        =   1
      Top             =   180
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   765
      Left            =   4845
      Picture         =   "frmABOAntiDBloodGroupControlHistory.frx":0FBF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   375
      Left            =   195
      TabIndex        =   3
      Top             =   360
      Width           =   1425
      _ExtentX        =   2514
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
      Format          =   183238657
      CurrentDate     =   37735
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   195
      TabIndex        =   19
      Top             =   8190
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Operator"
      Height          =   195
      Left            =   225
      TabIndex        =   20
      Top             =   7890
      Width           =   615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   225
      TabIndex        =   18
      Top             =   6480
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Time to View"
      Height          =   255
      Left            =   3075
      TabIndex        =   4
      Top             =   480
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2595
      Picture         =   "frmABOAntiDBloodGroupControlHistory.frx":1629
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmABOAntiDBloodGroupControlHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
10    Unload Me
End Sub



Private Sub cmdPrint_Click()
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.FontSize = 9
50    Printer.Orientation = vbPRORPortrait

      '****Report heading
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print "                                  ABO-Anti D Blood Group Control History";
90    Printer.Font.Bold = False
100   Printer.Print "       Page 1 of 1"
110   Printer.Font.Bold = True
120   Printer.Print "DateTime:    "; FormatString(dt, 10); FormatString(lstTime, 8)
130   Printer.Print "Batch #:     "; FormatString(tCardLot, 10); "          Expiry Date:  "; FormatString(tCardExpiry, 16)
140   Printer.Print "Lot Number:  "; FormatString(tCellsLot, 10); "          Expiry Date:  "; FormatString(tCellsExpiry, 16)
150   Printer.Print

      '****Report body


160   For i = 1 To 108
170       Printer.Print "_";
180   Next i
190   Printer.Print
200   Printer.Print FormatString(" ", 23, "|");
210   Printer.Print FormatString("A2", 20, "|");
220   Printer.Print FormatString("B", 20, "|");
230   Printer.Print FormatString("R2R2", 20, "|");
240   Printer.Print FormatString("rr", 20, "|")
250   Printer.Font.Bold = False
260   For i = 1 To 108
270       Printer.Print "-";
280   Next i
290   Printer.Print
300   For Y = 1 To g.Rows - 1
310       Printer.Print FormatString(g.TextMatrix(Y, 0), 23, "|");
320       Printer.Print FormatString(g.TextMatrix(Y, 1), 20, "|");
330       Printer.Print FormatString(g.TextMatrix(Y, 2), 20, "|");
340       Printer.Print FormatString(g.TextMatrix(Y, 2), 20, "|");
350       Printer.Print FormatString(g.TextMatrix(Y, 4), 20, "|")
 
360   Next


370   Printer.EndDoc



380   For Each Px In Printers
390     If Px.DeviceName = OriginalPrinter Then
400       Set Printer = Px
410       Exit For
420     End If
430   Next
End Sub

Private Sub dt_CloseUp()
10    FillList
End Sub

Private Sub Form_Load()
      Dim s As String

10    dt = Format(Now, "dd/mm/yyyy")

20    g.Font.Bold = True

30    s = "Anti-A" & vbTab
40    g.AddItem s

50    s = "Anti-B" & vbTab
60    g.AddItem s

70    s = "Anti-D" & vbTab
80    g.AddItem s

90    s = "Anti-D" & vbTab & vbTab & vbTab & vbTab
100   g.AddItem s

110   s = "A1" & vbTab & vbTab & vbTab & vbTab
120   g.AddItem s

130   s = "B" & vbTab & vbTab & vbTab & vbTab
140   g.AddItem s

150   g.RemoveItem 1
  
160   With g
170     .row = 1
180     .col = 3
190     .CellBackColor = &H8000000F
200     .Text = "X"
210     .col = 4
220     .CellBackColor = &H8000000F
230     .Text = "X"
240     .row = 2
250     .col = 3
260     .CellBackColor = &H8000000F
270     .Text = "X"
280     .col = 4
290     .CellBackColor = &H8000000F
300     .Text = "X"

310     .row = 3
320     .col = 1
330     .CellBackColor = &H8000000F
340     .Text = "X"
350     .col = 2
360     .CellBackColor = &H8000000F
370     .Text = "X"
380     .row = 4
390     .col = 1
400     .CellBackColor = &H8000000F
410     .Text = "X"
420     .col = 2
430     .CellBackColor = &H8000000F
440     .Text = "X"
  
450     .row = 5
460     .col = 1
470     .CellBackColor = &H8000000F
480     .Text = "X"
490     .col = 2
500     .CellBackColor = &H8000000F
510     .Text = "X"
520     .col = 3
530     .CellBackColor = &H8000000F
540     .Text = "X"
550     .col = 4
560     .CellBackColor = &H8000000F
570     .Text = "X"
  
580     .row = 6
590     .col = 1
600     .CellBackColor = &H8000000F
610     .Text = "X"
620     .col = 2
630     .CellBackColor = &H8000000F
640     .Text = "X"
650     .col = 3
660     .CellBackColor = &H8000000F
670     .Text = "X"
680     .col = 4
690     .CellBackColor = &H8000000F
700     .Text = "X"

710   End With

720   FillList

End Sub

Private Sub FillList()
      Dim sql As String
      Dim tb As Recordset

10    lstTime.Clear

20    sql = "Select distinct DateTime from ABOQC " & _
            "where DateTime between '" & _
            Format(dt, "dd/mmm/yyyy 00:00") & "' and '" & _
            Format(dt, "dd/mmm/yyyy 23:59") & "' " & _
            "order by DateTime asc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    Do While Not tb.EOF
60      lstTime.AddItem Format(tb!DateTime, "hh:mm:ss")
70      tb.MoveNext
80    Loop


End Sub



Private Sub lstTime_Click()
10    LoadDetails
End Sub

Private Sub LoadDetails()

      Dim tb As Recordset
      Dim sql As String

10    tCardLot = ""
20    tCellsLot = ""
30    tCardExpiry = ""
40    tCellsExpiry = ""

50    g.TextMatrix(1, 1) = ""
60    g.TextMatrix(1, 2) = ""
70    g.TextMatrix(2, 1) = ""
80    g.TextMatrix(2, 2) = ""

90    g.TextMatrix(3, 3) = ""
100   g.TextMatrix(3, 4) = ""
110   g.TextMatrix(4, 3) = ""
120   g.TextMatrix(4, 4) = ""


130   sql = "Select * from ABOQC where " & _
            "DateTime = '" & Format(dt & " " & lstTime, "dd/mmm/yyyy hh:mm:ss") & "'"
140   Set tb = New Recordset
150   RecOpenServerBB 0, tb, sql
160   If Not tb.EOF Then
      '  If lblOperator = "" Then
      '    lblOperator = tb!Operator & ""
      '  End If
  
170     tCardLot = tb!CardLot & ""
180     tCellsLot = tb!CellsLot & ""
190     tCardExpiry = Format(tb!CardExpiry, "dd/mm/yyyy")
200     tCellsExpiry = Format(tb!CellsExpiry, "dd/mm/yyyy")
210     txtComment = tb!Comment & ""
220     txtOperator = tb!Operator & ""
230     g.TextMatrix(1, 1) = tb!r11 & ""
240     g.TextMatrix(1, 2) = tb!r12 & ""
250     g.TextMatrix(2, 1) = tb!r21 & ""
260     g.TextMatrix(2, 2) = tb!r22 & ""
  
270     g.TextMatrix(3, 3) = tb!r33 & ""
280     g.TextMatrix(3, 4) = tb!r34 & ""
290     g.TextMatrix(4, 3) = tb!r43 & ""
300     g.TextMatrix(4, 4) = tb!r44 & ""

310   End If


End Sub


