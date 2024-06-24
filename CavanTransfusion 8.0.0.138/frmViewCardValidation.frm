VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmViewCardValidation 
   Caption         =   "NetAcquire - View Card Validation"
   ClientHeight    =   7530
   ClientLeft      =   675
   ClientTop       =   495
   ClientWidth     =   9045
   Icon            =   "frmViewCardValidation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   9045
   Begin VB.CommandButton bprint 
      Caption         =   "Print"
      Height          =   915
      Left            =   5730
      Picture         =   "frmViewCardValidation.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   915
      Left            =   3570
      Picture         =   "frmViewCardValidation.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   330
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   915
      Left            =   7800
      Picture         =   "frmViewCardValidation.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "bCancel"
      Top             =   330
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1065
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3045
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   675
         Left            =   1830
         Picture         =   "frmViewCardValidation.frx":18A8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   945
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   38373
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   38373
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5745
      Left            =   240
      TabIndex        =   0
      Top             =   1410
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   10134
      _Version        =   393216
      Cols            =   7
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
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "<Date/Time                  |<Sample ID |<ABO Batch #       |<ABO Expiry |<LISS Batch #         |<LISS Expiry |<Operator       "
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   7200
      Width           =   8595
      _ExtentX        =   15161
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
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      Top             =   630
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "frmViewCardValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    With g
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60    End With

70    sql = "Select * from CardValidation where " & _
            "DateTime between '" & Format$(dtFrom, "dd/mmm/yyyy") & _
            "' and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59'"
80    Set tb = New Recordset
90    RecOpenServerBB 0, tb, sql

100   Do While Not tb.EOF
110     s = Format$(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!SampleID & vbTab & _
            tb!ABOBatch & vbTab
120     If IsDate(tb!ABOExpiry) Then
130       s = s & Format$(tb!ABOExpiry, "dd/mm/yy")
140     End If
150     s = s & vbTab & tb!LISSBatch & vbTab
160     If IsDate(tb!LISSExpiry) Then
170       s = s & Format$(tb!LISSExpiry, "dd/mm/yy")
180     End If
190     s = s & vbTab & tb!Operator & ""
  
200     g.AddItem s
  
210     tb.MoveNext
220   Loop

230   If g.Rows > 2 Then
240     g.RemoveItem 1
250   End If

260   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmViewCardValidation", "FillG", intEL, strES, sql

  
End Sub

Private Sub bprint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.FontSize = 10
50    Printer.Font.Bold = True
60    Printer.Orientation = vbPRORPortrait

      '****Report heading

70    Printer.Print
80    Printer.Print FormatString("Card Validation", 99, , AlignCenter)
90    Printer.Print FormatString("Between " & Format(dtFrom, "dd/mmm/yyyy") & " - " & Format(dtTo, "dd/mmm/yyyy"), 99, , AlignCenter)
      '****Report body

100   Printer.Font.Size = 9
110   For i = 1 To 108
120       Printer.Print "-";
130   Next i
140   Printer.Print

150   Printer.Print FormatString("", 0, "|");
160   Printer.Print FormatString("Date Time", 18, "|");
170   Printer.Print FormatString("SampleID", 8, "|");
180   Printer.Print FormatString("ABO Batch", 14, "|");
190   Printer.Print FormatString("ABO Exp", 10, "|");
200   Printer.Print FormatString("LISS Batch", 14, "|");
210   Printer.Print FormatString("LISS Exp", 10, "|");
220   Printer.Print FormatString("Op", 26, "|")
230   Printer.Font.Bold = False
240   For i = 1 To 108
250       Printer.Print "-";
260   Next i
270   Printer.Print
280   For Y = 1 To g.Rows - 1
290       Printer.Print FormatString("", 0, "|");
300       Printer.Print FormatString(g.TextMatrix(Y, 0), 18, "|");
310       Printer.Print FormatString(g.TextMatrix(Y, 1), 8, "|");
320       Printer.Print FormatString(g.TextMatrix(Y, 2), 14, "|");
330       Printer.Print FormatString(g.TextMatrix(Y, 3), 10, "|");
340       Printer.Print FormatString(g.TextMatrix(Y, 4), 14, "|");
350       Printer.Print FormatString(g.TextMatrix(Y, 5), 10, "|");
360       Printer.Print FormatString(g.TextMatrix(Y, 6), 26, "|")
 
370   Next


380   Printer.EndDoc



390   For Each Px In Printers
400     If Px.DeviceName = OriginalPrinter Then
410       Set Printer = Px
420       Exit For
430     End If
440   Next
End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Card Validation" & vbCr
20    strHeading = strHeading & "Between " & Format(dtFrom, "dd/mmm/yyyy") & " - " & Format(dtTo, "dd/mmm/yyyy") & vbCr
30    strHeading = strHeading & " " & vbCr
40    ExportFlexGrid g, Me, strHeading

End Sub

Private Sub cmdSearch_Click()

10    FillG

End Sub

Private Sub Form_Load()

10    dtFrom = Format$(Now - 7, "dd/mm/yyyy")
20    dtTo = Format$(Now, "dd/mm/yyyy")

30    FillG

End Sub


