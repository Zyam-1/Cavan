VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fdaily 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Report"
   ClientHeight    =   8580
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   12900
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "FDAILY.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8580
   ScaleWidth      =   12900
   Begin VB.Frame Frame1 
      Caption         =   "Show"
      Height          =   735
      Left            =   3090
      TabIndex        =   7
      Top             =   120
      Width           =   1515
      Begin VB.OptionButton optDAT 
         Caption         =   "DAT"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   420
         Width           =   645
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   450
         Width           =   465
      End
      Begin VB.OptionButton optAN 
         Caption         =   "AN"
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   210
         Width           =   525
      End
      Begin VB.OptionButton optGH 
         Caption         =   "GH"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export"
      Height          =   900
      Left            =   6690
      Picture         =   "FDAILY.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   900
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   7185
      Left            =   150
      TabIndex        =   3
      Top             =   1020
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   12674
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"FDAILY.frx":0BD4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton bcontinue 
      Appearance      =   0  'Flat
      Caption         =   "&Search"
      Height          =   900
      Left            =   4650
      Picture         =   "FDAILY.frx":0CB0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   900
      Left            =   5670
      Picture         =   "FDAILY.frx":113B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   900
      Left            =   11895
      Picture         =   "FDAILY.frx":17A5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   900
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   150
      TabIndex        =   6
      Top             =   8250
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Between Dates"
      Height          =   735
      Left            =   150
      TabIndex        =   11
      Top             =   120
      Width           =   2955
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   60
         TabIndex        =   12
         Top             =   270
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Format          =   110166017
         CurrentDate     =   36963
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   285
         Left            =   1470
         TabIndex        =   13
         Top             =   270
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Format          =   110166017
         CurrentDate     =   36963
      End
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
      Left            =   7770
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "fdaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bcontinue_Click()

10    FillGrid

End Sub

Private Sub bprint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    On Error GoTo bprint_Click_Error

20    OriginalPrinter = Printer.DeviceName
30    If Not SetFormPrinter() Then Exit Sub
40    Printer.Orientation = vbPRORLandscape
50    Printer.Font.Name = "Courier New"
60    Printer.Font.Size = 10
      '****Report heading
70    Printer.Font.Bold = True
80    Printer.Print "Search Results For ";
90    If optGH Then
100       Printer.Print "Group & Hold"
110   ElseIf optDAT Then
120       Printer.Print "D.A.T."
130   ElseIf optAN Then
140       Printer.Print "Ante-Netal"
150   ElseIf optAll Then
160       Printer.Print "All"
170   End If
180   Printer.Print "From " & Format(dtFrom, "dd/mm/yyyy") & " To " & Format$(dtTo, "dd/mm/yyyy") & "."
190   Printer.Print
      Dim i As Integer
200   For i = 1 To 152
210       Printer.Print "-";
220   Next i
230   Printer.Print
      '****Report body
240   Printer.Font.Size = 9

250   Printer.Print FormatString("", 0, "|");
260   Printer.Print FormatString("Date", 16, "|", AlignCenter);
270   Printer.Print FormatString("Lab #", 10, "|", AlignCenter);
280   Printer.Print FormatString("Patient Name", 34, "|", AlignCenter);
290   Printer.Print FormatString("D.o.B.", 12, "|", AlignCenter);
300   Printer.Print FormatString("Sex", 5, "|", AlignCenter);
310   Printer.Print FormatString("M.R.N.", 10, "|", AlignCenter);
320   Printer.Print FormatString("A and E", 10, "|", AlignCenter);
330   Printer.Print FormatString("Location", 25, "|", AlignCenter);
340   Printer.Print FormatString("P. Group", 10, "|", AlignCenter);
350   Printer.Print FormatString("Request", 9, "|", AlignCenter)
360   Printer.Font.Bold = False
370   For i = 1 To 152
380       Printer.Print "-";
390   Next i
400   Printer.Print

410   For Y = 1 To Grid1.Rows - 1
420       Printer.Print FormatString("", 0, "|");
430       Printer.Print FormatString(Grid1.TextMatrix(Y, 0), 16, "|");
440       Printer.Print FormatString(Grid1.TextMatrix(Y, 1), 10, "|");
450       Printer.Print FormatString(Grid1.TextMatrix(Y, 2), 34, "|");
460       Printer.Print FormatString(Grid1.TextMatrix(Y, 3), 12, "|");
470       Printer.Print FormatString(Grid1.TextMatrix(Y, 4), 5, "|", AlignCenter);
480       Printer.Print FormatString(Grid1.TextMatrix(Y, 5), 10, "|");
490       Printer.Print FormatString(Grid1.TextMatrix(Y, 6), 10, "|");
500       Printer.Print FormatString(Grid1.TextMatrix(Y, 7), 25, "|");
510       Printer.Print FormatString(Grid1.TextMatrix(Y, 8), 10, "|", AlignCenter);
520       Printer.Print FormatString(Grid1.TextMatrix(Y, 9), 9, "|")
530   Next
540   Printer.EndDoc

550   For Each Px In Printers
560     If Px.DeviceName = OriginalPrinter Then
570       Set Printer = Px
580       Exit For
590     End If
600   Next

610   Exit Sub

bprint_Click_Error:

Dim strES As String
Dim intEL As Integer

620   intEL = Erl
630   strES = Err.Description
640   LogError "fdaily", "bprint_Click", intEL, strES

End Sub

Private Sub FillGrid()

      Dim sn As Recordset
      Dim final As Recordset
      Dim sql As String
      Dim strFromTime As String
      Dim strToTime As String
      Dim s As String

10    On Error GoTo FillGrid_Error

20    Grid1.Rows = 2

30    strFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
40    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59"

50    sql = "SELECT DISTINCT LabNumber " & _
            "FROM PatientDetails WHERE " & _
            "SampleDate BETWEEN '" & strFromTime & "' " & _
            "AND '" & strToTime & "' "
60    If optGH Then
70      sql = sql & "AND RequestFrom = 'G' "
80    ElseIf optAN Then
90      sql = sql & "AND RequestFrom = 'A' "
100   ElseIf optDAT Then
110     sql = sql & "AND RequestFrom = 'D' "
120   End If
130   sql = sql & "ORDER BY LabNumber"
140   Set sn = New Recordset
150   RecOpenServerBB 0, sn, sql

160   If sn.EOF Then
170     iMsg "No records found!", vbInformation
180     If TimedOut Then Unload Me: Exit Sub
190     Exit Sub
200   End If

210   Grid1.Visible = False

220   Do While Not sn.EOF
230     sql = "select * from patientdetails where " & _
              "labnumber = '" & sn!LabNumber & "' " & _
              "order by datetime"
240     Set final = New Recordset
250     RecOpenServerBB 0, final, sql
260     final.MoveLast
270     s = Format(final!DateTime, "dd/mm/yyyy") & vbTab & _
            final!LabNumber & vbTab & _
            final!Name & vbTab & _
            final!DoB & vbTab & _
            final!Sex & vbTab & _
            final!Patnum & vbTab & _
            final!AandE & vbTab & _
            final!Ward & vbTab & _
            final!fGroup & ""
280     s = s & vbTab
290     Select Case final!requestfrom
          Case "X": s = s & "X-M"
300       Case "G": s = s & "G/H"
310       Case "A": s = s & "A/N"
320       Case "D": s = s & "DAT"
330     End Select
340     Grid1.AddItem s
350     sn.MoveNext
360   Loop
370   Grid1.Visible = True
380   If Grid1.Rows > 2 Then
390     Grid1.RemoveItem 1
400   End If

410   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

420   intEL = Erl
430   strES = Err.Description
440   LogError "fdaily", "FillGrid", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Search Results For "
20    If optGH Then
30        strHeading = strHeading & "Group & Hold" & vbCr
40    ElseIf optDAT Then
50        strHeading = strHeading & "D.A.T." & vbCr
60    ElseIf optAN Then
70        strHeading = strHeading & "Ante-Netal" & vbCr
80    ElseIf optAll Then
90        strHeading = strHeading & "All" & vbCr
100   End If
110   strHeading = "From " & Format(dtFrom, "dd/mm/yyyy") & " To " & Format$(dtTo, "dd/mm/yyyy") & "." & vbCr
120   strHeading = strHeading & " " & vbCr
130   ExportFlexGrid Grid1, Me, strHeading

End Sub

Private Sub Form_Load()

10    dtTo = Format(Now, "dd/mmm/yyyy")
20    dtFrom = dtTo

End Sub

