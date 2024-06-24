VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fmonthly 
   Caption         =   "Monthly Analysis"
   ClientHeight    =   5295
   ClientLeft      =   1740
   ClientTop       =   1905
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FMONTHLY.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5295
   ScaleWidth      =   7785
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6000
      Picture         =   "FMONTHLY.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   450
      Width           =   1245
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   345
      Left            =   1410
      TabIndex        =   6
      Top             =   570
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   609
      _Version        =   327681
      Value           =   2012
      BuddyControl    =   "txtYear"
      BuddyDispid     =   196614
      OrigLeft        =   1350
      OrigTop         =   480
      OrigRight       =   1590
      OrigBottom      =   1065
      Max             =   2020
      Min             =   1997
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3465
      Left            =   360
      TabIndex        =   5
      Top             =   1350
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   6112
      _Version        =   393216
      Rows            =   14
      Cols            =   6
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
      FormatString    =   "<                       |^Samples    |^DAT          |^X matches  |^A/B IDs   |^+ve A/B's "
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4620
      Picture         =   "FMONTHLY.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   450
      Width           =   1245
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1860
      Picture         =   "FMONTHLY.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   450
      Width           =   1245
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   690
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "2012"
      Top             =   570
      Width           =   675
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3240
      Picture         =   "FMONTHLY.frx":1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   450
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   1170
      TabIndex        =   7
      Top             =   5040
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
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
      Left            =   5940
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   405
   End
End
Attribute VB_Name = "fmonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CalcDAT()

      Dim sn As Recordset
      Dim tb As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim n As Integer
      Dim lngTotal As Long
      Dim lngMonthTot As Long

10    On Error GoTo CalcDAT_Error

20    lngTotal = 0

30    For n = 1 To 12
40      FromDate = Format$("01/" & Format(n, "0#") & "/" & txtYear, "dd/mmm/yyyy") & " 00:00:00"
50      ToDate = DateAdd("m", 1, FromDate)
60      ToDate = Format(DateAdd("d", -1, ToDate), "dd/mmm/yyyy") & " 23:59:59"
  
70      sql = "select labnumber, DateTime from patientdetails where " & _
              "(dat0 = 1 or dat1 = 1 or dat2 = 1 or dat3 = 1 " & _
              "or dat4 = 1 or dat5 = 1 or dat6 = 1 or dat7 = 1 " & _
              "or dat8 = 1 or dat9 = 1 or requestfrom = 'D') and " & _
              "( datetime between '" & FromDate & "' and '" & ToDate & "')"
80      Set sn = New Recordset
90      RecOpenServerBB 0, sn, sql
100     If Not sn.EOF Then
110       lngMonthTot = 0
120       Do While Not sn.EOF
130         sql = "Select top 1 DateTime from PatientDetailsAudit where " & _
                  "LabNumber = '" & sn!LabNumber & "' " & _
                  "order by DateTime Asc"
140         Set tb = New Recordset
150         RecOpenServerBB 0, tb, sql
160         If Not tb.EOF Then
170           If DateDiff("s", tb!DateTime, FromDate) < 0 And DateDiff("s", tb!DateTime, ToDate) > 0 Then
180             lngMonthTot = lngMonthTot + 1
190           End If
200         Else
210           lngMonthTot = lngMonthTot + 1
220         End If
230         sn.MoveNext
240       Loop
250       If lngMonthTot <> 0 Then
260         g.TextMatrix(n, 2) = lngMonthTot
270       End If
280       lngTotal = lngTotal + lngMonthTot
290     End If
300   Next
310   g.TextMatrix(13, 2) = lngTotal

320   Exit Sub

CalcDAT_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "fmonthly", "CalcDAT", intEL, strES, sql


End Sub

Private Sub CalcSamples()

      Dim sn As Recordset
      Dim tb As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim n As Integer
      Dim lngTotal As Long
      Dim lngMonthTot As Long

10    On Error GoTo CalcSamples_Error

20    lngTotal = 0

30    For n = 1 To 12
40      FromDate = Format$("01/" & Format(n, "0#") & "/" & txtYear, "dd/mmm/yyyy") & " 00:00:00"
50      ToDate = DateAdd("m", 1, FromDate)
60      ToDate = Format(DateAdd("d", -1, ToDate), "dd/mmm/yyyy") & " 23:59:59"
  
70      sql = "select labnumber, datetime from patientdetails where " & _
              "SampleDate between '" & FromDate & "' and '" & ToDate & "'"
80      Set sn = New Recordset
90      RecOpenServerBB 0, sn, sql
100     If Not sn.EOF Then
110       lngMonthTot = 0
120       Do While Not sn.EOF
130         sql = "Select top 1 SampleDate from PatientDetailsAudit where " & _
                  "LabNumber = '" & sn!LabNumber & "' " & _
                  "order by DateTime Asc"
140         Set tb = New Recordset
150         RecOpenServerBB 0, tb, sql
160         If Not tb.EOF Then
170           If DateDiff("s", tb!SampleDate, FromDate) <= 0 And DateDiff("s", tb!SampleDate, ToDate) >= 0 Then
180             lngMonthTot = lngMonthTot + 1
190           End If
200         Else
210           lngMonthTot = lngMonthTot + 1
220         End If
230         sn.MoveNext
240       Loop
250       If lngMonthTot <> 0 Then
260         g.TextMatrix(n, 1) = lngMonthTot
270       End If
280       g.Refresh
290       lngTotal = lngTotal + lngMonthTot
300     End If
310   Next
320   g.TextMatrix(13, 1) = lngTotal

330   Exit Sub

CalcSamples_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "fmonthly", "CalcSamples", intEL, strES, sql


End Sub

Private Sub CalcAID()

      Dim sn As Recordset
      Dim tb As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim n As Integer
      Dim lngTotal As Long
      Dim lngMonthTot As Long

10    On Error GoTo CalcAID_Error

20    lngTotal = 0
30    For n = 1 To 12
40      FromDate = Format$("01/" & Format(n, "0#") & "/" & txtYear, "dd/mmm/yyyy") & " 00:00:00"
50      ToDate = DateAdd("m", 1, FromDate)
60      ToDate = Format(DateAdd("d", -1, ToDate), "dd/mmm/yyyy") & " 23:59:59"
  
70      sql = "select labnumber, SampleDate from patientdetails " & _
              "where not (aidr is null) and " & _
              "( SampleDate between '" & FromDate & "' and '" & ToDate & "')"
80      Set sn = New Recordset
90      RecOpenServerBB 0, sn, sql
100     If Not sn.EOF Then
110       lngMonthTot = 0
120       Do While Not sn.EOF
130         sql = "Select top 1 SampleDate from PatientDetailsAudit where " & _
                  "LabNumber = '" & sn!LabNumber & "' " & _
                  "order by SampleDate Asc"
140         Set tb = New Recordset
150         RecOpenServerBB 0, tb, sql
160         If Not tb.EOF Then
170           If DateDiff("s", tb!SampleDate, FromDate) <= 0 And DateDiff("s", tb!SampleDate, ToDate) >= 0 Then
180             lngMonthTot = lngMonthTot + 1
190           End If
200         Else
210           lngMonthTot = lngMonthTot + 1
220         End If
230         sn.MoveNext
240       Loop
250       If lngMonthTot <> 0 Then
260         g.TextMatrix(n, 4) = lngMonthTot
270       End If
280       lngTotal = lngTotal + lngMonthTot
290     End If
300   Next
310   g.TextMatrix(13, 4) = lngTotal

320   Exit Sub

CalcAID_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "fmonthly", "CalcAID", intEL, strES, sql


End Sub
Private Sub CalcAIDPos()

      Dim sn As Recordset
      Dim tb As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim n As Integer
      Dim lngTotal As Long
      Dim lngMonthTot As Long

10    On Error GoTo CalcAIDPos_Error

20    lngTotal = 0
30    For n = 1 To 12
40      FromDate = Format$("01/" & Format(n, "0#") & "/" & txtYear, "dd/mmm/yyyy") & " 00:00:00"
50      ToDate = DateAdd("m", 1, FromDate)
60      ToDate = Format(DateAdd("d", -1, ToDate), "dd/mmm/yyyy") & " 23:59:59"
  
70      sql = "select labnumber, SampleDate from patientdetails " & _
              "where aidr like '%Positive%' and " & _
              "( SampleDate between '" & FromDate & "' and '" & ToDate & "')"
80      Set sn = New Recordset
90      RecOpenServerBB 0, sn, sql
100     If Not sn.EOF Then
110       lngMonthTot = 0
120       Do While Not sn.EOF
130         sql = "Select top 1 SampleDate from PatientDetailsAudit where " & _
                  "LabNumber = '" & sn!LabNumber & "' " & _
                  "order by SampleDate Asc"
140         Set tb = New Recordset
150         RecOpenServerBB 0, tb, sql
160         If Not tb.EOF Then
170           If DateDiff("s", tb!SampleDate, FromDate) <= 0 And DateDiff("s", tb!SampleDate, ToDate) >= 0 Then
180             lngMonthTot = lngMonthTot + 1
190           End If
200         Else
210           lngMonthTot = lngMonthTot + 1
220         End If
230         sn.MoveNext
240       Loop
250       If lngMonthTot <> 0 Then
260         g.TextMatrix(n, 5) = lngMonthTot
270       End If
280       lngTotal = lngTotal + lngMonthTot
290     End If
300   Next
310   g.TextMatrix(13, 5) = lngTotal

320   Exit Sub

CalcAIDPos_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "fmonthly", "CalcAIDPos", intEL, strES, sql


End Sub


Private Sub CalcXM()

      Dim sn As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim n As Integer
      Dim lngTotal As Long

10    On Error GoTo CalcXM_Error

20    lngTotal = 0
30    For n = 1 To 12
40      FromDate = Format$("01/" & Format(n, "0#") & "/" & txtYear, "dd/mmm/yyyy") & " 00:00:00"
50      ToDate = DateAdd("m", 1, FromDate)
60      ToDate = Format(DateAdd("d", -1, ToDate), "dd/mmm/yyyy") & " 23:59:59"
  
70      sql = "select count(labnumber) as tot from product where " & _
              "event = 'X' and labnumber <> '' and " & _
              "( datetime between '" & FromDate & " ' and '" & ToDate & "') " & _
              "and labnumber in " & _
              "  (select labnumber from patientdetails where " & _
              "   SampleDate between '" & FromDate & "' and '" & ToDate & "')"
  
80      Set sn = New Recordset
90      RecOpenServerBB 0, sn, sql
100     If sn!Tot <> 0 Then
110       g.TextMatrix(n, 3) = sn!Tot
120       lngTotal = lngTotal + sn!Tot
130     End If
140   Next
150   g.TextMatrix(13, 3) = lngTotal

160   Exit Sub

CalcXM_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fmonthly", "CalcXM", intEL, strES, sql


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
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Print
40    Printer.FontName = "Courier New"
50    Printer.Font.Bold = True
60    Printer.Font.Size = 10

70    Printer.Print FormatString("Blood Transfusion Report", 99, , AlignCenter)
80    Printer.Print FormatString("Monthly Analysis For  " & txtYear, 99, , AlignCenter)
90    Printer.Print

100   Printer.Font.Size = 9

110   For i = 1 To 108
120       Printer.Print "-";
130   Next i
140   Printer.Print

150   Printer.Print FormatString("", 0, "|");
160   Printer.Print FormatString(" ", 26, "|", AlignCenter);
170   Printer.Print FormatString("Samples", 15, "|", AlignCenter);
180   Printer.Print FormatString("DAT", 15, "|", AlignCenter);
190   Printer.Print FormatString("X-Matches", 15, "|", AlignCenter);
200   Printer.Print FormatString("A/B IDs", 15, "|", AlignCenter);
210   Printer.Print FormatString("+VE A/B's", 15, "|", AlignCenter)

220   For i = 1 To 108
230       Printer.Print "-";
240   Next i
250   Printer.Print

260   Printer.Font.Bold = False

270   For Y = 1 To 12
280       Printer.Print FormatString("", 0, "|");
290       Printer.Print FormatString(g.TextMatrix(Y, 0), 26, "|");
300       Printer.Print FormatString(g.TextMatrix(Y, 1), 15, "|", AlignCenter);
310       Printer.Print FormatString(g.TextMatrix(Y, 2), 15, "|", AlignCenter);
320       Printer.Print FormatString(g.TextMatrix(Y, 3), 15, "|", AlignCenter);
330       Printer.Print FormatString(g.TextMatrix(Y, 4), 15, "|", AlignCenter);
340       Printer.Print FormatString(g.TextMatrix(Y, 5), 15, "|", AlignCenter)
350   Next

360   For i = 1 To 108
370       Printer.Print "-";
380   Next i
390   Printer.Print

400   Printer.Font.Bold = True
410   Printer.Print FormatString("", 0, "|");
420   Printer.Print FormatString(g.TextMatrix(13, 0), 26, "|", AlignCenter);
430   Printer.Print FormatString(g.TextMatrix(13, 1), 15, "|", AlignCenter);
440   Printer.Print FormatString(g.TextMatrix(13, 2), 15, "|", AlignCenter);
450   Printer.Print FormatString(g.TextMatrix(13, 3), 15, "|", AlignCenter);
460   Printer.Print FormatString(g.TextMatrix(13, 4), 15, "|", AlignCenter);
470   Printer.Print FormatString(g.TextMatrix(13, 5), 15, "|", AlignCenter)
480   Printer.Font.Bold = False

490   For i = 1 To 108
500       Printer.Print "-";
510   Next i
520   Printer.Print

530   Printer.EndDoc

540   For Each Px In Printers
550     If Px.DeviceName = OriginalPrinter Then
560       Set Printer = Px
570       Exit For
580     End If
590   Next

End Sub

Private Sub cmdStart_Click()

      Dim n As Integer
      Dim X As Integer

10    For n = 1 To 13
20      g.Row = n
30      For X = 1 To 5
40        g.Col = X
50        g = ""
60      Next
70    Next

80    CalcSamples
90    CalcDAT
100   CalcAID
110   CalcAIDPos
120   CalcXM

End Sub

Private Sub cmdXL_Click()
      Dim strHeading As String

10    strHeading = "Blood Transfusion Report" & vbCr
20    strHeading = strHeading & "Monthly Analysis For  " & txtYear & vbCr
30    strHeading = strHeading & " " & vbCr
40    ExportFlexGrid g, Me, strHeading
End Sub

Private Sub Form_Load()

      Dim n As Integer

10    txtYear = Format(Now, "yyyy")

20    For n = 1 To 12
30      g.TextMatrix(n, 0) = Format$("28/" & n & "/2004", "mmmm")
40    Next
50    g.TextMatrix(13, 0) = "Total"

End Sub

