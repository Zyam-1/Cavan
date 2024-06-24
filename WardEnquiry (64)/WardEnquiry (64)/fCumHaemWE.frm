VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form fCumHaemWE 
   Caption         =   "NetAcquire - Cumulative Haematology"
   ClientHeight    =   6720
   ClientLeft      =   375
   ClientTop       =   1395
   ClientWidth     =   12645
   HelpContextID   =   10025
   Icon            =   "fCumHaemWE.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6720
   ScaleWidth      =   12645
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   615
      Left            =   11250
      Picture         =   "fCumHaemWE.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   90
      Width           =   1245
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6660
      Top             =   0
   End
   Begin VB.CommandButton bFull 
      Caption         =   "Graphs"
      Height          =   675
      Left            =   7710
      Picture         =   "fCumHaemWE.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   8910
      Picture         =   "fCumHaemWE.frx":1616
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   810
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
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
      FormatString    =   $"fCumHaemWE.frx":1C80
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   6090
      TabIndex        =   6
      Top             =   60
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   344
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
      Left            =   10140
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblSex 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6090
      TabIndex        =   11
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label lblChart 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2670
      TabIndex        =   10
      Top             =   30
      Width           =   1605
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2670
      TabIndex        =   9
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblDoB 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4965
      TabIndex        =   8
      Top             =   60
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   4590
      TabIndex        =   7
      Top             =   90
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      Height          =   585
      Left            =   150
      Top             =   30
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1560
      Picture         =   "fCumHaemWE.frx":1D53
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Click on Sample ID to show full details"
      Height          =   405
      Left            =   210
      TabIndex        =   4
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2220
      TabIndex        =   3
      Top             =   390
      Width           =   420
   End
   Begin VB.Label lblChartTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   2250
      TabIndex        =   2
      Top             =   90
      Width           =   375
   End
End
Attribute VB_Name = "fCumHaemWE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub bcancel_Click()

Unload Me

End Sub


Private Sub bFull_Click()

With frmFullHaemWE
    .lblName = lblName
    .lblDoB = lblDoB
    .lblChart = lblChart
    .Show 1
End With

End Sub

Private Sub cmdXL_Click()

ExportFlexGrid g, Me

End Sub

Private Sub Form_Activate()

If LogOffNow Then
    Unload Me
End If

PBar.Max = LogOffDelaySecs
PBar = 0
SingleUserUpdateLoggedOn UserName

Timer1.Enabled = True

If Not Activated Then
    FillG
    Activated = True
End If

End Sub

Private Sub Form_Deactivate()

Timer1.Enabled = False

End Sub


Private Sub Form_Load()

g.Font.Bold = True

PBar.Max = LogOffDelaySecs
PBar = 0

LogAsViewed "H", "", frmMain.txtChart

Activated = False

End Sub


Private Sub FillG()

      Dim sn As Recordset
      Dim tb As Recordset
      Dim sql As String
      Dim S As String
      Dim OBS As Observations
      Dim RunSampleDiff As Boolean
      Dim sampleDate As String
      Dim n As Integer

10    On Error GoTo FillG_Error

20    lblChartTitle = "Chart"


30    With frmViewResultsWE
40        For n = 1 To .grd.Rows - 1
50            If n > 1 Then sql = sql & " Union "
60            sql = sql & "SELECT D.SampleID, D.RunDate, D.SampleDate FROM Demographics D JOIN HaemResults R ON D.SampleID = R.SampleID WHERE "
70            If Trim$(.grd.TextMatrix(n, 0)) <> "" Then
80                sql = sql & "Chart = '" & .grd.TextMatrix(n, 0) & "' "
90            Else
100               sql = sql & "(Chart is null or Chart = '') "
110           End If
120           sql = sql & "AND D.PatName = '" & AddTicks(.grd.TextMatrix(n, 2)) & "' "
130           If IsDate(.grd.TextMatrix(n, 1)) Then
140               sql = sql & "AND D.DoB = '" & Format(.grd.TextMatrix(n, 1), "dd/mmm/yyyy") & "' "
150           Else
160               sql = sql & "AND (D.DoB = '' or DoB is Null) "
170           End If
180           sql = sql & "AND D.RunDate > '" & Format(Now - Val(frmMain.txtLookBack), "dd/MMM/yyyy") & "' "
190       Next n
200       sql = sql & "ORDER BY D.SampleDate DESC,D.SampleID DESC"
210   End With

      'sql = "SELECT D.SampleID, D.RunDate, D.SampleDate FROM Demographics D JOIN HaemResults R ON D.SampleID = R.SampleID WHERE " & _
      '      "PatName = '" & AddTicks(lblName) & "'"
      'If Trim$(lblChart) <> "" Then
      '    sql = sql & "and Chart = '" & lblChart & "' "
      'Else
      '    sql = sql & "and (Chart is null or Chart = '') "
      'End If
      'If IsDate(lblDoB) Then
      '    sql = sql & "and DoB = '" & Format$(lblDoB, "dd/mmm/yyyy") & "' "
      'Else
      '    sql = sql & "and (DoB = '' or DoB is null )"
      'End If

220   g.Visible = False
230   g.Rows = 2
240   g.AddItem ""
250   g.RemoveItem 1
260   g.ColWidth(0) = 0
270   g.ColWidth(3) = 0
280   g.ColWidth(4) = 0

290   For n = 0 To intOtherHospitalsInGroup
300       Set sn = New Recordset
310       RecOpenServer n, sn, sql

320       RunSampleDiff = False
330       Do While Not sn.EOF
340           S = Format$(n) & vbTab & _
                  Format(sn!Rundate, "dd/mm/yyyy") & vbTab & _
                  sn!SampleID & vbTab
350           sampleDate = ""
360           If DateDiff("d", sn!sampleDate, sn!Rundate) <> 0 Then
370               sampleDate = Format(sn!sampleDate, "dd/mm/yy")
380               If InStr(sn!sampleDate, ":") Then
390                   sampleDate = sampleDate & " " & Format(sn!sampleDate, "hh:mm")
400               End If
410               S = S & sampleDate
420               RunSampleDiff = True
430           ElseIf InStr(sn!sampleDate, ":") Then
440               sampleDate = Format(sn!sampleDate, "dd/mm/yy hh:mm")
450               RunSampleDiff = True
460               S = S & sampleDate
470           End If
480           S = S & vbTab

490           sql = "Select * from HaemResults where " & _
                    "SampleID = '" & sn!SampleID & "'"
500           Set tb = New Recordset
510           RecOpenServer n, tb, sql
520           If Not tb.EOF Then
530               If IsNull(tb!Valid) Or tb!Valid = 0 Then
540                   S = S & vbTab & "Not" & vbTab & "Valid"
550               Else
560                   Set OBS = New Observations
570                   Set OBS = OBS.Load(sn!SampleID, "Haematology")
580                   If Not OBS Is Nothing Then
590                       S = S & " " & OBS.Item(1).Comment
600                       g.ColWidth(4) = TextWidth("Comment ")
610                   End If
                          'tb!RDWCV & vbTab
620                   S = S & vbTab & tb!rbc & vbTab & _
                          tb!Hgb & vbTab & _
                          tb!Hct & vbTab & _
                          tb!MCV & vbTab & _
                          tb!mch & vbTab & _
                          tb!mchc & vbTab & _
                          tb!plt & vbTab & _
                          tb!WBC & vbTab & _
                          tb!NeutA & " (" & tb!NeutP & "%)" & vbTab & _
                          tb!LymA & " (" & tb!LymP & "%)" & vbTab & _
                          tb!MonoA & " (" & tb!MonoP & "%)" & vbTab & _
                          tb!EosA & " (" & tb!EosP & "%)" & vbTab & _
                          tb!BasA & " (" & tb!BasP & "%)" & vbTab
630               End If
640           End If
650           g.AddItem S
660           sn.MoveNext
670       Loop
680   Next
690   If RunSampleDiff Then
700       g.ColWidth(3) = TextWidth("Sample    ")
710   End If

720   If g.Rows > 2 Then
730       g.RemoveItem 1
740   End If

750   Colourize

760   g.Visible = True

      '750   g.col = 1
      '760   g.Sort = 9

770   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

780   intEL = Erl
790   strES = Err.Description
800   LogError "fCumHaemWE", "FillG", intEL, strES, sql

End Sub

Private Sub Colourize()

Dim x As Integer
Dim y As Integer
Dim Value As Single
Dim Analyte As String
Dim Sex As String

On Error GoTo Colourize_Error

Sex = Trim$(UCase$(Left$(lblSex, 1) & " "))

For y = 1 To g.Rows - 1
    g.Row = y
    For x = 5 To 18
        g.col = x
        If Trim$(g.TextMatrix(y, x)) = "" Then
            g.CellBackColor = &HFFFFFF
            g.CellForeColor = &H0&
        ElseIf Trim$(g.TextMatrix(y, x)) = "(%)" Then
            g.TextMatrix(y, x) = ""
            g.CellBackColor = &HFFFFFF
            g.CellForeColor = &H0&
        Else
            Value = Val(g.TextMatrix(y, x))
            Analyte = Choose(x - 4, "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC", "RDWCV", "PLT", "WBC", "NEUTA", "LYMA", "MONOA", "EOSA", "BASA")

            Select Case InterpH(Value, Analyte, Sex, lblDoB, "")
            Case "X":
                g.CellBackColor = vbBlack
                g.CellForeColor = vbWhite
            Case "H":
                g.CellBackColor = &HFFFF&
                g.CellForeColor = &HFF&
            Case "L"
                g.CellBackColor = &HFFFF00
                g.CellForeColor = &HC00000
            Case Else
                g.CellBackColor = &HFFFFFF
                g.CellForeColor = &H0&
            End Select
        End If
    Next
Next

Exit Sub

Colourize_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "fCumHaemWE", "Colourize", intEL, strES

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

PBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

Activated = False

End Sub

Private Sub g_Click()

On Error GoTo g_Click_Error

If g.MouseCol = 3 Then
    If g.ColWidth(3) = TextWidth("Sample    ") Then
        g.ColWidth(3) = TextWidth("Sample Date/Time        ")
    Else
        g.ColWidth(3) = TextWidth("Sample    ")
    End If
    Exit Sub
End If

If g.MouseCol = 4 Then
    If g.ColWidth(4) = TextWidth("Comment ") Then
        g.ColWidth(4) = 6000
    Else
        g.ColWidth(4) = TextWidth("Comment ")
    End If
    Exit Sub
End If

If g = "" Then Exit Sub
If g.MouseCol <> 2 Then Exit Sub
If g.MouseRow = 0 Then Exit Sub

If g.TextMatrix(g.Row, 4) = "Not" Then
    iMsg "Cannot display " & vbCrLf & "Non Validated Results", vbInformation
    Exit Sub
End If

With fResultHaemWE
    .lblSampleID = g.TextMatrix(g.Row, 2)
    .lblChart = lblChart
    .lblName = lblName
    .lblDoB = lblDoB
    .lblCnxn = g.TextMatrix(g.Row, 0)
    .Show 1
End With

Exit Sub

g_Click_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "fCumHaemWE", "g_Click", intEL, strES

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, cmp As Integer)

Dim d1 As Date
Dim d2 As Date
Dim Column As Integer

With g
    Column = .col
    cmp = 0
    If IsDate(.TextMatrix(Row1, Column)) Then
        d1 = Format(.TextMatrix(Row1, Column), "dd/mmm/yyyy")
        If IsDate(.TextMatrix(Row2, Column)) Then
            d2 = Format(.TextMatrix(Row2, Column), "dd/mmm/yyyy")
            cmp = Sgn(DateDiff("d", d1, d2))
        End If
    End If
End With

End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

PBar = 0

End Sub


Private Sub Timer1_Timer()

'tmrRefresh.Interval set to 1000
PBar = PBar + 1

If PBar = PBar.Max Then
    LogOffNow = True
    Unload Me
End If

End Sub


