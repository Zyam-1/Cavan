VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frmReportViewer 
   Caption         =   "Netacquire - Report Viewer"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13695
   ControlBox      =   0   'False
   Icon            =   "frmReportViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   17.912
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   24.156
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHide 
      Caption         =   "Un-Hide this Report"
      Height          =   1125
      Left            =   690
      Picture         =   "frmReportViewer.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   180
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdPTimes 
      Height          =   1185
      Left            =   4200
      TabIndex        =   4
      Top             =   150
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2090
      _Version        =   393216
      Cols            =   5
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
      FormatString    =   "       |<Report Type              |<Printed Time                   |<Printed By     |<Counter"
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8715
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   15372
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReportViewer.frx":1794
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Re-Print this page"
      Height          =   1125
      Left            =   10680
      Picture         =   "frmReportViewer.frx":1816
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Re Print already Printed Results"
      Top             =   180
      Width           =   1605
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   1125
      Left            =   12420
      Picture         =   "frmReportViewer.frx":26E0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblDept 
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1050
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCounterSelected 
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1050
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgHidden 
      Height          =   225
      Left            =   330
      Picture         =   "frmReportViewer.frx":35AA
      Top             =   720
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   2670
      TabIndex        =   5
      Top             =   375
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      TabIndex        =   3
      Top             =   600
      Width           =   1995
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String
Private mDept As String

Private pPrintToPrinter As String
Private pInhibitChoosePrinter As Boolean

Private Sub AutoHide()

      Dim sql As String
      Dim tb As Recordset
      Dim TopTime As String

      'Hidden = 0 - auto set to not hidden
      '       = 1 - set to hidden by user
      '       = 2 - auto set to hidden
      '       = 3 - set to not hidden by user

10    On Error GoTo AutoHide_Error

20    sql = "SELECT Dept, PrintTime, RepNo,Counter FROM Reports WHERE " & _
            "SampleID = '" & mSampleID & "' " & _
            "AND Dept = 'Microbiology' " & _
            "ORDER BY PrintTime DESC"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    TopTime = ""
60    Do While Not tb.EOF
70      If TopTime = "" Then
80        TopTime = Format$(tb!PrintTime, "dd/MMM/yyyy HH:nn:ss")
90      Else
'100       If DateDiff("S", tb!PrintTime, TopTime) > 10 Then
'110         If tb!Hidden = 0 Then
'120           tb!Hidden = 2
'130           tb.Update
'140         End If
'150       End If
100     End If
110     tb.MoveNext
120   Loop

130   Exit Sub

AutoHide_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmReportViewer", "AutoHide", intEL, strES, sql

End Sub

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

      'Hidden = 0 - auto set to not hidden
      '       = 1 - set to hidden by user
      '       = 2 - auto set to hidden
      '       = 3 - set to not hidden by user

10    On Error GoTo FillG_Error

20    AutoHide

30    With grdPTimes
40      .ColWidth(4) = 0: .ColWidth(0) = 0
50      .Rows = 2
60      .AddItem ""
70      .RemoveItem 1

80    If mDept = "Microbiology" Then
90      sql = "SELECT * FROM Reports WHERE " & _
        "SampleID = '" & mSampleID & "' " & _
        "AND Dept = 'Microbiology' " & _
        "ORDER BY PrintTime DESC"
100   ElseIf mDept = "Semen" Then
110     sql = "SELECT * FROM Reports WHERE " & _
        "SampleID = '" & mSampleID & "' " & _
        "AND Dept = 'Semen' " & _
        "ORDER BY PrintTime DESC"
120   Else
130     sql = "SELECT * FROM Reports WHERE " & _
        "SampleID = '" & mSampleID & "' " & _
        "AND Dept <> 'Microbiology' " & _
        "AND Dept <> 'Semen' " & _
        "ORDER BY PrintTime DESC"
140   End If
              
150     Set tb = New Recordset
160     RecOpenServerBB 0, tb, sql
170     Do While Not tb.EOF
180       s = vbTab & _
              tb!Dept & vbTab & _
              Format$(tb!PrintTime, "dd/MM/yy HH:nn") & vbTab & _
              tb!Initiator & vbTab & _
              tb!Counter
190       .AddItem s
      '140       If tb!Hidden = 1 Or tb!Hidden = 2 Then
      '150         .Row = .Rows - 1
      '160         .Col = 0
      '170         Set .CellPicture = imgHidden.Picture
      '180         .CellPictureAlignment = flexAlignCenterCenter
      '190       End If
200       tb.MoveNext
210     Loop
220     If .Rows > 2 Then
230       .RemoveItem 1
240       .row = 1
250       HighlightRow
260       FillReport
270     End If

280   End With

290   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "frmReportViewer", "FillG", intEL, strES, sql

End Sub

Public Property Let InhibitChoosePrinter(ByVal blnNewValue As Boolean)

10        pInhibitChoosePrinter = blnNewValue

End Property

Public Property Let PrintToPrinter(ByVal strNewValue As String)

10        pPrintToPrinter = strNewValue

End Property


Public Property Get PrintToPrinter() As String

10        PrintToPrinter = pPrintToPrinter

End Property


Private Sub cmdHide_Click()

          Dim Hide As Integer
          Dim sql As String

10        On Error GoTo cmdHide_Click_Error

'Hidden = 0 - auto set to not hidden
'       = 1 - set to hidden by user
'       = 2 - auto set to hidden
'       = 3 - set to not hidden by user

20        If Left$(cmdHide.Caption, 1) = "H" Then
30            Hide = 1
40        Else
50            Hide = 3
60        End If

70        If grdPTimes.Rows = 2 And grdPTimes.TextMatrix(1, 4) = "" Then
80            Exit Sub
90        End If

100       sql = "UPDATE Reports " & _
                "SET Hidden = '" & Hide & "' " & _
                "WHERE Counter = '" & grdPTimes.TextMatrix(grdPTimes.row, 4) & "'"
110       Cnxn(0).Execute sql

120       If Hide = 1 Or Hide = 2 Then
130           Set grdPTimes.CellPicture = imgHidden.Picture
140           grdPTimes.CellPictureAlignment = flexAlignCenterCenter
150       Else
160           Set grdPTimes.CellPicture = Nothing
170       End If

180       FillReport

190       Exit Sub

cmdHide_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmReportViewer", "cmdHide_Click", intEL, strES, sql

End Sub

Private Sub cmdPrint_Click()
      Dim sql As String


10    On Error GoTo cmdPrint_Click_Error

12  If Not SetFormPrinter() Then Exit Sub

15  rtb.SelStart = 0
20  rtb.SelLength = 100000
30  rtb.SelPrint Printer.hDC

240   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmReportViewer", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Function getLabDeptCode(ByVal strDept As String) As String

10    getLabDeptCode = ""
20    Select Case UCase(strDept)
      Case "BIOCHEMISTRY": getLabDeptCode = "B"
30    Case "HAEMATOLOGY": getLabDeptCode = "H"
40    Case "COAGULATION": getLabDeptCode = "D"
50    Case "IMMUNOLOGY": getLabDeptCode = "B"
60    Case "BLOOD GAS": getLabDeptCode = "B"
70    Case "EXTERNALS": getLabDeptCode = "B"
80    Case "MICROBIOLOGY": getLabDeptCode = "M"
90    Case "SEMEN": getLabDeptCode = "M"
100   End Select

End Function




Private Sub HighlightRow()

          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer

10        With grdPTimes
20            ySave = .row

30            .col = 0
40            If .CellPicture = imgHidden Then
50                cmdHide.Caption = "Un-Hide this Report"
60            Else
70                cmdHide.Caption = "Hide this Report"
80            End If

90            For Y = 1 To .Rows - 1
100               .row = Y
110               If .CellBackColor = vbYellow Then
120                   For X = 0 To .Cols - 1
130                       .col = X
140                       .CellBackColor = 0
150                   Next
160                   Exit For
170               End If
180           Next

190           .row = ySave
200           For X = 0 To .Cols - 1
210               .col = X
220               .CellBackColor = vbYellow
230           Next
  
240           lblCounterSelected = .TextMatrix(.row, 4) 'Counter
250           lblDept = getLabDeptCode(.TextMatrix(.row, 1)) 'Lab Dept
260       End With

End Sub

Private Sub cmdExit_Click()

10        Unload Me

End Sub



Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    lblCounterSelected = ""
30    lblDept = ""

40    cmdHide.Visible = False
50    If UCase(UserMemberOf) = "MANAGERS" Or UCase(UserMemberOf) = "USERS" Then
60        cmdHide.Visible = True
70    End If

90    FillG

100   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmReportViewer", "Form_Load", intEL, strES

End Sub

Private Sub FillReport()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillReport_Error

20    rtb = ""
30    rtb.SelText = ""

40    If grdPTimes.Rows = 2 And grdPTimes.TextMatrix(1, 4) = "" Then
50      Exit Sub
60    End If

70    grdPTimes.col = 0
80    If grdPTimes.CellPicture <> imgHidden Then
90        sql = "SELECT Report FROM Reports WHERE " & _
                "Counter = '" & grdPTimes.TextMatrix(grdPTimes.row, 4) & "' "
100       Set tb = New Recordset
110       RecOpenServerBB 0, tb, sql
120       If Not tb.EOF Then
130         If Trim(tb!Report & "") <> "" Then
140           rtb.SelText = Trim(tb!Report)
150         End If
160       End If
170   End If

180   Exit Sub

FillReport_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmReportViewer", "FillReport", intEL, strES, sql

End Sub

Public Property Let SampleID(ByVal SID As String)

10        On Error GoTo SampleID_Error

20        mSampleID = SID
30        lblInfo = SID

40        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmReportViewer", "SampleID", intEL, strES

End Property

Public Property Let Dept(ByVal Department As String)

10        mDept = Department

End Property

Private Sub Form_Unload(Cancel As Integer)

10        pPrintToPrinter = ""
20    mDept = ""


End Sub

Private Sub grdPTimes_Click()

10        HighlightRow
20        FillReport

End Sub

