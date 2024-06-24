VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReportViewer 
   Caption         =   "Netacquire - Report Viewer"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13695
   ControlBox      =   0   'False
   Icon            =   "frmReportViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHide 
      Caption         =   "Un-Hide Report"
      Height          =   1100
      Left            =   60
      Picture         =   "frmReportViewer.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   90
      Width           =   1000
   End
   Begin VB.CommandButton cmdSetPrinter 
      Caption         =   "Choose Printer"
      Height          =   1100
      Left            =   10020
      Picture         =   "frmReportViewer.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   90
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grdPTimes 
      Height          =   1185
      Left            =   3000
      TabIndex        =   4
      Top             =   60
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   2090
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
      FormatString    =   "       |<Department              |<Printed Time                   |<Report Number |<Counter|                          "
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8715
      Left            =   60
      TabIndex        =   2
      Top             =   1290
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   15372
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReportViewer.frx":1A9E
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Re-Print"
      Height          =   1100
      Left            =   11265
      Picture         =   "frmReportViewer.frx":1B20
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Re Print already Printed Results"
      Top             =   90
      Width           =   1000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   1100
      Left            =   12510
      Picture         =   "frmReportViewer.frx":29EA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   1000
   End
   Begin VB.Label lblDept 
      Height          =   255
      Left            =   2100
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCounterSelected 
      Height          =   255
      Left            =   1140
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgHidden 
      Height          =   225
      Left            =   330
      Picture         =   "frmReportViewer.frx":38B4
      Top             =   630
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   300
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
      Left            =   1170
      TabIndex        =   3
      Top             =   510
      Width           =   1455
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

46120 On Error GoTo AutoHide_Error

46130 sql = "SELECT Dept, PrintTime, ReportNumber, Counter, Hidden FROM Reports WHERE " & _
            "SampleID = '" & mSampleID & "' " & _
            "AND Dept = 'Microbiology' " & _
            "ORDER BY PrintTime DESC"
46140 Set tb = New Recordset
46150 RecOpenServer 0, tb, sql
46160 TopTime = ""
46170 Do While Not tb.EOF
46180   If TopTime = "" Then
46190     TopTime = Format$(tb!PrintTime, "dd/MMM/yyyy HH:nn:ss")
46200   Else
46210     If DateDiff("S", tb!PrintTime, TopTime) > 10 Then
46220       If tb!Hidden = 0 Then
46230         tb!Hidden = 2
46240         tb.Update
46250       End If
46260     End If
46270   End If
46280   tb.MoveNext
46290 Loop

46300 Exit Sub

AutoHide_Error:

      Dim strES As String
      Dim intEL As Integer

46310 intEL = Erl
46320 strES = Err.Description
46330 LogError "frmReportViewer", "AutoHide", intEL, strES, sql

End Sub

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

      'Hidden = 0 - auto set to not hidden
      '       = 1 - set to hidden by user
      '       = 2 - auto set to hidden
      '       = 3 - set to not hidden by user

46340 On Error GoTo FillG_Error

46350 AutoHide

46360 With grdPTimes
46370   .ColWidth(4) = 0
46380   .Rows = 2
46390   .AddItem ""
46400   .RemoveItem 1

46410 If mDept = "Microbiology" Then
46420   sql = "SELECT Dept, PrintTime, ReportNumber, Counter, Hidden, ReportType FROM Reports WHERE " & _
        "SampleID = '" & mSampleID & "' " & _
        "AND Dept = 'Microbiology' " & _
        "ORDER BY PrintTime DESC"
46430 ElseIf mDept = "Semen" Then
46440   sql = "SELECT Dept, PrintTime, ReportNumber, Counter, Hidden, ReportType FROM Reports WHERE " & _
        "SampleID = '" & mSampleID & "' " & _
        "AND Dept = 'Semen' " & _
        "ORDER BY PrintTime DESC"
46450 Else
46460   sql = "SELECT Dept, PrintTime, ReportNumber, Counter, Hidden, ReportType FROM Reports WHERE " & _
        "SampleID = '" & mSampleID & "' " & _
        "AND Dept <> 'Microbiology' " & _
        "AND Dept <> 'Semen' " & _
        "ORDER BY PrintTime DESC"
46470 End If
              
46480   Set tb = New Recordset
46490   RecOpenServer 0, tb, sql
46500   Do While Not tb.EOF
46510     s = vbTab & tb!Dept & vbTab & _
              Format$(tb!PrintTime, "dd/MM/yy HH:nn") & vbTab & _
              tb!ReportNumber & vbTab & _
              tb!Counter & vbTab & tb!ReportType & ""
46520     .AddItem s
46530     If tb!Hidden = 1 Or tb!Hidden = 2 Then
46540       .row = .Rows - 1
46550       .Col = 0
46560       Set .CellPicture = imgHidden.Picture
46570       .CellPictureAlignment = flexAlignCenterCenter
46580     End If
46590     tb.MoveNext
46600   Loop
46610   If .Rows > 2 Then
46620     .RemoveItem 1
46630     .row = 1
46640     HighlightRow
46650     FillReport
46660   End If

46670 End With

46680 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

46690 intEL = Erl
46700 strES = Err.Description
46710 LogError "frmReportViewer", "FillG", intEL, strES, sql

End Sub

Public Property Let InhibitChoosePrinter(ByVal blnNewValue As Boolean)

46720     pInhibitChoosePrinter = blnNewValue

End Property

Public Property Let PrintToPrinter(ByVal strNewValue As String)

46730     pPrintToPrinter = strNewValue

End Property


Public Property Get PrintToPrinter() As String

46740     PrintToPrinter = pPrintToPrinter

End Property


Private Sub cmdHide_Click()

          Dim Hide As Integer
          Dim sql As String

46750     On Error GoTo cmdHide_Click_Error

      'Hidden = 0 - auto set to not hidden
      '       = 1 - set to hidden by user
      '       = 2 - auto set to hidden
      '       = 3 - set to not hidden by user

46760     If Left$(cmdHide.Caption, 1) = "H" Then
46770         Hide = 1
46780     Else
46790         Hide = 3
46800     End If

46810     If grdPTimes.Rows = 2 And grdPTimes.TextMatrix(1, 4) = "" Then
46820         Exit Sub
46830     End If

46840     sql = "UPDATE Reports " & _
                "SET Hidden = '" & Hide & "' " & _
                "WHERE Counter = '" & grdPTimes.TextMatrix(grdPTimes.row, 4) & "'"
46850     Cnxn(0).Execute sql

46860     If Hide = 1 Or Hide = 2 Then
46870         Set grdPTimes.CellPicture = imgHidden.Picture
46880         grdPTimes.CellPictureAlignment = flexAlignCenterCenter
46890     Else
46900         Set grdPTimes.CellPicture = Nothing
46910     End If

46920     FillReport

46930     Exit Sub

cmdHide_Click_Error:

          Dim strES As String
          Dim intEL As Integer

46940     intEL = Erl
46950     strES = Err.Description
46960     LogError "frmReportViewer", "cmdHide_Click", intEL, strES, sql

End Sub

Private Sub cmdPrint_Click()
      Dim sql As String
      Dim tb As Recordset


46970 On Error GoTo cmdPrint_Click_Error

46980 sql = "Select * from PrintPending where ReprintReportCounter = " & lblCounterSelected
46990 Set tb = New Recordset
47000 RecOpenClient 0, tb, sql
47010 If tb.EOF Then
47020     tb.AddNew
47030 End If
47040 tb!SampleID = lblInfo
47050 tb!ReprintReportCounter = lblCounterSelected
47060 tb!Department = lblDept
47070 tb!UsePrinter = pPrintToPrinter
47080 tb.Update

47090 If InStr(App.Path, "Ward") > 0 Then
47100     Select Case lblDept
              Case "B"
47110             LogAsViewed "I", lblInfo, ""
47120         Case "H"
47130             LogAsViewed "J", lblInfo, ""
47140         Case "D"
47150             LogAsViewed "K", lblInfo, ""
47160         Case "M"
47170             LogAsViewed "N", lblInfo, ""

47180     End Select
47190 End If
47200 cmdPrint.Enabled = False
47210 Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

47220 intEL = Erl
47230 strES = Err.Description
47240 LogError "frmReportViewer", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Function getLabDeptCode(ByVal strDept As String) As String

47250 getLabDeptCode = ""
47260 Select Case UCase(strDept)
      Case "BIOCHEMISTRY": getLabDeptCode = "B"
47270 Case "HAEMATOLOGY": getLabDeptCode = "H"
47280 Case "COAGULATION": getLabDeptCode = "D"
47290 Case "IMMUNOLOGY": getLabDeptCode = "B"
47300 Case "BLOOD GAS": getLabDeptCode = "B"
47310 Case "EXTERNALS": getLabDeptCode = "B"
47320 Case "MICROBIOLOGY": getLabDeptCode = "M"
47330 Case "SEMEN": getLabDeptCode = "M"
47340 End Select

End Function

Private Sub cmdSetPrinter_Click()

47350     frmForcePrinter.From = Me
47360     frmForcePrinter.Show 1

47370     If pPrintToPrinter = "Automatic Selection" Then
47380         pPrintToPrinter = ""
47390     End If

47400     If pPrintToPrinter <> "" Then
47410         cmdSetPrinter.BackColor = vbRed
47420         cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
47430     Else
47440         cmdSetPrinter.BackColor = vbButtonFace
47450         pPrintToPrinter = ""
47460         cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
47470     End If

End Sub


Private Sub HighlightRow()

          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer

47480     With grdPTimes
47490         ySave = .row

47500         .Col = 0
47510         If .CellPicture = imgHidden Then
47520             cmdHide.Caption = "Un-Hide this Report"
47530         Else
47540             cmdHide.Caption = "Hide this Report"
47550         End If

47560         For Y = 1 To .Rows - 1
47570             .row = Y
47580             If .CellBackColor = vbYellow Then
47590                 For X = 0 To .Cols - 1
47600                     .Col = X
47610                     .CellBackColor = 0
47620                 Next
47630                 Exit For
47640             End If
47650         Next

47660         .row = ySave
47670         For X = 0 To .Cols - 1
47680             .Col = X
47690             .CellBackColor = vbYellow
47700         Next
        
47710         lblCounterSelected = .TextMatrix(.row, 4) 'Counter
47720         lblDept = getLabDeptCode(.TextMatrix(.row, 1)) 'Lab Dept
47730     End With

End Sub

Private Sub cmdExit_Click()

47740     Unload Me

End Sub

Private Sub Form_Load()

47750 On Error GoTo Form_Load_Error

47760 lblCounterSelected = ""
47770 lblDept = ""

47780 cmdHide.Visible = False
47790 If UCase(UserMemberOf) = "MANAGERS" Then
47800     cmdHide.Visible = True
47810 Else
47820     If UserHasAuthority(UserMemberOf, "EnableBioReportHideButton") = True Then
47830         cmdHide.Visible = True
47840     End If
47850 End If

47860 cmdSetPrinter.Visible = Not pInhibitChoosePrinter

47870 FillG

47880 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

47890 intEL = Erl
47900 strES = Err.Description
47910 LogError "frmReportViewer", "Form_Load", intEL, strES

End Sub

Private Sub FillReport()

      Dim tb As Recordset
      Dim sql As String

47920 On Error GoTo FillReport_Error

47930 rtb = ""
47940 rtb.SelText = ""

47950 If grdPTimes.Rows = 2 And grdPTimes.TextMatrix(1, 4) = "" Then
47960   Exit Sub
47970 End If

47980 grdPTimes.Col = 0
47990 If grdPTimes.CellPicture <> imgHidden Then
48000     sql = "SELECT Report FROM Reports WHERE " & _
                "Counter = '" & grdPTimes.TextMatrix(grdPTimes.row, 4) & "' "
48010     Set tb = New Recordset
48020     RecOpenServer 0, tb, sql
48030     If Not tb.EOF Then
48040       If Trim(tb!Report & "") <> "" Then
48050         rtb.SelText = Trim(tb!Report)
48060       End If
48070     End If
48080 End If

48090 Exit Sub

FillReport_Error:

      Dim strES As String
      Dim intEL As Integer

48100 intEL = Erl
48110 strES = Err.Description
48120 LogError "frmReportViewer", "FillReport", intEL, strES, sql

End Sub

Public Property Let SampleID(ByVal SID As String)

48130     On Error GoTo SampleID_Error

48140     mSampleID = SID
48150     lblInfo = SID

48160     Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer

48170     intEL = Erl
48180     strES = Err.Description
48190     LogError "frmReportViewer", "SampleID", intEL, strES

End Property

Public Property Let Dept(ByVal Department As String)

48200     mDept = Department

End Property

Private Sub Form_Unload(Cancel As Integer)

48210     pPrintToPrinter = ""
48220 mDept = ""


End Sub

Private Sub grdPTimes_Click()

48230     HighlightRow
48240     FillReport
48250     If cmdPrint.Enabled = False Then
48260       cmdPrint.Enabled = True
48270     End If

End Sub

