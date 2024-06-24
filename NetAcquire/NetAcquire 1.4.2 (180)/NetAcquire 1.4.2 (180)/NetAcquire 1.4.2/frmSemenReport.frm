VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSemenReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Semen Analysis History"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   150
      TabIndex        =   3
      Top             =   390
      Width           =   9345
      Begin VB.Label lblDemogComment 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   8505
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   2445
         TabIndex        =   16
         Top             =   540
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   1
         Left            =   7620
         TabIndex        =   15
         Top             =   210
         Width           =   285
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2775
         TabIndex        =   14
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label lblAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7935
         TabIndex        =   13
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   3915
         TabIndex        =   12
         Top             =   540
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   1
         Left            =   5130
         TabIndex        =   11
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   10
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   615
         TabIndex        =   9
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4515
         TabIndex        =   8
         Top             =   510
         Width           =   4140
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6045
         TabIndex        =   7
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1065
         TabIndex        =   6
         Top             =   540
         Width           =   1200
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1065
         TabIndex        =   5
         Top             =   210
         Width           =   3540
      End
      Begin VB.Label lblClDetails 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   150
         TabIndex        =   4
         Top             =   1110
         Visible         =   0   'False
         Width           =   8505
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   885
      Left            =   9720
      Picture         =   "frmSemenReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   705
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9180
      Top             =   120
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   6482
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmSemenReport.frx":066A
   End
   Begin MSFlexGridLib.MSFlexGrid grdSID 
      Height          =   1935
      Left            =   150
      TabIndex        =   1
      Top             =   1380
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   7
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmSemenReport.frx":06EC
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   150
      TabIndex        =   18
      Top             =   180
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmSemenReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean




Private Sub FillGrid()

      Dim sqlBase As String
      Dim sql As String
      Dim tb As Recordset
      Dim s As String
      Dim Cn As Integer

54490 On Error GoTo FillGrid_Error

54500 With grdSID
54510   .ColWidth(6) = 0 'Cn
54520   .Rows = 2
54530   .AddItem ""
54540   .RemoveItem 1
54550 End With

54560 sqlBase = "SELECT D.*, P.PrintedDateTime " & _
                "FROM Demographics D LEFT JOIN PrintValidLog P " & _
                "ON D.SampleID = P.SampleID " & _
                "WHERE PatName = '" & AddTicks(lblName) & "' "
54570 If Trim$(lblChart) <> "" Then
54580   sqlBase = sqlBase & "AND Chart = '" & lblChart & "' "
54590 Else
54600   sqlBase = sqlBase & "AND ( Chart IS NULL OR Chart = '' ) "
54610 End If
54620 If IsDate(lblDoB) Then
54630   sqlBase = sqlBase & "AND DoB = '" & Format$(lblDoB, "dd/mmm/yyyy") & "' "
54640 Else
54650   sqlBase = sqlBase & "AND ( DoB IS NULL OR DoB = '' ) "
54660 End If
54670 sqlBase = sqlBase & "AND D.SampleID > "

54680 For Cn = 0 To intOtherHospitalsInGroup
54690   Set tb = New Recordset
54700   sql = sqlBase & sysOptSemenOffset(Cn) & " AND D.SampleID < " & sysOptMicroOffsetOLD(Cn) & " " & _
              "AND SampleDate > '01/Jan/2007'"
54710   RecOpenClient Cn, tb, sql

54720   Do While Not tb.EOF
54730     s = Format$(Val(tb!SampleID) - sysOptSemenOffset(Cn)) & vbTab & _
              tb!Rundate & vbTab
54740     If IsDate(tb!SampleDate) Then
54750       If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
54760         s = s & Format(tb!SampleDate, "dd/MM/yy hh:mm")
54770       Else
54780         s = s & Format(tb!SampleDate, "dd/MM/yy")
54790       End If
54800     Else
54810       s = s & "Not Specified"
54820     End If
54830     s = s & vbTab
54840     If Not IsNull(tb!PrintedDateTime) Then
54850       s = s & Format$(tb!PrintedDateTime, "dd/MM/yy HH:mm")
54860     End If
54870     s = s & vbTab
54880     s = s & "Semen Analysis"
54890     s = s & vbTab & HospName(Cn) & vbTab & Format$(Cn)
54900     grdSID.AddItem s
          
54910     lblAge = tb!Age & ""
54920     Select Case Left$(UCase$(tb!Sex & ""), 1)
            Case "M": lblSex = "Male"
54930       Case "F": lblSex = "Female"
54940       Case Else: lblSex = ""
54950     End Select
54960     lblAddress = tb!Addr0 & " " & tb!Addr1 & ""
54970     tb.MoveNext
54980   Loop
54990 Next

55000 If grdSID.Rows > 2 Then
55010   grdSID.RemoveItem 1
55020 End If
55030 grdSID.Col = 1
55040 grdSID.Sort = 9

55050 Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

55060 intEL = Erl
55070 strES = Err.Description
55080 LogError "frmSemenReport", "FillGrid", intEL, strES, sql


End Sub




Private Sub cmdCancel_Click()

55090 Unload Me

End Sub

Private Sub Form_Activate()

55100 pBar.max = LogOffDelaySecs
55110 pBar = 0

55120 Timer1.Enabled = True

55130 If Activated Then Exit Sub
55140 Activated = True

55150 FillGrid

End Sub

Private Sub Form_Deactivate()

55160 Timer1.Enabled = False

End Sub


Private Sub Form_Load()

55170 Activated = False

55180 pBar.max = LogOffDelaySecs
55190 pBar = 0

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55200 pBar = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

55210 Activated = False

End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55220 pBar = 0

End Sub


Private Sub grdSID_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim Cn As Integer

55230 rtb.Text = ""

55240 If grdSID.MouseRow = 0 Then
55250   If SortOrder Then
55260     grdSID.Sort = flexSortGenericAscending
55270   Else
55280     grdSID.Sort = flexSortGenericDescending
55290   End If
55300   SortOrder = Not SortOrder
55310   Exit Sub
55320 End If

55330 If grdSID.Rows = 2 And grdSID.TextMatrix(1, 0) = "" Then Exit Sub

55340 For Y = 1 To grdSID.Rows - 1
55350   grdSID.row = Y
55360   For X = 1 To grdSID.Cols - 1
55370     grdSID.Col = X
55380     grdSID.CellBackColor = 0
55390   Next
55400 Next

55410 grdSID.row = grdSID.MouseRow
55420 For X = 1 To grdSID.Cols - 1
55430   grdSID.Col = X
55440   grdSID.CellBackColor = vbYellow
55450 Next

55460 Cn = grdSID.TextMatrix(grdSID.row, 6)
55470 PrintSAReport

End Sub


Private Sub grdSID_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As Date
      Dim d2 As Date
      Dim Column As Integer

55480 With grdSID
55490   Column = .Col
55500   Cmp = 0
55510   If IsDate(.TextMatrix(Row1, Column)) Then
55520     d1 = Format(.TextMatrix(Row1, Column), "dd/mmm/yyyy")
55530     If IsDate(.TextMatrix(Row2, Column)) Then
55540       d2 = Format(.TextMatrix(Row2, Column), "dd/mmm/yyyy")
55550       Cmp = Sgn(DateDiff("d", d1, d2))
55560     End If
55570   End If
55580 End With

End Sub

Private Sub grdSID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55590 pBar = 0

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

55600 pBar = 0

End Sub


Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

55610 pBar = 0

End Sub


Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

55620 pBar = 0

End Sub


Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

55630 pBar = 0

End Sub


Private Sub Label6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

55640 pBar = 0

End Sub


Private Sub Label7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

55650 pBar = 0

End Sub


Private Sub lblAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55660 pBar = 0

End Sub


Private Sub lblAge_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55670 pBar = 0

End Sub


Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55680 pBar = 0

End Sub


Private Sub lblClDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55690 pBar = 0

End Sub


Private Sub lblDemogComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55700 pBar = 0

End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55710 pBar = 0

End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55720 pBar = 0

End Sub


Private Sub lblSex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

55730 pBar = 0

End Sub


Private Sub rtb_KeyPress(KeyAscii As Integer)

55740 KeyAscii = 0

End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

55750 rtb.SelLength = 0
End Sub

Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
55760 pBar = pBar + 1
        
55770 If pBar = pBar.max Then
55780   Unload Me
55790 End If

End Sub



Public Sub PrintSAReport()
       
      Dim tb As Recordset
      Dim sql As String
55800 ReDim Comments(1 To 4) As String
      Dim DoB As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim PrintTime As String
      Dim SRS As New SemenResults
      Dim SR As New SemenResult

55810 On Error GoTo PrintSAReport_Error

55820 PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

55830 With grdSID

55840   sql = "SELECT * FROM Demographics WHERE " & _
              "SampleID = '" & .TextMatrix(.row, 0) + sysOptSemenOffset(0) & "'"
55850   Set tb = New Recordset
55860   RecOpenServer 0, tb, sql
          
55870   DoB = Format(tb!DoB, "dd/MMM/yyyy")
55880   Rundate = Format(tb!Rundate, "dd/MMM/yyyy")
55890   SampleDate = Format(tb!SampleDate, "dd/MMM/yyyy hh:mm:ss")
          
55900   PrintTextRTB rtb, "Cl Details:", , True
55910   PrintTextRTB rtb, tb!ClDetails & vbCrLf
55920   PrintSemenComment tb!SampleID, "D"
55930   PrintTextRTB rtb, vbCrLf
          
55940   SRS.Load tb!SampleID & ""

55950   If SRS.Count > 0 Then
55960     Set SR = SRS("SpecimenType")
55970     If Not SR Is Nothing Then
55980       Select Case SR.Result
              Case "Infertility Analysis": PrintSAInfertility .TextMatrix(.row, 0) + sysOptSemenOffset(0), SRS
55990         Case "Post Vasectomy": PrintSAVasectomy .TextMatrix(.row, 0) + sysOptSemenOffset(0)
56000       End Select
56010     End If
56020   End If
          
56030 End With

56040 Exit Sub

PrintSAReport_Error:

      Dim strES As String
      Dim intEL As Integer

56050 intEL = Erl
56060 strES = Err.Description
56070 LogError "frmSemenReport", "PrintSAReport", intEL, strES, sql

End Sub

Private Sub PrintSAVasectomy(ByVal SampleID As String)

56080 On Error GoTo PrintSAVasectomy_Error

56090 PrintTextRTB rtb, FormatString("Specimen Type : ", 20, , AlignRight)
56100 PrintTextRTB rtb, FormatString("Semen Post Vasectomy Analysis" & vbCrLf, 40, , AlignLeft), , True
56110 PrintTextRTB rtb, vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
56120 PrintSemenComment SampleID, "P"

56130 Exit Sub

PrintSAVasectomy_Error:

      Dim strES As String
      Dim intEL As Integer

56140 intEL = Erl
56150 strES = Err.Description
56160 LogError "frmSemenReport", "PrintSAVasectomy", intEL, strES

End Sub


Private Sub PrintSAInfertility(ByVal SampleID As String, ByVal SRS As SemenResults)

      Dim SR As SemenResult
      Dim Result As String

56170 On Error GoTo PrintSAInfertility_Error

56180 PrintTextRTB rtb, FormatString("Specimen Type : ", 20, , AlignRight)
56190 PrintTextRTB rtb, FormatString("Semen Infertility Analysis", 30, , AlignLeft), , True
56200 PrintTextRTB rtb, vbCrLf & vbCrLf
56210 PrintTextRTB rtb, Space(15)
56220 PrintTextRTB rtb, FormatString("Test Values", 15, , AlignCenter), , True, , True
56230 PrintTextRTB rtb, Space(70)
56240 PrintTextRTB rtb, FormatString("Reference Value", 15, , AlignCenter), , True, , True
56250 PrintTextRTB rtb, vbCrLf & vbCrLf

56260 Set SR = SRS("pH")
56270 If Not SR Is Nothing Then
56280   PrintTextRTB rtb, FormatString("pH: ", 15, , AlignRight), , True
56290   PrintTextRTB rtb, FormatString(SR.Result, 30, , AlignLeft)
56300   PrintTextRTB rtb, FormatString("(pH:     7.2 or more)", 85, , AlignRight)
56310 End If
56320 PrintTextRTB rtb, vbCrLf

56330 Set SR = SRS("Volume")
56340 If Not SR Is Nothing Then
56350   PrintTextRTB rtb, FormatString("Volume: ", 15, , AlignRight), , True
56360   PrintTextRTB rtb, FormatString(SR.Result, 8, , AlignLeft)
56370   PrintTextRTB rtb, FormatString("mls", 22, , AlignLeft)
56380   PrintTextRTB rtb, FormatString("(Volume: >2.0 mls)", 85, , AlignRight)
56390 End If
56400 PrintTextRTB rtb, vbCrLf

56410 Set SR = SRS("Consistency")
56420 If Not SR Is Nothing Then
56430   PrintTextRTB rtb, FormatString("Viscosity: ", 15, , AlignRight), , True
56440   PrintTextRTB rtb, FormatString(SR.Result, 30, , AlignLeft)
56450 End If
56460 PrintTextRTB rtb, vbCrLf

56470 PrintTextRTB rtb, FormatString("Motility: ", 15, , AlignRight), , True
56480 PrintTextRTB rtb, Space(30)
56490 PrintTextRTB rtb, FormatString("(Motility: % Grades A+B >50%)", 85, , AlignRight)
56500 PrintTextRTB rtb, vbCrLf

56510 PrintTextRTB rtb, FormatString("  Grade A: ", 15, , AlignRight), , True
56520 Result = ""
56530 Set SR = SRS("GradeA")
56540 If Not SR Is Nothing Then
56550   Result = SR.Result
56560 End If
56570 PrintTextRTB rtb, FormatString(Result, 8, , AlignLeft)
56580 PrintTextRTB rtb, FormatString("% (Fast progressive)", 30, , AlignLeft)
56590 PrintTextRTB rtb, vbCrLf

56600 Result = ""
56610 Set SR = SRS("GradeB")
56620 If Not SR Is Nothing Then
56630   Result = SR.Result
56640 End If
56650 PrintTextRTB rtb, FormatString("  Grade B: ", 15, , AlignRight), , True
56660 PrintTextRTB rtb, FormatString(Result, 8, , AlignLeft)
56670 PrintTextRTB rtb, FormatString("% (Slow progressive)", 30, , AlignLeft)
56680 PrintTextRTB rtb, vbCrLf

56690 Result = ""
56700 Set SR = SRS("GradeC")
56710 If Not SR Is Nothing Then
56720   Result = SR.Result
56730 End If
56740 PrintTextRTB rtb, FormatString("  Grade C: ", 15, , AlignRight), , True
56750 PrintTextRTB rtb, FormatString(Result, 8, , AlignLeft)
56760 PrintTextRTB rtb, FormatString("% (motile non progressive)", 30, , AlignLeft)
56770 PrintTextRTB rtb, vbCrLf

56780 Result = ""
56790 Set SR = SRS("GradeD")
56800 If Not SR Is Nothing Then
56810   Result = SR.Result
56820 End If
56830 PrintTextRTB rtb, FormatString("  Grade D: ", 15, , AlignRight), , True
56840 PrintTextRTB rtb, FormatString(Result, 8, , AlignLeft)
56850 PrintTextRTB rtb, FormatString("% (non motile)", 30, , AlignLeft)
56860 PrintTextRTB rtb, vbCrLf

56870 Result = ""
56880 Set SR = SRS("Morphology")
56890 If Not SR Is Nothing Then
56900   Result = SR.Result
56910 End If
56920 PrintTextRTB rtb, FormatString("Morphology: ", 15, , AlignRight), , True
56930 PrintTextRTB rtb, FormatString(Result, 8, , AlignLeft)
56940 PrintTextRTB rtb, FormatString("% Normal", 22, , AlignLeft)
56950 PrintTextRTB rtb, FormatString("Morphology: >15% Normal)", 85, , AlignRight)
56960 PrintTextRTB rtb, vbCrLf

56970 Result = ""
56980 Set SR = SRS("SemenCount")
56990 If Not SR Is Nothing Then
57000   Result = SR.Result
57010 End If
57020 PrintTextRTB rtb, FormatString("Sperm Count: ", 15, , AlignRight), , True
57030 PrintTextRTB rtb, FormatString(Result, 8, , AlignLeft)
57040 PrintTextRTB rtb, FormatString("  million/ml", 22, , AlignLeft)
57050 PrintTextRTB rtb, FormatString("(Sperm Count: >20 million/ml)", 85, , AlignRight)
57060 PrintTextRTB rtb, vbCrLf & vbCrLf

57070 PrintSemenComment SampleID, "I"

57080 PrintTextRTB rtb, "Semen Analysis Test Values lower than the Reference Values are ", 9
57090 PrintTextRTB rtb, "ASSOCIATED", 9, True, , True
57100 PrintTextRTB rtb, " with decreased Fertility.", 9

57110 Exit Sub

PrintSAInfertility_Error:

      Dim strES As String
      Dim intEL As Integer

57120 intEL = Erl
57130 strES = Err.Description
57140 LogError "frmSemenReport", "PrintSAInfertility", intEL, strES

End Sub


Private Sub PrintSemenComment(ByVal SampleID As String, _
                              ByVal Source As String)

      Dim tb As Recordset
      Dim sql As String
57150 On Error GoTo PrintSemenComment_Error

57160 ReDim Comments(1 To 4) As String
      Dim n As Integer
      Dim pSource As String

57170 Select Case UCase$(Left$(Source, 1))
        Case "I": Source = "Semen": pSource = "Infertility Comment: "
57180   Case "P": Source = "Semen": pSource = "Post Vasectomy: "
57190   Case "D": Source = "Demographic": pSource = "Demographic Comment:"
57200 End Select

57210 sql = "Select * from Observations where " & _
            "SampleID = '" & Val(SampleID) & "' " & _
            "AND Discipline = '" & Source & "'"
57220 Set tb = New Recordset
57230 RecOpenServer 0, tb, sql
57240 rtb.SelFontSize = 10
57250 If Not tb.EOF Then
57260   If Trim$(tb!Comment & "") <> "" Then
57270     FillCommentLines pSource & tb!Comment, 4, Comments(), 97
57280     PrintTextRTB rtb, pSource, 9, True
57290     PrintTextRTB rtb, Mid$(Comments(1), Len(pSource) + 1) & vbCrLf
57300     For n = 2 To 4
57310       If Trim$(Comments(n)) <> "" Then
57320         PrintTextRTB rtb, Comments(n) & vbCrLf
57330       End If
57340     Next
57350   End If
57360 End If

57370 Exit Sub

PrintSemenComment_Error:

      Dim strES As String
      Dim intEL As Integer

57380 intEL = Erl
57390 strES = Err.Description
57400 LogError "modPrintSemen", "PrintSemenComment", intEL, strES, sql


End Sub


