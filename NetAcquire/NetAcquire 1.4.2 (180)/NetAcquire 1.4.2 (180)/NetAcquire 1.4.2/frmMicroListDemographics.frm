VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMicroListDemographics 
   Caption         =   "NetAcquire - Demographic Data"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10155
   Begin VB.Frame Frame1 
      Caption         =   "Between"
      Height          =   1125
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   3315
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1020
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1020
         TabIndex        =   8
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optDates 
         Alignment       =   1  'Right Justify
         Caption         =   "Dates"
         Height          =   225
         Left            =   780
         TabIndex        =   7
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optSIDs 
         Caption         =   "Sample Numbers"
         Height          =   225
         Left            =   1530
         TabIndex        =   6
         Top             =   0
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1020
         TabIndex        =   10
         Top             =   600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218955777
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   1020
         TabIndex        =   11
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218955777
         CurrentDate     =   38126
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.CommandButton breCalc 
      Caption         =   "Calculate"
      Height          =   825
      Left            =   3630
      Picture         =   "frmMicroListDemographics.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   6180
      Picture         =   "frmMicroListDemographics.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   825
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   7020
      Picture         =   "frmMicroListDemographics.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   825
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   60
      TabIndex        =   0
      Top             =   1230
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6435
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11351
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   $"frmMicroListDemographics.frx":0C7E
   End
   Begin VB.Label lblSortBy 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1020
      Visible         =   0   'False
      Width           =   4245
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
      Left            =   7860
      TabIndex        =   14
      Top             =   390
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmMicroListDemographics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim lngFrom As Long
      Dim lngTo As Long

38160 On Error GoTo FillG_Error

38170 g.Rows = 2
38180 g.AddItem ""
38190 g.RemoveItem 1

38200 If optDates Then
38210   If Abs(DateDiff("d", dtFrom, dtTo)) > 60 Then
38220     iMsg "Maximum 60 Days!", vbExclamation
38230     Exit Sub
38240   End If
38250 Else
38260   lngFrom = Val(txtFrom)
38270   lngTo = Val(txtTo)
38280   If lngTo < lngFrom Then
38290     txtFrom = Format(lngTo)
38300     txtTo = Format(lngFrom)
38310     lngFrom = Val(txtFrom)
38320     lngTo = Val(txtTo)
38330   End If
38340   If lngFrom < 1 Or lngFrom > 9999999 Then
38350     iMsg "Number <From> is incorrect!", vbExclamation
38360     txtFrom = ""
38370     Exit Sub
38380   End If
38390   If lngTo < 1 Or lngTo > 9999999 Then
38400     iMsg "Number <To> is incorrect!", vbExclamation
38410     txtTo = ""
38420     Exit Sub
38430   End If
38440   If lngTo - lngFrom > 5000 Then
38450     iMsg "Maximum 5000 Records!", vbExclamation
38460     Exit Sub
38470   End If
38480 End If

38490 lblSortBy.Caption = "Calculating - Please wait"
38500 lblSortBy.Visible = True
38510 lblSortBy.Refresh

38520 sql = "SELECT D.*, M.Site, M.SiteDetails " & _
            "FROM Demographics D JOIN SiteDetails50 M " & _
            "ON D.SampleID = M.SampleID " & _
            "WHERE "
38530 If optDates Then
38540   sql = sql & "D.Rundate BETWEEN '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
                    "AND '" & Format(dtTo, "dd/mmm/yyyy") & "' "
38550 Else
      '410     sql = sql & "D.SampleID BETWEEN '" & Format$(Val(txtFrom) + sysOptMicroOffset(0)) & "' " & _
      '                    "AND '" & Format$(Val(txtTo) + sysOptMicroOffset(0)) & "' "
38560   sql = sql & "D.SampleID BETWEEN '" & Format$(Val(txtFrom)) & "' " & _
                    "AND '" & Format$(Val(txtTo)) & "' "
38570 End If
38580 sql = sql & "ORDER BY D.SampleID"

38590 Set tb = New Recordset
38600 RecOpenClient 0, tb, sql
38610 If Not tb.EOF Then
38620   pb.max = tb.RecordCount
38630   pb = 0
38640   pb.Visible = True
38650   g.Visible = False
38660 End If
38670 Do While Not tb.EOF
        
38680   pb = pb + 1
38690   s = tb!PatName & vbTab & _
            tb!Chart & vbTab & _
            tb!DoB & vbTab
38700   If Left$(tb!Sex & "", 1) = "M" Then
38710     s = s & "Male"
38720   ElseIf Left$(tb!Sex & "", 1) = "F" Then
38730     s = s & "Female"
38740   End If
38750   s = s & vbTab & _
            tb!Addr0 & vbTab & _
            tb!Addr1 & vbTab & _
            tb!Ward & vbTab & _
            tb!Clinician & vbTab & _
            tb!GP & vbTab
      '610     If Val(tb!SampleID) > sysOptMicroOffset(0) Then
      '620       s = s & tb!SampleID - sysOptMicroOffset(0)
      '630     Else
38760     s = s & tb!SampleID
      '650     End If
38770   s = s & vbTab & _
            tb!Site & vbTab & _
            tb!SiteDetails & vbTab & _
            tb!SampleDate & vbTab & _
            tb!RecDate & vbTab & _
            tb!Rundate & ""
        
38780   g.AddItem s
        
38790   tb.MoveNext
38800 Loop

38810 If g.Rows > 2 Then
38820   g.RemoveItem 1
38830 End If

38840 g.Visible = True

38850 pb.Visible = False
38860 pb = 0
38870 lblSortBy.Visible = False

38880 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

38890 intEL = Erl
38900 strES = Err.Description
38910 LogError "frmMicroListDemographics", "FillG", intEL, strES, sql

End Sub

Private Sub breCalc_Click()

38920 FillG

End Sub

Private Sub cmdCancel_Click()

38930 Unload Me

End Sub


Private Sub cmdXL_Click()

38940 ExportFlexGrid g, Me

End Sub


Private Sub Form_Load()

38950 dtFrom = Format(Now, "dd/mmm/yyyy")
38960 dtTo = dtFrom

End Sub

Private Sub g_Click()

38970 Screen.MousePointer = vbHourglass

38980 If g.MouseRow = 0 Then
38990   If InStr(UCase$(g.TextMatrix(0, g.Col)), "DATE") <> 0 Then
39000     lblSortBy.Caption = "Sorting by " & g.TextMatrix(0, g.Col) & " ... Please wait."
39010     lblSortBy.Visible = True
39020     lblSortBy.Refresh
39030     g.Sort = 9
39040     lblSortBy.Visible = False
39050   Else
39060     If SortOrder Then
39070       g.Sort = flexSortGenericAscending
39080     Else
39090       g.Sort = flexSortGenericDescending
39100     End If
39110   End If
39120   SortOrder = Not SortOrder
39130 End If

39140 Screen.MousePointer = 0

End Sub

Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

39150 If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
39160   Cmp = 0
39170   Exit Sub
39180 End If

39190 If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
39200   Cmp = 0
39210   Exit Sub
39220 End If

39230 d1 = Format(g.TextMatrix(Row1, g.Col), "dd/mmm/yyyy hh:mm:ss")
39240 d2 = Format(g.TextMatrix(Row2, g.Col), "dd/mmm/yyyy hh:mm:ss")

39250 If SortOrder Then
39260   Cmp = Sgn(DateDiff("s", d1, d2))
39270 Else
39280   Cmp = Sgn(DateDiff("s", d2, d1))
39290 End If

End Sub

Private Sub optDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

39300 dtFrom.Visible = True
39310 dtTo.Visible = True
39320 txtFrom.Visible = False
39330 txtTo.Visible = False

End Sub


Private Sub optSIDs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

39340 dtFrom.Visible = False
39350 dtTo.Visible = False
39360 txtFrom.Visible = True
39370 txtTo.Visible = True

End Sub


