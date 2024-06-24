VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSignOffSamples 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ward Enquiry ---Unsigned Samples"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14430
   Icon            =   "frmSignOffSamples.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   14430
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   960
      Left            =   10860
      Picture         =   "frmSignOffSamples.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   210
      Width           =   1245
   End
   Begin VB.ComboBox cmbWard 
      Height          =   315
      ItemData        =   "frmSignOffSamples.frx":0316
      Left            =   7215
      List            =   "frmSignOffSamples.frx":031D
      TabIndex        =   12
      Text            =   "cmbWard"
      Top             =   525
      Width           =   2085
   End
   Begin VB.ComboBox CmbClinician 
      Height          =   315
      ItemData        =   "frmSignOffSamples.frx":0326
      Left            =   5010
      List            =   "frmSignOffSamples.frx":032D
      TabIndex        =   10
      Text            =   "CmbClinician"
      Top             =   525
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   960
      Left            =   13065
      Picture         =   "frmSignOffSamples.frx":0336
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   210
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1065
      Left            =   150
      TabIndex        =   3
      Top             =   165
      Width           =   2445
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   675
         TabIndex        =   4
         Top             =   270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         Format          =   111542273
         CurrentDate     =   43101
         MinDate         =   43101
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   675
         TabIndex        =   5
         Top             =   630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         Format          =   111542273
         CurrentDate     =   38631
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   675
         Width           =   195
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.ComboBox cmbDescipline 
      Height          =   315
      ItemData        =   "frmSignOffSamples.frx":09A0
      Left            =   2760
      List            =   "frmSignOffSamples.frx":09B6
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   525
      Width           =   2085
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   960
      Left            =   9510
      Picture         =   "frmSignOffSamples.frx":0A01
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6255
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   $"frmSignOffSamples.frx":0E43
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
   Begin VB.Label Label 
      Caption         =   "Click on sample ID to sign off"
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   16
      Top             =   1260
      Width           =   4920
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   300
      Left            =   10860
      TabIndex        =   15
      Top             =   1185
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label 
      Caption         =   "Ward"
      Height          =   195
      Index           =   4
      Left            =   7215
      TabIndex        =   13
      Top             =   255
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "Clinician"
      Height          =   195
      Index           =   3
      Left            =   5010
      TabIndex        =   11
      Top             =   255
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "Discipline"
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   255
      Width           =   870
   End
End
Attribute VB_Name = "frmSignOffSamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
10    On Error GoTo cmdCancel_Click_Error

20    Unload Me

30    Exit Sub

cmdCancel_Click_Error:
      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmSignOffSamples", "cmdCancel_Click", intEL, strES
End Sub

Private Sub cmdSearch_Click()
      Dim tb As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim tDiscipline As String
      Dim col As Integer
      Dim x As Integer
      Dim ExcludeCodes As String


10    On Error GoTo cmdSearch_Click_Error

20    FromDate = Format$(dtFrom, "dd/mmm/yyyy 00:00:00")
30    ToDate = Format$(dtTo, "dd/mmm/yyyy 23:59:59")

40    ExcludeCodes = "'1071','1072','1073','HbA1','REJ','REJEX'"

50    If cmbDescipline.List(cmbDescipline.ListIndex) = "All" Or cmbDescipline.List(cmbDescipline.ListIndex) = "Biochemistry" Then
60        sql = " SELECT DISTINCT COALESCE(r.signoff,0) AS signoff,d.SampleID,d.Chart,d.PatName,d.age,d.Clinician,d.ward ,d.SampleDate,'Biochemistry' as Discipline " & _
                "FROM BioResults R  LEFT JOIN demographics D ON d.SampleID=r.sampleid " & _
                "WHERE d.ward<>'gp' AND COALESCE( r.signoff,0)=0 AND (COALESCE (R.valid, 0) = 1)" & _
                " AND R.Code not in (" & ExcludeCodes & ") AND d.SampleDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' "
70        If CmbClinician.ListIndex > 0 Then
80            sql = sql & " And D.Clinician = '" & CmbClinician.List(CmbClinician.ListIndex) & "'"
90        End If
100       If cmbWard.ListIndex > 0 Then
110           sql = sql & " And D.Ward = '" & cmbWard.List(cmbWard.ListIndex) & "'"
120       End If
130   End If

140   If cmbDescipline.List(cmbDescipline.ListIndex) = "All" Or cmbDescipline.List(cmbDescipline.ListIndex) = "Coagulation" Then
150       If Len(sql) > 0 Then sql = sql & " Union "
160       sql = sql & " SELECT DISTINCT COALESCE(r.signoff,0) AS signoff,d.SampleID,d.Chart,d.PatName,d.age,d.Clinician,d.ward,d.SampleDate,'Coagulation' as Discipline " & _
                "FROM CoagResults R  LEFT JOIN demographics D ON d.SampleID=r.sampleid " & _
                "WHERE d.ward<>'gp' AND COALESCE( r.signoff,0)=0 AND (COALESCE (R.valid, 0) = 1)  AND d.SampleDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' "
170       If CmbClinician.ListIndex > 0 Then
180           sql = sql & " And D.Clinician = '" & CmbClinician.List(CmbClinician.ListIndex) & "'"
190       End If
200       If cmbWard.ListIndex > 0 Then
210           sql = sql & " And D.Ward = '" & cmbWard.List(cmbWard.ListIndex) & "'"
220       End If
230   End If

240   If cmbDescipline.List(cmbDescipline.ListIndex) = "All" Or cmbDescipline.List(cmbDescipline.ListIndex) = "Haematology" Then
250       If Len(sql) > 0 Then sql = sql & " Union "
260       sql = sql & " SELECT DISTINCT COALESCE(r.signoff,0) AS signoff,d.SampleID,d.Chart,d.PatName,d.age,d.Clinician,d.ward,d.SampleDate,'Haematology' as Discipline " & _
                "FROM HaemResults R  LEFT JOIN demographics D ON d.SampleID=r.sampleid " & _
                "WHERE d.ward<>'gp' AND COALESCE( r.signoff,0)=0 AND (COALESCE (R.valid, 0) = 1)  AND d.SampleDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' "
270       If CmbClinician.ListIndex > 0 Then
280           sql = sql & " And D.Clinician = '" & CmbClinician.List(CmbClinician.ListIndex) & "'"
290       End If
300       If cmbWard.ListIndex > 0 Then
310           sql = sql & " And D.Ward = '" & cmbWard.List(cmbWard.ListIndex) & "'"
320       End If
330   End If
      'Case "Endocrinology", "All"
      '    If Len(sql) > 0 Then sql = sql & " Union "
      '    sql = " SELECT COALESCE(r.signoff,0) AS signoff,d.SampleID,d.Chart,d.PatName,d.age,d.Clinician,d.ward,d.SampleDate " & _
           '          "FROM immResults R  LEFT JOIN demographics D ON d.SampleID=r.sampleid " & _
           ' "WHERE d.ward<>'gp' AND COALESCE( r.signoff,0)=0  AND d.SampleDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' "
      'Case "Immunology", "All"
340   If cmbDescipline.List(cmbDescipline.ListIndex) = "All" Or cmbDescipline.List(cmbDescipline.ListIndex) = "Immunology" Then
350       If Len(sql) > 0 Then sql = sql & " Union "
360       sql = sql & " SELECT DISTINCT COALESCE(r.signoff,0) AS signoff,d.SampleID,d.Chart,d.PatName,d.age,d.Clinician,d.ward,d.SampleDate,'Immunology' as Discipline " & _
                "FROM ImmResults R  LEFT JOIN demographics D ON d.SampleID=r.sampleid " & _
                "WHERE d.ward <> 'gp' AND COALESCE( r.signoff,0)=0 AND (COALESCE (R.valid, 0) = 1) AND d.SampleDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' "
370       If CmbClinician.ListIndex > 0 Then
380           sql = sql & " And D.Clinician = '" & CmbClinician.List(CmbClinician.ListIndex) & "'"
390       End If
400       If cmbWard.ListIndex > 0 Then
410           sql = sql & " And D.Ward = '" & cmbWard.List(cmbWard.ListIndex) & "'"
420       End If
430   End If
      'Case "Micro", "All"
440   If cmbDescipline.List(cmbDescipline.ListIndex) = "All" Or cmbDescipline.List(cmbDescipline.ListIndex) = "Microbiology" Then
450       If Len(sql) > 0 Then sql = sql & " Union "
460       sql = sql & " SELECT DISTINCT COALESCE(r.signoff,0) AS signoff,d.SampleID,d.Chart,d.PatName,d.age,d.Clinician,d.ward,d.SampleDate,'Microbiology' as Discipline " & _
                "FROM PrintValidLog R  LEFT JOIN demographics D ON d.SampleID=r.sampleid " & _
                "WHERE d.ward <> 'gp' AND COALESCE( r.signoff,0)=0 AND (COALESCE (R.valid, 0) = 1)  AND d.SampleDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' "
470       If CmbClinician.ListIndex > 0 Then
480           sql = sql & " And D.Clinician = '" & CmbClinician.List(CmbClinician.ListIndex) & "'"
490       End If
500       If cmbWard.ListIndex > 0 Then
510           sql = sql & " And D.Ward = '" & cmbWard.List(cmbWard.ListIndex) & "'"
520       End If
530   End If
540   Set tb = New Recordset
550   sql = sql & "Order by Discipline"
560   RecOpenClient Cn, tb, sql
'Text1.Text = sql


570   If tb.EOF Then
          'g.Clear
580       g.Rows = 1
590   Else
600       tb.MoveFirst
          'g.Clear
610       g.Rows = 1



620       Do Until tb.EOF

      '620           If tb!Discipline <> tDiscipline Then
      '630               g.AddItem vbTab & vbTab & vbTab & tb!Discipline
      '
      '640               For col = 0 To 6
      '650                   g.Row = g.Rows - 1
      '660                   g.col = col
      '670                   g.CellBackColor = &HFFFF80
      '680               Next
      '
      '690           End If
630           g.AddItem tb!SampleID & vbTab & tb!SampleDate & vbTab & tb!Chart & vbTab & tb!PatName & vbTab & tb!Age & vbTab & tb!Clinician & vbTab & tb!Ward
640           tDiscipline = tb!Discipline
650           tb.MoveNext
660       Loop
670   End If
      ''g.SelectionMode = flexSelectionByRow
      'For x = 0 To g.Cols - 1
      '    g.col = x
      '    g.CellBackColor = &H80000018
      'Next

680   Exit Sub

cmdSearch_Click_Error:
      Dim strES As String
      Dim intEL As Integer

690   intEL = Erl
700   strES = Err.Description
710   LogError "frmSignOffSamples", "cmdSearch_Click", intEL, strES, sql

End Sub
Private Sub MakeSql()

End Sub

Private Sub cmdXL_Click()
10    ExportFlexGrid g, Me
End Sub


Private Function SampleidToDOB(SampleID As String) As String
      Dim sql As String
      Dim tb As Recordset
10    On Error GoTo SampleidToDOB_Error

20    sql = "Select DOB from Demographics where sampleid='" & SampleID & " '"
30    Set tb = New Recordset
40    RecOpenClient Cn, tb, sql
50    If tb.EOF Then
60        SampleidToDOB = ""
70    Else
80        SampleidToDOB = tb!DoB & ""
90    End If

100   Exit Function

SampleidToDOB_Error:
      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmSignOffSamples", "SampleidToDOB", intEL, strES, sql

End Function
Private Sub dtFrom_Change()
'CmbClinician.ListIndex = -1
'cmbWard.ListIndex = -1
End Sub

Private Sub dtTo_Change()
'CmbClinician.ListIndex = -1
'cmbWard.ListIndex = -1
End Sub

Private Sub Form_Load()
      Dim tempdate As Date
10    On Error GoTo Form_Load_Error

20    tempdate = DateAdd("d", -7, Now)
30    dtFrom = Format(tempdate, "dd/MMM/yyyy")
40    dtTo = Format(Now, "dd/MMM/yyyy")
50    dtFrom.MinDate = DateAdd("d", Val(frmMain.txtLookBack) * -1, Now)
60    cmbDescipline.ListIndex = 0

70    FillCombo
80    CmbClinician.ListIndex = 0
90    cmbWard.ListIndex = 0
100   Exit Sub

Form_Load_Error:
      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmSignOffSamples", "Form_Load", intEL, strES
End Sub
Private Sub FillCombo()
      Dim tb As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String



10    On Error GoTo FillCombo_Error

      '20 FromDate = Format$(dtFrom, "dd/mmm/yyyy")
      '30 ToDate = Format$(dtTo, "dd/mmm/yyyy")

20    sql = "SELECT DISTINCT Text FROM Clinicians " & _
            "WHERE InUse = 1 ORDER BY Text"
30    Set tb = New Recordset
40    RecOpenClient Cn, tb, sql
50    CmbClinician.Clear
60    CmbClinician.AddItem "All"
70    If tb.EOF Then
80    Else
90        tb.MoveFirst
100       Do Until tb.EOF
110           CmbClinician.AddItem tb!Text
120           tb.MoveNext
130       Loop
140   End If


150   sql = " SELECT DISTINCT Text FROM Wards WHERE InUse = 1 ORDER BY Text  "
160   Set tb = New Recordset
170   RecOpenClient Cn, tb, sql
180   cmbWard.Clear
190   cmbWard.AddItem "All"
200   If tb.EOF Then
210   Else
220       tb.MoveFirst
230       Do Until tb.EOF
240           cmbWard.AddItem tb!Text
250           tb.MoveNext
260       Loop
270   End If

280   FixComboWidth CmbClinician
290   FixComboWidth cmbWard


300   Exit Sub

FillCombo_Error:
      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmSignOffSamples", "FillCombo", intEL, strES

End Sub
Private Sub g_Click()
      Static SortOrder As Boolean
10    On Error GoTo g_Click_Error

20    Debug.Print g.MouseRow
30    If g.MouseRow = 0 Then
          '    If cmbDescipline.ListIndex <> 0 Then
40        SortOrder = Not SortOrder
50        g.col = g.MouseCol
60        If SortOrder Then
70            g.Sort = flexSortGenericAscending
80        Else
90            g.Sort = flexSortGenericDescending
100       End If
          '    Else
          '    End If

110       Exit Sub
120   ElseIf g.MouseRow > 0 Then
130       If g.MouseCol = 0 Then
140           With frmViewResultsWE
150               .grd.AddItem ""
160               .grd.RemoveItem 1
170               .grd.AddItem g.TextMatrix(g.Row, 2) & vbTab & _
                               SampleidToDOB(g.TextMatrix(g.Row, 0)) & vbTab & _
                               g.TextMatrix(g.Row, 3)
180               .grd.RemoveItem 1

190               .lblSampleID = g.TextMatrix(g.Row, 0)
200               .lblChart = g.TextMatrix(g.Row, 2)
210               .lblName = g.TextMatrix(g.Row, 3)
220               .lblDoB = SampleidToDOB(g.TextMatrix(g.Row, 0))
230               .Show 1
240           End With
250       End If
260   End If
270   Exit Sub

g_Click_Error:
      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmSignOffSamples", "g_Click", intEL, strES
End Sub
