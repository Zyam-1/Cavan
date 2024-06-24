VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatHistory 
   Caption         =   "NetAcquire - Patient Search"
   ClientHeight    =   10440
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   14925
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShort 
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Top             =   1080
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   14340
      Top             =   1020
   End
   Begin VB.Frame Frame4 
      Caption         =   "Records"
      Height          =   1065
      Left            =   13440
      TabIndex        =   17
      Top             =   30
      Width           =   1335
      Begin VB.TextBox txtRecords 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmPatHistory.frx":0000
         Top             =   240
         Width           =   765
      End
      Begin ComCtl2.UpDown udRecords 
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   570
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         _Version        =   327681
         Value           =   25
         BuddyControl    =   "txtRecords"
         BuddyDispid     =   196612
         OrigLeft        =   150
         OrigTop         =   450
         OrigRight       =   915
         OrigBottom      =   690
         Increment       =   20
         Max             =   9999
         Min             =   5
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search For"
      Height          =   1275
      Left            =   7920
      TabIndex        =   16
      Top             =   30
      Width           =   2445
      Begin VB.OptionButton oFor 
         Caption         =   "Name+DoB"
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   22
         Top             =   615
         Width           =   1125
      End
      Begin VB.CheckBox chkSoundex 
         Caption         =   "Use Soundex"
         Height          =   195
         Left            =   1020
         TabIndex        =   21
         Top             =   300
         Width           =   1305
      End
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
         Caption         =   "D.o.B."
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   20
         Top             =   930
         Width           =   825
      End
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   615
         Width           =   735
      End
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   705
      Left            =   2280
      Picture         =   "frmPatHistory.frx":0005
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   570
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   3420
      Picture         =   "frmPatHistory.frx":0447
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   570
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Appearance      =   0  'Flat
      Caption         =   "Copy to &Edit"
      Enabled         =   0   'False
      Height          =   705
      Left            =   4590
      Picture         =   "frmPatHistory.frx":0AB1
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   570
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   885
      Left            =   120
      TabIndex        =   8
      Top             =   30
      Width           =   1275
      Begin VB.OptionButton optBoth 
         Alignment       =   1  'Right Justify
         Caption         =   "Both"
         Height          =   195
         Left            =   450
         TabIndex        =   25
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optDownLoad 
         Alignment       =   1  'Right Justify
         Caption         =   "Download"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   210
         Width           =   1035
      End
      Begin VB.OptionButton optHistoric 
         Alignment       =   1  'Right Justify
         Caption         =   "Historic"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   420
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   30
      Width           =   5385
      Begin VB.CheckBox chkRemote 
         Caption         =   "Also Search Remote"
         Height          =   195
         Left            =   3510
         TabIndex        =   7
         Top             =   210
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   90
         MaxLength       =   20
         TabIndex        =   6
         Top             =   150
         Width           =   3375
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "How to Search"
      Height          =   1275
      Left            =   10410
      TabIndex        =   1
      Top             =   30
      Width           =   2985
      Begin VB.OptionButton optAll 
         Caption         =   "All Characters   (Slow Search)"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   1005
         Width           =   2775
      End
      Begin VB.OptionButton optTrailing 
         Caption         =   "Trailing Characters"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   750
         Width           =   1665
      End
      Begin VB.OptionButton optLeading 
         Caption         =   "Leading Characters"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   495
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.OptionButton optExact 
         Caption         =   "Exact Match"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.TextBox txtDoB 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   2910
      TabIndex        =   0
      Text            =   "Date of Birth"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   8865
      Left            =   0
      TabIndex        =   11
      Top             =   1470
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   15637
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   16711680
      ForeColorSel    =   65280
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmPatHistory.frx":111B
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   7920
      TabIndex        =   26
      Top             =   1320
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblDept 
      Height          =   225
      Left            =   1380
      TabIndex        =   28
      Top             =   240
      Width           =   900
   End
   Begin VB.Label lblNoPrevious 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Previous Details"
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   6330
      TabIndex        =   15
      Top             =   750
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmPatHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmPatHistory
' DateTime  : 08/01/2007 13:00
' Author    : Administrator
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Private NoPrevious As Boolean
Private mFromEdit As Boolean
Private mEditScreen As Form

Private Activated As Boolean

Private mFromLookup As Boolean

Private pWithin As Integer    'Used for fuzzy DoB search

Private SortOrder As Boolean
Private m_ShortScreen As Boolean

Public Property Let EditScreen(ByVal f As Form)

23570 Set mEditScreen = f

End Property


Private Sub FillG()

      Dim n As Integer

23580 lblNoPrevious.Visible = False

23590 With g
23600     .Rows = 2
23610     .AddItem ""
23620     .RemoveItem 1
23630 End With

23640 If Trim$(txtName) = "" Then Exit Sub

23650 If optDownLoad Then
23660     FillGridDetailsDownload 0
23670 ElseIf optHistoric Then
23680     FillGridDetailsHistoric 0
23690 Else
23700     FillGridDetailsDownload 0
23710     FillGridDetailsHistoric 0
23720 End If

23730 If sysOptRemote(0) And chkRemote.Value = 1 Then
23740     For n = 1 To intOtherHospitalsInGroup
23750         If optDownLoad Then
23760             FillGridDetailsDownload n
23770         ElseIf optHistoric Then
23780             FillGridDetailsHistoric n
23790         Else
23800             FillGridDetailsDownload n
23810             FillGridDetailsHistoric n
23820         End If
23830     Next
23840 End If

23850 With g
23860     If .Rows > 2 Then
23870         .RemoveItem 1
23880         SortOrder = True
23890         .Col = 8
23900         .Sort = 9
23910         .row = 1
23920         .Col = 7
23930         .ColSel = .Cols - 1
23940         .RowSel = 1
23950         .Highlight = flexHighlightAlways
23960     Else
23970         NoPrevious = True
23980         lblNoPrevious.Visible = True
23990     End If
24000 End With

24010 cmdCopy.Enabled = mFromEdit

24020 g.Visible = True

24030 Screen.MousePointer = 0

End Sub
Public Property Let FromEdit(ByVal X As Boolean)

24040 mFromEdit = X

End Property



Public Property Let FromLookUp(ByVal bNewValue As Boolean)

24050 mFromLookup = bNewValue

End Property

Private Function GetRemoteChart(ByVal LocalChart As String) As String

      Dim sql As String
      Dim tb As Recordset
      Dim RegionalNumber As String

24060 On Error GoTo GetRemoteChart_Error

24070 sql = "select * from PatientIFs where " & _
            "Chart = '" & LocalChart & "' " & _
            "and Entity = '" & Entity & "'"

24080 Set tb = New Recordset
24090 RecOpenClient 0, tb, sql

24100 If Not tb.EOF Then
24110     RegionalNumber = Trim$(tb!RegionalNumber & "")
24120     If RegionalNumber <> "" Then
24130         sql = "select * from PatientIFs where " & _
                    "RegionalNumber = '" & RegionalNumber & "' " & _
                    "and Entity = '" & RemoteEntity & "'"
24140         Set tb = New Recordset
24150         RecOpenClient 0, tb, sql
24160         If Not tb.EOF Then
24170             GetRemoteChart = Trim$(tb!Chart & "")
24180         End If
24190     End If
24200 End If

24210 Exit Function

GetRemoteChart_Error:

      Dim strES As String
      Dim intEL As Integer

24220 intEL = Erl
24230 strES = Err.Description
24240 LogError "frmPatHistory", "GetRemoteChart", intEL, strES, sql

End Function
Private Function LocationIsActive(ByVal Location As String, _
                                  ByVal WardClinOrGPName As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

24250 On Error GoTo LocationIsActive_Error

24260 sql = "SELECT COUNT(*) Tot FROM " & Location & " WHERE " & _
            "Text = '" & AddTicks(WardClinOrGPName) & "' " & _
            "AND InUse = 1"
24270 Set tb = New Recordset
24280 RecOpenServer 0, tb, sql

24290 LocationIsActive = tb!Tot > 0

24300 Exit Function

LocationIsActive_Error:

      Dim strES As String
      Dim intEL As Integer

24310 intEL = Erl
24320 strES = Err.Description
24330 LogError "frmPatHistory", "LocationIsActive", intEL, strES, sql

End Function

Private Sub LoadHeading(ByVal Index As Integer)

      Dim n As Integer

24340 For n = 0 To 7
24350     g.ColWidth(n) = 250
24360 Next

24370 If Index = 0 Then

24380     g.Cols = 11
24390     g.FormatString = "<Chart       |<Name                             |<Date of Birth " & _
                           "|<Sex        |<Address                                                 |<                             " & _
                           "|<Ward                         |<Clinician                      |<Hospital                " & _
                           "|<GP                           |<Admission Date "
24400     g.Col = 2
24410     g.row = 0
24420     txtDoB.Left = g.Left + g.CellLeft
24430     txtDoB.width = g.CellWidth

24440 Else
24450     g.Cols = 21
24460     g.FormatString = "E|B|M|C|H|B|G|S|<Run Date |<Run #           |<Chart      |" & _
                           "<Name                             |<Date of Birth|<Age |<Sex    |" & _
                           "<Address        |<   |<Ward           |<Clinician           |<GP                |<Hospital     |<Lab No      "

24470     If Not sysOptDeptExt(0) Then
24480         g.ColWidth(0) = 0
24490     End If
24500     If GetOptionSetting("DeptMediBridge", "0") = "0" Then
24510         g.ColWidth(1) = 0
24520     End If
24530     If Not sysOptDeptMicro(0) Then
24540         g.ColWidth(2) = 0
24550     End If
24560     If Not sysOptDeptBga(0) Then
24570         g.ColWidth(6) = 0
24580     End If
24590     If Not sysOptDeptSemen(0) Then
24600         g.ColWidth(7) = 0
24610     End If
24620     g.Col = 12
24630     g.row = 0
24640     txtDoB.Left = g.Left + g.CellLeft
24650     txtDoB.width = g.CellWidth

24660 End If

End Sub

Private Sub FillGridDetailsHistoric(ByVal CnxnNumber As Integer)

      Dim sql As String
      Dim s As String
      Dim tb As Recordset
      Dim tbBGA As Recordset
      Dim div As Integer

24670 On Error GoTo FillGridDetailsHistoric_Error

24680 div = IIf(optBoth, 2, 1)

24690 sql = "SELECT TOP " & Format$(Val(txtRecords) \ div) & " " & _
            "CASE (SELECT COUNT(*) FROM ExtResults AS X WHERE X.SampleID = D.SampleID) " & _
            "  WHEN 0 THEN 0 ELSE 1 END AS ForExt, " & _
            "CASE (SELECT COUNT(*) FROM SiteDetails50 AS M WHERE M.SampleID = D.SampleID) " & _
            "  WHEN 0 THEN 0 ELSE 1 END  as ForMicro , " & _
            "CASE (SELECT COUNT(*) FROM Demographics AS M WHERE (M.SampleID > " & sysOptSemenOffset(0) & " AND M.SampleID < " & sysOptMicroOffsetOLD(0) & ") AND M.SampleID = D.SampleID) " & _
            "  WHEN 0 THEN 0 ELSE 1 END  as ForSemen , " & _
            "CASE (SELECT COUNT(*) FROM MedibridgeResults AS M WHERE M.SampleID = D.SampleID) " & _
            "  WHEN 0 THEN 0 ELSE 1 END AS ForMedibridge, " & _
            "CASE (SELECT COUNT(*) FROM CoagResults AS C WHERE C.SampleID = D.SampleID) " & _
            "  WHEN 0 THEN 0 ELSE 1 END AS ForCoag, " & _
            "CASE (SELECT COUNT(*) FROM BioResults AS B WHERE B.SampleID = D.SampleID) " & _
            "  WHEN 0 THEN 0 ELSE 1 END AS ForBio, " & _
            "CASE (SELECT COUNT(*) FROM HaemResults AS H WHERE H.SampleID = D.SampleID) " & _
            "  WHEN 0 THEN 0 ELSE 1 END AS ForHaem, " & _
            "CASE (SELECT cFilm FROM HaemResults AS H WHERE H.SampleID = D.SampleID) " & _
            "  WHEN 1 THEN 1 ELSE 0 END AS cFilm, " & _
            "D.Rundate, D.SampleID, D.Chart, D.PatName, D.DoB, D.Age, D.Sex, D.Addr0, D.Addr1, D.Ward, D.Clinician, D.GP,ISNULL(D.labNo,0) AS labNo " & _
            "FROM Demographics AS D WHERE "

24700 If oFor(0) Then
24710     If chkSoundex = 1 Then
24720         sql = sql & "SOUNDEX(PatName) = SOUNDEX('" & AddTicks(txtName) & "') "
24730     Else
24740         If optExact Then
24750             sql = sql & "PatName = '" & AddTicks(txtName) & "' "
24760         ElseIf optLeading Then
24770             sql = sql & "PatName LIKE '" & AddTicks(txtName) & "%' "
24780         ElseIf optAll Then
24790             sql = sql & "PatName LIKE '%" & AddTicks(txtName) & "%' "
24800         Else
24810             sql = sql & "PatName LIKE '%" & AddTicks(txtName) & "' "
24820         End If
24830     End If
24840 ElseIf oFor(1) Then
24850     sql = sql & "chart = '" & AddTicks(txtName) & "' "
24860 ElseIf oFor(2) Then
24870     txtName = Convert62Date(txtName, BACKWARD)
24880     If Not IsDate(txtName) Then
24890         Screen.MousePointer = 0
24900         iMsg "Invalid Date", vbExclamation, "Date of Birth Search"
24910         Exit Sub
24920     End If
24930     sql = sql & "DoB = '" & Format$(txtName, "dd/mmm/yyyy") & "'"
24940 Else    'Name+DoB
24950     If chkSoundex = 1 Then
24960         sql = sql & "SOUNDEX(PatName) = SOUNDEX('" & AddTicks(txtName) & "') "
24970     Else
24980         If optExact Then
24990             sql = sql & "PatName = '" & AddTicks(txtName) & "' "
25000         ElseIf optLeading Then
25010             sql = sql & "PatName like '" & AddTicks(txtName) & "%' "
25020         Else
25030             sql = sql & "PatName like '%" & AddTicks(txtName) & "' "
25040         End If
25050     End If
25060     If pWithin = 0 Then
25070         sql = sql & "AND DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "' "
25080     Else
25090         sql = sql & "AND DoB BETWEEN '" & Format$(DateAdd("yyyy", -pWithin, txtDoB), "dd/mmm/yyyy") & "' " & _
                    "AND '" & Format$(DateAdd("yyyy", pWithin, txtDoB), "dd/mmm/yyyy") & "' "
25100     End If
25110 End If

      'Exclude POCT samples from the search starting at 30000000
      '+++ Junaid
      '460   SQL = SQL & " AND D.sampleid < 30000000 "
      '460   Sql = Sql & " AND D.sampleid < 3000000 "
      '--- Junaid
25120 sql = sql & " ORDER BY D.RunDate desc"

25130 NoPrevious = False
25140 Set tb = New Recordset
25150 RecOpenClient CnxnNumber, tb, sql
25160 With tb

25170     g.Visible = False
25180     Do While Not .EOF
25190         s = vbTab & vbTab & vbTab & vbTab
25200         s = s & IIf(tb!cFilm, "F", "")
25210         s = s & vbTab & vbTab & vbTab & vbTab & _
                  Format$(!Rundate, "dd/mm/yy") & vbTab & _
                  Trim$(!SampleID & "") & vbTab & _
                  !Chart & vbTab & _
                  !PatName & vbTab & _
                  Format$(!DoB, "dd/mm/yyyy") & vbTab & _
                  !Age & vbTab & _
                  !Sex & vbTab & _
                  !Addr0 & vbTab & _
                  !Addr1 & vbTab & _
                  !Ward & vbTab & _
                  !Clinician & vbTab & _
                  !GP & vbTab & _
                  vbTab & _
                  IIf(!LabNo = 0, "", !LabNo)

      '        If Trim$(!Chart & "") <> "" And Trim$(!PatName & "") <> "" Then
      '            sql = "Select * from PatientIFs where " & _
      '                  "Chart = '" & !Chart & "' " & _
      '                  "and PatName = '" & AddTicks(!PatName) & "'"
      '            Set tbBGA = New Recordset
      '            RecOpenServer CnxnNumber, tbBGA, sql
      '            If Not tbBGA.EOF Then
      '                If tbBGA!Entity & "" <> "" Then
      '                    If tbBGA!Entity = "01" Then
      '                        s = s & "Cavan"
      '                        sql = "Update Demographics " & _
      '                              "set Hospital = 'Cavan' where " & _
      '                              "Chart = '" & !Chart & "' " & _
      '                              "and PatName = '" & AddTicks(!PatName) & "' AND Hospital <> 'Cavan'"
      '                        Cnxn(CnxnNumber).Execute Sql
      '                    ElseIf tbBGA!Entity = "31" Then
      '                        s = s & "Monaghan"
      '                        sql = "Update Demographics " & _
      '                              "set Hospital = 'Monaghan' where " & _
      '                              "Chart = '" & !Chart & "' " & _
      '                              "and PatName = '" & AddTicks(!PatName) & "' AND Hospital <> 'Monaghan'"
      '                        Cnxn(CnxnNumber).Execute Sql
      '                    End If
      '                End If
      '            End If
      '        End If
25220         If Me.chkShort.Value = 1 And FndDuplicteEntery(!Chart, !PatName, Format$(!DoB, "dd/mm/yyyy"), !Sex, IIf(!LabNo = 0, "", !LabNo)) = True Then
25230         Else
25240             g.AddItem s
25250             g.row = g.Rows - 1
25260             If sysOptDeptExt(0) Then
25270             If !ForExt Then
25280                 g.Col = 0
25290                 g.CellBackColor = vbRed
25300             End If
25310         End If

25320         If GetOptionSetting("DeptMediBridge", "0") <> "0" Then
25330             If !ForMedibridge Then
25340                 g.Col = 1
25350                 g.CellBackColor = vbRed
25360             End If
25370         End If

25380         If sysOptDeptMicro(0) Then
25390             If !ForMicro Then
      '750                   g.TextMatrix(g.row, 9) = Format$(Val(g.TextMatrix(g.row, 9)) - sysOptMicroOffset(0))
25400                 g.TextMatrix(g.row, 9) = Format$(Val(g.TextMatrix(g.row, 9)))
25410                 g.Col = 2
25420                 g.CellBackColor = vbRed
                      'ElseIf lngSID > sysOptSemenOffset(0) Then
                      'Semen Result
25430             End If
25440         End If
25450         If sysOptDeptSemen(0) Then
25460             If !ForSemen Then
25470                 g.TextMatrix(g.row, 9) = Format$(Val(g.TextMatrix(g.row, 9)) - sysOptSemenOffset(0))
25480                 g.Col = 7
25490                 g.CellBackColor = vbRed
25500             End If
25510         End If
25520         If !ForCoag Then
25530             g.Col = 3
25540             g.CellBackColor = vbRed
25550         End If
25560         If !ForHaem Then
25570             g.Col = 4
25580             g.CellBackColor = vbRed
25590         End If
25600         If !ForBio Then
25610             g.Col = 5
25620             g.CellBackColor = vbRed
25630         End If

25640         If sysOptDeptBga(0) Then
25650             sql = "Select * from BGAResults where SampleID = '" & !SampleID & "'"
25660             Set tbBGA = New Recordset
25670             RecOpenServer CnxnNumber, tbBGA, sql
25680             If Not tbBGA.EOF Then
25690                 g.Col = 6
25700                 g.CellBackColor = vbRed
25710             End If
25720         End If

25730         If CnxnNumber <> 0 Then
25740             g.Col = g.Cols - 1
25750             g.CellBackColor = vbYellow
25760         End If
25770         End If

              


25780         .MoveNext
25790     Loop

25800 End With

25810 Exit Sub

FillGridDetailsHistoric_Error:

      Dim strES As String
      Dim intEL As Integer

25820 intEL = Erl
25830 strES = Err.Description
25840 LogError "frmPatHistory", "FillGridDetailsHistoric", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : FndDuplicteEntery
' Author    : XPMUser
' Date      : 24/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function FndDuplicteEntery(ChartNo As String, Name As String, DoB As String, Sex As String, LabNo As String) As Boolean

25850 On Error GoTo FndDuplicteEntery_Error
      Dim i As Integer
25860 With g
25870     For i = 1 To .Rows - 1
25880         If UCase(.TextMatrix(i, 10)) = UCase(ChartNo) And _
                 UCase(.TextMatrix(i, 11)) = UCase(Name) And _
                 UCase(.TextMatrix(i, 12)) = UCase(DoB) And _
                 UCase(.TextMatrix(i, 14)) = UCase(Sex) And _
                 UCase(.TextMatrix(i, 21)) = UCase(LabNo) Then

25890             FndDuplicteEntery = True

25900             Exit Function
25910         End If
25920     Next i
25930 End With



25940 Exit Function


FndDuplicteEntery_Error:

      Dim strES As String
      Dim intEL As Integer

25950 intEL = Erl
25960 strES = Err.Description
25970 LogError "frmPatHistory", "FndDuplicteEntery", intEL, strES
End Function


Private Sub FillGridDetailsDownload(ByVal CnxnNumber As Integer)

      Dim sql As String
      Dim s As String
      Dim tb As Recordset
      Dim div As Integer

25980 On Error GoTo FillGridDetailsDownload_Error

25990 div = IIf(optBoth, 2, 1)

26000 sql = "SELECT TOP " & Format$(Val(txtRecords) \ div) & " * FROM " & _
            "PatientIFs WHERE "

26010 If oFor(0) Then
26020     If chkSoundex = 1 Then
26030         sql = sql & "SOUNDEX(PatName) = SOUNDEX('" & AddTicks(txtName) & "') "
26040     Else
26050         If optExact Then
26060             sql = sql & "PatName = '" & AddTicks(txtName) & "' "
26070         ElseIf optLeading Then
26080             sql = sql & "PatName LIKE '" & AddTicks(txtName) & "%' "
26090         Else
26100             sql = sql & "PatName LIKE '%" & AddTicks(txtName) & "' "
26110         End If
26120     End If
26130 ElseIf oFor(1) Then
26140     sql = sql & "chart = '" & AddTicks(txtName) & "' "
26150 ElseIf oFor(2) Then
26160     txtName = Convert62Date(txtName, BACKWARD)
26170     If Not IsDate(txtName) Then
26180         Screen.MousePointer = 0
26190         iMsg "Invalid Date", vbExclamation, "Date of Birth Search"
26200         Exit Sub
26210     End If
26220     sql = sql & "DoB = '" & Format$(txtName, "dd/mmm/yyyy") & "'"
26230 Else    'Name+DoB
26240     If chkSoundex = 1 Then
26250         sql = sql & "SOUNDEX(PatName) = SOUNDEX('" & AddTicks(txtName) & "') "
26260     Else
26270         If optExact Then
26280             sql = sql & "PatName = '" & AddTicks(txtName) & "' "
26290         ElseIf optLeading Then
26300             sql = sql & "PatName like '" & AddTicks(txtName) & "%' "
26310         Else
26320             sql = sql & "PatName like '%" & AddTicks(txtName) & "' "
26330         End If
26340     End If
26350     If pWithin = 0 Then
26360         sql = sql & "AND DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "' "
26370     Else
26380         sql = sql & "AND DoB BETWEEN '" & Format$(DateAdd("yyyy", -pWithin, txtDoB), "dd/mmm/yyyy") & "' " & _
                    "AND '" & Format$(DateAdd("yyyy", pWithin, txtDoB), "dd/mmm/yyyy") & "' "
26390     End If
26400 End If

26410 sql = sql & " ORDER BY DateTimeAmended desc"

26420 NoPrevious = False

26430 Set tb = New Recordset
26440 RecOpenClient CnxnNumber, tb, sql
26450 With tb

26460     g.Visible = False
26470     Do While Not .EOF
26480         If optBoth Then
26490             s = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
26500         Else
26510             s = ""
26520         End If
26530         s = s & !Chart & vbTab & _
                  !PatName & vbTab & _
                  Format$(!DoB, "dd/mm/yyyy") & vbTab
26540         If optBoth Then
26550             s = s & vbTab
26560         End If
26570         s = s & !Sex & vbTab & _
                  !Address0 & vbTab & _
                  !Address1 & vbTab & _
                  !Ward & vbTab & _
                  !Clinician & vbTab
26580         If optBoth Then
26590             s = s & !GP & vbTab
26600             If !Entity & "" <> "" Then
26610                 s = s & IIf(!Entity = "01", "Cavan", "Monaghan")
26620             End If
26630         Else
26640             If !Entity & "" <> "" Then
26650                 s = s & IIf(!Entity = "01", "Cavan", "Monaghan")
26660             End If
26670             If CnxnNumber <> 0 Then
26680                 s = s & HospName(CnxnNumber)
26690             End If
26700             s = s & vbTab & !GP & vbTab
26710             If IsDate(!AdmitDate) Then
26720                 s = s & Format$(!AdmitDate, "dd/MM/yyyy")
26730             End If
26740         End If
26750         g.AddItem s
26760         .MoveNext
26770     Loop
26780 End With

26790 Exit Sub

FillGridDetailsDownload_Error:

      Dim strES As String
      Dim intEL As Integer

26800 intEL = Erl
26810 strES = Err.Description
26820 LogError "frmPatHistory", "FillGridDetailsDownload", intEL, strES, sql

End Sub


Public Property Get NoPreviousDetails() As Variant

26830 NoPreviousDetails = NoPrevious

End Property


Private Sub bsearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26840 pBar = 0

End Sub

Private Sub chkRemote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26850 pBar = 0

End Sub


Private Sub chkSoundex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26860 pBar = 0

End Sub

Private Sub cmdCancel_Click()

26870 Unload Me

End Sub


Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26880 pBar = 0

End Sub


Private Sub cmdCopy_Click()

      Dim gRow As Integer
      Dim strWard As String
      Dim strGP As String
      Dim strClinician As String
      Dim strSex As String
      Dim strName As String
      Dim strDOB As String
      Dim strCHART As String

26890 On Error GoTo cmdCopy_Click_Error

26900 gRow = g.row

26910 With mEditScreen
26920     If optHistoric Or optBoth Then
26930         If .txtChart = "" Then
26940             .txtChart = g.TextMatrix(gRow, 10)
26950         End If
26960         strCHART = Trim$(g.TextMatrix(gRow, 10))
26970         strName = Initial2Upper(g.TextMatrix(gRow, 11))
26980         .txtSurName = UCase(SurName(strName))
26990         .txtForeName = UCase(ForeName(strName))
27000         .txtDoB = g.TextMatrix(gRow, 12)

27010         strDOB = g.TextMatrix(gRow, 12)
27020         .txtAge = CalcAge(.txtDoB, Now)
27030         strSex = g.TextMatrix(gRow, 14)
27040         If strSex = "" Then
27050             NameLostFocus SurName(strName), ForeName(strName), strSex
27060         End If
27070         .txtSex = strSex
27080         .txtAddress(0) = Initial2Upper(g.TextMatrix(gRow, 15))
27090         .txtAddress(1) = Initial2Upper(g.TextMatrix(gRow, 16))
27100         strWard = Initial2Upper(g.TextMatrix(gRow, 17))
27110         strGP = Initial2Upper(g.TextMatrix(gRow, 19))
27120         If strWard = "" And strGP <> "" Then
27130             strWard = "GP"
27140         End If
27150         If LocationIsActive("Wards", strWard) Then
27160             .cmbWard = strWard
27170         Else
27180             .cmbWard = ""
27190         End If
27200         If LocationIsActive("GPs", strGP) Then
27210             .cmbGP = strGP
27220         Else
27230             .cmbGP = ""
27240         End If

27250         strClinician = Initial2Upper(g.TextMatrix(gRow, 18))
27260         If LocationIsActive("Clinicians", strClinician) Then
27270             .cmbClinician = strClinician
27280         Else
27290             .cmbClinician = ""
27300         End If
27310         .txtLabNo = g.TextMatrix(gRow, 21)
      '        If g.TextMatrix(gRow, 21) <> "" And g.TextMatrix(gRow, 21) <> "0" Then
      '            .txtLabNo = g.TextMatrix(gRow, 21)
      '        End If
27320         LabNoUpdatePrviousData = ""
27330         If strCHART = "" Then
27340             LabNoUpdatePrviousData = "1"    'UCase(.txtSurName & .txtForeName & .txtDoB)
27350         End If
27360     ElseIf optDownLoad Then
27370         .txtChart = g.TextMatrix(gRow, 0)

27380         .txtSurName = UCase(SurName(Initial2Upper(g.TextMatrix(gRow, 1))))
27390         .txtForeName = UCase(ForeName(Initial2Upper(g.TextMatrix(gRow, 1))))
27400         .txtDoB = g.TextMatrix(gRow, 2)
27410         .txtAge = CalcAge(.txtDoB, Now)
27420         .txtSex = g.TextMatrix(gRow, 3)
27430         .txtAddress(0) = Initial2Upper(g.TextMatrix(gRow, 4))
27440         .txtAddress(1) = Initial2Upper(g.TextMatrix(gRow, 5))

27450         strWard = Initial2Upper(g.TextMatrix(gRow, 6))
27460         If LocationIsActive("Wards", strWard) Then
27470             .cmbWard = strWard
27480         Else
27490             .cmbWard = ""
27500         End If

27510         strClinician = Initial2Upper(g.TextMatrix(gRow, 7))
27520         If LocationIsActive("Clinicians", strClinician) Then
27530             .cmbClinician = strClinician
27540         Else
27550             .cmbClinician = ""
27560         End If

27570     End If

27580     Call CopyCCfromHistory(mEditScreen.txtSampleID, lblDept, strName, strDOB, strCHART)

27590 End With

27600 Unload Me

27610 Exit Sub

cmdCopy_Click_Error:

      Dim strES As String
      Dim intEL As Integer

27620 intEL = Erl
27630 strES = Err.Description
27640 LogError "frmPatHistory", "cmdCopy_Click", intEL, strES

End Sub

Private Sub CopyCCfromHistory(ByVal strCurrentSampleID As String, ByVal strDept As String, ByVal strName As String, ByVal strDOB As String, _
                              ByVal strCHART As String)

      Dim sql As String
      Dim tb As Recordset
      Dim sn As Recordset
      Dim cc As Recordset
      Dim strCurrentSID As String

27650 On Error GoTo CopyCCfromHistory_Click_Error

27660 If strDept = "M" Then
27670     strCurrentSID = Val(strCurrentSampleID) ' + sysOptMicroOffset(0)
27680 Else
27690     strCurrentSID = Val(strCurrentSampleID)
27700 End If

      'Select Last entry details for patient
27710 sql = "select top 1 SampleID from demographics where PatName = '" & AddTicks(strName) & "' and DoB = '" & Format(strDOB, "dd/mmm/yyyy") & "' "
27720 sql = sql & "and Chart = '" & strCHART & "' order by SampleDate desc"
27730 Set tb = New Recordset
27740 RecOpenClient 0, tb, sql
27750 If Not tb.EOF Then
27760     sql = "Select * from SendCopyTo where " & _
                "SampleID = '" & Val(tb!SampleID) & "'"
27770     Set sn = New Recordset
27780     RecOpenServer 0, sn, sql
27790     If Not sn.EOF Then
27800         sql = "Select * from SendCopyTo where " & _
                    "SampleID = '" & Val(strCurrentSID) & "'"
27810         Set cc = New Recordset
27820         RecOpenServer 0, cc, sql
27830         If cc.EOF Then
27840             cc.AddNew
27850             cc!SampleID = Trim$(strCurrentSID)
27860             cc!Ward = sn!Ward & ""
27870             cc!Clinician = sn!Clinician & ""
27880             cc!GP = sn!GP & ""
27890             cc!Device = sn!Device & ""
27900             cc!Destination = sn!Destination & ""
27910             cc.Update
27920         End If
27930     End If
27940 End If

27950 Exit Sub

CopyCCfromHistory_Click_Error:

      Dim strES As String
      Dim intEL As Integer

27960 intEL = Erl
27970 strES = Err.Description
27980 LogError "frmPatHistory", "CopyCCfromHistory", intEL, strES
End Sub




Private Sub bsearch_Click()

27990 FillG

End Sub


Private Sub chkSoundex_Click()

28000 If chkSoundex = 1 Then
28010     fraSearch.Visible = False
28020 Else
28030     fraSearch.Visible = True
28040 End If

28050 FillG

End Sub


Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28060 pBar = 0

End Sub

Private Sub Form_Activate()

28070 TimerBar.Enabled = True
28080 pBar = 0

28090 Me.Caption = "NetAcquire - Patient Search (" & HospName(0) & ")"
28100 If LogOffDelaySecs > 0 Then
28110     pBar.max = LogOffDelaySecs
28120 End If

28130 If Activated Then Exit Sub

28140 Activated = True

28150 txtName.SetFocus

End Sub

Private Sub Form_Deactivate()

28160 pBar = 0
28170 TimerBar.Enabled = False

End Sub


Private Sub Form_Load()

28180 Activated = False

28190 If optDownLoad Then
28200     LoadHeading 0
28210 Else
28220     LoadHeading 1
28230 End If

28240 chkRemote.Visible = False
28250 chkRemote.Value = 0

28260 If sysOptRemote(0) Then
28270     If HospName(1) <> "" Then
28280         chkRemote.Visible = True
28290         chkRemote.Value = Val(GetSetting("NetAcquire", "PatSearch", "cRemote", "1"))
28300     End If
28310 End If
28320 LabNoUpdatePrviousData = ""

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28330 pBar = 0

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : XPMUser
' Date      : 24/Sep/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub Form_Resize()
'On Error GoTo Form_Resize_Error
'
'Dim i As Integer
'
'If Me.ShortScreen = True Then
'    Me.Width = 7935
'    With g
'        .Width = Me.Width - 100
'        For i = 0 To 9
'            .ColWidth(i) = 0
'        Next i
'        .ColWidth(13) = 0
'        For i = 15 To 20
'            .ColWidth(i) = 0
'        Next i
'    End With
'Else
'    g.Width = 14805
'    Me.Width = 15030
'End If
'
'
'Exit Sub
'
'
'Form_Resize_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "frmPatHistory", "Form_Resize", intEL, strES
'End Sub

Private Sub Form_Unload(Cancel As Integer)
28340 chkShort.Value = 0
28350 Activated = False

28360 SaveSetting "NetAcquire", "PatSearch", "cRemote", Format$(chkRemote.Value)

End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28370 pBar = 0

End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28380 pBar = 0

End Sub


Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28390 pBar = 0

End Sub


Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28400 pBar = 0

End Sub


Private Sub fraSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28410 pBar = 0

End Sub


Private Sub g_Click()
      '
      'MsgBox g.Col
      'Exit Sub

      Dim tb As Recordset
      Dim sql As String
      Dim NewChart As String
      Dim PatName As String
      Dim DoB As String
      Dim f As Form

28420 On Error GoTo g_Click_Error

28430 If g.MouseRow = 0 Then
28440     If InStr(UCase$(g.TextMatrix(0, g.Col)), "DATE") <> 0 Then
28450         g.Sort = 9
28460     Else
28470         If SortOrder Then
28480             g.Sort = flexSortGenericAscending
28490         Else
28500             g.Sort = flexSortGenericDescending
28510         End If
28520     End If
28530     SortOrder = Not SortOrder
28540     Exit Sub
28550 End If

28560 If optDownLoad Then
28570     If g.Col > 1 And mFromEdit Then
28580         g.Col = 0
28590         g.ColSel = g.Cols - 1
28600         g.RowSel = g.row
28610         g.Highlight = flexHighlightAlways
28620         cmdCopy.Enabled = True
28630     ElseIf g.Col = 0 Then
28640         If Trim$(g.TextMatrix(g.row, 0)) = "" Then Exit Sub
28650         PatName = g.TextMatrix(g.row, 1)
28660         If Trim$(PatName) = "" Then Exit Sub
28670         DoB = g.TextMatrix(g.row, 2)
28680         If IsDate(DoB) Then
28690             DoB = Format$(DoB, "dd/mmm/yyyy")
28700         Else
28710             Exit Sub
28720         End If
28730         If iMsg("Do you want to change the Chart Number?", vbQuestion + vbYesNo) = vbYes Then
28740             NewChart = iBOX("New Chart Number", , g.TextMatrix(g.row, 0))
28750             sql = "Update Demographics " & _
                        "set Chart = '" & NewChart & "' where " & _
                        "PatName = '" & AddTicks(PatName) & "' " & _
                        "and dob = '" & DoB & "'"
28760             Set tb = New Recordset
28770             RecOpenClient 0, tb, sql
28780             FillG
28790         End If
28800     End If
28810 Else
28820     If g.Col = 10 Then    'chart
28830         If Trim$(g.TextMatrix(g.row, 10)) = "" Then Exit Sub
28840         PatName = g.TextMatrix(g.row, 11)
28850         If Trim$(PatName) = "" Then Exit Sub
28860         DoB = g.TextMatrix(g.row, 12)
28870         If IsDate(DoB) Then
28880             DoB = Format$(DoB, "dd/mmm/yyyy")
28890         Else
28900             Exit Sub
28910         End If
28920         If iMsg("Do you want to change the Chart Number?", vbQuestion + vbYesNo) = vbYes Then
28930             NewChart = iBOX("New Chart Number", , g.TextMatrix(g.row, 10))
28940             sql = "Update Demographics " & _
                        "set Chart = '" & NewChart & "' where " & _
                        "PatName = '" & AddTicks(PatName) & "' " & _
                        "and dob = '" & DoB & "'"
28950             Set tb = New Recordset
28960             RecOpenClient 0, tb, sql
28970             FillG
28980         End If
28990     ElseIf g.Col > 7 Then
29000         g.Col = 8
29010         g.ColSel = g.Cols - 1
29020         g.RowSel = g.row
29030         g.Highlight = flexHighlightAlways
29040         cmdCopy.Enabled = True
29050         If mFromEdit Then
29060             If g.TextMatrix(g.row, 20) = "Monaghan" Then
29070                 cmdCopy.Enabled = False
29080             End If
29090         End If
29100     Else
29110         cmdCopy.Enabled = False

29120         If g.CellBackColor <> vbRed Then Exit Sub

29130         If g.Col = 0 Then    'External
29140             With frmExternalReport
29150                 .lblChart = g.TextMatrix(g.row, 10)
29160                 .lblName = g.TextMatrix(g.row, 11)
29170                 .lblDoB = g.TextMatrix(g.row, 12)
29180                 .Show 1
29190             End With
29200         ElseIf g.Col = 1 Then
29210             With frmViewMedibridge
29220                 .SampleID = g.TextMatrix(g.row, 9)
29230                 .Show 1
29240             End With
29250         ElseIf g.Col = 2 Then    'Micro
29260             Set f = New frmMicroReport
29270             With f
29280                 .FromEdit = False
29290                 .lblChart = g.TextMatrix(g.row, 10)
29300                 .lblName = g.TextMatrix(g.row, 11)
29310                 .lblDoB = g.TextMatrix(g.row, 12)
29320                 .lblSex = Trim$(Left$(g.TextMatrix(g.row, 14) & " ", 1))
29330                 .Show 1
29340             End With
29350             Unload f
29360             Set f = Nothing
29370         ElseIf g.Col = 7 Then   'Semen
29380             With frmSemenReport
29390                 .lblChart = g.TextMatrix(g.row, 10)
29400                 .lblName = g.TextMatrix(g.row, 11)
29410                 .lblDoB = g.TextMatrix(g.row, 12)
29420                 .lblSex = Trim$(Left$(g.TextMatrix(g.row, 14) & " ", 1))
29430                 .Show 1
29440             End With
29450         Else
29460             With frmViewResults
29470                 .lblSampleID = g.TextMatrix(g.row, 9)
29480                 .lblChart = g.TextMatrix(g.row, 10)
29490                 .lblName = g.TextMatrix(g.row, 11)
29500                 .lblDoB = g.TextMatrix(g.row, 12)
29510                 .lblAge = g.TextMatrix(g.row, 13)
29520                 If Left$(g.TextMatrix(g.row, 14), 1) = "M" Then
29530                     .lblSex = "Male"
29540                 ElseIf Left$(g.TextMatrix(g.row, 14), 1) = "F" Then
29550                     .lblSex = "Female"
29560                 Else
29570                     .lblSex = ""
29580                 End If
29590                 .lblAddress = g.TextMatrix(g.row, 15) & " " & g.TextMatrix(g.row, 16)
29600                 .lblWard = g.TextMatrix(g.row, 17)
29610                 .lblGP = g.TextMatrix(g.row, 19)
29620                 .Show 1
29630             End With
29640         End If
29650     End If
29660 End If

29670 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

29680 intEL = Erl
29690 strES = Err.Description
29700 LogError "frmPatHistory", "g_Click", intEL, strES, sql

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

29710 If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
29720     Cmp = 0
29730     Exit Sub
29740 End If

29750 If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
29760     Cmp = 0
29770     Exit Sub
29780 End If

29790 d1 = Format(g.TextMatrix(Row1, g.Col), "dd/mmm/yyyy hh:mm:ss")
29800 d2 = Format(g.TextMatrix(Row2, g.Col), "dd/mmm/yyyy hh:mm:ss")

29810 If SortOrder Then
29820     Cmp = Sgn(DateDiff("s", d1, d2))
29830 Else
29840     Cmp = Sgn(DateDiff("s", d2, d1))
29850 End If

End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29860 pBar = 0

End Sub

Private Sub oFor_Click(Index As Integer)

      Dim f As Form

29870 Select Case Index

      Case 0: optLeading = True
29880     chkSoundex.Enabled = True
29890     txtDoB.Visible = False

29900 Case 1, 2: optExact = True
29910     chkSoundex.Enabled = False
29920     chkSoundex = 0
29930     txtDoB.Visible = False

29940 Case 3: optLeading = True
29950     chkSoundex.Enabled = True

29960     Set f = New frmGetDoB
29970     f.Show 1
29980     txtDoB = f.txtDoB
29990     If f.lblWithin.Enabled Then
30000         pWithin = f.lblWithin
30010     Else
30020         pWithin = 0
30030     End If
30040     Unload f
30050     Set f = Nothing

30060     If Not IsDate(txtDoB) Then
30070         oFor(0) = 1
30080         Exit Sub
30090     End If
30100     txtDoB.Visible = True

30110     If txtName.Visible Then
30120         txtName.SetFocus
30130     End If

30140 End Select


30150 g.Rows = 2
30160 g.AddItem ""
30170 g.RemoveItem 1
30180 txtName = ""

End Sub

Private Sub oFor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

30190 pBar = 0

End Sub


Private Sub optBoth_Click()

30200 g.Rows = 2
30210 g.AddItem ""
30220 g.RemoveItem 1

30230 LoadHeading 1

30240 If Not Activated Then Exit Sub

30250 FillG

End Sub


Private Sub optBoth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30260 pBar = 0

End Sub


Private Sub optDownLoad_Click()

30270 g.Rows = 2
30280 g.AddItem ""
30290 g.RemoveItem 1

30300 LoadHeading 0

30310 If Not Activated Then Exit Sub

30320 FillG

End Sub

Private Sub optDownLoad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30330 pBar = 0

End Sub


Private Sub optExact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30340 pBar = 0

End Sub


Private Sub optHistoric_Click()

30350 g.Rows = 2
30360 g.AddItem ""
30370 g.RemoveItem 1

30380 LoadHeading 1

30390 If Not Activated Then Exit Sub

30400 FillG

End Sub


Private Sub optHistoric_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30410 pBar = 0

End Sub


Private Sub optLeading_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30420 pBar = 0

End Sub


Private Sub optTrailing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30430 pBar = 0

End Sub


Private Sub pBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30440 pBar = 0

End Sub


Private Sub TimerBar_Timer()

30450 pBar = pBar + 1

30460 If pBar = pBar.max Then
30470     Unload Me
30480     Exit Sub
30490 End If

End Sub


Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)

30500 pBar = 0

30510 If oFor(0) Or oFor(3) Then
30520     If Len(Trim$(txtName)) > 3 Then
30530         FillG
30540     End If
30550 End If

End Sub


Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30560 pBar = 0

End Sub


Private Sub txtRecords_KeyPress(KeyAscii As Integer)

30570 pBar = 0

End Sub


Private Sub txtRecords_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30580 pBar = 0

End Sub


Private Sub udRecords_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30590 pBar = 0

End Sub

Private Sub udRecords_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

30600 FillG

End Sub



Public Property Get ShortScreen() As Boolean

30610 ShortScreen = m_ShortScreen

End Property

Public Property Let ShortScreen(ByVal bShortScreen As Boolean)

30620 m_ShortScreen = bShortScreen

End Property
