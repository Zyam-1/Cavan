VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatsExtInt 
   Caption         =   "NetAcquire"
   ClientHeight    =   4665
   ClientLeft      =   375
   ClientTop       =   825
   ClientWidth     =   10155
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   10155
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   6510
      Picture         =   "frmStatsExtInt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3540
      Width           =   1185
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   675
      Left            =   360
      Picture         =   "frmStatsExtInt.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3540
      Width           =   1185
   End
   Begin VB.CommandButton cmdSetLocations 
      Caption         =   "Set Locations"
      Height          =   405
      Left            =   8220
      TabIndex        =   5
      Top             =   240
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   7830
      TabIndex        =   0
      Top             =   960
      Width           =   1965
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   870
         Width           =   1365
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "Calculate"
         Height          =   825
         Left            =   420
         Picture         =   "frmStatsExtInt.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1380
         Width           =   1125
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Text            =   "cmbMonth"
         Top             =   510
         Width           =   1365
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   990
         TabIndex        =   3
         Top             =   180
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2005
         Alignment       =   0
         BuddyControl    =   "lblYear"
         BuddyDispid     =   196616
         OrigLeft        =   900
         OrigTop         =   180
         OrigRight       =   1515
         OrigBottom      =   495
         Max             =   2015
         Min             =   2000
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblYear 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2005"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   4
         Top             =   180
         Width           =   705
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2775
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   11
      Cols            =   7
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmStatsExtInt.frx":0C7E
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
   Begin MSComctlLib.ProgressBar pb 
      Height          =   195
      Left            =   7830
      TabIndex        =   13
      Top             =   3270
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not Specified"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   2
      Left            =   5730
      TabIndex        =   12
      Top             =   210
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "External"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   3900
      TabIndex        =   11
      Top             =   210
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "In-House"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Index           =   0
      Left            =   2070
      TabIndex        =   10
      Top             =   210
      Width           =   1815
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
      Left            =   1590
      TabIndex        =   9
      Top             =   3660
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7680
      Picture         =   "frmStatsExtInt.frx":0D65
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmStatsExtInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strNotSpecified As String
Private strInHouse As String
Private strExternal As String
Private strUnknown As String
Private Sub FillcmbDate()

      Dim DaysInMonth As Integer
      Dim n As Integer

60210 DaysInMonth = Day(DateSerial(lblYear, cmbMonth.ListIndex + 2, 0))

60220 cmbDate.Clear

60230 cmbDate.AddItem "Whole Month"
60240 For n = 1 To DaysInMonth
60250   cmbDate.AddItem Format$(n)
60260 Next
60270 cmbDate.ListIndex = 0

End Sub

Private Sub GenerateLists()

      Dim tb As Recordset
      Dim sql As String

60280 On Error GoTo GenerateLists_Error

60290 sql = "Select Text from Wards where " & _
            "Location <> 'In-House' and Location <> 'External'"
60300 Set tb = New Recordset
60310 RecOpenServer 0, tb, sql
60320 strUnknown = ""
60330 strNotSpecified = ""
60340 Do While Not tb.EOF
60350   strNotSpecified = strNotSpecified & " Ward = '" & AddTicks(tb!Text) & "' or"
60360   strUnknown = strUnknown & " Ward <> '" & AddTicks(tb!Text) & "' and"
60370   tb.MoveNext
60380 Loop
60390 If Len(strNotSpecified) > 0 Then
60400   strNotSpecified = Left$(strNotSpecified, Len(strNotSpecified) - 3)
60410 End If

60420 sql = "Select Text from Wards where " & _
            "Location = 'In-House'"
60430 Set tb = New Recordset
60440 RecOpenServer 0, tb, sql
60450 strInHouse = ""
60460 Do While Not tb.EOF
60470   strInHouse = strInHouse & " Ward = '" & AddTicks(tb!Text) & "' or"
60480   strUnknown = strUnknown & " Ward <> '" & AddTicks(tb!Text) & "' and"
60490   tb.MoveNext
60500 Loop
60510 If Len(strInHouse) > 0 Then
60520   strInHouse = Left$(strInHouse, Len(strInHouse) - 3)
60530 End If

60540 sql = "Select Text from Wards where " & _
            "Location = 'External'"
60550 Set tb = New Recordset
60560 RecOpenServer 0, tb, sql
60570 strExternal = ""
60580 Do While Not tb.EOF
60590   strExternal = strExternal & " Ward = '" & AddTicks(tb!Text) & "' or"
60600   strUnknown = strUnknown & " Ward <> '" & AddTicks(tb!Text) & "' and"
60610   tb.MoveNext
60620 Loop
60630 If Len(strExternal) > 0 Then
60640   strExternal = Left$(strExternal, Len(strExternal) - 3)
60650 End If
60660 If Len(strUnknown) > 0 Then
60670   strUnknown = Left$(strUnknown, Len(strUnknown) - 3)
60680 End If

60690 Exit Sub

GenerateLists_Error:

      Dim strES As String
      Dim intEL As Integer

60700 intEL = Erl
60710 strES = Err.Description
60720 LogError "frmStatsExtInt", "GenerateLists", intEL, strES, sql


End Sub

Private Sub cmbMonth_Click()

60730 FillcmbDate

End Sub

Private Sub cmdCalculate_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim MonthNumber As Integer

60740 On Error GoTo cmdCalculate_Click_Error

60750 g.Rows = 2
60760 g.AddItem ""
60770 g.RemoveItem 1
60780 g.FormatString = "|<Tests      |<Samples  " & _
                       "|<Tests      |<Samples  " & _
                       "|<Tests      |<Samples  " & _
                       ";Discipline            " & _
                       "|Biochemistry" & _
                       "|Endocrinology" & _
                       "|Coagulation" & _
                       "|Haem - FBC" & _
                       "|Haem - ESR" & _
                       "|Haem - MonoSpot" & _
                       "|Haem - Malaria" & _
                       "|Haem - Sickledex" & _
                       "|Haem - RA" & _
                       "|Haem - Films"

60790 MonthNumber = cmbMonth.ListIndex + 1

60800 FromDate = Format$("01/" & Format$(MonthNumber) & "/" & lblYear, "dd/mmm/yyyy")
60810 ToDate = Format$(DateAdd("m", 1, FromDate) - 1, "dd/mmm/yyyy")

60820 sql = "Select SampleID from Demographics where " & _
            "" & _
            "RunDate between '" & FromDate & "' and '" & ToDate & "'"
60830 Set tb = New Recordset
60840 RecOpenClient 0, tb, sql
60850 pb = 0
60860 If tb.EOF Then Exit Sub

60870 pb.max = 33
60880 pb.Visible = True

60890 CalcBio
60900 CalcCoag
60910 CalcFBC "RBC", 4
60920 CalcFBC "ESR", 5
60930 CalcFBC "MonoSpot", 6
60940 CalcFBC "Malaria", 7
60950 CalcFBC "Sickledex", 8
60960 CalcFBC "RA", 9
60970 CalcFBC "cFilm", 10

60980 pb.Visible = False

60990 Exit Sub

cmdCalculate_Click_Error:

      Dim strES As String
      Dim intEL As Integer

61000 intEL = Erl
61010 strES = Err.Description
61020 LogError "frmStatsExtInt", "cmdCalculate_Click", intEL, strES, sql


End Sub


Private Sub CalcBio()

      Dim tbC As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim MonthNumber As Integer
      Dim lngTests As Long
      Dim lngSamples As Long

61030 On Error GoTo CalcBio_Error

61040 MonthNumber = cmbMonth.ListIndex + 1

61050 FromDate = Format$("01/" & Format$(MonthNumber) & "/" & lblYear, "dd/mmm/yyyy")
61060 ToDate = Format$(DateAdd("m", 1, FromDate) - 1, "dd/mmm/yyyy")
61070 If cmbDate <> "Whole Month" Then
61080   FromDate = Format$(cmbDate & "/" & Format$(MonthNumber) & "/" & lblYear, "dd/mmm/yyyy")
61090   ToDate = FromDate
61100 End If

61110 If strNotSpecified <> "" Then
61120   lngTests = 0
61130   lngSamples = 0
61140   pb = 1
61150   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from BioResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strNotSpecified & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') "
61160   If UCase$(HospName(0)) = "CAVAN" Then
61170     sql = sql & "and (Analyser = 'A' or Analyser = 'B')"
61180   End If
61190   Set tbC = New Recordset
61200   RecOpenServer 0, tbC, sql
61210   lngTests = tbC!Tests
61220   lngSamples = tbC!Samples
61230   g.TextMatrix(1, 5) = Format$(lngTests)
61240   g.TextMatrix(1, 6) = Format$(lngSamples)
61250   g.Refresh
61260 End If
61270 If strUnknown <> "" Then
61280   lngTests = 0
61290   lngSamples = 0
61300   pb = 1
61310   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from BioResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strUnknown & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') "
61320   If UCase$(HospName(0)) = "CAVAN" Then
61330     sql = sql & "and (Analyser = 'A' or Analyser = 'B')"
61340   End If
61350   Set tbC = New Recordset
61360   RecOpenServer 0, tbC, sql
61370   lngTests = tbC!Tests
61380   lngSamples = tbC!Samples
61390   g.TextMatrix(1, 5) = Format$(Val(g.TextMatrix(1, 5)) + lngTests)
61400   g.TextMatrix(1, 6) = Format$(Val(g.TextMatrix(1, 6)) + lngSamples)
61410   g.Refresh
61420 End If



61430 If strNotSpecified <> "" Then
61440   lngTests = 0
61450   lngSamples = 0
61460   pb = 2
61470   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from BioResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strNotSpecified & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') " & _
              "and (Analyser = '4')"
61480   Set tbC = New Recordset
61490   RecOpenServer 0, tbC, sql
61500   lngTests = tbC!Tests
61510   lngSamples = tbC!Samples
61520   g.TextMatrix(2, 5) = Format$(lngTests)
61530   g.TextMatrix(2, 6) = Format$(lngSamples)
61540   g.Refresh
61550 End If

61560 If strInHouse <> "" Then
61570   lngTests = 0
61580   lngSamples = 0
61590   pb = 3
61600   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from BioResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strInHouse & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') "
61610   If UCase$(HospName(0)) = "CAVAN" Then
61620     sql = sql & "and (Analyser = 'A' or Analyser = 'B')"
61630   End If
61640   Set tbC = New Recordset
61650   RecOpenServer 0, tbC, sql
61660   lngTests = tbC!Tests
61670   lngSamples = tbC!Samples
61680   g.TextMatrix(1, 1) = Format$(lngTests)
61690   g.TextMatrix(1, 2) = Format$(lngSamples)
61700   g.Refresh
61710 End If

61720 If strInHouse <> "" Then
61730   lngTests = 0
61740   lngSamples = 0
61750   pb = 4
61760   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from BioResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strInHouse & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') " & _
              "and (Analyser = '4')"
61770   Set tbC = New Recordset
61780   RecOpenServer 0, tbC, sql
61790   lngTests = tbC!Tests
61800   lngSamples = tbC!Samples
61810   g.TextMatrix(2, 1) = Format$(lngTests)
61820   g.TextMatrix(2, 2) = Format$(lngSamples)
61830   g.Refresh
61840 End If

61850 If strExternal <> "" Then
61860   lngTests = 0
61870   lngSamples = 0
61880   pb = 5
61890   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from BioResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strExternal & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') "
61900   If UCase$(HospName(0)) = "CAVAN" Then
61910     sql = sql & "and (Analyser = 'A' or Analyser = 'B')"
61920   End If
61930   Set tbC = New Recordset
61940   RecOpenServer 0, tbC, sql
61950   lngTests = tbC!Tests
61960   lngSamples = tbC!Samples
61970   g.TextMatrix(1, 3) = Format$(lngTests)
61980   g.TextMatrix(1, 4) = Format$(lngSamples)
61990   g.Refresh
62000 End If

62010 If strExternal <> "" Then
62020   lngTests = 0
62030   lngSamples = 0
62040   pb = 6
62050   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from BioResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strExternal & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') " & _
              "and (Analyser = '4')"
62060   Set tbC = New Recordset
62070   RecOpenServer 0, tbC, sql
62080   lngTests = tbC!Tests
62090   lngSamples = tbC!Samples
62100   g.TextMatrix(2, 3) = Format$(lngTests)
62110   g.TextMatrix(2, 4) = Format$(lngSamples)
62120   g.Refresh
62130 End If

62140 Exit Sub

CalcBio_Error:

      Dim strES As String
      Dim intEL As Integer

62150 intEL = Erl
62160 strES = Err.Description
62170 LogError "frmStatsExtInt", "CalcBio", intEL, strES, sql


End Sub

Private Sub CalcCoag()

      Dim tbC As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim MonthNumber As Integer
      Dim lngTests As Long
      Dim lngSamples As Long

62180 On Error GoTo CalcCoag_Error

62190 MonthNumber = cmbMonth.ListIndex + 1

62200 FromDate = Format$("01/" & Format$(MonthNumber) & "/" & lblYear, "dd/mmm/yyyy")
62210 ToDate = Format$(DateAdd("m", 1, FromDate) - 1, "dd/mmm/yyyy")
62220 If cmbDate <> "Whole Month" Then
62230   FromDate = Format$(cmbDate & "/" & Format$(MonthNumber) & "/" & lblYear, "dd/mmm/yyyy")
62240   ToDate = FromDate
62250 End If

62260 If strInHouse <> "" Then
62270   lngTests = 0
62280   lngSamples = 0
62290   pb = 7
62300   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from CoagResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strInHouse & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "')"
62310   Set tbC = New Recordset
62320   RecOpenServer 0, tbC, sql
62330   lngTests = tbC!Tests
62340   lngSamples = tbC!Samples
        'g.TextMatrix(3, 1) = Format$(lngTests)
62350   g.TextMatrix(3, 2) = Format$(lngSamples)
62360   g.Refresh
62370 End If

62380 If strExternal <> "" Then
62390   lngTests = 0
62400   lngSamples = 0
62410   pb = 8
62420   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from CoagResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strExternal & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "')"
62430   Set tbC = New Recordset
62440   RecOpenServer 0, tbC, sql
62450   lngTests = tbC!Tests
62460   lngSamples = tbC!Samples
        'g.TextMatrix(3, 3) = Format$(lngTests)
62470   g.TextMatrix(3, 4) = Format$(lngSamples)
62480   g.Refresh
62490 End If

62500 If strNotSpecified <> "" Then
62510   lngTests = 0
62520   lngSamples = 0
62530   pb = 9
62540   sql = "Select count (SampleID) as Tests, " & _
              "count (distinct SampleID) as Samples from CoagResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strNotSpecified & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "')"
62550   Set tbC = New Recordset
62560   RecOpenServer 0, tbC, sql
62570   lngTests = tbC!Tests
62580   lngSamples = tbC!Samples
        'g.TextMatrix(3, 5) = Format$(lngTests)
62590   g.TextMatrix(3, 6) = Format$(lngSamples)
62600   g.Refresh
62610 End If

62620 Exit Sub

CalcCoag_Error:

      Dim strES As String
      Dim intEL As Integer

62630 intEL = Erl
62640 strES = Err.Description
62650 LogError "frmStatsExtInt", "CalcCoag", intEL, strES, sql


End Sub

Private Sub CalcFBC(ByVal Parameter As String, _
                    ByVal StartRow As Integer)

      Dim tbC As Recordset
      Dim sql As String
      Dim FromDate As String
      Dim ToDate As String
      Dim MonthNumber As Integer
      Dim lngTests As Long

62660 On Error GoTo CalcFBC_Error

62670 MonthNumber = cmbMonth.ListIndex + 1

62680 FromDate = Format$("01/" & Format$(MonthNumber) & "/" & lblYear, "dd/mmm/yyyy")
62690 ToDate = Format$(DateAdd("m", 1, FromDate) - 1, "dd/mmm/yyyy")
62700 If cmbDate <> "Whole Month" Then
62710   FromDate = Format$(cmbDate & "/" & Format$(MonthNumber) & "/" & lblYear, "dd/mmm/yyyy")
62720   ToDate = FromDate
62730 End If

62740 If strInHouse <> "" Then
62750   lngTests = 0
62760   pb = ((StartRow - 4) * 3) + 10
62770   sql = "Select count (SampleID) as Tests from HaemResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strInHouse & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') " & _
              "and " & Parameter & " <> '' and " & Parameter & " is not null"
62780   Set tbC = New Recordset
62790   RecOpenServer 0, tbC, sql
62800   lngTests = tbC!Tests
62810   g.TextMatrix(StartRow, 1) = Format$(lngTests)
62820   g.Refresh
62830 End If

62840 If strExternal <> "" Then
62850   lngTests = 0
62860   pb = pb + 1
62870   sql = "Select count (SampleID) as Tests from HaemResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strExternal & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') " & _
              "and " & Parameter & " <> '' and " & Parameter & " is not null"
62880   Set tbC = New Recordset
62890   RecOpenServer 0, tbC, sql
62900   lngTests = tbC!Tests
62910   g.TextMatrix(StartRow, 3) = Format$(lngTests)
62920   g.Refresh
62930 End If

62940 If strNotSpecified <> "" Then
62950   lngTests = 0
62960   pb = pb + 1
62970   sql = "Select count (SampleID) as Tests from HaemResults where " & _
              "SampleID in (" & _
              "  Select SampleID from Demographics where " & _
              "  ( " & strNotSpecified & " ) " & _
              "  and RunDate between '" & FromDate & "' and '" & ToDate & "') " & _
              "and " & Parameter & " <> '' and " & Parameter & " is not null"
62980   Set tbC = New Recordset
62990   RecOpenServer 0, tbC, sql
63000   lngTests = tbC!Tests
63010   g.TextMatrix(StartRow, 5) = Format$(lngTests)
63020   g.Refresh
63030 End If

63040 Exit Sub

CalcFBC_Error:

      Dim strES As String
      Dim intEL As Integer

63050 intEL = Erl
63060 strES = Err.Description
63070 LogError "frmStatsExtInt", "CalcFBC", intEL, strES, sql


End Sub



Private Sub cmdCancel_Click()

63080 Unload Me

End Sub

Private Sub cmdSetLocations_Click()

63090 frmWardList.Show 1
63100 GenerateLists

End Sub

Private Sub cmdXL_Click()

63110 ExportFlexGrid g, Me

End Sub


Private Sub Form_Load()

      Dim n As Integer

63120 lblYear = Format$(Now, "yyyy")

63130 For n = 1 To 12
63140   cmbMonth.AddItem Format$("01/" & Format$(n) & "/2005", "mmmm")
63150 Next
63160 cmbMonth.ListIndex = Month(Now) - 1

63170 FillcmbDate

63180 GenerateLists

End Sub



Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

63190 FillcmbDate

End Sub

