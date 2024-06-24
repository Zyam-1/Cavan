VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewRep 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - View Repeats"
   ClientHeight    =   6285
   ClientLeft      =   225
   ClientTop       =   1230
   ClientWidth     =   14940
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6285
   ScaleWidth      =   14940
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4935
      Left            =   180
      TabIndex        =   8
      Top             =   1140
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   27
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmViewRep.frx":0000
   End
   Begin VB.CommandButton bswap 
      Appearance      =   0  'Flat
      Caption         =   "&Swap Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   4230
      Picture         =   "frmViewRep.frx":00CF
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bdelete 
      Appearance      =   0  'Flat
      Caption         =   "&Delete Repeat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   7050
      Picture         =   "frmViewRep.frx":1A51
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bmove 
      Appearance      =   0  'Flat
      Caption         =   "&Move Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5640
      Picture         =   "frmViewRep.frx":33D3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   11700
      Picture         =   "frmViewRep.frx":4D55
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label lRunDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2310
      TabIndex        =   9
      Top             =   60
      Width           =   1635
   End
   Begin VB.Label lname 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1050
      TabIndex        =   4
      Top             =   360
      Width           =   2925
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   390
      Width           =   495
   End
   Begin VB.Label lSampleID 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      Top             =   60
      Width           =   1245
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sample ID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   885
   End
End
Attribute VB_Name = "frmViewRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEditForm As Form

Private Sub bcancel_Click()

34590 Unload Me

End Sub

Private Sub bDelete_Click()

      Dim sql As String

34600 On Error GoTo bDelete_Click_Error

34610 g.Col = 0
34620 sql = "DELETE FROM HaemRepeats WHERE " & _
            "RunDateTime = '" & _
            Format$(g.TextMatrix(g.row, 0), "dd/MMM/yyyy HH:mm:ss") & "' " & _
            "AND SampleID = '" & lSampleID & "'"
34630 Cnxn(0).Execute sql

34640 g.Highlight = False
34650 bmove.Visible = False
34660 bdelete.Visible = False
34670 bswap.Visible = False

34680 FillG

34690 Exit Sub

bDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

34700 intEL = Erl
34710 strES = Err.Description
34720 LogError "fviewrep", "bDelete_Click", intEL, strES, sql

End Sub

Private Sub bmove_Click()

      Dim sqlFrom As String
      Dim sqlTo As String
      Dim tbFrom As Recordset
      Dim tbTo As Recordset

34730 On Error GoTo bmove_Click_Error
34740 StrEvent = "Move Haematology Result"
34750 LogEvent StrEvent, "fviewrep", "bmove_Click"


34760 sqlFrom = "Select * from HaemRepeats where " & _
                "RunDateTime = '" & Format$(g.TextMatrix(g.row, 0), "dd/MMM/yyyy HH:mm:ss") & "' " & _
                "and SampleID = '" & lSampleID & "'"
34770 sqlTo = "Select * from HaemResults where " & _
              "SampleID = '" & lSampleID & "'"
34780 Set tbFrom = New Recordset
34790 RecOpenClient 0, tbFrom, sqlFrom
34800 Set tbTo = New Recordset
34810 RecOpenClient 0, tbTo, sqlTo

34820 If (Not tbFrom.EOF) And (Not tbTo.EOF) Then
34830   tbTo!rbc = tbFrom!rbc & ""
34840   tbTo!Hgb = tbFrom!Hgb & ""
34850   tbTo!MCV = tbFrom!MCV & ""
34860   tbTo!hct = tbFrom!hct & ""
34870   tbTo!RDWCV = tbFrom!RDWCV & ""
34880   tbTo!rdwsd = tbFrom!rdwsd & ""
34890   tbTo!mch = tbFrom!mch & ""
34900   tbTo!mchc = tbFrom!mchc & ""
34910   tbTo!plt = tbFrom!plt & ""
34920   tbTo!mpv = tbFrom!mpv & ""
34930   tbTo!plcr = tbFrom!plcr & ""
34940   tbTo!pdw = tbFrom!pdw & ""
34950   tbTo!WBC = tbFrom!WBC & ""
34960   tbTo!LymA = tbFrom!LymA & ""
34970   tbTo!LymP = tbFrom!LymP & ""
34980   tbTo!MonoA = tbFrom!MonoA & ""
34990   tbTo!MonoP = tbFrom!MonoP & ""
35000   tbTo!NeutA = tbFrom!NeutA & ""
35010   tbTo!NeutP = tbFrom!NeutP & ""
35020   tbTo!EosA = tbFrom!EosA & ""
35030   tbTo!EosP = tbFrom!EosP & ""
35040   tbTo!BasA = tbFrom!BasA & ""
35050   tbTo!BasP = tbFrom!BasP & ""
35060   tbTo!RetA = tbFrom!RetA & ""
35070   tbTo!RetP = tbFrom!RetP & ""
35080   tbTo!LongError = tbFrom!LongError
35090   tbTo!Rundate = Format$(tbFrom!Rundate, "dd/mmm/yyyy")
35100   tbTo!RunDateTime = Format$(tbFrom!RunDateTime, "dd/mmm/yyyy hh:mm:ss")
35110   tbTo.Update
35120 End If
35130 With mEditForm
35140   .tWBC = g.TextMatrix(g.row, 1)
35150   .tRBC = g.TextMatrix(g.row, 2)
35160   .tHgb = g.TextMatrix(g.row, 3)
35170   .tHct = g.TextMatrix(g.row, 4)
35180   .tMCV = g.TextMatrix(g.row, 5)
35190   .tMCH = g.TextMatrix(g.row, 6)
35200   .tMCHC = g.TextMatrix(g.row, 7)
35210   .tPlt = g.TextMatrix(g.row, 8)
35220   .tLymP = g.TextMatrix(g.row, 9)
35230   .tMonoP = g.TextMatrix(g.row, 10)
35240   .tNeutP = g.TextMatrix(g.row, 11)
35250   .tEosP = g.TextMatrix(g.row, 12)
35260   .tBasP = g.TextMatrix(g.row, 13)
35270   .tLymA = g.TextMatrix(g.row, 14)
35280   .tMonoA = g.TextMatrix(g.row, 15)
35290   .tNeutA = g.TextMatrix(g.row, 16)
35300   .tEosA = g.TextMatrix(g.row, 17)
35310   .tBasA = g.TextMatrix(g.row, 18)
35320   .tRDWCV = g.TextMatrix(g.row, 19)
      '  g.Col = 20
      '  .tRDWSD = g
      '  g.Col = 21
      '  .tPdw = g
35330   .tMPV = g.TextMatrix(g.row, 22)
      '  g.Col = 23
      '  .tPLCR = g
35340   .tRetA = g.TextMatrix(g.row, 24)
35350   .tRetP = g.TextMatrix(g.row, 25)
35360 End With

35370 bDelete_Click

35380 Unload Me

35390 Exit Sub

bmove_Click_Error:

      Dim strES As String
      Dim intEL As Integer

35400 intEL = Erl
35410 strES = Err.Description
35420 LogError "fviewrep", "bmove_Click", intEL, strES

End Sub

Private Sub bswap_Click()

      Dim sql As String
      Dim tb As Recordset

      Dim MonoSpot As String
      Dim cMonospot As Integer
      Dim ESR As String
      Dim cESR As Integer
      Dim Malaria As String
      Dim cMalaria As Integer
      Dim Sickledex As String
      Dim cSickledex As Integer
      Dim RA As String
      Dim cRA As Integer


      Dim HaemRepeatsColums As String
      Dim HaemResultsColums As String
      Dim TempHaemColums As String



35430 On Error GoTo bswap_Click_Error


      'HaemRepeatsColums = "SampleID, AnalysisError, NegPosError, PosDiff, PosMorph, PosCount, err_f, err_r, ipmessage, wbc, rbc, hgb, hct, mcv, mch, mchc, plt, lymp, monop, neutP, eosp, " & _
      '                    "  basp, lyma, monoa, neuta, eosa, basa, rdwcv, rdwSD, pdw, mpv, plcr, valid, printed, retics, monospot, wbccomment, cesr, cretics, cmonospot, ccoag, md0, md1, md2, " & _
      '                    "  md3, md4, md5, RunDate, RunDateTime, ESR, PT, PTControl, APTT, APTTControl, INR, FDP, FIB, Operator, FAXed, Warfarin, DDimers, TransmitTime, Pct, WIC, WOC, " & _
      '                    "  gWB1, gWB2, gRBC, gPlt, gWIC, LongError, cFilm, RetA, RetP, nrbcA, nrbcP, cMalaria, Malaria, cSickledex, Sickledex, RA, cRA, Val1, Val2, Val3, Val4, Val5, gRBCH, " & _
      '                    "  gPLTH, gPLTF, gV, gC, gS, DF1, IRF, NOPAS, Image, mi, an, ca, va, ho, he, ls, at, bl, pp, nl, mn, wp, ch, wb, hdw, LUCP, LUCA, LI, MPXI, ANALYSER, cAsot, tAsot, tRa, " & _
      '                    "  hyp , rbcf, rbcg, mpo, ig, lplt, pclm, ValidateTime, Healthlink, CD3A, CD4A, CD8A, CD3P, CD4P, CD8P, CD48, WVF, AnalyserMessage " & _
      '                    "  , SignOff, SignOffBy, SignOffDateTime "
      'HaemResultsColums = "sampleid, analysiserror, negposerror, posdiff, posmorph, poscount, err_f, err_r, ipmessage, wbc, rbc, hgb, hct, mcv, mch, mchc, plt, lymp, monop, neutP, eosp, basp, " & _
      '                    "  lyma, monoa, neuta, eosa, basa, rdwcv, rdwsd, pdw, mpv, plcr, valid, printed, retics, monospot, wbccomment, cesr, cretics, cmonospot, ccoag, md0, md1, md2, md3," & _
      '                    "  md4, md5, RunDate, RunDateTime, ESR, PT, PTControl, APTT, APTTControl, INR, FDP, FIB, Operator, FAXed, Warfarin, DDimers, TransmitTime, Pct, WIC, WOC, gWB1," & _
      '                    "  gWB2, gRBC, gPlt, gWIC, LongError, cFilm, RetA, RetP, nrbcA, nrbcP, cMalaria, Malaria, cSickledex, Sickledex, RA, cRA, Val1, Val2, Val3, Val4, Val5, gRBCH, gPLTH," & _
      '                    "  gPLTF, gV, gC, gS, DF1, IRF, NOPAS, Image, mi, an, ca, va, ho, he, ls, at, bl, pp, nl, mn, wp, ch, wb, hdw, LUCP, LUCA, LI, MPXI, ANALYSER, cAsot, tAsot, tRa, hyp, rbcf," & _
      '                    "  rbcg , mpo, ig, lplt, pclm, ValidateTime, Healthlink, CD3A, CD4A, CD8A, CD3P, CD4P, CD8P, CD48, WVF, AnalyserMessage " & _
      '                    "  , SignOff, SignOffBy, SignOffDateTime "
      '
      'TempHaemColums = "sampleid, analysiserror, negposerror, posdiff, posmorph, poscount, err_f, err_r, ipmessage, wbc, rbc, hgb, hct, mcv, mch, mchc, plt, lymp, monop, neutP, eosp, basp, " & _
      '                      " lyma, monoa, neuta, eosa, basa, rdwcv, rdwsd, pdw, mpv, plcr, valid, printed, retics, monospot, wbccomment, cesr, cretics, cmonospot, ccoag, md0, md1, md2, md3, " & _
      '                      " md4, md5, RunDate, RunDateTime, ESR, PT, PTControl, APTT, APTTControl, INR, FDP, FIB, Operator, FAXed, Warfarin, DDimers, TransmitTime, Pct, WIC, WOC, gWB1, " & _
      '                      " gWB2, gRBC, gPlt, gWIC, LongError, cFilm, RetA, RetP, nrbcA, nrbcP, cMalaria, Malaria, cSickledex, Sickledex, RA, cRA, Val1, Val2, Val3, Val4, Val5, gRBCH, gPLTH, " & _
      '                      " gPLTF, gV, gC, gS, DF1, IRF, NOPAS, Image, mi, an, ca, va, ho, he, ls, at, bl, pp, nl, mn, wp, ch, wb, hdw, LUCP, LUCA, LI, MPXI, ANALYSER, cAsot, tAsot, tRa, hyp, rbcf, " & _
      '                      " rbcg , mpo, ig, lplt, pclm, ValidateTime, Healthlink, CD3A, CD4A, CD8A, CD3P, CD4P, CD8P, CD48, WVF, AnalyserMessage " & _
      '                      "  , SignOff, SignOffBy, SignOffDateTime "


35440 g.Col = 0

35450 sql = "SELECT MonoSpot, COALESCE(cMonoSpot, 0) cMonoSpot, " & _
            "ESR, COALESCE(cESR, 0) cESR, " & _
            "Malaria, COALESCE(cMalaria, 0) cMalaria, " & _
            "Sickledex, COALESCE(cSickledex, 0) cSickledex, " & _
            "RA, COALESCE(cRA, 0) cRA " & _
            "FROM HaemResults WHERE " & _
            "SampleID = '" & lSampleID & "'"
35460 Set tb = New Recordset
35470 RecOpenServer 0, tb, sql
35480 If Not tb.EOF Then
35490   MonoSpot = tb!MonoSpot & ""
35500   cMonospot = tb!cMonospot
35510   ESR = tb!ESR & ""
35520   cESR = tb!cESR
35530   Malaria = tb!Malaria & ""
35540   cMalaria = tb!cMalaria
35550   Sickledex = tb!Sickledex & ""
35560   cSickledex = tb!cSickledex
35570   RA = tb!RA & ""
35580   cRA = tb!cRA
35590 End If

35600 sql = "TRUNCATE TABLE TempHaem " & _
            "INSERT INTO TempHaem " & _
            "    SELECT TOP 1 * FROM HaemRepeats WHERE " & _
            "    RunDateTime = '" & Format$(g, "dd/MMM/yyyy HH:mm:ss") & "' " & _
            "    AND SampleID = '" & lSampleID & "' " & _
            "DELETE FROM HaemRepeats WHERE " & _
            "RunDateTime = '" & Format$(g, "dd/MMM/yyyy HH:mm:ss") & "' " & _
            "AND SampleID = '" & lSampleID & "' " & _
            "INSERT INTO HaemRepeats " & _
            "    SELECT * FROM HaemResults WHERE " & _
            "    SampleID = '" & lSampleID & "' " & _
            "DELETE FROM HaemResults WHERE " & _
            "SampleID = '" & lSampleID & "' " & _
            "INSERT INTO HaemResults " & _
            "    SELECT * FROM TempHaem"
35610 Cnxn(0).Execute sql


      'sql = "TRUNCATE TABLE TempHaem " & _
      '      "INSERT INTO TempHaem " & _
      '      "    SELECT TOP 1 " & _
      '      HaemRepeatsColums & " FROM HaemRepeats WHERE " & _
      '      "    RunDateTime = '" & Format$(g, "dd/MMM/yyyy HH:mm:ss") & "' " & _
      '      "    AND SampleID = '" & lSampleID & "' " & _
      '      "DELETE FROM HaemRepeats WHERE " & _
      '      "RunDateTime = '" & Format$(g, "dd/MMM/yyyy HH:mm:ss") & "' " & _
      '      "AND SampleID = '" & lSampleID & "' " & _
      '      "INSERT INTO HaemRepeats " & _
      '      "    SELECT " & _
      '      HaemResultsColums & " FROM HaemResults WHERE " & _
      '      "    SampleID = '" & lSampleID & "' " & _
      '      "DELETE FROM HaemResults WHERE " & _
      '      "SampleID = '" & lSampleID & "' " & _
      '      "INSERT INTO HaemResults " & _
      '      "    SELECT " & TempHaemColums & " FROM TempHaem"
      'Cnxn(0).Execute Sql




35620 sql = "UPDATE HaemResults " & _
            "SET MonoSpot = '" & MonoSpot & "', " & _
            "cMonospot = '" & cMonospot & "', " & _
            "ESR = '" & ESR & "', " & _
            "cESR = '" & cESR & "', " & _
            "Malaria = '" & Malaria & "', " & _
            "cMalaria = '" & cMalaria & "', " & _
            "Sickledex = '" & Sickledex & "', " & _
            "cSickledex = '" & cSickledex & "', " & _
            "RA = '" & RA & "', " & _
            "cRA = '" & cRA & "' " & _
            "WHERE SampleID = '" & lSampleID & "'"
35630 Cnxn(0).Execute sql

35640 Unload Me

35650 Exit Sub

bswap_Click_Error:

      Dim strES As String
      Dim intEL As Integer

35660 intEL = Erl
35670 strES = Err.Description
35680 LogError "fviewrep", "bswap_Click", intEL, strES

End Sub
Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

35690 On Error GoTo FillG_Error

35700 sql = "select * from haemrepeats where " & _
            "sampleid = '" & lSampleID & "' " & _
            "order by rundatetime"
35710 Set tb = New Recordset
35720 RecOpenClient 0, tb, sql

35730 g.Rows = 2
35740 g.AddItem ""
35750 g.RemoveItem 1

35760 Do While Not tb.EOF
35770   s = Format$(tb!RunDateTime, "dd/MM/yyyy HH:mm:ss") & vbTab
35780   s = s & tb!WBC & vbTab
35790   s = s & tb!rbc & vbTab
35800   s = s & tb!Hgb & vbTab
35810   s = s & tb!hct & vbTab
35820   s = s & tb!MCV & vbTab
35830   s = s & tb!mch & vbTab
35840   s = s & tb!mchc & vbTab
35850   s = s & tb!plt & vbTab
35860   s = s & tb!LymP & vbTab
35870   s = s & tb!MonoP & vbTab
35880   s = s & tb!NeutP & vbTab
35890   s = s & tb!EosP & vbTab
35900   s = s & tb!BasP & vbTab
35910   s = s & tb!LymA & vbTab
35920   s = s & tb!MonoA & vbTab
35930   s = s & tb!NeutA & vbTab
35940   s = s & tb!EosA & vbTab
35950   s = s & tb!BasA & vbTab
35960   s = s & tb!RDWCV & vbTab
35970   s = s & tb!rdwsd & vbTab
35980   s = s & tb!pdw & vbTab
35990   s = s & tb!mpv & vbTab
36000   s = s & tb!plcr & vbTab
36010   s = s & tb!RetA & vbTab
36020   s = s & tb!RetP & vbTab
36030   g.AddItem s
36040   tb.MoveNext
36050 Loop

36060 sql = "select * from haemresults where " & _
            "sampleid = '" & lSampleID & "'"
36070 Set tb = New Recordset
36080 RecOpenClient 0, tb, sql

36090 If Not tb.EOF Then
36100   lRunDate = Format$(tb!RunDateTime, "dd/mm/yyyy")
36110   s = Format$(tb!RunDateTime, "hh:mm:ss") & vbTab
36120   s = s & tb!WBC & vbTab
36130   s = s & tb!rbc & vbTab
36140   s = s & tb!Hgb & vbTab
36150   s = s & tb!hct & vbTab
36160   s = s & tb!MCV & vbTab
36170   s = s & tb!mch & vbTab
36180   s = s & tb!mchc & vbTab
36190   s = s & tb!plt & vbTab
36200   s = s & tb!LymP & vbTab
36210   s = s & tb!MonoP & vbTab
36220   s = s & tb!NeutP & vbTab
36230   s = s & tb!EosP & vbTab
36240   s = s & tb!BasP & vbTab
36250   s = s & tb!LymA & vbTab
36260   s = s & tb!MonoA & vbTab
36270   s = s & tb!NeutA & vbTab
36280   s = s & tb!EosA & vbTab
36290   s = s & tb!BasA & vbTab
36300   s = s & tb!RDWCV & vbTab
36310   s = s & tb!rdwsd & vbTab
36320   s = s & tb!pdw & vbTab
36330   s = s & tb!mpv & vbTab
36340   s = s & tb!plcr & vbTab
36350   s = s & tb!RetA & vbTab
36360   s = s & tb!RetP & vbTab

36370   g.AddItem s, 1
36380 End If

36390 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

36400 intEL = Erl
36410 strES = Err.Description
36420 LogError "fviewrep", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Activate()

36430 FillG

End Sub

Private Sub g_Click()

36440 If g.row < 3 Then
36450   g.Highlight = flexHighlightNever
36460   bmove.Visible = False
36470   bdelete.Visible = False
36480   bswap.Visible = False
36490 Else
36500   g.Highlight = flexHighlightAlways
36510   bmove.Visible = True
36520   bdelete.Visible = True
36530   bswap.Visible = True
        
36540   g.Col = 1
36550   g.ColSel = g.Cols - 1
36560   g.RowSel = g.row
        
36570   bswap.SetFocus
36580 End If

End Sub


Public Property Let EditForm(ByVal EditForm As Form)

36590 Set mEditForm = EditForm

End Property
