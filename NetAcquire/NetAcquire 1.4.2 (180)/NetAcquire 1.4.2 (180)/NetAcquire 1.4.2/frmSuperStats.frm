VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmSuperStats 
   Caption         =   "NetAcquire - Statistics "
   ClientHeight    =   8460
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   14055
   Icon            =   "frmSuperStats.frx":0000
   LinkTopic       =   "frmBigStats"
   ScaleHeight     =   8460
   ScaleWidth      =   14055
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1065
      Left            =   150
      TabIndex        =   21
      Top             =   90
      Width           =   1545
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   22
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Format          =   220200961
         CurrentDate     =   38631
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   90
         TabIndex        =   23
         Top             =   630
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Format          =   220200961
         CurrentDate     =   38631
      End
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   945
      Left            =   3570
      Picture         =   "frmSuperStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   180
      Width           =   855
   End
   Begin Threed.SSCommand cmdStart 
      Height          =   945
      Left            =   1830
      TabIndex        =   15
      Top             =   180
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1667
      _StockProps     =   78
      Caption         =   "&Start"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1020
      Left            =   5940
      TabIndex        =   0
      Top             =   90
      Width           =   4245
      _Version        =   65536
      _ExtentX        =   7488
      _ExtentY        =   1799
      _StockProps     =   15
      Caption         =   "Discipline"
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton optDisp 
         Caption         =   "External"
         Height          =   195
         Index           =   6
         Left            =   3030
         TabIndex        =   24
         Tag             =   "Ext"
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Blood Gas"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   14
         Tag             =   "BGA"
         Top             =   750
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   13
         Tag             =   "Imm"
         Top             =   510
         Width           =   1305
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   12
         Tag             =   "End"
         Top             =   270
         Width           =   1305
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Haematology"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   11
         Tag             =   "Haem"
         Top             =   750
         Width           =   1245
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Coagulation"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Tag             =   "Coag"
         Top             =   510
         Width           =   1185
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Tag             =   "Bio"
         Top             =   270
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1020
      Left            =   11190
      TabIndex        =   1
      Top             =   90
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   1799
      _StockProps     =   15
      Caption         =   "Criteria"
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton optDoc 
         Caption         =   "Ward"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   4
         Top             =   750
         Width           =   690
      End
      Begin VB.OptionButton optDoc 
         Caption         =   "Clinician"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   3
         Top             =   510
         Width           =   915
      End
      Begin VB.OptionButton optDoc 
         Caption         =   "Gp"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1020
      Left            =   12390
      TabIndex        =   5
      Top             =   90
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   1799
      _StockProps     =   15
      Caption         =   "Hospital"
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   750
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   510
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   945
      Left            =   2670
      TabIndex        =   16
      Top             =   180
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1667
      _StockProps     =   78
      Caption         =   "E&xit"
   End
   Begin MSFlexGridLib.MSFlexGrid grdStats 
      Height          =   7035
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   12409
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
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
   Begin VB.Label lblUpTo 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   435
      Left            =   2880
      TabIndex        =   25
      Top             =   4020
      Width           =   6735
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
      Left            =   4560
      TabIndex        =   20
      Top             =   180
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label3 
      Caption         =   "The Report is being Generated.              Please Wait."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   2835
      TabIndex        =   17
      Top             =   2700
      Width           =   6765
   End
End
Attribute VB_Name = "frmSuperStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Choice As String
Private Dept As String

Private Hosp As String

Private Sub CalcHaemByMonth()
          
      Dim sql As String
      Dim tb As Recordset
      Dim Sel As String
      Dim Y As Integer
      Dim StartDate As String
      Dim EndDate As String
      Dim InterDate As String
      Dim Finished As Boolean

2500  On Error GoTo CalcHaemByMonth_Error

2510  StartDate = Format(dtFrom, "dd/MMM/yyyy")
2520  EndDate = Format(dtTo, "dd/MMM/yyyy")
2530  InterDate = StartDate

2540  Finished = False
2550  grdStats.Visible = True
2560  grdStats.Refresh
2570  Screen.MousePointer = vbHourglass
2580  lblUpTo.Visible = True
2590  lblUpTo.Refresh

2600  Do While Not Finished
        
2610    InterDate = Format(DateAdd("m", 1, InterDate), "dd/MMM/yyyy")
2620    If DateDiff("d", InterDate, EndDate) < 0 Then
2630      InterDate = Format(EndDate, "dd/MMM/yyyy")
2640      Finished = True
2650    End If

2660    lblUpTo = "Up to " & Format(InterDate, "dd/MM/yy")
2670    lblUpTo.Refresh

2680    Sel = "xxx"

2690    sql = "IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[CustomTable]') " & _
              "           AND OBJECTPROPERTY(id, N'IsUserTable') = 1) " & _
              "  DROP TABLE [CustomTable] " & _
              "SELECT DISTINCT D." & Choice & " Choice, " & _
              "CASE WBC WHEN '' THEN 0 ELSE COUNT (WBC) END WBC, " & _
              "CASE RetA WHEN '' THEN 0 ELSE COUNT (RetA) END RetA, " & _
              "CASE ESR WHEN '' THEN 0 ELSE COUNT (ESR) END ESR, " & _
              "CASE tRA WHEN '' THEN 0 ELSE COUNT (tRA) END tRA, " & _
              "CASE Malaria WHEN '' THEN 0 ELSE COUNT (Malaria) END Malaria, " & _
              "CASE Monospot WHEN '' THEN 0 ELSE COUNT (Monospot) END Monospot, " & _
              "CASE Sickledex WHEN '' THEN 0 ELSE COUNT (Sickledex) END Sickledex, " & _
              "CASE tASOT WHEN '' THEN 0 ELSE COUNT (tASOT) END tASOT " & _
              "INTO CustomTable " & _
              "FROM HaemResults R, Demographics D WHERE " & _
              "D.SampleID = R.SampleID " & _
              "AND D.Rundate BETWEEN '" & StartDate & "' AND '" & InterDate & "' " & _
              "AND (COALESCE(D.Hospital,'" & Hosp & "') = '" & Hosp & "' OR D.Hospital = '' ) " & _
              "GROUP BY D." & Choice & ", WBC, RetA, ESR, tRA, Malaria, Monospot, Sickledex, tASOT " & _
              "ORDER BY " & Choice & " "
2700    Cnxn(0).Execute sql
2710    sql = "SELECT Choice, " & _
              "SUM(WBC) WBC, " & _
              "SUM(RetA) RetA, " & _
              "SUM(ESR) ESR, " & _
              "sum(tRA) tRA, " & _
              "SUM(Malaria) Malaria, " & _
              "SUM(Monospot) Monospot, " & _
              "SUM(Sickledex) Sickledex, " & _
              "SUM(tASOT) tASOT " & _
              "FROM CustomTable " & _
              "GROUP BY Choice "
        
2720    Set tb = New Recordset
2730    RecOpenServer 0, tb, sql
          
2740    With grdStats
2750      Do While Not tb.EOF
2760        If UCase(Sel) <> UCase(Trim$(tb!Choice & "")) Then
2770          Sel = UCase(Trim$(tb!Choice & ""))
2780          For Y = 1 To grdStats.Rows - 1
2790            If UCase(Trim$(grdStats.TextMatrix(Y, 0))) = UCase(Trim$(tb!Choice & "")) Then
2800              Exit For
2810            End If
2820          Next
2830        End If
2840        If Y <> .Rows And Y <> 0 Then
2850          .TextMatrix(Y, 1) = Val(.TextMatrix(Y, 1)) + tb!WBC
2860          .TextMatrix(Y, 2) = Val(.TextMatrix(Y, 2)) + tb!RetA
2870          .TextMatrix(Y, 3) = Val(.TextMatrix(Y, 3)) + tb!ESR
2880          .TextMatrix(Y, 4) = Val(.TextMatrix(Y, 4)) + tb!tRa
2890          .TextMatrix(Y, 5) = Val(.TextMatrix(Y, 5)) + tb!Malaria
2900          .TextMatrix(Y, 6) = Val(.TextMatrix(Y, 6)) + tb!MonoSpot
2910          .TextMatrix(Y, 7) = Val(.TextMatrix(Y, 7)) + tb!Sickledex
2920          .TextMatrix(Y, 8) = Val(.TextMatrix(Y, 8)) + tb!tASOt
2930        End If

2940        tb.MoveNext
2950      Loop
2960    End With
2970    grdStats.Refresh
2980    StartDate = Format$(DateAdd("d", 1, InterDate), "dd/MMM/yyyy")
2990  Loop

3000  lblUpTo.Visible = False
3010  Screen.MousePointer = vbNormal

3020  Exit Sub

CalcHaemByMonth_Error:

      Dim strES As String
      Dim intEL As Integer

3030  intEL = Erl
3040  strES = Err.Description
3050  LogError "frmSuperStats", "CalcHaemByMonth", intEL, strES, sql


End Sub

Private Sub FillColsHaem()
           
3060  grdStats.Cols = 9
3070  grdStats.TextMatrix(0, 1) = "WBC"
3080  grdStats.TextMatrix(0, 2) = "Retics"
3090  grdStats.TextMatrix(0, 3) = "ESR"
3100  grdStats.TextMatrix(0, 4) = "RA"
3110  grdStats.TextMatrix(0, 5) = "Malaria"
3120  grdStats.TextMatrix(0, 6) = "Monospot"
3130  grdStats.TextMatrix(0, 7) = "Sickledex"
3140  grdStats.TextMatrix(0, 8) = "ASOT"

End Sub

Private Sub FillColsExt()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
          
3150  On Error GoTo FillColsExt_Error

3160  sql = "SELECT DISTINCT T.AnalyteName " & _
            "FROM Demographics D, ExtResults R, ExternalDefinitions T " & _
            "WHERE D.SampleID = R.SampleID " & _
            "AND R.Analyte = T.AnalyteName " & _
            "AND (D.RunDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' AND '" & Format(dtTo, "dd/MMM/yyyy") & "') " & _
            "AND D." & Choice & " <> '' " & _
            "ORDER BY T.AnalyteName"

3170  Set tb = New Recordset
3180  RecOpenServer 0, tb, sql
3190  Do While Not tb.EOF
3200    n = grdStats.Cols - 1
3210    grdStats.TextMatrix(0, n) = tb!AnalyteName & ""
3220    grdStats.Cols = grdStats.Cols + 1
3230    tb.MoveNext
3240  Loop
3250  If grdStats.Cols > 2 Then grdStats.Cols = grdStats.Cols - 1

3260  Exit Sub

FillColsExt_Error:

      Dim strES As String
      Dim intEL As Integer

3270  intEL = Erl
3280  strES = Err.Description
3290  LogError "frmSuperStats", "FillColsExt", intEL, strES, sql

End Sub


Private Sub cmdExit_Click()

3300  Unload Me

End Sub

Private Sub cmdStart_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
      Dim z As Long
      Dim TestTot As Long
      Dim s As String
      Dim FromDate As String
      Dim ToDate As String
      Dim ChoiceName As String
      Dim Sel As String

3310  On Error GoTo cmdStart_Click_Error

3320  grdStats.Visible = False

3330  lblInfo = ""
3340  lblInfo.Refresh
3350  Me.Refresh

3360  FromDate = Format(dtFrom, "dd/MMM/yyyy")
3370  ToDate = Format(dtTo, "dd/MMM/yyyy")

3380  For n = 0 To 6
3390    If optDisp(n) Then
3400      Dept = optDisp(n).Tag
3410    End If
3420  Next

3430  FillRows
3440  FillCols

3450  For z = 1 To grdStats.Cols - 1
3460    grdStats.ColWidth(z) = 700
3470  Next

3480  grdStats.TextMatrix(0, 0) = Choice & " Name"

3490  Sel = "xxx"

3500  Select Case Dept
        Case "Haem"
3510      CalcHaemByMonth
      '    sql = "IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[CustomTable]') " & _
      '          "           AND OBJECTPROPERTY(id, N'IsUserTable') = 1) " & _
      '          "  DROP TABLE [CustomTable] " & _
      '          "SELECT DISTINCT D." & Choice & " Choice, " & _
      '          "CASE WBC WHEN '' THEN 0 ELSE COUNT (WBC) END WBC, " & _
      '          "CASE RetA WHEN '' THEN 0 ELSE COUNT (RetA) END RetA, " & _
      '          "CASE ESR WHEN '' THEN 0 ELSE COUNT (ESR) END ESR, " & _
      '          "CASE tRA WHEN '' THEN 0 ELSE COUNT (tRA) END tRA, " & _
      '          "CASE Malaria WHEN '' THEN 0 ELSE COUNT (Malaria) END Malaria, " & _
      '          "CASE Monospot WHEN '' THEN 0 ELSE COUNT (Monospot) END Monospot, " & _
      '          "CASE Sickledex WHEN '' THEN 0 ELSE COUNT (Sickledex) END Sickledex, " & _
      '          "CASE tASOT WHEN '' THEN 0 ELSE COUNT (tASOT) END tASOT " & _
      '          "INTO CustomTable " & _
      '          "FROM HaemResults R, Demographics D WHERE " & _
      '          "D.SampleID = R.SampleID " & _
      '          "AND D.Rundate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' AND '" & Format(dtTo, "dd/MMM/yyyy") & "' " & _
      '          "AND (COALESCE(D.Hospital,'" & Hosp & "') = '" & Hosp & "' OR D.Hospital = '' ) " & _
      '          "GROUP BY D." & Choice & ", WBC, RetA, ESR, tRA, Malaria, Monospot, Sickledex, tASOT " & _
      '          "ORDER BY " & Choice & " "
      '    Cnxn(0).Execute Sql
      '    sql = "SELECT Choice, " & _
      '          "SUM(WBC) WBC, " & _
      '          "SUM(RetA) RetA, " & _
      '          "SUM(ESR) ESR, " & _
      '          "sum(tRA) tRA, " & _
      '          "SUM(Malaria) Malaria, " & _
      '          "SUM(Monospot) Monospot, " & _
      '          "SUM(Sickledex) Sickledex, " & _
      '          "SUM(tASOT) tASOT " & _
      '          "FROM CustomTable " & _
      '          "GROUP BY Choice "
      '
      '    Set tb = New Recordset
      '    RecOpenServer 0, tb, sql
      '    Do While Not tb.EOF
      '      If UCase(Sel) <> UCase(Trim$(tb!Choice & "")) Then
      '        Sel = UCase(Trim$(tb!Choice & ""))
      '        For y = 1 To grdStats.Rows - 1
      '          If UCase(Trim$(grdStats.TextMatrix(y, 0))) = UCase(Trim$(tb!Choice & "")) Then
      '            Exit For
      '          End If
      '        Next
      '      End If
      '      If y <> grdStats.Rows And y <> 0 Then
      '        grdStats.TextMatrix(y, 1) = tb!WBC
      '        grdStats.TextMatrix(y, 2) = tb!reta
      '        grdStats.TextMatrix(y, 3) = tb!esr
      '        grdStats.TextMatrix(y, 4) = tb!tRa
      '        grdStats.TextMatrix(y, 5) = tb!Malaria
      '        grdStats.TextMatrix(y, 6) = tb!Monospot
      '        grdStats.TextMatrix(y, 7) = tb!Sickledex
      '        grdStats.TextMatrix(y, 8) = tb!tASOt
      '      End If
      '      tb.MoveNext
      '    Loop
        
3520    Case "Bio", "Imm", "End", "Coag"
3530      For n = 1 To grdStats.Rows - 1
3540        ChoiceName = AddTicks(grdStats.TextMatrix(n, 0))
3550        sql = "SELECT COUNT(" & Dept & "Results.Code) AS Tot, Code FROM Demographics, " & Dept & "Results WHERE " & _
                  Choice & " = '" & ChoiceName & "' " & _
                  "AND Demographics.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                  "AND Demographics.SampleID = " & Dept & "Results.SampleID " & _
                  "GROUP BY Code"
3560        Set tb = New Recordset
3570        RecOpenServer 0, tb, sql
3580        Do While Not tb.EOF
3590          For z = 1 To grdStats.Cols - 1
3600            If grdStats.TextMatrix(0, z) = tb!Code Then
3610              grdStats.TextMatrix(n, z) = tb!Tot
3620              lblInfo = grdStats.TextMatrix(n, 0) & " " & tb!Tot
3630              lblInfo.Refresh
3640              Exit For
3650            End If
3660          Next
3670          tb.MoveNext
3680        Loop
3690      Next
          
3700      For n = 1 To grdStats.Cols - 1
3710        grdStats.TextMatrix(0, n) = GetSCode(grdStats.TextMatrix(0, n), Dept)
3720      Next
        
3730    Case "Ext"
3740      For n = 1 To grdStats.Rows - 1
3750        ChoiceName = AddTicks(grdStats.TextMatrix(n, 0))
3760        For z = 1 To grdStats.Cols - 1
          
3770          sql = "SELECT COUNT(ExtResults.Analyte) As Tot " & _
                    "FROM Demographics, ExtResults WHERE " & _
                    Choice & " = '" & ChoiceName & "' " & _
                    "AND Demographics.RunDate BETWEEN '" & FromDate & "' and '" & ToDate & "' " & _
                    "AND Demographics.SampleID = ExtResults.SampleID " & _
                    "AND ExtResults.Analyte = '" & grdStats.TextMatrix(0, z) & "'"
3780          Set tb = New Recordset
3790          RecOpenServer 0, tb, sql
3800          If tb!Tot <> 0 Then
3810            grdStats.TextMatrix(n, z) = tb!Tot
3820            lblInfo = grdStats.TextMatrix(n, 0) & " " & tb!Tot
3830            lblInfo.Refresh
3840          End If
3850        Next
3860      Next
        
3870  End Select

3880  grdStats.AddItem ""
3890  s = "TOTAL" & vbTab
3900  For n = 1 To grdStats.Cols - 1
3910    For z = 1 To grdStats.Rows - 1
3920      TestTot = TestTot + Val(grdStats.TextMatrix(z, n))
3930    Next
3940    s = s & TestTot & vbTab
3950    TestTot = 0
3960  Next
3970  grdStats.AddItem s

3980  grdStats.Visible = True

3990  Exit Sub

cmdStart_Click_Error:

      Dim strES As String
      Dim intEL As Integer

4000  intEL = Erl
4010  strES = Err.Description
4020  LogError "frmSuperStats", "cmdStart_Click", intEL, strES, sql

End Sub

Private Function GetSCode(ByVal Code As String, ByVal Dept As String) As String
      Dim tb As Recordset
      Dim sql As String

4030  On Error GoTo GetSCode_Error

4040  If Dept = "Coag" Then
4050    sql = "Select TestName as tName from CoagTestDefinitions where code = '" & Code & "' and inuse = 1"
4060  Else
4070    sql = "Select ShortName as tName from " & Dept & "testdefinitions where code = '" & Code & "' and inuse = 1"
4080  End If

4090  Set tb = New Recordset
4100  RecOpenServer 0, tb, sql
4110  If Not tb.EOF Then
4120    GetSCode = tb!tName
4130  End If

4140  Exit Function

GetSCode_Error:

      Dim strES As String
      Dim intEL As Integer

4150  intEL = Erl
4160  strES = Err.Description
4170  LogError "frmSuperStats", "GetSCode", intEL, strES, sql


End Function



Private Sub cmdXL_Click()

4180  ExportFlexGrid grdStats, Me

End Sub

Private Sub Form_Load()
      Dim n As Long

4190  For n = 0 To intOtherHospitalsInGroup
4200    optHosp(n).Visible = True
4210    optHosp(n).Caption = Initial2Upper(HospName(n))
4220  Next

4230  dtFrom = Format(Now, "dd/MMM/yyyy")
4240  dtTo = Format(Now, "dd/MMM/yyyy")



End Sub


Private Sub FillRows()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
      Dim FromDate As String
      Dim ToDate As String

4250  On Error GoTo FillRows_Error

4260  grdStats.ColWidth(0) = 2000

4270  With grdStats
4280    .Rows = 2
4290    .Cols = 2
4300    .TextMatrix(0, 1) = ""
4310    .AddItem ""
4320    .RemoveItem 1
4330  End With

4340  For n = 0 To 2
4350    If optHosp(n) Then
4360      Hosp = optHosp(n).Caption
4370    End If
4380  Next

4390  For n = 0 To 2
4400    If optDoc(n) Then
4410      Choice = optDoc(n).Caption
4420    End If
4430  Next

4440  FromDate = Format(dtFrom, "dd/MMM/yyyy")
4450  ToDate = Format(dtTo, "dd/MMM/yyyy")

4460  sql = "SELECT DISTINCT(" & Choice & ") FROM Demographics WHERE " & _
            Choice & " <> '' " & _
            "AND SampleID IN (SELECT DISTINCT(SampleID) FROM " & Dept & "Results WHERE " & _
            "                 RunDate BETWEEN '" & FromDate & "' and '" & ToDate & "') " & _
            "ORDER BY " & Choice
4470  Set tb = New Recordset
4480  RecOpenServer 0, tb, sql
4490  Do While Not tb.EOF
4500    If Trim(tb(Choice) & "") <> "" Then grdStats.AddItem tb(Choice)
4510    tb.MoveNext
4520  Loop

4530  If grdStats.Rows > 2 And grdStats.TextMatrix(1, 0) = "" Then
4540    grdStats.RemoveItem 1
4550  End If

4560  Exit Sub

FillRows_Error:

      Dim strES As String
      Dim intEL As Integer

4570  intEL = Erl
4580  strES = Err.Description
4590  LogError "frmSuperStats", "FillRows", intEL, strES, sql

End Sub


Private Sub FillCols()

4600  Select Case Dept
        
        Case "Bio": FillColsBio
4610    Case "Imm": FillColsImm
4620    Case "End": FillColsEnd
4630    Case "Coag": FillColsCoag
4640    Case "Haem": FillColsHaem
4650    Case "Ext": FillColsExt
          
4660  End Select

End Sub

Private Sub FillColsCoag()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
          
4670  On Error GoTo FillColsCoag_Error

4680  sql = "SELECT DISTINCT T.Code, T.TestName " & _
            "FROM Demographics D, CoagResults R, CoagTestDefinitions T " & _
            "WHERE D.SampleID = R.SampleID " & _
            "AND R.Code = T.Code " & _
            "AND T.InUse = 1 " & _
            "AND (D.RunDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' AND '" & Format(dtTo, "dd/MMM/yyyy") & "') " & _
            "" & _
            "ORDER BY T.TestName"

4690  Set tb = New Recordset
4700  RecOpenServer 0, tb, sql
4710  Do While Not tb.EOF
4720    n = grdStats.Cols - 1
4730    grdStats.TextMatrix(0, n) = tb!Code
4740    grdStats.Cols = grdStats.Cols + 1
4750    tb.MoveNext
4760  Loop
4770  If grdStats.Cols > 2 Then grdStats.Cols = grdStats.Cols - 1

4780  Exit Sub

FillColsCoag_Error:

      Dim strES As String
      Dim intEL As Integer

4790  intEL = Erl
4800  strES = Err.Description
4810  LogError "frmSuperStats", "FillColsCoag", intEL, strES, sql

End Sub

Private Sub FillColsImm()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
          
4820  On Error GoTo FillColsImm_Error

4830  sql = "SELECT DISTINCT T.Code, T.shortname " & _
            "FROM Demographics D, ImmResults R, ImmTestDefinitions T " & _
            "WHERE D.SampleID = R.SampleID " & _
            "AND R.Code = T.Code " & _
            "AND T.InUse = 1 " & _
            "AND (D.RunDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' AND '" & Format(dtTo, "dd/MMM/yyyy") & "') " & _
            "AND D." & Choice & " <> '' " & _
            "ORDER BY T.shortname"

4840  Set tb = New Recordset
4850  RecOpenServer 0, tb, sql
4860  Do While Not tb.EOF
4870    n = grdStats.Cols - 1
4880    grdStats.TextMatrix(0, n) = tb!Code
4890    grdStats.Cols = grdStats.Cols + 1
4900    tb.MoveNext
4910  Loop
4920  If grdStats.Cols > 2 Then grdStats.Cols = grdStats.Cols - 1

4930  Exit Sub

FillColsImm_Error:

      Dim strES As String
      Dim intEL As Integer

4940  intEL = Erl
4950  strES = Err.Description
4960  LogError "frmSuperStats", "FillColsImm", intEL, strES, sql

End Sub

Private Sub FillColsBio()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
          
4970  On Error GoTo FillColsBio_Error

4980  sql = "SELECT DISTINCT T.Code, T.shortname " & _
            "FROM Demographics D, BioResults R, BioTestDefinitions T " & _
            "WHERE D.SampleID = R.SampleID " & _
            "AND R.Code = T.Code " & _
            "AND T.InUse = 1 " & _
            "AND (D.RunDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' AND '" & Format(dtTo, "dd/MMM/yyyy") & "') " & _
            "AND D." & Choice & " <> '' " & _
            "ORDER BY T.shortname"

4990  Set tb = New Recordset
5000  RecOpenServer 0, tb, sql
5010  Do While Not tb.EOF
5020    n = grdStats.Cols - 1
5030    grdStats.TextMatrix(0, n) = tb!Code
5040    grdStats.Cols = grdStats.Cols + 1
5050    tb.MoveNext
5060  Loop
5070  If grdStats.Cols > 2 Then grdStats.Cols = grdStats.Cols - 1

5080  Exit Sub

FillColsBio_Error:

      Dim strES As String
      Dim intEL As Integer

5090  intEL = Erl
5100  strES = Err.Description
5110  LogError "frmSuperStats", "FillColsBio", intEL, strES, sql

End Sub

Private Sub FillColsEnd()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
          
5120  On Error GoTo FillColsEnd_Error

5130  sql = "SELECT DISTINCT T.Code, T.shortname " & _
            "FROM Demographics D, EndResults R, EndTestDefinitions T " & _
            "WHERE D.SampleID = R.SampleID " & _
            "AND R.Code = T.Code " & _
            "AND T.InUse = 1 " & _
            "AND (D.RunDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' AND '" & Format(dtTo, "dd/MMM/yyyy") & "') " & _
            "AND D." & Choice & " <> '' " & _
            "ORDER BY T.shortname"

5140  Set tb = New Recordset
5150  RecOpenServer 0, tb, sql
5160  Do While Not tb.EOF
5170    n = grdStats.Cols - 1
5180    grdStats.TextMatrix(0, n) = tb!Code
5190    grdStats.Cols = grdStats.Cols + 1
5200    tb.MoveNext
5210  Loop
5220  If grdStats.Cols > 2 Then grdStats.Cols = grdStats.Cols - 1

5230  Exit Sub

FillColsEnd_Error:

      Dim strES As String
      Dim intEL As Integer

5240  intEL = Erl
5250  strES = Err.Description
5260  LogError "frmSuperStats", "FillColsEnd", intEL, strES, sql

End Sub

