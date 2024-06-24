VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListingBioCD4TMImm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Run Dates Between"
      Height          =   3075
      Left            =   13290
      TabIndex        =   7
      Top             =   990
      Width           =   1755
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "Calculate"
         Height          =   975
         Left            =   330
         Picture         =   "frmListingBioCD4TMImm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1860
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   270
         TabIndex        =   12
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   216793089
         CurrentDate     =   40246
      End
      Begin VB.OptionButton optMonthly 
         Caption         =   "Monthly"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1530
         Width           =   945
      End
      Begin VB.OptionButton optWeekly 
         Caption         =   "Weekly"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   1260
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton optDaily 
         Caption         =   "Daily"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   990
         Width           =   675
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   270
         TabIndex        =   8
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   216793089
         CurrentDate     =   36966
      End
   End
   Begin VB.ComboBox cmbWard 
      Height          =   315
      Left            =   13230
      TabIndex        =   4
      Text            =   "cmbWard"
      Top             =   480
      Width           =   1995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   1100
      Left            =   13620
      Picture         =   "frmListingBioCD4TMImm.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1200
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1100
      Left            =   13620
      Picture         =   "frmListingBioCD4TMImm.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7935
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   13996
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
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
      FormatString    =   $"frmListingBioCD4TMImm.frx":209E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   13560
      TabIndex        =   6
      Top             =   5790
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ward"
      Height          =   195
      Left            =   14025
      TabIndex        =   5
      Top             =   240
      Width           =   390
   End
End
Attribute VB_Name = "frmListingBioCD4TMImm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ClearOptions()

10590     optDaily = 0
10600     optWeekly = 0
10610     optMonthly = 0

End Sub

Private Sub FillG()

          Dim sql As String
          Dim tb As Recordset
10620     ReDim sLine(0 To 0) As String
          Dim X As Integer
          Dim sInclude As String
10630     ReDim GridParam(0 To 0) As String
          Dim n As Integer
          Dim GridFormat As String
          Dim s As String
          Dim SurForeName() As String
          Dim FullName As String
          Dim FromDate As String
          Dim ToDate As String

10640     On Error GoTo FillG_Error

10650     g.Rows = 2
10660     g.AddItem ""
10670     g.RemoveItem 1
10680     g.Refresh
10690     g.Visible = False

10700     FromDate = Format$(dtFrom, "dd/MMM/yyyy") & " 00:00"
10710     ToDate = Format$(dtTo, "dd/MMM/yyyy") & " 23:59"

10720     GridFormat = "<SampleID      " & _
              "|<Patient Name       " & _
              "|<Chart        " & _
              "|<D.o.B.      " & _
              "|<Ward         " & _
              "|<Clinician    "

10730     If cmbWard <> "All" Then
10740         sql = "SELECT DISTINCT Code, ShortName FROM BioTestDefinitions WHERE " & _
                  "Code IN ( " & _
                  "  SELECT DISTINCT Code FROM BioResults R JOIN Demographics D " & _
                  "  ON R.SampleID = D.SampleID " & _
                  "  WHERE " & _
                  "  Ward = '" & cmbWard & "' " & _
                  "  AND R.RunTime BETWEEN '" & FromDate & "' AND '" & ToDate & "' ) "
10750     Else
10760         sql = "SELECT DISTINCT Code, ShortName FROM BioTestDefinitions WHERE " & _
                  "Code IN ( " & _
                  "  SELECT DISTINCT Code FROM BioResults R WHERE " & _
                  "  R.RunTime BETWEEN '" & FromDate & "' AND '" & ToDate & "' ) "
10770     End If
10780     Set tb = New Recordset
10790     RecOpenServer 0, tb, sql
10800     If tb.EOF Then
10810         g.Cols = 6
10820         g.FormatString = GridFormat
10830         g.Visible = True
10840         Exit Sub
10850     End If

10860     sInclude = ""
10870     X = -1
10880     Do While Not tb.EOF
10890         X = X + 1
10900         ReDim Preserve sLine(0 To X)
10910         ReDim Preserve GridParam(0 To X)
10920         sLine(X) = "(SELECT Result FROM BioResults WHERE " & _
                  "Code = '" & tb!Code & "' " & _
                  "AND SampleID = D.SampleID ) " & _
                  "[" & tb!ShortName & "], "
10930         sInclude = sInclude & "'" & tb!Code & "', "
10940         GridParam(X) = tb!ShortName & ""
10950         tb.MoveNext
10960     Loop
10970     sLine(UBound(sLine)) = Left$(sLine(UBound(sLine)), Len(sLine(UBound(sLine))) - 2) & " "
10980     sInclude = Left$(sInclude, Len(sInclude) - 2)

10990     For X = 0 To UBound(GridParam)
11000         GridFormat = GridFormat & "|<" & GridParam(X) & "   "
11010     Next
11020     g.Cols = 6 + UBound(GridParam) + 1
11030     g.FormatString = GridFormat

11040     sql = "SELECT  D.SampleID, D.PatName, D.Chart, D.DoB, D.Ward, D.Clinician, "
11050     For n = 0 To UBound(sLine)
11060         sql = sql & sLine(n)
11070     Next
11080     sql = sql & "FROM Demographics D JOIN BioResults R " & _
              "ON R.SampleID = D.SampleID " & _
              "WHERE "
11090     If cmbWard <> "All" Then
11100         sql = sql & "Ward = '" & cmbWard & "' AND "
11110     End If
11120     sql = sql & "( R.RunTime BETWEEN '" & FromDate & "' AND '" & ToDate & "') " & _
              "AND R.Code IN (" & sInclude & ") " & _
              "GROUP BY D.SampleID, R.SampleID, D.PatName, D.Chart, D.DoB, D.Ward, D.Clinician " & _
              "ORDER BY R.SampleID"
11130     Set tb = New Recordset
11140     RecOpenServer 0, tb, sql
11150     Do While Not tb.EOF
11160         s = tb!SampleID & vbTab
        
11170         FullName = tb!PatName & ""
11180         SurForeName = Split(tb!PatName & "", " ")
11190         If UBound(SurForeName) > 0 Then
11200             If SurForeName(0) = SurForeName(1) Then
11210                 FullName = SurForeName(0)
11220             End If
11230         End If
11240         s = s & FullName & vbTab & _
                  tb!Chart & vbTab & _
                  tb!DoB & vbTab & _
                  tb!Ward & vbTab & _
                  tb!Clinician & ""
        
11250         For X = 0 To UBound(GridParam)
11260             s = s & vbTab & tb(GridParam(X)) & ""
11270         Next
11280         g.AddItem s
11290         tb.MoveNext
11300     Loop

11310     If g.Rows > 2 Then
11320         g.RemoveItem 1
11330     End If
11340     g.Visible = True

11350     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

11360     intEL = Erl
11370     strES = Err.Description
11380     LogError "frmListingBioCD4TMImm", "FillG", intEL, strES, sql
11390     g.Visible = True

End Sub


Private Sub FillWard()

          Dim sql As String
          Dim tb As Recordset

11400     On Error GoTo FillWard_Error

11410     cmbWard.Clear

11420     sql = "SELECT DISTINCT Ward FROM Demographics " & _
              "WHERE COALESCE(Ward, '') <> '' " & _
              "ORDER BY Ward"
11430     Set tb = New Recordset
11440     RecOpenServer 0, tb, sql
11450     Do While Not tb.EOF
11460         cmbWard.AddItem tb!Ward
11470         tb.MoveNext
11480     Loop
11490     cmbWard.AddItem "All", 0
11500     cmbWard.ListIndex = 0

11510     Exit Sub

FillWard_Error:

          Dim strES As String
          Dim intEL As Integer

11520     intEL = Erl
11530     strES = Err.Description
11540     LogError "frmListingBioCD4TMImm", "FillWard", intEL, strES, sql

End Sub

Private Sub cmdCalculate_Click()

11550     If Abs(DateDiff("m", dtFrom, dtTo)) > 1 Then
11560         iMsg "Date difference too wide", vbInformation
11570         Exit Sub
11580     End If

11590     FillG

End Sub

Private Sub cmdCancel_Click()

11600     Unload Me

End Sub

Private Sub cmdXL_Click()

11610     On Error GoTo cmdXL_Click_Error

11620     ExportFlexGrid g, Me, "Biochemistry Results" & vbCr & _
              "Between " & dtFrom & " and " & dtTo & vbCr & _
              "For " & cmbWard & vbCr

11630     Exit Sub

cmdXL_Click_Error:

          Dim strES As String
          Dim intEL As Integer

11640     intEL = Erl
11650     strES = Err.Description
11660     LogError "frmListingBioCD4TMImm", "cmdXL_Click", intEL, strES

End Sub

Private Sub dtFrom_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

11670     ClearOptions

End Sub

Private Sub dtTo_CloseUp()

11680     ClearOptions

End Sub


Private Sub Form_Activate()

11690     Me.Caption = "NetAcquire Daily Report "

11700     FillWard

End Sub

Private Sub Form_Load()

11710     dtFrom = Format(Now - 7, "dd/MM/yyyy")
11720     dtTo = Format$(Now, "dd/MM/yyyy")

End Sub



Private Sub g_Click()

          Static SortOrder As Boolean

11730     On Error GoTo g_Click_Error

11740     If g.MouseRow = 0 Then
11750         If SortOrder Then
11760             g.Sort = flexSortStringAscending
11770         Else
11780             g.Sort = flexSortStringDescending
11790         End If
11800         SortOrder = Not SortOrder
11810         Exit Sub
11820     End If

11830     Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

11840     intEL = Erl
11850     strES = Err.Description
11860     LogError "frmListingBioCD4TMImm", "g_Click", intEL, strES

End Sub


Private Sub optDaily_Click()

11870     dtFrom = dtTo

End Sub


Private Sub optMonthly_Click()

11880     dtFrom = DateAdd("m", -1, dtTo)

End Sub


Private Sub optWeekly_Click()

11890     dtFrom = DateAdd("ww", -1, dtTo)

End Sub


