VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmViewWards 
   Caption         =   "NetAcquire"
   ClientHeight    =   7560
   ClientLeft      =   195
   ClientTop       =   435
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   12750
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7185
      Left            =   210
      TabIndex        =   4
      Top             =   150
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12674
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
      FormatString    =   $"frmViewWards.frx":0000
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
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1100
      Left            =   11220
      Picture         =   "frmViewWards.frx":008D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6090
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   3975
      Left            =   10710
      TabIndex        =   0
      Top             =   450
      Width           =   1785
      Begin VB.CheckBox optDept 
         Caption         =   "Log In/Out"
         Height          =   285
         Index           =   3
         Left            =   420
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox optDept 
         Caption         =   "Printing"
         Height          =   255
         Index           =   4
         Left            =   420
         TabIndex        =   9
         Top             =   2190
         Width           =   825
      End
      Begin VB.CheckBox optDept 
         Caption         =   "Haem"
         Height          =   285
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   1140
         Width           =   735
      End
      Begin VB.CheckBox optDept 
         Caption         =   "Bio"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   7
         Top             =   1410
         Width           =   525
      End
      Begin VB.CheckBox optDept 
         Caption         =   "Coag"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   6
         Top             =   1680
         Width           =   675
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   1155
         Left            =   390
         Picture         =   "frmViewWards.frx":0F57
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2610
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219152385
         CurrentDate     =   37679
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219152385
         CurrentDate     =   37679
      End
   End
End
Attribute VB_Name = "frmViewWards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

49760 On Error GoTo FillG_Error

49770 g.Rows = 2
49780 g.AddItem ""
49790 g.RemoveItem 1
49800 g.Refresh
49810 g.Visible = False

49820 If optDept(0) <> 1 And _
         optDept(1) <> 1 And _
         optDept(2) <> 1 And _
         optDept(3) <> 1 And _
         optDept(4) <> 1 Then Exit Sub

49830 sql = "SELECT DISTINCT " & _
            "CASE Discipline " & _
            "  WHEN 'A' THEN 'Results Overview' " & _
            "  WHEN 'B' THEN 'Biochemistry Result' " & _
            "  WHEN 'C' THEN 'Coagulation Result' " & _
            "  WHEN 'D' THEN 'Biochemistry History' " & _
            "  WHEN 'E' THEN 'Coagulation History' " & _
            "  WHEN 'F' THEN 'Haematology History' " & _
            "  WHEN 'G' THEN 'Haematology Graphs' " & _
            "  WHEN 'H' THEN 'Cumulative Haematology' " & _
            "  WHEN 'I' THEN 'Biochemistry Printout' " & _
            "  WHEN 'J' THEN 'Haematology Printout' " & _
            "  WHEN 'K' THEN 'Coagulation Printout' " & _
            "  WHEN 'L' THEN 'Log On' " & _
            "  WHEN 'M' THEN 'Manual Log Off' " & _
            "  WHEN 'N' THEN 'Microbiology Printout' " & _
            "  WHEN 'O' THEN 'Auto Log Off' " & _
            "  WHEN 'R' THEN 'Haematology Result' " & _
            "  WHEN 'X' THEN 'Close Program' END Disc, " & _
            "DateTime, Viewer, SampleID, Chart, Usercode FROM ViewedReports WHERE " & _
            "DateTime BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
            "AND '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
            "AND ("
49840 If optDept(0) Or optDept(1) Or optDept(2) Then
49850   sql = sql & "Discipline = 'A' or "
49860 End If
49870 If optDept(0) Then 'Haem
49880   sql = sql & "Discipline = 'F' or " & _
                    "Discipline = 'G' or " & _
                    "Discipline = 'H' or " & _
                    "Discipline = 'R' or "
49890 End If
49900 If optDept(1) Then 'Bio
49910   sql = sql & "Discipline = 'B' or " & _
                    "Discipline = 'D' or "
49920 End If
49930 If optDept(2) Then 'Coag
49940   sql = sql & "Discipline = 'C' or " & _
                    "Discipline = 'E' or "
49950 End If
49960 If optDept(3) Then 'LogIn/Out
49970   sql = sql & "Discipline = 'L' or " & _
                    "Discipline = 'M' or " & _
                    "Discipline = 'O' or " & _
                    "Discipline = 'X' or "
49980 End If
49990 If optDept(4) Then 'Printing
50000   sql = sql & "Discipline = 'I' or " & _
                    "Discipline = 'J' or " & _
                    "Discipline = 'K' or " & _
                    "Discipline = 'N' or "
50010 End If
50020 sql = Left$(sql, Len(sql) - 3) & ") " & _
            "ORDER BY DateTime DESC"
50030 Set tb = New Recordset
50040 RecOpenClient 0, tb, sql
50050 Do While Not tb.EOF
50060   s = tb!Chart & vbTab & _
            Format$(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!SampleID & vbTab & _
            tb!UserCode & vbTab & _
            tb!Viewer & vbTab & _
            tb!Disc & ""
50070   g.AddItem s
50080   tb.MoveNext
50090 Loop

50100 If g.Rows > 2 Then
50110   g.RemoveItem 1
50120 End If
50130 g.Visible = True

50140 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

50150 intEL = Erl
50160 strES = Err.Description
50170 LogError "fViewWards", "FillG", intEL, strES, sql
50180 g.Visible = True

End Sub

Private Sub bcancel_Click()

50190 Unload Me

End Sub


Private Sub cmdRefresh_Click()

50200 FillG

End Sub

Private Sub Form_Activate()

50210 Activated = True

End Sub

Private Sub Form_Load()

      Dim Checks As String
      Dim n As Integer

50220 Activated = False

50230 Checks = GetSetting("NetAcquire", "StartUp", "Check", "0000")
50240 For n = 0 To 3
50250   optDept(n) = IIf(Mid$(Checks, n + 1, 1) = "1", 1, 0)
50260 Next

50270 dtTo = Format$(Now, "dd/mm/yyyy")
50280 dtFrom = Format$(Now, "dd/mm/yyyy")

50290 FillG

End Sub


Private Sub Form_Unload(Cancel As Integer)

      Dim Checks As String
      Dim n As Integer

50300 Activated = False

50310 Checks = ""
50320 For n = 0 To 3
50330   Checks = Checks & IIf(optDept(n), "1", "0")
50340 Next
50350 SaveSetting "NetAcquire", "StartUp", "Check", Checks

50360 Activated = False

End Sub

Private Sub g_Click()

      Static SortOrder As Boolean

50370 If g.MouseRow = 0 Then
50380   If SortOrder Then
50390     g.Sort = flexSortGenericAscending
50400   Else
50410     g.Sort = flexSortGenericDescending
50420   End If
50430   SortOrder = Not SortOrder
50440 End If

End Sub


