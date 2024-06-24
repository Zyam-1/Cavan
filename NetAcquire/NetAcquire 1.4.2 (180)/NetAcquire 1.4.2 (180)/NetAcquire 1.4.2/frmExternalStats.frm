VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExternalStats 
   Caption         =   "NetAcquire - External Tests"
   ClientHeight    =   7950
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   8490
   Begin VB.CommandButton cmdAnalytes 
      Caption         =   "Export Analytes to Excel"
      Height          =   765
      Left            =   7020
      Picture         =   "frmExternalStats.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3150
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Between Dates"
      Height          =   1305
      Left            =   270
      TabIndex        =   12
      Top             =   90
      Width           =   6585
      Begin VB.CommandButton breCalc 
         Caption         =   "Calculate"
         Height          =   795
         Left            =   5070
         Picture         =   "frmExternalStats.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Today"
         Height          =   195
         Index           =   6
         Left            =   2550
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Year To Date"
         Height          =   195
         Index           =   5
         Left            =   3390
         TabIndex        =   22
         Top             =   960
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Quarter"
         Height          =   195
         Index           =   4
         Left            =   3390
         TabIndex        =   21
         Top             =   720
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Quarter"
         Height          =   195
         Index           =   3
         Left            =   3390
         TabIndex        =   20
         Top             =   480
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Month"
         Height          =   195
         Index           =   2
         Left            =   3390
         TabIndex        =   19
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Month"
         Height          =   195
         Index           =   1
         Left            =   2190
         TabIndex        =   18
         Top             =   705
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Week"
         Height          =   225
         Index           =   0
         Left            =   2250
         TabIndex        =   17
         Top             =   465
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   510
         TabIndex        =   13
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218824705
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   510
         TabIndex        =   14
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218824705
         CurrentDate     =   38126
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   780
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export Sources to Excel"
      Height          =   765
      Left            =   7020
      Picture         =   "frmExternalStats.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2370
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdAnalyte 
      Height          =   6165
      Left            =   4020
      TabIndex        =   8
      ToolTipText     =   "Click on Heading to Sort"
      Top             =   1620
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   10874
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      FormatString    =   "<Analyte                           |^Total     "
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   7020
      Picture         =   "frmExternalStats.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1620
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   7080
      Picture         =   "frmExternalStats.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7110
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   4905
      IntegralHeight  =   0   'False
      Left            =   8340
      TabIndex        =   4
      Top             =   210
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source"
      Height          =   525
      Left            =   270
      TabIndex        =   0
      Top             =   1560
      Width           =   3645
      Begin VB.OptionButton oSource 
         Caption         =   "G.P.s"
         Height          =   195
         Index           =   2
         Left            =   2310
         TabIndex        =   3
         Top             =   210
         Width           =   825
      End
      Begin VB.OptionButton oSource 
         Caption         =   "Wards"
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   2
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton oSource 
         Caption         =   "Clinicians"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5325
      Left            =   240
      TabIndex        =   7
      Top             =   2460
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   9393
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   2
      FormatString    =   "<Source               |<Samples |<Tests      |<T/S      "
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Generating Report - Please wait"
      Height          =   765
      Left            =   7020
      TabIndex        =   11
      Top             =   5430
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   6990
      TabIndex        =   10
      Top             =   3930
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   600
      Picture         =   "frmExternalStats.frx":15F2
      Stretch         =   -1  'True
      Top             =   2100
      Width           =   660
   End
End
Attribute VB_Name = "frmExternalStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim total As Long
          Dim Samples As Long
          Dim Tests As Long
          Dim n As Integer
          Dim tps As String
          Dim Source As String

35520     On Error GoTo FillGrid_Error

35530     If oSource(0) Then
35540         Source = "Clinician = '"
35550     ElseIf oSource(1) Then
35560         Source = "Ward = '"
35570     Else
35580         Source = "GP = '"
35590     End If

35600     g.Rows = 2
35610     g.AddItem ""
35620     g.RemoveItem 1

35630     For n = 0 To List1.ListCount - 1
35640         List1.Selected(n) = True
35650         Tests = 0
35660         sql = "SELECT COUNT(SampleID) AS Tests FROM ExtResults WHERE " & _
                  "SampleID IN (" & _
                  "  SELECT DISTINCT D.SampleID FROM Demographics AS D WHERE " & _
                  "  RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                  Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59'  " & _
                  "  AND " & Source & AddTicks(List1.List(n)) & "')"
35670         Set tb = New Recordset
35680         RecOpenServer 0, tb, sql
35690         Tests = tb!Tests
        
35700         Samples = 0
35710         sql = "SELECT COUNT(DISTINCT SampleID) AS Samples FROM ExtResults WHERE " & _
                  "SampleID IN (" & _
                  "  SELECT DISTINCT D.SampleID FROM Demographics AS D WHERE " & _
                  "  RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                  Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59'  " & _
                  "  AND " & Source & AddTicks(List1.List(n)) & "')"
35720         Set tb = New Recordset
35730         RecOpenServer 0, tb, sql
35740         Samples = tb!Samples
        
35750         If Tests <> 0 And Samples <> 0 Then
35760             tps = Format$(Tests / Samples, "##.00")
35770             g.AddItem List1.List(n) & vbTab & Samples & vbTab & Tests & vbTab & tps
35780             g.Refresh
35790         End If
35800     Next
35810     g.AddItem ""

35820     If g.Rows = 2 Then Exit Sub

35830     g.Col = 1
35840     total = 0
35850     For n = 2 To g.Rows - 1
35860         g.row = n
35870         total = total + Val(g)
35880     Next

35890     g.Col = 2
35900     Tests = 0
35910     For n = 2 To g.Rows - 1
35920         g.row = n
35930         Tests = Tests + Val(g)
35940     Next

35950     If total <> 0 And Tests <> 0 Then
35960         g.AddItem "Total" & vbTab & total & vbTab & Tests & vbTab & Format$(Tests / total, ".00")
35970     Else
35980         g.AddItem "Total"
35990     End If
36000     g.Refresh

36010     g.RemoveItem 1

36020     Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

36030     intEL = Erl
36040     strES = Err.Description
36050     LogError "frmExternalStats", "FillGrid", intEL, strES, sql

End Sub
Private Sub FillGridAnalyte()

          Dim tbRes As Recordset
          Dim sql As String
          Dim Tests As Long
          Dim n As Long
          Dim Found As Boolean

36060     On Error GoTo FillGridAnalyte_Error

36070     With grdAnalyte
36080         .Rows = 2
36090         .AddItem ""
36100         .RemoveItem 1
36110     End With

36120     sql = "SELECT * FROM ExtResults WHERE " & _
              "SampleID IN ( " & _
              "  SELECT SampleID FROM Demographics WHERE " & _
              "  RunDate BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' AND '" & _
              Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' )"
36130     Set tbRes = New Recordset
36140     RecOpenServer 0, tbRes, sql
36150     Do While Not tbRes.EOF
36160         Found = False
36170         For n = 1 To grdAnalyte.Rows - 1
36180             If Trim$(tbRes!Analyte & "") = grdAnalyte.TextMatrix(n, 0) Then
36190                 grdAnalyte.TextMatrix(n, 1) = Format$(Val(grdAnalyte.TextMatrix(n, 1)) + 1)
36200                 Found = True
36210                 Exit For
36220             End If
36230         Next
36240         If Not Found Then
36250             grdAnalyte.AddItem Trim$(tbRes!Analyte & "") & vbTab & "1"
36260         End If
36270         tbRes.MoveNext
36280     Loop

36290     Tests = 0
36300     With grdAnalyte
36310         For n = 1 To .Rows - 1
36320             Tests = Tests + Val(.TextMatrix(n, 1))
36330         Next
36340         .AddItem "Total" & vbTab & Tests

36350         If .Rows > 2 Then
36360             .RemoveItem 1
36370         End If
36380     End With

36390     Exit Sub

FillGridAnalyte_Error:

          Dim strES As String
          Dim intEL As Integer

36400     intEL = Erl
36410     strES = Err.Description
36420     LogError "frmExternalStats", "FillGridAnalyte", intEL, strES, sql

End Sub

Private Sub FillList()

          Dim sql As String
          Dim tb As Recordset
          Dim strSource As String
          Dim Y As Integer
          Dim Found As Boolean

36430     On Error GoTo FillList_Error

36440     List1.Clear

36450     strSource = Switch(oSource(0), "Clinician", _
              oSource(1), "Ward", _
              oSource(2), "GP")
        
36460     sql = "Select distinct " & strSource & " as Source " & _
              "from Demographics where " & _
              "RunDate between '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
              "and '" & Format(dtTo, "dd/mmm/yyyy") & "' " & _
              "Order by " & strSource
36470     Set tb = New Recordset
36480     RecOpenServer 0, tb, sql
36490     Do While Not tb.EOF
36500         Debug.Print tb!Source & ""
36510         Found = False
36520         For Y = 0 To List1.ListCount - 1
36530             If UCase$(Trim$(tb!Source & "")) = UCase$(List1.List(Y)) Then
36540                 Found = True
36550                 Exit For
36560             End If
36570         Next
36580         If Not Found Then
36590             List1.AddItem Trim$(tb!Source & "")
36600         End If
36610         tb.MoveNext
36620     Loop

36630     Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

36640     intEL = Erl
36650     strES = Err.Description
36660     LogError "frmExternalStats", "FillList", intEL, strES, sql


End Sub


Private Sub breCalc_Click()

36670     lblGen.Visible = True
36680     lblGen.Refresh
36690     FillList
36700     FillGrid
36710     FillGridAnalyte
36720     lblGen.Visible = False

End Sub


Private Sub cmdAnalytes_Click()

36730     ExportFlexGrid grdAnalyte, Me

End Sub

Private Sub cmdCancel_Click()

36740     Unload Me

End Sub


Private Sub cmdPrint_Click()

          Dim n As Integer
          Dim X As Integer

36750     Printer.Print "Totals:"; dtFrom; " to "; dtTo

36760     Printer.Print
36770     For n = 0 To g.Rows - 1
36780         g.row = n
36790         For X = 0 To 3
36800             g.Col = X
36810             Printer.Print Tab(Choose(X + 1, 1, 40, 50, 60)); g;
36820         Next
36830         Printer.Print
36840     Next

36850     Printer.Print
36860     For n = 0 To grdAnalyte.Rows - 1
36870         Printer.Print grdAnalyte.TextMatrix(n, 0); Tab(25); grdAnalyte.TextMatrix(n, 1)
36880     Next

36890     Printer.EndDoc

End Sub


Private Sub cmdXL_Click()

36900     ExportFlexGrid g, Me

End Sub

Private Sub dtFrom_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

36910     FillList

End Sub


Private Sub dtTo_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

36920     FillList

End Sub


Private Sub Form_Load()

36930     dtFrom = Format$(Now, "dd/mmm/yyyy")
36940     dtTo = dtFrom

End Sub


Private Sub grdAnalyte_Click()

          Static SortOrder As Boolean

36950     If SortOrder Then
36960         grdAnalyte.Sort = flexSortGenericAscending
36970     Else
36980         grdAnalyte.Sort = flexSortGenericDescending
36990     End If

37000     SortOrder = Not SortOrder

End Sub

Private Sub Label2_Click()

End Sub

Private Sub obetween_Click(Index As Integer)

          Dim UpTo As String

37010     dtFrom = BetweenDates(Index, UpTo)
37020     dtTo = UpTo

37030     FillList

End Sub


Private Sub oSource_Click(Index As Integer)

37040     FillList

End Sub


