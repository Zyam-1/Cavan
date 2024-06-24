VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUnfinished 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Unfinished Samples"
   ClientHeight    =   8550
   ClientLeft      =   2115
   ClientTop       =   1530
   ClientWidth     =   8535
   Icon            =   "frmUnfinished.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Department"
      Height          =   3075
      Left            =   6750
      TabIndex        =   5
      Top             =   1860
      Width           =   1545
      Begin VB.OptionButton optDept 
         Caption         =   "Red Sub"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   15
         Top             =   2730
         Width           =   945
      End
      Begin VB.OptionButton optDept 
         Caption         =   "C && S"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   14
         Top             =   2460
         Width           =   765
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Ova/Parasites"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   2190
         Width           =   1365
      End
      Begin VB.OptionButton optDept 
         Caption         =   "CSF"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   705
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Rota/Adeno"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1650
         Width           =   1215
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Urine"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1380
         Width           =   705
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Faeces"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1110
         Width           =   945
      End
      Begin VB.OptionButton optDept 
         Caption         =   "c. Diff"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   765
      End
      Begin VB.OptionButton optDept 
         Caption         =   "H. Pylori"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Width           =   945
      End
      Begin VB.OptionButton optDept 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   315
      Left            =   6810
      TabIndex        =   3
      Top             =   1410
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      Format          =   220200961
      CurrentDate     =   39730
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   720
      Left            =   6720
      Picture         =   "frmUnfinished.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7620
      Width           =   1470
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   8265
      Left            =   300
      TabIndex        =   0
      Top             =   45
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   14579
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Sample Number    |<Outstanding                     |<Date                      "
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "View Samples dated later than"
      Height          =   465
      Left            =   6810
      TabIndex        =   4
      Top             =   930
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6600
      Picture         =   "frmUnfinished.frx":0614
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Heading to Sort"
      Height          =   435
      Left            =   7080
      TabIndex        =   2
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "frmUnfinished"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG(ByVal DeptCode As String, _
                  ByVal DeptNumber As Integer)
          
      Dim sql As String
      Dim tb As Recordset
      Dim Department(1 To 9) As String
      Dim s As String

26990 On Error GoTo FillG_Error

27000 Department(1) = "H. Pylori"
27010 Department(2) = "c. Diff"
27020 Department(3) = "Faeces"
27030 Department(4) = "Urine"
27040 Department(5) = "Rota/Adeno"
27050 Department(6) = "CSF"
27060 Department(7) = "Ova/Parasites"
27070 Department(8) = "C & S"
27080 Department(9) = "Red Sub"

27090 sql = "SELECT P.SampleID, D.RunDate FROM PrintValidLog P, Demographics D WHERE " & _
            "P.Department = '" & DeptCode & "' " & _
            "AND P.Valid = 0 " & _
            "AND P.SampleID = D.SampleID " & _
            "AND D.RunDate >= '" & Format$(dt, "dd/MMM/yyyy") & "'"
27100 Set tb = New Recordset
27110 RecOpenServer 0, tb, sql
27120 Do While Not tb.EOF
27130   s = Format$(Val(tb!SampleID)) & vbTab & _
            Department(DeptNumber) & vbTab & _
            Format$(tb!Rundate, "dd/MM/yyyy")
27140   g.AddItem s
27150   tb.MoveNext
27160 Loop

27170 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

27180 intEL = Erl
27190 strES = Err.Description
27200 LogError "frmUnfinished", "FillG", intEL, strES, sql


End Sub

Private Sub RefreshGrid()
        
      Dim Dept As String
      Dim n As Integer
      Dim OptionSelected As Integer

27210 On Error GoTo RefreshGrid_Error

27220 g.Rows = 2
27230 g.AddItem ""
27240 g.RemoveItem 1
27250 g.Visible = False
27260 Screen.MousePointer = vbHourglass

27270 OptionSelected = 0
27280 For n = 0 To 9
27290   If optDept(n).Value Then
27300     OptionSelected = n
27310     Exit For
27320   End If
27330 Next

27340 Dept = "YGFUACODR"

27350 If OptionSelected = 0 Then
27360   For n = 1 To Len(Dept)
27370     FillG Mid$(Dept, n, 1), n
27380   Next
27390 Else
27400     FillG Mid$(Dept, OptionSelected, 1), n
27410 End If

27420 g.Col = 0
27430 g.Sort = flexSortGenericAscending
27440 If g.Rows > 2 Then
27450   g.RemoveItem 1
27460 End If
27470 g.Visible = True
27480 Screen.MousePointer = vbNormal

27490 Exit Sub

RefreshGrid_Error:

      Dim strES As String
      Dim intEL As Integer

27500 intEL = Erl
27510 strES = Err.Description
27520 LogError "frmUnfinished", "RefreshGrid", intEL, strES
27530 Screen.MousePointer = vbNormal

End Sub

Private Sub bcancel_Click()

27540 Unload Me

End Sub


Private Sub dt_CloseUp()

27550 RefreshGrid

End Sub


Private Sub Form_Activate()

27560 RefreshGrid

End Sub

Private Sub Form_Load()

27570 dt = Format$(Now - 3, "dd/MM/yyyy")

End Sub


Private Sub g_Click()


27580 On Error GoTo g_Click_Error

27590 If g.MouseRow = 0 Then
27600   If InStr(UCase$(g.TextMatrix(0, g.MouseCol)), "DATE") <> 0 Then
27610     g.Sort = 9
27620   ElseIf SortOrder Then
27630     g.Sort = flexSortGenericAscending
27640   Else
27650     g.Sort = flexSortGenericDescending
27660   End If
27670   SortOrder = Not SortOrder
27680   Exit Sub
27690 End If

      'frmEditMicrobiology.forcedSID = g.TextMatrix(g.Row, 0)
      'frmEditMicrobiologyNew.Show 1

27700 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

27710 intEL = Erl
27720 strES = Err.Description
27730 LogError "frmUnfinished", "g_Click", intEL, strES

End Sub
Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

27740 If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
27750   Cmp = 0
27760   Exit Sub
27770 End If

27780 If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
27790   Cmp = 0
27800   Exit Sub
27810 End If

27820 d1 = Format(g.TextMatrix(Row1, g.Col), "dd/MMM/yyyy HH:mm:ss")
27830 d2 = Format(g.TextMatrix(Row2, g.Col), "dd/MMM/yyyy HH:mm:ss")

27840 If SortOrder Then
27850   Cmp = Sgn(DateDiff("s", d1, d2))
27860 Else
27870   Cmp = Sgn(DateDiff("s", d2, d1))
27880 End If

End Sub


Private Sub optDept_Click(Index As Integer)

27890 RefreshGrid

End Sub


