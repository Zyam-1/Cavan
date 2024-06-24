VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBioEndoTotals 
   Caption         =   "Biochemistry / Endocrinology Totals"
   ClientHeight    =   5355
   ClientLeft      =   690
   ClientTop       =   510
   ClientWidth     =   7170
   LinkTopic       =   "Form2"
   ScaleHeight     =   5355
   ScaleWidth      =   7170
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   4560
      Picture         =   "frmBioEndoTotals.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   1485
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   4560
      Picture         =   "frmBioEndoTotals.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   270
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      Caption         =   "Between Dates"
      Height          =   2685
      Left            =   450
      TabIndex        =   1
      Top             =   180
      Width           =   3645
      Begin VB.CommandButton breCalc 
         Caption         =   "Calculate"
         Height          =   945
         Left            =   2070
         Picture         =   "frmBioEndoTotals.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1110
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Today"
         Height          =   195
         Index           =   6
         Left            =   990
         TabIndex        =   8
         Top             =   720
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Year To Date"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   7
         Top             =   2310
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Full Quarter"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   6
         Top             =   2040
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Quarter"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   5
         Top             =   1770
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Full Month"
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   4
         Top             =   1500
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Month"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   3
         Top             =   1230
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Week"
         Height          =   225
         Index           =   0
         Left            =   690
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218824705
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   330
         TabIndex        =   11
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218824705
         CurrentDate     =   38126
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   1305
      Left            =   510
      TabIndex        =   0
      Top             =   3150
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   2302
      _Version        =   393216
      Rows            =   4
      Cols            =   4
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
      FormatString    =   "<                                    |<Alpha + Beta |<Endocrinology |<HbA1c  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblEvents 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2850
      TabIndex        =   16
      Top             =   4590
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Labels Used"
      Height          =   195
      Left            =   1890
      TabIndex        =   15
      Top             =   4650
      Width           =   885
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   315
      Left            =   4560
      TabIndex        =   14
      Top             =   1110
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "frmBioEndoTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String

7370      On Error GoTo FillGrid_Error

7380      sql = "SELECT COUNT(DISTINCT sampleid) AS Tot From BioResults WHERE " & _
              "RunDate BETWEEN '" & Format$(dtFrom, "dd/MMM/yyyy") & _
              "' AND '" & Format$(dtTo, "dd/MMM/yyyy") & "' " & _
              "AND (Analyser = 'A' OR Analyser = 'B')"
7390      Set tb = New Recordset
7400      RecOpenClient 0, tb, sql
7410      g.TextMatrix(1, 1) = tb!Tot

7420      sql = "SELECT COUNT(DISTINCT sampleid) AS Tot From BioResults WHERE " & _
              "RunDate BETWEEN '" & Format$(dtFrom, "dd/MMM/yyyy") & _
              "' AND '" & Format$(dtTo, "dd/MMM/yyyy") & "' " & _
              "AND (Analyser = '4')"
7430      Set tb = New Recordset
7440      RecOpenClient 0, tb, sql
7450      g.TextMatrix(1, 2) = tb!Tot

7460      sql = "SELECT COUNT(DISTINCT sampleid) AS Tot From BioResults WHERE " & _
              "RunDate BETWEEN '" & Format$(dtFrom, "dd/MMM/yyyy") & _
              "' AND '" & Format$(dtTo, "dd/MMM/yyyy") & "' " & _
              "AND (Analyser is null)"
7470      Set tb = New Recordset
7480      RecOpenClient 0, tb, sql
7490      g.TextMatrix(1, 3) = tb!Tot




7500      sql = "SELECT COUNT(DISTINCT b.sampleid) AS Tot From BioResults as b, demographics as d WHERE " & _
              "b.RunDate BETWEEN '" & Format$(dtFrom, "dd/MMM/yyyy") & _
              "' AND '" & Format$(dtTo, "dd/MMM/yyyy") & "' " & _
              "AND (b.Analyser = 'A' OR b.Analyser = 'B') " & _
              "and b.SampleID = d.sampleid " & _
              "and d.Ward = 'GP'"

7510      Set tb = New Recordset
7520      RecOpenClient 0, tb, sql
7530      g.TextMatrix(2, 1) = tb!Tot

7540      sql = "SELECT COUNT(DISTINCT b.sampleid) AS Tot From BioResults as b, demographics as d WHERE " & _
              "b.RunDate BETWEEN '" & Format$(dtFrom, "dd/MMM/yyyy") & _
              "' AND '" & Format$(dtTo, "dd/MMM/yyyy") & "' " & _
              "AND (b.Analyser = '4') " & _
              "and b.SampleID = d.sampleid " & _
              "and d.Ward = 'GP'"
7550      Set tb = New Recordset
7560      RecOpenClient 0, tb, sql
7570      g.TextMatrix(2, 2) = tb!Tot

7580      sql = "SELECT COUNT(DISTINCT b.sampleid) AS Tot From BioResults as b, demographics as d WHERE " & _
              "b.RunDate BETWEEN '" & Format$(dtFrom, "dd/MMM/yyyy") & _
              "' AND '" & Format$(dtTo, "dd/MMM/yyyy") & "' " & _
              "AND (b.Analyser is null) " & _
              "and b.SampleID = d.SampleID " & _
              "and d.Ward = 'GP'"
7590      Set tb = New Recordset
7600      RecOpenClient 0, tb, sql
7610      g.TextMatrix(2, 3) = tb!Tot
        
7620      g.TextMatrix(3, 1) = Format$(Val(g.TextMatrix(1, 1)) - Val(g.TextMatrix(2, 1)))
7630      g.TextMatrix(3, 2) = Format$(Val(g.TextMatrix(1, 2)) - Val(g.TextMatrix(2, 2)))
7640      g.TextMatrix(3, 3) = Format$(Val(g.TextMatrix(1, 3)) - Val(g.TextMatrix(2, 3)))

7650      sql = "SELECT COUNT(DISTINCT sampleid) AS Tot From BioResults WHERE " & _
              "RunDate BETWEEN '" & Format$(dtFrom, "dd/MMM/yyyy") & _
              "' AND '" & Format$(dtTo, "dd/MMM/yyyy") & "'"
7660      Set tb = New Recordset
7670      RecOpenClient 0, tb, sql
7680      lblEvents = tb!Tot

7690      Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

7700      intEL = Erl
7710      strES = Err.Description
7720      LogError "frmBioEndoTotals", "FillGrid", intEL, strES, sql

End Sub
Private Sub Label13_Click()

End Sub

Private Sub Label5_Click()

End Sub


Private Sub breCalc_Click()

7730      FillGrid

End Sub

Private Sub cmdCancel_Click()

7740      Unload Me

End Sub


Private Sub cmdXL_Click()

7750      ExportFlexGrid g, Me

End Sub


Private Sub Form_Load()

7760      dtFrom = Format$(Now, "dd/mmm/yyyy")
7770      dtTo = dtFrom

7780      g.TextMatrix(1, 0) = "Total"
7790      g.TextMatrix(2, 0) = "GP"
7800      g.TextMatrix(3, 0) = "Total - GP"

End Sub


Private Sub obetween_Click(Index As Integer)

          Dim UpTo As String

7810      dtFrom = BetweenDates(Index, UpTo)
7820      dtTo = UpTo

End Sub


