VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCoagTotals 
   Caption         =   "NetAcquire - Coagulation - Totals"
   ClientHeight    =   6225
   ClientLeft      =   90
   ClientTop       =   570
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8820
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   765
      Left            =   7530
      Picture         =   "frmCoagTotals.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   240
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source"
      Height          =   1125
      Left            =   3840
      TabIndex        =   15
      Top             =   30
      Width           =   1365
      Begin VB.OptionButton oSource 
         Caption         =   "Clinicians"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton oSource 
         Caption         =   "Wards"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton oSource 
         Caption         =   "G.P.s"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   810
         Width           =   825
      End
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   2565
      Index           =   1
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   3555
      TabIndex        =   4
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton breCalc 
         Caption         =   "Calculate"
         Height          =   945
         Left            =   1980
         Picture         =   "frmCoagTotals.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   930
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Today"
         Height          =   195
         Index           =   6
         Left            =   900
         TabIndex        =   11
         Top             =   540
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Year To Date"
         Height          =   195
         Index           =   5
         Left            =   390
         TabIndex        =   10
         Top             =   2130
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Full Quarter"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   9
         Top             =   1860
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Quarter"
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   8
         Top             =   1590
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Full Month"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   1320
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Month"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   6
         Top             =   1050
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Week"
         Height          =   225
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   780
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1830
         TabIndex        =   13
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   216793089
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   216793089
         CurrentDate     =   38126
      End
   End
   Begin VB.ListBox List1 
      Height          =   3255
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   5400
      Picture         =   "frmCoagTotals.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   825
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   765
      Left            =   6450
      Picture         =   "frmCoagTotals.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   825
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4695
      Left            =   3840
      TabIndex        =   0
      Top             =   1350
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   6
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
      FormatString    =   "<Source                 |^PT   |^INR |^APTT |^FIB  |^D-D  "
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
      Left            =   7290
      TabIndex        =   20
      Top             =   990
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmCoagTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim sn As Recordset
          Dim sql As String
21870     On Error GoTo FillG_Error

21880     ReDim Count(1 To 5) As Long
21890     ReDim totals(1 To 5) As Long
          Dim t As Integer
          Dim n As Integer
          Dim s As String
          Dim p As Integer
          Dim CodePT As String
          Dim CodeINR As String
          Dim CodeAPTT As String
          Dim CodeFIB As String
          Dim CodeDD As String

21900     g.Rows = 2
21910     g.AddItem ""
21920     g.RemoveItem 1

21930     CodePT = "041"
21940     CodeINR = "044"
21950     CodeAPTT = "051"
21960     CodeFIB = "062"
21970     CodeDD = "612"

21980     If UCase$(HospName(0)) = "CAVAN" Then
21990         CodeDD = "352"
22000     End If

22010     For n = 0 To List1.ListCount - 1
22020         List1.Selected(n) = True
22030         For p = 1 To 5
22040             sql = "select count (R.Code) as Tot " & _
                      "from Coagresults as R, Demographics as D " & _
                      "where (R.Code = '" & _
                      Choose(p, CodePT, CodeINR, CodeAPTT, CodeFIB, CodeDD) & "') " & _
                      "and D.SampleID = R.SampleID " & _
                      "and (D.RunDate between '" & _
                      Format(dtFrom, "dd/mmm/yyyy") & "' and '" & _
                      Format(dtTo, "dd/mmm/yyyy") & "' )" & _
                      "and "
22050             If oSource(0) Then
22060                 sql = sql & "D.Clinician"
22070             ElseIf oSource(1) Then
22080                 sql = sql & "D.Ward"
22090             Else
22100                 sql = sql & "D.GP"
22110             End If
22120             sql = sql & " = '" & AddTicks(List1.List(n)) & "'"
22130             Set sn = New Recordset
22140             RecOpenServer 0, sn, sql
22150             Count(p) = sn!Tot
22160         Next
22170         If (Count(1) + Count(2) + Count(3) + Count(4) + Count(5)) <> 0 Then
22180             s = List1.List(n) & vbTab & _
                      Count(1) & vbTab & _
                      Count(2) & vbTab & _
                      Count(3) & vbTab & _
                      Count(4) & vbTab & _
                      Count(5)
22190             g.AddItem s
22200             For t = 1 To 5
22210                 totals(t) = totals(t) + Count(t)
22220             Next
22230         End If
22240     Next
22250     g.AddItem ""
22260     s = "Totals" & vbTab & _
              totals(1) & vbTab & _
              totals(2) & vbTab & _
              totals(3) & vbTab & _
              totals(4) & vbTab & _
              totals(5)
22270     g.AddItem s

22280     If g.Rows > 2 Then
22290         g.RemoveItem 1
22300     End If

22310     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

22320     intEL = Erl
22330     strES = Err.Description
22340     LogError "frmCoagTotals", "FillG", intEL, strES, sql

End Sub


Private Sub FillList()

          Dim tb As Recordset
          Dim Source As String
          Dim sql As String

22350     On Error GoTo FillList_Error

22360     g.Rows = 2
22370     g.AddItem ""
22380     g.RemoveItem 1

22390     If oSource(0) Then
22400         Source = "Clinician"
22410     ElseIf oSource(1) Then
22420         Source = "Ward"
22430     Else
22440         Source = "GP"
22450     End If

22460     List1.Clear

22470     sql = "Select distinct " & Source & " as Source from demographics where " & _
              "(RunDate between '" & _
              Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
              Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' ) "
22480     Set tb = New Recordset
22490     RecOpenServer 0, tb, sql
22500     Do While Not tb.EOF
22510         List1.AddItem tb!Source & ""
22520         tb.MoveNext
22530     Loop

22540     Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

22550     intEL = Erl
22560     strES = Err.Description
22570     LogError "frmCoagTotals", "FillList", intEL, strES, sql


End Sub


Private Sub breCalc_Click()

22580     FillList
22590     FillG

End Sub

Private Sub cmdCancel_Click()

22600     Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim n As Integer
          Dim X As Integer

22610     Printer.Font.Name = "Courier New"

22620     Printer.Print "Totals: "; dtFrom; " to "; dtTo

22630     Printer.Print
22640     For n = 0 To g.Rows - 1
22650         Printer.Print Left$(g.TextMatrix(n, 0), 24); Tab(25);
22660         For X = 1 To 5
22670             Printer.Print Left(g.TextMatrix(n, X) & "      ", 6);
22680         Next
22690         Printer.Print
22700     Next
22710     Printer.EndDoc

End Sub

Private Sub cmdXL_Click()

22720     ExportFlexGrid g, Me

End Sub

Private Sub dtFrom_CloseUp()

22730     FillList

End Sub

Private Sub dtTo_CloseUp()

22740     FillList

End Sub

Private Sub Form_Activate()

22750     FillList

End Sub

Private Sub Form_Load()

22760     dtFrom = Format(Now, "dd/mmm/yyyy")
22770     dtTo = dtFrom

End Sub

Private Sub obetween_Click(Index As Integer)

          Dim UpTo As String

22780     dtFrom = BetweenDates(Index, UpTo)
22790     dtTo = UpTo

22800     FillList

End Sub


Private Sub oSource_Click(Index As Integer)

22810     FillList

End Sub


