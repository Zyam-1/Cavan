VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMonthlyWastage 
   Caption         =   "NetAcquire"
   ClientHeight    =   4890
   ClientLeft      =   990
   ClientTop       =   1410
   ClientWidth     =   8355
   Icon            =   "frmMonthlyWastage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   8355
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   2910
      Picture         =   "frmMonthlyWastage.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3510
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2295
      Left            =   210
      TabIndex        =   7
      Top             =   990
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   "^Group        |^Received into Stock |^Wasted (Expired)    |^Wasted (Other)   "
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   5730
      Picture         =   "frmMonthlyWastage.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3510
      Width           =   1245
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "&Start"
      Height          =   705
      Left            =   5490
      Picture         =   "frmMonthlyWastage.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   1500
      Picture         =   "frmMonthlyWastage.frx":1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3510
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   345
      Left            =   2370
      TabIndex        =   3
      Top             =   210
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   37509
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   345
      Left            =   4050
      TabIndex        =   4
      Top             =   210
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   37509
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   1440
      TabIndex        =   10
      Top             =   4590
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   4260
      TabIndex        =   9
      Top             =   3780
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "and"
      Height          =   195
      Left            =   3720
      TabIndex        =   6
      Top             =   300
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Between"
      Height          =   195
      Left            =   1710
      TabIndex        =   5
      Top             =   270
      Width           =   630
   End
End
Attribute VB_Name = "frmMonthlyWastage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillReceived()

      Dim tb As Recordset
      Dim sql As String
      Dim grp As Integer

10    On Error GoTo FillReceived_Error

20    For grp = 1 To 8
30      sql = "Select count (*) as Tot from Product where " & _
              "DateTime between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
              "and Event = 'C' " & _
              "and GroupRh = '" & Group2Bar(g.TextMatrix(grp, 0)) & "'"
40      Set tb = New Recordset
50      RecOpenServerBB 0, tb, sql
60      If tb!Tot > 0 Then
70        g.TextMatrix(grp, 1) = tb!Tot
80      Else
90        g.TextMatrix(grp, 1) = ""
100     End If
110   Next

120   Exit Sub

FillReceived_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmMonthlyWastage", "FillReceived", intEL, strES, sql


End Sub

Private Sub FillWasted()

      Dim tb As Recordset
      Dim sql As String
      Dim grp As Integer
      Dim Expired As Integer
      Dim Other As Integer

10    On Error GoTo FillWasted_Error

20        For grp = 1 To 8
30          Expired = 0
40          Other = 0
50          sql = "Select * from Latest where " & _
                  "DateTime between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
                  "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                  "and Event = 'D' " & _
                  "and GroupRh = '" & Group2Bar(g.TextMatrix(grp, 0)) & "'"
60          Set tb = New Recordset
70          RecOpenServerBB 0, tb, sql
80          Do While Not tb.EOF
90            If Trim$(tb!Reason & "") <> "" Then
100             Debug.Print tb!Reason
110             If InStr(UCase$(tb!Reason), "EXPIRED") Or Left$(UCase$(tb!Reason), 3) = "EXP" Then
120               Expired = Expired + 1
130             Else
140               Other = Other + 1
150             End If
160           Else
170             Other = Other + 1
180           End If
190           tb.MoveNext
200         Loop
210         If Expired > 0 Then
220           g.TextMatrix(grp, 2) = Expired
230         Else
240           g.TextMatrix(grp, 2) = ""
250         End If
260         If Other > 0 Then
270           g.TextMatrix(grp, 3) = Other
280         Else
290           g.TextMatrix(grp, 3) = ""
300         End If
310       Next
320   Exit Sub

FillWasted_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmMonthlyWastage", "FillWasted", intEL, strES, sql

End Sub

Private Sub bprint_Click()

      Dim n As Integer

10    If Not SetFormPrinter() Then
20      iMsg "Can't find printer!", vbExclamation
30      If TimedOut Then Unload Me: Exit Sub
40      Exit Sub
50    End If

60        Printer.Print "Product Wastage"
70        Printer.Print "Between " & dtFrom & " and " & dtTo & "."
    
80        For n = 0 To g.Rows - 1
90          Printer.Print g.TextMatrix(n, 0);
100         Printer.Print Tab(12); g.TextMatrix(n, 1);
110         Printer.Print Tab(35); g.TextMatrix(n, 2);
120         Printer.Print Tab(53); g.TextMatrix(n, 3)
130       Next
140       Printer.EndDoc

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdStart_Click()

10    FillReceived
20    FillWasted

End Sub

Private Sub cmdXL_Click()

10    ExportFlexGrid g, Me

End Sub

Private Sub Form_Load()

      Dim n As Integer
10    InitGrid
20    dtFrom = Format(Now - 30, "dd/mm/yyyy")
30    dtTo = Format(Now, "dd/mm/yyyy")

40    For n = 1 To 8
50      g.AddItem Choose(n, "O Neg", "O Pos", "A Neg", "A Pos", _
                            "B Neg", "B Pos", "AB Neg", "AB Pos")
60    Next
70    g.RemoveItem 1

End Sub

Private Sub InitGrid()

10    With g
20        .Rows = 2: .Cols = 5
30        .AddItem ""
40        .RemoveItem 1
50        .ColWidth(0) = 1000: .TextMatrix(0, 0) = "Group": .ColAlignment(0) = flexAlignCenterCenter
60        .ColWidth(1) = 1900: .TextMatrix(0, 1) = "Received into Stock": .ColAlignment(1) = flexAlignLeftCenter
70        .ColWidth(2) = 1900: .TextMatrix(0, 2) = "Wasted (Expired)": .ColAlignment(2) = flexAlignLeftCenter
80        .ColWidth(3) = 2000: .TextMatrix(0, 3) = "Wasted (Other)": .ColAlignment(3) = flexAlignLeftCenter
90        .ColWidth(4) = 0
    
100   End With
  
End Sub

