VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOption 
   Caption         =   "NetAcquire 6.9 - System Options"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   6990
      Top             =   6060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Options"
      Height          =   1725
      Left            =   45
      TabIndex        =   3
      Top             =   6660
      Width           =   8700
      Begin VB.OptionButton optSys 
         Caption         =   "Users"
         Height          =   285
         Index           =   1
         Left            =   4410
         TabIndex        =   10
         Top             =   990
         Width           =   1680
      End
      Begin VB.OptionButton optSys 
         Caption         =   "System"
         Height          =   285
         Index           =   0
         Left            =   4410
         TabIndex        =   9
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   420
         Left            =   6435
         TabIndex        =   8
         Top             =   675
         Width           =   1905
      End
      Begin VB.TextBox txtContent 
         Height          =   375
         Left            =   1575
         TabIndex        =   7
         Top             =   945
         Width           =   2490
      End
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   1575
         TabIndex        =   4
         Top             =   495
         Width           =   2490
      End
      Begin VB.Label Label2 
         Caption         =   "Content"
         Height          =   240
         Left            =   270
         TabIndex        =   6
         Top             =   1035
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   240
         Left            =   270
         TabIndex        =   5
         Top             =   540
         Width           =   1095
      End
   End
   Begin Threed.SSCommand cmdCncel 
      Cancel          =   -1  'True
      Height          =   555
      Left            =   4590
      TabIndex        =   2
      Top             =   5985
      Width           =   1635
      _Version        =   65536
      _ExtentX        =   2884
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Cancel"
   End
   Begin Threed.SSCommand cmdUpdate 
      Height          =   555
      Left            =   2520
      TabIndex        =   1
      Top             =   5985
      Width           =   1770
      _Version        =   65536
      _ExtentX        =   3122
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Update"
      Enabled         =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdOpt 
      Height          =   5505
      Left            =   360
      TabIndex        =   0
      Top             =   270
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   9710
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Description                               |<Contents                                                                  "
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
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdAdd_Click()

      Dim sql As String
      Dim tb As New Recordset
      Dim sn As New Recordset

16740 On Error GoTo cmdAdd_Click_Error

16750 If optSys(1) Then
16760   sql = "Select name from users"
16770   Set tb = New Recordset
16780   RecOpenServer 0, tb, sql
16790   Do While Not tb.EOF
16800     sql = "select * from options where " & _
                "description = '" & txtDescription & "' " & _
                "and username = '" & AddTicks(tb!Name) & "'"
16810     Set sn = New Recordset
16820     RecOpenServer 0, sn, sql
16830     If sn.EOF Then sn.AddNew
16840       sn!Description = txtDescription
16850       sn!Contents = txtContent
16860       sn!UserName = tb!Name
16870       sn.Update
16880     tb.MoveNext
16890   Loop
16900 Else
16910     sql = "select * from options where " & _
                "description = '" & txtDescription & "' "
16920     Set sn = New Recordset
16930     RecOpenServer 0, sn, sql
16940     If sn.EOF Then sn.AddNew
16950       sn!Description = txtDescription
16960       sn!Contents = txtContent
16970       sn.Update
16980 End If

16990 txtDescription = ""
17000 txtContent = ""

17010 Load_Options

17020 Exit Sub

cmdAdd_Click_Error:

      Dim er As Long
      Dim ers As String

17030 er = Err.Number
17040 ers = Err.Description

      'LogError "frmOption/cmdAdd_Click:" & Format$(er) & ":" & ers

End Sub

Private Sub cmdCncel_Click()

17050 Unload Me

End Sub

Private Sub cmdUpdate_Click()
      Dim tb As New Recordset
      Dim sql As String
      Dim n As Integer


17060 On Error GoTo cmdUpdate_Click_Error

17070 For n = 1 To grdOpt.Rows - 1
17080   sql = "Select * from options where description = '" & grdOpt.TextMatrix(n, 0) & "'"
17090   Set tb = New Recordset
17100   RecOpenServer 0, tb, sql
17110   tb!Description = grdOpt.TextMatrix(n, 0)
17120   If grdOpt.TextMatrix(n, 1) = "True" Then
17130     tb!Contents = 1
17140   ElseIf grdOpt.TextMatrix(n, 1) = "False" Then
17150     tb!Contents = 0
17160   Else
17170     tb!Contents = grdOpt.TextMatrix(n, 1)
17180   End If
17190   tb.Update
17200 Next

17210 LoadOptions

17220 sql = "Insert into updates  (upd, dtime) values ('Option', '" & Format(Now, "dd/MMM/yyyy hh:mm") & "')"
17230 Cnxn(0).Execute sql

17240 Exit Sub

cmdUpdate_Click_Error:

      Dim er As Long
      Dim ers As String

17250 er = Err.Number
17260 ers = Err.Description

      'LogError "frmOption/cmdUpdate_Click:" & Format$(er) & ers

End Sub

Private Sub Form_DblClick()

      Dim c As Long

17270 cD.ShowColor
17280 c = cD.Color

End Sub

Private Sub Form_Load()

17290 Load_Options

End Sub

Private Sub grdOpt_Click()

      Dim s As String

17300 On Error GoTo grdOpt_Click_Error

17310 If grdOpt.Col = 1 Then
17320   If grdOpt.TextMatrix(grdOpt.row, 1) = "True" Then
17330     grdOpt.TextMatrix(grdOpt.row, 1) = "False"
17340   ElseIf grdOpt.TextMatrix(grdOpt.row, 1) = "False" Then
17350     grdOpt.TextMatrix(grdOpt.row, 1) = "True"
17360   Else
17370     s = iBOX("Change", , grdOpt.TextMatrix(grdOpt.row, 1), False)
17380     If s <> "" Then grdOpt.TextMatrix(grdOpt.row, 1) = s
17390   End If
17400 End If

17410 cmdUpdate.Enabled = True

17420 Exit Sub

grdOpt_Click_Error:

      Dim er As Long
      Dim ers As String

17430 er = Err.Number
17440 ers = Err.Description

      'LogError "frmOption/grdOpt_Click:" & Format$(er) & ers

End Sub

Private Sub Load_Options()

      Dim tb As New Recordset
      Dim sql As String
      Dim s As String

17450 On Error GoTo Load_Options_Error

17460 grdOpt.Rows = 2
17470 grdOpt.AddItem ""
17480 grdOpt.RemoveItem 1

17490 sql = "Select * from Options where " & _
            "UserName = '' " & _
            "or UserName is null " & _
            "order by Description"
17500 Set tb = New Recordset
17510 RecOpenServer 0, tb, sql

17520 Do While Not tb.EOF
17530   s = UCase$(Trim$(tb!Description & "")) & vbTab
17540   If Trim$(tb!Contents & "") = "1" Then
17550     s = s & "True" & vbTab
17560   ElseIf Trim$(tb!Contents & "") = "0" Then
17570       s = s & "False" & vbTab
17580   Else
17590       s = s & Trim$(tb!Contents & "") & vbTab
17600   End If
17610   grdOpt.AddItem s
17620   tb.MoveNext
17630 Loop

17640 If grdOpt.Rows > 2 Then
17650   grdOpt.RemoveItem 1
17660 End If

17670 Exit Sub

Load_Options_Error:

      Dim strES As String
      Dim intEL As Integer

17680 intEL = Erl
17690 strES = Err.Description
17700 LogError "frmOption", "Load_Options", intEL, strES, sql


End Sub

