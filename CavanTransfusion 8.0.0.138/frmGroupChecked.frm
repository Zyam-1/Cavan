VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGroupChecked 
   Caption         =   "NetAcquire"
   ClientHeight    =   6495
   ClientLeft      =   945
   ClientTop       =   1050
   ClientWidth     =   5685
   DrawWidth       =   10
   Icon            =   "frmGroupChecked.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   5685
   Begin VB.CheckBox chkXM 
      Caption         =   "Include XMatched"
      Height          =   255
      Left            =   3690
      TabIndex        =   3
      Top             =   330
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   4170
      TabIndex        =   2
      Top             =   2820
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Columns         =   2
      Height          =   5130
      Left            =   390
      TabIndex        =   0
      Top             =   900
      Width           =   3585
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   30
      TabIndex        =   4
      Top             =   6270
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The following Units are within Date and have been Group Checked"
      Height          =   525
      Left            =   390
      TabIndex        =   1
      Top             =   210
      Width           =   2595
   End
End
Attribute VB_Name = "frmGroupChecked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillList()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillList_Error

20    List1.Clear

30    sql = "Select ISBT128 from Latest where " & _
            "DateExpiry >= '" & Format$(Now, "dd/mmm/yyyy hh:mm") & "' " & _
            "and Checked = 1 " & _
            "and (Event = 'R' " & _
            "     or Event = 'C' "
40    If chkXM Then
50      sql = sql & "or Event = 'X' "
60    End If
70    sql = sql & ") Order by ISBT128"

80    Set tb = New Recordset
90    RecOpenServerBB 0, tb, sql
100   If Not tb.EOF Then
110     Do While Not tb.EOF
120       List1.AddItem tb!ISBT128 & ""
130       tb.MoveNext
140     Loop
150   Else
160     List1.AddItem "None Found"
170   End If

180   Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmGroupChecked", "FillList", intEL, strES, sql


End Sub

Private Sub chkXM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillList

End Sub


Private Sub cmdOK_Click()

10    Unload Me

End Sub


Private Sub Form_Load()
10    FillList
End Sub
