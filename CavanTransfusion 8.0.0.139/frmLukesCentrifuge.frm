VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLukesCentrifuge 
   Caption         =   "NetAcquire --- Centrifuge Speeds / Temperatures"
   ClientHeight    =   4080
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   6945
   Icon            =   "frmLukesCentrifuge.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6945
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1050
      TabIndex        =   6
      Top             =   2400
      Width           =   5655
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Previous"
      Height          =   765
      Left            =   1560
      Picture         =   "frmLukesCentrifuge.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2850
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   4320
      Picture         =   "frmLukesCentrifuge.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2850
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   765
      Left            =   2940
      Picture         =   "frmLukesCentrifuge.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2850
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemperature 
      Height          =   1065
      Left            =   300
      TabIndex        =   1
      Top             =   1230
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   1879
      _Version        =   393216
      Rows            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmLukesCentrifuge.frx":1C08
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
   Begin MSFlexGridLib.MSFlexGrid grdSpeed 
      Height          =   825
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   1455
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmLukesCentrifuge.frx":1C5B
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   660
      TabIndex        =   8
      Top             =   3840
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   2460
      Width           =   660
   End
   Begin VB.Label lblLastEntered 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   4230
      TabIndex        =   5
      Top             =   1470
      Width           =   2475
   End
End
Attribute VB_Name = "frmLukesCentrifuge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdLoad_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdLoad_Click_Error

20    sql = "Select top 1 * from StLukesCentrifuge " & _
            "Order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      txtComment = tb!Comment & ""
70      grdSpeed.TextMatrix(1, 1) = tb!Cent1Phase1 & ""
80      grdSpeed.TextMatrix(1, 2) = tb!Cent1Phase2 & ""
90      grdSpeed.TextMatrix(2, 1) = tb!Cent2Phase1 & ""
100     grdSpeed.TextMatrix(2, 2) = tb!Cent2Phase2 & ""
110     grdTemperature.TextMatrix(1, 1) = tb!BlockL & ""
120     grdTemperature.TextMatrix(2, 1) = tb!BlockR & ""
130     grdTemperature.TextMatrix(3, 1) = tb!BlockS & ""
140   Else
150     grdSpeed.TextMatrix(1, 1) = ""
160     grdSpeed.TextMatrix(1, 2) = ""
170     grdSpeed.TextMatrix(2, 1) = ""
180     grdSpeed.TextMatrix(2, 2) = ""
190     grdTemperature.TextMatrix(1, 1) = ""
200     grdTemperature.TextMatrix(2, 1) = ""
210     grdTemperature.TextMatrix(3, 1) = ""
220   End If
  
230   cmdSave.Enabled = True

240   Exit Sub

cmdLoad_Click_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmLukesCentrifuge", "cmdLoad_Click", intEL, strES, sql


End Sub

Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    sql = "Select  * from StLukesCentrifuge where " & _
            "DateTime = '01/01/2000'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    tb.AddNew
60    tb!Comment = txtComment
70    tb!Cent1Phase1 = grdSpeed.TextMatrix(1, 1)
80    tb!Cent1Phase2 = grdSpeed.TextMatrix(1, 2)
90    tb!Cent2Phase1 = grdSpeed.TextMatrix(2, 1)
100   tb!Cent2Phase2 = grdSpeed.TextMatrix(2, 2)
110   tb!BlockL = grdTemperature.TextMatrix(1, 1)
120   tb!BlockR = grdTemperature.TextMatrix(2, 1)
130   tb!BlockS = grdTemperature.TextMatrix(3, 1)
140   tb!DateTime = Format(Now, "dd/mm/yyyy hh:mm:ss")
150   tb!Operator = UserName
160   tb.Update

170   cmdSave.Enabled = False
180   Unload Me

190   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmLukesCentrifuge", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo Form_Load_Error

20    sql = "Select top 1 * from StLukesCentrifuge " & _
            "Order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      s = "Last Entered by " & tb!Operator & vbCrLf & _
            "on " & Format$(tb!DateTime, "dd/mm/yyyy") & _
            " at " & Format$(tb!DateTime, "hh:mm:ss")
70      lblLastEntered = s
80    End If

90    Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmLukesCentrifuge", "Form_Load", intEL, strES, sql


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Enabled Then
30      Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70      End If
80    End If

End Sub


Private Sub grdSpeed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim strIP As String
      Dim s As String

10    With grdSpeed
20      If .MouseRow <> 0 And .MouseCol <> 0 Then
30        s = "Enter Speed for " & .TextMatrix(0, .Col) & " " & _
              "for Centrifuge " & .TextMatrix(.Row, 0)
40        strIP = iBOX(s, , .TextMatrix(.Row, .Col))
50        If TimedOut Then Unload Me: Exit Sub
60        .TextMatrix(.Row, .Col) = strIP
70        cmdSave.Enabled = True
80      End If
90    End With

End Sub


Private Sub grdTemperature_Click()

      Dim strIP As String
      Dim s As String

10    With grdTemperature
20      If .MouseRow <> 0 And .MouseCol <> 0 Then
30        s = "Enter Temperature for " & .TextMatrix(.Row, 0)
40        strIP = iBOX(s, , .TextMatrix(.Row, 1))
50        If TimedOut Then Unload Me: Exit Sub
60        .TextMatrix(.Row, 1) = strIP
70        cmdSave.Enabled = True
80      End If
90    End With

End Sub


