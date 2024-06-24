VERSION 5.00
Begin VB.Form frmCoagPanelSequence 
   Caption         =   "NetAcquire"
   ClientHeight    =   4980
   ClientLeft      =   750
   ClientTop       =   735
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   4065
   Begin VB.ListBox lstPanel 
      Height          =   4545
      Left            =   120
      TabIndex        =   4
      Top             =   210
      Width           =   2145
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Save"
      Height          =   675
      Left            =   2550
      Picture         =   "frmCoagPanelSequence.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   2550
      Picture         =   "frmCoagPanelSequence.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3030
      Width           =   1185
   End
   Begin VB.CommandButton cmdMoveDown 
      Height          =   525
      Left            =   2280
      Picture         =   "frmCoagPanelSequence.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   810
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveUp 
      Height          =   555
      Left            =   2280
      Picture         =   "frmCoagPanelSequence.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   465
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2730
      Top             =   240
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2730
      Top             =   900
   End
End
Attribute VB_Name = "frmCoagPanelSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer
Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String

20480     On Error GoTo FillG_Error

20490     lstPanel.Clear
          
20500     sql = "Select distinct PanelName, ListOrder from CoagPanels " & _
              "order by ListOrder"
20510     Set tb = New Recordset
20520     RecOpenServer 0, tb, sql
20530     Do While Not tb.EOF
20540         lstPanel.AddItem tb!PanelName & ""
20550         tb.MoveNext
20560     Loop

20570     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

20580     intEL = Erl
20590     strES = Err.Description
20600     LogError "frmCoagPanelSequence", "FillG", intEL, strES, sql


End Sub

Private Sub FireDown()

          Dim s As String
          Dim Y As Integer

20610     If lstPanel.Selected(lstPanel.ListCount - 1) Then Exit Sub

20620     FireCounter = FireCounter + 1
20630     If FireCounter > 5 Then
20640         tmrDown.Interval = 100
20650     End If

20660     For Y = 0 To lstPanel.ListCount - 2
20670         If lstPanel.Selected(Y) Then
20680             s = lstPanel.List(Y)
20690             lstPanel.RemoveItem Y
20700             lstPanel.AddItem s, Y + 1
20710             lstPanel.Selected(Y + 1) = True
20720             Exit For
20730         End If
20740     Next

20750     cmdSave.Visible = True

End Sub

Private Sub FireUp()

          Dim s As String
          Dim Y As Integer

20760     FireCounter = FireCounter + 1
20770     If FireCounter > 5 Then
20780         tmrUp.Interval = 100
20790     End If

20800     If lstPanel.Selected(0) Then Exit Sub

20810     For Y = 1 To lstPanel.ListCount - 1
20820         If lstPanel.Selected(Y) Then
20830             s = lstPanel.List(Y)
20840             lstPanel.RemoveItem Y
20850             lstPanel.AddItem s, Y - 1
20860             lstPanel.Selected(Y - 1) = True
20870             Exit For
20880         End If
20890     Next

20900     cmdSave.Visible = True

End Sub

Private Sub cmdCancel_Click()

20910     Unload Me

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20920     FireDown

20930     tmrDown.Interval = 250
20940     FireCounter = 0

20950     tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

20960     tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20970     FireUp

20980     tmrUp.Interval = 250
20990     FireCounter = 0

21000     tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

21010     tmrUp.Enabled = False

End Sub


Private Sub cmdSave_Click()

          Dim sql As String
          Dim Y As Integer

21020     On Error GoTo cmdSave_Click_Error

21030     For Y = 0 To lstPanel.ListCount - 1
21040         sql = "Update CoagPanels " & _
                  "Set ListOrder = '" & Y & "' " & _
                  "where PanelName = '" & lstPanel.List(Y) & "'"
21050         Cnxn(0).Execute sql
21060     Next

21070     cmdSave.Visible = False

21080     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

21090     intEL = Erl
21100     strES = Err.Description
21110     LogError "frmCoagPanelSequence", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

21120     FillG

End Sub


Private Sub tmrDown_Timer()

21130     FireDown

End Sub


Private Sub tmrUp_Timer()

21140     FireUp

End Sub


