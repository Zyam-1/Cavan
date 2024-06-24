VERSION 5.00
Begin VB.Form frmOCPanelSequence 
   Caption         =   "NetAcquire - Order Comms Panel Sequence"
   ClientHeight    =   5565
   ClientLeft      =   360
   ClientTop       =   630
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   4890
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   660
      TabIndex        =   5
      Text            =   "cmbSampleType"
      Top             =   270
      Width           =   2145
   End
   Begin VB.ListBox lstPanel 
      Height          =   4545
      Left            =   660
      TabIndex        =   4
      Top             =   720
      Width           =   2145
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Save"
      Height          =   675
      Left            =   3090
      Picture         =   "frmOCPanelSequence.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2430
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   3090
      Picture         =   "frmOCPanelSequence.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3540
      Width           =   1185
   End
   Begin VB.CommandButton cmdMoveDown 
      Height          =   525
      Left            =   2820
      Picture         =   "frmOCPanelSequence.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveUp 
      Height          =   555
      Left            =   2820
      Picture         =   "frmOCPanelSequence.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   750
      Width           =   465
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3270
      Top             =   750
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3270
      Top             =   1410
   End
End
Attribute VB_Name = "frmOCPanelSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer

Private Sub FillG()

      Dim PanelType As String
      Dim tb As Recordset
      Dim sql As String

15890 On Error GoTo FillG_Error

15900 PanelType = ListCodeFor("ST", cmbSampleType)

15910 lstPanel.Clear
          
15920 sql = "Select distinct PanelName, Listorder from OCOrderContents where " & _
            "PanelType = '" & PanelType & "' " & _
            "order by ListOrder"
15930 Set tb = New Recordset
15940 RecOpenServer 0, tb, sql
15950 Do While Not tb.EOF
15960   lstPanel.AddItem tb!PanelName & ""
15970   tb.MoveNext
15980 Loop

15990 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

16000 intEL = Erl
16010 strES = Err.Description
16020 LogError "frmOCPanelSequence", "FillG", intEL, strES, sql


End Sub
Private Sub FireDown()

      Dim s As String
      Dim Y As Integer

16030 If lstPanel.Selected(lstPanel.ListCount - 1) Then Exit Sub

16040 FireCounter = FireCounter + 1
16050 If FireCounter > 5 Then
16060   tmrDown.Interval = 100
16070 End If

16080 For Y = 0 To lstPanel.ListCount - 2
16090   If lstPanel.Selected(Y) Then
16100     s = lstPanel.List(Y)
16110     lstPanel.RemoveItem Y
16120     lstPanel.AddItem s, Y + 1
16130     lstPanel.Selected(Y + 1) = True
16140     Exit For
16150   End If
16160 Next

16170 cmdSave.Visible = True

End Sub

Private Sub FireUp()

      Dim s As String
      Dim Y As Integer

16180 FireCounter = FireCounter + 1
16190 If FireCounter > 5 Then
16200   tmrUp.Interval = 100
16210 End If

16220 If lstPanel.Selected(0) Then Exit Sub

16230 For Y = 1 To lstPanel.ListCount - 1
16240   If lstPanel.Selected(Y) Then
16250     s = lstPanel.List(Y)
16260     lstPanel.RemoveItem Y
16270     lstPanel.AddItem s, Y - 1
16280     lstPanel.Selected(Y - 1) = True
16290     Exit For
16300   End If
16310 Next

16320 cmdSave.Visible = True

End Sub

Private Sub cmdCancel_Click()

16330 Unload Me

End Sub


Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

16340 FireDown

16350 tmrDown.Interval = 250
16360 FireCounter = 0

16370 tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16380 tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

16390 FireUp

16400 tmrUp.Interval = 250
16410 FireCounter = 0

16420 tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

16430 tmrUp.Enabled = False

End Sub


Private Sub cmdSave_Click()

      Dim PanelType As String
      Dim sql As String
      Dim Y As Integer

16440 On Error GoTo cmdSave_Click_Error

16450 PanelType = ListCodeFor("ST", cmbSampleType)

16460 For Y = 0 To lstPanel.ListCount - 1
16470   sql = "Update OCOrderContents " & _
              "Set ListOrder = '" & Y & "' " & _
              "where PanelName = '" & lstPanel.List(Y) & "' " & _
              "and PanelType = '" & PanelType & "'"
16480   Cnxn(0).Execute sql
16490 Next

16500 cmdSave.Visible = False

16510 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

16520 intEL = Erl
16530 strES = Err.Description
16540 LogError "frmOCPanelSequence", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

16550 On Error GoTo Form_Load_Error

16560 sql = "Select * from Lists where " & _
            "ListType = 'ST' and InUse = 1 " & _
            "Order by ListOrder"
16570 Set tb = New Recordset
16580 RecOpenServer 0, tb, sql

16590 cmbSampleType.Clear

16600 Do While Not tb.EOF
16610   cmbSampleType.AddItem tb!Text & ""
16620   tb.MoveNext
16630 Loop
16640 If cmbSampleType.ListCount > 0 Then
16650   cmbSampleType.ListIndex = 0
16660 End If

16670 FillG

16680 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

16690 intEL = Erl
16700 strES = Err.Description
16710 LogError "frmOCPanelSequence", "Form_Load", intEL, strES, sql


End Sub


Private Sub tmrDown_Timer()

16720 FireDown

End Sub


Private Sub tmrUp_Timer()

16730 FireUp

End Sub


