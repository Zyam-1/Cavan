VERSION 5.00
Begin VB.Form frmPanelSequence 
   Caption         =   "NetAcquire"
   ClientHeight    =   5535
   ClientLeft      =   600
   ClientTop       =   705
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   4185
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2850
      Top             =   1380
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2850
      Top             =   720
   End
   Begin VB.CommandButton cmdMoveUp 
      Height          =   555
      Left            =   2400
      Picture         =   "frmPanelSequence.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveDown 
      Height          =   525
      Left            =   2400
      Picture         =   "frmPanelSequence.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1290
      Width           =   465
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   2670
      Picture         =   "frmPanelSequence.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3510
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Save"
      Height          =   675
      Left            =   2670
      Picture         =   "frmPanelSequence.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ListBox lstPanel 
      Height          =   4545
      Left            =   240
      TabIndex        =   1
      Top             =   690
      Width           =   2145
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "cmbSampleType"
      Top             =   240
      Width           =   2145
   End
End
Attribute VB_Name = "frmPanelSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer

Private Sub FireDown()

      Dim s As String
      Dim Y As Integer

22720 If lstPanel.Selected(lstPanel.ListCount - 1) Then Exit Sub

22730 FireCounter = FireCounter + 1
22740 If FireCounter > 5 Then
22750   tmrDown.Interval = 100
22760 End If

22770 For Y = 0 To lstPanel.ListCount - 2
22780   If lstPanel.Selected(Y) Then
22790     s = lstPanel.List(Y)
22800     lstPanel.RemoveItem Y
22810     lstPanel.AddItem s, Y + 1
22820     lstPanel.Selected(Y + 1) = True
22830     Exit For
22840   End If
22850 Next

22860 cmdSave.Visible = True

End Sub

Private Sub FireUp()

      Dim s As String
      Dim Y As Integer

22870 FireCounter = FireCounter + 1
22880 If FireCounter > 5 Then
22890   tmrUp.Interval = 100
22900 End If

22910 If lstPanel.Selected(0) Then Exit Sub

22920 For Y = 1 To lstPanel.ListCount - 1
22930   If lstPanel.Selected(Y) Then
22940     s = lstPanel.List(Y)
22950     lstPanel.RemoveItem Y
22960     lstPanel.AddItem s, Y - 1
22970     lstPanel.Selected(Y - 1) = True
22980     Exit For
22990   End If
23000 Next

23010 cmdSave.Visible = True

End Sub



Private Sub FillG()

      Dim PanelType As String
      Dim tb As Recordset
      Dim sql As String

23020 On Error GoTo FillG_Error

23030 PanelType = ListCodeFor("ST", cmbSampleType)

23040 lstPanel.Clear
          
23050 sql = "Select distinct PanelName, Listorder from Panels where " & _
            "PanelType = '" & PanelType & "' " & _
            "and Hospital = '" & HospName(0) & "' " & _
            "order by ListOrder"
23060 Set tb = New Recordset
23070 RecOpenServer 0, tb, sql
23080 Do While Not tb.EOF
23090   lstPanel.AddItem tb!PanelName & ""
23100   tb.MoveNext
23110 Loop

23120 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

23130 intEL = Erl
23140 strES = Err.Description
23150 LogError "frmPanelSequence", "FillG", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

23160 Unload Me

End Sub


Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

23170 FireDown

23180 tmrDown.Interval = 250
23190 FireCounter = 0

23200 tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

23210 tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

23220 FireUp

23230 tmrUp.Interval = 250
23240 FireCounter = 0

23250 tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

23260 tmrUp.Enabled = False

End Sub


Private Sub cmdSave_Click()

      Dim PanelType As String
      Dim sql As String
      Dim Y As Integer

23270 On Error GoTo cmdSave_Click_Error

23280 PanelType = ListCodeFor("ST", cmbSampleType)

23290 For Y = 0 To lstPanel.ListCount - 1
23300   sql = "Update Panels " & _
              "Set ListOrder = '" & Y & "' " & _
              "where PanelName = '" & lstPanel.List(Y) & "' " & _
              "and PanelType = '" & PanelType & "' " & _
              "and Hospital = '" & HospName(0) & "'"
23310   Cnxn(0).Execute sql
23320 Next

23330 cmdSave.Visible = False

23340 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

23350 intEL = Erl
23360 strES = Err.Description
23370 LogError "frmPanelSequence", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

23380 On Error GoTo Form_Load_Error

23390 sql = "Select * from Lists where " & _
            "ListType = 'ST' and InUse = 1 " & _
            "Order by ListOrder"
23400 Set tb = New Recordset
23410 RecOpenServer 0, tb, sql

23420 cmbSampleType.Clear

23430 Do While Not tb.EOF
23440   cmbSampleType.AddItem tb!Text & ""
23450   tb.MoveNext
23460 Loop
23470 If cmbSampleType.ListCount > 0 Then
23480   cmbSampleType.ListIndex = 0
23490 End If

23500 FillG

23510 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

23520 intEL = Erl
23530 strES = Err.Description
23540 LogError "frmPanelSequence", "Form_Load", intEL, strES, sql


End Sub


Private Sub tmrDown_Timer()

23550 FireDown

End Sub


Private Sub tmrUp_Timer()

23560 FireUp

End Sub


