VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetSources 
   Caption         =   "NetAcquire"
   ClientHeight    =   5445
   ClientLeft      =   240
   ClientTop       =   465
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   9270
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   7800
      Picture         =   "frmSetSources.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   1245
   End
   Begin VB.OptionButton oSource 
      Caption         =   "GPs"
      Height          =   255
      Index           =   2
      Left            =   7950
      TabIndex        =   7
      Top             =   810
      Width           =   1035
   End
   Begin VB.OptionButton oSource 
      Caption         =   "Clinicians"
      Height          =   255
      Index           =   1
      Left            =   7950
      TabIndex        =   6
      Top             =   510
      Width           =   1035
   End
   Begin VB.OptionButton oSource 
      Caption         =   "Wards"
      Height          =   255
      Index           =   0
      Left            =   7950
      TabIndex        =   5
      Top             =   210
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.ListBox lstSource 
      Columns         =   2
      DragIcon        =   "frmSetSources.frx":066A
      Height          =   5130
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   4815
   End
   Begin VB.CommandButton baddnew 
      Caption         =   "Add New Panel"
      Height          =   735
      Left            =   7800
      Picture         =   "frmSetSources.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1245
   End
   Begin VB.CommandButton bRemovePanel 
      Caption         =   "Remove Panel"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7800
      Picture         =   "frmSetSources.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3540
      Width           =   1245
   End
   Begin VB.CommandButton bRemoveItem 
      Caption         =   "Remove Item"
      Enabled         =   0   'False
      Height          =   585
      Left            =   7800
      Picture         =   "frmSetSources.frx":1780
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1245
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   5145
      Left            =   4950
      TabIndex        =   0
      Top             =   90
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   9075
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSetSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillList()

      Dim tb As Recordset
      Dim sql As String
      Dim strSource As String

57920 On Error GoTo FillList_Error

57930 lstSource.Clear

57940 If oSource(0) Then
57950   strSource = "Wards"
57960 ElseIf oSource(1) Then
57970   strSource = "Clinicians"
57980 ElseIf oSource(2) Then
57990   strSource = "GPs"
58000 End If

58010 sql = "Select * from " & strSource & " " & _
            "Order by ListOrder"
58020 Set tb = New Recordset
58030 RecOpenServer 0, tb, sql
58040 Do While Not tb.EOF
58050   lstSource.AddItem tb!Text & ""
58060   tb.MoveNext
58070 Loop

58080 Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

58090 intEL = Erl
58100 strES = Err.Description
58110 LogError "fSetSources", "FillList", intEL, strES, sql


End Sub

Private Sub FillTree()

      Dim NodX As MSComctlLib.Node
      Dim Key As Integer
      Dim SourcePanelType As String
      Dim PanelName As String
      Dim tb As Recordset
      Dim tbP As Recordset
      Dim sql As String
      Dim t As Single

58120 On Error GoTo FillTree_Error

58130 If oSource(0) Then
58140   SourcePanelType = "W"
58150 ElseIf oSource(1) Then
58160   SourcePanelType = "C"
58170 ElseIf oSource(2) Then
58180   SourcePanelType = "G"
58190 End If

58200 t = Timer

58210 Tree.Nodes.Clear
58220 sql = "Select SourcePanelName from SourcePanels where " & _
            "SourcePanelType = '" & SourcePanelType & "' " & _
            "Group By SourcePanelName " & _
            "Order by min(ListOrder)"
58230 Set tb = New Recordset
58240 RecOpenClient 0, tb, sql

58250 Do While Not tb.EOF
58260   PanelName = tb!SourcePanelName & ""
58270   Set NodX = Tree.Nodes.Add(, , "key" & CStr(Key), PanelName)
58280   Key = Key + 1
58290   sql = "Select * from SourcePanels where " & _
              "SourcePanelType = '" & SourcePanelType & "' " & _
              "and SourcePanelName = '" & PanelName & "' " & _
              "Order by ListOrder"
58300   Set tbP = New Recordset
58310   RecOpenClient 0, tbP, sql
58320   Do While Not tbP.EOF
58330     Set NodX = Tree.Nodes.Add("key" & CStr(Key - 1), tvwChild, , tbP!Content & "")
58340     tbP.MoveNext
58350   Loop
58360   tb.MoveNext
58370 Loop

58380 Debug.Print Timer - t

58390 Exit Sub

FillTree_Error:

      Dim strES As String
      Dim intEL As Integer

58400 intEL = Erl
58410 strES = Err.Description
58420 LogError "fSetSources", "FillTree", intEL, strES, sql

End Sub

Private Sub baddnew_Click()
        
      Dim NodX As MSComctlLib.Node
      Static k As String

58430 If k = "" Then
58440   k = "1"
58450 Else
58460   k = CStr(Val(k) + 1)
58470 End If

58480 Set NodX = Tree.Nodes.Add(, , "newkey" & k, "New Panel" & k)

End Sub


Private Sub bcancel_Click()

58490 Unload Me

End Sub


Private Sub bRemoveItem_Click()

      Dim SourcePanelType As String
      Dim SourcePanelName As String
      Dim Content As String
      Dim sql As String

58500 On Error GoTo bRemoveItem_Click_Error

58510 If oSource(0) Then
58520   SourcePanelType = "W"
58530 ElseIf oSource(1) Then
58540   SourcePanelType = "C"
58550 ElseIf oSource(2) Then
58560   SourcePanelType = "G"
58570 End If

58580 SourcePanelName = Tree.SelectedItem.Parent.Text
58590 Content = Tree.SelectedItem.Text

58600 sql = "Delete from SourcePanels where " & _
            "SourcePanelType = '" & SourcePanelType & "' " & _
            "and SourcePanelName = '" & SourcePanelName & "' " & _
            "and Content = '" & Content & "'"
58610 Cnxn(0).Execute sql

58620 FillTree

58630 Exit Sub

bRemoveItem_Click_Error:

      Dim strES As String
      Dim intEL As Integer

58640 intEL = Erl
58650 strES = Err.Description
58660 LogError "fSetSources", "bRemoveItem_Click", intEL, strES, sql


End Sub

Private Sub bRemovePanel_Click()

      Dim SourcePanelType As String
      Dim sql As String

58670 On Error GoTo bRemovePanel_Click_Error

58680 If oSource(0) Then
58690   SourcePanelType = "W"
58700 ElseIf oSource(1) Then
58710   SourcePanelType = "C"
58720 ElseIf oSource(2) Then
58730   SourcePanelType = "G"
58740 End If

58750 sql = "Delete from SourcePanels where " & _
            "SourcePanelType = '" & SourcePanelType & "' " & _
            "and SourcePanelName = '" & Tree.SelectedItem.Text & "'"
58760 Cnxn(0).Execute sql

58770 FillTree

58780 Exit Sub

bRemovePanel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

58790 intEL = Erl
58800 strES = Err.Description
58810 LogError "fSetSources", "bRemovePanel_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

58820 FillList
58830 FillTree

End Sub


Private Sub lstSource_Click()
                                                                                                '
End Sub

Private Sub lstSource_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

58840 lstSource_Click
58850 lstSource.Drag

End Sub


Private Sub oSource_Click(Index As Integer)

58860 FillList
58870 FillTree

End Sub


Private Sub Tree_AfterLabelEdit(Cancel As Integer, NewString As String)

      Dim sql As String
      Dim SourcePanelType As String

58880 On Error GoTo Tree_AfterLabelEdit_Error

58890 If oSource(0) Then
58900   SourcePanelType = "W"
58910 ElseIf oSource(1) Then
58920   SourcePanelType = "C"
58930 ElseIf oSource(2) Then
58940   SourcePanelType = "G"
58950 End If

58960 If Trim$(NewString) = "" Then
58970   Cancel = True
58980   Exit Sub
58990 End If

59000 sql = "UPDATE SourcePanels SET " & _
            "SourcePanelName = '" & NewString & "' " & _
            "WHERE " & _
            "SourcePanelName = '" & Tree.SelectedItem.Text & "' " & _
            "and SourcePanelType = '" & SourcePanelType & "'"
59010 Cnxn(0).Execute sql

59020 Exit Sub

Tree_AfterLabelEdit_Error:

      Dim strES As String
      Dim intEL As Integer

59030 intEL = Erl
59040 strES = Err.Description
59050 LogError "fSetSources", "Tree_AfterLabelEdit", intEL, strES, sql

End Sub

Private Sub Tree_DragDrop(Source As Control, X As Single, Y As Single)
        
      Dim NodX As MSComctlLib.Node
      Dim Key
      Dim SourcePanelType As String
      Dim sql As String

59060 On Error GoTo Tree_DragDrop_Error

59070 If oSource(0) Then
59080   SourcePanelType = "W"
59090 ElseIf oSource(1) Then
59100   SourcePanelType = "C"
59110 ElseIf oSource(2) Then
59120   SourcePanelType = "G"
59130 End If

59140 Set NodX = Tree.HitTest(X, Y)
59150 If NodX Is Nothing Then Exit Sub

59160 If Tree.DropHighlight Is Nothing Then
59170   Set Tree.DropHighlight = Nothing
59180   Exit Sub
59190 Else
59200   If NodX = Tree.DropHighlight Then
59210     Key = NodX.Key
59220     If Key <> "" Then
59230       Set NodX = Tree.Nodes.Add(Key, tvwChild, , Source.Text)
59240       sql = "Insert into SourcePanels " & _
                  "(SourcePanelName, SourcePanelType, Content, ListOrder) VALUES " & _
                  "('" & Tree.DropHighlight.Text & "', " & _
                  "'" & SourcePanelType & "', " & _
                  "'" & AddTicks(Source.Text) & "', " & _
                  "'999')"
59250       Cnxn(0).Execute sql
59260       Set Tree.DropHighlight = Nothing
59270       Tree.Nodes(Key).Child.EnsureVisible
59280     End If
59290   End If
59300 End If

59310 Exit Sub

Tree_DragDrop_Error:

      Dim strES As String
      Dim intEL As Integer

59320 intEL = Erl
59330 strES = Err.Description
59340 LogError "fSetSources", "Tree_DragDrop", intEL, strES, sql

End Sub


Private Sub Tree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
        
59350 Set Tree.DropHighlight = Tree.HitTest(X, Y)

End Sub


Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
        
59360 bRemovePanel.Enabled = False
59370 bRemoveItem.Enabled = False

End Sub


Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)

      Dim s As String

59380 s = Node.Key
59390 If s = "" Then
59400   bRemovePanel.Enabled = False
59410   bRemoveItem.Enabled = True
59420 Else
59430   bRemovePanel.Enabled = True
59440   bRemoveItem.Enabled = False
59450 End If


End Sub


