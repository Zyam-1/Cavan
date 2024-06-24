VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOCPanels 
   Caption         =   "NetAcquire - Define Order Comms Panels"
   ClientHeight    =   5775
   ClientLeft      =   300
   ClientTop       =   750
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   10080
   Begin MSComctlLib.ListView ListView1 
      DragIcon        =   "frmOCPanels.frx":0000
      Height          =   5115
      Left            =   210
      TabIndex        =   7
      Top             =   300
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9022
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   8340
      Picture         =   "frmOCPanels.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4860
      Width           =   1245
   End
   Begin VB.CommandButton baddnew 
      Caption         =   "Add New Panel"
      Height          =   525
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1245
   End
   Begin VB.CommandButton bRemovePanel 
      Caption         =   "Remove Panel"
      Enabled         =   0   'False
      Height          =   525
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton bRemoveItem 
      Caption         =   "Remove Item"
      Enabled         =   0   'False
      Height          =   525
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   7710
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   270
      Width           =   2145
   End
   Begin VB.CommandButton cmdAmendSequence 
      Caption         =   "Amend Panel Sequence"
      Height          =   345
      Left            =   4980
      TabIndex        =   0
      Top             =   5070
      Width           =   2715
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   4785
      Left            =   4980
      TabIndex        =   1
      Top             =   270
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   8440
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmOCPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FillList()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleType As String

14320 On Error GoTo FillList_Error

14330 ListView1.ListItems.Clear

14340 SampleType = ListCodeFor("ST", cmbSampleType)

14350 sql = "Select distinct Shortname, PrintPriority from BioTestDefinitions where " & _
            "SampleType = '" & SampleType & "' " & _
            "and Hospital = '" & HospName(0) & "' " & _
            "Order by PrintPriority"
14360 Set tb = New Recordset
14370 RecOpenServer 0, tb, sql
14380 Do While Not tb.EOF
14390   ListView1.ListItems.Add , , tb!ShortName & ""
14400   tb.MoveNext
14410 Loop

14420 Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

14430 intEL = Erl
14440 strES = Err.Description
14450 LogError "frmOCPanels", "FillList", intEL, strES, sql


End Sub

Private Sub FillTree()

      Dim NodX As MSComctlLib.Node
      Dim Key As Integer
      Dim PanelType As String
      Dim PanelName As String
      Dim tb As Recordset
      Dim tc As Recordset
      Dim sql As String

14460 On Error GoTo FillTree_Error

14470 PanelType = ListCodeFor("ST", cmbSampleType)

14480 Tree.Nodes.Clear
          
14490 sql = "Select Distinct O.PanelName, P.Sequence from OCOrderContents as O, OCOrderPanel as P where " & _
            "O.PanelType = '" & PanelType & "' " & _
            "and P.ShortName = O.PanelName " & _
            "order by P.Sequence"
14500 Set tb = New Recordset
14510 RecOpenServer 0, tb, sql
14520 Do While Not tb.EOF
14530   PanelName = tb!PanelName & ""
          
14540   Set NodX = Tree.Nodes.Add(, , "key" & CStr(Key), PanelName)
14550   Key = Key + 1
          
14560   sql = "Select Analyte from OCOrderContents where " & _
              "PanelType = '" & PanelType & "' " & _
              "and PanelName = '" & PanelName & "'"
14570   Set tc = New Recordset
14580   RecOpenServer 0, tc, sql
14590   Do While Not tc.EOF
14600     Set NodX = Tree.Nodes.Add("key" & CStr(Key - 1), tvwChild, , tc!Analyte & "")
14610     tc.MoveNext
14620   Loop
        
14630   tb.MoveNext
14640 Loop

14650 Exit Sub

FillTree_Error:

      Dim strES As String
      Dim intEL As Integer

14660 intEL = Erl
14670 strES = Err.Description
14680 LogError "frmOCPanels", "FillTree", intEL, strES, sql

End Sub

Private Sub baddnew_Click()
        
      Dim NodX As MSComctlLib.Node
      Static k As String

14690 If k = "" Then
14700   k = "1"
14710 Else
14720   k = CStr(Val(k) + 1)
14730 End If

14740 Set NodX = Tree.Nodes.Add(, , "newkey" & k, "New Panel" & k)

End Sub


Private Sub bcancel_Click()

14750 Unload Me

End Sub


Private Sub bRemoveItem_Click()

      Dim pType As String
      Dim pName As String
      Dim pContent As String
      Dim sql As String

14760 On Error GoTo bRemoveItem_Click_Error

14770 pType = ListCodeFor("ST", cmbSampleType)
14780 pName = AddTicks(Tree.SelectedItem.Parent.Text)
14790 pContent = Tree.SelectedItem.Text

14800 sql = "Delete from OCOrderContents where " & _
            "and PanelType = '" & pType & "' " & _
            "and PanelName = '" & pName & "' " & _
            "and Analyte = '" & pContent & "'"
14810 Cnxn(0).Execute sql

14820 FillTree

14830 Exit Sub

bRemoveItem_Click_Error:

      Dim strES As String
      Dim intEL As Integer

14840 intEL = Erl
14850 strES = Err.Description
14860 LogError "frmOCPanels", "bRemoveItem_Click", intEL, strES, sql


End Sub


Private Sub bRemovePanel_Click()

      Dim PanelType As String
      Dim sql As String

14870 On Error GoTo bRemovePanel_Click_Error

14880 PanelType = ListCodeFor("ST", cmbSampleType)

14890 sql = "Delete from OCOrderContents where " & _
            "PanelType = '" & PanelType & "' " & _
            "and PanelName = '" & AddTicks(Tree.SelectedItem.Text) & "'"

14900 Cnxn(0).Execute sql

14910 FillTree

14920 Exit Sub

bRemovePanel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

14930 intEL = Erl
14940 strES = Err.Description
14950 LogError "frmOCPanels", "bRemovePanel_Click", intEL, strES, sql


End Sub


Private Sub cmbSampleType_Click()

14960 FillList
14970 FillTree

End Sub


Private Sub cmdAmendSequence_Click()

14980 frmOCPanelSequence.Show 1

14990 FillTree

End Sub


Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

15000 On Error GoTo Form_Load_Error

15010 sql = "Select * from Lists where " & _
            "ListType = 'ST' and InUse = 1 " & _
            "Order by ListOrder"
15020 Set tb = New Recordset
15030 RecOpenServer 0, tb, sql

15040 cmbSampleType.Clear

15050 Do While Not tb.EOF
15060   cmbSampleType.AddItem tb!Text & ""
15070   tb.MoveNext
15080 Loop
15090 If cmbSampleType.ListCount > 0 Then
15100   cmbSampleType.ListIndex = 0
15110 End If

15120 FillList
15130 FillTree

15140 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

15150 intEL = Erl
15160 strES = Err.Description
15170 LogError "frmOCPanels", "Form_Load", intEL, strES, sql


End Sub


Private Sub ListView1_Click()
      '
15180 Debug.Print "Click"
15190 Debug.Print ListView1.SelectedItem.Text
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

15200 ListView1_Click
15210 Debug.Print "MouseDown"
15220 Debug.Print ListView1.SelectedItem.Text
15230 Debug.Print ListView1.SelectedItem.Ghosted
      'ListView1.DragIcon = ListView1.SelectedItem.CreateDragImage
      'ListView1.Drag

End Sub


Private Sub Tree_AfterLabelEdit(Cancel As Integer, NewString As String)

      Dim sql As String
      Dim PanelType As String

15240 On Error GoTo Tree_AfterLabelEdit_Error

15250 PanelType = ListCodeFor("ST", cmbSampleType)

15260 If Trim$(NewString) = "" Then
15270   Cancel = True
15280   Exit Sub
15290 End If

15300 sql = "UPDATE OCOrderContents SET " & _
            "PanelName = '" & NewString & "', " & _
            "PanelType = '" & PanelType & "' " & _
            "WHERE PanelName = '" & Tree.SelectedItem.Text & "'"
15310 Cnxn(0).Execute sql

15320 Exit Sub

Tree_AfterLabelEdit_Error:

      Dim strES As String
      Dim intEL As Integer

15330 intEL = Erl
15340 strES = Err.Description
15350 LogError "frmOCPanels", "Tree_AfterLabelEdit", intEL, strES, sql

End Sub

Private Sub Tree_DragDrop(Source As Control, X As Single, Y As Single)
        
      Dim NodX As MSComctlLib.Node
      Dim Key
      Dim PanelType As String
      Dim ListOrder As Integer
      Dim tb As Recordset
      Dim sql As String

15360 On Error GoTo Tree_DragDrop_Error

15370 PanelType = ListCodeFor("ST", cmbSampleType)

15380 Set NodX = Tree.HitTest(X, Y)
15390 If NodX Is Nothing Then Exit Sub

15400 If Tree.DropHighlight Is Nothing Then
15410   Set Tree.DropHighlight = Nothing
15420   Exit Sub
15430 Else
15440   If NodX = Tree.DropHighlight Then
15450     Key = NodX.Key
15460     If Key <> "" Then
15470       Set NodX = Tree.Nodes.Add(Key, tvwChild, , Source.Text)

15480       sql = "Select * from OCOrderContents where " & _
                  "PanelType = '" & PanelType & "' " & _
                  "and PanelName = '" & Tree.DropHighlight.Text & "'"
15490       Set tb = New Recordset
15500       RecOpenServer 0, tb, sql
15510       If Not tb.EOF Then
15520         ListOrder = tb!ListOrder
15530       End If
15540       tb.AddNew
15550       tb!PanelName = Tree.DropHighlight.Text
15560       tb!ListOrder = ListOrder
15570       tb!PanelType = PanelType
15580       tb!Analyte = Source.Text
15590       tb.Update
15600       Set Tree.DropHighlight = Nothing
15610       Tree.Nodes(Key).Child.EnsureVisible
15620     End If
15630   End If
15640 End If

15650 Exit Sub

Tree_DragDrop_Error:

      Dim strES As String
      Dim intEL As Integer

15660 intEL = Erl
15670 strES = Err.Description
15680 LogError "frmOCPanels", "Tree_DragDrop", intEL, strES, sql

End Sub


Private Sub Tree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
        
15690 Set Tree.DropHighlight = Tree.HitTest(X, Y)

End Sub


Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
        
15700 bRemovePanel.Enabled = False
15710 bRemoveItem.Enabled = False

End Sub


Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)

      Dim s As String
      Dim strTitle As String

15720 s = Node.Key
15730 strTitle = Node.Text

15740 If s = "" Then
15750   bRemoveItem.Caption = "Remove " & strTitle
15760   bRemoveItem.Font.Bold = True
15770   bRemovePanel.Caption = "Remove Panel"
15780   bRemovePanel.Font.Bold = False
15790   bRemovePanel.Enabled = False
15800   bRemoveItem.Enabled = True
15810 Else
15820   bRemoveItem.Caption = "Remove Item"
15830   bRemoveItem.Font.Bold = False
15840   bRemovePanel.Caption = "Remove " & strTitle
15850   bRemovePanel.Font.Bold = True
15860   bRemovePanel.Enabled = True
15870   bRemoveItem.Enabled = False
15880 End If

End Sub


