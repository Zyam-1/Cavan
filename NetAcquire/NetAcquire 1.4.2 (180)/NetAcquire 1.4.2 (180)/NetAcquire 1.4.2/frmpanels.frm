VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Define Panels"
   ClientHeight    =   5775
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   10080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5775
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdAmendSequence 
      Caption         =   "Amend Panel Sequence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5070
      TabIndex        =   7
      Top             =   5070
      Width           =   2715
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   4785
      Left            =   5070
      TabIndex        =   6
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
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   270
      Width           =   2145
   End
   Begin VB.CommandButton bRemoveItem 
      Caption         =   "Remove Item"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton bRemovePanel 
      Caption         =   "Remove Panel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton baddnew 
      Caption         =   "Add New Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8340
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Columns         =   3
      DragIcon        =   "frmpanels.frx":0000
      Height          =   5130
      Left            =   240
      TabIndex        =   1
      Top             =   270
      Width           =   4815
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8430
      Picture         =   "frmpanels.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4860
      Width           =   1245
   End
End
Attribute VB_Name = "frmPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub baddnew_Click()
        
      Dim NodX As MSComctlLib.Node
      Static k As String

21160 If k = "" Then
21170   k = "1"
21180 Else
21190   k = CStr(Val(k) + 1)
21200 End If

21210 Set NodX = Tree.Nodes.Add(, , "newkey" & k, "New Panel" & k)

End Sub

Private Sub bcancel_Click()

21220 Unload Me

End Sub

Private Sub FillList()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleType As String

21230 On Error GoTo FillList_Error

21240 List1.Clear

21250 SampleType = ListCodeFor("ST", cmbSampleType)

21260 sql = "Select distinct Shortname, PrintPriority from BioTestDefinitions where " & _
            "SampleType = '" & SampleType & "' " & _
            "and Hospital = '" & HospName(0) & "' " & _
            "Order by PrintPriority"
21270 Set tb = New Recordset
21280 RecOpenServer 0, tb, sql
21290 Do While Not tb.EOF
21300   List1.AddItem tb!ShortName & ""
21310   tb.MoveNext
21320 Loop

21330 Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

21340 intEL = Erl
21350 strES = Err.Description
21360 LogError "fPanels", "FillList", intEL, strES, sql

End Sub

Private Sub bRemoveItem_Click()

      Dim pType As String
      Dim pName As String
      Dim pContent As String
      Dim sql As String

21370 On Error GoTo bRemoveItem_Click_Error

21380 pType = ListCodeFor("ST", cmbSampleType)
21390 pName = AddTicks(Tree.SelectedItem.Parent.Text)
21400 pContent = Tree.SelectedItem.Text

21410 sql = "Delete from Panels where " & _
            "Hospital = '" & HospName(0) & "' " & _
            "and PanelType = '" & pType & "' " & _
            "and PanelName = '" & pName & "' " & _
            "and Content = '" & pContent & "'"
21420 Cnxn(0).Execute sql

21430 FillTree

21440 Exit Sub

bRemoveItem_Click_Error:

      Dim strES As String
      Dim intEL As Integer

21450 intEL = Erl
21460 strES = Err.Description
21470 LogError "fPanels", "bRemoveItem_Click", intEL, strES, sql


End Sub

Private Sub bRemovePanel_Click()

      Dim PanelType As String
      Dim sql As String

21480 On Error GoTo bRemovePanel_Click_Error

21490 PanelType = ListCodeFor("ST", cmbSampleType)

21500 sql = "Delete from Panels where " & _
            "PanelType = '" & PanelType & "' " & _
            "and PanelName = '" & AddTicks(Tree.SelectedItem.Text) & "' " & _
            "and Hospital = '" & HospName(0) & "' "

21510 Cnxn(0).Execute sql

21520 FillTree

21530 Exit Sub

bRemovePanel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

21540 intEL = Erl
21550 strES = Err.Description
21560 LogError "fPanels", "bRemovePanel_Click", intEL, strES, sql


End Sub

Private Sub cmbSampleType_Click()

21570 FillList
21580 FillTree

End Sub


Private Sub cmdAmendSequence_Click()

21590 frmPanelSequence.Show 1

21600 FillTree

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

21610 On Error GoTo Form_Load_Error

21620 sql = "Select * from Lists where " & _
            "ListType = 'ST' and InUse = 1 " & _
            "Order by ListOrder"
21630 Set tb = New Recordset
21640 RecOpenServer 0, tb, sql

21650 cmbSampleType.Clear

21660 Do While Not tb.EOF
21670   cmbSampleType.AddItem tb!Text & ""
21680   tb.MoveNext
21690 Loop
21700 If cmbSampleType.ListCount > 0 Then
21710   cmbSampleType.ListIndex = 0
21720 End If

21730 FillList
21740 FillTree

21750 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

21760 intEL = Erl
21770 strES = Err.Description
21780 LogError "fPanels", "Form_Load", intEL, strES, sql


End Sub

Private Sub List1_Click()
                                                                                                            '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

21790 List1_Click
21800 List1.Drag

End Sub


Private Sub Tree_AfterLabelEdit(Cancel As Integer, NewString As String)

      Dim sql As String
      Dim PanelType As String

21810 On Error GoTo Tree_AfterLabelEdit_Error

21820 PanelType = ListCodeFor("ST", cmbSampleType)

21830 If Trim$(NewString) = "" Then
21840   Cancel = True
21850   Exit Sub
21860 End If

21870 sql = "UPDATE Panels SET " & _
            "PanelName = '" & NewString & "', " & _
            "PanelType = '" & PanelType & "' " & _
            "WHERE PanelName = '" & Tree.SelectedItem.Text & "'"
21880 Cnxn(0).Execute sql

21890 Exit Sub

Tree_AfterLabelEdit_Error:

      Dim strES As String
      Dim intEL As Integer

21900 intEL = Erl
21910 strES = Err.Description
21920 LogError "fPanels", "Tree_AfterLabelEdit", intEL, strES, sql

End Sub

Private Sub Tree_DragDrop(Source As Control, X As Single, Y As Single)
        
      Dim NodX As MSComctlLib.Node
      Dim Key
      Dim PanelType As String
      Dim BarCode As String
      Dim ListOrder As Integer
      Dim tb As Recordset
      Dim sql As String

21930 On Error GoTo Tree_DragDrop_Error

21940 PanelType = ListCodeFor("ST", cmbSampleType)

21950 Set NodX = Tree.HitTest(X, Y)
21960 If NodX Is Nothing Then Exit Sub

21970 If Tree.DropHighlight Is Nothing Then
21980   Set Tree.DropHighlight = Nothing
21990   Exit Sub
22000 Else
22010   If NodX = Tree.DropHighlight Then
22020     Key = NodX.Key
22030     If Key <> "" Then
22040       Set NodX = Tree.Nodes.Add(Key, tvwChild, , Source.Text)

22050       sql = "Select * from Panels where " & _
                  "PanelType = '" & PanelType & "' " & _
                  "and PanelName = '" & Tree.DropHighlight.Text & "' " & _
                  "and Hospital = '" & HospName(0) & "'"
                  '"and Content = '" & NodX.Child.Text & "' "
22060       Set tb = New Recordset
22070       RecOpenServer 0, tb, sql
22080       If Not tb.EOF Then
22090         BarCode = tb!BarCode & ""
22100         ListOrder = tb!ListOrder
22110       End If
22120       tb.AddNew
22130       tb!PanelName = Tree.DropHighlight.Text
22140       tb!BarCode = BarCode
22150       tb!ListOrder = ListOrder
22160       tb!PanelType = PanelType
22170       tb!Content = Source.Text
22180       tb!Hospital = HospName(0)
22190       tb.Update
22200       Set Tree.DropHighlight = Nothing
22210       Tree.Nodes(Key).Child.EnsureVisible
22220     End If
22230   End If
22240 End If

22250 Exit Sub

Tree_DragDrop_Error:

      Dim strES As String
      Dim intEL As Integer

22260 intEL = Erl
22270 strES = Err.Description
22280 LogError "fPanels", "Tree_DragDrop", intEL, strES, sql

End Sub

Private Sub Tree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
        
22290 Set Tree.DropHighlight = Tree.HitTest(X, Y)

End Sub
Private Sub FillTree()

      Dim NodX As MSComctlLib.Node
      Dim Key As Integer
      Dim PanelType As String
      Dim PanelName As String
      Dim tb As Recordset
      Dim tc As Recordset
      Dim sql As String

22300 On Error GoTo FillTree_Error

22310 PanelType = ListCodeFor("ST", cmbSampleType)

22320 Tree.Nodes.Clear
          
22330 sql = "Select distinct PanelName, Listorder from Panels where " & _
            "PanelType = '" & PanelType & "' " & _
            "and Hospital = '" & HospName(0) & "' " & _
            "order by ListOrder"
22340 Set tb = New Recordset
22350 RecOpenServer 0, tb, sql
22360 Do While Not tb.EOF
22370   PanelName = tb!PanelName & ""
          
22380   Set NodX = Tree.Nodes.Add(, , "key" & CStr(Key), PanelName)
22390   Key = Key + 1
          
22400   sql = "Select Content from Panels where " & _
              "PanelType = '" & PanelType & "' " & _
              "and PanelName = '" & PanelName & "' " & _
              "and Hospital = '" & HospName(0) & "'"
22410   Set tc = New Recordset
22420   RecOpenServer 0, tc, sql
22430   Do While Not tc.EOF
22440     Set NodX = Tree.Nodes.Add("key" & CStr(Key - 1), tvwChild, , tc!Content & "")
22450     tc.MoveNext
22460   Loop
        
22470   tb.MoveNext
22480 Loop

22490 Exit Sub

FillTree_Error:

      Dim strES As String
      Dim intEL As Integer

22500 intEL = Erl
22510 strES = Err.Description
22520 LogError "fPanels", "FillTree", intEL, strES, sql

End Sub

Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
        
22530 bRemovePanel.Enabled = False
22540 bRemoveItem.Enabled = False

End Sub

Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)

      Dim s As String
      Dim strTitle As String

22550 s = Node.Key
22560 strTitle = Node.Text

22570 If s = "" Then
22580   bRemoveItem.Caption = "Remove " & strTitle
22590   bRemoveItem.Font.Bold = True
22600   bRemovePanel.Caption = "Remove Panel"
22610   bRemovePanel.Font.Bold = False
22620   bRemovePanel.Enabled = False
22630   bRemoveItem.Enabled = True
22640 Else
22650   bRemoveItem.Caption = "Remove Item"
22660   bRemoveItem.Font.Bold = False
22670   bRemovePanel.Caption = "Remove " & strTitle
22680   bRemovePanel.Font.Bold = True
22690   bRemovePanel.Enabled = True
22700   bRemoveItem.Enabled = False
22710 End If

End Sub


