VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCoagPanels 
   Caption         =   "NetAcquire - Coagulation Panels"
   ClientHeight    =   3570
   ClientLeft      =   570
   ClientTop       =   900
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5985
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   4500
      Picture         =   "frmCoagPanels.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2340
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Columns         =   3
      DragIcon        =   "frmCoagPanels.frx":066A
      Height          =   3180
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   1785
   End
   Begin VB.CommandButton baddnew 
      Caption         =   "Add New Panel"
      Height          =   525
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   540
      Width           =   1245
   End
   Begin VB.CommandButton bRemovePanel 
      Caption         =   "Remove Panel"
      Enabled         =   0   'False
      Height          =   525
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1140
      Width           =   1245
   End
   Begin VB.CommandButton bRemoveItem 
      Caption         =   "Remove Item"
      Enabled         =   0   'False
      Height          =   525
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1740
      Width           =   1245
   End
   Begin VB.CommandButton cmdAmendSequence 
      Caption         =   "Amend Panel Sequence"
      Height          =   345
      Left            =   2070
      TabIndex        =   0
      Top             =   3000
      Width           =   2115
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   2805
      Left            =   2070
      TabIndex        =   1
      Top             =   150
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   4948
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmCoagPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub baddnew_Click()
        
          Dim NodX As MSComctlLib.Node
          Static k As String

19180     If k = "" Then
19190         k = "1"
19200     Else
19210         k = CStr(Val(k) + 1)
19220     End If

19230     Set NodX = Tree.Nodes.Add(, , "newkey" & k, "New Panel" & k)

End Sub

Private Sub bcancel_Click()

19240     Unload Me

End Sub


Private Sub bRemoveItem_Click()

          Dim pName As String
          Dim pContent As String
          Dim sql As String

19250     On Error GoTo bRemoveItem_Click_Error

19260     pName = AddTicks(Tree.SelectedItem.Parent.Text)
19270     pContent = Tree.SelectedItem.Text

19280     sql = "Delete from CoagPanels where " & _
              "PanelName = '" & pName & "' " & _
              "and Content = '" & pContent & "'"
19290     Cnxn(0).Execute sql

19300     FillTree

19310     Exit Sub

bRemoveItem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

19320     intEL = Erl
19330     strES = Err.Description
19340     LogError "frmCoagPanels", "bRemoveItem_Click", intEL, strES, sql


End Sub


Private Sub bRemovePanel_Click()

          Dim sql As String

19350     On Error GoTo bRemovePanel_Click_Error

19360     sql = "Delete from CoagPanels where " & _
              "PanelName = '" & AddTicks(Tree.SelectedItem.Text) & "'"

19370     Cnxn(0).Execute sql

19380     FillTree

19390     Exit Sub

bRemovePanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

19400     intEL = Erl
19410     strES = Err.Description
19420     LogError "frmCoagPanels", "bRemovePanel_Click", intEL, strES, sql


End Sub


Private Sub cmdAmendSequence_Click()

19430     frmCoagPanelSequence.Show 1

19440     FillTree

End Sub


Private Sub Form_Load()

19450     FillList
19460     FillTree

End Sub


Private Sub FillList()

          Dim tb As Recordset
          Dim sql As String

19470     On Error GoTo FillList_Error

19480     List1.Clear

19490     sql = "Select distinct TestName, PrintPriority " & _
              "from CoagTestDefinitions " & _
              "Order by PrintPriority"
19500     Set tb = New Recordset
19510     RecOpenServer 0, tb, sql
19520     Do While Not tb.EOF
19530         List1.AddItem tb!TestName & ""
19540         tb.MoveNext
19550     Loop

19560     Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

19570     intEL = Erl
19580     strES = Err.Description
19590     LogError "frmCoagPanels", "FillList", intEL, strES, sql


End Sub


Private Sub FillTree()

          Dim NodX As MSComctlLib.Node
          Dim Key As Integer
          Dim PanelName As String
          Dim tb As Recordset
          Dim tc As Recordset
          Dim sql As String

19600     On Error GoTo FillTree_Error

19610     Tree.Nodes.Clear
          
19620     sql = "Select distinct PanelName, Listorder from CoagPanels " & _
              "order by ListOrder"
19630     Set tb = New Recordset
19640     RecOpenServer 0, tb, sql
19650     Do While Not tb.EOF
19660         PanelName = tb!PanelName & ""
          
19670         Set NodX = Tree.Nodes.Add(, , "key" & CStr(Key), PanelName)
19680         Key = Key + 1
          
19690         sql = "Select Content from CoagPanels where " & _
                  "PanelName = '" & PanelName & "'"
19700         Set tc = New Recordset
19710         RecOpenServer 0, tc, sql
19720         Do While Not tc.EOF
19730             Set NodX = Tree.Nodes.Add("key" & CStr(Key - 1), tvwChild, , tc!Content & "")
19740             tc.MoveNext
19750         Loop
        
19760         tb.MoveNext
19770     Loop

19780     Exit Sub

FillTree_Error:

          Dim strES As String
          Dim intEL As Integer

19790     intEL = Erl
19800     strES = Err.Description
19810     LogError "frmCoagPanels", "FillTree", intEL, strES, sql

End Sub


Private Sub List1_Click()
    '
End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

19820     List1_Click
19830     List1.Drag

End Sub


Private Sub Tree_AfterLabelEdit(Cancel As Integer, NewString As String)

          Dim sql As String

19840     On Error GoTo Tree_AfterLabelEdit_Error

19850     If Trim$(NewString) = "" Then
19860         Cancel = True
19870         Exit Sub
19880     End If

19890     sql = "UPDATE CoagPanels SET " & _
              "PanelName = '" & NewString & "' " & _
              "WHERE PanelName = '" & Tree.SelectedItem.Text & "'"
19900     Cnxn(0).Execute sql

19910     Exit Sub

Tree_AfterLabelEdit_Error:

          Dim strES As String
          Dim intEL As Integer

19920     intEL = Erl
19930     strES = Err.Description
19940     LogError "frmCoagPanels", "Tree_AfterLabelEdit", intEL, strES, sql

End Sub

Private Sub Tree_DragDrop(Source As Control, X As Single, Y As Single)
        
          Dim NodX As MSComctlLib.Node
          Dim Key
          Dim BarCode As String
          Dim ListOrder As Integer
          Dim tb As Recordset
          Dim sql As String

19950     On Error GoTo Tree_DragDrop_Error

19960     Set NodX = Tree.HitTest(X, Y)
19970     If NodX Is Nothing Then Exit Sub

19980     If Tree.DropHighlight Is Nothing Then
19990         Set Tree.DropHighlight = Nothing
20000         Exit Sub
20010     Else
20020         If NodX = Tree.DropHighlight Then
20030             Key = NodX.Key
20040             If Key <> "" Then
20050                 Set NodX = Tree.Nodes.Add(Key, tvwChild, , Source.Text)

20060                 sql = "Select * from CoagPanels where " & _
                          "PanelName = '" & Tree.DropHighlight.Text & "'"
20070                 Set tb = New Recordset
20080                 RecOpenServer 0, tb, sql
20090                 If Not tb.EOF Then
20100                     BarCode = tb!BarCode & ""
20110                     ListOrder = tb!ListOrder
20120                 End If
20130                 tb.AddNew
20140                 tb!PanelName = Tree.DropHighlight.Text
20150                 tb!BarCode = BarCode
20160                 tb!ListOrder = ListOrder
20170                 tb!Content = Source.Text
20180                 tb.Update
20190                 Set Tree.DropHighlight = Nothing
20200                 Tree.Nodes(Key).Child.EnsureVisible
20210             End If
20220         End If
20230     End If

20240     Exit Sub

Tree_DragDrop_Error:

          Dim strES As String
          Dim intEL As Integer

20250     intEL = Erl
20260     strES = Err.Description
20270     LogError "frmCoagPanels", "Tree_DragDrop", intEL, strES, sql

End Sub


Private Sub Tree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
        
20280     Set Tree.DropHighlight = Tree.HitTest(X, Y)

End Sub


Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
        
20290     bRemovePanel.Enabled = False
20300     bRemoveItem.Enabled = False

End Sub


Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)

          Dim s As String
          Dim strTitle As String

20310     s = Node.Key
20320     strTitle = Node.Text

20330     If s = "" Then
20340         bRemoveItem.Caption = "Remove " & strTitle
20350         bRemoveItem.Font.Bold = True
20360         bRemovePanel.Caption = "Remove Panel"
20370         bRemovePanel.Font.Bold = False
20380         bRemovePanel.Enabled = False
20390         bRemoveItem.Enabled = True
20400     Else
20410         bRemoveItem.Caption = "Remove Item"
20420         bRemoveItem.Font.Bold = False
20430         bRemovePanel.Caption = "Remove " & strTitle
20440         bRemovePanel.Font.Bold = True
20450         bRemovePanel.Enabled = True
20460         bRemoveItem.Enabled = False
20470     End If

End Sub


