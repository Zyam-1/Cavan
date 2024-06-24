VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExtPanels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Define External Panels"
   ClientHeight    =   8175
   ClientLeft      =   1845
   ClientTop       =   660
   ClientWidth     =   8460
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
   ScaleHeight     =   8175
   ScaleWidth      =   8460
   Begin MSComctlLib.TreeView tv 
      Height          =   7665
      Left            =   210
      TabIndex        =   5
      Top             =   90
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   13520
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   6
      Appearance      =   1
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
      Height          =   765
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3930
      Width           =   1905
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
      Height          =   765
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3060
      Width           =   1905
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1620
      Width           =   1905
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   585
      Left            =   6480
      Picture         =   "frmExtPanels.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5790
      Width           =   1905
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   7695
      Left            =   3720
      TabIndex        =   4
      Top             =   60
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   13573
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmExtPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NodX As MSComctlLib.Node       ' Item that is being dragged.
Dim nodXText As String ' Items text

Private Sub FillTV()

          Dim NodX As MSComctlLib.Node
          Dim n As Integer
          Dim Relative As String
          Dim ThisNode As String
          Dim sql As String
          Dim tb As Recordset

40100     On Error GoTo FillTV_Error

40110     For n = Asc("A") To Asc("Z")
40120         Set NodX = tv.Nodes.Add(, , Chr$(n), Chr$(n))
40130     Next
40140     For n = Asc("0") To Asc("9")
40150         Set NodX = tv.Nodes.Add(, , "#" & Chr$(n), Chr$(n))
40160     Next

40170     sql = "Select * from ExternalDefinitions"
40180     Set tb = New Recordset
40190     RecOpenServer 0, tb, sql
40200     Do While Not tb.EOF
40210         If Trim$(tb!AnalyteName & "") <> "" Then
40220             Relative = UCase(Left(tb!AnalyteName, 1))
40230             If IsNumeric(Relative) Then Relative = "#" & Relative
40240             ThisNode = tb!AnalyteName
40250             Set NodX = tv.Nodes.Add(Relative, tvwChild, , ThisNode)
40260         End If
40270         tb.MoveNext
40280     Loop

40290     Exit Sub

FillTV_Error:

          Dim strES As String
          Dim intEL As Integer

40300     intEL = Erl
40310     strES = Err.Description
40320     LogError "frmExtPanels", "FillTV", intEL, strES, sql


End Sub

Private Sub baddnew_Click()
        
          Dim NodX As MSComctlLib.Node
          Static k As String
40330     On Error Resume Next

40340     If k = "" Then
40350         k = "1"
40360     Else
40370         k = CStr(Val(k) + 1)
40380     End If

40390     Set NodX = Tree.Nodes.Add(, , "newkey" & k, "New Panel" & k)

End Sub

Private Sub cmdCancel_Click()

40400     Unload Me

End Sub

Private Sub bRemoveItem_Click()

          Dim sql As String
          Dim p As String

40410     On Error GoTo bRemoveItem_Click_Error

40420     p = Tree.SelectedItem.Parent.Text

40430     sql = "DELETE FROM ExtPanels " & _
              "WHERE PanelName = '" & p & "' " & _
              "and Content = '" & Tree.SelectedItem.Text & "'"
40440     Cnxn(0).Execute sql

40450     bRemoveItem.Caption = "Remove Item"
40460     bRemoveItem.Font.Bold = False
40470     bRemoveItem.Enabled = False

40480     FillTree

40490     Exit Sub

bRemoveItem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40500     intEL = Erl
40510     strES = Err.Description
40520     LogError "frmExtPanels", "bRemoveItem_Click", intEL, strES, sql

End Sub

Private Sub bRemovePanel_Click()

          Dim sql As String

40530     On Error GoTo bRemovePanel_Click_Error

40540     sql = "DELETE FROM ExtPanels " & _
              "WHERE PanelName = '" & Tree.SelectedItem.Text & "'"
40550     Cnxn(0).Execute sql
        
40560     bRemovePanel.Caption = "Remove Panel"
40570     bRemovePanel.Font.Bold = False
40580     bRemovePanel.Enabled = False

40590     FillTree

40600     Exit Sub

bRemovePanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40610     intEL = Erl
40620     strES = Err.Description
40630     LogError "frmExtPanels", "bRemovePanel_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

40640     FillTV
40650     FillTree

End Sub

Private Sub Tree_AfterLabelEdit(Cancel As Integer, NewString As String)

          Dim sql As String

40660     On Error GoTo Tree_AfterLabelEdit_Error

40670     If Trim(NewString) = "" Then
40680         Cancel = True
40690         Exit Sub
40700     End If

40710     sql = "UPDATE ePanels SET " & _
              "PanelName = '" & NewString & "' " & _
              "WHERE PanelName = '" & Tree.SelectedItem.Text & "'"
40720     Cnxn(0).Execute sql

40730     Exit Sub

Tree_AfterLabelEdit_Error:

          Dim strES As String
          Dim intEL As Integer

40740     intEL = Erl
40750     strES = Err.Description
40760     LogError "frmExtPanels", "Tree_AfterLabelEdit", intEL, strES, sql

End Sub

Private Sub Tree_Collapse(ByVal Node As MSComctlLib.Node)

40770     bRemovePanel.Enabled = False
40780     bRemoveItem.Enabled = False

End Sub


Private Sub Tree_DragDrop(Source As Control, X As Single, Y As Single)
        
          Dim tb As Recordset
          Dim sql As String
          Dim nodT As MSComctlLib.Node
          Dim Key

40790     On Error GoTo Tree_DragDrop_Error

40800     Set nodT = Tree.HitTest(X, Y)
40810     If nodT Is Nothing Then Exit Sub

40820     If Tree.DropHighlight Is Nothing Then
40830         Exit Sub
40840     Else
40850         If nodT = Tree.DropHighlight Then
40860             Key = nodT.Key
40870             If Key <> "" Then
40880                 Set nodT = Tree.Nodes.Add(Key, tvwChild, , nodXText)
40890                 sql = "Select * from ExtPanels"
40900                 Set tb = New Recordset
40910                 RecOpenServer 0, tb, sql
40920                 tb.AddNew
40930                 tb!PanelName = Tree.DropHighlight.Text
40940                 tb!Content = nodXText
40950                 tb.Update
40960                 Set Tree.DropHighlight = Nothing
40970             End If
40980         End If
40990     End If
41000     nodXText = ""

41010     Exit Sub

Tree_DragDrop_Error:

          Dim strES As String
          Dim intEL As Integer

41020     intEL = Erl
41030     strES = Err.Description
41040     LogError "frmExtPanels", "Tree_DragDrop", intEL, strES, sql


End Sub


Private Sub Tree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
        
41050     Set Tree.DropHighlight = Tree.HitTest(X, Y)

End Sub
Private Sub FillTree()

          Dim sn As Recordset
          Dim snp As Recordset
          Dim sql As String
          Dim NodX As MSComctlLib.Node
          Dim Key As Integer

41060     On Error GoTo FillTree_Error

41070     Tree.Nodes.Clear
41080     sql = "SELECT DISTINCT PanelName FROM ExtPanels " & _
              "ORDER BY PanelName"
41090     Set sn = New Recordset
41100     RecOpenServer 0, sn, sql
41110     Key = 1
41120     Do While Not sn.EOF

41130         Set NodX = Tree.Nodes.Add(, , "key" & CStr(Key), sn!PanelName)
41140         Key = Key + 1
        
41150         sql = "SELECT * from ExtPanels " & _
                  "WHERE PanelName = '" & sn!PanelName & "'"
41160         Set snp = New Recordset
41170         RecOpenServer 0, snp, sql
41180         Do While Not snp.EOF
41190             Set NodX = Tree.Nodes.Add("key" & CStr(Key - 1), tvwChild, , snp!Content)
41200             snp.MoveNext
41210         Loop
41220         sn.MoveNext
41230     Loop

41240     Exit Sub

FillTree_Error:

          Dim strES As String
          Dim intEL As Integer

41250     intEL = Erl
41260     strES = Err.Description
41270     LogError "frmExtPanels", "FillTree", intEL, strES, sql

End Sub



Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)

41280     bRemovePanel.Enabled = False
41290     bRemoveItem.Enabled = False

End Sub





Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)

          Dim s As String
          Dim strTitle As String

41300     s = Node.Key
41310     strTitle = Node.Text

41320     If s = "" Then
41330         bRemoveItem.Caption = "Remove " & strTitle
41340         bRemoveItem.Font.Bold = True
41350         bRemovePanel.Caption = "Remove Panel"
41360         bRemovePanel.Font.Bold = False
41370         bRemovePanel.Enabled = False
41380         bRemoveItem.Enabled = True
41390     Else
41400         bRemoveItem.Caption = "Remove Item"
41410         bRemoveItem.Font.Bold = False
41420         bRemovePanel.Caption = "Remove " & strTitle
41430         bRemovePanel.Font.Bold = True
41440         bRemovePanel.Enabled = True
41450         bRemoveItem.Enabled = False
41460     End If

End Sub


Private Sub tv_Click()

41470     nodXText = tv.SelectedItem.Text
41480     Debug.Print "Click "; nodXText

End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

41490     On Error Resume Next

          'If Not tv.SelectedItem Is Nothing Then
          '  Set NodX = tv.SelectedItem
41500     nodXText = tv.HitTest(X, Y)
41510     Debug.Print "Mousedown "; nodXText
          'End If

End Sub


Private Sub tv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

41520     If Button = vbLeftButton Then
              'nodXText = tv.SelectedItem.Text
41530         Debug.Print "Mousemove "; nodXText
41540         If nodXText <> "0" Then
41550             tv.Drag vbBeginDrag
41560         End If
41570     End If

End Sub





Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)

41580     nodXText = tv.SelectedItem.Text
41590     Debug.Print "Nodeclick "; nodXText

End Sub
