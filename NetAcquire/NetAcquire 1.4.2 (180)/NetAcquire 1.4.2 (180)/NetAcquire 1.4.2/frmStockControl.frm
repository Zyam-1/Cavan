VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockControl 
   Caption         =   "NetAcquire - Select Parameter"
   ClientHeight    =   5295
   ClientLeft      =   450
   ClientTop       =   600
   ClientWidth     =   6690
   Icon            =   "frmStockControl.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   6690
   Begin VB.TextBox txtCurrentStock 
      Height          =   285
      Left            =   1590
      TabIndex        =   9
      Top             =   4830
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockControl.frx":000C
            Key             =   "Square"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockControl.frx":02F2
            Key             =   "SquareCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockControl.frx":05D8
            Key             =   "SquareTick"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockControl.frx":08BE
            Key             =   "Tubes"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkMonitor 
      Alignment       =   1  'Right Justify
      Caption         =   "Monitor Reagent Usage"
      Height          =   255
      Left            =   450
      TabIndex        =   8
      Top             =   4560
      Width           =   1995
   End
   Begin VB.TextBox txtEachAssay 
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   4530
      Width           =   915
   End
   Begin VB.TextBox txtAlarmAssays 
      Height          =   285
      Left            =   5040
      TabIndex        =   3
      Top             =   4830
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Top             =   210
      Width           =   855
   End
   Begin VB.ListBox lstReagents 
      DragIcon        =   "frmStockControl.frx":1308
      Height          =   3960
      Left            =   240
      TabIndex        =   1
      Top             =   555
      Width           =   2205
   End
   Begin MSComctlLib.TreeView tvwParameter 
      Height          =   4305
      Left            =   2970
      TabIndex        =   0
      Top             =   210
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   7594
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Each Assay uses"
      Height          =   195
      Left            =   3780
      TabIndex        =   7
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Current Stock"
      Height          =   195
      Left            =   570
      TabIndex        =   6
      Top             =   4890
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Alarm when stock less than"
      Height          =   195
      Left            =   3030
      TabIndex        =   5
      Top             =   4890
      Width           =   1965
   End
End
Attribute VB_Name = "frmStockControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddReagentsToParameter(ByVal NodX As MSComctlLib.Node, _
                                   ByVal ParameterName As String)
          
      Dim sql As String
      Dim tb As Recordset
      Dim intR As Integer
      Dim nodNew As MSComctlLib.Node
      Dim strKey As String
      Dim MonitorOK As Boolean

65510 On Error GoTo AddReagentsToParameter_Error

65520 sql = "Select C.Reagent, R.Monitor from StockControl as C, StockReagents as R where " & _
            "C.Parameter = '" & ParameterName & "' " & _
            "and C.Reagent = R.Reagent "
65530 Set tb = New Recordset
10    RecOpenServer 0, tb, sql
20    If tb.EOF Then
30      NodX.Image = "Square"
40    Else
50      MonitorOK = True
60      intR = 0
70      strKey = NodX.Key
80      Do While Not tb.EOF
90        intR = intR + 1
100       Set nodNew = tvwParameter.Nodes.Add(strKey, tvwChild, Format$(intR) & strKey, tb!Reagent)
110       If Not tb!Monitor Then
120         nodNew.Image = "SquareCross"
130         MonitorOK = False
140       Else
150         nodNew.Image = "SquareTick"
160       End If
170       tb.MoveNext
180     Loop
190     NodX.Image = IIf(MonitorOK, "SquareTick", "SquareCross")
200   End If

210   Exit Sub

AddReagentsToParameter_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmStockControl", "AddReagentsToParameter", intEL, strES, sql


End Sub

Private Sub FillReagentList()
        
      Dim sql As String
      Dim tb As Recordset

250   On Error GoTo FillReagentList_Error

260   sql = "Select Reagent from StockReagents"
270   Set tb = New Recordset
280   RecOpenServer 0, tb, sql

290   lstReagents.Clear
300   Do While Not tb.EOF
310     lstReagents.AddItem tb!Reagent
320     tb.MoveNext
330   Loop

340   Exit Sub

FillReagentList_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "frmStockControl", "FillReagentList", intEL, strES, sql


End Sub

Private Sub FillTree()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim NodX As MSComctlLib.Node
      Dim strKey As String

380   On Error GoTo FillTree_Error

390   With tvwParameter.Nodes
400     .Clear
410     .Add , , "BioRoot", "Biochemistry", "Tubes"
420     .Add , , "CoagRoot", "Coagulation", "Tubes"
430     .Add , , "CytoRoot", "Cytology", "Tubes"
440     .Add , , "HaemRoot", "Haematology", "Tubes"
450     .Add , , "HistoRoot", "Histology", "Tubes"
460     .Add , , "ImmunoRoot", "Immunology", "Tubes"

470     sql = "Select distinct LongName from BioTestDefinitions " & _
              "Order by LongName"
480     Set tb = New Recordset
490     RecOpenServer 0, tb, sql
500     n = 0
510     Do While Not tb.EOF
520       n = n + 1
530       strKey = "B" & Format$(n)
540       Set NodX = .Add("BioRoot", tvwChild, strKey, tb!LongName & "")
550       AddReagentsToParameter NodX, tb!LongName & ""
560       tb.MoveNext
570     Loop
        
580     Set NodX = .Add("HaemRoot", tvwChild, "H1", "FBC")
590     AddReagentsToParameter NodX, "FBC"
600     Set NodX = .Add("HaemRoot", tvwChild, "H2", "Retic")
610     AddReagentsToParameter NodX, "Retic"
620     Set NodX = .Add("HaemRoot", tvwChild, "H3", "ESR")
630     AddReagentsToParameter NodX, "ESR"
640     Set NodX = .Add("HaemRoot", tvwChild, "H4", "Film")
650     AddReagentsToParameter NodX, "Film"
660     Set NodX = .Add("HaemRoot", tvwChild, "H5", "Monospot")
670     AddReagentsToParameter NodX, "MonoSpot"
680     Set NodX = .Add("HaemRoot", tvwChild, "H6", "Malaria")
690     AddReagentsToParameter NodX, "Malaria"
700     Set NodX = .Add("HaemRoot", tvwChild, "H7", "Sickledex")
710     AddReagentsToParameter NodX, "Sickledex"
       
720     Set NodX = .Add("CoagRoot", tvwChild, "C1", "PT")
730     AddReagentsToParameter NodX, "PT"
740     Set NodX = .Add("CoagRoot", tvwChild, "C2", "APTT")
750     AddReagentsToParameter NodX, "APTT"
       
760   End With

770   Exit Sub

FillTree_Error:

      Dim strES As String
      Dim intEL As Integer

780   intEL = Erl
790   strES = Err.Description
800   LogError "frmStockControl", "FillTree", intEL, strES, sql


End Sub

Private Sub chkMonitor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim sql As String
      Dim n As Integer
      Dim Found As Boolean

810   On Error GoTo chkMonitor_MouseUp_Error

820   Found = False
830   For n = 0 To lstReagents.ListCount - 1
840     If lstReagents.Selected(n) Then
850       Found = True
860       Exit For
870     End If
880   Next
890   If Not Found Then Exit Sub

900   sql = "Update StockReagents " & _
            "set Monitor = '" & chkMonitor & "' " & _
            "where Reagent = '" & lstReagents.Text & "'"
910   Cnxn(0).Execute sql

920   Exit Sub

chkMonitor_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

930   intEL = Erl
940   strES = Err.Description
950   LogError "frmStockControl", "chkMonitor_MouseUp", intEL, strES, sql


End Sub

Private Sub cmdAdd_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim strReagent As String

960   On Error GoTo cmdAdd_Click_Error

970   strReagent = Trim$(iBOX("New Reagent Name?"))
980   If strReagent = "" Then Exit Sub

990   sql = "Select * from StockReagents where " & _
            "Reagent = '" & strReagent & "'"
1000  Set tb = New Recordset
1010  RecOpenServer 0, tb, sql
1020  If Not tb.EOF Then
1030    iMsg strReagent & " already exists!", vbExclamation
1040    Exit Sub
1050  End If
1060  tb.AddNew
1070  tb!Reagent = strReagent
1080  tb!CurrentStock = 0
1090  tb!AlarmBelow = 0
1100  tb!Monitor = 1
1110  tb.Update

1120  lstReagents.AddItem strReagent

1130  Exit Sub

cmdAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

1140  intEL = Erl
1150  strES = Err.Description
1160  LogError "frmStockControl", "cmdadd_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

1170  FillTree
1180  FillReagentList

End Sub

Private Sub lstReagents_Click()
        
      Dim sql As String
      Dim tb As Recordset

1190  On Error GoTo lstReagents_Click_Error

1200  sql = "Select * from StockReagents where " & _
            "Reagent = '" & lstReagents & "'"
1210  Set tb = New Recordset
1220  RecOpenServer 0, tb, sql

1230  If Not tb.EOF Then
1240    chkMonitor = IIf(tb!Monitor, 1, 0)
1250    txtCurrentStock = tb!CurrentStock
1260    txtAlarmAssays = tb!AlarmBelow
1270  Else
1280    chkMonitor = 0
1290    txtCurrentStock = ""
1300    txtAlarmAssays = ""
1310  End If

1320  Exit Sub

lstReagents_Click_Error:

      Dim strES As String
      Dim intEL As Integer

1330  intEL = Erl
1340  strES = Err.Description
1350  LogError "frmStockControl", "lstReagents_Click", intEL, strES, sql


End Sub

Private Sub lstReagents_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

1360  txtEachAssay = ""

1370  lstReagents_Click
1380  lstReagents.Drag

End Sub


Private Sub tvwParameter_DragDrop(Source As Control, X As Single, Y As Single)
        
      Dim NodX As MSComctlLib.Node
      Dim NodP As MSComctlLib.Node
      Dim Key
      Dim tb As Recordset
      Dim sql As String

1390  On Error GoTo tvwParameter_DragDrop_Error

1400  Set NodX = tvwParameter.HitTest(X, Y)
1410  If NodX Is Nothing Then Exit Sub
1420  If NodX.Parent Is Nothing Then Exit Sub
1430  If NodX.Parent.Parent Is Nothing Then
        
1440    If tvwParameter.DropHighlight Is Nothing Then
1450      Exit Sub
1460    Else
1470      If NodX = tvwParameter.DropHighlight Then
1480        Key = NodX.Key
1490        If Key <> "" Then
1500          Set NodX = Nothing
1510          Set NodX = tvwParameter.Nodes(Key & Source.Text)
1520          If NodX Is Nothing Then
          
1530            Set NodX = tvwParameter.Nodes.Add(Key, tvwChild, Key & Source.Text, Source.Text)
1540            Set NodP = NodX.Parent
1550            sql = "Select Monitor from StockReagents where " & _
                      "Reagent = '" & Source.Text & "'"
1560            Set tb = New Recordset
1570            RecOpenServer 0, tb, sql
1580            If tb!Monitor Then
1590              NodX.Image = "SquareTick"
1600              If NodP.Image <> "SquareCross" Then
1610                NodP.Image = "SquareTick"
1620              End If
1630            Else
1640              NodX.Image = "SquareCross"
1650              NodP.Image = "SquareCross"
1660            End If

1670            sql = "Select * from StockControl where " & _
                      "Parameter = '" & tvwParameter.DropHighlight & "' " & _
                      "and Reagent = '" & Source.Text & "'"
1680            Set tb = New Recordset
1690            RecOpenServer 0, tb, sql
1700            If tb.EOF Then tb.AddNew
1710            tb!Parameter = tvwParameter.DropHighlight
1720            tb!Reagent = Source.Text
1730            tb!UsePerTest = 0
1740            tb.Update
          
1750            Set tvwParameter.DropHighlight = Nothing
1760            tvwParameter.Nodes(Key).Child.EnsureVisible
1770          End If
1780        End If
1790      End If
1800    End If
1810  End If

1820  Exit Sub

tvwParameter_DragDrop_Error:

      Dim strES As String
      Dim intEL As Integer

1830  intEL = Erl
1840  strES = Err.Description
1850  LogError "frmStockControl", "tvwParameter_DragDrop", intEL, strES, sql


End Sub

Private Sub tvwParameter_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
        
1860  Set tvwParameter.DropHighlight = tvwParameter.HitTest(X, Y)

End Sub


Private Sub tvwParameter_NodeClick(ByVal Node As MSComctlLib.Node)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

1870  On Error GoTo tvwParameter_NodeClick_Error

1880  txtEachAssay = ""

1890  If tvwParameter.SelectedItem.Parent Is Nothing Then
1900    Exit Sub
1910  ElseIf tvwParameter.SelectedItem.Parent.Parent Is Nothing Then
1920    Exit Sub
1930  End If

1940  sql = "Select * from StockControl where " & _
            "Parameter = '" & tvwParameter.SelectedItem.Parent.Text & "' " & _
            "and Reagent = '" & tvwParameter.SelectedItem.Text & "'"
1950  Set tb = New Recordset
1960  RecOpenServer 0, tb, sql
1970  If Not tb.EOF Then
1980    txtEachAssay = tb!UsePerTest
1990  End If

2000  For n = 0 To lstReagents.ListCount - 1
2010    If lstReagents.List(n) = tvwParameter.SelectedItem.Text Then
2020      lstReagents.Selected(n) = True
2030      Exit For
2040    End If
2050  Next

2060  Exit Sub

tvwParameter_NodeClick_Error:

      Dim strES As String
      Dim intEL As Integer

2070  intEL = Erl
2080  strES = Err.Description
2090  LogError "frmStockControl", "tvwParameter_NodeClick", intEL, strES, sql


End Sub


Private Sub txtAlarmAssays_KeyUp(KeyCode As Integer, Shift As Integer)

      Dim sql As String
      Dim n As Integer
      Dim Found As Boolean

2100  On Error GoTo txtAlarmAssays_KeyUp_Error

2110  Found = False
2120  For n = 0 To lstReagents.ListCount - 1
2130    If lstReagents.Selected(n) Then
2140      Found = True
2150      Exit For
2160    End If
2170  Next
2180  If Not Found Then Exit Sub

2190  sql = "Update StockReagents " & _
            "set AlarmBelow = " & Val(txtAlarmAssays) & " " & _
            "where Reagent = '" & lstReagents.Text & "'"
2200  Cnxn(0).Execute sql

2210  Exit Sub

txtAlarmAssays_KeyUp_Error:

      Dim strES As String
      Dim intEL As Integer

2220  intEL = Erl
2230  strES = Err.Description
2240  LogError "frmStockControl", "txtAlarmAssays_KeyUp", intEL, strES, sql


End Sub


Private Sub txtCurrentStock_KeyUp(KeyCode As Integer, Shift As Integer)

      Dim sql As String
      Dim intN As Integer
      Dim blnFound As Boolean

2250  On Error GoTo txtCurrentStock_KeyUp_Error

2260  blnFound = False
2270  For intN = 0 To lstReagents.ListCount - 1
2280    If lstReagents.Selected(intN) Then
2290      blnFound = True
2300      Exit For
2310    End If
2320  Next
2330  If Not blnFound Then Exit Sub

2340  sql = "Update StockReagents " & _
            "set CurrentStock = " & Val(txtCurrentStock) & " " & _
            "where Reagent = '" & lstReagents.Text & "'"
2350  Cnxn(0).Execute sql

2360  Exit Sub

txtCurrentStock_KeyUp_Error:

      Dim strES As String
      Dim intEL As Integer

2370  intEL = Erl
2380  strES = Err.Description
2390  LogError "frmStockControl", "txtCurrentStock_KeyUp", intEL, strES, sql


End Sub


Private Sub txtEachAssay_KeyUp(KeyCode As Integer, Shift As Integer)

      Dim sql As String

2400  On Error GoTo txtEachAssay_KeyUp_Error

2410  If tvwParameter.SelectedItem Is Nothing Then Exit Sub
2420  If tvwParameter.SelectedItem.Parent Is Nothing Then Exit Sub
2430  If tvwParameter.SelectedItem.Parent.Parent Is Nothing Then Exit Sub

2440  sql = "Update StockControl " & _
            "set UsePerTest = " & Val(txtEachAssay) & " " & _
            "where Reagent = '" & tvwParameter.SelectedItem.Text & "' " & _
            "and Parameter = '" & tvwParameter.SelectedItem.Parent.Text & "'"
2450  Cnxn(0).Execute sql

2460  Exit Sub

txtEachAssay_KeyUp_Error:

      Dim strES As String
      Dim intEL As Integer

2470  intEL = Erl
2480  strES = Err.Description
2490  LogError "frmStockControl", "txtEachAssay_KeyUp", intEL, strES, sql


End Sub


