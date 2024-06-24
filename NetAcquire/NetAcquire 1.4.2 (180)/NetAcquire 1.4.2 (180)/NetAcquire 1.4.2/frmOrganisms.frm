VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrganisms 
   Caption         =   "NetAcquire - Organisms"
   ClientHeight    =   8820
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   10260
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   9120
      Picture         =   "frmOrganisms.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2580
      Width           =   825
   End
   Begin VB.TextBox txtReportName 
      Height          =   285
      Left            =   6960
      TabIndex        =   18
      Top             =   900
      Width           =   1515
   End
   Begin VB.TextBox txtShortName 
      Height          =   285
      Left            =   5340
      MaxLength       =   15
      TabIndex        =   15
      Top             =   900
      Width           =   1605
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9690
      Top             =   4890
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9690
      Top             =   5460
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   825
      Left            =   9150
      Picture         =   "frmOrganisms.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1410
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Top             =   900
      Width           =   1425
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   9150
      Picture         =   "frmOrganisms.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4590
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   9150
      Picture         =   "frmOrganisms.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5430
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid grdOrg 
      Height          =   7155
      Left            =   900
      TabIndex        =   10
      Top             =   1290
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   12621
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmOrganisms.frx":0FD0
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save Details"
      Height          =   825
      Left            =   9150
      Picture         =   "frmOrganisms.frx":1057
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6270
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   9150
      Picture         =   "frmOrganisms.frx":16C1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7620
      Width           =   795
   End
   Begin VB.CommandButton cmdNewGroup 
      Caption         =   "Add New Group"
      Height          =   315
      Left            =   5430
      TabIndex        =   8
      Top             =   90
      Width           =   1485
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   8520
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtNewOrg 
      Height          =   285
      Left            =   2340
      TabIndex        =   0
      Top             =   900
      Width           =   2985
   End
   Begin VB.ComboBox cmbGroups 
      Height          =   315
      Left            =   2370
      TabIndex        =   5
      Text            =   "cmbGroups"
      Top             =   60
      Width           =   2985
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8850
      TabIndex        =   20
      Top             =   3420
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Report Name"
      Height          =   195
      Left            =   6960
      TabIndex        =   17
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Short Name"
      Height          =   195
      Left            =   5370
      TabIndex        =   16
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   195
      Left            =   900
      TabIndex        =   13
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "New Organism"
      Height          =   195
      Left            =   2370
      TabIndex        =   7
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Members of this Group"
      Height          =   435
      Left            =   30
      TabIndex        =   6
      Top             =   1410
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Organism Group"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   90
      Width           =   1140
   End
End
Attribute VB_Name = "frmOrganisms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

17710 With grdOrg
17720   If .row = .Rows - 1 Then Exit Sub
17730   n = .row
        
17740   FireCounter = FireCounter + 1
17750   If FireCounter > 5 Then
17760     tmrDown.Interval = 100
17770   End If
        
17780   VisibleRows = .height \ .RowHeight(1) - 1
        
17790   .Visible = False
        
17800   s = ""
17810   For X = 0 To .Cols - 1
17820     s = s & .TextMatrix(n, X) & vbTab
17830   Next
17840   s = Left$(s, Len(s) - 1)
        
17850   .RemoveItem n
17860   If n < .Rows Then
17870     .AddItem s, n + 1
17880     .row = n + 1
17890   Else
17900     .AddItem s
17910     .row = .Rows - 1
17920   End If
        
17930   For X = 0 To .Cols - 1
17940     .Col = X
17950     .CellBackColor = vbYellow
17960   Next
        
17970   If Not .RowIsVisible(.row) Or .row = .Rows - 1 Then
17980     If .row - VisibleRows + 1 > 0 Then
17990       .TopRow = .row - VisibleRows + 1
18000     End If
18010   End If
        
18020   .Visible = True
18030 End With

18040 cmdSave.Visible = True

End Sub
Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

18050 With grdOrg
18060   If .row = 1 Then Exit Sub
        
18070   FireCounter = FireCounter + 1
18080   If FireCounter > 5 Then
18090     tmrUp.Interval = 100
18100   End If
        
18110   n = .row
        
18120   .Visible = False
        
18130   s = ""
18140   For X = 0 To .Cols - 1
18150     s = s & .TextMatrix(n, X) & vbTab
18160   Next
18170   s = Left$(s, Len(s) - 1)
        
18180   .RemoveItem n
18190   .AddItem s, n - 1
        
18200   .row = n - 1
18210   For X = 0 To .Cols - 1
18220     .Col = X
18230     .CellBackColor = vbYellow
18240   Next
        
18250   If Not .RowIsVisible(.row) Then
18260     .TopRow = .row
18270   End If
        
18280   .Visible = True
        
18290   cmdSave.Visible = True
18300 End With

End Sub





Private Sub FillGroups()

      Dim tb As Recordset
      Dim sql As String

18310 On Error GoTo FillGroups_Error

18320 cmbGroups.Clear

18330 sql = "Select * from Lists where " & _
            "ListType = 'OR' and InUse = 1 " & _
            "order by ListOrder"
18340 Set tb = New Recordset
18350 RecOpenServer 0, tb, sql
18360 Do While Not tb.EOF
18370   cmbGroups.AddItem tb!Text & ""
18380   tb.MoveNext
18390 Loop

18400 Exit Sub

FillGroups_Error:

      Dim strES As String
      Dim intEL As Integer

18410 intEL = Erl
18420 strES = Err.Description
18430 LogError "frmOrganisms", "FillGroups", intEL, strES, sql


End Sub

Private Sub FillList()

      Dim tb As Recordset
      Dim sql As String

18440 On Error GoTo FillList_Error

18450 grdOrg.Rows = 2
18460 grdOrg.AddItem ""
18470 grdOrg.RemoveItem 1

18480 sql = "Select * from Organisms where " & _
            "GroupName = '" & cmbGroups.Text & "' " & _
            "order by listorder"
18490 Set tb = New Recordset
18500 RecOpenClient 0, tb, sql
18510 Do While Not tb.EOF
18520   grdOrg.AddItem tb!Code & vbTab & _
                       tb!Name & vbTab & _
                       tb!ShortName & vbTab & _
                       tb!ReportName & ""
18530   tb.MoveNext
18540 Loop

18550 If grdOrg.Rows > 2 Then
18560   grdOrg.RemoveItem 1
18570 End If

18580 Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

18590 intEL = Erl
18600 strES = Err.Description
18610 LogError "frmOrganisms", "FillList", intEL, strES, sql


End Sub

Private Sub SaveDetails()

      Dim sql As String
      Dim n As Integer

18620 On Error GoTo SaveDetails_Error

18630   For n = 1 To grdOrg.Rows - 1
18640     If Trim$(grdOrg.TextMatrix(n, 1)) <> "" Then
18650       sql = "IF EXISTS (SELECT * FROM Organisms WHERE " & _
                  "           Name = '" & grdOrg.TextMatrix(n, 1) & "' " & _
                  "           AND GroupName = '" & cmbGroups.Text & "') " & _
                  "  UPDATE Organisms " & _
                  "  SET Code = '" & grdOrg.TextMatrix(n, 0) & "', " & _
                  "  ShortName = '" & grdOrg.TextMatrix(n, 2) & "', " & _
                  "  ReportName = '" & grdOrg.TextMatrix(n, 3) & "', " & _
                  "  ListOrder = '" & n & "' " & _
                  "  WHERE Name = '" & grdOrg.TextMatrix(n, 1) & "' " & _
                  "  AND GroupName = '" & cmbGroups.Text & "' " & _
                  "ELSE " & _
                  "  INSERT INTO Organisms (Code, Name, ShortName, ReportName, GroupName, ListOrder ) " & _
                  "  VALUES ('" & grdOrg.TextMatrix(n, 0) & "', " & _
                  "  '" & grdOrg.TextMatrix(n, 1) & "', " & _
                  "  '" & grdOrg.TextMatrix(n, 2) & "', " & _
                  "  '" & grdOrg.TextMatrix(n, 3) & "', " & _
                  "  '" & cmbGroups.Text & "', " & _
                  "  '" & n & "' )"
18660       Cnxn(0).Execute sql
18670     End If
18680   Next

18690 cmdSave.Visible = False

18700 Exit Sub

SaveDetails_Error:

      Dim strES As String
      Dim intEL As Integer

18710 intEL = Erl
18720 strES = Err.Description
18730 LogError "frmOrganisms", "SaveDetails", intEL, strES, sql

End Sub

Private Sub bMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18740 FireDown

18750 tmrDown.Interval = 250
18760 FireCounter = 0

18770 tmrDown.Enabled = True

End Sub


Private Sub bMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18780 tmrDown.Enabled = False

End Sub


Private Sub bMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18790 FireUp

18800 tmrUp.Interval = 250
18810 FireCounter = 0

18820 tmrUp.Enabled = True

End Sub


Private Sub bMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18830 tmrUp.Enabled = False

End Sub


Private Sub cmbGroups_Click()

18840 FillList

End Sub

Private Sub cmbGroups_GotFocus()

18850 If cmdSave.Visible Then
18860   If iMsg("Save Changes?", vbQuestion + vbYesNo) = vbYes Then
18870     SaveDetails
18880   End If
18890   cmdSave.Visible = False
18900 End If

End Sub


Private Sub cmdAdd_Click()

18910 grdOrg.AddItem UCase$(txtCode) & vbTab & _
                     txtNewOrg & vbTab & _
                     txtShortName & vbTab & _
                     txtReportName
18920 txtNewOrg = ""
18930 txtCode = ""
18940 txtShortName = ""
18950 txtReportName = ""

18960 cmdSave.Visible = True
18970 txtNewOrg.SetFocus

End Sub

Private Sub cmdCancel_Click()

18980 If cmdSave.Visible Then
18990   If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
19000     SaveDetails
19010   End If
19020   cmdSave.Visible = False
19030 End If

19040 Unload Me

End Sub


Private Sub cmdDelete_Click()

19050 On Error GoTo cmdDelete_Click_Error

19060 txtCode = grdOrg.TextMatrix(grdOrg.row, 0)
19070 txtNewOrg = grdOrg.TextMatrix(grdOrg.row, 1)
19080 txtShortName = grdOrg.TextMatrix(grdOrg.row, 2)
19090 txtReportName = grdOrg.TextMatrix(grdOrg.row, 3)

19100 If grdOrg.Rows = 2 Then
19110   grdOrg.AddItem ""
19120   grdOrg.RemoveItem 1
19130 Else
19140   grdOrg.RemoveItem grdOrg.row
19150 End If
19160 cmdDelete.Visible = False
19170 cmdSave.Visible = True
19180 cmdSave.SetFocus

19190 Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

19200 intEL = Erl
19210 strES = Err.Description
19220 LogError "frmOrganisms", "cmdDelete_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

19230 SaveDetails

End Sub

Private Sub cmdXL_Click()

19240 ExportFlexGrid grdOrg, Me

End Sub

Private Sub Form_DblClick()
                                                                                                '
                                                                                                'Dim tb As Recordset
                                                                                                'Dim sql As String
                                                                                                'Dim x() As String
                                                                                                'Dim s As String
                                                                                                '
                                                                                                'sql = "Select * from Organisms"
                                                                                                'Set tb = New Recordset
                                                                                                'RecOpenServer 0, tb, sql
                                                                                                'Do While Not tb.EOF
                                                                                                '  x = Split(tb!Name & "", " ")
                                                                                                '  If UBound(x) = 1 Then
                                                                                                '    s = LCase$(Left$(x(0), 1)) & "." & _
                                                                                                '        Left$(x(1), 16)
                                                                                                '    tb!ShortName = s
                                                                                                '    tb.Update
                                                                                                '  End If
                                                                                                '  tb.MoveNext
                                                                                                'Loop
                                                                                                '
End Sub

Private Sub Form_Load()

19250 FillGroups

End Sub

Private Sub grdOrg_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer
      Dim xSave As Integer

19260 ySave = grdOrg.row
19270 xSave = grdOrg.Col

19280 grdOrg.Visible = False
19290 grdOrg.Col = 0
19300 For Y = 1 To grdOrg.Rows - 1
19310   grdOrg.row = Y
19320   If grdOrg.CellBackColor = vbYellow Then
19330     For X = 0 To grdOrg.Cols - 1
19340       grdOrg.Col = X
19350       grdOrg.CellBackColor = 0
19360     Next
19370     Exit For
19380   End If
19390 Next
19400 grdOrg.row = ySave

19410 grdOrg.Visible = True

19420 If grdOrg.MouseRow = 0 Then
19430   If SortOrder Then
19440     grdOrg.Sort = flexSortGenericAscending
19450   Else
19460     grdOrg.Sort = flexSortGenericDescending
19470   End If
19480   SortOrder = Not SortOrder
19490   Exit Sub
19500 End If

19510 For X = 0 To grdOrg.Cols - 1
19520   grdOrg.Col = X
19530   grdOrg.CellBackColor = vbYellow
19540 Next

19550 grdOrg.Col = xSave
19560 If grdOrg.Col = 2 Then
19570   grdOrg.Enabled = False
19580   grdOrg = iBOX("Short Name", , grdOrg)
19590   grdOrg.Enabled = True
19600   cmdSave.Enabled = True
19610   cmdSave.Visible = True
19620 ElseIf grdOrg.Col = 3 Then
19630   grdOrg.Enabled = False
19640   grdOrg = iBOX("Report Name", , grdOrg)
19650   grdOrg.Enabled = True
19660   cmdSave.Enabled = True
19670   cmdSave.Visible = True
19680 End If

19690 bMoveUp.Enabled = True
19700 bMoveDown.Enabled = True
19710 cmdDelete.Visible = True

End Sub

Private Sub tmrDown_Timer()

19720 FireDown

End Sub

Private Sub tmrUp_Timer()

19730 FireUp

End Sub


Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)

19740 cmdAdd.Visible = False

19750 If Trim$(txtCode) <> "" Then
19760   If Trim$(txtNewOrg) <> "" Then
19770     cmdAdd.Visible = True
19780   End If
19790 End If

End Sub


Private Sub txtNewOrg_KeyUp(KeyCode As Integer, Shift As Integer)

19800 cmdAdd.Visible = False

19810 If Trim$(txtCode) <> "" Then
19820   If Trim$(txtNewOrg) <> "" Then
19830     cmdAdd.Visible = True
19840   End If
19850 End If

End Sub


