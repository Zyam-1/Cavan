VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmComments 
   Caption         =   "NetAcquire - Comments"
   ClientHeight    =   8730
   ClientLeft      =   795
   ClientTop       =   660
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   10080
   Begin VB.ComboBox cmbListItems 
      Height          =   315
      ItemData        =   "fComments.frx":0000
      Left            =   6960
      List            =   "fComments.frx":0002
      TabIndex        =   22
      Top             =   1500
      Width           =   1875
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      Height          =   825
      Left            =   9180
      Picture         =   "fComments.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3780
      Width           =   795
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8850
      Top             =   5970
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8850
      Top             =   5190
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   9180
      Picture         =   "fComments.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2040
      Width           =   795
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   9180
      Picture         =   "fComments.frx":0F38
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2910
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Clinician"
      Height          =   1305
      Left            =   150
      TabIndex        =   7
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   570
         MaxLength       =   50
         TabIndex        =   10
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   570
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   660
         Width           =   4215
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "&Add"
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   840
         Width           =   315
      End
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move &Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   9180
      Picture         =   "fComments.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move &Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   9180
      Picture         =   "fComments.frx":19E4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   795
   End
   Begin VB.CommandButton bSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   9180
      Picture         =   "fComments.frx":1E26
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7650
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Type"
      Height          =   1305
      Left            =   5190
      TabIndex        =   0
      Top             =   120
      Width           =   4125
      Begin VB.OptionButton optType 
         Caption         =   "Film Comments"
         Height          =   225
         Index           =   7
         Left            =   2340
         TabIndex        =   20
         Top             =   990
         Width           =   1545
      End
      Begin VB.OptionButton optType 
         Caption         =   "Micro Comments"
         Height          =   225
         Index           =   6
         Left            =   2340
         TabIndex        =   19
         Top             =   750
         Width           =   1545
      End
      Begin VB.OptionButton optType 
         Caption         =   "Clinical Details"
         Height          =   225
         Index           =   5
         Left            =   2340
         TabIndex        =   18
         Top             =   510
         Width           =   1335
      End
      Begin VB.OptionButton optType 
         Caption         =   "Semen Comments"
         Height          =   225
         Index           =   4
         Left            =   2340
         TabIndex        =   17
         Top             =   270
         Width           =   1605
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Caption         =   "Demographic Comments"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   16
         Top             =   990
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Caption         =   "Biochemistry Comments"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Caption         =   "Haematology Comments"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   510
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Caption         =   "Coagulation Comments"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   1
         Top             =   750
         Width           =   2115
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6795
      Left            =   180
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1830
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11986
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   3
      FormatString    =   "<Code   |Text                                                                                                              |Inuse"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Items to be displayed in list"
      Height          =   195
      Left            =   5040
      TabIndex        =   23
      Top             =   1560
      Width           =   1875
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Integer






Private Sub cmbListItems_Click()
32370     bsave.Enabled = True
End Sub

Private Sub cmbListItems_KeyPress(KeyAscii As Integer)
32380     KeyAscii = 0
End Sub
Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

32390     If g.row = g.Rows - 1 Then Exit Sub
32400     n = g.row

32410     FireCounter = FireCounter + 1
32420     If FireCounter > 5 Then
32430         tmrDown.Interval = 100
32440     End If

32450     VisibleRows = g.height \ g.RowHeight(1) - 1

32460     g.Visible = False

32470     s = ""
32480     For X = 0 To g.Cols - 1
32490         s = s & g.TextMatrix(n, X) & vbTab
32500     Next
32510     s = Left$(s, Len(s) - 1)

32520     g.RemoveItem n
32530     If n < g.Rows Then
32540         g.AddItem s, n + 1
32550         g.row = n + 1
32560     Else
32570         g.AddItem s
32580         g.row = g.Rows - 1
32590     End If

32600     For X = 0 To g.Cols - 1
32610         g.Col = X
32620         g.CellBackColor = vbYellow
32630     Next

32640     If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
32650         If g.row - VisibleRows + 1 > 0 Then
32660             g.TopRow = g.row - VisibleRows + 1
32670         End If
32680     End If

32690     g.Visible = True

32700     bsave.Visible = True

End Sub
Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

32710     If g.row = 1 Then Exit Sub

32720     FireCounter = FireCounter + 1
32730     If FireCounter > 5 Then
32740         tmrUp.Interval = 100
32750     End If

32760     n = g.row

32770     g.Visible = False

32780     s = ""
32790     For X = 0 To g.Cols - 1
32800         s = s & g.TextMatrix(n, X) & vbTab
32810     Next
32820     s = Left$(s, Len(s) - 1)

32830     g.RemoveItem n
32840     g.AddItem s, n - 1

32850     g.row = n - 1
32860     For X = 0 To g.Cols - 1
32870         g.Col = X
32880         g.CellBackColor = vbYellow
32890     Next

32900     If Not g.RowIsVisible(g.row) Then
32910         g.TopRow = g.row
32920     End If

32930     g.Visible = True

32940     bsave.Visible = True

End Sub




Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim LT As String
          Dim s As String

32950     On Error GoTo FillG_Error

32960     LT = Switch(optType(0), "BI", _
              optType(1), "HA", _
              optType(2), "CO", _
              optType(3), "DE", _
              optType(4), "SE", _
              optType(5), "CD", _
              optType(6), "BA", _
              optType(7), "FI")

32970     g.Rows = 2
32980     g.AddItem ""
32990     g.RemoveItem 1

33000     sql = "Select * from Lists where " & _
              "ListType = '" & LT & "' " & _
              "order by ListOrder"
33010     Set tb = New Recordset
33020     RecOpenServer 0, tb, sql
33030     Do While Not tb.EOF
33040         s = tb!Code & vbTab & tb!Text & vbTab & _
                  IIf(tb!InUse, "Yes", "No")
33050         g.AddItem s
33060         tb.MoveNext
33070     Loop

33080     If g.Rows > 2 Then
33090         g.RemoveItem 1
33100     End If

33110     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

33120     intEL = Erl
33130     strES = Err.Description
33140     LogError "frmComments", "FillG", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

33150     txtCode = Trim$(UCase$(txtCode))
33160     txtText = Trim$(txtText)

33170     If txtCode = "" Then
33180         Exit Sub
33190     End If

33200     If txtText = "" Then Exit Sub

33210     g.AddItem txtCode & vbTab & txtText & vbTab & "Yes"

33220     txtCode = ""
33230     txtText = ""

33240     bsave.Enabled = True

End Sub


Private Sub bcancel_Click()

33250     Unload Me

End Sub


Private Sub bMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

33260     FireDown

33270     tmrDown.Interval = 250
33280     FireCounter = 0

33290     tmrDown.Enabled = True

End Sub


Private Sub bMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

33300     tmrDown.Enabled = False

End Sub


Private Sub bMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

33310     FireUp

33320     tmrUp.Interval = 250
33330     FireCounter = 0

33340     tmrUp.Enabled = True

End Sub


Private Sub bMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

33350     tmrUp.Enabled = False

End Sub


Private Sub bPrint_Click()

          Dim LT As String

33360     LT = Switch(optType(0), "Biochemistry Comments.", _
              optType(1), "Haematology Comments.", _
              optType(2), "Coagulation Comments.", _
              optType(3), "Demographic Comments.", _
              optType(4), "Semen Comments.", _
              optType(5), "Clinical Details.", _
              optType(6), "Microbiology Comments.", _
              optType(7), "Film Comments")

33370     Printer.Print

33380     Printer.Print "List of "; LT

33390     g.Col = 0
33400     g.row = 1
33410     g.ColSel = g.Cols - 1
33420     g.RowSel = g.Rows - 1

33430     Printer.Print g.Clip

33440     Printer.EndDoc
33450     Screen.MousePointer = 0

End Sub


Private Sub bSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim LT As String
          Dim Y As Integer

33460     On Error GoTo bSave_Click_Error

33470     LT = Switch(optType(0), "BI", _
              optType(1), "HA", _
              optType(2), "CO", _
              optType(3), "DE", _
              optType(4), "SE", _
              optType(5), "CD", _
              optType(6), "BA", _
              optType(7), "FI")

33480     For Y = 1 To g.Rows - 1
33490         If g.TextMatrix(Y, 0) <> "" Then
        
33500             sql = "Select * from Lists where " & _
                      "ListType = '" & LT & "' " & _
                      "and Code = '" & AddTicks(g.TextMatrix(Y, 0)) & "'"
33510             Set tb = New Recordset
33520             RecOpenServer 0, tb, sql
33530             If tb.EOF Then
33540                 tb.AddNew
33550             End If
33560             tb!Code = g.TextMatrix(Y, 0)
33570             tb!ListType = LT
33580             tb!Text = g.TextMatrix(Y, 1)
33590             tb!InUse = (g.TextMatrix(Y, 2) = "Yes")
33600             tb!ListOrder = Y
33610             tb.Update
33620         End If
33630     Next

33640     Call SaveOptionSetting("CommentListLength", cmbListItems)

33650     FillG

33660     txtCode = ""
33670     txtText = ""
33680     txtCode.SetFocus
33690     bMoveUp.Enabled = False
33700     bMoveDown.Enabled = False
33710     bsave.Enabled = False

33720     Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

33730     intEL = Erl
33740     strES = Err.Description
33750     LogError "frmComments", "bSave_Click", intEL, strES, sql


End Sub


Private Sub cmdDelete_Click()

          Dim LT As String
33760     On Error GoTo cmdDelete_Click_Error

33770     If g.row = 0 Or g.Rows <= 2 Then Exit Sub

33780     LT = Switch(optType(0), "BI", _
              optType(1), "HA", _
              optType(2), "CO", _
              optType(3), "DE", _
              optType(4), "SE", _
              optType(5), "CD", _
              optType(6), "BA", _
              optType(7), "FI")

33790     If iMsg("Are you sure you wanted to delete selected comment?", vbYesNo) = vbYes Then
33800         Cnxn(0).Execute "DELETE FROM Lists WHERE ListType = '" & LT & "' " & _
                  "AND Code = '" & AddTicks(g.TextMatrix(g.row, 0)) & "' " & _
                  "AND Text = '" & AddTicks(g.TextMatrix(g.row, 1)) & "'"
33810         FillG
33820     End If


33830     Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

33840     intEL = Erl
33850     strES = Err.Description
33860     LogError "frmComments", "cmdDelete_Click", intEL, strES

End Sub

Private Sub Form_Activate()

33870     If Activated Then Exit Sub

33880     Activated = True

33890     FillG

End Sub

Private Sub Form_Load()

33900     Activated = False

33910     g.Font.Bold = True
        
33920     If sysOptDeptSemen(0) Then
33930         optType(4).Enabled = True
33940     Else
33950         optType(4).Enabled = False
33960     End If

          Dim i  As Integer
33970     cmbListItems.Clear
33980     For i = 8 To 32 Step 8
33990         cmbListItems.AddItem i
34000     Next i
34010     cmbListItems.Text = GetOptionSetting("CommentListLength", 8)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

34020     If bsave.Enabled Then
34030         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
34040             Cancel = True
34050             Exit Sub
34060         End If
34070     End If

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer

34080     ySave = g.row

34090     g.Visible = False
34100     g.Col = 0
34110     For Y = 1 To g.Rows - 1
34120         g.row = Y
34130         If g.CellBackColor = vbYellow Then
34140             For X = 0 To g.Cols - 1
34150                 g.Col = X
34160                 g.CellBackColor = 0
34170             Next
34180             Exit For
34190         End If
34200     Next
34210     g.row = ySave
34220     g.Visible = True

34230     If g.MouseRow = 0 Then
34240         If SortOrder Then
34250             g.Sort = flexSortGenericAscending
34260         Else
34270             g.Sort = flexSortGenericDescending
34280         End If
34290         SortOrder = Not SortOrder
34300         Exit Sub
34310     ElseIf g.MouseCol = 2 Then
34320         If g.TextMatrix(g.RowSel, 2) = "No" Then
34330             g.TextMatrix(g.RowSel, 2) = "Yes"
34340         Else
34350             g.TextMatrix(g.RowSel, 2) = "No"
34360         End If
34370         bsave.Enabled = True
34380     ElseIf g.MouseCol = 1 Then
              Dim Answer As String
34390         Answer = iBOX("Please enter new comment")
34400         If Answer <> "" Then
34410             g.TextMatrix(g.row, 1) = Answer
34420             bsave.Enabled = True
34430         End If
34440     End If

34450     For X = 0 To g.Cols - 1
34460         g.Col = X
34470         g.CellBackColor = vbYellow
34480     Next

34490     bMoveUp.Enabled = True
34500     bMoveDown.Enabled = True

End Sub


Private Sub optType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

34510     FillG

34520     FrameAdd.Caption = "Add New " & Left$(optType(Index).Caption, Len(optType(Index).Caption) - 1)

34530     txtCode = ""
34540     txtText = ""
34550     If txtCode.Visible Then
34560         txtCode.SetFocus
34570     End If

End Sub


Private Sub tmrDown_Timer()

34580     FireDown

End Sub

Private Sub tmrUp_Timer()

34590     FireUp

End Sub

