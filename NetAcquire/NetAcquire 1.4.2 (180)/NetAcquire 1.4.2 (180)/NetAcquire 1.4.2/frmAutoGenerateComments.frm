VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoGenerateComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Auto-Generate Comments"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   13800
      Top             =   3210
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   13800
      Top             =   4920
   End
   Begin VB.CommandButton cmdMoveDown 
      Height          =   615
      Left            =   13740
      Picture         =   "frmAutoGenerateComments.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4290
      Width           =   525
   End
   Begin VB.CommandButton cmdMoveUp 
      Height          =   615
      Left            =   13740
      Picture         =   "frmAutoGenerateComments.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3660
      Width           =   525
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   1185
      Left            =   10110
      Picture         =   "frmAutoGenerateComments.frx":3304
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1005
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3885
      Left            =   2250
      TabIndex        =   7
      Top             =   3090
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   6853
      _Version        =   393216
      Cols            =   6
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmAutoGenerateComments.frx":41CE
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1185
      Left            =   12720
      Picture         =   "frmAutoGenerateComments.frx":4292
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1185
      Left            =   11430
      Picture         =   "frmAutoGenerateComments.frx":515C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1005
   End
   Begin VB.ListBox lstParameter 
      Height          =   8190
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   1965
   End
   Begin VB.Frame Frame1 
      Height          =   2385
      Left            =   2250
      TabIndex        =   5
      Top             =   570
      Width           =   11445
      Begin VB.Frame fraAlpha 
         Caption         =   "Alphanumeric Results"
         Height          =   765
         Left            =   5970
         TabIndex        =   17
         Top             =   270
         Width           =   5085
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
            Height          =   285
            Left            =   3510
            TabIndex        =   24
            Top             =   300
            Width           =   1035
         End
         Begin VB.ComboBox cmbText 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   23
            Text            =   "cmbText"
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "If this Result"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   330
            Width           =   1095
         End
      End
      Begin VB.Frame fraNumeric 
         BackColor       =   &H8000000A&
         Caption         =   "Numeric Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   765
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Width           =   5385
         Begin VB.TextBox txtCriteria 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   3030
            TabIndex        =   20
            Top             =   300
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox txtCriteria 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   4110
            TabIndex        =   19
            Top             =   300
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.ComboBox cmbCriteria 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmAutoGenerateComments.frx":6026
            Left            =   1470
            List            =   "frmAutoGenerateComments.frx":6028
            TabIndex        =   18
            Text            =   "cmbCriteria"
            Top             =   300
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "If this result is"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   315
         Left            =   2550
         TabIndex        =   10
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   326172673
         CurrentDate     =   40093
      End
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   315
         Left            =   4380
         TabIndex        =   9
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   326172673
         CurrentDate     =   40093
      End
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         MaxLength       =   95
         TabIndex        =   1
         Top             =   1410
         Width           =   10335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "(Maximum 95 characters)"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2370
         TabIndex        =   15
         Top             =   1170
         Width           =   1770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "and"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   1920
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "This rule is active between"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1920
         Width           =   2310
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Print this comment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   1170
         Width           =   1575
      End
   End
   Begin VB.Label lblDiscipline 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biochemistry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7560
      TabIndex        =   8
      Top             =   150
      Width           =   1830
   End
End
Attribute VB_Name = "frmAutoGenerateComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String
Private pDisc As String

Private FireCounter As Integer

Private EntryMode As String
Private Sub ClearHighlight()

          Dim SaveY As Integer
          Dim Y As Integer
          Dim X As Integer

59160     SaveY = g.MouseRow

59170     For Y = 1 To g.Rows - 1
59180         g.Col = 1
59190         g.row = Y
59200         If g.CellBackColor = vbYellow Then
59210             For X = 1 To g.Cols - 1
59220                 g.Col = X
59230                 g.CellBackColor = 0
59240             Next
59250         End If
59260     Next

59270     g.row = SaveY

End Sub

Private Sub FillList()

          Dim tb As Recordset
          Dim sql As String

59280     On Error GoTo FillList_Error

59290     lstParameter.Clear

59300     sql = "SELECT DISTINCT ShortName FROM " & pDisc & "TestDefinitions " & _
              "ORDER BY ShortName"
59310     Set tb = New Recordset
59320     RecOpenServer 0, tb, sql
59330     Do While Not tb.EOF
59340         lstParameter.AddItem tb!ShortName & ""
59350         tb.MoveNext
59360     Loop

59370     Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

59380     intEL = Erl
59390     strES = Err.Description
59400     LogError "frmAutoGenerateComments", "FillList", intEL, strES, sql

End Sub


Private Sub SaveListOrder()

          Dim Y As Integer
          Dim sql As String

59410     On Error GoTo SaveListOrder_Error

59420     For Y = 1 To g.Rows - 1
59430         If g.TextMatrix(Y, 0) <> "" Then
59440             sql = "UPDATE AutoComments " & _
                      "SET ListOrder = '" & Y & "' " & _
                      "WHERE Parameter = '" & lstParameter & "' " & _
                      "AND Comment = '" & AddTicks(g.TextMatrix(Y, 3)) & "'"
59450             Cnxn(0).Execute sql
59460         End If
59470     Next

59480     Exit Sub

SaveListOrder_Error:

          Dim strES As String
          Dim intEL As Integer

59490     intEL = Erl
59500     strES = Err.Description
59510     LogError "frmAutoGenerateComments", "SaveListOrder", intEL, strES, sql


End Sub

Private Sub SetToNumeric()

59520     fraNumeric.FontBold = True
59530     fraNumeric.ForeColor = vbRed
59540     fraAlpha.FontBold = False
59550     fraAlpha.ForeColor = vbBlack
59560     EntryMode = "Numeric"

End Sub

Private Sub SetToAlpha()

59570     fraNumeric.FontBold = False
59580     fraNumeric.ForeColor = vbBlack
59590     fraAlpha.FontBold = True
59600     fraAlpha.ForeColor = vbRed
59610     EntryMode = "Alpha"

End Sub

Private Sub cmbCriteria_Click()

59620     ClearHighlight
59630     cmdDelete.Enabled = False

59640     Select Case cmbCriteria
              Case "Present": txtCriteria(0).Visible = False: txtCriteria(1).Visible = False
59650         Case "Equal to": txtCriteria(0).Visible = True
59660         Case "Greater than": txtCriteria(0).Visible = True
59670         Case "Less than": txtCriteria(0).Visible = True
59680         Case "Between": txtCriteria(0).Visible = True: txtCriteria(1).Visible = True
59690         Case "Not between": txtCriteria(0).Visible = True: txtCriteria(1).Visible = True
59700     End Select

End Sub


Private Sub cmbCriteria_GotFocus()

59710     SetToNumeric

End Sub


Private Sub cmbCriteria_KeyPress(KeyAscii As Integer)

59720     KeyAscii = 0

End Sub


Private Sub cmbText_GotFocus()

59730     SetToAlpha

End Sub


Private Sub cmbText_KeyPress(KeyAscii As Integer)

59740     KeyAscii = 0

End Sub


Private Sub cmdCancel_Click()

59750     Unload Me

End Sub

Private Sub cmdDelete_Click()

          Dim sql As String

59760     On Error GoTo cmdDelete_Click_Error

59770     sql = "DELETE FROM AutoComments WHERE " & _
              "Discipline = '" & pDiscipline & "' " & _
              "AND Parameter = '" & lstParameter & "' " & _
              "AND Criteria = '" & g.TextMatrix(g.row, 0) & "' " & _
              "AND Comment = '" & g.TextMatrix(g.row, 3) & "'"
59780     Cnxn(0).Execute sql

59790     FillG

59800     cmdSave.Enabled = False

59810     Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

59820     intEL = Erl
59830     strES = Err.Description
59840     LogError "frmAutoGenerateComments", "cmdDelete_Click", intEL, strES, sql

End Sub
Private Sub ClearEntry()

59850     cmbCriteria.ListIndex = 0
59860     txtCriteria(0) = ""
59870     txtCriteria(1) = ""
59880     txtComment = ""

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

59890     FireDown

59900     tmrDown.Interval = 250
59910     FireCounter = 0

59920     tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

59930     tmrDown.Enabled = False

59940     cmdSave.Enabled = True

End Sub


Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

59950     If g.row = g.Rows - 1 Or g.row = 0 Then Exit Sub
59960     n = g.row

59970     FireCounter = FireCounter + 1
59980     If FireCounter > 5 Then
59990         tmrDown.Interval = 100
60000     End If

60010     VisibleRows = g.height \ g.RowHeight(1) - 1

60020     g.Visible = False

60030     s = ""
60040     For X = 0 To g.Cols - 1
60050         s = s & g.TextMatrix(n, X) & vbTab
60060     Next
60070     s = Left$(s, Len(s) - 1)

60080     g.RemoveItem n
60090     If n < g.Rows Then
60100         g.AddItem s, n + 1
60110         g.row = n + 1
60120     Else
60130         g.AddItem s
60140         g.row = g.Rows - 1
60150     End If

60160     For X = 0 To g.Cols - 1
60170         g.Col = X
60180         g.CellBackColor = vbYellow
60190     Next

60200     If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
60210         If g.row - VisibleRows + 1 > 0 Then
60220             g.TopRow = g.row - VisibleRows + 1
60230         End If
60240     End If

60250     g.Visible = True

60260     cmdSave.Visible = True

End Sub

Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

60270     FireUp

60280     tmrUp.Interval = 250
60290     FireCounter = 0

60300     tmrUp.Enabled = True

End Sub


Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

60310     If g.row < 2 Then Exit Sub

60320     FireCounter = FireCounter + 1
60330     If FireCounter > 5 Then
60340         tmrUp.Interval = 100
60350     End If

60360     n = g.row

60370     g.Visible = False

60380     s = ""
60390     For X = 0 To g.Cols - 1
60400         s = s & g.TextMatrix(n, X) & vbTab
60410     Next
60420     s = Left$(s, Len(s) - 1)

60430     g.RemoveItem n
60440     g.AddItem s, n - 1

60450     g.row = n - 1
60460     For X = 0 To g.Cols - 1
60470         g.Col = X
60480         g.CellBackColor = vbYellow
60490     Next

60500     If Not g.RowIsVisible(g.row) Then
60510         g.TopRow = g.row
60520     End If

60530     g.Visible = True

60540     cmdSave.Visible = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

60550     tmrUp.Enabled = False

60560     cmdSave.Enabled = True

End Sub


Private Sub cmdSave_Click()

          Dim sql As String
          Dim Value0 As String
          Dim Value1 As String
          Dim Criteria As String

60570     On Error GoTo cmdSave_Click_Error

60580     If Trim$(txtComment) <> "" Then
60590         If EntryMode = "Numeric" Then
60600             Criteria = cmbCriteria
60610             If IsNumeric(txtCriteria(0)) Then
60620                 If cmbCriteria = "Greater than" Then
60630                     Value0 = Val(txtCriteria(0)) '- 0.00001
60640                 ElseIf cmbCriteria = "Not between" Or _
                          cmbCriteria = "Between" Or _
                          cmbCriteria = "Less than" Or _
                          cmbCriteria = "Equal to" Then
60650                     Value0 = Val(txtCriteria(0)) '+ 0.00001
60660                 End If
60670             Else
60680                 Value0 = ""
60690             End If
60700             If IsNumeric(txtCriteria(1)) Then
60710                 Value1 = Val(txtCriteria(1))
60720             Else
60730                 Value1 = ""
60740             End If
60750         ElseIf EntryMode = "Alpha" Then
60760             Criteria = cmbText
60770             Value0 = txtText
60780             Value1 = ""
60790         End If

60800         sql = "INSERT INTO AutoComments " & _
                  "(Discipline, Parameter, Criteria, Value0, Value1, Comment, DateStart, DateEnd) " & _
                  "VALUES " & _
                  "( '" & pDiscipline & "', " & _
                  "  '" & lstParameter & "', " & _
                  "  '" & Criteria & "', " & _
                  "  '" & AddTicks(Value0) & "', " & _
                  "  '" & AddTicks(Value1) & "', " & _
                  "  '" & AddTicks(txtComment) & "', " & _
                  "  '" & Format$(dtStart, "dd/MMM/yyyy") & "', " & _
                  "  '" & Format$(dtEnd, "dd/MMM/yyyy") & "')"
60810         Cnxn(0).Execute sql

60820         SaveListOrder

60830     Else
60840         txtComment.BackColor = vbRed
60850     End If

60860     cmdSave.Enabled = False

60870     FillG

60880     ClearEntry

60890     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60900     intEL = Erl
60910     strES = Err.Description
60920     LogError "frmAutoGenerateComments", "cmdSave_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()
        
60930     lblDiscipline = pDiscipline

60940     FillList

60950     ClearHighlight
60960     cmdDelete.Enabled = False

End Sub

Private Sub Form_Load()

60970     CheckAutoCommentsInDb

60980     EnsureColumnExists "AutoComments", "DateStart", "smalldatetime"
60990     EnsureColumnExists "AutoComments", "DateEnd", "smalldatetime"
61000     EnsureColumnExists "AutoComments", "ListOrder", "tinyint"

61010     cmbCriteria.AddItem "Present"
61020     cmbCriteria.AddItem "Equal to"
61030     cmbCriteria.AddItem "Less than"
61040     cmbCriteria.AddItem "Greater than"
61050     cmbCriteria.AddItem "Between"
61060     cmbCriteria.AddItem "Not between"
61070     cmbCriteria.ListIndex = 0

61080     cmbText.AddItem "Starts with"
61090     cmbText.AddItem "Contains Text"
61100     cmbText.ListIndex = 0

61110     dtStart = Format$(Now, "dd/MM/yyyy")
61120     dtEnd = Format$(DateAdd("m", 6, Now), "dd/MM/yyyy")

61130     EntryMode = "Numeric"

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

61140     If cmdSave.Enabled Then
61150         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
61160             Cancel = True
61170         End If
61180     End If

End Sub


Private Sub g_Click()

          Dim X As Integer

61190     ClearHighlight
61200     cmdDelete.Enabled = False

61210     If g.MouseRow > 0 And g.TextMatrix(g.MouseRow, 0) <> "" Then
       
61220         For X = 1 To g.Cols - 1
61230             g.Col = X
61240             g.CellBackColor = vbYellow
61250         Next
        
61260         cmdDelete.Enabled = True
61270     End If

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

61280     On Error GoTo FillG_Error

61290     g.Rows = 2
61300     g.AddItem ""
61310     g.RemoveItem 1

61320     sql = "SELECT * FROM AutoComments WHERE " & _
              "Discipline = '" & pDiscipline & "' " & _
              "AND Parameter = '" & lstParameter & "' " & _
              "ORDER BY ListOrder"
61330     Set tb = New Recordset
61340     RecOpenServer 0, tb, sql
61350     Do While Not tb.EOF
61360         s = tb!Criteria & vbTab & _
                  tb!Value0 & vbTab & _
                  tb!Value1 & vbTab & _
                  tb!Comment & vbTab
61370         If Not IsNull(tb!DateStart) Then
61380             s = s & Format$(tb!DateStart, "dd/MM/yyyy")
61390         End If
61400         s = s & vbTab
61410         If Not IsNull(tb!DateEnd) Then
61420             s = s & Format$(tb!DateEnd, "dd/MM/yyyy")
61430         End If

61440         g.AddItem s
61450         tb.MoveNext
61460     Loop

61470     If g.Rows > 2 Then
61480         g.RemoveItem 1
61490     End If

61500     ClearHighlight
61510     cmdDelete.Enabled = False

61520     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

61530     intEL = Erl
61540     strES = Err.Description
61550     LogError "frmAutoGenerateComments", "FillG", intEL, strES, sql

End Sub


Private Sub lstParameter_Click()

61560     FillG

61570     ClearEntry

End Sub

Private Sub txtComment_KeyUp(KeyCode As Integer, Shift As Integer)

61580     ClearHighlight
61590     cmdDelete.Enabled = False
61600     txtComment.BackColor = vbWhite

61610     cmdSave.Enabled = Len(Trim$(txtComment)) > 0

End Sub

Private Sub txtCriteria_Change(Index As Integer)

61620     ClearHighlight
61630     cmdDelete.Enabled = False

End Sub



Public Property Let Discipline(ByVal sNewValue As String)

61640     pDiscipline = sNewValue

61650     If pDiscipline = "Biochemistry" Then
61660         pDisc = "Bio"
61670     ElseIf pDiscipline = "Coagulation" Then
61680         pDisc = "Coag"
61690     End If

End Property

Private Sub txtCriteria_GotFocus(Index As Integer)

61700     SetToNumeric

End Sub

Private Sub txtText_GotFocus()
61710     SetToAlpha
End Sub

