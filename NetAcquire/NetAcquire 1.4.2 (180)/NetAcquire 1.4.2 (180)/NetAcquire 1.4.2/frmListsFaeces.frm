VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListsFaeces 
   Caption         =   "NetAcquire"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   8535
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   8535
   Begin VB.Frame Frame1 
      Caption         =   "List Type"
      Height          =   1065
      Left            =   5010
      TabIndex        =   11
      Top             =   330
      Width           =   1695
      Begin VB.OptionButton o 
         Caption         =   "SMAC"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   480
         Width           =   765
      End
      Begin VB.OptionButton o 
         Caption         =   "XLD/DCA"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton o 
         Caption         =   "Preston/CCDA"
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   720
         Width           =   1365
      End
   End
   Begin VB.CommandButton bSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   7260
      Picture         =   "frmListsFaeces.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7500
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7260
      Picture         =   "frmListsFaeces.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   795
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7260
      Picture         =   "frmListsFaeces.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Item"
      Height          =   1365
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   4365
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Top             =   330
         Width           =   645
      End
      Begin VB.TextBox tText 
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
         Left            =   660
         MaxLength       =   50
         TabIndex        =   4
         Top             =   900
         Width           =   3495
      End
      Begin VB.TextBox tCode 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   3
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   390
         Width           =   375
      End
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   7260
      Picture         =   "frmListsFaeces.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   795
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   7260
      Picture         =   "frmListsFaeces.frx":1558
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2370
      Width           =   795
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6990
      Top             =   6150
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6990
      Top             =   5370
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6675
      Left            =   180
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1650
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11774
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Code   |Text                                                                                      "
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
End
Attribute VB_Name = "frmListsFaeces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Integer
Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim LT As String
          Dim s As String

11900     On Error GoTo FillG_Error

11910     LT = Switch(o(0), "FX", _
              o(1), "FS", _
              o(2), "FP")

11920     g.Rows = 2
11930     g.AddItem ""
11940     g.RemoveItem 1

11950     sql = "Select * from Lists where " & _
              "ListType = '" & LT & "' and InUse = 1 order by ListOrder"
11960     Set tb = New Recordset
11970     RecOpenServer 0, tb, sql
11980     Do While Not tb.EOF
11990         s = tb!Code & vbTab & tb!Text & ""
12000         g.AddItem s
12010         tb.MoveNext
12020     Loop

12030     If g.Rows > 2 Then
12040         g.RemoveItem 1
12050     End If

12060     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

12070     intEL = Erl
12080     strES = Err.Description
12090     LogError "frmListsFaeces", "FillG", intEL, strES, sql


End Sub

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

12100     If g.row = g.Rows - 1 Then Exit Sub
12110     n = g.row

12120     VisibleRows = g.height \ g.RowHeight(1) - 1

12130     FireCounter = FireCounter + 1
12140     If FireCounter > 5 Then
12150         tmrDown.Interval = 100
12160     End If

12170     g.Visible = False

12180     s = ""
12190     For X = 0 To g.Cols - 1
12200         s = s & g.TextMatrix(n, X) & vbTab
12210     Next
12220     s = Left$(s, Len(s) - 1)

12230     g.RemoveItem n
12240     If n < g.Rows Then
12250         g.AddItem s, n + 1
12260         g.row = n + 1
12270     Else
12280         g.AddItem s
12290         g.row = g.Rows - 1
12300     End If

12310     For X = 0 To g.Cols - 1
12320         g.Col = X
12330         g.CellBackColor = vbYellow
12340     Next

12350     If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
12360         If g.row - VisibleRows + 1 > 0 Then
12370             g.TopRow = g.row - VisibleRows + 1
12380         End If
12390     End If

12400     g.Visible = True

12410     bsave.Visible = True

End Sub

Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

12420     If g.row = 1 Then Exit Sub

12430     FireCounter = FireCounter + 1
12440     If FireCounter > 5 Then
12450         tmrUp.Interval = 100
12460     End If

12470     n = g.row

12480     g.Visible = False

12490     s = ""
12500     For X = 0 To g.Cols - 1
12510         s = s & g.TextMatrix(n, X) & vbTab
12520     Next
12530     s = Left$(s, Len(s) - 1)

12540     g.RemoveItem n
12550     g.AddItem s, n - 1

12560     g.row = n - 1
12570     For X = 0 To g.Cols - 1
12580         g.Col = X
12590         g.CellBackColor = vbYellow
12600     Next

12610     If Not g.RowIsVisible(g.row) Then
12620         g.TopRow = g.row
12630     End If

12640     g.Visible = True

12650     bsave.Visible = True

End Sub


Private Sub bAdd_Click()

12660     tCode = Trim$(UCase$(tCode))
12670     tText = Trim$(tText)

12680     If tCode = "" Then
12690         Exit Sub
12700     End If

12710     If tText = "" Then Exit Sub

12720     g.AddItem tCode & vbTab & tText

12730     tCode = ""
12740     tText = ""

12750     bsave.Visible = True

End Sub


Private Sub bcancel_Click()

12760     Unload Me

End Sub

Private Sub bMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12770     FireDown

12780     tmrDown.Interval = 250
12790     FireCounter = 0

12800     tmrDown.Enabled = True

End Sub

Private Sub bMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12810     tmrDown.Enabled = False

End Sub



Private Sub bMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12820     FireUp

12830     tmrUp.Interval = 250
12840     FireCounter = 0

12850     tmrUp.Enabled = True

End Sub



Private Sub bMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12860     tmrUp.Enabled = False

End Sub


Private Sub bPrint_Click()

          Dim LT As String

12870     LT = Switch(o(0), "XLD/DCA.", _
              o(1), "SMAC.", _
              o(2), "Preston/CCDA.")

12880     Printer.Print

12890     Printer.Print "List of "; LT

12900     g.Col = 0
12910     g.row = 1
12920     g.ColSel = g.Cols - 1
12930     g.RowSel = g.Rows - 1

12940     Printer.Print g.Clip

12950     Printer.EndDoc
12960     Screen.MousePointer = 0

End Sub

Private Sub bSave_Click()

          Dim LT As String
          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String

12970     On Error GoTo bSave_Click_Error

12980     LT = Switch(o(0), "FX", _
              o(1), "FS", _
              o(2), "FP")

12990     For Y = 1 To g.Rows - 1
13000         sql = "Select * from Lists where " & _
                  "ListType = '" & LT & "' " & _
                  "and Code = '" & g.TextMatrix(Y, 0) & "' and InUse = 1"
13010         Set tb = New Recordset
13020         RecOpenServer 0, tb, sql
13030         If tb.EOF Then
13040             tb.AddNew
13050         End If
13060         tb!Code = g.TextMatrix(Y, 0)
13070         tb!ListType = LT
13080         tb!Text = g.TextMatrix(Y, 1)
13090         tb!ListOrder = Y
13100         tb!InUse = 1
13110         tb.Update
13120     Next

13130     FillG

13140     tCode = ""
13150     tText = ""
13160     tCode.SetFocus
13170     bMoveUp.Enabled = False
13180     bMoveDown.Enabled = False
13190     bsave.Visible = False

13200     Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

13210     intEL = Erl
13220     strES = Err.Description
13230     LogError "frmListsFaeces", "bSave_Click", intEL, strES, sql


End Sub



Private Sub Form_Activate()

13240     If Activated Then Exit Sub

13250     Activated = True

13260     FillG

End Sub

Private Sub Form_Load()

13270     g.Font.Bold = True

13280     Activated = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

13290     If bsave.Visible Then
13300         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
13310             Cancel = True
13320             Exit Sub
13330         End If
13340     End If

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer

13350     ySave = g.row

13360     g.Visible = False
13370     g.Col = 0
13380     For Y = 1 To g.Rows - 1
13390         g.row = Y
13400         If g.CellBackColor = vbYellow Then
13410             For X = 0 To g.Cols - 1
13420                 g.Col = X
13430                 g.CellBackColor = 0
13440             Next
13450             Exit For
13460         End If
13470     Next
13480     g.row = ySave
13490     g.Visible = True

13500     If g.MouseRow = 0 Then
13510         If SortOrder Then
13520             g.Sort = flexSortGenericAscending
13530         Else
13540             g.Sort = flexSortGenericDescending
13550         End If
13560         SortOrder = Not SortOrder
13570         Exit Sub
13580     End If

13590     For X = 0 To g.Cols - 1
13600         g.Col = X
13610         g.CellBackColor = vbYellow
13620     Next

13630     bMoveUp.Enabled = True
13640     bMoveDown.Enabled = True

End Sub

Private Sub o_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

13650     FillG

13660     FrameAdd.Caption = "Add New " & o(Index).Caption

13670     tCode = ""
13680     tText = ""
13690     If tCode.Visible Then
13700         tCode.SetFocus
13710     End If

End Sub


Private Sub tmrDown_Timer()

13720     FireDown

End Sub


Private Sub tmrUp_Timer()

13730     FireUp

End Sub



