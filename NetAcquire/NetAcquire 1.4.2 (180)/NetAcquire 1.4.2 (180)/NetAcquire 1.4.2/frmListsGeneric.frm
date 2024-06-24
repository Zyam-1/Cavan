VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListsGeneric 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - List of Generic"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   1065
      Left            =   7170
      Picture         =   "frmListsGeneric.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3810
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   7170
      Picture         =   "frmListsGeneric.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7230
      Width           =   945
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Generic"
      Height          =   975
      Left            =   90
      TabIndex        =   10
      Top             =   120
      Width           =   6735
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
         Left            =   180
         MaxLength       =   15
         TabIndex        =   0
         Top             =   420
         Width           =   1545
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
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   420
         Width           =   4065
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   825
         Left            =   6060
         Picture         =   "frmListsGeneric.frx":1D94
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   210
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   1830
         TabIndex        =   11
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   615
      Left            =   6840
      Picture         =   "frmListsGeneric.frx":3716
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   525
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   765
      Left            =   6840
      Picture         =   "frmListsGeneric.frx":5098
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2190
      Width           =   525
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   1065
      Left            =   7170
      Picture         =   "frmListsGeneric.frx":6A1A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6090
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   1065
      Left            =   7170
      Picture         =   "frmListsGeneric.frx":78E4
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4950
      Width           =   945
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7380
      Top             =   2460
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7380
      Top             =   1680
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   945
      Left            =   7170
      Picture         =   "frmListsGeneric.frx":87AE
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7125
      Left            =   90
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1170
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12568
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
      FormatString    =   "<Code                       |<Text                                                                  "
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
      Height          =   285
      Left            =   6930
      TabIndex        =   14
      Top             =   1170
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmListsGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Integer

Private pListTypeNames As String
Private pListTypeName As String
Private pListType As String
'pListType = "UN", "Units"
'            "ER", "Errors"
'            "ST", "SampleTypes"
'            "MB", "Specimen Sources"



Private Sub FillG()

          Dim s As String
          Dim tb As Recordset
          Dim sql As String

13740     On Error GoTo FillG_Error

13750     g.Rows = 2
13760     g.AddItem ""
13770     g.RemoveItem 1

13780     sql = "SELECT * FROM Lists WHERE " & _
              "ListType = '" & pListType & "' and InUse = 1 " & _
              "ORDER BY ListOrder"
13790     Set tb = New Recordset
13800     RecOpenServer 0, tb, sql
13810     Do While Not tb.EOF
13820         s = tb!Code & vbTab & tb!Text & ""
13830         g.AddItem s
13840         tb.MoveNext
13850     Loop

13860     If g.Rows > 2 Then
13870         g.RemoveItem 1
13880     End If

13890     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

13900     intEL = Erl
13910     strES = Err.Description
13920     LogError "frmListsGeneric", "FillG", intEL, strES, sql

End Sub

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

13930     If g.row = g.Rows - 1 Then Exit Sub
13940     n = g.row

13950     FireCounter = FireCounter + 1
13960     If FireCounter > 5 Then
13970         tmrDown.Interval = 100
13980     End If

13990     VisibleRows = g.height \ g.RowHeight(1) - 1

14000     g.Visible = False

14010     s = ""
14020     For X = 0 To g.Cols - 1
14030         s = s & g.TextMatrix(n, X) & vbTab
14040     Next
14050     s = Left$(s, Len(s) - 1)

14060     g.RemoveItem n
14070     If n < g.Rows Then
14080         g.AddItem s, n + 1
14090         g.row = n + 1
14100     Else
14110         g.AddItem s
14120         g.row = g.Rows - 1
14130     End If

14140     For X = 0 To g.Cols - 1
14150         g.Col = X
14160         g.CellBackColor = vbYellow
14170     Next

14180     If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
14190         If g.row - VisibleRows + 1 > 0 Then
14200             g.TopRow = g.row - VisibleRows + 1
14210         End If
14220     End If

14230     g.Visible = True

14240     cmdSave.Visible = True

End Sub

Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

14250     If g.row = 1 Then Exit Sub

14260     FireCounter = FireCounter + 1
14270     If FireCounter > 5 Then
14280         tmrUp.Interval = 100
14290     End If

14300     n = g.row

14310     g.Visible = False

14320     s = ""
14330     For X = 0 To g.Cols - 1
14340         s = s & g.TextMatrix(n, X) & vbTab
14350     Next
14360     s = Left$(s, Len(s) - 1)

14370     g.RemoveItem n
14380     g.AddItem s, n - 1

14390     g.row = n - 1
14400     For X = 0 To g.Cols - 1
14410         g.Col = X
14420         g.CellBackColor = vbYellow
14430     Next

14440     If Not g.RowIsVisible(g.row) Then
14450         g.TopRow = g.row
14460     End If

14470     g.Visible = True

14480     cmdSave.Visible = True

End Sub



Private Sub cmdAdd_Click()

14490     txtCode = Trim$(UCase$(txtCode))
14500     txtText = Trim$(txtText)

14510     If txtCode = "" Then
14520         Exit Sub
14530     End If

14540     If txtText = "" Then
14550         Exit Sub
14560     End If

14570     g.AddItem txtCode & vbTab & txtText

14580     If g.TextMatrix(1, 0) = "" Then
14590         g.RemoveItem 1
14600     End If

14610     txtCode = ""
14620     txtText = ""

14630     txtCode.SetFocus

14640     cmdSave.Visible = True

End Sub


Private Sub cmdCancel_Click()

14650     Unload Me

End Sub


Private Sub cmdDelete_Click()

          Dim Y As Integer
          Dim sql As String
          Dim s As String

14660     On Error GoTo cmdDelete_Click_Error

14670     g.Col = 0
14680     For Y = 1 To g.Rows - 1
14690         g.row = Y
14700         If g.CellBackColor = vbYellow Then
14710             s = "Delete " & g.TextMatrix(Y, 1) & vbCrLf & _
                      "From " & pListTypeNames & " ?"
14720             If iMsg(s, vbQuestion + vbYesNo) = vbYes Then
14730                 sql = "Delete from Lists where " & _
                          "ListType = '" & pListType & "' " & _
                          "and Code = '" & g.TextMatrix(Y, 0) & "'"
14740                 Cnxn(0).Execute sql
14750             End If
14760             Exit For
14770         End If
14780     Next

14790     cmdDelete.Enabled = False
14800     FillG

14810     Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

14820     intEL = Erl
14830     strES = Err.Description
14840     LogError "frmListsGeneric", "cmdDelete_Click", intEL, strES, sql


End Sub


Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14850     FireDown

14860     tmrDown.Interval = 250
14870     FireCounter = 0

14880     tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14890     tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14900     FireUp

14910     tmrUp.Interval = 250
14920     FireCounter = 0

14930     tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14940     tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

14950     Printer.Print

14960     Printer.Print "List of "; pListTypeNames

14970     g.Col = 0
14980     g.row = 1
14990     g.ColSel = g.Cols - 1
15000     g.RowSel = g.Rows - 1

15010     Printer.Print g.Clip

15020     Printer.EndDoc
15030     Screen.MousePointer = 0

End Sub


Private Sub cmdSave_Click()

          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String

15040     On Error GoTo cmdSave_Click_Error

15050     For Y = 1 To g.Rows - 1
15060         If g.TextMatrix(Y, 0) <> "" Then
15070             sql = "SELECT * FROM Lists WHERE " & _
                      "ListType = '" & pListType & "' " & _
                      "AND Code = '" & g.TextMatrix(Y, 0) & "' " & _
                      "AND InUse = 1"
15080             Set tb = New Recordset
15090             RecOpenServer 0, tb, sql
15100             If tb.EOF Then
15110                 tb.AddNew
15120             End If
15130             tb!Code = g.TextMatrix(Y, 0)
15140             tb!ListType = pListType
15150             tb!Text = g.TextMatrix(Y, 1)
15160             tb!ListOrder = Y
15170             tb!InUse = 1
15180             tb.Update
15190         End If
15200     Next

15210     FillG

15220     txtCode = ""
15230     txtText = ""
15240     txtCode.SetFocus
15250     cmdMoveUp.Enabled = False
15260     cmdMoveDown.Enabled = False
15270     cmdSave.Visible = False
15280     cmdDelete.Enabled = False

15290     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

15300     intEL = Erl
15310     strES = Err.Description
15320     LogError "frmListsGeneric", "cmdsave_Click", intEL, strES, sql

End Sub


Private Sub cmdXL_Click()

15330     ExportFlexGrid g, Me

End Sub


Private Sub Form_Activate()

15340     If Activated Then Exit Sub

15350     Activated = True

15360     FillG

End Sub

Private Sub Form_Load()

15370     g.Font.Bold = True

15380     If pListType = "" Then
15390         MsgBox "pListType not set"
15400     End If
15410     If pListTypeName = "" Then
15420         MsgBox "pListTypeName not set"
15430     End If
15440     If pListTypeNames = "" Then
15450         MsgBox "pListTypeNames not set"
15460     End If

15470     FrameAdd.Caption = "Add New " & pListTypeName
15480     Me.Caption = "NetAcquire - List of " & pListTypeNames

15490     Activated = False

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

15500     If cmdSave.Visible Then
15510         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
15520             Cancel = True
15530             Exit Sub
15540         End If
15550     End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

15560     pListType = ""
15570     pListTypeName = ""
15580     pListTypeNames = ""

End Sub

Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer

15590     On Error GoTo g_Click_Error

15600     ySave = g.row

15610     g.Visible = False
15620     g.Col = 0
15630     For Y = 1 To g.Rows - 1
15640         g.row = Y
15650         If g.CellBackColor = vbYellow Then
15660             For X = 0 To g.Cols - 1
15670                 g.Col = X
15680                 g.CellBackColor = 0
15690             Next
15700             Exit For
15710         End If
15720     Next
15730     g.row = ySave
15740     g.Visible = True

15750     If g.MouseRow = 0 Then
15760         If SortOrder Then
15770             g.Sort = flexSortGenericAscending
15780         Else
15790             g.Sort = flexSortGenericDescending
15800         End If
15810         SortOrder = Not SortOrder
15820         Exit Sub
15830     End If

15840     For X = 0 To g.Cols - 1
15850         g.Col = X
15860         g.CellBackColor = vbYellow
15870     Next

15880     cmdMoveUp.Enabled = True
15890     cmdMoveDown.Enabled = True
15900     cmdDelete.Enabled = True

15910     Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

15920     intEL = Erl
15930     strES = Err.Description
15940     LogError "frmListsGeneric", "g_Click", intEL, strES

End Sub



Private Sub tmrDown_Timer()

15950     FireDown

End Sub


Private Sub tmrUp_Timer()

15960     FireUp

End Sub



Public Property Let ListType(ByVal strNewValue As String)

15970     pListType = strNewValue

End Property
Public Property Let ListTypeName(ByVal strNewValue As String)

15980     pListTypeName = strNewValue

End Property

Public Property Let ListTypeNames(ByVal strNewValue As String)

15990     pListTypeNames = strNewValue

End Property

