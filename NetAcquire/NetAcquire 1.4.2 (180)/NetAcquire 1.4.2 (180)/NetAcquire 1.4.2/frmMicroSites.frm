VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMicroSites 
   Caption         =   "NetAcquire - Microbiology - Sites"
   ClientHeight    =   8460
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   7815
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7410
      Top             =   5730
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7410
      Top             =   4890
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3630
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5340
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4500
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":1330
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":199A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1710
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Site"
      Height          =   1365
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6705
      Begin VB.OptionButton optABs 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   4980
         TabIndex        =   17
         Top             =   810
         Width           =   405
      End
      Begin VB.OptionButton optABs 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   4350
         TabIndex        =   16
         Top             =   810
         Value           =   -1  'True
         Width           =   405
      End
      Begin VB.OptionButton optABs 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   15
         Top             =   810
         Width           =   345
      End
      Begin VB.OptionButton optABs 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   14
         Top             =   810
         Width           =   345
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   285
         Left            =   5910
         TabIndex        =   3
         Top             =   810
         Width           =   585
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
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   2
         Top             =   450
         Width           =   3495
      End
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
         Left            =   690
         MaxLength       =   5
         TabIndex        =   1
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Number of Antibiotics to Report"
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   1980
         TabIndex        =   5
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   210
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6795
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11986
      _Version        =   393216
      Cols            =   3
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
      FormatString    =   "<Code   |Text                                                                             |^AB's "
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
Attribute VB_Name = "frmMicroSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Integer

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

53110 If g.row = g.Rows - 1 Then Exit Sub
53120 n = g.row

53130 FireCounter = FireCounter + 1
53140 If FireCounter > 5 Then
53150   tmrDown.Interval = 100
53160 End If

53170 VisibleRows = g.height \ g.RowHeight(1) - 1

53180 g.Visible = False

53190 s = ""
53200 For X = 0 To g.Cols - 1
53210   s = s & g.TextMatrix(n, X) & vbTab
53220 Next
53230 s = Left$(s, Len(s) - 1)

53240 g.RemoveItem n
53250 If n < g.Rows Then
53260   g.AddItem s, n + 1
53270   g.row = n + 1
53280 Else
53290   g.AddItem s
53300   g.row = g.Rows - 1
53310 End If

53320 For X = 0 To g.Cols - 1
53330   g.Col = X
53340   g.CellBackColor = vbYellow
53350 Next

53360 If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
53370   If g.row - VisibleRows + 1 > 0 Then
53380     g.TopRow = g.row - VisibleRows + 1
53390   End If
53400 End If

53410 g.Visible = True

53420 cmdSave.Visible = True

End Sub

Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

53430 If g.row = 1 Then Exit Sub

53440 FireCounter = FireCounter + 1
53450 If FireCounter > 5 Then
53460   tmrUp.Interval = 100
53470 End If

53480 n = g.row

53490 g.Visible = False

53500 s = ""
53510 For X = 0 To g.Cols - 1
53520   s = s & g.TextMatrix(n, X) & vbTab
53530 Next
53540 s = Left$(s, Len(s) - 1)

53550 g.RemoveItem n
53560 g.AddItem s, n - 1

53570 g.row = n - 1
53580 For X = 0 To g.Cols - 1
53590   g.Col = X
53600   g.CellBackColor = vbYellow
53610 Next

53620 If Not g.RowIsVisible(g.row) Then
53630   g.TopRow = g.row
53640 End If

53650 g.Visible = True

53660 cmdSave.Visible = True

End Sub



Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

53670 On Error GoTo FillG_Error

53680 g.Rows = 2
53690 g.AddItem ""
53700 g.RemoveItem 1

53710 sql = "Select * from Lists where " & _
            "ListType = 'SI' and InUse = 1 " & _
            "order by ListOrder"
53720 Set tb = New Recordset
53730 RecOpenServer 0, tb, sql
53740 Do While Not tb.EOF
53750   s = tb!Code & vbTab & tb!Text & vbTab & tb!Default & ""
53760   g.AddItem s
53770   tb.MoveNext
53780 Loop

53790 If g.Rows > 2 Then
53800   g.RemoveItem 1
53810 End If

53820 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

53830 intEL = Erl
53840 strES = Err.Description
53850 LogError "frmMicroSites", "FillG", intEL, strES, sql


End Sub


Private Sub cmdAdd_Click()
53860 On Error GoTo ErrorHandler

      Dim n As Integer
      Dim DefaultABs As String

53870 txtCode = Trim$(UCase$(txtCode))
53880 txtText = Trim$(txtText)

53890 If txtCode = "" Then Exit Sub
53900 If txtText = "" Then Exit Sub

53910 For n = 0 To 3
53920   If optABs(n) Then
53930     DefaultABs = Format$(n + 1)
53940   End If
53950 Next

53960 g.AddItem txtCode & vbTab & txtText & vbTab & DefaultABs

53970 txtCode = ""
53980 txtText = ""
53990 optABs(sysOptDefaultABs(0)) = True

54000 cmdSave.Visible = True

54010     Exit Sub
ErrorHandler:
      Dim strES As String
      Dim intEL As Integer

54020 intEL = Erl
54030 strES = Err.Description
54040 LogError "frmMicroSites", "cmdAdd_Click", intEL, strES
End Sub


Private Sub cmdCancel_Click()

54050 Unload Me

End Sub


Private Sub cmdDelete_Click()

      Dim sql As String

54060 On Error GoTo cmdDelete_Click_Error

54070 sql = "Delete from Lists where " & _
            "ListType = 'SI' " & _
            "and Code = '" & g.TextMatrix(g.row, 0) & "'"
54080 Cnxn(0).Execute sql
        
54090 FillG

54100 Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

54110 intEL = Erl
54120 strES = Err.Description
54130 LogError "frmMicroSites", "cmdDelete_Click", intEL, strES, sql


End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

54140 FireDown

54150 tmrDown.Interval = 250
54160 FireCounter = 0

54170 tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

54180 tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

54190 FireUp

54200 tmrUp.Interval = 250
54210 FireCounter = 0

54220 tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

54230 tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

54240 Printer.Print

54250 Printer.Print "List of Sites."

54260 g.Col = 0
54270 g.row = 1
54280 g.ColSel = g.Cols - 1
54290 g.RowSel = g.Rows - 1

54300 Printer.Print g.Clip

54310 Printer.EndDoc
54320 Screen.MousePointer = 0

End Sub


Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim Y As Integer

54330 On Error GoTo cmdSave_Click_Error

54340 For Y = 1 To g.Rows - 1
54350   sql = "Select * from Lists where " & _
              "ListType = 'SI' " & _
              "and Code = '" & g.TextMatrix(Y, 0) & "' and InUse = 1"
54360   Set tb = New Recordset
54370   RecOpenServer 0, tb, sql
54380   If tb.EOF Then
54390     tb.AddNew
54400   End If
54410   tb!Code = g.TextMatrix(Y, 0)
54420   tb!ListType = "SI"
54430   tb!Text = g.TextMatrix(Y, 1)
54440   tb!Default = g.TextMatrix(Y, 2)
54450   tb!ListOrder = Y
54460   tb!InUse = 1
54470   tb.Update
54480 Next

54490 FillG

54500 txtCode = ""
54510 txtText = ""
54520 If txtCode.Visible And txtCode.Enabled Then txtCode.SetFocus
54530 cmdMoveUp.Enabled = False
54540 cmdMoveDown.Enabled = False
54550 cmdDelete.Enabled = False
54560 cmdSave.Visible = False

54570 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

54580 intEL = Erl
54590 strES = Err.Description
54600 LogError "frmMicroSites", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Activate()

54610 If Activated Then Exit Sub

54620 Activated = True

54630 FillG

End Sub

Private Sub Form_Load()

54640 g.Font.Bold = True

54650 Activated = False

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

54660 If cmdSave.Visible Then
54670   If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
54680     Cancel = True
54690     Exit Sub
54700   End If
54710 End If

End Sub


Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer

54720 If g.MouseRow = 0 Then
54730   If SortOrder Then
54740     g.Sort = flexSortGenericAscending
54750   Else
54760     g.Sort = flexSortGenericDescending
54770   End If
54780   SortOrder = Not SortOrder
54790   Exit Sub
54800 End If

54810 ySave = g.row

54820 If g.Col = 2 Then
54830   g.Enabled = False
54840   g = iBOX("Number of Antibiotics to Report?", , sysOptDefaultABs(0))
54850   g = Val(g)
54860   g.Enabled = True
54870   cmdSave.Visible = True
54880   Exit Sub
54890 End If

54900 g.Visible = False
54910 g.Col = 0
54920 For Y = 1 To g.Rows - 1
54930   g.row = Y
54940   If g.CellBackColor = vbYellow Then
54950     For X = 0 To g.Cols - 1
54960       g.Col = X
54970       g.CellBackColor = 0
54980     Next
54990     Exit For
55000   End If
55010 Next
55020 g.row = ySave
55030 g.Visible = True

55040 For X = 0 To g.Cols - 1
55050   g.Col = X
55060   g.CellBackColor = vbYellow
55070 Next

55080 cmdMoveUp.Enabled = True
55090 cmdMoveDown.Enabled = True
55100 cmdDelete.Enabled = True

End Sub


Private Sub optABs_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim sql As String

55110 On Error GoTo optABs_MouseUp_Error

55120 If sysOptDefaultABs(0) <> Index + 1 Then
55130   If iMsg("Do you want to reset the Default to " & Format(Index + 1) & " ?", vbQuestion + vbYesNo) = vbYes Then
55140     sql = "Update Options " & _
                "Set Contents = '" & Index + 1 & "' where Description = 'DefaultABs'"
55150     Cnxn(0).Execute sql
55160     sysOptDefaultABs(0) = Index + 1
55170   End If
55180 End If

55190 Exit Sub

optABs_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

55200 intEL = Erl
55210 strES = Err.Description
55220 LogError "frmMicroSites", "optABs_MouseUp", intEL, strES, sql


End Sub


Private Sub tmrDown_Timer()

55230 FireDown

End Sub


Private Sub tmrUp_Timer()

55240 FireUp

End Sub


