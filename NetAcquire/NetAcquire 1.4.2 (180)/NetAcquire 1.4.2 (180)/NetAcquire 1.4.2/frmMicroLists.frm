VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMicroLists 
   Caption         =   "NetAcquire"
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   360
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   8565
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7020
      Top             =   5370
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7020
      Top             =   6150
   End
   Begin VB.CommandButton cmdOrganisms 
      Caption         =   "Organisms"
      Height          =   315
      Left            =   7200
      TabIndex        =   14
      Top             =   1650
      Width           =   1215
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   7290
      Picture         =   "frmMicroLists.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2370
      Width           =   795
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   7290
      Picture         =   "frmMicroLists.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3240
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Organism Group"
      Height          =   1365
      Left            =   180
      TabIndex        =   21
      Top             =   150
      Width           =   4365
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
         TabIndex        =   0
         Top             =   330
         Width           =   975
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
         TabIndex        =   1
         Top             =   900
         Width           =   3495
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   300
         TabIndex        =   22
         Top             =   960
         Width           =   315
      End
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7260
      Picture         =   "frmMicroLists.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7290
      Picture         =   "frmMicroLists.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   795
   End
   Begin VB.CommandButton bSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   7290
      Picture         =   "frmMicroLists.frx":1558
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7500
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Type"
      Height          =   1455
      Left            =   4710
      TabIndex        =   20
      Top             =   60
      Width           =   3705
      Begin VB.OptionButton o 
         Alignment       =   1  'Right Justify
         Caption         =   "Gram Quantity"
         Height          =   225
         Index           =   10
         Left            =   390
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton o 
         Caption         =   "Identification Notes"
         Height          =   225
         Index           =   9
         Left            =   1800
         TabIndex        =   12
         Top             =   1200
         Width           =   1725
      End
      Begin VB.OptionButton o 
         Alignment       =   1  'Right Justify
         Caption         =   "Qualifiers"
         Height          =   225
         Index           =   8
         Left            =   750
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton o 
         Caption         =   "Miscellaneous"
         Height          =   225
         Index           =   7
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   1395
      End
      Begin VB.OptionButton o 
         Caption         =   "Crystals"
         Height          =   225
         Index           =   6
         Left            =   1800
         TabIndex        =   10
         Top             =   720
         Width           =   915
      End
      Begin VB.OptionButton o 
         Caption         =   "Casts"
         Height          =   225
         Index           =   5
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   1035
      End
      Begin VB.OptionButton o 
         Caption         =   "Wet Preps"
         Height          =   225
         Index           =   4
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton o 
         Alignment       =   1  'Right Justify
         Caption         =   "Gram Stains"
         Height          =   225
         Index           =   3
         Left            =   510
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton o 
         Alignment       =   1  'Right Justify
         Caption         =   "Organism Groups"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   1545
      End
      Begin VB.OptionButton o 
         Alignment       =   1  'Right Justify
         Caption         =   "Ovae"
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6675
      Left            =   210
      TabIndex        =   13
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
Attribute VB_Name = "frmMicroLists"
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

39380 If g.row = g.Rows - 1 Then Exit Sub
39390 n = g.row

39400 VisibleRows = g.height \ g.RowHeight(1) - 1

39410 FireCounter = FireCounter + 1
39420 If FireCounter > 5 Then
39430   tmrDown.Interval = 100
39440 End If

39450 g.Visible = False

39460 s = ""
39470 For X = 0 To g.Cols - 1
39480   s = s & g.TextMatrix(n, X) & vbTab
39490 Next
39500 s = Left$(s, Len(s) - 1)

39510 g.RemoveItem n
39520 If n < g.Rows Then
39530   g.AddItem s, n + 1
39540   g.row = n + 1
39550 Else
39560   g.AddItem s
39570   g.row = g.Rows - 1
39580 End If

39590 For X = 0 To g.Cols - 1
39600   g.Col = X
39610   g.CellBackColor = vbYellow
39620 Next

39630 If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
39640   If g.row - VisibleRows + 1 > 0 Then
39650     g.TopRow = g.row - VisibleRows + 1
39660   End If
39670 End If

39680 g.Visible = True

39690 bsave.Visible = True

End Sub

Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

39700 If g.row = 1 Then Exit Sub

39710 FireCounter = FireCounter + 1
39720 If FireCounter > 5 Then
39730   tmrUp.Interval = 100
39740 End If

39750 n = g.row

39760 g.Visible = False

39770 s = ""
39780 For X = 0 To g.Cols - 1
39790   s = s & g.TextMatrix(n, X) & vbTab
39800 Next
39810 s = Left$(s, Len(s) - 1)

39820 g.RemoveItem n
39830 g.AddItem s, n - 1

39840 g.row = n - 1
39850 For X = 0 To g.Cols - 1
39860   g.Col = X
39870   g.CellBackColor = vbYellow
39880 Next

39890 If Not g.RowIsVisible(g.row) Then
39900   g.TopRow = g.row
39910 End If

39920 g.Visible = True

39930 bsave.Visible = True

End Sub



Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim LT As String
      Dim s As String

39940 On Error GoTo FillG_Error

39950 LT = Switch(o(1), "OR", _
                  o(2), "OV", _
                  o(3), "GS", _
                  o(4), "WP", _
                  o(5), "CA", _
                  o(6), "CR", _
                  o(7), "MI", _
                  o(8), "MQ", _
                  o(9), "IN", _
                  o(10), "GQ")

39960 g.Rows = 2
39970 g.AddItem ""
39980 g.RemoveItem 1

39990 sql = "Select * from Lists where " & _
            "ListType = '" & LT & "' and InUse = 1 order by ListOrder"
40000 Set tb = New Recordset
40010 RecOpenServer 0, tb, sql
40020 Do While Not tb.EOF
40030   s = tb!Code & vbTab & tb!Text & ""
40040   g.AddItem s
40050   tb.MoveNext
40060 Loop

40070 If g.Rows > 2 Then
40080   g.RemoveItem 1
40090 End If

40100 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

40110 intEL = Erl
40120 strES = Err.Description
40130 LogError "frmMicroLists", "FillG", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

40140 tCode = Trim$(UCase$(tCode))
40150 tText = Trim$(tText)

40160 If tCode = "" Then
40170   Exit Sub
40180 End If

40190 If tText = "" Then Exit Sub

40200 g.AddItem tCode & vbTab & tText

40210 tCode = ""
40220 tText = ""
40230 tCode.SetFocus
40240 bsave.Visible = True

End Sub


Private Sub bcancel_Click()

40250 Unload Me

End Sub


Private Sub bMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

40260 FireDown

40270 tmrDown.Interval = 250
40280 FireCounter = 0

40290 tmrDown.Enabled = True

End Sub


Private Sub bMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

40300 tmrDown.Enabled = False

End Sub


Private Sub bMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

40310 FireUp

40320 tmrUp.Interval = 250
40330 FireCounter = 0

40340 tmrUp.Enabled = True

End Sub


Private Sub bMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

40350 tmrUp.Enabled = False

End Sub


Private Sub bPrint_Click()

      Dim LT As String

40360 LT = Switch(o(1), "Organisms.", _
                  o(2), "Ova.", _
                  o(3), "Gram Stains.", _
                  o(4), "Wet Preps.", _
                  o(5), "Casts.", _
                  o(6), "Crystals.", _
                  o(7), "Miscellaneous.", _
                  o(8), "Qualifiers", _
                  o(9), "Identification Notes")

40370 Printer.Print

40380 Printer.Print "List of "; LT

40390 g.Col = 0
40400 g.row = 1
40410 g.ColSel = g.Cols - 1
40420 g.RowSel = g.Rows - 1

40430 Printer.Print g.Clip

40440 Printer.EndDoc
40450 Screen.MousePointer = 0

End Sub


Private Sub bSave_Click()

      Dim LT As String
      Dim Y As Integer
      Dim tb As Recordset
      Dim sql As String

40460 On Error GoTo bSave_Click_Error

40470 LT = Switch(o(1), "OR", _
                  o(2), "OV", _
                  o(3), "GS", _
                  o(4), "WP", _
                  o(5), "CA", _
                  o(6), "CR", _
                  o(7), "MI", _
                  o(8), "MQ", _
                  o(9), "IN", _
                  o(10), "GQ")

40480 For Y = 1 To g.Rows - 1
40490   sql = "Select * from Lists where " & _
              "ListType = '" & LT & "' " & _
              "and Code = '" & g.TextMatrix(Y, 0) & "' and InUse = 1"
40500   Set tb = New Recordset
40510   RecOpenServer 0, tb, sql
40520   If tb.EOF Then
40530     tb.AddNew
40540   End If
40550   tb!Code = g.TextMatrix(Y, 0)
40560   tb!ListType = LT
40570   tb!Text = g.TextMatrix(Y, 1)
40580   tb!ListOrder = Y
40590   tb!InUse = 1
40600   tb.Update
40610 Next

40620 FillG

40630 tCode = ""
40640 tText = ""
40650 tCode.SetFocus
40660 bMoveUp.Enabled = False
40670 bMoveDown.Enabled = False
40680 bsave.Visible = False

40690 Exit Sub

bSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40700 intEL = Erl
40710 strES = Err.Description
40720 LogError "frmMicroLists", "bSave_Click", intEL, strES, sql


End Sub


Private Sub cmdOrganisms_Click()

40730 frmOrganisms.Show 1

End Sub

Private Sub Form_Activate()

40740 If Activated Then Exit Sub

40750 Activated = True

40760 FillG

End Sub

Private Sub Form_Load()

40770 g.Font.Bold = True

40780 Activated = False

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

40790 If bsave.Visible Then
40800   If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
40810     Cancel = True
40820     Exit Sub
40830   End If
40840 End If

End Sub


Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer

40850 ySave = g.row

40860 g.Visible = False
40870 g.Col = 0
40880 For Y = 1 To g.Rows - 1
40890   g.row = Y
40900   If g.CellBackColor = vbYellow Then
40910     For X = 0 To g.Cols - 1
40920       g.Col = X
40930       g.CellBackColor = 0
40940     Next
40950     Exit For
40960   End If
40970 Next
40980 g.row = ySave
40990 g.Visible = True

41000 If g.MouseRow = 0 Then
41010   If SortOrder Then
41020     g.Sort = flexSortGenericAscending
41030   Else
41040     g.Sort = flexSortGenericDescending
41050   End If
41060   SortOrder = Not SortOrder
41070   Exit Sub
41080 End If

41090 For X = 0 To g.Cols - 1
41100   g.Col = X
41110   g.CellBackColor = vbYellow
41120 Next

41130 bMoveUp.Enabled = True
41140 bMoveDown.Enabled = True

End Sub


Private Sub o_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

41150 FillG

41160 FrameAdd.Caption = "Add New " & Left$(o(Index).Caption, Len(o(Index).Caption) - 1)

41170 tCode = ""
41180 tText = ""
41190 If tCode.Visible Then
41200   tCode.SetFocus
41210 End If

End Sub


Private Sub tmrDown_Timer()

41220 FireDown

End Sub


Private Sub tmrUp_Timer()

41230 FireUp

End Sub


