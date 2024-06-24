VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUserRoles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - System Rights"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSystemRights 
      Height          =   315
      Left            =   6615
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   135
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.ComboBox cmbMemberOf 
      Height          =   315
      Left            =   2100
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   3435
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1000
      Left            =   1380
      Picture         =   "frmUserRoles.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7860
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   1000
      Left            =   6990
      Picture         =   "frmUserRoles.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7860
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1000
      Left            =   8190
      Picture         =   "frmUserRoles.frx":6ADC
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7860
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   1000
      Left            =   180
      Picture         =   "frmUserRoles.frx":79A6
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7860
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7125
      Left            =   180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   570
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   12568
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   325
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmUserRoles.frx":8870
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
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   5985
      Picture         =   "frmUserRoles.frx":88FA
      Top             =   270
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   5985
      Picture         =   "frmUserRoles.frx":8BD0
      Top             =   45
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select System Role"
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
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   1665
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
      Left            =   2880
      TabIndex        =   5
      Top             =   8218
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmUserRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdAdd_Click()

30590 txtCode = Trim$(UCase$(txtCode))
30600 txtText = Trim$(txtText)

30610 If txtCode = "" Then
30620     Exit Sub
30630 End If

30640 If txtText = "" Then
30650     Exit Sub
30660 End If

30670 g.AddItem txtCode & vbTab & txtText

30680 If g.TextMatrix(1, 0) = "" Then
30690     g.RemoveItem 1
30700 End If

30710 txtCode = ""
30720 txtText = ""

30730 txtCode.SetFocus

30740 cmdSave.Visible = True

End Sub


Private Sub cmbMemberOf_Change()

30750 On Error GoTo cmbMemberOf_Change_Error

30760 FillG

30770 Exit Sub

cmbMemberOf_Change_Error:

      Dim strES As String
      Dim intEL As Integer

30780 intEL = Erl
30790 strES = Err.Description
30800 LogError "frmUserRoles", "cmbMemberOf_Change", intEL, strES

End Sub

Private Sub cmbMemberOf_Click()

30810 On Error GoTo cmbMemberOf_Click_Error

30820 FillG

30830 Exit Sub

cmbMemberOf_Click_Error:

      Dim strES As String
      Dim intEL As Integer

30840 intEL = Erl
30850 strES = Err.Description
30860 LogError "frmUserRoles", "cmbMemberOf_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

30870 Unload Me

End Sub








Private Sub cmdPrint_Click()

30880 Printer.Print

30890 Printer.Print "NetAcquire Security Settings for "; cmbMemberOf

30900 g.Col = 0
30910 g.row = 1
30920 g.ColSel = g.Cols - 1
30930 g.RowSel = g.Rows - 1

30940 Printer.Print g.Clip

30950 Printer.EndDoc
30960 Screen.MousePointer = 0

End Sub


Private Sub cmdSave_Click()

      Dim Y As Integer
      Dim tb As Recordset
      Dim sql As String
      Dim CurrentUserRole As New UserRole

30970 On Error GoTo cmdSave_Click_Error

30980 For Y = 1 To g.Rows - 1
30990     g.row = Y
31000     g.Col = 2
31010     If g.CellPicture = imgGreenTick.Picture Then
31020         CurrentUserRole.Update cmbMemberOf, g.TextMatrix(Y, 0), 1, UserName
31030     Else
31040         CurrentUserRole.Update cmbMemberOf, g.TextMatrix(Y, 0), 0, UserName
31050     End If
31060 Next

31070 cmdSave.Visible = False
31080 FillG



31090 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

31100 intEL = Erl
31110 strES = Err.Description
31120 LogError "frmListsGeneric", "cmdsave_Click", intEL, strES, sql

End Sub


Private Sub cmdXL_Click()

31130 ExportFlexGrid g, Me

End Sub

Private Sub Form_Load()

      Dim i As Integer
      Dim J As Integer

31140 On Error GoTo Form_Load_Error

31150 FillGenericList cmbMemberOf, "SR"
31160 FillGenericList cmbSystemRights, "SystemRights"

31170 For i = 0 To cmbMemberOf.ListCount - 1
31180     For J = 0 To cmbSystemRights.ListCount - 1
31190         Set URole = New UserRole
31200         With URole
31210             If .GetUserRole(cmbMemberOf.List(i), cmbSystemRights.List(J), UserName) = False Then
31220                 .MemberOf = cmbMemberOf.List(i)
31230                 .SystemRole = cmbSystemRights.List(J)
31240                 .Description = "Grants access permission to " & cmbSystemRights.List(J)
31250                 .Enabled = 1

31260                 .Add ("Administrator")
31270             End If
31280         End With
31290     Next J
31300 Next i
31310 g.ColWidth(0) = 0
31320 g.ColWidth(1) = g.ColWidth(1) + 1000
31330 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

31340 intEL = Erl
31350 strES = Err.Description
31360 LogError "frmUserRoles", "Form_Load", intEL, strES

End Sub

Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer

31370 On Error GoTo g_Click_Error

31380 ySave = g.row

31390 g.Visible = False
31400 g.Col = 0
31410 For Y = 1 To g.Rows - 1
31420     g.row = Y
31430     If g.CellBackColor = vbYellow Then
31440         For X = 0 To g.Cols - 1
31450             g.Col = X
31460             g.CellBackColor = 0
31470         Next
31480         Exit For
31490     End If
31500 Next
31510 g.row = ySave
31520 g.Visible = True

31530 If g.MouseRow = 0 Then
31540     If SortOrder Then
31550         g.Sort = flexSortGenericAscending
31560     Else
31570         g.Sort = flexSortGenericDescending
31580     End If
31590     SortOrder = Not SortOrder
31600     Exit Sub
31610 End If

31620 If g.MouseCol = 2 Then
31630     g.row = g.MouseRow
31640     g.Col = 2
31650     If g.CellPicture = imgGreenTick.Picture Then
31660         Set g.CellPicture = imgRedCross.Picture
31670     Else
31680         Set g.CellPicture = imgGreenTick.Picture
31690     End If
31700     cmdSave.Visible = True
31710 End If

31720 For X = 0 To g.Cols - 1
31730     g.Col = X
31740     g.CellBackColor = vbYellow
31750 Next



31760 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

31770 intEL = Erl
31780 strES = Err.Description
31790 LogError "frmListsGeneric", "g_Click", intEL, strES

End Sub




Private Sub FillG()

      Dim s As String
      Dim tb As Recordset
      Dim sql As String
      Dim URoleList As UserRoleCollection
      Dim URole As New UserRole
      Dim i As Integer

31800 On Error GoTo FillG_Error


31810 ClearFGrid g

31820 Set URoleList = URole.GetUserRoleList(cmbMemberOf, UserName)
31830 If Not URoleList Is Nothing Then
31840     If URoleList.Count > 0 Then

31850         For i = 1 To URoleList.Count
31860             Set URole = New UserRole
31870             Set URole = URoleList.Item(i)
31880             g.AddItem URole.SystemRole & vbTab & _
                            URole.Description
31890             g.row = g.Rows - 1
31900             g.Col = 2
31910             g.CellPictureAlignment = 4
31920             If URole.Enabled Then
31930                 Set g.CellPicture = imgGreenTick.Picture
31940             Else
31950                 Set g.CellPicture = imgRedCross.Picture
31960             End If
31970         Next i
31980     End If
31990 End If
32000 FixG g



32010 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

32020 intEL = Erl
32030 strES = Err.Description
32040 LogError "frmListsGeneric", "FillG", intEL, strES, sql

End Sub

