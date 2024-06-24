VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestMasks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Masks"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCriteria 
      Height          =   315
      Left            =   3780
      TabIndex        =   6
      Text            =   "cmbCriteria"
      Top             =   180
      Width           =   3045
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1125
      Left            =   7200
      Picture         =   "frmTestMasks.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1155
      Left            =   7230
      Picture         =   "frmTestMasks.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   1155
      HelpContextID   =   10026
      Left            =   7230
      Picture         =   "frmTestMasks.frx":1C8C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7380
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7815
      HelpContextID   =   10090
      Left            =   180
      TabIndex        =   0
      Top             =   510
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   13785
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
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
      AllowUserResizing=   1
      FormatString    =   "<Code          |<Long Name                 |<Short Name      |^Old    |^  Lipaemic  |^  Icteric  |^  Haemolysed  "
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   180
      TabIndex        =   4
      Top             =   8370
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   8640
      Picture         =   "frmTestMasks.frx":2B56
      Top             =   2010
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   8640
      Picture         =   "frmTestMasks.frx":2E2C
      Top             =   2310
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   7230
      TabIndex        =   3
      Top             =   3810
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "frmTestMasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private pDiscipline As String
Private pSampleType As String

Private Lxs As New LIHs

Private Sub SaveMasks()

      Dim sql As String
      Dim Y As Integer
      Dim Lxs As New LIHs
      Dim Lx As LIH
      Dim Criteria As String
      Dim X As Integer
      Dim LIorH As String

12980 On Error GoTo SaveMasks_Error

12990 If cmbCriteria = "Do not Print Result if >=" Then
13000   Criteria = "P"
13010 Else
13020   Criteria = "W"
13030 End If

13040 g.Col = 3
13050 For Y = 1 To g.Rows - 1
13060   g.row = Y
13070   sql = "UPDATE " & pDiscipline & "TestDefinitions " & _
              "SET O = '" & IIf(g.CellPicture = imgSquareTick, 1, 0) & "' " & _
              "WHERE LongName = '" & g.TextMatrix(Y, 1) & "'"
13080   Cnxn(0).Execute sql

13090   For X = 1 To 3
13100     LIorH = Choose(X, "L", "I", "H")
13110     Set Lx = Lxs.Item(LIorH, g.TextMatrix(Y, 0), Criteria)
13120     If Lx Is Nothing Then
13130       Set Lx = New LIH
13140       Lx.LIorH = LIorH
13150       Lx.NoPrintOrWarning = Criteria
13160     End If
13170     Lx.Code = g.TextMatrix(Y, 0)
13180     Lx.CutOff = Val(g.TextMatrix(Y, X + 3))
13190     Lx.UserName = UserName
13200     Lx.Save
13210   Next

13220 Next

13230 cmdSave.Visible = False

13240 Exit Sub

SaveMasks_Error:

      Dim strES As String
      Dim intEL As Integer

13250 intEL = Erl
13260 strES = Err.Description
13270 LogError "frmTestMasks", "SaveMasks", intEL, strES, sql
           
End Sub



Public Property Let Discipline(ByVal strNewValue As String)

13280 pDiscipline = UCase$(strNewValue)

End Property


Private Sub cmbCriteria_Click()

13290 FillgMasks

End Sub

Private Sub cmbCriteria_KeyPress(KeyAscii As Integer)

13300 KeyAscii = 0

End Sub


Private Sub cmdCancel_Click()
        
13310 Unload Me

End Sub


Private Sub cmdSave_Click()

13320 SaveMasks
13330 Lxs.Load

13340 FillgMasks
End Sub

Private Sub cmdXL_Click()

13350 ExportFlexGrid g, Me

End Sub

Private Sub Form_Load()

13360 g.ColWidth(0) = 0

13370 cmbCriteria.Clear
13380 cmbCriteria.AddItem "Do not Print Result if >="
13390 cmbCriteria.AddItem "Issue Warning if >="
13400 cmbCriteria.ListIndex = 0

13410 Lxs.Load

13420 FillgMasks

End Sub


Private Sub FillgMasks()

      Dim tb As Recordset
      Dim sql As String
      Dim strS As String
      Dim Lx As LIH
      Dim Criteria As String
      Dim X As Integer
      Dim LIorH As String

13430 On Error GoTo FillgMasks_Error

13440 If cmbCriteria = "Do not Print Result if >=" Then
13450   Criteria = "P"
13460 Else
13470   Criteria = "W"
13480 End If

13490 With g
13500   .Visible = False
13510   .Rows = 2
13520   .AddItem ""
13530   .RemoveItem 1
13540   .Col = 3 'Old

13550   sql = "SELECT DISTINCT Code, LongName, ShortName, " & _
              "COALESCE (O, 0) AS O, " & _
              "PrintPriority " & _
              "FROM " & pDiscipline & "TestDefinitions " & _
              "ORDER BY PrintPriority"
13560   Set tb = New Recordset
13570   RecOpenServer 0, tb, sql
13580   Do While Not tb.EOF
13590     strS = tb!Code & vbTab & _
                 tb!LongName & vbTab & _
                 tb!ShortName & vbTab & vbTab
13600     .AddItem strS

13610     .row = .Rows - 1
13620     Set .CellPicture = IIf(tb!o <> 0, imgSquareTick.Picture, imgSquareCross.Picture)
13630     .CellPictureAlignment = flexAlignCenterCenter

13640     For X = 1 To 3
13650       LIorH = Choose(X, "L", "I", "H")
13660       Set Lx = Lxs.Item(LIorH, g.TextMatrix(.row, 0), Criteria)
13670       If Not Lx Is Nothing Then
13680         If Lx.CutOff > 0 Then
13690           g.TextMatrix(.row, X + 3) = Lx.CutOff
13700         Else
13710           g.TextMatrix(.row, X + 3) = ""
13720         End If
13730       End If
13740     Next
          
13750     tb.MoveNext
        
13760   Loop
        
13770   If .Rows > 2 Then
13780     .RemoveItem 1
13790   End If
13800   .Visible = True

13810 End With

13820 Exit Sub

FillgMasks_Error:

      Dim strES As String
      Dim intEL As Integer

13830 intEL = Erl
13840 strES = Err.Description
13850 LogError "frmTestMasks", "FillgMasks", intEL, strES, sql


End Sub



Private Sub g_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim v As Single

13860 If g.MouseRow = 0 Then Exit Sub

13870 Select Case g.Col
          
        Case 3 'Old
13880     If g.CellPicture = imgSquareCross.Picture Then
13890       Set g.CellPicture = imgSquareTick.Picture
13900     Else
13910       Set g.CellPicture = imgSquareCross.Picture
13920     End If
13930     cmdSave.Visible = True
          
13940   Case 4, 5, 6 'LIH
          
13950     v = Val(iBOX("Enter cut-off point for " & g.TextMatrix(0, g.Col), , g.TextMatrix(g.row, g.Col)))
13960     g.TextMatrix(g.row, g.Col) = Format$(v)
13970     cmdSave.Visible = True
        
13980 End Select

End Sub


