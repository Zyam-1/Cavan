VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSelectBioImmEnd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   705
      Left            =   4650
      Picture         =   "frmSelectBioImmEnd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4770
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   4650
      Picture         =   "frmSelectBioImmEnd.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5970
      Width           =   1395
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   705
      Left            =   4650
      Picture         =   "frmSelectBioImmEnd.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3090
      Width           =   1395
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6375
      Left            =   300
      TabIndex        =   0
      Top             =   330
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   3
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "Code|<Test Name                      |<Discipline                       "
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
      Left            =   4695
      TabIndex        =   3
      Top             =   3990
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmSelectBioImmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

53270 On Error GoTo FillG_Error

53280 g.Rows = 2

53290 sql = "SELECT DISTINCT (Code), ShortName, BIE FROM BioTestDefinitions " & _
            "ORDER BY ShortName"
53300 Set tb = New Recordset
53310 RecOpenServer 0, tb, sql
53320 Do While Not tb.EOF
53330   s = tb!Code & vbTab & _
            tb!ShortName & vbTab
53340   If tb!BIE = "B" Then
53350     s = s & "Biochemistry"
53360   ElseIf tb!BIE = "I" Then
53370     s = s & "Immunology"
53380   ElseIf tb!BIE = "E" Then
53390     s = s & "Endocrinology"
53400   Else
53410     s = s & "?"
53420   End If
53430   g.AddItem s
53440   tb.MoveNext
53450 Loop

53460 If g.Rows > 2 Then
53470   g.RemoveItem 1
53480 End If

53490 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

53500 intEL = Erl
53510 strES = Err.Description
53520 LogError "frmSelectBioImmEnd", "FillG", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

53530 If cmdSave.Visible Then
53540   If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
53550     Exit Sub
53560   End If
53570 End If

53580 Unload Me

End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim n As Integer

53590 On Error GoTo cmdSave_Click_Error

53600 For n = 1 To g.Rows - 1
53610   If g.TextMatrix(n, 2) <> "?" Then
53620     sql = "UPDATE BioTestDefinitions SET BIE = '" & Left$(g.TextMatrix(n, 2), 1) & "' WHERE Code = '" & g.TextMatrix(n, 0) & "'"
53630     Cnxn(0).Execute sql
53640   End If
53650 Next

53660 cmdSave.Visible = False

53670 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

53680 intEL = Erl
53690 strES = Err.Description
53700 LogError "frmSelectBioImmEnd", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()

53710 ExportFlexGrid g, Me

End Sub


Private Sub Form_Load()

53720 g.ColWidth(0) = 0

53730 FillG

End Sub


Private Sub g_Click()

53740 If g.MouseRow = 0 Then Exit Sub
53750 If g.MouseCol <> 2 Then Exit Sub

53760 Select Case g.TextMatrix(g.MouseRow, 2)
        Case "Biochemistry"
53770     g.TextMatrix(g.MouseRow, 2) = "Immunology"
53780   Case "Immunology"
53790     g.TextMatrix(g.MouseRow, 2) = "Endocrinology"
53800   Case Else
53810     g.TextMatrix(g.MouseRow, 2) = "Biochemistry"
53820 End Select

53830 cmdSave.Visible = True

End Sub


