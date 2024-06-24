VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAssignPanelsForHealthLink 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1125
      Left            =   5430
      Picture         =   "frmAssignPanelsForHealthLink.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4230
      Width           =   1065
   End
   Begin VB.CommandButton cndCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   1125
      Left            =   5430
      Picture         =   "frmAssignPanelsForHealthLink.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7590
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   210
      TabIndex        =   1
      Top             =   120
      Width           =   4785
      Begin VB.ComboBox cmbPanel 
         Height          =   315
         Left            =   3270
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   60
         Picture         =   "frmAssignPanelsForHealthLink.frx":1D94
         Top             =   90
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Click on Short Name to assign Panel "
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   2625
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7905
      Left            =   210
      TabIndex        =   0
      Top             =   780
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   13944
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   -2147483624
      ForeColorSel    =   255
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FormatString    =   "<Short Name          |<Long Name                    |<Panel             |<Code "
   End
End
Attribute VB_Name = "frmAssignPanelsForHealthLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String

56390     On Error GoTo FillG_Error

56400     g.Rows = 2
56410     g.AddItem ""
56420     g.RemoveItem 1

56430     sql = "SELECT DISTINCT ShortName, LongName, COALESCE(HealthLinkPanel, '') HealthLinkPanel, Code " & _
              "FROM BioTestDefinitions " & _
              "WHERE COALESCE(InUse, 1) = 1 " & _
              "ORDER BY ShortName"
56440     Set tb = New Recordset
56450     RecOpenServer 0, tb, sql
56460     Do While Not tb.EOF
56470         g.AddItem tb!ShortName & vbTab & _
                  tb!LongName & vbTab & _
                  tb!HealthLinkPanel & vbTab & _
                  tb!Code & ""
56480         tb.MoveNext
56490     Loop

56500     If g.Rows > 2 Then
56510         g.RemoveItem 1
56520     End If

56530     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

56540     intEL = Erl
56550     strES = Err.Description
56560     LogError "frmAssignPanelsForHealthLink", "FillG", intEL, strES, sql

End Sub

Private Sub FillPanels()

56570     With cmbPanel
56580         .Clear
56590         .AddItem ""
56600         .AddItem "LFT"
56610         .AddItem "U+E"
56620         .AddItem "Bone"
56630         .AddItem "TFT"
56640         .AddItem "Lipids"
56650         .AddItem "Cholesterol"
56660         .ListIndex = 0
56670     End With

End Sub

Private Sub cmdSave_Click()

          Dim n As Integer
          Dim sql As String

56680     On Error GoTo cmdSave_Click_Error

56690     For n = 1 To g.Rows - 1
56700         sql = "UPDATE BioTestDefinitions " & _
                  "SET HealthLinkPanel = '" & g.TextMatrix(n, 2) & "' " & _
                  "WHERE Code = '" & g.TextMatrix(n, 3) & "'"
56710         Cnxn(0).Execute sql
56720     Next

56730     cmdSave.Enabled = False

56740     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

56750     intEL = Erl
56760     strES = Err.Description
56770     LogError "frmAssignPanelsForHealthLink", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cndCancel_Click()

56780     Unload Me

End Sub

Private Sub Form_Load()

56790     g.ColWidth(3) = 0

56800     FillPanels
56810     FillG

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

56820     If cmdSave.Enabled Then
56830         If iMsg("Cancel without Saving", vbYesNo) = vbNo Then
56840             Cancel = True
56850         End If
56860     End If

End Sub


Private Sub g_Click()

          Dim s As String
          Dim R As Integer
          Dim c As Integer

56870     R = g.MouseRow
56880     c = g.MouseCol

56890     If c = 0 And R > 0 Then
56900         g.TextMatrix(R, 2) = cmbPanel
56910         cmdSave.Enabled = True
56920     ElseIf c = 2 And R > 0 Then
56930         s = iBOX("Panel Name?", , g.TextMatrix(R, c))
56940         g.TextMatrix(R, c) = s
56950         cmdSave.Enabled = True
56960     End If

End Sub

