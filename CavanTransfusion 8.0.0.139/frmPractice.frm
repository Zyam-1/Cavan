VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPractice 
   Caption         =   "NetAcquire - GP Practices"
   ClientHeight    =   5550
   ClientLeft      =   750
   ClientTop       =   600
   ClientWidth     =   7680
   Icon            =   "frmPractice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7680
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   300
      TabIndex        =   5
      Top             =   210
      Width           =   4905
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   345
         Left            =   3840
         TabIndex        =   10
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox txtFAX 
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   960
         Width           =   3225
      End
      Begin VB.TextBox txtPractice 
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   390
         Width           =   3225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "FAX Number"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Practice Name"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   180
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   5790
      Picture         =   "frmPractice.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1395
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   705
      Left            =   5790
      Picture         =   "frmPractice.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3195
      Left            =   330
      TabIndex        =   2
      Top             =   1800
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Practice Name                     |<FAX Number             "
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
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Text            =   "cmbHospital"
      Top             =   630
      Width           =   1995
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   330
      TabIndex        =   11
      Top             =   5160
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   5430
      TabIndex        =   1
      Top             =   420
      Width           =   570
   End
End
Attribute VB_Name = "frmPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from Practices where " & _
            "Hospital = '" & cmbHospital & "'"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    Do While Not tb.EOF
90      g.AddItem tb!Text & vbTab & tb!FAX & ""
100     tb.MoveNext
110   Loop

120   If g.Rows > 2 Then
130     g.RemoveItem 1
140   End If

150   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmPractice", "FillG", intEL, strES, sql


End Sub

Private Sub cmbHospital_Click()

10    FillG

End Sub



Private Sub cmdadd_Click()

10    If Trim$(txtPractice) = "" Then
20      iMsg "Require Practice Name", vbCritical
30      If TimedOut Then Unload Me: Exit Sub
40      Exit Sub
50    End If

60    If Trim$(txtFAX) = "" Then
70      iMsg "Require FAX Number", vbCritical
80      If TimedOut Then Unload Me: Exit Sub
90      Exit Sub
100   End If

110   g.AddItem txtPractice & vbTab & txtFAX

120   txtPractice = ""
130   txtFAX = ""

140   cmbHospital.Enabled = False
150   cmdSave.Visible = True

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

10    On Error GoTo cmdSave_Click_Error

20    For n = 1 To g.Rows - 1
30      sql = "Select * from Practices where " & _
              "Hospital = '" & cmbHospital & "' " & _
              "and [Text] = '" & g.TextMatrix(n, 0) & "'"
40      Set tb = New Recordset
50      RecOpenServer 0, tb, sql
60      If tb.EOF Then
70        tb.AddNew
80      End If
90      tb!Text = g.TextMatrix(n, 0)
100     tb!FAX = g.TextMatrix(n, 1)
110     tb!Hospital = cmbHospital
120     tb.Update
130   Next

140   cmdSave.Visible = False
150   cmbHospital.Enabled = True

160   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmPractice", "cmdSave_Click", intEL, strES, sql


End Sub




Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo Form_Load_Error

20    cmbHospital.Clear

30    sql = "Select * from Lists where " & _
            "ListType = 'HO' " & _
            "order by ListOrder"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70      cmbHospital.AddItem tb!Text & ""
80      tb.MoveNext
90    Loop
100   For n = 0 To cmbHospital.ListCount - 1
110     If cmbHospital.List(n) = HospName(0) Then
120       cmbHospital = HospName(0)
130     End If
140   Next

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
150       FillG
      '**************************************

160   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmPractice", "Form_Load", intEL, strES, sql


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Visible Then
30      Answer = iMsg("Cancel without saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70      End If
80    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub


