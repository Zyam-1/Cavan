VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPractice 
   Caption         =   "NetAcquire - GP Practices"
   ClientHeight    =   5235
   ClientLeft      =   750
   ClientTop       =   600
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
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
      Picture         =   "frmPractice.frx":0000
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
      Picture         =   "frmPractice.frx":066A
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

37600 On Error GoTo FillG_Error

37610 g.Rows = 2
37620 g.AddItem ""
37630 g.RemoveItem 1

37640 sql = "Select * from Practices where " & _
            "Hospital = '" & cmbHospital & "'"
37650 Set tb = New Recordset
37660 RecOpenServer 0, tb, sql
37670 Do While Not tb.EOF
37680   g.AddItem tb!Text & vbTab & tb!FAX & ""
37690   tb.MoveNext
37700 Loop

37710 If g.Rows > 2 Then
37720   g.RemoveItem 1
37730 End If

37740 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

37750 intEL = Erl
37760 strES = Err.Description
37770 LogError "frmPractice", "FillG", intEL, strES, sql


End Sub

Private Sub cmbHospital_Click()

37780 FillG

End Sub



Private Sub cmdAdd_Click()

37790 If Trim$(txtPractice) = "" Then
37800   iMsg "Require Practice Name", vbCritical
37810   Exit Sub
37820 End If

37830 If Trim$(txtFAX) = "" Then
37840   iMsg "Require FAX Number", vbCritical
37850   Exit Sub
37860 End If

37870 g.AddItem txtPractice & vbTab & txtFAX

37880 txtPractice = ""
37890 txtFAX = ""

37900 cmbHospital.Enabled = False
37910 cmdSave.Visible = True

End Sub

Private Sub cmdCancel_Click()

37920 Unload Me

End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

37930 On Error GoTo cmdSave_Click_Error

37940 For n = 1 To g.Rows - 1
37950   sql = "Select * from Practices where " & _
              "Hospital = '" & cmbHospital & "' " & _
              "and [Text] = '" & g.TextMatrix(n, 0) & "'"
37960   Set tb = New Recordset
37970   RecOpenServer 0, tb, sql
37980   If tb.EOF Then
37990     tb.AddNew
38000   End If
38010   tb!Text = g.TextMatrix(n, 0)
38020   tb!FAX = g.TextMatrix(n, 1)
38030   tb!Hospital = cmbHospital
38040   tb.Update
38050 Next

38060 cmdSave.Visible = False
38070 cmbHospital.Enabled = True

38080 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

38090 intEL = Erl
38100 strES = Err.Description
38110 LogError "frmPractice", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Activate()

38120 If Activated Then
38130   Exit Sub
38140 End If

38150 Activated = True

38160 FillG

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

38170 On Error GoTo Form_Load_Error

38180 cmbHospital.Clear

38190 sql = "Select * from Lists where " & _
            "ListType = 'HO' and InUse = 1 " & _
            "order by ListOrder"
38200 Set tb = New Recordset
38210 RecOpenServer 0, tb, sql
38220 Do While Not tb.EOF
38230   cmbHospital.AddItem tb!Text & ""
38240   tb.MoveNext
38250 Loop
38260 For n = 0 To cmbHospital.ListCount - 1
38270   If cmbHospital.List(n) = HospName(0) Then
38280     cmbHospital = HospName(0)
38290   End If
38300 Next

38310 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

38320 intEL = Erl
38330 strES = Err.Description
38340 LogError "frmPractice", "Form_Load", intEL, strES, sql


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

38350 If cmdSave.Visible Then
38360   If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
38370     Cancel = True
38380   End If
38390 End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

38400 Activated = False

End Sub


