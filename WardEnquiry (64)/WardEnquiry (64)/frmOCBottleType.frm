VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmOCBottleType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5010
   ClientLeft      =   3675
   ClientTop       =   3225
   ClientWidth     =   9135
   Icon            =   "frmOCBottleType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2925
      Left            =   450
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1890
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   5159
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Code |<Bottle Name                              |<Colour    |<Anticoagulant     |<Size in ml   "
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   8010
      Picture         =   "frmOCBottleType.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3750
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add new Bottle Type"
      Height          =   1485
      Left            =   450
      TabIndex        =   2
      Top             =   240
      Width           =   7365
      Begin VB.TextBox txtBottleName 
         Height          =   285
         Left            =   1140
         TabIndex        =   13
         Top             =   330
         Width           =   4875
      End
      Begin VB.ComboBox cmbAnticoagulant 
         Height          =   315
         Left            =   4260
         TabIndex        =   11
         Text            =   "cmbAnticoagulant"
         Top             =   660
         Width           =   1755
      End
      Begin VB.ComboBox cmbVolume 
         Height          =   315
         Left            =   4260
         TabIndex        =   10
         Text            =   "cmbVolume"
         Top             =   1020
         Width           =   1755
      End
      Begin VB.ComboBox cmbColour 
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Text            =   "cmbColour"
         Top             =   1020
         Width           =   1755
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   765
         Left            =   6270
         Picture         =   "frmOCBottleType.frx":1D94
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   885
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1140
         MaxLength       =   1
         TabIndex        =   0
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bottle Name"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Anticoagulant"
         Height          =   195
         Left            =   3210
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   690
         TabIndex        =   5
         Top             =   690
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Size in ml"
         Height          =   195
         Left            =   3510
         TabIndex        =   4
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Colour"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   1050
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmOCBottleType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "SELECT * FROM OCBottleType"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    Do While Not tb.EOF

90      g.AddItem tb!Code & vbTab & _
                  tb!BottleName & vbTab & _
                  tb!Colour & vbTab & _
                  tb!Anticoagulant & vbTab & _
                  tb!Volume & ""
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
180   LogError "frmOCBottleType", "FillG", intEL, strES, sql

End Sub

Private Sub cmbColour_Click()

10    cmbColour.BackColor = vbWindowBackground

End Sub

Private Sub cmbColour_KeyDown(KeyCode As Integer, Shift As Integer)

10    cmbColour.BackColor = vbWindowBackground

End Sub


Private Sub cmdAdd_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdAdd_Click_Error

20    txtBottleName = Trim$(txtBottleName)
30    If txtBottleName = "" Then
40      txtBottleName.BackColor = vbRed
50      txtBottleName.SetFocus
60      Exit Sub
70    End If

80    txtCode = UCase$(Trim$(txtCode))
90    If txtCode = "" Then
100     txtCode.BackColor = vbRed
110     txtCode.SetFocus
120     Exit Sub
130   End If

140   If cmbColour = "" Then
150     cmbColour.BackColor = vbRed
160     cmbColour.SetFocus
170     Exit Sub
180   End If

190   If cmbAnticoagulant = "" Then
200     cmbAnticoagulant.SetFocus
210     Exit Sub
220   End If

230   If cmbVolume = "" Then
240     cmbVolume.SetFocus
250     Exit Sub
260   End If

270   sql = "SELECT * FROM OCBottleType WHERE " & _
            "Code = '" & txtCode & "'"
280   Set tb = New Recordset
290   RecOpenServer 0, tb, sql
300   If tb.EOF Then
310     tb.AddNew
320   End If
330   tb!Code = txtCode
340   tb!BottleName = txtBottleName
350   tb!Colour = cmbColour
360   tb!Anticoagulant = cmbAnticoagulant
370   tb!Volume = cmbVolume
380   tb!UserName = UserName
390   tb.Update

400   FillG
  
410   txtCode = ""
420   cmbColour = ""
430   cmbAnticoagulant = ""
440   cmbVolume = ""
450   txtBottleName = ""
460   txtBottleName.SetFocus

470   Exit Sub

cmdAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "frmOCBottleType", "cmdAdd_Click", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub Form_Load()

10    FillG

End Sub


Private Sub g_Click()

      Static SortOrder As Boolean

10    If g.MouseRow = 0 Then
20      If SortOrder Then
30        g.Sort = flexSortGenericAscending
40      Else
50        g.Sort = flexSortGenericDescending
60      End If
70      SortOrder = Not SortOrder
80      Exit Sub
90    End If

      'g.Col = 0
      'txtCode = g
      'colBottles.Delete g
      'g.Col = 1
      'tColour = g
      'g.Col = 2
      'tAnticoagulant = g
      'g.Col = 3
      'tSize = g
      '
      'If g.Rows = 2 Then
      '  g.AddItem ""
      '  g.RemoveItem 1
      'Else
      '  g.RemoveItem g.Row
      'End If

End Sub


Private Sub txtBottleName_KeyDown(KeyCode As Integer, Shift As Integer)

10    txtBottleName.BackColor = vbWindowBackground

End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)

10    txtCode.BackColor = vbWindowBackground

End Sub


