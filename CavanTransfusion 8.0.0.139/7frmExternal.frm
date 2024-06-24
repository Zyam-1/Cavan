VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExternal 
   Caption         =   "NetAcquire - External Reports"
   ClientHeight    =   5835
   ClientLeft      =   1650
   ClientTop       =   1110
   ClientWidth     =   6585
   Icon            =   "7frmExternal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   675
      Left            =   3720
      Picture         =   "7frmExternal.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   5130
      Picture         =   "7frmExternal.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1245
   End
   Begin VB.TextBox tReport 
      Height          =   4545
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   870
      Width           =   6435
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   480
      TabIndex        =   5
      Top             =   5580
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblChart 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123456789"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "frmExternal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChart As String
Private mAandE As String

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdSave_Click()

      Dim sql As String
      Dim sqlA As String
      Dim tb As Recordset

10    On Error GoTo cmdSave_Click_Error

20    If mChart <> "" Then
30      sqlA = "MRN = '" & mChart & "'"
40    ElseIf mAandE <> "" Then
50      sqlA = "AandE = '" & mAandE & "'"
60    Else
70      Exit Sub
80    End If

90    sql = "Select * from ExternalNotes where " & sqlA
100   Set tb = New Recordset
110   RecOpenServerBB 0, tb, sql
120   If Trim$(treport) <> "" Then
130     If tb.EOF Then
140       tb.AddNew
150     End If
160     tb!MRN = mChart
170     tb!AandE = mAandE
180     tb!Notes = treport & vbCrLf & _
                  "Opertator : " & UserName & vbCrLf & _
                  "Date Time : " & Format(Now, "dd/mmm/yyyy hh:mm:ss")
190     tb.Update
200   Else
210     If Not tb.EOF Then
220       tb.Delete
230     End If
240   End If

250   Unload Me

260   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmExternal", "cmdSave_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()

      Dim sql As String
      Dim sqlA As String
      Dim tb As Recordset

10    On Error GoTo Form_Activate_Error

20    If mChart <> "" Then
30      lblTitle = "Chart"
40      lblChart = mChart
50      sqlA = "MRN = '" & mChart & "'"
60    ElseIf mAandE <> "" Then
70      lblTitle = "A/E"
80      lblChart = mAandE
90      sqlA = "AandE = '" & mAandE & "'"
100   Else
110     Exit Sub
120   End If

130   sql = "SELECT * FROM ExternalNotes WHERE " & sqlA
140   Set tb = New Recordset
150   RecOpenServerBB 0, tb, sql
160   If Not tb.EOF Then
170     treport = tb!Notes & ""
180   End If

190   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmExternal", "Form_Activate", intEL, strES, sql


End Sub

Public Property Let Chart(ByVal NewValue As String)

10    mChart = NewValue

End Property

Public Property Let AandE(ByVal NewValue As String)

10    mAandE = NewValue

End Property

