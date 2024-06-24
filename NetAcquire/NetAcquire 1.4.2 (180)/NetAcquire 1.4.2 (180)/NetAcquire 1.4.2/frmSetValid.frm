VERSION 5.00
Begin VB.Form frmSetValid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSetValid 
      Height          =   315
      Left            =   570
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   4800
      Picture         =   "frmSetValid.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveValid 
      Caption         =   "&Save Details"
      Height          =   705
      Left            =   3180
      Picture         =   "frmSetValid.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Validated Date/Time"
      Height          =   195
      Left            =   570
      TabIndex        =   4
      Top             =   960
      Width           =   1485
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   570
      TabIndex        =   3
      Top             =   360
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   570
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSetValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalendar_Click()

End Sub


Private Sub bcancel_Click()

End Sub

Private Sub cmdCancel_Click()

59460 Unload Me

End Sub


Private Sub cmdSaveValid_Click()

      Dim sql As String

59470 On Error GoTo cmdSaveValid_Click_Error

59480 If Not IsDate(cmbSetValid) Then
59490   iMsg "Not a valid Date!", vbExclamation
59500   Exit Sub
59510 End If
        
59520 sql = "INSERT INTO PrintValidLogArc " & _
            "  SELECT PrintValidLog.*, " & _
            "  '" & AddTicks(UserName) & "', " & _
            "  '" & Format$(Now, "dd/MMM/yyyy hh:mm:ss") & "' " & _
            "  FROM PrintValidLog WHERE " & _
            "  SampleID = '" & Val(lblSampleID) & "' " & _
            "  AND Department = 'M' "
59530 Cnxn(0).Execute sql

59540 sql = "UPDATE PrintValidLog " & _
            "SET ValidatedDateTime = '" & Format$(cmbSetValid, "dd/MMM/yyyy HH:mm:ss") & "' " & _
            "WHERE SampleID = '" & Val(lblSampleID) & "'"
59550 Cnxn(0).Execute sql

59560 Unload Me

59570 Exit Sub

cmdSaveValid_Click_Error:

      Dim strES As String
      Dim intEL As Integer

59580 intEL = Erl
59590 strES = Err.Description
59600 LogError "frmSetValid", "cmdSaveValid_Click", intEL, strES, sql

End Sub

Private Sub Form_Activate()

      Dim tb As Recordset
      Dim sql As String

59610 On Error GoTo Form_Activate_Error

59620 cmbSetValid.Clear

59630 sql = "SELECT * FROM PrintValidLog WHERE " & _
            "SampleID = '" & Val(lblSampleID) & "' " & _
            "AND ValidatedDateTime IS NOT NULL"
59640 Set tb = New Recordset
59650 RecOpenServer 0, tb, sql
59660 If Not tb.EOF Then
59670   cmbSetValid.AddItem tb!ValidatedDateTime
59680 End If

59690 sql = "SELECT * FROM PrintValidLogArc WHERE " & _
            "SampleID = '" & Val(lblSampleID) & "' " & _
            "AND ValidatedDateTime IS NOT NULL " & _
            "ORDER BY ValidatedDateTime desc"
59700 Set tb = New Recordset
59710 RecOpenServer 0, tb, sql
59720 Do While Not tb.EOF
59730   cmbSetValid.AddItem tb!ValidatedDateTime
59740   tb.MoveNext
59750 Loop

59760 Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

59770 intEL = Erl
59780 strES = Err.Description
59790 LogError "frmSetValid", "Form_Activate", intEL, strES, sql


End Sub

