VERSION 5.00
Begin VB.Form frmOverRideDSN 
   Caption         =   "NetAcquire - DSN OverRide"
   ClientHeight    =   3150
   ClientLeft      =   7110
   ClientTop       =   2730
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5220
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   525
      Left            =   2010
      TabIndex        =   4
      Top             =   2250
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select DSN"
      Height          =   1455
      Left            =   300
      TabIndex        =   0
      Top             =   450
      Width           =   4545
      Begin VB.OptionButton optDSN 
         Caption         =   "Not Set"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   990
         Width           =   3465
      End
      Begin VB.OptionButton optDSN 
         Caption         =   "Not Set"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   660
         Width           =   3465
      End
      Begin VB.OptionButton optDSN 
         Caption         =   "Not Set"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmOverRideDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ptb As Recordset

Private pDSN As String

Private Sub cmdContinue_Click()

If optDSN(0) Then
  pDSN = "DSN"
ElseIf optDSN(1) Then
  pDSN = "Live69DSN"
ElseIf optDSN(2) Then
  pDSN = "Test69DSN"
End If

Unload Me

End Sub

Private Sub Form_Activate()

Dim n As Integer

n = 0

If Trim$(ptb!DSN & "") <> "" Then
  optDSN(0).Caption = Trim$(ptb!DSN)
  optDSN(0).Enabled = True
  n = 1
End If
If Trim$(ptb!Live69DSN & "") <> "" Then
  optDSN(1).Caption = Trim$(ptb!Live69DSN)
  optDSN(1).Enabled = True
  n = n + 1
End If
If Trim$(ptb!Test69DSN & "") <> "" Then
  optDSN(2).Caption = Trim$(ptb!Test69DSN)
  optDSN(2).Enabled = True
  n = n + 1
End If

If n = 1 Then
  pDSN = "DSN"
  Unload Me
End If

End Sub

Public Property Let Rec(ByVal R As Recordset)

Set ptb = R

End Property

Public Property Get DSN() As String

DSN = pDSN

End Property

