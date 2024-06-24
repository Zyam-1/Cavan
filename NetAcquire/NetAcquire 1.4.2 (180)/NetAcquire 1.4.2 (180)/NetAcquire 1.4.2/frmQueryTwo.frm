VERSION 5.00
Begin VB.Form frmQueryTwo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   1845
   ClientLeft      =   7260
   ClientTop       =   5355
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelect 
      Caption         =   "cmdSelect(1)"
      Height          =   405
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   720
      Width           =   2985
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "cmdSelect(0)"
      Height          =   405
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   2985
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Niether"
      Height          =   405
      Left            =   300
      MaskColor       =   &H80000000&
      TabIndex        =   0
      Top             =   1320
      Width           =   2985
   End
End
Attribute VB_Name = "frmQueryTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pReturnVal As String

Private pShowNeither As Boolean
Private Sub cmdCancel_Click()

42380 pReturnVal = ""
42390 Me.Hide

End Sub

Private Sub cmdSelect_Click(Index As Integer)

42400 pReturnVal = cmdSelect(Index).Caption

42410 Me.Hide

End Sub

Public Property Get ReturnVal() As String

42420 ReturnVal = pReturnVal

End Property

Private Sub Form_Activate()

42430 If pShowNeither Then
42440   Me.height = 2250
42450 Else
42460   Me.height = 1650
42470 End If

End Sub

Public Property Let ShowNeither(ByVal blnValue As Boolean)

42480 pShowNeither = blnValue

End Property
