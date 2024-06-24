VERSION 5.00
Begin VB.Form frmComment 
   Caption         =   "NetAcquire - Comments"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   825
      Left            =   3105
      Picture         =   "frmComment.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2565
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   825
      Left            =   405
      Picture         =   "frmComment.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2565
      Width           =   1275
   End
   Begin VB.TextBox txtComment 
      Height          =   2130
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   315
      Width           =   4425
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pComment As String
Private Sub cmdCancel_Click()

22820     Me.Hide

End Sub

Private Sub cmdSave_Click()

22830     pComment = txtComment

22840     Me.Hide

End Sub

Public Property Get Comment() As String

22850     Comment = pComment

End Property

Public Property Let Comment(ByVal sNewValue As String)

22860     pComment = sNewValue
22870     txtComment = pComment

End Property

