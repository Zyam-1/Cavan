VERSION 5.00
Begin VB.Form fcdrInputBox 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   1995
   ClientTop       =   2265
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInput 
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Top             =   1950
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O. K."
      Default         =   -1  'True
      Height          =   645
      Left            =   4200
      Picture         =   "fcdrInputBox.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4200
      Picture         =   "fcdrInputBox.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   660
      TabIndex        =   2
      Top             =   180
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fcdrInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ReturnValue As String
Private mPass As Boolean


Private Sub cmdCancel_Click()

10    ReturnValue = ""
20    Unload Me

End Sub

Private Sub cmdOK_Click()

10    ReturnValue = txtInput
20    Unload Me

End Sub

Public Property Get Retval() As String

10    Retval = ReturnValue

End Property





Private Sub Form_Activate()

10    If mPass Then
20      txtInput.PasswordChar = "*"
30    Else
40      txtInput.PasswordChar = ""
50    End If

60    txtInput.SelStart = 0
70    txtInput.SelLength = Len(txtInput)
80    txtInput.SetFocus

End Sub

Public Property Let PassWord(ByVal blnNewValue As Boolean)

10    mPass = blnNewValue

End Property

Private Sub Form_Deactivate()

10    mPass = False

End Sub


