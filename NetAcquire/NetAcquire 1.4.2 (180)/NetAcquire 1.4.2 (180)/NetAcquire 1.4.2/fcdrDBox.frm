VERSION 5.00
Begin VB.Form fcdrDBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   2700
   ClientLeft      =   3600
   ClientTop       =   2535
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox optOptions 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Text            =   "optOptions"
      Top             =   2040
      Width           =   3405
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   4320
      Picture         =   "fcdrDBox.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1245
   End
   Begin VB.CommandButton bOK 
      Caption         =   "O. K."
      Default         =   -1  'True
      Height          =   645
      Left            =   4320
      Picture         =   "fcdrDBox.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1245
   End
   Begin VB.Label lPrompt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   480
      TabIndex        =   2
      Top             =   210
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fcdrDBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RetVal As String

Private pListOrCombo

Private pDefault As String

Private Sub bcancel_Click()

31530     RetVal = ""
31540     Me.Hide

End Sub

Public Property Get ReturnValue() As String

31550     ReturnValue = RetVal

End Property

Private Sub bOK_Click()

31560     RetVal = Trim$(optOptions)
31570     Me.Hide

End Sub


Public Property Let Options(ByRef varOptions As Variant)

          Dim n As Integer

31580     optOptions.Clear
31590     For n = 0 To UBound(varOptions)
31600         optOptions.AddItem varOptions(n)
31610     Next

End Property

Public Property Let Prompt(ByVal strPrompt As String)

31620     lPrompt = strPrompt

End Property

Public Property Let ListOrCombo(ByVal strListOrCombo As String)

31630     pListOrCombo = strListOrCombo

End Property

Private Sub Form_Activate()

31640     optOptions = pDefault

End Sub

Private Sub optOptions_KeyPress(KeyAscii As Integer)

31650     If pListOrCombo = "List" Then
31660         KeyAscii = 0
31670     End If

End Sub



Public Property Let Default(ByVal sNewValue As String)

31680     pDefault = sNewValue

End Property
