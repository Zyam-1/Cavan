VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   4
      Top             =   2490
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ComboBox optOptions 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Text            =   "optOptions"
      Top             =   2040
      Width           =   3405
   End
   Begin VB.CommandButton cmdCancel 
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

Private retval As String

Private pListOrCombo
Private Sub cmdCancel_Click()

10    retval = ""
20    Me.Hide

End Sub

Public Property Get ReturnValue() As String

10    ReturnValue = retval

End Property

Private Sub bOK_Click()

10    retval = Trim$(optOptions)
20    Me.Hide

End Sub



Public Property Let Options(ByRef varOptions As Variant)

      Dim n As Integer

10    optOptions.Clear
20    For n = 0 To UBound(varOptions)
30      optOptions.AddItem varOptions(n)
40    Next

End Property

Public Property Let Prompt(ByVal strPrompt As String)

10    lPrompt = strPrompt

End Property

Public Property Let ListOrCombo(ByVal strListOrCombo As String)

10    pListOrCombo = strListOrCombo

End Property

Private Sub optOptions_KeyPress(KeyAscii As Integer)

10    If pListOrCombo = "List" Then
20      KeyAscii = 0
30    End If

End Sub


