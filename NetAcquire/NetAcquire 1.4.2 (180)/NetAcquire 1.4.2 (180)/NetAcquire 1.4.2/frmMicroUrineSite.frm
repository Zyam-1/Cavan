VERSION 5.00
Begin VB.Form frmMicroUrineSite 
   Caption         =   "NetAcquire"
   ClientHeight    =   2550
   ClientLeft      =   7200
   ClientTop       =   4170
   ClientWidth     =   2145
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   2145
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   1170
      Picture         =   "frmMicroUrineSite.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1500
      Width           =   765
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   825
      Left            =   180
      Picture         =   "frmMicroUrineSite.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1500
      Width           =   765
   End
   Begin VB.Frame Frame2 
      Caption         =   "Urine Sample"
      Height          =   1095
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   1785
      Begin VB.OptionButton optU 
         Caption         =   "SPA"
         Height          =   195
         Index           =   3
         Left            =   930
         TabIndex        =   6
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "BSU"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Caption         =   "CSU"
         Height          =   195
         Index           =   1
         Left            =   930
         TabIndex        =   4
         Top             =   360
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "MSU"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmMicroUrineSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()

Dim n As Integer

For n = 0 To 3
  optU(n).Value = False
Next
Me.Hide

End Sub


Private Sub cmdsave_Click()

Me.Hide

End Sub


Public Property Get Details() As String

Dim n As Integer

For n = 0 To 3
  If optU(n) Then
    Details = optU(n).Caption
  End If
Next

End Property

