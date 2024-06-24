VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectPrinter 
   Caption         =   "NetAcquire - Select Printer"
   ClientHeight    =   5250
   ClientLeft      =   4095
   ClientTop       =   2430
   ClientWidth     =   4755
   Icon            =   "frmSelectPrinter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   4755
   Begin VB.ListBox lstPrinter 
      Height          =   3765
      Left            =   330
      TabIndex        =   1
      Top             =   330
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   3210
      TabIndex        =   0
      Top             =   4260
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   2
      Top             =   5010
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmSelectPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

      Dim Px As Printer

10    lstPrinter.Clear

20    For Each Px In Printers
30      lstPrinter.AddItem Px.DeviceName
40    Next

End Sub


