VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowse 
   Caption         =   "NetAcquire"
   ClientHeight    =   4740
   ClientLeft      =   2865
   ClientTop       =   3525
   ClientWidth     =   5805
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   5805
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   2490
      Picture         =   "frmBrowse.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4020
      Width           =   1245
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   3090
      TabIndex        =   2
      Top             =   120
      Width           =   2505
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   180
      TabIndex        =   1
      Top             =   510
      Width           =   2835
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2865
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   90
      TabIndex        =   5
      Top             =   4530
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblPathAndFile 
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   3600
      Width           =   5445
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Me.Hide

End Sub

Private Sub Dir1_Change()
10    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
10    Dir1.Path = Drive1.Drive
End Sub


Private Sub File1_Click()

10    lblPathAndFile = File1.Path & "\" & File1.FileName

End Sub

Private Sub File1_DblClick()

10    lblPathAndFile = File1.Path & "\" & File1.FileName

20    Me.Hide

End Sub

Private Sub File1_PathChange()

10    Debug.Print File1.Path

End Sub


Private Sub Form_Load()

10    File1.Pattern = "*.wav"

End Sub


