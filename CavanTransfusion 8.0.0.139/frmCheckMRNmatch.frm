VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmCheckMRNmatch 
   Caption         =   "NetAcquire - MRN Check"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTypeMRN 
      Height          =   315
      Left            =   1305
      TabIndex        =   0
      Top             =   1365
      Width           =   2385
   End
   Begin VB.TextBox txtScanMRN 
      Height          =   315
      Left            =   1305
      TabIndex        =   1
      Top             =   2205
      Width           =   2385
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Default         =   -1  'True
      Height          =   1155
      Left            =   1305
      Picture         =   "frmCheckMRNmatch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2775
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1155
      Left            =   3780
      Picture         =   "frmCheckMRNmatch.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2745
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   90
      TabIndex        =   6
      Top             =   4305
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblPackNumber 
      Caption         =   "MRN Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1815
      TabIndex        =   7
      Top             =   345
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1410
      X2              =   3810
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type in MRN"
      Height          =   195
      Left            =   1335
      TabIndex        =   5
      Top             =   1170
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Scan MRN Barcode"
      Height          =   195
      Left            =   1335
      TabIndex        =   4
      Top             =   1965
      Width           =   1440
   End
End
Attribute VB_Name = "frmCheckMRNmatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnReturnValue As Boolean

Public Property Get retval() As Integer

10    retval = blnReturnValue

End Property

Private Sub cmdCancel_Click()
10    Unload Me
End Sub

Private Sub cmdCheck_Click()

10    If txtTypeMRN <> "" And txtScanMRN <> "" Then
20        If UCase(Trim$(txtTypeMRN)) = UCase(Trim$(txtScanMRN)) Then
30            blnReturnValue = True
40            Unload Me
50        Else
60            blnReturnValue = False
70            Unload Me
80        End If
90    End If

End Sub


Private Sub Form_Activate()
10    txtTypeMRN.SetFocus
End Sub

