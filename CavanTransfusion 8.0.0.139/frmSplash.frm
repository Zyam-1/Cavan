VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.PictureBox picUpdate 
         BackColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   3270
         ScaleHeight     =   765
         ScaleWidth      =   3495
         TabIndex        =   6
         Top             =   270
         Width           =   3555
         Begin ComctlLib.ProgressBar pbUpdate 
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   510
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Checking Database Integrity. Please wait ..."
            Height          =   225
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   3585
         End
         Begin VB.Label lblUpdate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Updating : "
            Height          =   225
            Left            =   0
            TabIndex        =   8
            Top             =   240
            Width           =   3585
         End
      End
      Begin VB.Image Image 
         Height          =   2220
         Left            =   165
         Picture         =   "frmSplash.frx":08CA
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2310
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright: Custom Software Ltd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4515
         TabIndex        =   1
         Top             =   3705
         Width           =   2460
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   2
         Top             =   3045
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows NT/Citrix"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4110
         TabIndex        =   3
         Top             =   2685
         Width           =   2745
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   765
         Left            =   2940
         TabIndex        =   4
         Top             =   1650
         Width           =   2430
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   5
      Top             =   4170
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
10        Unload Me
End Sub

Private Sub Form_Load()

10    ReDim Cnxn(0 To 0) As Connection
20    ReDim CnxnBB(0 To 0) As Connection
30    ReDim HospName(0 To 0) As String

40    On Error GoTo Form_Load_Error

50    CheckIDE

60    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
70    lblProductName.Caption = App.Title
80    Me.Show     ' Display startup form.
90    DoEvents    ' Ensure startup form is painted.

100   GetINI

105   Load_RBCRT011
'100   ConnectToDatabase
110   Load frmMain  ' Load main application fom.
120   Unload Me   ' Unload startup form.

130   frmMain.Show 1 ' Display main form.

140   Exit Sub

Form_Load_Error:

Dim strES As String
Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   MsgBox "Error in frmSplash, Form_Load Line " & intEL

End Sub

Private Sub Frame1_Click()
10        Unload Me
End Sub

