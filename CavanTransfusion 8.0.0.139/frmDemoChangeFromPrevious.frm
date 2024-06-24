VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDemoChangeFromPrevious 
   Caption         =   "Demographics change"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Reject"
      Height          =   555
      Left            =   5160
      TabIndex        =   1
      Top             =   4155
      Width           =   1800
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   555
      Left            =   2595
      TabIndex        =   0
      Top             =   4155
      Width           =   1800
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   195
      Left            =   165
      TabIndex        =   23
      Top             =   4890
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line 
      X1              =   30
      X2              =   9615
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   22
      Top             =   165
      Width           =   1800
   End
   Begin VB.Label lblSex 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6300
      TabIndex        =   21
      Top             =   3420
      Width           =   2835
   End
   Begin VB.Label lblSex 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1710
      TabIndex        =   20
      Top             =   3345
      Width           =   2835
   End
   Begin VB.Label lblDOB 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   6300
      TabIndex        =   19
      Top             =   2925
      Width           =   2835
   End
   Begin VB.Label lblDOB 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1710
      TabIndex        =   18
      Top             =   2925
      Width           =   2835
   End
   Begin VB.Label lblPatientName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   6300
      TabIndex        =   17
      Top             =   2445
      Width           =   2835
   End
   Begin VB.Label lblPatientName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1710
      TabIndex        =   16
      Top             =   2385
      Width           =   2835
   End
   Begin VB.Label lblSampleId 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   6300
      TabIndex        =   15
      Top             =   1680
      Width           =   2835
   End
   Begin VB.Label lblSampleId 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1710
      TabIndex        =   14
      Top             =   1650
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Chart :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   405
      TabIndex        =   13
      Top             =   210
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Sex :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   4920
      TabIndex        =   12
      Top             =   3435
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Sex :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   255
      TabIndex        =   11
      Top             =   3375
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Birth :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   4920
      TabIndex        =   10
      Top             =   2970
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Birth :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   255
      TabIndex        =   9
      Top             =   2895
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "Patient Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   4920
      TabIndex        =   8
      Top             =   2430
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Patient Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   2415
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Sample Id :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4920
      TabIndex        =   6
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Sample Id :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   255
      TabIndex        =   5
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Current Demographics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Top             =   1185
      Width           =   2880
   End
   Begin VB.Label Label1 
      Caption         =   "Demographic details for patient have changed since previous sample"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   405
      TabIndex        =   3
      Top             =   585
      Width           =   6420
   End
   Begin VB.Label Label1 
      Caption         =   "Previous Demographics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   1185
      Width           =   2880
   End
End
Attribute VB_Name = "frmDemoChangeFromPrevious"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ReturnValue As Integer


Public Property Get retval() As Integer

10    retval = ReturnValue

End Property



Private Sub cmdAccept_Click()

      Dim PW As String

10    PW = iBOX("Your Password?", "NetAcquire", , True)
      'If colTechnicians.PasswordCorrectForUser(PW, UserName) Then
20        ReturnValue = 1
30        Unload Me
      'End If


End Sub

Private Sub cmdExit_Click()

10    ReturnValue = 0
20    Unload Me

End Sub

