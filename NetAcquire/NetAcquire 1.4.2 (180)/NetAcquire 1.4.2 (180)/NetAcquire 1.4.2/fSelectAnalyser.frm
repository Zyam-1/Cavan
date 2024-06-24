VERSION 5.00
Begin VB.Form fSelectAnalyser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire"
   ClientHeight    =   3555
   ClientLeft      =   3915
   ClientTop       =   2280
   ClientWidth     =   3285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3285
   Begin VB.Label lblArchitect 
      AutoSize        =   -1  'True
      Caption         =   "Architect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   2685
      Width           =   1095
   End
   Begin VB.Label lblIntegraB 
      AutoSize        =   -1  'True
      Caption         =   "Integra Beta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   645
      TabIndex        =   5
      Top             =   1800
      Width           =   1530
   End
   Begin VB.Label lblIntegraA 
      AutoSize        =   -1  'True
      Caption         =   "Integra Alpha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   4
      Top             =   960
      Width           =   1650
   End
   Begin VB.Label lblR 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   825
      Left            =   2250
      TabIndex        =   3
      Top             =   2430
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Analyser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   330
      TabIndex        =   2
      Top             =   150
      Width           =   2520
   End
   Begin VB.Label lblB 
      AutoSize        =   -1  'True
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   885
      Left            =   2280
      TabIndex        =   1
      Top             =   1515
      Width           =   570
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   630
      Width           =   570
   End
End
Attribute VB_Name = "fSelectAnalyser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAnalyser As String

Public Property Get Analyser() As String

10    Analyser = mAnalyser

End Property

Private Sub lblA_Click()

10    mAnalyser = "A"
20    Unload Me

End Sub


Private Sub lblArchitect_Click()

10    mAnalyser = "R"
20    Unload Me

End Sub

Private Sub lblB_Click()

10    mAnalyser = "B"
20    Unload Me

End Sub


Private Sub lblIntegraA_Click()

10    mAnalyser = "A"
20    Unload Me

End Sub

Private Sub lblIntegraB_Click()

10    mAnalyser = "B"
20    Unload Me

End Sub


Private Sub lblR_Click()

10    mAnalyser = "R"
20    Unload Me

End Sub



