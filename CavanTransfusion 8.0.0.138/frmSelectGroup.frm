VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectGroup 
   Caption         =   "NetAcquire - Select Group"
   ClientHeight    =   3345
   ClientLeft      =   2565
   ClientTop       =   1170
   ClientWidth     =   3645
   Icon            =   "frmSelectGroup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   3645
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   1740
      Picture         =   "frmSelectGroup.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2220
      Width           =   1605
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   675
      Left            =   360
      Picture         =   "frmSelectGroup.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2220
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rhesus"
      Height          =   1665
      Left            =   1740
      TabIndex        =   1
      Top             =   180
      Width           =   1605
      Begin VB.OptionButton optRh 
         Caption         =   "Negative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1425
      End
      Begin VB.OptionButton optRh 
         Caption         =   "Positive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   390
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group"
      Height          =   1665
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   1275
      Begin VB.OptionButton optGroup 
         Caption         =   "AB"
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
         Index           =   3
         Left            =   330
         TabIndex        =   5
         Top             =   1170
         Width           =   735
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "B"
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
         Index           =   2
         Left            =   330
         TabIndex        =   4
         Top             =   880
         Width           =   555
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "A"
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
         Index           =   1
         Left            =   330
         TabIndex        =   3
         Top             =   590
         Width           =   585
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "O"
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
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   300
         Width           =   615
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   360
      TabIndex        =   10
      Top             =   3060
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmSelectGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pGRh As String

Private Sub cmdCancel_Click()

10    pGRh = ""
20    Me.Hide

End Sub

Private Sub cmdOK_Click()

      Dim n As Integer
      Dim Group As String
      Dim Found As Boolean
      Dim s As String

10    s = "Both Group and Rhesus must be selected"

20    For n = 0 To 3
30      If optGroup(n) Then
40        Group = optGroup(n).Caption
50      End If
60    Next
70    If Group = "" Then
80      iMsg s, vbCritical
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110   End If

120   Found = False
130   For n = 0 To 1
140     If optRh(n) Then
150       Found = True
160       Group = Group & Left$(optRh(n).Caption, 3)
170     End If
180   Next
190   If Not Found Then
200     iMsg s, vbCritical
210     If TimedOut Then Unload Me: Exit Sub
220     Exit Sub
230   End If

240   pGRh = Group
250   Me.Hide

End Sub



Public Property Get GRh() As String

10    GRh = pGRh

End Property

Private Sub Form_Activate()

      Dim n As Integer

10    For n = 0 To 3
20      optGroup(n) = False
30    Next
40    For n = 0 To 1
50      optRh(n) = False
60    Next

End Sub

