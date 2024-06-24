VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmGetDoB 
   Caption         =   "Date of Birth"
   ClientHeight    =   4005
   ClientLeft      =   5205
   ClientTop       =   1320
   ClientWidth     =   3045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   3045
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   2160
      Picture         =   "frmGetDoB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel"
      Top             =   3090
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Continue"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   330
      Picture         =   "frmGetDoB.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3090
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2085
      Left            =   330
      TabIndex        =   2
      Top             =   690
      Width           =   2445
      Begin ComCtl2.UpDown udWithin 
         Height          =   285
         Left            =   1170
         TabIndex        =   7
         Top             =   1560
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblWithin"
         BuddyDispid     =   196615
         OrigLeft        =   2490
         OrigTop         =   1320
         OrigRight       =   2730
         OrigBottom      =   1965
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.OptionButton optFuzzy 
         Caption         =   "Use a 'Fuzzy' Search"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton optExact 
         Caption         =   "Use Exact Date of Birth"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   390
         Width           =   2055
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         Caption         =   "Search +/-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   1290
         Width           =   930
      End
      Begin VB.Label lblWithin 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label lblYears 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1710
         TabIndex        =   5
         Top             =   1290
         Width           =   405
      End
   End
   Begin VB.TextBox txtDoB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date of Birth"
      Height          =   195
      Left            =   450
      TabIndex        =   0
      Top             =   210
      Width           =   885
   End
End
Attribute VB_Name = "frmGetDoB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

50320     txtDoB = ""
50330     Unload Me

End Sub

Private Sub cmdOK_Click()

50340     Me.Hide

End Sub

Private Sub optExact_Click()

50350     lblSearch.Enabled = False
50360     lblSearch.Font.Bold = False
50370     lblWithin.Enabled = False
50380     lblYears.Enabled = False
50390     lblYears.Font.Bold = False
50400     udWithin.Enabled = False

End Sub

Private Sub optFuzzy_Click()

50410     lblSearch.Enabled = True
50420     lblSearch.Font.Bold = True
50430     lblWithin.Enabled = True
50440     lblYears.Enabled = True
50450     lblYears.Font.Bold = True
50460     udWithin.Enabled = True

End Sub

Private Sub txtDoB_Change()

50470     txtDoB = Convert62Date(txtDoB, BACKWARD)
50480     If IsDate(txtDoB) Then
50490         cmdOK.Enabled = True
50500     Else
50510         cmdOK.Enabled = False
50520     End If

End Sub

Private Sub udWithin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

50530     If lblWithin = "1" Then
50540         lblYears = "Year"
50550     Else
50560         lblYears = "Years"
50570     End If

End Sub


