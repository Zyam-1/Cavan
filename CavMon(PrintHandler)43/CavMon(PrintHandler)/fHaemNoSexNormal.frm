VERSION 5.00
Begin VB.Form fHaemNoSexNormal 
   Caption         =   "NetAcquire"
   ClientHeight    =   5610
   ClientLeft      =   1635
   ClientTop       =   1155
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6465
   Begin VB.TextBox tNeut 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   15
      Top             =   4410
      Width           =   2265
   End
   Begin VB.TextBox tRDW 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   14
      Top             =   3420
      Width           =   2265
   End
   Begin VB.TextBox tMCV 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   13
      Top             =   2430
      Width           =   2265
   End
   Begin VB.TextBox tMCH 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   12
      Top             =   2760
      Width           =   2265
   End
   Begin VB.TextBox tRBC 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   11
      Top             =   1440
      Width           =   2265
   End
   Begin VB.TextBox tWBC 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   10
      Top             =   1110
      Width           =   2265
   End
   Begin VB.TextBox tMCHC 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   9
      Top             =   3090
      Width           =   2265
   End
   Begin VB.TextBox tHb 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   1770
      Width           =   2265
   End
   Begin VB.TextBox tPlt 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   3750
      Width           =   2265
   End
   Begin VB.TextBox tMPV 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Top             =   4080
      Width           =   2265
   End
   Begin VB.TextBox tLymp 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   4740
      Width           =   2265
   End
   Begin VB.TextBox tMono 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   5070
      Width           =   2265
   End
   Begin VB.TextBox tHct 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   2100
      Width           =   2265
   End
   Begin VB.CommandButton bSave 
      Height          =   525
      Left            =   4650
      Picture         =   "fHaemNoSexNormal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   1860
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Height          =   525
      Left            =   4650
      Picture         =   "fHaemNoSexNormal.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   2910
      Width           =   1245
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "WBC"
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
      Left            =   1140
      TabIndex        =   28
      Top             =   1140
      Width           =   435
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Hb"
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
      Left            =   1320
      TabIndex        =   27
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "MCV"
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
      Left            =   1170
      TabIndex        =   26
      Top             =   2460
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "MCHC"
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
      Left            =   1035
      TabIndex        =   25
      Top             =   3120
      Width           =   540
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Plt"
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
      Left            =   1335
      TabIndex        =   24
      Top             =   3750
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Neutrophils"
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
      Left            =   600
      TabIndex        =   23
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Lymphocytes"
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
      Left            =   465
      TabIndex        =   22
      Top             =   4770
      Width           =   1110
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Monocytes"
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
      Left            =   645
      TabIndex        =   21
      Top             =   5130
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "MPV"
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
      Left            =   1170
      TabIndex        =   20
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "RDW"
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
      Left            =   1110
      TabIndex        =   19
      Top             =   3450
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MCH"
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
      Left            =   1155
      TabIndex        =   18
      Top             =   2790
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hct"
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
      Left            =   1260
      TabIndex        =   17
      Top             =   2130
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "RBC"
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
      Left            =   1185
      TabIndex        =   16
      Top             =   1470
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "These Normal Ranges only apply when the Age/Sex Related option is Disabled."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   450
      TabIndex        =   0
      Top             =   180
      Width           =   5685
   End
End
Attribute VB_Name = "fHaemNoSexNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

10    Unload Me

End Sub

Private Sub bSave_Click()

10    SaveSetting "NetAcquire", "PrintHandler", "WBC", tWBC
20    SaveSetting "NetAcquire", "PrintHandler", "RBC", tRBC
30    SaveSetting "NetAcquire", "PrintHandler", "HB", tHb
40    SaveSetting "NetAcquire", "PrintHandler", "HCT", tHct
50    SaveSetting "NetAcquire", "PrintHandler", "MCV", tMCV
60    SaveSetting "NetAcquire", "PrintHandler", "MCH", tMCH
70    SaveSetting "NetAcquire", "PrintHandler", "MCHC", tMCHC
80    SaveSetting "NetAcquire", "PrintHandler", "RDW", tRDW
90    SaveSetting "NetAcquire", "PrintHandler", "PLT", tPlt
100   SaveSetting "NetAcquire", "PrintHandler", "MPV", tMPV
110   SaveSetting "NetAcquire", "PrintHandler", "NEUT", tNeut
120   SaveSetting "NetAcquire", "PrintHandler", "LYMP", tLymp
130   SaveSetting "NetAcquire", "PrintHandler", "MONO", tMono

140   bSave.Visible = False

End Sub

Private Sub Form_Load()

10    tWBC = GetSetting("NetAcquire", "PrintHandler", "WBC")
20    tRBC = GetSetting("NetAcquire", "PrintHandler", "RBC")
30    tHb = GetSetting("NetAcquire", "PrintHandler", "HB")
40    tHct = GetSetting("NetAcquire", "PrintHandler", "HCT")
50    tMCV = GetSetting("NetAcquire", "PrintHandler", "MCV")
60    tMCH = GetSetting("NetAcquire", "PrintHandler", "MCH")
70    tMCHC = GetSetting("NetAcquire", "PrintHandler", "MCHC")
80    tRDW = GetSetting("NetAcquire", "PrintHandler", "RDW")
90    tPlt = GetSetting("NetAcquire", "PrintHandler", "PLT")
100   tMPV = GetSetting("NetAcquire", "PrintHandler", "MPV")
110   tNeut = GetSetting("NetAcquire", "PrintHandler", "NEUT")
120   tLymp = GetSetting("NetAcquire", "PrintHandler", "LYMP")
130   tMono = GetSetting("NetAcquire", "PrintHandler", "MONO")

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If bSave.Visible Then
20      If iMsg("Cancel without Saving?", vbYesNo) = vbNo Then
30        Cancel = True
40      End If
50    End If

End Sub


Private Sub tHb_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tHct_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tLymp_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMCH_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMCHC_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMCV_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMono_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMPV_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tNeut_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tPlt_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tRBC_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tRDW_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tWBC_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


