VERSION 5.00
Begin VB.Form frmSetLIH 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Set H/I/L"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   705
   End
   Begin VB.TextBox txtL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   5
      Top             =   720
      Width           =   705
   End
   Begin VB.TextBox txtI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   10
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   11
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   16
      Top             =   1710
      Width           =   705
   End
   Begin VB.TextBox txtH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   17
      Top             =   1710
      Width           =   705
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   6450
      Picture         =   "frmSetLIH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   900
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3930
      TabIndex        =   15
      Top             =   1710
      Width           =   705
   End
   Begin VB.TextBox txtH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3090
      TabIndex        =   14
      Top             =   1710
      Width           =   705
   End
   Begin VB.TextBox txtH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2250
      TabIndex        =   13
      Top             =   1710
      Width           =   705
   End
   Begin VB.TextBox txtH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   12
      Top             =   1710
      Width           =   705
   End
   Begin VB.TextBox txtI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3930
      TabIndex        =   9
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3090
      TabIndex        =   8
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2250
      TabIndex        =   7
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   6
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox txtL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3930
      TabIndex        =   3
      Top             =   720
      Width           =   705
   End
   Begin VB.TextBox txtL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3090
      TabIndex        =   2
      Top             =   720
      Width           =   705
   End
   Begin VB.TextBox txtL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2250
      TabIndex        =   1
      Top             =   720
      Width           =   705
   End
   Begin VB.TextBox txtL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   0
      Top             =   720
      Width           =   705
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   7620
      Picture         =   "frmSetLIH.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   900
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "5+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   5040
      TabIndex        =   28
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "6+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   5880
      TabIndex        =   27
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "4+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4170
      TabIndex        =   26
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "3+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3330
      TabIndex        =   25
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "2+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2490
      TabIndex        =   24
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1680
      TabIndex        =   23
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Haemolysed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   22
      Top             =   1740
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Icteric"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   21
      Top             =   1260
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lipaemic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   690
      TabIndex        =   20
      Top             =   765
      Width           =   630
   End
End
Attribute VB_Name = "frmSetLIH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Clear()

      Dim n As Integer

57410 For n = 1 To 6
57420     txtL(n) = ""
57430     txtI(n) = ""
57440     txtH(n) = ""
57450 Next

End Sub

Private Sub LoadDetails()

      Dim X As Integer
      Dim txt As TextBox
      Dim v As Single

57460 On Error GoTo LoadDetails_Error

57470 Clear

57480 For X = 1 To 6
57490   Set txt = txtL(X)
57500   v = Val(GetOptionSetting("LIH_L" & Format$(X), "0"))
57510   If v > 0 Then
57520     txt = Format$(v)
57530   End If
57540 Next
57550 For X = 1 To 6
57560   Set txt = txtI(X)
57570   v = Val(GetOptionSetting("LIH_I" & Format$(X), "0"))
57580   If v > 0 Then
57590     txt = Format$(v)
57600   End If
57610 Next
57620 For X = 1 To 6
57630   Set txt = txtH(X)
57640   v = Val(GetOptionSetting("LIH_H" & Format$(X), "0"))
57650   If v > 0 Then
57660     txt = Format$(v)
57670   End If
57680 Next
        
57690 Exit Sub

LoadDetails_Error:

      Dim strES As String
      Dim intEL As Integer

57700 intEL = Erl
57710 strES = Err.Description
57720 LogError "frmSetLIH", "LoadDetails", intEL, strES

End Sub

Private Sub cmdExit_Click()

57730 Unload Me

End Sub


Private Sub cmdSave_Click()

            Dim X As Integer
            Dim txt As TextBox

57740 For X = 1 To 6
57750   Set txt = txtL(X)
57760   SaveOptionSetting "LIH_L" & Format$(X), txt
57770 Next
57780 For X = 1 To 6
57790   Set txt = txtI(X)
57800   SaveOptionSetting "LIH_I" & Format$(X), txt
57810 Next
57820 For X = 1 To 6
57830   Set txt = txtH(X)
57840   SaveOptionSetting "LIH_H" & Format$(X), txt
57850 Next

End Sub

Private Sub Form_Load()

57860 LoadDetails

End Sub

Private Sub txtH_GotFocus(Index As Integer)

57870 txtH(Index).SelStart = 0
57880 txtH(Index).SelLength = Len(txtH(Index))

End Sub

Private Sub txtH_KeyPress(Index As Integer, KeyAscii As Integer)

57890 cmdSave.Visible = True

End Sub

Private Sub txtI_KeyPress(Index As Integer, KeyAscii As Integer)

57900 cmdSave.Visible = True

End Sub

Private Sub txtL_KeyPress(Index As Integer, KeyAscii As Integer)

57910 cmdSave.Visible = True

End Sub

