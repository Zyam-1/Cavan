VERSION 5.00
Begin VB.Form fcdrMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   2115
   ClientLeft      =   1590
   ClientTop       =   2775
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton b 
      Caption         =   "&Ignore"
      Height          =   525
      Index           =   5
      Left            =   4410
      TabIndex        =   7
      Top             =   1350
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&Retry"
      Height          =   525
      Index           =   4
      Left            =   4410
      TabIndex        =   6
      Top             =   750
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&Abort"
      Height          =   525
      Index           =   3
      Left            =   4410
      TabIndex        =   5
      Top             =   150
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&Yes"
      Height          =   525
      Index           =   6
      Left            =   4410
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&Cancel"
      Height          =   525
      Index           =   2
      Left            =   4410
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "O. K."
      Height          =   525
      Index           =   1
      Left            =   4410
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&No"
      Height          =   525
      Index           =   7
      Left            =   4410
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   48
      Left            =   180
      Picture         =   "fcdrMsgBox.frx":0000
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   32
      Left            =   180
      Picture         =   "fcdrMsgBox.frx":0442
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   64
      Left            =   180
      Picture         =   "fcdrMsgBox.frx":0884
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   16
      Left            =   180
      Picture         =   "fcdrMsgBox.frx":0CC6
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   1695
      Left            =   840
      TabIndex        =   1
      Top             =   150
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fcdrMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ReturnValue As Integer

Private mDefaultButton As Long
Private mMsgFontSize As Long

Private mButtons As Integer

Private mIcon As Integer
Private mMessage As String



Public Property Get RetVal() As Integer

31840     RetVal = ReturnValue

End Property


Private Sub b_Click(Index As Integer)

31850     ReturnValue = Index
31860     Unload Me

End Sub


Private Sub Form_Activate()

31870     If mDefaultButton > 0 Then
31880         b(mDefaultButton).Default = True
31890     End If

31900     If mMsgFontSize <> 0 Then
31910         lblMessage.Font.size = mMsgFontSize
31920     End If
        
31930     Select Case mIcon
              Case vbCritical:
31940             i(mIcon).Visible = True
31950             PlaySound sysOptSoundCritical(0), ByVal 0&, SND_FILENAME Or SND_ASYNC
31960         Case vbExclamation:
31970             i(mIcon).Visible = True
31980             PlaySound sysOptSoundSevere(0), ByVal 0&, SND_FILENAME Or SND_ASYNC
31990         Case vbInformation:
32000             i(mIcon).Visible = True
32010             PlaySound sysOptSoundInformation(0), ByVal 0&, SND_FILENAME Or SND_ASYNC
32020         Case vbQuestion:
32030             i(mIcon).Visible = True
32040             PlaySound sysOptSoundQuestion(0), ByVal 0&, SND_FILENAME Or SND_ASYNC
32050         Case Else:
32060     End Select
        
32070     lblMessage = mMessage
        
32080     Select Case mButtons
              Case 0: 'MB_OK 0 Display OK button only.
32090             b(1).Visible = True
32100             b(1).Cancel = True
32110         Case 1: 'MB_OKCANCEL 1 Display OK and Cancel buttons.
32120             b(1).Visible = True
32130             b(2).Visible = True
32140             b(2).Cancel = True
                  'Select Case DefaultButton
                  '  Case 0: .b(1).Default = True
                  '  Case 256: .b(2).Default = True
                  'End Select
        
32150         Case 2: 'MB_ABORTRETRYIGNORE 2 Display Abort, Retry, and Ignore buttons.
32160             b(3).Visible = True
32170             b(4).Visible = True
32180             b(5).Visible = True
                  '      Select Case DefaultButton
                  '        Case 0: .b(3).Default = True
                  '        Case 256: .b(4).Default = True
                  '        Case 512: .b(5).Default = True
                  '      End Select
        
32190         Case 3: 'MB_YESNOCANCEL  3 Display Yes, No, and Cancel buttons.
32200             b(6).Visible = True
32210             b(7).Visible = True
32220             b(2).Visible = True
32230             b(2).Cancel = True
                  'Select Case DefaultButton
                  '  Case 0: .b(6).Default = True
                  '  Case 256: .b(7).Default = True
                  '  Case 512: .b(2).Default = True
                  'End Select
        
32240         Case 4: 'MB_YESNO  4 Display Yes and No buttons.
32250             b(6).Visible = True
32260             b(7).Visible = True
                  'Select Case DefaultButton
                  '  Case 0: .b(6).Default = True
                  '  Case 256: .b(7).Default = True
                  'End Select
        
32270         Case 5: 'MB_RETRYCANCEL  5 Display Retry and Cancel buttons.
32280             b(4).Visible = True
32290             b(2).Visible = True
32300             b(2).Cancel = True
                  'Select Case DefaultButton
                  '  Case 0: .b(4).Default = True
                  '  Case 256: .b(2).Default = True
                  'End Select

32310     End Select

End Sub

Public Property Let DefaultButton(ByVal lngButton As Long)

32320     mDefaultButton = lngButton

End Property

Public Property Let DisplayButtons(ByVal intButtons As Integer)

32330     mButtons = intButtons

End Property

Public Property Let Message(ByVal strMessage As String)

32340     mMessage = strMessage

End Property


Public Property Let ShowIcon(ByVal intIcon As Integer)

32350     mIcon = intIcon

End Property


Public Property Let MsgFontSize(ByVal FntSize As Long)

32360     mMsgFontSize = FntSize

End Property


