VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fcdrMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FcdrMsgBox"
   ClientHeight    =   2430
   ClientLeft      =   1590
   ClientTop       =   2775
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
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
      Default         =   -1  'True
      Height          =   525
      Index           =   7
      Left            =   4410
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   210
      TabIndex        =   8
      Top             =   2100
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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



Public Property Get retval() As Integer

10    retval = ReturnValue

End Property


Private Sub b_Click(Index As Integer)

10    ReturnValue = Index
20    Unload Me

End Sub


Private Sub Form_Activate()

10    If mDefaultButton > 0 Then
20      b(mDefaultButton).Default = True
30    End If

40    If mMsgFontSize <> 0 Then
50      lblMessage.Font.Size = mMsgFontSize
60    End If
  
70    Select Case mIcon
        Case vbCritical:
80        i(mIcon).Visible = True
90        PlaySound sysOptSoundCritical(0), ByVal 0&, SND_FILENAME Or SND_ASYNC
100     Case vbExclamation:
110       i(mIcon).Visible = True
120       PlaySound sysOptSoundSevere(0), ByVal 0&, SND_FILENAME Or SND_ASYNC
130     Case vbInformation:
140       i(mIcon).Visible = True
150       PlaySound sysOptSoundInformation(0), ByVal 0&, SND_FILENAME Or SND_ASYNC
160     Case vbQuestion:
170       i(mIcon).Visible = True
180       PlaySound sysOptSoundQuestion(0), ByVal 0&, SND_FILENAME Or SND_ASYNC
190     Case Else:
200   End Select
  
210   lblMessage = mMessage
  
220   Select Case mButtons
        Case 0: 'MB_OK 0 Display OK button only.
230       b(1).Visible = True
240       b(1).Cancel = True
250     Case 1: 'MB_OKCANCEL 1 Display OK and Cancel buttons.
260       b(1).Visible = True
270       b(2).Visible = True
280       b(2).Cancel = True
          'Select Case DefaultButton
          '  Case 0: .b(1).Default = True
          '  Case 256: .b(2).Default = True
          'End Select
  
290     Case 2: 'MB_ABORTRETRYIGNORE 2 Display Abort, Retry, and Ignore buttons.
300       b(3).Visible = True
310       b(4).Visible = True
320       b(5).Visible = True
      '      Select Case DefaultButton
      '        Case 0: .b(3).Default = True
      '        Case 256: .b(4).Default = True
      '        Case 512: .b(5).Default = True
      '      End Select
  
330     Case 3: 'MB_YESNOCANCEL  3 Display Yes, No, and Cancel buttons.
340       b(6).Visible = True
350       b(7).Visible = True
360       b(2).Visible = True
370       b(2).Cancel = True
          'Select Case DefaultButton
          '  Case 0: .b(6).Default = True
          '  Case 256: .b(7).Default = True
          '  Case 512: .b(2).Default = True
          'End Select
  
380     Case 4: 'MB_YESNO  4 Display Yes and No buttons.
390       b(6).Visible = True
400       b(7).Visible = True
          'Select Case DefaultButton
          '  Case 0: .b(6).Default = True
410        b(7).Default = True
          'End Select
  
420     Case 5: 'MB_RETRYCANCEL  5 Display Retry and Cancel buttons.
430       b(4).Visible = True
440       b(2).Visible = True
450       b(2).Cancel = True
          'Select Case DefaultButton
          '  Case 0: .b(4).Default = True
          '  Case 256: .b(2).Default = True
          'End Select

460   End Select

End Sub

Public Property Let DefaultButton(ByVal lngButton As Long)

10    mDefaultButton = lngButton

End Property

Public Property Let DisplayButtons(ByVal intButtons As Integer)

10    mButtons = intButtons

End Property

Public Property Let Message(ByVal strMessage As String)

10    mMessage = strMessage

End Property


Public Property Let ShowIcon(ByVal intIcon As Integer)

10    mIcon = intIcon

End Property


Public Property Let MsgFontSize(ByVal FntSize As Long)

10    mMsgFontSize = FntSize

End Property


