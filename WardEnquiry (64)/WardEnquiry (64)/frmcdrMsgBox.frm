VERSION 5.00
Begin VB.Form fcdrMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FcdrMsgBox"
   ClientHeight    =   2115
   ClientLeft      =   1920
   ClientTop       =   2505
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Picture         =   "frmcdrMsgBox.frx":0000
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   32
      Left            =   180
      Picture         =   "frmcdrMsgBox.frx":0442
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   64
      Left            =   180
      Picture         =   "frmcdrMsgBox.frx":0884
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   16
      Left            =   180
      Picture         =   "frmcdrMsgBox.frx":0CC6
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

10    RetVal = ReturnValue

End Property



Private Sub b_Click(Index As Integer)

10    ReturnValue = Index
20    Unload Me

End Sub


Private Sub Form_Activate()
10    SingleUserUpdateLoggedOn UserName

20    If mDefaultButton > 0 Then
30      b(mDefaultButton).Default = True
40    End If

50    If mMsgFontSize <> 0 Then
60      lblMessage.Font.Size = mMsgFontSize
70    End If
  
80    Select Case mIcon
        Case 16, 32, 48, 64: i(mIcon).Visible = True
90      Case Else:
100   End Select
  
110   lblMessage = mMessage
  
120   Select Case mButtons
        Case 0: 'MB_OK 0 Display OK button only.
130       b(1).Visible = True
140       b(1).Cancel = True
150     Case 1: 'MB_OKCANCEL 1 Display OK and Cancel buttons.
160       b(1).Visible = True
170       b(2).Visible = True
180       b(2).Cancel = True
          'Select Case DefaultButton
          '  Case 0: .b(1).Default = True
          '  Case 256: .b(2).Default = True
          'End Select
  
190     Case 2: 'MB_ABORTRETRYIGNORE 2 Display Abort, Retry, and Ignore buttons.
200       b(3).Visible = True
210       b(4).Visible = True
220       b(5).Visible = True
      '      Select Case DefaultButton
      '        Case 0: .b(3).Default = True
      '        Case 256: .b(4).Default = True
      '        Case 512: .b(5).Default = True
      '      End Select
  
230     Case 3: 'MB_YESNOCANCEL  3 Display Yes, No, and Cancel buttons.
240       b(6).Visible = True
250       b(7).Visible = True
260       b(2).Visible = True
270       b(2).Cancel = True
          'Select Case DefaultButton
          '  Case 0: .b(6).Default = True
          '  Case 256: .b(7).Default = True
          '  Case 512: .b(2).Default = True
          'End Select
  
280     Case 4: 'MB_YESNO  4 Display Yes and No buttons.
290       b(6).Visible = True
300       b(7).Visible = True
          'Select Case DefaultButton
          '  Case 0: .b(6).Default = True
          '  Case 256: .b(7).Default = True
          'End Select
  
310     Case 5: 'MB_RETRYCANCEL  5 Display Retry and Cancel buttons.
320       b(4).Visible = True
330       b(2).Visible = True
340       b(2).Cancel = True
          'Select Case DefaultButton
          '  Case 0: .b(4).Default = True
          '  Case 256: .b(2).Default = True
          'End Select

350   End Select

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


