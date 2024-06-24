VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckMatch 
   Caption         =   "NetAcquire"
   ClientHeight    =   4230
   ClientLeft      =   4245
   ClientTop       =   1740
   ClientWidth     =   5160
   Icon            =   "frmCheckMatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   5160
   Begin VB.CommandButton cmdHistory 
      Caption         =   "View History"
      Height          =   1155
      Left            =   900
      Picture         =   "frmCheckMatch.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2430
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1155
      Left            =   3270
      Picture         =   "frmCheckMatch.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2430
      Width           =   1035
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Default         =   -1  'True
      Height          =   1155
      Left            =   2070
      Picture         =   "frmCheckMatch.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2430
      Width           =   1035
   End
   Begin VB.TextBox txtBack 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   1890
      Width           =   2385
   End
   Begin VB.TextBox txtFront 
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Top             =   1050
      Width           =   2385
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   900
      TabIndex        =   6
      Top             =   3750
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1260
      X2              =   3660
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label lblPackNumber 
      Caption         =   "Pack # 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   1710
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Scan Back of Pack"
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   1650
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Scan Front of Pack"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1380
   End
End
Attribute VB_Name = "frmCheckMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdCheck_Click()

      Dim f As String
      Dim b As String
      Dim s As String
        
10    f = UCase$(Trim$(txtFront))
20    b = UCase$(Trim$(txtBack))

30    If f = "" Or b = "" Then
40      Exit Sub
50    End If

60    If Left$(f, 1) = "=" Then 'Barcode scanning entry
70     s = ISOmod37_2(Mid$(f, 2, 13))
80     f = Mid$(f, 2, 13) & " " & s
90    End If

100   If Left$(b, 1) = "=" Then 'Barcode scanning entry
110    s = ISOmod37_2(Mid$(b, 2, 13))
120    b = Mid$(b, 2, 13) & " " & s
130   End If

140   If f = b Then
150     LogReasonWhy "Check Front/Back - OK (" & b & ")", "XM"
160     iMsg "Front and Back of Pack match." & vbCr & "You may Proceed.", vbInformation
170     If TimedOut Then Unload Me: Exit Sub
180     Unload Me: Exit Sub
190   Else
200     LogReasonWhy "Check Front/Back - Mis-match (" & f & "/" & b & ")", "XM"
210     iMsg "Scan Mis-match (" & f & " / " & b & ")" & vbCrLf & "Re-enter", vbCritical, , vbRed, 16
220     If TimedOut Then Unload Me: Exit Sub
230   End If
  
240   txtFront = ""
250   txtBack = ""
260   txtFront.SetFocus

End Sub

Private Sub cmdCheck_GotFocus()
10    cmdCheck_Click
End Sub


Private Sub cmdHistory_Click()

10    frmUnlockReasons.Show 1

End Sub


Public Property Let DisplayNumber(ByVal sNewValue As String)

10    lblPackNumber = sNewValue
20    lblPackNumber.Visible = True
30    Line1.Visible = True

End Property



Private Sub txtBack_LostFocus()
'Dim s As String

'If Left$(txtBack, 1) <> "=" Then
'        txtBack = ""
'Else
'       txtBack = UCase(txtBack)
'       s = ISOmod37_2(Mid$(txtBack, 2, 13))
'       txtBack = Mid$(txtBack, 2, 13) & " " & s
'End If

End Sub

Private Sub txtFront_LostFocus()
Dim s As String

If Left$(txtFront, 1) <> "=" Then
        txtFront = ""
Else
       txtFront = UCase(txtFront)
       s = ISOmod37_2(Mid$(txtFront, 2, 13))
       txtFront = Mid$(txtFront, 2, 13) & " " & s
End If

End Sub
