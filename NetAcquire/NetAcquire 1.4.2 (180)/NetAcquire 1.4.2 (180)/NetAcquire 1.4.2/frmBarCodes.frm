VERSION 5.00
Begin VB.Form frmBarCodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - BarCodes"
   ClientHeight    =   5625
   ClientLeft      =   1005
   ClientTop       =   900
   ClientWidth     =   4455
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMonoSpot 
      Height          =   315
      Left            =   1230
      TabIndex        =   10
      Top             =   5010
      Width           =   1575
   End
   Begin VB.TextBox txtFBC 
      Height          =   315
      Left            =   1230
      TabIndex        =   7
      Top             =   4020
      Width           =   1575
   End
   Begin VB.TextBox txtRetics 
      Height          =   315
      Left            =   1230
      TabIndex        =   9
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtESR 
      Height          =   315
      Left            =   1230
      TabIndex        =   8
      Top             =   4350
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1185
      Left            =   3030
      Picture         =   "frmBarCodes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2070
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtB 
      Height          =   315
      Left            =   1230
      TabIndex        =   6
      Top             =   3540
      Width           =   1575
   End
   Begin VB.TextBox txtA 
      Height          =   315
      Left            =   1230
      TabIndex        =   5
      Top             =   3210
      Width           =   1575
   End
   Begin VB.TextBox txtFasting 
      Height          =   315
      Left            =   1230
      TabIndex        =   4
      Top             =   2790
      Width           =   1575
   End
   Begin VB.TextBox txtRandom 
      Height          =   315
      Left            =   1230
      TabIndex        =   3
      Top             =   2460
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1185
      Left            =   3030
      Picture         =   "frmBarCodes.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4140
      Width           =   1095
   End
   Begin VB.TextBox txtClear 
      Height          =   315
      Left            =   1230
      TabIndex        =   2
      Top             =   2010
      Width           =   1575
   End
   Begin VB.TextBox txtSave 
      Height          =   315
      Left            =   1230
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtCancel 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Retics"
      Height          =   195
      Left            =   720
      TabIndex        =   24
      Top             =   4740
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "MonoSpot"
      Height          =   195
      Left            =   420
      TabIndex        =   23
      Top             =   5070
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "ESR"
      Height          =   195
      Left            =   825
      TabIndex        =   22
      Top             =   4410
      Width           =   330
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "FBC"
      Height          =   195
      Left            =   855
      TabIndex        =   21
      Top             =   4080
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Random"
      Height          =   195
      Left            =   555
      TabIndex        =   20
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Fasting"
      Height          =   195
      Left            =   645
      TabIndex        =   19
      Top             =   2850
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Set Analyser 'B'"
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   3570
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Set Analyser 'A'"
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   3270
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cancel"
      Height          =   195
      Left            =   660
      TabIndex        =   16
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Save"
      Height          =   195
      Left            =   780
      TabIndex        =   15
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Clear"
      Height          =   195
      Left            =   795
      TabIndex        =   14
      Top             =   2070
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scan Entries using BarCode Reader"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Left            =   1260
      TabIndex        =   13
      Top             =   90
      Width           =   1560
   End
End
Attribute VB_Name = "frmBarCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

64690     Unload Me

End Sub


Private Sub cmdSave_Click()

          Dim sql As String

64700     On Error GoTo cmdSave_Click_Error

64710     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtCancel & "' " & _
              "Where Text = 'ctlCancel'"
64720     Cnxn(0).Execute sql

64730     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtSave & "' " & _
              "Where Text = 'ctlSave'"
64740     Cnxn(0).Execute sql

64750     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtClear & "' " & _
              "Where Text = 'ctlClear'"
64760     Cnxn(0).Execute sql

64770     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtRandom & "' " & _
              "Where Text = 'ctlRandom'"
64780     Cnxn(0).Execute sql

64790     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtFasting & "' " & _
              "Where Text = 'ctlFasting'"
64800     Cnxn(0).Execute sql

64810     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtA & "' " & _
              "Where Text = 'ctlA'"
64820     Cnxn(0).Execute sql

64830     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtB & "' " & _
              "Where Text = 'ctlB'"
64840     Cnxn(0).Execute sql

64850     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtFBC & "' " & _
              "Where Text = 'ctlFBC'"
64860     Cnxn(0).Execute sql

64870     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtESR & "' " & _
              "Where Text = 'ctlESR'"
64880     Cnxn(0).Execute sql

64890     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtRetics & "' " & _
              "Where Text = 'ctlRetics'"
64900     Cnxn(0).Execute sql

64910     sql = "Update BarCodeControl " & _
              "Set Code = '" & txtMonoSpot & "' " & _
              "Where Text = 'ctlMonoSpot'"
64920     Cnxn(0).Execute sql

64930     cmdSave.Visible = False

64940     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

64950     intEL = Erl
64960     strES = Err.Description
64970     LogError "frmBarCodes", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

          Dim sql As String
          Dim tb As Recordset

64980     On Error GoTo Form_Load_Error

64990     sql = "Select * from BarCodeControl"
65000     Set tb = New Recordset
65010     RecOpenClient 0, tb, sql
65020     With tb
        
65030         Do While Not .EOF
65040             Select Case UCase$(Trim$(!Text))
                      Case "CTLCANCEL": txtCancel = !Code & ""
65050                 Case "CTLSAVE": txtSave = !Code & ""
65060                 Case "CTLCLEAR": txtClear = !Code & ""
65070                 Case "CTLRANDOM": txtRandom = !Code & ""
65080                 Case "CTLFASTING": txtFasting = !Code & ""
65090                 Case "CTLA": txtA = !Code & ""
65100                 Case "CTLB": txtB = !Code & ""
65110                 Case "CTLFBC": txtFBC = !Code & ""
65120                 Case "CTLESR": txtESR = !Code & ""
65130                 Case "CTLRETICS": txtRetics = !Code & ""
65140                 Case "CTLMONOSPOT": txtMonoSpot = !Code & ""
65150             End Select
65160             .MoveNext
65170         Loop
65180     End With

65190     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

65200     intEL = Erl
65210     strES = Err.Description
65220     LogError "frmBarCodes", "Form_Load", intEL, strES, sql


End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

65230     If cmdSave.Visible Then
65240         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
65250             Cancel = True
65260         End If
65270     End If

End Sub


Private Sub txtA_KeyPress(KeyAscii As Integer)

65280     cmdSave.Visible = True

End Sub


Private Sub txtB_KeyPress(KeyAscii As Integer)

65290     cmdSave.Visible = True

End Sub


Private Sub txtCancel_KeyPress(KeyAscii As Integer)

65300     cmdSave.Visible = True

End Sub


Private Sub txtClear_KeyPress(KeyAscii As Integer)

65310     cmdSave.Visible = True

End Sub


Private Sub txtESR_KeyPress(KeyAscii As Integer)

65320     cmdSave.Visible = True

End Sub


Private Sub txtFasting_KeyPress(KeyAscii As Integer)

65330     cmdSave.Visible = True

End Sub


Private Sub txtFBC_KeyPress(KeyAscii As Integer)

65340     cmdSave.Visible = True

End Sub


Private Sub txtMonoSpot_KeyPress(KeyAscii As Integer)

65350     cmdSave.Visible = True

End Sub


Private Sub txtRandom_KeyPress(KeyAscii As Integer)

65360     cmdSave.Visible = True

End Sub


Private Sub txtRetics_KeyPress(KeyAscii As Integer)

65370     cmdSave.Visible = True

End Sub


Private Sub txtSave_KeyPress(KeyAscii As Integer)

65380     cmdSave.Visible = True

End Sub


