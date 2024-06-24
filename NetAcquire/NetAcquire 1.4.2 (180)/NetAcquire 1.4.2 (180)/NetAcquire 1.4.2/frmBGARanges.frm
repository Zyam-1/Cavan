VERSION 5.00
Begin VB.Form frmBGARanges 
   Caption         =   "NetAcquire - Blood Gas - Normal Ranges"
   ClientHeight    =   5115
   ClientLeft      =   1290
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   5580
   Begin VB.CommandButton cmdHistory 
      Caption         =   "View History"
      Height          =   855
      Left            =   3900
      Picture         =   "frmBGARanges.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3990
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   705
      Left            =   3900
      Picture         =   "frmBGARanges.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1380
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   3900
      Picture         =   "frmBGARanges.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2670
      Width           =   1065
   End
   Begin VB.TextBox txtPh 
      Alignment       =   2  'Center
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
      Height          =   330
      Index           =   1
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   1
      Top             =   600
      Width           =   1050
   End
   Begin VB.TextBox txtPco2 
      Alignment       =   2  'Center
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
      Height          =   330
      Index           =   1
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1005
      Width           =   1050
   End
   Begin VB.TextBox txtPo2 
      Alignment       =   2  'Center
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
      Height          =   330
      Index           =   1
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1410
      Width           =   1050
   End
   Begin VB.TextBox txtHco3 
      Alignment       =   2  'Center
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
      Height          =   330
      Index           =   1
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1815
      Width           =   1050
   End
   Begin VB.TextBox txtO2Sat 
      Alignment       =   2  'Center
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
      Height          =   330
      Index           =   1
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   11
      Top             =   2625
      Width           =   1050
   End
   Begin VB.TextBox txtBE 
      Alignment       =   2  'Center
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
      Height          =   330
      Index           =   1
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2220
      Width           =   1050
   End
   Begin VB.TextBox txtTotCo2 
      Alignment       =   2  'Center
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
      Height          =   330
      Index           =   1
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   13
      Top             =   3030
      Width           =   1050
   End
   Begin VB.TextBox txtPh 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   0
      Top             =   600
      Width           =   1050
   End
   Begin VB.TextBox txtPco2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1005
      Width           =   1050
   End
   Begin VB.TextBox txtPo2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1410
      Width           =   1050
   End
   Begin VB.TextBox txtHco3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1815
      Width           =   1050
   End
   Begin VB.TextBox txtO2Sat 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   10
      Top             =   2625
      Width           =   1050
   End
   Begin VB.TextBox txtBE 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2220
      Width           =   1050
   End
   Begin VB.TextBox txtTotCo2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   1065
      MaxLength       =   6
      TabIndex        =   12
      Top             =   3030
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   450
      X2              =   4950
      Y1              =   3630
      Y2              =   3630
   End
   Begin VB.Label lblTimeAmended 
      Caption         =   "88:88"
      Height          =   195
      Left            =   2220
      TabIndex        =   31
      Top             =   4410
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "at"
      Height          =   195
      Left            =   2040
      TabIndex        =   30
      Top             =   4410
      Width           =   135
   End
   Begin VB.Label lblDateAmended 
      Caption         =   "88/88/8888"
      Height          =   195
      Left            =   1050
      TabIndex        =   28
      Top             =   4410
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "on"
      Height          =   195
      Left            =   750
      TabIndex        =   27
      Top             =   4410
      Width           =   180
   End
   Begin VB.Label lblAmendedBy 
      Caption         =   "qqqqq"
      Height          =   195
      Left            =   1050
      TabIndex        =   26
      Top             =   4140
      Width           =   2370
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Entered By"
      Height          =   195
      Left            =   150
      TabIndex        =   25
      Top             =   4140
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "High"
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
      Height          =   195
      Left            =   2700
      TabIndex        =   23
      Top             =   330
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Low"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1410
      TabIndex        =   22
      Top             =   330
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "pH"
      Height          =   195
      Index           =   1
      Left            =   645
      TabIndex        =   21
      Top             =   675
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "pCO2"
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   20
      Top             =   1080
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "pO2"
      Height          =   195
      Index           =   3
      Left            =   555
      TabIndex        =   19
      Top             =   1485
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "HCO3"
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   18
      Top             =   1890
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BE"
      Height          =   195
      Index           =   5
      Left            =   645
      TabIndex        =   17
      Top             =   2295
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "O2Sat"
      Height          =   195
      Index           =   6
      Left            =   405
      TabIndex        =   16
      Top             =   2700
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tot CO2"
      Height          =   195
      Index           =   7
      Left            =   255
      TabIndex        =   15
      Top             =   3105
      Width           =   600
   End
End
Attribute VB_Name = "frmBGARanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

950       If cmdSave.Visible Then
960           If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
970               Exit Sub
980           End If
990       End If

1000      Unload Me

End Sub

Private Sub cmdHistory_Click()

1010      frmBGARangeHistory.Show 1

End Sub

Private Sub cmdSave_Click()

          Dim sql As String

1020      On Error GoTo cmdSave_Click_Error

1030      sql = "Insert into BGADefinitions " & _
              "(pH, PCO2, PO2, HCO3, BE, O2SAT, TotCO2, " & _
              " pHLow, PCO2Low, PO2Low, HCO3Low, BELow, O2SATLow, TotCO2Low, " & _
              " pHHigh, PCO2High, PO2High, HCO3High, BEHigh, O2SATHigh, TotCO2High, " & _
              " DateTimeAmended, AmendedBy ) VALUES " & _
              "('', '',   '',  '',   '', '',    '', " & _
              " '" & txtpH(0) & "', '" & txtPco2(0) & "', '" & txtPo2(0) & "', '" & txtHco3(0) & "', " & _
              " '" & txtBE(0) & "', '" & txtO2Sat(0) & "', '" & txtTotCo2(0) & "', " & _
              " '" & txtpH(1) & "', '" & txtPco2(1) & "', '" & txtPo2(1) & "', '" & txtHco3(1) & "', " & _
              " '" & txtBE(1) & "', '" & txtO2Sat(1) & "', '" & txtTotCo2(1) & "', " & _
              " '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "', '" & UserName & "')"

1040      Cnxn(0).Execute sql

1050      Unload Me

1060      Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

1070      intEL = Erl
1080      strES = Err.Description
1090      LogError "frmBGARanges", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

1100      On Error GoTo Form_Load_Error

1110      sql = "Select top 1 * from BGADefinitions " & _
              "Order by DateTimeAmended Desc"
1120      Set tb = New Recordset
1130      RecOpenServer 0, tb, sql

1140      If Not tb.EOF Then
1150          txtpH(0) = tb!pHLow & ""
1160          txtPco2(0) = tb!pCO2Low & ""
1170          txtPo2(0) = tb!PO2Low & ""
1180          txtHco3(0) = tb!HCO3Low & ""
1190          txtBE(0) = tb!BELow & ""
1200          txtO2Sat(0) = tb!O2SATLow & ""
1210          txtTotCo2(0) = tb!totCO2Low & ""
1220          txtpH(1) = tb!phhigh & ""
1230          txtPco2(1) = tb!pCO2High & ""
1240          txtPo2(1) = tb!po2high & ""
1250          txtHco3(1) = tb!hco3high & ""
1260          txtBE(1) = tb!BEHigh & ""
1270          txtO2Sat(1) = tb!O2SatHigh & ""
1280          txtTotCo2(1) = tb!TotCO2High & ""
1290          lblAmendedBy = tb!AmendedBy & ""
1300          lblDateAmended = Format$(tb!DateTimeAmended, "dd/mm/yyyy")
1310          lblTimeAmended = Format$(tb!DateTimeAmended, "hh:mm")
1320      End If

1330      Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

1340      intEL = Erl
1350      strES = Err.Description
1360      LogError "frmBGARanges", "Form_Load", intEL, strES, sql

        
End Sub


Private Sub txtBE_KeyPress(Index As Integer, KeyAscii As Integer)

1370      cmdSave.Visible = True

End Sub


Private Sub txtHco3_KeyPress(Index As Integer, KeyAscii As Integer)

1380      cmdSave.Visible = True

End Sub


Private Sub txtO2Sat_KeyPress(Index As Integer, KeyAscii As Integer)

1390      cmdSave.Visible = True

End Sub


Private Sub txtPco2_KeyPress(Index As Integer, KeyAscii As Integer)

1400      cmdSave.Visible = True

End Sub


Private Sub txtpH_KeyPress(Index As Integer, KeyAscii As Integer)

1410      cmdSave.Visible = True

End Sub


Private Sub txtPo2_KeyPress(Index As Integer, KeyAscii As Integer)

1420      cmdSave.Visible = True

End Sub


Private Sub txtTotCo2_KeyPress(Index As Integer, KeyAscii As Integer)

1430      cmdSave.Visible = True

End Sub


