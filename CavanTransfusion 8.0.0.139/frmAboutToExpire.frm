VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAboutToExpire 
   Caption         =   "NetAcquire"
   ClientHeight    =   4875
   ClientLeft      =   1185
   ClientTop       =   2295
   ClientWidth     =   8040
   Icon            =   "frmAboutToExpire.frx":0000
   ScaleHeight     =   4875
   ScaleWidth      =   8040
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   765
      Left            =   6105
      Picture         =   "frmAboutToExpire.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1950
      Width           =   1395
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   435
      Left            =   5760
      TabIndex        =   1
      Top             =   840
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   767
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   7
      SelStart        =   1
      Value           =   1
      TextPosition    =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3795
      Left            =   270
      TabIndex        =   0
      Top             =   630
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Unit Number          |<Group Rh   |<Date Expiry            "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   270
      TabIndex        =   6
      Top             =   4620
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "These Units are about to Expire "
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
      Left            =   480
      TabIndex        =   4
      Top             =   180
      Width           =   3990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "1    2    3    4    5    6    7"
      Height          =   195
      Left            =   5910
      TabIndex        =   3
      Top             =   1290
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number Of Days Before Expiry"
      Height          =   195
      Left            =   5685
      TabIndex        =   2
      Top             =   630
      Width           =   2145
   End
End
Attribute VB_Name = "frmAboutToExpire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ReCalc()

Dim Expiry As String
Dim s As String
Dim p As Product
Dim Ps As New Products

10    On Error GoTo ReCalc_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    Expiry = Format$(DateAdd("d", Slider1.Value, Now), "dd/mmm/yyyy")

60    Ps.LoadLatestBetweenExpiryDates Now, Expiry
70    For Each p In Ps
80      If InStr("CXRP", p.PackEvent) > 0 Then
  
90        s = p.ISBT128 & vbTab & _
              Bar2Group(p.GroupRh & "") & vbTab & _
              Format(p.DateExpiry, "dd/mm/yyyy HH:mm")
100       g.AddItem s
110     End If
120   Next

130   If g.Rows > 2 Then
140     g.RemoveItem 1
150   End If

160   Exit Sub

ReCalc_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmAboutToExpire", "ReCalc", intEL, strES

End Sub

Private Sub cmdOK_Click()

10    Unload Me

End Sub

Private Sub Form_Activate()

10    ReCalc

End Sub

Private Sub Form_Load()

10    Slider1.Value = Val(sysOptTransfusionExpiry(0))

End Sub


Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo Slider1_MouseUp_Error

20    For n = 1 To 4
30      sql = "Select * from Options where " & _
              "Description = 'TRANSFUSIONEXPIRY'"
40      Set tb = New Recordset
50      RecOpenServer 0, tb, sql
60      If tb.EOF Then
70        tb.AddNew
80        tb!Description = "TRANSFUSIONEXPIRY"
90      End If
100     tb!Contents = Format$(Val(Slider1.Value))
110     tb.Update
120   Next
130   sysOptTransfusionExpiry(0) = Format$(Val(Slider1.Value))

140   ReCalc

150   Exit Sub

Slider1_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmAboutToExpire", "Slider1_MouseUp", intEL, strES, sql


End Sub


