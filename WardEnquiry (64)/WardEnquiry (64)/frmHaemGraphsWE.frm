VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHaemGraphs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Haematology Cell Distribution"
   ClientHeight    =   5475
   ClientLeft      =   450
   ClientTop       =   705
   ClientWidth     =   6855
   HelpContextID   =   10033
   Icon            =   "frmHaemGraphsWE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6855
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6000
      Top             =   4830
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   4500
      Picture         =   "frmHaemGraphsWE.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1185
   End
   Begin MSChart20Lib.MSChart gRBC 
      Height          =   2385
      Left            =   90
      OleObjectBlob   =   "frmHaemGraphsWE.frx":1534
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   3435
   End
   Begin MSChart20Lib.MSChart gWBC 
      Height          =   2385
      Left            =   3330
      OleObjectBlob   =   "frmHaemGraphsWE.frx":2DF6
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   3435
   End
   Begin MSChart20Lib.MSChart gPla 
      Height          =   2385
      Left            =   90
      OleObjectBlob   =   "frmHaemGraphsWE.frx":4DC6
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3030
      Width           =   3435
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   3570
      TabIndex        =   7
      Top             =   4950
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "White Cell Distribution"
      Height          =   225
      Left            =   3600
      TabIndex        =   5
      Top             =   60
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Platelet Distribution"
      Height          =   225
      Left            =   330
      TabIndex        =   4
      Top             =   2790
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Red Cell Distribution"
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   60
      Width           =   2895
   End
End
Attribute VB_Name = "frmHaemGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String
Private Activated As Boolean

Private Sub LoadGraphs()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

      Dim gDataRBC(1 To 64, 1 To 1) As Variant
      Dim gDataWBC(1 To 64, 1 To 3) As Variant
      Dim gDataPLa(1 To 64, 1 To 1) As Variant
      Dim PltVal As Single

10    On Error GoTo LoadGraphs_Error

20    sql = "Select * from HaemResults where " & _
            "SampleID = '" & mSampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If tb.EOF Then
60      Exit Sub
70    End If

80    For n = 1 To 64
90      gDataRBC(n, 1) = Asc(Mid$(tb!gRBC & String$(64, 1), n, 1))
100     gDataWBC(n, 1) = Asc(Mid$(tb!gwb1 & String$(64, 1), n, 1))
110     gDataWBC(n, 2) = Asc(Mid$(tb!gwb2 & String$(64, 1), n, 1))
120     gDataWBC(n, 3) = Asc(Mid$(tb!gwic & String$(64, 1), n, 1))
130     gDataPLa(n, 1) = Asc(Mid$(tb!gplt & String$(64, 1), n, 1))
140   Next

150   gRBC.ChartData = gDataRBC
160   gWBC.ChartData = gDataWBC
170   gPla.ChartData = gDataPLa
180   PltVal = Val(tb!plt & "")
190   If PltVal < 100 Then
200     gPla.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
210     gPla.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 250
220   Else
230     gPla.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
240   End If

250   Exit Sub

LoadGraphs_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "fHaemGraphs", "LoadGraphs", intEL, strES, sql


End Sub
Public Property Let SampleID(ByVal sNewValue As String)

10    mSampleID = sNewValue

End Property

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub Form_Activate()

10    If LogOffNow Then
20      Unload Me
30    End If

40    PBar.Max = LogOffDelaySecs
50    PBar = 0
60    SingleUserUpdateLoggedOn UserName

70    Timer1.Enabled = True

80    If Activated Then Exit Sub
90    Activated = True

100   LogAsViewed "G", mSampleID, frmMain.txtChart
110   LoadGraphs

End Sub

Private Sub Form_Click()

10    Unload Me

End Sub


Private Sub Form_Load()

10    Activated = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub

Private Sub gPla_Click()

10    Unload Me

End Sub

Private Sub gRBC_Click()

10    Unload Me

End Sub

Private Sub gWBC_Click()

10    Unload Me

End Sub

Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10    PBar = PBar + 1
  
20    If PBar = PBar.Max Then
30      LogOffNow = True
40      Unload Me
50    End If

End Sub


