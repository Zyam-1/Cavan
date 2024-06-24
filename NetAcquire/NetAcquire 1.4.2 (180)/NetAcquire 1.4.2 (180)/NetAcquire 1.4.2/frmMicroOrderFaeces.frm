VERSION 5.00
Begin VB.Form frmMicroOrderFaeces 
   Caption         =   "NetAcquire"
   ClientHeight    =   5070
   ClientLeft      =   8655
   ClientTop       =   4080
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   5940
   Begin VB.Frame Frame2 
      Caption         =   "BD Max"
      Height          =   2475
      Left            =   150
      TabIndex        =   17
      Top             =   2460
      Width           =   4215
      Begin VB.CheckBox chkFaecal 
         Caption         =   "BD Max Ent Bac"
         Height          =   315
         Index           =   13
         Left            =   480
         TabIndex        =   19
         Tag             =   "BD MAX Ent Bac"
         Top             =   855
         Width           =   1995
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "BD Max C. Difficile"
         Height          =   315
         Index           =   12
         Left            =   480
         TabIndex        =   18
         Tag             =   "BD MAX Cdiff"
         Top             =   480
         Width           =   1995
      End
      Begin VB.Line Line1 
         X1              =   960
         X2              =   960
         Y1              =   1200
         Y2              =   2160
      End
      Begin VB.Label Label1 
         Caption         =   "Salmonella species"
         Height          =   255
         Index           =   3
         Left            =   1020
         TabIndex        =   23
         Top             =   1965
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Shiga toxin (E coli O157 + others)"
         Height          =   255
         Index           =   2
         Left            =   1020
         TabIndex        =   22
         Top             =   1455
         Width           =   2595
      End
      Begin VB.Label Label1 
         Caption         =   "Campylobacter"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   21
         Top             =   1710
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Shigella species"
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   20
         Top             =   1200
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1100
      Left            =   4560
      Picture         =   "frmMicroOrderFaeces.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2955
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1100
      Left            =   4560
      Picture         =   "frmMicroOrderFaeces.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1740
      Width           =   1200
   End
   Begin VB.Frame frFaeces 
      Caption         =   "Faecal Requests"
      Height          =   1755
      Left            =   150
      TabIndex        =   2
      Top             =   570
      Width           =   4245
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Reducing Substances"
         Height          =   195
         Index           =   11
         Left            =   1890
         TabIndex        =   16
         Tag             =   "RedSub"
         Top             =   810
         Width           =   1875
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "E/P Coli"
         Height          =   195
         Index           =   9
         Left            =   1890
         TabIndex        =   13
         Tag             =   "EPColi"
         Top             =   1230
         Width           =   885
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   930
         TabIndex        =   12
         Tag             =   "CS"
         Top             =   300
         Width           =   705
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "O/P"
         Height          =   195
         Index           =   2
         Left            =   1020
         TabIndex        =   11
         Tag             =   "OP"
         Top             =   570
         Width           =   615
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "Rota/Adeno"
         Height          =   195
         Index           =   6
         Left            =   450
         TabIndex        =   10
         Tag             =   "RotaAdeno"
         Top             =   810
         Width           =   1185
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "Coli 0157"
         Height          =   195
         Index           =   8
         Left            =   660
         TabIndex        =   9
         Tag             =   "Coli0157"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C.Difficile"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   8
         Tag             =   "CDIFF"
         Top             =   300
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "Toxin A"
         Height          =   195
         Index           =   7
         Left            =   780
         TabIndex        =   7
         Tag             =   "ToxinA"
         Top             =   1230
         Width           =   855
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check9"
         Height          =   195
         Index           =   3
         Left            =   1890
         TabIndex        =   6
         Tag             =   "OB0"
         Top             =   570
         Width           =   255
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check10"
         Height          =   195
         Index           =   4
         Left            =   2130
         TabIndex        =   5
         Tag             =   "OB1"
         Top             =   570
         Width           =   225
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Occult Blood"
         Height          =   195
         Index           =   5
         Left            =   2340
         TabIndex        =   4
         Tag             =   "OB2"
         Top             =   570
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "S/S Screen"
         Height          =   195
         Index           =   10
         Left            =   1890
         TabIndex        =   3
         Tag             =   "SSScreen"
         Top             =   1440
         Width           =   1245
      End
   End
   Begin VB.TextBox txtSampleID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1710
      MaxLength       =   12
      TabIndex        =   0
      Top             =   180
      Width           =   1545
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "frmMicroOrderFaeces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFaecalOrders As FaecesRequests
Private Sub chkFaecal_Click(Index As Integer)

41240 cmdSave.Enabled = True

End Sub

Private Sub cmdCancel_Click()

41250 If cmdSave.Enabled Then
41260     If iMsg("Cancel without Saving?", vbQuestion = vbYesNo) = vbNo Then
41270         Exit Sub
41280     End If
41290 End If

41300 Me.Hide

End Sub


Private Sub cmdSave_Click()

41310 SaveDetails

41320 Me.Hide

End Sub


Private Sub Form_Activate()

41330 LoadDetails

End Sub

Private Sub Form_Load()

41340 If sysOptShortFaeces(0) Then
41350     frFaeces.height = 1185
41360 Else
41370     frFaeces.height = 1755
41380 End If

End Sub


Private Sub LoadDetails()

      Dim SampleIDWithOffset As Long
      Dim n As Integer
      Dim Fx As FaecesRequest
      Dim Fxs As New FaecesRequests

41390 On Error GoTo LoadDetails_Error

      '20    SampleIDWithOffset = Val(txtSampleID) + sysOptMicroOffset(0)
41400 SampleIDWithOffset = Val(txtSampleID)
41410 For n = 0 To 13
41420     chkFaecal(n) = 0
41430 Next
      '+++ Junaid 20-05-2024
      '60    Fxs.Load SampleIDWithOffset
41440 Fxs.Load Trim(txtSampleID.Text)
      '--- Junaid
41450 For Each Fx In Fxs
41460     If Fx.Analyser = "B" Then
41470         For n = 12 To 13
41480             If chkFaecal(n).Tag = Fx.Request Then
41490                 chkFaecal(n) = 1
41500                 Exit For
41510             End If
41520         Next
41530     Else
41540         For n = 0 To 11
41550             If chkFaecal(n).Tag = Fx.Request Then
41560                 chkFaecal(n) = 1
41570                 Exit For
41580             End If
41590         Next
41600     End If
41610 Next

41620 cmdSave.Enabled = False

41630 Exit Sub

LoadDetails_Error:

      Dim strES As String
      Dim intEL As Integer

41640 intEL = Erl
41650 strES = Err.Description
41660 LogError "frmMicroOrderFaeces", "LoadDetails", intEL, strES

End Sub

Private Sub SaveDetails()

      Dim n As Integer
      Dim SampleIDWithOffset As Long
      Dim Fx As FaecesRequest
      Dim Fxs As New FaecesRequests

41670 On Error GoTo SaveDetails_Error

      '20    SampleIDWithOffset = Val(txtSampleID) + sysOptMicroOffset(0)
41680 SampleIDWithOffset = Val(txtSampleID)

41690 For n = 0 To 13
41700     If chkFaecal(n) Then
41710         Set Fx = New FaecesRequest
      '+++ Junaid 20-05-2024
      '60            Fx.SampleID = SampleIDWithOffset
41720         Fx.SampleID = Trim(txtSampleID.Text)
      '--- Junaid
41730         Fx.Request = chkFaecal(n).Tag
41740         Fx.UserName = UserName
41750         If n >= 12 And n <= 13 Then
41760             Fx.Analyser = "B"
41770         Else
41780             Fx.Analyser = ""
41790         End If
41800         Fx.Programmed = False
41810         Fxs.Add Fx
41820     End If
41830 Next

41840 If Fxs.Count = 0 Then
41850     iMsg "Nothing to Save!", vbExclamation
41860     cmdSave.Enabled = False
41870     Exit Sub
41880 End If
      '+++ Junaid 20-05-2024
      '230   Fxs.Save SampleIDWithOffset
41890 Fxs.Save Trim(txtSampleID.Text)
      '--- Junaid

41900 Set mFaecalOrders = Fxs

41910 cmdSave.Enabled = False

41920 Exit Sub

SaveDetails_Error:

      Dim strES As String
      Dim intEL As Integer

41930 intEL = Erl
41940 strES = Err.Description
41950 LogError "frmMicroOrderFaeces", "SaveDetails", intEL, strES

End Sub



Private Sub txtsampleid_LostFocus()

41960 LoadDetails

End Sub



Public Property Get FaecalOrders() As FaecesRequests

41970 Set FaecalOrders = mFaecalOrders

End Property

