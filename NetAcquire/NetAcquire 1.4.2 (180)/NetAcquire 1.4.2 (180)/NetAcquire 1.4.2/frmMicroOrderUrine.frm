VERSION 5.00
Begin VB.Form frmMicroOrderUrine 
   Caption         =   "NetAcquire"
   ClientHeight    =   3810
   ClientLeft      =   5925
   ClientTop       =   5670
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4215
   Begin VB.Frame Frame2 
      Caption         =   "Urine Sample"
      Height          =   885
      Left            =   600
      TabIndex        =   11
      Top             =   570
      Width           =   3105
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "MSU"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   15
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "CSU"
         Height          =   195
         Index           =   1
         Left            =   810
         TabIndex        =   14
         Top             =   540
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Caption         =   "BSU"
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   13
         Top             =   270
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Caption         =   "SPA"
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   12
         Top             =   540
         Width           =   615
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
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   9
      Top             =   120
      Width           =   1545
   End
   Begin VB.Frame frUrine 
      Caption         =   "Urine Requests"
      Height          =   1155
      Left            =   210
      TabIndex        =   2
      Top             =   1530
      Width           =   3765
      Begin VB.CheckBox chkUrine 
         Alignment       =   1  'Right Justify
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Tag             =   "CS"
         Top             =   270
         Width           =   765
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Pregnancy"
         Height          =   195
         Index           =   1
         Left            =   1950
         TabIndex        =   7
         Tag             =   "Pregnancy"
         Top             =   270
         Width           =   1155
      End
      Begin VB.CheckBox chkUrine 
         Alignment       =   1  'Right Justify
         Caption         =   "Bence Jones"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   6
         Tag             =   "BJ"
         Top             =   870
         Width           =   1245
      End
      Begin VB.CheckBox chkUrine 
         Alignment       =   1  'Right Justify
         Caption         =   "Fat Globules"
         Height          =   195
         Index           =   2
         Left            =   630
         TabIndex        =   5
         Tag             =   "FatGlobules"
         Top             =   630
         Width           =   1215
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Specific Gravity"
         Height          =   195
         Index           =   4
         Left            =   1950
         TabIndex        =   4
         Tag             =   "SG"
         Top             =   630
         Width           =   1455
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Urinary HCG"
         Height          =   195
         Index           =   5
         Left            =   1950
         TabIndex        =   3
         Tag             =   "U-HCG"
         Top             =   870
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   2190
      Picture         =   "frmMicroOrderUrine.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2790
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   735
      Left            =   750
      Picture         =   "frmMicroOrderUrine.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2790
      Width           =   1275
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   780
      TabIndex        =   10
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmMicroOrderUrine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mUrineOrders As UrineRequests

Public Property Get UrineOrders() As UrineRequests

42950 Set UrineOrders = mUrineOrders

End Property

Public Property Get SiteDetails() As String

      Dim n As Integer

42960 For n = 0 To 3
42970   If optU(n) Then
42980     SiteDetails = optU(n).Caption
42990   End If
43000 Next

End Property
Private Sub chkUrine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

43010 cmdSave.Enabled = True

End Sub


Private Sub cmdCancel_Click()

43020 If cmdSave.Enabled Then
43030   If iMsg("Cancel without Saving?", vbQuestion = vbYesNo) = vbNo Then
43040     Exit Sub
43050   End If
43060 End If

43070 Me.Hide

End Sub


Private Sub cmdSave_Click()

43080 SaveDetails
43090 Me.Hide

End Sub


Private Sub Form_Activate()

43100 LoadDetails

End Sub

Private Sub LoadDetails()

      Dim SampleIDWithOffset As Long
      Dim n As Integer
      Dim Ux As UrineRequest
      Dim Uxs As New UrineRequests

43110 On Error GoTo LoadDetails_Error

43120 SampleIDWithOffset = Val(txtSampleID) ' + sysOptMicroOffset(0)
43130 For n = 0 To 5
43140   chkUrine(n) = 0
43150 Next
      '+++ Junaid 20-05-2024
      '60    Uxs.Load SampleIDWithOffset
43160 Uxs.Load Trim(txtSampleID.Text)
      '--- JunaID
43170 For Each Ux In Uxs
43180   For n = 0 To 5
43190     If chkUrine(n).Tag = Ux.Request Then
43200       chkUrine(n) = 1
43210       Exit For
43220     End If
43230   Next
43240 Next
        
43250 cmdSave.Enabled = False

43260 Exit Sub

LoadDetails_Error:

      Dim strES As String
      Dim intEL As Integer

43270 intEL = Erl
43280 strES = Err.Description
43290 LogError "frmMicroOrderUrine", "LoadDetails", intEL, strES

End Sub

Private Sub SaveDetails()

      Dim n As Integer
      Dim SampleIDWithOffset As Long
      Dim Ux As UrineRequest
      Dim Uxs As New UrineRequests

43300 On Error GoTo SaveDetails_Error

43310 SampleIDWithOffset = Val(txtSampleID) ' + sysOptMicroOffset(0)

43320 For n = 0 To 5
43330   If chkUrine(n) Then
43340     Set Ux = New UrineRequest
      '+++ Junaid 20-05-2024
      '60        Ux.SampleID = SampleIDWithOffset
43350     Ux.SampleID = Trim(txtSampleID.Text)
      '--- Junaid
43360     Ux.Request = chkUrine(n).Tag
43370     Ux.UserName = UserName
43380     Uxs.Add Ux
43390   End If
43400 Next

43410 If Uxs.Count = 0 Then
43420   iMsg "Nothing to Save!", vbExclamation
43430   cmdSave.Enabled = False
43440   Exit Sub
43450 End If
      '+++ JUnaid 20-05-2024
      '170   Uxs.Save SampleIDWithOffset
43460 Uxs.Save SampleIDWithOffset
      '--- Junaid

43470 Set mUrineOrders = Uxs

43480 cmdSave.Enabled = False

43490 Exit Sub

SaveDetails_Error:

      Dim strES As String
      Dim intEL As Integer

43500 intEL = Erl
43510 strES = Err.Description
43520 LogError "frmMicroOrderUrine", "SaveDetails", intEL, strES

End Sub

Private Sub Form_Load()

43530 If sysOptShortUrine(0) Then
43540   frUrine.height = 615
43550 Else
43560   frUrine.height = 1155
43570 End If

End Sub


Private Sub optU_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

43580 cmdSave.Enabled = True

End Sub


Private Sub txtsampleid_LostFocus()

43590 LoadDetails

End Sub


