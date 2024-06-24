VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmEligibility 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNotEligible 
      Caption         =   "Patient is not Eligible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   6570
      Picture         =   "frmEligibility.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3930
      Width           =   1995
   End
   Begin VB.CommandButton cmdEligible 
      Caption         =   "Patient is Eligible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   6570
      Picture         =   "frmEligibility.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2700
      Width           =   1995
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   16
      Top             =   30
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Vision Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2910
      TabIndex        =   22
      Top             =   5385
      Width           =   1035
   End
   Begin VB.Image imgResultAbnormalFlags 
      Height          =   570
      Left            =   4140
      Picture         =   "frmEligibility.frx":0294
      Stretch         =   -1  'True
      Top             =   5235
      Width           =   570
   End
   Begin VB.Image imgQuestion 
      Height          =   570
      Left            =   9900
      Picture         =   "frmEligibility.frx":0B5E
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   585
   End
   Begin VB.Label lblSampleID 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8765432092345"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3690
      TabIndex        =   21
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label Label10 
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   2910
      TabIndex        =   20
      Top             =   900
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   330
      X2              =   8985
      Y1              =   1515
      Y2              =   1500
   End
   Begin VB.Label lblSampleDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88/88/8888 88:88"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1500
      TabIndex        =   19
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblSampleDateTitle 
      Caption         =   "Sample Date"
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   900
      Width           =   915
   End
   Begin VB.Image imgPrevNotEligible 
      Height          =   570
      Left            =   4140
      Picture         =   "frmEligibility.frx":2B93
      Stretch         =   -1  'True
      Top             =   4620
      Width           =   570
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Previously Eligible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2355
      TabIndex        =   17
      Top             =   4800
      Width           =   1665
   End
   Begin VB.Label lblDoB 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88/88/8888"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1500
      TabIndex        =   15
      Top             =   397
      Width           =   1335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   1080
      TabIndex        =   14
      Top             =   450
      Width           =   315
   End
   Begin VB.Label lblPatName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "fred bloggs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5565
      TabIndex        =   13
      Top             =   390
      Width           =   3450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   5070
      TabIndex        =   12
      Top             =   450
      Width           =   450
   End
   Begin VB.Label lblChart 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12345"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3690
      TabIndex        =   11
      Top             =   420
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   3060
      TabIndex        =   10
      Top             =   450
      Width           =   405
   End
   Begin VB.Label lblNotEligible 
      AutoSize        =   -1  'True
      Caption         =   "Disagree"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5550
      TabIndex        =   9
      Top             =   4260
      Width           =   840
   End
   Begin VB.Label lblEligible 
      AutoSize        =   -1  'True
      Caption         =   "Agree"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5580
      TabIndex        =   8
      Top             =   3060
      Width           =   870
   End
   Begin VB.Image imgAdverse 
      Height          =   570
      Left            =   4140
      Picture         =   "frmEligibility.frx":345D
      Stretch         =   -1  'True
      Top             =   4020
      Width           =   570
   End
   Begin VB.Image imgPreviousAB 
      Height          =   570
      Left            =   4140
      Picture         =   "frmEligibility.frx":3D27
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   570
   End
   Begin VB.Image imgCurrentAB 
      Height          =   570
      Left            =   4140
      Picture         =   "frmEligibility.frx":45F1
      Stretch         =   -1  'True
      Top             =   2820
      Width           =   570
   End
   Begin VB.Image imgGroup 
      Height          =   570
      Left            =   4140
      Picture         =   "frmEligibility.frx":4EBB
      Stretch         =   -1  'True
      Top             =   2220
      Width           =   570
   End
   Begin VB.Image imgPrevious 
      Height          =   570
      Left            =   4140
      Picture         =   "frmEligibility.frx":5785
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   570
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Suggestion: This Patient is NOT Eligible for Electronic Issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   585
      Left            =   5595
      TabIndex        =   5
      Top             =   1680
      Width           =   3420
   End
   Begin VB.Image imgRedFlag 
      Height          =   570
      Left            =   10080
      Picture         =   "frmEligibility.frx":604F
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   570
   End
   Begin VB.Image imgGreenFlag 
      Height          =   570
      Left            =   9930
      Picture         =   "frmEligibility.frx":6199
      Stretch         =   -1  'True
      Top             =   4290
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Current and Previous Group agreement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   4
      Top             =   2385
      Width           =   3480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Previous Sample(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2250
      TabIndex        =   3
      Top             =   1785
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No Previous Adverse Reactions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1140
      TabIndex        =   2
      Top             =   4185
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "All Previous Antibody Screens Negative"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   1
      Top             =   3585
      Width           =   3570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Sample Antibody Screen Negative"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   2985
      Width           =   3780
   End
   Begin VB.Image imgRedCross 
      Height          =   570
      Left            =   9960
      Picture         =   "frmEligibility.frx":62E3
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   570
   End
   Begin VB.Image imgGreenTick 
      Height          =   570
      Left            =   9960
      Picture         =   "frmEligibility.frx":71AD
      Stretch         =   -1  'True
      Top             =   270
      Width           =   570
   End
End
Attribute VB_Name = "frmEligibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Chart As String
Private m_DoB As String
Private m_Name As String
Private m_SampleDate As String
Private m_SampleID As String

Private m_EI As ElectronicIssue

Private Sub SetInfo()

10    On Error GoTo SetInfo_Error

20    If m_EI.PreviousSample = 1 And _
         m_EI.PreviousGroupAgreement = 1 And _
         m_EI.CurrentNegativeAB = 1 And _
         m_EI.PreviousNegativeAB = 1 And _
         m_EI.AdverseReactions = 0 And m_EI.PreviousSampleEligible = 1 And m_EI.ResultAbnormalFlags = 1 Then
30      lblInfo = "Suggestion: This Patient is Eligible for Electronic Issue"
40      lblInfo.ForeColor = &H8000&
50      lblEligible = "Agree"
60      lblEligible.Font.Bold = True
70      lblEligible.Font.Size = 14
80      cmdEligible.Font.Bold = True
90      cmdEligible.Font.Size = 14
100     lblNotEligible = "Disagree"
110     cmdNotEligible.Font.Bold = False
120     cmdNotEligible.Font.Size = 8
130     lblNotEligible.Font.Bold = False
140     lblNotEligible.Font.Size = 8
150   Else
160     lblInfo = "Suggestion: This Patient is NOT Eligible for Electronic Issue"
170     lblInfo.ForeColor = vbRed
180     lblEligible = "Disagree"
190     cmdEligible.Font.Bold = False
200     cmdEligible.Font.Size = 8
210     lblEligible.Font.Bold = False
220     lblEligible.Font.Size = 8
230     lblNotEligible = "Agree"
240     cmdNotEligible.Font.Bold = True
250     cmdNotEligible.Font.Size = 14
260     lblNotEligible.Font.Bold = True
270     lblNotEligible.Font.Size = 14
280   End If

290   Exit Sub

SetInfo_Error:

Dim strES As String
Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "frmEligibility", "SetInfo", intEL, strES

End Sub

Private Sub cmdEligible_Click()

      Dim s As String

10    On Error GoTo cmdEligible_Click_Error

20    If m_EI.PreviousSample <> 1 Or _
         m_EI.PreviousGroupAgreement <> 1 Or _
         m_EI.CurrentNegativeAB <> 1 Or _
         m_EI.PreviousNegativeAB <> 1 Or _
         m_EI.AdverseReactions <> 0 Or m_EI.PreviousSampleEligible <> 1 Or m_EI.ResultAbnormalFlags <> 1 Then
30      s = iBOX("Why do you disagree?")
40      If TimedOut Then
50        Exit Sub
60        Unload Me
70      End If
80      LogReasonWhy "Chart " & lblChart & " : Forced to eligible. " & s, "XM"
90      m_EI.ForcedEligible = 1
100     m_EI.ForcedNotEligible = 0
102   Else
106     s = iBOX("Enter reason?")
108     If TimedOut Then
112         Exit Sub
114         Unload Me
118     End If
120     LogReasonWhy "Chart " & lblChart & " : Forced to eligible. " & s, "XM"
125   End If

128   If m_EI.ForcedNotEligible = 1 Then
130     m_EI.ForcedEligible = 1
140     m_EI.ForcedNotEligible = 0
150   End If

160   Unload Me

170   Exit Sub

cmdEligible_Click_Error:

Dim strES As String
Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmEligibility", "cmdEligible_Click", intEL, strES


End Sub

Private Sub cmdNotEligible_Click()

      Dim s As String

10    On Error GoTo cmdNotEligible_Click_Error

20    If (m_EI.PreviousSample = 1 And _
         m_EI.PreviousGroupAgreement = 1 And _
         m_EI.CurrentNegativeAB = 1 And _
         m_EI.PreviousNegativeAB = 1 And _
         m_EI.AdverseReactions = 0) And m_EI.PreviousSampleEligible = 1 And m_EI.ResultAbnormalFlags = 1 Then
30      s = iBOX("Why do you disagree?")
40      If TimedOut Then
50        Exit Sub
60        Unload Me
70      End If
80      LogReasonWhy "Chart " & lblChart & " : Forced to not eligible. " & s, "XM"
90      m_EI.ForcedEligible = 0
100     m_EI.ForcedNotEligible = 1
102   Else
106     s = iBOX("Enter reason?")
110     If TimedOut Then
112         Exit Sub
114         Unload Me
116     End If
118     LogReasonWhy "Chart " & lblChart & " : Forced to eligible. " & s, "XM"
119   End If

120   If m_EI.ForcedEligible = 1 Then
130     m_EI.ForcedNotEligible = 1
140     m_EI.ForcedEligible = 0
150   End If

160   Unload Me

170   Exit Sub

cmdNotEligible_Click_Error:

Dim strES As String
Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmEligibility", "cmdNotEligible_Click", intEL, strES

End Sub

Private Sub Form_Activate()



10    If IsDate(m_SampleDate) And Trim$(m_Chart) <> "" Then
20      Set m_EI = New ElectronicIssue
30      m_EI.Chart = m_Chart
40      m_EI.SampleDate = m_SampleDate
50      m_EI.SampleID = m_SampleID
60      m_EI.Load
  
70      Select Case m_EI.PreviousSample
          Case 0: imgPrevious.Picture = imgRedCross.Picture
80        Case 1: imgPrevious.Picture = imgGreenTick.Picture
90        Case 2: imgPrevious.Picture = imgQuestion.Picture
100     End Select
  
110     Select Case m_EI.PreviousGroupAgreement
          Case 0: imgGroup.Picture = imgRedCross.Picture
120       Case 1: imgGroup.Picture = imgGreenTick.Picture
130       Case 2: imgGroup.Picture = imgQuestion.Picture
140     End Select
  
150     Select Case m_EI.CurrentNegativeAB
          Case 0: imgCurrentAB.Picture = imgRedCross.Picture
160       Case 1: imgCurrentAB.Picture = imgGreenTick.Picture
170       Case 2: imgCurrentAB.Picture = imgQuestion.Picture
180     End Select
  
190     Select Case m_EI.PreviousNegativeAB
          Case 0: imgPreviousAB.Picture = imgRedCross.Picture
200       Case 1: imgPreviousAB.Picture = imgGreenTick.Picture
210       Case 2: imgPreviousAB.Picture = imgQuestion.Picture
220     End Select
  
230     Select Case m_EI.AdverseReactions
          Case 0: imgAdverse.Picture = imgGreenTick.Picture
240       Case 1: imgAdverse.Picture = imgRedCross.Picture
250       Case 2: imgAdverse.Picture = imgQuestion.Picture
260     End Select
  
270     Select Case m_EI.PreviousSampleEligible
          Case 0: imgPrevNotEligible.Picture = imgRedCross.Picture
280       Case 1: imgPrevNotEligible.Picture = imgGreenTick.Picture
290       Case 2: imgPrevNotEligible.Picture = imgQuestion.Picture
300     End Select

302     Select Case m_EI.ResultAbnormalFlags
           Case 0: imgResultAbnormalFlags.Picture = imgRedCross.Picture
305        Case 1: imgResultAbnormalFlags.Picture = imgGreenTick.Picture
306        Case 2: imgResultAbnormalFlags.Picture = imgQuestion.Picture
308     End Select
        
310   End If

320   SetInfo

End Sub

Public Property Let Chart(ByVal sNewValue As String)

10    lblChart = sNewValue
20    m_Chart = sNewValue

End Property

Public Property Let SampleDate(ByVal sNewValue As String)

10    lblSampleDate = sNewValue
20    m_SampleDate = Format$(sNewValue, "dd/MMM/yyyy HH:nn")

End Property

Public Property Let SampleID(ByVal sNewValue As String)

10    lblSampleID = sNewValue
20    m_SampleID = sNewValue

End Property

Public Property Let PatName(ByVal sNewValue As String)

10    lblPatName = sNewValue
20    m_Name = sNewValue

End Property

Public Property Let DoB(ByVal sNewValue As String)

10    lblDoB = sNewValue
20    m_DoB = sNewValue

End Property


Private Sub Form_Unload(Cancel As Integer)

      Dim sql As String

10    On Error GoTo Form_Unload_Error

20    m_EI.Save

30    Exit Sub

Form_Unload_Error:

Dim strES As String
Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmEligibility", "Form_Unload", intEL, strES, sql

End Sub


