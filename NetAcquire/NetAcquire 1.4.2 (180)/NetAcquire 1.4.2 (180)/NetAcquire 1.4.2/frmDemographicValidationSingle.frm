VERSION 5.00
Begin VB.Form frmDemographicValidationSingle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Demographic Validation"
   ClientHeight    =   4095
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Height          =   975
      Left            =   4995
      Picture         =   "frmDemographicValidationSingle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "cmdExit"
      Top             =   2850
      Width           =   960
   End
   Begin VB.TextBox txtDummy 
      Height          =   765
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmDemographicValidationSingle.frx":0ECA
      Top             =   1350
      Width           =   1575
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   6210
      TabIndex        =   0
      Top             =   840
      Width           =   1005
   End
   Begin VB.CommandButton cmdNoMatch 
      Height          =   975
      Left            =   3540
      Picture         =   "frmDemographicValidationSingle.frx":0EFC
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2850
      Width           =   960
   End
   Begin VB.CommandButton cmdValidate 
      Height          =   975
      Left            =   1590
      Picture         =   "frmDemographicValidationSingle.frx":1DC6
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2850
      Width           =   960
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Entered Date"
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
      Left            =   8250
      TabIndex        =   30
      Top             =   2820
      Width           =   1185
   End
   Begin VB.Label lblEnteredDate 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9570
      TabIndex        =   29
      Top             =   2790
      Width           =   2940
   End
   Begin VB.Label lblValidatedBy 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9570
      TabIndex        =   28
      Top             =   3120
      Width           =   2940
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Validated by "
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
      Left            =   8250
      TabIndex        =   27
      Top             =   3150
      Width           =   1185
   End
   Begin VB.Label lblSampleID 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   25
      Top             =   330
      Width           =   1440
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "txtInput =>"
      Height          =   195
      Left            =   5190
      TabIndex        =   24
      Top             =   870
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblUnknown 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample ID Unknown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   3030
      TabIndex        =   23
      Top             =   330
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Label lblEnteredBy 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1560
      TabIndex        =   22
      Top             =   2190
      Width           =   2940
   End
   Begin VB.Label lblClinician 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9570
      TabIndex        =   21
      Top             =   2160
      Width           =   2940
   End
   Begin VB.Label lblWard 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9570
      TabIndex        =   20
      Top             =   1860
      Width           =   2940
   End
   Begin VB.Label lblGP 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9570
      TabIndex        =   19
      Top             =   1560
      Width           =   2940
   End
   Begin VB.Label lblSex 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9570
      TabIndex        =   18
      Top             =   1260
      Width           =   2940
   End
   Begin VB.Label lblDoB 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1560
      TabIndex        =   17
      Top             =   1710
      Width           =   2940
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9570
      TabIndex        =   16
      Top             =   660
      Width           =   2940
   End
   Begin VB.Label lblPatName 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1560
      TabIndex        =   15
      Top             =   1260
      Width           =   2940
   End
   Begin VB.Label lblChart 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1560
      TabIndex        =   14
      Top             =   810
      Width           =   2940
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
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
      Left            =   9060
      TabIndex        =   13
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Details Entered by"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2250
      Width           =   1290
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Clinician"
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
      Left            =   8655
      TabIndex        =   11
      Top             =   2190
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Ward"
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
      Left            =   8910
      TabIndex        =   10
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "GP"
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
      Left            =   9120
      TabIndex        =   9
      Top             =   1590
      Width           =   285
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Address"
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
      Left            =   8640
      TabIndex        =   8
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Date of Birth"
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
      Left            =   315
      TabIndex        =   7
      Top             =   1740
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Patient Name"
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
      Left            =   180
      TabIndex        =   6
      Top             =   1290
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
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
      Left            =   930
      TabIndex        =   5
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
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
      Top             =   420
      Width           =   945
   End
   Begin VB.Menu mnuLists 
      Caption         =   "&Lists"
      Begin VB.Menu mnuAccept 
         Caption         =   "&Accept Barcode"
      End
      Begin VB.Menu mnuReject 
         Caption         =   "&Reject Barcode"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel Barcode"
      End
   End
End
Attribute VB_Name = "frmDemographicValidationSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_BarcodeAccept As String
Private m_BarcodeReject As String
Private m_BarcodeCancel As String

Private m_InShowMessage As Boolean
Private m_Valid As Boolean

Private DVs As DemogValidations

Private Activated As Boolean

Private Sub Accept()

          Dim t As Single
          Dim DV As DemogValidation

32860     On Error GoTo Accept_Error

32870     If DVs Is Nothing Then Exit Sub

32880     If lblSampleID = "" Or lblPatName = "" Or lblDoB = "" Or lblEnteredBy = "" Then Exit Sub

32890     If UCase$(Trim$(lblEnteredBy)) <> UCase$(Trim$(UserName)) Then    ' I didn't make the first entry
32900         If Not m_Valid Then  'not valid
32910             Set DV = New DemogValidation
32920             DV.SampleID = lblSampleID
32930             DV.EnteredBy = lblEnteredBy
32940             DV.EnteredDateTime = lblEnteredDate
32950             DV.ValidatedBy = UserName
32960             DVs.Add DV
32970             DVs.Save DV
32980             lblChart.BackColor = vbGreen
32990             lblPatName.BackColor = vbGreen
33000             lblDoB.BackColor = vbGreen
33010             t = Timer
33020             Do While Timer < t + 1: DoEvents: Loop
33030         End If
33040     End If

33050     lblChart.BackColor = &H80000018
33060     lblPatName.BackColor = &H80000018
33070     lblDoB.BackColor = &H80000018
33080     lblEnteredBy.BackColor = &H80000018

33090     lblChart = ""
33100     lblPatName = ""
33110     lblDoB = ""
33120     lblSampleID = ""
33130     lblEnteredBy = ""

33140     Exit Sub

Accept_Error:

          Dim strES As String
          Dim intEL As Integer

33150     intEL = Erl
33160     strES = Err.Description
33170     LogError "frmDemographicValidationSingle", "Accept", intEL, strES

End Sub
Private Sub ClearDetails()

33180     lblChart = ""
33190     lblPatName = ""
33200     lblDoB = ""
33210     lblEnteredBy = ""
33220     lblEnteredDate = ""
33230     lblValidatedBy = ""

33240     lblUnknown.Visible = False

33250     lblEnteredBy.BackColor = &H80000018
33260     lblEnteredBy.ForeColor = &H80000012

33270     m_Valid = False

End Sub

Private Sub FillDetails()

          Dim Ds As New Demographics
          Dim DV As DemogValidation

33280     On Error GoTo FillDetails_Error

33290     ClearDetails

33300     If Val(lblSampleID) = 0 Then Exit Sub

33310     Set DVs = New DemogValidations
33320     DVs.LoadSingle lblSampleID
33330     Set DV = DVs(lblSampleID)
33340     If Not DV Is Nothing Then
33350         m_Valid = True
33360         lblValidatedBy = DV.ValidatedBy
33370     End If

33380     Ds.Load lblSampleID
33390     If Ds.Count = 0 Then
33400         lblUnknown.Visible = True
33410     Else
33420         With Ds(1)
33430             lblChart = .Chart
33440             lblPatName = .PatName
33450             lblDoB = .DoB
33460             lblEnteredBy = .Operator
33470             If UCase$(Trim$(lblEnteredBy)) = UCase$(Trim$(UserName)) Then
33480                 lblEnteredBy.BackColor = vbRed
33490                 lblEnteredBy.ForeColor = vbYellow
33500             End If
33510             lblEnteredDate = .DateTimeDemographics
33520         End With
33530     End If

33540     Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer

33550     intEL = Erl
33560     strES = Err.Description
33570     LogError "frmDemographicValidationSingle", "FillDetails", intEL, strES


End Sub
Private Sub Reject()

          Dim t As Single

33580     lblChart.BackColor = vbRed
33590     lblPatName.BackColor = vbRed
33600     lblDoB.BackColor = vbRed

33610     t = Timer
33620     Do While Timer < t + 1: DoEvents: Loop

33630     lblChart.BackColor = &H80000018
33640     lblPatName.BackColor = &H80000018
33650     lblDoB.BackColor = &H80000018

33660     ClearDetails
33670     lblSampleID = ""

End Sub


Private Sub cmdExit_Click()

33680     Unload Me

End Sub

Private Sub cmdNoMatch_Click()

33690     Reject

End Sub

Private Sub cmdValidate_Click()

33700     Accept

End Sub

Private Sub Form_Activate()
          
33710     On Error GoTo Form_Activate_Error

33720     If Activated Then Exit Sub

33730     Activated = True

33740     m_InShowMessage = True

33750     If m_BarcodeCancel = "" Then
33760         m_BarcodeCancel = iBOX("Scan barcode for 'Cancel'", , m_BarcodeCancel)
33770         If Trim$(m_BarcodeCancel) = "" Then
33780             iMsg "Barcode for 'Cancel' must be specified.", vbCritical
33790             Unload Me
33800             Exit Sub

33810         End If
33820         SaveOptionSetting "AutoValBarcodeCancel", m_BarcodeCancel
33830     End If

33840     If m_BarcodeAccept = "" Then
33850         m_BarcodeAccept = iBOX("Scan barcode for 'Accept'", , m_BarcodeAccept)
33860         If Trim$(m_BarcodeAccept) = "" Then
33870             iMsg "Barcode for 'Accept' must be specified.", vbCritical
33880             Unload Me
33890             Exit Sub
33900         End If
33910         SaveOptionSetting "AutoValBarcodeAccept", m_BarcodeAccept
33920     End If

33930     If m_BarcodeReject = "" Then
33940         m_BarcodeReject = iBOX("Scan barcode for 'Reject'", , m_BarcodeReject)
33950         If Trim$(m_BarcodeReject) = "" Then
33960             iMsg "Barcode for 'Reject' must be specified.", vbCritical
33970             Unload Me
33980             Exit Sub
33990         End If
34000         SaveOptionSetting "AutoValBarcodeReject", m_BarcodeReject
34010     End If

34020     m_InShowMessage = False
34030     txtInput.SetFocus

34040     Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

34050     intEL = Erl
34060     strES = Err.Description
34070     LogError "frmDemographicValidationSingle", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

34080     Me.height = 4755
34090     If Not IsIDE Then
34100         Me.width = 6300
34110     End If

34120     m_BarcodeAccept = GetOptionSetting("AutoValBarcodeAccept", "")
34130     m_BarcodeReject = GetOptionSetting("AutoValBarcodeReject", "")
34140     m_BarcodeCancel = GetOptionSetting("AutoValBarcodeCancel", "")

End Sub

Private Sub Form_Unload(Cancel As Integer)

34150     Activated = False

End Sub

Private Sub mnuAccept_Click()

34160     On Error GoTo mnuAccept_Click_Error

34170     m_InShowMessage = True
34180     m_BarcodeAccept = iBOX("Scan barcode for 'Accept'", , m_BarcodeAccept)
34190     SaveOptionSetting "AutoValBarcodeAccept", m_BarcodeAccept
34200     m_InShowMessage = False

34210     Exit Sub

mnuAccept_Click_Error:

          Dim strES As String
          Dim intEL As Integer

34220     intEL = Erl
34230     strES = Err.Description
34240     LogError "frmDemographicValidationSingle", "mnuAccept_Click", intEL, strES

End Sub


Private Sub mnuCancel_Click()

34250     On Error GoTo mnuCancel_Click_Error

34260     m_InShowMessage = True
34270     m_BarcodeCancel = iBOX("Scan barcode for 'Cancel'", , m_BarcodeCancel)
34280     SaveOptionSetting "AutoValBarcodeCancel", m_BarcodeCancel
34290     m_InShowMessage = False

34300     Exit Sub

mnuCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

34310     intEL = Erl
34320     strES = Err.Description
34330     LogError "frmDemographicValidationSingle", "mnuCancel_Click", intEL, strES

End Sub

Private Sub mnuReject_Click()

34340     On Error GoTo mnuReject_Click_Error

34350     m_InShowMessage = True
34360     m_BarcodeReject = iBOX("Scan barcode for 'Reject'", , m_BarcodeReject)
34370     SaveOptionSetting "AutoValBarcodeReject", m_BarcodeReject
34380     m_InShowMessage = False

34390     Exit Sub

mnuReject_Click_Error:

          Dim strES As String
          Dim intEL As Integer

34400     intEL = Erl
34410     strES = Err.Description
34420     LogError "frmDemographicValidationSingle", "mnuReject_Click", intEL, strES

End Sub




Private Sub txtInput_LostFocus()

34430     If Me.ActiveControl.Tag = "cmdExit" Then
34440         Unload Me
34450         Exit Sub
34460     End If

34470     If Not m_InShowMessage Then
34480         If txtInput = m_BarcodeAccept Then
34490             Accept
34500         ElseIf txtInput = m_BarcodeReject Then
34510             Reject
34520         ElseIf txtInput = m_BarcodeCancel Then
34530             Unload Me
34540             Exit Sub
34550         Else
34560             lblSampleID = txtInput
34570             FillDetails
34580         End If
34590         txtInput = ""
34600         txtInput.SetFocus
34610     End If

End Sub


