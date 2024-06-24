VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPhoneLog 
   Caption         =   "NetAcquire - Transfusion Phone Log"
   ClientHeight    =   6585
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6810
   Icon            =   "frmPhoneLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optCallOut 
      Caption         =   "Call Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   510
      Left            =   3495
      TabIndex        =   26
      Top             =   60
      Value           =   -1  'True
      Width           =   1845
   End
   Begin VB.OptionButton optCallIn 
      Caption         =   "Call In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   510
      Left            =   1320
      TabIndex        =   25
      Top             =   60
      Width           =   1845
   End
   Begin VB.Frame fraReason 
      Caption         =   "Reason For Call"
      Height          =   1065
      Left            =   330
      TabIndex        =   23
      Top             =   1950
      Width           =   6105
      Begin VB.ComboBox cmbReasonForCall 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmPhoneLog.frx":08CA
         Left            =   210
         List            =   "frmPhoneLog.frx":08CC
         TabIndex        =   24
         Text            =   "cmbReasonForCall"
         Top             =   390
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Phoned To"
      Height          =   945
      Left            =   330
      TabIndex        =   18
      Top             =   3120
      Width           =   6105
      Begin VB.ComboBox cmbWardClinGP 
         Height          =   315
         Left            =   1590
         TabIndex        =   22
         Text            =   "cmbWardClinGP"
         Top             =   390
         Width           =   4215
      End
      Begin VB.OptionButton optClinician 
         Caption         =   "Clinician"
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   450
         Width           =   885
      End
      Begin VB.OptionButton optWard 
         Caption         =   "Ward"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   240
         Width           =   765
      End
      Begin VB.OptionButton optGP 
         Caption         =   "GP"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   660
         Value           =   -1  'True
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Results Phoned"
      Height          =   1155
      Left            =   330
      TabIndex        =   5
      Top             =   660
      Width           =   3525
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Group Rh"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   240
         Width           =   1125
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Antibody Screen / ID"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   9
         Top             =   510
         Width           =   1875
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "DAT"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   780
         Width           =   765
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Kleihauer"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1065
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Genotype"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   6
         Top             =   510
         Width           =   1185
      End
   End
   Begin VB.TextBox txtComment 
      Height          =   1065
      Left            =   330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4260
      Width           =   6105
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Details"
      Enabled         =   0   'False
      Height          =   700
      Left            =   330
      Picture         =   "frmPhoneLog.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5490
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   700
      Left            =   4935
      Picture         =   "frmPhoneLog.frx":0F38
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5490
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   8010
      Top             =   2220
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "View Previous Details"
      Height          =   700
      Left            =   2382
      Picture         =   "frmPhoneLog.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5490
      Visible         =   0   'False
      Width           =   2000
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   330
      TabIndex        =   0
      Top             =   6330
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   420
      TabIndex        =   17
      Top             =   4470
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   3990
      TabIndex        =   16
      Top             =   915
      Width           =   735
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4770
      TabIndex        =   15
      Top             =   870
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Operator"
      Height          =   195
      Left            =   4110
      TabIndex        =   14
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblPhonedBy 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4770
      TabIndex        =   13
      Top             =   1155
      Width           =   1635
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Date/Time"
      Height          =   195
      Left            =   3960
      TabIndex        =   12
      Top             =   1485
      Width           =   765
   End
   Begin VB.Label lblDateTime 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4770
      TabIndex        =   11
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Menu mnuList 
      Caption         =   "&List"
      Begin VB.Menu mnuReasonsForCall 
         Caption         =   "&Reasons For Call"
      End
   End
End
Attribute VB_Name = "frmPhoneLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Private pWardClinGPText As String
Private pWardClinGPType As String

Private Sub FillWardClinGP(ByVal WCG As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillWardClinGP_Error

20    cmbWardClinGP.Clear

30    If WCG = "W" Then
40      sql = "SELECT Text FROM Wards ORDER BY ListOrder"
50    ElseIf WCG = "C" Then
60      sql = "SELECT Text FROM Clinicians ORDER BY ListOrder"
70    Else
80      sql = "SELECT Text FROM GPs ORDER BY ListOrder"
90    End If

100   Set tb = New Recordset
110   RecOpenServer 0, tb, sql
120   Do While Not tb.EOF
130     cmbWardClinGP.AddItem tb!Text & ""
140     tb.MoveNext
150   Loop

160   Exit Sub

FillWardClinGP_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmPhoneLog", "FillWardClinGP", intEL, strES, sql

End Sub
Private Sub FillReasonsForCall()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillReasonsForCall_Error

20    cmbReasonForCall.Clear

30    sql = "SELECT Text FROM lists WHERE ListType = 'RFC' ORDER BY ListOrder"

40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    Do While Not tb.EOF
70      cmbReasonForCall.AddItem tb!Text & ""
80      tb.MoveNext
90    Loop

100   Exit Sub

FillReasonsForCall_Error:

Dim strES As String
Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmPhoneLog", "FillReasonsForCall", intEL, strES, sql

End Sub
Public Property Let WardClinGPType(ByVal strNewValue As String)
      'either "W", "C", "G"
10    pWardClinGPType = strNewValue

End Property


Public Property Let SampleID(ByVal strNewValue As String)

10    pSampleID = strNewValue

End Property


Public Property Let WardClinGPText(ByVal strNewValue As String)

10    pWardClinGPText = strNewValue

End Property



Private Sub chkDiscipline_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10    cmdSave.Enabled = True

End Sub


Private Sub cmbWardclingp_Click()

10    cmdSave.Enabled = True

End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdHistory_Click()

10    frmPhoneLogHistory.SampleID = pSampleID
20    frmPhoneLogHistory.Show 1

End Sub


Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Disc As String

      'Check if at least one discipline is selected
10    On Error GoTo cmdSave_Click_Error

20    Disc = ""
30    For n = 0 To 4
40      If chkDiscipline(n).Value = 1 Then
50        Disc = Disc & Mid$("GADKT", n + 1, 1)
60      Else
70        Disc = Disc & " "
80      End If
90    Next
100   If optCallOut And Trim$(Disc) = "" Then
110     iMsg "Select Group/RH, AB Screen etc.", vbCritical
120     If TimedOut Then Unload Me: Exit Sub
130     Exit Sub
140   ElseIf optCallIn And Trim$(cmbReasonForCall) = "" Then
150     iMsg "Select reason for call", vbCritical
160     If TimedOut Then Unload Me: Exit Sub
170     Exit Sub
180   End If
  
190   If Trim$(cmbWardClinGP) = "" Then
200     iMsg "Fill in 'Phone To'", vbCritical
210     If TimedOut Then Unload Me: Exit Sub
220     Exit Sub
230   End If

240   sql = "Select * from PhoneLog where " & _
            "SampleID = '" & pSampleID & "'"
250   Set tb = New Recordset
260   RecOpenServerBB 0, tb, sql
270   tb.AddNew

280   tb!SampleID = pSampleID
290   tb!PhonedTo = cmbWardClinGP
300   tb!PhonedBy = UserName
310   tb!DateTime = Format$(Now, "dd/mmm/yyyy hh:mm:ss")
320   tb!Comment = Left$(txtComment, 50)
330   If optCallIn Then
340       tb!ReasonForCall = cmbReasonForCall
350   ElseIf optCallOut Then
360       tb!Discipline = Disc
370   End If
380   tb.Update

390   Unload Me

400   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "frmPhoneLog", "cmdSave_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error

20    lblSampleID = pSampleID
30    lblPhonedBy = UserName
40    lblDateTime = Format$(Now, "dd/mmm/yyyy HH:nn:ss")
  
50    Select Case pWardClinGPType
        Case "W":  optWard.Value = True
60      Case "C": optClinician.Value = True
70      Case "G": optGP.Value = True
80    End Select
90    txtComment = ""

100   If CheckPhoneLog(pSampleID).SampleID <> "0" Then
110     cmdHistory.Visible = True
120   Else
130     cmdHistory.Visible = False
140   End If

150   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmPhoneLog", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

      Dim n As Integer

10    For n = 0 To 4
20      chkDiscipline(n).Value = 0
30    Next

40    FillWardClinGP pWardClinGPType
50    cmbWardClinGP = pWardClinGPText

60    FillReasonsForCall

70    SelectCallInOut

End Sub


Private Sub mnuReasonsForCall_Click()

10    flists.ListName = "RFC"
20    flists.oList(6) = True
30    flists.Show 1

40    FillReasonsForCall

End Sub

Private Sub optCallIn_Click()
10    SelectCallInOut
End Sub

Private Sub optCallOut_Click()
10    SelectCallInOut
End Sub

Private Sub optClinician_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillWardClinGP "C"
20    cmbWardClinGP = ""

End Sub


Private Sub optGP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillWardClinGP "G"
20    cmbWardClinGP = ""

End Sub


Private Sub optWard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillWardClinGP "W"
20    cmbWardClinGP = ""

End Sub


Private Sub Timer1_Timer()

10    lblDateTime = Format$(Now, "dd/mmm/yyyy HH:nn:ss")

End Sub


Private Sub txtComment_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub
Private Sub SelectCallInOut()

10    If optCallIn Then
20        Frame1.Enabled = False
30        fraReason.Enabled = True
40        Frame2.Caption = "Phone From"
50    ElseIf optCallOut Then
60        Frame1.Enabled = True
70        fraReason.Enabled = False
80        Frame2.Caption = "Phone To"
90    End If
End Sub

