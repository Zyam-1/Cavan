VERSION 5.00
Begin VB.Form frmPhoneLog 
   Caption         =   "NetAcquire - Phone Log"
   ClientHeight    =   5805
   ClientLeft      =   285
   ClientTop       =   885
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   7425
   Begin VB.ComboBox cmbComment 
      Height          =   315
      Left            =   1080
      TabIndex        =   28
      Text            =   "cmbComment"
      Top             =   4110
      Width           =   6045
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   5490
      TabIndex        =   25
      Top             =   2490
      Width           =   1635
      Begin VB.OptionButton optCallOut 
         Caption         =   "Call OUT"
         Height          =   195
         Left            =   420
         TabIndex        =   27
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optCallIn 
         Caption         =   "Call IN"
         Height          =   195
         Left            =   420
         TabIndex        =   26
         Top             =   210
         Width           =   765
      End
   End
   Begin VB.ComboBox cmbWard 
      Height          =   315
      Left            =   1080
      TabIndex        =   22
      Text            =   "cmbWard"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.OptionButton optWard 
      Caption         =   "Wards"
      Height          =   195
      Left            =   2640
      TabIndex        =   21
      Top             =   3510
      Width           =   765
   End
   Begin VB.OptionButton optGP 
      Alignment       =   1  'Right Justify
      Caption         =   "GP's"
      Height          =   195
      Left            =   1980
      TabIndex        =   20
      Top             =   3510
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "View Previous Details"
      Height          =   675
      Left            =   3210
      Picture         =   "frmPhoneLog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4650
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   6480
      Top             =   1500
   End
   Begin VB.ComboBox cmbGP 
      Height          =   315
      Left            =   1080
      TabIndex        =   16
      Text            =   "cmbGP"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   5640
      Picture         =   "frmPhoneLog.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4650
      Width           =   1485
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Details"
      Enabled         =   0   'False
      Height          =   675
      Left            =   1110
      Picture         =   "frmPhoneLog.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4650
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   3135
      Left            =   1080
      TabIndex        =   0
      Top             =   150
      Width           =   3015
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Results Not Phoned"
         Height          =   255
         Index           =   8
         Left            =   210
         TabIndex        =   24
         Top             =   2730
         Width           =   1725
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Endocrinology Results Phoned"
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   210
         TabIndex        =   23
         Top             =   2130
         Width           =   2565
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Microbiology Results Phoned"
         Height          =   255
         Index           =   6
         Left            =   210
         TabIndex        =   19
         Top             =   1860
         Width           =   2415
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "External Results Phoned"
         Height          =   255
         Index           =   5
         Left            =   210
         TabIndex        =   6
         Top             =   1590
         Width           =   2085
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Blood Gas Results Phoned"
         Height          =   255
         Index           =   4
         Left            =   210
         TabIndex        =   5
         Top             =   1320
         Width           =   2265
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Immunology Results Phoned"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   4
         Top             =   1050
         Width           =   2355
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Coagulation Results Phoned"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   780
         Width           =   2325
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Biochemistry Results Phoned"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   510
         Width           =   2385
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Haematology Results Phoned"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Width           =   2505
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   2760
         Y1              =   2550
         Y2              =   2550
      End
   End
   Begin VB.Label lblPhoneTo 
      AutoSize        =   -1  'True
      Caption         =   "Phone To"
      Height          =   195
      Left            =   330
      TabIndex        =   17
      Top             =   3750
      Width           =   705
   End
   Begin VB.Label lblDateTime 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5490
      TabIndex        =   15
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Date/Time"
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   1620
      Width           =   765
   End
   Begin VB.Label lblPhonedBy 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5490
      TabIndex        =   13
      Top             =   1050
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Phoned By"
      Height          =   195
      Left            =   4665
      TabIndex        =   12
      Top             =   1095
      Width           =   780
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5490
      TabIndex        =   11
      Top             =   540
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   4710
      TabIndex        =   10
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   375
      TabIndex        =   7
      Top             =   4170
      Width           =   660
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCommentList 
         Caption         =   "&Comment List"
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
Private pGP As String

Private pWardOrGp As String

Private m_Discipline As String
Private Sub AddPhoneComment(ByVal Discipline As String)

      Dim sql As String
      Dim s As String
      Dim OB As Observation
      Dim OBs As Observations

34870 On Error GoTo AddPhoneComment_Error

34880 If Discipline <> "Demographic" Then
34890     If chkDiscipline(8).Value = 0 Then
34900         If Discipline = "MicroGeneral" Then Discipline = "Demographic"
34910         s = "Results Phoned to " & _
                  IIf(optGP, cmbGP, cmbWard) & _
                  " at " & Format$(Now, "hh:mm") & _
                  " on " & Format$(Now, "dd/MM/yyyy") & _
                  " by " & UserName

34920         Set OBs = New Observations
34930         Set OBs = OBs.Load(pSampleID, Discipline)
34940         If OBs Is Nothing Then
34950             Set OBs = New Observations
34960             OBs.Save pSampleID, True, Discipline, s
34970         Else
34980             Set OB = OBs.Item(1)
34990             If InStr(OB.Comment, "Results Phoned") = 0 Then
35000                 Set OBs = New Observations
35010                 OBs.Save pSampleID, True, Discipline, OB.Comment & " " & s
35020             End If
35030         End If
35040     Else
35050         If Discipline = "MicroGeneral" Then Discipline = "Demographic"
35060         s = "Results Not Phoned. " & cmbComment & " " & _
                  "Logged at " & Format$(Now, "hh:mm") & _
                  " on " & Format$(Now, "dd/MM/yyyy") & _
                  " by " & UserName
35070         Set OBs = New Observations
35080         Set OBs = OBs.Load(pSampleID, Discipline)
35090         If OBs Is Nothing Then
35100             Set OBs = New Observations
35110             OBs.Save pSampleID, True, Discipline, s
35120         Else
35130             Set OB = OBs.Item(1)
35140             If InStr(OB.Comment, "Results not Phoned") = 0 Then
35150                 Set OBs = New Observations
35160                 OBs.Save pSampleID, True, Discipline, OB.Comment & " " & s
35170             End If
35180         End If
35190     End If
35200 End If

35210 Exit Sub

AddPhoneComment_Error:

      Dim strES As String
      Dim intEL As Integer

35220 intEL = Erl
35230 strES = Err.Description
35240 LogError "frmPhoneLog", "AddPhoneComment", intEL, strES, sql

End Sub

Private Sub FillComments()

      Dim Lxs As New Lists
      Dim Lx As List

35250 Lxs.Load "PhoneComment"

35260 cmbComment.Clear
35270 For Each Lx In Lxs
35280   cmbComment.AddItem Lx.Text
35290 Next
        
35300 cmbComment.ListIndex = -1
        
End Sub

Private Sub chkDiscipline_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

35310 If Index = 8 And chkDiscipline(8).Value = 1 Then
35320   cmbWard.Visible = False
35330   cmbGP.Visible = False
35340   cmbWard = ""
35350   cmbGP = ""
35360   optWard.Visible = False
35370   optGP.Visible = False
35380   lblPhoneTo.Visible = False
35390 Else
35400   optWard.Visible = True
35410   optGP.Visible = True
35420   lblPhoneTo.Visible = True
35430   chkDiscipline(8).Value = 0
35440   If optGP.Value = True Then
35450     cmbGP.Visible = True
35460   Else
35470     cmbWard.Visible = True
35480   End If
35490 End If

35500 cmdSave.Enabled = True

End Sub


Private Sub cmbComment_KeyPress(KeyAscii As Integer)

35510 cmdSave.Enabled = True

End Sub


Private Sub cmbGP_Click()

35520 cmdSave.Enabled = True

End Sub


Private Sub cmbWard_Click()

35530 cmdSave.Enabled = True

End Sub


Private Sub cmdCancel_Click()

35540 Unload Me
       
End Sub


Private Sub cmdHistory_Click()

35550 frmPhoneLogHistory.SampleID = pSampleID
35560 frmPhoneLogHistory.Show 1

End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim n As Integer
      Dim Disc As String
      Dim Discipline As String
      Dim PhonedTo As String
      Dim isDiscSelect As Boolean

      'Check if at least one discipline is selected
35570 On Error GoTo cmdSave_Click_Error
35580 isDiscSelect = False
35590 For n = 0 To 7
35600     If chkDiscipline(n).Value = 1 Then
35610         isDiscSelect = True
35620         Exit For
35630     End If
35640 Next


35650 If isDiscSelect = False Then
35660     iMsg "Please Select a discipline", vbInformation
35670     Exit Sub
35680 End If



35690 Disc = ""
35700 For n = 0 To 8
35710     If chkDiscipline(n).Value = 1 Then
35720         Disc = Disc & Mid$("HBCIGEMDN", n + 1, 1)
35730         Select Case n
              Case 0: AddPhoneComment "Haematology"
35740         Case 1: AddPhoneComment "Biochemistry"
35750         Case 2: AddPhoneComment "Coagulation"
35760         Case 3: AddPhoneComment "Immunology"
35770         Case 4: AddPhoneComment "BloodGas"
35780         Case 5:    'addphonecomment "external"
35790         Case 6: AddPhoneComment "MicroGeneral"
35800         Case 7: AddPhoneComment "Endocrinology"
35810         Case 8: AddPhoneComment "Demographic"
35820         End Select
35830     Else
35840         Disc = Disc & " "
35850     End If
35860 Next
35870 If Trim$(Disc) = "" Then
35880     iMsg "Select Discipline.", vbCritical
35890     Exit Sub
35900 End If

35910 If Trim$(cmbGP) = "" And optGP And InStr(Disc, "N") = 0 Then
35920     iMsg "Fill in 'Phone To'", vbCritical
35930     Exit Sub
35940 ElseIf Trim$(cmbWard) = "" And optWard And InStr(Disc, "N") = 0 Then
35950     iMsg "Fill in 'Phone To'", vbCritical
35960     Exit Sub
35970 End If

35980 If optGP Then
35990     PhonedTo = cmbGP
36000 Else
36010     PhonedTo = cmbWard
36020 End If

36030 sql = "INSERT INTO PhoneLog " & _
            "([SampleID], [DateTime], [PhonedTo], [PhonedBy], [Comment], [Discipline], [Direction], [Year]) " & _
            "VALUES " & _
            "('" & pSampleID & "', " & _
            " '" & Format$(Now, "dd/mmm/yyyy HH:mm") & "', " & _
            " '" & AddTicks(PhonedTo) & "', " & _
            " '" & AddTicks(UserName) & "', " & _
            " '" & Left$(AddTicks(cmbComment), 50) & "', " & _
            " '" & Disc & "', " & _
            " '" & IIf(optCallIn.Value = True, "IN", "OUT") & "', " & _
            " '" & Format$(Now, "yyyy") & "')"
36040 Cnxn(0).Execute sql

      '310   sql = "Select * from PhoneLog where " & _
       '            "SampleID = '" & pSampleID & "'"
      '320   Set tb = New Recordset
      '330   RecOpenServer 0, tb, sql
      '340   tb.AddNew
      '
      '350   tb!SampleID = pSampleID
      '360   If optGP Then
      '370     tb!PhonedTo = cmbGP
      '380   Else
      '390     tb!PhonedTo = cmbWard
      '400   End If
      '410   tb!PhonedBy = UserName
      '420   tb!DateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
      '430   tb!Comment = Left$(cmbComment, 50)
      '440   tb!Discipline = Disc
      '450   tb!Direction = IIf(optCallIn.Value = True, "IN", "OUT")
      '460   tb.Update

36050 For n = 0 To 8
36060     Discipline = ""
36070     If chkDiscipline(n).Value = 1 Then
36080         Select Case n
              Case 0: Discipline = "Haematology"
36090         Case 1: Discipline = "Biochemistry"
36100         Case 2: Discipline = "Coagulation"
36110         Case 3: Discipline = "Immunology"
36120         Case 4: Discipline = "BloodGas"
36130         Case 5:    'addphonecomment "external"
36140         Case 6: Discipline = "MicroGeneral"
36150         Case 7: Discipline = "Endocrinology"
36160         Case 8: Discipline = "Demographic"
36170         End Select
36180     End If
36190     If Discipline <> "" Then
36200         sql = "DELETE FROM PhoneAlert " & _
                    "WHERE SampleID = '" & pSampleID & "' " & _
                    "AND Discipline = '" & Discipline & "'"
36210         Cnxn(0).Execute sql
36220     End If
36230 Next

36240 Unload Me

36250 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

36260 intEL = Erl
36270 strES = Err.Description
36280 LogError "frmPhoneLog", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub Form_Activate()

36290 lblSampleID = pSampleID
36300 lblPhonedBy = UserName
36310 lblDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
36320 If pWardOrGp = "GP" Then
36330   cmbGP = pGP
36340   cmbGP.Visible = True
36350   optGP.Value = True
36360   cmbWard.Visible = False
36370 Else
36380   cmbWard = pGP
36390   cmbGP.Visible = False
36400   optWard.Value = True
36410   cmbWard.Visible = True
36420 End If
36430 cmbComment = ""

36440 If CheckPhoneLog(pSampleID).SampleID <> 0 Then
36450   cmdHistory.Visible = True
36460 Else
36470   cmdHistory.Visible = False
36480 End If

End Sub

Private Sub Form_Load()

      Dim n As Integer

36490 If pWardOrGp = "GP" Then
36500   FillGPs cmbGP, HospName(0)
36510   cmbWard.Visible = False
36520   cmbGP.Visible = True
36530 Else
36540   FillWards cmbWard, HospName(0)
36550   cmbWard = ""
36560   cmbWard.Visible = True
36570   cmbGP.Visible = False
36580 End If

36590 For n = 0 To 7
36600   chkDiscipline(n).Value = 0
36610 Next

36620 FillComments

End Sub



Public Property Let SampleID(ByVal strNewValue As String)

36630 pSampleID = strNewValue

End Property

Public Property Let Discipline(ByVal strNewValue As String)

36640 m_Discipline = strNewValue
36650 If m_Discipline = "Micro" Then
36660   chkDiscipline(6).Value = 1
36670 End If

End Property
Public Property Let GP(ByVal strNewValue As String)

36680 pGP = strNewValue

End Property

Public Property Let WardOrGP(ByVal strNewValue As String)

36690 pWardOrGp = strNewValue

End Property


Private Sub mnuCommentList_Click()

      Dim CSave As String

36700 CSave = cmbComment

36710 With frmListsGeneric
36720   .ListType = "PhoneComment"
36730   .ListTypeName = "Phone Comment"
36740   .ListTypeNames = "Phone Comments"
36750   .Show 1
36760 End With

36770 FillComments
36780 cmbComment = CSave

End Sub

Private Sub optGP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

36790 FillGPs cmbGP, HospName(0)
36800 cmbWard.Visible = False
36810 cmbGP.Visible = True

End Sub


Private Sub optWard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

36820 FillWards cmbWard, HospName(0)
36830 cmbWard = ""
36840 cmbWard.Visible = True
36850 cmbGP.Visible = False

End Sub


Private Sub Timer1_Timer()

36860 lblDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")

End Sub

