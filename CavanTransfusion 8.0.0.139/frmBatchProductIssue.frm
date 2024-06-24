VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchProductIssue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Batch Issue"
   ClientHeight    =   5580
   ClientLeft      =   270
   ClientTop       =   645
   ClientWidth     =   11775
   ControlBox      =   0   'False
   Icon            =   "frmBatchProductIssue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmeProducts 
      Caption         =   "Product Names"
      Height          =   2565
      Left            =   3825
      TabIndex        =   50
      Top             =   1290
      Visible         =   0   'False
      Width           =   4125
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   735
         Left            =   3240
         TabIndex        =   53
         Top             =   1710
         Width           =   765
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   735
         Left            =   2430
         TabIndex        =   52
         Top             =   1710
         Width           =   765
      End
      Begin MSFlexGridLib.MSFlexGrid flxProducts 
         Height          =   2235
         Left            =   150
         TabIndex        =   51
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   5
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4755
      Left            =   420
      TabIndex        =   12
      Top             =   330
      Width           =   5085
      Begin VB.ComboBox cmbPatientGroup 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "cmbPatientGroup"
         Top             =   630
         Width           =   1215
      End
      Begin VB.TextBox txtAddr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox txtAddr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         Top             =   2340
         Width           =   3615
      End
      Begin VB.TextBox txtAddr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         Top             =   2040
         Width           =   3615
      End
      Begin VB.ComboBox cmbWard 
         BackColor       =   &H80000018&
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3300
         Width           =   3615
      End
      Begin VB.ComboBox cmbClinician 
         BackColor       =   &H80000018&
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3630
         Width           =   3615
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   21
         Top             =   1350
         Width           =   3615
      End
      Begin VB.TextBox txtChart 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   20
         Top             =   660
         Width           =   1590
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2820
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   19
         Top             =   1680
         Width           =   705
      End
      Begin VB.TextBox txtDoB 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1680
         Width           =   1365
      End
      Begin VB.TextBox txtAandE 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   990
         Width           =   1590
      End
      Begin VB.TextBox txtReceivedDateTime 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   4380
         Width           =   2115
      End
      Begin VB.TextBox txtSampleDateTime 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4020
         Width           =   2115
      End
      Begin VB.TextBox txtSampleID 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1590
      End
      Begin VB.TextBox txtTypenex 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Group"
         Height          =   195
         Left            =   2940
         TabIndex        =   44
         Top             =   660
         Width           =   435
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   195
         Index           =   23
         Left            =   840
         TabIndex        =   43
         Top             =   2700
         Width           =   90
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Index           =   24
         Left            =   840
         TabIndex        =   42
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Addr 1"
         Height          =   195
         Index           =   28
         Left            =   480
         TabIndex        =   41
         Top             =   2100
         Width           =   465
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
         Height          =   195
         Index           =   31
         Left            =   360
         TabIndex        =   40
         Top             =   3660
         Width           =   585
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Index           =   33
         Left            =   540
         TabIndex        =   39
         Top             =   3360
         Width           =   390
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   34
         Left            =   480
         TabIndex        =   38
         Top             =   1410
         Width           =   420
      End
      Begin VB.Label lblsex 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3930
         TabIndex        =   37
         Top             =   1680
         Width           =   705
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   14
         Left            =   3630
         TabIndex        =   36
         Top             =   1710
         Width           =   270
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   13
         Left            =   570
         TabIndex        =   35
         Top             =   720
         Width           =   375
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   16
         Left            =   2460
         TabIndex        =   34
         Top             =   1710
         Width           =   285
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         Height          =   195
         Index           =   15
         Left            =   510
         TabIndex        =   33
         Top             =   1740
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "A/E"
         Height          =   195
         Left            =   660
         TabIndex        =   32
         Top             =   1020
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sample Date/Time"
         Height          =   195
         Left            =   840
         TabIndex        =   31
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Received Date/Time"
         Height          =   195
         Left            =   660
         TabIndex        =   30
         Top             =   4410
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lab Number"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   300
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Typenex"
         Height          =   195
         Left            =   2730
         TabIndex        =   28
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdPrintBoth 
      Caption         =   "Print &Both"
      Height          =   1245
      Left            =   5850
      Picture         =   "frmBatchProductIssue.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3780
      Width           =   1005
   End
   Begin VB.CommandButton cmdPrintLabel 
      Caption         =   "Print &Label"
      Height          =   1245
      Left            =   7080
      Picture         =   "frmBatchProductIssue.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3780
      Width           =   1005
   End
   Begin VB.CommandButton cmdPrintForm 
      Caption         =   "Print &Form"
      Height          =   1245
      Left            =   8370
      Picture         =   "frmBatchProductIssue.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3780
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Issue Product"
      Height          =   2835
      Left            =   5850
      TabIndex        =   1
      Top             =   330
      Width           =   5385
      Begin VB.TextBox txtIdentifier 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   570
         Width           =   2415
      End
      Begin VB.CommandButton cmdIssue 
         Caption         =   "Issue to Patient"
         Enabled         =   0   'False
         Height          =   1515
         Left            =   4110
         Picture         =   "frmBatchProductIssue.frx":1DA8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   570
         Width           =   1005
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   49
         Top             =   2220
         Width           =   3675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   780
         TabIndex        =   48
         Top             =   2250
         Width           =   450
      End
      Begin VB.Label lblUnitGroup 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblBatchNumber 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   46
         Top             =   1380
         Width           =   2415
      End
      Begin VB.Label lblProduct 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   45
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "BatchNumber"
         Height          =   195
         Left            =   255
         TabIndex        =   11
         Top             =   1410
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         Height          =   195
         Left            =   675
         TabIndex        =   10
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Identifier"
         Height          =   195
         Left            =   630
         TabIndex        =   9
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Group"
         Height          =   195
         Left            =   795
         TabIndex        =   3
         Top             =   1830
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   1245
      Left            =   10230
      Picture         =   "frmBatchProductIssue.frx":2C72
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3780
      Width           =   1005
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   420
      TabIndex        =   7
      Top             =   5160
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmBatchProductIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String
Private mTypenex As String

Private TimeIssued As String
Private boolFormLoaded As Boolean
Private m_Flag As Boolean
Private m_UID As String

Private Const fcsLine_NO = 0
Private Const fcsRID = 1
Private Const fcsSID = 2
Private Const fcsUID = 3
Private Const fcsPrd = 4


Private Sub FormatGrid()
    On Error GoTo ERROR_FormatGrid
    
    flxProducts.Rows = 1
    flxProducts.row = 0
    
    flxProducts.ColWidth(fcsLine_NO) = 250
    
    flxProducts.TextMatrix(0, fcsRID) = ""
    flxProducts.ColWidth(fcsRID) = 0
    flxProducts.ColAlignment(fcsRID) = flexAlignLeftCenter
    
    flxProducts.TextMatrix(0, fcsSID) = ""
    flxProducts.ColWidth(fcsSID) = 0
    flxProducts.ColAlignment(fcsSID) = flexAlignLeftCenter
    
    flxProducts.TextMatrix(0, fcsUID) = ""
    flxProducts.ColWidth(fcsUID) = 0
    flxProducts.ColAlignment(fcsUID) = flexAlignLeftCenter
    
    flxProducts.TextMatrix(0, fcsPrd) = "Products"
    flxProducts.ColWidth(fcsPrd) = 1550
    flxProducts.ColAlignment(fcsPrd) = flexAlignLeftCenter
    
        
    Exit Sub
ERROR_FormatGrid:
    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmMain", "FormatGrid", intEL, strES
End Sub

Private Sub ClearDetails()

10    lblProduct = ""
20    lblBatchNumber = ""
30    lblUnitGroup = ""
40    lblStatus = ""
50    lblStatus.BackColor = &HC0FFFF

End Sub

Public Property Let SampleID(ByVal sNewValue As String)

10    mSampleID = sNewValue

End Property


Public Property Let Typenex(ByVal sNewValue As String)

10    mTypenex = sNewValue

End Property
Private Sub FillLists()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillLists_Error

20    cmbWard.Clear
30    cmbClinician.Clear

40    sql = "Select * from Wards order by ListOrder"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    Do While Not tb.EOF
80      cmbWard.AddItem tb!Text & ""
90      tb.MoveNext
100   Loop

110   sql = "Select * from Clinicians order by listorder"
120   Set tb = New Recordset
130   RecOpenServer 0, tb, sql
140   Do While Not tb.EOF
150     cmbClinician.AddItem tb!Text & ""
160     tb.MoveNext
170   Loop

180   Exit Sub

FillLists_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmBatchProductIssue", "FillLists", intEL, strES, sql

End Sub

Private Sub btnCancel_Click()
    fmeProducts.Visible = False
    m_Flag = False
End Sub

Private Sub btnOK_Click()
    m_Flag = True
    m_UID = flxProducts.TextMatrix(flxProducts.row, fcsUID)
    Call cmdIssue_Click
    fmeProducts.Visible = False
End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdIssue_Click()

      Dim DoB As String
      Dim s As String
      Dim BP As BatchProduct
      Dim BPs As New BatchProducts

10    On Error GoTo cmdIssue_Click_Error
      
      If m_Flag = False Then
        If txtSampleID.Text <> "" Then
'            Call GetProducts(Trim(txtSampleID.Text))
'            fmeProducts.Visible = True
            Exit Sub
        End If
      End If
      
20    If Left$(cmbPatientGroup & "  ", 2) <> Left$(lblUnitGroup & "  ", 2) And cmbPatientGroup <> "" And lblUnitGroup <> "" Then
30      s = "Patient and Unit Group Differ." & vbCrLf & _
            "Continue anyway?"
40      Answer = iMsg(s, vbYesNo + vbQuestion)
50      If TimedOut Then Unload Me: Exit Sub
60      If Answer = vbNo Then
70        Exit Sub
80      End If
90      LogReasonWhy "Batch Product(" & txtSampleID & "): Patient/Unit Group Mis-match. Proceeded", "XM"
100   End If

110   If Trim$(txtSampleID) = "" Then
120     iMsg "Specify Lab Number!", vbCritical
130     If TimedOut Then Unload Me: Exit Sub
140     Exit Sub
150   End If

160   If Trim$(txtName) = "" Then
170     iMsg "Patients Name?", vbQuestion
180     If TimedOut Then Unload Me: Exit Sub
190     Exit Sub
200   End If

210   s = "Issue " & txtName & " " & _
          "with Unit of " & _
          lblProduct & " " & _
          "Batch " & lblBatchNumber & "?"
220   Answer = iMsg(s, vbQuestion + vbYesNo)
230   If TimedOut Then Unload Me: Exit Sub
240   If Answer = vbNo Then
250     Exit Sub
260   End If

270   BPs.LoadSpecificIdentifierLatest txtIdentifier

280   Set BP = BPs.Item(1)
      
290   If IsDate(txtDoB) Then
300     DoB = Format(txtDoB, "dd/MMM/yyyy")
310   Else
320     DoB = vbNullString
330   End If

340   BP.Chart = txtChart
350   BP.PatName = txtName
360   BP.DoB = DoB
370   BP.Age = txtAge
380   BP.Sex = Left$(lblsex, 1)
390   BP.Addr0 = txtAddr(0)
400   BP.Addr1 = txtAddr(1)
410   BP.Addr2 = txtAddr(2)
420   BP.Ward = cmbWard
430   BP.Clinician = cmbClinician
440   BP.UserName = UserName
450   BP.RecordDateTime = Format(Now, "dd/MMM/yyyy HH:nn:ss")
460   BP.SampleID = txtSampleID
470   BP.Typenex = txtTypenex
480   BP.PatientGroup = cmbPatientGroup
490   BP.EventCode = "I"
500   BPs.Update BP

      If m_Flag Then
        If m_UID <> "" Then
            'Call UpdateIdentifier(m_UID, txtIdentifier.Text)
            'Call CountAndUpdate(Trim(txtSampleID.Text))
            m_Flag = False
            m_UID = ""
        End If
      End If

510   lblStatus = ""
520   lblUnitGroup = ""
530   lblBatchNumber = ""
540   lblProduct = ""
550   txtIdentifier = ""

560   Exit Sub

cmdIssue_Click_Error:

      Dim strES As String
      Dim intEL As Integer

570   intEL = Erl
580   strES = Err.Description
590   LogError "frmBatchProductIssue", "cmdIssue_Click", intEL, strES, , "Identifier=" & txtIdentifier

End Sub

Private Sub cmdPrintBoth_Click()

      Dim TwoForms As Boolean

10    On Error GoTo cmdPrintBoth_Click_Error

20    TwoForms = False
30    Answer = iMsg("Do you want to print Two Forms?", vbQuestion + vbYesNo)
40    If TimedOut Then Unload Me: Exit Sub
50    If Answer = vbYes Then
60      TwoForms = True
70    End If

80    PrintBatchForm txtSampleID, TimeIssued
90    If TwoForms Then
100     PrintBatchForm txtSampleID, TimeIssued
110   End If
120   PrintBatchLabels txtSampleID

130   Exit Sub

cmdPrintBoth_Click_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmBatchProductIssue", "cmdPrintBoth_Click", intEL, strES

End Sub

Private Sub cmdPrintForm_Click()

10    On Error GoTo cmdPrintForm_Click_Error

20    CurrentReceivedDate = txtReceivedDateTime
30    PrintBatchForm txtSampleID, TimeIssued

40    Exit Sub

cmdPrintForm_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmBatchProductIssue", "cmdPrintForm_Click", intEL, strES

End Sub

Private Sub cmdPrintLabel_Click()

10    PrintBatchLabels txtSampleID

End Sub





Private Sub Form_Activate()

10    If Not boolFormLoaded Then
20        If Trim$(mSampleID) <> "" Then
30          txtSampleID = mSampleID
40          LoadLabNumber
50        End If
60        boolFormLoaded = True
70    End If
      Call FormatGrid
      fmeProducts.Visible = False
      m_Flag = False

End Sub

Private Sub Form_Load()

10    With cmbPatientGroup
20      .Clear
30      .AddItem "O"
40      .AddItem "A"
50      .AddItem "B"
60      .AddItem "AB"
70      .ListIndex = -1
80    End With

90    FillLists

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    mSampleID = ""
20    boolFormLoaded = False

End Sub

Private Sub lblSex_Click()

10    Select Case lblsex
        Case "": lblsex = "Male"
20      Case "Male": lblsex = "Female"
30      Case "Female": lblsex = ""
40    End Select

End Sub


Private Sub txtChart_LostFocus()

      Dim tb As Recordset
      Dim sql As String

10    If Trim$(txtChart) = "" Then Exit Sub

20    sql = "select * from patientdetails where " & _
            "patnum = '" & txtChart & "' "
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If tb.EOF Then
60      txtName = ""
70      txtDoB = ""
80      txtAge = ""
90      lblsex = ""
100     txtAddr(0) = ""
110     txtAddr(1) = ""
120     txtAddr(2) = ""
130     cmbWard = ""
140     cmbClinician = ""
150   Else
160     txtName = tb!Name & ""
170     txtDoB = tb!DoB & ""
180     txtAge = CalcAge(txtDoB)
190     Select Case tb!Sex & ""
          Case "M": lblsex = "Male"
200       Case "F": lblsex = "Female"
210       Case Else: lblsex = ""
220     End Select
230     txtAddr(0) = tb!Addr1 & ""
240     txtAddr(1) = tb!Addr2 & ""
250     txtAddr(2) = tb!Addr3 & ""
260     cmbWard = tb!Ward & ""
270     cmbClinician = tb!Clinician & ""
280   End If
End Sub


Private Sub txtDoB_LostFocus()

10    txtDoB = Convert62Date(txtDoB, BACKWARD)
20    txtAge = CalcAge(txtDoB)

End Sub


Private Sub txtIdentifier_LostFocus()

      Dim BPs As New BatchProducts
      Dim BP As BatchProduct

10    On Error GoTo txtIdentifier_LostFocus_Error

20    ClearDetails
30    cmdIssue.Enabled = False

40    BPs.LoadSpecificIdentifierLatest txtIdentifier
50    If BPs.Count > 0 Then
60      Set BP = BPs.Item(1)
70      lblProduct = BP.Product
80      lblBatchNumber = BP.BatchNumber
90      lblUnitGroup = BP.UnitGroup
100     lblStatus = gEVENTCODES(BP.EventCode).Text
 
110     If BP.EventCode = "C" Or BP.EventCode = "R" Or BP.EventCode = "E" Then
120       lblStatus.BackColor = &HC0FFFF
130       cmdIssue.Enabled = True
140     Else
150       lblStatus.BackColor = vbRed
160     End If
  
170   Else
  
180     iMsg "Identifier not found.", vbCritical
190     If TimedOut Then Unload Me: Exit Sub
200     ClearDetails

210   End If

220   Exit Sub

txtIdentifier_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmBatchProductIssue", "txtIdentifier_LostFocus", intEL, strES

End Sub


Private Sub txtSampleID_LostFocus()

10    LoadLabNumber

End Sub


Private Sub LoadLabNumber()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadLabNumber_Error

20    sql = "select * from patientdetails where " & _
            "labnumber = '" & txtSampleID & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      txtChart = tb!Patnum & ""
70      txtAandE = tb!AandE & ""
80      txtName = tb!Name & ""
90      lblsex = tb!Sex & ""
100     txtAddr(0) = tb!Addr1 & ""
110     txtAddr(1) = tb!Addr2 & ""
120     txtAddr(2) = tb!Addr3 & ""
130     txtSampleDateTime = tb!SampleDate & ""
140     txtReceivedDateTime = tb!DateReceived & ""
150     If Not IsNull(tb!DoB) Then
160       txtDoB = Format(tb!DoB, "dd/mm/yyyy")
170     Else
180       txtDoB = ""
190     End If
200     txtAge = tb!Age & ""
210     cmbWard = tb!Ward & ""
220     cmbClinician = tb!Clinician & ""
230     cmbPatientGroup = tb!fGroup & ""
240   End If

250   Exit Sub

LoadLabNumber_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmBatchProductIssue", "LoadLabNumber", intEL, strES, sql

End Sub
'Private Sub GetProducts(p_SampleID As String)
'    On Error GoTo ERROR_Handler
'
'    Dim sql As String
'    Dim tb As ADODB.Recordset
'    Dim Str As String
'
'    sql = "Select * from ocmRequestDetails Where SampleID = '" & p_SampleID & "' And transA = '1' And Status = 'Process'"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    If Not tb Is Nothing Then
'        If Not tb.EOF Then
'            While Not tb.EOF
'                Str = "" & vbTab & tb!RequestID & vbTab & tb!SampleID & vbTab & tb!UID & vbTab & tb!TestCode
'                flxProducts.AddItem (Str)
'                tb.MoveNext
'            Wend
'        End If
'    End If
'
'    Exit Sub
'ERROR_Handler:
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    LogError "frmBatchProductIssue", "GetProducts", intEL, strES
'End Sub
'
'Private Sub UpdateIdentifier(p_UID As String, p_indentifier As String)
'    On Error GoTo ERROR_Handler
'
'    Dim sql As String
'    Dim tb As ADODB.Recordset
'    Dim l_ID As Integer
'
'    sql = "Update ocmRequestDetails Set indentifier = '" & p_indentifier & "' Where UID = '" & p_UID & "'"
'    Cnxn(0).Execute sql
'    DoEvents
'    DoEvents
'
'    sql = "Select IsNULL(Max(id),0) ID from ocmbtproductsissued"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    If Not tb Is Nothing Then
'        If Not tb.EOF Then
'            l_ID = tb!ID
'        End If
'    End If
'
'    l_ID = l_ID + 1
'
'    sql = "Insert Into ocmbtproductsissued(id,uid,identifier,units) "
'    sql = sql & "Values(" & l_ID & ",'" & p_UID & "','" & p_indentifier & "','1')"
'    Cnxn(0).Execute sql
'    DoEvents
'    DoEvents
'
'    Exit Sub
'ERROR_Handler:
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
''    MsgBox Err.Description
'    LogError "frmBatchProductIssue", "UpdateIdentifier", intEL, strES
'End Sub
'
'Private Sub CountAndUpdate(p_SID As String)
'    On Error GoTo ERROR_Handler
'
'    Dim sql As String
'    Dim tb As ADODB.Recordset
'    Dim tb2 As ADODB.Recordset
'
'    sql = "Select Units, UID from ocmRequestDetails Where SampleID = '" & p_SID & "' And Status In ('Pending','Process')"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    If Not tb Is Nothing Then
'        If Not tb.EOF Then
'            While Not tb.EOF
'                    sql = "Select Count(UID) CUID from ocmbtproductsissued Where UID = '" & tb!UID & "'"
'                    Set tb2 = New Recordset
'                    RecOpenServer 0, tb2, sql
'                    If Not tb2 Is Nothing Then
'                        If Not tb2.EOF Then
'                            If tb2!CUID = tb!Units Then
'                                sql = "Update ocmRequestDetails Set Status = 'Issued' Where UID = '" & tb!UID & "'"
'                                Cnxn(0).Execute sql
'                            End If
'                        End If
'                    End If
'                tb.MoveNext
'            Wend
'        End If
'    End If
'
'    Exit Sub
'ERROR_Handler:
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
''    MsgBox Err.Description
'    LogError "frmBatchProductIssue", "CountAndUpdate", intEL, strES
'End Sub
