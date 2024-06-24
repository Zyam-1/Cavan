VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmBatchIssue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Batch Issue"
   ClientHeight    =   5205
   ClientLeft      =   270
   ClientTop       =   645
   ClientWidth     =   10020
   ControlBox      =   0   'False
   Icon            =   "7frmBatchIssue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkPreview 
      Caption         =   "Print Preview"
      Height          =   255
      Left            =   7200
      TabIndex        =   45
      Top             =   2010
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   3915
      Left            =   990
      TabIndex        =   18
      Top             =   900
      Width           =   4695
      Begin VB.TextBox txtSampleDateTime 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   3390
         Width           =   1365
      End
      Begin VB.TextBox txtReceivedDateTime 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   3660
         Width           =   1365
      End
      Begin VB.TextBox txtAandE 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox tDoB 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   660
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   29
         Top             =   1320
         Width           =   1365
      End
      Begin VB.TextBox tAge 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   28
         Top             =   1320
         Width           =   705
      End
      Begin VB.TextBox txtChart 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   660
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   27
         Top             =   300
         Width           =   1365
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   660
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   26
         Top             =   990
         Width           =   3615
      End
      Begin VB.ComboBox cClinician 
         BackColor       =   &H80000018&
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3270
         Width           =   3615
      End
      Begin VB.ComboBox cWard 
         BackColor       =   &H80000018&
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2940
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   660
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   660
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1980
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   660
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox tAddr 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   660
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2580
         Width           =   3615
      End
      Begin VB.ComboBox cGroup 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "cGroup"
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "A/E"
         Height          =   195
         Left            =   300
         TabIndex        =   43
         Top             =   660
         Width           =   285
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   42
         Top             =   1380
         Width           =   450
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   16
         Left            =   2100
         TabIndex        =   41
         Top             =   1350
         Width           =   285
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   13
         Left            =   210
         TabIndex        =   40
         Top             =   360
         Width           =   375
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   14
         Left            =   3270
         TabIndex        =   39
         Top             =   1350
         Width           =   270
      End
      Begin VB.Label lsex 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3570
         TabIndex        =   38
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   34
         Left            =   120
         TabIndex        =   37
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Index           =   33
         Left            =   180
         TabIndex        =   36
         Top             =   3000
         Width           =   390
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Clin"
         Height          =   195
         Index           =   31
         Left            =   300
         TabIndex        =   35
         Top             =   3300
         Width           =   255
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Addr 1"
         Height          =   195
         Index           =   28
         Left            =   120
         TabIndex        =   34
         Top             =   1740
         Width           =   465
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Index           =   24
         Left            =   480
         TabIndex        =   33
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   195
         Index           =   23
         Left            =   480
         TabIndex        =   32
         Top             =   2340
         Width           =   90
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   195
         Index           =   17
         Left            =   480
         TabIndex        =   31
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Group"
         Height          =   195
         Left            =   3000
         TabIndex        =   30
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdPrintBoth 
      Caption         =   "Print &Both"
      Height          =   795
      Left            =   5850
      Picture         =   "7frmBatchIssue.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2340
      Width           =   1245
   End
   Begin VB.TextBox tTypenex 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   540
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrintLabel 
      Caption         =   "Print &Label"
      Height          =   795
      Left            =   7200
      Picture         =   "7frmBatchIssue.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2340
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrintForm 
      Caption         =   "Print &Form"
      Height          =   795
      Left            =   8550
      Picture         =   "7frmBatchIssue.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2340
      Width           =   1245
   End
   Begin VB.TextBox txtSampleID 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   540
      Width           =   2040
   End
   Begin VB.ComboBox cmbProduct 
      Height          =   315
      Left            =   990
      TabIndex        =   7
      Text            =   "cmbProduct"
      Top             =   120
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Issue Batch Number"
      Height          =   1755
      Left            =   5820
      TabIndex        =   1
      Top             =   60
      Width           =   3945
      Begin VB.ComboBox cUnitGroup 
         Height          =   315
         Left            =   1470
         TabIndex        =   10
         Text            =   "cGroup"
         Top             =   1230
         Width           =   825
      End
      Begin ComCtl2.UpDown udBottles 
         Height          =   225
         Left            =   2010
         TabIndex        =   5
         Top             =   840
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   397
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "tBottles"
         BuddyDispid     =   196634
         OrigLeft        =   1200
         OrigTop         =   780
         OrigRight       =   1515
         OrigBottom      =   975
         Max             =   25
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox tBottles 
         Height          =   285
         Left            =   1500
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "1"
         Top             =   810
         Width           =   510
      End
      Begin VB.ComboBox cBatchNumber 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2115
      End
      Begin VB.CommandButton cmdIssue 
         Caption         =   "Issue"
         Height          =   675
         Left            =   2910
         Picture         =   "7frmBatchIssue.frx":1548
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Group"
         Height          =   195
         Left            =   1020
         TabIndex        =   9
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   870
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   8550
      Picture         =   "7frmBatchIssue.frx":198A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3540
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   990
      TabIndex        =   46
      Top             =   4950
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Typenex"
      Height          =   195
      Left            =   3780
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Lab Number"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   600
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   390
      TabIndex        =   8
      Top             =   180
      Width           =   555
   End
End
Attribute VB_Name = "frmBatchIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String

Private TimeIssued As String
Private boolFormLoaded As Boolean

Public Property Let SampleID(ByVal sNewValue As String)

10    mSampleID = sNewValue

End Property


Private Sub FillLists()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillLists_Error

20    cWard.Clear
30    cClinician.Clear

40    sql = "Select * from Wards order by ListOrder"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    Do While Not tb.EOF
80      cWard.AddItem tb!Text & ""
90      tb.MoveNext
100   Loop

110   sql = "Select * from Clinicians order by listorder"
120   Set tb = New Recordset
130   RecOpenServer 0, tb, sql
140   Do While Not tb.EOF
150     cClinician.AddItem tb!Text & ""
160     tb.MoveNext
170   Loop

180   Exit Sub

FillLists_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmBatchIssue", "FillLists", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdIssue_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim DoB As String
      Dim s As String

10    On Error GoTo cmdIssue_Click_Error

20    If Left$(cGroup & "  ", 2) <> Left$(cUnitGroup & "  ", 2) And cGroup <> "" And cUnitGroup <> "" Then
30      s = "Patient and Unit Group Differ." & vbCrLf & _
            "Continue anyway?"
40      Answer = iMsg(s, vbYesNo + vbQuestion)
50      If TimedOut Then Unload Me: Exit Sub
60      If Answer = vbNo Then
70        Exit Sub
80      End If
90      LogReasonWhy "Batch Product(" & txtSampleID & "): Patient/Unit Group Mis-match. Proceeded", "XM"
100   End If

110   If cBatchNumber = "" Then
120     iMsg "Batch Number?", vbQuestion
130     If TimedOut Then Unload Me: Exit Sub
140     Exit Sub
150   End If

160   If Trim$(txtSampleID) = "" Then
170     iMsg "Specify Lab Number!", vbCritical
180     If TimedOut Then Unload Me: Exit Sub
190     Exit Sub
200   End If

210   If Val(tBottles) = 0 Then
220     iMsg "Number of Units to issue?", vbQuestion
230     If TimedOut Then Unload Me: Exit Sub
240     Exit Sub
250   End If

260   If Trim$(txtName) = "" Then
270     iMsg "Patients Name?", vbQuestion
280     If TimedOut Then Unload Me: Exit Sub
290     Exit Sub
300   End If

310   s = "Issue " & txtName & " with " & Format(tBottles) & " Unit"
320   If Val(tBottles) > 1 Then
330     s = s & "s"
340   End If
350   s = s & " of " & cmbProduct & " Batch " & cBatchNumber & "?"
360   Answer = iMsg(s, vbQuestion + vbYesNo)
370   If TimedOut Then Unload Me: Exit Sub
380   If Answer = vbNo Then
390     Exit Sub
400   End If

410   sql = "Select * from BatchProductList where " & _
            "BatchNumber = '" & cBatchNumber & "' " & _
            "and Product = '" & cmbProduct & "'"
420   Set tb = New Recordset
430   RecOpenServerBB 0, tb, sql
440   If tb.EOF Then
450     iMsg "Batch Number not known!", vbExclamation
460     If TimedOut Then Unload Me: Exit Sub
470     Exit Sub
480   End If
490   If tb!CurrentStock < Val(tBottles) Then
500     iMsg "Only " & tb!CurrentStock & " Units of " & cmbProduct & " in stock!", vbCritical
510     If TimedOut Then Unload Me: Exit Sub
520     tBottles = tb!CurrentStock
530     Exit Sub
540   End If

550   tb!CurrentStock = tb!CurrentStock - Val(tBottles)
560   tb.Update

570   If IsDate(tDoB) Then
580     DoB = Format(tDoB, "dd/mmm/yyyy")
590   Else
600     DoB = vbNullString
610   End If

620   TimeIssued = Format(Now, "dd/mmm/yyyy hh:mm:ss")

630   sql = "Select * from BatchDetails where " & _
            "Chart = 'xxx'"
640   Set tb = New Recordset
650   RecOpenServerBB 0, tb, sql
660   tb.AddNew
670   tb!Chart = txtChart
680   tb!Name = txtName
690   If IsDate(DoB) Then
700     tb!DoB = DoB
710   Else
720     tb!DoB = Null
730   End If
740   tb!Age = tAge
750   tb!Sex = Left$(lSex, 1)
760   tb!Addr0 = tAddr(0)
770   tb!Addr1 = tAddr(1)
780   tb!Addr2 = tAddr(2)
790   tb!Addr3 = tAddr(3)
800   tb!Ward = cWard
810   tb!Clinician = cClinician
820   tb!BatchNumber = cBatchNumber
830   tb!UserCode = UserCode
840   tb!Date = TimeIssued
850   tb!Bottles = Val(tBottles)
860   tb!SampleID = txtSampleID
870   tb!Product = cmbProduct
880   tb!Typenex = tTypenex
890   tb!PatientGroup = cGroup
900   tb!Event = "I"
910   tb.Update

920   tBottles = ""

930   Exit Sub

cmdIssue_Click_Error:

      Dim strES As String
      Dim intEL As Integer

940   intEL = Erl
950   strES = Err.Description
960   LogError "frmBatchIssue", "cmdIssue_Click", intEL, strES, sql


End Sub

Private Sub cmdPrintBoth_Click()

      Dim TwoForms As Boolean

10    TwoForms = False
20    Answer = iMsg("Do you want to print Two Forms?", vbQuestion + vbYesNo)
30    If TimedOut Then Unload Me: Exit Sub
40    If Answer = vbYes Then
50      TwoForms = True
60    End If

70    PrintBatchForm txtSampleID, TimeIssued
80    If TwoForms Then
90      PrintBatchForm txtSampleID, TimeIssued
100   End If
110   PrintBatchLabels txtSampleID

End Sub

Private Sub cmdPrintForm_Click()

10    CurrentReceivedDate = txtReceivedDateTime
20    PrintBatchForm txtSampleID, TimeIssued

End Sub

Private Sub cmdPrintLabel_Click()

10    PrintBatchLabels txtSampleID

End Sub


Private Sub cBatchNumber_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cBatchNumber_Click_Error

20    sql = "Select [Group] from BatchProductList where " & _
            "BatchNumber = '" & cBatchNumber & "' " & _
            "and Product = '" & cmbProduct & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60      cUnitGroup = tb!Group & ""
70    End If

80    Exit Sub

cBatchNumber_Click_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmBatchIssue", "cBatchNumber_Click", intEL, strES, sql


End Sub


Private Sub cmbProduct_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmbProduct_Click_Error

20    sql = "Select * from BatchProductList where " & _
            "DateExpiry > '" & Format$(Now, "dd/mmm/yyyy") & "' " & _
            "and Product = '" & cmbProduct & "' " & _
            "and CurrentStock > 0"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    cBatchNumber.Clear
60    Do While Not tb.EOF
70      cBatchNumber.AddItem tb!BatchNumber & ""
80      tb.MoveNext
90    Loop

100   Exit Sub

cmbProduct_Click_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmBatchIssue", "cmbProduct_Click", intEL, strES, sql


End Sub


Private Sub cmbProduct_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub

Private Sub cUnitGroup_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub




Private Sub Form_Activate()
10    If Not boolFormLoaded Then
20        If Trim$(mSampleID) <> "" Then
30          txtSampleID = mSampleID
40          LoadLabNumber
50        End If
60        boolFormLoaded = True
70    End If
End Sub

Private Sub Form_Load()

10    FillcmbProduct

20    With cGroup
30      .Clear
40      .AddItem "O"
50      .AddItem "A"
60      .AddItem "B"
70      .AddItem "AB"
80      .ListIndex = -1
90    End With

100   With cUnitGroup
110     .Clear
120     .AddItem "O"
130     .AddItem "A"
140     .AddItem "B"
150     .AddItem "AB"
160     .ListIndex = -1
170   End With

180   FillLists


  




End Sub

Private Sub FillcmbProduct()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillcmbProduct_Error

20    sql = "Select * from Lists where " & _
            "ListType = 'B' " & _
            "Order by ListOrder"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    cmbProduct.Clear
60    Do While Not tb.EOF
70      cmbProduct.AddItem tb!Text & ""
80      tb.MoveNext
90    Loop

100   Exit Sub

FillcmbProduct_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmBatchIssue", "FillcmbProduct", intEL, strES, sql


End Sub

Private Sub Form_Unload(Cancel As Integer)


10    mSampleID = ""
20    boolFormLoaded = False
End Sub

Private Sub lsex_Click()

10    Select Case lSex
        Case "": lSex = "Male"
20      Case "Male": lSex = "Female"
30      Case "Female": lSex = ""
40    End Select

End Sub


Private Sub txtChart_LostFocus()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo txtChart_LostFocus_Error

20    If Trim$(txtChart) = "" Then Exit Sub

30    sql = "select * from patientdetails where " & _
            "patnum = '" & txtChart & "' "
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If tb.EOF Then
70      txtName = ""
80      tDoB = ""
90      tAge = ""
100     lSex = ""
110     tAddr(0) = ""
120     tAddr(1) = ""
130     tAddr(2) = ""
140     tAddr(3) = ""
150     cWard = ""
160     cClinician = ""
170   Else
180     txtName = tb!Name & ""
190     tDoB = tb!DoB & ""
200     tAge = CalcAge(tDoB)
210     Select Case tb!Sex & ""
          Case "M": lSex = "Male"
220       Case "F": lSex = "Female"
230       Case Else: lSex = ""
240     End Select
250     tAddr(0) = tb!Addr1 & ""
260     tAddr(1) = tb!Addr2 & ""
270     tAddr(2) = tb!Addr3 & ""
280     tAddr(3) = tb!addr4 & ""
290     cWard = tb!Ward & ""
300     cClinician = tb!Clinician & ""
310   End If

320   Exit Sub

txtChart_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmBatchIssue", "txtChart_LostFocus", intEL, strES, sql


End Sub


Private Sub tDoB_LostFocus()

10    tDoB = Convert62Date(tDoB, BACKWARD)
20    tAge = CalcAge(tDoB)

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
90      lSex = tb!Sex & ""
100     tAddr(0) = tb!Addr1 & ""
110     tAddr(1) = tb!Addr2 & ""
120     tAddr(2) = tb!Addr3 & ""
130     tAddr(3) = tb!addr4 & ""
140     txtSampleDateTime = tb!SampleDate & ""
150     txtReceivedDateTime = tb!DateReceived & ""
160     If Not IsNull(tb!DoB) Then
170       tDoB = Format(tb!DoB, "dd/mm/yyyy")
180     Else
190       tDoB = ""
200     End If
210     tAge = tb!Age & ""
220     cWard = tb!Ward & ""
230     cClinician = tb!Clinician & ""
240     cGroup = tb!fGroup & ""
250   End If

260   Exit Sub

LoadLabNumber_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmBatchIssue", "LoadLabNumber", intEL, strES, sql

End Sub


