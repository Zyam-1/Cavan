VERSION 5.00
Begin VB.Form frmBioAddAnalyte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Add New Analyte"
   ClientHeight    =   4065
   ClientLeft      =   705
   ClientTop       =   1275
   ClientWidth     =   5190
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4065
   ScaleWidth      =   5190
   Begin VB.ComboBox cmbAnalyser 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1590
      Width           =   1965
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Text            =   "cmbSampleType"
      Top             =   2760
      Width           =   1965
   End
   Begin VB.ComboBox cmbUnits 
      Height          =   315
      Left            =   1530
      TabIndex        =   5
      Text            =   "cunits"
      Top             =   3360
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   1200
      Left            =   3840
      Picture         =   "frmBioAddAnalyte.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Details"
      Default         =   -1  'True
      Height          =   1200
      Left            =   3840
      Picture         =   "frmBioAddAnalyte.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2190
      Width           =   1965
   End
   Begin VB.TextBox txtShortName 
      Height          =   285
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   0
      Top             =   480
      Width           =   1965
   End
   Begin VB.TextBox txtLongName 
      Height          =   285
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1050
      Width           =   1965
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Analyser"
      Height          =   195
      Left            =   765
      TabIndex        =   13
      Top             =   1650
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "SampleType"
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   2790
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Units"
      Height          =   195
      Left            =   1005
      TabIndex        =   11
      Top             =   3420
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Analyser Code"
      Height          =   195
      Left            =   345
      TabIndex        =   10
      Top             =   2250
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Long Name"
      Height          =   195
      Left            =   540
      TabIndex        =   9
      Top             =   1110
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Short Name"
      Height          =   195
      Left            =   525
      TabIndex        =   8
      Top             =   510
      Width           =   840
   End
End
Attribute VB_Name = "frmBioAddAnalyte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String

Private Sub FillLists()

          Dim tb As Recordset
          Dim sql As String

2950      On Error GoTo FillLists_Error

2960      FillAnalyserList

2970      cmbUnits.Clear
2980      cmbSampleType.Clear

2990      sql = "SELECT ListType, Text FROM Lists WHERE " & _
              "(ListType = 'ST' OR ListType = 'UN') " & _
              "AND InUse = 1 " & _
              "ORDER BY ListOrder"
3000      Set tb = New Recordset
3010      RecOpenServer 0, tb, sql
3020      Do While Not tb.EOF
3030          Select Case tb!ListType & ""
                  Case "ST": cmbSampleType.AddItem tb!Text & ""
3040              Case "UN": cmbUnits.AddItem tb!Text & ""
3050          End Select
3060          tb.MoveNext
3070      Loop

3080      Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

3090      intEL = Erl
3100      strES = Err.Description
3110      LogError "frmBioAddAnalyte", "FillLists", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

3120      Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleType As String

3130      On Error GoTo cmdSave_Click_Error

3140      If Trim$(cmbSampleType) = "" Then
3150          iMsg "Select Sample Type.", vbCritical
3160          Exit Sub
3170      End If

3180      If Trim$(txtCode) = "" Then
3190          iMsg "Enter Code.", vbCritical
3200          Exit Sub
3210      End If

3220      If Trim$(txtShortName) = "" Then
3230          iMsg "Enter Short Name.", vbCritical
3240          Exit Sub
3250      End If

3260      If Trim$(txtLongName) = "" Then
3270          iMsg "Enter Long Name.", vbCritical
3280          Exit Sub
3290      End If

3300      If Trim$(cmbUnits) = "" Then
3310          If iMsg("Do you want to enter Units?", vbQuestion + vbYesNo) = vbYes Then
3320              Exit Sub
3330          End If
3340      End If

3350      If Trim$(cmbAnalyser) = "" Then
3360          iMsg "Select Analyser"
3370          Exit Sub
3380      End If

3390      SampleType = ListCodeFor("ST", cmbSampleType)

3400      sql = "SELECT * FROM " & pDiscipline & "TestDefinitions WHERE 0 = 1"
3410      Set tb = New Recordset
3420      RecOpenClient 0, tb, sql

3430      With tb
3440          .AddNew
3450          !Code = txtCode
3460          !ArchitectCode = txtCode
3470          !ShortName = txtShortName
3480          !LongName = txtLongName
3490          !DoDelta = False
3500          !DeltaLimit = 0
3510          !PrintPriority = 999
3520          !DP = 1
3530          !BarCode = ""
3540          !Units = cmbUnits
3550          !H = False
3560          !s = False
3570          !l = False
3580          !o = False
3590          !g = False
3600          !J = False
3610          !MaleLow = 0
3620          !MaleHigh = 9999
3630          !FemaleLow = 0
3640          !FemaleHigh = 9999
3650          !FlagMaleLow = 0
3660          !FlagMaleHigh = 9999
3670          !FlagFemaleLow = 0
3680          !FlagFemaleHigh = 9999
3690          !SampleType = SampleType
3700          !Category = "Human"
3710          !LControlLow = 0
3720          !LControlHigh = 9999
3730          !NControlLow = 0
3740          !NControlHigh = 9999
3750          !HControlLow = 0
3760          !HControlHigh = 9999
3770          !Printable = False
3780          !PlausibleLow = 0
3790          !PlausibleHigh = 9999
3800          !InUse = 1
3810          !AgeFromDays = 0
3820          !AgeFromText = "0 Days"
3830          !AgeToDays = 43830
3840          !AgeToText = "120 Years"
3850          !Analyser = cmbAnalyser
3860          !KnownToAnalyser = 0
3870          !Hospital = HospName(0)
3880          !ActiveFromDate = Format$("01/01/2006 00:00:01", "dd/MMM/yyyy HH:mm:ss")
3890          !ActiveToDate = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
3900          !PrintRefRange = 1
3910          .Update
        
3920      End With

3930      cmbSampleType.ListIndex = -1
3940      cmbAnalyser.ListIndex = -1
3950      txtCode = ""
3960      txtShortName = ""
3970      txtLongName = ""
3980      cmbUnits.ListIndex = -1

3990      Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

4000      intEL = Erl
4010      strES = Err.Description
4020      LogError "frmBioAddAnalyte", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

4030      FillLists

End Sub


Private Sub FillAnalyserList()

          Dim sql As String
          Dim tb As Recordset

4040      On Error GoTo FillAnalyserList_Error

4050      cmbAnalyser.Clear
4060      sql = "SELECT Text FROM Lists " & _
              "WHERE ListType = 'Analyser' " & _
              "ORDER BY ListOrder"
4070      Set tb = New Recordset
4080      RecOpenServer 0, tb, sql
4090      Do While Not tb.EOF
4100          cmbAnalyser.AddItem tb!Text & ""
4110          tb.MoveNext
4120      Loop
4130      cmbAnalyser.ListIndex = -1

4140      Exit Sub

FillAnalyserList_Error:

          Dim strES As String
          Dim intEL As Integer

4150      intEL = Erl
4160      strES = Err.Description
4170      LogError "frmBioAddAnalyte", "FillAnalyserList", intEL, strES, sql

End Sub

Private Sub txtCode_LostFocus()

          Dim sql As String
          Dim tb As Recordset

4180      On Error GoTo txtCode_LostFocus_Error

4190      If LTrim(RTrim(txtCode)) = "" Then
4200          Exit Sub
4210      End If

4220      sql = "SELECT Count(ArchitectCode) Tot FROM " & pDiscipline & "TestDefinitions " & _
              "WHERE ArchitectCode = '" & AddTicks(txtCode) & "'"
4230      Set tb = New Recordset
4240      RecOpenServer 0, tb, sql
4250      If tb!Tot > 0 Then
4260          iMsg "Code """ & txtCode & """ exists.", vbCritical, , vbRed
4270          txtCode = ""
4280          txtCode.SetFocus
4290      End If

4300      Exit Sub

txtCode_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

4310      intEL = Erl
4320      strES = Err.Description
4330      LogError "frmBioAddAnalyte", "txtCode_LostFocus", intEL, strES, sql

End Sub


Private Sub txtLongName_LostFocus()

          Dim sql As String
          Dim tb As Recordset

4340      On Error GoTo txtLongName_LostFocus_Error

4350      sql = "SELECT Count(LongName) Tot FROM " & pDiscipline & "TestDefinitions " & _
              "WHERE LongName = '" & AddTicks(txtLongName) & "'"
4360      Set tb = New Recordset
4370      RecOpenServer 0, tb, sql
4380      If tb!Tot > 0 Then
4390          iMsg "Long Name """ & txtLongName & """ exists.", vbCritical, , vbRed
4400          txtLongName = ""
4410          txtLongName.SetFocus
4420      End If

4430      Exit Sub

txtLongName_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

4440      intEL = Erl
4450      strES = Err.Description
4460      LogError "frmBioAddAnalyte", "txtLongName_LostFocus", intEL, strES, sql

End Sub


Private Sub txtShortName_LostFocus()

          Dim sql As String
          Dim tb As Recordset

4470      On Error GoTo txtShortName_LostFocus_Error

4480      sql = "SELECT Count(ShortName) Tot FROM " & pDiscipline & "TestDefinitions " & _
              "WHERE ShortName = '" & AddTicks(txtShortName) & "'"
4490      Set tb = New Recordset
4500      RecOpenServer 0, tb, sql
4510      If tb!Tot > 0 Then
4520          iMsg "Short Name """ & txtShortName & """ exists.", vbCritical, , vbRed
4530          txtShortName = ""
4540          txtShortName.SetFocus
4550      End If

4560      Exit Sub

txtShortName_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

4570      intEL = Erl
4580      strES = Err.Description
4590      LogError "frmBioAddAnalyte", "txtShortName_LostFocus", intEL, strES, sql

End Sub



Public Property Let Discipline(ByVal sNewValue As String)

4600      pDiscipline = sNewValue

End Property
