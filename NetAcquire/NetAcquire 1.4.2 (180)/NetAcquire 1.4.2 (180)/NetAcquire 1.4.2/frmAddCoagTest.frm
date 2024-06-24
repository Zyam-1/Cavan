VERSION 5.00
Begin VB.Form frmAddCoagTest 
   Caption         =   "NetAcquire - Add Coagulation Test"
   ClientHeight    =   2055
   ClientLeft      =   2400
   ClientTop       =   3615
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   6795
   Begin VB.CommandButton bAddUnit 
      Caption         =   "Add New Unit"
      Height          =   315
      Left            =   3030
      TabIndex        =   8
      Top             =   1350
      Width           =   1125
   End
   Begin VB.TextBox tTestName 
      Height          =   285
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   4
      Top             =   840
      Width           =   3105
   End
   Begin VB.TextBox tcode 
      Height          =   285
      Left            =   1035
      MaxLength       =   3
      TabIndex        =   3
      Top             =   300
      Width           =   825
   End
   Begin VB.CommandButton bsave 
      Appearance      =   0  'Flat
      Caption         =   "&Save Details"
      Default         =   -1  'True
      Height          =   525
      Left            =   4650
      TabIndex        =   2
      Top             =   390
      Width           =   1965
   End
   Begin VB.ComboBox cunits 
      Height          =   315
      Left            =   1035
      TabIndex        =   1
      Text            =   "cunits"
      Top             =   1350
      Width           =   1965
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel without Saving"
      Height          =   525
      Left            =   4650
      TabIndex        =   0
      Top             =   1170
      Width           =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Test Name"
      Height          =   195
      Left            =   30
      TabIndex        =   7
      Top             =   870
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   195
      Left            =   495
      TabIndex        =   6
      Top             =   330
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Units"
      Height          =   195
      Left            =   510
      TabIndex        =   5
      Top             =   1410
      Width           =   360
   End
End
Attribute VB_Name = "frmAddCoagTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bAddUnit_Click()

38860     With frmListsGeneric
38870         .ListType = "UN"
38880         .ListTypeName = "Unit"
38890         .ListTypeNames = "Units"
38900         .Show 1
38910     End With

38920     FillLists
        
End Sub


Private Sub bcancel_Click()

38930     Unload Me

End Sub


Private Sub bSave_Click()

          Dim sql As String
          Dim tb As Recordset

38940     On Error GoTo bSave_Click_Error

38950     If Trim$(tCode) = "" Then
38960         iMsg "Enter Code.", vbCritical
38970         Exit Sub
38980     End If

38990     If Trim$(tTestName) = "" Then
39000         iMsg "Enter Test Name.", vbCritical
39010         Exit Sub
39020     End If

39030     If Len(cUnits) = 0 Then
39040         If iMsg("Should Units be blank", vbYesNo) = vbNo Then
39050             iMsg "Select Units.", vbCritical
39060             Exit Sub
39070         End If
39080     End If

39090     sql = "Select * from CoagTestDefinitions where " & _
              "Code = '" & tCode & "'"
39100     Set tb = New Recordset
39110     RecOpenServer 0, tb, sql
39120     If Not tb.EOF Then
39130         iMsg "Code already used.", vbCritical
39140         Exit Sub
39150     Else
       
39160         With tb
39170             .AddNew
39180             !Code = tCode
39190             !TestName = tTestName
39200             !LongName = tTestName
39210             !ShortName = tTestName
39220             !DoDelta = False
39230             !DeltaLimit = 0
39240             !PrintPriority = 999
39250             !DP = 1
39260             !Units = cUnits
39270             !MaleLow = 0
39280             !MaleHigh = 9999
39290             !FemaleLow = 0
39300             !FemaleHigh = 9999
39310             !FlagMaleLow = 0
39320             !FlagMaleHigh = 9999
39330             !FlagFemaleLow = 0
39340             !FlagFemaleHigh = 9999
39350             !Category = ""
39360             !Printable = 1
39370             !PlausibleLow = 0
39380             !PlausibleHigh = 9999
39390             !InUse = 1
39400             !AgeFromDays = 0
39410             !AgeToDays = MaxAgeToDays
39420             !Displayable = 1
39430             !Hospital = "Cavan"
39440             .Update
39450         End With

39460     End If

39470     Unload Me

39480     Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

39490     intEL = Erl
39500     strES = Err.Description
39510     LogError "fAddCoagTest", "bSave_Click", intEL, strES, sql

End Sub


Private Sub Form_Load()

39520     FillLists

End Sub

Private Sub FillLists()

          Dim tb As Recordset
          Dim sql As String

39530     On Error GoTo FillLists_Error

39540     cUnits.Clear
39550     sql = "Select * from Lists where " & _
              "ListType = 'UN' and InUse = 1 " & _
              "order by ListOrder"
39560     Set tb = New Recordset
39570     RecOpenServer 0, tb, sql
39580     Do While Not tb.EOF
39590         cUnits.AddItem tb!Text & ""
39600         tb.MoveNext
39610     Loop

39620     Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

39630     intEL = Erl
39640     strES = Err.Description
39650     LogError "fAddCoagTest", "FillLists", intEL, strES, sql


End Sub




