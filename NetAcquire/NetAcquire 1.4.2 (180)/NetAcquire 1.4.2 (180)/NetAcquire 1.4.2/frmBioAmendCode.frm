VERSION 5.00
Begin VB.Form frmBioAmendCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Amend Biochemistry Code"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNewCode 
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2460
      Width           =   1875
   End
   Begin VB.TextBox txtCurrentCode 
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   5
      Top             =   240
      Width           =   1875
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Details"
      Default         =   -1  'True
      Height          =   1245
      Left            =   4110
      Picture         =   "frmBioAmendCode.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel without Saving"
      Height          =   1245
      Left            =   4110
      Picture         =   "frmBioAmendCode.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1530
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "New Analyser Code"
      Height          =   195
      Left            =   345
      TabIndex        =   8
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label lblShortName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1110
      Width           =   1875
   End
   Begin VB.Label lblLongName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   1530
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Short Name"
      Height          =   195
      Left            =   945
      TabIndex        =   4
      Top             =   1140
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Long Name"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Current Analyser Code"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   300
      Width           =   1575
   End
End
Attribute VB_Name = "frmBioAmendCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String

Private Sub cmdCancel_Click()

4610      Unload Me

End Sub


Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim tbNew As Recordset
          Dim f As Field
          Dim sql As String

4620      On Error GoTo cmdSave_Click_Error

4630      If Trim$(txtCurrentCode) = "" Then
4640          iMsg "Enter Current Code.", vbCritical
4650          Exit Sub
4660      End If

4670      If Trim$(txtNewCode) = "" Then
4680          iMsg "Enter New Code.", vbCritical
4690          Exit Sub
4700      End If

4710      sql = "SELECT * FROM " & pDiscipline & "TestDefinitions " & _
              "WHERE Code = '" & txtCurrentCode & "'"
4720      Set tb = New Recordset
4730      RecOpenClient 0, tb, sql
4740      sql = "SELECT * FROM " & pDiscipline & "TestDefinitions WHERE 0 = 1"
4750      Set tbNew = New Recordset
4760      RecOpenClient 0, tbNew, sql
4770      Do While Not tb.EOF 'for each age range
          
4780          tbNew.AddNew
4790          For Each f In tb.Fields
4800              If UCase(f.Name) = "ARCHITECTCODE" Then
4810                  tbNew!ArchitectCode = txtNewCode
4820              ElseIf UCase(f.Name) = "CODE" Then
4830                  tbNew!Code = txtNewCode
4840              Else
4850                  tbNew(f.Name) = tb(f.Name)
4860              End If
4870          Next
4880          tbNew.Update
          
4890          tb.MoveNext
        
4900      Loop

4910      txtCurrentCode = ""
4920      lblShortName = ""
4930      lblLongName = ""
4940      txtNewCode = ""

4950      Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

4960      intEL = Erl
4970      strES = Err.Description
4980      LogError "frmBioAmendCode", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub txtCurrentCode_LostFocus()

          Dim sql As String
          Dim tb As Recordset

4990      On Error GoTo txtCurrentCode_LostFocus_Error
        
5000      If Trim$(txtCurrentCode) = "" Then Exit Sub

5010      lblLongName = ""
5020      lblShortName = ""
5030      txtNewCode = ""

5040      sql = "SELECT LongName, ShortName FROM " & pDiscipline & "TestDefinitions " & _
              "WHERE Code = '" & AddTicks(txtCurrentCode) & "'"
5050      Set tb = New Recordset
5060      RecOpenServer 0, tb, sql
5070      If tb.EOF Then
5080          iMsg "Code """ & txtCurrentCode & """ does not exist.", vbCritical, , vbRed
5090          txtCurrentCode = ""
5100      Else
5110          lblLongName = tb!LongName & ""
5120          lblShortName = tb!ShortName & ""
5130      End If

5140      Exit Sub

txtCurrentCode_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

5150      intEL = Erl
5160      strES = Err.Description
5170      LogError "frmBioAmendCode", "txtCurrentCode_LostFocus", intEL, strES, sql

End Sub


Private Sub txtNewCode_LostFocus()

          Dim sql As String
          Dim tb As Recordset

5180      On Error GoTo txtNewCode_LostFocus_Error
        
5190      If Trim$(txtNewCode) = "" Then Exit Sub

5200      sql = "SELECT COUNT(Code) Tot FROM " & pDiscipline & "TestDefinitions " & _
              "WHERE Code = '" & AddTicks(txtNewCode) & "'"
5210      Set tb = New Recordset
5220      RecOpenServer 0, tb, sql
5230      If tb!Tot > 0 Then
5240          iMsg "Code """ & txtNewCode & """ already exists.", vbCritical, , vbRed
5250          txtNewCode = ""
5260      End If

5270      Exit Sub

txtNewCode_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

5280      intEL = Erl
5290      strES = Err.Description
5300      LogError "frmBioAmendCode", "txtNewCode_LostFocus", intEL, strES, sql

End Sub



Public Property Let Discipline(ByVal sNewValue As String)

5310      pDiscipline = sNewValue

End Property
