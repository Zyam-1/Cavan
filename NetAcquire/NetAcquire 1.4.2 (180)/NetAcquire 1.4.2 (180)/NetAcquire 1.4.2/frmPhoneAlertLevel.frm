VERSION 5.00
Begin VB.Form frmPhoneAlertLevel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   2940
      TabIndex        =   7
      Top             =   570
      Width           =   2925
      Begin VB.TextBox txtGreaterThan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         TabIndex        =   0
         Top             =   690
         Width           =   1005
      End
      Begin VB.TextBox txtLessThan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1140
         TabIndex        =   1
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Must be Phoned if the result is:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   2670
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "or <"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   570
         TabIndex        =   8
         Top             =   1185
         Width           =   405
      End
   End
   Begin VB.ListBox lstParameter 
      Height          =   5580
      IntegralHeight  =   0   'False
      Left            =   150
      TabIndex        =   6
      Top             =   660
      Width           =   2655
   End
   Begin VB.ComboBox cmbDiscipline 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   180
      Width           =   1905
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1185
      Left            =   4080
      Picture         =   "frmPhoneAlertLevel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3210
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1185
      Left            =   4080
      Picture         =   "frmPhoneAlertLevel.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Discipline"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmPhoneAlertLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillListBio()

      Dim tb As Recordset
      Dim sql As String

33840 On Error GoTo FillListBio_Error

33850 lstParameter.Clear

33860 sql = "SELECT DISTINCT ShortName FROM BioTestDefinitions " & _
            "ORDER BY ShortName"
33870 Set tb = New Recordset
33880 RecOpenServer 0, tb, sql
33890 Do While Not tb.EOF
33900   lstParameter.AddItem tb!ShortName & ""
33910   tb.MoveNext
33920 Loop

33930 Exit Sub

FillListBio_Error:

      Dim strES As String
      Dim intEL As Integer

33940 intEL = Erl
33950 strES = Err.Description
33960 LogError "frmPhoneAlertLevel", "FillListBio", intEL, strES, sql


End Sub

Private Sub FillListCoag()

      Dim tb As Recordset
      Dim sql As String

33970 On Error GoTo FillListCoag_Error

33980 lstParameter.Clear

33990 sql = "SELECT DISTINCT TestName FROM CoagTestDefinitions " & _
            "ORDER BY TestName"
34000 Set tb = New Recordset
34010 RecOpenServer 0, tb, sql
34020 Do While Not tb.EOF
34030   lstParameter.AddItem tb!TestName & ""
34040   tb.MoveNext
34050 Loop

34060 Exit Sub

FillListCoag_Error:

      Dim strES As String
      Dim intEL As Integer

34070 intEL = Erl
34080 strES = Err.Description
34090 LogError "frmPhoneAlertLevel", "FillListCoag", intEL, strES, sql


End Sub

Private Sub FillListHaem()

      Dim tb As Recordset
      Dim sql As String

34100 On Error GoTo FillListHaem_Error

34110 lstParameter.Clear

34120 sql = "SELECT DISTINCT AnalyteName FROM HaemTestDefinitions " & _
            "ORDER BY AnalyteName"
34130 Set tb = New Recordset
34140 RecOpenServer 0, tb, sql
34150 Do While Not tb.EOF
34160   lstParameter.AddItem tb!AnalyteName & ""
34170   tb.MoveNext
34180 Loop

34190 Exit Sub

FillListHaem_Error:

      Dim strES As String
      Dim intEL As Integer

34200 intEL = Erl
34210 strES = Err.Description
34220 LogError "frmPhoneAlertLevel", "FillListHaem", intEL, strES, sql


End Sub

Private Sub cmbDiscipline_Click()

34230 Select Case cmbDiscipline
        Case "Haematology": FillListHaem
34240   Case "Biochemistry": FillListBio
34250   Case "Coagulation": FillListCoag
34260 End Select

End Sub

Private Sub cmdCancel_Click()

34270 Unload Me

End Sub

Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim Saveable As Boolean

34280 On Error GoTo cmdSave_Click_Error

34290 Saveable = (Val(txtGreaterThan) > 0) Or (Val(txtLessThan) > 0)

34300 sql = "SELECT * FROM PhoneAlertLevel WHERE " & _
            "Discipline = '" & cmbDiscipline & "' " & _
            "AND Parameter = '" & lstParameter & "'"
34310 Set tb = New Recordset
34320 RecOpenServer 0, tb, sql
34330 If tb.EOF And Saveable Then
34340   tb.AddNew
34350   tb!Discipline = cmbDiscipline
34360   tb!Parameter = lstParameter
34370   If Val(txtLessThan) > 0 Then
34380     tb!LessThan = Val(txtLessThan)
34390   Else
34400     tb!LessThan = Null
34410   End If
34420   If Val(txtGreaterThan) > 0 Then
34430     tb!GreaterThan = Val(txtGreaterThan)
34440   Else
34450     tb!GreaterThan = Null
34460   End If
34470   tb.Update
34480 ElseIf Not tb.EOF And Saveable Then
34490   tb!Discipline = cmbDiscipline
34500   tb!Parameter = lstParameter
34510   tb!LessThan = Val(txtLessThan)
34520   tb!GreaterThan = Val(txtGreaterThan)
34530   tb.Update
34540 ElseIf Not tb.EOF Then
34550   tb.Delete
34560 End If

34570 cmdSave.Enabled = False

34580 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

34590 intEL = Erl
34600 strES = Err.Description
34610 LogError "frmPhoneAlertLevel", "cmdsave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

34620 cmbDiscipline.AddItem "Haematology"
34630 cmbDiscipline.AddItem "Biochemistry"
34640 cmbDiscipline.AddItem "Coagulation"

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

34650 If cmdSave.Enabled Then
34660   If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
34670     Cancel = True
34680   End If
34690 End If

End Sub


Private Sub lstParameter_Click()

      Dim tb As Recordset
      Dim sql As String

34700 On Error GoTo lstParameter_Click_Error

34710 sql = "SELECT " & _
            "COALESCE(CAST(LessThan AS nvarchar(10)), '') LT, " & _
            "COALESCE(CAST(GreaterThan AS nvarchar(10)), '') GT " & _
            "FROM PhoneAlertLevel WHERE " & _
            "Discipline = '" & cmbDiscipline & "' " & _
            "AND Parameter = '" & lstParameter & "'"
34720 Set tb = New Recordset
34730 RecOpenServer 0, tb, sql
34740 If Not tb.EOF Then
34750   txtLessThan = tb!LT
34760   txtGreaterThan = tb!GT
34770 Else
34780   txtLessThan = ""
34790   txtGreaterThan = ""
34800 End If

34810 Exit Sub

lstParameter_Click_Error:

      Dim strES As String
      Dim intEL As Integer

34820 intEL = Erl
34830 strES = Err.Description
34840 LogError "frmPhoneAlertLevel", "lstParameter_Click", intEL, strES, sql


End Sub


Private Sub txtGreaterThan_KeyPress(KeyAscii As Integer)

34850 cmdSave.Enabled = True

End Sub


Private Sub txtLessThan_KeyPress(KeyAscii As Integer)

34860 cmdSave.Enabled = True

End Sub


