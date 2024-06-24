VERSION 5.00
Begin VB.Form frmMergeClinsWards 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5295
   ClientLeft      =   1455
   ClientTop       =   885
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraReplace 
      Height          =   2325
      Left            =   3450
      TabIndex        =   5
      Top             =   1470
      Visible         =   0   'False
      Width           =   2865
      Begin VB.TextBox txtReplace 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   8
         Top             =   1050
         Width           =   2355
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Yes"
         Height          =   675
         Left            =   900
         Picture         =   "frmMergeClinsWards.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "with"
         Height          =   195
         Left            =   1260
         TabIndex        =   10
         Top             =   840
         Width           =   285
      End
      Begin VB.Label lblReplace 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   9
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Replace All occurrances of"
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   210
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   4320
      Picture         =   "frmMergeClinsWards.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4260
      Width           =   1155
   End
   Begin VB.ListBox lstSource 
      Height          =   4740
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   3045
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source"
      Height          =   1035
      Left            =   4020
      TabIndex        =   0
      Top             =   150
      Width           =   1995
      Begin VB.OptionButton optSource 
         Caption         =   "Wards"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   2
         Top             =   630
         Width           =   975
      End
      Begin VB.OptionButton optSource 
         Caption         =   "Clinicians"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   540
         TabIndex        =   1
         Top             =   330
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmMergeClinsWards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillList()

      Dim Source As String
      Dim tb As Recordset
      Dim sql As String

36260 On Error GoTo FillList_Error

36270 Source = IIf(optSource(0), "Clinician", "Ward")
        
36280 lstSource.Clear

36290 sql = "Select distinct " & Source & " as Source " & _
            "From Demographics " & _
            "where " & Source & " <> '' " & _
            "and " & Source & " is not null " & _
            "Order by " & Source
36300 Set tb = New Recordset
36310 RecOpenServer 0, tb, sql
36320 Do While Not tb.EOF
36330   lstSource.AddItem tb!Source
36340   tb.MoveNext
36350 Loop

36360 Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

36370 intEL = Erl
36380 strES = Err.Description
36390 LogError "frmMergeClinsWards", "FillList", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

36400 Unload Me

End Sub

Private Sub cmdReplace_Click()

      Dim s As String
      Dim sql As String
      Dim Source As String
      Dim Records As Long

36410 On Error GoTo cmdReplace_Click_Error

36420 Source = IIf(optSource(0), "Clinician", "Ward")

36430 txtReplace = Replace(txtReplace, "'", " ")

36440 s = "If you continue, you will replace" & vbCrLf & _
          "all occurrances of" & vbCrLf & _
          lblReplace & " with" & vbCrLf & _
          txtReplace & "." & vbCrLf & _
          "You will not be able to undo this change!" & vbCrLf & _
          "Do you want to procede?"
36450 If iMsg(s, vbQuestion + vbYesNo, , vbRed) = vbYes Then
36460   s = iBOX("Enter your Password", , , True)
36470   If UCase$(s) <> UCase$(TechnicianPassFor(UserName)) Then
36480     iMsg "Invalid Password." & vbCrLf & "Operation cancelled!", vbInformation
36490     Exit Sub
36500   End If
36510   s = "Replace " & lblReplace & " with" & vbCrLf & _
            txtReplace & "." & vbCrLf & _
            "Are you sure?"
36520   If iMsg(s, vbQuestion + vbYesNo, , vbRed) = vbYes Then
36530     LogToEventLog "Replaced " & lblReplace & " with " & txtReplace
36540     sql = "Update Demographics " & _
                "Set " & Source & " = '" & AddTicks(txtReplace) & "' " & _
                "where " & Source & " = '" & AddTicks(lblReplace) & "'"
36550     Cnxn(0).Execute sql, Records
       
36560     iMsg Format$(Records) & " Record(s) affected."
36570     FillList
36580     fraReplace.Visible = False
36590   Else
36600     iMsg "Operation cancelled!", vbInformation
36610   End If
36620 Else
36630   iMsg "Operation cancelled!", vbInformation
36640 End If

36650 Exit Sub

cmdReplace_Click_Error:

      Dim strES As String
      Dim intEL As Integer

36660 intEL = Erl
36670 strES = Err.Description
36680 LogError "frmMergeClinsWards", "cmdReplace_Click", intEL, strES, sql

          
End Sub

Private Sub Form_Load()

36690 optSource(0) = False
36700 optSource(1) = False

36710 CheckEventLogInDb Cnxn(0)

End Sub

Private Sub lstSource_Click()

36720 fraReplace.Visible = True
36730 lblReplace = lstSource
36740 txtReplace = lstSource

End Sub

Private Sub optSource_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

36750 If Index = 0 Then
36760   If optSource(0) Then
36770     optSource(0).ForeColor = vbBlue
36780     optSource(0).Font.Bold = True
36790     optSource(1).ForeColor = vbBlack
36800     optSource(1).Font.Bold = False
36810   End If
36820 Else
36830   If optSource(1) Then
36840     optSource(1).ForeColor = vbBlue
36850     optSource(1).Font.Bold = True
36860     optSource(0).ForeColor = vbBlack
36870     optSource(0).Font.Bold = False
36880   End If
36890 End If

36900 FillList

End Sub


