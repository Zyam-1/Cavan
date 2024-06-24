VERSION 5.00
Begin VB.Form frmCommentsTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Comment Templates"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   1100
      Left            =   5940
      Picture         =   "frmCommentsTemplate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4500
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1100
      Left            =   5940
      Picture         =   "frmCommentsTemplate.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   1100
      Left            =   5940
      Picture         =   "frmCommentsTemplate.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5695
      Width           =   1200
   End
   Begin VB.CheckBox chkInactive 
      Caption         =   "Inactive"
      Height          =   195
      Left            =   4740
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtCommentTemplate 
      Height          =   4755
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2040
      Width           =   5595
   End
   Begin VB.ComboBox cmbCommentName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   5595
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "lblHeading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   2805
      TabIndex        =   10
      Top             =   180
      Width           =   1305
   End
   Begin VB.Label lblCommentID 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6780
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Comment Template Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   1830
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Comment Template"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type in name of comment template or select from drop down list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   4575
   End
End
Attribute VB_Name = "frmCommentsTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CommentDepartment As String

Private Sub SetHeading()

22880     Select Case CommentDepartment
              Case "B": lblHeading = "Biochemistry Comment Templates"
22890         Case "C": lblHeading = "Coagulation Comment Templates"
22900     End Select

End Sub

Private Sub FillCommentNames()

          Dim tb As Recordset
          Dim sql As String

22910     On Error GoTo FillCommentNames_Error

22920     sql = "SELECT * FROM CommentsTemplate WHERE " & _
              "Department = '" & CommentDepartment & "' " & _
              "ORDER BY CommentName"
22930     Set tb = New Recordset
22940     RecOpenClient 0, tb, sql
22950     If Not tb.EOF Then

22960         With cmbCommentName
22970             .Clear
22980             Do While Not tb.EOF
22990                 .AddItem tb!CommentName & ""
23000                 .ItemData(.NewIndex) = tb!CommentID
23010                 tb.MoveNext
23020             Loop
23030         End With

23040     End If

23050     Exit Sub

FillCommentNames_Error:

          Dim strES As String
          Dim intEL As Integer

23060     intEL = Erl
23070     strES = Err.Description
23080     LogError "frmCommentsTemplate", "FillCommentNames", intEL, strES, sql

End Sub



Private Sub chkInactive_Click()
          
23090     cmdSave.Enabled = True

End Sub

Private Sub cmbCommentName_Click()

          Dim sql As String
          Dim tb As Recordset

23100     On Error GoTo cmbCommentName_Click_Error

23110     If cmbCommentName.ListIndex < 0 Then Exit Sub
23120     sql = "Select * From CommentsTemplate Where CommentID = " & cmbCommentName.ItemData(cmbCommentName.ListIndex)
23130     Set tb = New Recordset
23140     RecOpenClient 0, tb, sql
23150     If Not tb.EOF Then
23160         lblCommentID = tb!CommentID
23170         cmbCommentName = tb!CommentName
23180         txtCommentTemplate = tb!CommentTemplate
23190         chkInactive.Value = tb!Inactive
23200     End If


23210     Exit Sub

cmbCommentName_Click_Error:

          Dim strES As String
          Dim intEL As Integer

23220     intEL = Erl
23230     strES = Err.Description
23240     LogError "frmCommentsTemplate", "cmbCommentName_Click", intEL, strES, sql

End Sub

Private Sub cmbCommentName_KeyPress(KeyAscii As Integer)
23250     On Error GoTo cmbCommentName_KeyPress_Error

23260     If Len(cmbCommentName.Text) > 50 Then
23270         KeyAscii = 0
              'cmbCommentName = Left(cmbCommentName, 200)
23280         Exit Sub
23290     End If

23300     lblCommentID = ""
23310     txtCommentTemplate = ""

23320     Exit Sub

cmbCommentName_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

23330     intEL = Erl
23340     strES = Err.Description
23350     LogError "frmCommentsTemplate", "cmbCommentName_KeyPress", intEL, strES

End Sub

Private Sub cmbCommentName_LostFocus()
23360     cmbCommentName_Click
End Sub

Private Sub cmdClear_Click()

23370     On Error GoTo cmdClear_Click_Error

23380     lblCommentID = ""
23390     cmbCommentName.Text = ""
23400     txtCommentTemplate = ""
23410     chkInactive.Value = 0

23420     Exit Sub

cmdClear_Click_Error:

          Dim strES As String
          Dim intEL As Integer

23430     intEL = Erl
23440     strES = Err.Description
23450     LogError "frmCommentsTemplate", "cmdClear_Click", intEL, strES

End Sub

Private Sub cmdExit_Click()

23460     Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String

23470     On Error GoTo cmdSave_Click_Error

23480     If cmbCommentName = "" Then
23490         iMsg "Please enter comment name or select from the list"
23500         cmbCommentName.SetFocus
23510         cmdSave.Enabled = False
23520         Exit Sub
23530     End If

23540     If txtCommentTemplate = "" Then
23550         iMsg "Comment template cannot be empty"
23560         txtCommentTemplate.SetFocus
23570         cmdSave.Enabled = False
23580         Exit Sub
23590     End If

23600     sql = "Select * From CommentsTemplate Where CommentID = " & Val(lblCommentID)
23610     Set tb = New Recordset
23620     RecOpenServer 0, tb, sql
23630     If tb.EOF Then
23640         tb.AddNew
23650     End If

23660     tb!CommentName = cmbCommentName
23670     tb!CommentTemplate = txtCommentTemplate
23680     tb!Inactive = IIf(chkInactive.Value = 1, 1, 0)
23690     tb!UserName = UserName
23700     tb!Department = CommentDepartment
23710     tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")
23720     tb.Update

23730     cmdClear_Click

23740     cmdSave.Enabled = False

23750     FillCommentNames
23760     cmbCommentName.SetFocus

23770     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

23780     intEL = Erl
23790     strES = Err.Description
23800     LogError "frmCommentsTemplate", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

23810     FillCommentNames
23820     SetHeading

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

23830     If cmdSave.Enabled Then
23840         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
23850             Cancel = True
23860         End If
23870     End If

End Sub


Private Sub txtCommentTemplate_KeyUp(KeyCode As Integer, Shift As Integer)
          
23880     cmdSave.Enabled = True

End Sub


