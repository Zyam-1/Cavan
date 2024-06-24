VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUnitNotes 
   Caption         =   "NetAcquire - Unit Notes"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   DrawWidth       =   10
   Icon            =   "frmUnitNotes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtExpiry 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   540
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4020
      TabIndex        =   5
      Top             =   150
      Width           =   825
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   150
      Width           =   825
   End
   Begin VB.TextBox txtNotes 
      Height          =   1515
      Left            =   1080
      TabIndex        =   0
      Top             =   930
      Width           =   3795
   End
   Begin VB.TextBox txtUnitNumber 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   180
      Width           =   1545
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   60
      TabIndex        =   7
      Top             =   3000
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Left            =   615
      TabIndex        =   9
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Notes entered by Fred Bloggs dd/mm/yyyy hh:mm:ss"
      Height          =   405
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   3795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Notes"
      Height          =   195
      Left            =   570
      TabIndex        =   3
      Top             =   900
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit Number"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   210
      Width           =   885
   End
End
Attribute VB_Name = "frmUnitNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormLoaded As Boolean

Private Sub FillDetails()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillDetails_Error

20    sql = "Select * from UnitNotes where " & _
            "UnitNumber = '" & txtUnitNumber & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60      txtNotes = tb!Notes & ""
70      lblInfo = "Notes Entered by " & _
                  tb!Technician & " " & _
                  Format$(tb!DateTime, "dd/mm/yyyy hh:mm:ss")
80    Else
90      lblInfo = ""
100     txtNotes = ""
110   End If

120   cmdSave.Enabled = False

130   Exit Sub

FillDetails_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmUnitNotes", "FillDetails", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    If Trim$(txtNotes) = "" Then
30      sql = "Delete from UnitNotes where " & _
              "UnitNumber = '" & txtUnitNumber & "' And DateExpiry = '" & Format(txtExpiry, "dd/mmm/yyyy") & "'"
40      CnxnBB(0).Execute sql
50    Else
60      sql = "Select * from UnitNotes where " & _
              "UnitNumber = '" & txtUnitNumber & "' And DateExpiry = '" & Format(txtExpiry, "dd/mmm/yyyy") & "'"
70      Set tb = New Recordset
80      RecOpenServerBB 0, tb, sql
90      If tb.EOF Then
100       tb.AddNew
110     End If
120     tb!Notes = txtNotes
130     tb!UnitNumber = txtUnitNumber
140     tb!DateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
150     tb!Technician = UserName
160     tb!DateExpiry = Format(txtExpiry, "dd/mmm/yyyy")
170     tb.Update
180   End If

190   Unload Me

200   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmUnitNotes", "cmdSave_Click", intEL, strES, sql


End Sub




Private Sub Form_Activate()
10    If Not FormLoaded Then
20        FormLoaded = True
30        FillDetails
40    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
10    FormLoaded = False
End Sub

Private Sub txtNotes_KeyUp(KeyCode As Integer, Shift As Integer)

10    If txtUnitNumber.Tag = "ISBT128" Then
20    cmdSave.Enabled = True
30    End If

End Sub


Private Sub txtUnitNumber_LostFocus()

10    FillDetails

End Sub

