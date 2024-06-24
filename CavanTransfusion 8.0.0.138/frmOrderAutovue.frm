VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmOrderAutovue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   4215
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   1275
      Left            =   6660
      Picture         =   "frmOrderAutovue.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1125
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "&Order on AutoVue"
      Height          =   1275
      Left            =   6660
      Picture         =   "frmOrderAutovue.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   660
      Width           =   1125
   End
   Begin VB.ListBox lstAutoVue 
      Columns         =   3
      Height          =   3375
      IntegralHeight  =   0   'False
      Left            =   300
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   660
      Width           =   5745
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   330
      TabIndex        =   3
      Top             =   150
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblSampleID 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6165
      TabIndex        =   4
      Top             =   60
      Width           =   2100
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAddTests 
         Caption         =   "&Test List"
      End
   End
End
Attribute VB_Name = "frmOrderAutovue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Private Sub FillKnownOrder()
    
      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

10    On Error GoTo FillKnownOrder_Error

20    sql = "SELECT TestRequired FROM BBOrderComms " & _
            "WHERE SampleID = '" & pSampleID & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    Do While Not tb.EOF
60      For n = 0 To lstAutoVue.ListCount - 1
70        If UCase$(tb!TestRequired & "") = UCase$(lstAutoVue.List(n)) Then
80          lstAutoVue.Selected(n) = True
90          Exit For
100       End If
110     Next
120     tb.MoveNext
130   Loop

140   Exit Sub

FillKnownOrder_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmOrderAutovue", "FillKnownOrder", intEL, strES, sql

End Sub

Private Sub cmdExit_Click()

10    Unload Me

End Sub

Private Sub FillList()

Dim tb As Recordset
Dim sql As String

10    On Error GoTo FillList_Error

20    lstAutoVue.Clear

30    sql = "SELECT Text FROM Lists WHERE " & _
            "ListType = 'AV' " & _
            "ORDER BY ListOrder"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    Do While Not tb.EOF
70      lstAutoVue.AddItem tb!Text & ""
80      tb.MoveNext
90    Loop

100   Exit Sub

FillList_Error:

Dim strES As String
Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmOrderAutovue", "FillList", intEL, strES

End Sub



Public Property Get SampleID() As String

10    SampleID = pSampleID

End Property

Public Property Let SampleID(ByVal sNewValue As String)

10    pSampleID = sNewValue

20    lblSampleID = sNewValue

End Property

Private Sub cmdOrder_Click()

      Dim sql As String
      Dim n As Integer

10    On Error GoTo cmdOrder_Click_Error

20    If lstAutoVue.SelCount = 0 Then
30      iMsg "Select Test to be Ordered", vbExclamation
40      If TimedOut Then Exit Sub: Unload Me
50    End If

60    For n = 0 To lstAutoVue.ListCount - 1
70      If lstAutoVue.Selected(n) Then
80        sql = "IF NOT EXISTS(SELECT * FROM BBOrderComms " & _
                "              WHERE SampleID = '" & pSampleID & "' AND TestRequired = '" & lstAutoVue.List(n) & "') " & _
                "  INSERT INTO BBOrderComms " & _
                "  (TestRequired, UnitNumber, SampleID, Programmed) VALUES " & _
                "  ('" & lstAutoVue.List(n) & "', '','" & pSampleID & "', 0 )"
90        CnxnBB(0).Execute sql
100     End If
110   Next

120   Unload Me

130   Exit Sub

cmdOrder_Click_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmOrderAutovue", "cmdOrder_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10    FillList
20    FillKnownOrder

End Sub
Private Sub mnuAddTests_Click()

10    With flists
20      .ListName = "AV"
30      .oList(7).Value = True
40      .Show 1
50    End With

60    FillList

End Sub


