VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmABConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Antibody Confirmation"
   ClientHeight    =   3960
   ClientLeft      =   1470
   ClientTop       =   2010
   ClientWidth     =   6960
   ControlBox      =   0   'False
   DrawWidth       =   10
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmABConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3960
   ScaleWidth      =   6960
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   330
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3165
      Left            =   300
      TabIndex        =   5
      Top             =   450
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5583
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
      AllowBigSelection=   0   'False
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   525
      Left            =   5310
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   5310
      TabIndex        =   0
      Top             =   420
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   300
      TabIndex        =   6
      Top             =   3720
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lnumber 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1590
      TabIndex        =   4
      Top             =   90
      Width           =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lab Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmABConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim growsave As Integer
Private m_sLabNum As String
Option Compare Binary

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdSave_Click()


      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo cmdSave_Click_Error

20    For n = 1 To g.Rows - 1
30      g.row = n
40      g.col = 0
50      If g <> "" Then
60        sql = "Select * from abconfirmation where " & _
                "LabNumber = '" & LabNum & "' " & _
                "and Antibody = '" & g & "'"
70        Set tb = New Recordset
80        RecOpenServerBB 0, tb, sql
90        If tb.EOF Then
100         tb.AddNew
110       End If
120       tb!LabNumber = LabNum
130       tb!antibody = g
140       g.col = 1
150       tb!Patient = g
160       g.col = 2
170       tb!Positive = g
180       g.col = 3
190       tb!Negative = g
200       tb.Update
210     End If
220   Next

230   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmABConfirm", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub FillG()


      Dim sn As Recordset
      Dim sql As String
      Dim s As String
      Dim n As Integer

10    On Error GoTo FillG_Error

20    sql = "select * from abconfirmation where " & _
            "labnumber = '" & LabNum & "'"
30    Set sn = New Recordset
40    RecOpenServerBB 0, sn, sql
50    n = 1
60    Do While Not sn.EOF
70      s = sn!antibody & vbTab & _
            sn!Patient & vbTab & _
            sn!Positive & vbTab & _
            sn!Negative & ""
80      g.AddItem s, n
90      n = n + 1
100     sn.MoveNext
110   Loop

120   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmABConfirm", "FillG", intEL, strES, sql


End Sub

Private Sub FillList()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillList_Error

20    List1.Clear

30    sql = "select * from reagents where " & _
            "Inuse = 1 " & _
            "and Block = 'Anti Sera' " & _
            "order by listorder"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    Do While Not tb.EOF
70      List1.AddItem tb!Name & ""
80      tb.MoveNext
90    Loop

100   Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmABConfirm", "FillList", intEL, strES, sql


End Sub



Private Sub Form_Load()
      '*****NOTE
          'this code might be dependent on many components so for any future
          'update in code try to keep this code on bottom most line of form load.
          Dim n As Integer

10        g.row = 0
20        For n = 0 To 3
30          g.col = n
40          g = Choose(n + 1, "Antiserum", "Patient", "Pos Control", "Neg Control")
50          g.ColWidth(n) = Choose(n + 1, 1605, 840, 1080, 1080)
60        Next
    
70        FillList
80        FillG
90        lNumber = LabNum
      '**************************************
End Sub

Private Sub g_Click()

      Dim n As Integer
      Dim full As Integer

10    If g.row = 0 Then Exit Sub

20    growsave = g.row

30    Select Case g.col
        Case 0: List1.Visible = True
  Case Else:
40        List1.Visible = False
50        Select Case g
            Case "": g = "Negative"
60          Case "Negative": g = "Positive"
70          Case "Positive": g = ""
80        End Select
90    End Select

100   full = True
110   For n = 0 To 3
120     g.col = n
130     If g = "" Then full = False
140   Next
150   cmdSave.Visible = full

End Sub

Private Sub List1_Click()

10    g.row = growsave
20    g.col = 0
30    g = List1
40    List1.Visible = False

End Sub


Public Property Get LabNum() As String

10        LabNum = m_sLabNum

End Property

Public Property Let LabNum(ByVal sLabNum As String)

10        m_sLabNum = sLabNum

End Property
