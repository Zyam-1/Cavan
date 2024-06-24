VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGeneFreq 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gene Frequency"
   ClientHeight    =   3360
   ClientLeft      =   1830
   ClientTop       =   1875
   ClientWidth     =   5145
   ControlBox      =   0   'False
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
   Icon            =   "7frmGeneFreq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   5145
   Begin VB.CommandButton bupdate 
      Appearance      =   0  'Flat
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2820
      TabIndex        =   20
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2820
      TabIndex        =   19
      Top             =   1860
      Width           =   1215
   End
   Begin VB.ComboBox crace 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2100
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   660
      Width           =   2475
   End
   Begin VB.TextBox tgf 
      BackColor       =   &H8000000E&
      Height          =   285
      Index           =   7
      Left            =   900
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2460
      Width           =   915
   End
   Begin VB.TextBox tgf 
      BackColor       =   &H8000000E&
      Height          =   285
      Index           =   6
      Left            =   900
      MaxLength       =   7
      TabIndex        =   8
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox tgf 
      BackColor       =   &H8000000E&
      Height          =   285
      Index           =   5
      Left            =   900
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1860
      Width           =   915
   End
   Begin VB.TextBox tgf 
      BackColor       =   &H8000000E&
      Height          =   285
      Index           =   4
      Left            =   900
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox tgf 
      BackColor       =   &H8000000E&
      Height          =   285
      Index           =   3
      Left            =   900
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1260
      Width           =   915
   End
   Begin VB.TextBox tgf 
      BackColor       =   &H8000000E&
      Height          =   285
      Index           =   2
      Left            =   900
      MaxLength       =   7
      TabIndex        =   4
      Top             =   960
      Width           =   915
   End
   Begin VB.TextBox tgf 
      BackColor       =   &H8000000E&
      Height          =   285
      Index           =   1
      Left            =   900
      MaxLength       =   7
      TabIndex        =   3
      Top             =   660
      Width           =   915
   End
   Begin VB.TextBox tgf 
      BackColor       =   &H8000000E&
      Height          =   285
      Index           =   0
      Left            =   900
      MaxLength       =   7
      TabIndex        =   2
      Top             =   360
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2820
      TabIndex        =   0
      Top             =   2460
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   210
      TabIndex        =   21
      Top             =   3120
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Race"
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
      Left            =   2100
      TabIndex        =   18
      Top             =   420
      Width           =   390
   End
   Begin VB.Label Label8 
      Caption         =   "CDE"
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
      Left            =   420
      TabIndex        =   16
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "CDe"
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
      Left            =   420
      TabIndex        =   15
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "CdE"
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
      Left            =   420
      TabIndex        =   14
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Cde"
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
      Left            =   420
      TabIndex        =   13
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "cDE"
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
      Left            =   420
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "cDe"
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
      Left            =   420
      TabIndex        =   11
      Top             =   1020
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "cdE"
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
      Left            =   420
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "cde"
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
      Left            =   420
      TabIndex        =   1
      Top             =   420
      Width           =   375
   End
End
Attribute VB_Name = "frmGeneFreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bupdate_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo bupdate_Click_Error

20    sql = "Select * from genefrequency where " & _
            "race = '" & crace.Text & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If tb.EOF Then tb.AddNew

60    tb("race") = crace.Text

70    For n = 0 To 7
80      tb(n + 1) = tgf(n)
90    Next

100   tb.Update

110   fillcrace

120   Exit Sub

bupdate_Click_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmGeneFreq", "bupdate_Click", intEL, strES, sql


End Sub

Private Sub crace_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo crace_Click_Error

20    sql = "Select * from GeneFrequency where " & _
            "Race = '" & crace.Text & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If tb.EOF Then
60      iMsg "Details not found.", vbInformation
70      If TimedOut Then Unload Me: Exit Sub
80    End If

90    For n = 0 To 7
100     tgf(n) = tb(n + 1)
110   Next

120   Exit Sub

crace_Click_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmGeneFreq", "crace_Click", intEL, strES, sql


End Sub

Private Sub fillcrace()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo fillcrace_Error

20    crace.Clear
30    sql = "Select * from GeneFrequency"

40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    Do While Not tb.EOF
70      crace.AddItem tb("race")
80      tb.MoveNext
90    Loop


100   Exit Sub

fillcrace_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmGeneFreq", "fillcrace", intEL, strES, sql


End Sub

Private Sub Form_Load()

10    fillcrace
End Sub

