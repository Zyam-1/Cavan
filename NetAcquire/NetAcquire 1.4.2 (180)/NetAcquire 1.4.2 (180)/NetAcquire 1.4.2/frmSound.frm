VERSION 5.00
Begin VB.Form frmSound 
   Caption         =   "NetAcquire Sounds"
   ClientHeight    =   3225
   ClientLeft      =   1620
   ClientTop       =   2520
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   8580
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Height          =   465
      Left            =   7170
      Picture         =   "frmSound.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2520
      Width           =   1155
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   7170
      TabIndex        =   14
      Top             =   2010
      Width           =   465
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   345
      Index           =   3
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2010
      Width           =   615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   7170
      TabIndex        =   11
      Top             =   1440
      Width           =   465
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   345
      Index           =   2
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   7170
      TabIndex        =   8
      Top             =   900
      Width           =   465
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   345
      Index           =   1
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   900
      Width           =   615
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   345
      Index           =   0
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   7170
      TabIndex        =   5
      Top             =   360
      Width           =   465
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   15
      Top             =   2010
      Width           =   5385
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   12
      Top             =   1460
      Width           =   5385
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   910
      Width           =   5385
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   5385
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Question"
      Height          =   195
      Left            =   780
      TabIndex        =   3
      Top             =   2070
      Width           =   630
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   16
      Left            =   210
      Picture         =   "frmSound.frx":066A
      Top             =   270
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   64
      Left            =   210
      Picture         =   "frmSound.frx":0AAC
      Top             =   1380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   32
      Left            =   240
      Picture         =   "frmSound.frx":0EEE
      Top             =   1950
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   48
      Left            =   210
      Picture         =   "frmSound.frx":1330
      Top             =   810
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Severe Error"
      Height          =   195
      Left            =   750
      TabIndex        =   2
      Top             =   975
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Information"
      Height          =   195
      Left            =   750
      TabIndex        =   1
      Top             =   1515
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Critical Error"
      Height          =   195
      Left            =   750
      TabIndex        =   0
      Top             =   420
      Width           =   840
   End
End
Attribute VB_Name = "frmSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()
          
      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Description As String

59800 On Error GoTo bCancel_Click_Error

59810 For n = 1 To 4
59820   Description = Choose(n, "SOUNDCRITICAL", "SOUNDSEVERE", _
                                "SOUNDINFORMATION", "SOUNDQUESTION")
59830   sql = "Select * from Options where " & _
              "Description = '" & Description & "'"
59840   Set tb = New Recordset
59850   RecOpenServer 0, tb, sql
59860   If tb.EOF Then
59870     tb.AddNew
59880     tb!Description = Description
59890   End If
59900   tb!Contents = lblPath(n - 1)
59910   tb.Update
59920 Next
          
59930 sysOptSoundCritical(0) = lblPath(0)
59940 sysOptSoundSevere(0) = lblPath(1)
59950 sysOptSoundInformation(0) = lblPath(2)
59960 sysOptSoundQuestion(0) = lblPath(3)

59970 Unload Me

59980 Exit Sub

bCancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

59990 intEL = Erl
60000 strES = Err.Description
60010 LogError "frmSound", "bcancel_Click", intEL, strES, sql


End Sub

Private Sub cmdBrowse_Click(Index As Integer)

      Dim f As Form

60020 Set f = New frmBrowse

60030 f.Show 1

60040 lblPath(Index) = f.lblPathAndFile

60050 Set f = Nothing

End Sub

Private Sub cmdTest_Click(Index As Integer)

60060 PlaySound lblPath(Index), ByVal 0&, SND_FILENAME Or SND_ASYNC

End Sub

Private Sub Form_Load()

60070 lblPath(0) = sysOptSoundCritical(0)
60080 lblPath(1) = sysOptSoundSevere(0)
60090 lblPath(2) = sysOptSoundInformation(0)
60100 lblPath(3) = sysOptSoundQuestion(0)

End Sub


