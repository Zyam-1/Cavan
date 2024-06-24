VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fTempsQC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Q.C. Temperatures"
   ClientHeight    =   4050
   ClientLeft      =   555
   ClientTop       =   375
   ClientWidth     =   7395
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "fTempsQC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4050
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Fridges"
      Height          =   1485
      Left            =   180
      TabIndex        =   29
      Top             =   1890
      Width           =   3435
      Begin VB.TextBox t 
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
         Index           =   4
         Left            =   930
         TabIndex        =   32
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox t 
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
         Index           =   5
         Left            =   930
         TabIndex        =   31
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox t 
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
         Index           =   3
         Left            =   930
         TabIndex        =   30
         Top             =   300
         Width           =   975
      End
      Begin VB.PictureBox s 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   5
         Left            =   630
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   1020
         Width           =   315
      End
      Begin VB.PictureBox s 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   630
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   660
         Width           =   315
      End
      Begin VB.PictureBox s 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   3
         Left            =   630
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   35
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(  < - 75  )"
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
         Height          =   195
         Index           =   5
         Left            =   2100
         TabIndex        =   41
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(3.5 - 4.5)"
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
         Height          =   195
         Index           =   4
         Left            =   2100
         TabIndex        =   40
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(3.5 - 4.5)"
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
         Height          =   195
         Index           =   3
         Left            =   2100
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
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
         Height          =   195
         Left            =   330
         TabIndex        =   38
         Top             =   1050
         Width           =   120
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
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
         Height          =   195
         Left            =   330
         TabIndex        =   37
         Top             =   690
         Width           =   120
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
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
         Height          =   195
         Left            =   330
         TabIndex        =   36
         Top             =   330
         Width           =   120
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Blocks"
      Height          =   1485
      Left            =   3720
      TabIndex        =   16
      Top             =   210
      Width           =   3435
      Begin VB.TextBox t 
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
         Index           =   8
         Left            =   870
         TabIndex        =   19
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox t 
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
         Index           =   7
         Left            =   870
         TabIndex        =   18
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox t 
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
         Index           =   6
         Left            =   870
         TabIndex        =   17
         Top             =   300
         Width           =   975
      End
      Begin VB.PictureBox s 
         BackColor       =   &H80000005&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   8
         Left            =   630
         ScaleHeight     =   225
         ScaleWidth      =   195
         TabIndex        =   20
         Top             =   1020
         Width           =   255
      End
      Begin VB.PictureBox s 
         BackColor       =   &H80000005&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   7
         Left            =   630
         ScaleHeight     =   225
         ScaleWidth      =   195
         TabIndex        =   21
         Top             =   660
         Width           =   255
      End
      Begin VB.PictureBox s 
         BackColor       =   &H80000005&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   6
         Left            =   630
         ScaleHeight     =   225
         ScaleWidth      =   195
         TabIndex        =   22
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(24.5 - 25.5)"
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
         Height          =   195
         Index           =   8
         Left            =   2010
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(36.5 - 37.5)"
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
         Height          =   195
         Index           =   7
         Left            =   2010
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(36.5 - 37.5)"
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
         Height          =   195
         Index           =   6
         Left            =   2010
         TabIndex        =   26
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
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
         Height          =   195
         Left            =   270
         TabIndex        =   25
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
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
         Height          =   195
         Left            =   270
         TabIndex        =   24
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
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
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   360
         Width           =   120
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Waterbaths"
      Height          =   1485
      Left            =   180
      TabIndex        =   3
      Top             =   210
      Width           =   3435
      Begin VB.PictureBox SpinButton3 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   630
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   15
         Top             =   1020
         Width           =   285
      End
      Begin VB.PictureBox SpinButton2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   630
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   14
         Top             =   660
         Width           =   285
      End
      Begin VB.PictureBox SpinButton1 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   630
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   13
         Top             =   300
         Width           =   285
      End
      Begin VB.TextBox t 
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
         Index           =   2
         Left            =   930
         TabIndex        =   6
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox t 
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
         Index           =   1
         Left            =   930
         TabIndex        =   5
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox t 
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
         Index           =   0
         Left            =   930
         TabIndex        =   4
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(24.5 - 25.5)"
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
         Height          =   195
         Index           =   2
         Left            =   2070
         TabIndex        =   12
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(36.5 - 37.5)"
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
         Height          =   195
         Index           =   1
         Left            =   2070
         TabIndex        =   11
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(36.5 - 37.5)"
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
         Height          =   195
         Index           =   0
         Left            =   2070
         TabIndex        =   10
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
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
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
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
         Height          =   195
         Left            =   330
         TabIndex        =   8
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
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
         Height          =   195
         Left            =   330
         TabIndex        =   7
         Top             =   360
         Width           =   120
      End
   End
   Begin VB.CommandButton bprint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   4110
      TabIndex        =   2
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   5700
      TabIndex        =   1
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   4095
      TabIndex        =   0
      Top             =   2970
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   195
      TabIndex        =   42
      Top             =   3750
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "fTempsQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdSave_Click()

      Dim mt As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    sql = "Select * from TempsQC"
30    Set mt = New Recordset
40    RecOpenServerBB 0, mt, sql

50    mt.AddNew
60    mt("t0") = t(0)
70    mt("t1") = t(1)
80    mt("t2") = t(2)
90    mt("t3") = t(3)
100   mt("t4") = t(4)
110   mt("t5") = t(5)
120   mt("t6") = t(6)
130   mt("t7") = t(7)
140   mt("t8") = t(8)

150   mt("s0") = lt(0)
160   mt("s1") = lt(1)
170   mt("s2") = lt(2)
180   mt("s3") = lt(3)
190   mt("s4") = lt(4)
200   mt("s5") = lt(5)
210   mt("s6") = lt(6)
220   mt("s7") = lt(7)
230   mt("s8") = lt(8)

240   mt!DateTime = Now
250   mt!Operator = UserCode
260   mt.Update

270   Unload Me

280   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "fTempsQC", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Activate()

      Dim mt As Recordset
      Dim sql As String

10    On Error GoTo Form_Activate_Error

20    sql = "Select * from TempsQC"
30    Set mt = New Recordset
40    RecOpenServerBB 0, mt, sql
50    If Not mt.EOF Then
60      mt.MoveLast

70      t(0) = mt("t0") & ""
80      t(1) = mt("t1") & ""
90      t(2) = mt("t2") & ""
100     t(3) = mt("t3") & ""
110     t(4) = mt("t4") & ""
120     t(5) = mt("t5") & ""
130     t(6) = mt("t6") & ""
140     t(7) = mt("t7") & ""
150     t(8) = mt("t8") & ""

160     lt(0) = mt("s0") & ""
170     lt(1) = mt("s1") & ""
180     lt(2) = mt("s2") & ""
190     lt(3) = mt("s3") & ""
200     lt(4) = mt("s4") & ""
210     lt(5) = mt("s5") & ""
220     lt(6) = mt("s6") & ""
230     lt(7) = mt("s7") & ""
240     lt(8) = mt("s8") & ""
250   End If

260   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "fTempsQC", "Form_Activate", intEL, strES, sql


End Sub

Private Sub lt_DblClick(Index As Integer)

      Dim s As String

10    s = iBOX("Permitted Range.", "Temperature Q.C.", lt(Index))
20    If TimedOut Then Unload Me: Exit Sub

30    If s = "" Then s = "--------"
40    lt(Index).Caption = s

End Sub

Private Sub s_SpinDown(Index As Integer)

10    t(Index) = Format(Val(t(Index)) - 0.1)

End Sub

Private Sub s_SpinUp(Index As Integer)

10    t(Index) = Format(Val(t(Index)) + 0.1)

End Sub

Private Sub t_KeyPress(Index As Integer, KeyAscii As Integer)

      Dim Valid As Integer

10    Valid = InStr("0123456789.-+", Chr(KeyAscii))

20    If Valid Then Exit Sub

30    Beep
40    KeyAscii = 0

End Sub

Private Sub t_LostFocus(Index As Integer)

10    t(Index) = Format(Val(t(Index)))

End Sub


