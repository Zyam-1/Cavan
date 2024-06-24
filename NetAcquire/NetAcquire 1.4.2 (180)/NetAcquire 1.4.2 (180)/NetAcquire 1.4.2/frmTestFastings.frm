VERSION 5.00
Begin VB.Form frmTestFastings 
   Caption         =   "NetAcquire - Fasting Ranges"
   ClientHeight    =   3945
   ClientLeft      =   705
   ClientTop       =   1440
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6540
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   5220
      Picture         =   "frmTestFastings.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   360
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1215
      Left            =   5220
      Picture         =   "frmTestFastings.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2430
      Width           =   1065
   End
   Begin VB.Frame Frame3 
      Caption         =   "Glucose"
      Height          =   1035
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Width           =   4605
      Begin VB.TextBox tText 
         Height          =   285
         Index           =   0
         Left            =   2580
         TabIndex        =   17
         Top             =   510
         Width           =   1815
      End
      Begin VB.TextBox tHigh 
         Height          =   285
         Index           =   0
         Left            =   1380
         TabIndex        =   6
         Top             =   510
         Width           =   975
      End
      Begin VB.TextBox tLow 
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printout Text"
         Height          =   195
         Index           =   0
         Left            =   3030
         TabIndex        =   18
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   3
         Top             =   270
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Triglyceride"
      Height          =   1035
      Left            =   270
      TabIndex        =   1
      Top             =   2670
      Width           =   4605
      Begin VB.TextBox tText 
         Height          =   285
         Index           =   2
         Left            =   2580
         TabIndex        =   20
         Top             =   570
         Width           =   1815
      End
      Begin VB.TextBox tHigh 
         Height          =   285
         Index           =   2
         Left            =   1380
         TabIndex        =   15
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox tLow 
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printout Text"
         Height          =   195
         Index           =   2
         Left            =   3030
         TabIndex        =   22
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   13
         Top             =   330
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   12
         Top             =   330
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cholesterol"
      Height          =   1035
      Left            =   270
      TabIndex        =   0
      Top             =   1380
      Width           =   4605
      Begin VB.TextBox tText 
         Height          =   285
         Index           =   1
         Left            =   2580
         TabIndex        =   19
         Top             =   510
         Width           =   1815
      End
      Begin VB.TextBox tHigh 
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   11
         Top             =   510
         Width           =   975
      End
      Begin VB.TextBox tLow 
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printout Text"
         Height          =   195
         Index           =   1
         Left            =   3030
         TabIndex        =   21
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   8
         Top             =   270
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmTestFastings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

9830  Unload Me

End Sub


Private Sub cmdSave_Click()

      Dim Fx As Fasting
      Dim n As Integer

9840  For n = 0 To 2
9850    Set Fx = New Fasting
9860    Fx.FastingLow = Format$(Val(tLow(n)))
9870    Fx.FastingHigh = Format$(Val(tHigh(n)))
9880    Fx.FastingText = tText(n)
9890    Fx.TestName = Choose(n + 1, "GLU", "CHO", "TRI")
9900    colFastings.Add Fx
9910  Next

9920  colFastings.Refresh

9930  cmdSave.Enabled = False

End Sub

Private Sub Form_Load()

      Dim Fx As Fasting

9940  For Each Fx In colFastings
9950    Select Case Fx.TestName
          Case "GLU"
9960        tLow(0) = Fx.FastingLow
9970        tHigh(0) = Fx.FastingHigh
9980        tText(0) = Fx.FastingText
9990      Case "CHO"
10000       tLow(1) = Fx.FastingLow
10010       tHigh(1) = Fx.FastingHigh
10020       tText(1) = Fx.FastingText
10030     Case "TRI"
10040       tLow(2) = Fx.FastingLow
10050       tHigh(2) = Fx.FastingHigh
10060       tText(2) = Fx.FastingText
10070   End Select
10080 Next

End Sub


Private Sub tHigh_Change(Index As Integer)

10090 tText(Index) = "( " & tLow(Index) & " - " & tHigh(Index) & " )"

End Sub

Private Sub tHigh_KeyPress(Index As Integer, KeyAscii As Integer)

10100 cmdSave.Enabled = True

End Sub


Private Sub tLow_Change(Index As Integer)

10110 tText(Index) = "( " & tLow(Index) & " - " & tHigh(Index) & " )"

End Sub

Private Sub tLow_KeyPress(Index As Integer, KeyAscii As Integer)

10120 cmdSave.Enabled = True

End Sub


Private Sub tText_KeyPress(Index As Integer, KeyAscii As Integer)

10130 cmdSave.Enabled = True

End Sub


