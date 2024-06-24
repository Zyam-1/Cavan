VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewAB3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2310
   ClientLeft      =   210
   ClientTop       =   945
   ClientWidth     =   11775
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
   Icon            =   "7frmViewAB3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2310
   ScaleWidth      =   11775
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Default         =   -1  'True
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
      Left            =   10680
      Picture         =   "7frmViewAB3.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   150
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   21
      Cols            =   51
      FixedCols       =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   1
      FormatString    =   "<Donor # |<ABO |<Rh=Hr |^Cell |     "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   30
      TabIndex        =   11
      Top             =   2100
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lbluser 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9870
      TabIndex        =   9
      Top             =   1590
      Width           =   645
   End
   Begin VB.Label lblsupplier 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4830
      TabIndex        =   8
      Top             =   1590
      Width           =   2850
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
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
      Left            =   4260
      TabIndex        =   7
      Top             =   1620
      Width           =   570
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8190
      TabIndex        =   6
      Top             =   1590
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dated"
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
      Left            =   7740
      TabIndex        =   5
      Top             =   1620
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Expires"
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
      Left            =   2250
      TabIndex        =   4
      Top             =   1620
      Width           =   510
   End
   Begin VB.Label lblexpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2790
      TabIndex        =   3
      Top             =   1590
      Width           =   1425
   End
   Begin VB.Label lblissued 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   690
      TabIndex        =   2
      Top             =   1590
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Issued"
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
      Left            =   180
      TabIndex        =   1
      Top             =   1620
      Width           =   465
   End
End
Attribute VB_Name = "frmViewAB3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PanelNotFound As Boolean

Private mLotNumber As String

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub Form_Activate()

10    If PanelNotFound Then cmdCancel.Value = True

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String
      Dim Y As Integer
      Dim Pattern() As String
      Dim Position As Integer

10    On Error GoTo Form_Load_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1
50    g.Rows = 21

60    If mLotNumber = "" Then
70      iMsg "Specify Panel Number.", vbExclamation
80      If TimedOut Then Unload Me: Exit Sub
90    End If

100   PanelNotFound = False

110   sql = "Select * from AntibodyPanels where " & _
            "LotNumber = '" & mLotNumber & "'"

120   Set tb = New Recordset
130   RecOpenServerBB 0, tb, sql

140   If tb.EOF Then
150     iMsg "Panel Number not found.", vbInformation
160     If TimedOut Then Unload Me: Exit Sub
170     PanelNotFound = True
180     Exit Sub
190   End If

200   lblsupplier = tb!Supplier & ""
210   lblissued = tb!IssuedDate & ""
220   lblExpiry = tb!ExpiryDate & ""
230   lbluser = tb!EnteredBy & ""
240   lbldate = tb!DateEntered & ""

250   sql = "Select * from AntibodyPatterns where " & _
            "LotNumber = '" & mLotNumber & "' " & _
            "order by Position"
260   Set tb = New Recordset
270   RecOpenServerBB 0, tb, sql

280   Do While Not tb.EOF
290     Position = tb!Position
300     Pattern = Split(tb!Pattern, vbTab)
310     For Y = 0 To UBound(Pattern)
320       g.TextMatrix(Y, Position) = Pattern(Y)
330     Next
340     If Position > 3 Then
350       g.ColWidth(Position) = TextWidth(Pattern(0) & "W")
360     End If
370     tb.MoveNext
380   Loop

390   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "frmViewAB3", "Form_Load", intEL, strES, sql


End Sub


Public Property Let LotNumber(ByVal strNewValue As String)

10    mLotNumber = strNewValue

End Property
