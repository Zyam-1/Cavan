VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUnlockReasons 
   Caption         =   "NetAcquire - Unlock Reasons"
   ClientHeight    =   6225
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   9060
   Icon            =   "7frmUnlockReasons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9060
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   180
      TabIndex        =   8
      Top             =   6090
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   495
      Left            =   3990
      Picture         =   "7frmUnlockReasons.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   495
      Left            =   6390
      Picture         =   "7frmUnlockReasons.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   7770
      Picture         =   "7frmUnlockReasons.frx":1376
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   345
      Left            =   960
      TabIndex        =   2
      Top             =   150
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   37509
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   345
      Left            =   2640
      TabIndex        =   1
      Top             =   150
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   37509
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5355
      Left            =   180
      TabIndex        =   0
      Top             =   690
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   9446
      _Version        =   393216
      Cols            =   3
      RowHeightMin    =   500
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   3
      FormatString    =   $"7frmUnlockReasons.frx":18A8
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "and"
      Height          =   195
      Left            =   2310
      TabIndex        =   4
      Top             =   240
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Between"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   210
      Width           =   630
   End
End
Attribute VB_Name = "frmUnlockReasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from UnlockReason where " & _
            "DateTime between '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
            "and '" & Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
            "order by DateTime desc"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF
90      s = Format(tb!DateTime, "dd/mm hh:mm:ss") & vbTab & _
            tb!UserName & vbTab & _
            tb!Reason & ""
100     g.AddItem s
110     tb.MoveNext
120   Loop

130   If g.Rows > 2 Then
140     g.RemoveItem 1
150   End If

160   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmUnlockReasons", "FillG", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdPrint_Click()

      Dim Y As Integer

10    If Not SetFormPrinter() Then Exit Sub

20    Printer.Font.Name = "Courier New"

30    For Y = 0 To g.Rows - 1
40      Printer.Print g.TextMatrix(Y, 0);
50      Printer.Print Tab(20); Left$(g.TextMatrix(Y, 1), 9);
60      Printer.Print Tab(30); g.TextMatrix(Y, 2)
70    Next
80    Printer.EndDoc

End Sub

Private Sub cmdSearch_Click()

10    FillG

End Sub

Private Sub dtFrom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    g.Rows = 2
20    g.AddItem ""
30    g.RemoveItem 1

End Sub


Private Sub dtTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    g.Rows = 2
20    g.AddItem ""
30    g.RemoveItem 1

End Sub




Private Sub Form_Load()

10    Activated = False

20    dtFrom = Format(Now - 7, "dd/mm/yyyy")
30    dtTo = Format(Now, "dd/mm/yyyy")

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
40        FillG
      '**************************************
End Sub





