VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReRunTimes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire       Re-Run Days "
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Order by"
      Height          =   1095
      Left            =   3900
      TabIndex        =   2
      Top             =   240
      Width           =   1395
      Begin VB.OptionButton optOrderBy 
         Caption         =   "List Order"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   660
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optOrderBy 
         Caption         =   "Alphabetical"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   330
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1125
      Left            =   4080
      Picture         =   "frmReRunTimes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "cancel"
      Top             =   5430
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6315
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   11139
      _Version        =   393216
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
      AllowUserResizing=   1
      FormatString    =   "<Analyte                              |^ReRun Days "
   End
End
Attribute VB_Name = "frmReRunTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim OrderBy As String

48280 On Error GoTo FillG_Error

48290 g.Rows = 2
48300 g.AddItem ""
48310 g.RemoveItem 1

48320 If optOrderBy(0) Then
48330   OrderBy = "LongName"
48340 Else
48350   OrderBy = "PrintPriority"
48360 End If

48370 sql = "SELECT DISTINCT LongName, COALESCE(ReRunDays, '') RR, PrintPriority " & _
            "FROM BioTestDefinitions " & _
            "ORDER BY " & OrderBy
48380 Set tb = New Recordset
48390 RecOpenServer 0, tb, sql
48400 Do While Not tb.EOF
48410   g.AddItem tb!LongName & vbTab & _
                  IIf(Val(tb!RR) > 0, tb!RR, "")
48420   tb.MoveNext
48430 Loop

48440 If g.Rows > 2 Then
48450   g.RemoveItem 1
48460 End If

48470 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

48480 intEL = Erl
48490 strES = Err.Description
48500 LogError "frmReRunTimes", "FillG", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

48510 Unload Me

End Sub


Private Sub Form_Load()

48520 FillG

End Sub

Private Sub g_Click()

      Dim sql As String
      Dim s As String

48530 On Error GoTo g_Click_Error

48540 If g.MouseRow = 0 Then Exit Sub

48550 s = "Analyte :- " & g.TextMatrix(g.row, 0) & vbCrLf & _
          "Enter number of days before Re-Runs are allowed."

48560 g.Enabled = False
48570 cmdCancel.Enabled = False

48580 g.TextMatrix(g.row, 1) = Format$(Val(iBOX(s, , g.TextMatrix(g.row, 1))))

48590 sql = "UPDATE BioTestDefinitions " & _
            "SET ReRunDays = " & Val(g.TextMatrix(g.row, 1)) & " " & _
            "WHERE LongName = '" & g.TextMatrix(g.row, 0) & "'"

48600 Cnxn(0).Execute sql

48610 FillG

48620 g.Enabled = True
48630 cmdCancel.Enabled = True

48640 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

48650 intEL = Erl
48660 strES = Err.Description
48670 LogError "frmReRunTimes", "g_Click", intEL, strES, sql

48680 g.Enabled = True
48690 cmdCancel.Enabled = True

End Sub


Private Sub optOrderByAlpha_Click()

End Sub


Private Sub optOrderByAlpha_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


End Sub


Private Sub optOrderBy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

48700 FillG

End Sub


