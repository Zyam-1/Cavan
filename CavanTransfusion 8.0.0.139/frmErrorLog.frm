VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmErrorLog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NetAcquire - Error Log"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13980
   Icon            =   "frmErrorLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   13980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   525
      Left            =   10740
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ComboBox cmbMachine 
      Height          =   315
      Left            =   7410
      TabIndex        =   7
      Text            =   "cmbMachine"
      Top             =   390
      Width           =   2445
   End
   Begin VB.ComboBox cmbProcedure 
      Height          =   315
      Left            =   4860
      TabIndex        =   6
      Text            =   "cmbProcedure"
      Top             =   390
      Width           =   2445
   End
   Begin VB.ComboBox cmbModule 
      Height          =   315
      Left            =   2310
      TabIndex        =   5
      Text            =   "cmbModule"
      Top             =   390
      Width           =   2445
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   525
      Left            =   12630
      TabIndex        =   4
      Top             =   240
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Max Records"
      Height          =   675
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   1845
      Begin MSComCtl2.UpDown udRecords 
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   270
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   450
         _Version        =   393216
         Value           =   50
         BuddyControl    =   "lblMaxRecords"
         BuddyDispid     =   196616
         OrigLeft        =   1170
         OrigTop         =   270
         OrigRight       =   1680
         OrigBottom      =   525
         Increment       =   50
         Max             =   10000
         Min             =   50
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMaxRecords 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   915
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdError 
      Height          =   5955
      Left            =   90
      TabIndex        =   0
      Top             =   930
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   10504
      _Version        =   393216
      Cols            =   12
      RowHeightMin    =   400
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   3
      FormatString    =   $"frmErrorLog.frx":08CA
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Module"
      Height          =   195
      Left            =   2340
      TabIndex        =   10
      Top             =   210
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Procedure"
      Height          =   195
      Left            =   4860
      TabIndex        =   9
      Top             =   210
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Machine"
      Height          =   195
      Left            =   7440
      TabIndex        =   8
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "frmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SortOrder As Boolean

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo FillG_Error

20    grdError.Rows = 2
30    grdError.AddItem ""
40    grdError.RemoveItem 1
50    grdError.Visible = False

60    sql = "SELECT TOP " & Val(lblMaxRecords) & " * FROM ErrorLog "
70    If cmbModule <> "" Then
80      sql = sql & "WHERE ModuleName = '" & cmbModule & "'"
90    ElseIf cmbProcedure <> "" Then
100     sql = sql & "WHERE ProcedureName = '" & cmbProcedure & "'"
110   ElseIf cmbMachine <> "" Then
120     sql = sql & "WHERE MachineName = '" & cmbMachine & "'"
130   End If
140   sql = sql & " ORDER BY DateTime DESC"
150   Set tb = New Recordset
160   RecOpenServer 0, tb, sql
170   Do While Not tb.EOF
180     s = vbTab & _
            Format$(tb!DateTime, "General Date") & vbTab & _
            tb!ModuleName & vbTab & _
            tb!ProcedureName & vbTab & _
            tb!ErrorLineNumber & vbTab & _
            tb!SQLStatement & vbTab
190     If tb!ErrorDescription & "" = "" Then
200       s = s & tb!MSG & vbTab
210     Else
220       s = s & tb!ErrorDescription & vbTab
230     End If
240     s = s & tb!UserName & vbTab & _
            tb!MachineName & vbTab & _
            tb!EventDesc & vbTab & _
            tb!AppName & vbTab & _
            tb!Guid & ""
250     grdError.AddItem s
260     tb.MoveNext
270   Loop

280   If grdError.Rows > 2 Then
290     grdError.RemoveItem 1
300   Else
310     s = vbTab & Format$(Now, "General Date") & vbTab & vbTab & "No Entries"
320     grdError.AddItem s
330     grdError.RemoveItem 1
340   End If
350   grdError.Visible = True

360   cmdDelete.Visible = False

370   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "frmErrorLog", "FillG", intEL, strES, sql

End Sub

Private Sub FillCombos()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillCombos_Error

20    cmbModule.Clear
30    cmbProcedure.Clear
40    cmbMachine.Clear

50    sql = "SELECT DISTINCT ModuleName FROM ErrorLog " & _
            "ORDER BY ModuleName"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    Do While Not tb.EOF
90      cmbModule.AddItem tb!ModuleName & ""
100     tb.MoveNext
110   Loop
120   cmbModule.AddItem "", 0

130   sql = "SELECT DISTINCT ProcedureName FROM ErrorLog " & _
            "ORDER BY ProcedureName"
140   Set tb = New Recordset
150   RecOpenServer 0, tb, sql
160   Do While Not tb.EOF
170     cmbProcedure.AddItem tb!ProcedureName & ""
180     tb.MoveNext
190   Loop
200   cmbProcedure.AddItem "", 0

210   sql = "SELECT DISTINCT MachineName FROM ErrorLog " & _
            "ORDER BY MachineName"
220   Set tb = New Recordset
230   RecOpenServer 0, tb, sql
240   Do While Not tb.EOF
250     cmbMachine.AddItem tb!MachineName & ""
260     tb.MoveNext
270   Loop
280   cmbMachine.AddItem "", 0

290   Exit Sub

FillCombos_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "frmErrorLog", "FillCombos", intEL, strES, sql

End Sub

Private Sub cmbMachine_Click()

10    cmbModule = ""
20    cmbProcedure = ""

30    FillG

End Sub


Private Sub cmbModule_Click()

10    cmbProcedure = ""
20    cmbMachine = ""

30    FillG

End Sub


Private Sub cmbProcedure_Click()

10    cmbModule = ""
20    cmbMachine = ""

30    FillG

End Sub


Private Sub cmdDelete_Click()

      Dim sql As String
      Dim StartRow As Integer
      Dim StopRow As Integer
      Dim n As Integer

10    If grdError.RowSel > grdError.Row Then
20      StartRow = grdError.Row
30      StopRow = grdError.RowSel
40    Else
50      StartRow = grdError.RowSel
60      StopRow = grdError.Row
70    End If

80    For n = StartRow To StopRow
90      sql = "DELETE FROM ErrorLog WHERE " & _
              "GUID = '" & grdError.TextMatrix(n, 11) & "'"
100     Cnxn(0).Execute sql
110   Next

120   FillG

End Sub

Private Sub cmdExit_Click()

10    Unload Me

End Sub

Private Sub Form_Load()

10    grdError.ColWidth(11) = 0

20    FillCombos
30    FillG

End Sub


Private Sub grdError_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    With grdError
20      If Not IsDate(.TextMatrix(Row1, .Col)) Then
30        Cmp = 0
40        Exit Sub
50      End If

60      If Not IsDate(.TextMatrix(Row2, .Col)) Then
70        Cmp = 0
80        Exit Sub
90      End If

100     d1 = Format(.TextMatrix(Row1, .Col), "General Date")
110     d2 = Format(.TextMatrix(Row2, .Col), "General Date")
120   End With

130   If SortOrder Then
140     Cmp = Sgn(DateDiff("s", d1, d2))
150   Else
160     Cmp = Sgn(DateDiff("s", d2, d1))
170   End If

End Sub


Private Sub grdError_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    With grdError
  
20      If .MouseRow = 0 And .MouseCol > 0 Then
    
30        If .MouseCol = 1 Then
40          .Sort = 9
50        Else
60          If SortOrder Then
70            .Sort = flexSortGenericAscending
80          Else
90            .Sort = flexSortGenericDescending
100         End If
110       End If
120       SortOrder = Not SortOrder
    
130       .ColSel = .Col
140       .RowSel = .Row
150       cmdDelete.Visible = False
    
160     Else
170       If (.ColSel <> .Col) Or (.RowSel <> .Row) Then
180         cmdDelete.Visible = True
190       Else
200         cmdDelete.Visible = False
210       End If
220     End If

230   End With

End Sub

Private Sub udRecords_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillG

End Sub


