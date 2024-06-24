VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmYearlyStats 
   Caption         =   "NetAcquire"
   ClientHeight    =   6015
   ClientLeft      =   765
   ClientTop       =   450
   ClientWidth     =   8730
   Icon            =   "frmYearlyStats.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8730
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   735
      Left            =   6300
      Picture         =   "frmYearlyStats.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   7350
      Picture         =   "frmYearlyStats.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   780
      Width           =   975
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   750
      TabIndex        =   9
      Top             =   510
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2012
      BuddyControl    =   "lblYear"
      BuddyDispid     =   196619
      OrigLeft        =   4440
      OrigTop         =   2670
      OrigRight       =   5205
      OrigBottom      =   2910
      Max             =   2020
      Min             =   1995
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   5220
      Picture         =   "frmYearlyStats.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   780
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3525
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   465
         Left            =   2130
         TabIndex        =   6
         Top             =   750
         Width           =   1095
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Received"
         Height          =   255
         Index           =   4
         Left            =   210
         TabIndex        =   5
         Top             =   1350
         Width           =   1095
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Destroyed"
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   4
         Top             =   1065
         Width           =   1455
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Expired"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   795
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Transfused"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Crossmatched"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1365
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3225
      Left            =   180
      TabIndex        =   10
      Top             =   2280
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   5689
      _Version        =   393216
      Rows            =   13
      Cols            =   9
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
      ScrollBars      =   0
      FormatString    =   $"frmYearlyStats.frx":18A8
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   210
      TabIndex        =   13
      Top             =   5700
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   7170
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2012"
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   810
      Width           =   585
   End
End
Attribute VB_Name = "frmYearlyStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private StartDate As String
Private StopDate As String

Private Sub ClearGrid()

      Dim grp As Integer
      Dim Y As Integer

10    For grp = 1 To 8
20      For Y = 1 To 12
30        g.TextMatrix(Y, grp) = ""
40      Next
50    Next

End Sub

Private Sub GetDates(ByVal MonthNum As Integer)

10    StartDate = Format$("01/" & Format$(MonthNum) & "/" & lblYear, "dd/mmm/yyyy")
20    StopDate = Format$(DateAdd("m", 1, StartDate) - 1, "dd/mmm/yyyy")

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName
20        If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.FontSize = 10
50    Printer.Orientation = vbPRORPortrait

      '****Report heading
60    Printer.Font.Bold = True
70    Printer.Print

      Dim strTitle As String
80    If Opt(0) Then
90        strTitle = Opt(0).Caption
100   ElseIf Opt(1) Then
110       strTitle = Opt(1).Caption
120   ElseIf Opt(2) Then
130       strTitle = Opt(2).Caption
140   ElseIf Opt(3) Then
150       strTitle = Opt(3).Caption
160   ElseIf Opt(4) Then
170       strTitle = Opt(4).Caption
180   End If
190   Printer.Print FormatString("Search Results For " & strTitle, 99, , AlignCenter)
200   Printer.Print FormatString("Yearly Analysis For " & lblYear, 99, , AlignCenter)

      '****Report body
210   Printer.Font.Size = 9

220   For i = 1 To 108
230       Printer.Print "-";
240   Next i
250   Printer.Print

260   Printer.Print FormatString("", 0, "|");
270   Printer.Print FormatString("Month", 34, "|");
280   Printer.Print FormatString("O Pos", 8, "|", AlignCenter);
290   Printer.Print FormatString("A Pos", 8, "|", AlignCenter);
300   Printer.Print FormatString("B Pos", 8, "|", AlignCenter);
310   Printer.Print FormatString("AB Pos", 8, "|", AlignCenter);
320   Printer.Print FormatString("O Neg", 8, "|", AlignCenter);
330   Printer.Print FormatString("A Neg", 8, "|", AlignCenter);
340   Printer.Print FormatString("B Neg", 8, "|", AlignCenter);
350   Printer.Print FormatString("AB Neg", 8, "|", AlignCenter)
360   Printer.Font.Bold = False
370   For i = 1 To 108
380       Printer.Print "-";
390   Next i
400   Printer.Print
410   For Y = 1 To g.Rows - 1
420       Printer.Print FormatString("", 0, "|");
430       Printer.Print FormatString(g.TextMatrix(Y, 0), 34, "|");
440       Printer.Print FormatString(g.TextMatrix(Y, 1), 8, "|", AlignCenter);
450       Printer.Print FormatString(g.TextMatrix(Y, 2), 8, "|", AlignCenter);
460       Printer.Print FormatString(g.TextMatrix(Y, 3), 8, "|", AlignCenter);
470       Printer.Print FormatString(g.TextMatrix(Y, 4), 8, "|", AlignCenter);
480       Printer.Print FormatString(g.TextMatrix(Y, 5), 8, "|", AlignCenter);
490       Printer.Print FormatString(g.TextMatrix(Y, 6), 8, "|", AlignCenter);
500       Printer.Print FormatString(g.TextMatrix(Y, 7), 8, "|", AlignCenter);
510       Printer.Print FormatString(g.TextMatrix(Y, 8), 8, "|", AlignCenter)
 
520   Next

530   For i = 1 To 108
540       Printer.Print "-";
550   Next i
560   Printer.EndDoc



570   For Each Px In Printers
580     If Px.DeviceName = OriginalPrinter Then
590       Set Printer = Px
600       Exit For
610     End If
620   Next
End Sub

Private Sub cmdSearch_Click()

      Dim n As Integer

10    For n = 0 To 4
20      If Opt(n).Value = True Then
30        Select Case n
            Case 0: FillXM
40          Case 1: FillTransfused
50          Case 2: FillExpired
60          Case 3: FillDestroyed
70          Case 4: FillReceived
80        End Select
90      End If
100   Next
    
End Sub

Private Sub FillDestroyed()

      Dim tb As Recordset
      Dim sql As String
      Dim grp As Integer
      Dim Other As Integer
      Dim Y As Integer
      Dim subsql As String

10    On Error GoTo FillDestroyed_Error

20    For grp = 1 To 8
30      For Y = 1 To 12
40        GetDates Y
50        Other = 0
60        subsql = "Select ISBT128 from Latest where " & _
                "DateTime between '" & StartDate & "' " & _
                "and '" & StopDate & " 23:59:59' " & _
                "and Event = 'D' Or Event = 'T' " & _
                "and GroupRh = '" & Group2Bar(g.TextMatrix(0, grp)) & "'"
    
          'Count Returned to supplier for credit
70        sql = "Select Count(*) As CD from Destroy where " & _
                "Unit IN (" & subsql & ") " & _
                "and Reason Not Like '%EXPIRED%' "
80        Set tb = New Recordset
90        RecOpenServerBB 0, tb, sql
    
100       If tb!CD > 0 Then
110           g.TextMatrix(Y, grp) = tb!CD
120       Else
130           g.TextMatrix(Y, grp) = ""
140       End If
    
      '    sql = "Select * from Latest where " & _
      '          "DateTime between '" & StartDate & "' " & _
      '          "and '" & StopDate & " 23:59:59' " & _
      '          "and Event = 'D' " & _
      '          "and GroupRh = '" & Group2Bar(g.TextMatrix(0, grp)) & "'"
      '    Set tb = New Recordset
      '    RecOpenServerBB 0, tb, sql
      '    Do While Not tb.EOF
      '      sql = "Select * from Destroy where Unit = '" & tb!Number & "'"
      '      Set tbD = New Recordset
      '      RecOpenServerBB 0, tbD, sql
      '      If Not tbD.EOF Then
      '        If Trim$(tbD!Reason & "") <> "" Then
      '          Debug.Print tbD!Reason
      '          If InStr(UCase$(tbD!Reason), "EXPIRED") Then
      '            Expired = Expired + 1
      '          Else
      '            Other = Other + 1
      '          End If
      '        Else
      '          Other = Other + 1
      '        End If
      '      End If
      '      tb.MoveNext
      '    Loop
      '    If Other > 0 Then
      '        g.TextMatrix(Y, grp) = Other
      '    Else
      '        g.TextMatrix(Y, grp) = ""
      '    End If
150     Next
160   Next

170   Exit Sub

FillDestroyed_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmYearlyStats", "FillDestroyed", intEL, strES, sql


End Sub

Private Sub FillExpired()

      Dim tb As Recordset
      Dim sql As String
      Dim grp As Integer
      Dim Expired As Integer
      Dim Y As Integer
      Dim subsql As String

10    On Error GoTo FillExpired_Error

20    For grp = 1 To 8
30      For Y = 1 To 12
40        GetDates Y
50        Expired = 0
60        subsql = "Select ISBT128 from Latest where " & _
                "DateTime between '" & StartDate & "' " & _
                "and '" & StopDate & " 23:59:59' " & _
                "and Event = 'D' Or Event = 'T' " & _
                "and GroupRh = '" & Group2Bar(g.TextMatrix(0, grp)) & "'"
    
          'Count Returned to supplier for credit
70        sql = "Select Count(*) As CD from Destroy where " & _
                "Unit IN (" & subsql & ") " & _
                "and Reason Like '%EXPIRED%' "
80        Set tb = New Recordset
90        RecOpenServerBB 0, tb, sql
    
100       If tb!CD > 0 Then
110           g.TextMatrix(Y, grp) = tb!CD
120       Else
130           g.TextMatrix(Y, grp) = ""
140       End If
      '    sql = "Select number from Latest where " & _
      '          "DateTime between '" & StartDate & "' " & _
      '          "and '" & StopDate & " 23:59:59' " & _
      '          "and Event = 'D' " & _
      '          "and GroupRh = '" & Group2Bar(g.TextMatrix(0, grp)) & "'"
      '    Set tb = New Recordset
      '    RecOpenServerBB 0, tb, sql
      '    Do While Not tb.EOF
      '      sql = "Select * from Destroy where Unit = '" & tb!Number & "'"
      '      Set tbD = New Recordset
      '      RecOpenServerBB 0, tbD, sql
      '      If Not tbD.EOF Then
      '        If Trim$(tbD!Reason & "") <> "" Then
      '          Debug.Print tbD!Reason
      '          If InStr(UCase$(tbD!Reason), "EXPIRED") Then
      '            Expired = Expired + 1
      '          End If
      '        End If
      '      End If
      '      tb.MoveNext
      '    Loop
      '    If Expired > 0 Then
      '      g.TextMatrix(Y, grp) = Expired
      '    Else
      '      g.TextMatrix(Y, grp) = ""
      '    End If
150     Next
160   Next

170   Exit Sub

FillExpired_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmYearlyStats", "FillExpired", intEL, strES, sql


End Sub

Private Sub cmdXL_Click()
      Dim strHeading As String
      Dim strTitle As String
10    If Opt(0) Then
20        strTitle = Opt(0).Caption
30    ElseIf Opt(1) Then
40        strTitle = Opt(1).Caption
50    ElseIf Opt(2) Then
60        strTitle = Opt(2).Caption
70    ElseIf Opt(3) Then
80        strTitle = Opt(3).Caption
90    ElseIf Opt(4) Then
100       strTitle = Opt(4).Caption
110   End If
120   strHeading = "Search Results For " & strTitle & vbCr
130   strHeading = strHeading & "Yearly Analysis For " & lblYear & vbCr
140   strHeading = strHeading & " " & vbCr
150   ExportFlexGrid g, Me, strHeading

End Sub

Private Sub Form_Load()

      Dim n As Integer
      Dim s As String

10    g.Rows = 2
20    g.AddItem ""
30    g.RemoveItem 1

40    For n = 1 To 12
50      s = Format$("1/" & Format$(n) & "/2005", "mmmm")
60      g.AddItem s
70    Next

80    g.RemoveItem 1

End Sub


Private Sub FillReceived()

      Dim tb As Recordset
      Dim sql As String
      Dim grp As Integer
      Dim Y As Integer

10    On Error GoTo FillReceived_Error

20    For grp = 1 To 8
30      For Y = 1 To 12
40        GetDates Y
50        sql = "Select count (*) as Tot from Product where " & _
                "DateTime between '" & StartDate & "' " & _
                "and '" & StopDate & " 23:59:59' " & _
                "and Event = 'R' " & _
                "and GroupRh = '" & Group2Bar(g.TextMatrix(0, grp)) & "'"
60          Set tb = New Recordset
70          RecOpenServerBB 0, tb, sql
80        If tb!Tot > 0 Then
90          g.TextMatrix(Y, grp) = tb!Tot
100       Else
110         g.TextMatrix(Y, grp) = ""
120       End If
130     Next
140   Next

150   Exit Sub

FillReceived_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmYearlyStats", "FillReceived", intEL, strES, sql


End Sub

Private Sub FillTransfused()

      Dim tb As Recordset
      Dim sql As String
      Dim grp As Integer
      Dim Y As Integer

10    On Error GoTo FillTransfused_Error

20    For grp = 1 To 8
30      For Y = 1 To 12
40        GetDates Y
50        sql = "Select count (*) as Tot from Latest where " & _
                "DateTime between '" & StartDate & "' " & _
                "and '" & StopDate & " 23:59:59' " & _
                "and Event = 'S' " & _
                "and GroupRh = '" & Group2Bar(g.TextMatrix(0, grp)) & "'"
60          Set tb = New Recordset
70          RecOpenServerBB 0, tb, sql
80        If tb!Tot > 0 Then
90          g.TextMatrix(Y, grp) = tb!Tot
100       Else
110         g.TextMatrix(Y, grp) = ""
120       End If
130     Next
140   Next

150   Exit Sub

FillTransfused_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmYearlyStats", "FillTransfused", intEL, strES, sql


End Sub
Private Sub FillXM()

      Dim tb As Recordset
      Dim sql As String
      Dim grp As Integer
      Dim Y As Integer

10    On Error GoTo FillXM_Error

20    For grp = 1 To 8
30      For Y = 1 To 12
40        GetDates Y
50        sql = "Select count (*) as Tot from Product where " & _
                "DateTime between '" & StartDate & "' " & _
                "and '" & StopDate & " 23:59:59' " & _
                "and Event = 'X' " & _
                "and GroupRh = '" & Group2Bar(g.TextMatrix(0, grp)) & "'"
60          Set tb = New Recordset
70          RecOpenServerBB 0, tb, sql
80        If tb!Tot > 0 Then
90          g.TextMatrix(Y, grp) = tb!Tot
100       Else
110         g.TextMatrix(Y, grp) = ""
120       End If
130     Next
140   Next

150   Exit Sub

FillXM_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmYearlyStats", "FillXM", intEL, strES, sql


End Sub
Private Sub Opt_Click(Index As Integer)

10    ClearGrid

End Sub


Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    ClearGrid

End Sub


