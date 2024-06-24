VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Administrator Password"
      Height          =   1245
      Left            =   9210
      Picture         =   "frmAdmin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Local Policy"
      Height          =   1845
      Left            =   3750
      TabIndex        =   22
      Top             =   330
      Width           =   6525
      Begin VB.Frame Frame3 
         Caption         =   "Passwords Expire after"
         Height          =   885
         Left            =   3690
         TabIndex        =   27
         Top             =   180
         Width           =   2475
         Begin VB.OptionButton optNever 
            Caption         =   "Never"
            Height          =   195
            Left            =   1320
            TabIndex        =   31
            Top             =   510
            Width           =   735
         End
         Begin VB.OptionButton opt180 
            Caption         =   "180 Days"
            Height          =   195
            Left            =   1320
            TabIndex        =   30
            Top             =   270
            Width           =   975
         End
         Begin VB.OptionButton opt90 
            Alignment       =   1  'Right Justify
            Caption         =   "90 Days"
            Height          =   195
            Left            =   270
            TabIndex        =   29
            Top             =   510
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton opt60 
            Alignment       =   1  'Right Justify
            Caption         =   "60 Days"
            Height          =   195
            Left            =   270
            TabIndex        =   28
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.CheckBox chkAlpha 
         Alignment       =   1  'Right Justify
         Caption         =   "Include Alpha characters"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2085
      End
      Begin MSComCtl2.UpDown udLength 
         Height          =   405
         Left            =   2400
         TabIndex        =   25
         Top             =   1290
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   714
         _Version        =   393216
         Value           =   6
         BuddyControl    =   "lblLength"
         BuddyDispid     =   196624
         OrigLeft        =   2670
         OrigTop         =   1200
         OrigRight       =   2910
         OrigBottom      =   1695
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkShowUser 
         Alignment       =   1  'Right Justify
         Caption         =   "Show User Name"
         Height          =   195
         Left            =   780
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CheckBox chkNumeric 
         Alignment       =   1  'Right Justify
         Caption         =   "Include numeric characters"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   840
         Width           =   2235
      End
      Begin VB.CheckBox chkUpperLower 
         Alignment       =   1  'Right Justify
         Caption         =   "Upper/Lower case mix"
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label lblReUse 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No"
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
         Left            =   5670
         TabIndex        =   33
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password can be Re-Used"
         Height          =   195
         Left            =   3720
         TabIndex        =   32
         Top             =   1380
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Minimum Password Length"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   1380
         Width           =   1890
      End
      Begin VB.Label lblLength 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         Height          =   255
         Left            =   2070
         TabIndex        =   23
         Top             =   1350
         Width           =   330
      End
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9660
      Top             =   4740
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9660
      Top             =   5520
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   885
      Left            =   9270
      Picture         =   "frmAdmin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7110
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   885
      Left            =   9060
      Picture         =   "frmAdmin.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5370
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   885
      Left            =   9060
      Picture         =   "frmAdmin.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4470
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New User"
      Height          =   1845
      Left            =   150
      TabIndex        =   14
      Top             =   330
      Width           =   3555
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2430
         Picture         =   "frmAdmin.frx":1330
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   420
         Width           =   915
      End
      Begin VB.TextBox txtCode 
         DataField       =   "opcode"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   150
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1440
         Width           =   675
      End
      Begin VB.TextBox txtName 
         DataField       =   "opname"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   150
         MaxLength       =   20
         TabIndex        =   0
         Top             =   420
         Width           =   2175
      End
      Begin VB.ComboBox cmbMemberOf 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   900
         Width           =   2205
      End
      Begin VB.TextBox txtAutoLogOff 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Text            =   "5"
         Top             =   1440
         Width           =   465
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   1440
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   503
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "txtAutoLogOff"
         BuddyDispid     =   196635
         OrigLeft        =   2400
         OrigTop         =   2130
         OrigRight       =   2895
         OrigBottom      =   2370
         Max             =   999
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Member Of"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Auto Log Off"
         Height          =   195
         Left            =   1470
         TabIndex        =   18
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
         Height          =   195
         Left            =   2430
         TabIndex        =   17
         Top             =   1485
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Log Off Administrator"
      Height          =   1035
      Left            =   9270
      Picture         =   "frmAdmin.frx":199A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8190
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   12
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6945
      Left            =   150
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2280
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   12250
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
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
      ScrollBars      =   2
      FormatString    =   "^In Use|<Operator Name                             |<Code    |<Member Of             |^Log Off Delay |<Password   |"
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
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer

Private pAdminName As String

Private Function GeneratePassword() As String

      Dim s As String
      Dim MinLen As Integer
      Dim n As Integer

10    MinLen = GetOptionSetting("LogOnMinPassLen", "1")

20    If MinLen < 6 Then
30      MinLen = 6
40    End If

50    Randomize

60    s = Chr$(Int((Asc("Z") - Asc("A") + 1) * Rnd + Asc("A")))
70    For n = 2 To MinLen - 1
80      s = s & Chr$(Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a")))
90    Next
100   s = s & Chr$(Int((Asc("9") - Asc("0") + 1) * Rnd + Asc("0")))

110   GeneratePassword = s

End Function

Private Function GetAdminCount() As Integer
    
      Dim ySave As Integer
      Dim Y As Integer
      Dim Counter As Integer

10    ySave = g.Row

20    Counter = 0
30    For Y = 1 To g.Rows - 1
40      If g.TextMatrix(Y, 3) = "Administrators" Then
50        Counter = Counter + 1
60      End If
70    Next

80    g.Row = ySave
90    GetAdminCount = Counter

End Function

Private Sub LoadLocalPolicy()

      Dim Alpha As Boolean
      Dim PasswordExpiry As String

10    Alpha = GetOptionSetting("LogOnAlpha", False)
20    If Alpha Then
30      chkAlpha.Value = 1
40      chkUpperLower.Enabled = True
50      chkUpperLower.Value = IIf(GetOptionSetting("LogOnUpperLower", False), 1, 0)
60    Else
70      chkAlpha.Value = 0
80      chkUpperLower.Enabled = False
90      chkUpperLower.Value = 0
100   End If

110   chkNumeric.Value = IIf(GetOptionSetting("LogOnNumeric", False), 1, 0)

120   chkShowUser.Value = IIf(GetOptionSetting("LogOnShowUser", False), 1, 0)

130   lblLength = GetOptionSetting("LogOnMinPassLen", "1")

140   PasswordExpiry = GetOptionSetting("PasswordExpiry", "90")
150   Select Case PasswordExpiry
        Case "60": opt60.Value = True
160     Case "90": opt90.Value = True
170     Case "180": opt180.Value = True
180     Case "36525": optNever.Value = True
190   End Select

200   lblReUse.Caption = GetOptionSetting("PasswordReUse", "No")

End Sub

Private Sub chkAlpha_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    SaveOptionSetting "LogOnAlpha", chkAlpha.Value

20    If chkAlpha.Value = False Then
30      chkUpperLower.Enabled = False
40      chkUpperLower.Value = 0
50      SaveOptionSetting "LogOnUpperLower", 0
60    Else
70      chkUpperLower.Enabled = True
80    End If

End Sub


Private Sub chkNumeric_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    SaveOptionSetting "LogOnNumeric", chkNumeric.Value

End Sub


Private Sub chkShowUser_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    SaveOptionSetting "LogOnShowUser", chkShowUser.Value

End Sub


Private Sub chkUpperLower_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    SaveOptionSetting "LogOnUpperLower", chkUpperLower.Value

End Sub


Private Sub cmdadd_Click()

      Dim sql As String
      Dim Password As String

10    On Error GoTo cmdAdd_Click_Error

20    txtName = Trim$(txtName)
30    If txtName = "" Then
40      iMsg "Enter Name of new user", vbExclamation
50      If TimedOut Then Unload Me: Exit Sub
60      Exit Sub
70    End If

80    txtCode = Trim$(txtCode)
90    If txtCode = "" Then
100     iMsg "Enter Code of new user", vbExclamation
110     If TimedOut Then Unload Me: Exit Sub
120     Exit Sub
130   End If

140   If NameHasBeenUsed(txtName) Then
150     iMsg "Name has been used!", vbExclamation
160     If TimedOut Then Unload Me: Exit Sub
170     Exit Sub
180   End If

190   If CodeHasBeenUsed(txtCode) Then
200     iMsg "Code has been used!", vbExclamation
210     If TimedOut Then Unload Me: Exit Sub
220     Exit Sub
230   End If
  
240   Password = GeneratePassword()

250   sql = "INSERT INTO Users (Password, Name, Code, InUse, MemberOf, LogOffDelay, ListOrder, PassDate, ExpiryDate) " & _
            "VALUES ( " & _
            "'" & Password & "', " & _
            "'" & AddTicks(txtName) & "', " & _
            "'" & txtCode & "', " & _
            "'1', " & _
            "'" & cmbMemberOf & "', " & _
            "'" & txtAutoLogOff & "', " & _
            "'1', " & _
            "'" & Format$(Now, "dd/MMM/yyyy") & "', " & _
            "'" & Format$(Now, "dd/MMM/yyyy") & "')"
260   Cnxn(0).Execute sql

270   FillG

280   iMsg "Password assigned to" & vbCrLf & txtName & vbCrLf & Password, vbInformation, , , 12

290   Exit Sub

cmdAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "frmAdmin", "cmdAdd_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdChangePassword_Click()

      Dim NewPass As String
      Dim Confirm As String
      Dim tb As Recordset
      Dim sql As String
      Dim MinLength As Integer

10    On Error GoTo cmdChangePassword_Click_Error

20    NewPass = iBOX("Enter new password", , , True)
30    If TimedOut Then Unload Me: Exit Sub
40    Confirm = iBOX("Confirm password", , , True)
50    If TimedOut Then Unload Me: Exit Sub

60    If NewPass <> Confirm Then
70      iMsg "Passwords don't match!", vbExclamation
80      If TimedOut Then Unload Me: Exit Sub
90      Exit Sub
100   End If

110   MinLength = Val(GetOptionSetting("LogOnMinPassLen", "1"))
120   If Len(NewPass) < MinLength Then
130     iMsg "Passwords must have a minimum of " & Format(MinLength) & " characters!", vbExclamation
140     If TimedOut Then Unload Me: Exit Sub
150     Exit Sub
160   End If

170   If GetOptionSetting("LogOnUpperLower", False) Then
180     If AllLowerCase(NewPass) Or AllUpperCase(NewPass) Then
190       iMsg "Passwords must have a mixture of UPPER CASE and lower case letters!", vbExclamation
200       If TimedOut Then Unload Me: Exit Sub
210       Exit Sub
220     End If
230   End If

240   If GetOptionSetting("LogOnNumeric", False) Then
250     If Not ContainsNumeric(NewPass) Then
260       iMsg "Passwords must contain a numeric character!", vbExclamation
270       If TimedOut Then Unload Me: Exit Sub
280       Exit Sub
290     End If
300   End If

310   If GetOptionSetting("LogOnAlpha", False) Then
320     If Not ContainsAlpha(NewPass) Then
330       iMsg "Passwords must contain an alphabetic character!", vbExclamation
340       If TimedOut Then Unload Me: Exit Sub
350       Exit Sub
360     End If
370   End If

380   If PasswordHasBeenUsed(NewPass) Then
390     iMsg "Password has been used!", vbExclamation
400     If TimedOut Then Unload Me: Exit Sub
410     Exit Sub
420   End If

430   sql = "SELECT * FROM Users WHERE " & _
            "Name = '" & AddTicks(pAdminName) & "'"
440   ArchiveTable "Users", "", sql
450   Cnxn(0).Execute sql
  
460   Set tb = New Recordset
470   RecOpenServer 0, tb, sql
480   If Not tb.EOF Then
490     sql = "UPDATE Users SET " & _
              "PassWord = '" & NewPass & "' WHERE " & _
              "Name = '" & AddTicks(pAdminName) & "'"
500     Cnxn(0).Execute sql
  
510     iMsg "Your Password has been changed.", vbInformation
520     If TimedOut Then Unload Me: Exit Sub
  
530   End If

540   Exit Sub

cmdChangePassword_Click_Error:

      Dim strES As String
      Dim intEL As Integer

550   intEL = Erl
560   strES = Err.Description
570   LogError "frmAdmin", "cmdChangePassword_Click", intEL, strES, sql

End Sub


Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    FireDown

20    tmrDown.Interval = 250
30    FireCounter = 0

40    tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    tmrDown.Enabled = False

End Sub


Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim x As Integer
      Dim VisibleRows As Integer

10    If g.Row = g.Rows - 1 Then Exit Sub
20    n = g.Row

30    FireCounter = FireCounter + 1
40    If FireCounter > 5 Then
50      tmrDown.Interval = 100
60    End If

70    VisibleRows = g.Height \ g.RowHeight(1) - 1

80    g.Visible = False

90    s = ""
100   For x = 0 To g.Cols - 1
110     s = s & g.TextMatrix(n, x) & vbTab
120   Next
130   s = Left$(s, Len(s) - 1)

140   g.RemoveItem n
150   If n < g.Rows Then
160     g.AddItem s, n + 1
170     g.Row = n + 1
180   Else
190     g.AddItem s
200     g.Row = g.Rows - 1
210   End If

220   For x = 0 To g.Cols - 1
230     g.Col = x
240     g.CellBackColor = vbYellow
250   Next

260   If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
270     If g.Row - VisibleRows + 1 > 0 Then
280       g.TopRow = g.Row - VisibleRows + 1
290     End If
300   End If

310   g.Visible = True

320   cmdSave.Visible = True

End Sub
Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim x As Integer

10    If g.Row = 1 Then Exit Sub

20    FireCounter = FireCounter + 1
30    If FireCounter > 5 Then
40      tmrUp.Interval = 100
50    End If

60    n = g.Row

70    g.Visible = False

80    s = ""
90    For x = 0 To g.Cols - 1
100     s = s & g.TextMatrix(n, x) & vbTab
110   Next
120   s = Left$(s, Len(s) - 1)

130   g.RemoveItem n
140   g.AddItem s, n - 1

150   g.Row = n - 1
160   For x = 0 To g.Cols - 1
170     g.Col = x
180     g.CellBackColor = vbYellow
190   Next

200   If Not g.RowIsVisible(g.Row) Then
210     g.TopRow = g.Row
220   End If

230   g.Visible = True

240   cmdSave.Visible = True

End Sub





Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    FireUp

20    tmrUp.Interval = 250
30    FireCounter = 0

40    tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    tmrUp.Enabled = False

End Sub


Private Sub cmdSave_Click()

      Dim sql As String
      Dim Y As Integer

10    On Error GoTo cmdSave_Click_Error

20    For Y = 1 To g.Rows - 1
30      sql = "UPDATE Users " & _
              "SET ListOrder = " & Y & " WHERE " & _
              "Name = '" & AddTicks(g.TextMatrix(Y, 1)) & "' " & _
              "AND Password = '" & g.TextMatrix(Y, 6) & "' " & _
              "COLLATE SQL_Latin1_General_CP1_CS_AS"
40      Cnxn(0).Execute sql
50    Next
60    cmdSave.Visible = False

70    Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmAdmin", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10    FillG
20    LoadLocalPolicy

30    cmbMemberOf.Clear
40    cmbMemberOf.AddItem "Administrators"
50    cmbMemberOf.AddItem "Managers"
60    cmbMemberOf.AddItem "Users"
70    cmbMemberOf.AddItem "LookUp"

End Sub

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "SELECT * FROM Users WHERE " & _
            "MemberOf = 'Administrators' " & _
            "OR MemberOf = 'Managers' " & _
            "OR MemberOf = 'Users' " & _
            "ORDER BY ListOrder"
60    Set tb = New Recordset
70    RecOpenClient 0, tb, sql
80    With tb
90      Do While Not .EOF
100       s = IIf(!InUse, "Yes", "No") & vbTab & _
              !Name & vbTab & _
              !code & vbTab & _
              !MemberOf & vbTab & _
              !LogOffDelay & vbTab & _
              "*****" & vbTab & _
              !Password
110       g.AddItem s
120       .MoveNext
130     Loop
140   End With

150   If g.Rows > 2 Then
160     g.RemoveItem 1
170   End If

180   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmAdmin", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    If cmdSave.Visible Then
20      Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
30      If TimedOut Then Unload Me: Exit Sub
40      If Answer = vbNo Then
50        Cancel = True
60      End If
70    End If

End Sub

Private Sub g_Click()

      Static SortOrder As Boolean
      Dim LogOff As String
      Dim lngNewLogOff As Long
      Dim sql As String
      Dim MemberOf As String
      Dim f As Form
      Dim s(0 To 3) As String
      Dim AdminCount As Integer
      Dim NewPass As String
      Dim x As Integer
      Dim Y As Integer
      Dim ySave As Integer
      Dim xSave As Integer

10    On Error GoTo g_Click_Error

20    If g.MouseRow = 0 Then
30      If SortOrder Then
40        g.Sort = flexSortGenericAscending
50      Else
60        g.Sort = flexSortGenericDescending
70      End If
80      SortOrder = Not SortOrder
90      cmdSave.Visible = True
100     Exit Sub
110   End If

120   cmdMoveUp.Enabled = False
130   cmdMoveDown.Enabled = False
140   ySave = g.Row
150   xSave = g.Col
160   g.Visible = False
170   g.Col = 0
180   For Y = 1 To g.Rows - 1
190     g.Row = Y
200     If g.CellBackColor = vbYellow Then
210       For x = 0 To g.Cols - 1
220         g.Col = x
230         g.CellBackColor = 0
240       Next
250       Exit For
260     End If
270   Next
280   g.Row = ySave
290   g.Col = xSave
300   g.Visible = True

310   Select Case g.Col
        Case 0: 'In Use
320       If g.TextMatrix(g.Row, 0) = "No" Then 'Mark as InUse
330         g.TextMatrix(g.Row, 0) = "Yes"
340         NewPass = GeneratePassword
350         sql = "UPDATE Users " & _
                  "SET InUse = 1, " & _
                  "PassDate = '" & Format$(Now, "dd/MMM/yyyy") & "', " & _
                  "Password = '" & NewPass & "' " & _
                  "WHERE Name = '" & AddTicks(g.TextMatrix(g.Row, 1)) & "' " & _
                  "AND Password = '" & g.TextMatrix(g.Row, 6) & "' " & _
                  "COLLATE SQL_Latin1_General_CP1_CS_AS"
360         Cnxn(0).Execute sql
370         iMsg "Password has been changed for " & g.TextMatrix(g.Row, 1) & vbCrLf & _
                 "New Password : " & NewPass, vbInformation
380         If TimedOut Then Unload Me: Exit Sub
390         FillG
400       Else
410         If g.TextMatrix(g.Row, 3) = "Administrators" Then
420           AdminCount = GetAdminCount()
430           If AdminCount = 1 Then
440             iMsg "At least 1 Administrator must be In-Use", vbCritical
450             If TimedOut Then Unload Me: Exit Sub
460             Exit Sub
470           Else
480             g.TextMatrix(g.Row, 0) = "No"
490             sql = "UPDATE Users " & _
                      "SET InUse = 0 " & _
                      "WHERE Name = '" & AddTicks(g.TextMatrix(g.Row, 1)) & "' " & _
                      "AND Password = '" & g.TextMatrix(g.Row, 6) & "' " & _
                      "COLLATE SQL_Latin1_General_CP1_CS_AS"
500             Cnxn(0).Execute sql
510           End If
520         Else
530           g.TextMatrix(g.Row, 0) = "No"
540           sql = "UPDATE Users " & _
                    "SET InUse = 0 " & _
                    "WHERE Name = '" & AddTicks(g.TextMatrix(g.Row, 1)) & "' " & _
                    "AND Password = '" & g.TextMatrix(g.Row, 6) & "' " & _
                    "COLLATE SQL_Latin1_General_CP1_CS_AS"
550           Cnxn(0).Execute sql
560         End If
570       End If
    
580     Case 1: 'Name
590       For x = 0 To g.Cols - 1
600         g.Col = x
610         g.CellBackColor = vbYellow
620       Next
630       cmdMoveUp.Enabled = True
640       cmdMoveDown.Enabled = True
  
650     Case 3: 'MemberOf
  
660       AdminCount = GetAdminCount()
670       g.Enabled = False
680       s(0) = "Administrators"
690       s(1) = "Managers"
700       s(2) = "Users"
710       Set f = New fcdrDBox
720       With f
730         .Options = s
740         .Prompt = "Enter Member Of Group."
750         .Show 1
760         If TimedOut Then Unload Me: Exit Sub
770         MemberOf = .ReturnValue
780       End With
790       Set f = Nothing
800       If MemberOf <> "" Then
810         If AdminCount = 1 And _
               g.TextMatrix(g.Row, 3) = "Administrators" And _
               MemberOf <> "Administrators" Then
820           iMsg "Cannot demote Administrator."
830           If TimedOut Then Unload Me: Exit Sub
840         Else
850           If g.TextMatrix(g.Row, 3) <> MemberOf Then
860             g.TextMatrix(g.Row, 3) = MemberOf
870             sql = "UPDATE Users " & _
                      "SET MemberOf = '" & MemberOf & "' " & _
                      "WHERE Name = '" & g.TextMatrix(g.Row, 1) & "'"
880             Cnxn(0).Execute sql
890           End If
900         End If
910       End If
920       g.Enabled = True
    
930     Case 4: 'Log Off Delay
940       g.Enabled = False
950       LogOff = g.TextMatrix(g.Row, 4)
960       lngNewLogOff = Val(iBOX("Log Off Delay. (Minutes)", , LogOff))
970       If TimedOut Then Unload Me: Exit Sub
980       If lngNewLogOff > 0 Then
990         g.TextMatrix(g.Row, 4) = Format$(lngNewLogOff)
1000        sql = "UPDATE Users " & _
                  "SET LogOffDelay = " & lngNewLogOff & " " & _
                  "WHERE Name = '" & g.TextMatrix(g.Row, 1) & "'"
1010        Cnxn(0).Execute sql
1020      End If

1030  End Select

1040  Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

1050  intEL = Erl
1060  strES = Err.Description
1070  LogError "frmAdmin", "g_Click", intEL, strES, sql

End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    If g.MouseRow > 0 And g.MouseCol = 5 Then
20      g.ToolTipText = "Password:" & g.TextMatrix(g.MouseRow, 6)
30      Exit Sub
40    ElseIf g.MouseRow > 0 And g.MouseCol = 0 Then
50      g.ToolTipText = "Click to Toggle Yes/No"
60      Exit Sub
70    Else
80      g.ToolTipText = ""
90    End If

End Sub


Private Sub lblReUse_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    lblReUse.Caption = IIf(lblReUse.Caption = "No", "Yes", "No")

20    SaveOptionSetting "PasswordReUse", lblReUse.Caption

End Sub


Private Sub opt180_Click()

10    SaveOptionSetting "PasswordExpiry", "180"

End Sub

Private Sub opt60_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    SaveOptionSetting "PasswordExpiry", "60"

End Sub


Private Sub opt90_Click()

10    SaveOptionSetting "PasswordExpiry", "90"

End Sub


Private Sub optNever_Click()

10    SaveOptionSetting "PasswordExpiry", "36525" '100 years

End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)

10    If KeyAscii = Asc("'") Then
20      KeyAscii = 0
30      Beep
40    End If

End Sub

Private Sub tmrDown_Timer()

10    FireDown

End Sub


Private Sub tmrUp_Timer()

10    FireUp

End Sub


Private Sub udLength_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    SaveOptionSetting "LogOnMinPassLen", lblLength.Caption

End Sub



Public Property Let AdminName(ByVal sNewValue As String)

10    pAdminName = sNewValue

End Property
