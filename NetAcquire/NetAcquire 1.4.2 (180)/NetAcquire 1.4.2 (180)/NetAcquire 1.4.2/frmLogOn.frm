VERSION 5.00
Begin VB.Form frmLogOn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   885
      Left            =   4350
      Picture         =   "frmLogOn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   510
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   870
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1050
      Width           =   3105
   End
   Begin VB.ComboBox cmbUserName 
      Height          =   315
      Left            =   870
      TabIndex        =   0
      Top             =   270
      Width           =   3105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   390
      TabIndex        =   3
      Top             =   330
      Width           =   420
   End
End
Attribute VB_Name = "frmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillUserNames()

          Dim sql As String
          Dim tb As Recordset

16000     On Error GoTo FillUserNames_Error

16010     cmbUserName.Clear

16020     sql = "SELECT Name FROM Users WHERE " & _
              "MemberOf = 'Administrators' " & _
              "OR MemberOf = 'Managers' " & _
              "OR MemberOf = 'Users' " & _
              "AND InUse = 1 " & _
              "ORDER BY ListOrder"
16030     Set tb = New Recordset
16040     RecOpenServer 0, tb, sql
16050     Do While Not tb.EOF
16060         cmbUserName.AddItem tb!Name & ""
16070         tb.MoveNext
16080     Loop

16090     Exit Sub

FillUserNames_Error:

          Dim strES As String
          Dim intEL As Integer

16100     intEL = Erl
16110     strES = Err.Description
16120     LogError "frmLogOn", "FillUserNames", intEL, strES, sql

End Sub

Private Sub LogOnWithName()

          Dim sql As String
          Dim tb As Recordset
          Dim Diff As Long
          Static LockOutCounter As Integer

16130     On Error GoTo LogOnWithName_Error

16140     sql = "SELECT Code, MemberOf, LogOffDelay, " & _
              "COALESCE(PassDate, getdate()-1) AS PassDate FROM Users WHERE " & _
              "Name = '" & AddTicks(cmbUserName) & "' " & _
              "AND Password = '" & txtPassword & "' "
16150     If GetOptionSetting("LogOnUpperLower", False) Then
16160         sql = sql & "COLLATE SQL_Latin1_General_CP1_CS_AS"
16170     End If
16180     Set tb = New Recordset
16190     RecOpenServer 0, tb, sql
16200     If tb.EOF Then
16210         iMsg "Incorrect Password!", vbExclamation
16220         txtPassword = ""
16230         LockOutCounter = LockOutCounter + 1
16240         If LockOutCounter > 2 Then
16250             sql = "UPDATE Users SET InUse = 0 WHERE " & _
                      "Name = '" & AddTicks(cmbUserName) & "' " & _
                      "AND MemberOf <> 'Administrators'"
16260             Cnxn(0).Execute sql
16270             iMsg "Your account has been de-activated." & vbCrLf & _
                      "Please contact your administrator.", vbInformation
16280             LockOutCounter = 0
16290             Unload Me
16300         End If
16310     Else
16320         LockOutCounter = 0
16330         If tb!MemberOf = "Administrators" Then
16340             frmAdmin.AdminName = cmbUserName
16350             cmbUserName = ""
16360             txtPassword = ""
16370             frmAdmin.Show 1
16380             Unload Me
16390             Exit Sub
16400         Else
16410             Diff = DateDiff("d", Now, tb!PassDate)
16420             If Diff < 0 Then
16430                 iMsg "Your password has expired.", vbInformation
                      '320         If TimedOut Then Unload Me: Exit Sub
16440                 UserName = ""
16450                 UserCode = ""
16460                 UserMemberOf = ""
16470                 Unload Me: Exit Sub
16480             ElseIf Diff = 0 Then
16490                 iMsg "YOUR PASSWORD EXPIRES TODAY!", vbInformation
                      '390         If TimedOut Then Unload Me: Exit Sub
16500             ElseIf Diff < 14 Then
16510                 iMsg "Your password will expire " & IIf(Diff = 1, "tomorrow!", "in " & Diff & " days"), vbInformation
                      '420         If TimedOut Then Unload Me: Exit Sub
16520             End If

16530             UserName = cmbUserName
16540             UserCode = tb!Code & ""
16550             UserMemberOf = Trim$(tb!MemberOf & "")
16560             If Not IsNull(tb!LogOffDelay) Then
16570                 LogOffDelayMin = tb!LogOffDelay
16580             Else
16590                 LogOffDelayMin = 5
16600             End If
16610             LogOffDelaySecs = LogOffDelayMin * 60
16620             If LogOffDelaySecs > 0 And LogOffDelaySecs <= 32767 Then
16630                 frmMain.pBar.max = LogOffDelaySecs
16640             Else
16650                 frmMain.pBar.max = 1
16660             End If

16670             AddActivity "", "NetAcquire Login", "Login", "", "", "", ""
16680             Unload Me

16690         End If
16700     End If

16710     Exit Sub

LogOnWithName_Error:

          Dim strES As String
          Dim intEL As Integer

16720     intEL = Erl
16730     strES = Err.Description
16740     LogError "frmLogOn", "LogOnWithName", intEL, strES, sql

End Sub

Private Sub LogOnNoName()

          Dim sql As String
          Dim tb As Recordset
          Dim Diff As Long
          Static LockOutCounter As Integer

16750     On Error GoTo LogOnNoName_Error

16760     sql = "SELECT Code, MemberOf, LogOffDelay, Name, " & _
              "COALESCE(PassDate, getdate()-1) AS PassDate FROM Users WHERE inuse = 1 and " & _
              "Password = '" & txtPassword & "' "
16770     If GetOptionSetting("LogOnUpperLower", False) Then
16780         sql = sql & "COLLATE SQL_Latin1_General_CP1_CS_AS"
16790     End If
16800     Set tb = New Recordset
16810     RecOpenServer 0, tb, sql
16820     If tb.EOF Then
16830         iMsg "Password unknown!", vbExclamation
16840         txtPassword = ""
16850         LockOutCounter = LockOutCounter + 1
16860         If LockOutCounter > 2 Then
16870             iMsg "Please contact your administrator.", vbInformation
16880             LockOutCounter = 0
16890             Unload Me
16900         End If
16910     Else
16920         LockOutCounter = 0
16930         If tb!MemberOf = "Administrators" Then
16940             frmAdmin.AdminName = "Administrator"
16950             cmbUserName = ""
16960             txtPassword = ""
16970             frmAdmin.Show 1
16980             Unload Me
16990             Exit Sub
17000         ElseIf tb!MemberOf <> "LookUp" Then
17010             Diff = DateDiff("d", Now, tb!PassDate)
17020             If Diff < 0 Then
17030                 iMsg "Your password has expired.", vbInformation
17040                 UserName = ""
17050                 UserCode = ""
17060                 UserMemberOf = ""
17070                 Unload Me: Exit Sub
17080             ElseIf Diff = 0 Then
17090                 iMsg "YOUR PASSWORD EXPIRES TODAY!", vbInformation
17100             ElseIf Diff < 14 Then
17110                 iMsg "Your password will expire " & IIf(Diff = 1, "tomorrow!", "in " & Diff & " days"), vbInformation
17120             End If

17130             UserName = tb!Name
17140             UserCode = tb!Code & ""
17150             UserMemberOf = Trim$(tb!MemberOf & "")
17160             If Not IsNull(tb!LogOffDelay) Then
17170                 LogOffDelayMin = tb!LogOffDelay
17180             Else
17190                 LogOffDelayMin = 5
17200             End If
17210             LogOffDelaySecs = LogOffDelayMin * 60
17220             If LogOffDelaySecs > 0 And LogOffDelaySecs <= 32767 Then
17230                 frmMain.pBar.max = LogOffDelaySecs
17240             Else
17250                 frmMain.pBar.max = 1
17260             End If

17270             AddActivity "", "NetAcquire Login", "Login", "", "", "", ""
17280             Unload Me

17290         Else
17300             UserName = ""
17310             UserCode = ""
17320             UserMemberOf = ""

17330             Unload Me
17340         End If
17350     End If

17360     Exit Sub

LogOnNoName_Error:

          Dim strES As String
          Dim intEL As Integer

17370     intEL = Erl
17380     strES = Err.Description
17390     LogError "frmLogOn", "LogOnNoName", intEL, strES, sql

End Sub


Private Sub cmbUserName_LostFocus()

          Dim tb As Recordset
          Dim sql As String

17400     On Error GoTo cmbUserName_LostFocus_Error

17410     If cmbUserName = "" Then Exit Sub

17420     sql = "SELECT Name FROM Users WHERE " & _
              "(Code = '" & AddTicks(cmbUserName) & "' " & _
              "  OR Name = '" & AddTicks(cmbUserName) & "') " & _
              "AND (MemberOf = 'Administrators' " & _
              "  OR MemberOf = 'Managers' " & _
              "  OR MemberOf = 'Users') " & _
              "AND InUse = 1"
17430     Set tb = New Recordset
17440     RecOpenServer 0, tb, sql
17450     If Not tb.EOF Then
17460         cmbUserName = tb!Name
17470     Else
17480         cmbUserName = ""
17490     End If

17500     Exit Sub

cmbUserName_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

17510     intEL = Erl
17520     strES = Err.Description
17530     LogError "frmLogOn", "cmbUserName_LostFocus", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

17540     On Error GoTo cmdCancel_Click_Error

17550     If GetOptionSetting("LogOnShowName", "1") = "1" Then
17560         LogOnWithName
17570     Else
17580         LogOnNoName
17590     End If

17600     Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

17610     intEL = Erl
17620     strES = Err.Description
17630     LogError "frmLogOn", "cmdCancel_Click", intEL, strES

End Sub

Private Sub Form_Activate()

          Dim CurrentUser As String

17640     CurrentUser = cmbUserName

17650     FillUserNames

17660     cmbUserName = CurrentUser

End Sub

Private Sub Form_Load()

          ' unique usernames - policy length 6 characters
          ' minimum length - 6 characters
          ' alphanumeric
          ' a user can change their password whenever they want
          ' password expiry 90 days - prompt to change 2-3 weeks before expiry date
          ' consecutive passwords not allowed - password history between 5-8 (recommended 8)
          ' account lockout after 3 incorrect logons, contact system administrator to reactivate the account and reset the password.
          ' after account lockout the user will be prompted to change their password

          'All above from transfusion system
          'For NetAcquire and WardEnq - option to show/hide name

17670     If GetOptionSetting("LogOnShowName", "1") = "0" Then
17680         lblName.Visible = False
17690         cmbUserName.Visible = False
17700     End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

17710     If KeyAscii = Asc("'") Then
17720         KeyAscii = 0
17730         Beep
17740     End If

End Sub


