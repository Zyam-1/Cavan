VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmLogOn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   885
      Left            =   3030
      Picture         =   "frmLogOn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1470
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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
   Begin VB.Label Label1 
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

10    On Error GoTo FillUserNames_Error

20    cmbUserName.Clear

'30    sql = "SELECT Name FROM Users WHERE " & _
            "MemberOf = 'Administrators' " & _
            "OR MemberOf = 'Managers' " & _
            "OR MemberOf = 'Users' " & _
            "AND InUse = 1 " & _
            "ORDER BY ListOrder"
            
30    sql = "SELECT U.Name FROM Users as U, UserRole as UR  WHERE  U.InUse = 1 and U.MemberOf = UR.MemberOf  and UR.Systemrole = 'TransfusionAccess' and  UR.enabled = 1 ORDER BY ListOrder"
          
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70      cmbUserName.AddItem tb!Name & ""
80      tb.MoveNext
90    Loop

100   Exit Sub

FillUserNames_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmLogOn", "FillUserNames", intEL, strES, sql

End Sub

Private Sub cmbUserName_LostFocus()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmbUserName_LostFocus_Error

20    If cmbUserName = "" Then Exit Sub

'30    sql = "SELECT Name FROM Users WHERE " & _
            "(Code = '" & AddTicks(cmbUserName) & "' " & _
            "  OR Name = '" & AddTicks(cmbUserName) & "') " & _
            "AND (MemberOf = 'Administrators' " & _
            "  OR MemberOf = 'Managers' " & _
            "  OR MemberOf = 'Users') " & _
            "AND InUse = 1"
            
30        sql = "SELECT U.Name FROM Users as U, UserRole as UR  WHERE (U.Code = '" & AddTicks(cmbUserName) & "' " & _
            " OR U.Name = '" & AddTicks(cmbUserName) & "') and U.InUse = 1 and U.MemberOf = UR.MemberOf  " & _
            "and UR.Systemrole = 'TransfusionAccess' and  UR.enabled = 1 ORDER BY ListOrder"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      cmbUserName = tb!Name
80    Else
90      cmbUserName = ""
100   End If

110   Exit Sub

cmbUserName_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmLogOn", "cmbUserName_LostFocus", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim Diff As Long
      Static LockOutCounter As Integer

10    On Error GoTo cmdCancel_Click_Error

20    sql = "SELECT Code, MemberOf, LogOffDelay, " & _
            "COALESCE(PassDate, getdate()-1) AS PassDate FROM Users WHERE " & _
            "Name = '" & AddTicks(cmbUserName) & "' " & _
            "AND Password = '" & txtPassword & "' "
30    If GetOptionSetting("LogOnUpperLower", False) Then
40      sql = sql & "COLLATE SQL_Latin1_General_CP1_CS_AS"
50    End If
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    If tb.EOF Then
90      iMsg "Incorrect Password!", vbExclamation
100     txtPassword = ""
110     LockOutCounter = LockOutCounter + 1
120     If LockOutCounter > 2 Then
130       sql = "UPDATE Users SET InUse = 0 WHERE " & _
                "Name = '" & AddTicks(cmbUserName) & "' " & _
                "AND MemberOf <> 'Administrators'"
140       Cnxn(0).Execute sql
150       iMsg "Your account has been de-activated." & vbCrLf & _
               "Please contact your administrator.", vbInformation
160       LockOutCounter = 0
170       Unload Me
180     End If
190   Else
200     LockOutCounter = 0
210     If tb!MemberOf = "Administrators" Then
220       frmAdmin.AdminName = cmbUserName
230       cmbUserName = ""
240       txtPassword = ""
250       frmAdmin.Show 1
260       Unload Me
270       Exit Sub
280     Else
290       Diff = DateDiff("d", Now, tb!PassDate)
300       If Diff < 0 Then
310         iMsg "Your password has expired.", vbInformation
320         If TimedOut Then Unload Me: Exit Sub
330         UserName = ""
340         UserCode = ""
350         UserMemberOf = ""
360         Unload Me: Exit Sub
370       ElseIf Diff = 0 Then
380         iMsg "YOUR PASSWORD EXPIRES TODAY!", vbInformation
390         If TimedOut Then Unload Me: Exit Sub
400       ElseIf Diff < 14 Then
410         iMsg "Your password will expire " & IIf(Diff = 1, "tomorrow!", "in " & Diff & " days"), vbInformation
420         If TimedOut Then Unload Me: Exit Sub
430       End If
    
440       UserName = cmbUserName
450       UserCode = tb!code & ""
460       UserMemberOf = Trim$(tb!MemberOf & "")
470       If Not IsNull(tb!LogOffDelay) Then
480         LogOffDelayMin = tb!LogOffDelay
490       Else
500         LogOffDelayMin = 5
510       End If
520       LogOffDelaySecs = LogOffDelayMin * 60
530       If LogOffDelaySecs > 0 And LogOffDelaySecs <= 32767 Then
540         frmMain.pBar.max = LogOffDelaySecs
550       Else
560         frmMain.pBar.max = 1
570       End If
580       blnBTCdownWarningDisplayed = False
590       frmMain.timAnalyserHeartBeat.Enabled = True
600       Unload Me
    
610     End If
620   End If

630   Exit Sub

cmdCancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

640   intEL = Erl
650   strES = Err.Description
660   LogError "frmLogOn", "cmdCancel_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()

      Dim CurrentUser As String

10    CurrentUser = cmbUserName

20    FillUserNames

30    cmbUserName = CurrentUser

End Sub

Private Sub Form_Load()
                                                            '" unique usernames - policy length 6 characters
                                                            '" minimum length - 6 characters
                                                            '" alphanumeric
                                                            '" a user can change their password whenever they want
                                                            '" password expiry 90 days - prompt to change 2-3 weeks before expiry date
                                                            '" consecutive passwords not allowed - password history between 5-8 (recommended 8)
                                                            '" account lockout after 3 incorrect logons, contact system administrator to reactivate the account and reset the password.
                                                            '" after account lockout the user will be prompted to change their password

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

10    If KeyAscii = Asc("'") Then
20      KeyAscii = 0
30      Beep
40    End If

End Sub


