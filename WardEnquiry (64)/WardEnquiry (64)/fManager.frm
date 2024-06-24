VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form fManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Laboratory Result Look Up"
   ClientHeight    =   7635
   ClientLeft      =   1350
   ClientTop       =   1650
   ClientWidth     =   9510
   ClipControls    =   0   'False
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
   Icon            =   "fManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7635
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbRoleList 
      Height          =   315
      Left            =   810
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   4620
      Picture         =   "fManager.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   150
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O. K."
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
      Height          =   1065
      Left            =   3450
      Picture         =   "fManager.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   1005
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   840
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   510
      Width           =   2535
   End
   Begin VB.ComboBox cmbUserName 
      Height          =   315
      Left            =   8010
      TabIndex        =   1
      Text            =   "cmbUserName"
      Top             =   660
      Visible         =   0   'False
      Width           =   2235
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3495
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   8
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
      AllowUserResizing=   1
      FormatString    =   "<In Use |<Operator Name                 |<Code          |<Member Of |^Ward Print |^Log Off Delay |<Password   |"
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   7080
      Picture         =   "fManager.frx":2C5E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1890
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Operator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   5565
      Begin VB.TextBox txtAutoLogOff 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         TabIndex        =   20
         Text            =   "5"
         Top             =   2100
         Width           =   495
      End
      Begin VB.ComboBox cmbMemberOf 
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
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1650
         Width           =   2205
      End
      Begin VB.TextBox tConfirm 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1650
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox tName 
         DataField       =   "opname"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   3
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox tCode 
         DataField       =   "opcode"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   4
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox tPass 
         DataField       =   "oppass"
         DataSource      =   "Data1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1650
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   870
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   4350
         Picture         =   "fManager.frx":3B28
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   855
         Width           =   1005
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2145
         TabIndex        =   21
         Top             =   2100
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "txtAutoLogOff"
         BuddyDispid     =   196616
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
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
         Left            =   2640
         TabIndex        =   22
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Auto Log Off in"
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
         Left            =   480
         TabIndex        =   19
         Top             =   2130
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Member Of"
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
         Left            =   750
         TabIndex        =   18
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
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
         Left            =   270
         TabIndex        =   15
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Left            =   1110
         TabIndex        =   13
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   840
         TabIndex        =   12
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
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
         TabIndex        =   11
         Top             =   330
         Width           =   375
      End
   End
   Begin VB.Label lblUserRole 
      Caption         =   "Please Select a Profile"
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   210
      Picture         =   "fManager.frx":49F2
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "fManager.frx":4CC8
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Password"
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
      Left            =   90
      TabIndex        =   17
      Top             =   540
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Last Log On"
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
      Left            =   8610
      TabIndex        =   16
      Top             =   450
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "fManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Temp, Pass As String

Private mLookUp As Boolean
Private mOperator As Boolean
Private mManager As Boolean
Private mAdministrator As Boolean

Private blnRefreshOnActivate As Boolean

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim S As String

10    g.Rows = 2
20    g.AddItem ""
30    g.RemoveItem 1

40    cmbUserName.Clear

50    sql = "SELECT TOP 20 Password, Name, Code, InUse, MemberOf, LogOffDelay, " & _
            "COALESCE(Prints, 0) Prints FROM Users " & _
            "WHERE InUse IS NOT NULL " & _
            "ORDER BY ListOrder"
60    Set tb = New Recordset
70    RecOpenClient 0, tb, sql
80    With tb
90        Do While Not .EOF

100           If !InUse Then
110               If mLookUp And Trim$(!MemberOf) = "LookUp" Or _
                     mOperator And Trim$(!MemberOf) = "Users" Or _
                     mOperator And Trim$(!MemberOf) = "User" Or _
                     mOperator And Trim$(!MemberOf) = "Operators" Or _
                     mOperator And Trim$(!MemberOf) = "Operator" Or _
                     mManager And Trim$(!MemberOf) = "Managers" Or _
                     mManager And Trim$(!MemberOf) = "Manager" Or _
                     mAdministrator And Trim$(!MemberOf) = "Administrators" Or _
                     mAdministrator And Trim$(!MemberOf) = "Administrator" Then
120                   cmbUserName.AddItem !Name
130               End If
140           End If

150           S = IIf(!InUse, "Yes", "No") & vbTab & _
                  !Name & vbTab & _
                  !Code & vbTab & _
                  !MemberOf & vbTab & vbTab & _
                  !LogOffDelay & vbTab & _
                  "*****" & vbTab & _
                  !Password
160           g.AddItem S
170           If !MemberOf = "LookUp" Then
180               g.Row = g.Rows - 1
190               g.col = 4
200               g.CellPictureAlignment = flexAlignCenterCenter
210               If !Prints Then
220                   Set g.CellPicture = imgGreenTick.Picture
230               Else
240                   Set g.CellPicture = imgRedCross.Picture
250               End If
260           End If
270           .MoveNext
280       Loop
290   End With

300   If g.Rows > 2 Then
310       g.RemoveItem 1
320   End If

End Sub

Private Sub cmbRoleList_Click()

          Dim sql As String
          Dim Aa As String

10        Aa = MsgBox("You have selected " & cmbRoleList.Text & " as your profile is this correct?", vbYesNo, "Question")
20        If Aa = vbYes Then
30            sql = "UPDATE Users " & _
                    "SET RoleName = '" & cmbRoleList.Text & "' " & _
                    "Where Password = '" & Pass & "' " & _
                    "and InUse = 1"
40            Cnxn(0).Execute sql
50            Pass = ""
60            cmdOK.Enabled = True
70        Else

80        End If

End Sub

Private Sub cmdCancel_Click()

      '10    FillG

10    cmbUserName.ListIndex = -1
20    txtPassword = ""
30    Me.Width = 6165
40    Me.Height = 1750

End Sub

Private Sub cmdHide_Click()

10    Unload Me

End Sub


Private Sub cmdOK_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Diff As Long

10        If cmbRoleList.Visible Then Unload Me

20        On Error GoTo cmdOK_Click_Error

30        sql = "Select * from Users where " & _
                "Password = '" & AddTicks(txtPassword) & "' " & _
                "and InUse = 1"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        If Not tb.EOF Then
70            Pass = AddTicks(txtPassword)
80            If IsNull(tb!RoleName) Or tb!RoleName = "" Then
90                FillRoles
100               cmbRoleList.Visible = True
110               cmdOK.Enabled = False
120               lblUserRole.Visible = True
                  'txtPassword.Enabled = False
130           Else
140               cmbUserName = tb!Name & ""
150           End If
160       End If


170       sql = "SELECT top 20 Password, Name, Code, InUse, MemberOf, RoleName, " & _
                "COALESCE(LogOffDelay, 5) LogOffDelayMin, " & _
                "COALESCE(Prints, 0) UserCanPrint, " & _
                "COALESCE(PassDate, getdate()-1) PassDate " & _
                "FROM Users WHERE " & _
                "Name = '" & AddTicks(cmbUserName) & "' " & _
                "AND Password = '" & AddTicks(txtPassword) & "' " & _
                "AND InUse = 1"

180       Set tb = New Recordset
190       RecOpenClient 0, tb, sql
200       If Not tb.EOF Then
210           If Trim$(tb!MemberOf & "") = "Administrators" Then
220               Screen.MousePointer = vbHourglass
230               FillG
240               Screen.MousePointer = vbDefault
250               Me.Width = 9600
260               Me.Height = 8040
270               Exit Sub
280           End If

290           Diff = DateDiff("d", Now, tb!PassDate)
300           If Diff < 0 Then
310               iMsg "Your password has expired.", vbInformation
320               UserName = ""
330               UserCode = ""
340               UserMemberOf = ""
350               UserRoleName = ""
360               Unload Me: Exit Sub
370           ElseIf Diff = 0 Then
380               iMsg vbCrLf & "YOUR PASSWORD EXPIRES TODAY!" & vbCrLf & vbCrLf & "YOU MUST CHANGE YOUR PASSWORD NOW!", vbInformation
390           ElseIf Diff < 14 Then
400               iMsg vbCrLf & "Your password will expire " & IIf(Diff = 1, "tomorrow!", "in " & Diff & " days") & vbCrLf & vbCrLf & "YOU MUST CHANGE YOUR PASSWORD NOW!", vbInformation
410           End If

420           UserName = cmbUserName
430           UserCanPrint = tb!UserCanPrint
440           'SaveOptionSetting "WardLogOnUserName", UserName, vbGetComputerName
450           UserCode = Trim$(tb!Code & "")
460           UserPass = Trim$(tb!Password & "")
470           UserMemberOf = Trim$(tb!MemberOf & "")
480           UserRoleName = Trim$(tb!RoleName & "")
490           LogOffDelayMin = tb!LogOffDelayMin
500           LogOffDelaySecs = LogOffDelayMin * 60
510           If LogOffDelaySecs > 0 And LogOffDelaySecs <= 32767 Then
520               frmMain.PBar.Max = LogOffDelaySecs
530           Else
540               frmMain.PBar.Max = 1
550           End If
              DoEvents
              DoEvents
              If m_RunExe = "OCM" Then
                m_FindChart = True
              End If
560           Unload Me
570       Else
580           txtPassword = ""
590       End If
            

600       Exit Sub

cmdOK_Click_Error:

          Dim strES As String
          Dim intEL As Integer

610       intEL = Erl
620       strES = Err.Description
630       LogError "fManager", "cmdOK_Click", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    tName = Trim$(tName)
30    tCode = UCase$(Trim$(tCode))
40    tPass = UCase$(Trim$(tPass))
50    tConfirm = UCase$(Trim$(tConfirm))

60    If Val(txtAutoLogOff) < 1 Then
70        txtAutoLogOff = "5"
80    End If

90    If tName = "" Or tCode = "" Or tPass = "" Then
100       iMsg "Must have Name, Code," & vbCrLf & "and Password.", vbCritical
110       Exit Sub
120   End If

130   If cmbMemberOf = "" Then
140       iMsg "Member Of ???", vbCritical
150       Exit Sub
160   End If

170   If tPass <> tConfirm Then
180       tPass = ""
190       tConfirm = ""
200       iMsg "Password/Confirm don't match." & vbCrLf & "Retype Password and Confirmation", vbCritical
210       Exit Sub
220   End If

230   sql = "Select * from Users where " & _
            "Name = '" & AddTicks(tName) & "'"
240   Set tb = New Recordset
250   RecOpenServer 0, tb, sql
260   If Not tb.EOF Then
270       iMsg "Name already used.", vbExclamation
280       tName = ""
290       Exit Sub
300   End If

310   sql = "Select * from Users where " & _
            "Code = '" & AddTicks(tCode) & "'"
320   Set tb = New Recordset
330   RecOpenServer 0, tb, sql
340   If Not tb.EOF Then
350       iMsg "Code already used.", vbExclamation
360       tName = ""
370       Exit Sub
380   End If

390   sql = "Select * from Users where " & _
            "Password = '" & AddTicks(tPass) & "'"
400   Set tb = New Recordset
410   RecOpenServer 0, tb, sql
420   If Not tb.EOF Then
430       iMsg "Password already used.", vbExclamation
440       tName = ""
450       Exit Sub
460   End If

470   sql = "INSERT INTO Users " & _
            "(PassWord, Name, Code, InUse, MemberOf, LogOffDelay, ListOrder, Prints, PassDate, TransUser, ExpiryDate, LogOffDelaySec, RoleName) " & _
            "VALUES " & _
            "('" & tPass & "', " & _
            "'" & tName & "', " & _
            "'" & tCode & "', " & _
            "'1', " & _
            "'" & cmbMemberOf & "', " & _
            "'" & Val(txtAutoLogOff) & "', " & _
            "'99', " & _
            "'0', " & _
            "'" & Format$(Now, "dd/MMM/yyyy") & "', " & _
            "'0' ," & _
            "'" & Format$(Now, "dd/MMM/yyyy") & "', " & _
            "'" & Val(txtAutoLogOff) * 60 & "', " & _
            "'')"
480   Cnxn(0).Execute sql

490   FillG

500   tCode = ""
510   tName = ""
520   tPass = ""
530   tConfirm = ""
540   txtAutoLogOff = "5"

550   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

560   intEL = Erl
570   strES = Err.Description
580   LogError "fManager", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmbUserName_Click()

'txtPassword = ""

End Sub


Private Sub Form_Activate()

      Dim n As Integer
      Dim TempUser As String

10    If Not blnRefreshOnActivate Then
20        blnRefreshOnActivate = True
30        Exit Sub
40    End If

'50    TempUser = GetOptionSetting("WardLogOnUserName", "", vbGetComputerName)
'
'60    For n = 0 To cmbUserName.ListCount - 1
'70        If cmbUserName.List(n) = TempUser Then
'80            cmbUserName.ListIndex = n
'90            Exit For
'100       End If
'110   Next

120   If cmbUserName <> "Administrator" Then
130       Me.Width = 6165
140       Me.Height = 1750
150   End If

160   txtPassword = ""
    DoEvents
    DoEvents
    If m_RunExe = "OCM" Then
        txtPassword.Text = m_Pass
        DoEvents
        DoEvents
        Call txtPassword_LostFocus
        Call cmdOK_Click
    End If

End Sub

Private Sub Form_Load()

10    Me.Width = 5715
20    Me.Height = 1750

30    g.ColWidth(7) = 0

40    txtPassword = ""

50    With cmbMemberOf
60        .Clear
70        .AddItem "LookUp"
80        .AddItem "Users"
90        .AddItem "Managers"
100       .AddItem "Administrators"
110   End With

      '120   FillG

120   blnRefreshOnActivate = True

End Sub

Private Sub g_Click()

      Static SortOrder As Boolean
      Dim LogOff As String
      Dim lngNewLogOff As Long
      Dim sql As String
      Dim gy As Long

10    If g.MouseRow = 0 Then
20        If SortOrder Then
30            g.Sort = flexSortGenericAscending
40        Else
50            g.Sort = flexSortGenericDescending
60        End If
70        SortOrder = Not SortOrder
80        For gy = 1 To g.Rows - 1
90            sql = "Update Users " & _
                    "Set ListOrder = '" & gy & "' " & _
                    "where Password = '" & g.TextMatrix(gy, 6) & "'"
100           Cnxn(0).Execute sql
110       Next
120       Exit Sub
130   End If

140   Select Case g.col
          Case 0:    'In Use
150           If g.TextMatrix(g.Row, 1) <> "Administrator" Then
160               g.TextMatrix(g.Row, 0) = IIf(g.TextMatrix(g.Row, 0) = "No", "Yes", "No")
170               sql = "Update Users " & _
                        "Set InUse = '" & IIf(g.TextMatrix(g.Row, 0) = "No", 0, 1) & "' " & _
                        "where Code = '" & g.TextMatrix(g.Row, 2) & "'"
180               Cnxn(0).Execute sql
190           End If
200       Case 4:    'Ward Print
210           g.Row = g.MouseRow
220           If g.CellPicture = imgGreenTick.Picture Then
230               Set g.CellPicture = imgRedCross.Picture
240               sql = "UPDATE Users SET Prints = 0 " & _
                        "WHERE Code = '" & g.TextMatrix(g.Row, 2) & "'"
250               Cnxn(0).Execute sql
260           ElseIf g.CellPicture = imgRedCross.Picture Then
270               Set g.CellPicture = imgGreenTick.Picture
280               sql = "UPDATE Users SET Prints = 1 " & _
                        "WHERE Code = '" & g.TextMatrix(g.Row, 2) & "'"
290               Cnxn(0).Execute sql
300           End If

310       Case 5:    'Log Off Delay
320           g.Enabled = False
330           LogOff = g.TextMatrix(g.Row, 5)
340           lngNewLogOff = Val(iBOX("Log Off Delay. (Minutes)", , LogOff))
350           If lngNewLogOff > 0 Then
360               g.TextMatrix(g.Row, 5) = Format$(lngNewLogOff)
370               sql = "Update Users " & _
                        "Set LogOffDelay = " & lngNewLogOff & " " & _
                        "where Code = '" & g.TextMatrix(g.Row, 2) & "'"
380               Cnxn(0).Execute sql
390           End If
400           g.Enabled = True
410   End Select

420   blnRefreshOnActivate = False

End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

      Dim n As Integer
      Static PrevY As Integer

10    If g.MouseRow > 0 And g.MouseCol = 6 Then
20        g.ToolTipText = g.TextMatrix(g.MouseRow, 7)
30        Exit Sub
40    ElseIf g.MouseRow > 0 And g.MouseCol = 0 Then
50        g.ToolTipText = "Click to Toggle Yes/No"
60        Exit Sub
70    ElseIf g.MouseRow > 0 And g.MouseCol = 1 Then
80        g.ToolTipText = "Drag to change List Order"
90    Else
100       g.ToolTipText = ""
110   End If
120   If Button = vbLeftButton And g.MouseRow > 0 And g.MouseCol = 1 Then
130       If Temp = "" Then
140           PrevY = g.MouseRow
150           For n = 0 To g.Cols - 1
160               Temp = Temp & g.TextMatrix(g.Row, n) & vbTab
170           Next
180           Temp = Left$(Temp, Len(Temp) - 1)
190           Exit Sub
200       Else
210           If g.MouseRow <> PrevY Then
220               g.RemoveItem PrevY
230               If g.MouseRow <> PrevY Then
240                   g.AddItem Temp, g.MouseRow
250                   PrevY = g.MouseRow
260               Else
270                   g.AddItem Temp
280                   PrevY = g.Rows - 1
290               End If
300           End If
310       End If
320   End If

End Sub

Private Sub g_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

      Dim gy As Long
      Dim sql As String

10    For gy = 1 To g.Rows - 1
20        sql = "Update Users " & _
                "Set ListOrder = '" & gy & "' " & _
                "where Password = '" & g.TextMatrix(gy, 6) & "'"
30        Cnxn(0).Execute sql
40    Next

50    Temp = ""

End Sub


Private Sub txtPassword_LostFocus()

10    txtPassword = UCase$(txtPassword)

End Sub



Public Property Let LookUp(ByVal ShowLookUp As Boolean)

10    mLookUp = ShowLookUp

End Property
Public Property Let Operator(ByVal ShowOperator As Boolean)

10    mOperator = ShowOperator

End Property

Public Property Let Manager(ByVal ShowManager As Boolean)

10    mManager = ShowManager

End Property


Public Property Let Administrator(ByVal ShowAdministrator As Boolean)

10    mAdministrator = ShowAdministrator

End Property

Private Sub FillRoles()

    Dim tl As Recordset
    Dim sql As String
    Dim S As String

    On Error GoTo FillRoles_Error

    cmbRoleList.Clear

    sql = "Select * from Lists where " & _
          "ListType = 'RL' " & _
          "and InUse = 1 order by ListOrder"
    Set tl = New Recordset
    RecOpenClient 0, tl, sql
    
    Do While Not tl.EOF
        With tl
            cmbRoleList.AddItem !Text
        End With
        tl.MoveNext
    Loop

    Exit Sub

FillRoles_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "fManager", "FillRoles", intEL, strES, sql

End Sub

