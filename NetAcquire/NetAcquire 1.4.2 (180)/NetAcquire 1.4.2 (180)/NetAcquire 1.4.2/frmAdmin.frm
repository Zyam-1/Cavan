VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSystemRights 
      Height          =   315
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   180
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CommandButton cmdUserRights 
      Caption         =   "System Rights"
      Height          =   375
      Left            =   5880
      TabIndex        =   34
      Top             =   120
      Width           =   2000
   End
   Begin VB.CommandButton cmdSystemRoles 
      Caption         =   "System Roles"
      Height          =   375
      Left            =   3750
      TabIndex        =   33
      Top             =   120
      Width           =   2000
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Administrator Password"
      Height          =   1200
      Left            =   9175
      Picture         =   "frmAdmin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2580
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "Local Policy"
      Height          =   1845
      Left            =   3750
      TabIndex        =   21
      Top             =   630
      Width           =   6525
      Begin VB.Frame Frame3 
         Caption         =   "Passwords Expire after"
         Height          =   885
         Left            =   3690
         TabIndex        =   26
         Top             =   180
         Width           =   2475
         Begin VB.OptionButton optNever 
            Caption         =   "Never"
            Height          =   195
            Left            =   1320
            TabIndex        =   30
            Top             =   510
            Width           =   735
         End
         Begin VB.OptionButton opt180 
            Caption         =   "180 Days"
            Height          =   195
            Left            =   1320
            TabIndex        =   29
            Top             =   270
            Width           =   975
         End
         Begin VB.OptionButton opt90 
            Alignment       =   1  'Right Justify
            Caption         =   "90 Days"
            Height          =   195
            Left            =   270
            TabIndex        =   28
            Top             =   510
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton opt60 
            Alignment       =   1  'Right Justify
            Caption         =   "60 Days"
            Height          =   195
            Left            =   270
            TabIndex        =   27
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
         TabIndex        =   24
         Top             =   1290
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   714
         _Version        =   393216
         Value           =   6
         BuddyControl    =   "lblLength"
         BuddyDispid     =   196626
         OrigLeft        =   2670
         OrigTop         =   1200
         OrigRight       =   2910
         OrigBottom      =   1695
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkShowName 
         Alignment       =   1  'Right Justify
         Caption         =   "Show User Name"
         Height          =   195
         Left            =   780
         TabIndex        =   9
         Top             =   1080
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
         TabIndex        =   32
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password can be Re-Used"
         Height          =   195
         Left            =   3720
         TabIndex        =   31
         Top             =   1380
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Minimum Password Length"
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   1380
         Width           =   1890
      End
      Begin VB.Label lblLength 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         Height          =   255
         Left            =   2070
         TabIndex        =   22
         Top             =   1350
         Width           =   330
      End
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9660
      Top             =   5040
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9660
      Top             =   5820
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   1000
      Left            =   9175
      Picture         =   "frmAdmin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7410
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   885
      Left            =   9060
      Picture         =   "frmAdmin.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5670
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
      Top             =   4770
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New User"
      Height          =   2445
      Left            =   150
      TabIndex        =   13
      Top             =   30
      Width           =   3555
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   800
         Left            =   2430
         Picture         =   "frmAdmin.frx":1330
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1000
      End
      Begin VB.TextBox txtCode 
         DataField       =   "opcode"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   150
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1980
         Width           =   675
      End
      Begin VB.TextBox txtName 
         DataField       =   "opname"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   150
         MaxLength       =   20
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cmbMemberOf 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1260
         Width           =   2205
      End
      Begin VB.TextBox txtAutoLogOff 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "5"
         Top             =   1980
         Width           =   465
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   1980
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   503
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "txtAutoLogOff"
         BuddyDispid     =   196637
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
         TabIndex        =   20
         Top             =   1740
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Member Of"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Auto Log Off"
         Height          =   195
         Left            =   1470
         TabIndex        =   17
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
         Height          =   195
         Left            =   2430
         TabIndex        =   16
         Top             =   2025
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Log Off Administrator"
      Height          =   1000
      Left            =   9175
      Picture         =   "frmAdmin.frx":199A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8490
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6945
      Left            =   150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2580
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

42860     MinLen = GetOptionSetting("LogOnMinPassLen", "1")

42870     If MinLen < 6 Then
42880         MinLen = 6
42890     End If

42900     Randomize

42910     s = Chr$(Int((Asc("Z") - Asc("A") + 1) * Rnd + Asc("A")))
42920     For n = 2 To MinLen - 1
42930         s = s & Chr$(Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a")))
42940     Next
42950     s = s & Chr$(Int((Asc("9") - Asc("0") + 1) * Rnd + Asc("0")))

42960     GeneratePassword = s

End Function

Private Function GetAdminCount() As Integer

          Dim ySave As Integer
          Dim Y As Integer
          Dim Counter As Integer

42970     ySave = g.row

42980     Counter = 0
42990     For Y = 1 To g.Rows - 1
43000         If g.TextMatrix(Y, 3) = "Administrators" Then
43010             Counter = Counter + 1
43020         End If
43030     Next

43040     g.row = ySave
43050     GetAdminCount = Counter

End Function

Private Sub LoadLocalPolicy()

          Dim Alpha As Boolean
          Dim PasswordExpiry As String

43060     On Error GoTo LoadLocalPolicy_Error

43070     Alpha = GetOptionSetting("LogOnAlpha", False)
43080     If Alpha Then
43090         chkAlpha.Value = 1
43100         chkUpperLower.Enabled = True
43110         chkUpperLower.Value = IIf(GetOptionSetting("LogOnUpperLower", False), 1, 0)
43120     Else
43130         chkAlpha.Value = 0
43140         chkUpperLower.Enabled = False
43150         chkUpperLower.Value = 0
43160     End If

43170     chkNumeric.Value = IIf(GetOptionSetting("LogOnNumeric", False), 1, 0)

43180     chkShowName.Value = IIf(GetOptionSetting("LogOnShowName", "1"), 1, 0)

43190     lblLength = GetOptionSetting("LogOnMinPassLen", "1")

43200     PasswordExpiry = GetOptionSetting("PasswordExpiry", "90")
43210     Select Case PasswordExpiry
              Case "60": opt60.Value = True
43220         Case "90": opt90.Value = True
43230         Case "180": opt180.Value = True
43240         Case "365": optNever.Value = True
43250     End Select

43260     lblReUse.Caption = GetOptionSetting("PasswordReUse", "No")

43270     Exit Sub

LoadLocalPolicy_Error:

          Dim strES As String
          Dim intEL As Integer

43280     intEL = Erl
43290     strES = Err.Description
43300     LogError "frmAdmin", "LoadLocalPolicy", intEL, strES

End Sub

Private Sub chkAlpha_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

43310     SaveOptionSetting "LogOnAlpha", chkAlpha.Value

43320     If chkAlpha.Value = False Then
43330         chkUpperLower.Enabled = False
43340         chkUpperLower.Value = 0
43350         SaveOptionSetting "LogOnUpperLower", 0
43360     Else
43370         chkUpperLower.Enabled = True
43380     End If

End Sub


Private Sub chkNumeric_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

43390     SaveOptionSetting "LogOnNumeric", chkNumeric.Value

End Sub


Private Sub chkShowName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

43400     SaveOptionSetting "LogOnShowName", chkShowName.Value

End Sub


Private Sub chkUpperLower_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

43410     SaveOptionSetting "LogOnUpperLower", chkUpperLower.Value

End Sub


Private Sub cmdAdd_Click()

          Dim sql As String
          Dim Password As String

43420     On Error GoTo cmdAdd_Click_Error

43430     txtName = Trim$(txtName)
43440     If txtName = "" Then
43450         iMsg "Enter Name of new user", vbExclamation
43460         Exit Sub
43470     End If

43480     txtCode = Trim$(txtCode)
43490     If txtCode = "" Then
43500         iMsg "Enter Code of new user", vbExclamation
43510         Exit Sub
43520     End If

43530     If NameHasBeenUsed(txtName) Then
43540         iMsg "Name has been used!", vbExclamation
43550         Exit Sub
43560     End If

43570     If CodeHasBeenUsed(txtCode) Then
43580         iMsg "Code has been used!", vbExclamation
43590         Exit Sub
43600     End If

43610     Password = GeneratePassword()

43620     sql = "INSERT INTO Users (Password, Name, Code, InUse, MemberOf, LogOffDelay, ListOrder, PassDate) " & _
              "VALUES ( " & _
              "'" & Password & "', " & _
              "'" & AddTicks(txtName) & "', " & _
              "'" & txtCode & "', " & _
              "'1', " & _
              "'" & cmbMemberOf & "', " & _
              "'" & txtAutoLogOff & "', " & _
              "'1', " & _
              "'" & Format$(Now, "dd/MMM/yyyy") & "')"
43630     Cnxn(0).Execute sql

43640     FillG

43650     iMsg "Password assigned to" & vbCrLf & txtName & vbCrLf & Password, vbInformation, , , 12

43660     Exit Sub

cmdAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

43670     intEL = Erl
43680     strES = Err.Description
43690     LogError "frmAdmin", "cmdAdd_Click", intEL, strES, sql

End Sub

Public Function CodeHasBeenUsed(ByVal Code As String) As Boolean

          Dim sql As String
          Dim tb As Recordset

43700     On Error GoTo CodeHasBeenUsed_Error

43710     CodeHasBeenUsed = False

43720     sql = "SELECT Code FROM Users WHERE " & _
              "Code = '" & Code & "'"
43730     Set tb = New Recordset
43740     RecOpenServer 0, tb, sql
43750     If Not tb.EOF Then
43760         CodeHasBeenUsed = True
43770     End If

43780     Exit Function

CodeHasBeenUsed_Error:

          Dim strES As String
          Dim intEL As Integer

43790     intEL = Erl
43800     strES = Err.Description
43810     LogError "frmAdmin", "CodeHasBeenUsed", intEL, strES, sql

End Function




Private Sub cmdCancel_Click()

43820     Unload Me

End Sub


Private Sub cmdChangePassword_Click()

          Dim NewPass As String
          Dim Confirm As String
          Dim tb As Recordset
          Dim sql As String
          Dim MinLength As Integer

43830     On Error GoTo cmdChangePassword_Click_Error

43840     NewPass = iBOX("Enter new password", , , True)
43850     Confirm = iBOX("Confirm password", , , True)

43860     If NewPass <> Confirm Then
43870         iMsg "Passwords don't match!", vbExclamation
43880         Exit Sub
43890     End If

43900     MinLength = Val(GetOptionSetting("LogOnMinPassLen", "1"))
43910     If Len(NewPass) < MinLength Then
43920         iMsg "Passwords must have a minimum of " & Format(MinLength) & " characters!", vbExclamation
43930         Exit Sub
43940     End If

43950     If GetOptionSetting("LogOnUpperLower", False) Then
43960         If AllLowerCase(NewPass) Or AllUpperCase(NewPass) Then
43970             iMsg "Passwords must have a mixture of UPPER CASE and lower case letters!", vbExclamation
43980             Exit Sub
43990         End If
44000     End If

44010     If GetOptionSetting("LogOnNumeric", False) Then
44020         If Not ContainsNumeric(NewPass) Then
44030             iMsg "Passwords must contain a numeric character!", vbExclamation
44040             Exit Sub
44050         End If
44060     End If

44070     If GetOptionSetting("LogOnAlpha", False) Then
44080         If Not ContainsAlpha(NewPass) Then
44090             iMsg "Passwords must contain an alphabetic character!", vbExclamation
44100             Exit Sub
44110         End If
44120     End If

44130     If PasswordHasBeenUsed(NewPass) Then
44140         iMsg "Password has been used!", vbExclamation
44150         Exit Sub
44160     End If

44170     Set tb = New Recordset
44180     RecOpenServer 0, tb, sql
44190     If Not tb.EOF Then
44200         sql = "UPDATE Users SET " & _
                  "PassWord = '" & NewPass & "' WHERE " & _
                  "Name = '" & AddTicks(pAdminName) & "'"
44210         Cnxn(0).Execute sql

44220         iMsg "Your Password has been changed.", vbInformation

44230     End If

44240     Exit Sub

cmdChangePassword_Click_Error:

          Dim strES As String
          Dim intEL As Integer

44250     intEL = Erl
44260     strES = Err.Description
44270     LogError "frmAdmin", "cmdChangePassword_Click", intEL, strES, sql

End Sub


Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

44280     FireDown

44290     tmrDown.Interval = 250
44300     FireCounter = 0

44310     tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

44320     tmrDown.Enabled = False

End Sub


Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

44330     If g.row = g.Rows - 1 Then Exit Sub
44340     n = g.row

44350     FireCounter = FireCounter + 1
44360     If FireCounter > 5 Then
44370         tmrDown.Interval = 100
44380     End If

44390     VisibleRows = g.height \ g.RowHeight(1) - 1

44400     g.Visible = False

44410     s = ""
44420     For X = 0 To g.Cols - 1
44430         s = s & g.TextMatrix(n, X) & vbTab
44440     Next
44450     s = Left$(s, Len(s) - 1)

44460     g.RemoveItem n
44470     If n < g.Rows Then
44480         g.AddItem s, n + 1
44490         g.row = n + 1
44500     Else
44510         g.AddItem s
44520         g.row = g.Rows - 1
44530     End If

44540     For X = 0 To g.Cols - 1
44550         g.Col = X
44560         g.CellBackColor = vbYellow
44570     Next

44580     If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
44590         If g.row - VisibleRows + 1 > 0 Then
44600             g.TopRow = g.row - VisibleRows + 1
44610         End If
44620     End If

44630     g.Visible = True

44640     cmdSave.Visible = True

End Sub
Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

44650     If g.row = 1 Then Exit Sub

44660     FireCounter = FireCounter + 1
44670     If FireCounter > 5 Then
44680         tmrUp.Interval = 100
44690     End If

44700     n = g.row

44710     g.Visible = False

44720     s = ""
44730     For X = 0 To g.Cols - 1
44740         s = s & g.TextMatrix(n, X) & vbTab
44750     Next
44760     s = Left$(s, Len(s) - 1)

44770     g.RemoveItem n
44780     g.AddItem s, n - 1

44790     g.row = n - 1
44800     For X = 0 To g.Cols - 1
44810         g.Col = X
44820         g.CellBackColor = vbYellow
44830     Next

44840     If Not g.RowIsVisible(g.row) Then
44850         g.TopRow = g.row
44860     End If

44870     g.Visible = True

44880     cmdSave.Visible = True

End Sub





Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

44890     FireUp

44900     tmrUp.Interval = 250
44910     FireCounter = 0

44920     tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

44930     tmrUp.Enabled = False

End Sub


Private Sub cmdSave_Click()

          Dim sql As String
          Dim Y As Integer

44940     On Error GoTo cmdSave_Click_Error

44950     For Y = 1 To g.Rows - 1
44960         sql = "UPDATE Users " & _
                  "SET ListOrder = " & Y & " WHERE " & _
                  "Name = '" & AddTicks(g.TextMatrix(Y, 1)) & "' " & _
                  "AND Password = '" & g.TextMatrix(Y, 6) & "' " & _
                  "COLLATE SQL_Latin1_General_CP1_CS_AS"
44970         Cnxn(0).Execute sql
44980     Next
44990     cmdSave.Visible = False

45000     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

45010     intEL = Erl
45020     strES = Err.Description
45030     LogError "frmAdmin", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmdSystemRoles_Click()

          Dim i As Integer
          Dim J As Integer
          Dim URole As UserRole

45040     On Error GoTo cmdSystemRoles_Click_Error

45050     With frmListsGeneric
45060         .ListType = "SR"
45070         .ListTypeName = "System Role"
45080         .ListTypeNames = "System Roles"
45090         .Show 1

45100     End With
45110     FillGenericList cmbMemberOf, "SR"
45120     FillGenericList cmbSystemRights, "SystemRights"

45130     For i = 0 To cmbMemberOf.ListCount - 1
45140         For J = 0 To cmbSystemRights.ListCount - 1
45150             Set URole = New UserRole
45160             With URole
45170                 If .GetUserRole(cmbMemberOf.List(i), cmbSystemRights.List(J), UserName) = False Then
45180                     .MemberOf = cmbMemberOf.List(i)
45190                     .SystemRole = cmbSystemRights.List(J)
45200                     .Description = "Grants access permission to " & cmbSystemRights.List(J)
45210                     .Enabled = 1

45220                     .Add ("Administrator")
45230                 End If
45240             End With
45250         Next J
45260     Next i

45270     Exit Sub

cmdSystemRoles_Click_Error:

          Dim strES As String
          Dim intEL As Integer

45280     intEL = Erl
45290     strES = Err.Description
45300     LogError "frmAdmin", "cmdSystemRoles_Click", intEL, strES

End Sub

Private Sub cmdUserRights_Click()

45310     On Error GoTo cmdUserRights_Click_Error

45320     frmUserRoles.Show 1

45330     Exit Sub

cmdUserRights_Click_Error:

          Dim strES As String
          Dim intEL As Integer

45340     intEL = Erl
45350     strES = Err.Description
45360     LogError "frmAdmin", "cmdUserRights_Click", intEL, strES

End Sub

Private Sub Form_Load()

45370     FillG
45380     LoadLocalPolicy
45390     FillGenericList cmbMemberOf, "SR"

          'cmbMemberOf.Clear
          'cmbMemberOf.AddItem "Administrators"
          'cmbMemberOf.AddItem "Managers"
          'cmbMemberOf.AddItem "Users"
          'cmbMemberOf.AddItem "LookUp"

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

45400     On Error GoTo FillG_Error

45410     g.Rows = 2
45420     g.AddItem ""
45430     g.RemoveItem 1

45440     sql = "SELECT * FROM Users WHERE " & _
              "MemberOf <> 'LookUp' " & _
              "ORDER BY ListOrder"
          'sql = "SELECT * FROM Users WHERE " & _
          '      "MemberOf = 'Administrators' " & _
          '      "OR MemberOf = 'Managers' " & _
          '      "OR MemberOf = 'Users' " & _
          '      "ORDER BY ListOrder"

45450     Set tb = New Recordset
45460     RecOpenClient 0, tb, sql
45470     With tb
45480         Do While Not .EOF
45490             s = IIf(!InUse, "Yes", "No") & vbTab & _
                      !Name & vbTab & _
                      !Code & vbTab & _
                      !MemberOf & vbTab & _
                      !LogOffDelay & vbTab & _
                      "*****" & vbTab & _
                      !Password
45500             g.AddItem s
45510             .MoveNext
45520         Loop
45530     End With

45540     If g.Rows > 2 Then
45550         g.RemoveItem 1
45560     End If

45570     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

45580     intEL = Erl
45590     strES = Err.Description
45600     LogError "frmAdmin", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Unload(Cancel As Integer)

          Dim Answer As Integer

45610     If cmdSave.Visible Then
45620         Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
              '30      If TimedOut Then Unload Me: Exit Sub
45630         If Answer = vbNo Then
45640             Cancel = True
45650         End If
45660     End If

End Sub

Private Sub g_Click()

          Static SortOrder As Boolean
          Dim LogOff As String
          Dim lngNewLogOff As Long
          Dim sql As String
          Dim MemberOf As String
          Dim f As Form
          Dim s() As String
          Dim AdminCount As Integer
          Dim NewPass As String
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer
          Dim xSave As Integer
          Dim i  As Integer

45670     On Error GoTo g_Click_Error

45680     If g.MouseRow = 0 Then
45690         If SortOrder Then
45700             g.Sort = flexSortGenericAscending
45710         Else
45720             g.Sort = flexSortGenericDescending
45730         End If
45740         SortOrder = Not SortOrder
45750         cmdSave.Visible = True
45760         Exit Sub
45770     End If

45780     cmdMoveUp.Enabled = False
45790     cmdMoveDown.Enabled = False
45800     ySave = g.row
45810     xSave = g.Col
45820     g.Visible = False
45830     g.Col = 0
45840     For Y = 1 To g.Rows - 1
45850         g.row = Y
45860         If g.CellBackColor = vbYellow Then
45870             For X = 0 To g.Cols - 1
45880                 g.Col = X
45890                 g.CellBackColor = 0
45900             Next
45910             Exit For
45920         End If
45930     Next
45940     g.row = ySave
45950     g.Col = xSave
45960     g.Visible = True

45970     Select Case g.Col
              Case 0:    'In Use
45980             If g.TextMatrix(g.row, 0) = "No" Then    'Mark as InUse
45990                 g.TextMatrix(g.row, 0) = "Yes"
46000                 NewPass = GeneratePassword
46010                 sql = "UPDATE Users " & _
                          "SET InUse = 1, " & _
                          "PassDate = '" & Format$(Now, "dd/MMM/yyyy") & "', " & _
                          "Password = '" & NewPass & "' " & _
                          "WHERE Name = '" & AddTicks(g.TextMatrix(g.row, 1)) & "' " & _
                          "AND Password = '" & g.TextMatrix(g.row, 6) & "' " & _
                          "COLLATE SQL_Latin1_General_CP1_CS_AS"
46020                 Cnxn(0).Execute sql
46030                 iMsg "Password has been changed for " & g.TextMatrix(g.row, 1) & vbCrLf & _
                          "New Password : " & NewPass, vbInformation
46040                 FillG
46050             Else
46060                 If g.TextMatrix(g.row, 3) = "Administrators" Then
46070                     AdminCount = GetAdminCount()
46080                     If AdminCount = 1 Then
46090                         iMsg "At least 1 Administrator must be In-Use", vbCritical
46100                         Exit Sub
46110                     Else
46120                         g.TextMatrix(g.row, 0) = "No"
46130                         sql = "UPDATE Users " & _
                                  "SET InUse = 0 " & _
                                  "WHERE Name = '" & AddTicks(g.TextMatrix(g.row, 1)) & "' " & _
                                  "AND Password = '" & g.TextMatrix(g.row, 6) & "' " & _
                                  "COLLATE SQL_Latin1_General_CP1_CS_AS"
46140                         Cnxn(0).Execute sql
46150                     End If
46160                 Else
46170                     g.TextMatrix(g.row, 0) = "No"
46180                     sql = "UPDATE Users " & _
                              "SET InUse = 0 " & _
                              "WHERE Name = '" & AddTicks(g.TextMatrix(g.row, 1)) & "' " & _
                              "AND Password = '" & g.TextMatrix(g.row, 6) & "' " & _
                              "COLLATE SQL_Latin1_General_CP1_CS_AS"
46190                     Cnxn(0).Execute sql
46200                 End If
46210             End If

46220         Case 1:    'Name
46230             For X = 0 To g.Cols - 1
46240                 g.Col = X
46250                 g.CellBackColor = vbYellow
46260             Next
46270             cmdMoveUp.Enabled = True
46280             cmdMoveDown.Enabled = True

46290         Case 3:    'MemberOf

46300             AdminCount = GetAdminCount()
46310             g.Enabled = False
46320             ReDim s(0 To cmbMemberOf.ListCount - 1) As String
46330             For i = 0 To cmbMemberOf.ListCount - 1
46340                 s(i) = cmbMemberOf.List(i)
46350             Next i
            
46360             Set f = New fcdrDBox
46370             With f
46380                 .Options = s
46390                 .Prompt = "Enter Member Of Group."
46400                 .Show 1
46410                 MemberOf = .ReturnValue
46420             End With
46430             Set f = Nothing
46440             If MemberOf <> "" Then
46450                 If AdminCount = 1 And _
                          g.TextMatrix(g.row, 3) = "Administrators" And _
                          MemberOf <> "Administrators" Then
46460                     iMsg "Cannot demote Administrator."
46470                 Else
46480                     If g.TextMatrix(g.row, 3) <> MemberOf Then
46490                         g.TextMatrix(g.row, 3) = MemberOf
46500                         sql = "UPDATE Users " & _
                                  "SET MemberOf = '" & MemberOf & "' " & _
                                  "WHERE Name = '" & g.TextMatrix(g.row, 1) & "'"
46510                         Cnxn(0).Execute sql
46520                     End If
46530                 End If
46540             End If
46550             g.Enabled = True

46560         Case 4:    'Log Off Delay
46570             g.Enabled = False
46580             LogOff = g.TextMatrix(g.row, 4)
46590             lngNewLogOff = Val(iBOX("Log Off Delay. (Minutes)", , LogOff))
46600             If lngNewLogOff > 0 Then
46610                 g.TextMatrix(g.row, 4) = Format$(lngNewLogOff)
46620                 sql = "UPDATE Users " & _
                          "SET LogOffDelay = " & lngNewLogOff & " " & _
                          "WHERE Name = '" & g.TextMatrix(g.row, 1) & "'"
46630                 Cnxn(0).Execute sql
46640             End If

46650     End Select

46660     Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

46670     intEL = Erl
46680     strES = Err.Description
46690     LogError "frmAdmin", "g_Click", intEL, strES, sql

End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

46700     If g.MouseRow > 0 And g.MouseCol = 5 Then
46710         g.ToolTipText = "Password:" & g.TextMatrix(g.MouseRow, 6)
46720         Exit Sub
46730     ElseIf g.MouseRow > 0 And g.MouseCol = 0 Then
46740         g.ToolTipText = "Click to Toggle Yes/No"
46750         Exit Sub
46760     Else
46770         g.ToolTipText = ""
46780     End If

End Sub


Private Sub lblReUse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

46790     lblReUse.Caption = IIf(lblReUse.Caption = "No", "Yes", "No")

46800     SaveOptionSetting "PasswordReUse", lblReUse.Caption

End Sub


Private Sub opt180_Click()

46810     SaveOptionSetting "PasswordExpiry", "180"

End Sub

Private Sub opt60_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

46820     SaveOptionSetting "PasswordExpiry", "60"

End Sub


Private Sub opt90_Click()

46830     SaveOptionSetting "PasswordExpiry", "90"

End Sub


Private Sub optNever_Click()

46840     SaveOptionSetting "PasswordExpiry", "365"    '1 year

End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)

46850     If KeyAscii = Asc("'") Then
46860         KeyAscii = 0
46870         Beep
46880     End If

End Sub

Private Sub tmrDown_Timer()

46890     FireDown

End Sub


Private Sub tmrUp_Timer()

46900     FireUp

End Sub


Private Sub udLength_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

46910     SaveOptionSetting "LogOnMinPassLen", lblLength.Caption

End Sub



Public Property Let AdminName(ByVal sNewValue As String)

46920     pAdminName = sNewValue

End Property


