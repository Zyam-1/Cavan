VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActivityLog 
   Caption         =   "Activity Log Viewer"
   ClientHeight    =   12030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   12030
   ScaleWidth      =   19875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraAdvanceFind 
      Height          =   5415
      Left            =   7335
      TabIndex        =   8
      Top             =   3360
      Width           =   6435
      Begin VB.ComboBox cmbAction 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1680
         Width           =   4920
      End
      Begin VB.CheckBox chkToDate 
         Caption         =   "To Date"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3660
         Width           =   975
      End
      Begin VB.ComboBox cmbAppVersion 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   4440
         Width           =   4920
      End
      Begin VB.ComboBox cmbMachine 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   4080
         Width           =   4920
      End
      Begin VB.CheckBox chkFromDate 
         Caption         =   "For Date"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3180
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtDate1 
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   218365953
         CurrentDate     =   43339
      End
      Begin VB.ComboBox cmbUserName 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2760
         Width           =   4920
      End
      Begin VB.TextBox txtNotes 
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   2400
         Width           =   4920
      End
      Begin VB.ComboBox cmbActionType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   4920
      End
      Begin VB.TextBox txtReason 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   2040
         Width           =   4920
      End
      Begin VB.TextBox txtPatientID 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtSubmissionID 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtSampleID 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdFindAdvanceFind 
         Caption         =   "Find"
         Height          =   375
         Left            =   4605
         TabIndex        =   10
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelAdvanceFind 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5505
         TabIndex        =   9
         Top             =   4920
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtDate2 
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   218365953
         CurrentDate     =   43339
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Action"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   1740
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "App. Version"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   4500
         Width           =   900
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Machine"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   4140
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   2820
         Width           =   795
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Notes"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   2445
         Width           =   420
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Action Type"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Reason"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   2085
         Width           =   555
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Patient ID"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Submission ID"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   1005
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.Frame fraMain 
      Height          =   11895
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   19725
      Begin VB.CommandButton cmdExport 
         Height          =   885
         Left            =   180
         Picture         =   "frmActivityLog.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   10560
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Print"
         DisabledPicture =   "frmActivityLog.frx":0ECA
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   16500
         Picture         =   "frmActivityLog.frx":1D94
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "bprint"
         Top             =   10560
         Width           =   1200
      End
      Begin VB.CommandButton cmdAdvanceFind 
         Caption         =   "Advance Find"
         Height          =   375
         Left            =   9600
         TabIndex        =   7
         Top             =   315
         Width           =   1695
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   8640
         TabIndex        =   5
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Default         =   -1  'True
         Height          =   375
         Left            =   7560
         TabIndex        =   4
         Top             =   315
         Width           =   975
      End
      Begin VB.TextBox txtFind 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   3
         Top             =   300
         Width           =   6255
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   18465
         Picture         =   "frmActivityLog.frx":24C5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   10560
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   9615
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   19545
         _ExtentX        =   34475
         _ExtentY        =   16960
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
         AllowUserResizing=   1
      End
      Begin VB.Label lblExcelInfo 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Exporting..."
         Height          =   285
         Left            =   180
         TabIndex        =   38
         Top             =   11475
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Find Log"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   405
         Width           =   615
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   19695
      _ExtentX        =   34740
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmActivityLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkFromDate_Click()
36260     If chkFromDate.Value = vbChecked Then
36270         dtDate1.Enabled = True
36280         chkToDate.Enabled = True
36290         dtDate2.Enabled = True
36300     Else
36310         dtDate1.Enabled = False
36320         chkToDate.Enabled = False
36330         dtDate1.Enabled = False
36340     End If
End Sub

Private Sub chkToDate_Click()
36350     If chkToDate.Value = vbChecked Then
36360         dtDate2.Enabled = True
36370         chkFromDate.Caption = "From Date"
36380     Else
36390         dtDate2.Enabled = False
36400         chkFromDate.Caption = "For Date"
36410     End If
End Sub

Private Sub cmdAdvanceFind_Click()
36420     fraAdvanceFind.Visible = True
36430     fraMain.Enabled = False
36440     cmdFindAdvanceFind.Default = True
36450     cmdCancelAdvanceFind.Cancel = True
End Sub

Private Sub cmdCancel_Click()
36460     Unload Me
End Sub

Private Sub cmdCancelAdvanceFind_Click()
36470     fraAdvanceFind.Visible = False
36480     fraMain.Enabled = True
36490     cmdFindAdvanceFind.Default = True
36500     cmdCancelAdvanceFind.Cancel = True
End Sub

Private Sub cmdClear_Click()
36510     txtFind = ""
36520     FillGrid
End Sub

Private Sub cmdExport_Click()
36530     ExportFlexGrid Grid, Me, "Activity Log" & vbCr
End Sub

Private Sub cmdFind_Click()
          Dim tb As New Recordset
          Dim sql As String
          Dim gRow As String

36540     On Error GoTo cmdFind_Click_Error

36550     If Trim$(txtFind) <> "" Then
36560         sql = "Select * from ActivityLog " & _
                  "WHERE SampleID + SubmissionID + PatientID + ActionType + Action + Reason + Notes + UserName + MachineName + ApplicationVersion " & _
                  "Like '%" & txtFind & "%' order by DateTimeOfRecord desc"
36570         ReadyGrid
36580         RecOpenServer 0, tb, sql
36590         Do While Not tb.EOF
36600             s = tb!ActivityID & vbTab & _
                      tb!DateTimeOfRecord & "" & vbTab & _
                      tb!SampleID & "" & vbTab & _
                      tb!SubmissionID & "" & vbTab & _
                      tb!PatientID & "" & vbTab & _
                      tb!ActionType & "" & vbTab & _
                      tb!Action & "" & vbTab & _
                      tb!Reason & "" & vbTab & _
                      tb!Notes & "" & vbTab & _
                      tb!UserName & "" & vbTab
                  'tb!MachineName & "" & vbTab & _
                  'tb!ApplicationVersion
36610             Grid.AddItem s, Grid.Rows - 1
36620             tb.MoveNext
36630         Loop
36640     End If

36650     Exit Sub

cmdFind_Click_Error:

          Dim strES As String
          Dim intEL As Integer

36660     intEL = Erl
36670     strES = Err.Description
36680     LogError "frmActivityLog", "cmdFind_Click", intEL, strES, sql

End Sub

Private Sub cmdFindAdvanceFind_Click()
          Dim S1 As String
          Dim S2 As String
          Dim sql As String
        
36690     On Error GoTo cmdFindAdvanceFind_Click_Error

36700     S1 = "Select * from ActivityLog "
36710     If Len(Trim(txtSampleID)) > 0 Then
36720         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "SampleID Like '%" & txtSampleID & "%'"
36730     End If
36740     If Len(Trim(txtSubmissionID)) > 0 Then
36750         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "SubmissionID Like '%" & txtSubmissionID & "%'"
36760     End If
36770     If Len(Trim(txtPatientID)) > 0 Then
36780         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "PatientID Like '%" & txtPatientID & "%'"
36790     End If
36800     If cmbActionType.ListIndex > 0 Then
36810         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "ActionType = '" & cmbActionType.Text & "'"
36820     End If
36830     If cmbAction.ListIndex > 0 Then
36840         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "Action = '" & cmbAction.Text & "'"
36850     End If
36860     If Len(Trim(txtReason)) > 0 Then
36870         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "Reason Like '%" & AddTicks(txtReason) & "%'"
36880     End If
36890     If Len(Trim(txtNotes)) > 0 Then
36900         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "Notes Like '%" & AddTicks(txtNotes) & "%'"
36910     End If
36920     If cmbUserName.ListIndex > 0 Then
36930         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "UserName = '" & cmbUserName.Text & "'"
36940     End If
36950     If cmbMachine.ListIndex > 0 Then
36960         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "MachineName = '" & cmbMachine.Text & "'"
36970     End If
36980     If cmbAppVersion.ListIndex > 0 Then
36990         S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & "ApplicationVersion = '" & cmbAppVersion.Text & "'"
37000     End If
37010     If chkFromDate.Value = vbChecked Then
37020         If chkToDate.Value = vbUnchecked Then
37030             S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & " Convert(varchar(10),DateTimeOfRecord,105) ='" & Format(dtDate1.Value, "dd-MM-yyyy") & "'"
37040         Else
37050             S2 = S2 & IIf(Len(S2) > 0, " AND ", "") & " Convert(varchar(10),DateTimeOfRecord,105) BETWEEN '" & Format(dtDate1.Value, "dd-MM-yyyy") & "' AND '" & Format(dtDate2.Value, "dd-MM-yyyy") & "'"
37060         End If
37070     End If
37080     sql = S1 & IIf(Len(S2) > 0, " WHERE " & S2, "")
37090     sql = sql & " order by DateTimeOfRecord desc"

          Dim tb As New Recordset
37100     ReadyGrid
37110     RecOpenServer 0, tb, sql
37120     Do While Not tb.EOF
37130         s = tb!ActivityID & vbTab & _
                  tb!DateTimeOfRecord & "" & vbTab & _
                  tb!SampleID & "" & vbTab & _
                  tb!SubmissionID & "" & vbTab & _
                  tb!PatientID & "" & vbTab & _
                  tb!ActionType & "" & vbTab & _
                  tb!Action & "" & vbTab & _
                  tb!Reason & "" & vbTab & _
                  tb!Notes & "" & vbTab & _
                  tb!UserName & "" & vbTab
              'tb!MachineName & "" & vbTab & _
              'tb!ApplicationVersion
37140         Grid.AddItem s, Grid.Rows - 1
37150         tb.MoveNext
37160     Loop

37170     fraAdvanceFind.Visible = False
37180     fraMain.Enabled = True
37190     cmdFindAdvanceFind.Default = True
37200     cmdCancelAdvanceFind.Cancel = True

37210     Exit Sub

cmdFindAdvanceFind_Click_Error:

          Dim strES As String
          Dim intEL As Integer

37220     intEL = Erl
37230     strES = Err.Description
37240     LogError "frmActivityLog", "cmdFindAdvanceFind_Click", intEL, strES, sql

End Sub

Private Sub cmdPrint_Click()

37250     On Error GoTo cmdPrint_Click_Error

37260     If Grid.TextMatrix(1, 1) = "" Then
37270         iMsg "Nothing to print"
37280         If TimedOut Then Exit Sub
37290         Exit Sub
37300     End If

37310     Printer.Orientation = vbPRORLandscape

37320     Printer.Print

37330     Printer.FontName = "Courier New"
37340     Printer.FontSize = 10
37350     Printer.FontBold = True

37360     Printer.Print FormatString("Activity Log - " & Format(Now, "dd/mmm/yyyy hh:mm"), 120, , AlignCenter)
37370     Printer.Print

37380     Printer.Font.Bold = False
37390     Printer.FontSize = 4
37400     For n = 1 To 333
37410         Printer.Print "-";
37420     Next n
37430     Printer.Print

37440     Printer.FontSize = 8
37450     Printer.Font.Bold = True

37460     Printer.Print FormatString("", 0, "|", AlignCenter);
37470     Printer.Print FormatString("Date\Time", 19, "|", AlignCenter);
37480     Printer.Print FormatString("Sample Id", 10, "|", AlignCenter);
37490     Printer.Print FormatString("Submission Id", 14, "|", AlignCenter);
37500     Printer.Print FormatString("Patient Id", 11, "|", AlignCenter);
37510     Printer.Print FormatString("Action Type", 40, "|", AlignCenter);
37520     Printer.Print FormatString("Action", 30, "|", AlignCenter);
37530     Printer.Print FormatString("User Name", 35, "|", AlignCenter)

37540     Printer.Font.Bold = False
37550     Printer.FontSize = 4
37560     For n = 1 To 333
37570         Printer.Print "-";
37580     Next n
37590     Printer.Print

37600     Printer.FontSize = 8
37610     For n = 1 To Grid.Rows - 1
37620         If Grid.TextMatrix(n, 1) <> "" Then
37630             Printer.Print FormatString("", 0, " ", AlignCenter);
37640             Printer.Print FormatString(Format(Grid.TextMatrix(n, 1), "dd/mm/yyyy hh:mm"), 19, "|", AlignLeft);
37650             Printer.Print FormatString(Grid.TextMatrix(n, 2), 10, "|", AlignLeft);
37660             Printer.Print FormatString(Grid.TextMatrix(n, 3), 14, "|", AlignLeft);
37670             Printer.Print FormatString(Grid.TextMatrix(n, 4), 11, "|", AlignLeft);
37680             Printer.Print FormatString(Grid.TextMatrix(n, 5), 40, "|", AlignLeft);
37690             Printer.Print FormatString(Grid.TextMatrix(n, 6), 30, "|", AlignLeft);
37700             Printer.Print FormatString(Grid.TextMatrix(n, 9), 35, "|", AlignLeft)
37710             If Trim$(Grid.TextMatrix(n, 7)) <> "" Then
37720                 Printer.Print FormatString(" Reaseon: " & Grid.TextMatrix(n, 7), 170, "", AlignLeft)
37730             End If
37740             If Trim$(Grid.TextMatrix(n, 8)) <> "" Then
37750                 Printer.Print FormatString(" Notes: " & Grid.TextMatrix(n, 8), 170, "", AlignLeft)
37760             End If
37770         End If
37780     Next n

37790     Printer.Print
37800     Printer.Font.Bold = False
37810     Printer.FontSize = 4
37820     For n = 1 To 333
37830         Printer.Print "-";
37840     Next n
37850     Printer.Print

37860     Printer.EndDoc

37870     Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

37880     intEL = Erl
37890     strES = Err.Description
37900     LogError "frmActivityLog", "cmdPrint_Click", intEL, strES

End Sub



Private Sub Form_Activate()
37910     Me.Move (Screen.width - Me.width) / 2, (Screen.height - Me.height) / 2
End Sub

Private Sub Form_Load()
37920     fraAdvanceFind.Visible = False
37930     fraMain.Enabled = True
37940     ReadyGrid
37950     FillGrid
37960     FillCombos
37970     dtDate1.Value = Date
37980     dtDate2.Value = Date
End Sub
Sub FillCombos()
          Dim tb As New Recordset
          Dim sql As String

37990     On Error GoTo FillCombos_Error

38000     cmbActionType.Clear
38010     sql = "Select Distinct ActionType from ActivityLog order by ActionType"
38020     RecOpenServer 0, tb, sql
38030     cmbActionType.AddItem "--Any--"
38040     Do Until tb.EOF
38050         cmbActionType.AddItem tb!ActionType
38060         tb.MoveNext
38070     Loop
38080     cmbActionType.ListIndex = 0
38090     tb.Close

38100     cmbAction.Clear
38110     sql = "Select Distinct Action from ActivityLog order by Action"
38120     RecOpenServer 0, tb, sql
38130     cmbAction.AddItem "--Any--"
38140     Do Until tb.EOF
38150         cmbAction.AddItem tb!Action
38160         tb.MoveNext
38170     Loop
38180     cmbAction.ListIndex = 0
38190     tb.Close

38200     cmbUserName.Clear
38210     sql = "Select Distinct UserName from ActivityLog order by UserName"
38220     RecOpenServer 0, tb, sql
38230     cmbUserName.AddItem "--Any--"
38240     Do Until tb.EOF
38250         cmbUserName.AddItem tb!UserName
38260         tb.MoveNext
38270     Loop
38280     cmbUserName.ListIndex = 0
38290     tb.Close

38300     cmbMachine.Clear
38310     sql = "Select Distinct MachineName from ActivityLog order by MachineName"
38320     RecOpenServer 0, tb, sql
38330     cmbMachine.AddItem "--Any--"
38340     Do Until tb.EOF
38350         cmbMachine.AddItem tb!MachineName
38360         tb.MoveNext
38370     Loop
38380     cmbMachine.ListIndex = 0
38390     tb.Close

38400     cmbAppVersion.Clear
38410     sql = "Select Distinct ApplicationVersion from ActivityLog order by ApplicationVersion"
38420     RecOpenServer 0, tb, sql
38430     cmbAppVersion.AddItem "--Any--"
38440     Do Until tb.EOF
38450         cmbAppVersion.AddItem tb!ApplicationVersion
38460         tb.MoveNext
38470     Loop
38480     cmbAppVersion.ListIndex = 0
38490     tb.Close

38500     Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer

38510     intEL = Erl
38520     strES = Err.Description
38530     LogError "frmActivityLog", "FillCombos", intEL, strES, sql

End Sub

Sub ReadyGrid()

38540     On Error GoTo ReadyGrid_Error

38550     With Grid
38560         .Clear: .FixedCols = 0: .FixedRows = 1: .Rows = 2: .Cols = 10
38570         .ColWidth(0) = 0: .TextMatrix(0, 0) = "ID"
38580         .ColWidth(1) = 1800: .TextMatrix(0, 1) = "Date Time": .ColAlignment(1) = flexAlignLeftCenter
38590         .ColWidth(2) = 1000: .TextMatrix(0, 2) = "Sample ID": .ColAlignment(2) = flexAlignLeftCenter
38600         .ColWidth(3) = 0: .TextMatrix(0, 3) = "Submission ID"
38610         .ColWidth(4) = 1100: .TextMatrix(0, 4) = "Patient ID"
38620         .ColWidth(5) = 4000: .TextMatrix(0, 5) = "Action Type"
38630         .ColWidth(6) = 2150: .TextMatrix(0, 6) = "Action"
38640         .ColWidth(7) = 3350: .TextMatrix(0, 7) = "Reason"
38650         .ColWidth(8) = 4300: .TextMatrix(0, 8) = "Notes"
38660         .ColWidth(9) = 1400: .TextMatrix(0, 9) = "User"
              '140       .ColWidth(10) = 1200: .TextMatrix(0, 10) = "Machine": .ColAlignment(10) = flexAlignLeftCenter
              '150       .ColWidth(11) = 900: .TextMatrix(0, 11) = "App Ver": .ColAlignment(11) = flexAlignLeftCenter
38670         .SelectionMode = flexSelectionByRow
38680     End With
          'Grid.RemoveItem 1
38690     Exit Sub

ReadyGrid_Error:

          Dim strES As String
          Dim intEL As Integer

38700     intEL = Erl
38710     strES = Err.Description
38720     LogError "frmActivityLog", "ReadyGrid", intEL, strES
End Sub

Sub FillGrid()
          Dim tb As New Recordset
          Dim sql As String
          Dim gRow As String

38730     On Error GoTo FillGrid_Error

38740     ReadyGrid
38750     sql = "Select top 35 * from ActivityLog order by DateTimeOfRecord desc"
38760     RecOpenServer 0, tb, sql
38770     Do While Not tb.EOF
38780         s = tb!ActivityID & vbTab & _
                  tb!DateTimeOfRecord & "" & vbTab & _
                  tb!SampleID & "" & vbTab & _
                  tb!SubmissionID & "" & vbTab & _
                  tb!PatientID & "" & vbTab & _
                  tb!ActionType & "" & vbTab & _
                  tb!Action & "" & vbTab & _
                  tb!Reason & "" & vbTab & _
                  tb!Notes & "" & vbTab & _
                  tb!UserName & "" & vbTab
              'tb!MachineName & "" & vbTab & _
              'tb!ApplicationVersion
38790         Grid.AddItem s, Grid.Rows - 1
        
38800         tb.MoveNext
38810     Loop

38820     Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

38830     intEL = Erl
38840     strES = Err.Description
38850     LogError "frmActivityLog", "FillGrid", intEL, strES, sql

End Sub

Private Sub Form_Resize()
    '      Dim TWidth As Integer
    '      Dim n As Integer
    '      Dim AutoSizeColIndex As Integer
    '10    AutoSizeColIndex = 9 'Notes column
    '20    For n = 0 To Grid.Cols - 1
    '30        TWidth = TWidth + Grid.ColWidth(n)
    '40    Next
    '50    TWidth = TWidth - Grid.ColWidth(AutoSizeColIndex)
    '60    Grid.ColWidth(AutoSizeColIndex) = Grid.Width - TWidth

End Sub
