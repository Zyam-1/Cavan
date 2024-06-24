VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmGrpConf 
   Caption         =   "Group Confirm"
   ClientHeight    =   4770
   ClientLeft      =   1800
   ClientTop       =   1470
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "7frmGrpConf.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   5790
   Begin VB.TextBox txtSampleID 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   2835
   End
   Begin VB.OptionButton optGroupDetails 
      Caption         =   "Group && Details Check"
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   120
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.OptionButton optDetailsOnly 
      Caption         =   "Details Check Only"
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   420
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   915
      Left            =   4470
      Picture         =   "7frmGrpConf.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1830
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   915
      Left            =   4470
      Picture         =   "7frmGrpConf.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   810
      Width           =   1005
   End
   Begin VB.Frame Frame7 
      Caption         =   "Grouping"
      Height          =   3165
      Left            =   270
      TabIndex        =   6
      Top             =   1260
      Width           =   3765
      Begin VB.TextBox txtExpiry 
         Height          =   285
         Left            =   1110
         TabIndex        =   4
         Top             =   2085
         Width           =   2415
      End
      Begin VB.TextBox txtCardUsed 
         Height          =   285
         Left            =   1110
         TabIndex        =   3
         Top             =   1740
         Width           =   2415
      End
      Begin VB.ComboBox lstfg 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Text            =   "lstfg"
         Top             =   1230
         Width           =   1275
      End
      Begin MSFlexGridLib.MSFlexGrid gFG 
         Height          =   555
         Left            =   1320
         TabIndex        =   1
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   979
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   0
         FormatString    =   "^A    |^B    |^D    "
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Date Time"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label lblDateTime 
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1110
         TabIndex        =   19
         Top             =   2775
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Expiry"
         Height          =   195
         Left            =   615
         TabIndex        =   14
         Top             =   2130
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lot Number"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   1785
         Width           =   825
      End
      Begin VB.Label lblVal 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1110
         TabIndex        =   12
         Top             =   2430
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Validated By"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   2475
         Width           =   885
      End
      Begin VB.Label lblsuggestfg 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1350
         TabIndex        =   9
         Top             =   930
         Width           =   1035
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Suggest"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Report"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   7
         Top             =   1290
         Width           =   480
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   15
      Top             =   4560
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   405
      TabIndex        =   18
      Top             =   885
      Width           =   735
   End
End
Attribute VB_Name = "frmGrpConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim UserIn As String
Dim UserInCode As String
Private m_sSampleID As String

Private Sub cmdSave_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim fpat As String
          Dim rpat As String
          Dim ID As Long

10        On Error GoTo cmdSave_Click_Error

20        If txtSampleID = "" Then
30            iMsg "Please input Sample Number first"
40            If TimedOut Then Unload Me: Exit Sub
50            Exit Sub
60        End If

70        If optGroupDetails.Value = True Then    'Group & Details Check

80            Set tb = New Recordset
90            sql = "Select fGroup From PatientDetails Where LabNumber='" & txtSampleID & "'"
100           RecOpenClientBB 0, tb, sql
110           If tb.EOF Or IsNull(tb!fGroup) Then
120               iMsg "Invalid Sample Number"
130               If TimedOut Then Unload Me: Exit Sub
140               Exit Sub
150           Else
160               If lstfg <> tb!fGroup Then
170                   iMsg "Group does not match! Please re check.", vbCritical
180                   If TimedOut Then Unload Me: Exit Sub
190                   Exit Sub
200               End If
210               fpat = ""
220               rpat = ""

230               If Len(Trim$(txtCardUsed)) = 0 Then
240                   iMsg "Please input Lot number!"
250                   If TimedOut Then Unload Me: Exit Sub
260                   Exit Sub
270               End If

280               If Len(Trim$(txtExpiry)) = 0 Or Not IsDate(txtExpiry) Then
290                   iMsg "Please input correct expiry date!"
300                   If TimedOut Then Unload Me: Exit Sub
310                   Exit Sub
320               End If
330           End If
340       End If


350       sql = "select top 1 ID from GroupVal order by ID desc"
360       Set tb = New Recordset
370       RecOpenServerBB 0, tb, sql
380       If tb.EOF Then
390           ID = 1
400       Else
410           ID = tb!ID + 1
420       End If

430       sql = "select * from Groupval where labnumber = '" & txtSampleID & "'"
440       Set tb = New Recordset
450       RecOpenServerBB 0, tb, sql

460       If tb.EOF Then tb.AddNew

470       If optGroupDetails.Value = True Then    'Group & Details Check
480           tb!LabNumber = txtSampleID
490           tb!fg = lstfg
500           tb!Operator = UserInCode
510           tb!fpat = fpat
520           tb!ID = ID
530           tb!CardUsed = txtCardUsed
540           tb!Expiry = txtExpiry
550           tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")
560           tb.Update

570       ElseIf optDetailsOnly.Value = True Then

580           tb!LabNumber = txtSampleID
590           tb!Operator = UserInCode
600           tb!ID = ID
610           tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")
620           tb.Update
630       End If


640       CnxnBB(0).Execute "Update PatientDetails Set Checker = '" & TechnicianNameForCode(UserInCode) & "' Where labnumber = '" & txtSampleID & "'"
650       If UCase(Trim$(frmxmatch.tLabNum)) = UCase(Trim$(txtSampleID)) Then
660           frmxmatch.lblgrpchecker = TechnicianNameForCode(UserInCode)
670       End If
680       frmxmatch.cmdSave.Enabled = True
690       frmxmatch.bHold.Enabled = True

700       Unload Me

710       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

720       intEL = Erl
730       strES = Err.Description
740       LogError "frmGrpConf", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub Form_Activate()

10        If UserIn = "" Then
20            Unload Me
30        End If

End Sub

Private Sub Form_Load()

          Dim n As Integer
          Dim s As String
          Dim sql As String

10        On Error GoTo Form_Load_Error

20        For n = 0 To 14
30            s = Choose(n + 1, "", "O Neg", "O Pos", _
                         "A Neg", "A Pos", _
                         "B Neg", "B Pos", _
                         "AB Neg", "AB Pos", _
                         "O D- C/E+", "A D- C/E+", _
                         "B D- C/E+", "AB D-C/E+", _
                         "Control ?", "Error")
40            lstfg.AddItem s, n
50        Next

60        lstfg = ""

70        txtSampleID = Me.SampleID
80        If Me.SampleID <> "" Then
90            txtSampleID.Enabled = False
100       End If
110       Check_Val
120       If lblVal = "" Then
130           UserIn = ""
140           GetLog
150           lblVal = UserIn
160       End If


170       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmGrpConf", "Form_Load", intEL, strES, sql

End Sub
Private Sub GetLog()

          Dim Trial As String
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo GetLog_Error

20        Trial = Trim$(iBOX("Your Password", , , True))
30        If TimedOut Then Unload Me: Exit Sub

40        If Trial = "" Then Exit Sub

50        sql = "Select * from Users where " & _
                "UPPER(Password) = '" & AddTicks(UCase$(Trial)) & "' " & _
                "COLLATE SQL_Latin1_General_CP1_CS_AS"

60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql

80        If tb.EOF Then
90            iMsg "Incorrect Password!", vbExclamation
100           GetLog
110       Else
120           UserIn = Trim$(tb!Name)
130           UserInCode = tb!code
140       End If


150       Exit Sub

GetLog_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmGrpConf", "GetLog", intEL, strES, sql


End Sub
Private Sub Check_Val()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer

10        On Error GoTo Check_Val_Error

20        If Trim$(txtSampleID) = "" Then Exit Sub

30        sql = "Select fGroup From PatientDetails Where LabNumber='" & txtSampleID & "'"
40        Set tb = New Recordset
50        RecOpenClientBB 0, tb, sql
60        If tb.EOF Then
70            iMsg "Sample number unknown"
80            If TimedOut Then Unload Me: Exit Sub
90            BlankDetails
100           Frame7.Enabled = False
110           optGroupDetails.Enabled = False
120           optDetailsOnly.Enabled = False
130           Exit Sub
140       Else
150           If tb!fGroup = "" Then
160               iMsg "Forward group details unknown"
170               If TimedOut Then Unload Me: Exit Sub
180               BlankDetails
190               Frame7.Enabled = False
200               optGroupDetails.Enabled = False
210               optDetailsOnly.Enabled = False
220               Exit Sub
230           End If
240       End If

250       sql = "select * from Groupval where labnumber = '" & txtSampleID & "'"
260       Set tb = New Recordset
270       RecOpenServerBB 0, tb, sql

280       If Not tb.EOF Then

290           lstfg = tb!fg & ""
300           For n = 0 To 2
310               gFG.TextMatrix(1, n) = Mid$(tb!fpat & "", n + 1, 1)
320           Next
330           lblVal = TechnicianNameForCode(tb!Operator & "")
340           txtCardUsed = tb!CardUsed & ""
350           txtExpiry = tb!Expiry & ""
360           lblDateTime = tb!DateTimeOfRecord & ""
370           UserIn = lblVal
380           Frame7.Enabled = False
390           optGroupDetails.Enabled = False
400           optDetailsOnly.Enabled = False
410           txtSampleID.Enabled = False
420           cmdSave.Enabled = False
430       Else
440           BlankDetails
450           Frame7.Enabled = True
460           optGroupDetails.Enabled = True
470           optDetailsOnly.Enabled = True
480           txtSampleID.Enabled = True
490           cmdSave.Enabled = True
500       End If

510       If lblVal <> "" And lstfg <> "" Then
520           optDetailsOnly.Value = True
530       End If

540       Exit Sub


550       Exit Sub

Check_Val_Error:

          Dim strES As String
          Dim intEL As Integer

560       intEL = Erl
570       strES = Err.Description
580       LogError "frmGrpConf", "Check_Val", intEL, strES, sql


End Sub



Private Sub Form_Unload(Cancel As Integer)
10        Me.SampleID = ""
End Sub

Private Sub lstfg_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub
Private Sub gFG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim Filled As Boolean
          Dim Pattern As String
          Dim n As Integer
          Dim OAB As String
          Dim PN As String

10        If gFG.MouseRow = 0 Then Exit Sub

20        Select Case Trim$(gFG.TextMatrix(1, gFG.col))
          Case "": gFG.TextMatrix(1, gFG.col) = "O"
30        Case "O": gFG.TextMatrix(1, gFG.col) = "1"
40        Case "1": gFG.TextMatrix(1, gFG.col) = "2"
50        Case "2": gFG.TextMatrix(1, gFG.col) = "3"
60        Case "3": gFG.TextMatrix(1, gFG.col) = "4"
70        Case "4": gFG.TextMatrix(1, gFG.col) = ""
80        Case Else: gFG.TextMatrix(1, gFG.col) = ""
90        End Select

100       Filled = True
110       If Val(txtSampleID) > NewFormatFGNumber Then
120           For n = 0 To 2
130               If gFG.TextMatrix(1, n) = "" Then
140                   Filled = False
150                   Exit For
160               End If
170           Next
180       Else
190           For n = 0 To 2
200               If gFG.TextMatrix(1, n) = "" Then
210                   Filled = False
220                   Exit For
230               End If
240           Next
250       End If
260       If Not Filled Then
270           lstfg = ""
280           lblsuggestfg = ""
290           Exit Sub
300       End If

310       Pattern = ""
320       For n = 0 To 2
330           If Trim$(gFG.TextMatrix(1, n)) = "" Then
340               PN = " "
350           ElseIf gFG.TextMatrix(1, n) = "O" Then
360               PN = "O"
370           Else
380               PN = "+"
390           End If
400           Pattern = Pattern & PN
410       Next

420       Select Case Left$(Pattern, 4)
          Case "+OO": OAB = "A Neg"
430       Case "+O+": OAB = "A Pos"
440       Case "O+O": OAB = "B Neg"
450       Case "O++": OAB = "B Pos"
460       Case "++O": OAB = "AB Neg"
470       Case "+++": OAB = "AB Pos"
480       Case "OOO": OAB = "O Neg"
490       Case "OO+": OAB = "O Pos"
500       Case Else: OAB = "Error"
510       End Select

520       lstfg = OAB

530       lblsuggestfg = lstfg

End Sub

Private Sub optDetailsOnly_Click()
10        Frame7.Enabled = False
          'cmdSave.Enabled = True
End Sub

Private Sub optGroupDetails_Click()
10        Frame7.Enabled = True

End Sub

Private Sub txtCardUsed_LostFocus()
      'St Lukes "02120748038771132107"
          Dim Expiry As String
          Dim lot As String

10        If Len(txtCardUsed) = 19 Then
20            txtExpiry = Mid$(txtCardUsed, 10, 2) & "." & _
                          Mid$(txtCardUsed, 12, 2)
30            txtCardUsed = Left$(txtCardUsed, 5) & "." & _
                            Mid$(txtCardUsed, 6, 2) & "." & _
                            Mid$(txtCardUsed, 8, 2)
40        ElseIf Len(txtCardUsed) = 20 Then
50            If Mid$(txtCardUsed, 7, 2) <> "48" Then
60                iMsg "This is not an ABO DD Blood Grouping Card", vbCritical
70                If TimedOut Then Unload Me: Exit Sub
80                Exit Sub
90            End If

100           Expiry = Left$(txtCardUsed, 2) & "/" & Mid$(txtCardUsed, 3, 2) & "/" & Mid$(txtCardUsed, 5, 2)
110           Expiry = Format$(Expiry, "dd/MMM/yyyy")
120           txtExpiry = Expiry

130           lot = Mid$(txtCardUsed, 15, 3)
140           txtCardUsed = "ADD" & lot & "A"
150       End If

End Sub



Public Property Get SampleID() As String

10        SampleID = m_sSampleID

End Property

Public Property Let SampleID(ByVal sSampleID As String)

10        m_sSampleID = sSampleID

End Property



Private Sub txtSampleID_LostFocus()
10        Check_Val
End Sub

Private Sub BlankDetails()
10        gFG.TextMatrix(1, 0) = ""
20        gFG.TextMatrix(1, 1) = ""
30        gFG.TextMatrix(1, 2) = ""
40        lstfg = ""
50        txtCardUsed = ""
60        txtExpiry = ""

End Sub
