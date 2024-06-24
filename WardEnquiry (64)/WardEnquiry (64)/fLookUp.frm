VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Laboratory Result Look Up"
   ClientHeight    =   5520
   ClientLeft      =   3840
   ClientTop       =   1755
   ClientWidth     =   7590
   ControlBox      =   0   'False
   HelpContextID   =   10004
   Icon            =   "fLookUp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLookBack 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   3915
      TabIndex        =   23
      Top             =   4680
      Width           =   675
   End
   Begin VB.CommandButton cmdViewManualLab 
      Caption         =   "Lab User Manual"
      Height          =   1065
      Left            =   5370
      Picture         =   "fLookUp.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3420
      Width           =   1245
   End
   Begin VB.CommandButton cmdViewManual 
      Caption         =   "Ward Enquiry Manual"
      Height          =   1065
      Left            =   3930
      Picture         =   "fLookUp.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3420
      Width           =   1245
   End
   Begin VB.Timer tmrSingleUser 
      Interval        =   30000
      Left            =   7080
      Top             =   750
   End
   Begin VB.Frame fraChartLocation 
      Height          =   435
      Left            =   2025
      TabIndex        =   16
      Top             =   1305
      Visible         =   0   'False
      Width           =   3285
      Begin VB.OptionButton optHospChart 
         Caption         =   "Cavan"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   150
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optHospChart 
         Caption         =   "Both"
         Height          =   240
         Index           =   2
         Left            =   2340
         TabIndex        =   18
         Top             =   150
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.OptionButton optHospChart 
         Caption         =   "Monaghan"
         Height          =   240
         Index           =   1
         Left            =   1050
         TabIndex        =   17
         Top             =   150
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   5145
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21/06/2024"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "00:30"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7355
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton optSoundex 
      Caption         =   "Use Soundex"
      Height          =   225
      Left            =   5580
      TabIndex        =   12
      Top             =   2175
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.OptionButton optLeading 
      Caption         =   "Leading Characters"
      Height          =   195
      Left            =   5580
      TabIndex        =   11
      Top             =   1935
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtDoB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2040
      TabIndex        =   10
      Top             =   2460
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2040
      TabIndex        =   8
      Top             =   1890
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   735
      Left            =   5580
      Picture         =   "fLookUp.frx":14DE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   810
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7050
      Top             =   1290
   End
   Begin VB.CommandButton cmdLogOn 
      Caption         =   "Click to Log On"
      Height          =   1065
      Left            =   390
      Picture         =   "fLookUp.frx":1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "cmdLogOn"
      Top             =   3420
      Width           =   3285
   End
   Begin VB.TextBox txtChart 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   2010
      TabIndex        =   0
      Top             =   810
      Visible         =   0   'False
      Width           =   3285
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Lab results will be shown for last              days (Micro excluded)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   24
      Top             =   4725
      Width           =   6480
   End
   Begin VB.Label lblChartLocation 
      Caption         =   "Chart Location"
      Height          =   225
      Left            =   600
      TabIndex        =   20
      Top             =   1455
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   60
      TabIndex        =   14
      Top             =   2190
      Visible         =   0   'False
      Width           =   7425
   End
   Begin VB.Label lblDPA 
      Caption         =   "** DATA PROTECTION ACT 1988 **  (PLEASE READ BEFORE PROCEEDING)"
      Height          =   2715
      Left            =   840
      TabIndex        =   13
      Top             =   180
      Width           =   5805
   End
   Begin VB.Label lblDoB 
      Alignment       =   2  'Center
      Caption         =   "       Date of Birth      (ddmmyy or dd/mm/yy)"
      Height          =   435
      Left            =   90
      TabIndex        =   9
      Top             =   2475
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "       Name         (Surname Forename)"
      Height          =   405
      Left            =   180
      TabIndex        =   7
      Top             =   1950
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter Search Criteria then press <Search>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   405
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Chart Number"
      Height          =   195
      Index           =   2
      Left            =   690
      TabIndex        =   3
      Top             =   990
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lInfo 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Details Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2010
      TabIndex        =   5
      Top             =   1380
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuviewUnsignedSamples 
         Caption         =   "View Unsigned Samples"
      End
      Begin VB.Menu mNull 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSetUp 
      Caption         =   "&Set Up"
      Begin VB.Menu mnuPrinter 
         Caption         =   "&Printer"
      End
      Begin VB.Menu mnuBottleType 
         Caption         =   "&Bottle Types"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function SingleUserAlreadyLoggedOn(ByVal UserName As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo SingleUserAlreadyLoggedOn_Error

20    sql = "SELECT COUNT(*) Tot FROM WardEnqUsers WHERE " & _
            "UserName = '" & AddTicks(UserName) & "'"
30    Set tb = New Recordset
40    Set tb = Cnxn(0).Execute(sql)
50    SingleUserAlreadyLoggedOn = tb!Tot > 0

60    Exit Function

SingleUserAlreadyLoggedOn_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "frmMain", "SingleUserAlreadyLoggedOn", intEL, strES, sql

End Function

Private Function GetRunningInArea() As String

      Dim P As String
      Dim S() As String

10    On Error GoTo GetRunningInArea_Error

20    P = App.Path
30    S = Split(P, "\")

40    GetRunningInArea = S(UBound(S))

50    Exit Function

GetRunningInArea_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmMain", "GetRunningInArea", intEL, strES

End Function

Private Sub LogOff()

10    SingleUserLogOff UserName

20    frmMain.HelpContextID = 10004
30    lbl(0).Visible = False
40    lbl(2).Visible = False
50    txtChart.Visible = False
60    txtChart = ""
70    lInfo.Visible = False
80    cmdSearch.Visible = False
90    lblName.Visible = False
100   txtName.Visible = False
110   lblDoB.Visible = False
120   txtDoB.Visible = False
130   optSoundex.Visible = False
140   optLeading.Visible = False
150   mnuviewUnsignedSamples.Visible = False

160   mnuChangePassword.Enabled = False
170   StatusBar1.Panels(3).Text = ""

180   cmdLogOn.Caption = "Click to Log On"

190   UserName = ""
200   UserCode = ""
210   UserPass = ""

220   If UCase$(HospName(0)) = "CAVAN" Then
230       lblDPA.Visible = True
240       lblChartLocation.Visible = False
250       fraChartLocation.Visible = False
260   ElseIf UCase$(HospName(0)) = "MALLOW" Then
270       lblDPA.Visible = True
280       lblWarning.Visible = True
290   Else
300       lblDPA.Visible = False
310       lblWarning.Visible = False
320   End If

End Sub

Private Sub SingleUserLogOff(ByVal UserName As String)

      Dim sql As String

10    sql = "DELETE FROM WardEnqUsers WHERE " & _
            "UserName = '" & AddTicks(UserName) & "'"
20    Cnxn(0).Execute sql

End Sub

Private Sub cmdLogOn_Click()

lblDPA.Visible = False
lblWarning.Visible = False

If cmdLogOn.Caption = "Click to Log On" Then
    fManager.LookUp = True
    fManager.Administrator = True
    fManager.Show 1
    AskUserQuestion "UQ1"
    If LogOffDelaySecs <> 0 Then
        PBar.Max = LogOffDelaySecs
    End If
    PBar = 0
    If Trim$(UserName) <> "" Then

        '120       If SingleUserAlreadyLoggedOn(UserName) Then
        '130         iMsg "You are already Logged on elsewhere.", vbCritical
        '140         cmdLogOn.Caption = "Click to Log On"
        '150         UserName = ""
        '160         UserCode = ""
        '170         UserPass = ""
        '180         Exit Sub
        '190       Else
        '200         SingleUserUpdateLoggedOn (UserName)
        '210       End If
        '
        frmMain.HelpContextID = 10018
        lbl(0).Visible = True
        lbl(2).Visible = True
        txtChart.Visible = True
        cmdSearch.Visible = True
        cmdLogOn.Caption = "Log Off Now"
        LogAsViewed "L", "", ""
        If sysOptWardSearchName(0) Then
            lblName.Visible = True
            txtName.Visible = True
            optSoundex.Visible = False
            optLeading.Visible = False
        End If
        If sysOptWardSearchDoB(0) Then
            lblDoB.Visible = True
            txtDoB.Visible = True
        End If
        If sysOptViewUnsignedSamples(0) Then
            mnuviewUnsignedSamples.Visible = True
        End If
        If sysOptWardChartLocation(0) Then
            lblChartLocation.Visible = True
            fraChartLocation.Visible = False
            'Set Default option
            If UCase$(HospName(0)) = "CAVAN" Then
                optHospChart(0).Value = True
            ElseIf UCase$(HospName(0)) = "MONAGHAN" Then
                optHospChart(1).Value = True    'Monaghan
            Else
                optHospChart(2).Value = True    'Both
            End If

            StatusBar1.Panels(3).Text = UserName & " Logged On  "

        End If
        txtLookBack = GetOptionSetting("WardEnquiryLookBackDays", "35", "")
        mnuChangePassword.Enabled = True
    Else
        LogOff
    End If
Else
    LogAsViewed "M", "", ""
    LogOff
End If

End Sub



Public Sub cmdSearch_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim sqlCount As String
          Dim tbLatest As Recordset
          Dim SQLLatest As String
          Dim SearchFor As String
          Dim f As Form
          Dim strChart As String
          Dim strName As String
          Dim strDOB As String
          Dim Cn As Integer
10        ReDim cdn(0 To 0) As udtChartDoBName
          Dim TestCounter As Integer
          Dim n As Integer
          Dim Found As Boolean
          Dim strHospitalChart As String
          Dim RecordCount As Long

20        lInfo.Visible = False
30        strName = ""
40        strChart = ""

50        If Trim$(txtChart) <> "" Then
60            SearchFor = "C"
70            If sysOptWardChartLocation(0) Then    'Chart Location
80                If optHospChart(0) Then
90                    strHospitalChart = "Cavan"
100               ElseIf optHospChart(1) Then
110                   strHospitalChart = "Monaghan"
120               Else
130                   strHospitalChart = ""
140               End If

150               sql = "Select distinct PatName, Chart, DoB " & _
                      "from Demographics where " & _
                      "Chart = '" & AddTicks(txtChart) & "'"
160               If Len(strHospitalChart) > 0 Then
170                   sql = sql & " and Hospital = '" & strHospitalChart & "' "
180               End If

190               sql = sql & "AND RunDate > '" & Format(Now - Val(txtLookBack), "dd/MMM/yyyy") & "' "

200               SQLLatest = "Select top 1 rundate, PatName, Chart, DoB " & _
                      "from Demographics where " & _
                      "Chart = '" & AddTicks(txtChart) & "'"
210               sqlCount = "SELECT COUNT (DISTINCT PatName) Tot " & _
                      "FROM Demographics WHERE " & _
                      "Chart = '" & AddTicks(txtChart) & "'"

220               If Len(strHospitalChart) > 0 Then
230                   SQLLatest = SQLLatest & " and Hospital = '" & strHospitalChart & "' "
240               End If
250               SQLLatest = SQLLatest & "Order by RunDate Desc"

260           Else    'No Chart location
270               sql = "Select distinct PatName, Chart, DoB " & _
                      "from Demographics where " & _
                      "Chart = '" & AddTicks(txtChart) & "'"

280               sql = sql & "AND RunDate > '" & Format(Now - Val(txtLookBack), "dd/MMM/yyyy") & "' "
290               SQLLatest = "Select top 1 rundate, PatName, Chart, DoB " & _
                      "from Demographics where " & _
                      "Chart = '" & AddTicks(txtChart) & "' " & _
                      "Order by RunDate Desc"
300               sqlCount = "SELECT COUNT (DISTINCT PatName) Tot " & _
                      "FROM Demographics WHERE " & _
                      "Chart = '" & AddTicks(txtChart) & "'"
310           End If
320       ElseIf Trim$(txtName) <> "" Then
330           SearchFor = "N"
340           sql = "SELECT DISTINCT TOP 100 PatName, DoB, Chart " & _
                  "FROM Demographics WHERE "
350           If optSoundex Then
360               sql = sql & "SOUNDEX(PatName) = SOUNDEX('" & AddTicks(txtName) & "')"
370               SQLLatest = "Select top 1 rundate, PatName, Chart, DoB " & _
                      "from Demographics where " & _
                      "Soundex(PatName) = Soundex('" & AddTicks(txtName) & "')" & _
                      "Order by RunDate Desc"
380               sqlCount = "SELECT COUNT (DISTINCT PatName) Tot " & _
                      "FROM Demographics WHERE " & _
                      "Soundex(PatName) = Soundex('" & AddTicks(txtName) & "')"
390           Else
400               sql = sql & "PatName like '" & AddTicks(txtName) & "%'"
410               SQLLatest = "Select top 1 rundate, PatName, Chart, DoB " & _
                      "from Demographics where " & _
                      "PatName like '" & AddTicks(txtName) & "%'" & _
                      "Order by RunDate Desc"
420               sqlCount = "SELECT COUNT (DISTINCT PatName) Tot " & _
                      "FROM Demographics WHERE " & _
                      "PatName LIKE '" & AddTicks(txtName) & "%'"
430           End If
440       ElseIf Trim$(txtDoB) <> "" Then
450           SearchFor = "D"
460           If Len(Trim$(txtDoB)) = 6 Then
470               txtDoB = Trim$(txtDoB)
480               txtDoB = Left$(txtDoB, 2) & "/" & Mid$(txtDoB, 3, 2) & "/" & Right$(txtDoB, 2)
490           End If
500           If IsDate(txtDoB) Then
510               txtDoB = Format$(txtDoB, "dd/mmm/yyyy")
520           Else
530               txtDoB = ""
540               Exit Sub
550           End If
560           If DateDiff("d", txtDoB, Now) < 0 Then
570               txtDoB = DateAdd("yyyy", -100, txtDoB)
580           End If

590           sql = "Select distinct PatName, Chart, DoB " & _
                  "from Demographics where " & _
                  "DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "'"
600           SQLLatest = "Select top 1 rundate, PatName, Chart, DoB " & _
                  "from Demographics where " & _
                  "DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "'" & _
                  "Order by RunDate Desc"
610           sqlCount = "SELECT COUNT (DISTINCT PatName) Tot " & _
                  "FROM Demographics WHERE " & _
                  "DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "'"
620       Else
630           Exit Sub
640       End If

650       Set tb = New Recordset
660       RecOpenServer 0, tb, sqlCount
670       If Not tb Is Nothing Then
680           RecordCount = tb!Tot
690       End If


700       TestCounter = -1
710       For Cn = 0 To intOtherHospitalsInGroup
720           Set tb = New Recordset
730           RecOpenClient Cn, tb, sql

740           Do While Not tb.EOF
750               Found = False
760               For n = 0 To TestCounter
770                   If cdn(n).Chart = Trim$(tb!Chart & "") And _
                          cdn(n).Name = Trim$(tb!PatName & "") And _
                          cdn(n).DoB = Format$(tb!DoB, "dd/mmm/yyyy") Then
780                       Found = True
790                       cdn(n).Hospital = "Multiple"
800                       cdn(n).Cn = -1
810                       Exit For
820                   End If
830               Next
840               If Not Found Then
850                   TestCounter = TestCounter + 1
860                   ReDim Preserve cdn(0 To TestCounter) As udtChartDoBName
870                   cdn(TestCounter).Chart = Trim$(tb!Chart & "")
880                   cdn(TestCounter).Name = Trim$(tb!PatName & "")
890                   cdn(TestCounter).DoB = Format$(tb!DoB, "dd/mmm/yyyy")
900                   cdn(TestCounter).Hospital = HospName(Cn)
910                   cdn(TestCounter).Cn = Cn
920               End If
930               tb.MoveNext
940           Loop
950       Next


          '***********************************TRANSFUSION SEARCH START

960       If TestCounter = -1 Then
              'IF NO PATIENTS FOUND IN DEMOGRAPHICS, SEARCH IN PATIENT DETAILS (TRANSFUSION)
970           If Trim$(txtChart) <> "" Then
980               SearchFor = "C"
990               If sysOptWardChartLocation(0) Then    'Chart Location
1000                  If optHospChart(0) Then
1010                      strHospitalChart = "C"
1020                  ElseIf optHospChart(1) Then
1030                      strHospitalChart = "M"
1040                  Else
1050                      strHospitalChart = ""
1060                  End If

1070                  sql = "Select distinct name PatName, patnum Chart, DoB " & _
                          "from PatientDetails where " & _
                          "patnum = '" & AddTicks(txtChart) & "'"
1080                  If Len(strHospitalChart) > 0 Then
1090                      sql = sql & " and Hospital = '" & strHospitalChart & "' "
1100                  End If

1110                  SQLLatest = "Select top 1 datetime rundate, name PatName, patnum Chart, DoB " & _
                          "from PatientDetails where " & _
                          "patnum = '" & AddTicks(txtChart) & "'"
1120                  sqlCount = "SELECT COUNT (DISTINCT name) Tot " & _
                          "FROM PatientDetails WHERE " & _
                          "patnum = '" & AddTicks(txtChart) & "'"

1130                  If Len(strHospitalChart) > 0 Then
1140                      SQLLatest = SQLLatest & " and Hospital = '" & strHospitalChart & "' "
1150                  End If
1160                  SQLLatest = SQLLatest & "Order by datetime Desc"

1170              Else    'No Chart location
1180                  sql = "Select distinct name PatName, patnum Chart, DoB " & _
                          "from PatientDetails where " & _
                          "patnum = '" & AddTicks(txtChart) & "'"
1190                  SQLLatest = "Select top 1 datetime rundate, name PatName, patnum Chart, DoB " & _
                          "from PatientDetails where " & _
                          "patnum = '" & AddTicks(txtChart) & "' " & _
                          "Order by datetime Desc"
1200                  sqlCount = "SELECT COUNT (DISTINCT name) Tot " & _
                          "FROM PatientDetails WHERE " & _
                          "patnum = '" & AddTicks(txtChart) & "'"
1210              End If
1220          ElseIf Trim$(txtName) <> "" Then
1230              SearchFor = "N"
1240              sql = "SELECT DISTINCT TOP 100 name PatName, DoB, patnum Chart " & _
                      "FROM PatientDetails WHERE "
1250              If optSoundex Then
1260                  sql = sql & "SOUNDEX(Name) = SOUNDEX('" & AddTicks(txtName) & "')"
1270                  SQLLatest = "Select top 1 datetime rundate, name PatName, patnum Chart, DoB " & _
                          "from PatientDetails where " & _
                          "Soundex(Name) = Soundex('" & AddTicks(txtName) & "')" & _
                          "Order by datetime Desc"
1280                  sqlCount = "SELECT COUNT (DISTINCT Name) Tot " & _
                          "FROM PatientDetails WHERE " & _
                          "Soundex(Name) = Soundex('" & AddTicks(txtName) & "')"
1290              Else
1300                  sql = sql & "Name like '" & AddTicks(txtName) & "%'"
1310                  SQLLatest = "Select top 1 datetime rundate, name PatName, patnum Chart, DoB " & _
                          "from PatientDetails where " & _
                          "Name like '" & AddTicks(txtName) & "%'" & _
                          "Order by datetime Desc"
1320                  sqlCount = "SELECT COUNT (DISTINCT Name) Tot " & _
                          "FROM PatientDetails WHERE " & _
                          "Name LIKE '" & AddTicks(txtName) & "%'"
1330              End If
1340          ElseIf Trim$(txtDoB) <> "" Then
1350              SearchFor = "D"
1360              If Len(Trim$(txtDoB)) = 6 Then
1370                  txtDoB = Trim$(txtDoB)
1380                  txtDoB = Left$(txtDoB, 2) & "/" & Mid$(txtDoB, 3, 2) & "/" & Right$(txtDoB, 2)
1390              End If
1400              If IsDate(txtDoB) Then
1410                  txtDoB = Format$(txtDoB, "dd/mmm/yyyy")
1420              Else
1430                  txtDoB = ""
1440                  Exit Sub
1450              End If
1460              If DateDiff("d", txtDoB, Now) < 0 Then
1470                  txtDoB = DateAdd("yyyy", -100, txtDoB)
1480              End If

1490              sql = "Select distinct name PatName, patnum Chart, DoB " & _
                      "from PatientDetails where " & _
                      "DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "'"
1500              SQLLatest = "Select top 1 datetime rundate, name PatName, patnum Chart, DoB " & _
                      "from PatientDetails where " & _
                      "DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "'" & _
                      "Order by datetime Desc"
1510              sqlCount = "SELECT COUNT (DISTINCT name) Tot " & _
                      "FROM PatientDetails WHERE " & _
                      "DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "'"
1520          Else
1530              Exit Sub
1540          End If

1550          Set tb = New Recordset
1560          RecOpenServerBB 0, tb, sqlCount
1570          If Not tb Is Nothing Then
1580              RecordCount = tb!Tot
1590          End If

1600          For Cn = 0 To intOtherHospitalsInGroup
1610              Set tb = New Recordset
1620              RecOpenClientBB Cn, tb, sql

1630              Do While Not tb.EOF
1640                  Found = False
1650                  For n = 0 To TestCounter
1660                      If cdn(n).Chart = Trim$(tb!Chart & "") And _
                              cdn(n).Name = Trim$(tb!PatName & "") And _
                              cdn(n).DoB = Format$(tb!DoB, "dd/mmm/yyyy") Then
1670                          Found = True
1680                          cdn(n).Hospital = "Multiple"
1690                          cdn(n).Cn = -1
1700                          Exit For
1710                      End If
1720                  Next
1730                  If Not Found Then
1740                      TestCounter = TestCounter + 1
1750                      ReDim Preserve cdn(0 To TestCounter) As udtChartDoBName
1760                      cdn(TestCounter).Chart = Trim$(tb!Chart & "")
1770                      cdn(TestCounter).Name = Trim$(tb!PatName & "")
1780                      cdn(TestCounter).DoB = Format$(tb!DoB, "dd/mmm/yyyy")
1790                      cdn(TestCounter).Hospital = HospName(Cn)
1800                      cdn(TestCounter).Cn = Cn
1810                  End If
1820                  tb.MoveNext
1830              Loop
1840          Next

1850      End If

          '**************************TRANSFUSION SEARCH END

1860      If TestCounter = -1 Then
1870          lInfo.Visible = True
1880          Select Case SearchFor
                  Case "C": txtChart.SelStart = 0: txtChart.SelLength = Len(txtChart): txtChart.SetFocus
1890              Case "D": txtDoB.SelStart = 0: txtDoB.SelLength = Len(txtDoB): txtDoB.SetFocus
1900              Case "N": txtName.SelStart = 0: txtName.SelLength = Len(txtName): txtName.SetFocus
1910          End Select
1920      ElseIf TestCounter = 0 Then
1930          lInfo.Visible = False
1940          With frmViewResultsWE
1950              .grd.Rows = 2
1960              .grd.AddItem ""
1970              .grd.RemoveItem 1
1980              .grd.AddItem cdn(0).Chart & vbTab & _
                      cdn(0).DoB & vbTab & _
                      cdn(0).Name
1990              .grd.RemoveItem 1
2000              .lblChart = cdn(0).Chart
2010              .lblName = cdn(0).Name
2020              .lblDoB = cdn(0).DoB
2030              .Show 1
2040          End With
2050          Select Case SearchFor
                  Case "C": txtChart = "": txtChart.SetFocus
2060              Case "D": txtDoB = "": txtDoB.SetFocus
2070              Case "N": txtName = "": txtName.SetFocus
2080          End Select
2090      Else
2100          Set f = New frmConflict
2110          With f.grd
2120              .Rows = 2
2130              .AddItem ""
2140              .RemoveItem 1

2150              For n = 0 To UBound(cdn)
2160                  .AddItem cdn(n).Chart & vbTab & _
                          cdn(n).DoB & vbTab & _
                          cdn(n).Name & vbTab & _
                          cdn(n).Hospital & vbTab & _
                          cdn(n).Cn

2170              Next

2180              .RemoveItem 1
2190          End With

2200          With f

                  '        .CountWarning = RecordCount

2210              Set tbLatest = New Recordset
2220              If InStr(1, SQLLatest, "Demographics") > 0 Then
2230                  RecOpenServer 0, tbLatest, SQLLatest
2240              Else
2250                  RecOpenServerBB 0, tbLatest, SQLLatest
2260              End If
2270              If Not tbLatest.EOF Then
2280                  .RecentPatName = Trim$(tbLatest!PatName & "")
2290                  .RecentChart = Trim$(tbLatest!Chart & "")
2300                  .RecentDoB = Format$(tbLatest!DoB, "dd/mmm/yyyy")
2310                  .RecentDate = Format$(tbLatest!Rundate, "dd/mmm/yyyy")
2320              End If

2330              .Show 1

2340              strName = .PatName
2350              strChart = .Chart
2360              strDOB = .DoB
2370          End With
              'go thru grid and find which patients are selected

2380          With frmViewResultsWE
2390              .grd.Rows = 2
2400              .grd.AddItem ""
2410              .grd.RemoveItem 1
2420              For n = 1 To f.grd.Rows - 1
2430                  f.grd.Row = n
2440                  f.grd.col = 5
2450                  If f.grd.CellPicture = f.imgGreenTick.Picture Then
2460                      .grd.AddItem f.grd.TextMatrix(n, 0) & vbTab & _
                              f.grd.TextMatrix(n, 1) & vbTab & _
                              f.grd.TextMatrix(n, 2)
2470                  End If

2480              Next n
2490              Set f = Nothing
2500              If .grd.Rows > 2 Then
2510                  .grd.RemoveItem 1
2520                  .Show 1
2530              End If

2540          End With


              '        If strName <> "" Then
              '            With frmViewResultsWE
              '                .lblChart = strChart
              '                .lblName = strName
              '                .lblDoB = strDoB
              '                .Show 1
              '            End With
              '        End If
2550          Select Case SearchFor
                  Case "C": txtChart = "": If txtChart.Visible Then txtChart.SetFocus
2560              Case "D": txtDoB = "": If txtDoB.Visible Then txtDoB.SetFocus
2570              Case "N": txtName = "": If txtName.Visible Then txtName.SetFocus
2580          End Select
2590      End If

End Sub

Private Sub cmdTest_Click()

End Sub

Private Sub cmdViewManual_Click()

      Dim PathToDoc As String

10    On Error GoTo cmdViewManual_Click_Error

20    PathToDoc = App.Path & "\ward enquiry user guide.doc"

30    ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5

40    Exit Sub

cmdViewManual_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmMain", "cmdViewManual_Click", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdViewManualLab_Click
' Author    : XPMUser
' Date      : 28/Jan/15
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdViewManualLab_Click()

10      On Error GoTo cmdViewManualLab_Click_Error

      Dim PathToDoc As String

20    PathToDoc = App.Path & "\Lab User Manual.pdf"

30    ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5


       
40    Exit Sub

       
cmdViewManualLab_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmMain", "cmdViewManualLab_Click", intEL, strES
End Sub

Private Sub Form_Activate()

      Dim Path As String
      Dim strVersion As String
10    If Not IsIDE Then
20        Path = CheckNewEXE("WardEnq")
30        If Path <> "" Then
40            Shell App.Path & "\CustomStart.exe WardEnq"
50            End
60            Exit Sub
70        End If
80    End If

90    intOtherHospitalsInGroup = 0

100   LoadOptions

110   PBar = 0
120   SingleUserUpdateLoggedOn UserName

130   Timer1.Enabled = True

140   RunningInArea = GetRunningInArea()
150   WardEnqForcedPrinter = GetOptionSetting("WardEnqForcedPrinter", "", RunningInArea)

160   strVersion = App.Major & "." & App.Minor & "." & App.Revision
170   Me.Caption = "NetAcquire - Laboratory Result Look Up. V. " & _
                   strVersion & " (" & RunningInArea & ")"

180   If txtChart.Visible And txtChart.Enabled Then
190       txtChart.SetFocus
          If m_RunExe = "OCM" Then
            If m_FindChart = True Then
                txtChart.Text = m_Chart
                DoEvents
                DoEvents
                Call cmdSearch_Click
                m_RunExe = "NA"
                m_FindChart = False
            End If
          End If
200   End If
        If m_FindChart = False Then
            If m_RunExe = "OCM" Then
                Call cmdLogOn_Click
            End If
        End If

End Sub

Private Sub Form_Deactivate()

10    Timer1.Enabled = False

End Sub


Private Sub Form_Load()

      Dim S(0 To 2) As String

10    If App.PrevInstance Then End

20    App.HelpFile = App.Path & "\WardEnq.chm"

30    On Error Resume Next

40    CheckIDE

50    ConnectToDatabase

60    S(0) = "** DATA PROTECTION ACT 1988 **  (PLEASE READ BEFORE PROCEEDING)" & vbCrLf & _
             vbCrLf & _
             "You are about to gain access to confidential, sensitive client" & vbCrLf & _
             "information. The Data Protection Act, 1988, obliges you to" & vbCrLf & _
             "safeguard client information whilst using this terminal." & vbCrLf & _
             "Consequently:-" & vbCrLf & _
             "       1) Do not disclose your personal password to ANYBODY" & vbCrLf & _
             "       2) Do not leave the terminal unattended without logging out" & vbCrLf & _
             "       3) Do not allow unauthorised personnel to see information on screen" & vbCrLf & _
             "       4) Do not pass on personal information to unauthorised personnel"

70    S(1) = S(0)

80    S(2) = "Warning: Patient results may have been entered with or without identifiers (MRN, D.O.B., Etc)." & vbCrLf & _
             "A thorough search must include MRN, Name and DoB Searches. (Including variations - Margaret," & vbCrLf & _
             "Mgt, Peggy, OBrien, O'Brien, O Brien, McCarthy, Mc Carthy, MacCarthy, Etc)" & vbCrLf & _
             "The operator must ensure the test results are specific for the date and patient required." & vbCrLf & _
             "Note: Multiple results of the SAME date may not be easily distinguished."

90    If UCase$(HospName(0)) = "CAVAN" Then
100       lblDPA.Visible = True
110       lblDPA.Caption = S(1)
120   ElseIf UCase$(HospName(0)) = "MALLOW" Then
130       lblDPA.Visible = True
140       lblDPA.Caption = S(0)
150       lblWarning.Visible = True
160       lblWarning.Caption = S(2)
170   ElseIf UCase$(HospName(0)) = "HOGWARTS" Then
180       lblDPA.Visible = True
190       lblDPA.Caption = S(0)
200   Else
210       lblDPA.Visible = False
220       lblWarning.Visible = False
230   End If

240   txtLookBack = GetOptionSetting("WardEnquiryLookBackDays", "35", "")


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub

Private Sub lInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub

Private Sub mAbout_Click()

10    frmAbout.Show 1

End Sub

Private Sub mExit_Click()

      Dim f As Form

10    LogAsViewed "X", "", ""

20    For Each f In Forms
30        Debug.Print f.Caption
40        Unload f
50    Next

60    Unload Me

End Sub


Private Sub mnuBottleType_Click()

10    frmOCBottleType.Show 1

End Sub

Private Sub mnuChangePassword_Click()

      Dim NewPass As String
      Dim Confirm As String
      Dim tb As Recordset
      Dim sql As String
      Dim MinLength As Integer
      Dim Current As String
      Dim PasswordExpiry As Long
      Dim AllowReUse As String

10    On Error GoTo mnuChangePassword_Click_Error

20    Current = iBOX("Enter your current Password", , , True)
30    sql = "SELECT * FROM Users WHERE " & _
            "Name = '" & AddTicks(UserName) & "' " & _
            "AND Password = '" & AddTicks(Current) & "' "
40    If GetOptionSetting("LogOnUpperLower", False, "") Then
50        sql = sql & "COLLATE SQL_Latin1_General_CP1_CS_AS"
60    End If
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If Not tb.EOF Then

100       NewPass = iBOX("Enter new password" & vbCrLf & vbCrLf & "Password must have at least 6 characters, contain letters and numbers", , , True)

110       MinLength = Val(GetOptionSetting("LogOnMinPassLen", "1", ""))
120       If Len(NewPass) < MinLength Then
130           iMsg "Passwords must have a minimum of " & Format(MinLength) & " characters!", vbExclamation
140           Exit Sub
150       End If

160       If GetOptionSetting("LogOnUpperLower", False, "") Then
170           If AllLowerCase(NewPass) Or AllUpperCase(NewPass) Then
180               iMsg "Passwords must have a mixture of UPPER CASE and lower case letters!", vbExclamation
190               Exit Sub
200           End If
210       End If

220       If GetOptionSetting("LogOnNumeric", False, "") Then
230           If Not ContainsNumeric(NewPass) Then
240               iMsg "Passwords must contain a numeric character!", vbExclamation
250               Exit Sub
260           End If
270       End If

280       If GetOptionSetting("LogOnAlpha", False, "") Then
290           If Not ContainsAlpha(NewPass) Then
300               iMsg "Passwords must contain an alphabetic character!", vbExclamation
310               Exit Sub
320           End If
330       End If

340       AllowReUse = GetOptionSetting("PasswordReUse", "No", "")
350       If AllowReUse = "No" Then
360           If PasswordHasBeenUsed(NewPass) Then
370               iMsg vbCrLf & "Password has been used!" & vbCrLf & vbCrLf & "Please use a different password!", vbExclamation
380               Exit Sub
390           End If
400       End If

410       Confirm = iBOX("Confirm password", , , True)

420       If NewPass <> Confirm Then
430           iMsg "Passwords don't match!", vbExclamation
440           Exit Sub
450       End If

460       Cnxn(0).Execute sql

470       PasswordExpiry = Val(GetOptionSetting("PasswordExpiry", "90", ""))

480       sql = "UPDATE Users SET " & _
                "PassWord = '" & NewPass & "', " & _
                "PassDate = '" & Format$(Now + PasswordExpiry, "dd/MMM/yyyy") & "', " & _
                "ExpiryDate = '" & Format$(Now + PasswordExpiry, "dd/MMM/yyyy") & "' " & _
                "WHERE " & _
                "Name = '" & AddTicks(UserName) & "'"
490       Cnxn(0).Execute sql

500       iMsg "Your Password has been changed.", vbInformation

510   End If

520   Exit Sub

mnuChangePassword_Click_Error:

      Dim strES As String
      Dim intEL As Integer

530   intEL = Erl
540   strES = Err.Description
550   LogError "frmMain", "mnuChangePassword_Click", intEL, strES, sql

End Sub

Private Sub mnuPrinter_Click()

      Dim Px As String

10    If UCase$(iBOX("Password", , , True)) = "TEMO" Then

20        Px = iBOX("Enter Printer Path and Name", , WardEnqForcedPrinter)
30        If Trim$(Px) <> "" Then
40            SaveOptionSetting "WardEnqForcedPrinter", Px, RunningInArea
50        End If

60    Else
70        iMsg "Incorrect Password"
80    End If

End Sub

Private Sub mnutest_Click()

End Sub

Private Sub mnuviewUnsignedSamples_Click()
10    frmSignOffSamples.Show
End Sub

Private Sub tmrSingleUser_Timer()

      '30 sec timer
      Static Counter As Integer
      Dim sql As String

10    On Error GoTo tmrSingleUser_Timer_Error

20    Counter = Counter + 1

30    If Counter > 10 Then
40        sql = "DELETE FROM WardEnqUsers WHERE " & _
                "DATEDIFF(minute, DateTimeOfRecord, GETDATE()) > 5"
50        Cnxn(0).Execute sql
60        Counter = 0
70    End If

80    Exit Sub

tmrSingleUser_Timer_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmMain", "tmrSingleUser_Timer", intEL, strES, sql

End Sub

Private Sub txtChart_KeyUp(KeyCode As Integer, Shift As Integer)

10    lInfo.Visible = False

20    txtName = ""
30    txtDoB = ""

40    PBar = 0

End Sub

Private Sub txtChart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub

Private Sub Timer1_Timer()

10    If LogOffNow Then
20        LogOffNow = False
30        LogAsViewed "O", "", ""
40        LogOff
50        Exit Sub
60    End If

      'tmrRefresh.Interval set to 1000
70    If cmdLogOn.Caption = "Log Off Now" Then
80        PBar.Visible = True

90        PBar = PBar + 1

100       If PBar = PBar.Max Then
110           LogAsViewed "O", "", ""
120           LogOff
130       End If
140   Else
150       PBar.Visible = False
160       PBar = 0
170   End If

End Sub

Private Sub txtDoB_KeyUp(KeyCode As Integer, Shift As Integer)

10    lInfo.Visible = False

20    txtChart = ""
30    txtName = ""

40    PBar = 0

End Sub

Private Sub txtDoB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


Private Sub txtLookBack_KeyPress(KeyAscii As Integer)
10    On Error GoTo txtLookBack_KeyPress_Error

20    KeyAscii = VI(KeyAscii, AlphaNumeric)

30    Exit Sub
txtLookBack_KeyPress_Error:
         
40    LogError "frmMain", "txtLookBack_KeyPress", Erl, Err.Description


End Sub


Private Sub txtLookBack_LostFocus()
10    On Error GoTo txtLookBack_LostFocus_Error

20    If Val(txtLookBack) < 1 Or Val(txtLookBack) > 3650 Then txtLookBack = 35
30    'SaveOptionSetting "WardEnquiryLookBackDays", Val(txtLookBack), ""

40    Exit Sub
txtLookBack_LostFocus_Error:
         
50    LogError "frmMain", "txtLookBack_LostFocus", Erl, Err.Description


End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

10    On Error GoTo txtName_KeyPress_Error

20    KeyAscii = VI(KeyAscii, AlphaAndSpaceAposDash)

30    Exit Sub

txtName_KeyPress_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmMain", "txtName_KeyPress", intEL, strES
          
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)

10    lInfo.Visible = False

20    txtChart = ""
30    txtDoB = ""

40    PBar = 0

End Sub

Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

10    PBar = 0

End Sub


