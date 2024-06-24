VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmScaneSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Scan new samples"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   14400
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox txtDateTime 
      Height          =   315
      Left            =   1380
      TabIndex        =   5
      ToolTipText     =   "Time of Sample"
      Top             =   60
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   585
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6300
      Width           =   1500
   End
   Begin VB.CommandButton cmdFinishScan 
      Caption         =   "&Place Order"
      Height          =   585
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6300
      Width           =   1500
   End
   Begin MSFlexGridLib.MSFlexGrid flxSampleDetails 
      Height          =   5655
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   10
      RowHeightMin    =   315
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSampleID 
      Height          =   345
      Left            =   6015
      TabIndex        =   0
      Top             =   90
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID:"
      Height          =   195
      Left            =   5175
      TabIndex        =   1
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmScaneSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private grd As MSFlexGrid
Dim m_SampleID As String

Private Sub LoadControls(grd As MSFlexGrid, txtText As MaskEdBox)
50360 On Error GoTo LoadControls_Error

50370 txtText.Visible = False
      'txtText = ""
      'gRD.SetFocus
50380 If grd.TextMatrix(grd.row, grd.Col) = "" Then Exit Sub

50390 txtText.Move grd.Left + grd.CellLeft + 5, _
                   grd.Top + grd.CellTop + 5, _
                   grd.CellWidth - 20, grd.CellHeight - 20
50400 txtText.Text = grd.TextMatrix(grd.row, grd.Col)
50410 txtText.Visible = True
50420 txtText.SelStart = 0
50430 txtText.SelLength = Len(txtText)
50440 txtText.SetFocus


50450 Exit Sub

LoadControls_Error:

      Dim strES As String
      Dim intEL As Integer

50460 intEL = Erl
50470 strES = Err.Description
50480 LogError "frmOptions", "LoadControls", intEL, strES

End Sub

Private Sub FormatGrid()
50490 On Error GoTo FormatGrid_Error

50500 flxSampleDetails.Rows = 1
50510 flxSampleDetails.row = 0

50520 flxSampleDetails.ColWidth(0) = 0

50530 flxSampleDetails.TextMatrix(0, 1) = "Sample ID"
50540 flxSampleDetails.ColWidth(1) = 1100
50550 flxSampleDetails.ColAlignment(1) = flexAlignLeftCenter

50560 flxSampleDetails.TextMatrix(0, 2) = "Patient Name"
50570 flxSampleDetails.ColWidth(2) = 3100
50580 flxSampleDetails.ColAlignment(3) = flexAlignLeftCenter

50590 flxSampleDetails.TextMatrix(0, 3) = "Date of Birth"
50600 flxSampleDetails.ColWidth(3) = 1400
50610 flxSampleDetails.ColAlignment(3) = flexAlignLeftCenter

50620 flxSampleDetails.TextMatrix(0, 4) = "Profile"
50630 flxSampleDetails.ColWidth(4) = 3700
50640 flxSampleDetails.ColAlignment(4) = flexAlignLeftCenter

50650 flxSampleDetails.TextMatrix(0, 5) = "Urgent"
50660 flxSampleDetails.ColWidth(5) = 800
50670 flxSampleDetails.ColAlignment(5) = flexAlignCenterCenter

50680 flxSampleDetails.TextMatrix(0, 6) = "Sample Date"
50690 flxSampleDetails.ColWidth(6) = 1800
50700 flxSampleDetails.ColAlignment(6) = flexAlignCenterCenter

50710 flxSampleDetails.TextMatrix(0, 7) = "Received Date"
50720 flxSampleDetails.ColWidth(7) = 1800
50730 flxSampleDetails.ColAlignment(7) = flexAlignCenterCenter

50740 flxSampleDetails.TextMatrix(0, 8) = ""
50750 flxSampleDetails.ColWidth(8) = 0
50760 flxSampleDetails.ColAlignment(8) = flexAlignLeftCenter

50770 flxSampleDetails.TextMatrix(0, 9) = ""
50780 flxSampleDetails.ColWidth(9) = 250
50790 flxSampleDetails.ColAlignment(9) = flexAlignLeftCenter



50800 Exit Sub
FormatGrid_Error:

      Dim strES As String
      Dim intEL As Integer

50810 intEL = Erl
50820 strES = Err.Description
50830 LogError "frmScan", "Form_Load", intEL, strES
End Sub

Private Sub cmdFinishScan_Click()
50840 On Error GoTo cmdCancel_Click_Error

      Dim i As Integer

50850 With flxSampleDetails
50860     For i = 1 To .Rows - 1
50870         PlaceOrder .TextMatrix(i, 1), .TextMatrix(i, 6), .TextMatrix(i, 7)
50880         UpdateRequestStatus .TextMatrix(i, 0), "Received in the Lab"
50890     Next
50900 End With

50910 Unload Me

50920 Exit Sub
cmdCancel_Click_Error:

50930 LogError "frmScaneSample", "cmdFinishScan_Click", Erl, Err.Description
End Sub

Private Sub PlaceOrder(ByVal SampleID As String, grdSampleDate As String, grdSampleReceivedDate As String)

      Dim Dept As String
      Dim Code As String
      Dim ST As String
      Dim Analyser As String
      Dim sql As String
      Dim tb As Recordset
      Dim DelHaeReq As Boolean

50940 On Error GoTo PlaceOrder_Error

50950 sql = "SELECT * FROM ocmRequestDetails WHERE SampleID = '" & SampleID & "' AND Programmed = '0'"

50960 Set tb = New Recordset
50970 RecOpenServer 0, tb, sql
50980 Do While Not tb.EOF

50990     Dept = GetOCMMapping("Department", "Cavan", tb!DepartmentID & "")
51000     Code = GetOCMMapping("TestCode", "Cavan", tb!TestCode & "")
51010     ST = GetOCMMapping("SampleType", "Cavan", tb!SampleType & "")

51020     If Dept <> "" Then

51030         If Code = "" Then Code = tb!TestCode

51040         Select Case UCase(Dept)
              Case "BIO"
51050             Analyser = AnalyserFor(Dept, Code)
51060             UpDateRequestBio SampleID, Code, ST, Analyser, 0

51070         Case "HAEM"
                  'delete old entries
51080             If DelHaeReq = False Then
51090                 sql = "Delete from  HaeRequests "
51100                 sql = sql & " WHERE SAMPLEID ='" & SampleID & "'"
51110                 Cnxn(0).Execute sql
51120                 DelHaeReq = True
51130             End If
51140             SaveHae SampleID, Code, UCase("Hae"), ST

51150         Case "MICRO"
                  Dim MicroSite As String
                  Dim SampleIDWithOffset As String
      '            MsgBox "Micro"
51160             MicroSite = GetOCMMapping("Site", "Cavan", tb!ProfileID)
51170             SaveSiteDetails SampleID, Dept, MicroSite, tb!TestDescription & "", grdSampleReceivedDate
51180             Select Case Code
                      Case "CS"
      '                    sql = "INSERT INTO [UrineRequests50] ([SampleID],[Request],[DateTimeOfRecord],[UserName]) " & _
      '                    "Values ('" & sampleid & "' ,'CS' ,'" & grdSampleReceivedDate & "' ,'')"
      '                    Cnxn(0).Execute Sql

51190             End Select
51200         Case "COAG"
51210             UpDateRequestCoag SampleID, Code
      '            MsgBox "COAG"
51220         Case "IMM"
51230             Analyser = AnalyserFor(Dept, Code)
51240             UpDateRequestImm SampleID, Code, ST, Analyser, 0

51250         End Select
51260         UpdateDemographic SampleID, grdSampleDate, grdSampleReceivedDate

51270     End If
51280     tb.MoveNext
51290 Loop

51300 UpdateRequestDetail SampleID


51310 Exit Sub
PlaceOrder_Error:
51320 LogError "frmScaneSample", "PlaceOrder", Erl, Err.Description, sql
End Sub
Private Sub SaveSiteDetails(ByVal SampleID As String, ByVal Dep As String, ByVal SampleType As String, ByVal SiteDetails As String, ByVal grdSampleReceivedDate As String)
51330 On Error GoTo SaveSiteDetails_Error

      Dim n As Integer
      Dim TestName As String
      Dim sql As String

51340 Cnxn(0).Execute "delete from siteDetails50 where sampleid= '" & 90031 & " ' and site = '" & SampleType & "'"

      'UpDateRequestsHae "Hae", TestName, Code

51350 sql = "INSERT INTO SiteDetails50" & vbNewLine
51360 sql = sql & "(SampleId, Site, SiteDetails, Username, DateTimeOfRecord) " & vbNewLine
51370 sql = sql & " VALUES ( " & vbNewLine
51380 sql = sql & " '" & SampleID & "', " & vbNewLine
51390 sql = sql & "        '" & SampleType & "', '" & SiteDetails & "' , getdate(), " & vbNewLine
51400 sql = sql & "         '" & Format(grdSampleReceivedDate, "dd/MM/yyyy HH:mm") & "' " & vbNewLine

51410 sql = sql & " )"

51420 Cnxn(0).Execute sql




51430 Exit Sub


SaveSiteDetails_Error:

      Dim strES As String
      Dim intEL As Integer

51440 intEL = Erl
51450 strES = Err.Description
51460 LogError "frmScaneSample", "SaveSiteDetails", intEL, strES
End Sub
Private Sub SaveHae(ByVal SampleID As String, ByVal Code As String, ByVal Dep As String, ByVal SampleType As String)
51470 On Error GoTo SaveHae_Error

      Dim n As Integer
      Dim TestName As String
      Dim sql As String



      'UpDateRequestsHae "Hae", TestName, Code

51480     sql = "INSERT INTO " & Dep & "Requests " & vbNewLine
51490     sql = sql & "(SampleId, Code, DateTimeOfRecord, SampleType, Analyser,Programmed) " & vbNewLine
51500     sql = sql & " VALUES ( " & vbNewLine
51510     sql = sql & " '" & SampleID & "', " & vbNewLine
51520     sql = sql & "        '" & Code & "' , getdate(), " & vbNewLine
51530     sql = sql & "        '" & SampleType & "' ,  " & vbNewLine
51540     sql = sql & "       'IPU',0 " & vbNewLine
51550     sql = sql & " )"

51560     Cnxn(0).Execute sql




51570 Exit Sub


SaveHae_Error:

      Dim strES As String
      Dim intEL As Integer

51580 intEL = Erl
51590 strES = Err.Description
51600 LogError "frmScaneSample", "SaveHae", intEL, strES
End Sub

Private Sub cmdExit_Click()
51610 Unload Me
End Sub

Private Sub flxSampleDetails_Click()
51620 On Error GoTo flxSampleDetails_Click_Error

      Dim sql As String

51630 If flxSampleDetails.MouseRow > 0 Then
51640     If flxSampleDetails.ColSel = 6 Or flxSampleDetails.ColSel = 7 Then

51650         LoadControls flxSampleDetails, txtDateTime
51660         Exit Sub
51670     End If
51680 End If

51690 If flxSampleDetails.Col = 9 Then
51700     If flxSampleDetails.Rows > 1 Then
51710         If MsgBox("Are you sure to delete this row ?", vbInformation + vbYesNo) = vbYes Then
      '            sql = "Delete from ocmRequestDetails Where ID = '" & flxSampleDetails.TextMatrix(flxSampleDetails.row, 8) & "'"
      '            Cnxn(0).Execute Sql
      ''            MsgBox sql
      '            flxSampleDetails.Rows = 1
      '            flxSampleDetails.row = 0
      '            If txtSampleID.Text = "" Then
      '                Call ShowRecords(m_SampleID, "1")
      '            Else
      '                Call ShowRecords(txtSampleID.Text, "0")
      '            End If
51720             Call flxSampleDetails.RemoveItem(flxSampleDetails.row)
51730             DoEvents
51740             DoEvents
51750         End If
51760     End If
51770 End If
51780 txtSampleID.SetFocus

51790 Exit Sub
flxSampleDetails_Click_Error:

51800 LogError "frmScaneSample", "flxSampleDetails_Click", Erl, Err.Description
End Sub

Private Sub flxSampleDetails_LeaveCell()
51810 If txtDateTime.Visible Then
51820     If IsDate(txtDateTime) Then
51830         flxSampleDetails.TextMatrix(flxSampleDetails.row, flxSampleDetails.Col) = txtDateTime
51840         txtDateTime.Visible = False
51850     Else
51860         iMsg "Date or Time is invalid"
51870         txtDateTime.SetFocus
51880     End If
          
51890 End If
End Sub

Private Sub flxSampleDetails_Scroll()
51900 txtDateTime.Visible = False
End Sub

Private Sub Form_Activate()
51910 On Error GoTo Form_Activate_Error
51920 If txtDateTime.Visible = False Then
51930     txtSampleID.SetFocus
51940 End If
51950 Exit Sub
Form_Activate_Error:

51960 LogError "frmScaneSample", "Form_Activate", Erl, Err.Description
End Sub

Private Sub Form_Load()
51970 On Error GoTo Form_Load_Error

51980 Call FormatGrid

51990 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

52000 intEL = Erl
52010 strES = Err.Description
52020 LogError "frmScan", "Form_Load", intEL, strES
End Sub

Private Sub txtsampleid_KeyPress(KeyAscii As Integer)
52030 On Error GoTo txtSampleID_KeyPress_Error

52040 If KeyAscii = 13 Then
52050     If txtSampleID.Text <> "" Then
52060         Call ShowRecords(txtSampleID.Text, "0")
52070         m_SampleID = txtSampleID.Text
52080         txtSampleID.Text = ""
52090         txtSampleID.SetFocus
52100     Else
52110         MsgBox "Please enter sample id.", vbInformation
52120     End If
52130 End If

52140 Exit Sub

txtSampleID_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

52150 intEL = Erl
52160 strES = Err.Description
52170 LogError "frmScan", "Form_Load", intEL, strES
End Sub

Public Sub ShowRecords(SampleID As String, AddOn As String)
52180 On Error GoTo ShowRecords_Error

      Dim sql As String
      Dim l_str As String
      Dim tb As Recordset
      Dim l_Count As Integer

52190 m_SampleID = SampleID
52200 For l_Count = 1 To flxSampleDetails.Rows - 1
52210     If flxSampleDetails.TextMatrix(l_Count, 1) = txtSampleID.Text Then
52220         MsgBox "Sample ID already exist.", vbInformation
52230         Exit Sub
52240     End If
52250 Next

      '    sql = "Select Distinct SampleID, PatName, IsNull(D.Sex,'M') Sex, IsNull(D.DOB,0) DOB, '' OrderingClinician, '' ProfileID, IsNull(D.RooH,0) RooH, IsNull(D.Urgent,0) Urgent, '' DepartmentID From ocmDemographic D "
      ''    Sql = Sql & "Inner Join ocmRequestDetails R ON R.SampleID = D.SampleID "
      ''    Sql = Sql & "Inner Join ocmRequest Q ON Q.RequestID = R.RequestID "
      '    sql = sql & "Where D.SampleID = " & txtSampleID.Text

52260 sql = "SELECT DISTINCT R.RequestID, RD.SampleID, R.Chart,R.PatName,R.Sex,R.DoB, R.OrderingClinician, Rd.ProfileID, R.RooH, R.Urgent,rd.sampledate " & _
            "FROM ocmRequest R " & _
            "LEFT JOIN ocmRequestDetails RD ON R.RequestID = RD.RequestID " & _
            "Where RD.SampleID = " & SampleID & " And RD.Programmed = 0 "
52270       If AddOn = "1" Then
52280         sql = sql & " And IsNULL(RD.Addon,'0') = '1'"
52290       Else
52300         sql = sql & " And IsNULL(RD.Addon,'0') = '0'"
52310       End If

52320 Set tb = New Recordset
52330 RecOpenServer 0, tb, sql
52340 If Not tb Is Nothing Then
52350     If Not tb.EOF Then
52360         While Not tb.EOF
52370             l_str = tb!RequestID & vbTab & tb!SampleID & vbTab & tb!PatName & vbTab & tb!DoB & vbTab & _
                  tb!ProfileID & vbTab & IIf(tb!Urgent = 1, "Yes", "No") & vbTab & _
                  Format(tb!SampleDate, "dd/MM/yyyy HH:mm") & vbTab & Format(Now, "dd/MM/YYYY HH:mm") & vbTab & "" & vbTab & "X"
52380             flxSampleDetails.AddItem (l_str)
52390             tb.MoveNext
52400         Wend
52410         For l_Count = 1 To flxSampleDetails.Rows - 1
52420             If flxSampleDetails.TextMatrix(l_Count, 5) = "Yes" Then
52430                 flxSampleDetails.Col = 5
52440                 flxSampleDetails.row = l_Count
52450                 flxSampleDetails.CellBackColor = &HFF&
52460             Else
52470                 flxSampleDetails.Col = 5
52480                 flxSampleDetails.row = l_Count
52490                 flxSampleDetails.CellBackColor = &HFF00&
52500             End If
52510         Next
52520     Else
52530         MsgBox "Sample ID not found.", vbInformation
52540     End If
52550 Else
52560     MsgBox "Sample ID not found.", vbInformation
52570 End If

52580 Exit Sub

ShowRecords_Error:

      Dim strES As String
      Dim intEL As Integer

52590 intEL = Erl
52600 strES = Err.Description
52610 LogError "frmScanSample", "ShowRecords", intEL, strES
End Sub


Private Sub UpDateRequestBio(ByVal SampleID As String, ByVal Code As String, _
                             ByVal SampleType As String, ByVal Analyser As String, Optional Gbottle As Integer)

      Dim sql As String

52620 On Error GoTo UpDateRequestBio_Error

52630 sql = "IF EXISTS(SELECT * FROM BioRequests WHERE SampleID = " & SampleID & " AND Code = '" & Code & "') "
52640 sql = sql & "UPDATE BioRequests SET Programmed = 0 WHERE SampleID = " & SampleID & " AND Code = '" & Code & "' "
52650 sql = sql & "ELSE "
52660 sql = sql & "INSERT INTO BioRequests " & _
            "(SampleID, Code, DateTime, SampleType, Programmed, AddOn, AnalyserID, Gbottle) " & _
            "Values (" & SampleID & ",'" & Code & "', getdate(),'" & SampleType & "', 0, 0 ,'" & Analyser & "', " & Gbottle & ") "
52670 Cnxn(0).Execute sql

52680 Exit Sub

UpDateRequestBio_Error:

      Dim strES As String
      Dim intEL As Integer

52690 intEL = Erl
52700 strES = Err.Description
52710 LogError "frmScaneSample", "UpDateRequestBio", intEL, strES, sql

End Sub

Private Sub UpDateRequestImm(ByVal SampleID As String, ByVal Code As String, _
                             ByVal SampleType As String, ByVal Analyser As String, Optional Gbottle As Integer)

      Dim sql As String

52720 On Error GoTo UpDateRequestImm_Error

52730 sql = "IF EXISTS(SELECT * FROM ImmRequests WHERE SampleID = " & SampleID & " AND Code = '" & Code & "') "
52740 sql = sql & "UPDATE ImmRequests SET Programmed = 0 WHERE SampleID = " & SampleID & " AND Code = '" & Code & "' "
52750 sql = sql & "ELSE "
52760 sql = sql & "INSERT INTO ImmRequests " & _
            "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID) " & _
            "Values (" & SampleID & ",'" & Code & "', getdate(),'" & SampleType & "',0,'" & Analyser & "') "
52770 Cnxn(0).Execute sql

52780 Exit Sub

UpDateRequestImm_Error:

      Dim strES As String
      Dim intEL As Integer

52790 intEL = Erl
52800 strES = Err.Description
52810 LogError "frmScaneSample", "UpDateRequestImm", intEL, strES, sql

End Sub


Private Sub UpDateRequestCoag(ByVal SampleID As String, ByVal Code As String)

      Dim sql As String

52820 On Error GoTo UpDateRequestCoag_Error

52830 sql = "IF NOT EXISTS(SELECT * FROM CoagRequests WHERE SampleID = " & SampleID & " AND Code = '" & Code & "') "
52840 sql = sql & "INSERT INTO CoagRequests " & _
            "(SampleID, Code) Values (" & SampleID & ",'" & Code & "') "
52850 Cnxn(0).Execute sql

52860 Exit Sub

UpDateRequestCoag_Error:

      Dim strES As String
      Dim intEL As Integer

52870 intEL = Erl
52880 strES = Err.Description
52890 LogError "frmScaneSample", "UpDateRequestCoag", intEL, strES, sql

End Sub

Private Sub UpdateDemographic(ByVal SampleID As String, ByVal grdSampleDate As String, ByVal grdSampleReceivedDate As String)

      Dim tb As Recordset
      Dim sql As String

52900 On Error GoTo UpdateDemographic_Error


52910 sql = "SELECT R.*, RD.* " & _
            "FROM ocmRequest R " & _
            "LEFT JOIN ocmRequestDetails RD ON R.RequestID = RD.RequestID " & _
            "Where RD.SampleID = '" & SampleID & "' And RD.Programmed = '0' "
52920 Set tb = New Recordset
52930 RecOpenServer 0, tb, sql

52940 If Not tb.EOF Then

52950     sql = "IF EXISTS(SELECT * FROM Demographics WHERE SampleID = '" & SampleID & "') "
52960     sql = sql & "UPDATE [dbo].[demographics] SET " & _
                "[SampleID] = '" & SampleID & "', " & _
                "[Chart] = '" & AddTicks(tb!Chart & "") & "', " & _
                "[PatName] = '" & AddTicks(tb!PatName & "") & "', " & _
                "[Age] = '" & CalcAge(tb!DoB, grdSampleDate) & "', " & _
                "[Sex] = '" & AddTicks(tb!Sex & "") & "', " & _
                "[RunDate] = '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "', " & _
                "[DoB] = '" & Format(tb!DoB, "dd/MMM/yyyy") & "', " & _
                "[Addr0] = '" & AddTicks(tb!Addr0 & "") & "', " & _
                "[Addr1] = '" & AddTicks(tb!Addr1 & "") & "', " & _
                "[Ward] = '" & AddTicks(tb!Ward & "") & "', " & _
                "[Clinician] = '" & AddTicks(tb!Clinician & "") & "', "
52970     sql = sql & "[SampleDate] = '" & Format(grdSampleDate, "dd/MMM/yyyy HH:mm:ss") & "', " & _
                "[ClDetails] = '" & AddTicks(tb!ClDetails & "") & "', " & _
                "[Hospital] = '" & AddTicks(tb!Hospital & "") & "', " & _
                "[RooH] = '" & tb!RooH & "', " & _
                "[DateTimeDemographics] = '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "', " & _
                "[AandE] = '" & AddTicks(tb!AandE & "") & "', " & _
                "[RecordDateTime] = '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "', " & _
                "[Operator] = '" & AddTicks(UserCode) & "', " & _
                "[Username] = '" & AddTicks(UserName) & "', " & _
                "[Urgent] = '" & tb!Urgent & "', " & _
                "[Valid] = '0', "
52980     sql = sql & "[ForMicro] = 0, " & _
                "[SentToEMedRenal] = 0, " & _
                "[SurName] = '', " & _
                "[ForeName] = '', " & _
                "[ExtSampleID] = '', " & _
                "[Healthlink] = '0' " & _
                " WHERE " & _
                "[SampleID] = '" & SampleID & "' "
52990     sql = sql & "ELSE "
53000     sql = sql & "INSERT INTO [dbo].[demographics] "
53010     sql = sql & "([SampleID], [Chart], [PatName], [Age], [Sex], [RunDate], [DoB], [Addr0], [Addr1], [Ward], [Clinician],  [SampleDate], [ClDetails], [Hospital], [RooH], [DateTimeDemographics], [AandE], [RecordDateTime], [Operator], [Username], [Urgent], [Valid], [ForMicro], [SentToEMedRenal], [SurName], [ForeName], [ExtSampleID], [Healthlink]) "
53020     sql = sql & " VALUES "
53030     sql = sql & "('" & SampleID & "', '" & AddTicks(tb!Chart & "") & "', '" & AddTicks(tb!PatName & "") & "', '" & CalcAge(tb!DoB, grdSampleDate) & "', '" & AddTicks(tb!Sex & "") & "', '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "', '" & Format(tb!DoB, "dd/MMM/yyyy HH:mm:ss") & "', '" & AddTicks(tb!Addr0 & "") & "', '" & AddTicks(tb!Addr1 & "") & "', '" & AddTicks(tb!Ward & "") & "', '" & AddTicks(tb!Clinician & "") & "',  '" & Format(grdSampleDate, "dd/MMM/yyyy HH:mm:ss") & "', '" & AddTicks(tb!ClDetails & "") & "', '" & AddTicks(tb!Hospital & "") & "','" & tb!RooH & "', '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "', '" & AddTicks(tb!AandE & "") & "', '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "', '" & AddTicks(UserCode) & "', '" & AddTicks(UserName) & "', '" & tb!Urgent & "', '0', "
53040     sql = sql & "0 , 0 , '', '', '', 0) "
53050     Cnxn(0).Execute sql




53060 End If

53070 Exit Sub

UpdateDemographic_Error:
53080 LogError "frmScaneSample", "UpdateDemographic", Erl, Err.Description, sql
End Sub

Private Sub UpdateRequestStatus(ByVal RequestID As String, ByVal RequestState As String)

      Dim sql As String
      Dim tb As Recordset

53090 On Error GoTo UpdateDemographic_Error

53100 sql = "UPDATE ocmRequest SET RequestState = '" & RequestState & "' WHERE RequestID = " & RequestID
53110 Cnxn(0).Execute sql


53120 Exit Sub
UpdateDemographic_Error:

53130 LogError "frmScaneSample", "UpdateDemographic", Erl, Err.Description, sql



End Sub

Private Sub UpdateRequestDetail(SampleID As String)
53140 On Error GoTo GetCode_Error

      Dim sql As String

53150 sql = "Update ocmRequestDetails Set [Programmed] = 1 Where [SampleID] = '" & SampleID & "'"
53160 Cnxn(0).Execute sql


53170 Exit Sub
GetCode_Error:
53180 LogError "frmScaneSample", "UpdateRequestDetail", Erl, Err.Description
End Sub

Private Sub txtDateTime_KeyUp(KeyCode As Integer, Shift As Integer)
53190 If KeyCode = vbKeyUp Then
          'GoOneRowUp
53200 ElseIf KeyCode = vbKeyDown Then
          'GoOneRowDown
53210 ElseIf KeyCode = 13 Then
53220     txtDateTime.Visible = False
53230 Else
53240     flxSampleDetails.TextMatrix(flxSampleDetails.row, flxSampleDetails.Col) = txtDateTime
      '60        If IsDate(txtDateTime) Then
      '70            flxSampleDetails.TextMatrix(flxSampleDetails.row, flxSampleDetails.Col) = txtDateTime
      '80        Else
      '90            iMsg "Date or Time is invalid"
      '100       End If
53250 End If
End Sub

Private Sub txtDateTime_LostFocus()
53260 txtDateTime.Visible = False
End Sub


