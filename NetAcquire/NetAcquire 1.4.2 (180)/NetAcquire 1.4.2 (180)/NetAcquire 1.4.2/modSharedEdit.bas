Attribute VB_Name = "modSharedEdit"
Option Explicit

Public Type PhoneLog
    SampleID As Long
    DateTime As Date
    PhonedTo As String
    PhonedBy As String
    Comment As String
    Discipline As String
End Type

'---------------------------------------------------------------------------------------
' Procedure : CheckTimes
' Author    : Masood
' Date      : 23/Oct/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function CheckTimes(ByRef SampleTime As MaskEdBox, _
                           ByRef DemographicComment As TextBox, _
                           ByRef ReceivedTime As MaskEdBox, Optional Hospital As String, Optional Ward As String) As Boolean

          Dim strTime As String

          'returns true if ok
20710   On Error GoTo CheckTimes_Error


20720     CheckTimes = True






20730     If Not IsDate(SampleTime) Then
20740         If (UCase(GetWardLocation(Ward)) = UCase("In-House")) And InStr(DemographicComment, "Sample Time Unknown.") = 0 Then
20750             strTime = iTIME("Sample Time?")
20760             If IsDate(strTime) Then
20770                 SampleTime = strTime
20780             Else
20790                 CheckTimes = False
20800                 Exit Function
20810             End If
20820         End If

20830         If InStr(DemographicComment, "Sample Time Unknown.") = 0 Then
20840             If iMsg("Is Sample Time unknown?", vbQuestion + vbYesNo) = vbYes Then
20850                 DemographicComment = DemographicComment & " Sample Time Unknown."
20860             Else
20870                 strTime = iTIME("Sample Time?")
20880                 If IsDate(strTime) Then
20890                     SampleTime = strTime
20900                 Else
20910                     CheckTimes = False
20920                     Exit Function
20930                 End If
20940             End If
20950         End If
          
20960     End If

20970     If Not IsDate(ReceivedTime) Then
20980         strTime = iTIME("Received Time?")
20990         If IsDate(strTime) Then
21000             ReceivedTime = strTime
21010         Else
21020             CheckTimes = False
21030             Exit Function
21040         End If
21050     End If


21060         If SampleTime = "00:00" Then
21070             iMsg "Sample Time 00:00 not allowed. Please change sample time", vbInformation
21080             CheckTimes = False
21090             Exit Function
21100         End If

       
21110 Exit Function

       
CheckTimes_Error:

      Dim strES As String
      Dim intEL As Integer

21120 intEL = Erl
21130 strES = Err.Description
21140 LogError "modSharedEdit", "CheckTimes", intEL, strES

End Function
Public Function CheckPhoneLog(ByVal SID As String) As PhoneLog

      'Returns PhoneLog.SampleID = 0 if no entry in phone log

      Dim tb As Recordset
      Dim sql As String
      Dim PL As PhoneLog

21150 On Error GoTo CheckPhoneLog_Error

21160 sql = "Select * from PhoneLog where " & _
            " Cast(SampleID as varchar (100)) = '" & Val(SID) & "'"
21170 Set tb = Cnxn(0).Execute(sql)
21180 If tb.EOF Then
21190     CheckPhoneLog.SampleID = 0
21200 Else
21210     With PL
21220         .SampleID = Val(SID)
21230         .Comment = tb!Comment & ""
21240         .DateTime = tb!DateTime
21250         .Discipline = tb!Discipline & ""
21260         .PhonedBy = tb!PhonedBy & ""
21270         .PhonedTo = tb!PhonedTo & ""
21280     End With
21290     CheckPhoneLog = PL
21300 End If

21310 Exit Function

CheckPhoneLog_Error:

      Dim strES As String
      Dim intEL As Integer

21320 intEL = Erl
21330 strES = Err.Description
21340 LogError "modSharedEdit", "CheckPhoneLog", intEL, strES, sql

End Function



Public Sub FillClinicians(ByVal cmb As ComboBox, ByVal HospitalName As String)

      Dim strHospitalCode As String
      Dim tb As Recordset
      Dim sql As String

21350 On Error GoTo FillClinicians_Error

21360 strHospitalCode = ListCodeFor("HO", HospitalName)

21370 sql = "SELECT DISTINCT Text, MIN(ListOrder) L FROM Clinicians WHERE " & _
            "HospitalCode = '" & strHospitalCode & "' " & _
            "AND InUse = 1 " & _
            "AND (Text IS NOT NULL AND Text <> '') " & _
            "GROUP BY Text " & _
            "ORDER BY L"
21380 Set tb = New Recordset
21390 RecOpenServer 0, tb, sql

21400 With cmb
21410     .Clear
21420     .AddItem ""
21430     Do While Not tb.EOF
21440         .AddItem ConvertNull(tb!Text, "") & ""
21450         tb.MoveNext
21460     Loop
21470 End With

21480 Exit Sub

FillClinicians_Error:

      Dim strES As String
      Dim intEL As Integer

21490 intEL = Erl
21500 strES = Err.Description
21510 LogError "modSharedEdit", "FillClinicians", intEL, strES, sql


End Sub
Public Sub FillWards(ByVal cmb As ComboBox, ByVal HospitalName As String)

      Dim strHospitalCode As String
      Dim tb As Recordset
      Dim sql As String

21520 On Error GoTo FillWards_Error

21530 strHospitalCode = ListCodeFor("HO", HospitalName)

21540 sql = "SELECT DISTINCT Text, MIN(ListOrder) L FROM Wards WHERE " & _
            "HospitalCode = '" & strHospitalCode & "' " & _
            "AND InUse = 1 " & _
            "GROUP BY Text " & _
            "ORDER BY L"
21550 Set tb = New Recordset
21560 RecOpenServer 0, tb, sql

21570 With cmb
21580     .Clear
21590     Do While Not tb.EOF
21600         .AddItem tb!Text & ""
21610         tb.MoveNext
21620     Loop
21630 End With

21640 cmb = "GP"

21650 Exit Sub

FillWards_Error:

      Dim strES As String
      Dim intEL As Integer

21660 intEL = Erl
21670 strES = Err.Description
21680 LogError "modSharedEdit", "FillWards", intEL, strES, sql

End Sub

Public Sub FillGPs(ByVal cmb As ComboBox, ByVal HospitalName As String)

      Dim strHospitalCode As String
      Dim GXs As New GPs
      Dim Gx As GP

21690 On Error GoTo FillGPs_Error

21700 strHospitalCode = ListCodeFor("HO", HospitalName)

21710 GXs.Load strHospitalCode, True
21720 cmb.Clear
21730 For Each Gx In GXs
21740     cmb.AddItem Gx.Text
21750 Next

21760 Exit Sub

FillGPs_Error:

      Dim strES As String
      Dim intEL As Integer

21770 intEL = Erl
21780 strES = Err.Description
21790 LogError "modSharedEdit", "FillGPs", intEL, strES

End Sub

Public Sub FillMRU(ByVal f As Form)

      Dim sql As String
      Dim tb As Recordset

21800 On Error GoTo FillMRU_Error

21810 sql = "Select top 10 * from MRU where " & _
            "UserCode = '" & AddTicks(UserCode) & "' " & _
            "Order by DateTime desc"
21820 Set tb = New Recordset
21830 RecOpenClient 0, tb, sql

21840 With f.cMRU
21850     .Clear
21860     Do While Not tb.EOF
21870         .AddItem Trim$(tb!SampleID & "")
21880         tb.MoveNext
21890     Loop
21900     If .ListCount > 0 Then
21910         .Text = ""
21920     End If
21930 End With

21940 Exit Sub

FillMRU_Error:

      Dim strES As String
      Dim intEL As Integer

21950 intEL = Erl
21960 strES = Err.Description
21970 LogError "modSharedEdit", "FillMRU", intEL, strES, sql

End Sub


Public Sub FlashNoPrevious(ByVal f As Form)

      Dim t As Single
      Dim n As Integer

21980 With f.lNoPrevious
21990     For n = 1 To 5
22000         .Visible = True
22010         .Refresh
22020         t = Timer
22030         Do While Timer - t < 0.1: DoEvents: Loop
22040         .Visible = False
22050         .Refresh
22060         t = Timer
22070         Do While Timer - t < 0.1: DoEvents: Loop
22080     Next
22090 End With

End Sub


Public Sub LoadPatientFromChart(ByVal f As Form, ByVal NewRecord As Boolean)

          Dim tbPatIF As Recordset
          Dim tbDemog As Recordset
          Dim sql As String
          Dim RooH As Boolean
          Dim strPatientEntity As String
          Dim X As Long
          Dim CurrentHospital As String
          Dim HospCode As String

22100     On Error GoTo LoadPatientFromChart_Error

22110     f.bViewBB.Enabled = False

22120     If InStr(UCase$(f.lblChartNumber), "CAVAN") Then
22130         CurrentHospital = "Cavan"
22140         strPatientEntity = "01"
22150     ElseIf InStr(UCase$(f.lblChartNumber), "MONAGHAN") Then
22160         strPatientEntity = "31"
22170         CurrentHospital = "Monaghan"
22180     Else
22190         strPatientEntity = ""
22200         CurrentHospital = HospName(0)
22210     End If

22220     HospCode = ListCodeFor("HO", CurrentHospital)

22230     sql = "Select * from PatientIFs where " & _
                "Chart = '" & AddTicks(f.txtChart) & "' "
22240     If strPatientEntity <> "" Then
22250         sql = sql & "and (Entity = '" & strPatientEntity & "' OR Entity = 'CGH')"
22260     End If
22270     Set tbPatIF = New Recordset
22280     RecOpenServer 0, tbPatIF, sql

22290     sql = "select top 1 * from demographics where " & _
                "Chart = '" & AddTicks(f.txtChart) & "' " & _
                "and Hospital = '" & CurrentHospital & "' " & _
                "order by RecordDateTime desc"
22300     Set tbDemog = New Recordset
22310     RecOpenServer 0, tbDemog, sql

22320     If tbPatIF.EOF And tbDemog.EOF Then

22330         f.txtSurName = ""
22340         f.txtForeName = ""
22350         f.txtAddress(0) = ""
22360         f.txtAddress(1) = ""
22370         f.txtSex = ""
22380         f.txtDoB = ""
22390         f.txtAge = ""
      '310           f.cmbWard = "GP"
      '320           f.cmbClinician = ""
      '330           f.cmbGP = ""
22400         f.txtDemographicComment = ""
22410         f.tSampleTime.Mask = ""
22420         f.tSampleTime.Text = ""
22430         f.tSampleTime.Mask = "##:##"
22440     ElseIf tbDemog.EOF Then
22450         With tbPatIF
22460             f.txtChart = !Chart & ""
                  '410           f.txtLabNo = !LabNo & ""
22470             f.txtSurName = SurName(Initial2Upper(!PatName & ""))
22480             f.txtForeName = ForeName(Initial2Upper(!PatName & ""))

22490             Select Case Left$(UCase$(!Sex & ""), 1)
                  Case "M": f.txtSex = "Male"
22500             Case "F": f.txtSex = "Female"
22510             Case Else: f.txtSex = ""
22520             End Select

22530             If Not IsNull(!DoB) Then
22540                 If IsDate(!DoB) Then
22550                     f.txtDoB = Format$(!DoB, "dd/MM/yyyy")
22560                 Else
22570                     f.txtDoB = ""
22580                 End If
22590             Else
22600                 f.txtDoB = ""
22610             End If
22620             f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
22630             f.cmbWard = GetWard(!Ward & "", HospCode)
22640             f.cmbClinician = GetClinician(!Clinician & "", HospCode)
22650             f.txtAddress(0) = Initial2Upper(!Address0 & "")
22660             f.txtAddress(1) = Initial2Upper(!Address1 & "")
22670         End With
22680     ElseIf tbPatIF.EOF Then
22690         If UCase(CurrentHospital) = "CAVAN" Then
22700             If iMsg("Are you sure this chart number is correct", vbQuestion + vbYesNo) = vbNo Then
22710                 f.txtChart = ""
22720                 f.txtChart.SetFocus
22730                 Exit Sub
22740             End If
22750         End If
22760         If NewRecord Then
22770             RooH = IsRoutine()
22780             f.cRooH(0) = RooH
22790             f.cRooH(1) = Not RooH
22800         Else
22810             If tbDemog!SampleID = f.txtSampleID Then
22820                 f.cRooH(0) = tbDemog!RooH
22830                 f.cRooH(1) = Not tbDemog!RooH
22840             End If
22850         End If
22860         f.txtSurName = SurName(tbDemog!PatName & "")
22870         f.txtForeName = ForeName(tbDemog!PatName & "")
22880         f.txtAddress(0) = tbDemog!Addr0 & ""
22890         f.txtAddress(1) = tbDemog!Addr1 & ""
22900         Select Case Left$(UCase$(tbDemog!Sex & ""), 1)
              Case "M": f.txtSex = "Male"
22910         Case "F": f.txtSex = "Female"
22920         Case Else: f.txtSex = ""
22930         End Select
22940         f.txtChart = tbDemog!Chart & ""
22950         If Val(tbDemog!LabNo & "") <> 0 Then
22960             f.txtLabNo = tbDemog!LabNo & ""
22970         End If
22980         f.txtAge = tbDemog!Age & ""
22990         f.txtDoB = Format$(tbDemog!DoB, "dd/mm/yyyy")
23000         f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
      '950           f.cmbWard = tbDemog!Ward & ""
      '960           f.cmbClinician = tbDemog!Clinician & ""
      '970           f.cmbGP = tbDemog!GP & ""
23010     Else
23020         If IsNull(tbDemog!RecordDateTime) Or IsNull(tbPatIF!DateTimeAmended) Then
23030             X = 0
23040         Else
23050             X = DateDiff("h", tbDemog!RecordDateTime, tbPatIF!DateTimeAmended)
23060         End If
23070         If X < 1 Then
23080             If NewRecord Then
23090                 RooH = IsRoutine()
23100                 f.cRooH(0) = RooH
23110                 f.cRooH(1) = Not RooH
23120             Else
23130                 If tbDemog!SampleID = f.txtSampleID Then
23140                     f.cRooH(0) = tbDemog!RooH
23150                     f.cRooH(1) = Not tbDemog!RooH
23160                 End If
23170             End If
23180             f.txtSurName = SurName(tbDemog!PatName & "")
23190             f.txtForeName = ForeName(tbDemog!PatName & "")
23200             f.txtAddress(0) = tbDemog!Addr0 & ""
23210             f.txtAddress(1) = tbDemog!Addr1 & ""
23220             Select Case Left$(UCase$(tbDemog!Sex & ""), 1)
                  Case "M": f.txtSex = "Male"
23230             Case "F": f.txtSex = "Female"
23240             Case Else: f.txtSex = ""
23250             End Select
23260             f.txtChart = tbDemog!Chart & ""
23270             If Val(tbDemog!LabNo & "") <> 0 Then
23280                 f.txtLabNo = tbDemog!LabNo & ""
23290             End If
23300             f.txtAge = tbDemog!Age & ""
23310             f.txtDoB = Format$(tbDemog!DoB, "dd/mm/yyyy")
23320             f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
      '1300              f.cmbWard = tbDemog!Ward & ""
      '1310              f.cmbClinician = tbDemog!Clinician & ""
      '1320              f.cmbGP = tbDemog!GP & ""
23330         Else
23340             With tbPatIF
23350                 f.txtChart = !Chart & ""
                      '1330              f.txtLabNo = !LabNo & ""
23360                 f.txtSurName = SurName(Initial2Upper(!PatName & ""))
23370                 f.txtForeName = ForeName(Initial2Upper(!PatName & ""))
23380                 Select Case Left$(UCase$(!Sex & ""), 1)
                      Case "M": f.txtSex = "Male"
23390                 Case "F": f.txtSex = "Female"
23400                 Case Else: f.txtSex = ""
23410                 End Select
23420                 If Not IsNull(!DoB) Then
23430                     If IsDate(!DoB) Then
23440                         f.txtDoB = Format$(!DoB, "dd/MM/yyyy")
23450                     Else
23460                         f.txtDoB = ""
23470                     End If
23480                 Else
23490                     f.txtDoB = ""
23500                 End If
23510                 f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
23520                 f.cmbWard = GetWard(!Ward & "", HospCode)
23530                 f.cmbClinician = GetClinician(!Clinician & "", HospCode)
23540                 f.txtAddress(0) = Initial2Upper(!Address0 & "")
23550                 f.txtAddress(1) = Initial2Upper(!Address1 & "")

23560             End With

23570             If Val(tbDemog!LabNo & "") <> 0 Then
23580                 f.txtLabNo = tbDemog!LabNo & ""
23590             End If
23600         End If
23610     End If

23620     If sysOptBloodBank(0) Then
23630         If Trim$(f.txtChart) <> "" Then
23640             sql = "Select * from PatientDetails where " & _
                        "PatNum = '" & f.txtChart & "' " & _
                        "and Name = '" & AddTicks(f.txtSurName & " " & f.txtForeName) & "'"
23650             Set tbDemog = New Recordset
23660             RecOpenClientBB 0, tbDemog, sql
23670             f.bViewBB.Enabled = Not tbDemog.EOF
23680         End If
23690     End If

23700     Exit Sub

LoadPatientFromChart_Error:

          Dim strES As String
          Dim intEL As Integer

23710     intEL = Erl
23720     strES = Err.Description
23730     LogError "modSharedEdit", "LoadPatientFromChart", intEL, strES, sql

End Sub



Public Sub LockDemographics(ByVal f As Form, ByVal Lockit As Boolean)

23740 With f
23750     .txtChart.Enabled = Not Lockit
23760     .txtForeName.Enabled = Not Lockit
23770     .txtSurName.Enabled = Not Lockit
23780     .txtDoB.Enabled = Not Lockit
23790     .txtAge.Enabled = Not Lockit
23800     .txtSex.Enabled = Not Lockit
23810     .cmbHospital.Enabled = Not Lockit
23820     .cmbWard.Enabled = Not Lockit
23830     .cmbClinician.Enabled = Not Lockit
23840     .cmbGP.Enabled = Not Lockit
23850     .cmdUnLock.Visible = Lockit
23860 End With

End Sub

Public Sub NameLostFocus(ByRef strSurName As String, _
                         ByRef strForeName As String, _
                         ByRef strSex As String)

      Dim tb As Recordset
      Dim sql As String
      Dim ForeName As String
      Dim Sex As String

23870 On Error GoTo NameLostFocus_Error

23880 strSurName = Replace(strSurName, ",", "")
23890 strForeName = Replace(strForeName, ",", "")

23900 ForeName = AddTicks(strForeName)
23910 Sex = UCase$(Left$(strSex, 1))

23920 If CBool(GetOptionSetting("DEMOGRAPHICSNAMECAPS", 0)) Then
23930     strSurName = UCase$(strSurName)
23940     strForeName = UCase$(strForeName)
23950 Else
23960     strSurName = Initial2Upper(strSurName)
23970     strForeName = Initial2Upper(strForeName)
23980 End If

23990 sql = "SELECT Sex FROM SexNames WHERE Name = '" & ForeName & "'"
24000 Set tb = New Recordset
24010 Set tb = Cnxn(0).Execute(sql)
24020 If Not tb.EOF Then
24030     Select Case UCase(tb!Sex & "")
              Case "M": strSex = "Male"
24040         Case "F": strSex = "Female"
24050     End Select
24060 End If

24070 Exit Sub

NameLostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

24080 intEL = Erl
24090 strES = Err.Description
24100 LogError "modSharedEdit", "NameLostFocus", intEL, strES, sql


End Sub


Public Sub UpdateMRU(ByVal f As Form)

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer
      Dim Found As Boolean
      Dim NewMRU(0 To 9, 0 To 1) As String
      '(x,0) SampleID
      '(x,1) DateTime

24110 On Error GoTo UpdateMRU_Error

24120 sql = "Select top 10 * from MRU where " & _
            "UserCode = '" & AddTicks(UserCode) & "' " & _
            "Order by DateTime desc"
24130 Set tb = New Recordset
24140 RecOpenClient 0, tb, sql

24150 n = -1
24160 Do While Not tb.EOF
24170     n = n + 1
24180     NewMRU(n, 0) = Trim$(tb!SampleID)
24190     NewMRU(n, 1) = tb!DateTime
24200     tb.MoveNext
24210 Loop

24220 Found = False
24230 For n = 0 To 9
24240     If f.txtSampleID = NewMRU(n, 0) Then
24250         sql = "Update MRU " & _
                    "Set DateTime = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' " & _
                    "where SampleID = '" & f.txtSampleID & "' " & _
                    "and UserCode = '" & AddTicks(UserCode) & "'"
24260         Cnxn(0).Execute sql
24270         Found = True
24280         Exit For
24290     End If
24300 Next

24310 If Not Found Then
24320     sql = "Delete from MRU where " & _
                "UserCode = '" & AddTicks(UserCode) & "'"
24330     Cnxn(0).Execute sql
24340     For n = 0 To 8
24350         If NewMRU(n, 0) <> "" Then
24360             sql = "Insert into MRU " & _
                        "(SampleID, DateTime, UserCode ) VALUES " & _
                        "('" & NewMRU(n, 0) & "', " & _
                        "'" & Format$(NewMRU(n, 1), "dd/mmm/yyyy hh:mm:ss") & "', " & _
                        "'" & AddTicks(UserCode) & "')"
24370             Cnxn(0).Execute sql
24380         End If
24390     Next
24400     sql = "Insert into MRU " & _
                "(SampleID, DateTime, UserCode ) VALUES " & _
                "('" & f.txtSampleID & "', " & _
                "'" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
                "'" & AddTicks(UserCode) & "')"
24410     Cnxn(0).Execute sql
24420 End If

24430 sql = "Select top 10 * from MRU where " & _
            "UserCode = '" & AddTicks(UserCode) & "' " & _
            "Order by DateTime desc"
24440 Set tb = New Recordset
24450 RecOpenClient 0, tb, sql

24460 With f.cMRU
24470     .Clear
24480     Do While Not tb.EOF
24490         .AddItem Trim$(tb!SampleID & "")
24500         tb.MoveNext
24510     Loop
24520     If .ListCount > 0 Then
24530         .Text = ""
24540     End If
24550 End With

24560 Exit Sub

UpdateMRU_Error:

      Dim strES As String
      Dim intEL As Integer

24570 intEL = Erl
24580 strES = Err.Description
24590 LogError "modSharedEdit", "UpdateMRU", intEL, strES, sql


End Sub
Public Sub LoadPatientFromOrderCom(ByVal f As Form, ByVal NewRecord As Boolean, Optional SampleIDExternal As String)

          Dim tbPatIF As Recordset
          Dim tbDemog As Recordset
          Dim tbGpOrd As Recordset
          Dim sql As String
          Dim RooH As Boolean
          Dim strPatientEntity As String
          Dim X As Long
          Dim CurrentHospital As String
          Dim HospCode As String
          
          Dim GPName As String

24600     On Error GoTo LoadPatientFromOrderCom_Error

24610     f.bViewBB.Enabled = False

24620     If InStr(UCase$(f.lblChartNumber), "CAVAN") Then
24630         CurrentHospital = "Cavan"
24640         strPatientEntity = "01"
24650     ElseIf InStr(UCase$(f.lblChartNumber), "MONAGHAN") Then
24660         strPatientEntity = "31"
24670         CurrentHospital = "Monaghan"
24680     Else
24690         strPatientEntity = ""
24700         CurrentHospital = HospName(0)
24710     End If

24720     HospCode = ListCodeFor("HO", CurrentHospital)

24730     If SampleIDExternal = "" Then
24740         sql = "Select * from GPOrderPatient where " & _
                    "GPNumber = '" & AddTicks(f.txtChart) & "' "
24750     Else
24760         sql = "Select * from GPOrderPatient where " & _
                  " SampleIDExternal = '" & SampleIDExternal & "' "
24770     End If

24780     Set tbPatIF = New Recordset
24790     RecOpenServer 0, tbPatIF, sql



24800     If tbPatIF.EOF Then

24810         f.txtSurName = ""
24820         f.txtForeName = ""
24830         f.txtAddress(0) = ""
24840         f.txtAddress(1) = ""
24850         f.txtSex = ""
24860         f.txtDoB = ""
24870         f.txtAge = ""
24880         f.cmbWard = "GP"
24890         f.cmbClinician = ""
24900         f.cmbGP = ""
24910         f.txtDemographicComment = ""
24920         f.tSampleTime.Mask = ""
24930         f.tSampleTime.Text = ""
24940         f.tSampleTime.Mask = "##:##"
24950     Else
24960         With tbPatIF
24970             f.txtChart = ""
                  '410           f.txtLabNo = !LabNo & ""
24980             f.txtSurName = Initial2Upper(!PatientSurName & "")
24990             f.txtForeName = Initial2Upper(!PatientForeName & "")

25000             Select Case Left$(UCase$(!Sex & ""), 1)
                  Case "M": f.txtSex = "Male"
25010             Case "F": f.txtSex = "Female"
25020             Case Else: f.txtSex = ""
25030             End Select

25040             If Not IsNull(!DoB) Then
25050                 If IsDate(!DoB) Then
25060                     f.txtDoB = Format$(!DoB, "dd/MM/yyyy")
25070                 Else
25080                     f.txtDoB = ""
25090                 End If
25100             Else
25110                 f.txtDoB = ""
25120             End If
25130             f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
                  

25140             f.txtAddress(0) = Initial2Upper(!Addr1 & "")
25150             f.txtAddress(1) = Initial2Upper(!addr2 & "")
                  
                 
25160             f.cmbWard = "GP"
25170             f.cmbGP = GetGPNameFromMcNumber(!PracticeID & "")
                  'f.cmbGP = !GPSurName & "" & !GPForeName
                  


25180             sql = "Select Top 1 *  from GPOrders where " & _
                      " SampleIDExternal = '" & SampleIDExternal & "' " & _
                      "  ORDER BY SampleDate DESC "
25190             Set tbGpOrd = New Recordset
25200             RecOpenServer 0, tbGpOrd, sql
25210             If tbGpOrd.EOF = False Then
                  
                  
25220                 f.dtSampleDate = Format$(tbGpOrd!SampleDate, "dd/mm/yyyy")
25230                 f.tSampleTime = Format$(tbGpOrd!SampleDate, "hh:mm")

25240                 f.dtRecDate = Format$(tbGpOrd!SampleDate, "dd/mm/yyyy")
      '                f.tRecTime = Format$(tbGpOrd!SampleDate, "hh:mm")' Masood 23-Oct-2015
                      
                      
25250                 If UCase(f.Name) = UCase("frmeditall") Then
25260                 f.cClDetails = tbGpOrd!ClinicalDetails
25270                 End If
25280             End If



25290         End With
25300     End If

          'If sysOptBloodBank(0) Then
          '    If Trim$(f.txtChart) <> "" Then
          '        sql = "Select * from PatientDetails where " & _
                   '              "PatNum = '" & f.txtChart & "' " & _
                   '              "and Name = '" & AddTicks(f.txtSurName & " " & f.txtForeName) & "'"
          '        Set tbDemog = New Recordset
          '        RecOpenClientBB 0, tbDemog, sql
          '        f.bViewBB.Enabled = Not tbDemog.EOF
          '    End If
          'End If
          '600       Call LoadGPOrders(f.txtSampleID, SampleIDExternal)
25310     Exit Sub

LoadPatientFromOrderCom_Error:

          Dim strES As String
          Dim intEL As Integer

25320     intEL = Erl
25330     strES = Err.Description
25340     LogError "modSharedEdit", "LoadPatientFromOrderCom", intEL, strES, sql

End Sub





'---------------------------------------------------------------------------------------
' Procedure : LoadGPOrders
' Author    : Masood
' Date      : 08/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LoadGPOrders(SampleID As String, SampleIDExternal As String)
          Dim sql As String
          Dim tb As Recordset


25350     On Error GoTo LoadGPOrders_Error




25360     sql = "SELECT     O.ShortName, O.LongName, O.SampleIDExternal, P.Department, P.NetAcquirePanel " & _
              " FROM GPOrders AS O INNER JOIN " & _
              " GpordersProfile AS P ON O.ShortName = P.GPTestCode " & _
              " where " & _
                "O.SampleIDExternal = '" & SampleIDExternal & "' "
25370     Set tb = New Recordset
25380     RecOpenClient 0, tb, sql

25390     Do While Not tb.EOF
25400         If tb!NetAcquirePanel <> "" And tb!Department = "Biochemistry" Then
25410             Call SaveBioRequest(SampleID, tb!NetAcquirePanel)
25420         Else
25430             MsgBox tb!ShortName
25440         End If
25450         tb.MoveNext
25460     Loop


25470     Exit Sub


LoadGPOrders_Error:

          Dim strES As String
          Dim intEL As Integer

25480     intEL = Erl
25490     strES = Err.Description
25500     LogError "modSharedEdit", "LoadGPOrders", intEL, strES, sql
End Sub



Private Sub SaveBioRequest(SampleID As String, PanelNetacquire As String)

          Dim Code As String
          Dim sql As String
          Dim tb As Recordset


25510     On Error GoTo SaveBio_Error

25520     Cnxn(0).Execute ("DELETE FROM BioRequests WHERE " & _
                           "SampleID = '" & SampleID & "' " & _
                           "AND Programmed = 0")

25530     sql = "SELECT * FROM Panels where PanelName ='" & PanelNetacquire & "' "


25540     sql = " SELECT     P.PanelName, P.Content, P.BarCode,D.longname,D.Code,D.SampleType"
25550     sql = sql & " FROM         Panels AS P INNER JOIN"
25560     sql = sql & " BioTestDefinitions AS D ON P.Content = D.shortname"
25570     sql = sql & " WHERE     P.PanelName = '" & PanelNetacquire & "'"



25580     Set tb = New Recordset
25590     RecOpenClient 0, tb, sql

25600     Do While Not tb.EOF
25610         UpDateRequests "Bio", tb!Code, tb!SampleType, SampleID, 0, 0
25620         tb.MoveNext
25630     Loop
25640     MsgBox "UPDATED"
25650     Exit Sub

          '







          'For n = 0 To lstSerumTests.ListCount - 1
          '    If lstSerumTests.Selected(n) Then
          '        Code = lstSerumCodes.List(n)
          '        Found = False
          '        For e = 0 To lstExistingBio.ListCount - 1
          '            If lstExistingBio.List(e) = Code Then
          '                Found = True
          '                Exit For
          '            End If
          '        Next
          '        If Not Found Then
          '            UpDateRequests "Bio", Code, "S", IIf(Code = FndOptionSettingGlucose(Code), chkGBottle.Value, 0)
          '        End If
          '    End If
          'Next
          '
          'For n = 0 To lstSweatTests.ListCount - 1
          '    If lstSweatTests.Selected(n) Then
          '        Code = lstSweatCodes.List(n)
          '        UpDateRequests "Bio", Code, "SW"
          '    End If
          'Next
          '
          'For n = 0 To lstBloodTests.ListCount - 1
          '    If lstBloodTests.Selected(n) Then
          '        Code = lstBloodCodes.List(n)
          '        UpDateRequests "Bio", Code, "B"
          '    End If
          'Next
          '
          'For n = 0 To lstUrineTests.ListCount - 1
          '    If lstUrineTests.Selected(n) Then
          '        Code = lstUrineCodes.List(n)
          '        UpDateRequests "Bio", Code, "U"
          '    End If
          'Next
          '
          'For n = 0 To lstCSFTests.ListCount - 1
          '    If lstCSFTests.Selected(n) Then
          '        Code = lstCSFCodes.List(n)
          '        UpDateRequests "Bio", Code, "C"
          '    End If
          'Next

25660     sql = "SELECT * FROM demographics WHERE " & _
                "SampleID = '" & SampleID & "'"
25670     Set tb = New Recordset
25680     RecOpenClient 0, tb, sql
25690     If tb.EOF Then
25700         tb.AddNew
25710         tb!Rundate = Format$(Now, "dd/mmm/yyyy")
25720         tb!SampleID = SampleID
25730         tb!FAXed = 0
25740         tb!RooH = 0
25750     End If
          'If chkUrgent.Value = 1 Then
          '    tb!Urgent = 1
          'Else
25760     tb!Urgent = 0
          'End If
25770     tb!Fasting = 0    'IIf(oSorF(1), 1, 0)
25780     tb.Update

25790     Exit Sub

SaveBio_Error:

          Dim strES As String
          Dim intEL As Integer

25800     intEL = Erl
25810     strES = Err.Description
25820     LogError "fNewOrder", "SaveBio", intEL, strES, sql

End Sub


Private Sub UpDateRequests(ByVal Discipline As String, _
                           ByVal Code As String, _
                           ByVal SampleType As String, SampleID As String, AddOn As String, Optional Gbottle As Integer)

      Dim sql As String

25830 On Error GoTo UpDateRequests_Error

25840 sql = "INSERT INTO " & Discipline & "Requests " & _
            "(SampleID, Code, DateTime, SampleType, Programmed, AddOn, AnalyserID,Gbottle) " & _
            "SELECT DISTINCT '" & SampleID & "', " & _
            "       '" & Code & "', getdate(), " & _
            "       '" & SampleType & "', '0', '" & AddOn & "', " & _
            "       Analyser ," & Gbottle & "  FROM " & Discipline & "TestDefinitions " & _
            "        " & _
            " WHERE Code = '" & Code & "' " & _
            " AND InUse = 1"
25850 Cnxn(0).Execute sql

25860 Exit Sub

UpDateRequests_Error:

      Dim strES As String
      Dim intEL As Integer

25870 intEL = Erl
25880 strES = Err.Description
25890 LogError "frmNewOrder", "UpDateRequests", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : DemographicsUniLabNoInsertValues
' Author    : Masood
' Date      : 02/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub DemographicsUniLabNoInsertValues(SampleID As String, User As String, PatName As String, DoB As String, Sex As String, Chart As String, LabNo As String)
          Dim sql As String
25900     On Error GoTo DemographicsUniLabNoInsertValues_Error

25910     sql = "insert into DemographicsUniLabNo (SampleID, [User], PatName, DoB, Sex, Chart, LabNo,DateTimeOfRecord ) Values (" & vbNewLine
25920     sql = sql & "'" & SampleID & "','" & User & "','" & AddTicks(PatName) & "','" & Format(DoB, "YYYY-MM-DD") & "','" & Sex & "','" & Chart & "'," & LabNo & ",'" & Format(Now, "YYYY-MM-DD hh:mm:ss") & "')"

25930     Cnxn(0).Execute sql

25940     Exit Sub


DemographicsUniLabNoInsertValues_Error:

          Dim strES As String
          Dim intEL As Integer

25950     intEL = Erl
25960     strES = Err.Description
25970     LogError "modSharedEdit", "DemographicsUniLabNoInsertValues", intEL, strES, sql
End Sub

Public Function GetGPNameFromMcNumber(ByVal McNumber As String) As String

      Dim sql As String
      Dim tb As Recordset

25980 On Error GoTo GetGPNameFromMcNumber_Error

25990 sql = "SELECT TOP 1 COALESCE([Text], '') [Text] FROM GPs WHERE MCNumber = '" & McNumber & "'"
26000 Set tb = New Recordset
26010 RecOpenServer 0, tb, sql
26020 If Not tb.EOF Then
26030     GetGPNameFromMcNumber = tb!Text
26040 End If


26050 Exit Function

GetGPNameFromMcNumber_Error:

       Dim strES As String
       Dim intEL As Integer

26060  intEL = Erl
26070  strES = Err.Description
26080  LogError "modSharedEdit", "GetGPNameFromMcNumber", intEL, strES, sql
          
End Function
