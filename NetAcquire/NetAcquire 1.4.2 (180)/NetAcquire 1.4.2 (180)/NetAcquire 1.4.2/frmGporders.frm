VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGporders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ss"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8640
      Begin VB.CommandButton cmdSaveOrders 
         Caption         =   "&Save"
         Height          =   765
         Left            =   6075
         Picture         =   "frmGporders.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4230
         Width           =   1155
      End
      Begin VB.CommandButton bcancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   765
         Left            =   7290
         Picture         =   "frmGporders.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4230
         Width           =   1155
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   2475
         TabIndex        =   4
         Top             =   390
         Width           =   3525
      End
      Begin VB.TextBox txtForeName 
         Height          =   285
         Left            =   165
         TabIndex        =   3
         Top             =   390
         Width           =   2295
      End
      Begin VB.TextBox txtAddr1 
         Height          =   285
         Left            =   4275
         TabIndex        =   2
         Top             =   750
         Width           =   4125
      End
      Begin VB.TextBox txtAddr0 
         Height          =   285
         Left            =   735
         TabIndex        =   1
         Top             =   750
         Width           =   3525
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   2805
         Left            =   135
         TabIndex        =   7
         Top             =   1230
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   4948
         _Version        =   393216
         Cols            =   16
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Surname"
         Height          =   195
         Left            =   2505
         TabIndex        =   14
         Top             =   150
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Forename"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   150
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   750
         Width           =   570
      End
      Begin VB.Label lblSex 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7695
         TabIndex        =   11
         Top             =   390
         Width           =   705
      End
      Begin VB.Label lblDoB 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6075
         TabIndex        =   10
         Top             =   390
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   28
         Left            =   6075
         TabIndex        =   9
         Top             =   150
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   48
         Left            =   7695
         TabIndex        =   8
         Top             =   150
         Width           =   270
      End
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   9360
      Picture         =   "frmGporders.frx":1FEC
      Top             =   630
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   9360
      Picture         =   "frmGporders.frx":22C2
      Top             =   420
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmGporders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sSampleIDExt As String
Private m_sSampleID As String

Private m_sClinicalDetails As String

Private m_objEditScreen As Form

Private m_sDisiplinesQuery As String

Private m_sMicroSite As String

Private Sub Form_Unload(Cancel As Integer)

54220     Set m_objEditScreen = Nothing

End Sub


Public Property Get SampleIDExt() As String

54230     SampleIDExt = m_sSampleIDExt

End Property

Public Property Let SampleIDExt(ByVal sSampleIDExt As String)

54240     m_sSampleIDExt = sSampleIDExt

End Property

'---------------------------------------------------------------------------------------
' Procedure : bCancel_Click
' Author    : Masood
' Date      : 09/Apr/2015
' Purpose   :s
'---------------------------------------------------------------------------------------
'
Private Sub bcancel_Click()
54250     On Error GoTo bCancel_Click_Error


54260     If UCase(EditScreen.Caption) = UCase("frmeditall") Then
54270         EditScreen.CancelFromGpCom = True
54280     End If

54290     Unload Me




54300     Exit Sub


bCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

54310     intEL = Erl
54320     strES = Err.Description
54330     LogError "frmGporders", "bCancel_Click", intEL, strES
End Sub



Private Sub cmdSaveOrders_Click()
54340     On Error GoTo cmdSaveOrders_Click_Error
          Dim i As Integer
          Dim DisiplineName As String
          Dim DisiplinePanelName As String
          Dim OrderTestHaem As String


54350     With g
54360         For i = 1 To .Rows - 1
54370             .row = i
54380             .Col = 9
54390             If (UCase(.TextMatrix(i, 8)) = UCase("biochemistry") Or UCase(.TextMatrix(i, 8)) = UCase("Coagulation")) And .CellPicture = imgGreenTick And .TextMatrix(i, 6) <> "" Then

54400                 If UCase(.TextMatrix(i, 8)) = UCase("biochemistry") Then
54410                     DisiplineName = "Bio"
54420                     DisiplinePanelName = ""
54430                 ElseIf UCase(.TextMatrix(i, 8)) = UCase("Coagulation") Then
54440                     DisiplineName = "Coag"
54450                     DisiplinePanelName = "Coag"
54460                 End If

54470                 Call SaveRequest(SampleID, .TextMatrix(i, 6), .TextMatrix(i, 7), DisiplineName, DisiplinePanelName, .TextMatrix(i, 0))

54480             ElseIf UCase(.TextMatrix(i, 8)) = UCase("Haematology") And .CellPicture = imgGreenTick Then
54490                 Call SaveRequestHaem(.TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 0))

54500             ElseIf UCase(.TextMatrix(i, 8)) = UCase("External") And .CellPicture = imgGreenTick Then
54510                 Call FindExternalDetail(.TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 0))
54520             ElseIf UCase(.TextMatrix(i, 8)) = UCase("Microbiology") And .CellPicture = imgGreenTick Then
54530                 Call SaveRequestHaemMicro(.TextMatrix(i, 6), "", "", IIf((UCase(MicroSite) = UCase("Urine")), True, False), .TextMatrix(i, 0))
54540             End If
54550         Next i

54560     End With

54570     If SaveDemographics = True Then
54580         Call LoadPatientFromOrderCom(EditScreen, False, SampleIDExt)

54590         If UCase(EditScreen.Name) = UCase("frmeditall") Then
                  'EditScreen.cmdSaveDemographics.Value = True
54600             EditScreen.SavedDemoFromGPCom = True
54610         Else
54620             EditScreen.cmdSaveDemographics.Value = True
54630         End If
54640     Else
54650         EditScreen.LoadAllDetails
54660     End If
54670     Unload Me

54680     Exit Sub


cmdSaveOrders_Click_Error:

          Dim strES As String
          Dim intEL As Integer

54690     intEL = Erl
54700     strES = Err.Description
54710     LogError "frmGporders", "cmdSaveOrders_Click", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SaveMicro
' Author    : Masood
' Date      : 22/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub SaveRequestHaemMicro(ByVal TestName As String, Analyser As String, Programmed As String, IsUrine As Boolean, TestShortNameGP As String)

54720     On Error GoTo SaveMicro_Error

          Dim sql As String
          Dim SampleIDWithOffset As String
          '+++ Junaid 20-05-2024
          '20        SampleIDWithOffset = Val(SampleID) + Val(sysOptMicroOffset(0))
54730     SampleIDWithOffset = Val(SampleID)
          '--- Junaid
          ' If Panel is true the urine else Faeces
54740     If IsUrine = True Then

54750         sql = "Delete from UrineRequests50 where " & _
                  "SampleID = '" & SampleIDWithOffset & "' AND Request = '" & TestName & "'"
54760         Cnxn(0).Execute sql

54770         sql = "INSERT INTO UrineRequests50 (SampleID, Request, UserName) " & _
                  "VALUES " & _
                  "('" & SampleIDWithOffset & "', " & _
                  " '" & TestName & "', " & _
                  " '" & AddTicks(UserName) & "')"

54780         sql = sql & " UPDATE GPOrders SET Programmed = 1 WHERE SHORTNAME = '" & TestShortNameGP & "' AND SampleIDExternal = '" & SampleIDExt & "'"

54790         Cnxn(0).Execute sql

54800     End If


54810     If IsUrine = False Then

54820         sql = "Delete from FaecesRequests50 where " & _
                  "SampleID = '" & SampleIDWithOffset & "' AND Request = '" & TestName & "'"
54830         Cnxn(0).Execute sql


54840         sql = "INSERT INTO FaecesRequests50 (SampleID, Request, UserName, Analyser, Programmed) " & _
                  "VALUES " & _
                  "('" & SampleIDWithOffset & "', " & _
                  " '" & TestName & "', " & _
                  " '" & UserName & "', " & _
                  " '" & Analyser & "', " & _
                  " '" & Programmed & "')"
54850         sql = sql & " UPDATE GPOrders SET Programmed = 1 WHERE SHORTNAME = '" & TestShortNameGP & "' AND SampleIDExternal = '" & SampleIDExt & "'"

54860         Cnxn(0).Execute sql
54870     End If



54880     Exit Sub


SaveMicro_Error:

          Dim strES As String
          Dim intEL As Integer

54890     intEL = Erl
54900     strES = Err.Description
54910     LogError "frmGporders", "SaveMicro", intEL, strES, sql
End Sub

Private Sub FindExternalDetail(PanelName As String, IsPanel As Boolean, TestShortNameGP As String)

54920     On Error GoTo FndExtPanel_Error
          Dim sql As String
          Dim tb As New ADODB.Recordset

54930     If IsPanel = True Then
54940         sql = " SELECT     D.AnalyteName,D.SendTo,D.MBCode,BiomnisCode , P.PanelName"
54950         sql = sql & " FROM         ExternalDefinitions AS D INNER JOIN"
54960         sql = sql & "         ExtPanels AS P ON D.AnalyteName = P.[Content]"
54970         sql = sql & "  WHERE P.PanelName ='" & PanelName & "'"
54980     Else
54990         sql = "SELECT  D.AnalyteName ,D.SendTo,D.MBCode,BiomnisCode "
55000         sql = sql & " FROM         ExternalDefinitions AS D"
55010         sql = sql & " WHERE D.AnalyteName = '" & PanelName & "'"
55020     End If


55030     Set tb = New Recordset
55040     RecOpenClient 0, tb, sql

55050     Do While Not tb.EOF
55060         Call SaveRequestExternal(tb!AnalyteName, tb!SendTo, TestShortNameGP)
55070         tb.MoveNext
55080     Loop

55090     Exit Sub


FndExtPanel_Error:

          Dim strES As String
          Dim intEL As Integer

55100     intEL = Erl
55110     strES = Err.DescriptionS
55120     LogError "frmGporders", "FndExtPanel", intEL, strES, sql
End Sub

Private Sub SaveRequestExternal(AnalyteName As String, Destination As String, TestShortNameGP As String)
          Dim sql As String

55130     On Error GoTo SaveRequestExternal_Error


55140     If ListCodeFor("BiomnisEnableLabs", Destination) <> "" Then
              'Order to biomnis requests
55150         sql = "DECLARE @Code nvarchar(50) " & _
                  "DECLARE @Dept nvarchar(50) " & _
                  "DECLARE @SampleType nvarchar(50) " & _
                  "SET @Code = (SELECT BiomnisCode FROM ExternalDefinitions " & _
                  "             WHERE AnalyteName = '" & AnalyteName & "') "

55160         sql = sql & _
                  "SET @Dept = (SELECT Department FROM ExternalDefinitions " & _
                  "             WHERE AnalyteName = '" & AnalyteName & "') "

55170         sql = sql & _
                  "SET @SampleType = (SELECT SampleType FROM ExternalDefinitions " & _
                  "             WHERE AnalyteName = '" & AnalyteName & "') "
55180         sql = sql & _
                  "IF EXISTS(SELECT * FROM ExternalDefinitions " & _
                  "          WHERE AnalyteName = '" & AnalyteName & "' " & _
                  "          AND COALESCE(BiomnisCode, '') <> '') " & _
                  "  BEGIN IF NOT EXISTS(SELECT * FROM BiomnisRequests " & _
                  "                  WHERE SampleID = '" & SampleID & "' " & _
                  "                  AND TestCode = @Code)" & _
                  "      INSERT INTO BiomnisRequests (SampleID, TestCode, TestName, SampleType, SampleDateTime, Department, RequestedBy, SendTo, Status) " & _
                  "      VALUES " & _
                  "     ('" & SampleID & "', " & _
                  "      @Code, " & _
                  "      '" & AnalyteName & "', " & _
                  "      @SampleType, " & _
                  "      '" & Format(Now, "dd/MMM/yyyy hh:mm:ss") & "', " & _
                  "      @Dept, " & _
                  "      '" & UserCode & "^" & UserName & "', " & _
                  "      '" & Destination & "', " & _
                  "      'OutStanding') END"

55190         sql = sql & " UPDATE GPOrders SET Programmed = 1 WHERE ShortName = '" & TestShortNameGP & "' AND SampleIDExternal = '" & SampleIDExt & "'"
55200         Cnxn(0).Execute sql
55210     Else
              'Order to other external laboratories
55220         If ListCodeFor("MBEnableLabs", Destination) <> "" Then
55230             sql = "DECLARE @Code nvarchar(50) " & _
                      "DECLARE @Dept nvarchar(50) " & _
                      "DECLARE @SampleType nvarchar(50) " & _
                      "SET @Code = (SELECT MBCode FROM ExternalDefinitions " & _
                      "             WHERE AnalyteName = '" & AnalyteName & "') " & _
                      "SET @Dept = (SELECT Department FROM ExternalDefinitions " & _
                      "             WHERE AnalyteName = '" & AnalyteName & "') " & _
                      "SET @SampleType = (SELECT SampleType FROM ExternalDefinitions WHERE AnalyteName = '" & AnalyteName & "') " & _
                      "IF EXISTS(SELECT * FROM ExternalDefinitions " & _
                      "          WHERE AnalyteName = '" & AnalyteName & "' " & _
                      "          AND COALESCE(MBCode, '') <> '') " & _
                      "  BEGIN IF NOT EXISTS(SELECT * FROM MediBridgeRequests " & _
                      "                  WHERE SampleID = '" & SampleID & "' " & _
                      "                  AND TestCode = @Code)" & _
                      "      INSERT INTO MedibridgeRequests (SampleID, TestCode, TestName, SampleDateTime, ClinDetails, Orderer, Dept, SpecimenSource, Status ) " & _
                      "      VALUES " & _
                      "     ('" & SampleID & "', " & _
                      "      @Code, " & _
                      "      '" & AnalyteName & "', " & _
                      "      '" & Format(Now, "dd/MMM/yyyy hh:mm:ss") & "', " & _
                      "      '" & ClinicalDetails & "', " & _
                      "      '" & UserCode & "^" & UserName & "', " & _
                      "      @Dept, " & _
                      "      '" & Destination & "', " & _
                      "      'Requested') END"
55240             sql = sql & " UPDATE GPOrders SET Programmed = 1 WHERE ShortName = '" & TestShortNameGP & "' AND SampleIDExternal = '" & SampleIDExt & "'"
55250             Cnxn(0).Execute sql
55260         End If
55270     End If


55280     sql = "IF NOT EXISTS (SELECT * FROM ExtResults " & _
              "             WHERE SampleID = '" & SampleID & "' " & _
              "             AND Analyte = '" & AnalyteName & "') " & _
              "    INSERT INTO ExtResults " & _
              "    (SampleID, Analyte, Result, SendTo, Units, SentDate) " & _
              "    SELECT '" & SampleID & "',  '" & AnalyteName & "', '', SendTo, Units, GETDATE() " & _
              "    FROM ExternalDefinitions WHERE AnalyteName = '" & AnalyteName & "'"
55290     Cnxn(0).Execute sql

55300     Exit Sub


SaveRequestExternal_Error:

          Dim strES As String
          Dim intEL As Integer

55310     intEL = Erl
55320     strES = Err.Description
55330     LogError "frmGporders", "SaveRequestExternal", intEL, strES, sql
End Sub





'---------------------------------------------------------------------------------------
' Procedure : SaveRequestHaem
' Author    : Masood
' Date      : 14/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveRequestHaem(TestName As String, IsPanel As Boolean, TestShortNameGpOrder As String)
55340     On Error GoTo SaveRequestHaem_Error
          Dim sql As String
          Dim tb As ADODB.Recordset
          Dim HaemColumforUpd As String

          ' IF panel is true then request for same as gporder , if panel is false then update feild of Results


55350     If IsPanel = True Then
55360         sql = "SELECT * FROM HaeRequests WHERE " & _
                  "SampleID = '" & SampleID & "'"
55370         Set tb = New Recordset
55380         RecOpenServer 0, tb, sql
55390         If tb.EOF Then
55400             tb.AddNew
55410         End If
55420         tb!SampleID = Val(SampleID)
55430         tb!Code = TestName
55440         tb!Programmed = 0
55450         tb!SampleType = "Blood EDTA"
55460         tb!Analyser = "IPU"
55470         tb!UserName = UserName
55480         tb.Update
55490         tb.Close
55500     End If

55510     If IsPanel = False Then
55520         sql = "IF NOT EXISTS (SELECT * FROM HaemResults WHERE " & _
                  "           SampleID = '" & SampleID & "') " & _
                  "  INSERT INTO HaemResults (SampleID," & TestName & ") VALUES " & _
                  "  ('" & SampleID & "',1) " & _
                  " Else " & _
                  " UPDATE HaemResults SET " & TestName & " = 1 where SampleID = '" & SampleID & "'"
55530         Cnxn(0).Execute (sql)
55540     End If

          '        If UCase(TestName) = UCase("ESR") Then
          '            HaemColumforUpd = "cESR"
          '        ElseIf UCase(TestName) = UCase("Malaria Screen") Then
          '            HaemColumforUpd = "cMalaria"
          '        ElseIf UCase(TestName) = UCase("Monospot") Then
          '            HaemColumforUpd = "cMonospot"
          '        ElseIf UCase(TestName) = UCase("Sickle cell anaemia") Then
          '            HaemColumforUpd = "cSickledex"
          '        End If
          '
          '        If HaemColumforUpd <> "" Then
          '            sql = "UPDATE HaemResults SET " & HaemColumforUpd & " = 1 WHERE SampleID = '" & SampleID & "'"
          '            Cnxn(0).Execute Sql
          '        End If
          '    End If

55550     sql = " UPDATE GPOrders SET Programmed = 1 WHERE ShortName = '" & TestShortNameGpOrder & "' AND SampleIDExternal = '" & SampleIDExt & "'"
55560     Cnxn(0).Execute sql









          '150   If SaveDemoghrapic = True Then
          '160     sql = "SELECT * FROM demographics WHERE " & _
          '              "SampleID = '" & SampleID & "'"
          '170     Set tb = New Recordset
          '180     RecOpenClient 0, tb, sql
          '190     If tb.EOF Then
          '200         tb.AddNew
          '210         tb!Rundate = Format$(Now, "dd/mmm/yyyy")
          '220         tb!SampleID = SampleID
          '230         tb!FAXed = 0
          '240         tb!RooH = 0
          '250     End If
          '260     tb!Urgent = 0
          '270     tb!Fasting = 0    'IIf(oSorF(1), 1, 0)
          '280     tb.Update
          '290   End If

55570     Exit Sub


SaveRequestHaem_Error:

          Dim strES As String
          Dim intEL As Integer

55580     intEL = Erl
55590     strES = Err.Description
55600     LogError "frmGporders", "SaveRequestHaem", intEL, strES, sql
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Masood
' Date      : 09/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

55610     On Error GoTo Form_Load_Error

55620     With Me
55630         .Caption = "Gp Orders"
55640     End With
55650     GridHead
55660     LoadPatientGp (SampleIDExt)

55670     Exit Sub


Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

55680     intEL = Erl
55690     strES = Err.Description
55700     LogError "frmGporders", "Form_Load", intEL, strES
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadPatientGp
' Author    : Masood
' Date      : 09/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LoadPatientGp(SampleIDExternal As String)
          Dim sql As String
          Dim tb As ADODB.Recordset

55710     On Error GoTo LoadPatientGp_Error

55720     ClearDemographic

55730     sql = "SELECT     GPName, GPNumber, DateTimeOfMessage, PatientID, PatientSurName, PatientForeName, DoB, Sex, "
55740     sql = sql & " Addr1, Addr2, Addr3, Addr4, Addr5, PracticeID, GPSurName, "
55750     sql = sql & " GPForeName , Pregnant, GID, FileName, SampleIDExternal"
55760     sql = sql & " FROM         GPOrderPatient"
55770     sql = sql & " WHERE SampleIDExternal = '" & SampleIDExternal & "'"

55780     Set tb = New Recordset
55790     RecOpenClient 0, tb, sql

55800     Do While Not tb.EOF

55810         txtForeName.Text = tb!PatientSurName
55820         txtSurName.Text = tb!PatientForeName
55830         lblDoB = Format(tb!DoB, "dd/mm/yyyy")
55840         txtAddr0.Text = tb!Addr1
55850         txtAddr1.Text = tb!addr2
55860         lblSex = tb!Sex
55870         tb.MoveNext
55880     Loop

55890     If txtForeName <> "" Then
55900         LoadGPOrders (SampleIDExternal)
55910     End If
55920     Exit Sub


LoadPatientGp_Error:

          Dim strES As String
          Dim intEL As Integer

55930     intEL = Erl
55940     strES = Err.Description
55950     LogError "frmGporders", "LoadPatientGp", intEL, strES, sql
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ClearDemographic
' Author    : Masood
' Date      : 15/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ClearDemographic()

55960     On Error GoTo ClearDemographic_Error

55970     txtForeName.Text = ""
55980     txtSurName.Text = ""
55990     lblDoB = ""
56000     txtAddr0.Text = ""
56010     txtAddr1.Text = ""
56020     lblSex = ""




56030     Exit Sub


ClearDemographic_Error:

          Dim strES As String
          Dim intEL As Integer

56040     intEL = Erl
56050     strES = Err.Description
56060     LogError "frmGporders", "ClearDemographic", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GridHead
' Author    : Masood
' Date      : 09/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub GridHead()
56070     On Error GoTo GridHead_Error
56080     With g
56090         .Cols = 10
56100         .Rows = 1
56110         .FormatString = " Short Name     |<Long Name                          |<Clinical Details     |<Sample Type Code     |<Sample Type   |<Priority  |<Netacquire Panel |<IsPanel |<Department "

56120         .ColWidth(0) = 2000
56130         .ColAlignment(0) = 1

56140         .ColWidth(1) = 3300

56150         .ColWidth(2) = 0
56160         .ColWidth(3) = 0
56170         .ColWidth(4) = 0
56180         .ColWidth(5) = 0
56190         .ColWidth(6) = 0
56200         .ColWidth(7) = 0
56210         .ColWidth(8) = 2400
56220         .ColWidth(9) = imgRedCross.width
56230     End With
56240     Exit Sub


GridHead_Error:

          Dim strES As String
          Dim intEL As Integer

56250     intEL = Erl
56260     strES = Err.Description
56270     LogError "frmGporders", "GridHead", intEL, strES
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadGPOrders
' Author    : Masood
' Date      : 09/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LoadGPOrders(SampleIDExternal As String)

56280     On Error GoTo LoadGPOrders_Error

          Dim sql As String
          Dim tb As ADODB.Recordset
          Dim s As String
56290     GridHead

          '30        sql = "SELECT     * "
          '40        sql = sql & " FROM         GPOrders"
          '50        sql = sql & " WHERE SampleIDExternal = '" & SampleIDExternal & "'"




56300     sql = "SELECT     O.ShortName, O.LongName, O.ClinicalDetails, O.SampleTypeCode, O.SampleType,O.Priority,O.SampleIDExternal,isnull(P.Panel ,0) as Panel " & vbNewLine
56310     sql = sql & " ,P.Department, P.NetAcquirePanel " & vbNewLine
56320     sql = sql & " FROM GPOrders AS O LEFT OUTER JOIN " & vbNewLine
56330     sql = sql & " GpordersProfile AS P ON O.ShortName = P.GPTestCode " & vbNewLine
56340     sql = sql & " where " & vbNewLine
56350     sql = sql & " ISNULL(Programmed,0)  = 0 and " & vbNewLine
56360     sql = sql & " O.SampleIDExternal = '" & SampleIDExternal & "' " & vbNewLine
56370     sql = sql & " " & DisiplinesQuery & vbNewLine




56380     Set tb = New Recordset
56390     RecOpenClient 0, tb, sql

56400     Do While Not tb.EOF
56410         s = tb!ShortName & vbTab & tb!LongName & vbTab & tb!ClinicalDetails & vbTab & tb!SampleTypeCode & vbTab & tb!SampleType & vbTab & tb!Priority & vbTab & tb!NetAcquirePanel & vbTab & tb!Panel & vbTab & tb!Department
56420         g.AddItem (s)
56430         g.row = g.Rows - 1
56440         g.Col = 9
56450         Set g.CellPicture = imgRedCross.Picture

56460         tb.MoveNext
56470     Loop

56480     Exit Sub


LoadGPOrders_Error:

          Dim strES As String
          Dim intEL As Integer

56490     intEL = Erl
56500     strES = Err.Description
56510     LogError "frmGporders", "LoadGPOrders", intEL, strES, sql
End Sub

'---------------------------------------------------------------------------------------
' Procedure : g_Click
' Author    : Masood
' Date      : 09/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub g_Click()

56520     On Error GoTo g_Click_Error

56530     With g

56540         If .ColSel = 9 And .TextMatrix(.RowSel, 6) <> "" Then

56550             If .CellPicture = imgGreenTick.Picture Then
56560                 Set .CellPicture = imgRedCross.Picture
56570             Else
56580                 Set .CellPicture = imgGreenTick.Picture
56590             End If
56600         End If
56610     End With


56620     Exit Sub


g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

56630     intEL = Erl
56640     strES = Err.Description
56650     LogError "frmGporders", "g_Click", intEL, strES
End Sub

Public Property Get SampleID() As String

56660     SampleID = m_sSampleID

End Property

Public Property Let SampleID(ByVal sSampleID As String)

56670     m_sSampleID = sSampleID

End Property


Private Sub SaveRequest(SampleID As String, PanelNetacquire As String, IsPanel As Boolean, Disipline As String, DisiplinePanel As String, TestShortNameGP As String)

          Dim Code As String
          Dim sql As String
          Dim tb As Recordset
          Dim Gbottle As Integer


56680     On Error GoTo SaveBio_Error


56690     Gbottle = 0

56700     If IsPanel = True Then
56710         sql = " SELECT     P.PanelName, P.Content, P.BarCode,D.longname,D.Code,D.SampleType"
56720         sql = sql & " FROM " & DisiplinePanel & "Panels AS P INNER JOIN"
56730         sql = sql & " " & Disipline & "TestDefinitions AS D ON P.Content = D.shortname"
56740         sql = sql & " WHERE     P.PanelName = '" & PanelNetacquire & "'"
56750     Else
56760         sql = "select TOP 1 D.Code,D.SampleType from " & Disipline & "TestDefinitions AS D "
56770         sql = sql & " WHERE D.Code = '" & PanelNetacquire & "'"
56780     End If

56790     If Disipline = "Coag" Then
56800         sql = Replace(sql, "D.SampleType", "'' as SampleType ")
56810     End If


56820     Set tb = New Recordset
56830     RecOpenClient 0, tb, sql

56840     Do While Not tb.EOF
56850         If UCase(Disipline) = "BIO" And (FndOptionSettingGlucose(tb!Code & "") <> "") Then
56860             If GetOptionSetting("DisableGBottleDetection", 0) = 1 Then
56870                 Gbottle = 0
56880             Else
56890                 Gbottle = 1
56900             End If
56910         End If

56920         UpDateRequests Disipline, tb!Code, tb!SampleType, SampleID, 0, Gbottle, TestShortNameGP
56930         tb.MoveNext
56940     Loop



56950     Exit Sub

SaveBio_Error:

          Dim strES As String
          Dim intEL As Integer

56960     intEL = Erl
56970     strES = Err.Description
56980     LogError "fNewOrder", "SaveBio", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveDemographics
' Author    : Masood
' Date      : 16/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function SaveDemographics() As Boolean

56990     On Error GoTo SaveDemographics_Error
          Dim sql As String
          Dim tb As ADODB.Recordset

          Dim SampleIDWithOffset As String

57000     If UCase(EditScreen.Name) = UCase("frmEditMicrobiology") Then
              '+++ Junaid 20-05-2024
              '30            SampleIDWithOffset = Val(SampleID) + Val(sysOptMicroOffset(0))
57010         SampleIDWithOffset = Val(SampleID)
              '--- Junaid
57020     Else
57030         SampleIDWithOffset = Val(SampleID)
57040     End If

57050     SaveDemographics = False
57060     sql = "SELECT * FROM demographics WHERE " & _
              "SampleID = '" & SampleIDWithOffset & "'"
57070     Set tb = New Recordset
57080     RecOpenClient 0, tb, sql
57090     If tb.EOF Then
57100         tb.AddNew
57110         SaveDemographics = True
57120         tb!Rundate = Format$(Now, "dd/mmm/yyyy")
57130         tb!SampleID = SampleID
57140         tb!ExtSampleID = SampleIDExt
57150         tb!FAXed = 0
57160         tb!RooH = 0
57170     End If
57180     tb!Urgent = 0
57190     tb!Fasting = 0    'IIf(oSorF(1), 1, 0)
57200     tb.Update


57210     Exit Function


SaveDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

57220     intEL = Erl
57230     strES = Err.Description
57240     LogError "frmGporders", "SaveDemographics", intEL, strES, sql
End Function


Private Sub UpDateRequests(ByVal Discipline As String, _
          ByVal Code As String, _
          ByVal SampleType As String, SampleID As String, AddOn As String, Optional Gbottle As Integer, Optional TestShortNameGP As String)

          Dim sql As String

57250     On Error GoTo UpDateRequests_Error
57260     If Discipline = "Bio" Then
57270         sql = "INSERT INTO " & Discipline & "Requests " & _
                  "(SampleID, Code, DateTime, SampleType, Programmed, AddOn, AnalyserID,Gbottle) " & _
                  "SELECT DISTINCT '" & SampleID & "', " & _
                  "       '" & Code & "', getdate(), " & _
                  "       '" & SampleType & "', '0', '" & AddOn & "', " & _
                  "       Analyser ," & Gbottle & "  FROM " & Discipline & "TestDefinitions " & _
                  "        " & _
                  " WHERE Code = '" & Code & "' "
              '& _
              '" AND InUse = 1 ; "




57280     ElseIf Discipline = "Coag" Then
57290         sql = "Insert into CoagRequests " & _
                  "(SampleID, Code) VALUES " & _
                  "('" & SampleID & "', " & _
                  "'" & Code & "') ; "
57300     Else
57310         Exit Sub

57320     End If

57330     sql = sql & vbNewLine & " UPDATE GPOrders SET Programmed = 1 WHERE ShortName = '" & TestShortNameGP & "' AND SampleIDExternal = '" & SampleIDExt & "'"

57340     Cnxn(0).Execute sql

57350     Exit Sub

UpDateRequests_Error:

          Dim strES As String
          Dim intEL As Integer

57360     intEL = Erl
57370     strES = Err.Description
57380     LogError "frmNewOrder", "UpDateRequests", intEL, strES, sql

End Sub








Public Property Get ClinicalDetails() As String

57390     ClinicalDetails = m_sClinicalDetails

End Property

Public Property Let ClinicalDetails(ByVal sClinicalDetails As String)

57400     m_sClinicalDetails = sClinicalDetails

End Property

Public Property Get EditScreen() As Form

57410     Set EditScreen = m_objEditScreen

End Property

Public Property Set EditScreen(objEditScreen As Form)

57420     Set m_objEditScreen = objEditScreen

End Property

Public Property Get DisiplinesQuery() As String

57430     DisiplinesQuery = m_sDisiplinesQuery

End Property

Public Property Let DisiplinesQuery(ByVal sDisiplinesQuery As String)

57440     m_sDisiplinesQuery = sDisiplinesQuery

End Property

Public Property Get MicroSite() As String

57450     MicroSite = m_sMicroSite

End Property

Public Property Let MicroSite(ByVal sMicroSite As String)

57460     m_sMicroSite = sMicroSite

End Property
