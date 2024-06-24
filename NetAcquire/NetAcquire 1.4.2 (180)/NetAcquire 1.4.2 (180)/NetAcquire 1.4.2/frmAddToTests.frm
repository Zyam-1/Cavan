VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddToTests 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add to Tests Requested"
   ClientHeight    =   7815
   ClientLeft      =   960
   ClientTop       =   480
   ClientWidth     =   14160
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7815
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNpex 
      Caption         =   "Npex"
      Height          =   735
      Left            =   13020
      TabIndex        =   2
      Top             =   3270
      Width           =   975
   End
   Begin VB.Frame fmeDetail 
      Caption         =   "Detail"
      Height          =   3255
      Left            =   4133
      TabIndex        =   11
      Top             =   2700
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtSite 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   3
         Top             =   300
         Width           =   4035
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2985
         TabIndex        =   16
         Top             =   2790
         Width           =   1545
      End
      Begin VB.CommandButton btnShow 
         Caption         =   "Show"
         Height          =   375
         Left            =   1425
         TabIndex        =   15
         Top             =   2790
         Width           =   1545
      End
      Begin VB.TextBox txtComments 
         Height          =   1665
         Left            =   300
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   990
         Width           =   5325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Site Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   330
         Width           =   1395
      End
   End
   Begin VB.CommandButton btnExtReport 
      Caption         =   "External Order Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12990
      Picture         =   "frmAddToTests.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5310
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton cmdOrderExternal 
      Caption         =   "Order External"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   12990
      Picture         =   "frmAddToTests.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton cmdOrderBiomnis 
      Caption         =   "Order via &Biomnis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   12990
      Picture         =   "frmAddToTests.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2100
      Visible         =   0   'False
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5385
      Left            =   6390
      TabIndex        =   6
      Top             =   120
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   9499
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Analyte                               |<Sample Type  |<Destination                 "
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   7515
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   13256
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order via &Medibridge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   12990
      Picture         =   "frmAddToTests.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   900
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.ListBox lstPanels 
      Height          =   5730
      IntegralHeight  =   0   'False
      Left            =   4350
      TabIndex        =   1
      Top             =   1950
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   12990
      Picture         =   "frmAddToTests.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6595
      Width           =   1000
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmAddToTests.frx":31F2
      Top             =   90
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   630
      Picture         =   "frmAddToTests.frx":3634
      Top             =   270
      Width           =   480
   End
   Begin VB.Label lblComment 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   2145
      Left            =   6390
      TabIndex        =   7
      Top             =   5550
      Width           =   6405
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAddToTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Dim NodX As MSComctlLib.Node
Public FromEdit As Boolean

Private Activated As Boolean

Private m_Sex As String
Private m_SampleID As String
Private m_SampleDateTime As String
Private m_ClinicalDetails As String

Private PrvTestDetail As String

Public Sub OpenChrome(ByVal pURL As String)
          Dim sChromePath As String
          Dim sTmp As String
          Dim sProgramFiles As String
          Dim bNotFound As Boolean
          '
          ' check for 32/64 bit version
          '
39660     sProgramFiles = Environ("ProgramFiles")
39670     sChromePath = sProgramFiles & "\Google\Chrome\Application\chrome.exe"
39680     If Dir$(sChromePath) = vbNullString Then
              ' if not found, search for 32bit version
39690         sProgramFiles = Environ("ProgramFiles(x86)")
39700         If sProgramFiles > vbNullString Then
39710             sChromePath = sProgramFiles & "\Google\Chrome\Application\chrome.exe"
39720             If Dir$(sChromePath) = vbNullString Then
39730                 bNotFound = True
39740             End If
39750         Else
39760             bNotFound = True
39770         End If
39780     End If
39790     If bNotFound = True Then
39800         MsgBox "Chrome.exe not found"
39810         Exit Sub
39820     End If
39830     ShellExecute 0, "open", sChromePath, pURL, vbNullString, 1

End Sub

Private Sub FillDSSI()
    '
    '10    cmbDSSI.AddItem "AU -Audiology"
    '20    cmbDSSI.AddItem "BG -Blood gases"
    '30    cmbDSSI.AddItem "BLB-Blood bank"
    '40    cmbDSSI.AddItem "CUS-Cardiac Ultrasound"
    '50    cmbDSSI.AddItem "CTH-Cardiac catheterization"
    '60    cmbDSSI.AddItem "CT -CAT scan"
    '70    cmbDSSI.AddItem "CH -Chemistry"
    '80    cmbDSSI.AddItem "CP -Cytopathology"
    '90    cmbDSSI.AddItem "EC -Electrocardiac (e.g., EKG, EEC, Holter)"
    '100   cmbDSSI.AddItem "EN -Electroneuro(EEG, EMG, EP, PSG)"
    '110   cmbDSSI.AddItem "HM -Hematology"
    '120   cmbDSSI.AddItem "ICU-Bedside ICU Monitoring"
    '130   cmbDSSI.AddItem "IMG-Diagnostic Imaging"
    '140   cmbDSSI.AddItem "IMM-Immunology"
    '150   cmbDSSI.AddItem "LAB-Laboratory"
    '160   cmbDSSI.AddItem "MB -Microbiology"
    '170   cmbDSSI.AddItem "MCB-Mycobacteriology"
    '180   cmbDSSI.AddItem "MYC-Mycology"
    '190   cmbDSSI.AddItem "NMS-Nuclear medicine scan"
    '200   cmbDSSI.AddItem "NMR-Nuclear magnetic resonance"
    '210   cmbDSSI.AddItem "NRS-Nursing service measures"
    '220   cmbDSSI.AddItem "OUS-OB Ultrasound"
    '230   cmbDSSI.AddItem "OT -Occupational Therapy"
    '240   cmbDSSI.AddItem "OTH-Other"
    '250   cmbDSSI.AddItem "OSL-Outside Lab"
    '260   cmbDSSI.AddItem "PAR-Parasitology"
    '270   cmbDSSI.AddItem "PAT-Pathology(gross & histopath, Not Surgical)"
    '280   cmbDSSI.AddItem "PHR-Pharmacy"
    '290   cmbDSSI.AddItem "PT -Physical Therapy"
    '300   cmbDSSI.AddItem "PHY-Physician (Hx. Dx, admission note, etc.)"
    '310   cmbDSSI.AddItem "PF -Pulmonary function"
    '320   cmbDSSI.AddItem "RAD-Radiology"
    '330   cmbDSSI.AddItem "RX -Radiograph"
    '340   cmbDSSI.AddItem "RUS-Radiology ultrasound"
    '350   cmbDSSI.AddItem "RC -Respiratory Care (therapy)"
    '360   cmbDSSI.AddItem "RT -Radiation therapy"
    '370   cmbDSSI.AddItem "SR -Serology"
    '380   cmbDSSI.AddItem "SP -Surgical"
    '390   cmbDSSI.AddItem "TX -Toxicology"
    '400   cmbDSSI.AddItem "URN-Urinalysis"
    '410   cmbDSSI.AddItem "VUS-Vascular Ultrasound"
    '420   cmbDSSI.AddItem "VR -Virology"
    '430   cmbDSSI.AddItem "XRC-Cineradiograph"
    '
    '440   cmbDSSI = "LAB-Laboratory"

End Sub

Sub FillPanels()
Attribute FillPanels.VB_Description = "Load Panels"

          Dim tb As Recordset
          Dim sql As String

39840     On Error GoTo FillPanels_Error

39850     sql = "SELECT DISTINCT PanelName FROM ExtPanels " & _
              "ORDER BY PanelName"
39860     Set tb = New Recordset
39870     RecOpenServer 0, tb, sql

39880     lstPanels.Clear

39890     Do While Not tb.EOF
39900         lstPanels.AddItem tb!PanelName
39910         tb.MoveNext
39920     Loop

39930     Exit Sub

FillPanels_Error:

          Dim strES As String
          Dim intEL As Integer

39940     intEL = Erl
39950     strES = Err.Description
39960     LogError "frmAddToTests", "FillPanels", intEL, strES, sql

End Sub

Private Sub btnCancel_Click()
39970     fmeDetail.Visible = False
End Sub

Public Sub btnExtReport_Click()

          Dim l_SampleID As String
          
39980     l_SampleID = frmEditAll.txtSampleID.Text
39990     fmeDetail.Visible = True
40000     Call ShowSites
          
End Sub

Public Sub btnShow_Click()
40010     Call ShowReport(frmEditAll.txtSampleID.Text)
End Sub

Private Sub cmdCancel_Click()

          Dim sql As String
          Dim n As Integer
          Dim Test As String

40020     On Error GoTo cmdCancel_Click_Error

          'For n = 1 To g.Rows - 1
          '    Test = g.TextMatrix(n, 0)
          '
          '    sql = "IF NOT EXISTS (SELECT * FROM ExtResults " & _
          '          "             WHERE SampleID = '" & m_SampleID & "' " & _
          '          "             AND Analyte = '" & Test & "') " & _
          '          "    INSERT INTO ExtResults " & _
          '          "    (SampleID, Analyte, Result, SendTo, Units, SentDate) " & _
          '          "    SELECT '" & m_SampleID & "',  '" & Test & "', '', SendTo, Units, GETDATE() " & _
          '          "    FROM ExternalDefinitions WHERE AnalyteName = '" & Test & "'"
          '    Cnxn(0).Execute Sql
          'Next

40030     If IsChangeExits(PrvTestDetail) = True Then
40040         If iMsg("Do you want to exit without saving changes?", vbYesNo + vbQuestion) = vbNo Then
40050             Exit Sub
40060         Else
                  '        cmdOrderExternal_Click
                  '        Exit Sub
40070         End If
40080     End If

40090     Unload Me

40100     Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40110     intEL = Erl
40120     strES = Err.Description
40130     LogError "frmAddToTests", "cmdCancel_Click", intEL, strES, sql

End Sub

Private Sub cmdNpex_Click()
          'Shell "C:\Users\Public\Desktop\Google Chrome https://www.google.com/"
          'OpenChrome ("www.google.com")

          Dim url As String
40140     url = "www.google.com" ' Change this URL to the one you want to open
          
40150     OpenInChrome url



End Sub

Private Sub OpenInChrome(url As String)
          Dim chromePath As String
40160     chromePath = GetChromePath()
          
40170     If chromePath <> "" Then
40180         Shell """" & chromePath & """ """ & url & """", vbNormalFocus
40190     Else
40200         MsgBox "Chrome is not installed or its path could not be found."
40210     End If
End Sub

Private Function GetChromePath() As String
          Dim objShell As Object
          Dim regPath As String
          Dim chromePath As String
          
          ' Create a Shell object
40220     Set objShell = CreateObject("WScript.Shell")
          
          ' Registry path for Chrome installation
40230     regPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\"
          
          ' Read the registry value
40240     On Error Resume Next ' Ignore errors if the registry key doesn't exist
40250     chromePath = objShell.RegRead(regPath)
40260     On Error GoTo 0 ' Turn error handling back on
          
          ' Release the Shell object
40270     Set objShell = Nothing
          
          ' Return the Chrome path
40280     GetChromePath = chromePath
End Function

Private Sub cmdOrder_Click()

          Dim sql As String
          Dim n As Integer

40290     On Error GoTo cmdOrder_Click_Error

40300     For n = 1 To g.Rows - 1
40310         sql = "DECLARE @Code nvarchar(50) " & _
                  "DECLARE @Dept nvarchar(50) " & _
                  "SET @Code = (SELECT MBCode FROM ExternalDefinitions " & _
                  "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') " & _
                  "SET @Dept = (SELECT Department FROM ExternalDefinitions " & _
                  "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') " & _
                  "IF EXISTS(SELECT * FROM ExternalDefinitions " & _
                  "          WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "' " & _
                  "          AND COALESCE(MBCode, '') <> '') " & _
                  "  BEGIN IF NOT EXISTS(SELECT * FROM MediBridgeRequests " & _
                  "                  WHERE SampleID = '" & m_SampleID & "' " & _
                  "                  AND TestCode = @Code)" & _
                  "      INSERT INTO MedibridgeRequests (SampleID, TestCode, TestName, SampleDateTime, ClinDetails, Orderer, Dept, SpecimenSource) " & _
                  "      VALUES " & _
                  "     ('" & m_SampleID & "', " & _
                  "      @Code, " & _
                  "      '" & g.TextMatrix(n, 0) & "', " & _
                  "      '" & m_SampleDateTime & "', " & _
                  "      '" & m_ClinicalDetails & "', " & _
                  "      '" & UserCode & "^" & UserName & "', " & _
                  "      @Dept, " & _
                  "      '" & g.TextMatrix(n, 1) & "') END"
40320         Cnxn(0).Execute sql

40330     Next

40340     Exit Sub

cmdOrder_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40350     intEL = Erl
40360     strES = Err.Description
40370     LogError "frmAddToTests", "cmdOrder_Click", intEL, strES, sql

End Sub

Private Sub cmdOrderBiomnis_Click()
          Dim sql As String
          Dim n As Integer


40380     On Error GoTo cmdOrderBiomnis_Click_Error

40390     For n = 1 To g.Rows - 1
40400         sql = "DECLARE @Code nvarchar(50) " & _
                  "DECLARE @Dept nvarchar(50) " & _
                  "DECLARE @SampleType nvarchar(50) " & _
                  "SET @Code = (SELECT BiomnisCode FROM ExternalDefinitions " & _
                  "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "

40410         sql = sql & _
                  "SET @Dept = (SELECT Department FROM ExternalDefinitions " & _
                  "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "

40420         sql = sql & _
                  "SET @SampleType = (SELECT SampleType FROM ExternalDefinitions " & _
                  "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "
40430         sql = sql & _
                  "IF EXISTS(SELECT * FROM ExternalDefinitions " & _
                  "          WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "' " & _
                  "          AND COALESCE(BiomnisCode, '') <> '') " & _
                  "  BEGIN IF NOT EXISTS(SELECT * FROM BiomnisRequests " & _
                  "                  WHERE SampleID = '" & m_SampleID & "' " & _
                  "                  AND TestCode = @Code)" & _
                  "      INSERT INTO BiomnisRequests (SampleID, TestCode, TestName, SampleType, SampleDateTime, Department, RequestedBy, SendTo, Status) " & _
                  "      VALUES " & _
                  "     ('" & m_SampleID & "', " & _
                  "      @Code, " & _
                  "      '" & g.TextMatrix(n, 0) & "', " & _
                  "      @SampleType, " & _
                  "      '" & m_SampleDateTime & "', " & _
                  "      @Dept, " & _
                  "      '" & UserCode & "^" & UserName & "', " & _
                  "      '" & g.TextMatrix(n, 2) & "', " & _
                  "      'OutStanding') END"
40440         Cnxn(0).Execute sql

40450     Next

40460     Exit Sub

40470     Exit Sub

cmdOrderBiomnis_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40480     intEL = Erl
40490     strES = Err.Description
40500     LogError "frmAddToTests", "cmdOrderBiomnis_Click", intEL, strES, sql

End Sub



Private Sub cmdOrderExternal_Click()

          Dim sql As String
          Dim n As Integer


40510     On Error GoTo cmdOrderExternal_Click_Error


40520     For n = 1 To g.Rows - 1
              'If UCase(g.TextMatrix(n, 2)) = UCase(GetOptionSetting("ExternalBiomnisName", "Biomnis")) Then
              'UnComments
              'MsgBox ("Biomnis " & ListCodeFor("BiomnisEnableLabs", g.TextMatrix(n, 2)))
              'MsgBox ("NVRL " & ListCodeFor("MBEnableLabs", g.TextMatrix(n, 2)))
40530         If ListCodeFor("BiomnisEnableLabs", g.TextMatrix(n, 2)) <> "" Then
                  'Order to biomnis requests
40540             sql = "DECLARE @Code nvarchar(50) " & _
                      "DECLARE @Dept nvarchar(50) " & _
                      "DECLARE @SampleType nvarchar(50) " & _
                      "SET @Code = (SELECT BiomnisCode FROM ExternalDefinitions " & _
                      "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "

40550             sql = sql & _
                      "SET @Dept = (SELECT Department FROM ExternalDefinitions " & _
                      "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "

40560             sql = sql & _
                      "SET @SampleType = (SELECT SampleType FROM ExternalDefinitions " & _
                      "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "
40570             sql = sql & _
                      "IF EXISTS(SELECT * FROM ExternalDefinitions " & _
                      "          WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "' " & _
                      "          AND COALESCE(BiomnisCode, '') <> '') " & _
                      "  BEGIN IF NOT EXISTS(SELECT * FROM BiomnisRequests " & _
                      "                  WHERE SampleID = '" & m_SampleID & "' " & _
                      "                  AND TestCode = @Code)" & _
                      "      INSERT INTO BiomnisRequests (SampleID, TestCode, TestName, SampleType, SampleDateTime, Department, RequestedBy, SendTo, Status) " & _
                      "      VALUES " & _
                      "     ('" & m_SampleID & "', " & _
                      "      @Code, " & _
                      "      '" & g.TextMatrix(n, 0) & "', " & _
                      "      @SampleType, " & _
                      "      '" & m_SampleDateTime & "', " & _
                      "      @Dept, " & _
                      "      '" & UserCode & "^" & UserName & "', " & _
                      "      '" & g.TextMatrix(n, 2) & "', " & _
                      "      '" & IIf(g.TextMatrix(n, 2) = "NPEX", "Ordered", "OutStanding") & "') END"
40580             Cnxn(0).Execute sql
40590         Else
                  'Order to other external laboratories
40600             If ListCodeFor("MBEnableLabs", g.TextMatrix(n, 2)) <> "" Then
40610                 sql = "DECLARE @Code nvarchar(50) " & _
                          "DECLARE @Dept nvarchar(50) " & _
                          "SET @Code = (SELECT MBCode FROM ExternalDefinitions " & _
                          "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') " & _
                          "SET @Dept = (SELECT Department FROM ExternalDefinitions " & _
                          "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') " & _
                          "IF EXISTS(SELECT * FROM ExternalDefinitions " & _
                          "          WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "' " & _
                          "          AND COALESCE(MBCode, '') <> '') " & _
                          "  BEGIN IF NOT EXISTS(SELECT * FROM MediBridgeRequests " & _
                          "                  WHERE SampleID = '" & m_SampleID & "' " & _
                          "                  AND TestCode = @Code)" & _
                          "      INSERT INTO MedibridgeRequests (SampleID, TestCode, TestName, SampleDateTime, ClinDetails, Orderer, Dept, SpecimenSource, Status ) " & _
                          "      VALUES " & _
                          "     ('" & m_SampleID & "', " & _
                          "      @Code, " & _
                          "      '" & g.TextMatrix(n, 0) & "', " & _
                          "      '" & m_SampleDateTime & "', " & _
                          "      '" & m_ClinicalDetails & "', " & _
                          "      '" & UserCode & "^" & UserName & "', " & _
                          "      @Dept, " & _
                          "      '" & g.TextMatrix(n, 1) & "', " & _
                          "      'Requested') END"
40620                 Cnxn(0).Execute sql
40630             End If

40640         End If
              'Test = g.TextMatrix(n, 0)

40650         sql = "IF NOT EXISTS (SELECT * FROM ExtResults " & _
                  "             WHERE SampleID = '" & m_SampleID & "' " & _
                  "             AND Analyte = '" & g.TextMatrix(n, 0) & "') " & _
                  "    INSERT INTO ExtResults " & _
                  "    (SampleID, Analyte, Result, SendTo, Units, SentDate) " & _
                  "    SELECT '" & m_SampleID & "',  '" & g.TextMatrix(n, 0) & "', '', SendTo, Units, GETDATE() " & _
                  "    FROM ExternalDefinitions WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "'"
40660         Cnxn(0).Execute sql
40670     Next
40680     Unload Me

40690     Exit Sub

cmdOrderExternal_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40700     intEL = Erl
40710     strES = Err.Description
40720     LogError "frmAddToTests", "cmdOrderExternal_Click", intEL, strES, sql

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()

40730     If Not Activated Then
40740         FillTV
40750         FillPanels
40760         FillOrders
40770         FillDSSI
40780         Activated = True
40790     End If


40800     tv.SetFocus
40810     tv.SelectedItem.Expanded = False

End Sub
Private Sub FillOrders()

          Dim sql As String
          Dim tb As Recordset

40820     On Error GoTo FillOrders_Error

40830     g.Rows = 2
40840     g.AddItem ""
40850     g.RemoveItem 1
40860     PrvTestDetail = ""
40870     sql = "Select * from ExtResults where " & _
              "SampleID = '" & Val(frmEditAll.txtSampleID) & "'"
40880     Set tb = New Recordset
40890     RecOpenServer 0, tb, sql
40900     Do While Not tb.EOF
              'g.AddItem tb!Analyte & ""
40910         g.AddItem tb!Analyte & "" & vbTab & _
                  "" & vbTab & _
                  tb!SendTo & ""
40920         PrvTestDetail = PrvTestDetail & tb!Analyte
40930         tb.MoveNext
40940     Loop

40950     If g.Rows > 2 Then
40960         g.RemoveItem 1
40970     End If

40980     Exit Sub

FillOrders_Error:

          Dim strES As String
          Dim intEL As Integer

40990     intEL = Erl
41000     strES = Err.Description
41010     LogError "frmAddToTests", "FillOrders", intEL, strES, sql

End Sub

Sub FillTV()
Attribute FillTV.VB_Description = "Fill Ndal List"

          Dim NodX As MSComctlLib.Node
          Dim n As Integer
          Dim Relative As String
          Dim ThisNode As String
          Dim sql As String
          Dim tb As Recordset
          Dim Key As String
          Dim NodeText As String

41020     On Error GoTo FillTV_Error

41030     For n = Asc("A") To Asc("Z")
41040         Key = Chr$(n)
41050         NodeText = Chr$(n)
41060         Set NodX = tv.Nodes.Add(, , Key, NodeText)
41070     Next
41080     For n = Asc("0") To Asc("9")
41090         Set NodX = tv.Nodes.Add(, , "#" & Chr$(n), Chr$(n))
41100     Next

41110     sql = "SELECT AnalyteName FROM ExternalDefinitions where InUse=1 " & _
              "ORDER BY AnalyteName"
41120     Set tb = New Recordset
41130     RecOpenServer 0, tb, sql
41140     Do While Not tb.EOF
41150         If Trim$(tb!AnalyteName & "") <> "" Then
41160             Relative = UCase(Left(tb!AnalyteName, 1))
41170             If IsNumeric(Relative) Then Relative = "#" & Relative
41180             ThisNode = tb!AnalyteName
41190             Set NodX = tv.Nodes.Add(Relative, tvwChild, , ThisNode)
41200         End If
41210         tb.MoveNext
41220     Loop

41230     Exit Sub

FillTV_Error:

          Dim strES As String
          Dim intEL As Integer

41240     intEL = Erl
41250     strES = Err.Description
41260     LogError "frmAddToTests", "FillTV", intEL, strES, sql

End Sub


Private Sub Form_Load()

41270     Activated = False

41280     cmdOrder.Visible = False

41290     If GetOptionSetting("DeptMediBridge", "0") <> "0" Then
41300         cmdOrder.Visible = False
41310     End If

41320     EnsureColumnExists "MedibridgeRequests", "DateTimeOfRecord", "datetime DEFAULT getdate()"
41330     EnsureColumnExists "MedibridgeRequests", "Status", "nvarchar(50) DEFAULT 'Requested'"
41340     If frmEditAll.m_ShowDoc Then
41350         Call btnExtReport_Click
41360         txtComments.Text = frmEditAll.m_Notes
41370         DoEvents
41380         Call btnShow_Click
41390         DoEvents
41400         Unload Me
41410     End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

41420     Activated = False

End Sub

Private Sub g_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
          Dim S As String
          Dim n As Integer
          Dim R As Integer

41430     R = g.MouseRow

41440     S = "Remove " & g.TextMatrix(R, 0) & " from tests requested?"
41450     n = iMsg(S, vbYesNo + vbQuestion)
41460     If n = vbYes Then
41470         If g.Rows = 2 Then
41480             g.AddItem ""
41490             g.RemoveItem 1
41500         Else
41510             g.RemoveItem R
41520         End If
41530     End If

End Sub


Private Sub lstpanels_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Found As Boolean
          Dim S As String

41540     On Error GoTo lstpanels_Click_Error

41550     sql = "SELECT P.Content, D.SampleType, D.SendTo, D.Comment from ExtPanels P join ExternalDefinitions D " & _
              "ON P.Content = D.AnalyteName WHERE PanelName = '" & lstPanels & "'"
41560     Set tb = New Recordset
41570     RecOpenServer 0, tb, sql
41580     Do While Not tb.EOF
41590         Found = False
41600         For n = 1 To g.Rows - 1
41610             If g.TextMatrix(n, 0) = tb!Content & "" Then
41620                 Found = True
41630                 Exit For
41640             End If
41650         Next
41660         If Not Found Then
41670             S = tb!Content & vbTab & tb!SampleType & vbTab & tb!SendTo & ""
41680             g.AddItem S
41690             If Trim$(tb!Comment & "") <> "" Then
41700                 S = tb!Content & " : " & tb!Comment & vbCrLf
41710                 lblComment = lblComment & S
41720             End If
41730         End If

41740         tb.MoveNext
41750     Loop

41760     If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
41770         g.RemoveItem 1
41780     End If

41790     Exit Sub

lstpanels_Click_Error:

          Dim strES As String
          Dim intEL As Integer

41800     intEL = Erl
41810     strES = Err.Description
41820     LogError "frmAddToTests", "lstpanels_Click", intEL, strES, sql

End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Set nodX = tv.SelectedItem

End Sub


Private Sub tv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If Button = vbLeftButton Then
    '  tv.Drag vbBeginDrag
    'End If

End Sub


Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim Found As Boolean
          Dim S As String

41830     On Error GoTo tv_NodeClick_Error

41840     sql = "Select * from ExternalDefinitions where " & _
              "AnalyteName = '" & Node.Text & "'"
41850     Set tb = New Recordset
41860     RecOpenServer 0, tb, sql
41870     If Not tb.EOF Then
41880         Found = False
41890         For n = 0 To g.Rows - 1
41900             If g.TextMatrix(n, 0) = Node.Text Then
41910                 Found = True
41920                 Exit For
41930             End If
41940         Next
41950         If Not Found Then
41960             S = Node.Text & vbTab & tb!SampleType & vbTab & tb!SendTo & ""
41970             g.AddItem S
41980             If Trim$(tb!Comment & "") <> "" Then
41990                 S = Node.Text & " : " & tb!Comment & vbCrLf
42000                 lblComment = lblComment & S
42010             End If
42020         End If
42030     End If

42040     If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
42050         g.RemoveItem 1
42060     End If

42070     Exit Sub

tv_NodeClick_Error:

          Dim strES As String
          Dim intEL As Integer

42080     intEL = Erl
42090     strES = Err.Description
42100     LogError "frmAddToTests", "tv_NodeClick", intEL, strES, sql

End Sub



Public Property Let Sex(ByVal sNewValue As String)

42110     m_Sex = sNewValue

End Property

Public Property Let SampleID(ByVal sNewValue As String)

42120     m_SampleID = sNewValue

End Property

Public Property Let SampleDateTime(ByVal sNewValue As String)

42130     m_SampleDateTime = sNewValue

End Property

Public Property Let ClinicalDetails(ByVal sNewValue As String)

42140     m_ClinicalDetails = sNewValue

End Property

'---------------------------------------------------------------------------------------
' Procedure : IsChangeExits
' Author    : XPMUser
' Date      : 12/Nov/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function IsChangeExits(PrevValue As String) As Boolean

42150     On Error GoTo IsChangeExits_Error
          Dim i As Integer
          Dim CurVal As String

42160     IsChangeExits = False
42170     CurVal = ""

42180     With g
42190         For i = 1 To g.Rows - 1
42200             CurVal = CurVal & .TextMatrix(i, 0)
42210         Next i
42220     End With

42230     If CurVal <> PrevValue Then
42240         IsChangeExits = True
42250     End If
42260     Exit Function


IsChangeExits_Error:

          Dim strES As String
          Dim intEL As Integer

42270     intEL = Erl
42280     strES = Err.Description
42290     LogError "frmAddToTests", "IsChangeExits", intEL, strES
End Function

Private Sub ShowSites()
42300     On Error GoTo ErrorHandler
          
42310     txtSite.Text = g.TextMatrix(1, 2)
          
42320     Exit Sub
ErrorHandler:

          Dim strES As String
          Dim intEL As Integer

42330     intEL = Erl
42340     strES = Err.Description
42350     LogError "frmAddToTests", "ShowSites", intEL, strES
End Sub

Private Sub ShowReport(p_SampleID As String)
42360     On Error Resume Next
          'On Error GoTo ErrorHandler
          
          Dim sql As String
          Dim tb As ADODB.Recordset
          
42370     sql = "Select SampleID, PatName, DoB, Chart, Addr0, Addr1, '" & GetAnalyteName(p_SampleID) & "' AnalyteName, Ward, Clinician, (Select top(1) R.CLDetails from ocmRequest R Inner Join ocmRequestDetails D On R.RequestID = D.RequestID Where D.SampleID = '" & p_SampleID & "') CLDetails From demographics Where SampleID = '" & p_SampleID & "'"
42380     Set tb = New Recordset
42390     RecOpenServer 0, tb, sql
42400     If Not tb Is Nothing Then
42410         If Not tb.EOF Then
42420             Set rptTestReport.DataSource = tb
42430             rptTestReport.Sections(1).Controls("Label6").Caption = txtSite.Text
42440             rptTestReport.Sections(1).Controls("Label7").Caption = rptTestReport.Sections(1).Controls("Label7").Caption & " " & ConvertNull(tb!Clinician, "")
42450             rptTestReport.Sections(1).Controls("Label25").Caption = ConvertNull(tb!SampleID, "")
42460             rptTestReport.Sections(1).Controls("Label9").Caption = ConvertNull(tb!PatName, "")
42470             rptTestReport.Sections(1).Controls("Label11").Caption = Format$(ConvertNull(tb!DoB, ""), "dd-mm-yyyy")
42480             rptTestReport.Sections(1).Controls("Label13").Caption = ConvertNull(tb!Chart, "")
42490             rptTestReport.Sections(1).Controls("Label15").Caption = ConvertNull(tb!Addr0, "")
42500             rptTestReport.Sections(1).Controls("Label17").Caption = ConvertNull(tb!Addr1, "")
42510             rptTestReport.Sections(1).Controls("Label18").Caption = ConvertNull(tb!AnalyteName, "")
42520             rptTestReport.Sections(1).Controls("Label20").Caption = txtComments.Text
42530             rptTestReport.Sections(1).Controls("Label22").Caption = UserName
42540             rptTestReport.Sections(1).Controls("Label23").Caption = ConvertNull(tb!Ward, "")
42550             rptTestReport.Sections(1).Controls("Label26").Caption = ConvertNull(tb!Clinician, "")
42560             rptTestReport.Sections(1).Controls("Label27").Caption = ConvertNull(tb!ClDetails, "")
42570             rptTestReport.Show 1
42580             rptTestReport.SetFocus
42590         End If
42600     End If
          
42610     Exit Sub
ErrorHandler:

          Dim strES As String
          Dim intEL As Integer

42620     intEL = Erl
42630     strES = Err.Description
          'MsgBox strES
42640     LogError "frmAddToTests", "ShowReport", intEL, strES
End Sub

Private Function GetAnalyteName(p_SampleID As String) As String
42650     On Error GoTo ErrorHandler
          
          Dim sql As String
          Dim tb As ADODB.Recordset
          
42660     GetAnalyteName = ""
42670     sql = "Select Analyte From ExtResults Where SampleID = '" & p_SampleID & "'"
42680     Set tb = New Recordset
42690     RecOpenServer 0, tb, sql
42700     If Not tb Is Nothing Then
42710         If Not tb.EOF Then
42720             While Not tb.EOF
42730                 If GetAnalyteName = "" Then
42740                     GetAnalyteName = ConvertNull(tb!Analyte, "")
42750                 Else
42760                     GetAnalyteName = GetAnalyteName & ", " & ConvertNull(tb!Analyte, "")
42770                 End If
42780                 tb.MoveNext
42790             Wend
42800         End If
42810     End If
          
42820     Exit Function
ErrorHandler:

          Dim strES As String
          Dim intEL As Integer

42830     intEL = Erl
42840     strES = Err.Description
42850     LogError "frmAddToTests", "GetAnalyteName", intEL, strES
End Function

