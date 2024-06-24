Attribute VB_Name = "BBankShared"
Option Explicit

Public gEVENTCODES As New EventCodes

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long

Public blnEndApp As Boolean
Public strLatestVersion As String

Public blnPrintingWithPreview As Boolean

Public Cnxn() As Connection
Public CnxnBB() As Connection

Public intOtherHospitalsInGroup As Long

Public mGroup As String
Public mRh As String

Public HospName() As String
Public UserName As String
Public UserCode As String
Public UserInitials As String
Public SecondUserName As String

Public Training As Boolean

Public LogOffDelay As Long
Public LogOffDelayMin As Long
Public LogOffDelaySecs As Long

Global gbitmap As String

Public UserMemberOf As String

Global Const XMATCH = 0
Global Const GROUP_HOLD = 1
Global Const ANTENATAL = 3
Global Const DAT = 2

Public CancelCode As String
Public ValidateCode As String

Global Dept As Integer

Public Const MaxAgeToDays As Long = 43830

Public Const FORWARD = -1
Public Const BACKWARD = 1
Public Const DONTCARE = 0

Public colSexNames As New SexNames

Public Entity As String

Public TransfusionForm As String
Public TransfusionLabel As String
Public TransfusionPDF As String

Public TimedOut As Boolean
Public Answer As Long


Public Type PhoneLog
    SampleID As String
    DateTime As Date
    PhonedTo As String
    PhonedBy As String
    Comment As String
End Type

Public Enum InputValidation
    Numericfullstopdash = 0
    Char = 1
    YorN = 2
    AlphaNumeric_NoApos = 3
    AlphaNumeric_AllowApos = 4
    Numeric_Only = 5
End Enum
Public Enum PrintAlignContants
    Alignleft = 0
    AlignCenter = 1
    AlignRight = 2
End Enum

Public dateQC As Date

Public Enum AsMemberOf
    LookUp = 0
    Users = 1
    Managers = 2
    Administrators = 3

End Enum

Public Type udtRS
    UnitNumber As String
    ProductCode As String
    StorageLocation As String
    UnitExpiryDate As String
    UnitGroup As String
    StockComment As String
    ActionText As String
    Chart As String
    PatientHealthServiceNumber As String
    SurName As String
    ForeName As String
    DoB As String
    Sex As String
    PatientGroup As String
    DeReservationDateTime As String
    RERStatus As String
    RERExpiry As String
    UserName As String
    SampleStatus As String
    SampleValidDateTime As String
End Type

Public strBTCourier_StorageLocation_StockFridge As String
Public strBTCourier_StorageLocation_RoomTempIssueFridge As String
Public strBTCourier_StorageLocation_HemoSafeFridge As String

Public CurrentReceivedDate As String
Public Const NewFormatFGNumber As Long = 6869

Public blnBTCdownWarningDisplayed As Boolean

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val0 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val3 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val4 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val5 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val6 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val7 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val8 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val9 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos2_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos2_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos3_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos3_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos4_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos4_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos5_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos5_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos6_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos6_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos7_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos7_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos8_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos8_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos9_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos9_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos10_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos10_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos11_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos11_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos12_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos12_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos13_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos13_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos14_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos14_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos15_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos15_Antigen2 As String

Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos16_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos16_Antigen2 As String


Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value0_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value0_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value1_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value1_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value2_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value2_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value3_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value3_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value4_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value4_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value5_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value5_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value6_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value6_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value7_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value7_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen2 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value9_Antigen1 As String
Public pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value9_Antigen2 As String

Public pubStrRT014Interpretation_Pos9r1 As String
Public pubStrRT014Interpretation_Pos10r1 As String
Public pubStrRT014Interpretation_Pos11r1 As String
Public pubStrRT014Interpretation_Pos12r1 As String
Public pubStrRT014Interpretation_Pos13r1 As String
Public pubStrRT014Interpretation_Pos14r1 As String
Public pubStrRT014Interpretation_Pos15r1 As String
Public pubStrRT014Interpretation_Pos16r1 As String

Public pubStrRT014Interpretation_Pos9r2 As String
Public pubStrRT014Interpretation_Pos10r2 As String
Public pubStrRT014Interpretation_Pos11r2 As String
Public pubStrRT014Interpretation_Pos12r2 As String
Public pubStrRT014Interpretation_Pos13r2 As String
Public pubStrRT014Interpretation_Pos14r2 As String
Public pubStrRT014Interpretation_Pos15r2 As String
Public pubStrRT014Interpretation_Pos16r2 As String

Public pubStrRT014Interpretation_Value1r1 As String
Public pubStrRT014Interpretation_Value1r2 As String
Public pubStrRT014Interpretation_Value2r1 As String
Public pubStrRT014Interpretation_Value2r2 As String
Public pubStrRT014Interpretation_Value3r1 As String
Public pubStrRT014Interpretation_Value3r2 As String
Public pubStrRT014Interpretation_Value4r1 As String
Public pubStrRT014Interpretation_Value4r2 As String
Public pubStrRT014Interpretation_Value5r1 As String
Public pubStrRT014Interpretation_Value5r2 As String
Public pubStrRT014Interpretation_Value6r1 As String
Public pubStrRT014Interpretation_Value6r2 As String
Public pubStrRT014Interpretation_Value7r1 As String
Public pubStrRT014Interpretation_Value7r2 As String
Public pubStrRT014Interpretation_Value8r1 As String
Public pubStrRT014Interpretation_Value8r2 As String

Public pubStrStockSupplierName_IRL As String

Public Function CheckBadReaction(ByVal MRN As String) As Boolean

      Dim tb As Recordset
      Dim sql As String
      Dim retval As Boolean

10    On Error GoTo CheckBadReaction_Error

20    retval = False

30    If Trim$(MRN) <> "" Then

40        sql = "Select * from BadReact where " & _
                "PatNo = '" & MRN & "'"
50        Set tb = New Recordset
60        RecOpenServerBB 0, tb, sql

70        If Not tb.EOF Then
80            retval = True
90        End If

100   End If

110   CheckBadReaction = retval

120   Exit Function

CheckBadReaction_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "BBankShared", "CheckBadReaction", intEL, strES, sql

End Function

Public Function CheckCharacterForBatch(ByVal Identifier As String) As String

      Dim n As Integer
      Dim Sum As Long

10    Identifier = UCase$(Identifier)

20    Sum = 0
30    For n = 1 To Len(Identifier)
40        Sum = Sum + Asc(Mid$(Identifier, n, 1))
50    Next
60    Sum = Sum Mod 8
70    CheckCharacterForBatch = Chr$(Asc("A") + Sum)

End Function

Public Function AuthenticateUser(UserRole As AsMemberOf) As Boolean

      Dim tb As Recordset
      Dim sql As String
      Dim Pass As String
      Dim MemOf As String

10    On Error GoTo AuthenticateUser_Error

20    Pass = iBOX("Please enter password", , , True)
30    If Pass = "" Then
40        AuthenticateUser = False
50    Else
60        Select Case UserRole
          Case 0:
70            MemOf = "LookUp"
80        Case 1:
90            MemOf = "Users"
100       Case 2:
110           MemOf = "Managers"
120       Case 3:
130           MemOf = "Administrators"
140       End Select
150       sql = "Select * From Users Where Password = '" & Pass & "'" & _
                " And MemberOf = '" & MemOf & "' And InUse = 1"
160       Set tb = New Recordset
170       RecOpenClient 0, tb, sql
180       If tb.EOF Then
190           AuthenticateUser = False
200       Else
210           AuthenticateUser = True
220           SecondUserName = tb!Name
230       End If

240   End If

250   Exit Function

AuthenticateUser_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "BBankShared", "AuthenticateUser", intEL, strES, sql

End Function

Public Function GetConfirmDateTime(ByVal SampleID As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetConfirmDateTime_Error

20    GetConfirmDateTime = ""

30    sql = "SELECT TOP 1 * FROM ConfirmDetails WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "ORDER BY DateTimeOfRecord DESC"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70        GetConfirmDateTime = Format$(tb!DateTimeOfRecord, "dd/MM/yyyy HH:nn:ss")
80    End If

90    Exit Function

GetConfirmDateTime_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BBankShared", "GetConfirmDateTime", intEL, strES, sql

End Function

Public Function MinDate(ByVal d1 As Date, ByVal d2 As Date) As Date

10    If DateDiff("s", d1, d2) > 0 Then
20        MinDate = d1
30    Else
40        MinDate = d2
50    End If

End Function

Public Function PatientDetailsConfirmed(ByVal SampleID As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PatientDetailsConfirmed_Error

20    sql = "Select Top 1 * From ConfirmDetails Where SampleID = '" & SampleID & "' " & _
            "ORDER BY DateTimeOfRecord DESC"
30    Set tb = New Recordset
40    RecOpenClientBB 0, tb, sql

50    If tb.EOF Then
60        PatientDetailsConfirmed = False
70    Else
80        If IsNull(tb!Confirmed) Then
90            PatientDetailsConfirmed = False
100       Else
110           PatientDetailsConfirmed = (tb!Confirmed = 1)
120       End If
130   End If
140   Exit Function

PatientDetailsConfirmed_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "BBankShared", "PatientDetailsConfirmed", intEL, strES, sql

End Function

Public Function ConfirmDetailsCount(ByVal SampleID As String) As Integer

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ConfirmDetailsCount_Error

20    sql = "Select Count(*) As CNT From ConfirmDetails Where SampleID = '" & SampleID & "' " & _
            "And Confirmed = 1"
30    Set tb = New Recordset
40    RecOpenClientBB 0, tb, sql

50    ConfirmDetailsCount = tb!CNT

60    Exit Function

ConfirmDetailsCount_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BBankShared", "ConfirmDetailsCount", intEL, strES, sql

End Function

Public Sub HighlightGridRow(ByVal g As MSFlexGrid)

      Dim xSave As Integer
      Dim ySave As Integer
      Dim X As Integer
      Dim Y As Integer

10    On Error GoTo HighlightGridRow_Error

20    With g
30        ySave = .row
40        xSave = .col
50        .col = 0
60        For Y = 1 To .Rows - 1
70            .row = Y
80            If .CellBackColor = vbYellow Then
90                For X = 0 To .Cols - 1
100                   .col = X
110                   .CellBackColor = 0
120               Next
130               Exit For
140           End If
150       Next
160       .row = ySave
170       For X = 0 To .Cols - 1
180           .col = X
190           .CellBackColor = vbYellow
200       Next
210       .col = xSave
220   End With

230   Exit Sub

HighlightGridRow_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "BBankShared", "HighlightGridRow", intEL, strES

End Sub

Public Function IsTableInDatabaseBB(ByVal TableName As String) As Boolean

      Dim tbExists As Recordset
      Dim sql As String
      Dim retval As Boolean

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist
      'if it has a record then the table does exist.

10    On Error GoTo IsTableInDatabaseBB_Error

20    sql = "SELECT name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = '" & TableName & "'"
30    Set tbExists = CnxnBB(0).Execute(sql)

40    retval = True

50    If tbExists.EOF Then    'There is no table <TableName> in database
60        retval = False
70    End If
80    IsTableInDatabaseBB = retval

90    Exit Function

IsTableInDatabaseBB_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BBankShared", "IsTableInDatabaseBB", intEL, strES, sql


End Function


Public Sub CheckKleihauerQCInDb()

      Dim sql As String

10    On Error GoTo CheckKleihauerQCInDb_Error

20    If IsTableInDatabaseBB("KleihauerQC") = False Then    'There is no table  in database
30        sql = "CREATE TABLE KleihauerQC " & _
                "( SampleID numeric(9), " & _
                "  Positive nvarchar(1), " & _
                "  Negative nvarchar(1), " & _
                "  Operator nvarchar(50), " & _
                "  DateTime datetime, " & _
                "  Rhesus nvarchar(50) )"
40        CnxnBB(0).Execute sql
50    End If

60    Exit Sub

CheckKleihauerQCInDb_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BBankShared", "CheckKleihauerQCInDb", intEL, strES, sql

End Sub

Public Sub CheckConfirmDetailsInDb()

      Dim sql As String

10    On Error GoTo CheckConfirmDetailsInDb_Error

20    If IsTableInDatabaseBB("ConfirmDetails") = False Then    'There is no table  in database
30        sql = "CREATE TABLE ConfirmDetails " & _
                "( SampleID nvarchar(20), " & _
                "  AandE nvarchar(20) NOT NULL DEFAULT '', " & _
                "  Chart nvarchar(20) NOT NULL DEFAULT '', " & _
                "  Typenex nvarchar(20), " & _
                "  Sex nvarchar(20), " & _
                "  FwdGroup nvarchar(20), " & _
                "  Name nvarchar(50), " & _
                "  Notes nvarchar(50), " & _
                "  Operator nvarchar(50), " & _
                "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
                "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
40        CnxnBB(0).Execute sql
50    End If

60    Exit Sub

CheckConfirmDetailsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BBankShared", "CheckConfirmDetailsInDb", intEL, strES, sql

End Sub

Public Sub CheckStLukesCentrifugeInDb()

      Dim sql As String

10    On Error GoTo CheckStLukesCentrifugeInDb_Error

20    If IsTableInDatabaseBB("StLukesCentrifuge") = False Then    'There is no table  in database

30        sql = "CREATE TABLE [dbo].[StLukesCentrifuge]( " & _
                "[Cent1Phase1] [nvarchar](10) NULL, " & _
                "[Cent1Phase2] [nvarchar](10) NULL, " & _
                "[Cent2Phase1] [nvarchar](10) NULL, " & _
                "[Cent2Phase2] [nvarchar](10) NULL, " & _
                "[BlockL] [nvarchar](10) NULL, " & _
                "[BlockR] [nvarchar](10) NULL, " & _
                "[BlockS] [nvarchar](10) NULL, " & _
                "[DateTime] [datetime] NULL, " & _
                "[Operator] [nvarchar](20) NULL, " & _
                "[Comment] [ntext] NULL )"
40        CnxnBB(0).Execute sql
50    End If

60    Exit Sub

CheckStLukesCentrifugeInDb_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BBankShared", "CheckStLukesCentrifugeInDb", intEL, strES, sql

End Sub

Public Sub CheckPhoneLogInDb()

      Dim sql As String

10    On Error GoTo CheckPhoneLogInDb_Error

20    If IsTableInDatabaseBB("PhoneLog") = False Then    'There is no table  in database
30        sql = "CREATE TABLE PhoneLog " & _
                "( SampleID nvarchar(50), " & _
                "  PhonedTo nvarchar(50), " & _
                "  PhonedBy nvarchar(50), " & _
                "  Comment nvarchar(50), " & _
                "  DateTime datetime, " & _
                "  Discipline nvarchar(50) )"
40        CnxnBB(0).Execute sql
50    End If

60    Exit Sub

CheckPhoneLogInDb_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BBankShared", "CheckPhoneLogInDb", intEL, strES, sql

End Sub


Public Sub CheckPhenoTypeInDb()

      Dim sql As String

10    On Error GoTo CheckPhenoTypeInDb_Error

20    If IsTableInDatabaseBB("StLukesPhenoType") = False Then    'There is no table  in database
30        sql = "CREATE TABLE StLukesPhenoType " & _
                "( DateTime datetime, " & _
                "  AntiKLotNumber nvarchar(50), " & _
                "  AntiKExpiry smalldatetime, " & _
                "  AntiE0LotNumber nvarchar(50), " & _
                "  AntiE0Expiry smalldatetime, " & _
                "  AntiE1LotNumber nvarchar(50), " & _
                "  AntiE1Expiry smalldatetime, " & _
                "  AntiC0LotNumber nvarchar(50), " & _
                "  AntiC0Expiry smalldatetime, " & _
                "  AntiC1LotNumber nvarchar(50), " & _
                "  AntiC1Expiry smalldatetime, " & _
                "  Comment nvarchar(50), " & _
                "  Operator nvarchar(50))"

40        CnxnBB(0).Execute sql
50    End If

60    Exit Sub

CheckPhenoTypeInDb_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BBankShared", "CheckPhenoTypeInDb", intEL, strES, sql

End Sub

Public Function CheckPhoneLog(ByVal SID As String) As PhoneLog

      'Returns PhoneLog.SampleID = 0 if no entry in phone log

      Dim tb As Recordset
      Dim sql As String
      Dim PL As PhoneLog

10    On Error GoTo CheckPhoneLog_Error

20    sql = "Select * from PhoneLog where " & _
            "SampleID = '" & SID & "'"
30    Set tb = CnxnBB(0).Execute(sql)
40    If tb.EOF Then
50        CheckPhoneLog.SampleID = "0"
60    Else
70        With PL
80            .SampleID = SID
90            .Comment = tb!Comment & ""
100           .DateTime = tb!DateTime
110           .PhonedBy = tb!PhonedBy & ""
120           .PhonedTo = tb!PhonedTo & ""
130       End With
140       CheckPhoneLog = PL
150   End If

160   Exit Function

CheckPhoneLog_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "BBankShared", "CheckPhoneLog", intEL, strES, sql

End Function



Public Function EnsureColumnExistsBB(ByVal TableName As String, _
                                     ByVal ColumnName As String, _
                                     ByVal Definition As String) _
                                     As Boolean

      'Return 1 if column created
      '       0 if column already exists

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo EnsureColumnExistsBB_Error

20    frmSplash.lblUpdate.Caption = "Checking column '" & ColumnName & "' in '" & TableName & "' table"
30    frmSplash.pbUpdate.Value = frmSplash.pbUpdate.Value + 1
40    DoEvents
50    sql = "IF NOT EXISTS " & _
            "    (SELECT * FROM syscolumns WHERE " & _
            "    id = object_id('" & TableName & "') " & _
            "    AND name = '" & ColumnName & "') " & _
            "  BEGIN " & _
            "    ALTER TABLE " & TableName & " " & _
            "    ADD " & ColumnName & " " & Definition & " " & _
            "    SELECT 1 AS RetVal " & _
            "  END " & _
            "ELSE " & _
            "  SELECT 0 AS RetVal"

60    Set tb = CnxnBB(0).Execute(sql)

70    EnsureColumnExistsBB = tb!retval

80    Exit Function

EnsureColumnExistsBB_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "BBankShared", "EnsureColumnExistsBB", intEL, strES, sql


End Function

Public Sub CheckUsersArcInDb()

      Dim sql As String

10    If IsTableInDatabase("UsersArc") = False Then    'There is no table  in database
20        sql = "CREATE TABLE UsersArc " & _
                "( PassWord nvarchar(50), " & _
                "  Name nvarchar(50), " & _
                "  Code nvarchar(5), " & _
                "  InUse bit, " & _
                "  MemberOf nvarchar(25), " & _
                "  LogOffDelay numeric, " & _
                "  ListOrder int, " & _
                "  Prints bit, " & _
                "  PassDate datetime, " & _
                "  ArchiveDateTime datetime NOT NULL DEFAULT getdate(), " & _
                "  ArchivedBy nvarchar(50), " & _
                "  RowGUID uniqueidentifier ROWGUIDCOL NOT NULL DEFAULT newid() )"

30        Cnxn(0).Execute sql
40    End If

50    Exit Sub

CheckHealthlinkInDb_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "modDbDesign", "CheckHealthlinkInDb", intEL, strES, sql

End Sub

Public Function EnsureColumnExists(ByVal TableName As String, _
                                   ByVal ColumnName As String, _
                                   ByVal Definition As String) _
                                   As Boolean

      'Return 1 if column created
      '       0 if column already exists

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo EnsureColumnExists_Error

20    sql = "IF NOT EXISTS " & _
            "    (SELECT * FROM syscolumns WHERE " & _
            "    id = object_id('" & TableName & "') " & _
            "    AND name = '" & ColumnName & "') " & _
            "  BEGIN " & _
            "    ALTER TABLE " & TableName & " " & _
            "    ADD " & ColumnName & " " & Definition & " " & _
            "    SELECT 1 AS RetVal " & _
            "  END " & _
            "ELSE " & _
            "  SELECT 0 AS RetVal"

30    Set tb = Cnxn(0).Execute(sql)

40    EnsureColumnExists = tb!retval

50    Exit Function

EnsureColumnExists_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "BBankShared", "EnsureColumnExists", intEL, strES, sql

End Function

Public Function FormatGroup(ByVal GivenGroup As String) As String

      Dim strGroup As String
      Dim strRhesus As String

      'I want the group in the format eg "AB,+" or "O,-"
10    GivenGroup = UCase$(GivenGroup)

20    strGroup = Trim$(Left$(GivenGroup & "  ", 2))
30    Select Case strGroup

      Case "A", "B", "O", "AB"
40        If InStr(GivenGroup, "POS") <> 0 Then
50            strRhesus = "+"
60        ElseIf InStr(GivenGroup, "NEG") <> 0 Then
70            strRhesus = "-"
80        Else
90            strGroup = ""
100           strRhesus = ""
110       End If

120   Case Else
130       strGroup = ""
140       strRhesus = ""

150   End Select

160   FormatGroup = strGroup & "," & strRhesus

End Function
Public Function FullSex(ByVal MorF As String) As String

10    Select Case MorF
      Case "M": FullSex = "Male"
20    Case "F": FullSex = "Female"
30    Case Else: FullSex = "Unknown"
40    End Select

End Function

Public Function PrinterNameForMapping(ByVal MappedTo As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrinterNameForMapping_Error

20    sql = "Select * from Printers where " & _
            "MappedTo = '" & MappedTo & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        PrinterNameForMapping = UCase$(tb!PrinterName & "")
70    End If

80    Exit Function

PrinterNameForMapping_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "BBankShared", "PrinterNameForMapping", intEL, strES, sql


End Function

Public Function ReadableGroup(ByVal strGroup As String) As String

      Dim s As String

10    strGroup = UCase$(strGroup)

20    strGroup = Replace(strGroup, ",", " ")

30    s = Trim$(Left$(strGroup, 2))
40    If InStr(strGroup, "+") <> 0 Then
50        s = s & " Positive"
60    ElseIf InStr(strGroup, "-") <> 0 Then
70        s = s & " Negative"
80    End If

90    ReadableGroup = s

End Function
Public Function ReadableStatus(ByVal strStatus As String) As String
      '0 - Not Eligible
      '1 - Eligible
      '2 - Unknown Patient

10    Select Case strStatus
      Case "0": ReadableStatus = "Not Eligible"
20    Case "1": ReadableStatus = "Eligible"
30    Case "2": ReadableStatus = "Unknown Patient"
40    End Select

End Function


Public Function ReadableValidity(ByVal strValid As String) As String
      '0 - Not Eligible
      '1 - Eligible
      '2 - Unknown Patient

10    Select Case strValid
      Case "0": ReadableValidity = "Not Valid"
20    Case "1": ReadableValidity = "Valid"
30    Case "2": ReadableValidity = "Unknown Patient"
40    End Select

End Function


Public Function FormatTime(ByVal dt As String) As String

      Dim d As String
      Dim t As String
      Dim result As String

10    result = ""
20    t = ""
30    d = ""
40    If Len(dt) = 10 Then
50        d = Replace(dt, ".", "/")
60        d = Format$(d, "dd/mmm/yyyy")
70    ElseIf Len(dt) = 16 Then
80        d = Left$(dt, 10)
90        d = Replace(d, ".", "/")
100       d = Format$(d, "dd/mmm/yyyy")

110       t = Right$(dt, 5)
120       t = Replace(t, ".", ":")

130   End If

140   result = Trim$(d & " " & t)

150   If IsDate(result) Then
160       FormatTime = result
170   End If

End Function


Public Function GroupKnown(ByVal Chart As String, _
                           ByVal SampleID As String, _
                           PatName As String) _
                           As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GroupKnown_Error

20    GroupKnown = ""

30    If Trim$(Chart) = "" Then Exit Function

40    sql = "Select * from PatientDetails where " & _
            "PatNum = '" & Chart & "' " & _
            "and Name = '" & AddTicks(PatName) & "' " & _
            "and LabNumber <> '" & SampleID & "' " & _
            "Order by datetime desc"
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    If tb.EOF Then Exit Function
80    GroupKnown = tb!fGroup & ""

90    Exit Function

GroupKnown_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BBankShared", "GroupKnown", intEL, strES, sql


End Function


Public Sub LogReasonWhy(ByVal Reason As String, _
                        ByVal Test As String)

      Dim sql As String

10    On Error GoTo LogReasonWhy_Error

20    sql = "Insert into UnlockReason " & _
            "(Reason, Test, DateTime, UserName) VALUES " & _
            "('" & AddTicks(Reason) & "', " & _
            "'" & Test & "', " & _
            "'" & Format(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
            "'" & UserName & "');"
30    CnxnBB(0).Execute sql

40    Exit Sub

LogReasonWhy_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "BBankShared", "LogReasonWhy", intEL, strES, sql


End Sub
Public Function IsPartialPresent(ByVal PackNumber As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo IsPartialPresent_Error

20    IsPartialPresent = False

30    sql = "Select * from PartialPacks where " & _
            "Number = '" & PackNumber & "'"

40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql

60    If Not tb.EOF Then
70        IsPartialPresent = True
80    End If

90    Exit Function

IsPartialPresent_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BBankShared", "IsPartialPresent", intEL, strES, sql


End Function

Public Function TagIsPresent(ByVal UnitNumber As String, ByVal DateExpiry As Date) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo TagIsPresent_Error

20    sql = "Select * from UnitNotes where " & _
            "UnitNumber = '" & UnitNumber & "' And DateExpiry = '" & Format(DateExpiry, "dd/MMM/yyyy HH:mm") & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    TagIsPresent = Not tb.EOF

60    Exit Function

TagIsPresent_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "BBankShared", "TagIsPresent", intEL, strES, sql


End Function
Public Function AntigenDescription(ByVal code As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo AntigenDescription_Error

20    sql = "Select Antigen from AntibodyAntigen where " & _
            "Code = '" & code & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If Not tb.EOF Then
60        AntigenDescription = tb!Antigen & ""
70    Else
80        AntigenDescription = ""
90    End If

100   Exit Function

AntigenDescription_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "BBankShared", "AntigenDescription", intEL, strES, sql


End Function

Public Sub ArchiveTable(ByVal TableName As String, _
                        ByVal SampleID As String, _
                        Optional ByVal sql As String = "")

      Dim tb As Recordset
      Dim tbArc As Recordset
      Dim f As Field

10    On Error GoTo ArchiveTable_Error

20    If sql = "" Then
30        sql = "SELECT * FROM " & TableName & " WHERE " & _
                "SampleID = '" & SampleID & "'"
40    End If

50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        sql = "SELECT * FROM " & TableName & "Arc WHERE 0 = 1"
90        Set tbArc = New Recordset
100       RecOpenServer 0, tbArc, sql

110       Do While Not tb.EOF
120           tbArc.AddNew

130           For Each f In tbArc.Fields
140               If UCase$(f.Name) <> "ROWGUID" And _
                     UCase$(f.Name) <> "ARCHIVEDATETIME" And _
                     UCase$(f.Name) <> "ARCHIVEDBY" Then

150                   tbArc(f.Name) = tb(f.Name)

160               End If
170           Next
180           tbArc!ArchiveDateTime = Format$(Now, "yyyy mm dd hh:mm:ss")
190           tbArc!ArchivedBy = UserName
200           tbArc.Update

210           tb.MoveNext
220       Loop
230   End If

240   Exit Sub

ArchiveTable_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "BBankShared", "ArchiveTable", intEL, strES, sql

End Sub

Public Sub ArchiveTableBB(ByVal TableName As String, _
                          ByVal SampleID As String, _
                          Optional ByVal sql As String = "")

      Dim tb As Recordset
      Dim tbArc As Recordset
      Dim f As Field

10    On Error GoTo ArchiveTableBB_Error

20    If sql = "" Then
30        sql = "SELECT * FROM " & TableName & " WHERE " & _
                "SampleID = '" & SampleID & "'"
40    End If

50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    If Not tb.EOF Then
80        sql = "SELECT * FROM " & TableName & "Arc WHERE 0 = 1"
90        Set tbArc = New Recordset
100       RecOpenServerBB 0, tbArc, sql

110       Do While Not tb.EOF
120           tbArc.AddNew

130           For Each f In tbArc.Fields
140               If UCase$(f.Name) <> "ROWGUID" And _
                     UCase$(f.Name) <> "ARCHIVEDATETIME" And _
                     UCase$(f.Name) <> "ARCHIVEDBY" Then

150                   tbArc(f.Name) = tb(f.Name)

160               End If
170           Next
180           tbArc!ArchiveDateTime = Format$(Now, "yyyy mm dd hh:mm:ss")
190           tbArc!ArchivedBy = UserName
200           tbArc.Update

210           tb.MoveNext
220       Loop
230   End If

240   Exit Sub

ArchiveTableBB_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "BBankShared", "ArchiveTableBB", intEL, strES, sql

End Sub



Public Function AddTicks(ByVal s As String) As String

10    s = Trim$(s)

20    s = Replace(s, "'", "''")

30    AddTicks = s

End Function
Public Function PasswordHasBeenUsed(ByVal Password As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo PasswordHasBeenUsed_Error

20    PasswordHasBeenUsed = False

30    sql = "SELECT Password FROM Users WHERE " & _
            "Password = '" & Password & "' " & _
            "COLLATE SQL_Latin1_General_CP1_CS_AS"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        PasswordHasBeenUsed = True
80    Else
90        sql = "SELECT Password FROM UsersArc WHERE " & _
                "Password = '" & Password & "' " & _
                "COLLATE SQL_Latin1_General_CP1_CS_AS"
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       If Not tb.EOF Then
130           PasswordHasBeenUsed = True
140       End If
150   End If

160   Exit Function

PasswordHasBeenUsed_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "BBankShared", "PasswordHasBeenUsed", intEL, strES, sql

End Function


Public Function CodeHasBeenUsed(ByVal code As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CodeHasBeenUsed_Error

20    CodeHasBeenUsed = False

30    sql = "SELECT Code FROM Users WHERE " & _
            "Code = '" & code & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        CodeHasBeenUsed = True
80    End If

90    Exit Function

CodeHasBeenUsed_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BBankShared", "CodeHasBeenUsed", intEL, strES, sql

End Function



Public Function NameHasBeenUsed(ByVal UserName As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo NameHasBeenUsed_Error

20    NameHasBeenUsed = False

30    sql = "SELECT Name FROM Users WHERE " & _
            "Name = '" & AddTicks(UserName) & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        NameHasBeenUsed = True
80    End If

90    Exit Function

NameHasBeenUsed_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "BBankShared", "NameHasBeenUsed", intEL, strES, sql

End Function

Public Function AllLowerCase(stringToCheck As String) As Boolean

10    AllLowerCase = StrComp(stringToCheck, LCase$(stringToCheck), vbBinaryCompare) = 0

End Function


Public Function AllUpperCase(stringToCheck As String) As Boolean

10    AllUpperCase = StrComp(stringToCheck, UCase$(stringToCheck), vbBinaryCompare) = 0

End Function


Public Function ContainsAlpha(ByVal s As String) As Boolean

      Dim n As Integer
      Dim strTestL As String
      Dim strTestU As String

10    strTestL = "abcdefghijklmnopqrstuvwxyz"
20    strTestU = UCase$(strTestL)
30    ContainsAlpha = False

40    For n = 1 To Len(s)
50        If InStr(strTestL, Mid$(s, n, 1)) Or InStr(strTestU, Mid$(s, n, 1)) Then
60            ContainsAlpha = True
70            Exit Function
80        End If
90    Next

End Function
Public Function ContainsNumeric(ByVal s As String) As Boolean

      Dim n As Integer

10    ContainsNumeric = False
20    For n = 1 To Len(s)
30        If InStr("0123456789", Mid$(s, n, 1)) Then
40            ContainsNumeric = True
50            Exit Function
60        End If
70    Next

End Function



Public Function CheckPreviousABScreen(ByVal MRN As String) As String

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CheckPreviousABScreen_Error

20    CheckPreviousABScreen = ""

30    If Trim$(MRN) = "" Then Exit Function

40    sql = "Select AIDR from PatientDetails where " & _
            "PatNum = '" & MRN & "'"
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    If tb.EOF Then Exit Function

80    Do While Not tb.EOF
90        If Trim$(tb!AIDR & "") <> "" Then

100           If UCase$(Trim$(tb!AIDR & "")) <> "NEGATIVE" Then
110               CheckPreviousABScreen = tb!AIDR
120               Exit Function
130           End If
140       End If
150       tb.MoveNext
160   Loop

170   Exit Function

CheckPreviousABScreen_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "BBankShared", "CheckPreviousABScreen", intEL, strES, sql


End Function

Public Function QueryKnownGP(ByVal strCodeOrText As String, _
                             ByVal strHospital As String) _
                             As String
      'Returns either strCodeOrText = not known
      '        or CodeOrText = known

      Dim strOriginal As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo QueryKnownGP_Error

20    QueryKnownGP = ""
30    strOriginal = strCodeOrText

40    strCodeOrText = Trim$(UCase$(strCodeOrText))
50    If strCodeOrText = "" Then Exit Function

60    sql = "Select * from GPs where " & _
            "Code = '" & strCodeOrText & "' " & _
            "or Text = '" & strCodeOrText & "' " & _
            "and HospitalCode = '" & _
            ListCodeFor("HO", strHospital) & "'"
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If Not tb.EOF Then
100       QueryKnownGP = tb!Text & ""
110   Else
120       QueryKnownGP = strOriginal
130   End If

140   Exit Function

QueryKnownGP_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "BBankShared", "QueryKnownGP", intEL, strES, sql


End Function

Public Function QueryKnownClinician(ByVal strCodeOrText As String, _
                                    ByVal strHospital As String) _
                                    As String
      'Returns either strCodeOrText = not known
      '        or Text = known

      Dim strOriginal As String
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo QueryKnownClinician_Error

20    QueryKnownClinician = ""
30    strOriginal = strCodeOrText

40    strCodeOrText = Trim$(UCase$(strCodeOrText))
50    If strCodeOrText = "" Then Exit Function

60    sql = "Select * from Clinicians where " & _
            "Code = '" & strCodeOrText & "' " & _
            "or Text = '" & strCodeOrText & "' " & _
            "and HospitalCode = '" & _
            ListCodeFor("HO", strHospital) & "'"
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If Not tb.EOF Then
100       QueryKnownClinician = tb!Text & ""
110   Else
120       QueryKnownClinician = strOriginal
130   End If

140   Exit Function

QueryKnownClinician_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "BBankShared", "QueryKnownClinician", intEL, strES, sql


End Function


Public Function ParseForeName(ByVal Name As String) As String

10    Name = Trim$(UCase$(Name))
      Dim n As Integer
      Dim Temp As String

20    If InStr(Name, "B/O") Or _
         InStr(Name, "BABY") Then
30        ParseForeName = ""
40        Exit Function
50    End If

60    n = InStr(Name, " ")
70    If n = 0 Then
80        ParseForeName = ""
90        Exit Function
100   End If

110   Temp = Mid$(Name, n + 1)
120   If InStr(Temp, " ") Or _
         Temp Like "*[!A-Z]*" Or _
         Len(Temp) = 1 Then
130       ParseForeName = ""
140   Else
150       ParseForeName = Temp
160   End If

End Function


Public Function IsRoutine() As Boolean

      'Returns True if time now is between
      '09:30 and 16:30 Mon to Fri
      'else returns False

10    IsRoutine = False

20    If Weekday(Now) <> vbSaturday And Weekday(Now) <> vbSunday Then
30        If TimeValue(Now) > TimeValue("09:29") And _
             TimeValue(Now) < TimeValue("16:31") Then
40            IsRoutine = True
50        End If
60    End If

End Function
Public Function CheckJulian(ByVal JDate As String) As String

      'Convert Julian Date to DD/MM/YYYY

      Dim yyyy As String
      Dim dmy As String

10    CheckJulian = JDate

20    If Len(JDate) = 10 Then
30        If Left$(UCase$(JDate), 2) = "A2" And Right$(UCase$(JDate), 2) = "4A" Then
40            JDate = Mid$(JDate, 3, 6)
50        Else
60            Exit Function
70        End If
80    ElseIf Len(JDate) = 8 Then
90        If Left$(JDate, 1) = "2" And Right$(JDate, 1) = "4" Then
100           JDate = Mid$(JDate, 2, 6)
110       Else
120           Exit Function
130       End If
140   Else
150       Exit Function
160   End If

170   If Left$(JDate, 1) = "9" Then
180       yyyy = "19"
190   ElseIf Left$(JDate, 1) = "0" Then
200       yyyy = "20"
210   Else
220       Exit Function
230   End If

240   yyyy = yyyy & Mid$(JDate, 2, 2)

250   dmy = Format(DateAdd("d", Val(Mid$(JDate, 4, 3)) - 1, "01/01/" & yyyy), "dd/mm/yyyy")

260   CheckJulian = dmy

End Function



Public Function Bar2Group(ByVal Group As String) As String

      Dim s As String

10    Select Case Group
      Case "51": s = "O Pos"
20    Case "62": s = "A Pos"
30    Case "73": s = "B Pos"
40    Case "84": s = "AB Pos"
50    Case "95": s = "O Neg"
60    Case "06": s = "A Neg"
70    Case "17": s = "B Neg"
80    Case "28": s = "AB Neg"
85    Case "55": s = "O"
92    Case "66": s = "A"
94    Case "77": s = "B"
96    Case "88": s = "AB"
90    Case Else: s = ""
100   End Select

110   Bar2Group = s

End Function

Public Function bcrAllowUnit(ByVal UnitNo As String) As String

10    bcrAllowUnit = ""

20    UnitNo = UCase$(UnitNo)

30    Select Case Len(Trim$(UnitNo))
      Case 6, 7
40        bcrAllowUnit = UnitNo
50    Case 9
60        If Left$(UnitNo, 1) = "D" And Right$(UnitNo, 1) = "D" Then    'pack
70            bcrAllowUnit = Mid$(UnitNo, 2, 7)
80        End If
90    End Select

End Function





Public Function BetweenDates(ByVal Index As Integer, upto As String) As String

      Dim From As String
      Dim m As Integer

10    Select Case Index
      Case 0:    'last week
20        From = Format(DateAdd("ww", -1, Now), "dd/mm/yyyy")
30        upto = Format(Now, "dd/mm/yyyy")
40    Case 1:    'last month
50        From = Format(DateAdd("m", -1, Now), "dd/mm/yyyy")
60        upto = Format(Now, "dd/mm/yyyy")
70    Case 2:    'last fullmonth
80        From = Format(DateAdd("m", -1, Now), "dd/mm/yyyy")
90        From = "01/" & Mid$(From, 4)
100       upto = DateAdd("m", 1, From)
110       upto = Format(DateAdd("d", -1, upto), "dd/mm/yyyy")
120   Case 3:    'last quarter
130       From = Format(DateAdd("q", -1, Now), "dd/mm/yyyy")
140       upto = Format(Now, "dd/mm/yyyy")
150   Case 4:    'last full quarter
160       From = Format(DateAdd("q", -1, Now), "dd/mm/yyyy")
170       m = Val(Mid$(From, 4, 2))
180       m = ((m - 1) \ 3) * 3 + 1
190       From = "01/" & Format(m, "00") & Mid$(From, 6)
200       upto = DateAdd("q", 1, From)
210       upto = Format(DateAdd("d", -1, upto), "dd/mm/yyyy")
220   Case 5:    'year to date
230       From = "01/01/" & Format(Now, "yyyy")
240       upto = Format(Now, "dd/mm/yyyy")
250   Case 6:    'today
260       From = Format(Now, "dd/mm/yyyy")
270       upto = From
280   Case 7:    'last full year
290       From = "01/01/" & Format(DateAdd("yyyy", -1, Now), "yyyy")
300       upto = "31/12/" & Format(DateAdd("yyyy", -1, Now), "yyyy")
310   End Select

320   BetweenDates = From

End Function

Public Function CalcAge(ByVal DoB As String) As String

      Dim A As String
      Dim DaysOld As Long
      Dim MonthsOld As Long
      Dim YearsOld As Single

10    If Not IsDate(DoB) Then
20        A = ""
30    Else
40        DaysOld = DateDiff("d", DoB, Now)
50        If DaysOld < 8 Then
60            A = Format$(DaysOld) & " D"
70        Else
80            MonthsOld = DaysOld / 30.4375
90            If MonthsOld < 13 Then
100               A = Format$(MonthsOld, "##") & "M"
110           Else
120               YearsOld = Int(DaysOld / 365.25)
130               A = Format$(YearsOld, "##") & "Y"
140           End If
150       End If
160   End If

170   CalcAge = A

End Function


Public Function Convert62Date(ByVal d As String, _
                              ByVal Direction As Integer, _
                              Optional ByVal EarliestDate As Variant, _
                              Optional ByVal LatestDate As Variant) _
                              As String
      Dim dd As String
      Dim mm As String
      Dim yy As String

10    If Len(d) = 6 Then
20        dd = Left$(d, 2)
30        mm = Mid$(d, 3, 2)
40        yy = Right$(d, 2)
50        If Val(dd) < 1 Or Val(dd) > 31 Or _
             Val(mm) < 1 Or Val(mm) > 12 Then
60            Convert62Date = d
70            Exit Function
80        End If
90        d = dd & "/" & mm & "/" & yy
100   ElseIf Len(d) = 8 Then
110       d = Left$(d, 2) & "/" & Mid$(d, 3, 2) & "/" & Right$(d, 4)
120   End If
130   If Not IsDate(d) Then
140       d = ""
150       Convert62Date = d
160       Exit Function
170   Else
180       d = Format(d, "dd/mm/yyyy")
190       If Direction = FORWARD Then
200           If DateDiff("d", Now, d) < 0 Then
210               d = Format(DateAdd("yyyy", 100, d), "dd/mm/yyyy")
220           End If
230       Else
240           If DateDiff("d", Now, d) > 0 Then
250               d = Format(DateAdd("yyyy", -100, d), "dd/mm/yyyy")
260           End If
270       End If
280   End If

290   Convert62Date = d

300   If Not IsMissing(EarliestDate) Then
310       If DateDiff("d", EarliestDate, d) < 0 Then
320           Convert62Date = ""
330       End If
340   End If
350   If Not IsMissing(LatestDate) Then
360       If DateDiff("d", LatestDate, d) > 0 Then
370           Convert62Date = ""
380       End If
390   End If

End Function



Public Sub DeleteWorkList(f As Form)

      Dim Y As Integer
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo DeleteWorkList_Error

20    For Y = 1 To f.g.Rows - 1
30        f.g.row = Y
40        f.g.col = 0
50        sql = "SELECT * FROM PatientDetails WHERE " & _
                "LabNumber = '" & f.g & "' " & _
                "AND RequestFrom = "
60        Select Case Dept
          Case ANTENATAL: sql = sql & "'A'"
70        Case XMATCH: sql = sql & "'X'"
80        Case DAT: sql = sql & "'D'"
90        Case GROUP_HOLD: sql = sql & "'G'"
100       End Select
110       sql = sql & " AND Hold = 1"
120       Set tb = New Recordset
130       RecOpenServerBB 0, tb, sql
140       Do While Not tb.EOF
150           tb!Hold = False
160           tb.Update
170           tb.MoveNext
180       Loop
190   Next

200   Exit Sub

DeleteWorkList_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "BBankShared", "DeleteWorkList", intEL, strES, sql

End Sub

Public Function Event2Code(ByVal s As String) As String

      Dim E As String

10    Select Case UCase$(Left$(s & "         ", 9))
      Case "ALLOCATED": E = "A"
20    Case "BEING TRA": E = "B"
30    Case "RECEIVED ": E = "C"
40    Case "DESTROYED": E = "D"
50    Case "AMENDMENT": E = "E"
60    Case "DISPATCHE": E = "F"
70    Case "E RECEIVED": E = "G"
80    Case "E ISSUED": E = "V"
90    Case "ISSUED   ": E = "I"
100   Case "EXPIRED  ": E = "J"
110   Case "AWAITING ": E = "K"
120   Case "PENDING  ": E = "P"
130   Case "IN STOCK ": E = "Q"
140   Case "RESTOCKED": E = "R"
150   Case "TRANSFUSE": E = "S"
160   Case "RETURNED ": E = "T"
170   Case "XMATCHED ": E = "X"
180   Case "CROSS MAT": E = "X"
190   Case "REMOVED P": E = "Y"    'removed pending transfusion
200   Case "E TRANSFUSE": E = "Z"
          '  Case "MOVED": e = "M"
210   Case Else: E = " "
220   End Select
      'Transfusion
230   Event2Code = E

End Function
Public Function ChkDig(ByVal Unit As String) As String

      'see https://www.the-stationery-office.co.uk/nbs/rdbk2001/blood38.htm for explanation

      Dim n1 As Integer
      Dim n2 As Integer
      Dim n3 As Integer
      Dim n4 As Integer
      Dim n5 As Integer
      Dim n6 As Integer
      Dim Tot As Integer


10    n1 = Val(Left$(Unit, 1))
20    n2 = Val(Mid$(Unit, 2, 1))
30    n3 = Val(Mid$(Unit, 3, 1))
40    n4 = Val(Mid$(Unit, 4, 1))
50    n5 = Val(Mid$(Unit, 5, 1))
60    n6 = Val(Mid$(Unit, 6, 1))


70    n1 = n1 * 7
80    n2 = n2 * 6
90    n3 = n3 * 5
100   n4 = n4 * 4
110   n5 = n5 * 3
120   n6 = n6 * 2

130   Tot = n1 + n2 + n3 + n4 + n5 + n6
140   n1 = (Tot / 11)
150   n2 = (11 * n1) - Tot


160   If n2 < 0 Then
170       If n2 = -1 Then
180           ChkDig = "X"
190           Exit Function
200       Else
210           n2 = 11 + n2
220           ChkDig = n2
230           Exit Function
240       End If
250   End If

260   n2 = Abs(n2)

270   ChkDig = n2

End Function

Public Function EventCode2Text(ByVal strEventCode As String) _
       As String

10    strEventCode = UCase$(strEventCode)

20    Select Case strEventCode
      Case "A": EventCode2Text = "Allocated"
30    Case "B": EventCode2Text = "Being Transfused"
40    Case "C": EventCode2Text = "Received into Stock"
50    Case "D": EventCode2Text = "Destroyed"
60    Case "E": EventCode2Text = "Amendment"
70    Case "F": EventCode2Text = "Dispatched"
80    Case "G": EventCode2Text = "Received into Emergency Stock"
90    Case "H": EventCode2Text = "Issued as Emergency"
100   Case "I": EventCode2Text = "Issued"
110   Case "J": EventCode2Text = "Expired"
120   Case "K": EventCode2Text = "Awaiting Release"
130   Case "P": EventCode2Text = "Pending"
140   Case "R": EventCode2Text = "Restocked"
150   Case "S": EventCode2Text = "Transfused"
160   Case "T": EventCode2Text = "Returned to Supplier"
165   Case "V": EventCode2Text = "E Issued"
170   Case "X": EventCode2Text = "Cross matched"
180   Case "Y": EventCode2Text = "Removed Pending Transfusion"
190   Case "Z": EventCode2Text = "Transfused as Emergency"
          '  Case "M": EventCode2Text = "Moved to Emergency ONeg"
          '  Case "L": EventCode2Text = "Labeled for Emergency ONeg"
200   Case Else: EventCode2Text = "?????"
210   End Select

End Function






Public Function From2Text(ByVal strCode As String) As String

10    Select Case strCode
      Case "X": From2Text = "X-Match"
20    Case "G": From2Text = "G & H"
30    Case "A": From2Text = "A/Natal"
40    Case "D": From2Text = "D.A.T."
50    Case Else: From2Text = ""
60    End Select

End Function

Sub grh2image(Gr As String, Rh As String)

      Dim s As String
      Dim Rhesus As String

10    Rhesus = Rh
20    If Rh = "O" Then Rhesus = "-"
30    If UCase$(Rh) = "POS" Then Rhesus = "+"
40    If UCase$(Rh) = "NEG" Then Rhesus = "-"

50    Gr = Replace(Gr, "+", "")
60    Gr = Replace(Gr, "-", "")

70    s = "gr"
80    s = s & Trim$(Gr)
90    If Rhesus = "-" Then s = s & "Neg"
100   If Rhesus = "+" Then s = s & "Pos"
110   gbitmap = LCase$(Trim$(Gr)) & Rhesus
120   If Rhesus = "" Or Gr = "" Or Gr = "Er" Then s = "grPrev": gbitmap = ""

130   'frmxmatch.iprevious.Picture = frmMain.ImageList1.ListImages(s).Picture

End Sub

Public Function ImageForGRh(ByVal Group As String, _
                            ByVal Rhesus As String) _
                            As IPictureDisp

      Dim s As String
      Dim Rh As String

10    Rh = Rhesus
20    If Rhesus = "O" Then Rh = "-"
30    If UCase$(Rhesus) = "POS" Then Rh = "+"
40    If UCase$(Rhesus) = "NEG" Then Rh = "-"

50    s = "gr"
60    s = s & Trim$(Group)
70    If Rh = "-" Then s = s & "Neg"
80    If Rh = "+" Then s = s & "Pos"
90    gbitmap = LCase$(Trim$(Group)) & Rh
100   If Rh = "" Or Group = "" Or Group = "Er" Then s = "grPrev": gbitmap = ""

110   'Set ImageForGRh = frmMain.ImageList1.ListImages(s).Picture

End Function


Public Function Group2Bar(ByVal strGrpRH As String) _
       As String

      Dim strBar As String

10    strGrpRH = UCase$(Left$(strGrpRH & "      ", 6))
20    Select Case strGrpRH
      Case "O POS ": strBar = "51"
30    Case "A POS ": strBar = "62"
40    Case "B POS ": strBar = "73"
50    Case "AB POS": strBar = "84"
60    Case "O NEG ": strBar = "95"
70    Case "A NEG ": strBar = "06"
80    Case "B NEG ": strBar = "17"
90    Case "AB NEG": strBar = "28"
91    Case "O     ": strBar = "55"
92    Case "A     ": strBar = "66"
93    Case "B     ": strBar = "77"
94    Case "AB    ": strBar = "88"
100   Case Else: Beep: strBar = ""
110   End Select

120   Group2Bar = strBar

End Function

Public Function Group2Index(ByVal strGroup As String) As Integer

      Dim intOAB As Integer

10    Select Case strGroup
      Case "O Neg": intOAB = 1
20    Case "O Pos": intOAB = 2
30    Case "A Neg": intOAB = 3
40    Case "A Pos": intOAB = 4
50    Case "B Neg": intOAB = 5
60    Case "B Pos": intOAB = 6
70    Case "AB Neg": intOAB = 7
80    Case "AB Pos": intOAB = 8
90    Case "O Du Pos", "O D- C/E+": intOAB = 9
100   Case "A Du Pos", "A D- C/E+": intOAB = 10
110   Case "B Du Pos", "B D- C/E+": intOAB = 11
120   Case "AB Du Pos", "AB D-C/E+": intOAB = 12
130   Case Else: intOAB = 0
140   End Select

150   Group2Index = intOAB

End Function



Sub image2grh(pgr As String, prh As String)

10    Select Case LCase$(gbitmap)
      Case "o-": pgr = "O": prh = "-"
20    Case "o+": pgr = "O": prh = "+"
30    Case "a-": pgr = "A": prh = "-"
40    Case "a+": pgr = "A": prh = "+"
50    Case "b-": pgr = "B": prh = "-"
60    Case "b+": pgr = "B": prh = "+"
70    Case "ab-": pgr = "AB": prh = "-"
80    Case "ab+": pgr = "AB": prh = "+"
90    Case Else: pgr = "": prh = ""
100   End Select

End Sub

Public Function Initial2Upper(ByVal strAnyCase As String) _
       As String

      Dim intN As Integer

10    strAnyCase = Trim$(strAnyCase)
20    If strAnyCase = "" Then
30        Initial2Upper = ""
40        Exit Function
50    End If

60    strAnyCase = LCase(strAnyCase)
70    strAnyCase = UCase$(Left$(strAnyCase, 1)) & Mid$(strAnyCase, 2)

80    For intN = 1 To Len(strAnyCase) - 1
90        If Mid$(strAnyCase, intN, 1) = " " Or Mid$(strAnyCase, intN, 1) = "'" Then
100           strAnyCase = Left$(strAnyCase, intN) & UCase$(Mid$(strAnyCase, intN + 1, 1)) & Mid$(strAnyCase, intN + 2)
110       End If
120       If intN > 1 Then
130           If Mid$(strAnyCase, intN, 1) = "c" And Mid$(strAnyCase, intN - 1, 1) = "M" Then
140               strAnyCase = Left$(strAnyCase, intN) & UCase$(Mid$(strAnyCase, intN + 1, 1)) & Mid$(strAnyCase, intN + 2)
150           End If
160       End If
170   Next

180   Initial2Upper = strAnyCase

End Function

Public Sub LoadGenotype(ByVal strLabNum As String)

      Dim tb As Recordset
      Dim sql As String
      Dim intN As Integer

10    On Error GoTo LoadGenotype_Error

20    sql = "select * from genotype where " & _
            "labnumber = '" & strLabNum & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60        For intN = 0 To 4
70            If Not IsNull(tb(intN + 1)) Then
80                fgenotype.p(intN) = tb(intN + 1)
90            End If
100       Next
110   End If

120   fgenotype.lLabNumber = strLabNum

130   Exit Sub

LoadGenotype_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "BBankShared", "LoadGenotype", intEL, strES, sql


End Sub

Public Function QueryChecked(ByVal strBarCode As String, _
                             ByVal strPackNumber As String) _
                             As Integer

      'return true if checked

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo QueryChecked_Error

20    QueryChecked = False

30    sql = "Select Checked from Latest where " & _
            "BarCode = '" & strBarCode & "' " & _
            "and ISBT128 = '" & strPackNumber & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70        If Not IsNull(tb!Checked) Then
80            QueryChecked = tb!Checked
90        End If
100   End If

110   Exit Function

QueryChecked_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "BBankShared", "QueryChecked", intEL, strES, sql


End Function


Public Function SetFormPrinter() As Boolean
      'returns true if ok

      Dim TargetPrinter As String
      Dim xFound As Boolean
      Dim Px As Printer

10    On Error GoTo SetFormPrinter_Error

20    xFound = False
30    TargetPrinter = UCase$(TransfusionForm)
40    For Each Px In Printers
50        If UCase$(Px.DeviceName) = TargetPrinter Then
60            Set Printer = Px
70            xFound = True
80            Exit For
90        End If
100   Next

110   SetFormPrinter = xFound

120   Exit Function

SetFormPrinter_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "BBankShared", "SetFormPrinter", intEL, strES

End Function

Public Sub LogError(ByVal ModuleName As String, _
                    ByVal ProcedureName As String, _
                    ByVal ErrorLineNumber As Integer, _
                    ByVal ErrorDescription As String, _
                    Optional ByVal SQLStatement As String, _
                    Optional ByVal EventDesc As String)

      Dim sql As String
      Dim MyMachineName As String
      Dim Vers As String
      Dim UID As String

10    On Error Resume Next

20    UID = AddTicks(UserName)

30    SQLStatement = AddTicks(SQLStatement)

40    ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "[MSSQL]")
50    ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver]", "[SQL]")
60    ErrorDescription = AddTicks(ErrorDescription)

70    Vers = App.Major & "-" & App.Minor & "-" & App.Revision

80    MyMachineName = vbGetComputerName()

90    sql = "IF NOT EXISTS " & _
            "    (SELECT * FROM ErrorLog WHERE " & _
            "     ModuleName = '" & ModuleName & "' " & _
            "     AND ProcedureName = '" & ProcedureName & "' " & _
            "     AND ErrorLineNumber = '" & ErrorLineNumber & "' " & _
            "     AND AppName = '" & App.EXEName & "' " & _
            "     AND AppVersion = '" & Vers & "' ) " & _
            "  INSERT INTO ErrorLog (" & _
            "    ModuleName, ProcedureName, ErrorLineNumber, SQLStatement, " & _
            "    ErrorDescription, UserName, MachineName, Eventdesc, AppName, AppVersion, EventCounter, eMailed) " & _
            "  VALUES  ('" & ModuleName & "', " & _
            "           '" & ProcedureName & "', " & _
            "           '" & ErrorLineNumber & "', " & _
            "           '" & SQLStatement & "', " & _
            "           '" & ErrorDescription & "', " & _
            "           '" & UID & "', " & _
            "           '" & MyMachineName & "', " & _
            "           '" & AddTicks(EventDesc) & "', " & _
            "           '" & App.EXEName & "', " & _
            "           '" & Vers & "', " & _
            "           '1', '0') " & _
      "ELSE "
100   sql = sql & "  UPDATE ErrorLog " & _
            "  SET SQLStatement = '" & SQLStatement & "', " & _
            "  ErrorDescription = '" & ErrorDescription & "', " & _
            "  MachineName = '" & MyMachineName & "', " & _
            "  DateTime = getdate(), " & _
            "  UserName = '" & UID & "', " & _
            "  EventCounter = COALESCE(EventCounter, 0) + 1 " & _
            "  WHERE ModuleName = '" & ModuleName & "' " & _
            "  AND ProcedureName = '" & ProcedureName & "' " & _
            "  AND ErrorLineNumber = '" & ErrorLineNumber & "' " & _
            "  AND AppName = '" & App.EXEName & "' " & _
            "  AND AppVersion = '" & Vers & "'"

110   Cnxn(0).Execute sql

End Sub

Public Function vbGetComputerName() As String

      'Gets the name of the machine
      Const MAXSIZE As Integer = 256
      Dim sTmp As String * MAXSIZE
      Dim lLen As Long

10    lLen = MAXSIZE - 1
20    If (GetComputerName(sTmp, lLen)) Then
30        vbGetComputerName = Left$(sTmp, lLen)
40    Else
50        vbGetComputerName = ""
60    End If

End Function

Public Function SetLabelPrinter() As Boolean
      'returns true if ok

      Dim TargetPrinter As String
      Dim xFound As Boolean
      Dim Px As Printer

10    xFound = False

20    TargetPrinter = UCase$(TransfusionLabel)
30    For Each Px In Printers
40        If UCase$(Px.DeviceName) = TargetPrinter Then
50            Set Printer = Px
60            xFound = True
70            Exit For
80        End If
90    Next

100   SetLabelPrinter = xFound

End Function
Public Function StripComment(ByVal strComment As String) As String

      Dim intN As Integer
      Dim strNew As String

10    strComment = Trim$(strComment)
20    If strComment = "" Then StripComment = "": Exit Function

30    strNew = ""
40    For intN = 1 To Len(strComment)
50        If Mid$(strComment, intN, 1) <> " " Then
60            strNew = strNew & Mid$(strComment, intN, 1)
70        Else
80            Do While Mid$(strComment, intN + 1, 1) = " "
90                intN = intN + 1
100           Loop
110           strNew = strNew & Mid$(strComment, intN, 1)
120       End If
130   Next

140   StripComment = strNew

End Function

Public Function FormatString(strDestString As String, _
                             intNumChars As Integer, _
                             Optional strSeperator As String = "", _
                             Optional intAlign As PrintAlignContants = Alignleft) As String

      '**************intAlign = 0 --> Left Align
      '**************intAlign = 1 --> Center Align
      '**************intAlign = 2 --> Right Align
      Dim intPadding As Integer
10    intPadding = 0

20    If Len(strDestString) > intNumChars Then
30        FormatString = Mid$(strDestString, 1, intNumChars) & strSeperator
40    ElseIf Len(strDestString) < intNumChars Then
          Dim i As Integer
          Dim intStringLength As String
50        intStringLength = Len(strDestString)
60        intPadding = intNumChars - intStringLength

70        If intAlign = PrintAlignContants.Alignleft Then
80            strDestString = strDestString & String$(intPadding, " ")  '& " "
90        ElseIf intAlign = AlignCenter Then
100           If (intPadding Mod 2) = 0 Then
110               strDestString = String$(intPadding / 2, " ") & strDestString & String$(intPadding / 2, " ")
120           Else
130               strDestString = String$((intPadding - 1) / 2, " ") & strDestString & String$((intPadding - 1) / 2 + 1, " ")
140           End If
150       ElseIf intAlign = AlignRight Then
160           strDestString = String$(intPadding, " ") & strDestString
170       End If

180       strDestString = strDestString & strSeperator
190       FormatString = strDestString
200   Else
210       strDestString = strDestString & strSeperator
220       FormatString = strDestString
230   End If

End Function
Public Sub FillWards(ByVal cmb As ComboBox, ByVal HospitalName As String)

      Dim strHospitalCode As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillWards_Error

20    strHospitalCode = ListCodeFor("HO", HospitalName)

30    sql = "SELECT DISTINCT Text, MIN(ListOrder) L FROM Wards WHERE " & _
            "HospitalCode = '" & strHospitalCode & "' " & _
            "AND InUse = 1 " & _
            "GROUP BY Text " & _
            "ORDER BY L"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    With cmb
70        .Clear
80        Do While Not tb.EOF
90            .AddItem tb!Text & ""
100           tb.MoveNext
110       Loop
120   End With

130   cmb = "GP"

140   Exit Sub

FillWards_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "BBankShared", "FillWards", intEL, strES, sql

End Sub

Public Function MarkGridRow(ByVal flxGrid As MSFlexGrid, _
                            ByVal GridRow As Integer, _
                            ByVal BackColor As Long, _
                            ByVal ForeColor As Long, _
                            ByVal FontStrikeThru As Boolean, _
                            ByVal FontBold As Boolean, _
                            ByVal FontItalic As Boolean) As Boolean

      Dim X As Integer

10    If GridRow > flxGrid.Rows Then Exit Function
20    If GridRow = 0 Then Exit Function
30    flxGrid.row = GridRow
40    For X = 0 To flxGrid.Cols - 1
50        With flxGrid
60            .col = X
70            .CellBackColor = BackColor
80            .CellForeColor = ForeColor
90            .CellFontStrikeThrough = FontStrikeThru
100           .CellFontBold = FontBold
110           .CellFontItalic = FontItalic
120       End With
130   Next

140   MarkGridRow = True

End Function
Public Sub AutosizeGridColumns(ByRef msFG As MSFlexGrid)
      Dim intRow As Long, intCol As Long
      Dim intColumnSize As Single
      Const sngPadding As Single = 200
10    intColumnSize = 0
20    With msFG
30        For intCol = 0 To .Cols - 1
40            intColumnSize = 0
50            For intRow = 1 To .Rows - 1
60                If intColumnSize < Printer.TextWidth(.TextMatrix(intRow, intCol)) Then
70                    intColumnSize = Printer.TextWidth(.TextMatrix(intRow, intCol))
80                End If

90            Next intRow
100           .ColWidth(intCol) = intColumnSize + sngPadding
110       Next intCol
120   End With
End Sub

Public Function VI(KeyAscii As Integer, _
                   iv As InputValidation, _
                   Optional NextFieldOnEnter As Boolean) As Integer

      Dim sTemp As String

10    sTemp = Chr$(KeyAscii)
20    If KeyAscii = 13 Then    'Enter Key
30        If NextFieldOnEnter = True Then
40            VI = 9    'Return Tab Keyascii if User Selected NextFieldOnEnter Option
50        Else
60            VI = 13
70        End If
80        Exit Function
90    ElseIf KeyAscii = 8 Then    'BackSpace
100       VI = 8
110       Exit Function
120   End If

      ' turn input to upper case

130   Select Case iv
      Case 0:    'NumbersFullstopDash
140       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", "-", "<", ">"
150           VI = Asc(sTemp)
160       Case Else
170           VI = 0
180       End Select

190   Case 1:    'Characters Only
200       Select Case sTemp
          Case " ", "-"
210           VI = Asc(sTemp)
220       Case "A" To "Z"
230           VI = Asc(sTemp)
240       Case "a" To "z"
250           VI = Asc(sTemp)
260       Case Else
270           VI = 0
280       End Select

290   Case 2:    'Y or N Only
300       sTemp = UCase$(Chr$(KeyAscii))
310       Select Case sTemp
          Case "Y", "N"
320           VI = Asc(sTemp)
330       Case Else
340           VI = 0
350       End Select

360   Case 3:    'AlphaNumeric Only...No Apostrophe
370       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", _
               " ", "/", ";", ":", "\", "-", ">", "<", "(", ")", "@", _
               "%", "!", """", "+", "^", "~", "`", "Ç", "´", "Ã", "Á", _
               "Â", "È", "É", "Ê", "Ì", "Í", "Î", "Ò", "Ó", "Ô", "Õ", _
               "Ù", "Ú", "Û", "Ü", "à", "á", "â", "ã", "ç", "è", "é", _
               "ê", "ì", "í", "î", "ò", "ó", "ô", "õ", "ö", "ù", "ú", _
               "û", "ü", "Æ", "æ", ",", "?", "=", "*", "#"
380           VI = Asc(sTemp)
390       Case "A" To "Z"
400           VI = Asc(sTemp)
410       Case "a" To "z"
420           VI = Asc(sTemp)
430       Case Else
440           VI = 0
450       End Select

460   Case 4:    'AlphaNumeric Only...With Apostophe
470       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", " ", "'"
480           VI = Asc(sTemp)
490       Case "A" To "Z"
500           VI = Asc(sTemp)
510       Case "a" To "z"
520           VI = Asc(sTemp)
530       Case Else
540           VI = 0
550       End Select

560   Case 5:    'Numbers Only
570       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
580           VI = Asc(sTemp)
590       Case Else
600           VI = 0
610       End Select
620   Case 6:    'Date only
630       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/"
640           VI = Asc(sTemp)
650       Case Else
660           VI = 0
670       End Select
680   End Select

690   If VI = 0 Then Beep

End Function

Public Sub PrintText(ByVal Cx As Integer, ByVal cy As Integer, ByVal Text As String, _
                     Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                     Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                     Optional FontColor As ColorConstants = vbBlack)

      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------

10    Printer.CurrentX = Cx
20    Printer.CurrentY = cy
30    Printer.FontSize = FontSize
40    Printer.FontBold = FontBold
50    Printer.FontItalic = FontItalic
60    Printer.FontUnderLine = FontUnderLine
70    Printer.ForeColor = FontColor
80    Printer.Print Text;

End Sub


Public Sub TextOut(ByVal Text As String, _
                   Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                   Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                   Optional FontColor As ColorConstants = vbBlack)

      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------

10    Printer.FontSize = FontSize
20    Printer.FontBold = FontBold
30    Printer.FontItalic = FontItalic
40    Printer.FontUnderLine = FontUnderLine
50    Printer.ForeColor = FontColor
60    Printer.Print Text;

End Sub

Public Function MaskInput(KeyAscii As Integer, Text As String, InputMask As String) As Integer

      '---------------------------------------------------------------------------------------
      ' Procedure : MaskInput
      ' DateTime  : 06/06/2008 15:28
      ' Author    : Babar Shahzad
      ' Purpose   : Masks input and doesnt allow user to add more than masked number of chars.
      '               X,x = Alphabets
      '               # = Number
      '               all other chars should match mask
      '---------------------------------------------------------------------------------------

10    If KeyAscii = 8 Then
20        MaskInput = KeyAscii
30        Exit Function
40    End If

50    If Len(Text) = Len(InputMask) Then
60        MaskInput = 0
70        Exit Function
80    End If

90    If Mid$(InputMask, Len(Text) + 1, 1) = "X" Then
100       If KeyAscii >= 65 And KeyAscii <= 90 Then
110           MaskInput = KeyAscii
120           Exit Function
130       ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
140           MaskInput = KeyAscii - 32
150           Exit Function
160       Else
170           MaskInput = 0
180           Exit Function
190       End If
200   ElseIf Mid$(InputMask, Len(Text) + 1, 1) = "x" Then
210       If KeyAscii >= 65 And KeyAscii <= 90 Then
220           MaskInput = KeyAscii + 32
230           Exit Function
240       ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
250           MaskInput = KeyAscii
260           Exit Function
270       Else
280           MaskInput = 0
290           Exit Function
300       End If
310   ElseIf Mid$(InputMask, Len(Text) + 1, 1) = "#" Then
320       If KeyAscii >= 48 And KeyAscii <= 57 Then
330           MaskInput = KeyAscii
340           Exit Function
350       Else
360           MaskInput = 0
370           Exit Function
380       End If
390   Else
          'FOR ALL OTHER CHARACTERS
400       If KeyAscii = Asc(Mid$(InputMask, Len(Text) + 1, 1)) Then
410           MaskInput = KeyAscii
420           Exit Function
430       Else
440           MaskInput = 0
450           Exit Function
460       End If
470   End If

End Function



Public Function LogCourierInterface(Identifier As String, MSG As udtRS) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo LogCourierInterface_Error

20    MSG.UnitNumber = Replace(MSG.UnitNumber, " ", "")

30    sql = "Select top 1 * From CourierInterface Where " & _
            "UnitNumber = '" & MSG.UnitNumber & "' And " & _
            "ProductCode = '" & MSG.ProductCode & "' And " & _
            "UnitExpiry = '" & Format(MSG.UnitExpiryDate, "dd/MMM/yyyy hh:mm:ss") & "' "

40    If Identifier = "FT" Then
50        sql = sql & " And SampleStatus = '" & MSG.SampleStatus & "'"
60    End If
70    sql = sql & " order by DateTime desc"

80    Set tb = New Recordset
90    RecOpenClientBB 0, tb, sql

100   If tb.EOF Then
110       tb.AddNew    'No Record exits for this unit then add new message record
120   Else
130       If UCase(Trim$(tb!Identifier)) <> UCase(Identifier) Then    'If latest record identifier <> this message identifier then ADD new record.
140           tb.AddNew
150       End If
160   End If

170   With tb
180       !Identifier = Identifier
190       !UnitNumber = MSG.UnitNumber
200       !ProductCode = MSG.ProductCode
210       !UnitExpiry = MSG.UnitExpiryDate
220       !UserName = MSG.UserName
230       !ActionText = MSG.ActionText
240       !Processed = 0
250       !DateTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")

260       Select Case Identifier
          Case "RS3":
270           !Location = MSG.StorageLocation
280           !UnitGroup = MSG.UnitGroup
290           !StockComment = MSG.StockComment
300           !Chart = MSG.Chart
310           !SurName = MSG.SurName
320           !ForeName = MSG.ForeName
330           If IsDate(MSG.DoB) Then
340               !DoB = MSG.DoB
350           End If
360           !Sex = MSG.Sex
370           !PatientGroup = MSG.PatientGroup
380           If IsDate(MSG.DeReservationDateTime) Then
390               !DeReservationDateTime = MSG.DeReservationDateTime
400           End If

410       Case "SM":
420           !Location = MSG.StorageLocation
430       Case "RTS":
440           !Location = MSG.StorageLocation
              '!SampleStatus = MSG.SampleStatus
450       Case "FT":
460           !Location = MSG.StorageLocation
470           !SampleStatus = MSG.SampleStatus
480           If MSG.SampleStatus = "T" Then
490               !Chart = MSG.Chart
500               !SurName = MSG.SurName
510               !ForeName = MSG.ForeName
520               If MSG.DoB <> "" Then !DoB = MSG.DoB
530               !Sex = MSG.Sex
540           End If
              'Transfusion location is missing. hardcode it or add new field in udtRS
550       Case "SU3":
560           !Location = MSG.StorageLocation
570           !UnitGroup = MSG.UnitGroup
580           !DeReservationDateTime = MSG.DeReservationDateTime
590       End Select

600       .Update

610   End With


620   Exit Function

LogCourierInterface_Error:

      Dim strES As String
      Dim intEL As Integer

630   intEL = Erl
640   strES = Err.Description
650   LogError "BBankShared", "LogCourierInterface", intEL, strES, sql

End Function



Public Function RBCAntigensGeneral(ByVal RbcBarcode As String) As String    'RedBloodCellAntigenValue
    Dim code As String
    Dim Position As Integer
    Dim Interpretation As String

10    On Error GoTo RBCAntigensGeneral_Error

20    code = Mid(RbcBarcode, 3, 16)

30    For Position = 1 To Len(code)
        'Position = Position + 1
40      Interpretation = Trim(Interpretation) & " " & RBCRT009Interpretation(Position, Mid(code, Position, 1))
50    Next
60    code = Mid(RbcBarcode, 19, 2)
70    Interpretation = Trim(Interpretation) & " " & RBCRT011(code)
80    RBCAntigensGeneral = Interpretation

90    Exit Function

RBCAntigensGeneral_Error:

    Dim strES As String
    Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "fNewStockBar", "RBCAntigensGeneral", intEL, strES

End Function



Public Function PlateletHLAandPlateletSpecificAntigens(ByVal Barcode As String) As String
          Dim code As String
          Dim Position As Integer
          Dim Interpretation As String

'Data Structure:  &{AAAABBBBCCCCCCCCDE
'&{ Data identifier
'AAAA (0-9)   HLA-A antigens
'BBBB (0-9)   HLA-B antigens
'CCCCCCCC (0-9)  Platelet specific anitgens, IgA antigen and CMV antibody status
'D (0-9)  For future use. default 0
'E (0-9)  Info about high titered antibodies to A and B antigens

'AA
10    On Error GoTo PlateletHLAandPlateletSpecificAntigens_Error

20    PlateletHLAandPlateletSpecificAntigens = getHLA_AA(Mid(Barcode, 3, 2))
'AA
30    PlateletHLAandPlateletSpecificAntigens = PlateletHLAandPlateletSpecificAntigens & " " & getHLA_AA(Mid(Barcode, 5, 2))
'BB
40    PlateletHLAandPlateletSpecificAntigens = PlateletHLAandPlateletSpecificAntigens & " " & getHLA_BB(Mid(Barcode, 7, 2))
'BB
50    PlateletHLAandPlateletSpecificAntigens = PlateletHLAandPlateletSpecificAntigens & " " & getHLA_BB(Mid(Barcode, 9, 2)) & " "

'CCCCCCCC
60    code = Mid(Barcode, 11, 8)

70    For Position = 9 To Len(code) + 8 '8 positions
        'Position = Position + 1
80      Interpretation = Trim(Interpretation) & " " & RT014Interpretation(Position, Mid(code, Position - 8, 1))
90      Interpretation = Trim(Interpretation)
100   Next
110   PlateletHLAandPlateletSpecificAntigens = PlateletHLAandPlateletSpecificAntigens & Interpretation

'position E

120   PlateletHLAandPlateletSpecificAntigens = PlateletHLAandPlateletSpecificAntigens & " " & getRT044_Position18(Mid(Barcode, 20, 1))


130   Exit Function

PlateletHLAandPlateletSpecificAntigens_Error:

 Dim strES As String
 Dim intEL As Integer

140    intEL = Erl
150    strES = Err.Description
160    LogError "BBankShared", "PlateletHLAandPlateletSpecificAntigens", intEL, strES

End Function

Public Function getHLA_AA(ByVal strValueAA As String) As String

          Dim sql As String
          Dim tb As Recordset
          Dim strSearch As String
          
10    On Error GoTo getHLA_AA_Error

20        getHLA_AA = ""
30        strSearch = "RT014HLA_A" & strValueAA

40        sql = "Select * From Options Where " & _
              "Description = '" & strSearch & "' "
              
50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql
70        If Not tb.EOF Then
80             getHLA_AA = tb!Contents & ""
90        End If

100   Exit Function

getHLA_AA_Error:

 Dim strES As String
 Dim intEL As Integer

110    intEL = Erl
120    strES = Err.Description
130    LogError "BBankShared", "getHLA_AA", intEL, strES, sql

End Function

Public Function getHLA_BB(ByVal strValueBB As String) As String

          Dim sql As String
          Dim tb As Recordset
          Dim strSearch As String
          
10    On Error GoTo getHLA_BB_Error

20        getHLA_BB = ""
30        strSearch = "RT014HLA_B" & strValueBB

40        sql = "Select * From Options Where " & _
              "Description = '" & strSearch & "' "
              
50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql
70        If Not tb.EOF Then
80             getHLA_BB = tb!Contents & ""
90        End If

100   Exit Function

getHLA_BB_Error:

 Dim strES As String
 Dim intEL As Integer

110    intEL = Erl
120    strES = Err.Description
130    LogError "BBankShared", "getHLA_BB", intEL, strES, sql

End Function

Public Function getRT044_Position18(ByVal strValueP18 As String) As String
          Dim sql As String
          Dim tb As Recordset
          Dim strSearch As String
          
10    On Error GoTo getRT044_Position18_Error

20        getRT044_Position18 = ""
30        strSearch = "RT044Position18_Value" & strValueP18

40        sql = "Select * From Options Where " & _
              "Description like '" & strSearch & "' "
              
50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql
70        If Not tb.EOF Then
80             getRT044_Position18 = tb!Contents & ""
90        End If

100   Exit Function

getRT044_Position18_Error:

 Dim strES As String
 Dim intEL As Integer

110    intEL = Erl
120    strES = Err.Description
130    LogError "BBankShared", "getRT044_Position18", intEL, strES, sql

End Function



'[RT009]
Public Function RBCRT009Interpretation(ByVal Position As String, ByVal Value As String) As String    'GetAntigenInterpretation

          Dim Antigen1 As String
          Dim Antigen2 As String
          Dim ReturnValue As String


10    On Error GoTo RBCRT009Interpretation_Error

20    Select Case Position
          Case "1"
30      If Value = 0 Then
40          ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val0 '"C+c-E+e-"
50      ElseIf Value = 1 Then
60          ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val1 '"C+c+E+e-"
70      ElseIf Value = 2 Then
80          ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val2  '"C-c+E+e-"
90      ElseIf Value = 3 Then
100         ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val3 '"C+c-E+e+"
110     ElseIf Value = 4 Then
120         ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val4 '"C+c+E+e+"
130     ElseIf Value = 5 Then
140         ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val5 '"C-c+E+e+"
150     ElseIf Value = 6 Then
160         ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val6 '"C+c-E-e+"
170     ElseIf Value = 7 Then
180         ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val7 '"C+c+E-e+"
190     ElseIf Value = 8 Then
200         ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val8 '"C-c+E-e+"
210     Else
220         ReturnValue = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos1_Val9 ' blank / ni
230     End If

240   Case "2"    'position 2
250     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos2_Antigen1 '"K"
260     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos2_Antigen2 '"k"
270   Case "3"
280     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos3_Antigen1 '"Cw"
290     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos3_Antigen2 '"Mia"
300   Case "4"
310     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos4_Antigen1 '"M"
320     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos4_Antigen2 '"N"
330   Case "5"
340     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos5_Antigen1 '"S"
350     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos5_Antigen2 '"s"
360   Case "6"
370     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos6_Antigen1 '"U"
380     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos6_Antigen2 '"P1"
390   Case "7"
400     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos7_Antigen1 '"Lua"
410     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos7_Antigen2 '"Kpa"
420   Case "8"
430     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos8_Antigen1 '"Lea"
440     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos8_Antigen2 '"Leb"
450   Case "9"
460     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos9_Antigen1 '"Fya"
470     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos9_Antigen2 '"Fyb"
480   Case "10"
490     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos10_Antigen1 '"Jka"
500     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos10_Antigen2 '"Jkb"
510   Case "11"
520     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos11_Antigen1 '"Doa"
530     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos11_Antigen2 '"Dob"
540   Case "12"
550     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos12_Antigen1 '"Ina"
560     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos12_Antigen2 '"Cob"
570   Case "13"
580     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos13_Antigen1 '"Dia"
590     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos13_Antigen2 '"VS/V"
600   Case "14"
610     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos14_Antigen1 '"Jsa"
620     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos14_Antigen2 '"C"
630   Case "15"
640     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos15_Antigen1 '"c"
650     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos15_Antigen2 '"E"
660   Case "16"
670     Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos16_Antigen1 '"e"
680     Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Pos16_Antigen2 '"CMV"

690   End Select
700   If Val(Position) > 1 Then
710     Select Case Value
        Case "0"
720         Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value0_Antigen1 '""
730         Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value0_Antigen2 '""
740     Case "1"
750         Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value1_Antigen1 '""
760         Antigen2 = Antigen2 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value1_Antigen2 '"-"
770     Case "2"
780         Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value2_Antigen1 '""
790         Antigen2 = Antigen2 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value2_Antigen2 '"+"
800     Case "3"
810         Antigen1 = Antigen1 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value3_Antigen1 '"-"
820         Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value3_Antigen2 '""
830     Case "4"
840         Antigen1 = Antigen1 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value4_Antigen1 '"-"
850         Antigen2 = Antigen2 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value4_Antigen2 '"-"
860     Case "5"
870         Antigen1 = Antigen1 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value5_Antigen1 '"-"
880         Antigen2 = Antigen2 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value5_Antigen2 '"+"
890     Case "6"
900         Antigen1 = Antigen1 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value6_Antigen1 '"+"
910         Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value6_Antigen2 '""
920     Case "7"
930         Antigen1 = Antigen1 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value7_Antigen1 '"+"
940         Antigen2 = Antigen2 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value7_Antigen2 '"-"
950     Case "8"
960         Antigen1 = Antigen1 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen1 '"+"
970         Antigen2 = Antigen2 & pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value8_Antigen2 '"+"
980     Case "9"
990         Antigen1 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value9_Antigen1 '""
1000        Antigen2 = pubStrDS12SpecialTestingRedBloodCellAntigensGeneral_Value9_Antigen2 '""

1010    End Select
1020    ReturnValue = Antigen1 & Antigen2
1030  End If
1040  RBCRT009Interpretation = ReturnValue

1050  Exit Function

RBCRT009Interpretation_Error:

 Dim strES As String
 Dim intEL As Integer

1060   intEL = Erl
1070   strES = Err.Description
1080   LogError "BBankShared", "RBCRT009Interpretation", intEL, strES

End Function




'[RT011]
Public Function RBCRT011(ByVal Value As String) As String    'GetAntigenInterpretation2

    Dim ReturnValue As String

10    On Error GoTo RBCRT011_Error

20    Select Case Value
    Case "00"
30      ReturnValue = RBCRT011_ReturnValue(0)    ' "Information Elsewhere"

40    Case "01"
50      ReturnValue = RBCRT011_ReturnValue(1)    ' Ena

60    Case "02"
70      ReturnValue = RBCRT011_ReturnValue(2)    ' ‘N’

80    Case "03"
90      ReturnValue = RBCRT011_ReturnValue(3)    ' Vw

100   Case "04"
110     ReturnValue = RBCRT011_ReturnValue(4)  'Mur*

120   Case "05"
130     ReturnValue = RBCRT011_ReturnValue(5)  'Hut"

140   Case "06"
150     ReturnValue = RBCRT011_ReturnValue(6)    'Hil"

160   Case "07"
170     ReturnValue = RBCRT011_ReturnValue(7)    'P

180   Case "08"
190     ReturnValue = RBCRT011_ReturnValue(8)    'PP1Pk

200   Case "09"
210     ReturnValue = RBCRT011_ReturnValue(9)    'hrS

220   Case "10"
230     ReturnValue = RBCRT011_ReturnValue(10)    'hrB"

240   Case "11"
250     ReturnValue = RBCRT011_ReturnValue(11)    'f"

260   Case "12"
270     ReturnValue = RBCRT011_ReturnValue(12)    'Ce"

280   Case "13"
290     ReturnValue = RBCRT011_ReturnValue(13)    'G"

300   Case "14"
310     ReturnValue = RBCRT011_ReturnValue(14)    'Hro"

320   Case "15"
330     ReturnValue = RBCRT011_ReturnValue(15)    'CE"

340   Case "16"
350     ReturnValue = RBCRT011_ReturnValue(16)    'Ce"

360   Case "17"
370     ReturnValue = RBCRT011_ReturnValue(17)    'Cx"

380   Case "18"
390     ReturnValue = RBCRT011_ReturnValue(18)    'Ew"

400   Case "19"
410     ReturnValue = RBCRT011_ReturnValue(19)    'Dw"

420   Case "20"
430     ReturnValue = RBCRT011_ReturnValue(20)    'hrH"

440   Case "21"
450     ReturnValue = RBCRT011_ReturnValue(21)    'Goa"

460   Case "22"
470     ReturnValue = RBCRT011_ReturnValue(23)    'Rh32"

480   Case "23"
490     ReturnValue = RBCRT011_ReturnValue(23)    'Rh33"

500   Case "24"
510     ReturnValue = RBCRT011_ReturnValue(24)    'Tar

520   Case "25"
530     ReturnValue = RBCRT011_ReturnValue(25)    'Kpb"

540   Case "26"
550     ReturnValue = RBCRT011_ReturnValue(26)    'Kpc

560   Case "27"
570     ReturnValue = RBCRT011_ReturnValue(27)    'Jsb"

580   Case "28"
590     ReturnValue = RBCRT011_ReturnValue(28)    'Ula"

600   Case "29"
610     ReturnValue = RBCRT011_ReturnValue(29)    'K11

620   Case "30"
630     ReturnValue = RBCRT011_ReturnValue(30)    'K12

640   Case "31"
650     ReturnValue = RBCRT011_ReturnValue(31)    'K13

660   Case "32"
670     ReturnValue = RBCRT011_ReturnValue(32)    'K14

680   Case "33"
690     ReturnValue = RBCRT011_ReturnValue(33)    'K17

700   Case "34"
710     ReturnValue = RBCRT011_ReturnValue(34)    'K18

720   Case "35"
730     ReturnValue = RBCRT011_ReturnValue(25)    'K19

740   Case "36"
750     ReturnValue = RBCRT011_ReturnValue(36)    'K22

760   Case "37"
770     ReturnValue = RBCRT011_ReturnValue(37)    'K23

780   Case "38"
790     ReturnValue = RBCRT011_ReturnValue(38)    'K24

800   Case "39"
810     ReturnValue = RBCRT011_ReturnValue(39)    'Lub

820   Case "40"
830     ReturnValue = RBCRT011_ReturnValue(40)    'Lu3

840   Case "41"
850     ReturnValue = RBCRT011_ReturnValue(41)    'Lu4

860   Case "42"
870     ReturnValue = RBCRT011_ReturnValue(42)    'Lu5

880   Case "43"
890     ReturnValue = RBCRT011_ReturnValue(43)    'Lu6

900   Case "44"
910     ReturnValue = RBCRT011_ReturnValue(44)    'Lu7

920   Case "45"
930     ReturnValue = RBCRT011_ReturnValue(45)    'Lu8

940   Case "46"
950     ReturnValue = RBCRT011_ReturnValue(46)    'Lu11

960   Case "47"
970     ReturnValue = RBCRT011_ReturnValue(47)    'Lu12

980   Case "48"
990     ReturnValue = RBCRT011_ReturnValue(48)    'Lu13

1000  Case "49"
1010    ReturnValue = RBCRT011_ReturnValue(49)    'Lu20

1020  Case "50"
1030    ReturnValue = RBCRT011_ReturnValue(50)    'Aua

1040  Case "51"
1050    ReturnValue = RBCRT011_ReturnValue(51)    'Aub

1060  Case "52"
1070    ReturnValue = RBCRT011_ReturnValue(52)    'Fy4

1080  Case "53"
1090    ReturnValue = RBCRT011_ReturnValue(53)    'Fy5

1100  Case "54"
1110    ReturnValue = RBCRT011_ReturnValue(54)    'Fy6

1120  Case "55"
1130    ReturnValue = RBCRT011_ReturnValue(55)    'Dib

1140  Case "56"
1150    ReturnValue = RBCRT011_ReturnValue(56)    'Sda

1160  Case "57"
1170    ReturnValue = RBCRT011_ReturnValue(57)    'Wrb

1180  Case "58"
1190    ReturnValue = RBCRT011_ReturnValue(58)    'Ytb

1200  Case "59"
1210    ReturnValue = RBCRT011_ReturnValue(59)    'Xga

1220  Case "60"
1230    ReturnValue = RBCRT011_ReturnValue(60)    'Sc1

1240  Case "61"
1250    ReturnValue = RBCRT011_ReturnValue(61)    'Sc2

1260  Case "62"
1270    ReturnValue = RBCRT011_ReturnValue(62)    'Sc3

1280  Case "63"
1290    ReturnValue = RBCRT011_ReturnValue(63)    'Joa

1300  Case "64"
1310    ReturnValue = RBCRT011_ReturnValue(64)    'removed

1320  Case "65"
1330    ReturnValue = RBCRT011_ReturnValue(65)    'Hy

1340  Case "66"
1350    ReturnValue = RBCRT011_ReturnValue(66)    'Gya

1360  Case "67"
1370    ReturnValue = RBCRT011_ReturnValue(67)    'Co3

1380  Case "68"
1390    ReturnValue = RBCRT011_ReturnValue(68)    'LWa

1400  Case "69"
1410    ReturnValue = RBCRT011_ReturnValue(69)    'LWb

1420  Case "70"
1430    ReturnValue = RBCRT011_ReturnValue(70)    'Kx

1440  Case "71"
1450    ReturnValue = RBCRT011_ReturnValue(71)    'Ge2

1460  Case "72"
1470    ReturnValue = RBCRT011_ReturnValue(72)    'Ge3

1480  Case "73"
1490    ReturnValue = RBCRT011_ReturnValue(73)    'Wb

1500  Case "74"
1510    ReturnValue = RBCRT011_ReturnValue(74)    'Lsa

1520  Case "75"
1530    ReturnValue = RBCRT011_ReturnValue(75)    'Ana

1540  Case "76"
1550    ReturnValue = RBCRT011_ReturnValue(76)    'Dha

1560  Case "77"
1570    ReturnValue = RBCRT011_ReturnValue(77)    'Cra

1580  Case "78"
1590    ReturnValue = RBCRT011_ReturnValue(78)    'IFC

1600  Case "79"
1610    ReturnValue = RBCRT011_ReturnValue(79)    'Kna

1620  Case "80"
1630    ReturnValue = RBCRT011_ReturnValue(80)    'Inb

1640  Case "81"
1650    ReturnValue = RBCRT011_ReturnValue(81)    'Csa

1660  Case "82"
1670    ReturnValue = RBCRT011_ReturnValue(82)    'I

1680  Case "83"
1690    ReturnValue = RBCRT011_ReturnValue(83)    'Era

1700  Case "84"
1710    ReturnValue = RBCRT011_ReturnValue(84)    'Vel

1720  Case "85"
1730    ReturnValue = RBCRT011_ReturnValue(85)    'Lan

1740  Case "86"
1750    ReturnValue = RBCRT011_ReturnValue(86)    'Ata

1760  Case "87"
1770    ReturnValue = RBCRT011_ReturnValue(87)    'Jra

1780  Case "88"
1790    ReturnValue = RBCRT011_ReturnValue(88)    'Oka

1800  Case "89"
1810    ReturnValue = RBCRT011_ReturnValue(89)    'Wra

1820  Case "90"
1830    ReturnValue = RBCRT011_ReturnValue(90)    'blank

1840  Case "91"
1850    ReturnValue = RBCRT011_ReturnValue(91)    'blank

1860  Case "92"
1870    ReturnValue = RBCRT011_ReturnValue(92)    'blank

1880  Case "93"
1890    ReturnValue = RBCRT011_ReturnValue(93)    'blank

1900  Case "94"
1910    ReturnValue = RBCRT011_ReturnValue(94)    'blank

1920  Case "95"
1930    ReturnValue = RBCRT011_ReturnValue(95)    'blank

1940  Case "96"
1950    ReturnValue = RBCRT011_ReturnValue(96)    'HbS-  'Hemoglobin S negative

1960  Case "97"
1970    ReturnValue = RBCRT011_ReturnValue(97)    'parvovirus B19 antibody present

1980  Case "98"
1990    ReturnValue = RBCRT011_ReturnValue(98)    'IgA deficient

2000  Case "99"
2010    ReturnValue = RBCRT011_ReturnValue(99)    'No Information Provided
2020  Case Else
2030    ReturnValue = ""


2040  End Select
2050  RBCRT011 = ReturnValue

2060  Exit Function

RBCRT011_Error:

 Dim strES As String
 Dim intEL As Integer

2070   intEL = Erl
2080   strES = Err.Description
2090   LogError "BBankShared", "RBCRT011", intEL, strES
End Function


'REF 13- Special Testing:Red Blood Cell Antigens -- General
Public Function RBCAntigensFinnish(ByVal RbcBarcode As String) As String
      Dim code As String
      Dim Position As Integer
10    On Error GoTo RBCAntigensFinnish_Error

20    code = Mid(RbcBarcode, 2, 16)
      Dim Interpretation As String

30    For Position = 1 To Len(code)
40        Interpretation = Trim(Interpretation) & " " & RBCRT010Interpretation(Position, Mid(code, Position, 1))
50    Next
60    code = Mid(RbcBarcode, 18, 2)
70    Interpretation = Trim(Interpretation) & " " & RBCRT012(code)
80    RBCAntigensFinnish = Interpretation

90    Exit Function

RBCAntigensFinnish_Error:

 Dim strES As String
 Dim intEL As Integer

100    intEL = Erl
110    strES = Err.Description
120    LogError "BBankShared", "RBCAntigensFinnish", intEL, strES

End Function

'[RT010]Finnish
Public Function RBCRT010Interpretation(ByVal Position As String, ByVal Value As String) As String    'GetAntigenInterpretation

      Dim Antigen1 As String
      Dim Antigen2 As String
      Dim ReturnValue As String


10    On Error GoTo RBCRT010Interpretation_Error

20    Select Case Position
      Case "1"
30        If Value = 0 Then
40            ReturnValue = "C+c-E+e-"

50        ElseIf Value = 1 Then
60            ReturnValue = "C+c+E+e-"

70        ElseIf Value = 2 Then
80            ReturnValue = "C-c+E+e-"

90        ElseIf Value = 3 Then
100           ReturnValue = "C+c-E+e+"

110       ElseIf Value = 4 Then
120           ReturnValue = "C+c+E+e+"

130       ElseIf Value = 5 Then
140           ReturnValue = "C-c+E+e+"

150       ElseIf Value = 6 Then
160           ReturnValue = "C+c-E-e+"

170       ElseIf Value = 7 Then
180           ReturnValue = "C+c+E-e+"

190       ElseIf Value = 8 Then
200           ReturnValue = "C-c+E-e+"

210       Else
220           ReturnValue = ""
230       End If

240   Case "2"    'position 2
250       Antigen1 = "K"
260       Antigen2 = "k"
270   Case "3"
280       Antigen1 = "Cw"
290       Antigen2 = "Mia"
300   Case "4"
310       Antigen1 = "M"
320       Antigen2 = "N"
330   Case "5"
340       Antigen1 = "S"
350       Antigen2 = "s"
360   Case "6"
370       Antigen1 = "U"
380       Antigen2 = "P1"
390   Case "7"
400       Antigen1 = "Lua"
410       Antigen2 = "Kpa"
420   Case "8"
430       Antigen1 = "Lea"
440       Antigen2 = "Leb"
450   Case "9"
460       Antigen1 = "Fya"
470       Antigen2 = "Fyb"
480   Case "10"
490       Antigen1 = "Jka"
500       Antigen2 = "Jkd"
510   Case "11"
520       Antigen1 = "Doa"
530       Antigen2 = "Dod"
540   Case "12"
550       Antigen1 = "Ina"
560       Antigen2 = "Cod"
570   Case "13"
580       Antigen1 = "DiA"
590       Antigen2 = "VS/V"
600   Case "14"
610       Antigen1 = "Jsa"
620       Antigen2 = "C"
630   Case "15"
640       Antigen1 = "c"
650       Antigen2 = "E"
660   Case "16"
670       Antigen1 = "e"
680       Antigen2 = "CMV"


690   End Select
700   If Val(Position) > 1 Then
710       Select Case Value
          Case "0"
720           Antigen1 = ""
730           Antigen2 = ""
740       Case "1"
750           Antigen1 = ""
760           Antigen2 = Antigen2 & "-"
770       Case "2"
780           Antigen1 = ""
790           Antigen2 = Antigen2 & "+"
800       Case "3"
810           Antigen1 = Antigen1 & "-"
820           Antigen2 = ""
830       Case "4"
840           Antigen1 = Antigen1 & "-"
850           Antigen2 = Antigen2 & "-"
860       Case "5"
870           Antigen1 = Antigen1 & "-"
880           Antigen2 = Antigen2 & "+"
890       Case "6"
900           Antigen1 = Antigen1 & "+"
910           Antigen2 = ""
920       Case "7"
930           Antigen1 = Antigen1 & "+"
940           Antigen2 = Antigen2 & "-"
950       Case "8"
960           Antigen1 = Antigen1 & "+"
970           Antigen2 = Antigen2 & "+"
980       Case "9"
990           Antigen1 = ""
1000          Antigen2 = ""

1010      End Select
1020      ReturnValue = Antigen1 & Antigen2
1030  End If
1040  RBCRT010Interpretation = ReturnValue

1050  Exit Function

RBCRT010Interpretation_Error:

 Dim strES As String
 Dim intEL As Integer

1060   intEL = Erl
1070   strES = Err.Description
1080   LogError "BBankShared", "RBCRT010Interpretation", intEL, strES

End Function

'[RT012] Finnish
Public Function RBCRT012(ByVal Value As String) As String


      Dim ReturnValue As String

10    On Error GoTo RBCRT012_Error

20    Select Case Value
      Case "00"
30        ReturnValue = "Information Elsewhere"

40    Case "01"
50        ReturnValue = "Ena"

60    Case "02"
70        ReturnValue = "‘N’"

80    Case "03"
90        ReturnValue = "Vw"

100   Case "04"
110       ReturnValue = "Mur* "

120   Case "05"
130       ReturnValue = "Hut"

140   Case "06"
150       ReturnValue = "Hil"

160   Case "07"
170       ReturnValue = "P"

180   Case "08"
190       ReturnValue = "PP1Pk"

200   Case "09"
210       ReturnValue = "hrS"

220   Case "10"
230       ReturnValue = "hrB"

240   Case "11"
250       ReturnValue = "f"

260   Case "12"
270       ReturnValue = "Ce"

280   Case "13"
290       ReturnValue = "G"

300   Case "14"
310       ReturnValue = "Hro"

320   Case "15"
330       ReturnValue = "CE"

340   Case "16"
350       ReturnValue = "Ce"

360   Case "17"
370       ReturnValue = "Cx"

380   Case "18"
390       ReturnValue = "Ew"

400   Case "19"
410       ReturnValue = "Dw"

420   Case "20"
430       ReturnValue = "hrH"

440   Case "21"
450       ReturnValue = "Goa"

460   Case "22"
470       ReturnValue = "Rh32"

480   Case "23"
490       ReturnValue = "Rh33"

500   Case "24"
510       ReturnValue = "Tar"

520   Case "25"
530       ReturnValue = "Kpb"

540   Case "26"
550       ReturnValue = "Kpc"

560   Case "27"
570       ReturnValue = "Jsb"

580   Case "28"
590       ReturnValue = "Ula"

600   Case "29"
610       ReturnValue = "K11"

620   Case "30"
630       ReturnValue = "K12"

640   Case "31"
650       ReturnValue = "K13"

660   Case "32"
670       ReturnValue = "K14"

680   Case "33"
690       ReturnValue = "K17"

700   Case "34"
710       ReturnValue = "K18"

720   Case "35"
730       ReturnValue = "K19"

740   Case "36"
750       ReturnValue = "K22"

760   Case "37"
770       ReturnValue = "K23"

780   Case "38"
790       ReturnValue = "K24"

800   Case "39"
810       ReturnValue = "Lub"

820   Case "40"
830       ReturnValue = "Lu3"

840   Case "41"
850       ReturnValue = "Lu4"

860   Case "42"
870       ReturnValue = "Lu5"

880   Case "43"
890       ReturnValue = "Lu6"

900   Case "44"
910       ReturnValue = "Lu7"

920   Case "45"
930       ReturnValue = "Lu8"

940   Case "46"
950       ReturnValue = "Lu11"

960   Case "47"
970       ReturnValue = "Lu12"

980   Case "48"
990       ReturnValue = "Lu13"

1000  Case "49"
1010      ReturnValue = "Lu20"

1020  Case "50"
1030      ReturnValue = "Aua"

1040  Case "51"
1050      ReturnValue = "Aub"

1060  Case "52"
1070      ReturnValue = "Fy4"

1080  Case "53"
1090      ReturnValue = "Fy5"

1100  Case "54"
1110      ReturnValue = "Fy6"

1120  Case "55"
1130      ReturnValue = ""    'removed

1140  Case "56"
1150      ReturnValue = "Sda"

1160  Case "57"
1170      ReturnValue = "Wrb"

1180  Case "58"
1190      ReturnValue = "Ytb"

1200  Case "59"
1210      ReturnValue = "Xga"

1220  Case "60"
1230      ReturnValue = "Sc1"

1240  Case "61"
1250      ReturnValue = "Sc2"

1260  Case "62"
1270      ReturnValue = "Sc3"

1280  Case "63"
1290      ReturnValue = "Joa"

1300  Case "64"
1310      ReturnValue = "Dob"

1320  Case "65"
1330      ReturnValue = "Hy"

1340  Case "66"
1350      ReturnValue = "Gya"

1360  Case "67"
1370      ReturnValue = "Co3"

1380  Case "68"
1390      ReturnValue = "LWa"

1400  Case "69"
1410      ReturnValue = "LWb"

1420  Case "70"
1430      ReturnValue = "Kx"

1440  Case "71"
1450      ReturnValue = "Ge2"

1460  Case "72"
1470      ReturnValue = "Ge3"

1480  Case "73"
1490      ReturnValue = "Wb"

1500  Case "74"
1510      ReturnValue = "Lsa"

1520  Case "75"
1530      ReturnValue = "Ana"

1540  Case "76"
1550      ReturnValue = "Dha"

1560  Case "77"
1570      ReturnValue = "Cra"

1580  Case "78"
1590      ReturnValue = "IFC"

1600  Case "79"
1610      ReturnValue = "Kna"

1620  Case "80"
1630      ReturnValue = "Inb"

1640  Case "81"
1650      ReturnValue = "Csa"

1660  Case "82"
1670      ReturnValue = "I"

1680  Case "83"
1690      ReturnValue = "Era"

1700  Case "84"
1710      ReturnValue = "Vel"

1720  Case "85"
1730      ReturnValue = "Lan"

1740  Case "86"
1750      ReturnValue = "Ata"

1760  Case "87"
1770      ReturnValue = "Jra"

1780  Case "88"
1790      ReturnValue = "Oka"

1800  Case "89"
1810      ReturnValue = "Wra"

1820  Case "90"
1830      ReturnValue = " "

1840  Case "91"
1850      ReturnValue = " "

1860  Case "92"
1870      ReturnValue = " "

1880  Case "93"
1890      ReturnValue = " "

1900  Case "94"
1910      ReturnValue = " "

1920  Case "95"
1930      ReturnValue = " "

1940  Case "96"
1950      ReturnValue = " "

1960  Case "97"
1970      ReturnValue = " "

1980  Case "98"
1990      ReturnValue = "IgA deficient"

2000  Case "99"
2010      ReturnValue = "No Information Provided "
2020  Case Else
2030      ReturnValue = ""

2040  End Select
2050  RBCRT012 = ReturnValue

2060  Exit Function

RBCRT012_Error:

 Dim strES As String
 Dim intEL As Integer

2070   intEL = Erl
2080   strES = Err.Description
2090   LogError "BBankShared", "RBCRT012", intEL, strES
End Function



Public Function RT014Interpretation(Position As Integer, Value As String) As String
          Dim Test1Name As String
          Dim Test2Name As String
          Dim Test1Value As String
          Dim Test2Value As String

10    On Error GoTo RT014Interpretation_Error

20    Select Case Value
          Case "1"
30      Test1Value = pubStrRT014Interpretation_Value1r1 ' "" 'nt - not tested
40      Test2Value = pubStrRT014Interpretation_Value1r2 '"Neg"
50    Case "2"
60      Test1Value = pubStrRT014Interpretation_Value2r1 'nt - not tested
70      Test2Value = pubStrRT014Interpretation_Value2r2 '"Pos"
80    Case "3"
90      Test1Value = pubStrRT014Interpretation_Value3r1 '"Neg"
100     Test2Value = pubStrRT014Interpretation_Value3r2 '"" 'nt - not tested
110   Case "4"
120     Test1Value = pubStrRT014Interpretation_Value4r1 '"Neg"
130     Test2Value = pubStrRT014Interpretation_Value4r2 '"Neg"
140   Case "5"
150     Test1Value = pubStrRT014Interpretation_Value5r1 '"Neg"
160     Test2Value = pubStrRT014Interpretation_Value5r2 '"Pos"
170   Case "6"
180     Test1Value = pubStrRT014Interpretation_Value6r1 '"Pos"
190     Test2Value = pubStrRT014Interpretation_Value6r2 '"" 'nt - not tested
200   Case "7"
210     Test1Value = pubStrRT014Interpretation_Value7r1 '"Pos"
220     Test2Value = pubStrRT014Interpretation_Value7r2 '"Neg"
230   Case "8"
240     Test1Value = pubStrRT014Interpretation_Value8r1 '"Pos"
250     Test2Value = pubStrRT014Interpretation_Value8r2 '"Pos"
260   Case Else 'including case zero
270     Test1Value = ""
280     Test2Value = ""
290   End Select

300   Select Case Position
          Case "9"
310     Test1Name = pubStrRT014Interpretation_Pos9r1 '"HPA-1a"
320     Test2Name = pubStrRT014Interpretation_Pos9r2 '"HPA-1b"

330   Case "10"
340     Test1Name = pubStrRT014Interpretation_Pos10r1 '"HPA-2a"
350     Test2Name = pubStrRT014Interpretation_Pos10r2 '"HPA-2b"
360   Case "11"
370     Test1Name = pubStrRT014Interpretation_Pos11r1 '"HPA-3a"
380     Test2Name = pubStrRT014Interpretation_Pos11r2 '"HPA-3b"
390   Case "12"
400     Test1Name = pubStrRT014Interpretation_Pos12r1 '"HPA-4a"
410     Test2Name = pubStrRT014Interpretation_Pos12r2 ' "HPA-4b"
420   Case "13"
430     Test1Name = pubStrRT014Interpretation_Pos13r1 '"HPA-5a"
440     Test2Name = pubStrRT014Interpretation_Pos13r2 '"HPA-5b"
450   Case "14"
460     Test1Name = pubStrRT014Interpretation_Pos14r1 '"HPA-15a"
470     Test2Name = pubStrRT014Interpretation_Pos14r2 '"HPA-6bw"
480   Case "15"
490     Test1Name = pubStrRT014Interpretation_Pos15r1 '"HPA-15b"
500     Test2Name = pubStrRT014Interpretation_Pos15r2 '"HPA-7bw"
510   Case "16"
520     Test1Name = pubStrRT014Interpretation_Pos16r1 '"IgA"
530     Test2Name = pubStrRT014Interpretation_Pos16r2 '"CMV"
540   Case Else
550     Test1Name = ""
560     Test2Name = ""


570   End Select

580   If Trim(Test1Value) = "" Then Test1Name = ""
590   If Trim(Test2Value) = "" Then Test2Name = ""


600   RT014Interpretation = Trim(Test1Name & " " & Test1Value & " " & Test2Name & " " & Test2Value)


610   Exit Function

RT014Interpretation_Error:

 Dim strES As String
 Dim intEL As Integer

620    intEL = Erl
630    strES = Err.Description
640    LogError "BBankShared", "RT014Interpretation", intEL, strES

End Function


Public Sub GetPatientForNameSurName(ByRef strSurName As String, ByRef strForName As String, ByVal strLabNumber As String)
    Dim tb As Recordset
    Dim sql As String

10    On Error GoTo GetPatientForNameSurName_Error

20    If Len(strLabNumber) > 0 Then
30      sql = "Select PatSurName, PatForeName from PatientDetails where LabNumber = '" & strLabNumber & "'"
40      Set tb = New Recordset
50      RecOpenServerBB 0, tb, sql
60      If Not tb.EOF Then
70          strSurName = Trim$(tb!PatSurName & "")
80          strForName = Trim$(tb!PatForeName & "")
90      Else
100         strSurName = ""
110         strForName = ""
120     End If
130      tb.Close
140   End If

150   Exit Sub

GetPatientForNameSurName_Error:

 Dim strES As String
 Dim intEL As Integer

160    intEL = Erl
170    strES = Err.Description
180    LogError "BBankShared", "GetPatientForNameSurName", intEL, strES, sql
    
End Sub

Public Function getPatientForeSurNamefromIPMS(ByVal strChart As String, ByRef strForeName As String, ByRef strSurName As String) As Boolean
    Dim sql As String
    Dim rsRec As Recordset

10    On Error GoTo getPatientForeSurNamefromIPMS_Error

20    getPatientForeSurNamefromIPMS = False

30    strForeName = ""
40    strSurName = ""

50    If Trim$(strChart) <> "" Then

60      If UCase(Left(strChart, 1)) = "M" Then 'Mullingar Chart
70        strChart = Mid(strChart, 2)
80      End If
      
90      sql = "Select PatForeName, PatSurName from PatientIFs where chart = '" & strChart & "' "
100     Set rsRec = New Recordset
110     RecOpenServer 0, rsRec, sql
120     If Not rsRec.EOF Then
130         strForeName = Trim$(rsRec!PatForeName & "")
140         strSurName = Trim$(rsRec!PatSurName & "")
150         getPatientForeSurNamefromIPMS = True
160     End If
170   End If

180   Exit Function

getPatientForeSurNamefromIPMS_Error:

 Dim strES As String
 Dim intEL As Integer

190    intEL = Erl
200    strES = Err.Description
210    LogError "BBankShared", "getPatientForeSurNamefromIPMS", intEL, strES, sql

End Function





