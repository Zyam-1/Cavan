Attribute VB_Name = "modDbDesign"
Option Explicit




'---------------------------------------------------------------------------------------
' Procedure : CheckHaePanelsInDb
' Author    : Masood
' Date      : 04/Nov/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckHaePanelsInDb()

      Dim sql As String


60260   On Error GoTo CheckHaePanelsInDb_Error


60270 If Not IsTableInDatabase("HaePanels") Then
60280     sql = " CREATE TABLE [dbo].[HaePanels]( " & _
          " [PanelName] [nvarchar](50) NULL, " & _
          " [Content] [nvarchar](50) NULL, " & _
          " [BarCode] [nvarchar](20) NULL, " & _
          " [PanelType] [nvarchar](2) NULL, " & _
          " [Hospital] [char](10) NULL, " & _
          " [ListOrder] [int] NULL " & _
          " ) ON [PRIMARY]"

          
60290     Cnxn(0).Execute sql
60300 End If


       
60310 Exit Sub

       
CheckHaePanelsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

60320 intEL = Erl
60330 strES = Err.Description
60340 LogError "modDbDesign", "CheckHaePanelsInDb", intEL, strES, sql
          
End Sub



Public Sub CheckPatientNotepadInDb()

      Dim sql As String

60350 On Error GoTo CheckPatientNotepadInDb_Error

60360 If Not IsTableInDatabase("PatientNotePad") Then
60370     sql = "CREATE TABLE [dbo].[PatientNotePad] " & _
                "([SampleID] [numeric](18, 0) NOT NULL, " & _
                "[DateTimeofRecord] [datetime] NOT NULL, " & _
                "[Comment] [nvarchar](4000) NOT NULL, " & _
                "[UserName] [nvarchar](20) NOT NULL, " & _
                "[Descipline] [nvarchar](20), " & _
                "[LabNo] [numeric](18, 0) )" & _
                "ON [PRIMARY]"
60380     Cnxn(0).Execute sql
60390 End If

60400 Exit Sub

CheckPatientNotepadInDb_Error:

       Dim strES As String
       Dim intEL As Integer

60410  intEL = Erl
60420  strES = Err.Description
60430  LogError "modDbDesign", "CheckPatientNotepadInDb", intEL, strES, sql
          
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckDemographicsUniLabNoInDb
' Author    : Masood
' Date      : 02/Jun/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckDemographicsUniLabNoInDb()
      Dim sql As String


60440   On Error GoTo CheckDemographicsUniLabNoInDb_Error


60450 If Not IsTableInDatabase("DemographicsUniLabNo") Then
60460     sql = "CREATE TABLE [dbo].[DemographicsUniLabNo] " & _
                "([SampleID] [nvarchar](20) NOT NULL, " & _
                "[DateTimeOfRecord] [datetime] NOT NULL, " & _
                "[User] [nvarchar](200) NOT NULL, " & _
                "[PatName] [nvarchar](200) NOT NULL, " & _
                "[DoB] [datetime] NOT NULL, " & _
                "[Sex] [nvarchar](10) NOT NULL, " & _
                "[Chart] [nvarchar](20), " & _
                "[LabNo] [numeric](18, 0) )" & _
                "ON [PRIMARY]"
60470     Cnxn(0).Execute sql
60480 End If

       
60490 Exit Sub

       
CheckDemographicsUniLabNoInDb_Error:

      Dim strES As String
      Dim intEL As Integer

60500 intEL = Erl
60510 strES = Err.Description
60520 LogError "modDbDesign", "CheckDemographicsUniLabNoInDb", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckGpordersProfileInDb
' Author    : Masood
' Date      : 01/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckGpordersProfileInDb()

60530     On Error GoTo CheckGpordersProfileInDb_Error
          Dim sql As String

60540     If Not IsTableInDatabase("GpordersProfile") Then
60550         sql = "CREATE TABLE GpordersProfile(" & _
                    "GPTestCode nvarchar(200) NULL, " & _
                    "GPTestName nvarchar(200) NULL, " & _
                    "Department nvarchar(200) NULL, " & _
                    "NetAcquirePanel nvarchar(200) NULL, " & _
                    "Counter int NULL, " & _
                    "Panel bit NULL " & _
                    ")"
60560         Cnxn(0).Execute sql
60570     End If


60580     Exit Sub


CheckGpordersProfileInDb_Error:

          Dim strES As String
          Dim intEL As Integer

60590     intEL = Erl
60600     strES = Err.Description
60610     LogError "modDbDesign", "CheckGpordersProfileInDb", intEL, strES, sql
End Sub



Public Sub CheckScanViewLogInDb()

      Dim sql As String

60620 On Error GoTo CheckScanViewLogInDb_Error

60630 If Not IsTableInDatabase("ScanViewLog") Then
60640     sql = "CREATE TABLE [dbo].[ScanViewLog]( " & _
                "[SampleID] [nvarchar](50) NULL, " & _
                "[ScanName] [nvarchar](50) NULL, " & _
                "[Department] [nvarchar](50) NULL, " & _
                "[Viewed] [int] NULL, " & _
                "[Username] [nvarchar](50) NULL, " & _
                "[DateTimeOfRecord] [datetime] NULL, " & _
                "[Counter] [int] IDENTITY(1,1) NOT NULL " & _
                ") ON [PRIMARY] "
60650     Cnxn(0).Execute sql
      '
60660 End If

60670 Exit Sub

CheckScanViewLogInDb_Error:

      Dim strES As String
      Dim intEL As Integer

60680 intEL = Erl
60690 strES = Err.Description
60700 LogError "modDbDesign", "CheckScanViewLogInDb", intEL, strES, sql

End Sub


Public Sub CheckAndUpdateLockStatus()

      Dim sql As String

60710 sql = "IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[LockStatus]') " & _
            "               AND OBJECTPROPERTY(id, N'IsUserTable') = 1) " & _
            "SELECT SampleID, " & _
            "cast(v as int) & 1 as Demo , " & _
            "cast(v as int) & 2 as Microscopy, " & _
            "cast(v as int) & 4 as Ident, " & _
            "cast(v as int) & 8 as Faeces, " & _
            "cast(v as int) & 16 as CandS , " & _
            "cast(v as int) & 32 as FOB, " & _
            "cast(v as int) & 64 as RotaAdeno, " & _
            "cast(v as int) & 128  as RedSub, " & _
            "cast(v as int) & 256 as RSV , " & _
            "cast(v as int) & 512 as CSF, " & _
            "cast(v as int) & 1024 as CDiff, " & _
            "cast(v as int) & 2048 as OP, " & _
            "cast(v as int) & 4096 as Identification " & _
            "INTO LockStatus FROM PrintValid " & _
            "IF NOT EXISTS (SELECT name FROM sysindexes WHERE name = 'idx_LS_SampleID') " & _
            "  CREATE CLUSTERED INDEX [idx_LS_SampleID] ON [dbo].[LockStatus] ([SampleID])"

60720 Cnxn(0).Execute sql

End Sub

Public Sub CheckAndUpdatePrintedStatus()

      Dim sql As String

60730 sql = "IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[PrintedStatus]') " & _
            "               AND OBJECTPROPERTY(id, N'IsUserTable') = 1) " & _
            "SELECT SampleID, " & _
            "cast(P as int) & 1 as Demo , " & _
            "cast(P as int) & 2 as Microscopy, " & _
            "cast(P as int) & 4 as Ident, " & _
            "cast(P as int) & 8 as Faeces, " & _
            "cast(P as int) & 16 as CandS , " & _
            "cast(P as int) & 32 as FOB, " & _
            "cast(P as int) & 64 as RotaAdeno, " & _
            "cast(P as int) & 128  as RedSub, " & _
            "cast(P as int) & 256 as RSV , " & _
            "cast(P as int) & 512 as CSF, " & _
            "cast(P as int) & 1024 as CDiff, " & _
            "cast(P as int) & 2048 as OP, " & _
            "cast(P as int) & 4096 as Identification " & _
            "INTO PrintedStatus FROM PrintValid " & _
            "IF NOT EXISTS (SELECT name FROM sysindexes WHERE name = 'idx_PS_SampleID') " & _
            "  CREATE CLUSTERED INDEX [idx_PS_SampleID] ON [dbo].[PrintedStatus] ([SampleID])"

60740 Cnxn(0).Execute sql

End Sub

Public Sub CheckHaemControlsInDb()

      Dim sql As String

60750 On Error GoTo CheckHaemControlsInDb_Error

60760 If IsTableInDatabase("HaemControls") = False Then 'There is no table  in database
60770   sql = "CREATE TABLE HaemControls " & _
              " ( Rundate datetime, " & _
              "   RunDateTime datetime, " & _
              "   SampleID numeric, " & _
              "   RBC nvarchar (6), " & _
              "   WBC nvarchar (6), " & _
              "   Hgb nvarchar (6), " & _
              "   MCV nvarchar (6), " & _
              "   Hct nvarchar (6), " & _
              "   MCH nvarchar (6), " & _
              "   MCHC nvarchar (6), " & _
              "   RDWCV nvarchar (6), " & _
              "   Plt nvarchar (6), " & _
              "   MPV nvarchar (6), " & _
              "   Plcr nvarchar (6), " & _
              "   PDW nvarchar (6), "
60780   sql = sql & " LymA nvarchar (6), " & _
              "   LymP nvarchar (6), " & _
              "   MonoA nvarchar (6), " & _
              "   MonoP nvarchar (6), " & _
              "   NeutA nvarchar (6), " & _
              "   NeutP nvarchar (6) )"
60790   Cnxn(0).Execute sql
60800 End If

60810 Exit Sub

CheckHaemControlsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

60820 intEL = Erl
60830 strES = Err.Description
60840 LogError "modDbDesign", "CheckHaemControlsInDb", intEL, strES, sql

End Sub

Public Sub CheckCodeTranslationInDb()

      Dim sql As String


60850 On Error GoTo CheckCodeTranslationInDb_Error

60860 If IsTableInDatabase("CodeTranslation") = False Then
          
60870     sql = "CREATE TABLE CodeTranslation " & _
                  "(LisCode nvarchar (50) ," & _
                  "HostCode nvarchar (50) ," & _
                  "Discipline nvarchar (50) ," & _
                  "AnalyserID nvarchar (50) ," & _
                  "CodePrintName nvarchar (200)" & _
                ") ON [PRIMARY]"
60880 Cnxn(0).Execute sql

60890 End If
60900 Exit Sub

CheckCodeTranslationInDb_Error:

      Dim strES As String
      Dim intEL As Integer

60910 intEL = Erl
60920 strES = Err.Description
60930 LogError "modDbDesign", "CheckCodeTranslationInDb", intEL, strES, sql


End Sub

Public Sub CheckHaemControlDefinitionsInDb()

      Dim sql As String

60940 On Error GoTo CheckHaemControlDefinitionsInDb_Error

60950 If IsTableInDatabase("HaemControlDefinitions") = False Then 'There is no table  in database

60960   Cnxn(0).Execute sql
60970 End If

60980 Exit Sub

CheckHaemControlDefinitionsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

60990 intEL = Erl
61000 strES = Err.Description
61010 LogError "modDbDesign", "CheckHaemControlDefinitionsInDb", intEL, strES, sql

End Sub


Public Function EnsureColumnExists(ByVal TableName As String, _
                                   ByVal ColumnName As String, _
                                   ByVal Definition As String) _
                                   As Boolean

      'Return 1 if column created
      '       0 if column already exists

      Dim sql As String
      Dim tb As Recordset

61020 On Error GoTo EnsureColumnExists_Error

61030 sql = "IF NOT EXISTS " & _
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

61040 Set tb = Cnxn(0).Execute(sql)

61050 EnsureColumnExists = tb!RetVal

61060 Exit Function

EnsureColumnExists_Error:

      Dim strES As String
      Dim intEL As Integer

61070 intEL = Erl
61080 strES = Err.Description
61090 LogError "modDbDesign", "EnsureColumnExists", intEL, strES, sql


End Function

Public Function EnsureListEntryExists(ByVal Code As String, ByVal ListText As String, ListType As String, Optional ByVal InUse As Boolean = True) As Boolean

      Dim sql As String
      Dim tb As Recordset


61100 On Error GoTo EnsureListEntryExists_Error

61110 sql = "IF NOT EXISTS " & _
            "    (SELECT * FROM Lists WHERE " & _
            "    Code = '" & Code & "' " & _
            "    AND Text = '" & ListText & "' " & _
            "    AND ListType = '" & ListType & "') " & _
            "  BEGIN " & _
            "    INSERT INTO Lists (Code, Text, ListType, InUse, ListOrder) " & _
            "    VALUES ('" & Code & "', '" & ListText & "', '" & ListType & "', '" & IIf(InUse, 1, 0) & "', 999) " & _
            "  END " & _
            "ELSE " & _
            "  SELECT 0 AS RetVal"

61120 Set tb = Cnxn(0).Execute(sql)
61130 EnsureListEntryExists = True

61140 Exit Function

EnsureListEntryExists_Error:

      Dim strES As String
      Dim intEL As Integer

61150 intEL = Erl
61160 strES = Err.Description
61170 LogError "modDbDesign", "EnsureListEntryExists", intEL, strES, sql

End Function


Public Function EnsureOptionExists(ByVal Description As String, _
                                   ByVal Contents As String) _
                                   As Boolean

      'Return 1 if column created
      '       0 if column already exists

      Dim sql As String
      Dim tb As Recordset



61180 On Error GoTo EnsureOptionExists_Error

61190 sql = "IF NOT EXISTS " & _
            "    (SELECT * FROM Options WHERE " & _
            "    Description = '" & Description & "') " & _
            "  BEGIN " & _
            "    INSERT INTO Options (Description, Contents, UserName) " & _
            "    VALUES ('" & Description & "', '" & Contents & "', 'System') " & _
            "  END " & _
            "ELSE " & _
            "  SELECT 0 AS RetVal"

61200 Set tb = Cnxn(0).Execute(sql)

61210 EnsureOptionExists = True




61220 Exit Function

EnsureOptionExists_Error:

      Dim strES As String
      Dim intEL As Integer

61230 intEL = Erl
61240 strES = Err.Description
61250 LogError "modDbDesign", "EnsureOptionExists", intEL, strES, sql


End Function

Public Function IsTableInDatabase(ByVal TableName As String) As Boolean

      Dim tbExists As Recordset
      Dim sql As String
      Dim RetVal As Boolean

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist
      'if it has a record then the table does exist.

61260 On Error GoTo IsTableInDatabase_Error

61270 sql = "SELECT name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = '" & TableName & "'"

61280 Set tbExists = Cnxn(0).Execute(sql)
      '

61290 RetVal = True

61300 If tbExists.EOF Then 'There is no table <TableName> in database
61310   RetVal = False
61320 End If
61330 IsTableInDatabase = RetVal

61340 Exit Function

IsTableInDatabase_Error:

      Dim strES As String
      Dim intEL As Integer

61350 intEL = Erl
61360 strES = Err.Description
61370 LogError "modDbDesign", "IsTableInDatabase", intEL, strES, sql

        
End Function


Public Sub CheckFaxLogInDb()

      Dim sql As String

61380 On Error GoTo CheckFaxLogInDb_Error

61390 If IsTableInDatabase("FaxLog") = False Then 'There is no table  in database
61400   sql = "CREATE TABLE FaxLog " & _
              "( SampleID  numeric(9), " & _
              "  DateTime datetime, " & _
              "  FaxedTo nvarchar(50), " & _
              "  FaxedBy nvarchar(50), " & _
              "  Comment nvarchar(50), " & _
              "  Discipline nvarchar(10) )"
61410   Cnxn(0).Execute sql
61420 End If

61430 Exit Sub

CheckFaxLogInDb_Error:

      Dim strES As String
      Dim intEL As Integer

61440 intEL = Erl
61450 strES = Err.Description
61460 LogError "modDbDesign", "CheckFaxLogInDb", intEL, strES, sql

End Sub

Public Sub CheckPrintInhibitInDb()

      Dim sql As String

61470 On Error GoTo CheckPrintInhibitInDb_Error

61480 If IsTableInDatabase("PrintInhibit") = False Then 'There is no table  in database
61490   sql = "CREATE TABLE PrintInhibit " & _
              "( [SampleID] [numeric](18, 0) NULL, " & _
              "  [Discipline] [nvarchar](50) NULL, " & _
              "  [Parameter] [nvarchar](50) NULL )"
61500   Cnxn(0).Execute sql
61510 End If

61520 Exit Sub

CheckPrintInhibitInDb_Error:

      Dim strES As String
      Dim intEL As Integer

61530 intEL = Erl
61540 strES = Err.Description
61550 LogError "modDbDesign", "CheckPrintInhibitInDb", intEL, strES, sql

End Sub

Public Sub CheckAnalyserMessagesInDb()

      Dim sql As String

61560 On Error GoTo CheckAnalyserMessagesInDb_Error

61570 If IsTableInDatabase("AnalyserMessages") = False Then 'There is no table  in database
61580   sql = "CREATE TABLE AnalyserMessages " & _
              "( Analyser nvarchar(50) NOT NULL, " & _
              "  SampleID  numeric(9) NOT NULL, " & _
              "  Message nvarchar(100) NOT NULL, " & _
              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
61590   Cnxn(0).Execute sql
61600 End If

61610 Exit Sub

CheckAnalyserMessagesInDb_Error:

      Dim strES As String
      Dim intEL As Integer

61620 intEL = Erl
61630 strES = Err.Description
61640 LogError "modDbDesign", "CheckAnalyserMessagesInDb", intEL, strES, sql

End Sub

Public Sub CheckDemogValidationInDb()

      Dim sql As String

61650 If IsTableInDatabase("DemogValidation") = False Then 'There is no table  in database
61660   sql = "CREATE TABLE DemogValidation " & _
              "( SampleID nvarchar(50) NOT NULL, " & _
              "  EnteredBy nvarchar(50) NOT NULL, " & _
              "  ValidatedBy nvarchar(50) NOT NULL, " & _
              "  EnteredDateTime datetime NOT NULL, " & _
              "  ValidatedDateTime datetime NOT NULL DEFAULT getdate() )"
61670   Cnxn(0).Execute sql
61680 End If
End Sub


Public Sub CheckPOCTPatientLiveInDb()

      Dim sql As String

61690 On Error GoTo CheckPOCTPatientLiveInDb_Error

61700 If IsTableInDatabase("POCTPatientLive") = False Then 'There is no table  in database
61710   sql = "CREATE TABLE [dbo].[POCTPatientLive]( " & _
              "  [Chart] [nvarchar](50) NULL, " & _
              "  [ForeName] [nvarchar](50) NULL, " & _
              "  [Surname] [nvarchar](50) NULL, " & _
              "  [AlternateID] [nvarchar](50) NULL, " & _
              "  [AccountNo] [nvarchar](50) NULL, " & _
              "  [Visit] [nvarchar](50) NULL, " & _
              "  [Sex] [nvarchar](50) NULL, " & _
              "  [DoB] [smalldatetime] NULL, " & _
              "  [Location] [nvarchar](50) NULL, " & _
              "  [Doctor] [nvarchar](50) NULL, " & _
              "  [DateTimeOfRecord] datetime NOT NULL DEFAULT getdate() " & _
              ")"
61720   Cnxn(0).Execute sql
61730 End If

61740 Exit Sub

CheckPOCTPatientLiveInDb_Error:

      Dim strES As String
      Dim intEL As Integer

61750 intEL = Erl
61760 strES = Err.Description
61770 LogError "modDbDesign", "CheckPOCTPatientLiveInDb", intEL, strES, sql

End Sub

Public Sub CheckPOCTPatientTempInDb()

      Dim sql As String

61780 On Error GoTo CheckPOCTPatientTempInDb_Error

61790 If IsTableInDatabase("POCTPatientTemp") = False Then 'There is no table  in database
61800   sql = "CREATE TABLE [dbo].[POCTPatientTemp]( " & _
              "  [Chart] [nvarchar](50) NULL, " & _
              "  [ForeName] [nvarchar](50) NULL, " & _
              "  [Surname] [nvarchar](50) NULL, " & _
              "  [AlternateID] [nvarchar](50) NULL, " & _
              "  [AccountNo] [nvarchar](50) NULL, " & _
              "  [Visit] [nvarchar](50) NULL, " & _
              "  [Sex] [nvarchar](50) NULL, " & _
              "  [DoB] [smalldatetime] NULL, " & _
              "  [Location] [nvarchar](50) NULL, " & _
              "  [Doctor] [nvarchar](50) NULL, " & _
              "  [DateTimeOfRecord] datetime NOT NULL DEFAULT getdate() " & _
              ")"
61810   Cnxn(0).Execute sql
61820 End If

61830 Exit Sub

CheckPOCTPatientTempInDb_Error:

      Dim strES As String
      Dim intEL As Integer

61840 intEL = Erl
61850 strES = Err.Description
61860 LogError "modDbDesign", "CheckPOCTPatientTempInDb", intEL, strES, sql

End Sub
Public Sub CheckLIHValuesInDb()

      Dim sql As String

61870 On Error GoTo CheckLIHValuesInDb_Error

61880 If IsTableInDatabase("LIHValues") = False Then 'There is no table  in database
61890   sql = "CREATE TABLE [dbo].[LIHValues]( " & _
              "  [LIH] nvarchar(50) NOT NULL, " & _
              "  [Code] nvarchar(50) NOT NULL, " & _
              "  [CutOff] real NOT NULL DEFAULT 0, " & _
              "  [NoPrintOrWarning] nvarchar(50) NOT NULL, " & _
              "  [UserName] nvarchar(50) NOT NULL, " & _
              "  [DateTimeOfRecord] datetime NOT NULL DEFAULT getdate() " & _
              ")"
61900   Cnxn(0).Execute sql
61910 End If

61920 Exit Sub

CheckLIHValuesInDb_Error:

      Dim strES As String
      Dim intEL As Integer

61930 intEL = Erl
61940 strES = Err.Description
61950 LogError "modDbDesign", "CheckLIHValuesInDb", intEL, strES, sql

End Sub

Public Sub CheckPOCTPatientsInDb()

      Dim sql As String

61960 On Error GoTo CheckPOCTPatientsInDb_Error

61970 If IsTableInDatabase("POCTPatients") = False Then 'There is no table  in database
61980   sql = "CREATE TABLE [dbo].[POCTPatients]( " & _
              "  [Chart] [nvarchar](50) NULL, " & _
              "  [ForeName] [nvarchar](50) NULL, " & _
              "  [Surname] [nvarchar](50) NULL, " & _
              "  [AlternateID] [nvarchar](50) NULL, " & _
              "  [AccountNo] [nvarchar](50) NULL, " & _
              "  [Visit] [nvarchar](50) NULL, " & _
              "  [Sex] [nvarchar](50) NULL, " & _
              "  [DoB] [smalldatetime] NULL, " & _
              "  [Location] [nvarchar](50) NULL, " & _
              "  [Doctor] [nvarchar](50) NULL, " & _
              "  [DateTimeOfRecord] datetime NOT NULL DEFAULT getdate() " & _
              ")"
61990   Cnxn(0).Execute sql
62000 End If

62010 Exit Sub

CheckPOCTPatientsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62020 intEL = Erl
62030 strES = Err.Description
62040 LogError "modDbDesign", "CheckPOCTPatientsInDb", intEL, strES, sql

End Sub
Public Sub CheckPOCTResultsInDb()

      Dim sql As String

62050 On Error GoTo CheckPOCTResultsInDb_Error

62060 If IsTableInDatabase("POCTResults") = False Then 'There is no table  in database
62070   sql = "CREATE TABLE [dbo].[POCTResults]( " & _
              "  [Parameter] [nvarchar](50) NULL, " & _
              "  [Result] [nvarchar](100) NULL, " & _
              "  [Units] [nvarchar](50) NULL, " & _
              "  [PatientUI] [nvarchar](50) NULL, " & _
              "  [DateTimeOfRecord] [datetime] NOT NULL DEFAULT getdate(), " & _
              "  [Chart] [nvarchar](50) NULL, " & _
              "  [FileName] [nvarchar](100) NULL " & _
              ")"
62080   Cnxn(0).Execute sql
62090 End If

62100 Exit Sub

CheckPOCTResultsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62110 intEL = Erl
62120 strES = Err.Description
62130 LogError "modDbDesign", "CheckPOCTResultsInDb", intEL, strES, sql

End Sub

Public Sub CheckBioDefIndexInDb()

      Dim sql As String

62140 On Error GoTo CheckBioDefIndexInDb_Error

62150 If IsTableInDatabase("BioDefIndex") = False Then 'There is no table  in database
62160   sql = "CREATE TABLE BioDefIndex " & _
              "( DefIndex numeric IDENTITY(0, 1), " & _
              "  NormalLow real NOT NULL, " & _
              "  NormalHigh real NOT NULL, " & _
              "  FlagLow real NOT NULL, " & _
              "  FlagHigh real NOT NULL, " & _
              "  PlausibleLow real NOT NULL, " & _
              "  PlausibleHigh real NOT NULL, " & _
              "  AutoValLow real NOT NULL, " & _
              "  AutoValHigh real NOT NULL)"
62170   Cnxn(0).Execute sql

62180   sql = "INSERT INTO BioDefIndex " & _
              "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
              "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
              "VALUES (0, 9999, 0, 9999, 0, 9999, 0, 9999)"
62190   Cnxn(0).Execute sql
        
62200 End If

62210 Exit Sub

CheckBioDefIndexInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62220 intEL = Erl
62230 strES = Err.Description
62240 LogError "modDbDesign", "CheckBioDefIndexInDb", intEL, strES, sql

End Sub

Public Sub CheckAutoCommentsInDb()

      Dim sql As String

62250 On Error GoTo CheckAutoCommentsInDb_Error

62260 If IsTableInDatabase("AutoComments") = False Then 'There is no table  in database
62270   sql = "CREATE TABLE AutoComments " & _
              "( Discipline nvarchar(50) NOT NULL, " & _
              "  Parameter nvarchar(50) NOT NULL, " & _
              "  Criteria nvarchar(50) NOT NULL, " & _
              "  Value0 nvarchar(50), " & _
              "  Value1 nvarchar(50), " & _
              "  Comment nvarchar(80), " & _
              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate() )"
62280   Cnxn(0).Execute sql
62290 End If

62300 Exit Sub

CheckAutoCommentsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62310 intEL = Erl
62320 strES = Err.Description
62330 LogError "modDbDesign", "CheckAutoCommentsInDb", intEL, strES, sql

End Sub

Public Sub CheckBiomnisRequestInDb()

      Dim sql As String

62340 On Error GoTo CheckBiomnisRequestInDb_Error

62350 If IsTableInDatabase("BiomnisRequests") = False Then
62360     sql = "CREATE TABLE [dbo].[BiomnisRequests]( " & _
                "[SampleID] [nvarchar](50) NOT NULL, " & _
                "[TestCode] [nvarchar](50) NOT NULL, " & _
                "[TestName] [nvarchar](50) NULL, " & _
                "[SampleType] [nvarchar](50) NULL, " & _
                "[SampleDateTime] [datetime] NULL, " & _
                "[Department] [nvarchar](50) NULL, " & _
                "[RequestedBy] [nvarchar](50) NULL, " & _
                "[SendTo] [nvarchar](200) NULL, " & _
                "[Status] [nvarchar](50) NULL, " & _
                "DateTimeOfRecord datetime NOT NULL DEFAULT getdate() " & _
                ") ON [PRIMARY]"
62370     Cnxn(0).Execute sql
62380 End If

62390 Exit Sub

CheckBiomnisRequestInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62400 intEL = Erl
62410 strES = Err.Description
62420 LogError "modDbDesign", "CheckBiomnisRequestInDb", intEL, strES, sql

End Sub

Public Sub CheckMicroAutoCommentAlertInDb()

      Dim sql As String

62430 On Error GoTo CheckMicroAutoCommentAlertInDb_Error

62440 If IsTableInDatabase("MicroAutoCommentAlert") = False Then

62450     sql = "CREATE TABLE [dbo].[MicroAutoCommentAlert]( " & _
                  "[OrganismName] [nvarchar](100) NOT NULL, " & _
                  "[Site] [nvarchar](50) NULL, " & _
                  "[PatientLocation] [nvarchar](100) NULL, " & _
                  "[PatientAgeFrom] [int] NULL, " & _
                  "[PatientAgeTo] [int] NULL, " & _
                  "[DateStart] [smalldatetime] NULL, " & _
                  "[DateEnd] [smalldatetime] NULL, " & _
                  "[Comment] [nvarchar](95) NULL, " & _
                  "[PhoneAlert] [bit] NULL, " & _
                  "[PhoneAlertDateTime] [datetime] NULL, " & _
                  "[ListOrder] [int] NOT NULL, " & _
                  "[Counter] [int] IDENTITY(1,1) NOT NULL " & _
                 ") ON [PRIMARY] "
62460            Cnxn(0).Execute sql
62470 End If

62480 Exit Sub

CheckMicroAutoCommentAlertInDb_Error:

       Dim strES As String
       Dim intEL As Integer

62490  intEL = Erl
62500  strES = Err.Description
62510  LogError "modDbDesign", "CheckMicroAutoCommentAlertInDb", intEL, strES, sql
          
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckConsultantListInDb
' Author    : Masood
' Date      : 28/Apr/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckConsultantListLogInDb()
62520   On Error GoTo CheckConsultantListInDb_Error
      Dim sql As String


62530 If IsTableInDatabase("ConsultantListLog") = False Then
62540     sql = "CREATE TABLE [dbo].[ConsultantListLog]( " & _
                "[SampleID] [nvarchar](50) NOT NULL, " & _
                "[UserName] [nvarchar](50) NOT NULL, " & _
                "[Status] [nvarchar](50) NULL, " & _
                "DateTimeOfRecord datetime NOT NULL DEFAULT getdate() " & _
                ") ON [PRIMARY]"
62550     Cnxn(0).Execute sql
62560 End If

       
62570 Exit Sub

       
CheckConsultantListInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62580 intEL = Erl
62590 strES = Err.Description
62600 LogError "modDbDesign", "CheckConsultantListInDb", intEL, strES, sql
End Sub


Public Sub CheckUserRoleInDb()

      Dim sql As String

62610 On Error GoTo CheckUserRoleInDb_Error

62620 If IsTableInDatabase("UserRole") = False Then
62630     sql = "CREATE TABLE [dbo].[UserRole]( " & _
                "[MemberOf] [nvarchar](50) NULL, " & _
                "[SystemRole] [nvarchar](50) NULL, " & _
                "[Description] [nvarchar](50) NULL, " & _
                "[Enabled] [tinyint] NULL, " & _
                "[Username] [nvarchar](50) NULL, " & _
                "[DateTimeOfRecord] [datetime] NULL, " & _
                "[Counter] [int] IDENTITY(1,1) NOT NULL " & _
                ") ON [PRIMARY] "

62640     Cnxn(0).Execute sql
62650 End If

62660 Exit Sub

CheckUserRoleInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62670 intEL = Erl
62680 strES = Err.Description
62690 LogError "modDbDesign", "CheckUserRoleInDb", intEL, strES, sql

End Sub

Public Sub CheckTempHaemInDb()

      Dim sql As String

62700 On Error GoTo CheckTempHaemInDb_Error

62710 If IsTableInDatabase("TempHaem") = False Then 'There is no table  in database

62720   sql = "SELECT * INTO TempHaem " & _
              "FROM HaemResults WHERE 0 = 1"
        
      '30      sql = "CREATE TABLE TempHaem (" & _
      '              "SampleID numeric, " & _
      '              "AnalysisError nvarchar(1), " & _
      '              "NegPosError nvarchar(1), " & _
      '              "PosDiff nvarchar(1), " & _
      '              "PosMorph  nvarchar(1), " & _
      '              "PosCount  nvarchar(1), " & _
      '              "err_f nvarchar(1), " & _
      '              "err_r nvarchar(1), " & _
      '              "ipmessage nvarchar(6), " & _
      '              "wbc nvarchar(6), " & _
      '              "rbc nvarchar(5), " & _
      '              "hgb nvarchar(5), " & _
      '              "hct nvarchar(5), " & _
      '              "mcv nvarchar(5), " & _
      '              "mch nvarchar(5), " & _
      '              "mchc nvarchar(5), " & _
      '              "plt nvarchar(5), " & _
      '              "lymp  nvarchar(6), " & _
      '              "monop nvarchar(6), " & _
      '              "neutP nvarchar(6), " & _
      '              "eosp  nvarchar(6), " & _
      '              "basp  nvarchar(6), " & _
      '              "lyma  nvarchar(6), " & _
      '              "monoa nvarchar(6), "
      '40      sql = sql & "neuta nvarchar(6), " & _
      '              "eosa  nvarchar(6), " & _
      '              "basa  nvarchar(6), " & _
      '              "rdwcv nvarchar(5), " & _
      '              "rdwSD nvarchar(5), " & _
      '              "pdw nvarchar(5), " & _
      '              "mpv nvarchar(5), " & _
      '              "plcr nvarchar(5), " & _
      '              "valid bit, " & _
      '              "printed tinyint, " & _
      '              "retics nvarchar(4), " & _
      '              "monospot nvarchar(1), " & _
      '              "wbccomment nvarchar(30), " & _
      '              "cesr bit, " & _
      '              "cretics bit, " & _
      '              "cmonospot bit, " & _
      '              "ccoag bit, " & _
      '              "md0 nvarchar(30), " & _
      '              "md1 nvarchar(30), " & _
      '              "md2 nvarchar(30), " & _
      '              "md3 nvarchar(30), " & _
      '              "md4 nvarchar(30), " & _
      '              "md5 nvarchar(30), " & _
      '              "RunDate smalldatetime, " & _
      '              "RunDateTime datetime, "
      '50      sql = sql & "ESR nvarchar(5), "
      '60      sql = sql & "PT  nvarchar(5), " & _
      '              "PTControl nvarchar(5), " & _
      '              "APTT  nvarchar(5), " & _
      '              "APTTControl nvarchar(5), " & _
      '              "INR nvarchar(5), " & _
      '              "FDP nvarchar(10), " & _
      '              "FIB nvarchar(5), " & _
      '              "Operator nvarchar(5), " & _
      '              "FAXed bit, " & _
      '              "Warfarin nvarchar(5), " & _
      '              "DDimers nvarchar(5), " & _
      '              "TransmitTime datetime, " & _
      '              "Pct char(10), " & _
      '              "WIC char(10), " & _
      '              "WOC char(10), " & _
      '              "gWB1 ntext, " & _
      '              "gWB2 ntext, " & _
      '              "gRBC ntext, " & _
      '              "gPlt ntext, " & _
      '              "gWIC ntext, " & _
      '              "LongError numeric, " & _
      '              "cFilm bit, " & _
      '              "RetA nvarchar(5), " & _
      '              "RetP nvarchar(5), " & _
      '              "nrbcA char(5), "
      '70      sql = sql & "nrbcP char(5), " & _
      '              "cMalaria bit, " & _
      '              "Malaria char(8), " & _
      '              "cSickledex bit, " & _
      '              "Sickledex char(8), " & _
      '              "RA nvarchar(1), " & _
      '              "cRA bit, " & _
      '              "Val1 int, " & _
      '              "Val2 int, " & _
      '              "Val3 int, " & _
      '              "Val4 int, " & _
      '              "Val5 int, " & _
      '              "gRBCH ntext, " & _
      '              "gPLTH ntext, " & _
      '              "gPLTF ntext, " & _
      '              "gV ntext, " & _
      '              "gC ntext, " & _
      '              "gS ntext, " & _
      '              "DF1 ntext, " & _
      '              "IRF nvarchar(6), " & _
      '              "Image image, " & _
      '              "mi char(10), " & _
      '              "an char(10), " & _
      '              "ca char(10), "
      '80      sql = sql & "va char(10), " & _
      '              "ho char(10), " & _
      '              "he char(10), " & _
      '              "ls char(10), " & _
      '              "[at] char(10), " & _
      '              "bl char(10), " & _
      '              "pp char(10), " & _
      '              "nl char(10), " & _
      '              "mn char(10), " & _
      '              "wp char(10), " & _
      '              "ch char(10), " & _
      '              "wb char(10), " & _
      '              "hdw char(10), " & _
      '              "LUCP char(10), " & _
      '              "LUCA char(10), " & _
      '              "LI char(10), " & _
      '              "MPXI char(10), " & _
      '              "ANALYSER char(3), " & _
      '              "cAsot bit, "
      '90      sql = sql & "tAsot char(10), " & _
      '              "tRa char(10), " & _
      '              "hyp char(10), " & _
      '              "rbcf char(10), " & _
      '              "rbcg char(10), " & _
      '              "mpo char(10), " & _
      '              "ig char(10), " & _
      '              "lplt char(10), " & _
      '              "pclm char(10), " & _
      '              "ValidateTime datetime, " & _
      '              "Healthlink bit, " & _
      '              "CD3A nvarchar(50), " & _
      '              "CD4A nvarchar(50), " & _
      '              "CD8A nvarchar(50), " & _
      '              "CD3P nvarchar(50), " & _
      '              "CD4P nvarchar(50), " & _
      '              "CD8P nvarchar(50), " & _
      '              "CD48 nvarchar(50) )"
62730   Cnxn(0).Execute sql
62740 End If

62750 Exit Sub

CheckTempHaemInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62760 intEL = Erl
62770 strES = Err.Description
62780 LogError "modDbDesign", "CheckTempHaemInDb", intEL, strES, sql

End Sub

Public Sub CheckSapphireRequestsInDb()

      Dim sql As String

62790 On Error GoTo CheckSapphireRequestsInDb_Error

62800 If IsTableInDatabase("SapphireRequests") = False Then 'There is no table  in database
62810   sql = "CREATE TABLE SapphireRequests " & _
              "( SampleID  numeric(9), " & _
              "  OrderString nvarchar(50), " & _
              "  Programmed int )"
62820   Cnxn(0).Execute sql
62830 End If

62840 Exit Sub

CheckSapphireRequestsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62850 intEL = Erl
62860 strES = Err.Description
62870 LogError "modDbDesign", "CheckSapphireRequestsInDb", intEL, strES, sql

End Sub

Public Sub CheckTestDefinitionsArcInDb()

      Dim sql As String
      Dim n As Integer
      Dim tbName As String

62880 On Error GoTo CheckTestDefinitionsArcInDb_Error

62890 For n = 1 To 3
62900   tbName = Choose(n, "BioTestDefinitionsArc", "ImmTestDefinitionsArc", "EndTestDefinitionsArc")
        
62910   If IsTableInDatabase(tbName) = False Then 'There is no table  in database
62920     sql = "CREATE TABLE " & tbName & " " & _
                "(LongName nvarchar(50), ShortName nvarchar(50), " & _
                " DoDelta bit, DeltaLimit real, PrintPriority smallint, " & _
                " DP smallint, BarCode nvarchar(50), Units nvarchar(50), " & _
                " H bit, S bit, L bit, O bit, G bit, J bit, " & _
                " Category nvarchar(50), Code nvarchar(50), Printable bit, " & _
                " PlausibleLow real, PlausibleHigh real, " & _
                " KnownToAnalyser bit, SampleType nvarchar(50), InUse bit, " & _
                " MaleLow real, MaleHigh real, FemaleLow real, FemaleHigh real, " & _
                " FlagMaleLow real, FlagMaleHigh real, FlagFemaleLow real, FlagFemaleHigh real, " & _
                " LControlLow real, LControlHigh real, NControlLow real, NControlHigh real, HControlLow real, HControlHigh real, " & _
                " AgeFromDays int, AgeToDays int, AutoValLow real, AutoValHigh real, " & _
                " Hospital nvarchar(50), Analyser nvarchar(50), ImmunoCode nvarchar(50), " & _
                " SplitList int, EOD bit, LIH int, " & _
                " ActiveFromDate smalldatetime, ActiveToDate smalldatetime, " & _
                " DateTimeOfArchive datetime, ArchivedBy nvarchar(50) )"
62930     Cnxn(0).Execute sql
62940   End If

62950 Next

62960 Exit Sub

CheckTestDefinitionsArcInDb_Error:

      Dim strES As String
      Dim intEL As Integer

62970 intEL = Erl
62980 strES = Err.Description
62990 LogError "modDbDesign", "CheckTestDefinitionsArcInDb", intEL, strES, sql

End Sub

Public Sub CheckGroupedHospitalsInDb()

      Dim sql As String

63000 On Error GoTo CheckGroupedHospitalsInDb_Error

63010 If IsTableInDatabase("GroupedHospitals") = False Then 'There is no table  in database
63020   sql = "CREATE TABLE GroupedHospitals " & _
              "( [HospName] nvarchar(50), " & _
              "  [Connect] nvarchar(250), " & _
              "  [ConnectBB] nvarchar(250), " & _
              "  [UseInIDE] bit )"
63030   Cnxn(0).Execute sql
63040 End If

63050 Exit Sub

CheckGroupedHospitalsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63060 intEL = Erl
63070 strES = Err.Description
63080 LogError "modDbDesign", "CheckGroupedHospitalsInDb", intEL, strES, sql

End Sub


Public Sub CheckAssociatedIDsInDb()

      Dim sql As String

63090 On Error GoTo CheckAssociatedIDsInDb_Error

63100 If IsTableInDatabase("AssociatedIDs") = False Then 'There is no table  in database
63110   sql = "CREATE TABLE AssociatedIDs " & _
              "( SampleID  numeric(9), " & _
              "  AssID  numeric(9) )"
63120   Cnxn(0).Execute sql
63130 End If

63140 Exit Sub

CheckAssociatedIDsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63150 intEL = Erl
63160 strES = Err.Description
63170 LogError "modDbDesign", "CheckAssociatedIDsInDb", intEL, strES, sql

End Sub

Public Sub CheckPrintValidLogInDb()

      Dim sql As String

63180 On Error GoTo CheckPrintValidLogInDb_Error

63190 If IsTableInDatabase("PrintValidLog") = False Then 'There is no table  in database
63200   sql = "CREATE TABLE PrintValidLog " & _
              "( SampleID numeric(9), " & _
              "  Department nvarchar(1), " & _
              "  Printed tinyint, " & _
              "  Valid tinyint, " & _
              "  PrintedBy nvarchar(50), " & _
              "  PrintedDateTime datetime, " & _
              "  ValidatedBy nvarchar(50), " & _
              "  ValidatedDateTime datetime )"
63210   Cnxn(0).Execute sql
63220 End If

63230 Exit Sub

CheckPrintValidLogInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63240 intEL = Erl
63250 strES = Err.Description
63260 LogError "modDbDesign", "CheckPrintValidLogInDb", intEL, strES, sql

End Sub
Public Sub CheckSensitivitiesArcInDb()

      Dim sql As String

63270 On Error GoTo CheckSensitivitiesArcInDb_Error

63280 If IsTableInDatabase("SensitivitiesArc") = False Then 'There is no table  in database
63290   sql = "CREATE TABLE SensitivitiesArc " & _
              "( SampleID numeric, " & _
              "  IsolateNumber int, " & _
              "  AntibioticCode nvarchar(20), " & _
              "  Result nvarchar(10), " & _
              "  Report bit, " & _
              "  CPOFlag nvarchar(1), " & _
              "  RunDate datetime, " & _
              "  RunDateTime datetime, " & _
              "  RSI char(1), " & _
              "  UserCode nvarchar(5), " & _
              "  Forced bit, " & _
              "  Secondary bit, " & _
              "  Valid bit, " & _
              "  AuthoriserCode nvarchar(5), " & _
              "  ArchiveDateTime datetime, " & _
              "  ArchivedBy nvarchar(50) )"
63300   Cnxn(0).Execute sql
63310 End If

63320 Exit Sub

CheckSensitivitiesArcInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63330 intEL = Erl
63340 strES = Err.Description
63350 LogError "modDbDesign", "CheckSensitivitiesArcInDb", intEL, strES, sql

End Sub

Public Sub CheckSensitivitiesRepeatsInDb()

      Dim sql As String

63360 On Error GoTo CheckSensitivitiesRepeatsInDb_Error

63370 If IsTableInDatabase("SensitivitiesRepeats") = False Then 'There is no table  in database
63380   sql = "CREATE TABLE SensitivitiesRepeats " & _
              "( SampleID numeric, " & _
              "  IsolateNumber int, " & _
              "  AntibioticCode nvarchar(20), " & _
              "  Result nvarchar(10), " & _
              "  Report bit, " & _
              "  CPOFlag nvarchar(1), " & _
              "  RunDate datetime, " & _
              "  RunDateTime datetime, " & _
              "  RSI char(1), " & _
              "  UserCode nvarchar(5), " & _
              "  Forced bit, " & _
              "  Secondary bit, " & _
              "  Valid bit, " & _
              "  AuthoriserCode nvarchar(5) )"
63390   Cnxn(0).Execute sql
63400 End If

63410 Exit Sub

CheckSensitivitiesRepeatsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63420 intEL = Erl
63430 strES = Err.Description
63440 LogError "modDbDesign", "CheckSensitivitiesRepeatsInDb", intEL, strES, sql

End Sub


Public Sub CheckIsolatesArcInDb()

      Dim sql As String

63450 On Error GoTo CheckIsolatesArcInDb_Error

63460 If IsTableInDatabase("IsolatesArc") = False Then 'There is no table  in database
63470   sql = "CREATE TABLE IsolatesArc " & _
              "( SampleID numeric, " & _
              "  IsolateNumber int, " & _
              "  OrganismGroup nvarchar(50), " & _
              "  OrganismName nvarchar(50), " & _
              "  Qualifier nvarchar(50), " & _
              "  ArchiveDateTime datetime, " & _
              "  ArchivedBy nvarchar(50) )"
63480   Cnxn(0).Execute sql
63490 End If

63500 Exit Sub

CheckIsolatesArcInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63510 intEL = Erl
63520 strES = Err.Description
63530 LogError "modDbDesign", "CheckIsolatesArcInDb", intEL, strES, sql

End Sub


Public Sub CheckIsolatesRepeatsInDb()

      Dim sql As String

63540 On Error GoTo CheckIsolatesRepeatsInDb_Error

63550 If IsTableInDatabase("IsolatesRepeats") = False Then 'There is no table  in database
63560   sql = "CREATE TABLE IsolatesRepeats " & _
              "( SampleID numeric, " & _
              "  IsolateNumber int, " & _
              "  OrganismGroup nvarchar(50), " & _
              "  OrganismName nvarchar(50), " & _
              "  Qualifier nvarchar(50) )"
63570   Cnxn(0).Execute sql
63580 End If

63590 Exit Sub

CheckIsolatesRepeatsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63600 intEL = Erl
63610 strES = Err.Description
63620 LogError "modDbDesign", "CheckIsolatesRepeatsInDb", intEL, strES, sql

End Sub
Public Sub CheckUProInDb()

      Dim sql As String

63630 On Error GoTo CheckUProInDb_Error

63640 If IsTableInDatabase("UPro") = False Then 'There is no table  in database
63650   sql = "CREATE TABLE UPro " & _
              "( SampleID numeric(9), " & _
              "  CollectionPeriod tinyint, " & _
              "  TotalVolume int, " & _
              "  UPgPerL real, " & _
              "  UP24H real, " & _
              "  PrintedDateTime datetime, " & _
              "  PrintedBy nvarchar(50) )"
63660   Cnxn(0).Execute sql
63670 End If

63680 Exit Sub

CheckUProInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63690 intEL = Erl
63700 strES = Err.Description
63710 LogError "modDbDesign", "CheckUProInDb", intEL, strES, sql

End Sub


Public Sub CheckIncludeEGFRInDb()

      Dim sql As String

63720 On Error GoTo CheckIncludeEGFRInDb_Error

63730 If IsTableInDatabase("IncludeEGFR") = False Then 'There is no table  in database
63740   sql = "CREATE TABLE [IncludeEGFR] (" & _
              "[SourceType] [nvarchar] (50) NULL, " & _
              "[Hospital] [nvarchar] (50) NULL, " & _
              "[SourceName] [nvarchar] (50) NULL, " & _
              "[Include] [bit] NULL, " & _
              "[Counter] [numeric](18, 0) IDENTITY (1, 1) NOT NULL )"
63750   Cnxn(0).Execute sql
63760 End If

63770 Exit Sub

CheckIncludeEGFRInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63780 intEL = Erl
63790 strES = Err.Description
63800 LogError "modDbDesign", "CheckIncludeEGFRInDb", intEL, strES, sql

End Sub

Public Sub CheckIncludeAutoValUrineInDb()

      Dim sql As String

63810 On Error GoTo CheckIncludeAutoValUrineInDb_Error

63820 If IsTableInDatabase("IncludeAutoValUrine") = False Then 'There is no table  in database
63830   sql = "CREATE TABLE [IncludeAutoValUrine] (" & _
              "[SourceType] [nvarchar] (50) NOT NULL, " & _
              "[Hospital] [nvarchar] (50) NOT NULL, " & _
              "[SourceName] [nvarchar] (50) NOT NULL, " & _
              "[Include] [bit] NOT NULL DEFAULT 0, " & _
              "[Counter] [numeric](18, 0) IDENTITY (1, 1) NOT NULL )"
63840   Cnxn(0).Execute sql
63850 End If

63860 Exit Sub

CheckIncludeAutoValUrineInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63870 intEL = Erl
63880 strES = Err.Description
63890 LogError "modDbDesign", "CheckIncludeAutoValUrineInDb", intEL, strES, sql

End Sub

Public Sub CheckLoggedOnUsersInDb()

      Dim sql As String

63900 On Error GoTo CheckLoggedOnUsersInDb_Error

63910 If IsTableInDatabase("LoggedOnUsers") = False Then 'There is no table  in database
63920   sql = "CREATE TABLE dbo.[LoggedOnUsers](" & _
              "  [MachineName] [nvarchar](50) NULL, " & _
              "  [UserName] [nvarchar](50) NULL, " & _
              "  [AppName] [nvarchar](50) NULL)"
63930   Cnxn(0).Execute sql
63940 End If

63950 Exit Sub

CheckLoggedOnUsersInDb_Error:

      Dim strES As String
      Dim intEL As Integer

63960 intEL = Erl
63970 strES = Err.Description
63980 LogError "modDbDesign", "CheckLoggedOnUsersInDb", intEL, strES, sql

End Sub


Public Sub CheckObservationsInDb()

      Dim sql As String

63990 On Error GoTo CheckObservationsInDb_Error

64000 If IsTableInDatabase("Observations") = False Then 'There is no table  in database
64010   sql = "CREATE TABLE dbo.[Observations](" & _
              "  [SampleID] [numeric](18,0) NOT NULL, " & _
              "  [Discipline] [nvarchar](50) NOT NULL, " & _
              "  [Comment] [nvarchar](4000) NULL, " & _
              "  [UserName] [nvarchar] (50) NULL, " & _
              "  [DateTimeOfRecord] datetime NOT NULL DEFAULT getdate())"
64020   Cnxn(0).Execute sql
64030 End If

64040 Exit Sub

CheckObservationsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64050 intEL = Erl
64060 strES = Err.Description
64070 LogError "modDbDesign", "CheckObservationsInDb", intEL, strES, sql

End Sub


Public Sub CheckCustomEventLogInDb()

      Dim sql As String

64080 On Error GoTo CheckCustomEventLogInDb_Error

64090 If IsTableInDatabase("CustomEventLog") = False Then 'There is no table  in database
64100   sql = "CREATE TABLE CustomEventLog " & _
              "( MSG ntext, " & _
              "  DateTime datetime, " & _
              "  Dept nvarchar(50), " & _
              "  ModuleName  nvarchar(50), " & _
              "  ProcedureName nvarchar(50), " & _
              "  UserName nvarchar(50), " & _
              "  MachineName nvarchar(50) )"
64110   Cnxn(0).Execute sql
64120 End If

64130 Exit Sub

CheckCustomEventLogInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64140 intEL = Erl
64150 strES = Err.Description
64160 LogError "modDbDesign", "CheckCustomEventLogInDb", intEL, strES, sql

End Sub

Public Sub CheckPhoneLogInDb()

      Dim sql As String

64170 On Error GoTo CheckPhoneLogInDb_Error

64180 If IsTableInDatabase("PhoneLog") = False Then 'There is no table  in database
64190   sql = "CREATE TABLE PhoneLog " & _
              "( SampleID numeric, " & _
              "  DateTime datetime, " & _
              "  PhonedTo nvarchar(50), " & _
              "  PhonedBy nvarchar(50), " & _
              "  Comment nvarchar(50), " & _
              "  Discipline nvarchar(50) )"

64200 Cnxn(0).Execute sql
64210 End If

64220 Exit Sub

CheckPhoneLogInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64230 intEL = Erl
64240 strES = Err.Description
64250 LogError "modDbDesign", "CheckPhoneLogInDb", intEL, strES, sql

End Sub

Public Sub CheckUserProfilesInDb()

      Dim sql As String

64260 On Error GoTo CheckUserProfilesInDb_Error

64270 If IsTableInDatabase("UserProfiles") = False Then 'There is no table  in database
64280   sql = "CREATE TABLE UserProfiles " & _
              "( ProfileName nvarchar(50), " & _
              "  ProfileFunction nvarchar(50), " & _
              "  UserName nvarchar(50), " & _
              "  DateTimeOfRecord datetime DEFAULT getdate() )"

64290 Cnxn(0).Execute sql
64300 End If

64310 Exit Sub

CheckUserProfilesInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64320 intEL = Erl
64330 strES = Err.Description
64340 LogError "modDbDesign", "CheckUserProfilesInDb", intEL, strES, sql

End Sub


Public Sub CheckCSFResultsInDb()

      Dim sql As String

64350 On Error GoTo CheckCSFResultsInDb_Error

64360 If IsTableInDatabase("CSFResults") = False Then 'There is no table  in database
64370   sql = "CREATE TABLE CSFResults " & _
              "( SampleID numeric NOT NULL, " & _
              "  Gram nvarchar(50), " & _
              "  WCCDiff0 nvarchar(50), " & _
              "  WCCDiff1 nvarchar(50), " & _
              "  Appearance0 nvarchar(50), " & _
              "  Appearance1 nvarchar(50), " & _
              "  Appearance2 nvarchar(50), " & _
              "  WCC0 nvarchar(50), " & _
              "  WCC1 nvarchar(50), " & _
              "  WCC2 nvarchar(50), " & _
              "  RCC0 nvarchar(50), " & _
              "  RCC1 nvarchar(50), " & _
              "  RCC2 nvarchar(50), " & _
              "  UserName nvarchar(50) NOT NULL, " & _
              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"

64380 Cnxn(0).Execute sql
64390 End If

64400 Exit Sub

CheckCSFResultsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64410 intEL = Erl
64420 strES = Err.Description
64430 LogError "modDbDesign", "CheckCSFResultsInDb", intEL, strES, sql

End Sub

Public Sub CheckPrintValidLogArcInDb()

      Dim sql As String

64440 On Error GoTo CheckPrintValidLogArcInDb_Error

64450 If IsTableInDatabase("PrintValidLogArc") = False Then 'There is no table  in database
64460   sql = "CREATE TABLE PrintValidLogArc " & _
              "( SampleID numeric(9), " & _
              "  Department nvarchar(1), " & _
              "  Printed tinyint, " & _
              "  Valid tinyint, " & _
              "  PrintedBy nvarchar(50), " & _
              "  PrintedDateTime datetime, " & _
              "  ValidatedBy nvarchar(50), " & _
              "  ValidatedDateTime datetime, " & _
              "  ArchivedBy nvarchar(50), " & _
              "  ArchivedDateTime datetime )"
64470   Cnxn(0).Execute sql
64480 End If

64490 Exit Sub

CheckPrintValidLogArcInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64500 intEL = Erl
64510 strES = Err.Description
64520 LogError "modDbDesign", "CheckPrintValidLogArcInDb", intEL, strES, sql

End Sub


Public Sub CheckHealthlinkInDb()

      Dim sql As String

64530 On Error GoTo CheckHealthlinkInDb_Error

64540 If IsTableInDatabase("Healthlink") = False Then 'There is no table  in database
64550   sql = "CREATE TABLE Healthlink " & _
              "( RunTime datetime, " & _
              "  Message ntext )"
64560   Cnxn(0).Execute sql
64570 End If

64580 Exit Sub

CheckHealthlinkInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64590 intEL = Erl
64600 strES = Err.Description
64610 LogError "modDbDesign", "CheckHealthlinkInDb", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckGPOrderPatientInDb
' Author    : Masood
' Date      : 31/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckGPOrderPatientInDb()

64620     On Error GoTo CheckGPOrderPatientInDb_Error

          Dim sql As String


64630     If IsTableInDatabase("GPOrderPatient") = False Then
64640         sql = "CREATE TABLE GPOrderPatient ( " & _
                  " [GPName] [nvarchar](20) NULL," & _
                  " [GPNumber] [nvarchar](10) NULL," & _
                  " [DateTimeOfMessage] [datetime] NULL," & _
                  " [PatientID] [nvarchar](10) NULL," & _
                  " [PatientSurName] [nvarchar](20) NULL," & _
                  " [PatientForeName] [nvarchar](20) NULL," & _
                  " [DoB] [date] NULL," & _
                  " [Sex] [nvarchar](6) NULL," & _
                  " [Addr1] [nvarchar](20) NULL," & _
                  " [Addr2] [nvarchar](20) NULL," & _
                  " [Addr3] [nvarchar](20) NULL," & _
                  " [Addr4] [nvarchar](20) NULL," & _
                  " [Addr5] [nvarchar](20) NULL," & _
                  " [PracticeID] [nvarchar](10) NULL," & _
                  " [GPSurName] [nvarchar](20) NULL," & _
                  " [GPForeName] [nvarchar](20) NULL," & _
                  " [Pregnant] [nvarchar](25) NULL," & _
                  " [GID] [nvarchar](50) NULL," & _
                  " [FileName] [nvarchar](50) NULL," & _
                  " [SampleIDExternal] [nvarchar](50) NULL" & _
                  " ) "
64650         Cnxn(0).Execute sql
64660     End If


64670     Exit Sub


CheckGPOrderPatientInDb_Error:

          Dim strES As String
          Dim intEL As Integer

64680     intEL = Erl
64690     strES = Err.Description
64700     LogError "modDbDesign", "CheckGPOrderPatientInDb", intEL, strES, sql
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckGPOrdersInDb
' Author    : Masood
' Date      : 31/Mar/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckGPOrdersInDb()

64710     On Error GoTo CheckGPOrdersInDb_Error

        Dim sql As String
        
64720     If IsTableInDatabase("GPOrders") = False Then
64730         sql = "CREATE TABLE GPOrders ( " & _
                  " ShortName nvarchar(20) NULL," & _
                    "LongName nvarchar(50) NULL," & _
                    "ClinicalDetails nvarchar(20) NULL," & _
                    "SampleTypeCode nvarchar(20) NULL," & _
                    "SampleType nvarchar(50) NULL," & _
                    "Priority nvarchar(10) NULL," & _
                    "GID nvarchar(50) NULL," & _
                    "FileName nvarchar(50) NULL," & _
                    "SampleIDExternal nvarchar(50) NULL," & _
                    " SampleDate [datetime] NULL " & _
                    ") "
64740         Cnxn(0).Execute sql
64750     End If


64760     Exit Sub


CheckGPOrdersInDb_Error:

          Dim strES As String
          Dim intEL As Integer

64770     intEL = Erl
64780     strES = Err.Description
64790     LogError "modDbDesign", "CheckGPOrdersInDb", intEL, strES, sql
End Sub



Public Sub CheckAU400BottleLotInDb()

      Dim sql As String

64800 On Error GoTo CheckAU400BottleLotInDb_Error

64810 If IsTableInDatabase("AU400BottleLot") = False Then 'There is no table  in database
64820   sql = "CREATE TABLE AU400BottleLot " & _
              "(SampleID  numeric, " & _
              " Code  nvarchar(50), " & _
              " R1Bottle nvarchar(50), " & _
              " R1Lot nvarchar(50), " & _
              " R2Bottle nvarchar(50), " & _
              " R2Lot nvarchar(50), " & _
              " [DateTime]  datetime, " & _
              " GUID uniqueidentifier NOT NULL DEFAULT NEWID() )"
64830   Cnxn(0).Execute sql
64840 End If

64850 Exit Sub

CheckAU400BottleLotInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64860 intEL = Erl
64870 strES = Err.Description
64880 LogError "modDbDesign", "CheckAU400BottleLotInDb", intEL, strES, sql

End Sub

Public Sub CheckPhoneAlertInDb()

      Dim sql As String

64890 On Error GoTo CheckPhoneAlertInDb_Error

64900 If IsTableInDatabase("PhoneAlert") = False Then
64910     sql = "CREATE TABLE [dbo].[PhoneAlert] (" & _
                  "[SampleID] [decimal](18, 0) NULL , " & _
                  "[Discipline] [nvarchar] (50) NULL , " & _
                  "[Parameter] [nvarchar] (50) NULL )"
64920     Cnxn(0).Execute sql
64930 End If

64940 Exit Sub

CheckPhoneAlertInDb_Error:

      Dim strES As String
      Dim intEL As Integer

64950 intEL = Erl
64960 strES = Err.Description
64970 LogError "modDbDesign", "CheckPhoneAlertInDb", intEL, strES, sql

End Sub





Public Sub CheckPhoneAlertLevelInDb()

      Dim sql As String

64980 On Error GoTo CheckPhoneAlertLevelInDb_Error

64990 If IsTableInDatabase("PhoneAlertLevel") = False Then
65000     sql = "CREATE TABLE [dbo].[PhoneAlertLevel] (" & _
                  "[Discipline] [nvarchar] (50) NOT NULL , " & _
                  "[Parameter] [nvarchar] (50) NOT NULL , " & _
                  "[LessThan] [real] NULL , " & _
                  "[GreaterThan] [real] NULL )"
65010     Cnxn(0).Execute sql
65020 End If

65030 Exit Sub

CheckPhoneAlertLevelInDb_Error:

      Dim strES As String
      Dim intEL As Integer

65040 intEL = Erl
65050 strES = Err.Description
65060 LogError "modDbDesign", "CheckPhoneAlertLevelInDb", intEL, strES, sql

End Sub

Public Sub CheckCommentTemplateInDb()

      Dim sql As String

65070 On Error GoTo CheckCommentTemplateInDb_Error

65080 If IsTableInDatabase("CommentsTemplate") = False Then
65090     sql = "CREATE TABLE [dbo].[CommentsTemplate] (" & _
                  "[CommentID] [int] IDENTITY (1, 1) NOT NULL ," & _
                  "[CommentName] [nvarchar] (200) COLLATE Latin1_General_CI_AS NULL ," & _
                  "[CommentTemplate] [ntext] COLLATE Latin1_General_CI_AS NULL ," & _
                  "[Inactive] [tinyint] NULL ," & _
                  "[Department] [nvarchar] (10) COLLATE Latin1_General_CI_AS NULL ," & _
                  "[Username] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL ," & _
                  "[DateTimeOfRecord] [nvarchar] (50) COLLATE Latin1_General_CI_AS NULL)"
65100     Cnxn(0).Execute sql
65110 End If

65120 Exit Sub

CheckCommentTemplateInDb_Error:

      Dim strES As String
      Dim intEL As Integer

65130 intEL = Erl
65140 strES = Err.Description
65150 LogError "modDbDesign", "CheckCommentTemplateInDb", intEL, strES, sql

End Sub

Public Sub CheckIdentificationInDb()

      Dim sql As String

65160 On Error GoTo CheckIdentificationInDb_Error

65170 If IsTableInDatabase("Identification") = False Then
65180     sql = "CREATE TABLE [dbo].[Identification]( " & _
                  "[SampleID] [nvarchar](50) NULL, " & _
                  "[TestType] [nvarchar](50) NULL, " & _
                  "[TestName] [nvarchar](50) NULL, " & _
                  "[Result] [nvarchar](50) NULL, " & _
                  "[TestDateTime] [datetime] NULL, " & _
                  "[Valid] [tinyint] NULL, " & _
                  "[Printed] [tinyint] NULL, " & _
                  "[Username] [nvarchar](50) NULL, " & _
                  "[DateTimeOfRecord] [datetime] NULL )"
          
65190     Cnxn(0).Execute sql
65200 End If
65210 Exit Sub

CheckIdentificationInDb_Error:

      Dim strES As String
      Dim intEL As Integer

65220 intEL = Erl
65230 strES = Err.Description
65240 LogError "modDbDesign", "CheckIdentificationInDb", intEL, strES, sql

End Sub

Public Sub CheckIdentificationArcInDb()

      Dim sql As String

65250 On Error GoTo CheckIdentificationArcInDb_Error

65260 If IsTableInDatabase("IdentificationArc") = False Then

65270     sql = "CREATE TABLE [dbo].[IdentificationArc]( " & _
                "[SampleID] [nvarchar](50) NULL, " & _
                "[TestType] [nvarchar](50) NULL, " & _
                "[TestName] [nvarchar](50) NULL, " & _
                "[Result] [nvarchar](50) NULL, " & _
                "[TestDateTime] [datetime] NULL, " & _
                "[Valid] [tinyint] NULL, " & _
                "[Printed] [tinyint] NULL, " & _
                "[Username] [nvarchar](50) NULL, " & _
                "[DateTimeOfRecord] [datetime] NULL, " & _
                "[ArchivedBy] [nvarchar](50) NULL, " & _
                "[ArchiveDateTime] [datetime] NULL )"

65280     Cnxn(0).Execute sql
65290 End If
65300 Exit Sub

CheckIdentificationArcInDb_Error:

      Dim strES As String
      Dim intEL As Integer

65310 intEL = Erl
65320 strES = Err.Description
65330 LogError "modDbDesign", "CheckIdentificationArcInDb", intEL, strES, sql

End Sub


Public Sub CheckScannedImagesInDb()

      Dim sql As String

65340 On Error GoTo CheckScannedImagesInDb_Error

65350 If IsTableInDatabase("ScannedImages") = False Then
65360     sql = "CREATE TABLE [dbo].[ScannedImages]( " & _
                  "[SampleID] [nvarchar](50) NOT NULL, " & _
                  "[ScannedName] [nvarchar](50) NOT NULL, " & _
                  "[ScannedImage] [image] NOT NULL, " & _
                  "[RemoveFromLisDisplay] [bit] NOT NULL, " & _
                  "[RowGUID] [uniqueidentifier] NOT NULL DEFAULT newid(), " & _
                  "[DateTimeOfRecord] [datetime] Not NULL " & _
              ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] "
          
65370     Cnxn(0).Execute sql
65380 End If


65390 Exit Sub

CheckScannedImagesInDb_Error:

       Dim strES As String
       Dim intEL As Integer

65400  intEL = Erl
65410  strES = Err.Description
65420  LogError "modDbDesign", "CheckScannedImagesInDb", intEL, strES, sql
          
End Sub

