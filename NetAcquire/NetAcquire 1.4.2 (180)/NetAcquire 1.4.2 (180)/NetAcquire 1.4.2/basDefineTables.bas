Attribute VB_Name = "basDefineTables"
Option Explicit


Public Type FieldDefs
  ColumnName As String
  DataType As String
  Length As Long
  NoNull As Boolean
  DirectionASC As Boolean
End Type

Public Design() As FieldDefs

Public Sub DefineABDefinitions()

ReDim Design(0 To 4) As FieldDefs

FillDesignL 0, "AntibioticName", "nvarchar", 50
FillDesignL 1, "OrganismGroup", "nvarchar", 50
FillDesignL 2, "Site", "nvarchar", 50
FillDesignL 3, "ListOrder", "int"
FillDesignL 4, "PriSec", "nvarchar", 50

If IsTableInDB("ABDefinitions") = False Then 'There is no table  in database
  CreateTable "ABDefinitions"
Else
  DoTableAnalysis "ABDefinitions"
End If

End Sub

Public Sub DefineAntibiotics()

ReDim Design(0 To 7) As FieldDefs

FillDesignL 0, "AntibioticName", "nvarchar", 50
FillDesignL 1, "ListOrder", "int"
FillDesignL 2, "AllowIfPregnant", "int"
FillDesignL 3, "AllowIfOutPatient", "int"
FillDesignL 4, "AllowIfChild", "int"
FillDesignL 5, "ABC", "nvarchar", 50
FillDesignL 6, "Code", "nvarchar", 50
FillDesignL 7, "AllowIfPenAll", "int"

If IsTableInDB("Antibiotics") = False Then 'There is no table  in database
  CreateTable "Antibiotics"
Else
  DoTableAnalysis "Antibiotics"
End If

End Sub


Public Sub DefineAnswers()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "Text", "int"

If IsTableInDB("Answers") = False Then 'There is no table  in database
  CreateTable "Answers"
Else
  DoTableAnalysis "Answers"
End If

End Sub
Public Sub DefineInterp()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "Val", "float"
FillDesignL 0, "Interp", "nvarchar", 50
FillDesignL 1, "Sign", "nvarchar", 50
FillDesignL 1, "ListOrder", "int"
 
If IsTableInDB("Interp") = False Then 'There is no table  in database
  CreateTable "Interp"
Else
  DoTableAnalysis "Interp"
End If

End Sub

Public Sub DefineBactTestDefinitions()

ReDim Design(0 To 16) As FieldDefs

FillDesignL 0, "ListOrder", "smallint"
FillDesignL 1, "Printable", "int"
FillDesignL 2, "AssociatedText", "bit"
FillDesignL 3, "TextPos", "nvarchar", 50
FillDesignL 4, "TextNeg", "nvarchar", 50
FillDesignL 5, "TextInd", "nvarchar", 50
FillDesignL 6, "CodeUsed", "bit"
FillDesignL 7, "Antibiotic", "bit"
FillDesignL 8, "ShowAs", "nvarchar", 50
FillDesignL 9, "InUse", "int"
FillDesignL 10, "CounterID", "int"
FillDesignL 11, "Welcans", "real"
FillDesignL 12, "Triple", "bit"
FillDesignL 13, "SampleCode", "nvarchar", 50
FillDesignL 14, "TestCode", "nvarchar", 50
FillDesignL 15, "IsolateCode", "nvarchar", 50
FillDesignL 16, "RequestCode", "nvarchar", 50


If IsTableInDB("BactTestDefinitions") = False Then 'There is no table  in database
  CreateTable "BactTestDefinitions"
Else
  DoTableAnalysis "BactTestDefinitions"
End If

End Sub

Public Sub DefineBarCodeControl()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Text", "nvarchar", 50
FillDesignL 1, "Code", "nvarchar", 50

If IsTableInDB("BarCodeControl") = False Then 'There is no table  in database
  CreateTable "BarCodeControl"
Else
  DoTableAnalysis "BarCodeControl"
End If

End Sub


Public Sub DefineBarCodes()

ReDim Design(0 To 10) As FieldDefs

FillDesignL 0, "SaveRequests", "nvarchar", 50
FillDesignL 1, "ClearRequests", "nvarchar", 50
FillDesignL 2, "Cancel", "nvarchar", 50
FillDesignL 3, "Random", "nvarchar", 50
FillDesignL 4, "Fasting", "nvarchar", 50
FillDesignL 5, "A", "nvarchar", 50
FillDesignL 6, "B", "nvarchar", 50
FillDesignL 7, "FBC", "nvarchar", 50
FillDesignL 8, "ESR", "nvarchar", 50
FillDesignL 9, "Retics", "nvarchar", 50
FillDesignL 10, "Monospot", "nvarchar", 50

If IsTableInDB("BarCodes") = False Then 'There is no table  in database
  CreateTable "BarCodes"
Else
  DoTableAnalysis "BarCodes"
End If

End Sub


Public Sub DefineResults()

ReDim Design(0 To 18) As FieldDefs
Dim n As Integer
Dim pTable As String

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "Code", "nvarchar", 50, True
FillDesignL 2, "Result", "nvarchar", 50
FillDesignL 3, "Valid", "tinyint"
FillDesignL 4, "Printed", "tinyint"
FillDesignL 5, "RunTime", "datetime"
FillDesignL 6, "RunDate", "datetime"
FillDesignL 7, "Operator", "nvarchar", 50
FillDesignL 8, "Flags", "nvarchar", 50
FillDesignL 9, "Units", "nvarchar", 50
FillDesignL 10, "SampleType", "nvarchar", 50
FillDesignL 11, "Analyser", "nvarchar", 50
FillDesignL 12, "Faxed", "tinyint"
FillDesignL 13, "NOPAS", "nvarchar", 50
FillDesignL 14, "Authorised", "tinyint"
FillDesignL 15, "Comment", "nvarchar", 100
FillDesignL 16, "PC", "nvarchar", 50
FillDesignL 17, "Healthlink", "tinyint"
FillDesignL 18, "DefIndex", "int"

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "Results") = False Then 'There is no table  in database
    CreateTable pTable & "Results"
  Else
    DoTableAnalysis pTable & "Results"
  End If
Next

End Sub
Public Sub DefineExtResults()

ReDim Design(0 To 12) As FieldDefs

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "Analyte", "nvarchar", 50, True
FillDesignL 2, "Result", "nvarchar", 50
FillDesignL 3, "SendTo", "nvarchar", 50
FillDesignL 4, "Units", "nvarchar", 50
FillDesignL 5, "Date", "datetime"
FillDesignL 6, "RetDate", "datetime"
FillDesignL 7, "SentDate", "datetime"
FillDesignL 8, "SAPCode", "nvarchar", 50
FillDesignL 9, "HealthLink", "int"
FillDesignL 10, "OrderList", "int"
FillDesignL 11, "UserName", "nvarchar", 50
FillDesignL 12, "SaveTime", "datetime"

If IsTableInDB("ExtResults") = False Then 'There is no table  in database
  CreateTable "ExtResults"
Else
  DoTableAnalysis "ExtResults"
End If

End Sub

Public Sub DefineExtResultsArc()

ReDim Design(0 To 12) As FieldDefs

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "Analyte", "nvarchar", 50, True
FillDesignL 2, "Result", "nvarchar", 50
FillDesignL 3, "SendTo", "nvarchar", 50
FillDesignL 4, "Units", "nvarchar", 50
FillDesignL 5, "Date", "datetime"
FillDesignL 6, "RetDate", "datetime"
FillDesignL 7, "SentDate", "datetime"
FillDesignL 8, "SAPCode", "nvarchar", 50
FillDesignL 9, "HealthLink", "int"
FillDesignL 10, "OrderList", "int"
FillDesignL 11, "UserName", "nvarchar", 50
FillDesignL 12, "SaveTime", "datetime"

If IsTableInDB("ExtResultsArc") = False Then 'There is no table  in database
  CreateTable "ExtResultsArc"
Else
  DoTableAnalysis "ExtResultsArc"
End If

End Sub

Public Sub DefineExtTests()

ReDim Design(0 To 4) As FieldDefs

FillDesignL 0, "TestNumber", "int"
FillDesignL 1, "TestName", "nvarchar", 50
FillDesignL 2, "SendTo", "nvarchar", 50
FillDesignL 3, "Units", "tinyint", 50
FillDesignL 4, "Normal", "nvarchar", 50

If IsTableInDB("ExtTests") = False Then 'There is no table  in database
  CreateTable "ExtTests"
Else
  DoTableAnalysis "ExtTests"
End If

End Sub

Public Sub DefineExtAddress()

ReDim Design(0 To 6) As FieldDefs

FillDesignL 0, "Addr0", "nvarchar", 50
FillDesignL 1, "Addr1", "nvarchar", 50
FillDesignL 2, "Addr2", "nvarchar", 50
FillDesignL 3, "Addr3", "nvarchar", 50
FillDesignL 4, "Phone", "nvarchar", 50
FillDesignL 5, "Fax", "nvarchar", 50
FillDesignL 6, "Code", "nvarchar", 50

If IsTableInDB("ExtAddress") = False Then 'There is no table  in database
  CreateTable "ExtAddress"
Else
  DoTableAnalysis "ExtAddress"
End If

End Sub


Public Sub DefineExternalDefinitions()

ReDim Design(0 To 9) As FieldDefs

FillDesignL 0, "AnalyteName", "nvarchar", 50
FillDesignL 1, "PrintPriority", "int"
FillDesignL 2, "Units", "nvarchar", 50
FillDesignL 3, "MaleLow", "real"
FillDesignL 4, "MaleHigh", "real"
FillDesignL 5, "FemaleLow", "real"
FillDesignL 6, "FemaleHigh", "real"
FillDesignL 7, "SendTo", "nvarchar", 50
FillDesignL 8, "MBCode", "nvarchar", 50
FillDesignL 9, "SampleType", "nvarchar", 50

If IsTableInDB("ExternalDefinitions") = False Then 'There is no table  in database
  CreateTable "ExternalDefinitions"
Else
  DoTableAnalysis "ExternalDefinitions"
End If

End Sub


Public Sub DefinePanels()

ReDim Design(0 To 5) As FieldDefs
Dim pTable As String
Dim n As Integer

FillDesignL 0, "PanelName", "nvarchar", 50
FillDesignL 1, "Content", "nvarchar", 50
FillDesignL 2, "BarCode", "nvarchar", 20
FillDesignL 3, "PanelType", "nvarchar", 2
FillDesignL 4, "Hospital", "nvarchar", 10
FillDesignL 5, "ListOrder", "int"

For n = 1 To 7
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag", "Ext")
  If IsTableInDB(pTable & "Panels") = False Then 'There is no table  in database
    CreateTable pTable & "Panels"
  Else
    DoTableAnalysis pTable & "Panels"
  End If
Next

End Sub

Public Sub DefinePatientIFs()

ReDim Design(0 To 17) As FieldDefs

FillDesignL 0, "Chart", "nvarchar", 50
FillDesignL 1, "PatName", "nvarchar", 50
FillDesignL 2, "Sex", "nvarchar", 50
FillDesignL 3, "DoB", "datetime"
FillDesignL 4, "Ward", "nvarchar", 50
FillDesignL 5, "Clinician", "nvarchar", 50
FillDesignL 6, "Address0", "nvarchar", 50
FillDesignL 7, "Address1", "nvarchar", 50
FillDesignL 8, "Entity", "nvarchar", 50
FillDesignL 9, "Episode", "nvarchar", 50
FillDesignL 10, "RegionalNumber", "nvarchar", 50
FillDesignL 11, "DateTimeAmended", "datetime"
FillDesignL 12, "NewEntry", "bit"
FillDesignL 13, "NOPAS", "nvarchar", 50
FillDesignL 14, "AandE", "nvarchar", 50
FillDesignL 15, "MRN", "nvarchar", 50
FillDesignL 16, "Address2", "nvarchar", 50
FillDesignL 17, "Address3", "nvarchar", 50

If IsTableInDB("PatientIFs") = False Then  'There is no table  in database
  CreateTable "PatientIFs"
Else
  DoTableAnalysis "PatientIFs"
End If

End Sub
Public Function IsTableInDB(ByVal TableName As String) As Boolean

Dim tbExists As Recordset
Dim sql As String
Dim Retval As Boolean

'How to find if a table exists in a database
'open a recordset with the following sql statement:
'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
'If the recordset it at eof then the table doesn't exist
'if it has a record then the table does exist.

On Error GoTo IsTableInDB_Error

sql = "SELECT name FROM sysobjects WHERE " & _
      "xtype = 'U' " & _
      "AND name = '" & TableName & "'"
Set tbExists = Cnxn(0).Execute(sql)

Retval = True

If tbExists.EOF Then 'There is no table <TableName> in database
  Retval = False
End If
IsTableInDB = Retval

Exit Function

IsTableInDB_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogErrorA "modDbDesign", "IsTableInDB", intEL, strES, sql
  
End Function

Public Sub DoTableAnalysis(ByVal TableName As String)

Dim n As Long
Dim f As Long
Dim Found As Boolean
Dim Matching As Boolean
Dim sql As String
Dim tb As Recordset
Dim s As String

sql = "Select top 1 * from [" & TableName & "]"
Set tb = New Recordset
RecOpenServer 0, tb, sql
For n = 0 To UBound(Design)
  Found = False
  Matching = False
  For f = 0 To tb.Fields.Count - 1
    If UCase$(tb.Fields(f).Name) = UCase$(Design(n).ColumnName) Then
      Found = True
      If ((tb.Fields(f).Type = 2 And UCase$(Design(n).DataType) = "SMALLINT") Or _
          (tb.Fields(f).Type = 4 And UCase$(Design(n).DataType) = "REAL") Or _
          (tb.Fields(f).Type = 5 And UCase$(Design(n).DataType) = "FLOAT") Or _
          (tb.Fields(f).Type = 16 And UCase$(Design(n).DataType) = "TINYINT") Or _
          (tb.Fields(f).Type = 17 And UCase$(Design(n).DataType) = "TINYINT") Or _
          (tb.Fields(f).Type = 203 And UCase$(Design(n).DataType) = "NTEXT") Or _
          (tb.Fields(f).Type = 205 And UCase$(Design(n).DataType) = "IMAGE") Or _
          (tb.Fields(f).Type = 135 And UCase$(Design(n).DataType) = "DATETIME") Or _
          (tb.Fields(f).Type = 131 And UCase$(Design(n).DataType) = "NUMERIC") Or _
          (tb.Fields(f).Type = 11 And UCase$(Design(n).DataType) = "BIT") Or _
          (tb.Fields(f).Type = 129 And UCase$(Design(n).DataType) = "CHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
          (tb.Fields(f).Type = 200 And UCase$(Design(n).DataType) = "VARCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
          (tb.Fields(f).Type = 130 And UCase$(Design(n).DataType) = "NCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
          (tb.Fields(f).Type = 202 And UCase$(Design(n).DataType) = "NVARCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
          (tb.Fields(f).Type = 3 And UCase$(Design(n).DataType) = "INT")) Then
        Matching = True
      End If
      Exit For
    End If
  Next
  s = ""
  If Not Found Then
    sql = "ALTER TABLE [" & TableName & "] " & _
          "ADD [" & Design(n).ColumnName & "] " & Design(n).DataType & " "
    If Design(n).Length <> 0 Then
      sql = sql & "(" & Design(n).Length & ") "
    End If
    If Design(n).NoNull = True Then
      sql = sql & "NOT "
    End If
    sql = sql & "NULL "
    Cnxn(0).Execute sql
  ElseIf Not Matching Then
    If tb.Fields(f).DefinedSize < Design(n).Length Then
      s = TableName & vbCrLf & "Column '" & tb.Fields(f).Name & "' " & _
          "should be " & Design(n).DataType & "(" & Design(n).Length & ")"
      iMsg s
    End If
  End If
Next
  
If Right$(TableName, 3) = "Arc" Then
  Found = False
  For f = 0 To tb.Fields.Count - 1
    If tb.Fields(f).Name = "ArchiveDateTime" Then
      Found = True
      Exit For
    End If
  Next
  If Not Found Then
    sql = "ALTER TABLE [" & TableName & "] ADD " & _
          "[ArchiveDateTime] datetime"
    Cnxn(0).Execute sql
  End If
  Found = False
  For f = 0 To tb.Fields.Count - 1
    If tb.Fields(f).Name = "ArchivedBy" Then
      Found = True
      Exit For
    End If
  Next
  If Not Found Then
    sql = "ALTER TABLE [" & TableName & "] ADD " & _
          "[ArchivedBy] nvarchar (50)"
    Cnxn(0).Execute sql
  End If
End If


End Sub

Public Sub CreateTable(ByVal TableName As String)
  
Dim sql As String
Dim n As Integer

sql = "CREATE TABLE " & TableName & " ( "
For n = 0 To UBound(Design)
  sql = sql & "[" & Design(n).ColumnName & "] " & _
              Design(n).DataType & " "
  If Design(n).Length <> 0 Then
    sql = sql & "(" & Design(n).Length & ") "
  End If
  sql = sql & IIf(Design(n).NoNull, " NOT NULL, ", "NULL, ")
Next
sql = Left$(sql, Len(sql) - 2) & ")"
Cnxn(0).Execute sql

If Right$(TableName, 3) = "Arc" Then
  sql = "ALTER TABLE [" & TableName & "] ADD " & _
        "[ArchiveDateTime] datetime, " & _
        "[ArchivedBy] nvarchar (50)"
  Cnxn(0).Execute sql
End If

End Sub


Public Sub CreateIndex(ByVal TableName As String, _
                       ByVal IndexName As String, _
                       ByRef cD() As FieldDefs, _
                       ByVal Unique As Boolean, _
                       ByVal Clustered As Boolean)

Dim sql As String
Dim n As Integer

sql = "IF NOT EXISTS (SELECT * FROM sysindexes WHERE " & _
      "               name = '" & IndexName & "') " & _
      "BEGIN " & _
      "  CREATE " & IIf(Unique, "UNIQUE", "") & " " & _
      "  " & IIf(Clustered, "CLUSTERED", "NONCLUSTERED") & " " & _
      "  INDEX " & IndexName & " " & _
      "  ON " & TableName & " " & _
      "  ("
For n = 0 To UBound(Design)
  sql = sql & cD(n).ColumnName & " " & IIf(cD(n).DirectionASC, "ASC", "DESC") & ","
Next
sql = Left$(sql, Len(sql) - 1) & ") END"
Cnxn(0).Execute sql

End Sub

Public Sub CheckBuild()

DefineErrorLog
DefineEventLog
DefineABDefinitions
DefineAnswers
DefineAntibiotics
DefineBacteriology
DefineBactIsolateDefinitions
DefineBactRequestDefinitions
DefineBactResults
DefineBactTestDefinitions
DefineBarCodeControl
DefineBarCodes
DefineRepeats
DefineRepeatsArc
DefineResults
DefineResultsArc
DefineTestDefinitions
DefineRequests
DefineBiochemistryQC
DefineBioFlags
DefineBioFlagsRep
DefineBioQCDefs
DefineBottleType
DefineCategorys
DefineCCMs
DefineClinDetails
DefineClinicians
DefineCoagTranslation
DefineComments
DefineCommentsArc
DefineConsultantList
DefineControls
DefineCreatinine
DefineCulture
DefineCytoResults
DefineCytoResultsArc
DefineDemographics
DefineDemographicsArc
DefineDifferentials
DefineDifferentialTitles
DefineDisease
DefineETC
DefineExtAddress
DefineExternalDefinitions
DefineExtResults
DefineExtResultsArc
DefineFaecalRequests
DefineFaeces
DefineFastings
DefineFilmReports
DefineForcedABReport
DefineGPs
'defineHaemAdvia
DefineHaemCondition
DefineHaemFlags
DefineHaemFlagsRep
DefineHbA1c
DefineHISErrors
DefineHistoBlock
DefineHistoComments
DefineHistoResults
DefineHistoResultsArc
DefineHistory
DefineHistoSpecimen
DefineHistoStain
DefineHMRU
DefineHospitals
DefineINRHistory
DefineInterp
DefineInstalledPrinters
DefineIsolates
DefineLabName
DefineLists
DefineMasks
DefineMaxMMessages
DefineMedibridgeRequests
DefineMedibridgeResults
DefineMicroRequests
DefineMicroSiteDetails
DefineMRU
DefineNameExclusions
DefineOP
DefineOptions
DefineOrganisms
DefinePanels
DefinePatientIFs
DefinePatientUpdates
DefinePhoneLog
DefinePractices
DefinePrinters
DefinePrintPending
DefineReagentArchive
DefineReagentLevel
DefineReagentList
DefineReagentLotNumbers
DefineReagentTestLevel
DefineReports
DefineSemenResults
DefineSendCopyTo
DefineSensitivities
DefineSexNames
DefineTabIndex
DefineTrace
DefineTrackBatchNumbers
DefineTrackMessage
DefineUpdates
DefineUrine
DefineUrineIdent
DefineUsers
DefineUsersArc
DefineViewedReports
DefineWards

IndexComments
IndexCommentsArc
IndexDemographics
IndexDemographicsArc
IndexErrorLog
IndexEventLog
IndexResults
IndexResultsArc
IndexRepeats
IndexRepeatsArc
IndexSexNames

frmImportHaem.Show 1

End Sub
Public Sub DefinePatientUpdates()

ReDim Design(0 To 6) As FieldDefs

FillDesignL 0, "Chart", "nvarchar", 50
FillDesignL 1, "PatName", "nvarchar", 50
FillDesignL 2, "Sex", "nvarchar", 50
FillDesignL 3, "DoB", "datetime"
FillDesignL 4, "NewChart", "nvarchar", 50
FillDesignL 5, "DateTime", "datetime"
FillDesignL 6, "Operator", "nvarchar", 50

If IsTableInDB("PatientUpdates") = False Then  'There is no table  in database
  CreateTable "PatientUpdates"
Else
  DoTableAnalysis "PatientUpdates"
End If


End Sub


Public Sub DefinePhoneLog()

ReDim Design(0 To 5) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "DateTime", "datetime"
FillDesignL 2, "PhonedTo", "nvarchar", 50
FillDesignL 3, "PhonedBy", "nvarchar", 50
FillDesignL 4, "Comment", "nvarchar", 50
FillDesignL 5, "Discipline", "nvarchar", 50

If IsTableInDB("PhoneLog") = False Then  'There is no table  in database
  CreateTable "PhoneLog"
Else
  DoTableAnalysis "PhoneLog"
End If


End Sub

Public Sub DefinePractices()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "Text", "nvarchar", 50
FillDesignL 1, "FAX", "nvarchar", 50
FillDesignL 2, "Hospital", "nvarchar", 50

If IsTableInDB("Practices") = False Then  'There is no table  in database
  CreateTable "Practices"
Else
  DoTableAnalysis "Practices"
End If

End Sub

Public Sub DefinePrinters()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "MappedTo", "nvarchar", 50
FillDesignL 1, "PrinterName", "nvarchar", 100

If IsTableInDB("Printers") = False Then  'There is no table  in database
  CreateTable "Printers"
Else
  DoTableAnalysis "Printers"
End If

End Sub

Public Sub DefinePrintPending()

ReDim Design(0 To 10) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Department", "nvarchar", 50
FillDesignL 2, "Initiator", "nvarchar", 50
FillDesignL 3, "UsePrinter", "nvarchar", 50
FillDesignL 4, "FaxNumber", "nvarchar", 50
FillDesignL 5, "ptime", "datetime"
FillDesignL 6, "UseConnection", "nvarchar", 50
FillDesignL 7, "Hyear", "nvarchar", 50
FillDesignL 8, "Ward", "nvarchar", 50
FillDesignL 9, "Clinician", "nvarchar", 50
FillDesignL 10, "GP", "nvarchar", 50

If IsTableInDB("PrintPending") = False Then  'There is no table  in database
  CreateTable "PrintPending"
Else
  DoTableAnalysis "PrintPending"
End If

End Sub
Public Sub DefineReagentArchive()

ReDim Design(0 To 5) As FieldDefs

FillDesignL 0, "Reagent", "nvarchar", 50
FillDesignL 1, "Amount", "numeric"
FillDesignL 2, "UserName", "nvarchar", 50
FillDesignL 3, "DateAdded", "datetime"
FillDesignL 4, "Comment", "nvarchar", 300
FillDesignL 5, "Dept", "nvarchar", 50

If IsTableInDB("ReagentArchive") = False Then  'There is no table  in database
  CreateTable "ReagentArchive"
Else
  DoTableAnalysis "ReagentArchive"
End If

End Sub

Public Sub DefineReagentLevel()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "Reagent", "nvarchar", 50
FillDesignL 1, "RLevel", "numeric"
FillDesignL 2, "Min", "numeric"
FillDesignL 3, "Dept", "nvarchar", 50

If IsTableInDB("ReagentLevel") = False Then  'There is no table  in database
  CreateTable "ReagentLevel"
Else
  DoTableAnalysis "ReagentLevel"
End If

End Sub

Public Sub DefineReagentTestLevel()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "Reagent", "nvarchar", 50
FillDesignL 1, "Test", "nvarchar", 50
FillDesignL 2, "Amount", "int"
FillDesignL 3, "Dept", "nvarchar", 50

If IsTableInDB("ReagentTestLevel") = False Then  'There is no table  in database
  CreateTable "ReagentTestLevel"
Else
  DoTableAnalysis "ReagentTestLevel"
End If

End Sub

Public Sub DefineReports()

ReDim Design(0 To 9) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Name", "nvarchar", 50
FillDesignL 2, "Dept", "nvarchar", 50
FillDesignL 3, "Initiator", "nvarchar", 50
FillDesignL 4, "PrintTime", "datetime"
FillDesignL 5, "RepNo", "nvarchar", 50
FillDesignL 6, "PageOne", "ntext"
FillDesignL 7, "PageTwo", "ntext"
FillDesignL 8, "Printer", "nvarchar", 50
FillDesignL 9, "Printed", "int"

If IsTableInDB("Reports") = False Then  'There is no table  in database
  CreateTable "Reports"
Else
  DoTableAnalysis "Reports"
End If

End Sub

Public Sub DefineSemenResults()

ReDim Design(0 To 9) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Volume", "nvarchar", 50
FillDesignL 2, "SemenCount", "nvarchar", 50
FillDesignL 3, "MotilityPro", "nvarchar", 50
FillDesignL 4, "MotilityNonPro", "nvarchar", 50
FillDesignL 5, "MotilityNonMotile", "nvarchar", 50
FillDesignL 6, "Consistency", "nvarchar", 50
FillDesignL 7, "Valid", "bit"
FillDesignL 8, "Operator", "nvarchar", 50
FillDesignL 9, "Printed", "bit"

If IsTableInDB("SemenResults") = False Then  'There is no table  in database
  CreateTable "SemenResults"
Else
  DoTableAnalysis "SemenResults"
End If

End Sub
Public Sub DefineSendCopyTo()

ReDim Design(0 To 5) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Ward", "nvarchar", 50
FillDesignL 2, "Clinician", "nvarchar", 50
FillDesignL 3, "GP", "nvarchar", 50
FillDesignL 4, "Device", "nvarchar", 50
FillDesignL 5, "Destination", "nvarchar", 50

If IsTableInDB("SendCopyTo") = False Then  'There is no table  in database
  CreateTable "SendCopyTo"
Else
  DoTableAnalysis "SendCopyTo"
End If

End Sub

Public Sub DefineSensitivities()

ReDim Design(0 To 16) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "IsolateNumber", "int"
FillDesignL 2, "AntibioticCode", "nvarchar", 50
FillDesignL 3, "Result", "nvarchar", 50
FillDesignL 4, "Report", "bit"
FillDesignL 5, "CPOFlag", "nvarchar", 50
FillDesignL 6, "RunDate", "datetime"
FillDesignL 7, "RunDateTime", "datetime"
FillDesignL 8, "RSI", "nvarchar", 50
FillDesignL 9, "UserCode", "nvarchar", 50
FillDesignL 10, "Forced", "bit"
FillDesignL 11, "Secondary", "bit"
FillDesignL 12, "Valid", "bit"
FillDesignL 13, "AuthoriserCode", "nvarchar", 50
FillDesignL 14, "OrgIndex", "int"
FillDesignL 15, "Organism", "nvarchar", 50
FillDesignL 16, "Antibiotic", "nvarchar", 50

If IsTableInDB("Sensitivities") = False Then  'There is no table  in database
  CreateTable "Sensitivities"
Else
  DoTableAnalysis "Sensitivities"
End If

End Sub


Public Sub DefineSexNames()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Name", "nvarchar", 50
FillDesignL 1, "Sex", "nvarchar", 50

If IsTableInDB("SexNames") = False Then  'There is no table  in database
  CreateTable "SexNames"
Else
  DoTableAnalysis "SexNames"
End If

End Sub



Public Sub DefineTabIndex()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "Form", "nvarchar", 50
FillDesignL 1, "Control", "nvarchar", 50
FillDesignL 2, "TabIndex", "int"
FillDesignL 3, "UserName", "nvarchar", 50

If IsTableInDB("TabIndex") = False Then  'There is no table  in database
  CreateTable "TabIndex"
Else
  DoTableAnalysis "TabIndex"
End If

End Sub


Public Sub DefineTrace()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "Trace", "ntext"
FillDesignL 1, "DateTime", "datetime"
FillDesignL 2, "Analyser", "nvarchar", 50

If IsTableInDB("Trace") = False Then  'There is no table  in database
  CreateTable "Trace"
Else
  DoTableAnalysis "Trace"
End If

End Sub

Public Sub DefineUrine()

ReDim Design(0 To 22) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Pregnancy", "nvarchar", 50
FillDesignL 2, "HCGLevel", "nvarchar", 50
FillDesignL 3, "BenceJones", "nvarchar", 50
FillDesignL 4, "SG", "nvarchar", 50
FillDesignL 5, "FatGlobules", "nvarchar", 50
FillDesignL 6, "pH", "nvarchar", 50
FillDesignL 7, "Protein", "nvarchar", 50
FillDesignL 8, "Glucose", "nvarchar", 50
FillDesignL 9, "Ketones", "nvarchar", 50
FillDesignL 10, "Urobilinogen", "nvarchar", 50
FillDesignL 11, "Bilirubin", "nvarchar", 50
FillDesignL 12, "BloodHb", "nvarchar", 50
FillDesignL 13, "WCC", "nvarchar", 50
FillDesignL 14, "RCC", "nvarchar", 50
FillDesignL 15, "Crystals", "nvarchar", 50
FillDesignL 16, "Casts", "nvarchar", 50
FillDesignL 17, "Misc0", "nvarchar", 50
FillDesignL 18, "Misc1", "nvarchar", 50
FillDesignL 19, "Misc2", "nvarchar", 50
FillDesignL 20, "Valid", "bit"
FillDesignL 21, "Bacteria", "nvarchar", 50
FillDesignL 22, "Count", "nvarchar", 50

If IsTableInDB("Urine") = False Then  'There is no table  in database
  CreateTable "Urine"
Else
  DoTableAnalysis "Urine"
End If

End Sub

Public Sub DefineUrineIdent()

ReDim Design(0 To 17) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Gram", "nvarchar", 50
FillDesignL 2, "WetPrep", "nvarchar", 50
FillDesignL 3, "Coagulase", "nvarchar", 50
FillDesignL 4, "Catalase", "nvarchar", 50
FillDesignL 5, "Oxidase", "nvarchar", 50
FillDesignL 6, "API0", "nvarchar", 50
FillDesignL 7, "API1", "nvarchar", 50
FillDesignL 8, "Ident0", "nvarchar", 50
FillDesignL 9, "Ident1", "nvarchar", 50
FillDesignL 10, "Rapidec", "nvarchar", 50
FillDesignL 11, "Chromogenic", "nvarchar", 50
FillDesignL 12, "Reincubation", "nvarchar", 50
FillDesignL 13, "UrineSensitivity", "nvarchar", 50
FillDesignL 14, "ExtraSensitivity", "nvarchar", 50
FillDesignL 15, "Valid", "bit"
FillDesignL 16, "Isolate", "int"
FillDesignL 17, "Notes", "nvarchar", 50

If IsTableInDB("UrineIdent") = False Then  'There is no table  in database
  CreateTable "UrineIdent"
Else
  DoTableAnalysis "UrineIdent"
End If

End Sub


Public Sub DefineUsers()

ReDim Design(0 To 16) As FieldDefs

FillDesignL 0, "PassWord", "nvarchar", 50
FillDesignL 1, "Name", "nvarchar", 50
FillDesignL 2, "Code", "nvarchar", 50
FillDesignL 3, "InUse", "bit"
FillDesignL 4, "MemberOf", "nvarchar", 50
FillDesignL 5, "LogOffDelay", "numeric"
FillDesignL 6, "ListOrder", "int"
FillDesignL 7, "Prints", "bit"
FillDesignL 8, "PassDate", "datetime"
FillDesignL 9, "Bio", "bit"
FillDesignL 10, "Haem", "bit"
FillDesignL 11, "Coag", "bit"
FillDesignL 12, "End", "bit"
FillDesignL 13, "Imm", "bit"
FillDesignL 14, "Ext", "bit"
FillDesignL 15, "Micro", "bit"
FillDesignL 16, "Histo", "bit"

If IsTableInDB("Users") = False Then  'There is no table  in database
  CreateTable "Users"
Else
  DoTableAnalysis "Users"
End If

End Sub

Public Sub DefineUsersArc()

ReDim Design(0 To 16) As FieldDefs

FillDesignL 0, "PassWord", "nvarchar", 50
FillDesignL 1, "Name", "nvarchar", 50
FillDesignL 2, "Code", "nvarchar", 50
FillDesignL 3, "InUse", "bit"
FillDesignL 4, "MemberOf", "nvarchar", 50
FillDesignL 5, "LogOffDelay", "numeric"
FillDesignL 6, "ListOrder", "int"
FillDesignL 7, "Prints", "bit"
FillDesignL 8, "PassDate", "datetime"
FillDesignL 9, "Bio", "bit"
FillDesignL 10, "Haem", "bit"
FillDesignL 11, "Coag", "bit"
FillDesignL 12, "End", "bit"
FillDesignL 13, "Imm", "bit"
FillDesignL 14, "Ext", "bit"
FillDesignL 15, "Micro", "bit"
FillDesignL 16, "Histo", "bit"

If IsTableInDB("UsersArc") = False Then  'There is no table  in database
  CreateTable "UsersArc"
Else
  DoTableAnalysis "UsersArc"
End If

End Sub


Public Sub DefineViewedReports()

ReDim Design(0 To 4) As FieldDefs

FillDesignL 0, "Discipline", "nvarchar", 50
FillDesignL 1, "DateTime", "datetime"
FillDesignL 2, "Viewer", "nvarchar", 50
FillDesignL 3, "SampleID", "numeric"
FillDesignL 4, "Chart", "nvarchar", 50

If IsTableInDB("ViewedReports") = False Then  'There is no table  in database
  CreateTable "ViewedReports"
Else
  DoTableAnalysis "ViewedReports"
End If

End Sub



Public Sub DefineWards()

ReDim Design(0 To 7) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "Text", "nvarchar", 50
FillDesignL 2, "InUse", "bit"
FillDesignL 3, "HospitalCode", "nvarchar", 50
FillDesignL 4, "FAX", "nvarchar", 50
FillDesignL 5, "ListOrder", "int"
FillDesignL 6, "PrinterAddress", "ntext"
FillDesignL 7, "Location", "nvarchar", 50

If IsTableInDB("Wards") = False Then  'There is no table  in database
  CreateTable "Wards"
Else
  DoTableAnalysis "Wards"
End If

End Sub



Public Sub DefineUpdates()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Upd", "nvarchar", 50
FillDesignL 1, "dtime", "datetime"

If IsTableInDB("Updates") = False Then  'There is no table  in database
  CreateTable "Updates"
Else
  DoTableAnalysis "Updates"
End If

End Sub


Public Sub DefineReagentList()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "Reagent", "nvarchar", 50
FillDesignL 1, "Code", "nvarchar", 50
FillDesignL 2, "Unit", "nvarchar", 50
FillDesignL 3, "Dept", "nvarchar", 50

If IsTableInDB("ReagentList") = False Then  'There is no table  in database
  CreateTable "ReagentList"
Else
  DoTableAnalysis "ReagentList"
End If

End Sub

Public Sub DefineReagentLotNumbers()

ReDim Design(0 To 4) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Analyte", "nvarchar", 50
FillDesignL 2, "Expiry", "datetime"
FillDesignL 3, "LotNumber", "nvarchar", 50
FillDesignL 4, "EntryDateTime", "datetime"

If IsTableInDB("LotNumbers") = False Then  'There is no table  in database
  CreateTable "LotNumbers"
Else
  DoTableAnalysis "LotNumbers"
End If

End Sub

Public Sub DefineFaecalRequests()

ReDim Design(0 To 13) As FieldDefs

FillDesignL 0, "OP", "bit"
FillDesignL 1, "Rota", "bit"
FillDesignL 2, "Adeno", "bit"
FillDesignL 3, "EPC", "bit"
FillDesignL 4, "Culture", "bit"
FillDesignL 5, "CDiff", "bit"
FillDesignL 6, "ToxinA", "bit"
FillDesignL 7, "Coli0157", "bit"
FillDesignL 8, "OB0", "bit"
FillDesignL 9, "OB1", "bit"
FillDesignL 10, "OB2", "bit"
FillDesignL 11, "ssScreen", "nvarchar", 50
FillDesignL 12, "SampleID", "numeric", , True
FillDesignL 13, "CultureDate", "datetime"

If IsTableInDB("FaecalRequests") = False Then 'There is no table  in database
  CreateTable "FaecalRequests"
Else
  DoTableAnalysis "FaecalRequests"
End If

End Sub

Public Sub DefineFaeces()

ReDim Design(0 To 39) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "pcDone", "bit"
FillDesignL 2, "pc", "nvarchar", 50
FillDesignL 3, "SeleniteDone", "bit"
FillDesignL 4, "Selenite", "nvarchar", 50
FillDesignL 5, "Screen", "int"
FillDesignL 6, "Purity", "int"
FillDesignL 7, "API", "int"
FillDesignL 8, "APICode0", "nvarchar", 50
FillDesignL 9, "APICode1", "nvarchar", 50
FillDesignL 10, "APICode2", "nvarchar", 50
FillDesignL 11, "APICode3", "nvarchar", 50
FillDesignL 12, "APICode4", "nvarchar", 50
FillDesignL 13, "APIName0", "ntext"
FillDesignL 14, "APIName1", "ntext"
FillDesignL 15, "APIName2", "ntext"
FillDesignL 16, "APIName3", "ntext"
FillDesignL 17, "APIName4", "ntext"
FillDesignL 18, "Lact", "nvarchar", 50
FillDesignL 19, "Urea", "nvarchar", 50
FillDesignL 20, "Camp", "nvarchar", 50
FillDesignL 21, "CampLatex", "nvarchar", 50
FillDesignL 22, "Gram", "ntext"
FillDesignL 23, "CampCulture", "ntext"
FillDesignL 24, "PC0157", "nvarchar", 50
FillDesignL 25, "PC0157Latex", "nvarchar", 50
FillDesignL 26, "PC0157Report", "ntext"
FillDesignL 27, "EPC", "nvarchar", 50
FillDesignL 28, "chkOccult", "int"
FillDesignL 29, "Occult", "nvarchar", 50
FillDesignL 30, "Rota", "nvarchar", 50
FillDesignL 31, "Adeno", "nvarchar", 50
FillDesignL 32, "ToxinAL", "nvarchar", 50
FillDesignL 33, "ToxinATA", "nvarchar", 50
FillDesignL 34, "Aus", "nvarchar", 50
FillDesignL 35, "OP0", "ntext"
FillDesignL 36, "OP1", "ntext"
FillDesignL 37, "OP2", "ntext"
FillDesignL 38, "Valid", "bit"
FillDesignL 39, "CRP", "nvarchar", 50

If IsTableInDB("Faeces") = False Then 'There is no table  in database
  CreateTable "Faeces"
Else
  DoTableAnalysis "Faeces"
End If

End Sub


Public Sub DefineMasks()

ReDim Design(0 To 8) As FieldDefs
Dim n As Integer
Dim pTable As String

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "H", "bit", , True
FillDesignL 2, "S", "bit", , True
FillDesignL 3, "L", "bit", , True
FillDesignL 4, "O", "bit", , True
FillDesignL 5, "G", "bit", , True
FillDesignL 6, "J", "bit", , True
FillDesignL 7, "LIH", "int"
FillDesignL 8, "RunDate", "datetime", , True

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "Masks") = False Then 'There is no table  in database
    CreateTable pTable & "Masks"
  Else
    DoTableAnalysis pTable & "Masks"
  End If
Next

End Sub

Public Sub DefineFastings()

ReDim Design(0 To 4) As FieldDefs

FillDesignL 0, "TestName", "nvarchar", 50
FillDesignL 1, "FastingLow", "nvarchar", 50
FillDesignL 2, "FastingHigh", "nvarchar", 50
FillDesignL 3, "FastingText", "nvarchar", 50
FillDesignL 4, "Code", "nvarchar", 50

If IsTableInDB("Fastings") = False Then  'There is no table  in database
  CreateTable "Fastings"
Else
  DoTableAnalysis "Fastings"
End If

End Sub


Public Sub DefineFilmReports()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Report", "ntext"
FillDesignL 2, "Valid", "int"

If IsTableInDB("FilmReports") = False Then  'There is no table  in database
  CreateTable "FilmReports"
Else
  DoTableAnalysis "FilmReports"
End If

End Sub

Public Sub DefineForcedABReport()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "ABName", "nvarchar", 50
FillDesignL 2, "Report", "bit"
FillDesignL 3, "Index", "int"

If IsTableInDB("ForcedABReport") = False Then  'There is no table  in database
  CreateTable "ForcedABReport"
Else
  DoTableAnalysis "ForcedABReport"
End If

End Sub


Public Sub DefineGPs()

ReDim Design(0 To 14) As FieldDefs

FillDesignL 0, "Text", "nvarchar", 50
FillDesignL 1, "Code", "nvarchar", 50
FillDesignL 2, "Addr0", "nvarchar", 50
FillDesignL 3, "Addr1", "nvarchar", 50
FillDesignL 4, "InUse", "bit"
FillDesignL 5, "Title", "nvarchar", 50
FillDesignL 6, "ForeName", "nvarchar", 50
FillDesignL 7, "SurName", "nvarchar", 50
FillDesignL 8, "Phone", "nvarchar", 50
FillDesignL 9, "FAX", "nvarchar", 50
FillDesignL 10, "Practice", "nvarchar", 50
FillDesignL 11, "Compiled", "bit"
FillDesignL 12, "HospitalCode", "nvarchar", 50
FillDesignL 13, "ListOrder", "int"
FillDesignL 14, "HealthLink", "bit"

If IsTableInDB("GPs") = False Then  'There is no table  in database
  CreateTable "GPs"
Else
  DoTableAnalysis "GPs"
End If

End Sub

Public Sub DefineHaemCondition()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Chart", "nvarchar", 50
FillDesignL 1, "Condition", "ntext"

If IsTableInDB("HaemCondition") = False Then  'There is no table  in database
  CreateTable "HaemCondition"
Else
  DoTableAnalysis "HaemCondition"
End If

End Sub


Public Sub DefineRequests()

ReDim Design(0 To 6) As FieldDefs
Dim n As Integer
Dim pTable As String

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "Code", "nvarchar", 50, True
FillDesignL 2, "Programmed", "int"
FillDesignL 3, "DateTime", "datetime", , True
FillDesignL 4, "SampleType", "nvarchar", 50
FillDesignL 5, "AnalyserID", "nvarchar", 50
FillDesignL 6, "Units", "nvarchar", 50

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "Requests") = False Then 'There is no table  in database
    CreateTable pTable & "Requests"
  Else
    DoTableAnalysis pTable & "Requests"
  End If
Next

End Sub

Public Sub DefineConsultantList()

ReDim Design(0 To 0) As FieldDefs

FillDesignL 0, "SampleId", "numeric", , True

If IsTableInDB("ConsultantList") = False Then  'There is no table  in database
  CreateTable "ConsultantList"
Else
  DoTableAnalysis "ConsultantList"
End If

End Sub

Public Sub DefineControls()

ReDim Design(0 To 6) As FieldDefs
Dim n As Integer
Dim pTable As String

FillDesignL 0, "CName", "nvarchar", 50
FillDesignL 1, "Parameter", "int"
FillDesignL 2, "Value", "real"
FillDesignL 3, "DateTime", "datetime"
FillDesignL 4, "ControlName", "nvarchar", 50
FillDesignL 5, "SD1", "real"
FillDesignL 6, "Mean", "real"

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "Controls") = False Then 'There is no table  in database
    CreateTable pTable & "Controls"
  Else
    DoTableAnalysis pTable & "Controls"
  End If
Next

End Sub

Public Sub DefineCulture()

ReDim Design(0 To 21) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "PCDone", "bit"
FillDesignL 2, "PC", "nvarchar", 50
FillDesignL 3, "SeleniteDone", "bit"
FillDesignL 4, "Selenite", "nvarchar", 50
FillDesignL 5, "Screen", "int"
FillDesignL 6, "LactNeg", "int"
FillDesignL 7, "LactPos", "int"
FillDesignL 8, "UreaNeg", "int"
FillDesignL 9, "UreaPos", "int"
FillDesignL 10, "Purity", "int"
FillDesignL 11, "API", "int"
FillDesignL 12, "APIC1", "nvarchar", 50
FillDesignL 13, "APIC2", "nvarchar", 50
FillDesignL 14, "APIC3", "nvarchar", 50
FillDesignL 15, "APIC4", "nvarchar", 50
FillDesignL 16, "APIC5", "nvarchar", 50
FillDesignL 17, "apiT1", "nvarchar", 50
FillDesignL 18, "apiT2", "nvarchar", 50
FillDesignL 19, "apiT3", "nvarchar", 50
FillDesignL 20, "apiT4", "nvarchar", 50
FillDesignL 21, "apiT5", "nvarchar", 50

If IsTableInDB("Culture") = False Then  'There is no table  in database
  CreateTable "Culture"
Else
  DoTableAnalysis "Culture"
End If

End Sub

Public Sub DefineMasksArc()

ReDim Design(0 To 8) As FieldDefs
Dim n As Integer
Dim pTable As String

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "H", "bit", , True
FillDesignL 2, "S", "bit", , True
FillDesignL 3, "L", "bit", , True
FillDesignL 4, "O", "bit", , True
FillDesignL 5, "G", "bit", , True
FillDesignL 6, "J", "bit", , True
FillDesignL 7, "LIH", "int"
FillDesignL 8, "RunDate", "datetime", , True

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "MasksArc") = False Then 'There is no table  in database
    CreateTable pTable & "MasksArc"
  Else
    DoTableAnalysis pTable & "MasksArc"
  End If
Next

End Sub

Public Sub DefineRepeats()

ReDim Design(0 To 18) As FieldDefs
Dim n As Integer
Dim pTable As String

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "Code", "nvarchar", 50, True
FillDesignL 2, "Result", "nvarchar", 50
FillDesignL 3, "Valid", "tinyint"
FillDesignL 4, "Printed", "tinyint"
FillDesignL 5, "RunTime", "datetime"
FillDesignL 6, "RunDate", "datetime"
FillDesignL 7, "Operator", "nvarchar", 50
FillDesignL 8, "Flags", "nvarchar", 50
FillDesignL 9, "Units", "nvarchar", 50
FillDesignL 10, "SampleType", "nvarchar", 50
FillDesignL 11, "Analyser", "nvarchar", 50
FillDesignL 12, "Faxed", "tinyint"
FillDesignL 13, "NOPAS", "nvarchar", 50
FillDesignL 14, "Authorised", "tinyint"
FillDesignL 15, "Comment", "nvarchar", 100
FillDesignL 16, "PC", "nvarchar", 50
FillDesignL 17, "Healthlink", "tinyint"
FillDesignL 18, "DefIndex", "int"

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "Repeats") = False Then 'There is no table  in database
    CreateTable pTable & "Repeats"
  Else
    DoTableAnalysis pTable & "Repeats"
  End If
Next

End Sub


Public Sub DefineResultsArc()

ReDim Design(0 To 18) As FieldDefs
Dim n As Integer
Dim pTable As String

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "Code", "nvarchar", 50, True
FillDesignL 2, "Result", "nvarchar", 50
FillDesignL 3, "Valid", "tinyint"
FillDesignL 4, "Printed", "tinyint"
FillDesignL 5, "RunTime", "datetime"
FillDesignL 6, "RunDate", "datetime"
FillDesignL 7, "Operator", "nvarchar", 50
FillDesignL 8, "Flags", "nvarchar", 50
FillDesignL 9, "Units", "nvarchar", 50
FillDesignL 10, "SampleType", "nvarchar", 50
FillDesignL 11, "Analyser", "nvarchar", 50
FillDesignL 12, "Faxed", "tinyint"
FillDesignL 13, "NOPAS", "nvarchar", 50
FillDesignL 14, "Authorised", "tinyint"
FillDesignL 15, "Comment", "nvarchar", 100
FillDesignL 16, "PC", "nvarchar", 50
FillDesignL 17, "Healthlink", "tinyint"
FillDesignL 18, "DefIndex", "int"

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "ResultsArc") = False Then 'There is no table  in database
    CreateTable pTable & "ResultsArc"
  Else
    DoTableAnalysis pTable & "ResultsArc"
  End If
Next

End Sub

Public Sub DefineRepeatsArc()

ReDim Design(0 To 18) As FieldDefs
Dim n As Integer
Dim pTable As String

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "Code", "nvarchar", 50, True
FillDesignL 2, "Result", "nvarchar", 50
FillDesignL 3, "Valid", "tinyint"
FillDesignL 4, "Printed", "tinyint"
FillDesignL 5, "RunTime", "datetime"
FillDesignL 6, "RunDate", "datetime"
FillDesignL 7, "Operator", "nvarchar", 50
FillDesignL 8, "Flags", "nvarchar", 50
FillDesignL 9, "Units", "nvarchar", 50
FillDesignL 10, "SampleType", "nvarchar", 50
FillDesignL 11, "Analyser", "nvarchar", 50
FillDesignL 12, "Faxed", "tinyint"
FillDesignL 13, "NOPAS", "nvarchar", 50
FillDesignL 14, "Authorised", "tinyint"
FillDesignL 15, "Comment", "nvarchar", 100
FillDesignL 16, "PC", "nvarchar", 50
FillDesignL 17, "Healthlink", "tinyint"
FillDesignL 18, "DefIndex", "int"

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "RepeatsArc") = False Then 'There is no table  in database
    CreateTable pTable & "RepeatsArc"
  Else
    DoTableAnalysis pTable & "RepeatsArc"
  End If
Next

End Sub


Public Sub DefineTestDefinitions()

Dim n As Integer
Dim pTable As String
ReDim Design(0 To 48) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50, True
FillDesignL 1, "SampleType", "nvarchar", 50
FillDesignL 2, "AgeFromDays", "int", , True
FillDesignL 3, "AgeToDays", "int", , True
FillDesignL 4, "DefIndex", "int", , True
FillDesignL 5, "KnownToAnalyser", "tinyint", , False
FillDesignL 6, "InUse", "tinyint", , True
FillDesignL 7, "LongName", "nvarchar", 50
FillDesignL 8, "ShortName", "nvarchar", 50
FillDesignL 9, "DoDelta", "bit"
FillDesignL 10, "DeltaLimit", "real"
FillDesignL 11, "PrintPriority", "smallint"
FillDesignL 12, "DP", "tinyint"
FillDesignL 13, "BarCode", "nvarchar", 50
FillDesignL 14, "Units", "nvarchar", 50
FillDesignL 15, "H", "bit", , True
FillDesignL 16, "S", "bit", , True
FillDesignL 17, "L", "bit", , True
FillDesignL 18, "O", "bit", , True
FillDesignL 19, "G", "bit", , True
FillDesignL 20, "J", "bit", , True
FillDesignL 21, "Category", "nvarchar", 50
FillDesignL 22, "Printable", "bit"
FillDesignL 23, "PlausibleLow", "real"
FillDesignL 24, "PlausibleHigh", "real"
FillDesignL 25, "MaleLow", "real"
FillDesignL 26, "MaleHigh", "real"
FillDesignL 27, "FemaleLow", "real"
FillDesignL 28, "FemaleHigh", "real"
FillDesignL 29, "FlagMaleLow", "real"
FillDesignL 30, "FlagMaleHigh", "real"
FillDesignL 31, "FlagFemaleLow", "real"
FillDesignL 32, "FlagFemaleHigh", "real"
FillDesignL 33, "LControlLow", "real"
FillDesignL 34, "LControlHigh", "real"
FillDesignL 35, "NControlLow", "real"
FillDesignL 36, "NControlHigh", "real"
FillDesignL 37, "HControlLow", "real"
FillDesignL 38, "HControlHigh", "real"
FillDesignL 39, "AutoValLow", "real"
FillDesignL 40, "AutoValHigh", "real"
FillDesignL 41, "Hospital", "nvarchar", 50
FillDesignL 42, "Analyser", "nvarchar", 50
FillDesignL 43, "ImmunoCode", "nvarchar", 50
FillDesignL 44, "SplitList", "int"
FillDesignL 45, "EOD", "bit"
FillDesignL 46, "LIH", "int"
FillDesignL 47, "ViewOnWard", "int"
FillDesignL 48, "PrnRR", "int"

For n = 1 To 6
  pTable = Choose(n, "Bio", "Imm", "End", "Haem", "BGA", "Coag")
  If IsTableInDB(pTable & "TestDefinitions") = False Then 'There is no table  in database
    CreateTable pTable & "TestDefinitions"
  Else
    DoTableAnalysis pTable & "TestDefinitions"
  End If
Next

End Sub




Public Sub DefineBactIsolateDefinitions()

ReDim Design(0 To 4) As FieldDefs

FillDesignL 0, "SampleCode", "nvarchar", 50
FillDesignL 1, "RequestCode", "nvarchar", 50
FillDesignL 2, "IsolateCode", "nvarchar", 50
FillDesignL 3, "ListOrder", "smallint"
FillDesignL 4, "InUse", "smallint"

If IsTableInDB("BactIsolateDefinitions") = False Then 'There is no table  in database
  CreateTable "BactIsolateDefinitions"
Else
  DoTableAnalysis "BactIsolateDefinitions"
End If

End Sub

Public Sub DefineBiochemistryQC()

ReDim Design(0 To 7) As FieldDefs

FillDesignL 0, "ControlName", "nvarchar", 20
FillDesignL 1, "Code", "nvarchar", 4
FillDesignL 2, "result", "nvarchar", 9
FillDesignL 3, "RunTime", "datetime"
FillDesignL 4, "RunDate", "datetime"
FillDesignL 5, "Units", "nvarchar", "15 "
FillDesignL 6, "SampleType", "nvarchar", 2
FillDesignL 7, "AliasName", "nvarchar", 20

If IsTableInDB("BiochemistryQC") = False Then 'There is no table  in database
  CreateTable "BiochemistryQC"
Else
  DoTableAnalysis "BiochemistryQC"
End If

End Sub

Public Sub DefineBioFlags()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "SampleID", "numeric"
FillDesignL 1, "Flags", "ntext"
FillDesignL 2, "DateTime", "datetime"

If IsTableInDB("BioFlags") = False Then 'There is no table  in database
  CreateTable "BioFlags"
Else
  DoTableAnalysis "BioFlags"
End If

End Sub

Public Sub DefineBioFlagsRep()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "SampleID", "numeric"
FillDesignL 1, "Flags", "ntext"
FillDesignL 2, "DateTime", "datetime"

If IsTableInDB("BioFlagsRep") = False Then 'There is no table  in database
  CreateTable "BioFlagsRep"
Else
  DoTableAnalysis "BioFlagsRep"
End If

End Sub

Public Sub DefineBioQCDefs()

ReDim Design(0 To 6) As FieldDefs

FillDesignL 0, "ControlName", "nvarchar", 50
FillDesignL 1, "ParameterName", "nvarchar", 50
FillDesignL 2, "High", "float"
FillDesignL 3, "Low", "float"
FillDesignL 4, "Mean", "float"
FillDesignL 5, "SD", "float"
FillDesignL 6, "AliasName", "nvarchar", 50

If IsTableInDB("BioQCDefs") = False Then 'There is no table  in database
  CreateTable "BioQCDefs"
Else
  DoTableAnalysis "BioQCDefs"
End If

End Sub
Public Sub DefineClinDetails()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "cdText", "nvarchar", 50
FillDesignL 2, "ListOrder", "int"

If IsTableInDB("ClinDetails") = False Then 'There is no table  in database
  CreateTable "ClinDetails"
Else
  DoTableAnalysis "ClinDetails"
End If

End Sub

Public Sub DefineComments()

ReDim Design(0 To 15) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Demographic", "ntext"
FillDesignL 2, "Biochemistry", "ntext"
FillDesignL 3, "Haematology", "ntext"
FillDesignL 4, "Coagulation", "ntext"
FillDesignL 5, "Semen", "ntext"
FillDesignL 6, "MicroCS", "ntext"
FillDesignL 7, "BloodGas", "ntext"
FillDesignL 8, "MicroGeneral", "ntext"
FillDesignL 9, "MicroIdent", "ntext"
FillDesignL 10, "Immunology", "ntext"
FillDesignL 11, "MicroConsultant", "ntext"
FillDesignL 12, "Film", "ntext"
FillDesignL 13, "Endocrinology", "ntext"
FillDesignL 14, "Histology", "ntext"
FillDesignL 15, "Cytology", "ntext"

If IsTableInDB("Comments") = False Then 'There is no table  in database
  CreateTable "Comments"
Else
  DoTableAnalysis "Comments"
End If

End Sub

Public Sub DefineCommentsArc()

ReDim Design(0 To 15) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Demographic", "ntext"
FillDesignL 2, "Biochemistry", "ntext"
FillDesignL 3, "Haematology", "ntext"
FillDesignL 4, "Coagulation", "ntext"
FillDesignL 5, "Semen", "ntext"
FillDesignL 6, "MicroCS", "ntext"
FillDesignL 7, "BloodGas", "ntext"
FillDesignL 8, "MicroGeneral", "ntext"
FillDesignL 9, "MicroIdent", "ntext"
FillDesignL 10, "Immunology", "ntext"
FillDesignL 11, "MicroConsultant", "ntext"
FillDesignL 12, "Film", "ntext"
FillDesignL 13, "Endocrinology", "ntext"
FillDesignL 14, "Histology", "ntext"
FillDesignL 15, "Cytology", "ntext"

If IsTableInDB("CommentsArc") = False Then 'There is no table  in database
  CreateTable "CommentsArc"
Else
  DoTableAnalysis "CommentsArc"
End If

End Sub

Public Sub DefineCytoResults()

ReDim Design(0 To 7) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "CytoComment", "nvarchar", 50
FillDesignL 2, "NatureOfSpecimen", "nvarchar", 50
FillDesignL 3, "NatureOfSpecimenB", "nvarchar", 50
FillDesignL 4, "NatureOfSpecimenC", "nvarchar", 50
FillDesignL 5, "NatureOfSpecimenD", "nvarchar", 50
FillDesignL 6, "CytoReport", "nvarchar", 50
FillDesignL 7, "HYear", "nvarchar", 50

If IsTableInDB("CytoResults") = False Then 'There is no table  in database
  CreateTable "CytoResults"
Else
  DoTableAnalysis "CytoResults"
End If

End Sub

Public Sub DefineDemographics()

ReDim Design(0 To 50) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Chart", "nvarchar", 50
FillDesignL 2, "PatName", "nvarchar", 50
FillDesignL 3, "Age", "nvarchar", 50
FillDesignL 4, "Sex", "nvarchar", 50
FillDesignL 5, "TimeTaken", "datetime"
FillDesignL 6, "Source", "nvarchar", 50
FillDesignL 7, "RunDate", "datetime"
FillDesignL 8, "DoB", "datetime"
FillDesignL 9, "Addr0", "nvarchar", 50
FillDesignL 10, "Addr1", "nvarchar", 50
FillDesignL 11, "Ward", "nvarchar", 50
FillDesignL 12, "Clinician", "nvarchar", 50
FillDesignL 13, "GP", "nvarchar", 50
FillDesignL 14, "SampleDate", "datetime"
FillDesignL 15, "ClDetails", "ntext"
FillDesignL 16, "Hospital", "nvarchar", 50
FillDesignL 17, "RooH", "bit"
FillDesignL 18, "FAXed", "bit"
FillDesignL 19, "Fasting", "bit"
FillDesignL 20, "OnWarfarin", "bit"
FillDesignL 21, "DateTimeDemographics", "datetime"
FillDesignL 22, "DateTimeHaemPrinted", "datetime"
FillDesignL 23, "DateTimeBioPrinted", "datetime"
FillDesignL 24, "DateTimeCoagPrinted", "datetime"
FillDesignL 25, "Pregnant", "bit"
FillDesignL 26, "AandE", "nvarchar", 50
FillDesignL 27, "NOPAS", "nvarchar", 50
FillDesignL 28, "RecDate", "datetime"
FillDesignL 29, "RecordDateTime", "datetime"
FillDesignL 30, "Category", "nvarchar", 50
FillDesignL 31, "HistoValid", "bit"
FillDesignL 32, "CytoValid", "bit"
FillDesignL 33, "Mrn", "nvarchar", 50
FillDesignL 34, "Username", "nvarchar", 50
FillDesignL 35, "Urgent", "int"
FillDesignL 36, "Valid", "bit"
FillDesignL 37, "HYear", "nvarchar", 50
FillDesignL 38, "SentToEMedRenal", "int"
FillDesignL 39, "SurName", "nvarchar", 50
FillDesignL 40, "ForeName", "nvarchar", 50
FillDesignL 41, "NameOfMother", "nvarchar", 50
FillDesignL 42, "NameOfFather", "nvarchar", 50
FillDesignL 43, "ClinicType", "nvarchar", 50
FillDesignL 44, "BreastFed", "int"
FillDesignL 45, "BreastFedEnd", "datetime"
FillDesignL 46, "Race", "nvarchar", 50
FillDesignL 47, "Profession", "nvarchar", 50
FillDesignL 48, "WhereBorn", "nvarchar", 50
FillDesignL 49, "Phone", "nvarchar", 50
FillDesignL 50, "HealthFacility", "nvarchar", 50

If IsTableInDB("Demographics") = False Then 'There is no table  in database
  CreateTable "Demographics"
Else
  DoTableAnalysis "Demographics"
End If

End Sub

Public Sub IndexDemographics()

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "SampleID"
Design(0).DirectionASC = True
CreateIndex "Demographics", "SampleID", Design(), True, True

ReDim Design(0 To 1) As FieldDefs
Design(0).ColumnName = "SurName"
Design(0).DirectionASC = True
Design(1).ColumnName = "ForeName"
Design(1).DirectionASC = True
CreateIndex "Demographics", "FullName", Design(), False, False

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "DoB"
Design(0).DirectionASC = True
CreateIndex "Demographics", "DoB", Design(), False, False

End Sub

Public Sub IndexDemographicsArc()

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "SampleID"
Design(0).DirectionASC = True
CreateIndex "DemographicsArc", "SampleID", Design(), False, False

ReDim Design(0 To 1) As FieldDefs
Design(0).ColumnName = "SurName"
Design(0).DirectionASC = True
Design(1).ColumnName = "ForeName"
Design(1).DirectionASC = True
CreateIndex "DemographicsArc", "FullName", Design(), False, False

End Sub

Public Sub IndexResults()

ReDim Design(0 To 1) As FieldDefs
Design(0).ColumnName = "SampleID"
Design(0).DirectionASC = True
Design(1).ColumnName = "Code"
Design(1).DirectionASC = True
CreateIndex "BGAResults", "SampleIDCode", Design(), True, False
CreateIndex "BioResults", "SampleIDCode", Design(), True, False
CreateIndex "CoagResults", "SampleIDCode", Design(), True, False
CreateIndex "EndResults", "SampleIDCode", Design(), True, False
CreateIndex "HaemResults", "SampleIDCode", Design(), True, False
CreateIndex "ImmResults", "SampleIDCode", Design(), True, False

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "RunDate"
Design(0).DirectionASC = True
CreateIndex "BGAResults", "RunDate", Design(), False, False
CreateIndex "BioResults", "RunDate", Design(), False, False
CreateIndex "CoagResults", "RunDate", Design(), False, False
CreateIndex "EndResults", "RunDate", Design(), False, False
CreateIndex "HaemResults", "RunDate", Design(), False, False
CreateIndex "ImmResults", "RunDate", Design(), False, False

End Sub

Public Sub IndexRepeats()

ReDim Design(0 To 1) As FieldDefs
Design(0).ColumnName = "SampleID"
Design(0).DirectionASC = True
Design(1).ColumnName = "Code"
Design(1).DirectionASC = True
CreateIndex "BGARepeats", "SampleIDCode", Design(), False, False
CreateIndex "BioRepeats", "SampleIDCode", Design(), False, False
CreateIndex "CoagRepeats", "SampleIDCode", Design(), False, False
CreateIndex "EndRepeats", "SampleIDCode", Design(), False, False
CreateIndex "HaemRepeats", "SampleIDCode", Design(), False, False
CreateIndex "ImmRepeats", "SampleIDCode", Design(), False, False

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "RunDate"
Design(0).DirectionASC = True
CreateIndex "BGARepeats", "RunDate", Design(), False, False
CreateIndex "BioRepeats", "RunDate", Design(), False, False
CreateIndex "CoagRepeats", "RunDate", Design(), False, False
CreateIndex "EndRepeats", "RunDate", Design(), False, False
CreateIndex "HaemRepeats", "RunDate", Design(), False, False
CreateIndex "ImmRepeats", "RunDate", Design(), False, False

End Sub

Public Sub IndexErrorLog()

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "ModuleName"
Design(0).DirectionASC = True
CreateIndex "ErrorLog", "ModuleName", Design(), False, False

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "ProcedureName"
Design(0).DirectionASC = True
CreateIndex "ErrorLog", "ProcedureName", Design(), False, False

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "DateTime"
Design(0).DirectionASC = False
CreateIndex "ErrorLog", "DateTime", Design(), False, False

End Sub
Public Sub IndexSexNames()

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "Name"
Design(0).DirectionASC = True
CreateIndex "SexNames", "Name", Design(), True, False

End Sub

Public Sub IndexEventLog()

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "DateTime"
Design(0).DirectionASC = False
CreateIndex "EventLog", "DateTime", Design(), False, False

End Sub

Public Sub IndexComments()

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "SampleID"
Design(0).DirectionASC = True
CreateIndex "Comments", "SampleID", Design(), True, False

End Sub
Public Sub IndexCommentsArc()

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "SampleID"
Design(0).DirectionASC = True
CreateIndex "CommentsArc", "SampleID", Design(), False, False

End Sub

Public Sub IndexRepeatsArc()

ReDim Design(0 To 1) As FieldDefs
Design(0).ColumnName = "SampleID"
Design(0).DirectionASC = True
Design(1).ColumnName = "Code"
Design(1).DirectionASC = True
CreateIndex "BGARepeatsArc", "SampleIDCode", Design(), False, False
CreateIndex "BioRepeatsArc", "SampleIDCode", Design(), False, False
CreateIndex "CoagRepeatsArc", "SampleIDCode", Design(), False, False
CreateIndex "EndRepeatsArc", "SampleIDCode", Design(), False, False
CreateIndex "HaemRepeatsArc", "SampleIDCode", Design(), False, False
CreateIndex "ImmRepeatsArc", "SampleIDCode", Design(), False, False

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "RunDate"
Design(0).DirectionASC = True
CreateIndex "BGARepeatsArc", "RunDate", Design(), False, False
CreateIndex "BioRepeatsArc", "RunDate", Design(), False, False
CreateIndex "CoagRepeatsArc", "RunDate", Design(), False, False
CreateIndex "EndRepeatsArc", "RunDate", Design(), False, False
CreateIndex "HaemRepeatsArc", "RunDate", Design(), False, False
CreateIndex "ImmRepeatsArc", "RunDate", Design(), False, False

End Sub


Public Sub IndexResultsArc()

ReDim Design(0 To 1) As FieldDefs
Design(0).ColumnName = "SampleID"
Design(0).DirectionASC = True
Design(1).ColumnName = "Code"
Design(1).DirectionASC = True
CreateIndex "BGAResultsArc", "SampleIDCode", Design(), False, False
CreateIndex "BioResultsArc", "SampleIDCode", Design(), False, False
CreateIndex "CoagResultsArc", "SampleIDCode", Design(), False, False
CreateIndex "EndResultsArc", "SampleIDCode", Design(), False, False
CreateIndex "HaemEndResultsArc", "SampleIDCode", Design(), False, False
CreateIndex "ImmResultsArc", "SampleIDCode", Design(), False, False

ReDim Design(0 To 0) As FieldDefs
Design(0).ColumnName = "RunDate"
Design(0).DirectionASC = True
CreateIndex "BGAResultsArc", "RunDate", Design(), False, False
CreateIndex "BioResultsArc", "RunDate", Design(), False, False
CreateIndex "CoagResultsArc", "RunDate", Design(), False, False
CreateIndex "EndResultsArc", "RunDate", Design(), False, False
CreateIndex "HaemResultsArc", "RunDate", Design(), False, False
CreateIndex "ImmResultsArc", "RunDate", Design(), False, False

End Sub

Public Sub DefineTrackBatchNumbers()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "DistrictBatchNumber", "nvarchar", 50
FillDesignL 2, "ProvincialBatchNumber", "nvarchar", 50
FillDesignL 3, "GUID", "nvarchar", 50

If IsTableInDB("TrackBatchNumbers") = False Then 'There is no table  in database
  CreateTable "TrackBatchNumbers"
Else
  DoTableAnalysis "TrackBatchNumbers"
End If

End Sub

Public Sub DefineTrackMessage()

ReDim Design(0 To 42) As FieldDefs

FillDesignL 0, "MessageType", "nvarchar", 50
FillDesignL 1, "GUID", "nvarchar", 50, True
FillDesignL 2, "NID", "nvarchar", 50
FillDesignL 3, "Surname", "nvarchar", 50
FillDesignL 4, "Forename", "nvarchar", 50
FillDesignL 5, "NameOfMother", "nvarchar", 50
FillDesignL 6, "NameOfFather", "nvarchar", 50
FillDesignL 7, "Race", "nvarchar", 50
FillDesignL 8, "Profession", "nvarchar", 50
FillDesignL 9, "WhereBorn", "nvarchar", 50
FillDesignL 10, "Address0", "nvarchar", 50
FillDesignL 11, "Address1", "nvarchar", 50
FillDesignL 12, "Phone", "nvarchar", 50
FillDesignL 13, "Sex", "nvarchar", 50
FillDesignL 14, "DoB", "datetime"
FillDesignL 15, "Age", "nvarchar", 50
FillDesignL 16, "BreastFed", "int"
FillDesignL 17, "HealthFacility", "nvarchar", 50
FillDesignL 18, "Province", "nvarchar", 50
FillDesignL 19, "District", "nvarchar", 50
FillDesignL 20, "ClinicType", "nvarchar", 50
FillDesignL 21, "Clinician", "nvarchar", 50
FillDesignL 22, "CollectionDate", "datetime"
FillDesignL 23, "BreastFedEnd", "datetime"
FillDesignL 24, "DateSent", "datetime"
FillDesignL 25, "DateResultSent", "datetime"
FillDesignL 26, "DateResultReceived", "datetime"
FillDesignL 27, "TestRequired", "nvarchar", 50
FillDesignL 28, "TestRequiredCode", "nvarchar", 50
FillDesignL 29, "Discipline", "nvarchar", 50
FillDesignL 30, "SampleType", "nvarchar", 50
FillDesignL 31, "Result", "nvarchar", 50
FillDesignL 32, "Units", "nvarchar", 50
FillDesignL 33, "Flags", "nvarchar", 50
FillDesignL 34, "NormalRange", "nvarchar", 50
FillDesignL 35, "PreviousPCR", "nvarchar", 50
FillDesignL 36, "PreviousPCRResult", "nvarchar", 50
FillDesignL 37, "DemographicsComment", "nvarchar", 50
FillDesignL 38, "ResultComment", "nvarchar", 50
FillDesignL 39, "DistrictBatchNumber", "nvarchar", 50
FillDesignL 40, "ProvincialBatchNumber", "nvarchar", 50
FillDesignL 41, "SampleID", "numeric"
FillDesignL 42, "AnalysedAt", "nvarchar", 50

If IsTableInDB("TrackMessage") = False Then 'There is no table  in database
  CreateTable "TrackMessage"
Else
  DoTableAnalysis "TrackMessage"
End If

End Sub

Public Sub DefineDemographicsArc()

ReDim Design(0 To 50) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Chart", "nvarchar", 50
FillDesignL 2, "PatName", "nvarchar", 50
FillDesignL 3, "Age", "nvarchar", 50
FillDesignL 4, "Sex", "nvarchar", 50
FillDesignL 5, "TimeTaken", "datetime"
FillDesignL 6, "Source", "nvarchar", 50
FillDesignL 7, "RunDate", "datetime"
FillDesignL 8, "DoB", "datetime"
FillDesignL 9, "Addr0", "nvarchar", 50
FillDesignL 10, "Addr1", "nvarchar", 50
FillDesignL 11, "Ward", "nvarchar", 50
FillDesignL 12, "Clinician", "nvarchar", 50
FillDesignL 13, "GP", "nvarchar", 50
FillDesignL 14, "SampleDate", "datetime"
FillDesignL 15, "ClDetails", "ntext"
FillDesignL 16, "Hospital", "nvarchar", 50
FillDesignL 17, "RooH", "bit"
FillDesignL 18, "FAXed", "bit"
FillDesignL 19, "Fasting", "bit"
FillDesignL 20, "OnWarfarin", "bit"
FillDesignL 21, "DateTimeDemographics", "datetime"
FillDesignL 22, "DateTimeHaemPrinted", "datetime"
FillDesignL 23, "DateTimeBioPrinted", "datetime"
FillDesignL 24, "DateTimeCoagPrinted", "datetime"
FillDesignL 25, "Pregnant", "bit"
FillDesignL 26, "AandE", "nvarchar", 50
FillDesignL 27, "NOPAS", "nvarchar", 50
FillDesignL 28, "RecDate", "datetime"
FillDesignL 29, "RecordDateTime", "datetime"
FillDesignL 30, "Category", "nvarchar", 50
FillDesignL 31, "HistoValid", "bit"
FillDesignL 32, "CytoValid", "bit"
FillDesignL 33, "Mrn", "nvarchar", 50
FillDesignL 34, "Username", "nvarchar", 50
FillDesignL 35, "Urgent", "int"
FillDesignL 36, "Valid", "bit"
FillDesignL 37, "HYear", "nvarchar", 50
FillDesignL 38, "SentToEMedRenal", "int"
FillDesignL 39, "SurName", "nvarchar", 50
FillDesignL 40, "ForeName", "nvarchar", 50
FillDesignL 41, "NameOfMother", "nvarchar", 50
FillDesignL 42, "NameOfFather", "nvarchar", 50
FillDesignL 43, "ClinicType", "nvarchar", 50
FillDesignL 44, "BreastFed", "int"
FillDesignL 45, "BreastFedEnd", "datetime"
FillDesignL 46, "Race", "nvarchar", 50
FillDesignL 47, "Profession", "nvarchar", 50
FillDesignL 48, "WhereBorn", "nvarchar", 50
FillDesignL 49, "Phone", "nvarchar", 50
FillDesignL 50, "HealthFacility", "nvarchar", 50

If IsTableInDB("DemographicsArc") = False Then 'There is no table  in database
  CreateTable "DemographicsArc"
Else
  DoTableAnalysis "DemographicsArc"
End If

End Sub


Public Sub DefineCoagTranslation()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "Units", "nvarchar", 50
FillDesignL 2, "TranCode", "nvarchar", 50

If IsTableInDB("CoagTranslation") = False Then 'There is no table  in database
  CreateTable "CoagTranslation"
Else
  DoTableAnalysis "CoagTranslation"
End If

End Sub

Public Sub DefineCreatinine()

ReDim Design(0 To 11) As FieldDefs

FillDesignL 0, "SerumNumber", "numeric", , True
FillDesignL 1, "UrineNumber", "numeric", , True
FillDesignL 2, "UrineVolume", "nvarchar", 50
FillDesignL 3, "SerumCreat", "nvarchar", 50
FillDesignL 4, "UrineCreat", "nvarchar", 50
FillDesignL 5, "UrineProL", "nvarchar", 50
FillDesignL 6, "UrinePro24hr", "nvarchar", 50
FillDesignL 7, "CCL", "nvarchar", 50
FillDesignL 8, "Name", "nvarchar", 50
FillDesignL 9, "Chart", "nvarchar", 50
FillDesignL 10, "Comment", "nvarchar", 50
FillDesignL 11, "DoB", "datetime"

If IsTableInDB("Creatinine") = False Then 'There is no table  in database
  CreateTable "Creatinine"
Else
  DoTableAnalysis "Creatinine"
End If

End Sub

Public Sub DefineHaemFlags()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Flags", "ntext"
FillDesignL 2, "DateTime", "datetime"

If IsTableInDB("HaemFlags") = False Then 'There is no table  in database
  CreateTable "HaemFlags"
Else
  DoTableAnalysis "HaemFlags"
End If

End Sub
Public Sub DefineHaemFlagsRep()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Flags", "ntext"
FillDesignL 2, "DateTime", "datetime"

If IsTableInDB("HaemFlagsRep") = False Then 'There is no table  in database
  CreateTable "HaemFlagsRep"
Else
  DoTableAnalysis "HaemFlagsRep"
End If

End Sub

Public Sub DefineCytoResultsArc()

ReDim Design(0 To 7) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "CytoComment", "nvarchar", 50
FillDesignL 2, "NatureOfSpecimen", "nvarchar", 50
FillDesignL 3, "NatureOfSpecimenB", "nvarchar", 50
FillDesignL 4, "NatureOfSpecimenC", "nvarchar", 50
FillDesignL 5, "NatureOfSpecimenD", "nvarchar", 50
FillDesignL 6, "CytoReport", "nvarchar", 50
FillDesignL 7, "HYear", "nvarchar", 50

If IsTableInDB("CytoResultsArc") = False Then 'There is no table  in database
  CreateTable "CytoResultsArc"
Else
  DoTableAnalysis "CytoResultsArc"
End If

End Sub

Public Sub DefineBottleType()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "Size", "nvarchar", 50
FillDesignL 1, "Anticoagulant", "nvarchar", 50
FillDesignL 2, "Colour", "nvarchar", 50
FillDesignL 3, "Code", "nvarchar", 50

If IsTableInDB("BottleType") = False Then 'There is no table  in database
  CreateTable "BottleType"
Else
  DoTableAnalysis "BottleType"
End If

End Sub

Public Sub DefineCategorys()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Cat", "nvarchar", 50
FillDesignL 1, "ListOrder", "int"

If IsTableInDB("Categorys") = False Then 'There is no table  in database
  CreateTable "Categorys"
Else
  DoTableAnalysis "Categorys"
End If

End Sub


Public Sub DefineCCMs()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "CCMType", "nvarchar", 50
FillDesignL 2, "Text", "nvarchar", 50
FillDesignL 3, "ListOrder", "int"

If IsTableInDB("CCMs") = False Then 'There is no table  in database
  CreateTable "CCMs"
Else
  DoTableAnalysis "CCMs"
End If

End Sub

Public Sub DefineClinicians()

ReDim Design(0 To 8) As FieldDefs

FillDesignL 0, "Text", "nvarchar", 50
FillDesignL 1, "InUse", "bit"
FillDesignL 2, "HospitalCode", "nvarchar", 50
FillDesignL 3, "Code", "nvarchar", 50
FillDesignL 4, "Ward", "nvarchar", 50
FillDesignL 5, "Title", "nvarchar", 50
FillDesignL 6, "ForeName", "nvarchar", 50
FillDesignL 7, "SurName", "nvarchar", 50
FillDesignL 8, "ListOrder", "int"

If IsTableInDB("Clinicians") = False Then 'There is no table  in database
  CreateTable "Clinicians"
Else
  DoTableAnalysis "Clinicians"
End If

End Sub
Public Sub DefineBactRequestDefinitions()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "SampleCode", "nvarchar", 50
FillDesignL 1, "RequestCode", "nvarchar", 50
FillDesignL 2, "ListOrder", "smallint"
FillDesignL 3, "InUse", "smallint"

If IsTableInDB("BactRequestDefinitions") = False Then 'There is no table  in database
  CreateTable "BactRequestDefinitions"
Else
  DoTableAnalysis "BactRequestDefinitions"
End If

End Sub


Public Sub DefineBactResults()

ReDim Design(0 To 11) As FieldDefs

FillDesignL 0, "SampleId", "numeric", , True
FillDesignL 1, "Type", "nvarchar", 50
FillDesignL 2, "Done", "bit"
FillDesignL 3, "X", "int"
FillDesignL 4, "Y", "int"
FillDesignL 5, "TriplePosition", "smallint"
FillDesignL 6, "IsolateName", "nvarchar", 50
FillDesignL 7, "TestName", "nvarchar", 50
FillDesignL 8, "Result", "nvarchar", 50
FillDesignL 9, "Valid", "int", , True
FillDesignL 10, "IsolateSortOrder", "smallint"
FillDesignL 11, "TestSortOrder", "smallint"

If IsTableInDB("BactResults") = False Then 'There is no table  in database
  CreateTable "BactResults"
Else
  DoTableAnalysis "BactResults"
End If

End Sub

Public Sub DefineBacteriology()

ReDim Design(0 To 18) As FieldDefs

FillDesignL 0, "AntibioticName", "nvarchar", 50
FillDesignL 1, "OrganismGroup", "nvarchar", 50
FillDesignL 2, "Site", "nvarchar", 50
FillDesignL 3, "ListOrder", "int"
FillDesignL 4, "PriSec", "nvarchar", 50
FillDesignL 5, "SampleID", "numeric", , True
FillDesignL 6, "Type", "nvarchar", 50
FillDesignL 7, "Done", "bit"
FillDesignL 8, "X", "int"
FillDesignL 9, "Y", "int"
FillDesignL 10, "TriplePosition", "smallint"
FillDesignL 11, "Result", "nvarchar", 50
FillDesignL 12, "Valid", "int"
FillDesignL 13, "IsolateCode", "nvarchar", 50
FillDesignL 14, "TestCode", "nvarchar", 50
FillDesignL 15, "RequestCode", "nvarchar", 50
FillDesignL 16, "TestSortOrder", "smallint"
FillDesignL 17, "IsolateSortOrder", "smallint"
FillDesignL 18, "SampleCode", "nvarchar", 50

If IsTableInDB("Bacteriology") = False Then 'There is no table  in database
  CreateTable "Bacteriology"
Else
  DoTableAnalysis "Bacteriology"
End If

End Sub

Public Sub DefineErrorLog()

ReDim Design(0 To 8) As FieldDefs

FillDesignL 0, "ModuleName", "nvarchar", 50
FillDesignL 1, "ProcedureName", "nvarchar", 50
FillDesignL 2, "ErrorLineNumber", "int"
FillDesignL 3, "SQLStatement", "ntext"
FillDesignL 4, "ErrorDescription", "ntext"
FillDesignL 5, "DateTime", "datetime", , True
FillDesignL 6, "UserName", "nvarchar", 50
FillDesignL 7, "MachineName", "nvarchar", 50
FillDesignL 8, "EventDesc", "ntext"

If IsTableInDB("ErrorLog") = False Then 'There is no table  in database
  CreateTable "ErrorLog"
Else
  DoTableAnalysis "ErrorLog"
End If

End Sub

Public Sub DefineDisease()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Chart", "nvarchar", 50
FillDesignL 2, "Name", "nvarchar", 50
FillDesignL 3, "LabID", "nvarchar", 50

If IsTableInDB("Disease") = False Then 'There is no table  in database
  CreateTable "Disease"
Else
  DoTableAnalysis "Disease"
End If

End Sub

Public Sub DefineETC()

ReDim Design(0 To 9) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "etc0", "ntext"
FillDesignL 2, "etc1", "ntext"
FillDesignL 3, "etc2", "ntext"
FillDesignL 4, "etc3", "ntext"
FillDesignL 5, "etc4", "ntext"
FillDesignL 6, "etc5", "ntext"
FillDesignL 7, "etc6", "ntext"
FillDesignL 8, "etc7", "ntext"
FillDesignL 9, "etc8", "ntext"

If IsTableInDB("ETC") = False Then 'There is no table  in database
  CreateTable "ETC"
Else
  DoTableAnalysis "ETC"
End If

End Sub

Public Sub DefineDifferentialTitles()

ReDim Design(0 To 24) As FieldDefs

FillDesignL 0, "K0", "nvarchar", 50
FillDesignL 1, "K1", "nvarchar", 50
FillDesignL 2, "K2", "nvarchar", 50
FillDesignL 3, "K3", "nvarchar", 50
FillDesignL 4, "K4", "nvarchar", 50
FillDesignL 5, "K5", "nvarchar", 50
FillDesignL 6, "K6", "nvarchar", 50
FillDesignL 7, "K7", "nvarchar", 50
FillDesignL 8, "K8", "nvarchar", 50
FillDesignL 9, "K9", "nvarchar", 50
FillDesignL 10, "K10", "nvarchar", 50
FillDesignL 11, "K11", "nvarchar", 50
FillDesignL 12, "K12", "nvarchar", 50
FillDesignL 13, "K13", "nvarchar", 50
FillDesignL 14, "K14", "nvarchar", 50
FillDesignL 15, "C5", "nvarchar", 50
FillDesignL 16, "C6", "nvarchar", 50
FillDesignL 17, "C7", "nvarchar", 50
FillDesignL 18, "C8", "nvarchar", 50
FillDesignL 19, "C9", "nvarchar", 50
FillDesignL 20, "C10", "nvarchar", 50
FillDesignL 21, "C11", "nvarchar", 50
FillDesignL 22, "C12", "nvarchar", 50
FillDesignL 23, "C13", "nvarchar", 50
FillDesignL 24, "C14", "nvarchar", 50

If IsTableInDB("DifferentialTitles") = False Then 'There is no table  in database
  CreateTable "DifferentialTitles"
Else
  DoTableAnalysis "DifferentialTitles"
End If

End Sub

Public Sub DefineDifferentials()

ReDim Design(0 To 62) As FieldDefs

FillDesignL 0, "Key0", "nvarchar", 50
FillDesignL 1, "Key1", "nvarchar", 50
FillDesignL 2, "Key2", "nvarchar", 50
FillDesignL 3, "Key3", "nvarchar", 50
FillDesignL 4, "Key4", "nvarchar", 50
FillDesignL 5, "Key5", "nvarchar", 50
FillDesignL 6, "Key6", "nvarchar", 50
FillDesignL 7, "Key7", "nvarchar", 50
FillDesignL 8, "Key8", "nvarchar", 50
FillDesignL 9, "Key9", "nvarchar", 50
FillDesignL 10, "Key10", "nvarchar", 50
FillDesignL 11, "Key11", "nvarchar", 50
FillDesignL 12, "Key12", "nvarchar", 50
FillDesignL 13, "Key13", "nvarchar", 50
FillDesignL 14, "Key14", "nvarchar", 50
FillDesignL 15, "Wording0", "nvarchar", 50
FillDesignL 16, "Wording1", "nvarchar", 50
FillDesignL 17, "Wording2", "nvarchar", 50
FillDesignL 18, "Wording3", "nvarchar", 50
FillDesignL 19, "Wording4", "nvarchar", 50
FillDesignL 20, "Wording5", "nvarchar", 50
FillDesignL 21, "Wording6", "nvarchar", 50
FillDesignL 22, "Wording7", "nvarchar", 50
FillDesignL 23, "Wording8", "nvarchar", 50
FillDesignL 24, "Wording9", "nvarchar", 50
FillDesignL 25, "Wording10", "nvarchar", 50
FillDesignL 26, "Wording11", "nvarchar", 50
FillDesignL 27, "Wording12", "nvarchar", 50
FillDesignL 28, "Wording13", "nvarchar", 50
FillDesignL 29, "Wording14", "nvarchar", 50
FillDesignL 30, "P0", "smallint"
FillDesignL 31, "P1", "smallint"
FillDesignL 32, "P2", "smallint"
FillDesignL 33, "P3", "smallint"
FillDesignL 34, "P4", "smallint"
FillDesignL 35, "P5", "smallint"
FillDesignL 36, "P6", "smallint"
FillDesignL 37, "P7", "smallint"
FillDesignL 38, "P8", "smallint"
FillDesignL 39, "P9", "smallint"
FillDesignL 40, "P10", "smallint"
FillDesignL 41, "P11", "smallint"
FillDesignL 42, "P12", "smallint"
FillDesignL 43, "P13", "smallint"
FillDesignL 44, "P14", "smallint"
FillDesignL 45, "A0", "real"
FillDesignL 46, "A1", "real"
FillDesignL 47, "A2", "real"
FillDesignL 48, "A3", "real"
FillDesignL 49, "A4", "real"
FillDesignL 50, "A5", "real"
FillDesignL 51, "A6", "real"
FillDesignL 52, "A7", "real"
FillDesignL 53, "A8", "real"
FillDesignL 54, "A9", "real"
FillDesignL 55, "A10", "real"
FillDesignL 56, "A11", "real"
FillDesignL 57, "A12", "real"
FillDesignL 58, "A13", "real"
FillDesignL 59, "A14", "real"
FillDesignL 60, "PrnDiff", "bit"
FillDesignL 61, "Operator", "nvarchar", 50
FillDesignL 62, "SampleID", "numeric", , True

If IsTableInDB("Differentials") = False Then 'There is no table  in database
  CreateTable "Differentials"
Else
  DoTableAnalysis "Differentials"
End If

End Sub
Public Sub DefineEventLog()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "Description", "ntext"
FillDesignL 1, "DateTime", "datetime", , True
FillDesignL 2, "UserName", "nvarchar", 50

If IsTableInDB("EventLog") = False Then 'There is no table  in database
  CreateTable "EventLog"
Else
  DoTableAnalysis "EventLog"
End If

End Sub


Public Sub DefineHbA1c()

ReDim Design(0 To 21) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Block", "ntext"
FillDesignL 2, "HbA1", "nvarchar", 50
FillDesignL 3, "HbF", "nvarchar", 50
FillDesignL 4, "Ps1", "nvarchar", 50
FillDesignL 5, "Pa1", "nvarchar", 50
FillDesignL 6, "Pp1", "nvarchar", 50
FillDesignL 7, "Ps2", "nvarchar", 50
FillDesignL 8, "Pa2", "nvarchar", 50
FillDesignL 9, "Pp2", "nvarchar", 50
FillDesignL 10, "Ps3", "nvarchar", 50
FillDesignL 11, "Pa3", "nvarchar", 50
FillDesignL 12, "Pp3", "nvarchar", 50
FillDesignL 13, "Ps4", "nvarchar", 50
FillDesignL 14, "Pa4", "nvarchar", 50
FillDesignL 15, "Pp4", "nvarchar", 50
FillDesignL 16, "Ps5", "nvarchar", 50
FillDesignL 17, "Pa5", "nvarchar", 50
FillDesignL 18, "Pp5", "nvarchar", 50
FillDesignL 19, "Ps6", "nvarchar", 50
FillDesignL 20, "Pa6", "nvarchar", 50
FillDesignL 21, "Pp6", "nvarchar", 50

If IsTableInDB("HbA1c") = False Then 'There is no table  in database
  CreateTable "HbA1c"
Else
  DoTableAnalysis "HbA1c"
End If

End Sub


Public Sub DefineHISErrors()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "DateTime", "datetime"
FillDesignL 1, "Hospital", "nvarchar", 50
FillDesignL 2, "RawData", "ntext"

If IsTableInDB("HISErrors") = False Then 'There is no table  in database
  CreateTable "HISErrors"
Else
  DoTableAnalysis "HISErrors"
End If

End Sub



Public Sub DefineHistoBlock()

ReDim Design(0 To 6) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Specimen", "nvarchar", 50
FillDesignL 2, "Block", "int"
FillDesignL 3, "Pieces", "int"
FillDesignL 4, "Type", "nvarchar", 50
FillDesignL 5, "HYear", "nvarchar", 50
FillDesignL 6, "PiComm", "nvarchar", 50

If IsTableInDB("HistoBlock") = False Then 'There is no table  in database
  CreateTable "HistoBlock"
Else
  DoTableAnalysis "HistoBlock"
End If

End Sub

Public Sub DefineHistoComments()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "Text", "nvarchar", 50

If IsTableInDB("HistoComments") = False Then 'There is no table  in database
  CreateTable "HistoComments"
Else
  DoTableAnalysis "HistoComments"
End If

End Sub

Public Sub DefineHistoResults()

ReDim Design(0 To 9) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "HistoComment", "nvarchar", 50
FillDesignL 2, "NatureOfSpecimen", "nvarchar", 50
FillDesignL 3, "NatureOfSpecimenB", "nvarchar", 50
FillDesignL 4, "NatureOfSpecimenC", "nvarchar", 50
FillDesignL 5, "NatureOfSpecimenD", "nvarchar", 50
FillDesignL 6, "NatureOfSpecimenE", "nvarchar", 50
FillDesignL 7, "NatureOfSpecimenF", "nvarchar", 50
FillDesignL 8, "HistoReport", "ntext"
FillDesignL 9, "HYear", "nvarchar", 50

If IsTableInDB("HistoResults") = False Then 'There is no table  in database
  CreateTable "HistoResults"
Else
  DoTableAnalysis "HistoResults"
End If

End Sub

Public Sub DefineHistoResultsArc()

ReDim Design(0 To 9) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "HistoComment", "nvarchar", 50
FillDesignL 2, "NatureOfSpecimen", "nvarchar", 50
FillDesignL 3, "NatureOfSpecimenB", "nvarchar", 50
FillDesignL 4, "NatureOfSpecimenC", "nvarchar", 50
FillDesignL 5, "NatureOfSpecimenD", "nvarchar", 50
FillDesignL 6, "NatureOfSpecimenE", "nvarchar", 50
FillDesignL 7, "NatureOfSpecimenF", "nvarchar", 50
FillDesignL 8, "HistoReport", "ntext"
FillDesignL 9, "HYear", "nvarchar", 50

If IsTableInDB("HistoResultsArc") = False Then 'There is no table  in database
  CreateTable "HistoResultsArc"
Else
  DoTableAnalysis "HistoResultsArc"
End If

End Sub

Public Sub DefineHistory()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Chart", "nvarchar", 50
FillDesignL 1, "History", "ntext"

If IsTableInDB("History") = False Then 'There is no table  in database
  CreateTable "History"
Else
  DoTableAnalysis "History"
End If

End Sub

Public Sub DefineHistoSpecimen()

ReDim Design(0 To 6) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Specimen", "nvarchar", 50
FillDesignL 2, "Type", "nvarchar", 50
FillDesignL 3, "Blocks", "nvarchar", 50
FillDesignL 4, "Rundate", "datetime"
FillDesignL 5, "Remark", "ntext"
FillDesignL 6, "HYear", "nvarchar", 50

If IsTableInDB("HistoSpecimen") = False Then 'There is no table  in database
  CreateTable "HistoSpecimen"
Else
  DoTableAnalysis "HistoSpecimen"
End If

End Sub


Public Sub DefineHistoStain()

ReDim Design(0 To 7) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Stain", "nvarchar", 50
FillDesignL 2, "Result", "nvarchar", 50
FillDesignL 3, "Block", "int"
FillDesignL 4, "Grid", "int"
FillDesignL 5, "Specimen", "int"
FillDesignL 6, "HYear", "nvarchar", 50
FillDesignL 7, "ResComm", "nvarchar", 50

If IsTableInDB("HistoStain") = False Then 'There is no table  in database
  CreateTable "HistoStain"
Else
  DoTableAnalysis "HistoStain"
End If

End Sub

Public Sub DefineHMRU()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "DateTime", "datetime"
FillDesignL 1, "SampleID", "numeric", , True
FillDesignL 2, "UserCode", "nvarchar", 50

If IsTableInDB("HMRU") = False Then 'There is no table  in database
  CreateTable "HMRU"
Else
  DoTableAnalysis "HMRU"
End If

End Sub


Public Sub DefineMRU()

ReDim Design(0 To 2) As FieldDefs

FillDesignL 0, "DateTime", "datetime"
FillDesignL 1, "SampleID", "numeric", , True
FillDesignL 2, "UserCode", "nvarchar", 50

If IsTableInDB("MRU") = False Then 'There is no table  in database
  CreateTable "MRU"
Else
  DoTableAnalysis "MRU"
End If

End Sub


Public Sub DefineOptions()

ReDim Design(0 To 5) As FieldDefs

FillDesignL 0, "Description", "nvarchar", 50
FillDesignL 1, "Contents", "nvarchar", 50
FillDesignL 2, "UserName", "nvarchar", 50
FillDesignL 3, "ListOrder", "int"
FillDesignL 4, "optType", "nvarchar", 50
FillDesignL 5, "HospitalNumber", "int"

If IsTableInDB("Options") = False Then 'There is no table  in database
  CreateTable "Options"
Else
  DoTableAnalysis "Options"
End If

End Sub

Public Sub DefineOrganisms()

ReDim Design(0 To 5) As FieldDefs

FillDesignL 0, "Name", "nvarchar", 50
FillDesignL 1, "GroupName", "nvarchar", 50
FillDesignL 2, "ListOrder", "int"
FillDesignL 3, "Code", "nvarchar", 50
FillDesignL 4, "ShortName", "nvarchar", 50
FillDesignL 5, "ReportName", "nvarchar", 50

If IsTableInDB("Organisms") = False Then 'There is no table  in database
  CreateTable "Organisms"
Else
  DoTableAnalysis "Organisms"
End If

End Sub

Public Sub DefineOP()

ReDim Design(0 To 26) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "AUS", "nvarchar", 50
FillDesignL 2, "ToxinALatex", "nvarchar", 50
FillDesignL 3, "ToxinA", "nvarchar", 50
FillDesignL 4, "ob1", "nvarchar", 50
FillDesignL 5, "ob2", "nvarchar", 50
FillDesignL 6, "ob3", "nvarchar", 50
FillDesignL 7, "EPC1", "nvarchar", 50
FillDesignL 8, "EPC2", "nvarchar", 50
FillDesignL 9, "EPC3", "nvarchar", 50
FillDesignL 10, "EPC4", "nvarchar", 50
FillDesignL 11, "Rota", "nvarchar", 50
FillDesignL 12, "Adeno", "nvarchar", 50
FillDesignL 13, "Camp", "nvarchar", 50
FillDesignL 14, "CampLatex", "nvarchar", 50
FillDesignL 15, "Latex0157", "nvarchar", 50
FillDesignL 16, "PC0157", "nvarchar", 50
FillDesignL 17, "TCCulture", "nvarchar", 50
FillDesignL 18, "Gram", "nvarchar", 50
FillDesignL 19, "LongEColi", "nvarchar", 50
FillDesignL 20, "OP0", "nvarchar", 50
FillDesignL 21, "OP1", "nvarchar", 50
FillDesignL 22, "OP2", "nvarchar", 50
FillDesignL 23, "o1", "bit"
FillDesignL 24, "o2", "bit"
FillDesignL 25, "o3", "bit"
FillDesignL 26, "CRP", "nvarchar", 50

If IsTableInDB("OP") = False Then 'There is no table  in database
  CreateTable "OP"
Else
  DoTableAnalysis "OP"
End If

End Sub

Public Sub DefineNameExclusions()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "GivenName", "nvarchar", 50
FillDesignL 1, "ReportName", "nvarchar", 50

If IsTableInDB("NameExclusions") = False Then 'There is no table  in database
  CreateTable "NameExclusions"
Else
  DoTableAnalysis "NameExclusions"
End If

End Sub

Public Sub DefineHospitals()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "Text", "nvarchar", 50

If IsTableInDB("Hospitals") = False Then 'There is no table  in database
  CreateTable "Hospitals"
Else
  DoTableAnalysis "Hospitals"
End If

End Sub


Public Sub DefineINRHistory()

ReDim Design(0 To 3) As FieldDefs

FillDesignL 0, "Chart", "nvarchar", 50
FillDesignL 1, "LowerTarget", "nvarchar", 50
FillDesignL 2, "UpperTarget", "nvarchar", 50
FillDesignL 3, "Condition", "nvarchar", 50

If IsTableInDB("INRHistory") = False Then 'There is no table  in database
  CreateTable "INRHistory"
Else
  DoTableAnalysis "INRHistory"
End If

End Sub



Public Sub DefineInstalledPrinters()

ReDim Design(0 To 0) As FieldDefs

FillDesignL 0, "PrinterName", "nvarchar", 50

If IsTableInDB("InstalledPrinters") = False Then 'There is no table  in database
  CreateTable "InstalledPrinters"
Else
  DoTableAnalysis "InstalledPrinters"
End If

End Sub

Public Sub DefineLabName()

ReDim Design(0 To 0) As FieldDefs

FillDesignL 0, "Laboratory", "nvarchar", 50

If IsTableInDB("LabName") = False Then 'There is no table  in database
  CreateTable "LabName"
Else
  DoTableAnalysis "LabName"
End If

End Sub

Public Sub DefineLists()

ReDim Design(0 To 5) As FieldDefs

FillDesignL 0, "Code", "nvarchar", 50
FillDesignL 1, "Text", "nvarchar", 150
FillDesignL 2, "ListType", "nvarchar", 50
FillDesignL 3, "ListOrder", "int"
FillDesignL 4, "InUse", "bit"
FillDesignL 5, "Default", "nvarchar", 50

If IsTableInDB("Lists") = False Then 'There is no table  in database
  CreateTable "Lists"
Else
  DoTableAnalysis "Lists"
End If

End Sub

Public Sub DefineMaxMMessages()

ReDim Design(0 To 1) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Message", "nvarchar", 100

If IsTableInDB("MaxMMessages") = False Then 'There is no table  in database
  CreateTable "MaxMMessages"
Else
  DoTableAnalysis "MaxMMessages"
End If

End Sub

Public Sub DefineMedibridgeRequests()

ReDim Design(0 To 7) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "TestCode", "nvarchar", 50
FillDesignL 2, "TestName", "nvarchar", 50
FillDesignL 3, "SampleDateTime", "datetime"
FillDesignL 4, "ClinDetails", "nvarchar", 50
FillDesignL 5, "Orderer", "nvarchar", 50
FillDesignL 6, "Dept", "nvarchar", 50
FillDesignL 7, "SpecimenSource", "nvarchar", 50

If IsTableInDB("MedibridgeRequests") = False Then 'There is no table  in database
  CreateTable "MedibridgeRequests"
Else
  DoTableAnalysis "MedibridgeRequests"
End If

End Sub

Public Sub DefineMedibridgeResults()

ReDim Design(0 To 7) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "PatName", "nvarchar", 50
FillDesignL 2, "DoB", "datetime"
FillDesignL 3, "Result", "ntext"
FillDesignL 4, "Request", "nvarchar", 50
FillDesignL 5, "MessageTime", "datetime"
FillDesignL 6, "ClinDetails", "ntext"
FillDesignL 7, "Sex", "nvarchar", 50

If IsTableInDB("MedibridgeResults") = False Then 'There is no table  in database
  CreateTable "MedibridgeResults"
Else
  DoTableAnalysis "MedibridgeResults"
End If

End Sub

Public Sub DefineMicroSiteDetails()

ReDim Design(0 To 6) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "Site", "ntext"
FillDesignL 2, "SiteDetails", "ntext"
FillDesignL 3, "PCA0", "ntext"
FillDesignL 4, "PCA1", "ntext"
FillDesignL 5, "PCA2", "ntext"
FillDesignL 6, "PCA3", "ntext"

If IsTableInDB("MicroSiteDetails") = False Then 'There is no table  in database
  CreateTable "MicroSiteDetails"
Else
  DoTableAnalysis "MicroSiteDetails"
End If

End Sub

Public Sub DefineMicroRequests()

ReDim Design(0 To 4) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "RequestDate", "datetime"
FillDesignL 2, "Faecal", "numeric"
FillDesignL 3, "Urine", "numeric"
FillDesignL 4, "Valid", "bit"

If IsTableInDB("MicroRequests") = False Then 'There is no table  in database
  CreateTable "MicroRequests"
Else
  DoTableAnalysis "MicroRequests"
End If

End Sub
Public Sub DefineIsolates()

ReDim Design(0 To 5) As FieldDefs

FillDesignL 0, "SampleID", "numeric", , True
FillDesignL 1, "IsolateNumber", "int"
FillDesignL 2, "OrganismGroup", "nvarchar", 50
FillDesignL 3, "OrganismName", "nvarchar", 50
FillDesignL 4, "Qualifier", "nvarchar", 50
FillDesignL 5, "Valid", "bit"

If IsTableInDB("Isolates") = False Then 'There is no table  in database
  CreateTable "Isolates"
Else
  DoTableAnalysis "Isolates"
End If

End Sub


Public Sub FillDesignL(ByVal ColumnIndex As Long, _
                        ByVal ColumnName As String, _
                        ByVal DataType As String, _
                        Optional ByVal DataLength As Integer, _
                        Optional ByVal NoNull As Boolean)

Design(ColumnIndex).ColumnName = ColumnName
Design(ColumnIndex).DataType = DataType
Design(ColumnIndex).Length = DataLength
Design(ColumnIndex).NoNull = NoNull

End Sub


