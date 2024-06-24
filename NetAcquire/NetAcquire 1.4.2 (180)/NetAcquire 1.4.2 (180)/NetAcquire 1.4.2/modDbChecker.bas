Attribute VB_Name = "modDbChecker"
Option Explicit



Private Sub DoTableAnalysis(ByVal TableName As String)
'
'Dim n As Long
'Dim f As Long
'Dim Found As Boolean
'Dim Matching As Boolean
'Dim sql As String
'Dim tb As Recordset
'Dim tbErr As Recordset
'Dim er As Long
'Dim es As String
'Dim s As String
'Dim blnCheckLength As Boolean
'Dim strRpt As String
'
'
'    sql = "Select top 1 * from [" & TableName & "]"
'    Set tb = New Recordset
'    RecOpenServer tb, sql
'    If tb.EOF Then
'      frmMain.lstNoRows.AddItem TableName
'    End If
'    For f = 0 To tb.Fields.Count - 1
'      Found = False
'      Matching = False
'      For n = 0 To UBound(Design)
'        If UCase$(tb.Fields(f).Name) = UCase$(Design(n).ColumnName) Then
'          Found = True
'          If ((tb.Fields(f).Type = 2 And UCase$(Design(n).DataType) = "SMALLINT") Or _
'             (tb.Fields(f).Type = 4 And UCase$(Design(n).DataType) = "REAL") Or _
'             (tb.Fields(f).Type = 5 And UCase$(Design(n).DataType) = "FLOAT") Or _
'             (tb.Fields(f).Type = 16 And UCase$(Design(n).DataType) = "TINYINT") Or _
'             (tb.Fields(f).Type = 17 And UCase$(Design(n).DataType) = "TINYINT") Or _
'             (tb.Fields(f).Type = 203 And UCase$(Design(n).DataType) = "NTEXT") Or _
'             (tb.Fields(f).Type = 205 And UCase$(Design(n).DataType) = "IMAGE") Or _
'             (tb.Fields(f).Type = 135 And UCase$(Design(n).DataType) = "DATETIME") Or _
'             (tb.Fields(f).Type = 131 And UCase$(Design(n).DataType) = "NUMERIC") Or _
'             (tb.Fields(f).Type = 11 And UCase$(Design(n).DataType) = "BIT") Or _
'             (tb.Fields(f).Type = 129 And UCase$(Design(n).DataType) = "CHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
'             (tb.Fields(f).Type = 200 And UCase$(Design(n).DataType) = "VARCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
'             (tb.Fields(f).Type = 130 And UCase$(Design(n).DataType) = "NCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
'             (tb.Fields(f).Type = 202 And UCase$(Design(n).DataType) = "NVARCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
'             (tb.Fields(f).Type = 3 And UCase$(Design(n).DataType) = "INT") _
'            ) Then
'            Matching = True
'          End If
'          Exit For
'        End If
'      Next
'      s = ""
'      If Not Found Then
'        s = TableName & "." & tb.Fields(f).Name & " not in definitions."
'      ElseIf Not Matching Then
''        blnCheckLength = False
''        strRpt = FieldTypeOf(tb.Fields(f).Type, blnCheckLength)
'
''        If UCase$(strRpt) = UCase$(Design(n).DataType) Then
''          If tb.Fields(f).DefinedSize < Design(n).Length Then
''            sql = "ALTER TABLE " & TableName & " " & _
''                  "ALTER COLUMN " & tb.Fields(f).Name & " " & _
''                  Design(n).DataType & "(" & Design(n).Length & ")"
''            frmMain.txtReport = frmMain.txtReport & sql & vbCrLf
''            Set tbErr = Cnxn.Execute(sql)
''            If Err.Number <> 0 Then
''              frmMain.txtReport = frmMain.txtReport & Err.Description & vbCrLf
''            End If
''          End If
''        ElseIf UCase$(strRpt) = "CHAR" And UCase$(Design(n).DataType) = "NVARCHAR" Then
''          sql = "ALTER TABLE " & TableName & " " & _
''                "ALTER COLUMN " & tb.Fields(f).Name & " " & _
''                "nvarchar(" & tb.Fields(f).DefinedSize & ")"
''          frmMain.txtReport = frmMain.txtReport & "Executing " & sql & vbCrLf
''          frmMain.txtReport.Refresh
''          Err.Clear
''          Cnxn.Execute sql
''          frmMain.txtReport = frmMain.txtReport & sql & vbCrLf
''          frmMain.txtReport.Refresh
''          If Err.Number <> 0 Then
''           frmMain.txtReport = frmMain.txtReport & Err.Description & vbCrLf
''          End If
''        Else
'          s = TableName & "." & tb.Fields(f).Name & " " & _
'              strRpt & " "
'          If blnCheckLength Then
'            s = s & "(" & tb.Fields(f).DefinedSize & ") "
'          End If
'          s = s & "Defined as " & Design(n).DataType & " "
'          If blnCheckLength Then
'            s = s & "(" & Design(n).Length & ") "
'          End If
'        End If
''      End If
'      If s <> "" Then frmMain.txtReport = frmMain.txtReport & s & vbCrLf
'    Next
'
'    For n = 0 To UBound(Design)
'      Found = False
'      Matching = False
'      For f = 0 To tb.Fields.Count - 1
'        If UCase$(tb.Fields(f).Name) = UCase$(Design(n).ColumnName) Then
'          Found = True
'          Exit For
'        End If
'      Next
'      s = ""
'      If Not Found Then
'        s = TableName & "." & Design(n).ColumnName & " " & Design(n).DataType & " (" & Design(n).Length & ") not in " & TableName & "."
'      End If
'      If s <> "" Then frmMain.txtReport = frmMain.txtReport & s & vbCrLf
'    Next
'
'  Case Else:
'    er = Err.Number
'    es = Err.Description
'    MsgBox es
'    Exit Sub
'
'End Select

End Sub
Private Function FieldTypeOf(ByVal intTypeNumber As Integer, _
                             ByRef blnCheckLength As Boolean) _
                             As String

Select Case intTypeNumber
  Case 0: FieldTypeOf = "EMPTY"
  Case 2: FieldTypeOf = "SMALLINT"
  Case 3: FieldTypeOf = "INT"
  Case 4:  FieldTypeOf = "REAL"
  Case 5:  FieldTypeOf = "FLOAT"
  Case 6:  FieldTypeOf = "CURRENCY"
  Case 7:  FieldTypeOf = "DATE"
  Case 8:  FieldTypeOf = "BSTR"
  Case 9:  FieldTypeOf = "DISPATCH"
  Case 10:  FieldTypeOf = "ERROR"
  Case 11: FieldTypeOf = "BIT"
  Case 12: FieldTypeOf = "VARIANT"
  Case 13: FieldTypeOf = "UNKNOWN"
  Case 14: FieldTypeOf = "DECIMAL"
  Case 16: FieldTypeOf = "TINYINT"
  Case 17: FieldTypeOf = "TINYINT"
  Case 20: FieldTypeOf = "BIGINT"
  Case 64: FieldTypeOf = "FILETIME"
  Case 72: FieldTypeOf = "GUID"
  Case 130: FieldTypeOf = "NCHAR": blnCheckLength = True
  Case 131: FieldTypeOf = "NUMERIC"
  Case 133: FieldTypeOf = "DBDATE"
  Case 134: FieldTypeOf = "DBTIME"
  Case 135: FieldTypeOf = "DATETIME"
  Case 136: FieldTypeOf = "CHAPTER"
  Case 128: FieldTypeOf = "BINARY"
  Case 129: FieldTypeOf = "CHAR": blnCheckLength = True
  Case 200: FieldTypeOf = "VARCHAR": blnCheckLength = True
  Case 201: FieldTypeOf = "LONGVARCHAR": blnCheckLength = True
  Case 202: FieldTypeOf = "NVARCHAR": blnCheckLength = True
  Case 203: FieldTypeOf = "NTEXT"
  Case 204: FieldTypeOf = "VARBINARY": blnCheckLength = True
  Case 205: FieldTypeOf = "IMAGE": blnCheckLength = True
  Case 8192: FieldTypeOf = "ARRAY"
End Select

End Function



Public Sub FillFieldDefs()

'DefineErrorLog
'DefineABDefinitions
'DefineAges
'DefineAntibiotics
'DefineArcBGARepeats
'DefineArcBGAResults
'DefineArcBioRepeats
'DefineArcBioResults
'DefineArcCoagRepeats
'DefineArcCoagResults
'DefineArcComments
'DefineArcCytoResults
'DefineArcDemographics
'DefineArcEndRepeats
'DefineArcEndResults
'DefineArcExtResults
'DefineArcHaemRepeats
'DefineArcHaemResults
'DefineArcHistoResults
'DefineArcImmRepeats
'DefineArcImmResults
'DefineArcMasks
'DefineArcUsers
'DefineBarCodeControl
'DefineBarCodes
'DefineBGADefinitions
'DefineBGARepeats
'DefineBGAResults
'DefineBiochemistryQC
'DefineBioFlags
'DefineBioFlagsRep
'DefineBioQCDefs
'DefineBioRepeats
'DefineBioRequests
'DefineBioResults
'DefineBioTestDefinitions
'DefineClinDetails
'DefineClinicians
'DefineCoagControls
'DefineCoagDefault
'DefineCoagPanels
'DefineCoagRepeats
'DefineCoagRequests
'DefineCoagResults
'DefineCoagTestDefinitions
'DefineComments
'DefineConsultantList
'DefineControls
'DefineCreatinine
'DefineCulture
'DefineCytoResults
'DefineDemographics
'DefineDifferentials
'DefineDifferentialTitles
'DefineEAddress
'DefineEndMasks
'DefineEndRepeats
'DefineEndRequests
'DefineEndResults
'DefineEndTestDefinitions
'DefineETC
'DefineETests
'DefineExternalDefinitions
'DefineExtPanels
'DefineExtResults
'DefineFaecalRequests
'DefineFaeces
'DefineFastings
'DefineForcedABReport
'DefineGPs
'DefineHaemCondition
'DefineHaemFlags
'DefineHaemFlagsRep
'DefineHaemRepeats
'DefineHaemRequests
'DefineHaemResults
'DefineHaemTestDefinitions
'DefineHbA1c
'DefineHISErrors
'DefineHistoBlock
'DefineHistoComments
'DefineHistoResults
'DefineHistory
'DefineHistoSpecimen
'DefineHistoStain
'DefineHMRU
'DefineHospitals
'DefineImmMasks
'DefineImmRepeats
'DefineImmRequests
'DefineImmResults
'DefineImmTestDefinitions
'DefineINRHistory
'DefineInstalledPrinters
'DefineInterp
'DefineIPanels
'DefineIsolates
'DefineLabName
'DefineLists
'DefineMasks
'DefineMaxMMessages
'DefineMedibridgeRequests
'DefineMedibridgeResults
'DefineMicroRequests
'DefineMicroSiteDetails
'DefineMRU
'DefineNameExclusions
'DefineNINChart
'DefineNinEvents
'DefineNinRequests
'DefineOP
'DefineOptions
'DefineOCAssignedAnswers
'DefineOCAssignedAnswersArc
'DefineOCAssignedQAN
'DefineOCAssignedQANArc
'DefineOCOrderContents
'DefineOCOrderContentsArc
'DefineOCOrderPanel
'DefineOCOrderPanelArc
'DefineOrganisms
'DefinePanels
'DefinePatientIFs
'DefinePatientUpdates
'DefinePhoneLog
'DefinePractices
'DefinePrinters
'DefinePrintPending
'DefineReagentLotNumbers
'DefineSemenResults
'DefineSendCopyTo
'DefineSensitivities
'DefineSexNames
'DefineSourcePanels
'DefineStock
'DefineStockControl
'DefineStockReagents
'DefineTrace
'DefineUnits
'DefineUpdates
'DefineUrine
'DefineUrineIdent
'DefineUsers
'DefineViewedReports
'DefineWards

End Sub


