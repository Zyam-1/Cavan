VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sensitivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String
Private pIsolateNumber As Integer
Private pAntibioticCode As String
Private pAntibioticName As String
Private pResult As String
Private pReport As Integer
Private pCPOFlag As String
Private pRunDate As String
Private pRunDateTime As String
Private pRSI As String
Private pUserCode As String
Private pForced As Integer
Private pSecondary As Integer
Private pValid As Integer
Private pAuthoriserCode As String
Private pDateTimeOfRecord As String
Private pArchivedBy As String
Private pArchiveDateTime As String

Public Sub Save()

      Dim sql As String

10    On Error GoTo Save_Error

20    sql = "IF EXISTS (SELECT * FROM Sensitivities WHERE " & _
            "           SampleID = '" & pSampleID & "' " & _
            "           AND IsolateNumber = '" & pIsolateNumber & "' " & _
            "           AND AntibioticCode = '" & pAntibioticCode & "') " & _
            "  UPDATE Sensitivities " & _
            "  SET Result = '" & pResult & "', " & _
            "  Report = '" & pReport & "', " & _
            "  CPOFlag = '" & pCPOFlag & "', " & _
            "  RunDate = '" & Format$(pRunDate, "dd/MMM/yyyy") & "', " & _
            "  RunDateTime = getdate(), " & _
            "  RSI = '" & pRSI & "', " & _
            "  UserCode = '" & pUserCode & "', " & _
            "  Forced = '" & pForced & "', " & _
            "  Secondary = '" & pSecondary & "', " & _
            "  Valid = '" & pValid & "', " & _
            "  AuthoriserCode = '" & pAuthoriserCode & "' " & _
            "  WHERE SampleID = '" & pSampleID & "' " & _
            "  AND IsolateNumber = '" & pIsolateNumber & "' " & _
      "  AND AntibioticCode = '" & pAntibioticCode & "' "
30    sql = sql & "ELSE " & _
            "  INSERT INTO Sensitivities " & _
            "  (SampleID, IsolateNumber, AntibioticCode, Result, Report, CPOFlag, RunDate, " & _
            "  RunDateTime, RSI, UserCode, Forced, Secondary, Valid, AuthoriserCode) " & _
            "  VALUES " & _
            "  ('" & pSampleID & "', " & _
            "  '" & pIsolateNumber & "', " & _
            "  '" & pAntibioticCode & "', " & _
            "  '" & pResult & "', " & _
            "  '" & pReport & "', " & _
            "  '" & pCPOFlag & "', " & _
            "  '" & Format$(pRunDate, "dd/MMM/yyyy") & "', " & _
            "  '" & Format$(pRunDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
            "  '" & pRSI & "', " & _
            "  '" & pUserCode & "', " & _
            "  '" & pForced & "', " & _
            "  '" & pSecondary & "', " & _
            "  '" & pValid & "', " & _
            "  '" & pAuthoriserCode & "')"

40    Cnxn(0).Execute sql

50    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "Sensitivity", "Save", intEL, strES, sql

End Sub



Public Property Get IsolateNumber() As Integer

10    IsolateNumber = pIsolateNumber

End Property

Public Property Get SampleID() As String

10    SampleID = pSampleID

End Property


Public Property Get AntibioticCode() As String

10    AntibioticCode = pAntibioticCode

End Property

Public Property Get ArchivedBy() As String

10    ArchivedBy = pArchivedBy

End Property
Public Property Get ArchiveDateTime() As String

10    ArchiveDateTime = pArchiveDateTime

End Property

Public Property Get AntibioticName() As String

10    AntibioticName = pAntibioticName

End Property


Public Property Get Result() As String

10    Result = pResult

End Property
Public Property Get Report() As Integer

10    Report = pReport

End Property

Public Property Get CPOFlag() As String

10    CPOFlag = pCPOFlag

End Property
Public Property Get Rundate() As String

10    Rundate = pRunDate

End Property

Public Property Get RunDateTime() As String

10    RunDateTime = pRunDateTime

End Property


Public Property Get RSI() As String

10    RSI = pRSI

End Property


Public Property Get UserCode() As String

10    UserCode = pUserCode

End Property


Public Property Get Forced() As Integer

10    Forced = pForced

End Property

Public Property Get Secondary() As Integer

10    Secondary = pSecondary

End Property

Public Property Get Valid() As Integer

10    Valid = pValid

End Property

Public Property Get AuthoriserCode() As String

10    AuthoriserCode = pAuthoriserCode

End Property

Public Property Get DateTimeOfRecord() As String

10    DateTimeOfRecord = pDateTimeOfRecord

End Property

Public Property Let SampleID(ByVal sNewValue As String)

10    pSampleID = sNewValue

End Property
Public Property Let IsolateNumber(ByVal iNewValue As Integer)

10    pIsolateNumber = iNewValue

End Property

Public Property Let AntibioticCode(ByVal sNewValue As String)

10    pAntibioticCode = sNewValue

End Property
Public Property Let ArchivedBy(ByVal sNewValue As String)

10    pArchivedBy = sNewValue

End Property

Public Property Let ArchiveDateTime(ByVal sNewValue As String)

10    pArchiveDateTime = sNewValue

End Property


Public Property Let AntibioticName(ByVal sNewValue As String)

10    pAntibioticName = sNewValue

End Property

Public Property Let Result(ByVal sNewValue As String)

10    pResult = sNewValue

End Property

Public Property Let Report(ByVal bNewValue As Integer)

10    pReport = IIf(bNewValue <> 0, 1, 0)

End Property

Public Property Let CPOFlag(ByVal sNewValue As String)

10    pCPOFlag = sNewValue

End Property
Public Property Let DateTimeOfRecord(ByVal sNewValue As String)

10    pDateTimeOfRecord = sNewValue

End Property

Public Property Let Rundate(ByVal sNewValue As String)

10    pRunDate = sNewValue

End Property

Public Property Let RunDateTime(ByVal sNewValue As String)

10    pRunDateTime = sNewValue

End Property
Public Property Let RSI(ByVal sNewValue As String)

10    pRSI = sNewValue

End Property

Public Property Let UserCode(ByVal sNewValue As String)

10    pUserCode = sNewValue

End Property
Public Property Let Forced(ByVal bNewValue As Integer)

10    pForced = IIf(bNewValue <> 0, 1, 0)

End Property
Public Property Let Secondary(ByVal bNewValue As Integer)

10    pSecondary = IIf(bNewValue <> 0, 1, 0)

End Property

Public Property Let Valid(ByVal bNewValue As Integer)

10    pValid = IIf(bNewValue <> 0, 1, 0)

End Property
Public Property Let AuthoriserCode(ByVal sNewValue As String)

10    pAuthoriserCode = sNewValue

End Property

