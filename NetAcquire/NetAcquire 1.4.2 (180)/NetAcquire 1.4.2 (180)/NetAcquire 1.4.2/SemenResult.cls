VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SemenResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String
Private pTestName As String
Private pResult As String
Private pValid As Integer
Private pValidatedBy As String
Private pValidatedDateTime As String
Private pPrinted As Integer
Private pPrintedBy As String
Private pPrintedDateTime As String
Private pUserName As String
Private pDateTimeOfRecord As String
Private pHealthLink As Integer
Public Sub Save()

          Dim sql As String
          Dim ValDate As String
          Dim PrnDate As String

23310     On Error GoTo Save_Error

23320     If IsDate(pValidatedDateTime) Then
23330         ValDate = Format$(pValidatedDateTime, "'dd/MMM/yyyy HH:nn:ss'")
23340     Else
23350         ValDate = "NULL"
23360     End If
23370     If IsDate(pPrintedDateTime) Then
23380         PrnDate = Format$(pPrintedDateTime, "'dd/MMM/yyyy HH:nn:ss'")
23390     Else
23400         PrnDate = "NULL"
23410     End If

23420     If Trim$(pResult) = "" Then
23430         sql = "Delete from SemenResults50 where " & _
                  "SampleID = '" & pSampleID & "' " & _
                  "and TestName = '" & pTestName & "'"
23440     Else
23450         sql = "IF EXISTS (SELECT * FROM SemenResults50 WHERE " & _
                  "           SampleID = '" & pSampleID & "' " & _
                  "           AND TestName = '" & pTestName & "') " & _
                  "  UPDATE SemenResults50 " & _
                  "  SET Result = '" & pResult & "', " & _
                  "  UserName = '" & AddTicks(pUserName) & "', " & _
                  "  Valid = '" & pValid & "', " & _
                  "  ValidatedBy = '" & AddTicks(pValidatedBy) & "', " & _
                  "  ValidatedDateTime = " & ValDate & ", " & _
                  "  Printed = '" & pPrinted & "', " & _
                  "  PrintedBy = '" & AddTicks(pPrintedBy) & "', " & _
                  "  PrintedDateTime = " & PrnDate & " " & _
                  "  WHERE SampleID = '" & pSampleID & "' " & _
                  "  AND TestName = '" & pTestName & "' "
23460         sql = sql & "ELSE " & _
                  "  INSERT INTO SemenResults50 (SampleID, TestName, Result, Valid, ValidatedBy, ValidatedDateTime, " & _
                  "                              Printed, PrintedBy, PrintedDateTime, UserName) " & _
                  "  VALUES " & _
                  "  ('" & pSampleID & "', " & _
                  "  '" & pTestName & "', " & _
                  "  '" & pResult & "', " & _
                  "  '" & pValid & "', " & _
                  "  '" & AddTicks(pValidatedBy) & "', " & _
                  "  " & ValDate & ", " & _
                  "  '" & pPrinted & "', " & _
                  "  '" & AddTicks(pPrintedBy) & "', " & _
                  "  " & PrnDate & ", " & _
                  "  '" & AddTicks(pUserName) & "')"
23470     End If
23480     Cnxn(0).Execute sql

23490     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

23500     intEL = Erl
23510     strES = Err.Description
23520     LogError "SemenResult", "Save", intEL, strES, sql

End Sub




Public Property Get SampleID() As String

23530     SampleID = pSampleID

End Property

Public Property Let SampleID(ByVal sNewValue As String)

23540     pSampleID = sNewValue

End Property

Public Property Get TestName() As String

23550     TestName = pTestName

End Property

Public Property Get ValidatedBy() As String

23560     ValidatedBy = pValidatedBy

End Property


Public Property Get PrintedBy() As String

23570     PrintedBy = pPrintedBy

End Property


Public Property Get Result() As String

23580     Result = pResult

End Property

Public Property Get DateTimeOfRecord() As String

23590     DateTimeOfRecord = pDateTimeOfRecord

End Property

Public Property Get Valid() As Integer

23600     Valid = pValid

End Property


Public Property Get HealthLink() As Integer

23610     HealthLink = pHealthLink

End Property


Public Property Get Printed() As Integer

23620     Printed = pPrinted

End Property


Public Property Get ValidatedDateTime() As String

23630     ValidatedDateTime = pValidatedDateTime

End Property

Public Property Get UserName() As String

23640     UserName = pUserName

End Property


Public Property Get PrintedDateTime() As String

23650     PrintedDateTime = pPrintedDateTime

End Property


Public Property Let ValidatedBy(ByVal sNewValue As String)

23660     pValidatedBy = sNewValue

End Property
Public Property Let TestName(ByVal sNewValue As String)

23670     pTestName = sNewValue

End Property

Public Property Let UserName(ByVal sNewValue As String)

23680     pUserName = sNewValue

End Property

Public Property Let PrintedBy(ByVal sNewValue As String)

23690     pPrintedBy = sNewValue

End Property

Public Property Let DateTimeOfRecord(ByVal sNewValue As String)

23700     pDateTimeOfRecord = sNewValue

End Property


Public Property Let Valid(ByVal NewValue As Integer)

23710     pValid = NewValue

End Property

Public Property Let ValidatedDateTime(ByVal NewValue As String)

23720     pValidatedDateTime = NewValue

End Property

Public Property Let PrintedDateTime(ByVal NewValue As String)

23730     pPrintedDateTime = NewValue

End Property


Public Property Let Printed(ByVal NewValue As Integer)

23740     pPrinted = NewValue

End Property

Public Property Let HealthLink(ByVal NewValue As Integer)

23750     pHealthLink = NewValue

End Property

Public Property Let Result(ByVal sNewValue As String)

23760     pResult = sNewValue

End Property

