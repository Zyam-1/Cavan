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

10    On Error GoTo Save_Error

20    If IsDate(pValidatedDateTime) Then
30      ValDate = Format$(pValidatedDateTime, "'dd/MMM/yyyy HH:nn:ss'")
40    Else
50      ValDate = "NULL"
60    End If
70    If IsDate(pPrintedDateTime) Then
80      PrnDate = Format$(pPrintedDateTime, "'dd/MMM/yyyy HH:nn:ss'")
90    Else
100     PrnDate = "NULL"
110   End If

120   If Trim$(pResult) = "" Then
130     sql = "Delete from SemenResults50 where " & _
              "SampleID = '" & pSampleID & "' " & _
              "and TestName = '" & pTestName & "'"
140   Else
150     sql = "IF EXISTS (SELECT * FROM SemenResults50 WHERE " & _
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
160     sql = sql & "ELSE " & _
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
170   End If
180   Cnxn(0).Execute sql

190   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "SemenResult", "Save", intEL, strES, sql

End Sub




Public Property Get SampleID() As String

10    SampleID = pSampleID

End Property

Public Property Let SampleID(ByVal sNewValue As String)

10    pSampleID = sNewValue

End Property

Public Property Get TestName() As String

10    TestName = pTestName

End Property

Public Property Get ValidatedBy() As String

10    ValidatedBy = pValidatedBy

End Property


Public Property Get PrintedBy() As String

10    PrintedBy = pPrintedBy

End Property


Public Property Get Result() As String

10    Result = pResult

End Property

Public Property Get DateTimeOfRecord() As String

10    DateTimeOfRecord = pDateTimeOfRecord

End Property

Public Property Get Valid() As Integer

10    Valid = pValid

End Property


Public Property Get Healthlink() As Integer

10    Healthlink = pHealthLink

End Property


Public Property Get Printed() As Integer

10    Printed = pPrinted

End Property


Public Property Get ValidatedDateTime() As String

10    ValidatedDateTime = pValidatedDateTime

End Property

Public Property Get UserName() As String

10    UserName = pUserName

End Property


Public Property Get PrintedDateTime() As String

10    PrintedDateTime = pPrintedDateTime

End Property


Public Property Let ValidatedBy(ByVal sNewValue As String)

10    pValidatedBy = sNewValue

End Property
Public Property Let TestName(ByVal sNewValue As String)

10    pTestName = sNewValue

End Property

Public Property Let UserName(ByVal sNewValue As String)

10    pUserName = sNewValue

End Property

Public Property Let PrintedBy(ByVal sNewValue As String)

10    pPrintedBy = sNewValue

End Property

Public Property Let DateTimeOfRecord(ByVal sNewValue As String)

10    pDateTimeOfRecord = sNewValue

End Property


Public Property Let Valid(ByVal NewValue As Integer)

10    pValid = NewValue

End Property

Public Property Let ValidatedDateTime(ByVal NewValue As String)

10    pValidatedDateTime = NewValue

End Property

Public Property Let PrintedDateTime(ByVal NewValue As String)

10    pPrintedDateTime = NewValue

End Property


Public Property Let Printed(ByVal NewValue As Integer)

10    pPrinted = NewValue

End Property

Public Property Let Healthlink(ByVal NewValue As Integer)

10    pHealthLink = NewValue

End Property

Public Property Let Result(ByVal sNewValue As String)

10    pResult = sNewValue

End Property

