VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IdentResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"IdentResults"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private pSampleID As String 'local copy
Private pTestType As String 'local copy
Private pTestName As String 'local copy
Private pResult As String 'local copy
Private pTestDateTime As String 'local copy
Private pValid As Byte 'local copy
Private pPrinted As Byte 'local copy
Private pUsername As String 'local copy
Private pDateTimeOfRecord As String 'local copy

Private pIdentResults As IdentResults

Private Sub Populate(Ident As IdentResult, tb As Recordset)

10    On Error GoTo Populate_Error

20    With Ident
30        .SampleID = tb!SampleID
40        .TestType = tb!TestType
50        .TestName = tb!TestName
60        .Result = tb!Result
70        .TestDateTime = tb!TestDateTime
80        .Valid = tb!Valid
90        .Printed = tb!Printed
100       .UserName = tb!UserName
110       .DateTimeOfRecord = tb!DateTimeOfRecord
120   End With

130   Exit Sub

Populate_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "IdentResult", "Populate", intEL, strES

End Sub

Public Sub Save()

      Dim sql As String

10    On Error GoTo Save_Error

20    If Trim$(pResult) = "" Then
30      sql = "DELETE FROM Identification " & _
              "WHERE SampleID = '" & pSampleID & "' " & _
              "AND TestType = '" & pTestType & "' " & _
              "AND TestName = '" & pTestName & "'"
40    Else
50      sql = "IF EXISTS(SELECT * FROM Identification " & _
              "          WHERE SampleID = '" & pSampleID & "' " & _
              "          AND TestType = '" & pTestType & "' " & _
              "          AND TestName = '" & pTestName & "') " & _
              "  UPDATE Identification " & _
              "  SET Result = '" & pResult & "', " & _
              "  TestDateTime = '" & pTestDateTime & "', " & _
              "  Valid = '" & pValid & "', " & _
              "  Printed = '" & pPrinted & "', " & _
              "  UserName = '" & pUsername & "' " & _
              "  WHERE SampleID = '" & pSampleID & "' " & _
              "  AND TestType = '" & pTestType & "' " & _
              "  AND TestName = '" & pTestName & "' " & _
              "ELSE "
60      sql = sql & "  INSERT INTO Identification " & _
              "  (SampleID, TestType, TestName, Result, TestDateTime, Valid, Printed, UserName) VALUES " & _
              "  ('" & pSampleID & "', " & _
              "  '" & pTestType & "', " & _
              "  '" & pTestName & "', " & _
              "  '" & pResult & "', " & _
              "  '" & pTestDateTime & "', " & _
              "  '" & pValid & "', " & _
              "  '" & pPrinted & "', " & _
              "  '" & pUsername & "')"
70    End If

80    Cnxn(0).Execute sql
    
90    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "IdentResult", "Save", intEL, strES, sql

End Sub

Public Property Get IdentResults() As IdentResults

10        If pIdentResults Is Nothing Then
20            Set pIdentResults = New IdentResults
30        End If

40        Set IdentResults = pIdentResults

End Property


Public Property Set IdentResults(vData As IdentResults)
10        Set pIdentResults = vData
End Property
Private Sub Class_Terminate()
10        Set mvarIdentResults = Nothing
End Sub

Public Property Let DateTimeOfRecord(ByVal sData As String)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.DateTimeOfRecord = 5
10        pDateTimeOfRecord = sData
End Property


Public Property Get DateTimeOfRecord() As String
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.DateTimeOfRecord
10        DateTimeOfRecord = pDateTimeOfRecord
End Property



Public Property Let UserName(ByVal sData As String)

10        pUsername = sData

End Property


Public Property Get UserName() As String

10        UserName = pUsername

End Property



Public Property Let Printed(ByVal vData As Byte)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.Printed = 5
10        pPrinted = vData
End Property


Public Property Get Printed() As Byte
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.Printed
10        Printed = pPrinted
End Property



Public Property Let Valid(ByVal vData As Byte)

10        pValid = vData

End Property


Public Property Get Valid() As Byte

10        Valid = pValid

End Property



Public Property Let TestDateTime(ByVal sData As String)

10        pTestDateTime = sData

End Property


Public Property Get TestDateTime() As String

10        TestDateTime = pTestDateTime

End Property



Public Property Let Result(ByVal sData As String)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.Result = 5
10        pResult = sData
End Property


Public Property Get Result() As String
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.Result
10        Result = pResult
End Property



Public Property Let TestName(ByVal sData As String)

10        pTestName = sData

End Property


Public Property Get TestName() As String

10        TestName = pTestName

End Property



Public Property Let TestType(ByVal sData As String)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.TestType = 5
10        pTestType = sData

End Property


Public Property Get TestType() As String
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.TestType
10        TestType = pTestType
End Property



Public Property Let SampleID(ByVal sData As String)

10        pSampleID = sData

End Property


Public Property Get SampleID() As String

10        SampleID = pSampleID

End Property



