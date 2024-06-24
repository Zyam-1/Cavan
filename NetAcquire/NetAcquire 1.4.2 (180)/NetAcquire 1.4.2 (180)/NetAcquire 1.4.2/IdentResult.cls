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
Private pUserName As String 'local copy
Private pDateTimeOfRecord As String 'local copy

Private pIdentResults As IdentResults

Public Sub Save()

          Dim sql As String

19080     On Error GoTo Save_Error

19090     If Trim$(pResult) = "" Then
19100         sql = "DELETE FROM Identification " & _
                  "WHERE SampleID = '" & pSampleID & "' " & _
                  "AND TestType = '" & pTestType & "' " & _
                  "AND TestName = '" & pTestName & "'"
19110     Else
19120         sql = "IF EXISTS(SELECT * FROM Identification " & _
                  "          WHERE SampleID = '" & pSampleID & "' " & _
                  "          AND TestType = '" & pTestType & "' " & _
                  "          AND TestName = '" & pTestName & "') " & _
                  "  UPDATE Identification " & _
                  "  SET Result = '" & pResult & "', " & _
                  "  TestDateTime = '" & pTestDateTime & "', " & _
                  "  Valid = '" & pValid & "', " & _
                  "  Printed = '" & pPrinted & "', " & _
                  "  UserName = '" & AddTicks(pUserName) & "' " & _
                  "  WHERE SampleID = '" & pSampleID & "' " & _
                  "  AND TestType = '" & pTestType & "' " & _
                  "  AND TestName = '" & pTestName & "' " & _
                  "ELSE "
19130         sql = sql & "  INSERT INTO Identification " & _
                  "  (SampleID, TestType, TestName, Result, TestDateTime, Valid, Printed, UserName) VALUES " & _
                  "  ('" & pSampleID & "', " & _
                  "  '" & pTestType & "', " & _
                  "  '" & pTestName & "', " & _
                  "  '" & pResult & "', " & _
                  "  '" & pTestDateTime & "', " & _
                  "  '" & pValid & "', " & _
                  "  '" & pPrinted & "', " & _
                  "  '" & AddTicks(pUserName) & "')"
19140     End If

19150     Cnxn(0).Execute sql
          
19160     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

19170     intEL = Erl
19180     strES = Err.Description
19190     LogError "IdentResult", "Save", intEL, strES, sql

End Sub

Public Property Get IdentResults() As IdentResults

19200     If pIdentResults Is Nothing Then
19210         Set pIdentResults = New IdentResults
19220     End If

19230     Set IdentResults = pIdentResults

End Property


Public Property Set IdentResults(vData As IdentResults)
19240     Set pIdentResults = vData
End Property
Private Sub Class_Terminate()
19250     Set mvarIdentResults = Nothing
End Sub

Public Property Let DateTimeOfRecord(ByVal sData As String)
          'used when assigning a value to the property, on the left side of an assignment.
          'Syntax: X.DateTimeOfRecord = 5
19260     pDateTimeOfRecord = sData
End Property


Public Property Get DateTimeOfRecord() As String
          'used when retrieving value of a property, on the right side of an assignment.
          'Syntax: Debug.Print X.DateTimeOfRecord
19270     DateTimeOfRecord = pDateTimeOfRecord
End Property



Public Property Let UserName(ByVal sData As String)

19280     pUserName = sData

End Property


Public Property Get UserName() As String

19290     UserName = pUserName

End Property



Public Property Let Printed(ByVal vData As Byte)
          'used when assigning a value to the property, on the left side of an assignment.
          'Syntax: X.Printed = 5
19300     pPrinted = vData
End Property


Public Property Get Printed() As Byte
          'used when retrieving value of a property, on the right side of an assignment.
          'Syntax: Debug.Print X.Printed
19310     Printed = pPrinted
End Property



Public Property Let Valid(ByVal vData As Byte)

19320     pValid = vData

End Property


Public Property Get Valid() As Byte

19330     Valid = pValid

End Property



Public Property Let TestDateTime(ByVal sData As String)

19340     pTestDateTime = sData

End Property


Public Property Get TestDateTime() As String

19350     TestDateTime = pTestDateTime

End Property



Public Property Let Result(ByVal sData As String)
          'used when assigning a value to the property, on the left side of an assignment.
          'Syntax: X.Result = 5
19360     pResult = sData
End Property


Public Property Get Result() As String
          'used when retrieving value of a property, on the right side of an assignment.
          'Syntax: Debug.Print X.Result
19370     Result = pResult
End Property



Public Property Let TestName(ByVal sData As String)

19380     pTestName = sData

End Property


Public Property Get TestName() As String

19390     TestName = pTestName

End Property



Public Property Let TestType(ByVal sData As String)
          'used when assigning a value to the property, on the left side of an assignment.
          'Syntax: X.TestType = 5
19400     pTestType = sData

End Property


Public Property Get TestType() As String
          'used when retrieving value of a property, on the right side of an assignment.
          'Syntax: Debug.Print X.TestType
19410     TestType = pTestType
End Property



Public Property Let SampleID(ByVal sData As String)

19420     pSampleID = sData

End Property


Public Property Get SampleID() As String

19430     SampleID = pSampleID

End Property



