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
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSampleID As String 'local copy
Private mvarTestType As String 'local copy
Private mvarTestName As String 'local copy
Private mvarResult As String 'local copy
Private mvarTestDateTime As Date 'local copy
Private mvarValid As Byte 'local copy
Private mvarPrinted As Byte 'local copy
Private mvarUsername As String 'local copy
Private mvarDateTimeOfRecord As Date 'local copy

Public Enum IdentSaveAction
    IdentSaveActionAdd
    IdentSaveActionUpdate
End Enum

Private mvarIdentResults As IdentResults

Private Function PopulateIdentList(tb As Recordset) As IdentResults

      Dim Idents As New IdentResults
      Dim Ident As IdentResult

10    On Error GoTo PopulateIdentList_Error

20    While Not tb.EOF
30      Set Ident = New IdentResult
40      Ident.PopulateIdent Ident, tb
50      Idents.Add Ident
60      tb.MoveNext
70    Wend

80    Set PopulateIdentList = Idents

90    Exit Function

PopulateIdentList_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "IdentResult", "PopulateIdentList", intEL, strES

End Function

Friend Function PopulateIdent(Ident As IdentResult, tb As Recordset) As Boolean

10    On Error GoTo PopulateIdent_Error

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

130   Exit Function

PopulateIdent_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "IdentResult", "PopulateIdent", intEL, strES

End Function

Public Function SaveIdent(SaveAction As IdentSaveAction) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveIdent_Error

20    Set tb = New Recordset
30    If SaveAction = IdentSaveActionAdd Then
40      sql = "Select * From Identification Where 1 = 0"
50      RecOpenServer 0, tb, sql
60      tb.AddNew
70    ElseIf SaveAction = IdentSaveActionUpdate Then
80      sql = "SELECT * FROM Identification WHERE " & _
              "SampleID = '" & Me.SampleID & "' " & _
              "AND TestName = '" & Me.TestName & "'"
90      RecOpenServer 0, tb, sql
100   End If
    
110   tb!SampleID = Me.SampleID
120   tb!TestType = Me.TestType
130   tb!TestName = Me.TestName
140   tb!Result = Me.Result
150   tb!TestDateTime = Me.TestDateTime
160   tb!Valid = Me.Valid
170   tb!Printed = Me.Printed
180   tb!UserName = Me.UserName
190   tb!DateTimeOfRecord = Me.DateTimeOfRecord
200   tb.Update
    
210   Exit Function

SaveIdent_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "IdentResult", "SaveIdent", intEL, strES, sql

End Function

Public Function LoadIdent(ByVal SampleID As String, ByVal TestType As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadIdent_Error

20    sql = "SELECT * FROM Identification WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND TestType = '" & TestType & "'"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      LoadIdent = False
70    Else
80      PopulateIdent Me, tb
90      LoadIdent = True
100   End If

110   Exit Function

LoadIdent_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "IdentResult", "LoadIdent", intEL, strES, sql

End Function

Public Function LoadIdentList(ByVal SampleID As String, ByVal TestType As String) As IdentResults

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo LoadIdentList_Error

20    sql = "SELECT * FROM Identification WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND TestType = '" & TestType & "'"

30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60      Set LoadIdentList = PopulateIdentList(tb)
70    Else
80      Set LoadIdentList = Nothing
90    End If

100   Exit Function

LoadIdentList_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "IdentResult", "LoadIdentList", intEL, strES, sql

End Function




Public Property Get IdentResults() As IdentResults
10        If mvarIdentResults Is Nothing Then
20            Set mvarIdentResults = New IdentResults
30        End If


40        Set IdentResults = mvarIdentResults
End Property


Public Property Set IdentResults(vData As IdentResults)
10        Set mvarIdentResults = vData
End Property
Private Sub Class_Terminate()
10        Set mvarIdentResults = Nothing
End Sub

Public Property Let DateTimeOfRecord(ByVal vData As Date)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.DateTimeOfRecord = 5
10        mvarDateTimeOfRecord = vData
End Property


Public Property Get DateTimeOfRecord() As Date
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.DateTimeOfRecord
10        DateTimeOfRecord = mvarDateTimeOfRecord
End Property



Public Property Let UserName(ByVal vData As String)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.Username = 5
10        mvarUsername = vData
End Property


Public Property Get UserName() As String
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.Username
10        UserName = mvarUsername
End Property



Public Property Let Printed(ByVal vData As Byte)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.Printed = 5
10        mvarPrinted = vData
End Property


Public Property Get Printed() As Byte
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.Printed
10        Printed = mvarPrinted
End Property



Public Property Let Valid(ByVal vData As Byte)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.Valid = 5
10        mvarValid = vData
End Property


Public Property Get Valid() As Byte
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.Valid
10        Valid = mvarValid
End Property



Public Property Let TestDateTime(ByVal vData As Date)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.TestDateTime = 5
10        mvarTestDateTime = vData
End Property


Public Property Get TestDateTime() As Date
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.TestDateTime
10        TestDateTime = mvarTestDateTime
End Property



Public Property Let Result(ByVal vData As String)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.Result = 5
10        mvarResult = vData
End Property


Public Property Get Result() As String
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.Result
10        Result = mvarResult
End Property



Public Property Let TestName(ByVal vData As String)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.TestName = 5
10        mvarTestName = vData
End Property


Public Property Get TestName() As String
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.TestName
10        TestName = mvarTestName
End Property



Public Property Let TestType(ByVal vData As String)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.TestType = 5
10        mvarTestType = vData
End Property


Public Property Get TestType() As String
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.TestType
10        TestType = mvarTestType
End Property



Public Property Let SampleID(ByVal vData As String)
      'used when assigning a value to the property, on the left side of an assignment.
      'Syntax: X.SampleID = 5
10        mvarSampleID = vData
End Property


Public Property Get SampleID() As String
      'used when retrieving value of a property, on the right side of an assignment.
      'Syntax: Debug.Print X.SampleID
10        SampleID = mvarSampleID
End Property



