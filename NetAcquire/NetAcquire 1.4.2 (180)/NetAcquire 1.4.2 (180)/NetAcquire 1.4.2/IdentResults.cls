VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IdentResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"IdentResult"
Attribute VB_Ext_KEY = "Member0" ,"IdentResult"
'local variable to hold collection
Option Explicit

Private mColIdents As Collection
Public Property Get Item(vntIndexKey As Variant) As IdentResult
Attribute Item.VB_UserMemId = 0
          'used when referencing an element in the collection
          'vntIndexKey contains either the Index or Key to the collection,
          'this is why it is declared as a Variant
          'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
19440     On Error GoTo Item_Error

19450     Set Item = mColIdents(vntIndexKey)

19460     Exit Property

Item_Error:

          Dim strES As String
          Dim intEL As Integer

19470     intEL = Erl
19480     strES = Err.Description
19490     LogError "IdentResults", "Item", intEL, strES
End Property



Public Property Get Count() As Long
          'used when retrieving the number of elements in the
          'collection. Syntax: Debug.Print x.Count
19500     On Error GoTo Count_Error

19510     Count = mColIdents.Count

19520     Exit Property

Count_Error:

          Dim strES As String
          Dim intEL As Integer

19530     intEL = Erl
19540     strES = Err.Description
19550     LogError "IdentResults", "Count", intEL, strES
End Property


Public Sub Remove(vntIndexKey As Variant)
          'used when removing an element from the collection
          'vntIndexKey contains either the Index or Key, which is why
          'it is declared as a Variant
          'Syntax: x.Remove(xyz)


19560     On Error GoTo Remove_Error

19570     mColIdents.Remove vntIndexKey

19580     Exit Sub

Remove_Error:

          Dim strES As String
          Dim intEL As Integer

19590     intEL = Erl
19600     strES = Err.Description
19610     LogError "IdentResults", "Remove", intEL, strES
          
End Sub

Public Sub Add(NewItem As IdentResult)

19620     On Error GoTo Add_Error

19630     mColIdents.Add NewItem

19640     Exit Sub

Add_Error:

          Dim strES As String
          Dim intEL As Integer

19650     intEL = Erl
19660     strES = Err.Description
19670     LogError "IdentResults", "Add", intEL, strES

End Sub


Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim ID As IdentResult

19680     On Error GoTo Load_Error

19690     sql = "SELECT SampleID, TestType, TestName, Result, TestDateTime, Valid, Printed, Username, COALESCE(DateTimeOfRecord, getdate()) DateTimeOfRecord " & _
              "FROM Identification WHERE " & _
              "SampleID = '" & SampleID & "'"

19700     Set tb = New Recordset
19710     RecOpenClient 0, tb, sql
19720     Do While Not tb.EOF
19730         Set ID = New IdentResult
19740         With ID
19750             .SampleID = tb!SampleID
19760             .TestType = tb!TestType
19770             .TestName = tb!TestName
19780             .Result = tb!Result
19790             .TestDateTime = tb!TestDateTime
19800             .Valid = tb!Valid
19810             .Printed = tb!Printed
19820             .UserName = tb!UserName
19830             .DateTimeOfRecord = tb!DateTimeOfRecord
19840         End With
19850         mColIdents.Add ID
19860         tb.MoveNext
19870     Loop

19880     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

19890     intEL = Erl
19900     strES = Err.Description
19910     LogError "IdentResults", "Load", intEL, strES, sql

End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
          'this property allows you to enumerate
          'this collection with the For...Each syntax
19920     On Error GoTo NewEnum_Error

19930     Set NewEnum = mColIdents.[_NewEnum]

19940     Exit Property

NewEnum_Error:

          Dim strES As String
          Dim intEL As Integer

19950     intEL = Erl
19960     strES = Err.Description
19970     LogError "IdentResults", "NewEnum", intEL, strES
End Property


Private Sub Class_Initialize()
          'creates the collection when this class is created
19980     Set mColIdents = New Collection
End Sub


Private Sub Class_Terminate()
          'destroys collection when this class is terminated
19990     Set mColIdents = Nothing
End Sub

