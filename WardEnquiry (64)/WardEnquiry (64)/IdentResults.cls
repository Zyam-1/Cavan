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

Private mCol As Collection
Public Property Get Item(vntIndexKey As Variant) As IdentResult
Attribute Item.VB_UserMemId = 0
          'used when referencing an element in the collection
          'vntIndexKey contains either the Index or Key to the collection,
          'this is why it is declared as a Variant
          'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
10    On Error GoTo Item_Error

20      Set Item = mCol(vntIndexKey)

30    Exit Property

Item_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "IdentResults", "Item", intEL, strES
End Property



Public Property Get Count() As Long
          'used when retrieving the number of elements in the
          'collection. Syntax: Debug.Print x.Count
10    On Error GoTo Count_Error

20        Count = mCol.Count

30    Exit Property

Count_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "IdentResults", "Count", intEL, strES
End Property


Public Sub Remove(vntIndexKey As Variant)
          'used when removing an element from the collection
          'vntIndexKey contains either the Index or Key, which is why
          'it is declared as a Variant
          'Syntax: x.Remove(xyz)


10    On Error GoTo Remove_Error

20        mCol.Remove vntIndexKey

30    Exit Sub

Remove_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "IdentResults", "Remove", intEL, strES
    
End Sub

Public Sub Add(NewItem As IdentResult)

10    On Error GoTo Add_Error

20    mCol.Add NewItem

30    Exit Sub

Add_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "IdentResults", "Add", intEL, strES

End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
          'this property allows you to enumerate
          'this collection with the For...Each syntax
10    On Error GoTo NewEnum_Error

20        Set NewEnum = mCol.[_NewEnum]

30    Exit Property

NewEnum_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "IdentResults", "NewEnum", intEL, strES
End Property


Private Sub Class_Initialize()
          'creates the collection when this class is created
10        Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
          'destroys collection when this class is terminated
10        Set mCol = Nothing
End Sub

