VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserRoleCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mCol As Collection

Public Function Add(NewUserRole As UserRole, Optional sKey As String) As UserRole
          Dim objNewMember As UserRole
31320     On Error GoTo Add_Error

31330     Set objNewMember = New UserRole
31340     Set objNewMember = NewUserRole
31350     If Len(sKey) = 0 Then
31360         mCol.Add objNewMember
31370     Else
31380         mCol.Add objNewMember, sKey
31390     End If
31400     Set Add = objNewMember
31410     Set objNewMember = Nothing

31420     Exit Function

Add_Error:

          Dim strES As String
          Dim intEL As Integer

31430     intEL = Erl
31440     strES = Err.Description
31450     LogError "UserRoleCollection", "Add", intEL, strES

End Function
Public Property Get Item(vntIndexKey As Variant) As UserRole
31460     Set Item = mCol(vntIndexKey)
End Property
Public Property Get Count() As Long
31470     Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
31480     mCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
31490     Set NewEnum = mCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
31500     Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
31510     Set mCol = Nothing
31520     Set mvarFaecesMidlandsCollection = Nothing
End Sub

