VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SexNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'In a shared module put the following declaration
'Public colSexNames As New SexNames

Private mcolSexNames As New Collection

Private mSexNames As New SexNames
Public Sub Refresh()

10    Set mcolSexNames = Nothing
20    Load

End Sub


Private Sub Class_Initialize()

10    Load

End Sub


Public Sub Add(ByVal sName As SexName)
  
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo Add_Error

20    sName.Name = Trim$(UCase$(sName.Name))

30    sql = "Select * from SexNames where " & _
            "Name = '" & AddTicks(sName.Name) & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql

60    With sName
70      If tb.EOF Then
80        tb.AddNew
90      End If

100     tb!Name = .Name
110     tb!Sex = .Sex
120     tb.Update
130   End With

140   Exit Sub

Add_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "SexNames", "Add", intEL, strES, sql

End Sub

Public Sub Update(ByVal fName As String, _
                  ByVal Sex As String)
  
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo Update_Error

20    If UCase$(Left$(Sex, 1)) <> "F" _
         And UCase$(Left$(Sex, 1)) <> "M" _
         And Sex <> "D" Then
30      Exit Sub
40    End If

50    fName = UCase$(Trim$(fName))
60    If fName = "" Then Exit Sub

70    sql = "Select * from SexNames where " & _
            "Name = '" & AddTicks(fName) & "'"
80    Set tb = New Recordset
90    RecOpenClient 0, tb, sql

100   If tb.EOF Then
110     tb.AddNew
120   Else
130     mcolSexNames.Remove CStr(fName)
140   End If

150   tb!Name = fName
160   tb!Sex = UCase$(Left$(Sex, 1))
170   tb.Update
180   mcolSexNames.Add fName, CStr(fName)

190   Exit Sub

Update_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "SexNames", "Update", intEL, strES, sql

End Sub


Public Property Get SexNames() As SexNames
  
10    Set SexNames = mSexNames

End Property


Public Function Count() As Long
  
10    Count = mcolSexNames.Count

End Function


Public Sub Delete(ByVal sName As String)
  
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo Delete_Error

20    sql = "Select * from SexNames where " & _
            "Name = '" & AddTicks(sName) & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60      tb.Delete
70      mcolSexNames.Remove CStr(sName)
80    End If

90    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "SexNames", "Delete", intEL, strES, sql

End Sub


Public Function Item(ByVal sName As String) As SexName
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolSexNames.Item(UCase$(CStr(sName)))

End Function


Private Sub Load()

      Dim tb As Recordset
      Dim s As SexName
      Dim sql As String

10    On Error GoTo Load_Error

20    sql = "Select * from SexNames"
30    Set tb = New Recordset

40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set s = New SexName
70      With s
80        .Name = tb!Name & ""
90        .Sex = tb!Sex & ""
100       mcolSexNames.Add s, UCase$(CStr(.Name))
110     End With
120     tb.MoveNext
130   Loop

140   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "SexNames", "Load", intEL, strES, sql

End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolSexNames.[_NewEnum]

End Function





