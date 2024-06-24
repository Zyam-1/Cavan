VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UrineRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolUrineRequests As New Collection

Private mUrineRequests As New UrineRequests

Public Sub Add(ByVal Fx As UrineRequest)

27280     mcolUrineRequests.Add Fx, Fx.Request

End Sub

Public Function Count() As Long
        
27290     Count = mcolUrineRequests.Count

End Function

Public Sub Delete(ByVal SampleID As String, ByVal Request As String)
        
          Dim sql As String

27300     On Error GoTo Delete_Error

27310     sql = "DELETE FROM UrineRequests50 WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND Request = '" & Request & "'"
27320     Cnxn(0).Execute sql

27330     mcolUrineRequests.Remove Request

27340     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

27350     intEL = Erl
27360     strES = Err.Description
27370     LogError "UrineRequests", "Delete", intEL, strES, sql

End Sub
Public Property Get UrineRequests() As UrineRequests
        
27380     Set UrineRequests = mUrineRequests

End Property



Public Function Item(ByVal Request As String) As UrineRequest
        


          Dim i As Integer
27390     On Error GoTo Item_Error

27400     Set Item = Nothing
27410     For i = 1 To mcolUrineRequests.Count
27420         If mcolUrineRequests.Item(i).Request = Request Then
27430             Set Item = mcolUrineRequests(i)
27440         End If
27450     Next i

27460     Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer

27470     intEL = Erl
27480     strES = Err.Description
27490     LogError "UrineRequests", "Item", intEL, strES
          

End Function


Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Fx As UrineRequest

27500     On Error GoTo Load_Error

27510     sql = "SELECT SampleID, Request, UserName, DateTimeOfRecord " & _
              "FROM UrineRequests50 " & _
              "WHERE SampleID = '" & SampleID & "'"
27520     Set tb = New Recordset
27530     RecOpenClient 0, tb, sql
27540     Do While Not tb.EOF
27550         Set Fx = New UrineRequest
27560         With Fx
27570             .SampleID = tb!SampleID & ""
27580             .Request = tb!Request & ""
27590             .UserName = tb!UserName & ""
27600             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
27610             mcolUrineRequests.Add Fx, .Request
27620         End With
27630         tb.MoveNext
27640     Loop

27650     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

27660     intEL = Erl
27670     strES = Err.Description
27680     LogError "UrineRequests", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
27690     Set NewEnum = mcolUrineRequests.[_NewEnum]

End Function

Public Sub Save(ByVal SampleID As String)

          Dim sql As String
          Dim Fx As UrineRequest

27700     On Error GoTo Save_Error

27710     sql = "Delete from UrineRequests50 where " & _
              "SampleID = '" & SampleID & "'"
27720     Cnxn(0).Execute sql

27730     For Each Fx In mcolUrineRequests

27740         sql = "INSERT INTO UrineRequests50 (SampleID, Request, UserName) " & _
                  "VALUES " & _
                  "('" & Fx.SampleID & "', " & _
                  " '" & Fx.Request & "', " & _
                  " '" & AddTicks(Fx.UserName) & "')"
27750         Cnxn(0).Execute sql
27760     Next

27770     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

27780     intEL = Erl
27790     strES = Err.Description
27800     LogError "UrineRequests", "Save", intEL, strES, sql

End Sub





