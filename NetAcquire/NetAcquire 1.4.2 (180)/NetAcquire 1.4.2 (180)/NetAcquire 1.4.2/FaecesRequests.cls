VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaecesRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolFaecesRequests As New Collection

Private mFaecesRequests As New FaecesRequests

Public Sub Add(ByVal Fx As FaecesRequest)

12070     mcolFaecesRequests.Add Fx, Fx.Request

End Sub


Public Function Count() As Long
        
12080     Count = mcolFaecesRequests.Count

End Function




Public Sub Delete(ByVal SampleID As String, ByVal Request As String)
        
          Dim sql As String

12090     On Error GoTo Delete_Error

12100     sql = "DELETE FROM FaecesRequests50 WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND Request = '" & Request & "'"
12110     Cnxn(0).Execute sql

12120     mcolFaecesRequests.Remove Request

12130     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

12140     intEL = Erl
12150     strES = Err.Description
12160     LogError "FaecesRequests", "Delete", intEL, strES, sql

End Sub

Public Property Get FaecesRequests() As FaecesRequests
        
12170     Set FaecesRequests = mFaecesRequests

End Property


Public Function Item(ByVal Request As String) As FaecesRequest
        
12180     On Error Resume Next

12190     Set Item = mcolFaecesRequests.Item(Request)

End Function





Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Fx As FaecesRequest

12200     On Error GoTo Load_Error

12210     sql = "SELECT SampleID, Request, UserName, DateTimeOfRecord, Analyser, COALESCE(Programmed, 0) Programmed " & _
              "FROM FaecesRequests50 " & _
              "WHERE SampleID = '" & SampleID & "'"
12220     Set tb = New Recordset
12230     RecOpenClient 0, tb, sql
12240     Do While Not tb.EOF
12250         Set Fx = New FaecesRequest
12260         With Fx
12270             .SampleID = tb!SampleID & ""
12280             .Request = tb!Request & ""
12290             .UserName = tb!UserName & ""
12300             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
12310             .Analyser = tb!Analyser & ""
12320             .Programmed = CBool(tb!Programmed)
12330             mcolFaecesRequests.Add Fx, .Request
12340         End With
12350         tb.MoveNext
12360     Loop

12370     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

12380     intEL = Erl
12390     strES = Err.Description
12400     LogError "FaecesRequests", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
12410     Set NewEnum = mcolFaecesRequests.[_NewEnum]

End Function




Public Sub Save(ByVal SampleID As String)

          Dim sql As String
          Dim Fx As FaecesRequest

12420     On Error GoTo Save_Error

12430     sql = "Delete from FaecesRequests50 where " & _
              "SampleID = '" & SampleID & "'"
12440     Cnxn(0).Execute sql

12450     For Each Fx In mcolFaecesRequests

12460         sql = "INSERT INTO FaecesRequests50 (SampleID, Request, UserName, Analyser, Programmed) " & _
                  "VALUES " & _
                  "('" & Fx.SampleID & "', " & _
                  " '" & Fx.Request & "', " & _
                  " '" & Fx.UserName & "', " & _
                  " '" & Fx.Analyser & "', " & _
                  " '" & Fx.Programmed & "')"
12470         Cnxn(0).Execute sql
12480     Next

12490     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

12500     intEL = Erl
12510     strES = Err.Description
12520     LogError "FaecesRequests", "Save", intEL, strES, sql

End Sub




