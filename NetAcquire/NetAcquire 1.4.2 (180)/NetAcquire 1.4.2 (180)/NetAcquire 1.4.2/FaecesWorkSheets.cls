VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaecesWorkSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolFaecesWorkSheets As New Collection

Private mFaecesWorkSheets As New FaecesWorkSheets

Public Function Count() As Long
        
13310     Count = mcolFaecesWorkSheets.Count

End Function
Public Sub Delete(ByVal SampleID As String, ByVal Dayindex As String)
        
          Dim sql As String

13320     On Error GoTo Delete_Error

13330     sql = "DELETE FROM FaecesWorkSheets50 WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND DayIndex = '" & Dayindex & "'"
13340     Cnxn(0).Execute sql

13350     mcolFaecesWorkSheets.Remove Dayindex

13360     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

13370     intEL = Erl
13380     strES = Err.Description
13390     LogError "FaecesWorkSheets", "Delete", intEL, strES, sql

End Sub

Public Property Get FaecesWorkSheets() As FaecesWorkSheets
        
13400     Set FaecesWorkSheets = mFaecesWorkSheets

End Property

Public Function Item(ByVal Dayindex As String) As FaecesWorkSheet
Attribute Item.VB_UserMemId = 0
        
13410     On Error Resume Next

13420     Set Item = mcolFaecesWorkSheets.Item(Dayindex)

End Function

Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Fx As FaecesWorkSheet

13430     On Error GoTo Load_Error

13440     sql = "SELECT F.SampleID, F.DayIndex, F.Result, F.UserName, F.DateTimeOfRecord " & _
              "FROM FaecesWorkSheets50 F " & _
              "WHERE F.SampleID = '" & SampleID & "'"
13450     Set tb = New Recordset
13460     RecOpenClient 0, tb, sql
13470     Do While Not tb.EOF
13480         Set Fx = New FaecesWorkSheet
13490         With Fx
13500             .SampleID = tb!SampleID & ""
13510             .Dayindex = tb!Dayindex & ""
13520             .Result = tb!Result & ""
13530             .UserName = tb!UserName & ""
13540             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
13550             mcolFaecesWorkSheets.Add Fx, .Dayindex
13560         End With
13570         tb.MoveNext
13580     Loop

13590     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

13600     intEL = Erl
13610     strES = Err.Description
13620     LogError "FaecesWorkSheets", "Load", intEL, strES, sql

End Sub
Public Sub LoadAudit(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Fx As FaecesWorkSheet

13630     On Error GoTo LoadAudit_Error

13640     sql = "SELECT F.SampleID, F.DayIndex, F.Result, F.UserName, F.DateTimeOfRecord " & _
              "FROM FaecesWorkSheets50Audit F " & _
              "WHERE F.SampleID = '" & SampleID & "' " & _
              "ORDER BY F.DateTimeOfRecord desc"
13650     Set tb = New Recordset
13660     RecOpenClient 0, tb, sql
13670     Do While Not tb.EOF
13680         Set Fx = New FaecesWorkSheet
13690         With Fx
13700             .SampleID = tb!SampleID & ""
13710             .Dayindex = tb!Dayindex & ""
13720             .Result = tb!Result & ""
13730             .UserName = tb!UserName & ""
13740             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
13750             mcolFaecesWorkSheets.Add Fx
13760         End With
13770         tb.MoveNext
13780     Loop

13790     Exit Sub

LoadAudit_Error:

          Dim strES As String
          Dim intEL As Integer

13800     intEL = Erl
13810     strES = Err.Description
13820     LogError "FaecesWorkSheets", "LoadAudit", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
13830     Set NewEnum = mcolFaecesWorkSheets.[_NewEnum]

End Function

Public Sub Save()

          Dim Fx As FaecesWorkSheet

13840     On Error GoTo Save_Error

13850     For Each Fx In mcolFaecesWorkSheets
13860         Fx.Save
13870     Next

13880     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

13890     intEL = Erl
13900     strES = Err.Description
13910     LogError "FaecesWorkSheets", "Save", intEL, strES

End Sub


