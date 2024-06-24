VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Isolates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolIsolates As New Collection

Private mIsolates As New Isolates

Public Sub Add(ByVal Iso As Isolate)

20250     mcolIsolates.Add Iso, CStr(Iso.IsolateNumber)

End Sub

Public Function Count() As Long
        
20260     Count = mcolIsolates.Count

End Function
Public Sub Delete(ByVal SampleID As String, ByVal IsolateNumber As Integer)
        
          Dim sql As String

20270     On Error GoTo Delete_Error

20280     sql = "DELETE FROM Isolates WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND IsolateNumber = '" & IsolateNumber & "'"
20290     Cnxn(0).Execute sql

20300     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

20310     intEL = Erl
20320     strES = Err.Description
20330     LogError "Isolates", "Delete", intEL, strES, sql

End Sub
Public Property Get Isolates() As Isolates
        
20340     Set Isolates = mIsolates

End Property


Public Function Item(ByVal IsolateNumber As Integer) As Isolate
Attribute Item.VB_UserMemId = 0
        
20350     On Error Resume Next

20360     Set Item = mcolIsolates.Item(CStr(IsolateNumber))

End Function
Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Ix As Isolate
          Dim l_Count As Integer

20370     On Error GoTo Load_Error

20380     sql = "SELECT I.SampleID, COALESCE(I.IsolateNumber, 1) IsolateNumber, " & _
              "COALESCE(I.OrganismGroup, '') OrganismGroup, " & _
              "COALESCE(I.OrganismName, '') OrganismName, " & _
              "COALESCE(I.Qualifier, '') Qualifier, " & _
              "COALESCE(I.UserName, '') UserName, " & _
              "COALESCE(I.RecordDateTime, '1/1/1900') RecordDateTime, " & _
              "COALESCE(PV.Valid, 0) Valid, " & _
              "COALESCE(PV.Printed, 0) Printed " & _
              "FROM Isolates I " & _
              "LEFT JOIN PrintValidLog PV ON I.SampleID = PV.SampleID " & _
              "WHERE I.SampleID = '" & SampleID & "' " & _
              "ORDER BY I.IsolateNumber"
20390     Set tb = New Recordset
20400     RecOpenClient 0, tb, sql
          '+++ Junaid 12-01-2024
20410     If mcolIsolates.Count > 0 Then
20420         For l_Count = 0 To mcolIsolates.Count - 1
20430             mcolIsolates.Remove (l_Count)
20440         Next
20450     End If
          '--- Junaid
20460     Do While Not tb.EOF
20470         Set Ix = New Isolate
20480         With Ix
20490             .SampleID = tb!SampleID & ""
20500             .IsolateNumber = tb!IsolateNumber
20510             .OrganismGroup = tb!OrganismGroup
20520             .OrganismName = tb!OrganismName
20530             .Qualifier = tb!Qualifier
20540             .UserName = tb!UserName
20550             .RecordDateTime = tb!RecordDateTime
20560             .Valid = tb!Valid
20570             .Printed = tb!Printed
20580             mcolIsolates.Add Ix, CStr(.IsolateNumber)
20590         End With
20600         tb.MoveNext
20610     Loop

20620     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

20630     intEL = Erl
20640     strES = Err.Description
20650     LogError "Isolates", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
20660     Set NewEnum = mcolIsolates.[_NewEnum]

End Function
Public Sub Save()

          Dim Ix As Isolate

20670     On Error GoTo Save_Error

20680     For Each Ix In mcolIsolates
20690         Ix.Save
20700     Next

20710     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

20720     intEL = Erl
20730     strES = Err.Description
20740     LogError "Isolates", "Save", intEL, strES

End Sub



