VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UrineIdents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolUrineIdents As New Collection

Private mUrineIdents As New UrineIdents

Public Function Count() As Long
        
26670     Count = mcolUrineIdents.Count

End Function

Public Sub Delete(ByVal SampleID As String, _
          ByVal TestName As String, _
          ByVal Isolate As Integer)
        
          Dim sql As String

26680     On Error GoTo Delete_Error

26690     sql = "DELETE FROM UrineIdents50 WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND TestName = '" & TestName & "' " & _
              "AND Isolate = '" & Isolate & "'"
26700     Cnxn(0).Execute sql

26710     mcolUrineIdents.Remove TestName & Format$(Isolate)

26720     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

26730     intEL = Erl
26740     strES = Err.Description
26750     LogError "UrineIdents", "Delete", intEL, strES, sql

End Sub

Public Function Item(ByVal TestName As String, ByVal Isolate As Integer) As UrineIdent
        
          '10    On Error Resume Next
          'Set Item = mcolUrineIdents.Item(TestName & Format$(Isolate))
          Dim i As Integer


26760     Set Item = Nothing
26770     For i = 1 To mcolUrineIdents.Count
26780         If mcolUrineIdents.Item(i).TestName = TestName And mcolUrineIdents.Item(i).Isolate = Isolate Then
26790             Set Item = mcolUrineIdents(i)
26800         End If
26810     Next i



End Function

Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Ix As UrineIdent

26820     On Error GoTo Load_Error

26830     sql = "SELECT SampleID, TestName, Isolate, Result, " & _
              "UserName, DateTimeOfRecord " & _
              "FROM UrineIdent50 " & _
              "WHERE SampleID = '" & SampleID & "'"
26840     Set tb = New Recordset
26850     RecOpenClient 0, tb, sql
26860     Do While Not tb.EOF
26870         Set Ix = New UrineIdent
26880         With Ix
26890             .SampleID = tb!SampleID & ""
26900             .TestName = tb!TestName & ""
26910             .Isolate = tb!Isolate
26920             .Result = tb!Result & ""
26930             .UserName = tb!UserName & ""
26940             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
26950             mcolUrineIdents.Add Ix, .TestName & Format$(.Isolate)
26960         End With
26970         tb.MoveNext
26980     Loop

26990     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

27000     intEL = Erl
27010     strES = Err.Description
27020     LogError "UrineIdents", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
27030     Set NewEnum = mcolUrineIdents.[_NewEnum]

End Function




Public Sub Save(ByVal Ix As UrineIdent)

          Dim sql As String

27040     On Error GoTo Save_Error

27050     If Trim$(Ix.Result) = "" Then
              '30      Sql = "DELETE FROM UrineIdent50 WHERE " & _
              '              "SampleID = '" & Ix.SampleID & "' " & _
              '              "AND TestName = '" & Ix.TestName & "' " & _
              '              "AND Isolate = '" & Ix.Isolate & "'"
              '        Exit Sub
27060     Else
27070         sql = "IF EXISTS (SELECT * FROM UrineIdent50 WHERE " & _
                  "           SampleID = '" & Ix.SampleID & "' " & _
                  "           AND TestName = '" & Ix.TestName & "' " & _
                  "           AND Isolate = '" & Ix.Isolate & "') " & _
                  "  UPDATE UrineIdent50 " & _
                  "  SET Result = '" & AddTicks(Ix.Result) & "', " & _
                  "  UserName = '" & AddTicks(Ix.UserName) & "' " & _
                  "  WHERE SampleID = '" & Ix.SampleID & "' " & _
                  "  AND TestName = '" & Ix.TestName & "' " & _
                  "  AND Isolate = '" & Ix.Isolate & "' " & _
                  "ELSE " & _
                  "  INSERT INTO UrineIdent50 (SampleID, TestName, Result, Isolate, UserName,DateTimeOfRecord) " & _
                  "  VALUES " & _
                  "  ('" & Ix.SampleID & "', " & _
                  "  '" & Ix.TestName & "', " & _
                  "  '" & AddTicks(Ix.Result) & "', " & _
                  "  '" & Ix.Isolate & "', " & _
                  "  '" & AddTicks(Ix.UserName) & "', GetDate())"
27080         Cnxn(0).Execute sql: Call WriteToFile_Execution(sql)
27090     End If


27100     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

27110     intEL = Erl
27120     strES = Err.Description
27130     LogError "UrineIdents", "Save", intEL, strES, sql

End Sub
Public Property Get UrineIdents() As UrineIdents
        
27140     Set UrineIdents = mUrineIdents

End Property

Public Sub Clear()
          Dim Ix As UrineIdent
          
27150     Set Ix = New UrineIdent
27160     With Ix
27170         .SampleID = ""
27180         .TestName = ""
27190         .Isolate = 0
27200         .Result = ""
27210         .UserName = ""
27220         .DateTimeOfRecord = ""
27230     End With
          
27240     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer
          
27250     intEL = Erl
27260     strES = Err.Description
27270     LogError "UrineIdents", "Clear", intEL, strES
End Sub
