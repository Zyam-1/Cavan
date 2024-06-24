VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SemenResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolSemenResults As New Collection

Private mSemenResults As New SemenResults

Public Function Count() As Long
        
23770     Count = mcolSemenResults.Count

End Function

Public Sub Delete(ByVal SampleID As String, ByVal TestName As String)
        
          Dim sql As String

23780     On Error GoTo Delete_Error

23790     sql = "DELETE FROM SemenResults50 WHERE " & _
              "SampleID = '" & SampleID & "' " & _
              "AND TestName = '" & TestName & "'"
23800     Cnxn(0).Execute sql

23810     mcolSemenResults.Remove TestName

23820     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

23830     intEL = Erl
23840     strES = Err.Description
23850     LogError "SemenResults", "Delete", intEL, strES, sql

End Sub

Public Property Get SemenResults() As SemenResults
        
23860     Set SemenResults = mSemenResults

End Property

Public Function Item(ByVal TestName As String) As SemenResult
Attribute Item.VB_UserMemId = 0
        
23870     On Error Resume Next

23880     Set Item = mcolSemenResults.Item(TestName)

End Function

Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim sx As SemenResult

23890     On Error GoTo Load_Error

23900     sql = "SELECT SampleID, TestName, Result, " & _
              "UserName, DateTimeOfRecord, COALESCE(Valid, 0) Valid, COALESCE(Printed, 0) Printed, " & _
              "ValidatedBy, ValidatedDateTime, PrintedBy, PrintedDateTime " & _
              "FROM SemenResults50 WHERE SampleID = '" & SampleID & "'"
23910     Set tb = New Recordset
23920     RecOpenClient 0, tb, sql
23930     Do While Not tb.EOF
23940         Set sx = New SemenResult
23950         With sx
23960             .SampleID = tb!SampleID & ""
23970             .TestName = tb!TestName & ""
23980             .Result = tb!Result & ""
23990             .Valid = tb!Valid
24000             .Printed = tb!Printed
24010             .ValidatedBy = tb!ValidatedBy & ""
24020             .ValidatedDateTime = tb!ValidatedDateTime & ""
24030             .PrintedBy = tb!PrintedBy & ""
24040             .PrintedDateTime = tb!PrintedDateTime & ""
24050             .UserName = tb!UserName & ""
24060             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
24070             mcolSemenResults.Add sx, .TestName
24080         End With
24090         tb.MoveNext
24100     Loop

24110     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

24120     intEL = Erl
24130     strES = Err.Description
24140     LogError "SemenResults", "Load", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
24150     Set NewEnum = mcolSemenResults.[_NewEnum]

End Function

Public Sub Save()

          Dim sx As SemenResult

24160     On Error GoTo Save_Error

24170     For Each sx In mcolSemenResults
24180         sx.Save
24190     Next

24200     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

24210     intEL = Erl
24220     strES = Err.Description
24230     LogError "SemenResults", "Save", intEL, strES

End Sub


