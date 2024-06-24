VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LIHs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolLIHs As New Collection

Private mLIHs As New LIHs

Public Function Count() As Long
        
20940     Count = mcolLIHs.Count

End Function

Public Property Get LIHs() As LIHs
        
20950     Set LIHs = mLIHs

End Property

Public Function Item(ByVal LIorH As String, ByVal Code As String, ByVal NoPrintOrWarning As String) As LIH
        
          Dim i As Integer
20960     On Error GoTo Item_Error

20970     LoadMask LIorH, Code, NoPrintOrWarning
20980     Set Item = Nothing
20990     For i = 1 To mcolLIHs.Count
21000         If mcolLIHs.Item(i).LIorH = LIorH And mcolLIHs.Item(i).Code = Code And mcolLIHs.Item(i).NoPrintOrWarning = NoPrintOrWarning Then
21010             Set Item = mcolLIHs(i)
21020         End If
         
21030     Next i


21040     Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer

21050     intEL = Erl
21060     strES = Err.Description
21070     LogError "LIHs", "Item", intEL, strES
          

End Function

Public Sub Load()

          Dim tb As Recordset
          Dim sql As String
          Dim Lx As LIH

21080     On Error GoTo Load_Error

21090     Set mcolLIHs = New Collection

21100     sql = "SELECT LIH, Code, CutOff, NoPrintOrWarning, " & _
              "UserName, DateTimeOfRecord " & _
              "FROM LIHValues"
21110     Set tb = New Recordset
21120     RecOpenClient 0, tb, sql
21130     Do While Not tb.EOF
21140         Set Lx = New LIH
21150         With Lx
21160             .LIorH = tb!LIH
21170             .Code = tb!Code
21180             .CutOff = tb!CutOff
21190             .NoPrintOrWarning = tb!NoPrintOrWarning
21200             .UserName = tb!UserName
21210             .DateTimeOfRecord = tb!DateTimeOfRecord
21220             mcolLIHs.Add Lx, .LIorH & CStr(.Code) & .NoPrintOrWarning
21230         End With
21240         tb.MoveNext
21250     Loop

21260     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

21270     intEL = Erl
21280     strES = Err.Description
21290     LogError "LIHs", "Load", intEL, strES, sql

End Sub

Public Sub LoadMask(ByVal LIorH As String, ByVal Code As String, ByVal NoPrintOrWarning As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Lx As LIH

21300     On Error GoTo Load_Error

21310     Set mcolLIHs = New Collection

21320     sql = "SELECT LIH, Code, CutOff, NoPrintOrWarning, " & _
              "UserName, DateTimeOfRecord " & _
              "FROM LIHValues " & _
              "WHERE LIH = '" & LIorH & "' AND Code = '" & Code & "' AND NoPrintOrWarning = '" & NoPrintOrWarning & "'"
21330     Set tb = New Recordset
21340     RecOpenClient 0, tb, sql
21350     Do While Not tb.EOF
21360         Set Lx = New LIH
21370         With Lx
21380             .LIorH = tb!LIH
21390             .Code = tb!Code
21400             .CutOff = tb!CutOff
21410             .NoPrintOrWarning = tb!NoPrintOrWarning
21420             .UserName = tb!UserName
21430             .DateTimeOfRecord = tb!DateTimeOfRecord
21440             mcolLIHs.Add Lx, .LIorH & CStr(.Code) & .NoPrintOrWarning
21450         End With
21460         tb.MoveNext
21470     Loop

21480     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

21490     intEL = Erl
21500     strES = Err.Description
21510     LogError "LIHs", "Load", intEL, strES, sql

End Sub

Public Sub Clear()

          Dim n As Integer

21520     For n = 1 To mcolLIHs.Count
21530         mcolLIHs.Remove 1
21540     Next

End Sub


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
21550     Set NewEnum = mcolLIHs.[_NewEnum]

End Function




Public Sub Save()

21560     On Error GoTo Save_Error

          Dim Lx As LIH

21570     For Each Lx In mcolLIHs
21580         Lx.Save
21590     Next

21600     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

21610     intEL = Erl
21620     strES = Err.Description
21630     LogError "LIHs", "Save", intEL, strES

End Sub




Private Sub Class_Initialize()

    'Load

End Sub


