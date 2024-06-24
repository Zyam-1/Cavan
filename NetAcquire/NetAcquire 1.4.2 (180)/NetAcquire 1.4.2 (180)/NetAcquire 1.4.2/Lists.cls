VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolLists As New Collection

Private mlists As New Lists

Public Sub Add(ByVal Lx As List)

21920     mcolLists.Add Lx, Lx.ListType & " " & Lx.Code

End Sub

Public Function Count() As Long
        
21930     Count = mcolLists.Count

End Function

Public Sub Delete(ByVal ListType As String, ByVal Code As String)
        
          Dim sql As String

21940     On Error GoTo Delete_Error

21950     sql = "DELETE FROM Lists WHERE " & _
              "ListType = '" & ListType & "' " & _
              "AND Code = '" & Code & "'"
21960     Cnxn(0).Execute sql

21970     mcolLists.Remove ListType & " " & Code

21980     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

21990     intEL = Erl
22000     strES = Err.Description
22010     LogError "Lists", "Delete", intEL, strES, sql

End Sub
Public Function Item(ByVal ListType As String, ByVal Code As String) As List
        
22020     On Error Resume Next
22030     Set Item = mcolLists.Item(ListType & " " & Code)

End Function

Public Sub Load(ByVal ListType As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Lx As List

22040     On Error GoTo Load_Error

22050     sql = "SELECT Code, COALESCE(Text, '') Text, " & _
              "COALESCE(ListOrder, 999) ListOrder, " & _
              "COALESCE(InUse, 0) InUse, " & _
              "COALESCE([Default], '') DefaultValue " & _
              "FROM Lists " & _
              "WHERE ListType = '" & ListType & "'"
22060     Set tb = New Recordset
22070     RecOpenClient 0, tb, sql
22080     Do While Not tb.EOF
22090         Set Lx = New List
22100         With Lx
22110             .Code = tb!Code & ""
22120             .Text = tb!Text
22130             .ListType = ListType
22140             .ListOrder = tb!ListOrder
22150             .InUse = tb!InUse
22160             .DefaultValue = tb!DefaultValue
22170             mcolLists.Add Lx, .ListType & " " & .Code
22180         End With
22190         tb.MoveNext
22200     Loop

22210     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

22220     intEL = Erl
22230     strES = Err.Description
22240     LogError "Lists", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
22250     Set NewEnum = mcolLists.[_NewEnum]

End Function



Public Sub Save()

          Dim sql As String
          Dim Lx As List

22260     On Error GoTo Save_Error

22270     For Each Lx In mcolLists

22280         sql = "IF EXISTS (SELECT * FROM Lists " & _
                  "           WHERE ListType = '" & Lx.ListType & "' " & _
                  "           AND Code = '" & Lx.Code & "') " & _
                  "  UPDATE Lists " & _
                  "  SET Text = '" & AddTicks(Lx.Text) & "', " & _
                  "  ListOrder = '" & Lx.ListOrder & "', " & _
                  "  InUse = '" & Lx.InUse & "', " & _
                  "  Default = '" & Lx.DefaultValue & "' " & _
                  "  WHERE ListType = '" & Lx.ListType & "' " & _
                  "  AND Code = '" & Lx.Code & "' " & _
                  "ELSE " & _
                  "  INSERT INTO Lists (Code, Text, ListType, ListOrder, InUse, Default) " & _
                  "  VALUES " & _
                  "  ('" & Lx.Code & "', " & _
                  "  '" & AddTicks(Lx.Text) & "', " & _
                  "  '" & Lx.ListType & "', )" & _
                  "  '" & Lx.ListOrder & "', " & _
                  "  '" & Lx.InUse & "', " & _
                  "  '" & Lx.DefaultValue & "')"
22290         Cnxn(0).Execute sql
22300     Next

22310     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

22320     intEL = Erl
22330     strES = Err.Description
22340     LogError "Lists", "Save", intEL, strES, sql

End Sub



Public Property Get Lists() As Lists
        
22350     Set Lists = mlists

End Property




