Attribute VB_Name = "modLists"
Option Explicit

Public Function ListCodeFor(ByVal ListType As String, ByVal Text As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ListCodeFor_Error

20    ListCodeFor = ""
30    Text = UCase$(Trim$(Text))

40    sql = "Select * from Lists where " & _
            "ListType = '" & ListType & "' " & _
            "and Text = '" & AddTicks(Text) & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80      ListCodeFor = tb!code
90    End If

100   Exit Function

ListCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modLists", "ListCodeFor", intEL, strES, sql


End Function

Public Function BBListCodeFor(ByVal ListType As String, ByVal Text As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo BBListCodeFor_Error

20    BBListCodeFor = ""
30    Text = UCase$(Trim$(Text))

40    sql = "Select * from Lists where " & _
            "ListType = '" & ListType & "' " & _
            "and Text = '" & AddTicks(Text) & "'"
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    If Not tb.EOF Then
80      BBListCodeFor = tb!code & ""
90    End If

100   Exit Function

BBListCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modLists", "BBListCodeFor", intEL, strES, sql


End Function

Public Function BBListTextFor(ByVal ListType As String, ByVal code As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo BBListTextFor_Error

20    BBListTextFor = ""
30    code = UCase$(Trim$(code))

40    sql = "Select * from Lists where " & _
            "ListType = '" & ListType & "' " & _
            "and Code = '" & AddTicks(code) & "'"
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    If Not tb.EOF Then
80      BBListTextFor = tb!Text & ""
90    End If

100   Exit Function

BBListTextFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modLists", "BBListTextFor", intEL, strES, sql


End Function



Public Function ListTextFor(ByVal ListType As String, ByVal code As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ListTextFor_Error

20    ListTextFor = ""
30    code = UCase$(Trim$(code))

40    sql = "Select * from Lists where " & _
            "ListType = '" & ListType & "' " & _
            "and Code = '" & AddTicks(code) & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80      ListTextFor = tb!Text & ""
90    End If

100   Exit Function

ListTextFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modLists", "ListTextFor", intEL, strES, sql


End Function


