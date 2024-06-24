Attribute VB_Name = "modLists"
Option Explicit

Public Function ListCodeFor(ByVal ListType As String, ByVal Text As String) As String

      Dim tb As Recordset
      Dim sql As String

1910  On Error GoTo ListCodeFor_Error

1920  ListCodeFor = ""
1930  Text = UCase$(Trim$(Text))
      '+++ Junaid
1940  sql = "Select * from Lists where " & _
            "ListType = '" & ListType & "' " & _
            "and Text = '" & AddTicks(Text) & "' and InUse = 1"
      '40     Sql = "Select * from Lists where " & _
      '            "ListType = '" & ListType & "' " & _
      '            "and InUse = 1"
      '--- Junaid
1950  Set tb = New Recordset
1960  RecOpenServer 0, tb, sql
1970  If Not tb.EOF Then
1980    ListCodeFor = tb!Code
1990  End If

2000  Exit Function

ListCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

2010  intEL = Erl
2020  strES = Err.Description
2030  LogError "modLists", "ListCodeFor", intEL, strES, sql


End Function

Public Function BBListCodeFor(ByVal ListType As String, ByVal Text As String) As String

      Dim tb As Recordset
      Dim sql As String

2040  On Error GoTo BBListCodeFor_Error

2050  BBListCodeFor = ""
2060  Text = UCase$(Trim$(Text))

2070  sql = "Select * from Lists where " & _
            "ListType = '" & ListType & "' " & _
            "and Text = '" & AddTicks(Text) & "' and InUse = 1"
2080  Set tb = New Recordset
2090  RecOpenServerBB 0, tb, sql
2100  If Not tb.EOF Then
2110    BBListCodeFor = tb!Code & ""
2120  End If

2130  Exit Function

BBListCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

2140  intEL = Erl
2150  strES = Err.Description
2160  LogError "modLists", "BBListCodeFor", intEL, strES, sql


End Function

Public Function BBListTextFor(ByVal ListType As String, ByVal Code As String) As String

      Dim tb As Recordset
      Dim sql As String

2170  On Error GoTo BBListTextFor_Error

2180  BBListTextFor = ""
2190  Code = UCase$(Trim$(Code))

2200  sql = "Select * from Lists where " & _
            "ListType = '" & ListType & "' " & _
            "and Code = '" & AddTicks(Code) & "' and InUse = 1"
2210  Set tb = New Recordset
2220  RecOpenServerBB 0, tb, sql
2230  If Not tb.EOF Then
2240    BBListTextFor = tb!Text & ""
2250  End If

2260  Exit Function

BBListTextFor_Error:

      Dim strES As String
      Dim intEL As Integer

2270  intEL = Erl
2280  strES = Err.Description
2290  LogError "modLists", "BBListTextFor", intEL, strES, sql


End Function



Public Function ListTextFor(ByVal ListType As String, ByVal Code As String) As String

      Dim tb As Recordset
      Dim sql As String

2300  On Error GoTo ListTextFor_Error

2310  ListTextFor = ""
2320  Code = UCase$(Trim$(Code))

2330  sql = "SELECT * FROM Lists " & _
            "WHERE ListType = '" & ListType & "' " & _
            "AND Code = '" & AddTicks(Code) & "' " & _
            "AND InUse = 1"
2340  Set tb = New Recordset
2350  RecOpenServer 0, tb, sql
2360  If Not tb.EOF Then
2370    ListTextFor = tb!Text & ""
2380  End If

2390  Exit Function

ListTextFor_Error:

      Dim strES As String
      Dim intEL As Integer

2400  intEL = Erl
2410  strES = Err.Description
2420  LogError "modLists", "ListTextFor", intEL, strES, sql

End Function


