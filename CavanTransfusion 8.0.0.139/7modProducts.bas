Attribute VB_Name = "modProducts"
Option Explicit

Public Function ProductBarCodeFor(ByVal strProductName As String) _
                                  As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ProductBarCodeFor_Error

20    sql = "Select * from ProductList where " & _
            "Wording = '" & strProductName & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60      ProductBarCodeFor = tb!BarCode
70    Else
80      ProductBarCodeFor = "???"
90    End If

100   Exit Function

ProductBarCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modProducts", "ProductBarCodeFor", intEL, strES, sql


End Function


Public Function ProductWordingFor(ByVal BarCode As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ProductWordingFor_Error

20    sql = "Select Wording from ProductList where " & _
            "BarCode = '" & BarCode & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60      ProductWordingFor = tb!Wording
70    Else
80      ProductWordingFor = "???"
90    End If

100   Exit Function

ProductWordingFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modProducts", "ProductWordingFor", intEL, strES, sql


End Function


Public Function ProductGenericFor(ByVal BarCode As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ProductGenericFor_Error

20    sql = "Select Generic from ProductList where " & _
            "BarCode = '" & BarCode & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60      ProductGenericFor = tb!Generic
70    Else
80      ProductGenericFor = "???"
90    End If

100   Exit Function

ProductGenericFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modProducts", "ProductGenericFor", intEL, strES, sql


End Function


Public Function ProductGenericForWording(ByVal Wording As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ProductGenericForWording_Error

20    sql = "Select Generic from ProductList where " & _
            "Wording = '" & Wording & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60      ProductGenericForWording = tb!Generic
70    Else
80      ProductGenericForWording = "???"
90    End If

100   Exit Function

ProductGenericForWording_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modProducts", "ProductGenericForWording", intEL, strES, sql

End Function



