Attribute VB_Name = "modSupplier"
Option Explicit

Public Function SupplierCodeFor(ByVal Supplier As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SupplierCodeFor_Error

20    SupplierCodeFor = "???"

30    sql = "Select BarCode from Supplier where " & _
            "Supplier = '" & Supplier & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70      SupplierCodeFor = tb!BarCode & ""
80    End If

90    Exit Function

SupplierCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modSupplier", "SupplierCodeFor", intEL, strES, sql


End Function

Public Function SupplierNameFor(ByVal BarCode As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SupplierNameFor_Error

20    SupplierNameFor = "???"

30    sql = "Select Supplier from Supplier where " & _
            "BarCode = '" & BarCode & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70      SupplierNameFor = tb!Supplier & ""
80    End If

90    Exit Function

SupplierNameFor_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modSupplier", "SupplierNameFor", intEL, strES, sql


End Function



