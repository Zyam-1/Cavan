Attribute VB_Name = "modDatabase"
Option Explicit

Public Function IsTableInDatabase(ByVal TableName As String) As Boolean

      Dim tbExists As Recordset
      Dim sql As String
      Dim retval As Boolean

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist
      'if it has a record then the table does exist.

10    On Error GoTo IsTableInDatabase_Error

20    sql = "SELECT name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = '" & TableName & "'"
30    Set tbExists = Cnxn(0).Execute(sql)

40    retval = True

50    If tbExists.EOF Then 'There is no table <TableName> in database
60      retval = False
70    End If
80    IsTableInDatabase = retval

90    Exit Function

IsTableInDatabase_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modDatabase", "IsTableInDatabase", intEL, strES, sql
  
End Function

Private Sub EnsureEntriesExist(ByVal ListType As String, ByVal code As String, ByVal Text As String)

      Dim sql As String

10    On Error GoTo EnsureEntriesExist_Error

20    sql = "IF NOT EXISTS ( SELECT * FROM Lists " & _
            "                WHERE ListType = '" & ListType & "' " & _
            "                AND Code = '" & code & "' " & _
            "                AND Text = '" & Text & "') " & _
            "  INSERT INTO Lists (Code, Text, ListType) VALUES " & _
            "  ('" & code & "', " & _
            "  '" & Text & "', " & _
            "  '" & ListType & "')"
30    CnxnBB(0).Execute sql

40    Exit Sub

EnsureEntriesExist_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "modDatabase", "EnsureEntriesExist", intEL, strES, sql

End Sub

Public Sub EnsureEventCodesInDatabase()

10          EnsureEntriesExist "EventBarCodes", "A", "Allocated"
20          EnsureEntriesExist "EventBarCodes", "B", "Being Transfused"
30          EnsureEntriesExist "EventBarCodes", "C", "Received into Stock"
40          EnsureEntriesExist "EventBarCodes", "D", "Destroyed"
50          EnsureEntriesExist "EventBarCodes", "E", "Amendment"
60          EnsureEntriesExist "EventBarCodes", "F", "Unit Dispatched"
70          EnsureEntriesExist "EventBarCodes", "G", "Received into Emergency Stock"
80          EnsureEntriesExist "EventBarCodes", "H", "Issued as Emergency"
90          EnsureEntriesExist "EventBarCodes", "I", "Issued"
100         EnsureEntriesExist "EventBarCodes", "J", "Expired"
110         EnsureEntriesExist "EventBarCodes", "K", "Awaiting Release"
120         EnsureEntriesExist "EventBarCodes", "L", "Labeled for Emergency ONeg"
130         EnsureEntriesExist "EventBarCodes", "M", "Moved to Emergency ONeg"
140         EnsureEntriesExist "EventBarCodes", "N", "Transferred"
150         EnsureEntriesExist "EventBarCodes", "P", "Pending"
160         EnsureEntriesExist "EventBarCodes", "R", "Restocked"
170         EnsureEntriesExist "EventBarCodes", "S", "Transfused"
180         EnsureEntriesExist "EventBarCodes", "T", "Returned to Supplier"
185         EnsureEntriesExist "EventBarCodes", "V", "E Issued"
190         EnsureEntriesExist "EventBarCodes", "X", "Cross matched"
200         EnsureEntriesExist "EventBarCodes", "Y", "Removed Pending Transfusion"
210         EnsureEntriesExist "EventBarCodes", "Z", "Transfused as Emergency"
220         EnsureEntriesExist "EventBarCodes", "W", "Blocked - Group Check failed"

End Sub

Public Sub EnsureGroupBarCodesInDatabase()
            '
            '10    EnsureEntriesExist "GroupBarCodes", "51", "O Pos"
            '20    EnsureEntriesExist "GroupBarCodes", "62", "A Pos"
            '30    EnsureEntriesExist "GroupBarCodes", "73", "B Pos"
            '40    EnsureEntriesExist "GroupBarCodes", "84", "AB Pos"
            '50    EnsureEntriesExist "GroupBarCodes", "95", "O Neg"
            '60    EnsureEntriesExist "GroupBarCodes", "06", "A Neg"
            '70    EnsureEntriesExist "GroupBarCodes", "17", "B Neg"
            '80    EnsureEntriesExist "GroupBarCodes", "28", "AB Neg"
            '
End Sub


Public Function EnsureColumnExistsBBo(ByVal TableName As String, _
                                     ByVal ColumnName As String, _
                                     ByVal Definition As String) _
                                     As Boolean

      'Return 1 if column created
      '       0 if column already exists

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo EnsureColumnExistsBB_Error

20    sql = "IF NOT EXISTS " & _
            "    (SELECT * FROM syscolumns WHERE " & _
            "    id = object_id('" & TableName & "') " & _
            "    AND name = '" & ColumnName & "') " & _
            "  BEGIN " & _
            "    ALTER TABLE " & TableName & " " & _
            "    ADD " & ColumnName & " " & Definition & " " & _
            "    SELECT 1 AS RetVal " & _
            "  END " & _
            "ELSE " & _
            "  SELECT 0 AS RetVal"

30    Set tb = CnxnBB(0).Execute(sql)

40    EnsureColumnExistsBBo = tb!retval

50    Exit Function

EnsureColumnExistsBB_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "modDatabase", "EnsureColumnExists", intEL, strES, sql


End Function


