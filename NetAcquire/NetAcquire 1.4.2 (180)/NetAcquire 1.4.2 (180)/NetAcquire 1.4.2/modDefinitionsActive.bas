Attribute VB_Name = "modDefinitionsActive"
Option Explicit

Public Sub CheckCoagActive()

      Dim tb As Recordset
      Dim sql As String
          
65430 On Error GoTo CheckCoagActive_Error

65440 sql = "Select Distinct TestName from CoagTestDefinitions"
65450 Set tb = New Recordset
65460 RecOpenServer 0, tb, sql

65470 Do While Not tb.EOF
        'For each test name get the latest date
        'then update them all - just in case they have age related ranges
          
65480   sql = "Update CoagTestDefinitions " & _
              "Set ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "' where " & _
              "TestName = '" & tb!TestName & "' " & _
              "and ActiveToDate in " & _
              "( Select top 1 ActiveToDate from CoagTestDefinitions where " & _
              "  TestName = '" & tb!TestName & "' " & _
              "  Order by ActiveToDate Desc, ActiveFromDate Desc)"
65490   Cnxn(0).Execute sql
        
65500   tb.MoveNext

65510 Loop

65520 Exit Sub

CheckCoagActive_Error:

      Dim strES As String
      Dim intEL As Integer

65530 intEL = Erl
10    strES = Err.Description
20    LogError "modDefinitionsActive", "CheckCoagActive", intEL, strES, sql


End Sub

Public Sub CheckDisciplineActive(ByVal Discipline As String)
      'Discipline is either "Bio","Imm" or "End"
      Dim tb As Recordset
      Dim sql As String
          
30    On Error GoTo CheckDisciplineActive_Error

40    sql = "Select Distinct LongName, SampleType, Category, Hospital " & _
            "from " & Discipline & "TestDefinitions"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql

70    Do While Not tb.EOF
        'For each test name get the latest date
        'then update them all - just in case they have age related ranges
        
80      sql = "Update " & Discipline & "TestDefinitions " & _
              "Set ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "' where " & _
              "LongName = '" & tb!LongName & "' " & _
              "and SampleType = '" & tb!SampleType & "' " & _
              "and Category = '" & tb!Category & "' " & _
              "and Hospital = '" & tb!Hospital & "' " & _
              "and ActiveToDate in " & _
              "( Select top 1 ActiveToDate from " & Discipline & "TestDefinitions where " & _
              "  LongName = '" & tb!LongName & "' " & _
              "  and SampleType = '" & tb!SampleType & "' " & _
              "  and Category = '" & tb!Category & "' " & _
              "  and Hospital = '" & tb!Hospital & "' " & _
              "  Order by ActiveToDate Desc, ActiveFromDate Desc)"
90      Cnxn(0).Execute sql
        
100     tb.MoveNext

110   Loop

120   Exit Sub

CheckDisciplineActive_Error:

      Dim strES As String
      Dim intEL As Integer
      Dim lngEN As Long

130   intEL = Erl
140   strES = Err.Description
150   lngEN = Err.Number
160   If lngEN <> -2147467259 Then '[Microsoft][ODBC SQL Server Driver][SQL Server]
                                   'Transaction (Process ID 110) was deadlocked on {lock} resources
                                   'with another process and has been chosen as the deadlock victim.
                                   'Rerun the transaction.

170     LogError "modDefinitionsActive", "CheckDisciplineActive", intEL, strES, sql
180   End If

End Sub

