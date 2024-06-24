Attribute VB_Name = "modDefinitionsActive"
Option Explicit

Public Sub CheckCoagActive()

      Dim tb As Recordset
      Dim sql As String
    
10    On Error GoTo CheckCoagActive_Error

20    sql = "Select Distinct TestName from CoagTestDefinitions"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    Do While Not tb.EOF
        'For each test name get the latest date
        'then update them all - just in case they have age related ranges
    
60      sql = "Update CoagTestDefinitions " & _
              "Set ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "' where " & _
              "TestName = '" & tb!TestName & "' " & _
              "and ActiveToDate in " & _
              "( Select top 1 ActiveToDate from CoagTestDefinitions where " & _
              "  TestName = '" & tb!TestName & "' " & _
              "  Order by ActiveToDate Desc, ActiveFromDate Desc)"
70      Cnxn(0).Execute sql
  
80      tb.MoveNext

90    Loop

100   Exit Sub

CheckCoagActive_Error:

Dim strES As String
Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modDefinitionsActive", "CheckCoagActive", intEL, strES, sql

End Sub

Public Sub CheckDisciplineActive(ByVal Discipline As String)
      'Discipline is either "Bio","Imm" or "End"
      Dim tb As Recordset
      Dim sql As String
    
10    sql = "Select Distinct LongName, SampleType, Category, Hospital " & _
            "from " & Discipline & "TestDefinitions"
20    Set tb = New Recordset
30    RecOpenServer 0, tb, sql

40    Do While Not tb.EOF
        'For each test name get the latest date
        'then update them all - just in case they have age related ranges
  
50      sql = "Update " & Discipline & "TestDefinitions " & _
              "Set ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "' where " & _
              "LongName = '" & AddTicks(tb!LongName) & "' " & _
              "and SampleType = '" & tb!SampleType & "' " & _
              "and Category = '" & tb!Category & "' " & _
              "and Hospital = '" & tb!Hospital & "' " & _
              "and ActiveToDate in " & _
              "( Select top 1 ActiveToDate from " & Discipline & "TestDefinitions where " & _
              "  LongName = '" & AddTicks(tb!LongName) & "' " & _
              "  and SampleType = '" & tb!SampleType & "' " & _
              "  and Category = '" & tb!Category & "' " & _
              "  and Hospital = '" & tb!Hospital & "' " & _
              "  Order by ActiveToDate Desc, ActiveFromDate Desc)"
60      Cnxn(0).Execute sql
  
70      tb.MoveNext

80    Loop

End Sub

