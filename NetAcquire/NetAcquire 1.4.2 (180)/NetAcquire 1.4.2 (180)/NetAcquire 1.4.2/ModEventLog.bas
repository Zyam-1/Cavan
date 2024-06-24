Attribute VB_Name = "ModEventLog"
Option Explicit


Public Sub CheckEventLogInDb(ByVal Cx As Connection)

      Dim sql As String

300   On Error GoTo CheckEventLogInDb_Error

310   If IsTableInDatabase("EventLog") = False Then 'There is no table  in database
320     sql = "CREATE TABLE EventLog " & _
              "( Description  nvarchar(50), " & _
              "  DateTime datetime, " & _
              "  UserName nvarchar(50) )"
330     Cx.Execute sql
340   End If

350   Exit Sub

CheckEventLogInDb_Error:

      Dim strES As String
      Dim intEL As Integer

360   intEL = Erl
370   strES = Err.Description
380   LogError "ModEventLog", "CheckEventLogInDb", intEL, strES, sql

End Sub

Public Sub LogToEventLog(ByVal strDescription As String)

      Dim sql As String

390   On Error GoTo LogToEventLog_Error

400   sql = "Insert into EventLog " & _
            "(Description, DateTime, UserName) VALUES " & _
            "('" & AddTicks(strDescription) & "', " & _
            "'" & Format(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
            "'" & UserName & "');"
410   Cnxn(0).Execute sql

420   Exit Sub

LogToEventLog_Error:

      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "ModEventLog", "LogToEventLog", intEL, strES, sql


End Sub

'+++ Junaid 12-12-2023
Public Sub WriteToFile_Execution(Optional p_SQLQuery As String = "", Optional p_ErrorRased As Boolean, Optional p_Log As String = "")
460       On Error GoTo ErrorHandler
          Dim filehandle As Integer
          Dim l_FileName As String
      '    Exit Sub
470       filehandle = FreeFile()
480       l_FileName = Month(Date) & "-" & Day(Date) & "-" & Year(Date)
490       If p_ErrorRased Then
500           l_FileName = App.Path & "\" & l_FileName & "_ErrQueries.TXT"
510           Open l_FileName For Append As #filehandle
520       Else
530           l_FileName = App.Path & "\" & l_FileName & "_SqlQueries.TXT"
540           Open l_FileName For Append As #filehandle
550       End If
          
560       If p_Log <> "" Then
570           Print #filehandle, "--" & Now & " ; " & UserName & " ; " & App.Major & "." & App.Minor & "." & App.Revision & " ;" & p_Log & Chr(13)
580       End If
590       If p_SQLQuery <> "" Then
600           If Left(Trim(p_SQLQuery), 2) = "**" Then
610               Print #filehandle, "--" & Now & " ; " & UserName & " ; " & App.Major & "." & App.Minor & "." & App.Revision & " ;" & p_SQLQuery & Chr(13)
620           Else
630               Print #filehandle, "/*" & Now & " ; " & UserName & " ; " & App.Major & "." & App.Minor & "." & App.Revision & " ;*/" & p_SQLQuery & Chr(13)
640           End If
650       End If
          
660       Close #filehandle

670       Exit Sub
ErrorHandler:
680       Resume Next
End Sub
'--- Junaid
