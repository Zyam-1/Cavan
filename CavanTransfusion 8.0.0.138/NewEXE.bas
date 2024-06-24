Attribute VB_Name = "NewEXE"
Option Explicit


Public Function CheckNewEXE(ByVal NameOfExe As String) As String

      Dim FileName As String
      Dim Current As String
      Dim Found As Boolean
      Dim Path As String

10    Found = False

20    Path = App.Path & "\"
30    Current = UCase$(NameOfExe) & ".EXE"
40    FileName = UCase$(Dir(Path & NameOfExe & "*.exe", vbNormal))

50    Do While FileName <> ""
60      If FileName > Current Then
70        Current = FileName
80        Found = True
90      End If
100     FileName = UCase$(Dir)
110   Loop

120   If Found And UCase$(App.EXEName) & ".EXE" <> Current Then
130     CheckNewEXE = Path & Current
140   Else
150     CheckNewEXE = ""
160   End If

End Function

Public Sub CheckLabConfirmLogInDb(ByVal Cx As Connection)
      Dim sql As String
      Dim tbExists As Recordset

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist "
      'if it has a record then the table does exist.


10    On Error GoTo CheckLabConfirmLogInDb_Error

20    sql = "SELECT Name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = 'LabConfirmLog'"
30    Set tbExists = New Recordset
40    Set tbExists = Cx.Execute(sql)

50    If tbExists.EOF Then 'There is no table  in database
60      sql = "CREATE TABLE LabConfirmLog " & _
              "( SampleNumber  nvarchar(50), " & _
              "  DateTime  datetime, " & _
              "  FwdCardLotNo  nvarchar(50), " & _
              "  RevCardLotNo  nvarchar(50), " & _
              "  UserName  nvarchar(50) )"
70      Cx.Execute sql
80    End If

90    Exit Sub

CheckLabConfirmLogInDb_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "NewEXE", "CheckLabConfirmLogInDb", intEL, strES, sql


End Sub

Public Sub CheckBTSConfirmedInDb(ByVal Cx As Connection)

10    On Error GoTo CheckBTSConfirmedInDb_Error

      Dim sql As String
      Dim tbExists As Recordset

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist "
      'if it has a record then the table does exist.


20    sql = "SELECT Name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = 'BTSConfirmed'"
30    Set tbExists = New Recordset
40    Set tbExists = Cx.Execute(sql)

50    If tbExists.EOF Then 'There is no table  in database
60      sql = "CREATE TABLE BTSConfirmed " & _
              "( SampleNumber  nvarchar(50), " & _
              "  PackNumber  nvarchar(50), " & _
              "  PackExpiry  smalldatetime, " & _
              "  DateTime  datetime, " & _
              "  BTSConfirmed  nvarchar(50), " & _
              "  UserName  nvarchar(50) )"
70      Cx.Execute sql
80    End If

90    Exit Sub

CheckBTSConfirmedInDb_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "NewEXE", "CheckBTSConfirmedInDb", intEL, strES, sql

End Sub
Public Sub CheckVersionControlInDb(ByVal Cx As Connection)

      Dim sql As String
      Dim tbExists As Recordset

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist "
      'if it has a record then the table does exist.

10    On Error GoTo CheckVersionControlInDb_Error

20    sql = "SELECT Name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = 'VersionControl'"
30    Set tbExists = New Recordset
40    Set tbExists = Cx.Execute(sql)

50    If tbExists.EOF Then 'There is no table  in database
60      sql = "CREATE TABLE VersionControl " & _
              "( Filename  nvarchar(50), " & _
              "  File_Version  nvarchar(50), " & _
              "  File_DateCreated  datetime, " & _
              "  DateTime  datetime, " & _
              "  Deployed  bit, " & _
              "  Active  bit, " & _
              "  DoNotUse  bit )"
70      Cx.Execute sql
80    End If

90    Exit Sub

CheckVersionControlInDb_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "NewEXE", "CheckVersionControlInDb", intEL, strES, sql

End Sub

Public Function AllowedToActivateVersion(strFileName As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo AllowedToActivateVersion_Error

20    CheckVersionControlInDb Cnxn(0)

30    sql = "Select * from VersionControl where " & _
            "FileName = '" & strFileName & "' " & _
            "and DoNotUse = 1 "

40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql

60    If Not tb.EOF Then
70        AllowedToActivateVersion = False
80    Else
90        AllowedToActivateVersion = True
100   End If

110   Exit Function

AllowedToActivateVersion_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "NewEXE", "AllowedToActivateVersion", intEL, strES, sql


End Function

Public Sub CreateShortcut(strFileName As String)
   
       'Reference wshom.ocx before using this code.
       'Windows Script Host Object model
       'Dim oShell As New IWshRuntimeLibrary.IWShellsh_Class
       'Dim oShort As IWshRuntimeLibrary.IWshShortcut_Class
 
       Dim oShell As New IWshRuntimeLibrary.IWshShell_Class
       Dim oShort As IWshRuntimeLibrary.IWshShortcut_Class


       Dim strDesktopPath As String

       'Get the path to the desktop
10     strDesktopPath = oShell.SpecialFolders("Desktop")
       'Create a new shortcut
20     Set oShort = oShell.CreateShortcut(strDesktopPath & "\Transfusion.lnk")
30     oShort.Description = "NetAcquire Blood Transfusion"
40     oShort.IconLocation = strFileName & ", 0"
50     oShort.TargetPath = App.Path & "\" & strFileName
60     oShort.WorkingDirectory = App.Path
70     oShort.IconLocation = App.Path & "\" & strFileName & ", 0"
80     oShort.Save

End Sub

Public Function IsLiveSystem() As Boolean

10    IsLiveSystem = True

      'This runs test program in test directory
      'on test database
20    If InStr(UCase$(App.Path), "TEST") Then
30      IsLiveSystem = False
40    End If

      'This runs live program from live directory
      'with test database
50    If UCase$(Command$) = "TEST" Then
60      IsLiveSystem = False
70    End If

End Function


