Attribute VB_Name = "NewEXE"
Option Explicit


Public Function CheckNewEXE(ByVal NameOfExe As String) As String

          Dim FileName As String
          Dim Current As String
          Dim Found As Boolean
          Dim Path As String

26790     Found = False

26800     Path = App.Path & "\"
26810     Current = UCase$(NameOfExe) & ".EXE"
26820     FileName = UCase$(Dir(Path & NameOfExe & "*.exe", vbNormal))

26830     Do While FileName <> ""
26840         If FileName > Current Then
26850             Current = FileName
26860             Found = True
26870         End If
26880         FileName = UCase$(Dir)
26890     Loop

26900     If Found And UCase$(App.EXEName) & ".EXE" <> Current Then
26910         CheckNewEXE = Path & Current
26920     Else
26930         CheckNewEXE = ""
26940     End If

End Function

Public Sub CheckVersionControlInDb(ByVal Cx As Connection)

          Dim sql As String
          Dim tbExists As Recordset

          'How to find if a table exists in a database
          'open a recordset with the following sql statement:
          'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
          'If the recordset it at eof then the table doesn't exist "
          'if it has a record then the table does exist.

26950     On Error GoTo CheckVersionControlInDb_Error

26960     sql = "SELECT Name FROM sysobjects WHERE " & _
              "xtype = 'U' " & _
              "AND name = 'VersionControl'"
26970     Set tbExists = New Recordset
26980     Set tbExists = Cx.Execute(sql)

26990     If tbExists.EOF Then 'There is no table  in database
27000         sql = "CREATE TABLE VersionControl " & _
                  "( Filename  nvarchar(50), " & _
                  "  File_Version  nvarchar(50), " & _
                  "  File_DateCreated  datetime, " & _
                  "  DateTime  datetime, " & _
                  "  Deployed  bit, " & _
                  "  Active  bit, " & _
                  "  DoNotUse  bit )"
27010         Cx.Execute sql
27020     End If
27030     Exit Sub

27040     Exit Sub

CheckVersionControlInDb_Error:

          Dim strES As String
          Dim intEL As Integer

27050     intEL = Erl
27060     strES = Err.Description
27070     LogError "NewEXE", "CheckVersionControlInDb", intEL, strES, sql

End Sub

Public Function AllowedToActivateVersion(strFileName As String) As Boolean

          Dim sql As String
          Dim tb As Recordset

27080     On Error GoTo AllowedToActivateVersion_Error

27090     CheckVersionControlInDb Cnxn(0)

27100     sql = "Select * from VersionControl where " & _
              "FileName = '" & strFileName & "' " & _
              "and DoNotUse = 1 "

27110     Set tb = New Recordset
27120     RecOpenServerBB 0, tb, sql

27130     If Not tb.EOF Then
27140         AllowedToActivateVersion = False
27150     Else
27160         AllowedToActivateVersion = True
27170     End If

27180     Exit Function

AllowedToActivateVersion_Error:

          Dim strES As String
          Dim intEL As Integer

27190     intEL = Erl
27200     strES = Err.Description
27210     LogError "NewEXE", "AllowedToActivateVersion", intEL, strES, sql


End Function

Public Function IsLiveSystem() As Boolean

27220     IsLiveSystem = True

          'This runs test program in test directory
          'on test database
27230     If InStr(UCase$(App.Path), "TEST") Then
27240         IsLiveSystem = False
27250     End If

          'This runs live program from live directory
          'with test database
27260     If UCase$(Command$) = "TEST" Then
27270         IsLiveSystem = False
27280     End If

End Function


