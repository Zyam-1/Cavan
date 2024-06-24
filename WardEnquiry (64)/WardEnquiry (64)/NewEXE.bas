Attribute VB_Name = "NewEXE"
Option Explicit
'+++ Junaid
Public m_RunExe As String
Public m_FindChart As Boolean
Public m_Pass As String
Public m_Chart As String
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub Main()
    
    If Command$ = "" Then
        m_RunExe = "NA"
    Else
        m_RunExe = Command$
    End If
    If m_RunExe = "OCM" Then
        m_Pass = ReadINI("Password", "Pass", "")
        If m_Pass = "" Then
            MsgBox "Password not present.", vbInformation
            End
        End If
        m_Chart = ReadINI("ChartNo", "Chart", "")
        If m_Chart = "" Then
            MsgBox "Chart No. not present.", vbInformation
            End
        End If
        frmMain.Show
    Else
        frmMain.Show
    End If
    m_FindChart = False
    
End Sub

Public Function WriteINI(p_Section As String, p_Key As String, p_Value As String) As Boolean
    On Error Resume Next
    Dim l_ININame   As String
    
    l_ININame = App.Path & "\WE.INI"
    
    WriteINI = WritePrivateProfileString(p_Section, p_Key, p_Value, l_ININame)
End Function

'This functin reads from INI file
Public Function ReadINI(p_Section As String, p_Key As String, p_DefaultValue) As String
    Dim l_Count     As Long
    Dim l_ININame   As String
    Dim l_string    As String * 256
    
    l_ININame = App.Path & "\WE.INI"
    
    l_Count = GetPrivateProfileString(p_Section, p_Key, p_DefaultValue, l_string, 255, l_ININame)
    ReadINI = Left(l_string, l_Count)
End Function
'--- Junaid

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

Public Sub CheckVersionControlInDb(ByVal Cx As Connection)

      Dim sql As String
      Dim tbExists As Recordset

10    On Error GoTo CTIDErr

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist "
      'if it has a record then the table does exist.

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

CTIDErr:

      Dim er As Long
      Dim es As String
    
100   er = Err.Number
110   es = Err.Description
120   MsgBox es
130   Exit Sub


End Sub

Public Function AllowedToActivateVersion(strFileName As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    CheckVersionControlInDb Cnxn(0)

20    sql = "Select * from VersionControl where " & _
            "FileName = '" & strFileName & "' " & _
            "and DoNotUse = 1 "

30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60        AllowedToActivateVersion = False
70    Else
80        AllowedToActivateVersion = True
90    End If

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


