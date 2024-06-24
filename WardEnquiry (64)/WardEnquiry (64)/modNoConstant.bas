Attribute VB_Name = "modNoConstant"
Option Explicit
  
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" _
    (ByVal lpSectionName As String, ByVal lpKeyName As String, _
     ByVal lpDefault As String, ByVal lpbuffurnedString As String, _
     ByVal nBuffSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSectionNames Lib "Kernel32.dll" Alias _
    "GetPrivateProfileSectionNamesA" _
    (ByVal lpszReturnBuffer As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

Public Sub ConnectToDatabase()

      Dim dbDSN As String
      Dim dbDSNbb As String
      Dim dbRemoteDSNbb As String
      Dim tb As Recordset
      Dim TempCnxn As Connection
      Dim dbConnectRemoteBB As String
10    ReDim Cnxn(0 To 0) As Connection
20    ReDim CnxnBB(0 To 0) As Connection
30    ReDim CnxnRemoteBB(0 To 0) As Connection
40    ReDim HospName(0 To 0) As String
      Dim Con As String
      Dim ConBB As String

50    On Error Resume Next
60    HospName(0) = GetcurrentConnectInfo(Con, ConBB)
70    frmMain.StatusBar1.Panels(4) = "Custom Software Ltd." 'Note full stop

80    If IsIDE And HospName(0) = "" Then
90      MsgBox "INI Error"
100     End
110   ElseIf HospName(0) = "" Then
120     frmMain.StatusBar1.Panels(4) = "Custom Software Ltd" 'Note NO full stop
  
130     If GetConnectInfo("Active", Con, HospName(0)) Then
140       GetConnectInfo "BB", ConBB
150       GetConnectInfo "RemoteBB", dbConnectRemoteBB
160     Else
170       Set TempCnxn = New Connection
180       TempCnxn.Open "uid=sa;dsn=Constant;"
190       Set tb = New Recordset
200       With tb
210         .CursorLocation = adUseServer
220         .CursorType = adOpenDynamic
230         .LockType = adLockOptimistic
240         .ActiveConnection = TempCnxn
250         .Source = "Select * from Constant where active = 1"
260         .Open
270       End With
    
280       dbDSN = tb!DSN & ""
290       dbDSNbb = tb!DSNBB & ""
300       dbRemoteDSNbb = tb!RemoteDSNbb & ""
    
310       HospName(0) = Trim$(tb!Hosp & "")
    
320       Con = "uid=sa;pwd=;dsn=" & dbDSN & ";"
330       If dbDSNbb <> "" Then ConBB = "uid=sa;pwd=;dsn=" & dbDSNbb & ";"
340       If dbRemoteDSNbb <> "" Then dbConnectRemoteBB = "uid=sa;pwd=;dsn=" & dbRemoteDSNbb & ";"
350     End If

360     If dbConnectRemoteBB <> "" Then
370       Set CnxnRemoteBB(0) = New Connection
380       CnxnRemoteBB(0).Open dbConnectRemoteBB
390     End If
400   End If

410   Set Cnxn(0) = New Connection
420   Cnxn(0).Open Con
430   If ConBB <> "" Then
440     Set CnxnBB(0) = New Connection
450     CnxnBB(0).Open ConBB
460   End If

      'CheckGroupedHospitalsInDb
      'GetHospitalsInGroup

End Sub
Public Function GetConnectInfo(ByVal ConnectTo As String, _
                               ByRef ReturnConnectionString As String, _
                               Optional ByRef HospName As Variant) As Boolean

      'ConnectTo = "Active"
      '            "BB"
      '            "Active" & n - HospitalGroup
      '            "BB" & n - HospitalGroup

10    GetConnectInfo = False

20    If Not IsMissing(HospName) Then
30      HospName = GetSetting("NetAcquire", "HospName", ConnectTo, "")
40      If Left$(UCase$(HospName), 5) = "LOCAL" Then
50        HospName = Mid$(HospName, 6)
60      End If
70    End If

80    ReturnConnectionString = GetSetting("NetAcquire", "Cnxn", ConnectTo, "")

90    If Trim$(ReturnConnectionString) <> "" Then
  
100     ReturnConnectionString = Obfuscate(ReturnConnectionString)
  
110     GetConnectInfo = True
  
120   End If

End Function


Public Function GetcurrentConnectInfo(ByRef Con As String, ByRef ConBB As String) As String

          'Returns Hospital Name

          Dim HospitalNames() As String
          Dim n As Long
          Dim HospitalName As String
          Dim retHospitalName As String
          Dim ServerName As String
          Dim NetAcquireDB As String
          Dim TransfusionDB As String
          Dim UID As String
          Dim PWD As String
          Dim CurrentPath As String

          '10    If IsIDE Then
          '20      CurrentPath = "C:\ClientCode\NetAcquire.INI"
          '30    Else
10        CurrentPath = App.Path & "\NetAcquire.INI"
          '50    End If

20        HospitalNames = GetINISectionNames(CurrentPath, n)
30        HospitalName = HospitalNames(0)
40        If Left$(UCase$(HospitalName), 5) = "LOCAL" Then
50            retHospitalName = Mid$(HospitalName, 6)
60        Else
70            retHospitalName = HospitalName
80        End If

90        ServerName = ProfileGetItem(HospitalName, "N", "", CurrentPath)
100       NetAcquireDB = ProfileGetItem(HospitalName, "D", "", CurrentPath)
110       TransfusionDB = ProfileGetItem(HospitalName, "T", "", CurrentPath)
120       UID = ProfileGetItem(HospitalName, "U", "", CurrentPath)
130       PWD = ProfileGetItem(HospitalName, "P", "", CurrentPath)

140       Con = "DRIVER={SQL Server};" & _
              "Server=" & Obfuscate(ServerName) & ";" & _
              "Database=" & Obfuscate(NetAcquireDB) & ";" & _
              "uid=" & Obfuscate(UID) & ";" & _
              "pwd=" & Obfuscate(PWD) & ";"
          '      Con = "Provider=SQLOLEDB;" & _
          '      "Data Source=" & "JUNAID" & ";" & _
          '      "Initial Catalog=" & "Cavan" & ";" & _
          '      "User ID=" & "LabUser" & ";" & _
          '      "Password=" & "DfySiywtgtw$1>)=" & ";"
'181                 Con = "DRIVER={SQL Server};" & _
'                      "Server=" & "192.168.20.83" & ";" & _
'                      "Database=" & "Cavan" & ";" & _
'                      "uid=" & "zyam" & ";" & _
'                      "pwd=" & "zyam12345" & ";"


160       If TransfusionDB <> "" Then
170           ConBB = "DRIVER={SQL Server};" & _
                  "Server=" & Obfuscate(ServerName) & ";" & _
                  "Database=" & Obfuscate(TransfusionDB) & ";" & _
                  "uid=" & Obfuscate(UID) & ";" & _
                  "pwd=" & Obfuscate(PWD) & ";"


              '        ConBB = "Provider=SQLOLEDB;" & _
              '      "Data Source=" & "JUNAID" & ";" & _
              '      "Initial Catalog=" & "CavanTransfusion" & ";" & _
              '      "User ID=" & "LabUser" & ";" & _
              '      "Password=" & "DfySiywtgtw$1>)=" & ";"

190       End If

200       GetcurrentConnectInfo = retHospitalName

End Function


Private Function ProfileGetItem(ByRef sSection As String, _
                                ByRef sKeyName As String, _
                                ByRef sDefValue As String, _
                                ByRef sIniFile As String) As String

          'retrieves a value from an ini file
          'corresponding to the section and
          'key name passed.

      Dim dwSize As Integer
      Dim nBuffSize As Integer
      Dim buff As String
      Dim RetVal As String

      'Call the API with the parameters passed.
      'nBuffSize is the length of the string
      'in buff, including the terminating null.
      'If a default value was passed, and the
      'section or key name are not in the file,
      'that value is returned. If no default
      'value was passed (""), then dwSize
      'will = 0 if not found.
      '
      'pad a string large enough to hold the data
10    buff = Space(2048)
20    nBuffSize = Len(buff)
30    dwSize = GetPrivateProfileString(sSection, sKeyName, sDefValue, buff, nBuffSize, sIniFile)

40    If dwSize > 0 Then
50      RetVal = Left$(buff, dwSize)
60    End If

70    ProfileGetItem = RetVal

End Function

Private Function GetINISectionNames(ByRef inFile As String, ByRef outCount As Long) As String()

      Dim StrBuf As String
      Dim BufLen As Long
      Dim RetVal() As String
      Dim Count As Long

10    BufLen = 16

20    Do
30      BufLen = BufLen * 2
40      StrBuf = Space$(BufLen)
50      Count = GetPrivateProfileSectionNames(StrBuf, BufLen, inFile)
60    Loop While Count = BufLen - 2

70    If (Count) Then
80      RetVal = Split(Left$(StrBuf, Count - 1), vbNullChar)
90      outCount = UBound(RetVal) + 1
100   End If

110   GetINISectionNames = RetVal

End Function
  

Public Function Obfuscate(ByVal strData As String) As String

      Dim lngI As Long
      Dim lngJ As Long
   
10    For lngI = 0 To Len(strData) \ 4
20      For lngJ = 1 To 4
30         Obfuscate = Obfuscate & Mid$(strData, (4 * lngI) + 5 - lngJ, 1)
40      Next
50    Next

End Function

