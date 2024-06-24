Attribute VB_Name = "modINI"
Option Explicit

Public Sub GetINI()
          'On Error GoTo ErrorHandler

          Dim ServerName As String
          Dim DatabaseName As String
          Dim TransfusionDatabaseName As String
          Dim Con As String
          Dim ConBB As String

          Dim S As String
          Dim PathToINI As String
          Dim f As Integer

10        PathToINI = App.Path & "\INI.xml"
20        f = FreeFile
30        Open PathToINI For Input As #f

40        Do While Not EOF(f)
50            Line Input #f, S
60            S = Trim$(UCase$(S))
70            If InStr(S, "<SITE>") > 0 Then
80                S = Replace(S, "<SITE>", "")
90                S = Replace(S, "</SITE>", "")
100               HospName(0) = S
110           ElseIf InStr(S, "<SERVER>") > 0 Then
120               S = Replace(S, "<SERVER>", "")
130               S = Replace(S, "</SERVER>", "")
140               ServerName = S
150           ElseIf InStr(S, "<DATABASE>") > 0 Then
160               S = Replace(S, "<DATABASE>", "")
170               S = Replace(S, "</DATABASE>", "")
180               DatabaseName = S
190           ElseIf InStr(S, "<TRANSFUSIONDATABASE>") > 0 Then
200               S = Replace(S, "<TRANSFUSIONDATABASE>", "")
210               S = Replace(S, "</TRANSFUSIONDATABASE>", "")
220               TransfusionDatabaseName = S
230           End If
240       Loop
'          'Prod Connection
'250       Con = "Provider=SQLOLEDB;" & _
'              "Data Source=" & ServerName & ";" & _
'              "Initial Catalog=" & DatabaseName & ";" & _
'              "User ID=" & GetUID() & ";" & _
'              "Password=" & GetPass() & ";"
          'Local Connection
260       Con = "Provider=SQLOLEDB;" & _
              "Data Source=" & "DESKTOP-3OMS1N5\SQLEXPRESS" & ";" & _
              "Initial Catalog=" & "Cavan" & ";" & _
              "Integrated Security=SSPI;"
'          'Server Connection
'          Con = "Provider=SQLOLEDB;" & _
'                "Data Source=" & "192.168.20.83" & ";" & _
'                "Initial Catalog=" & "Cavan" & ";" & _
'                "User ID=" & "zyam" & ";" & _
'                "Password=" & "zyam12345" & ";"
          ''


270       Set Cnxn(0) = New Connection
280       Cnxn(0).Open Con


'
'290       ConBB = "Provider=SQLOLEDB;" & _
'              "Data Source=" & ServerName & ";" & _
'              "Initial Catalog=" & TransfusionDatabaseName & ";" & _
'              "User ID=" & GetUID() & ";" & _
'              "Password=" & GetPass() & ";"
              
'              Con = "Provider=SQLOLEDB;" & _
'                "Data Source=" & "192.168.20.83" & ";" & _
'                "Initial Catalog=" & "Transfusion" & ";" & _
'                "User ID=" & "zyam" & ";" & _
'                "Password=" & "zyam12345" & ";"
300       ConBB = "Provider=SQLOLEDB;" & _
              "Data Source=" & "DESKTOP-3OMS1N5\SQLEXPRESS" & ";" & _
              "Initial Catalog=" & "Transfusion" & ";" & _
              "Integrated Security=SSPI;"
310       Set CnxnBB(0) = New Connection
320       CnxnBB(0).Open ConBB

          'Exit Sub
          'ErrorHandler:
          'MsgBox Err.Description

End Sub
Private Function GetPass() As String

      Dim RetVal As String
      Dim a As String

1710  RetVal = ""
1720  a = GeneratePassString()

              '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
              '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
              '             1         2         3         4         5         6         7         8         9
              'DfySiywtgtw$1>)*
1730  RetVal = Mid$(a, 30, 1) & Mid$(a, 6, 1) & Mid$(a, 25, 1) & Mid$(a, 45, 1) & _
               Mid$(a, 9, 1) & Mid$(a, 25, 1) & Mid$(a, 23, 1) & Mid$(a, 20, 1) & _
               Mid$(a, 7, 1) & Mid$(a, 20, 1) & Mid$(a, 23, 1) & Mid$(a, 55, 1) & _
               Mid$(a, 85, 1) & Mid$(a, 63, 1) & Mid$(a, 61, 1) & Mid$(a, 59, 1)

1740  GetPass = RetVal

End Function
Private Function GetUID() As String

      Dim RetVal As String
      Dim a As String

1750  RetVal = ""
1760  a = GeneratePassString()

              '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
              '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
              'LabUser      1         2         3         4         5         6         7         8         9

1770  RetVal = Mid$(a, 38, 1) & Mid$(a, 1, 1) & Mid$(a, 2, 1) & Mid$(a, 47, 1) & _
               Mid$(a, 19, 1) & Mid$(a, 5, 1) & Mid$(a, 18, 1)

1780  GetUID = RetVal

End Function
    Private Function GeneratePassString() As String

        Dim RetVal As String
1790    RetVal = ""
        Dim n As Integer

1800    For n = 97 To 122
1810        RetVal = RetVal & Chr(n)
1820    Next
1830    For n = 65 To 90
1840        RetVal = RetVal & Chr(n)
1850    Next
1860    RetVal = RetVal & "!£$%^&*()<>-_+={}[]:@~||;'#,./?"
1870    For n = 48 To 57
1880        RetVal = RetVal & Chr(n)
1890    Next
1900    GeneratePassString = RetVal

    End Function

