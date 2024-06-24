Attribute VB_Name = "modINI"
Option Explicit

Public Sub GetINI()

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
          'ZyamComment
250                 Con = "Provider=SQLOLEDB;" & _
                        "Data Source=" & ServerName & ";" & _
                        "Initial Catalog=" & DatabaseName & ";" & _
                        "User ID=" & GetUID() & ";" & _
                        "Password=" & GetPass() & ";"
          'ZyamComment

          '260       Con = "Provider=SQLOLEDB;" & _
          '              "Data Source=" & "192.168.20.83" & ";" & _
          '              "Initial Catalog=" & "Cavan" & ";" & _
          '              "User ID=" & "zyam" & ";" & _
          '              "Password=" & "zyam12345" & ";"
'250       Con = "Provider=SQLOLEDB;" & _
'              "Data Source=" & "DESKTOP-3OMS1N5\SQLEXPRESS" & ";" & _
'              "Initial Catalog=" & "Cavan" & ";" & _
'              "Integrated Security=SSPI;"


          'Con = "Provider=SQLOLEDB;" & _
          '      "Data Source=" & "JUNAID" & ";" & _
          '      "Initial Catalog=" & "Cavan" & ";" & _
          '      "User ID=" & "sa" & ";" & _
          '      "Password=" & "angel" & ";"
260       Set Cnxn(0) = New Connection
270       Cnxn(0).Open Con

          ConBB = "Provider=SQLOLEDB;" & _
                  "Data Source=" & ServerName & ";" & _
                  "Initial Catalog=" & TransfusionDatabaseName & ";" & _
                  "User ID=" & GetUID() & ";" & _
                  "Password=" & GetPass() & ";"
          '300       ConBB = "Provider=SQLOLEDB;" & _
          '              "Data Source=" & "192.168.20.83" & ";" & _
          '              "Initial Catalog=" & "CavanTransfusion" & ";" & _
          '              "User ID=" & "zyam" & ";" & _
          '              "Password=" & "zyam12345" & ";"
'280       ConBB = "Provider=SQLOLEDB;" & _
'              "Data Source=" & "DESKTOP-3OMS1N5\SQLEXPRESS" & ";" & _
'              "Initial Catalog=" & "Transfusion" & ";" & _
'              "Integrated Security=SSPI;"
290       Set CnxnBB(0) = New Connection
300       CnxnBB(0).Open ConBB

End Sub
Private Function GetPass() As String

Dim RetVal As String
Dim A As String

RetVal = ""
A = GeneratePassString()

        '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
        '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
        '             1         2         3         4         5         6         7         8         9
        'DfySiywtgtw$1>)*
RetVal = Mid$(A, 30, 1) & Mid$(A, 6, 1) & Mid$(A, 25, 1) & Mid$(A, 45, 1) & _
         Mid$(A, 9, 1) & Mid$(A, 25, 1) & Mid$(A, 23, 1) & Mid$(A, 20, 1) & _
         Mid$(A, 7, 1) & Mid$(A, 20, 1) & Mid$(A, 23, 1) & Mid$(A, 55, 1) & _
         Mid$(A, 85, 1) & Mid$(A, 63, 1) & Mid$(A, 61, 1) & Mid$(A, 59, 1)

GetPass = RetVal

End Function
Private Function GetUID() As String

Dim RetVal As String
Dim A As String

RetVal = ""
A = GeneratePassString()

        '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
        '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
        'LabUser      1         2         3         4         5         6         7         8         9

RetVal = Mid$(A, 38, 1) & Mid$(A, 1, 1) & Mid$(A, 2, 1) & Mid$(A, 47, 1) & _
         Mid$(A, 19, 1) & Mid$(A, 5, 1) & Mid$(A, 18, 1)

GetUID = RetVal

End Function
    Private Function GeneratePassString() As String

10      Dim RetVal As String
20      RetVal = ""
30      Dim n As Integer

40      For n = 97 To 122
50          RetVal = RetVal & Chr(n)
60      Next
70      For n = 65 To 90
80          RetVal = RetVal & Chr(n)
90      Next
100     RetVal = RetVal & "!£$%^&*()<>-_+={}[]:@~||;'#,./?"
110     For n = 48 To 57
120         RetVal = RetVal & Chr(n)
130     Next
140     GeneratePassString = RetVal

    End Function

