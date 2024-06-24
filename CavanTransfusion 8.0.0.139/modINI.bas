Attribute VB_Name = "modINI"
Option Explicit

Public Sub GetINI()

Dim ServerName As String
Dim DatabaseName As String
Dim TransfusionDatabaseName As String
Dim Con As String
Dim ConBB As String

Dim s As String
Dim PathToINI As String
Dim f As Integer
'Zyam
PathToINI = App.Path & "\INI.xml"
f = FreeFile
Open PathToINI For Input As #f

Do While Not EOF(f)
  Line Input #f, s
  s = Trim$(UCase$(s))
  If InStr(s, "<SITE>") > 0 Then
    s = Replace(s, "<SITE>", "")
    s = Replace(s, "</SITE>", "")
    HospName(0) = s
  ElseIf InStr(s, "<SERVER>") > 0 Then
    s = Replace(s, "<SERVER>", "")
    s = Replace(s, "</SERVER>", "")
    ServerName = s
  ElseIf InStr(s, "<DATABASE>") > 0 Then
    s = Replace(s, "<DATABASE>", "")
    s = Replace(s, "</DATABASE>", "")
    DatabaseName = s
  ElseIf InStr(s, "<TRANSFUSIONDATABASE>") > 0 Then
    s = Replace(s, "<TRANSFUSIONDATABASE>", "")
    s = Replace(s, "</TRANSFUSIONDATABASE>", "")
    TransfusionDatabaseName = s
  End If
Loop
'Zyam
Con = "Provider=SQLOLEDB;" & _
      "Data Source=" & ServerName & ";" & _
      "Initial Catalog=" & DatabaseName & ";" & _
      "User ID=" & GetUID() & ";" & _
      "Password=" & GetPass() & ";"
'Con = "Provider=SQLOLEDB;" & _
'      "Data Source=" & "192.168.20.83" & ";" & _
'      "Initial Catalog=" & "Cavan" & ";" & _
'      "User ID=" & "zyam" & ";" & _
'      "Password=" & "zyam1234" & ";"
Set Cnxn(0) = New Connection
Cnxn(0).Open Con

ConBB = "Provider=SQLOLEDB;" & _
        "Data Source=" & ServerName & ";" & _
        "Initial Catalog=" & TransfusionDatabaseName & ";" & _
        "User ID=" & GetUID() & ";" & _
        "Password=" & GetPass() & ";"
'ConBB = "Provider=SQLOLEDB;" & _
'        "Data Source=" & "192.168.20.83" & ";" & _
'        "Initial Catalog=" & "CavanTransfusion" & ";" & _
'        "User ID=" & "zyam" & ";" & _
'        "Password=" & "zyam1234" & ";"
Set CnxnBB(0) = New Connection
CnxnBB(0).Open ConBB

End Sub
Private Function GetPass() As String

Dim retval As String
Dim A As String

retval = ""
A = GeneratePassString()

        '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
        '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
        '             1         2         3         4         5         6         7         8         9
        'DfySiywtgtw$1>)*
retval = Mid$(A, 30, 1) & Mid$(A, 6, 1) & Mid$(A, 25, 1) & Mid$(A, 45, 1) & _
         Mid$(A, 9, 1) & Mid$(A, 25, 1) & Mid$(A, 23, 1) & Mid$(A, 20, 1) & _
         Mid$(A, 7, 1) & Mid$(A, 20, 1) & Mid$(A, 23, 1) & Mid$(A, 55, 1) & _
         Mid$(A, 85, 1) & Mid$(A, 63, 1) & Mid$(A, 61, 1) & Mid$(A, 59, 1)

GetPass = retval

End Function
Private Function GetUID() As String

Dim retval As String
Dim A As String

retval = ""
A = GeneratePassString()

        '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
        '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
        'LabUser      1         2         3         4         5         6         7         8         9

retval = Mid$(A, 38, 1) & Mid$(A, 1, 1) & Mid$(A, 2, 1) & Mid$(A, 47, 1) & _
         Mid$(A, 19, 1) & Mid$(A, 5, 1) & Mid$(A, 18, 1)

GetUID = retval

End Function
    Private Function GeneratePassString() As String

10      Dim retval As String
20      retval = ""
30      Dim n As Integer

40      For n = 97 To 122
50          retval = retval & Chr(n)
60      Next
70      For n = 65 To 90
80          retval = retval & Chr(n)
90      Next
100     retval = retval & "!£$%^&*()<>-_+={}[]:@~||;'#,./?"
110     For n = 48 To 57
120         retval = retval & Chr(n)
130     Next
140     GeneratePassString = retval

    End Function

