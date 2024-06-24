Attribute VB_Name = "modPassword"
Option Explicit

Public Function NameHasBeenUsed(ByVal UserName As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

4660  On Error GoTo NameHasBeenUsed_Error

4670  NameHasBeenUsed = False

4680  sql = "SELECT Name FROM Users WHERE " & _
            "Name = '" & AddTicks(UserName) & "' "
4690  Set tb = New Recordset
4700  RecOpenServer 0, tb, sql
4710  If Not tb.EOF Then
4720    NameHasBeenUsed = True
4730  End If

4740  Exit Function

NameHasBeenUsed_Error:

      Dim strES As String
      Dim intEL As Integer

4750  intEL = Erl
4760  strES = Err.Description
4770  LogError "modPassword", "NameHasBeenUsed", intEL, strES, sql

End Function


Public Function AllLowerCase(stringToCheck As String) As Boolean

4780  AllLowerCase = StrComp(stringToCheck, LCase$(stringToCheck), vbBinaryCompare) = 0

End Function

Public Function ContainsNumeric(ByVal s As String) As Boolean

      Dim n As Integer

4790  ContainsNumeric = False
4800  For n = 1 To Len(s)
4810    If InStr("0123456789", Mid$(s, n, 1)) Then
4820      ContainsNumeric = True
4830      Exit Function
4840    End If
4850  Next

End Function

Public Function ContainsAlpha(ByVal s As String) As Boolean

      Dim n As Integer
      Dim strTestL As String
      Dim strTestU As String

4860  strTestL = "abcdefghijklmnopqrstuvwxyz"
4870  strTestU = UCase$(strTestL)
4880  ContainsAlpha = False

4890  For n = 1 To Len(s)
4900    If InStr(strTestL, Mid$(s, n, 1)) Or InStr(strTestU, Mid$(s, n, 1)) Then
4910      ContainsAlpha = True
4920      Exit Function
4930    End If
4940  Next

End Function
Public Function PasswordHasBeenUsed(ByVal Password As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

4950  On Error GoTo PasswordHasBeenUsed_Error

4960  PasswordHasBeenUsed = False

4970  sql = "SELECT Password FROM Users WHERE " & _
            "Password = '" & Password & "' " & _
            "COLLATE SQL_Latin1_General_CP1_CS_AS"
4980  Set tb = New Recordset
4990  RecOpenServer 0, tb, sql
5000  If Not tb.EOF Then
5010    PasswordHasBeenUsed = True
5020  Else
5030    sql = "SELECT Password FROM UsersArc WHERE " & _
              "Password = '" & Password & "' " & _
              "COLLATE SQL_Latin1_General_CP1_CS_AS"
5040    Set tb = New Recordset
5050    RecOpenServer 0, tb, sql
5060    If Not tb.EOF Then
5070      PasswordHasBeenUsed = True
5080    End If
5090  End If

5100  Exit Function

PasswordHasBeenUsed_Error:

      Dim strES As String
      Dim intEL As Integer

5110  intEL = Erl
5120  strES = Err.Description
5130  LogError "modPassword", "PasswordHasBeenUsed", intEL, strES, sql

End Function

Public Function AllUpperCase(stringToCheck As String) As Boolean

5140  AllUpperCase = StrComp(stringToCheck, UCase$(stringToCheck), vbBinaryCompare) = 0

End Function



