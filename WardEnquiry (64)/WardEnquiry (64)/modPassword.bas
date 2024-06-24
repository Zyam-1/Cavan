Attribute VB_Name = "modPassword"
Option Explicit

Public Function NameHasBeenUsed(ByVal UserName As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo NameHasBeenUsed_Error

20    NameHasBeenUsed = False

30    sql = "SELECT Name FROM Users WHERE " & _
            "Name = '" & AddTicks(UserName) & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      NameHasBeenUsed = True
80    End If

90    Exit Function

NameHasBeenUsed_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modPassword", "NameHasBeenUsed", intEL, strES, sql

End Function


Public Function AllLowerCase(stringToCheck As String) As Boolean

10    AllLowerCase = StrComp(stringToCheck, LCase$(stringToCheck), vbBinaryCompare) = 0

End Function

Public Function ContainsNumeric(ByVal S As String) As Boolean

      Dim n As Integer

10    ContainsNumeric = False
20    For n = 1 To Len(S)
30      If InStr("0123456789", Mid$(S, n, 1)) Then
40        ContainsNumeric = True
50        Exit Function
60      End If
70    Next

End Function

Public Function ContainsAlpha(ByVal S As String) As Boolean

      Dim n As Integer
      Dim strTestL As String
      Dim strTestU As String

10    strTestL = "abcdefghijklmnopqrstuvwxyz"
20    strTestU = UCase$(strTestL)
30    ContainsAlpha = False

40    For n = 1 To Len(S)
50      If InStr(strTestL, Mid$(S, n, 1)) Or InStr(strTestU, Mid$(S, n, 1)) Then
60        ContainsAlpha = True
70        Exit Function
80      End If
90    Next

End Function
Public Function PasswordHasBeenUsed(ByVal Password As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo PasswordHasBeenUsed_Error

20    PasswordHasBeenUsed = False

30    sql = "SELECT Password FROM Users WHERE " & _
            "Password = '" & Password & "' " & _
            "COLLATE SQL_Latin1_General_CP1_CS_AS"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      PasswordHasBeenUsed = True
80    Else
90      sql = "SELECT Password FROM UsersArc WHERE " & _
              "Password = '" & Password & "' " & _
              "COLLATE SQL_Latin1_General_CP1_CS_AS"
100     Set tb = New Recordset
110     RecOpenServer 0, tb, sql
120     If Not tb.EOF Then
130       PasswordHasBeenUsed = True
140     End If
150   End If

160   Exit Function

PasswordHasBeenUsed_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "modPassword", "PasswordHasBeenUsed", intEL, strES, sql

End Function

Public Function AllUpperCase(stringToCheck As String) As Boolean

10    AllUpperCase = StrComp(stringToCheck, UCase$(stringToCheck), vbBinaryCompare) = 0

End Function



