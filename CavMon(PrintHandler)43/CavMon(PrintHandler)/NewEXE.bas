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

