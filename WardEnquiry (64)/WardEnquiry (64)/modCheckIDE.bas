Attribute VB_Name = "modCheckIDE"
Option Explicit

'You can check this flag anywhere in your code
Public IsIDE As Boolean

Public Sub CheckIDE()
    
10    IsIDE = False
    
      'This line is only executed if
      'running in the IDE and then
      'returns True
20    Debug.Assert CheckIfInIDE
    
      'Use the IsIDE flag anywhere
      'For example
      '   If IsIDE Then
      '       MsgBox ("Running under IDE")
      '   Else
      '       MsgBox ("Running as EXE")
      '   End If

End Sub

Private Function CheckIfInIDE() As Boolean
    
      'This function will never be executed in an EXE

10    IsIDE = True        'set global flag

      'Set CheckIfInIDE or the Debug.Assert will Break
20    CheckIfInIDE = True

End Function
