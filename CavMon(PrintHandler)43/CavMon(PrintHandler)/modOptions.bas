Attribute VB_Name = "modOptions"
Option Explicit

Public sysOptMicroOffset(0 To 0) As Variant '20,000,000
Public sysOptSemenOffset As Variant '10,000,000

Public sysOptSoundCritical As String
Public sysOptSoundInformation As String
Public sysOptSoundQuestion As String
Public sysOptSoundSevere As String


Public Sub LoadOptions()

10    On Error GoTo LoadOptions_Error

20    sysOptMicroOffset(0) = Val(GetOptionSetting("MicroOffset", "200000000000"))
30    sysOptSemenOffset = Val(GetOptionSetting("SemenOffset", "100000000000"))

40    sysOptSoundCritical = GetOptionSetting("SOUNDCRITICAL", "")
50    sysOptSoundInformation = GetOptionSetting("SOUNDINFORMATION", "")
60    sysOptSoundQuestion = GetOptionSetting("SOUNDQUESTION", "")
70    sysOptSoundSevere = GetOptionSetting("SOUNDSEVERE", "")

80    Exit Sub

LoadOptions_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "modOptions", "LoadOptions", intEL, strES

End Sub


