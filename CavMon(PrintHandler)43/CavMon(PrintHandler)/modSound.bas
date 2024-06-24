Attribute VB_Name = "modSound"
Option Explicit

Public Const SND_APPLICATION = &H80  ' look for application specific association
Public Const SND_ALIAS = &H10000     ' name is a WIN.INI [sounds] entry
Public Const SND_ALIAS_ID = &H110000 ' name is a WIN.INI [sounds] entry identifier
Public Const SND_ASYNC = &H1         ' play asynchronously
Public Const SND_FILENAME = &H20000  ' name is a file name
Public Const SND_LOOP = &H8          ' loop the sound until next sndPlaySound
Public Const SND_MEMORY = &H4        ' lpszSoundName points to a memory file
Public Const SND_NODEFAULT = &H2     ' silence not default, if sound not found
Public Const SND_NOSTOP = &H10       ' don't stop any currently playing sound
Public Const SND_NOWAIT = &H2000     ' don't wait if the driver is busy
Public Const SND_PURGE = &H40        ' purge non-static events for task
Public Const SND_RESOURCE = &H40004  ' name is a resource name or atom
Public Const SND_SYNC = &H0          ' play synchronously (default)

'usage
'PlaySound "C:\WINDOWS\MEDIA\TADA.WAV", ByVal 0&, SND_FILENAME Or SND_ASYNC

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
  (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long


