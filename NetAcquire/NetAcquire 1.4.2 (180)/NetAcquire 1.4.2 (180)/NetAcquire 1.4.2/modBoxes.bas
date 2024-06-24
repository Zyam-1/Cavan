Attribute VB_Name = "modBoxes"
Option Explicit

Public Function iMsg(Optional ByVal Message As String, _
                     Optional ByVal t As Integer = 0, _
                     Optional ByVal Caption As String = "NetAcquire", _
                     Optional ByVal BckColour As Long = &HC0C000, _
                     Optional ByVal MsgFontSize As Long) _
                     As Integer

      Dim SafeMsgBox As New fcdrMsgBox

59180 With SafeMsgBox
59190   .MsgFontSize = MsgFontSize
59200   .BackColor = BckColour
59210   .DisplayButtons = t And &H7
59220   .DefaultButton = t And &H300
59230   .ShowIcon = t And &H70
59240   .Message = Message
59250   .Caption = Caption
59260   .Show vbModal
59270   iMsg = .RetVal
59280 End With

59290 Unload SafeMsgBox
59300 Set SafeMsgBox = Nothing

End Function

Public Function iBOX(ByVal Prompt As String, _
            Optional ByVal Title As String = "NetAcquire", _
            Optional ByVal Default As String, _
            Optional ByVal Pass As Boolean) As String

      Dim Box As New fcdrInputBox

59310 With Box
59320   .Password = Pass
59330   .Caption = Title
59340   .lblPrompt = Prompt
59350   .txtInput = Default
59360   .Show vbModal
59370   iBOX = .RetVal
59380 End With

59390 Unload Box
59400 Set Box = Nothing

End Function

Public Function iTIME(ByVal Prompt As String, _
            Optional ByVal Title As String = "NetAcquire", _
            Optional ByVal Default As String = "__:__") As String

      Dim Box As New frmcdrInputTime

59410 With Box
59420   .Caption = Title
59430   .lblPrompt = Prompt
59440   .txtIP = Default
59450   .Show vbModal
59460   iTIME = .RetVal
59470 End With

59480 Unload Box
59490 Set Box = Nothing

End Function


