Attribute VB_Name = "modBoxes"
Option Explicit

Public Function iMsg(Optional ByVal Message As String, _
                     Optional ByVal t As Integer = 0, _
                     Optional ByVal Caption As String = "NetAcquire", _
                     Optional ByVal BckColour As Long = &HC0C000, _
                     Optional ByVal MsgFontSize As Long) _
                     As Integer

      Dim SafeMsgBox As New fcdrMsgBox

10    With SafeMsgBox
20      .MsgFontSize = MsgFontSize
30      .BackColor = BckColour
40      .DisplayButtons = t And &H7
      '  .DefaultButton = (t And &H300) / 256
50      .ShowIcon = t And &H70
60      .Message = Message
70      .Caption = Caption
80      .Show vbModal
90      iMsg = .retval
100   End With

110   Unload SafeMsgBox
120   Set SafeMsgBox = Nothing

End Function

Public Function iBOX(ByVal Prompt As String, _
            Optional ByVal Title As String = "NetAcquire", _
            Optional ByVal Default As String, _
            Optional ByVal Pass As Boolean) As String

      Dim Box As New fcdrInputBox

10    With Box
20      .Password = Pass
30      .Caption = Title
40      .lblPrompt = Prompt
50      .txtInput = Default
60      .Show vbModal
70      iBOX = .retval
80    End With

90    Unload Box
100   Set Box = Nothing

End Function



