Attribute VB_Name = "modTimedMsgBox"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'needed public for the Timer event
Public hwndMsgBox As Long

'custom user-defined type to pass
'info between procedures - easier than
'passing a long list of variables.
'Needed public for the Timer event
Public Type CUSTOM_MSG_PARAMS
   hOwnerThread         As Long
   hOwnerWindow         As Long
   dwStyle              As Long
   bUseTimer            As Boolean
   dwTimerDuration      As Long
   dwTimerInterval      As Long
   dwTimerExpireButton  As Long
   dwTimerCountDown     As Long
   dwTimerID            As Long
   sTitle               As String
   sPrompt              As String
End Type

Public cmp As CUSTOM_MSG_PARAMS

'Windows-defined uType parameters
Public Const MB_ICONINFORMATION As Long = &H40&
Private Const MB_ABORTRETRYIGNORE As Long = &H2&
Private Const MB_TASKMODAL As Long = &H2000&

'a const we define to identify our timer
Private Const MBTIMERID = 999

'Windows-defined MessageBox return values
Private Const IDOK As Long = 1
Private Const IDCANCEL As Long = 2
Private Const IDABORT As Long = 3
Private Const IDRETRY As Long = 4
Private Const IDIGNORE As Long = 5
Private Const IDYES As Long = 6
Private Const IDNO As Long = 7

'This section contains user-defined constants
'to represent the buttons/actions we are
'creating, based on the existing MessageBox
'constants. Doing this makes the code in
'the calling procedures more readable, since
'the messages match the buttons we're creating.
Public Const MB_SELECTBEGINSKIP As Long = MB_ABORTRETRYIGNORE
Public Const IDSELECT As Long = IDABORT
Public Const IDBEGIN As Long = IDRETRY
Public Const IDSKIP As Long = IDIGNORE
Public Const IDPROMPT As Long = &HFFFF&
Public Const IDTIMEDOK = 2
'misc API constants
Private Const WH_CBT = 5
Private Const GWL_HINSTANCE As Long = (-6)
Private Const HCBT_ACTIVATE As Long = 5
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_TIMER As Long = &H113

'UDT for passing data through the hook
Private Type MSGBOX_HOOK_PARAMS
   hwndOwner   As Long
   hHook       As Long
End Type

'need this declared at module level as
'it is used in the call and the hook proc
Private MHP As MSGBOX_HOOK_PARAMS

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long

Public Declare Function GetDlgItem Lib "user32" _
  (ByVal hDlg As Long, _
   ByVal nIDDlgItem As Long) As Long
   
Private Declare Function MessageBox Lib "user32" _
   Alias "MessageBoxA" _
  (ByVal hwnd As Long, _
   ByVal lpText As String, _
   ByVal lpCaption As String, _
   ByVal wType As Long) As Long
   
Public Declare Function PostMessage Lib "user32" _
   Alias "PostMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, lParam As Long) As Long
      
Public Declare Function PutFocus Lib "user32" _
   Alias "SetFocus" _
  (ByVal hwnd As Long) As Long
  
Public Declare Function SetDlgItemText Lib "user32" _
   Alias "SetDlgItemTextA" _
  (ByVal hDlg As Long, _
   ByVal nIDDlgItem As Long, _
   ByVal lpString As String) As Long
      
Private Declare Function SetWindowsHookEx Lib "user32" _
   Alias "SetWindowsHookExA" _
  (ByVal idHook As Long, _
   ByVal lpfn As Long, _
   ByVal hmod As Long, _
   ByVal dwThreadId As Long) As Long
      
Private Declare Function SetWindowText Lib "user32" _
   Alias "SetWindowTextA" _
  (ByVal hwnd As Long, _
   ByVal lpString As String) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" _
   (ByVal hHook As Long) As Long
   
Private Declare Function SetTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long, _
   ByVal uElapse As Long, _
   ByVal lpTimerFunc As Long) As Long
   
Private Declare Function KillTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long) As Long
    

Public Function MsgBoxHookProc(ByVal uMsg As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long) As Long

  'When the message box is about to be shown
  'change the button captions
   If uMsg = HCBT_ACTIVATE Then
   
     'in a HCBT_ACTIVATE message, wParam holds
     'the handle to the messagebox - save that
     'for the timer event
      hwndMsgBox = wParam
        
     'the ID's of the buttons on the message box
     'correspond exactly to the values they return,
     'so the same values can be used to identify
     'specific buttons in a SetDlgItemText call.
      SetDlgItemText wParam, IDSELECT, "Select.."
      SetDlgItemText wParam, IDBEGIN, "Begin"
      SetDlgItemText wParam, IDSKIP, "Skip"

     'we're done with the dialog, so release the hook
      UnhookWindowsHookEx MHP.hHook
         
   End If
   
  'return False to let normal processing continue
   MsgBoxHookProc = False

End Function

Public Sub ShowTimedBox(ByVal CallingForm As Form, _
                        ByVal Path As String)
    
Dim s() As String
Dim n As Integer

  'Display wrapper message box,
  'passing the CUSTOM_MSG_PARAMS
  'struct as the parameter.
    With cmp
      .sTitle = "NetAcquire - New Program"
      .dwStyle = MB_ICONINFORMATION Or 1
      .bUseTimer = True               'True = update once per dwTimerInterval
      .dwTimerDuration = 20          'time to wait seconds
      .dwTimerInterval = 1000         'countdown interval in milliseconds
      .dwTimerExpireButton = IDTIMEDOK  'message to return if timeout occurs
      .dwTimerCountDown = 0           '(re)set to 0
      .hOwnerThread = CallingForm.hwnd         'handle of form owning the thread on which
                                      'execution is proceeding.
                                      'The thread owner is always the calling form.
      .hOwnerWindow = CallingForm.hwnd         'who owns the dialog (me.hwnd or desktop).
                                      'GetDesktopWindow allows user-interaction
                                      'with the form while the dialog is displayed.
                                      'This may not be desirable, so set accordingly.
      'to enable the countdown TimerProc routine
      'to update the message box, place a %T variable
      'inside the message string.

      s = Split(Path, "\")
      .sPrompt = "A New Program is Available." & vbCrLf
      For n = 0 To UBound(s)
        .sPrompt = .sPrompt & s(n) & "\" & vbCrLf
      Next
      'get rid of trailing "\"
      .sPrompt = Left$(.sPrompt, Len(.sPrompt) - 3) & vbCrLf
      .sPrompt = .sPrompt & "Please contact your Administrator." & vbCrLf
      .sPrompt = .sPrompt & "This message will close in %T seconds." & Space$(12)
    End With

    TimedMessageBoxH cmp

End Sub


Public Function TimedMessageBoxH(cmp As CUSTOM_MSG_PARAMS) As Long

   Dim hInstance As Long
   Dim hThreadId As Long
   
  'Set up the hook
   hInstance = GetWindowLong(cmp.hOwnerThread, GWL_HINSTANCE)
   hThreadId = GetCurrentThreadId()

  'set up the MSGBOX_HOOK_PARAMS values
  'By specifying a Windows hook as one
  'of the params, we can intercept messages
  'sent by Windows and thereby manipulate
  'the dialog
   With MHP
      .hwndOwner = cmp.hOwnerWindow
      .hHook = SetWindowsHookEx(WH_CBT, _
                                AddressOf MsgBoxHookProc, _
                                hInstance, hThreadId)
   End With
   
  '(re) set the countdown (or rather 'count-up') value to 0
   cmp.dwTimerCountDown = 0
   
  'if bUseTimer, enable the timer. Because the
  'MessageBox API acts just as the MsgBox function
  'does (that is, creates a modal dialog), control
  'won't return to the next line until the dialog
  'is closed. This necessitates our starting the
  'timer before making the call.
  '
  'However, timer events will execute once the
  'modal dialog is shown, allowing us to use the
  'timer to dynamically modify the on-screen message!
  '
  'The handle passed to SetTimer is the form hwnd.
  'The event ID is set to the const we defined.
  'The interval is 1000 milliseconds, and the
  'callback is TimerProc
   If cmp.bUseTimer Then
      cmp.dwTimerID = SetTimer(cmp.hOwnerThread, _
                               MBTIMERID, _
                               1000, _
                               AddressOf TimerProc)
   End If

  'call the MessageBox API and return the
  'value as the result of the function.
  '
  'Replace original '%T' variable in the
  'original prompt with starting duration.
   TimedMessageBoxH = MessageBox(cmp.hOwnerWindow, _
                                 Replace$(cmp.sPrompt, "%T", CStr(cmp.dwTimerDuration)), _
                                 cmp.sTitle, _
                                 cmp.dwStyle)

  'in case the timer event didn't
  'suspend the timer, do it now
   If cmp.bUseTimer Then
      Call KillTimer(cmp.hOwnerThread, MBTIMERID)
   End If
   
End Function


Public Function TimerProc(ByVal hwnd As Long, _
                          ByVal uMsg As Long, _
                          ByVal idEvent As Long, _
                          ByVal dwTime As Long) As Long


   Dim hWndTargetBtn As Long
   Dim sUpdatedPrompt As String

  'watch for the WM_TIMER message
   Select Case uMsg
      Case WM_TIMER
      
        'compare to our event ID of '999'
         If idEvent = MBTIMERID Then
      
           'assure that there is messagebox to update
            If hwndMsgBox <> 0 Then
      
              'increment the counter
              'and update the caption string
              'with the new time
              'Note: VB5 users see comments below
               cmp.dwTimerCountDown = cmp.dwTimerCountDown + 1
               sUpdatedPrompt = Replace$(cmp.sPrompt, "%T", CStr(cmp.dwTimerDuration - cmp.dwTimerCountDown))
   
              'update the prompt message with the countdown value
               SetDlgItemText hwndMsgBox, IDPROMPT, sUpdatedPrompt
   
              'if the timer has 'expired' (the
              'count=duration), we need to
              'programmatically 'press' the button
              'specified as the default on timeout
               If cmp.dwTimerCountDown = cmp.dwTimerDuration Then
         
                 'nothing more to do, so
                 'we can kill this timer
                  Call KillTimer(cmp.hOwnerThread, MBTIMERID)
            
                 'now obtain the handle to the
                 'button designated as default
                 'if the timer expires
                  hWndTargetBtn = GetDlgItem(hwndMsgBox, cmp.dwTimerExpireButton)
            
                  If hWndTargetBtn <> 0 Then
            
                    'set the focus to the target button and
                    'simulate a click to close the dialog and
                    'return the correct value
                     Call PutFocus(hWndTargetBtn)
              
                    'need a DoEvents to allow PutFocus
                    'to actually put focus
                     DoEvents
               
                    'pretend a rodent pushed the button
                     Call PostMessage(hWndTargetBtn, WM_LBUTTONDOWN, 0, ByVal 0&)
                     Call PostMessage(hWndTargetBtn, WM_LBUTTONUP, 0, ByVal 0&)
            
                  End If  'If hWndTargetBtn
               End If  'If cmp.dwTimerCountDown
            End If  'If hwndMsgBox
         End If  'If idEvent
      Case Else
   End Select
   
End Function


