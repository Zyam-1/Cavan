VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3825
   ClientLeft      =   2280
   ClientTop       =   1860
   ClientWidth     =   5715
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640.083
   ScaleMode       =   0  'User
   ScaleWidth      =   5366.681
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   3075
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   60
      TabIndex        =   7
      Top             =   3660
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.399
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "NetAcquire - Transfusion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1170
      Left            =   1050
      TabIndex        =   3
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.399
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":0BD4
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   270
      TabIndex        =   4
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
10      Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
10      Unload Me
End Sub

Private Sub Form_Load()
10        Me.Caption = "About " & App.Title
20        lblVersion.Caption = "Version " & App.Major & "." & App.Minor
30        lblTitle.Caption = App.Title
End Sub

Public Sub StartSysInfo()
10        On Error GoTo SysInfoErr
  
          Dim rc As Long
          Dim SysInfoPath As String
    
          ' Try To Get System Info Program Path\Name From Registry...
20        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
          ' Try To Get System Info Program Path Only From Registry...
30        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
              ' Validate Existance Of Known 32 Bit File Version
40            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
50                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"

              ' Error - File Can Not Be Found...
60            Else
70                GoTo SysInfoErr
80            End If
          ' Error - Registry Entry Can Not Be Found...
90        Else
100           GoTo SysInfoErr
110       End If
    
120       Call Shell(SysInfoPath, vbNormalFocus)
    
130       Exit Sub
SysInfoErr:
140       iMsg "System Information Is Unavailable At This Time", vbInformation
150       If TimedOut Then Unload Me: Exit Sub
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
          Dim i As Long                                           ' Loop Counter
          Dim rc As Long                                          ' Return Code
          Dim hKey As Long                                        ' Handle To An Open Registry Key
          Dim KeyValType As Long                                  ' Data Type Of A Registry Key
          Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
          Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
          '------------------------------------------------------------
          ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
          '------------------------------------------------------------
10        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
20        If (rc <> ERROR_SUCCESS) Then           ' Handle Error...
30          KeyVal = ""                                             ' Set Return Val To Empty String
40          GetKeyValue = False                                     ' Return Failure
50          rc = RegCloseKey(hKey) ' Close Registry Key
60          Exit Function
70        End If
    
80        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
90        KeyValSize = 1024                                       ' Mark Variable Size
    
          '------------------------------------------------------------
          ' Retrieve Registry Key Value...
          '------------------------------------------------------------
100       rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                               KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

110       If (rc <> ERROR_SUCCESS) Then          ' Handle Errors
120         KeyVal = ""                                             ' Set Return Val To Empty String
130         GetKeyValue = False                                     ' Return Failure
140         rc = RegCloseKey(hKey) ' Close Registry Key
150         Exit Function
160       End If
    
170       If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
180           tmpVal = Left$(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
190       Else                                                    ' WinNT Does NOT Null Terminate String...
200           tmpVal = Left$(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
210       End If
          '------------------------------------------------------------
          ' Determine Key Value Type For Conversion...
          '------------------------------------------------------------
220       Select Case KeyValType                                  ' Search Data Types...
          Case REG_SZ                                             ' String Registry Key Data Type
230           KeyVal = tmpVal                                     ' Copy String Value
240       Case REG_DWORD                                          ' Double Word Registry Key Data Type
250           For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
260               KeyVal = KeyVal + Hex$(Asc(Mid$(tmpVal, i, 1)))   ' Build Value Char. By Char.
270           Next
280           KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
290       End Select
    
300       GetKeyValue = True                                      ' Return Success
310       rc = RegCloseKey(hKey)                                  ' Close Registry Key

End Function
