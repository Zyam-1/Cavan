VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3555
   ClientLeft      =   2235
   ClientTop       =   2010
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Update Lab No"
      Height          =   465
      Left            =   4110
      TabIndex        =   7
      Top             =   1890
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   720
      Left            =   180
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   240
      Width           =   720
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
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":0ECA
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   3
      Top             =   1140
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
      X2              =   5309.398
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
      Caption         =   $"frmAbout.frx":105D
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
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
35690     Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
35700     Unload Me
End Sub

Private Sub Command1_Click()
          '10    On Error GoTo Command1_Click_Error
          '
          '      Dim sql As String
          '      Dim Qry As String
          '      Dim tb As New ADODB.Recordset
          '      Dim i As Integer
          '
          '20    i = 0
          '30    sql = "SELECT Chart,PatName,DoB FROM demographics " & _
          '            "   WHERE ISNULL(Chart,'')<>'' " & _
          '            " AND ISNULL(PatName,'')<>'' " & _
          '            " AND ISNULL(DoB,'')<>'' " & _
          '            " Group By Chart,PatName,DoB"
          '40    Set tb = New Recordset
          '50    RecOpenServer 0, tb, sql
          '60    Do While Not tb.EOF
          '
          '70        Qry = " UPDATE demographics SET LABNO = '" & Val(FndMaxID("demographics", "LabNo", "")) + 1 & "'"
          '80        Qry = Qry & " WHERE "
          '90        Qry = Qry & " Chart = '" & tb!Chart & "'"
          '100       Qry = Qry & " AND PatName  = '" & tb!PatName & "'"
          '110       Qry = Qry & " AND DoB  = '" & Format(tb!DoB, "dd/MMM/yyyy") & "'"
          '120       Cnxn(0).Execute Qry
          '130       i = i + 1
          '140       DoEvents
          '
          '150       tb.MoveNext
          '160   Loop
          '
          '170   iMsg ("Updatation is completed")


35710     Exit Sub


Command1_Click_Error:

          Dim strES As String
          Dim intEL As Integer

35720     intEL = Erl
35730     strES = Err.Description
35740     LogError "frmMain", "Command1_Click", intEL, strES
End Sub

Private Sub Form_Load()
35750     Me.Caption = "About " & App.Title
35760     lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
35770     lblTitle.Caption = App.Title
          
35780     If UCase(UserName) = UCase("CRutter") Then
35790         Command1.Visible = True
35800     End If
End Sub

Public Sub StartSysInfo()
35810     On Error GoTo SysInfoErr
        
          Dim rc As Long
          Dim SysInfoPath As String
          
          ' Try To Get System Info Program Path\Name From Registry...
35820     If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
              ' Try To Get System Info Program Path Only From Registry...
35830     ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
              ' Validate Existance Of Known 32 Bit File Version
35840         If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
35850             SysInfoPath = SysInfoPath & "\MSINFO32.EXE"

                  ' Error - File Can Not Be Found...
35860         Else
35870             GoTo SysInfoErr
35880         End If
              ' Error - Registry Entry Can Not Be Found...
35890     Else
35900         GoTo SysInfoErr
35910     End If
          
35920     Call Shell(SysInfoPath, vbNormalFocus)
          
35930     Exit Sub
SysInfoErr:
35940     iMsg "System Information Is Unavailable At This Time", vbInformation
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
35950     rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

35960     If (rc <> ERROR_SUCCESS) Then         ' Handle Error...
35970         KeyVal = ""                                             ' Set Return Val To Empty String
35980         GetKeyValue = False                                     ' Return Failure
35990         rc = RegCloseKey(hKey)                                  ' Close Registry Key
36000         Exit Function
36010     End If

36020     tmpVal = String$(1024, 0)                             ' Allocate Variable Space
36030     KeyValSize = 1024                                       ' Mark Variable Size

          '------------------------------------------------------------
          ' Retrieve Registry Key Value...
          '------------------------------------------------------------
36040     rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
              KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

36050     If (rc <> ERROR_SUCCESS) Then          ' Handle Errors
36060         KeyVal = ""                                             ' Set Return Val To Empty String
36070         GetKeyValue = False                                     ' Return Failure
36080         rc = RegCloseKey(hKey)                                  ' Close Registry Key
36090         Exit Function
36100     End If

36110     If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
36120         tmpVal = Left$(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
36130     Else                                                    ' WinNT Does NOT Null Terminate String...
36140         tmpVal = Left$(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
36150     End If
          '------------------------------------------------------------
          ' Determine Key Value Type For Conversion...
          '------------------------------------------------------------
36160     Select Case KeyValType                                  ' Search Data Types...
              Case REG_SZ                                             ' String Registry Key Data Type
36170             KeyVal = tmpVal                                     ' Copy String Value
36180         Case REG_DWORD                                          ' Double Word Registry Key Data Type
36190             For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
36200                 KeyVal = KeyVal + Hex$(Asc(Mid$(tmpVal, i, 1)))   ' Build Value Char. By Char.
36210             Next
36220             KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
36230     End Select

36240     GetKeyValue = True                                      ' Return Success
36250     rc = RegCloseKey(hKey)                                  ' Close Registry Key

End Function


